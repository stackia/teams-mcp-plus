vi.unmock("node:fs");

import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { authorityFor, TenantStore } from "../../tenants.js";
import { TENANT_A, TENANT_B, tenantProfile } from "../../test-utils/tenants.js";
import { FULL_SCOPES, GraphService, READ_ONLY_SCOPES } from "../graph.js";

vi.mock("@azure/msal-node", () => ({ PublicClientApplication: vi.fn() }));
vi.mock("@microsoft/microsoft-graph-client", () => ({ Client: { initWithMiddleware: vi.fn() } }));

let directory: string;
let store: TenantStore;
let graph: GraphService;
let accounts: Map<string, any[]>;
let acquire: ReturnType<typeof vi.fn>;
const providers: any[] = [];

beforeEach(async () => {
  vi.clearAllMocks();
  providers.length = 0;
  directory = await fs.mkdtemp(join(tmpdir(), "teams-mcp-graph-"));
  store = new TenantStore(directory);
  for (const id of [TENANT_A, TENANT_B]) await store.save(tenantProfile(id), "cache");
  accounts = new Map(
    [TENANT_A, TENANT_B].map((id) => {
      const profile = tenantProfile(id);
      return [
        id,
        [
          { homeAccountId: "wrong-account", localAccountId: "decoy", tenantId: id },
          {
            homeAccountId: profile.homeAccountId,
            localAccountId: profile.localAccountId,
            tenantId: id,
          },
        ],
      ];
    })
  );
  acquire = vi.fn().mockImplementation(async ({ account }) => ({
    account,
    tenantId: account.tenantId,
    accessToken: `token-${account.tenantId}`,
    expiresOn: new Date(Date.now() + 3600_000),
  }));
  vi.mocked(PublicClientApplication).mockImplementation(function (config: any) {
    const id = config.auth.authority.split("/").pop();
    return {
      getTokenCache: () => ({ getAllAccounts: async () => accounts.get(id) }),
      acquireTokenSilent: acquire,
    };
  } as any);
  vi.mocked(Client.initWithMiddleware).mockImplementation((config: any) => {
    providers.push(config.authProvider);
    return {
      api: () => ({
        get: async () => ({
          id: "current-user",
          displayName: await config.authProvider.getAccessToken(),
          userPrincipalName: "user@example.com",
        }),
      }),
    } as any;
  });
  graph = new GraphService({ store });
});
afterEach(async () => {
  await fs.rm(directory, { recursive: true, force: true });
});

describe("multi-tenant Graph service", () => {
  it("requires an explicit selection when multiple connections exist", async () => {
    await expect(graph.getClient()).rejects.toThrow("Multiple tenants");
    expect(PublicClientApplication).not.toHaveBeenCalled();
  });

  it("uses a configured default and lets an explicit tenant override it", async () => {
    graph = new GraphService({ store, defaultTenantId: TENANT_A.toUpperCase() });
    expect((await (await graph.forTenant()).getAuthStatus()).tenantId).toBe(TENANT_A);
    expect((await (await graph.forTenant(TENANT_B)).getAuthStatus()).tenantId).toBe(TENANT_B);
    expect((await graph.listTenants()).find((t) => t.tenantId === TENANT_A)?.isDefault).toBe(true);
  });

  it("automatically selects exactly one tenant, and handles an empty store", async () => {
    await store.logout(TENANT_B);
    expect((await graph.getAuthStatus()).tenantId).toBe(TENANT_A);
    await store.logout(TENANT_A);
    await expect(graph.forTenant()).rejects.toThrow("No tenants connected");
  });

  it("never falls back from an unknown or logged-out default", async () => {
    graph = new GraphService({ store, defaultTenantId: TENANT_A });
    await store.logout(TENANT_A);
    await expect(graph.forTenant()).rejects.toThrow("not connected");
    await expect(graph.forTenant("common")).rejects.toThrow();
    expect((await (await graph.forTenant(TENANT_B)).getAuthStatus()).isAuthenticated).toBe(true);
  });

  it("isolates simultaneous requests, picks the exact account, and pins refresh authority", async () => {
    const [a, b] = await Promise.all([graph.forTenant(TENANT_A), graph.forTenant(TENANT_B)]);
    const [sa, sb] = await Promise.all([a.getAuthStatus(), b.getAuthStatus()]);
    expect(sa.displayName).toBe(`token-${TENANT_A}`);
    expect(sb.displayName).toBe(`token-${TENANT_B}`);
    expect(sa.userId).toBe("current-user");
    expect(sb.userId).toBe("current-user");
    for (const id of [TENANT_A, TENANT_B]) {
      expect(acquire).toHaveBeenCalledWith(
        expect.objectContaining({
          authority: authorityFor(id),
          account: expect.objectContaining({ localAccountId: `local-${id}` }),
        })
      );
    }
    expect(PublicClientApplication).toHaveBeenCalledTimes(2);
  });

  it("coalesces concurrent initialization of the same tenant", async () => {
    const [a, b, c] = await Promise.all([
      graph.forTenant(TENANT_A),
      graph.forTenant(TENANT_A),
      graph.forTenant(TENANT_A),
    ]);
    const clients = await Promise.all([a.getClient(), b.getClient(), c.getClient()]);
    expect(clients[0]).toBe(clients[1]);
    expect(clients[1]).toBe(clients[2]);
    expect(PublicClientApplication).toHaveBeenCalledTimes(1);
  });

  it("rejects stale clients and auth providers after logout without affecting another tenant", async () => {
    const a = await graph.forTenant(TENANT_A);
    await a.getClient();
    const authProvider = providers[0];
    await store.logout(TENANT_A);
    await expect(a.getClient()).rejects.toThrow("not connected");
    await expect(authProvider.getAccessToken()).rejects.toThrow("not connected");
    expect((await (await graph.forTenant(TENANT_B)).getAuthStatus()).isAuthenticated).toBe(true);
  });

  it("rebuilds a session after re-login and prevents a request switching accounts midway", async () => {
    const before = await graph.forTenant(TENANT_A);
    await before.getClient();
    await store.save(tenantProfile(), "new-cache");
    await expect(before.getClient()).rejects.toThrow("Authentication changed");
    const after = await graph.forTenant(TENANT_A);
    expect(after).not.toBe(before);
    expect((await after.getAuthStatus()).isAuthenticated).toBe(true);
  });

  it("rejects tokens with a different tenant or account", async () => {
    const a = await graph.forTenant(TENANT_A);
    acquire.mockResolvedValueOnce({
      tenantId: TENANT_B,
      account: accounts.get(TENANT_B)?.[1],
      accessToken: "wrong",
    });
    await expect(a.getClient()).rejects.toThrow("does not match");
    acquire.mockResolvedValueOnce({
      tenantId: TENANT_A,
      account: accounts.get(TENANT_A)?.[0],
      accessToken: "wrong",
    });
    await expect(a.getClient()).rejects.toThrow("does not match");
    expect(Client.initWithMiddleware).not.toHaveBeenCalled();
  });

  it("does not use the first cached account when the saved account is missing", async () => {
    accounts.set(TENANT_A, [accounts.get(TENANT_A)?.[0]]);
    await expect((await graph.forTenant(TENANT_A)).getClient()).rejects.toThrow(
      "Saved account not found"
    );
    expect(acquire).not.toHaveBeenCalled();
  });

  it("retries failed initialization and refreshes via the auth provider", async () => {
    const a = await graph.forTenant(TENANT_A);
    acquire.mockRejectedValueOnce(new Error("Interaction required"));
    expect((await a.getAuthStatus()).isAuthenticated).toBe(false);
    await a.getClient();
    expect(await providers[0].getAccessToken()).toBe(`token-${TENANT_A}`);
    expect(acquire).toHaveBeenCalledTimes(3);
  });

  it("keeps read-only scopes separate from full mode and respects per-tenant grants", async () => {
    const full = tenantProfile();
    full.scopes = FULL_SCOPES;
    await store.save(full, "cache");
    const ro = new GraphService({ store, readOnly: true });
    await (await ro.forTenant(TENANT_A)).getClient();
    expect(acquire).toHaveBeenLastCalledWith(expect.objectContaining({ scopes: READ_ONLY_SCOPES }));
    await (await graph.forTenant(TENANT_A)).getClient();
    expect(acquire).toHaveBeenLastCalledWith(expect.objectContaining({ scopes: FULL_SCOPES }));
    await (await graph.forTenant(TENANT_B)).getClient();
    expect(acquire).toHaveBeenLastCalledWith(expect.objectContaining({ scopes: READ_ONLY_SCOPES }));
  });
});

function jwt(claims: Record<string, unknown>) {
  return `header.${Buffer.from(JSON.stringify({ aud: "https://graph.microsoft.com", tid: TENANT_A, exp: Math.floor(Date.now() / 1000) + 3600, ...claims })).toString("base64url")}.signature`;
}

describe("injected tokens", () => {
  it("binds AUTH_TOKEN to tid while other tenants keep using their own MSAL credentials", async () => {
    const token = jwt({});
    graph = new GraphService({ store, accessToken: token });
    const a = await graph.forTenant(TENANT_A);
    expect((await a.getAuthStatus()).displayName).toBe(token);
    expect(PublicClientApplication).not.toHaveBeenCalled();
    expect((await (await graph.forTenant(TENANT_B)).getAuthStatus()).displayName).toBe(
      `token-${TENANT_B}`
    );
    expect(PublicClientApplication).toHaveBeenCalledTimes(1);
    expect(await graph.listTenants()).toHaveLength(2);
  });

  it.each([
    "bad-token",
    jwt({ tid: "common" }),
    jwt({ exp: 0 }),
    jwt({ aud: "https://other.example" }),
    jwt({ exp: "invalid" }),
  ])("rejects malformed or wrongly scoped tokens", (accessToken) => {
    expect(() => new GraphService({ store, accessToken })).toThrow();
  });

  it("accepts the Graph application ID audience and rejects a token that expires at runtime", async () => {
    const now = Date.now();
    graph = new GraphService({
      store,
      accessToken: jwt({ aud: "00000003-0000-0000-c000-000000000000" }),
    });
    const a = await graph.forTenant(TENANT_A);
    await a.getClient();
    const clock = vi.spyOn(Date, "now").mockReturnValue(now + 7200_000);
    try {
      await expect(providers[0].getAccessToken()).rejects.toThrow("expired");
    } finally {
      clock.mockRestore();
    }
  });
});
