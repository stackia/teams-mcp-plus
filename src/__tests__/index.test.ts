vi.unmock("node:fs");

import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PublicClientApplication } from "@azure/msal-node";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { main } from "../cli.js";
import { GraphService, READ_ONLY_SCOPES } from "../services/graph.js";
import { TenantStore } from "../tenants.js";
import { TENANT_A, TENANT_B, tenantProfile } from "../test-utils/tenants.js";

vi.mock("@azure/msal-node", () => ({ PublicClientApplication: vi.fn() }));
let directory: string;
let store: TenantStore;
let login: ReturnType<typeof vi.fn>;
let log: ReturnType<typeof vi.spyOn>;

beforeEach(async () => {
  vi.clearAllMocks();
  directory = await fs.mkdtemp(join(tmpdir(), "teams-mcp-cli-"));
  vi.stubEnv("TEAMS_MCP_CONFIG_DIR", directory);
  vi.stubEnv("TEAMS_MCP_TENANT_ID", undefined);
  vi.stubEnv("TEAMS_MCP_READ_ONLY", undefined);
  vi.stubEnv("AUTH_TOKEN", undefined);
  store = new TenantStore(directory);
  log = vi.spyOn(console, "log").mockImplementation(() => undefined);
  const profile = tenantProfile();
  login = vi.fn().mockResolvedValue({
    tenantId: TENANT_A,
    scopes: READ_ONLY_SCOPES,
    account: {
      tenantId: TENANT_A,
      homeAccountId: profile.homeAccountId,
      localAccountId: profile.localAccountId,
      username: profile.username,
    },
  });
  vi.mocked(PublicClientApplication).mockImplementation(function () {
    return {
      acquireTokenByDeviceCode: login,
      getTokenCache: () => ({ serialize: () => "new-cache" }),
    };
  } as any);
});
afterEach(async () => {
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
  process.exitCode = 0;
  await fs.rm(directory, { recursive: true, force: true });
});

describe("multi-tenant CLI", () => {
  it("authenticates an explicit tenant and preserves other connections", async () => {
    const b = tenantProfile(TENANT_B);
    await store.save(b, "b-cache");
    await main(["--tenant", TENANT_A, "authenticate", "--read-only", "--name", "Work"]);
    expect(PublicClientApplication).toHaveBeenCalledWith({
      auth: expect.objectContaining({ authority: `https://login.microsoftonline.com/${TENANT_A}` }),
    });
    expect(login).toHaveBeenCalledWith(expect.objectContaining({ scopes: READ_ONLY_SCOPES }));
    expect((await store.read(TENANT_A)).name).toBe("Work");
    expect(await store.read(TENANT_B)).toEqual(b);
    expect(await fs.readFile(store.cachePath(b), "utf8")).toBe("b-cache");
  });

  it("keeps existing credentials when login fails or returns the wrong tenant", async () => {
    const profile = tenantProfile();
    await store.save(profile, "original");
    login.mockResolvedValueOnce({ tenantId: TENANT_B, account: { tenantId: TENANT_B } });
    await expect(main(["auth", "--tenant", TENANT_A])).rejects.toThrow("requested tenant");
    login.mockResolvedValueOnce(null);
    await expect(main(["auth", "--tenant", TENANT_A])).rejects.toThrow();
    login.mockRejectedValueOnce(new Error("cancelled"));
    await expect(main(["auth", "--tenant", TENANT_A])).rejects.toThrow("cancelled");
    expect(await store.read(TENANT_A)).toEqual(profile);
    expect(await fs.readFile(store.cachePath(profile), "utf8")).toBe("original");
  });

  it("uses CLI tenant over environment, and supports the environment login default", async () => {
    vi.stubEnv("TEAMS_MCP_TENANT_ID", TENANT_B);
    await main(["auth", "--tenant", TENANT_A]);
    expect((await store.read(TENANT_A)).tenantId).toBe(TENANT_A);
    vi.stubEnv("TEAMS_MCP_TENANT_ID", TENANT_A);
    await main(["auth"]);
    expect(login).toHaveBeenCalledTimes(2);
  });

  it.each([
    ["authenticate"],
    ["authenticate", "--tenant"],
    ["authenticate", "--tenant", "common"],
    ["logout"],
    ["logout", "--all", "--tenant", TENANT_A],
    ["check", "--all"],
    ["check", "--name", "work"],
    ["unknown"],
    ["check", "extra"],
    ["--unknown"],
  ])("rejects invalid or ambiguous arguments %j", async (...args) => {
    await expect(main(args)).rejects.toThrow();
  });

  it("lists connections without displaying cache or account identifiers", async () => {
    await store.save(tenantProfile(), "secret-cache");
    await main(["tenants"]);
    const output = log.mock.calls.flat().join(" ");
    expect(output).toContain(TENANT_A);
    expect(output).not.toContain("secret-cache");
    expect(output).not.toContain("homeAccountId");
  });

  it("checks every tenant independently and reports authentication failure via exit status", async () => {
    await store.save(tenantProfile(), "a");
    await store.save(tenantProfile(TENANT_B), "b");
    const status = vi
      .spyOn(GraphService.prototype, "getAuthStatus")
      .mockResolvedValueOnce({
        tenantId: TENANT_A,
        isAuthenticated: false,
        error: "Interaction required",
      })
      .mockResolvedValueOnce({ tenantId: TENANT_B, isAuthenticated: true });
    await main(["check"]);
    expect(status).toHaveBeenCalledTimes(2);
    expect(process.exitCode).toBe(1);
    expect(log.mock.calls.flat().join(" ")).toContain(TENANT_B);
  });

  it("checks a selected tenant only", async () => {
    await store.save(tenantProfile(), "a");
    await store.save(tenantProfile(TENANT_B), "b");
    const status = vi
      .spyOn(GraphService.prototype, "getAuthStatus")
      .mockResolvedValue({ tenantId: TENANT_B, isAuthenticated: true });
    await main(["check", "--tenant", TENANT_B]);
    expect(status).toHaveBeenCalledTimes(1);
  });

  it("logs out one tenant or all, but never implicitly uses an environment default", async () => {
    await store.save(tenantProfile(), "a");
    await store.save(tenantProfile(TENANT_B), "b");
    vi.stubEnv("TEAMS_MCP_TENANT_ID", TENANT_A);
    await expect(main(["logout"])).rejects.toThrow("explicit");
    await main(["logout", "--tenant", TENANT_A]);
    expect((await store.list()).map((p) => p.tenantId)).toEqual([TENANT_B]);
    await main(["logout", "--all"]);
    expect(await store.list()).toEqual([]);
  });

  it("prints help without requiring credentials", async () => {
    await main(["--help"]);
    expect(log.mock.calls.flat().join(" ")).toContain("list_tenants");
  });
});
