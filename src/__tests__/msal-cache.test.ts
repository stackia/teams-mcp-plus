vi.unmock("node:fs");

import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PublicClientApplication } from "@azure/msal-node";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createCachePlugin } from "../msal-cache.js";
import { authorityFor, TenantStore } from "../tenants.js";
import { TENANT_A, TENANT_B, tenantProfile } from "../test-utils/tenants.js";

let directory: string;
let store: TenantStore;
beforeEach(async () => {
  directory = await fs.mkdtemp(join(tmpdir(), "teams-mcp-test-"));
  store = new TenantStore(join(directory, "credentials"));
});
afterEach(async () => {
  await fs.rm(directory, { recursive: true, force: true });
});

const context = (data: string, changed = true) =>
  ({
    cacheHasChanged: changed,
    tokenCache: { deserialize: vi.fn(), serialize: vi.fn().mockReturnValue(data) },
  }) as any;

describe("tenant credential storage", () => {
  it("round-trips saved accounts through the real MSAL cache and persists removals", async () => {
    const profile = tenantProfile();
    const environment = "login.windows.net";
    const accountKey = `${profile.homeAccountId}-${environment}-${profile.tenantId}`;
    await store.save(
      profile,
      JSON.stringify({
        Account: {
          [accountKey]: {
            home_account_id: profile.homeAccountId,
            environment,
            realm: profile.tenantId,
            local_account_id: profile.localAccountId,
            username: profile.username,
            authority_type: "MSSTS",
          },
        },
        IdToken: {},
        AccessToken: {},
        RefreshToken: {},
        AppMetadata: {},
      })
    );
    const createApp = () =>
      new PublicClientApplication({
        auth: { clientId: profile.clientId, authority: authorityFor(profile.tenantId) },
        cache: { cachePlugin: createCachePlugin(store, profile) },
      });
    const cache = createApp().getTokenCache();
    const accounts = await cache.getAllAccounts();
    expect(accounts).toHaveLength(1);
    expect(accounts[0]).toMatchObject({
      homeAccountId: profile.homeAccountId,
      localAccountId: profile.localAccountId,
      tenantId: profile.tenantId,
    });
    await cache.removeAccount(accounts[0]);
    expect(await createApp().getTokenCache().getAllAccounts()).toEqual([]);
    expect(JSON.parse(await fs.readFile(store.cachePath(profile), "utf8")).Account).toEqual({});
  });

  it("isolates concurrent cache reads and writes, including accounts with the same home ID", async () => {
    const a = tenantProfile();
    const b = tenantProfile(TENANT_B);
    await Promise.all([store.save(a, "cache-a"), store.save(b, "cache-b")]);
    const ca = context("new-a");
    const cb = context("new-b");
    const pa = createCachePlugin(store, a);
    const pb = createCachePlugin(store, b);
    await Promise.all([pa.beforeCacheAccess(ca), pb.beforeCacheAccess(cb)]);
    expect(ca.tokenCache.deserialize).toHaveBeenCalledWith("cache-a");
    expect(cb.tokenCache.deserialize).toHaveBeenCalledWith("cache-b");
    await Promise.all([pa.afterCacheAccess(ca), pb.afterCacheAccess(cb)]);
    expect(await fs.readFile(store.cachePath(a), "utf8")).toBe("new-a");
    expect(await fs.readFile(store.cachePath(b), "utf8")).toBe("new-b");
  });

  it("creates private credential directories and files", async () => {
    const profile = tenantProfile();
    await store.save(profile, "cache");
    expect((await fs.stat(store.directory)).mode & 0o777).toBe(0o700);
    expect((await fs.stat(join(store.directory, TENANT_A))).mode & 0o777).toBe(0o700);
    expect((await fs.stat(store.cachePath(profile))).mode & 0o777).toBe(0o600);
    expect((await fs.stat(join(store.directory, TENANT_A, "profile.json"))).mode & 0o777).toBe(
      0o600
    );
  });

  it("logs out only the selected tenant and prevents stale cache access", async () => {
    const a = tenantProfile();
    const b = tenantProfile(TENANT_B);
    await store.save(a, "a");
    await store.save(b, "b");
    const plugin = createCachePlugin(store, a);
    await store.logout(TENANT_A);
    await expect(plugin.beforeCacheAccess(context("stale"))).rejects.toThrow("not connected");
    await expect(plugin.afterCacheAccess(context("stale"))).rejects.toThrow("not connected");
    expect(await store.list()).toEqual([b]);
    expect(await fs.readFile(store.cachePath(b), "utf8")).toBe("b");
  });

  it("does not let a stale process overwrite a new login", async () => {
    const previous = tenantProfile();
    const current = tenantProfile();
    await store.save(previous, "old-cache");
    await store.save(current, "new-cache");
    await expect(
      createCachePlugin(store, previous).afterCacheAccess(context("stale"))
    ).rejects.toThrow("Authentication changed");
    expect(await fs.readFile(store.cachePath(current), "utf8")).toBe("new-cache");
    await store.logout(TENANT_A);
    expect(await fs.readdir(store.directory)).toEqual([]);
  });

  it("does not write unchanged caches", async () => {
    const profile = tenantProfile();
    await store.save(profile, "old");
    await createCachePlugin(store, profile).afterCacheAccess(context("new", false));
    expect(await fs.readFile(store.cachePath(profile), "utf8")).toBe("old");
  });

  it("rejects unsafe tenant IDs and corrupt metadata instead of picking another account", async () => {
    await expect(store.read("../outside")).rejects.toThrow();
    await expect(store.logout("common")).rejects.toThrow();
    await store.save(tenantProfile(), "cache");
    await fs.writeFile(join(store.directory, TENANT_A, "profile.json"), "invalid-json");
    await expect(store.read(TENANT_A)).rejects.toThrow();
  });

  it("surfaces missing cache and filesystem errors", async () => {
    const profile = tenantProfile();
    await store.save(profile, "cache");
    await fs.unlink(store.cachePath(profile));
    await expect(
      createCachePlugin(store, profile).beforeCacheAccess(context("x"))
    ).rejects.toThrow();
    const invalidStore = new TenantStore(join(directory, "file"));
    await fs.writeFile(invalidStore.directory, "not a directory");
    await expect(invalidStore.save(profile, "cache")).rejects.toThrow();
  });
});

describe("connection discovery during credential changes", () => {
  it("returns an empty list for a new store and ignores unrelated or incomplete directories", async () => {
    expect(await store.list()).toEqual([]);
    await fs.mkdir(join(store.directory, TENANT_A), { recursive: true });
    await fs.writeFile(join(store.directory, "README.txt"), "unrelated");
    await store.save(tenantProfile(TENANT_B), "b");
    expect((await store.list()).map((p) => p.tenantId)).toEqual([TENANT_B]);
  });

  it("surfaces malformed profiles and mismatched tenant IDs instead of implicitly switching tenants", async () => {
    await store.save(tenantProfile(), "a");
    await fs.writeFile(
      join(store.directory, TENANT_A, "profile.json"),
      JSON.stringify(tenantProfile(TENANT_B))
    );
    await expect(store.list()).rejects.toThrow("Tenant profile ID mismatch");
  });

  it("surfaces directory read errors rather than treating them as an empty store", async () => {
    await fs.writeFile(store.directory, "file");
    await expect(store.list()).rejects.toThrow();
  });

  it("keeps the previous login active if publishing its replacement fails", async () => {
    const original = tenantProfile();
    await store.save(original, "old");
    const rename = vi.spyOn(fs, "rename").mockRejectedValueOnce(new Error("disk failure"));
    try {
      await expect(store.save(tenantProfile(), "new")).rejects.toThrow("disk failure");
    } finally {
      rename.mockRestore();
    }
    expect(await store.read(TENANT_A)).toEqual(original);
    expect(
      (await fs.readdir(join(store.directory, TENANT_A))).some((name) => name.endsWith(".tmp"))
    ).toBe(false);
  });
});
