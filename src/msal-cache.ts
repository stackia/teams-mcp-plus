import { promises as fs } from "node:fs";
import type { ICachePlugin } from "@azure/msal-node";
import { atomicWrite, type TenantProfile, type TenantStore } from "./tenants.js";

/** A cache is tied to one tenant and one login, including refreshes in long-running servers. */
export function createCachePlugin(store: TenantStore, profile: TenantProfile): ICachePlugin {
  return {
    async beforeCacheAccess(context) {
      await store.assertCurrent(profile);
      context.tokenCache.deserialize(await fs.readFile(store.cachePath(profile), "utf8"));
    },
    async afterCacheAccess(context) {
      if (!context.cacheHasChanged) return;
      await store.assertCurrent(profile);
      // Never recreate a directory after logout. A concurrent re-login uses a different path.
      await atomicWrite(store.cachePath(profile), context.tokenCache.serialize());
    },
  };
}
