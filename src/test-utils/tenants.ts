import { randomUUID } from "node:crypto";
import { READ_ONLY_SCOPES } from "../services/graph.js";
import { CLIENT_ID, type TenantProfile } from "../tenants.js";

export const TENANT_A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
export const TENANT_B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
export function tenantProfile(tenantId = TENANT_A): TenantProfile {
  return {
    tenantId,
    name: tenantId === TENANT_A ? "Work" : "Customer",
    clientId: CLIENT_ID,
    homeAccountId: "same-home-account-in-both-tenants",
    localAccountId: `local-${tenantId}`,
    username: "user@example.com",
    scopes: READ_ONLY_SCOPES,
    authenticatedAt: new Date().toISOString(),
    revision: randomUUID(),
  };
}
