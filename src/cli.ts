import { randomUUID } from "node:crypto";
import { parseArgs } from "node:util";
import { PublicClientApplication } from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { FULL_SCOPES, GraphService, READ_ONLY_SCOPES } from "./services/graph.js";
import { authorityFor, CLIENT_ID, TenantStore, tenantIdSchema } from "./tenants.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

export function createServer(graphService: GraphService): McpServer {
  const server = new McpServer({ name: "teams-mcp-plus", version: "1.0.1" });
  registerAuthTools(server, graphService, graphService.readOnlyMode);
  registerUsersTools(server, graphService, graphService.readOnlyMode);
  registerTeamsTools(server, graphService, graphService.readOnlyMode);
  registerChatTools(server, graphService, graphService.readOnlyMode);
  registerSearchTools(server, graphService, graphService.readOnlyMode);
  return server;
}

export async function main(args = process.argv.slice(2)): Promise<void> {
  const { values, positionals } = parseArgs({
    args,
    allowPositionals: true,
    options: {
      tenant: { type: "string" },
      name: { type: "string" },
      "read-only": { type: "boolean" },
      all: { type: "boolean" },
      help: { type: "boolean", short: "h" },
    },
  });
  if (values.help || positionals[0] === "help") {
    console.log(`Microsoft Graph Teams MCP Plus — multiple tenants

teams-mcp-plus authenticate --tenant <tenant-id> [--name <label>] [--read-only]
teams-mcp-plus tenants                         List saved connections (no network check)
teams-mcp-plus check [--tenant <tenant-id>]     Check one tenant, or all connections
teams-mcp-plus logout --tenant <tenant-id>     Remove one tenant's local credentials
teams-mcp-plus logout --all                    Remove all saved tenant credentials
teams-mcp-plus [--tenant <default-tenant-id>] [--read-only]  Start MCP server

Environment:
  TEAMS_MCP_TENANT_ID     Default tenant (CLI --tenant takes precedence)
  TEAMS_MCP_CONFIG_DIR    Credential directory (default: ~/.teams-mcp)
  TEAMS_MCP_READ_ONLY=true  Disable write tools
  AUTH_TOKEN             Pre-issued Graph JWT, bound only to its tid claim

Tenant IDs must be Microsoft Entra directory GUIDs. Each tenant requires its own login.
Call list_tenants to discover connections. Tool calls accept tenantId; without a default, selection is automatic only for one connection.
`);
    return;
  }
  if (positionals.length > 1) throw new Error("Expected at most one command");
  const command = positionals[0];
  if (values.all && command !== "logout") throw new Error("--all is only supported for logout");
  if (values.name && command !== "authenticate" && command !== "auth")
    throw new Error("--name is only supported for authenticate");
  const tenant = values.tenant ?? process.env.TEAMS_MCP_TENANT_ID;
  const tenantId = tenant === undefined ? undefined : tenantIdSchema.parse(tenant);
  const readOnly = values["read-only"] || process.env.TEAMS_MCP_READ_ONLY === "true";
  const store = new TenantStore();

  if (command === "authenticate" || command === "auth") {
    if (!tenantId)
      throw new Error("Authentication requires --tenant <tenant-id> (or TEAMS_MCP_TENANT_ID)");
    const app = new PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: authorityFor(tenantId) },
    });
    // Login is staged in memory. Cancellation or failure leaves existing credentials intact.
    const result = await app.acquireTokenByDeviceCode({
      scopes: readOnly ? READ_ONLY_SCOPES : FULL_SCOPES,
      deviceCodeCallback: (response) => console.log(`Tenant: ${tenantId}\n${response.message}`),
    });
    if (
      !result?.account ||
      result.tenantId.toLowerCase() !== tenantId ||
      result.account.tenantId.toLowerCase() !== tenantId
    ) {
      throw new Error(
        "Authentication did not return an account in the requested tenant; credentials were not saved"
      );
    }
    await store.save(
      {
        tenantId,
        name: values.name ?? tenantId,
        clientId: CLIENT_ID,
        homeAccountId: result.account.homeAccountId,
        localAccountId: result.account.localAccountId,
        username: result.account.username,
        scopes: result.scopes,
        authenticatedAt: new Date().toISOString(),
        revision: randomUUID(),
      },
      app.getTokenCache().serialize()
    );
    console.log(
      `Authenticated ${result.account.username} in tenant ${tenantId} (${readOnly ? "read-only" : "full access"}).`
    );
    return;
  }
  if (command === "logout") {
    if (values.all && values.tenant) throw new Error("Use either --tenant or --all, not both");
    // Never infer a destructive target from a default environment setting.
    if (!values.all && !values.tenant)
      throw new Error("Logout requires explicit --tenant <tenant-id> or --all");
    const ids = values.all ? (await store.list()).map((p) => p.tenantId) : [tenantId as string];
    for (const id of ids) await store.logout(id);
    console.log(`Removed local credentials for ${ids.length} tenant(s).`);
    return;
  }
  const graphService = new GraphService({
    store,
    defaultTenantId: tenantId,
    readOnly,
    accessToken: process.env.AUTH_TOKEN,
  });
  if (command === "tenants") {
    console.log(JSON.stringify(await graphService.listTenants(), null, 2));
    return;
  }
  if (command === "check") {
    const ids = tenantId ? [tenantId] : (await graphService.listTenants()).map((p) => p.tenantId);
    if (ids.length === 0)
      throw new Error("No tenants connected. Run authenticate --tenant <tenant-id>");
    let failed = false;
    for (const id of ids) {
      try {
        const status = await (await graphService.forTenant(id)).getAuthStatus();
        console.log(JSON.stringify(status, null, 2));
        if (!status.isAuthenticated) failed = true;
      } catch (error) {
        console.error(`${id}: ${error instanceof Error ? error.message : String(error)}`);
        failed = true;
      }
    }
    if (failed) process.exitCode = 1;
    return;
  }
  if (command !== undefined) throw new Error(`Unknown command: ${command}. Use --help.`);
  await createServer(graphService).connect(new StdioServerTransport());
  console.error(
    `Teams MCP Plus server started${readOnly ? " (read-only)" : ""}; default tenant: ${tenantId ?? "automatic only for a single connection"}`
  );
}
