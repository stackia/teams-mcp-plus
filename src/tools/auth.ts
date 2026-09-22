import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";
import { tenantInputSchema } from "../tenants.js";

export function registerAuthTools(
  server: McpServer,
  graphService: GraphService,
  _readOnly: boolean
) {
  const annotations = {
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  };
  server.registerTool(
    "list_tenants",
    {
      title: "List Tenants",
      description:
        "List connected Microsoft Teams tenants, account names, granted scopes, and configured default. Use the tenantId in subsequent tool calls. Saved connections may require re-authentication; use auth_status to check live access.",
      inputSchema: {},
      annotations,
    },
    async () => {
      try {
        return {
          content: [
            { type: "text", text: JSON.stringify(await graphService.listTenants(), null, 2) },
          ],
        };
      } catch (error) {
        return {
          isError: true,
          content: [{ type: "text", text: error instanceof Error ? error.message : String(error) }],
        };
      }
    }
  );

  server.registerTool(
    "auth_status",
    {
      title: "Auth Status",
      description:
        "Check live authentication for the selected tenant. Returns its tenant ID, user profile, token expiration, and any authentication error.",
      inputSchema: tenantInputSchema,
      annotations,
    },
    async ({ tenantId }) => {
      try {
        const status = await (await graphService.forTenant(tenantId)).getAuthStatus();
        return {
          isError: !status.isAuthenticated,
          content: [{ type: "text", text: JSON.stringify(status, null, 2) }],
        };
      } catch (error) {
        return {
          isError: true,
          content: [{ type: "text", text: error instanceof Error ? error.message : String(error) }],
        };
      }
    }
  );
}
