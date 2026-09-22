import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import { tenantInputSchema } from "../tenants.js";
import type { GraphApiResponse, User, UserSummary } from "../types/graph.js";

export function registerUsersTools(
  server: McpServer,
  graphServices: GraphService,
  _readOnly: boolean
) {
  // Search users
  server.registerTool(
    "search_users",
    {
      title: "Search Users",
      description: "Search tenant users by name or email, including IDs and mention text.",
      inputSchema: {
        ...tenantInputSchema,
        query: z.string().trim().min(1).describe("Name or email prefix"),
        limit: z.number().int().min(1).max(50).optional().default(10).describe("Maximum users"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ tenantId, query, limit = 10 }) => {
      try {
        const graphService = await graphServices.forTenant(tenantId);
        const client = await graphService.getClient();
        const escapedQuery = query.replaceAll("'", "''");
        const response = (await client
          .api("/users")
          .filter(
            `startswith(displayName,'${escapedQuery}') or startswith(mail,'${escapedQuery}') or startswith(userPrincipalName,'${escapedQuery}')`
          )
          .top(limit)
          .select("id,displayName,userPrincipalName,mail")
          .get()) as GraphApiResponse<User>;

        const userList = (response?.value ?? []).slice(0, limit).map((user: User) => ({
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
          mail: user.mail,
          id: user.id,
          mentionText:
            user.userPrincipalName?.split("@")[0] ||
            user.displayName?.toLowerCase().replace(/\s+/g, "") ||
            user.id,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(userList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Get specific user
  server.registerTool(
    "get_user",
    {
      title: "Get User",
      description: "Get a user profile; defaults to the current user.",
      inputSchema: {
        ...tenantInputSchema,
        userId: z.string().min(1).optional().describe("User ID or UPN. Omit for the current user."),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ tenantId, userId } = { tenantId: undefined, userId: undefined }) => {
      try {
        const graphService = await graphServices.forTenant(tenantId);
        const client = await graphService.getClient();
        const user = (await client
          .api(userId ? `/users/${encodeURIComponent(userId)}` : "/me")
          .get()) as User;

        const userSummary: UserSummary = {
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
          mail: user.mail,
          id: user.id,
          jobTitle: user.jobTitle,
          department: user.department,
          officeLocation: user.officeLocation,
        };

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(userSummary, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );
}
