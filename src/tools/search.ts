import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import { tenantInputSchema } from "../tenants.js";
import type { SearchHit, SearchRequest, SearchResponse } from "../types/graph.js";
import { formatMessageContent } from "../utils/html-to-markdown.js";

/**
 * Maps raw SearchHit objects from the Microsoft Search API into a
 * consistent, flat shape for tool responses.
 *
 * @param hits - Array of search hits from the Microsoft Search API
 * @param contentFormat - Format for message content: "markdown" or "raw"
 */
export function formatSearchHits(
  hits: SearchHit[],
  contentFormat: "raw" | "markdown" = "markdown"
) {
  return hits.map((hit) => ({
    id: hit.resource.id,
    summary: hit.summary,
    rank: hit.rank,
    content: formatMessageContent(hit.resource.body?.content, contentFormat),
    from: hit.resource.from?.user?.displayName,
    fromUserId: hit.resource.from?.user?.id,
    createdDateTime: hit.resource.createdDateTime,
    importance: hit.resource.importance,
    webLink: hit.resource.webLink,
    chatId: hit.resource.chatId,
    teamId: hit.resource.channelIdentity?.teamId,
    channelId: hit.resource.channelIdentity?.channelId,
  }));
}

export function registerSearchTools(
  server: McpServer,
  graphServices: GraphService,
  _readOnly: boolean
) {
  server.registerTool(
    "search_messages",
    {
      title: "Search Messages",
      description: "Search Teams messages with KQL and optional mention or time filters.",
      inputSchema: {
        ...tenantInputSchema,
        query: z
          .string()
          .trim()
          .min(1)
          .optional()
          .describe("KQL query; optional with mentionsMe or hours."),
        mentionsMe: z
          .boolean()
          .optional()
          .default(false)
          .describe("Only messages mentioning the current user."),
        hours: z
          .number()
          .min(1)
          .max(168)
          .optional()
          .describe(
            "Look back this many hours; defaults to 24 with mentionsMe, otherwise unlimited."
          ),
        from: z
          .number()
          .int()
          .min(0)
          .optional()
          .default(0)
          .describe("Search offset; use nextFrom for the next page."),
        size: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .default(25)
          .describe("Search page size before time filtering."),
        enableTopResults: z
          .boolean()
          .optional()
          .describe("Rank by relevance; defaults to false with mentionsMe, otherwise true."),
        contentFormat: z
          .enum(["raw", "markdown"])
          .optional()
          .default("markdown")
          .describe("Markdown or original message body."),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({
      tenantId,
      query,
      mentionsMe = false,
      hours,
      from = 0,
      size = 25,
      enableTopResults,
      contentFormat,
    }) => {
      try {
        const graphService = await graphServices.forTenant(tenantId);
        const client = await graphService.getClient();
        if (!query?.trim() && !mentionsMe && hours === undefined) {
          throw new Error("Provide query, mentionsMe or hours.");
        }
        const lookback = hours ?? (mentionsMe ? 24 : undefined);
        const now = Date.now();
        const since =
          lookback === undefined ? undefined : new Date(now - lookback * 3600000).toISOString();
        const clauses: string[] = [];
        if (query) clauses.push(mentionsMe || since ? `(${query})` : query);
        if (mentionsMe) clauses.push("IsMentioned:true");
        if (since) clauses.push(`sent>=${since.split("T")[0]}`);
        const queryString = clauses.join(" AND ");
        const searchRequest: SearchRequest = {
          entityTypes: ["chatMessage"],
          query: { queryString },
          from,
          size,
          enableTopResults: enableTopResults ?? !mentionsMe,
        };
        const response = (await client
          .api("/search/query")
          .post({ requests: [searchRequest] })) as SearchResponse;
        const container = response?.value?.[0]?.hitsContainers?.[0];
        const hits = container?.hits ?? [];
        // Teams documents date-level KQL. Enforce exact hours on each returned page.
        const filtered = since
          ? hits.filter((hit) => {
              const timestamp = Date.parse(hit.resource.createdDateTime ?? "");
              return timestamp >= Date.parse(since) && timestamp <= now;
            })
          : hits;
        const moreResultsAvailable = container?.moreResultsAvailable ?? false;
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  query: queryString,
                  from,
                  size,
                  total: container?.total ?? 0,
                  returned: filtered.length,
                  ...(since ? { since } : {}),
                  moreResultsAvailable,
                  nextFrom: moreResultsAvailable ? from + (hits.length || size) : null,
                  results: formatSearchHits(filtered, contentFormat ?? "markdown"),
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          isError: true,
          content: [{ type: "text", text: `❌ Error searching messages: ${errorMessage}` }],
        };
      }
    }
  );
}
