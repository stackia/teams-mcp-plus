import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import { formatSearchHits, registerSearchTools } from "../search.js";

// Mock the Graph service
const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

// Mock the MCP server
const mockServer = {
  registerTool: vi.fn(),
} as unknown as McpServer;

// Mock client responses
const mockClient = {
  api: vi.fn(),
} as unknown as Client;

// Helper to build a standard search response with hits
function makeSearchResponse(hits: any[], total = hits.length, moreResultsAvailable = false) {
  return {
    value: [{ hitsContainers: [{ hits, total, moreResultsAvailable }] }],
  };
}

// Reusable hit fixture
function makeHit(overrides: Record<string, any> = {}) {
  return {
    hitId: "hit1",
    rank: 1,
    summary: "Found message",
    resource: {
      "@odata.type": "#microsoft.graph.chatMessage",
      id: "msg1",
      body: { content: "Hello world" },
      from: { user: { displayName: "John Doe", id: "user1" } },
      createdDateTime: "2025-01-01T10:00:00Z",
      chatId: "chat123",
      importance: "normal",
      webLink: "https://teams.microsoft.com/msg1",
      ...overrides,
    },
  };
}

describe("Search Tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGraphService.forTenant = vi.fn().mockReturnThis();
    mockGraphService.getClient = vi.fn().mockResolvedValue(mockClient);
  });

  describe("registerSearchTools", () => {
    it("should register only search_messages", () => {
      registerSearchTools(mockServer, mockGraphService, false);

      expect(mockServer.registerTool).toHaveBeenCalledTimes(1);
      expect(mockServer.registerTool).toHaveBeenCalledWith(
        "search_messages",
        expect.objectContaining({
          title: "Search Messages",
          description: expect.any(String),
        }),
        expect.any(Function)
      );
    });
  });

  describe("formatSearchHits", () => {
    it("should map hit fields to a flat shape", () => {
      const hits = [
        makeHit({
          channelIdentity: { teamId: "team1", channelId: "channel1" },
        }),
      ];

      const results = formatSearchHits(hits as any);

      expect(results).toHaveLength(1);
      expect(results[0]).toEqual({
        id: "msg1",
        summary: "Found message",
        rank: 1,
        content: "Hello world",
        from: "John Doe",
        fromUserId: "user1",
        createdDateTime: "2025-01-01T10:00:00Z",
        importance: "normal",
        webLink: "https://teams.microsoft.com/msg1",
        chatId: "chat123",
        teamId: "team1",
        channelId: "channel1",
      });
    });

    it("should handle missing optional fields gracefully", () => {
      const hits = [
        {
          hitId: "hit1",
          rank: 1,
          summary: "",
          resource: {
            "@odata.type": "#microsoft.graph.chatMessage",
            id: "msg1",
          },
        },
      ];

      const results = formatSearchHits(hits as any);

      expect(results[0].content).toBeUndefined();
      expect(results[0].from).toBeUndefined();
      expect(results[0].webLink).toBeUndefined();
    });
  });

  describe("search_messages", () => {
    let handler: (args: any) => Promise<any>;

    beforeEach(() => {
      registerSearchTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.registerTool)
        .mock.calls.find(([name]) => name === "search_messages");
      handler = call?.[2] as unknown as (args: any) => Promise<any>;
    });

    it("should send a single search request with provided parameters", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "hello", from: 0, size: 25, enableTopResults: true });

      expect(mockClient.api).toHaveBeenCalledWith("/search/query");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        requests: [
          {
            entityTypes: ["chatMessage"],
            query: { queryString: "hello" },
            from: 0,
            size: 25,
            enableTopResults: true,
          },
        ],
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.query).toBe("hello");
      expect(parsed.total).toBe(1);
      expect(parsed.results).toHaveLength(1);
      expect(parsed.results[0].content).toBe("Hello world");
    });

    it("should pass KQL query strings through unmodified", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await handler({
        query: "from:bob hasAttachment:true sent>=2025-01-01",
      });

      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall.requests[0].query.queryString).toBe(
        "from:bob hasAttachment:true sent>=2025-01-01"
      );
    });

    it("should support pagination with from offset", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()], 50, true)),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "test", from: 25, size: 25 });

      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall.requests[0].from).toBe(25);
      expect(postCall.requests[0].size).toBe(25);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.moreResultsAvailable).toBe(true);
      expect(parsed.from).toBe(25);
    });

    it("should return friendly message when no results found", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "nonexistent" });
      expect(JSON.parse(result.content[0].text).results).toEqual([]);
    });

    it("should handle empty hitsContainers", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ value: [{ hitsContainers: [] }] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "nonexistent" });
      expect(JSON.parse(result.content[0].text).results).toEqual([]);
    });

    it("should return error message on API failure", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Search API error")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "error" });
      expect(result.content[0].text).toBe("❌ Error searching messages: Search API error");
    });

    it("should handle non-Error thrown values", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue("string error"),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "error" });
      expect(result.content[0].text).toBe("❌ Error searching messages: Unknown error occurred");
    });
  });

  describe("mention and time filters", () => {
    let handler: (args: any) => Promise<any>;
    let post: ReturnType<typeof vi.fn>;
    beforeEach(() => {
      vi.useFakeTimers();
      vi.setSystemTime(new Date("2026-09-23T12:30:00Z"));
      registerSearchTools(mockServer, mockGraphService, false);
      handler = vi.mocked(mockServer.registerTool).mock.calls[0][2] as any;
      post = vi.fn().mockResolvedValue(makeSearchResponse([]));
      mockClient.api = vi.fn().mockReturnValue({ post });
    });
    afterEach(() => vi.useRealTimers());

    it("combines an OR query with mention and date constraints, without a /me lookup", async () => {
      post.mockResolvedValue(
        makeSearchResponse([makeHit({ createdDateTime: "2026-09-23T12:00:00Z" })])
      );
      const result = await handler({ query: "alpha OR beta", mentionsMe: true });
      expect(mockClient.api).toHaveBeenCalledExactlyOnceWith("/search/query");
      expect(post.mock.calls[0][0].requests[0]).toMatchObject({
        query: { queryString: "(alpha OR beta) AND IsMentioned:true AND sent>=2026-09-22" },
        enableTopResults: false,
        from: 0,
        size: 25,
      });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.since).toBe("2026-09-22T12:30:00.000Z");
      expect(parsed.results).toHaveLength(1);
      expect(parsed).not.toHaveProperty("mentions");
    });

    it("filters at the exact hour boundary and paginates by raw hits", async () => {
      post.mockResolvedValue(
        makeSearchResponse(
          [
            makeHit({ id: "too-old", createdDateTime: "2026-09-23T11:29:59Z" }),
            makeHit({ id: "boundary", createdDateTime: "2026-09-23T11:30:00Z" }),
            makeHit({ id: "recent", createdDateTime: "2026-09-23T12:29:59Z" }),
            makeHit({ id: "unknown", createdDateTime: undefined }),
            makeHit({ id: "invalid", createdDateTime: "invalid" }),
          ],
          5,
          true
        )
      );
      const result = await handler({ hours: 1, from: 10, size: 5 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.results.map((hit: any) => hit.id)).toEqual(["boundary", "recent"]);
      expect(parsed.returned).toBe(2);
      expect(parsed.total).toBe(5);
      expect(parsed.nextFrom).toBe(15);
      expect(parsed.moreResultsAvailable).toBe(true);
    });

    it("keeps pagination when an entire page falls outside the time range", async () => {
      post.mockResolvedValue(makeSearchResponse([makeHit()], 1, true));
      const parsed = JSON.parse((await handler({ mentionsMe: true, from: 4 })).content[0].text);
      expect(parsed.results).toEqual([]);
      expect(parsed.nextFrom).toBe(5);
      expect(parsed.moreResultsAvailable).toBe(true);
    });

    it("supports a custom lookback and explicit ranking", async () => {
      const parsed = JSON.parse(
        (await handler({ mentionsMe: true, hours: 168, enableTopResults: true })).content[0].text
      );
      expect(parsed.since).toBe("2026-09-16T12:30:00.000Z");
      expect(parsed.results).toEqual([]);
      expect(parsed.nextFrom).toBeNull();
      expect(post.mock.calls[0][0].requests[0].enableTopResults).toBe(true);
    });

    it("does not add a time restriction to ordinary queries", async () => {
      const parsed = JSON.parse((await handler({ query: "test" })).content[0].text);
      expect(parsed).not.toHaveProperty("since");
      expect(post.mock.calls[0][0].requests[0]).toMatchObject({
        query: { queryString: "test" },
        enableTopResults: true,
      });
    });

    it("rejects requests without a query or filter", async () => {
      expect((await handler({})).isError).toBe(true);
      expect(post).not.toHaveBeenCalled();
    });

    it("surfaces mention-search failures", async () => {
      post.mockRejectedValue(new Error("Graph search failed"));
      const result = await handler({ mentionsMe: true });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Graph search failed");
    });
  });
});
