import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import { createMockMcpServer } from "../../test-utils/setup.js";
import { TENANT_B } from "../../test-utils/tenants.js";
import { registerChatTools } from "../chats.js";
import { registerTeamsTools } from "../teams.js";

describe("Chat details and read state", () => {
  let server: ReturnType<typeof createMockMcpServer>;
  const get = vi.fn();
  const post = vi.fn();
  const api = vi.fn().mockReturnValue({ get, post });
  const status = { isAuthenticated: true, userId: "current-user", tenantId: TENANT_B };
  const bound = {
    getClient: vi.fn().mockResolvedValue({ api }),
    getAuthStatus: vi.fn().mockResolvedValue(status),
  };
  const forTenant = vi.fn().mockResolvedValue(bound);
  const root = { forTenant } as unknown as GraphService;
  const call = (name: string, args: Record<string, unknown>) =>
    server.getTool(name).handler({ tenantId: TENANT_B, ...args });
  const message = {
    id: "message",
    body: { contentType: "html", content: "<b>Hello</b>" },
    from: { user: { id: "sender", displayName: "Sender" } },
    mentions: [{ id: 0, mentionText: "Someone", mentioned: { user: { id: "someone" } } }],
    reactions: [{ reactionType: "like", user: { user: { id: "reactor" } } }],
    attachments: [
      {
        id: "file",
        name: "doc.pdf",
        contentType: "reference",
        contentUrl: "https://example.com/doc.pdf",
      },
    ],
    lastModifiedDateTime: "2026-09-23T00:00:00Z",
  };

  beforeEach(() => {
    vi.clearAllMocks();
    get.mockReset();
    post.mockReset();
    api.mockReturnValue({ get, post });
    forTenant.mockResolvedValue(bound);
    bound.getClient.mockResolvedValue({ api });
    bound.getAuthStatus.mockResolvedValue(status);
    server = createMockMcpServer();
    registerChatTools(server as any, root, false);
    registerTeamsTools(server as any, root, false);
  });

  it("fetches a chat message directly regardless of list options, retaining details", async () => {
    get.mockResolvedValue(message);
    const result = await call("get_chat_messages", {
      chatId: "chat/#",
      messageId: "message",
      since: "2099-01-01T00:00:00Z",
      fromUser: "other",
      descending: false,
      fetchAll: true,
    });
    expect(forTenant).toHaveBeenCalledExactlyOnceWith(TENANT_B);
    expect(api).toHaveBeenCalledExactlyOnceWith("/chats/chat%2F%23/messages/message");
    expect(JSON.parse(result.content[0].text)).toMatchObject({
      totalReturned: 1,
      hasMore: false,
      messages: [
        {
          id: "message",
          content: "**Hello**",
          from: "Sender",
          sender: message.from,
          mentions: message.mentions,
          reactions: message.reactions,
          attachments: message.attachments,
          lastModifiedDateTime: message.lastModifiedDateTime,
        },
      ],
    });
    expect(post).not.toHaveBeenCalled();
  });

  it.each([undefined, "reply/#"])("reads a channel message or reply (%s)", async (replyId) => {
    get.mockResolvedValue(message);
    const result = await call("get_channel_messages", {
      teamId: "team",
      channelId: "channel",
      messageId: "message",
      replyId,
      contentFormat: "raw",
    });
    expect(api).toHaveBeenCalledExactlyOnceWith(
      `/teams/team/channels/channel/messages/message${replyId ? "/replies/reply%2F%23" : ""}`
    );
    expect(JSON.parse(result.content[0].text).messages[0].content).toBe("<b>Hello</b>");
  });

  it("does not interpret plain text as HTML", async () => {
    get.mockResolvedValue({
      id: "message",
      body: { contentType: "text", content: "<b>literal</b>" },
    });
    const result = await call("get_chat_messages", { chatId: "chat", messageId: "message" });
    expect(JSON.parse(result.content[0].text).messages[0].content).toBe("<b>literal</b>");
  });

  it("rejects a reply without a parent message", async () => {
    const result = await call("get_channel_messages", {
      teamId: "team",
      channelId: "channel",
      replyId: "reply",
    });
    expect(result.isError).toBe(true);
    expect(api).not.toHaveBeenCalled();
  });

  it("lists a thread's replies oldest first using the messages envelope", async () => {
    get.mockResolvedValue({
      value: [
        { ...message, id: "later", createdDateTime: "2026-09-23T12:00:00Z" },
        { ...message, id: "earlier", createdDateTime: "2026-09-23T11:00:00Z" },
      ],
      "@odata.nextLink": "next-page",
    });
    const result = await call("get_channel_messages", {
      teamId: "team",
      channelId: "channel",
      messageId: "root/#",
      listReplies: true,
      limit: 2,
    });
    expect(api).toHaveBeenCalledExactlyOnceWith(
      "/teams/team/channels/channel/messages/root%2F%23/replies?$top=2"
    );
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.parentMessageId).toBe("root/#");
    expect(parsed.hasMore).toBe(true);
    expect(parsed.messages.map((m: any) => m.id)).toEqual(["earlier", "later"]);
  });

  it.each([{ listReplies: true }, { listReplies: true, messageId: "root", replyId: "reply" }])(
    "rejects ambiguous reply-list arguments %j",
    async (args) => {
      const result = await call("get_channel_messages", {
        teamId: "team",
        channelId: "channel",
        ...args,
      });
      expect(result.isError).toBe(true);
      expect(api).not.toHaveBeenCalled();
    }
  );

  describe.each([
    {
      name: "get_chat_messages",
      field: "messageId",
      args: { chatId: "chat/#" },
      prefix: "/chats/chat%2F%23/messages/",
    },
    {
      name: "get_channel_messages",
      field: "messageId",
      args: { teamId: "team/#", channelId: "channel/#" },
      prefix: "/teams/team%2F%23/channels/channel%2F%23/messages/",
    },
    {
      name: "get_channel_messages",
      field: "replyId",
      args: { teamId: "team/#", channelId: "channel/#", messageId: "root/#" },
      prefix: "/teams/team%2F%23/channels/channel%2F%23/messages/root%2F%23/replies/",
    },
  ])("batch $name by $field", ({ name, field, args, prefix }) => {
    it("reads sequentially in input order, ignoring list limits and retaining message details", async () => {
      let resolveFirst!: (value: typeof message) => void;
      get
        .mockImplementationOnce(
          () =>
            new Promise((resolve) => {
              resolveFirst = resolve;
            })
        )
        .mockResolvedValueOnce({
          ...message,
          id: "second",
          body: { contentType: "text", content: "<b>literal</b>" },
        });
      const pending = call(name, {
        ...args,
        [field]: ["first/#", "second"],
        limit: 1,
        since: "2099-01-01T00:00:00Z",
        descending: false,
        fetchAll: true,
      });
      await vi.waitFor(() => expect(get).toHaveBeenCalledTimes(1));
      expect(api).toHaveBeenCalledExactlyOnceWith(`${prefix}first%2F%23`);
      resolveFirst({ ...message, id: "first/#" });
      const result = await pending;
      expect(result.isError).toBeUndefined();
      expect(forTenant).toHaveBeenCalledExactlyOnceWith(TENANT_B);
      expect(api.mock.calls.map(([path]) => path)).toEqual([
        `${prefix}first%2F%23`,
        `${prefix}second`,
      ]);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toMatchObject({
        totalRequested: 2,
        totalReturned: 2,
        hasMore: false,
        errors: [],
      });
      expect(parsed.messages).toMatchObject([
        {
          id: "first/#",
          content: "**Hello**",
          attachments: message.attachments,
          reactions: message.reactions,
          sender: message.from,
        },
        { id: "second", content: "<b>literal</b>" },
      ]);
      expect(post).not.toHaveBeenCalled();
    });

    it("retains successes and reports failures per requested ID", async () => {
      get
        .mockRejectedValueOnce(new Error("Graph denied access"))
        .mockResolvedValueOnce(null)
        .mockResolvedValueOnce({ ...message, id: "ok" });
      const result = await call(name, {
        ...args,
        [field]: ["denied", "missing", "ok"],
        contentFormat: "raw",
      });
      expect(result.isError).toBeUndefined();
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toMatchObject({
        totalRequested: 3,
        totalReturned: 1,
        hasMore: false,
        messages: [{ id: "ok", content: "<b>Hello</b>" }],
      });
      const target = (id: string) =>
        field === "replyId" ? { messageId: "root/#", replyId: id } : { messageId: id };
      expect(parsed.errors).toEqual([
        { ...target("denied"), error: "Graph denied access" },
        { ...target("missing"), error: "Message not found." },
      ]);
      expect(get).toHaveBeenCalledTimes(3);
    });

    it("marks an entirely failed batch as an error", async () => {
      get.mockRejectedValue(new Error("Not found"));
      const result = await call(name, { ...args, [field]: ["missing"] });
      expect(result.isError).toBe(true);
      expect(JSON.parse(result.content[0].text)).toMatchObject({
        totalRequested: 1,
        totalReturned: 0,
        messages: [],
        errors: [{ error: "Not found" }],
      });
    });

    it("accepts scalar IDs and nonempty arrays of at most 50 nonempty IDs", () => {
      const schema = server.getTool(name).schema[field];
      for (const value of ["id", ["id"], Array(50).fill("id")]) {
        expect(schema.safeParse(value).success).toBe(true);
      }
      for (const value of ["", [], [""], ["id", 123], Array(51).fill("id")]) {
        expect(schema.safeParse(value).success).toBe(false);
      }
    });
  });

  it.each([
    { messageId: ["root"], replyId: "reply" },
    { messageId: ["root"], replyId: ["reply"] },
    { messageId: ["root"], listReplies: true },
    { messageId: "root", replyId: ["reply"], listReplies: true },
    { replyId: ["reply"] },
  ])("rejects ambiguous channel batches %j", async (args) => {
    const result = await call("get_channel_messages", {
      teamId: "team",
      channelId: "channel",
      ...args,
    });
    expect(result.isError).toBe(true);
    expect(api).not.toHaveBeenCalled();
  });

  it("lists all member pages and distinguishes membership IDs from user IDs", async () => {
    const members = [
      {
        id: "membership-a",
        userId: "user-a",
        email: "a@example.com",
        tenantId: TENANT_B,
        roles: ["owner"],
        visibleHistoryStartDateTime: "2026-01-01T00:00:00Z",
      },
      { id: "membership-b", userId: "guest-b", tenantId: "external-tenant", roles: ["guest"] },
    ];
    get
      .mockResolvedValueOnce({
        value: [members[0]],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat/members?$skiptoken=next",
      })
      .mockResolvedValueOnce({ value: [members[1]] });
    const result = await call("list_chat_members", { chatId: "chat" });
    expect(JSON.parse(result.content[0].text)).toEqual(members);
    expect(api.mock.calls.map(([path]) => path)).toEqual([
      "/chats/chat/members",
      "https://graph.microsoft.com/v1.0/chats/chat/members?$skiptoken=next",
    ]);
  });

  it.each([true, false])(
    "marks isRead=%s for the authenticated user in the resolved tenant",
    async (isRead) => {
      const result = await call("set_chat_read_state", { chatId: "chat", isRead });
      expect(result.isError).toBeUndefined();
      expect(api).toHaveBeenCalledExactlyOnceWith(
        `/chats/chat/${isRead ? "markChatReadForUser" : "markChatUnreadForUser"}`
      );
      expect(post).toHaveBeenCalledExactlyOnceWith({
        user: { id: "current-user", tenantId: TENANT_B },
      });
    }
  );

  it("uses the resolved tenant even when the caller omits tenantId", async () => {
    await server.getTool("set_chat_read_state").handler({ chatId: "chat", isRead: true });
    expect(forTenant).toHaveBeenCalledExactlyOnceWith(undefined);
    expect(post).toHaveBeenCalledExactlyOnceWith({
      user: { id: "current-user", tenantId: TENANT_B },
    });
  });

  it("supports an unread boundary and rejects it when marking read", async () => {
    const lastMessageReadDateTime = "2026-09-22T10:00:00+08:00";
    await call("set_chat_read_state", { chatId: "chat", isRead: false, lastMessageReadDateTime });
    expect(post).toHaveBeenCalledExactlyOnceWith({
      user: { id: "current-user", tenantId: TENANT_B },
      lastMessageReadDateTime,
    });
    post.mockClear();
    const result = await call("set_chat_read_state", {
      chatId: "chat",
      isRead: true,
      lastMessageReadDateTime,
    });
    expect(result.isError).toBe(true);
    expect(post).not.toHaveBeenCalled();
    const schema = server.getTool("set_chat_read_state").schema.lastMessageReadDateTime;
    expect(schema.safeParse("yesterday").success).toBe(false);
    expect(schema.safeParse(lastMessageReadDateTime).success).toBe(true);
  });

  it("never writes without a verified current user", async () => {
    bound.getAuthStatus.mockResolvedValueOnce({ ...status, isAuthenticated: false });
    const result = await call("set_chat_read_state", { chatId: "chat", isRead: true });
    expect(result.isError).toBe(true);
    expect(post).not.toHaveBeenCalled();
  });

  it.each([
    ["get_chat_messages", { chatId: "chat", messageId: "message" }],
    ["get_channel_messages", { teamId: "team", channelId: "channel", messageId: "message" }],
    ["list_chat_members", { chatId: "chat" }],
    ["set_chat_read_state", { chatId: "chat", isRead: false }],
  ])("surfaces Graph errors for %s", async (name, args) => {
    get.mockRejectedValue(new Error("Graph denied access"));
    post.mockRejectedValue(new Error("Graph denied access"));
    const result = await call(name, args);
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("Graph denied access");
  });

  it("keeps reads available but omits state changes in read-only mode", () => {
    const readonly = createMockMcpServer();
    registerChatTools(readonly as any, root, true);
    registerTeamsTools(readonly as any, root, true);
    expect(readonly.getAllTools()).not.toContain("set_chat_read_state");
    for (const name of ["list_chat_members", "get_chat_messages", "get_channel_messages"]) {
      expect(readonly.getTool(name).annotations.readOnlyHint).toBe(true);
    }
  });
});
