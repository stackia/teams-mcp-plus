import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it, vi } from "vitest";
import { createServer } from "../cli.js";
import type { GraphService } from "../services/graph.js";
import { createMockMcpServer } from "../test-utils/setup.js";
import { TENANT_A, TENANT_B } from "../test-utils/tenants.js";
import { registerAuthTools } from "../tools/auth.js";
import { registerChatTools } from "../tools/chats.js";
import { registerSearchTools } from "../tools/search.js";
import { registerTeamsTools } from "../tools/teams.js";
import { registerUsersTools } from "../tools/users.js";
import { uploadFileToChat } from "../utils/file-upload.js";

vi.mock("../utils/file-upload.js", async (original) => ({
  ...(await original<typeof import("../utils/file-upload.js")>()),
  uploadFileToChat: vi.fn().mockResolvedValue({
    attachmentId: "item",
    webUrl: "https://example.com/file",
    fileName: "file.txt",
    fileSize: 5,
    mimeType: "text/plain",
  }),
}));

function services() {
  const clients = new Map<string, any>();
  const bound = new Map<string, any>();
  for (const id of [TENANT_A, TENANT_B]) {
    const request = {
      get: vi.fn().mockResolvedValue({ id, displayName: id, value: [] }),
      post: vi.fn().mockResolvedValue({ id: `sent-${id}` }),
    };
    const client = { api: vi.fn().mockReturnValue(request) };
    clients.set(id, client);
    bound.set(id, { getClient: vi.fn().mockResolvedValue(client) });
  }
  const root = {
    readOnlyMode: false,
    listTenants: vi.fn().mockResolvedValue([{ tenantId: TENANT_A }, { tenantId: TENANT_B }]),
    forTenant: vi.fn().mockImplementation(async (id) => {
      if (!bound.has(id)) throw new Error("Select a tenant with tenantId");
      return bound.get(id);
    }),
    getClient: vi.fn().mockRejectedValue(new Error("Unbound client must never be used")),
  };
  return { root: root as unknown as GraphService, clients, bound };
}

describe("MCP tenant routing", () => {
  it("exposes tenant selection over the real MCP transport and isolates concurrent calls", async () => {
    const { root, clients } = services();
    const server = createServer(root);
    const client = new Client({ name: "test", version: "1" });
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    await client.connect(clientTransport);
    try {
      const { tools } = await client.listTools();
      expect(tools).toHaveLength(31);
      for (const tool of tools.filter((t) => t.name !== "list_tenants")) {
        expect(tool.inputSchema.properties).toHaveProperty("tenantId");
        expect(tool.inputSchema.required ?? []).not.toContain("tenantId");
      }
      const [a, b] = await Promise.all([
        client.callTool({ name: "get_current_user", arguments: { tenantId: TENANT_A } }),
        client.callTool({ name: "get_current_user", arguments: { tenantId: TENANT_B } }),
      ]);
      expect(JSON.parse((a.content as any)[0].text).id).toBe(TENANT_A);
      expect(JSON.parse((b.content as any)[0].text).id).toBe(TENANT_B);
      const ambiguous = await client.callTool({ name: "list_teams", arguments: {} });
      expect(ambiguous.isError).toBe(true);
      const invalid = await client.callTool({
        name: "list_teams",
        arguments: { tenantId: "common" },
      });
      expect(invalid.isError).toBe(true);
      await client.callTool({
        name: "send_chat_message",
        arguments: { tenantId: TENANT_B, chatId: "chat-b", message: "test" },
      });
      expect(clients.get(TENANT_B).api).toHaveBeenCalledWith("/me/chats/chat-b/messages");
      expect(clients.get(TENANT_A).api).not.toHaveBeenCalledWith("/me/chats/chat-b/messages");
      expect(root.getClient).not.toHaveBeenCalled();
    } finally {
      await client.close();
      await server.close();
    }
  });

  it("every tenant-scoped handler resolves the requested tenant before doing work", async () => {
    const server = createMockMcpServer();
    const { root } = services();
    vi.mocked(root.forTenant).mockRejectedValue(new Error("tenant-selection-failed"));
    for (const register of [
      registerAuthTools,
      registerUsersTools,
      registerTeamsTools,
      registerChatTools,
      registerSearchTools,
    ]) {
      register(server as any, root, false);
    }
    for (const name of server.getAllTools().filter((name) => name !== "list_tenants")) {
      vi.mocked(root.forTenant).mockClear();
      const result = await server.getTool(name).handler({ tenantId: TENANT_B });
      expect(root.forTenant, name).toHaveBeenCalledExactlyOnceWith(TENANT_B);
      expect(JSON.stringify(result), name).toContain("tenant-selection-failed");
      expect(result.isError, name).toBe(true);
    }
    expect(root.getClient).not.toHaveBeenCalled();
  });

  it("passes the selected tenant into nested file uploads", async () => {
    const server = createMockMcpServer();
    const { root, bound, clients } = services();
    vi.mocked(uploadFileToChat).mockResolvedValueOnce({
      attachmentId: "item",
      webUrl: "https://example.com/file",
      fileName: "file.txt",
      fileSize: 5,
      mimeType: "text/plain",
    });
    registerChatTools(server as any, root, false);
    await server
      .getTool("send_file_to_chat")
      .handler({ tenantId: TENANT_B, chatId: "chat", filePath: "/tmp/file.txt" });
    expect(uploadFileToChat).toHaveBeenCalledWith(bound.get(TENANT_B), "/tmp/file.txt", undefined);
    expect(clients.get(TENANT_B).api).toHaveBeenCalledWith("/me/chats/chat/messages");
    expect(bound.get(TENANT_A).getClient).not.toHaveBeenCalled();
  });

  it("keeps write tools disabled in read-only servers while exposing all tenant selectors", async () => {
    const { root } = services();
    Object.defineProperty(root, "readOnlyMode", { value: true });
    const server = createServer(root);
    const client = new Client({ name: "test", version: "1" });
    const [st, ct] = InMemoryTransport.createLinkedPair();
    await server.connect(st);
    await client.connect(ct);
    try {
      const { tools } = await client.listTools();
      expect(tools.every((t) => t.annotations?.readOnlyHint)).toBe(true);
      expect(tools.some((t) => t.name === "list_tenants")).toBe(true);
      expect(tools.some((t) => t.name === "send_chat_message")).toBe(false);
    } finally {
      await client.close();
      await server.close();
    }
  });
});
