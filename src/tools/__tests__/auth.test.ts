import { beforeEach, describe, expect, it, vi } from "vitest";
import { createMockMcpServer } from "../../test-utils/setup.js";
import { TENANT_A } from "../../test-utils/tenants.js";
import { registerAuthTools } from "../auth.js";

let server: ReturnType<typeof createMockMcpServer>;
let service: any;
beforeEach(() => {
  server = createMockMcpServer();
  service = {
    forTenant: vi.fn().mockReturnThis(),
    listTenants: vi.fn().mockResolvedValue([{ tenantId: TENANT_A, name: "Work" }]),
    getAuthStatus: vi
      .fn()
      .mockResolvedValue({ tenantId: TENANT_A, isAuthenticated: true, displayName: "User" }),
  };
  registerAuthTools(server as any, service, false);
});

describe("tenant discovery and authentication tools", () => {
  it("lists tenants without requiring a default or refreshing tokens", async () => {
    const result = await server.getTool("list_tenants").handler();
    expect(JSON.parse(result.content[0].text)).toEqual([{ tenantId: TENANT_A, name: "Work" }]);
    expect(service.forTenant).not.toHaveBeenCalled();
    expect(service.getAuthStatus).not.toHaveBeenCalled();
  });

  it("includes the selected tenant in live authentication status", async () => {
    const result = await server.getTool("auth_status").handler({ tenantId: TENANT_A });
    expect(service.forTenant).toHaveBeenCalledWith(TENANT_A);
    expect(JSON.parse(result.content[0].text)).toMatchObject({
      tenantId: TENANT_A,
      isAuthenticated: true,
    });
    expect(result.isError).toBe(false);
  });

  it("reports ambiguous selection as an MCP error", async () => {
    service.forTenant.mockRejectedValueOnce(new Error("Multiple tenants connected"));
    const result = await server.getTool("auth_status").handler({});
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("Multiple tenants");
  });

  it("reports expired credentials as an MCP error", async () => {
    service.getAuthStatus.mockResolvedValueOnce({
      tenantId: TENANT_A,
      isAuthenticated: false,
      error: "Interaction required",
    });
    const result = await server.getTool("auth_status").handler({ tenantId: TENANT_A });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("Interaction required");
  });
});

it("reports credential discovery errors without claiming an empty tenant list", async () => {
  service.listTenants.mockRejectedValueOnce(new Error("Corrupt tenant profile"));
  const result = await server.getTool("list_tenants").handler();
  expect(result.isError).toBe(true);
  expect(result.content[0].text).toContain("Corrupt tenant profile");
});
