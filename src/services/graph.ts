import { type AuthenticationResult, PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { z } from "zod";
import { createCachePlugin } from "../msal-cache.js";
import {
  authorityFor,
  CLIENT_ID,
  type TenantProfile,
  TenantStore,
  tenantIdSchema,
} from "../tenants.js";

/** Scopes sufficient for read-only operations (no message sending, no file uploads). */
export const READ_ONLY_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "TeamMember.Read.All",
  "Chat.Read",
];

/** Full scopes including write operations. */
export const FULL_SCOPES = [
  ...READ_ONLY_SCOPES,
  "ChannelMessage.Send",
  "ChannelMessage.ReadWrite",
  "Chat.ReadWrite",
  "Files.ReadWrite.All",
];

export interface AuthStatus {
  tenantId: string;
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
  error?: string;
}

export interface GraphOptions {
  store?: TenantStore;
  defaultTenantId?: string | undefined;
  readOnly?: boolean;
  accessToken?: string | undefined;
}

const accessTokenClaimsSchema = z.object({
  aud: z.union([z.string(), z.array(z.string())]),
  tid: tenantIdSchema,
  exp: z.number(),
  preferred_username: z.string().optional(),
  upn: z.string().optional(),
  scp: z.string().optional(),
});

/** A root resolves tenants; bound instances never change tenant or account during a call. */
export class GraphService {
  readonly store: TenantStore;
  readonly readOnlyMode: boolean;
  readonly defaultTenantId: string | undefined;
  private readonly sessions = new Map<string, GraphService>();
  private profile: TenantProfile | undefined;
  private clientPromise: Promise<Client> | undefined;
  private tokenExpiresAt: Date | undefined;
  private injected: { profile: TenantProfile; token: string; expiresAt: Date } | undefined;

  constructor(options: GraphOptions = {}) {
    this.store = options.store ?? new TenantStore();
    this.readOnlyMode = options.readOnly ?? false;
    this.defaultTenantId =
      options.defaultTenantId === undefined
        ? undefined
        : tenantIdSchema.parse(options.defaultTenantId);
    if (options.accessToken) this.injected = this.parseAccessToken(options.accessToken);
  }

  get scopes(): string[] {
    return this.readOnlyMode ? READ_ONLY_SCOPES : FULL_SCOPES;
  }

  async listTenants() {
    const profiles = await this.store.list();
    if (this.injected) {
      const index = profiles.findIndex((p) => p.tenantId === this.injected?.profile.tenantId);
      if (index >= 0) profiles.splice(index, 1);
      profiles.push(this.injected.profile);
    }
    return profiles.map((profile) => ({
      tenantId: profile.tenantId,
      name: profile.name,
      username: profile.username,
      scopes: profile.scopes,
      authenticatedAt: profile.authenticatedAt,
      isDefault: profile.tenantId === this.defaultTenantId,
      source: profile === this.injected?.profile ? "AUTH_TOKEN" : "device-code",
    }));
  }

  async forTenant(tenantId?: string): Promise<GraphService> {
    let id = tenantId ?? this.profile?.tenantId ?? this.defaultTenantId;
    if (id === undefined) {
      const tenants = await this.listTenants();
      if (tenants.length === 0)
        throw new Error("No tenants connected. Run: teams-mcp authenticate --tenant <tenant-id>");
      if (tenants.length !== 1)
        throw new Error(
          "Multiple tenants connected. Specify tenantId from list_tenants, or configure --tenant / TEAMS_MCP_TENANT_ID."
        );
      id = tenants[0].tenantId;
    }
    id = tenantIdSchema.parse(id);
    const profile =
      id === this.injected?.profile.tenantId ? this.injected.profile : await this.store.read(id);
    const existing = this.sessions.get(id);
    if (existing?.profile?.revision === profile.revision) return existing;
    const service = new GraphService({ store: this.store, readOnly: this.readOnlyMode });
    service.profile = profile;
    service.injected = this.injected;
    this.sessions.set(id, service);
    return service;
  }

  private async assertCurrent(): Promise<void> {
    if (!this.profile) throw new Error("Tenant is not selected");
    if (this.profile === this.injected?.profile) {
      if (this.injected.expiresAt.getTime() <= Date.now())
        throw new Error("AUTH_TOKEN has expired; supply a new token and restart the server.");
    } else {
      await this.store.assertCurrent(this.profile);
    }
  }

  private validateResult(result: AuthenticationResult | null): AuthenticationResult {
    if (
      !result ||
      result.tenantId.toLowerCase() !== this.profile?.tenantId ||
      result.account?.homeAccountId !== this.profile.homeAccountId ||
      result.account?.localAccountId !== this.profile.localAccountId
    ) {
      throw new Error(
        "Token tenant/account does not match the selected tenant. Re-authenticate this tenant."
      );
    }
    this.tokenExpiresAt = result.expiresOn ?? undefined;
    return result;
  }

  private async initializeClient(): Promise<Client> {
    const profile = this.profile;
    if (!profile) throw new Error("Tenant is not selected");
    if (profile === this.injected?.profile) {
      const injected = this.injected;
      this.tokenExpiresAt = injected.expiresAt;
      return Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            await this.assertCurrent();
            return injected.token;
          },
        },
      });
    }
    const authority = authorityFor(profile.tenantId);
    const app = new PublicClientApplication({
      auth: { clientId: profile.clientId, authority },
      cache: { cachePlugin: createCachePlugin(this.store, profile) },
    });
    const accounts = await app.getTokenCache().getAllAccounts();
    const account = accounts.find(
      (candidate) =>
        candidate.homeAccountId === profile.homeAccountId &&
        candidate.localAccountId === profile.localAccountId &&
        candidate.tenantId.toLowerCase() === profile.tenantId
    );
    if (!account)
      throw new Error(
        `Saved account not found for tenant ${profile.tenantId}. Re-authenticate this tenant.`
      );
    // Respect the permission mode used at login, even in a server exposing write tools.
    const scopes = this.scopes.filter((scope) =>
      profile.scopes.some((granted) => granted.toLowerCase() === scope.toLowerCase())
    );
    if (scopes.length === 0)
      throw new Error("No usable Graph scopes; re-authenticate this tenant.");
    const acquireToken = async () => {
      await this.assertCurrent();
      const result = this.validateResult(
        await app.acquireTokenSilent({ scopes, account, authority })
      );
      await this.assertCurrent();
      return result.accessToken;
    };
    await acquireToken();
    return Client.initWithMiddleware({ authProvider: { getAccessToken: acquireToken } });
  }

  async getClient(): Promise<Client> {
    if (!this.profile) return (await this.forTenant()).getClient();
    await this.assertCurrent();
    if (!this.clientPromise) {
      this.clientPromise = this.initializeClient().catch((error) => {
        this.clientPromise = undefined;
        throw error;
      });
    }
    return this.clientPromise;
  }

  async getAuthStatus(): Promise<AuthStatus> {
    if (!this.profile) return (await this.forTenant()).getAuthStatus();
    try {
      const client = await this.getClient();
      const me = await client.api("/me").get();
      return {
        tenantId: this.profile.tenantId,
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName,
        displayName: me?.displayName,
        expiresAt: this.tokenExpiresAt?.toISOString(),
      };
    } catch (error) {
      return {
        tenantId: this.profile.tenantId,
        isAuthenticated: false,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  private parseAccessToken(token: string) {
    let claims: z.infer<typeof accessTokenClaimsSchema>;
    try {
      if (token.split(".").length !== 3) throw new Error("Invalid JWT");
      claims = accessTokenClaimsSchema.parse(
        JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8"))
      );
    } catch {
      throw new Error("AUTH_TOKEN must be a Microsoft Graph JWT with tid and exp claims");
    }
    const audiences = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
    if (
      !audiences.some((aud: string) =>
        ["https://graph.microsoft.com", "00000003-0000-0000-c000-000000000000"].includes(aud)
      )
    ) {
      throw new Error("AUTH_TOKEN audience must be Microsoft Graph");
    }
    const tenantId = tenantIdSchema.parse(claims.tid);
    if (
      typeof claims.exp !== "number" ||
      !Number.isFinite(claims.exp) ||
      claims.exp * 1000 <= Date.now()
    )
      throw new Error("AUTH_TOKEN is expired or has no valid exp claim");
    const profile: TenantProfile = {
      tenantId,
      name: "Injected access token",
      clientId: CLIENT_ID,
      homeAccountId: "injected",
      localAccountId: "injected",
      username: claims.preferred_username ?? claims.upn ?? "",
      scopes: typeof claims.scp === "string" ? claims.scp.split(" ") : [],
      authenticatedAt: "",
      revision: "injected",
    };
    return { profile, token, expiresAt: new Date(claims.exp * 1000) };
  }
}
