import { randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { z } from "zod";

export const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
export const tenantIdSchema = z
  .string()
  .regex(
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
    "Use the Microsoft Entra tenant ID (GUID), not common, organizations, or a domain"
  )
  .transform((id) => id.toLowerCase());
export const tenantInputSchema = {
  tenantId: tenantIdSchema
    .optional()
    .describe(
      "Target tenant ID from list_tenants. Omit only when a default is configured or exactly one tenant is connected. Resource IDs belong to this tenant."
    ),
};
export const profileSchema = z.object({
  tenantId: tenantIdSchema,
  name: z.string(),
  clientId: z.string(),
  homeAccountId: z.string().min(1),
  localAccountId: z.string().min(1),
  username: z.string(),
  scopes: z.array(z.string()),
  authenticatedAt: z.string(),
  revision: z.string().uuid(),
});
export type TenantProfile = z.infer<typeof profileSchema>;
export const authorityFor = (tenantId: string) =>
  `https://login.microsoftonline.com/${tenantIdSchema.parse(tenantId)}`;

/** Private, atomic files; callers create directories only during explicit login. */
export async function atomicWrite(path: string, data: string): Promise<void> {
  const temporary = `${path}.${randomUUID()}.tmp`;
  try {
    await fs.writeFile(temporary, data, { encoding: "utf8", mode: 0o600, flag: "wx" });
    await fs.rename(temporary, path);
  } finally {
    await fs.rm(temporary, { force: true });
  }
}

export class TenantStore {
  constructor(
    readonly directory = process.env.TEAMS_MCP_CONFIG_DIR || join(homedir(), ".teams-mcp-plus")
  ) {}

  private tenantDirectory(tenantId: string): string {
    return join(this.directory, tenantIdSchema.parse(tenantId));
  }

  cachePath(profile: TenantProfile): string {
    return join(
      this.tenantDirectory(profile.tenantId),
      `${profileSchema.parse(profile).revision}.cache.json`
    );
  }

  async read(tenantId: string): Promise<TenantProfile> {
    const id = tenantIdSchema.parse(tenantId);
    try {
      const profile = profileSchema.parse(
        JSON.parse(await fs.readFile(join(this.tenantDirectory(id), "profile.json"), "utf8"))
      );
      if (profile.tenantId !== id) throw new Error("Tenant profile ID mismatch");
      return profile;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") {
        throw new Error(
          `Tenant ${id} is not connected. Run: teams-mcp-plus authenticate --tenant ${id}`,
          { cause: error }
        );
      }
      throw error;
    }
  }

  async list(): Promise<TenantProfile[]> {
    let names: string[];
    try {
      names = await fs.readdir(this.directory);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
      throw error;
    }
    const profiles: TenantProfile[] = [];
    for (const name of names.sort()) {
      if (!tenantIdSchema.safeParse(name).success) continue;
      try {
        profiles.push(await this.read(name));
      } catch (error) {
        // A login may be publishing its first profile, or another process may have logged out.
        // Ignore only absent profiles; malformed credentials must never silently select another tenant.
        if (
          (error as Error).cause &&
          ((error as Error).cause as NodeJS.ErrnoException).code === "ENOENT"
        )
          continue;
        throw error;
      }
    }
    return profiles;
  }

  async save(profile: TenantProfile, cache: string): Promise<void> {
    profileSchema.parse(profile);
    await fs.mkdir(this.directory, { recursive: true, mode: 0o700 });
    const directory = this.tenantDirectory(profile.tenantId);
    await fs.mkdir(directory, { recursive: true, mode: 0o700 });
    // Each login gets an immutable cache identity. Old processes cannot overwrite a new login.
    await atomicWrite(this.cachePath(profile), cache);
    await atomicWrite(join(directory, "profile.json"), JSON.stringify(profile, null, 2));
  }

  async assertCurrent(profile: TenantProfile): Promise<void> {
    if ((await this.read(profile.tenantId)).revision !== profile.revision) {
      throw new Error(
        `Authentication changed for tenant ${profile.tenantId}; retry the tool call.`
      );
    }
  }

  async logout(tenantId: string): Promise<void> {
    await fs.rm(this.tenantDirectory(tenantId), { recursive: true, force: true });
  }
}
