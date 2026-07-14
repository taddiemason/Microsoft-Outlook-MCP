import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { existsSync, readFileSync } from "node:fs";

/**
 * Minimal .env loader (no dependency). Reads KEY=VALUE lines from a .env file
 * in the project root and populates process.env for any keys not already set.
 */
function loadDotEnv(): void {
  const here = dirname(fileURLToPath(import.meta.url));
  // dist/config.js -> project root is one level up from dist
  const root = join(here, "..");
  const envPath = join(root, ".env");
  if (!existsSync(envPath)) return;

  const text = readFileSync(envPath, "utf8");
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq === -1) continue;
    const key = line.slice(0, eq).trim();
    let value = line.slice(eq + 1).trim();
    // Strip surrounding quotes if present.
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (process.env[key] === undefined) process.env[key] = value;
  }
}

loadDotEnv();

function projectRoot(): string {
  const here = dirname(fileURLToPath(import.meta.url));
  return join(here, "..");
}

export interface Config {
  clientId: string;
  tenantId: string;
  scopes: string[];
  tokenCachePath: string;
}

export function loadConfig(): Config {
  const clientId = process.env.OUTLOOK_CLIENT_ID?.trim();
  if (!clientId) {
    throw new Error(
      "OUTLOOK_CLIENT_ID is not set. Copy .env.example to .env and set your " +
        "Azure AD application (client) ID.",
    );
  }

  const tenantId = process.env.OUTLOOK_TENANT_ID?.trim() || "common";

  const scopes = (
    process.env.OUTLOOK_SCOPES?.trim() ||
    "User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Contacts.ReadWrite"
  )
    .split(/\s+/)
    .filter(Boolean);

  const tokenCachePath =
    process.env.OUTLOOK_TOKEN_CACHE_PATH?.trim() ||
    join(projectRoot(), ".token-cache.json");

  return { clientId, tenantId, scopes, tokenCachePath };
}
