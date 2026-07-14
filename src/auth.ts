import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import { existsSync, readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname } from "node:path";
import type { Config } from "./config.js";

/**
 * IMPORTANT: everything here logs to stderr. In an MCP stdio server, stdout is
 * reserved exclusively for the JSON-RPC protocol stream.
 */
function log(msg: string): void {
  process.stderr.write(`[auth] ${msg}\n`);
}

/**
 * Persists the MSAL token cache to a JSON file on disk so refresh tokens
 * survive process restarts. The file contains sensitive tokens — it is
 * git-ignored and should be treated like a credential.
 */
function makeCachePlugin(cachePath: string): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (existsSync(cachePath)) {
        ctx.tokenCache.deserialize(readFileSync(cachePath, "utf8"));
      }
    },
    async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (ctx.cacheHasChanged) {
        mkdirSync(dirname(cachePath), { recursive: true });
        writeFileSync(cachePath, ctx.tokenCache.serialize(), { mode: 0o600 });
      }
    },
  };
}

export class AuthProvider {
  private readonly pca: PublicClientApplication;
  private readonly scopes: string[];

  constructor(config: Config) {
    this.scopes = config.scopes;

    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
      cache: {
        cachePlugin: makeCachePlugin(config.tokenCachePath),
      },
      system: {
        loggerOptions: {
          loggerCallback: (level, message) => {
            if (level <= LogLevel.Warning) log(message);
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Warning,
        },
      },
    };

    this.pca = new PublicClientApplication(msalConfig);
  }

  /**
   * Returns a valid access token, refreshing silently from the cache when
   * possible. Only falls back to interactive device-code sign-in when there is
   * no usable cached account. When `interactive` is false (the default for
   * request-time acquisition) a missing/expired session throws instead of
   * blocking on user input — that keeps tool calls from hanging.
   */
  async getAccessToken(interactive = false): Promise<string> {
    const cache = this.pca.getTokenCache();
    const accounts = await cache.getAllAccounts();

    if (accounts.length > 0) {
      try {
        const result = await this.pca.acquireTokenSilent({
          account: accounts[0],
          scopes: this.scopes,
        });
        if (result?.accessToken) return result.accessToken;
      } catch (err) {
        log(`silent token acquisition failed: ${(err as Error).message}`);
      }
    }

    if (!interactive) {
      throw new Error(
        "Not signed in (or the session expired). Run the login step first: " +
          "`npm run login`.",
      );
    }

    return this.deviceCodeLogin();
  }

  /** Interactive device-code sign-in. Prompts on stderr. */
  async deviceCodeLogin(): Promise<string> {
    const result = await this.pca.acquireTokenByDeviceCode({
      scopes: this.scopes,
      deviceCodeCallback: (info) => {
        // info.message already contains the URL + code + instructions.
        process.stderr.write(`\n${info.message}\n\n`);
      },
    });
    if (!result?.accessToken) {
      throw new Error("Device code sign-in returned no access token.");
    }
    log(`signed in as ${result.account?.username ?? "unknown account"}`);
    return result.accessToken;
  }

  async isSignedIn(): Promise<boolean> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    return accounts.length > 0;
  }

  async signOut(): Promise<void> {
    const cache = this.pca.getTokenCache();
    for (const account of await cache.getAllAccounts()) {
      await cache.removeAccount(account);
    }
    log("signed out; cached accounts cleared.");
  }
}
