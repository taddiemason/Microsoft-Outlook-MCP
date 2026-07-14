import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import {
  PersistenceCreator,
  PersistenceCachePlugin,
  DataProtectionScope,
} from "@azure/msal-node-extensions";
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
 * Preferred cache: OS-native encrypted storage via @azure/msal-node-extensions.
 *   - Windows -> DPAPI (encrypted, tied to the current Windows user)
 *   - macOS   -> Keychain
 *   - Linux   -> libsecret / GNOME Keyring
 * The persistence is verified at startup; if the native backend can't be
 * loaded or validated we fall back to the plaintext file so sign-in still
 * works (just without encryption at rest).
 */
async function buildCachePlugin(config: Config): Promise<ICachePlugin> {
  try {
    const persistence = await PersistenceCreator.createPersistence({
      cachePath: config.tokenCachePath,
      // Used by the Windows DPAPI backend.
      dataProtectionScope: DataProtectionScope.CurrentUser,
      // Used by the macOS Keychain / Linux libsecret backends.
      serviceName: "microsoft-outlook-mcp",
      accountName: "token-cache",
      // Do NOT silently store plaintext on Linux when libsecret is missing;
      // let it throw so we hit the explicit fallback below (which logs).
      usePlaintextFileOnLinux: false,
    });

    // Confirms the backend can actually read/write/encrypt before we rely on
    // it. Throws or returns false when the native layer is unusable.
    const okToUse = await persistence.verifyPersistence();
    if (!okToUse) throw new Error("verifyPersistence() returned false");

    log("token cache: OS-native encrypted storage (msal-node-extensions)");
    return new PersistenceCachePlugin(persistence);
  } catch (err) {
    log(
      `secure token storage unavailable (${(err as Error).message}); ` +
        "falling back to a plaintext file with restricted permissions.",
    );
    return makePlaintextCachePlugin(config.tokenCachePath);
  }
}

/**
 * Fallback cache: a plaintext JSON file. Written with 0600 where the OS honors
 * POSIX modes. Treat this file as a live credential.
 */
function makePlaintextCachePlugin(cachePath: string): ICachePlugin {
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

  private constructor(pca: PublicClientApplication, scopes: string[]) {
    this.pca = pca;
    this.scopes = scopes;
  }

  /**
   * Async factory — building the encrypted persistence backend is async, so
   * construction must be awaited. Use this instead of `new AuthProvider(...)`.
   */
  static async create(config: Config): Promise<AuthProvider> {
    const cachePlugin = await buildCachePlugin(config);

    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
      cache: { cachePlugin },
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

    return new AuthProvider(new PublicClientApplication(msalConfig), config.scopes);
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
