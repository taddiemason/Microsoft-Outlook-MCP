#!/usr/bin/env node
/**
 * Standalone interactive sign-in. Run `npm run login` in a terminal.
 * This performs the device-code flow and writes the token cache to disk so the
 * MCP server (which runs non-interactively over stdio) can use it silently.
 *
 * Also supports `npm run login -- --logout` to clear the cached session.
 */
import { loadConfig } from "./config.js";
import { AuthProvider } from "./auth.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const auth = await AuthProvider.create(config);

  if (process.argv.includes("--logout")) {
    await auth.signOut();
    return;
  }

  if (process.argv.includes("--status")) {
    const signedIn = await auth.isSignedIn();
    process.stderr.write(signedIn ? "Signed in.\n" : "Not signed in.\n");
    return;
  }

  if (await auth.isSignedIn()) {
    process.stderr.write(
      "Already signed in. Use `--logout` to clear, or continue.\n",
    );
    // Refresh silently to confirm the token still works.
    await auth.getAccessToken(false).catch(() => auth.deviceCodeLogin());
    process.stderr.write("Session is valid.\n");
    return;
  }

  process.stderr.write("Starting device-code sign-in...\n");
  await auth.deviceCodeLogin();
  process.stderr.write("Sign-in complete. Token cache written.\n");
}

main().catch((err) => {
  process.stderr.write(`login failed: ${(err as Error).message}\n`);
  process.exit(1);
});
