# Microsoft Outlook MCP

An [MCP](https://modelcontextprotocol.io) server that connects Claude (or any
MCP client) to Microsoft Outlook — **Mail, Calendar, and Contacts** — through
the [Microsoft Graph API](https://learn.microsoft.com/graph/) using OAuth 2.0
**device code flow**.

- Cloud-based via Microsoft Graph — works anywhere, for Microsoft 365 and
  personal Microsoft accounts.
- No client secret required. Device code flow uses a **public client** app.
- Tokens are cached on disk and refreshed silently, so you sign in once.

## Tools

| Area | Tools |
| --- | --- |
| Account | `whoami`, `auth_status` |
| Mail | `list_messages`, `get_message`, `send_mail`, `reply_to_message`, `update_message`, `move_message`, `delete_message` |
| Calendar | `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `respond_to_event` |
| Contacts | `list_contacts`, `get_contact`, `create_contact`, `update_contact`, `delete_contact` |
| Folders | `list_mail_folders`, `create_mail_folder`, `rename_mail_folder`, `delete_mail_folder` |
| Inbox rules | `list_message_rules`, `get_message_rule`, `create_message_rule`, `update_message_rule`, `delete_message_rule` |

## 1. Register an app in Azure AD (Entra ID)

1. Go to the [Azure Portal](https://portal.azure.com) →
   **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it e.g. `Outlook MCP`.
3. **Supported account types**: choose based on the accounts you'll use:
   - *Accounts in any organizational directory and personal Microsoft accounts*
     → tenant `common`
   - *This organizational directory only* → your specific tenant GUID
4. Leave **Redirect URI** blank (device code flow doesn't need one). Register.
5. On the app's **Overview** page, copy the **Application (client) ID** and
   **Directory (tenant) ID**.
6. **Authentication** → **Advanced settings** → set
   **Allow public client flows** to **Yes**. (Required for device code flow.)
7. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions**, and add:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`
   - `MailboxSettings.ReadWrite` (required for inbox rules)
   - `offline_access` (added automatically for refresh tokens)

   Click **Grant admin consent** if you're in an org tenant that requires it.

## 2. Configure

```powershell
cd C:\Users\zlalime\Documents\Github\Microsoft-Outlook-MCP
copy .env.example .env
```

Edit `.env` and set at least `OUTLOOK_CLIENT_ID` and `OUTLOOK_TENANT_ID`.

## 3. Build & sign in

```powershell
npm install
npm run build
npm run login
```

`npm run login` prints a URL and a code. Open
<https://microsoft.com/devicelogin>, enter the code, and approve. The session
is cached in `.token-cache.json` (git-ignored). Useful variants:

```powershell
npm run login -- --status   # check whether a session is cached
npm run login -- --logout   # clear the cached session
```

## 4. Connect a client

### Claude Code

```powershell
claude mcp add outlook --scope user -- node "C:\Users\zlalime\Documents\Github\Microsoft-Outlook-MCP\dist\index.js"
```

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:\\Users\\zlalime\\Documents\\Github\\Microsoft-Outlook-MCP\\dist\\index.js"]
    }
  }
}
```

The server reads `.env` from its own folder, so no `env` block is needed. If
you prefer, you can instead pass the values inline:

```json
"env": {
  "OUTLOOK_CLIENT_ID": "…",
  "OUTLOOK_TENANT_ID": "common"
}
```

Restart the client. Ask Claude to run `auth_status` or `whoami` to confirm.

## How auth works

- The MCP server runs **non-interactively** over stdio, so it never prompts.
  At request time it acquires tokens **silently** from the on-disk cache
  (`acquireTokenSilent`), refreshing with the stored refresh token as needed.
- Interactive **device-code sign-in** happens only in the separate
  `npm run login` step. Run it again any time a tool reports the session
  expired.
- The token cache is stored in **OS-native encrypted storage** via
  [`@azure/msal-node-extensions`](https://www.npmjs.com/package/@azure/msal-node-extensions):
  DPAPI on Windows, Keychain on macOS, libsecret on Linux. The backend is
  verified at startup; if the native layer can't be loaded, the server logs a
  warning and falls back to a restricted-permission plaintext file so sign-in
  still works.
- All diagnostics go to **stderr**; **stdout** carries only the MCP JSON-RPC
  stream.

## Environment variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `OUTLOOK_CLIENT_ID` | *(required)* | Azure AD application (client) ID |
| `OUTLOOK_TENANT_ID` | `common` | Tenant GUID or `common`/`organizations`/`consumers` |
| `OUTLOOK_SCOPES` | Mail/Calendar/Contacts set | Space-separated delegated Graph scopes |
| `OUTLOOK_TOKEN_CACHE_PATH` | `.token-cache.json` | Where the token cache is stored |

## Security notes

- The token cache holds live refresh/access tokens. By default it is
  **encrypted at rest** using your OS credential store (Windows DPAPI / macOS
  Keychain / Linux libsecret), tied to your user account. Only if that native
  backend is unavailable does it fall back to a git-ignored plaintext file
  written with `0600` permissions — watch the startup log to see which is in
  use. Either way, run `npm run login -- --logout` to revoke local access.
- **Client ID and tenant ID are not secrets** — this is a **public client**
  app, so no client secret is stored anywhere.
- Scopes are delegated: the server can only do what your signed-in account can.
- If you upgraded from an earlier version that used a plaintext
  `.token-cache.json`, delete that file and re-run `npm run login` so the cache
  is rewritten in the encrypted format.

## Troubleshooting

### Encrypted storage isn't being used (native module)

The encrypted cache is provided by `@azure/msal-node-extensions`, which relies
on a **native (N-API) addon**. `npm install` normally downloads a prebuilt
binary for your platform. If none is available for your OS/CPU/Node version,
the build falls back to compiling from source, which needs a toolchain.

You'll know the native layer failed to load if the startup log shows:

```
[auth] secure token storage unavailable (...); falling back to a plaintext file...
```

instead of:

```
[auth] token cache: OS-native encrypted storage (msal-node-extensions)
```

The server still works in this state — it just stores tokens in a
restricted-permission plaintext file. To get encryption working:

- **Windows** — install the
  [Visual Studio Build Tools](https://visualstudio.microsoft.com/downloads/)
  ("Desktop development with C++"), then `npm rebuild`.
- **macOS** — install the Xcode command line tools: `xcode-select --install`,
  then `npm rebuild`.
- **Linux** — install `build-essential`, `python3`, and libsecret headers
  (Debian/Ubuntu: `sudo apt-get install build-essential python3 libsecret-1-dev`);
  ensure a Secret Service provider (e.g. GNOME Keyring) is running, then
  `npm rebuild`.

After rebuilding, run `npm run login -- --status` and confirm the log reports
OS-native encrypted storage.

### `AADSTS7000218` or "public client flow" errors during login

Your app registration is missing the public-client setting. In Azure →
**Authentication** → **Advanced settings**, set **Allow public client flows**
to **Yes**.

### Tools return HTTP 401 / 403 after adding scopes

The cached token predates the new permission. Re-run `npm run login` (and grant
admin consent in Azure if your tenant requires it) so a fresh token carries the
added scope.

### `whoami` / tools say "Not signed in"

Run `npm run login` in a terminal from the project folder. The MCP server never
prompts for sign-in itself — it only reads the cached session.

## Project layout

```
src/
  index.ts          MCP server entry: registers tools, connects stdio transport
  login.ts          Standalone device-code sign-in (npm run login)
  config.ts         .env loader + config
  auth.ts           MSAL public client, device code + silent refresh, encrypted cache
  graph.ts          Minimal fetch-based Graph client (auth, paging, errors)
  tools/
    mail.ts         Mail tools
    calendar.ts     Calendar tools
    contacts.ts     Contacts tools
    rules.ts        Mail folder management + inbox rule tools
    util.ts         ok()/fail() result helpers
```

## License

MIT
