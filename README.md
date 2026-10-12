# grok-bot-m365

A local (stdio) [Model Context Protocol](https://modelcontextprotocol.io) server for Microsoft 365 mail, using Microsoft Graph with delegated permissions.

It handles two things many built-in Outlook connectors don't:

- **One mailbox, several addresses.** It lists your SMTP aliases, filters the inbox by the alias a message was sent to, and drafts mail from any alias.
- **Shared mailboxes.** It reads a shared mailbox and sends from it (Send As) with your own sign-in.

Sending always takes two steps. `create_draft` or `reply_draft` returns a readback of the draft, and only `send_draft` sends it. `send_draft` must restate the draft's recipients and subject, so your client's approval prompt shows where the mail is going; the server checks them against the draft and refuses to send on any mismatch. The server deliberately has no tool that sends in one step.

## Tools

All mail tools accept an optional `mailbox` (a shared mailbox address, which must be listed in `M365_SHARED_MAILBOXES`). By default they use your own mailbox (`/me`).

| Tool | Purpose |
|---|---|
| `auth_start` / `auth_status` | Device-code sign-in: returns a URL and code, then checks whether sign-in finished |
| `list_aliases` | Primary address and all SMTP proxy addresses |
| `list_messages` | `folder`, `to_address`, `unread_only`, `query` (Graph `$search`), `top` |
| `get_message` | One message with a plain-text body |
| `create_draft` | `to`, `subject`, `body`, `cc`, `bcc`, `from_address`, `reply_to_message_id`, `html` |
| `reply_draft` | Reply or reply-all draft; never sends |
| `send_draft` | `draft_id`, `recipients` (every To/Cc/Bcc address), `subject`; sends only if they match the draft |
| `move_message`, `mark_read` | Basic triage |

`from_address` must be one of your own SMTP addresses or an allowed shared mailbox; anything else is refused before a draft is created.

`to_address` filtering runs on the client side (an exact, case-insensitive match on To and Cc, checking up to 500 recent messages). Graph can't filter on recipient collections, and `$search` would also match message bodies.

## Setup

### 1. Register an app in Microsoft Entra
1. Go to https://entra.microsoft.com, then **App registrations > New registration**. Choose single tenant and leave the redirect URI empty.
2. Under **Authentication > Advanced settings**, set **Allow public client flows** to **Yes**. Device-code flow needs this, and no client secret is used.
3. Under **API permissions > Microsoft Graph > Delegated**, add `Mail.ReadWrite`, `Mail.Send`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`, `User.Read` and `offline_access`, then **Grant admin consent**.
4. Copy the **Application (client) ID** and the **Directory (tenant) ID**.

### 2. Allow sending from aliases (Exchange Online)
```powershell
Connect-ExchangeOnline
Set-OrganizationConfig -SendFromAliasEnabled $true
Get-OrganizationConfig | Format-List SendFromAliasEnabled
```
The change can take from one hour up to a day to take effect. Every alias domain must be an accepted domain in the same tenant (check with `Get-AcceptedDomain`), and each alias must be a proxy address on the mailbox:
```powershell
Get-Mailbox user@contoso.com | Select -Expand EmailAddresses
Set-Mailbox user@contoso.com -EmailAddresses @{add="user@example.com"}
```

### 3. Grant access to a shared mailbox
```powershell
Add-MailboxPermission   -Identity shared@contoso.com -User user@contoso.com -AccessRights FullAccess -InheritanceType All -AutoMapping $false
Add-RecipientPermission -Identity shared@contoso.com -Trustee user@contoso.com -AccessRights SendAs -Confirm:$false
```
You can also do this in the Exchange admin center under **Recipients > Mailboxes > (mailbox) > Delegation**.

### 4. Install and configure your MCP client
```bash
git clone https://github.com/<owner>/grok-bot-m365 && cd grok-bot-m365 && uv sync
```
```json
{
  "mcpServers": {
    "m365-mail": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/grok-bot-m365", "python", "server.py"],
      "env": {
        "M365_CLIENT_ID": "<application-client-id>",
        "M365_TENANT_ID": "<directory-tenant-id>",
        "M365_ALIASES": "user@contoso.com,user@example.com",
        "M365_SHARED_MAILBOXES": "shared@contoso.com"
      }
    }
  }
}
```
Then call `auth_start`, open the URL it returns, enter the code, and call `auth_status`.

## Environment variables

| Variable | Required | Default | Meaning |
|---|---|---|---|
| `M365_CLIENT_ID` | yes | | Entra application (client) ID |
| `M365_TENANT_ID` | no | `organizations` | Tenant ID or domain |
| `M365_TOKEN_CACHE` | no | `~/.config/grok-bot-m365/token_cache.bin` | Token cache location (encrypted, or a signal file when the OS keyring holds the data) |
| `M365_ALLOW_PLAINTEXT_CACHE` | no | off | Set to `1` to fall back to a plaintext `0600` file when no OS keyring is available |
| `M365_ALIASES` | no | | Comma-separated list of expected aliases; `list_aliases` reports any that are missing from the mailbox |
| `M365_SHARED_MAILBOXES` | no | | Comma-separated allowlist of shared mailboxes. Any other `mailbox` is refused, so with this unset only your own mailbox is reachable |

## Errors
- **403 on a shared mailbox:** the user is missing Full Access, or admin consent hasn't been granted for the `.Shared` scopes.
- **Send rejected for a From address:** either `SendFromAliasEnabled` is off, the address isn't a proxy address on the mailbox, or Send As is missing on the shared mailbox.

## Security notes
- The server uses delegated permissions only. It can do nothing beyond what the signed-in user can already do in Exchange, and it has no client secret or application permissions.
- Within that, it narrows what a model can reach: only your own mailbox plus those in `M365_SHARED_MAILBOXES`; `from_address` limited to your own addresses and those mailboxes; IDs are encoded as single URL segments so they can't address other Graph paths; and the access token is only ever sent to `graph.microsoft.com`.
- The token cache holds refresh tokens and is stored in the OS secret store (DPAPI on Windows, Keychain on macOS, libsecret on Linux) via `msal-extensions`. A plaintext `0600` file is used only if you set `M365_ALLOW_PLAINTEXT_CACHE=1`. If you ran an earlier version, delete the old plaintext `~/.config/grok-bot-m365/token_cache.json` and sign in again.
- The server never logs or returns tokens.
- The server has no one-shot send tool. Keep `send_draft` behind a human approval step in your client; its arguments show the recipients and subject being approved, and a mismatch with the draft blocks the send. The body is not part of that check, so read the readback.
- Mail content returned by these tools is untrusted input to whatever model reads it. A message can carry instructions aimed at the model. Besides sending, the tools that act without a send are `move_message` and `mark_read`, which a misled model could use to hide mail; gate them in your client if that matters to you.
- Everything these tools return (bodies, addresses, shared-mailbox content) goes to the model provider behind your MCP client. Check that provider's data retention and your organization's data-handling policy before connecting work mail.
- Device-code sign-in is blocked by Conditional Access in many tenants because it is a common phishing route. If yours blocks it, ask for a scoped exception for this app rather than loosening the policy.

## Development
```bash
uv sync
uv run pytest -q                 # mocked Graph tests (respx)
uv run python tests/stdio_smoke.py   # starts the server over stdio and lists tools
```

## License
MIT. See [LICENSE](LICENSE).
