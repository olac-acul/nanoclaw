---
name: add-outlook-readonly
description: Add read-only Microsoft 365 Outlook mail access to selected NanoClaw agent groups using delegated Mail.Read and device-code OAuth.
---

# Add Outlook Mail (read only)

Install a dependency-free stdio MCP server for selected agent groups. It uses
Microsoft Graph delegated OAuth with a fixed `Mail.Read` scope. It cannot send,
delete, move, flag, draft, reply, forward, or mark messages as read.

The tools are:

- `outlook_auth_status`, `outlook_auth_start`, `outlook_auth_complete`
- `outlook_list_messages`, `outlook_search_messages`, `outlook_get_message`
- `outlook_list_attachments`, `outlook_download_attachment`

Mail bodies are requested as plain text and capped at 100,000 characters.
Downloads default to 10 MiB maximum; executable and script attachments are
blocked. Every result labels email content as untrusted data.

## Phase 1: Microsoft Entra application

Create a single-tenant app registration with:

- **Supported account types:** Accounts in this organizational directory only
- **Authentication → Allow public client flows:** Yes
- **Microsoft Graph delegated permissions:** `Mail.Read` only
- **Client secret:** none

Remove the default `User.Read` permission if present. Do not add application
permissions, `Mail.Send`, or `Mail.ReadWrite`. Record the Application (client)
ID and Directory (tenant) ID locally. Both are identifiers, not credentials,
but do not commit installation-specific values.

## Phase 2: Select and prepare a group

List groups and choose the one that may read the mailbox:

```bash
ncl groups list
```

Resolve its folder with `ncl groups get --id <group-id>`. Define shell
variables `GROUP_ID`, `GROUP_FOLDER`, `OUTLOOK_CLIENT_ID`, and
`OUTLOOK_TENANT_ID` in the current terminal. Do not paste their values into a
tracked file.

Copy the two runtime files into the group's immutable plugin tree:

```bash
install -d -m 0755 "groups/${GROUP_FOLDER}/plugins/outlook-readonly"
install -m 0644 "${CLAUDE_SKILL_DIR}/outlook-readonly-core.ts" \
  "groups/${GROUP_FOLDER}/plugins/outlook-readonly/outlook-readonly-core.ts"
install -m 0644 "${CLAUDE_SKILL_DIR}/outlook-readonly-mcp.ts" \
  "groups/${GROUP_FOLDER}/plugins/outlook-readonly/outlook-readonly-mcp.ts"
install -d -m 0700 "groups/${GROUP_FOLDER}/plugin-data/outlook-readonly"
```

The runtime is standalone and uses Node/Bun built-ins only. It adds no npm
package and does not change the shared agent image.

## Phase 3: Register the MCP server

Register it for the selected group. Client and tenant IDs are stored in local
NanoClaw runtime state, not Git. The OAuth token cache is a `0600` file under
the selected group's `plugin-data/` directory.

```bash
ncl groups config add-mcp-server \
  --id "$GROUP_ID" \
  --name outlook_readonly \
  --command bun \
  --args '["run","/workspace/agent/plugins/outlook-readonly/outlook-readonly-mcp.ts"]' \
  --env "{\"OUTLOOK_CLIENT_ID\":\"$OUTLOOK_CLIENT_ID\",\"OUTLOOK_TENANT_ID\":\"$OUTLOOK_TENANT_ID\",\"OUTLOOK_TOKEN_CACHE\":\"/workspace/agent/plugin-data/outlook-readonly/token-cache.json\",\"OUTLOOK_DOWNLOAD_DIR\":\"/workspace/agent/downloads/outlook\",\"OUTLOOK_MAX_ATTACHMENT_BYTES\":\"10485760\"}"

ncl groups restart \
  --id "$GROUP_ID" \
  --message "Outlook read-only tools are installed. Check outlook_auth_status and, if needed, start device-code login."
```

This registration is runtime state and has no in-tree integration line. The
skill therefore needs no structural registration test. Its behavior and
read-only boundaries are tested directly from the self-contained source.

## Phase 4: Authenticate once

Ask the selected NanoClaw agent to connect Outlook in read-only mode. It calls
`outlook_auth_start` and returns Microsoft's verification URL and a short
code. Open the URL on a trusted device, enter the code, sign in with the
corporate mailbox, and review that the consent screen requests mail reading
only. Then tell the agent to complete the Outlook login.

The cache contains delegated access and refresh tokens. It is intentionally
outside Git and mode `0600`, but the selected agent container can use it. The
token itself is limited by Entra to `Mail.Read`; revoke it at any time from the
Microsoft account's app-consent page or by following `REMOVE.md`.

## Verify

Run the local safety tests before deployment:

```bash
pnpm exec vitest run --config .claude/skills/add-outlook-readonly/vitest.config.ts
```

Then ask the agent to list two unread Inbox messages, read one body, list its
attachments, and download one small non-executable attachment. Confirm that
the message remains unread in Outlook.

## Upgrade order

Use `/update-nanoclaw`, not a raw pull. Reapply this skill after an update by
copying its two runtime files over the prior plugin copies, rerun the test, and
restart the selected group. No upstream source file is modified.

## Troubleshooting

- `AADSTS65001` or consent failure: confirm delegated `Mail.Read` exists and
  public client flows are enabled. A tenant policy may still require an admin.
- `unauthorized_client`: confirm the tenant and client UUIDs belong to the same
  single-tenant app registration.
- Tool absent: inspect `ncl groups config get --id <group-id>`, then restart the
  group.
- Token-cache error: remove only the cache files documented in `REMOVE.md`,
  then authenticate again.
- Attachment rejected: raise the cap only up to 25 MiB, or inspect the file
  manually outside the agent. Executable/script extensions remain blocked.

To uninstall or revoke local access, follow [REMOVE.md](REMOVE.md).
