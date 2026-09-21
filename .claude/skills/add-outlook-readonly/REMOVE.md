# Remove Outlook read-only access

This procedure is idempotent. Replace `<group-id>` and `<group-folder>` with
the selected group's values.

## 1. Remove the MCP registration

```bash
ncl groups config remove-mcp-server --id <group-id> --name outlook_readonly
ncl groups restart --id <group-id>
```

## 2. Remove the local runtime and delegated tokens

Delete only these Outlook integration paths:

```bash
rm -f "groups/<group-folder>/plugin-data/outlook-readonly/token-cache.json" \
      "groups/<group-folder>/plugin-data/outlook-readonly/token-cache.json.device"
rm -f "groups/<group-folder>/plugins/outlook-readonly/outlook-readonly-core.ts" \
      "groups/<group-folder>/plugins/outlook-readonly/outlook-readonly-mcp.ts"
rmdir "groups/<group-folder>/plugin-data/outlook-readonly" 2>/dev/null || true
rmdir "groups/<group-folder>/plugins/outlook-readonly" 2>/dev/null || true
```

Downloaded attachments under `groups/<group-folder>/downloads/outlook/` are
ordinary user data. Review and remove them separately if no longer needed.

## 3. Revoke Microsoft consent

In the Microsoft work account security/app-consent page, revoke the app named
for this NanoClaw integration. Deleting the local token cache stops this
installation; revoking consent also invalidates the delegated grant at
Microsoft.
