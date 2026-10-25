# Secrets: Azure Key Vault and Sliplane env

Follow `AGENTS.md` and `docs/tc-vault-migration.md`. Infisical was retired in
October 2026; do not restore its bootstrap variables or entrypoint.

The selected source for future deploys is `tc-data-kv-tc` in subscription
`1ce4dbd9-c69b-45aa-9718-e29e4efd232d`, True Collective tenant. The public
server's `MS365_MCP_CLIENT_SECRET` maps to
`ms365-mcp-server-ms365-mcp-client-secret`. Runtime uses the existing Sliplane
env copy. Non-secret app/tenant/redirect configuration stays in Sliplane.

TC's own destination is `tc-data-kv-tc`, subscription
`1ce4dbd9-c69b-45aa-9718-e29e4efd232d`, tenant
`f783e0f3-09c8-458b-aaa1-d46d8e870e16`. The user selected this source on
25 October 2026 and explicitly waived the Claude browser callback test for
this source-only change; record it as untested, not passed. No automatic
external vault sync was identified. Future manual secret maintenance uses
this exact TC binding. Retain DC-hosted `tc-data-kv` as the rollback source.
Preserve DC's separate `dc-data-kv` and all its consumers.

Never commit or print values, token-bearing URLs or hashes of credentials.
Never treat `<redacted>` as a credential. Copy unchanged values only after
checking the exact source and target; do not overwrite differing target copies.
Changing Sliplane env or deploying requires the approved release flow and
functional checks; it is not part of storage staging.
