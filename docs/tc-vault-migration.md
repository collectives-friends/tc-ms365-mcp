# TC vault migration: public Microsoft 365 MCP

Selected source for future manual secret maintenance: TC vault `tc-data-kv-tc`.
The user authorized this source-only selection on 25 October 2026 and explicitly
waived the Claude browser callback test. This changes the operative source
procedure, not the running process or a provider-side env binding. No automatic
external vault sync was identified. The 24 October acceptance below remains
historical; a fresh 25 October comparison verified the same ten source/target
values and unchanged version IDs before activation. Both source and destination
subscriptions have been observed active. No service env, image, OAuth app
or runtime configuration has been changed by this migration. Existing token
caches were retained; normal token refresh was allowed during read tests.

The public app secret and non-secret app/tenant IDs now match the active
process environment. The destination secret was also compared independently
with its source version through Azure API reads. A delegated TC Graph token
authenticated to this public MCP server and `get-current-user` returned the
expected TC user; unauthenticated MCP calls returned 401. The full browser
OAuth callback flow through the existing Claude client was not tested and
was explicitly waived by the user for this source-only change. It is not
recorded as passed. No callback URI or consent configuration was changed.

| Binding                   | Retained rollback source                              | Selected TC source                                       |
| ------------------------- | ----------------------------------------------------- | -------------------------------------------------------- |
| `MS365_MCP_CLIENT_SECRET` | `tc-data-kv/ms365-mcp-server-ms365-mcp-client-secret` | `tc-data-kv-tc/ms365-mcp-server-ms365-mcp-client-secret` |

Source subscription: `1668f090-c347-4ef4-b896-b770b044afb0`, DC tenant
`0b178f2d-c843-487a-a89b-8dbebc045d81`.
Target subscription: `1ce4dbd9-c69b-45aa-9718-e29e4efd232d`, TC tenant
`f783e0f3-09c8-458b-aaa1-d46d8e870e16`.
Both vaults use resource group `tc-infra-swc` in their respective subscriptions.
DC's own `dc-data-kv` remains separate and is outside this migration.

The target's older `ms365-mcp-server--MS365-MCP-CLIENT-SECRET` has a different
name. Do not overwrite it or assume equality. Stage the current source under
the exact destination name above and compare values in memory, recording only
version IDs and equality. The ten-binding, read-only preflight is maintained
in `tc-multica/scripts/keyvault_migration.py` with `vault-migration.json`.
Its reviewed activation enables selected Multica adapter reads; that switch
does not configure this service's external deploy process. Select the exact
TC target above for future manual secret maintenance; retain the DC source for rollback.

Sliplane's secret field is redacted in the UI. A source/destination match is
not proof of live runtime equality. Accept an authorized runtime comparison
and a read-only Graph call. The direct authenticated server test passed on
24 October; the user waived the separate Claude browser callback test for
this source-only selection. Future manual secret maintenance must read the
exact TC vault/secret above. This does not install or modify an external deployer.

Keep client ID `ee95c1be-8f3e-49ef-9798-0f2675b0e5be`, the TC tenant ID and
existing redirect/public URLs unchanged. The private Multica MCP uses a
different app (`63d92b22-5465-4656-ad53-5b96fd7f3205`) and a different secret;
do not substitute one app's credentials for the other.

During staging, do not write Sliplane env or redeploy. Keep the old source and
bindings available for rollback. Rotation, hosting transfer and source cleanup
are later, separate operations. Do not enable `MS365_MCP_KEYVAULT_URL`: its
optional runtime reader uses different secret names and is not this deploy
source migration.
