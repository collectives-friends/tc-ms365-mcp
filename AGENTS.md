# tc-ms365-mcp · Microsoft 365 MCP-server (fork)

Fork af `softeria/ms-365-mcp-server` (npm-navn stadig `@softeria/ms-365-mcp-server`). MCP-server mod Microsoft Graph. Kører som container på Sliplane i DC-flåden. Origin er `collectives-friends/tc-ms365-mcp` (ingen upstream-remote wired; rebase manuelt fra softeria hvis nødvendigt).

> Dev-box-baseline (`~/.claude/METHODOLOGY.md`) dækker den tværgående metodik. Denne fil er kun det repo-specifikke.

## Stack · Node/TS (npm, ikke pnpm)

- Lockfile = `package-lock.json` → brug **npm**, ikke pnpm.
- Verify før "done": `npm run verify` (kører build + lint + test). Enkeltdele: `npm run build`, `npm run lint`, `npm test`, `npm run dev`.
- **Fork-disciplin:** hold ændringer minimale og rebasable. Rør ikke upstream-filer bredt uden grund, det gør fremtidig rebase fra softeria dyr.

## Domæne · DC MCP-flåde

- Nye/ændrede tools: følg Anthropic search+execute-pattern + korrekt auth/elicitation, ikke ad-hoc tool-lister.
- `.claude/rules/` findes allerede til path-scoped regler; læg finkornede regler dér.

## Secrets · Key Vault + Sliplane env

- Selected source for future deployments: Azure Key Vault `tc-data-kv-tc` in True Collective's tenant (subscription `1ce4dbd9-c69b-45aa-9718-e29e4efd232d`): `MS365_MCP_CLIENT_SECRET` → `ms365-mcp-server-ms365-mcp-client-secret`. The non-secret env (CLIENT_ID/TENANT_ID/ORG_MODE/PUBLIC_URL/ALLOWED_REDIRECT_URIS) lives only in Sliplane. TC's own vault `tc-data-kv-tc` (subscription `1ce4dbd9-c69b-45aa-9718-e29e4efd232d`) was accessible after its subscription was reactivated on 24/10/2026. The user authorized this source-only selection on 25/10/2026 after a fresh ten-binding value/version comparison and explicitly waived the Claude browser callback test. That test is unverified, not passed. Future manual secret maintenance must use the exact TC binding above. Retain legacy `tc-data-kv` in DC for rollback. No automatic external vault sync was identified, and no service-env or deployment is changed. See [the migration runbook](docs/tc-vault-migration.md).
- Runtime: values are set as Sliplane env (secret=true for CLIENT_SECRET). The container starts with `node dist/index.js --http :3000`. Infisical is no longer used (it was shut down 2026-10).
- Entra app: TC Claude Code (ee95c1be-…). On client-secret rotation: new secret → KV → Sliplane env → redeploy.

## Deploy

- GitOps via Sliplane/Komodo. Ingen manuel container-manipulation eller prod-redeploy uden Stefans godkendelse. Propose-only: branch + draft-PR.
