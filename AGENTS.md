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
- Source of truth: Azure Key Vault `tc-data-kv-tc`, names `ms365-mcp-server--<ENV-NAME>` (underscores → dashes), 1:1 with the env names (MS365_MCP_CLIENT_ID/_SECRET/_TENANT_ID/_ORG_MODE/_PUBLIC_URL/_ALLOWED_REDIRECT_URIS). Tagged with expires.
- Runtime: values are set as Sliplane env (secret=true for CLIENT_SECRET). The container starts with `node dist/index.js --http :3000`. Infisical is no longer used (it was shut down 2026-10).
- Entra app: TC Claude Code (ee95c1be-…). On client-secret rotation: new secret → KV → Sliplane env → redeploy.

## Deploy

- GitOps via Sliplane/Komodo. Ingen manuel container-manipulation eller prod-redeploy uden Stefans godkendelse. Propose-only: branch + draft-PR.
