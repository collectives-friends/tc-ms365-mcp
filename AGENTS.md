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

## Secrets · Infisical universal-auth

- Container-entry er `infisical-entrypoint.sh`: logger ind med `INFISICAL_CLIENT_ID`/`INFISICAL_CLIENT_SECRET` (universal-auth) mod `INFISICAL_API_URL`, henter `INFISICAL_PROJECT_ID`/`INFISICAL_ENV`, kører `node dist/index.js`.
- **Nul secrets i git.** `.env.example` viser formen; udfyld aldrig rigtige værdier. Aldrig server-secrets i klienten.

## Deploy

- GitOps via Sliplane/Komodo. Ingen manuel container-manipulation eller prod-redeploy uden Stefans godkendelse. Propose-only: branch + draft-PR.
