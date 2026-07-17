# tc-ms365-mcp

## Metodik

Denne stack følger den kanoniske udviklings- og deployment-tilgang i `claude-loadout/core/METHODOLOGY.md`: fingerprint stack og domæne først · vault-secrets (aldrig disk eller git) · git draft-PR · GitOps-deploy · propose-only (merge, deploy eller send aldrig uden Stefans godkendelse) · verifikation før "done". På dev-box er den fulde metodik i kontekst via user-memory. Repo-specifikke regler nedenfor vinder ved konflikt.

Repo-specifikke regler tilføjes her efter behov.
