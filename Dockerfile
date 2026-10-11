FROM node:24-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run generate
RUN npm run build

FROM node:24-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/node_modules /app/node_modules
COPY --from=builder /app/package*.json ./

ENV NODE_ENV=production

# Secrets come from plain env (Sliplane env, synced from Key Vault tc-data-kv-tc,
# names ms365-mcp-server--<ENV-NAME>). No Infisical at startup.
ENTRYPOINT ["node", "dist/index.js"]
