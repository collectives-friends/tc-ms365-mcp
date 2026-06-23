FROM node:24-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run generate
RUN npm run build

FROM node:24-alpine AS release

WORKDIR /app

# Infisical CLI — secrets injected from the vault at runtime
RUN apk add --no-cache wget ca-certificates && \
    ARCH=$(uname -m); case "$ARCH" in x86_64) A=amd64;; aarch64) A=arm64;; *) A=amd64;; esac; \
    wget -qO /tmp/inf.tar.gz "https://github.com/Infisical/cli/releases/download/v0.43.96/cli_0.43.96_linux_${A}.tar.gz" && \
    tar -xzf /tmp/inf.tar.gz -C /usr/local/bin infisical && \
    rm /tmp/inf.tar.gz && chmod +x /usr/local/bin/infisical

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/node_modules /app/node_modules
COPY --from=builder /app/package*.json ./
COPY infisical-entrypoint.sh /app/infisical-entrypoint.sh
RUN chmod +x /app/infisical-entrypoint.sh

ENV NODE_ENV=production

ENTRYPOINT ["/app/infisical-entrypoint.sh"]
