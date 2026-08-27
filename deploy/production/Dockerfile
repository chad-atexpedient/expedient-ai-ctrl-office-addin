# syntax=docker/dockerfile:1

FROM node:22-alpine AS build
WORKDIR /app
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
COPY index.html tsconfig.json vite.config.ts ./
COPY public ./public
COPY src ./src
COPY server ./server
COPY tools ./tools
RUN corepack enable && pnpm install --frozen-lockfile --ignore-scripts
RUN node node_modules/vite/bin/vite.js build

FROM node:22-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV HOST=0.0.0.0
ENV PORT=3000
# Apply latest Alpine security patches (e.g. openssl/libcrypto3/libssl3 fixes)
# and remove the bundled npm/npx CLI: the runtime only ever runs
# `node server/server.mjs` directly and never invokes npm, but the base
# image ships npm's own vendored dependency tree (tar, sigstore, picomatch,
# brace-expansion, ip-address, ...) which otherwise shows up as unused
# attack surface in container vulnerability scans.
RUN apk upgrade --no-cache && \
    rm -rf /usr/local/lib/node_modules/npm /usr/local/bin/npm /usr/local/bin/npx /usr/local/bin/corepack
COPY package.json ./package.json
COPY server ./server
COPY --from=build /app/dist ./dist
RUN addgroup -S app && adduser -S app -G app && chown -R app:app /app
USER app
EXPOSE 3000
CMD ["node", "server/server.mjs"]
