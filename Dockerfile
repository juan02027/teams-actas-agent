FROM node:22-bookworm-slim AS build
WORKDIR /app

RUN corepack enable && corepack prepare pnpm@10.4.1 --activate
COPY package.json pnpm-lock.yaml ./
COPY patches ./patches
RUN pnpm install --frozen-lockfile

COPY . .
# Vite embeds VITE_* values in the browser bundle at build time.
ARG VITE_MICROSOFT_CLIENT_ID
ARG VITE_MICROSOFT_TENANT_ID
ARG VITE_PUBLIC_APP_URL
ARG VITE_LOCAL_DEMO_MODE=true
ARG VITE_SHAREPOINT_SITE_URL
ARG VITE_SHAREPOINT_LIST_NAME
ARG VITE_SHAREPOINT_COLUMN_NAME
ENV VITE_MICROSOFT_CLIENT_ID=${VITE_MICROSOFT_CLIENT_ID}
ENV VITE_MICROSOFT_TENANT_ID=${VITE_MICROSOFT_TENANT_ID}
ENV VITE_PUBLIC_APP_URL=${VITE_PUBLIC_APP_URL}
ENV VITE_LOCAL_DEMO_MODE=${VITE_LOCAL_DEMO_MODE}
ENV VITE_SHAREPOINT_SITE_URL=${VITE_SHAREPOINT_SITE_URL}
ENV VITE_SHAREPOINT_LIST_NAME=${VITE_SHAREPOINT_LIST_NAME}
ENV VITE_SHAREPOINT_COLUMN_NAME=${VITE_SHAREPOINT_COLUMN_NAME}
RUN pnpm check && pnpm build

FROM node:22-bookworm-slim AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV HOST=0.0.0.0
ENV PORT=3000
ENV LOCAL_DATA_DIR=/app/data

RUN corepack enable && corepack prepare pnpm@10.4.1 --activate
COPY package.json pnpm-lock.yaml ./
COPY patches ./patches
# The server bundle keeps the Vite bridge external in this project, so the
# runtime must include Vite and the other dev dependencies as well.
RUN pnpm install --frozen-lockfile && pnpm store prune
COPY --from=build /app/dist ./dist
COPY --from=build /app/server ./server
COPY --from=build /app/shared ./shared
COPY --from=build /app/client ./client

RUN mkdir -p /app/data/recordings /app/data/documents && chown -R node:node /app/data
USER node
VOLUME ["/app/data"]
EXPOSE 3000
CMD ["node", "dist/index.js"]
