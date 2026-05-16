# ── Stage 1: Install dependencies ──
FROM node:20-alpine AS deps
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm ci --production

# ── Stage 2: Final image ──
FROM node:20-alpine
WORKDIR /app

# OpenShift runs containers as arbitrary non-root UIDs.
# Ensure /app/data is writable by any UID.
RUN mkdir -p /app/data && chmod 777 /app/data

COPY --from=deps /app/node_modules ./node_modules
COPY server.js ./
COPY public/ ./public/

# Default seed data — will be copied to /app/data on first run if missing
COPY public/scs_data.json ./public/scs_data.json

EXPOSE 8080

# Run as non-root (OpenShift compat)
USER 1001

CMD ["node", "server.js"]
