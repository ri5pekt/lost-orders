# ── Stage 1: build Vue client ─────────────────────────────────────────────────
FROM node:20-slim AS builder

WORKDIR /app/client
COPY client/package*.json ./
RUN npm install
COPY client/ ./
RUN npm run build

# ── Stage 2: production image ─────────────────────────────────────────────────
FROM node:20-slim

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

# Copy built Vue SPA
COPY --from=builder /app/client/dist ./client/dist

COPY server.js ./

EXPOSE 3000
CMD ["node", "server.js"]
