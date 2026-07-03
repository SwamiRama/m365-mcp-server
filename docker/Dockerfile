# syntax=docker/dockerfile:1

# =============================================================================
# Build stage
# =============================================================================
FROM node:26-alpine@sha256:725aeba2364a9b16beae49e180d83bd597dbd0b15c47f1f28875c290bfd255b9 AS builder

WORKDIR /app

# Install dependencies
COPY package*.json ./
RUN npm ci --ignore-scripts

# Copy source and build
COPY tsconfig.json ./
COPY src ./src
RUN npm run build

# Remove dev dependencies
RUN npm prune --production

# =============================================================================
# Production stage
# =============================================================================
FROM node:26-alpine@sha256:725aeba2364a9b16beae49e180d83bd597dbd0b15c47f1f28875c290bfd255b9 AS production

# Security: Run as non-root user
RUN addgroup -g 1001 -S nodejs && \
    adduser -S mcp -u 1001 -G nodejs

WORKDIR /app

# Copy built artifacts
COPY --from=builder --chown=mcp:nodejs /app/dist ./dist
COPY --from=builder --chown=mcp:nodejs /app/node_modules ./node_modules
COPY --from=builder --chown=mcp:nodejs /app/package.json ./

# Set environment
ENV NODE_ENV=production
ENV PORT=3000

# No HEALTHCHECK - Azure Container Apps uses its own TCP probes.
# A Docker HEALTHCHECK conflicts with the platform's probe mechanism
# and causes unnecessary container restarts.

# Switch to non-root user
USER mcp

# Expose port
EXPOSE 3000

# Start server
CMD ["node", "dist/index.js"]
