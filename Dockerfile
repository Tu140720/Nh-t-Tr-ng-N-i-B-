FROM node:22-bookworm-slim

WORKDIR /app

RUN apt-get update && \
    apt-get install -y --no-install-recommends python3 make g++ && \
    rm -rf /var/lib/apt/lists/*

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY . .

RUN mkdir -p /app/data /app/bootstrap-data && \
    cp -a /app/data/. /app/bootstrap-data/ 2>/dev/null || true && \
    sed -i 's/\r$//' /app/scripts/render-entrypoint.sh && \
    chmod +x /app/scripts/render-entrypoint.sh

ENV NODE_ENV=production
ENV PORT=10000

EXPOSE 10000

ENTRYPOINT ["/app/scripts/render-entrypoint.sh"]
CMD ["node", "server.mjs"]
