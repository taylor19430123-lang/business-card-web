FROM node:22-bookworm-slim

WORKDIR /app

ENV NODE_ENV=production
ENV HOST=0.0.0.0
ENV PORT=3010

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY assets ./assets
COPY public ./public
COPY templates ./templates
COPY server.js ./server.js

EXPOSE 3010

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:' + (process.env.PORT || 3010) + '/healthz').then((res) => { if (!res.ok) process.exit(1); }).catch(() => process.exit(1));"

CMD ["npm", "start"]
