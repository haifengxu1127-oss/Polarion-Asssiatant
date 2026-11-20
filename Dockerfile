FROM node:22-bookworm-slim

WORKDIR /app
COPY server/package.json server/package-lock.json ./server/
WORKDIR /app/server
RUN npm ci --omit=dev
WORKDIR /app
COPY . .

ENV NODE_ENV=production
EXPOSE 3000
CMD ["node", "--env-file-if-exists=server/.env", "server/server.js"]
