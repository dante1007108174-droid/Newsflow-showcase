FROM node:22-alpine AS build  
LABEL "language"="nodejs"  
LABEL "framework"="vite"  
WORKDIR /app  
COPY . .  
RUN npm install -g pnpm && pnpm install  
WORKDIR /app/frontend  
RUN pnpm run build  
FROM zeabur/caddy-static  
COPY --from=build /app/frontend/dist /usr/share/caddy  
EXPOSE 8080  
