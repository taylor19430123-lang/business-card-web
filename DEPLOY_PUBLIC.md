# Public Deployment Guide

This project can now be deployed as a public web app and embedded in Feishu Miaoda.

## What changed

- Fonts no longer depend on local Windows paths.
- The app exposes `GET /healthz` for health checks.
- Upload size, host, proxy, and iframe embedding are configurable with environment variables.
- Security headers now allow Feishu pages to embed the app by default.

## Required environment

Copy [.env.example](/D:/code/business-card-web/.env.example) and fill in values as needed.

Important variables:

- `HOST=0.0.0.0`
- `PORT=3010`
- `TRUST_PROXY=true` if you run behind Nginx / Caddy / a cloud load balancer
- `FEISHU_APP_ID` and `FEISHU_APP_SECRET` only if you need Feishu sending

## Docker deployment

Build:

```bash
docker build -t business-card-web .
```

Run:

```bash
docker run -d \
  --name business-card-web \
  -p 3010:3010 \
  --env-file .env \
  business-card-web
```

Health check:

```bash
curl http://127.0.0.1:3010/healthz
```

## Reverse proxy / HTTPS

Put the app behind HTTPS before embedding it in Miaoda.

Recommended setup:

1. A public domain, for example `https://cards.example.com`
2. Nginx / Caddy / cloud ingress forwarding traffic to the container on port `3010`
3. Keep the app response header `Content-Security-Policy` intact, because it already includes Feishu frame ancestors

## Miaoda embed

After deployment succeeds:

1. Open your public URL directly and confirm the app loads
2. Open `https://your-domain/healthz` and confirm it returns `{ "ok": true }`
3. Put the public HTTPS URL into the Miaoda embed file `APP_URL`

## Notes

- Multi-employee export downloads a ZIP, with one PDF per employee name.
- Single-employee export downloads a single employee-named PDF.
- If a reverse proxy injects `X-Frame-Options: SAMEORIGIN`, Miaoda embedding will fail. Remove that header at the proxy layer.
