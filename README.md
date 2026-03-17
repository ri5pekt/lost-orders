# Lost Orders — Invoice Exporter

Web app that lets staff log in with their `@particleformen.com` Google account, paste WooCommerce order IDs, and download a combined PDF of all matching invoice emails from Gmail.

## Architecture

```
┌─────────────────────────────────┐
│  web  (Node.js + Vue 3)         │  :3000
│  ├─ Google OAuth login          │
│  ├─ JWT session (httpOnly)      │
│  └─ proxies export requests     │
└────────────────┬────────────────┘
                 │ http://pdf-service:5000
┌────────────────▼────────────────┐
│  pdf-service  (Python + Flask)  │
│  ├─ Gmail API (searches emails) │
│  ├─ Playwright → per-order PDFs │
│  └─ merges → combined PDF       │
└─────────────────────────────────┘
```

## Quick start

### 1. Prerequisites

- Docker + Docker Compose on the server
- A Google Cloud project with:
  - **Gmail API** enabled (existing `credentials.json` + `token.pickle`)
  - A new **OAuth 2.0 Web Client ID** for user login (see below)

### 2. Google OAuth Web Client

Go to [console.cloud.google.com](https://console.cloud.google.com) → APIs & Services → Credentials → **Create Credentials → OAuth client ID**:
- Application type: **Web application**
- Authorised redirect URI: `https://yourdomain.com/auth/google/callback`

Copy the Client ID and Secret into `.env`.

### 3. Environment

```bash
cp .env.example .env
# fill in: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_CALLBACK_URL, JWT_SECRET
```

### 4. Credentials

Place `credentials.json` and `token.pickle` in the project root (they are gitignored and must be copied to the server manually):

```bash
scp credentials.json token.pickle root@yourserver:/var/www/lost-orders/
```

### 5. Run

```bash
docker compose up -d --build
```

App runs on port **3000**. Put Nginx in front for HTTPS:

```nginx
server {
    listen 443 ssl;
    server_name yourdomain.com;

    location / {
        proxy_pass http://localhost:3000;
        proxy_read_timeout 600s;
        proxy_send_timeout 600s;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## Legacy scripts

The original standalone Python scripts (batch export, Gmail extractor) are in `scripts/`.
See `scripts/README.md` for usage.
