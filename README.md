# M.O.M Integration

Minutes of Meeting web app with Zoho Projects integration, PDF export, print support, and Outlook draft handoff.

## Features

- Dashboard with recent Zoho projects and project status badges
- M.O.M editor aligned to your sheet format
- Zoho project selection from dropdown while creating a new M.O.M
- Auto-fill of project fields from Zoho (project name/number and team users where available)
- Agenda and attendees row management
- PDF generation and browser print flow
- Microsoft Graph server-side draft creation (with PDF attachment) when configured
- Outlook compose deeplink fallback when Graph is not configured

## Stack

- Node.js + Express
- Vanilla HTML/CSS/JS
- PDFKit

## Project Structure

- `src/server.js` - API and app server
- `src/zohoClient.js` - Zoho Projects integration
- `src/pdfService.js` - PDF generation
- `public/index.html` - App UI
- `public/app.js` - Frontend logic
- `public/styles.css` - Styling
- `scripts/smoke-test.js` - smoke checks

## Local Setup

1. Install dependencies

```bash
npm install
```

2. Create environment file

```bash
cp .env.example .env
```

3. Configure Zoho (live mode)

- Set `ZOHO_USE_MOCK=false`
- Set `ZOHO_PORTAL_ID`
- Set `ZOHO_CLIENT_ID`
- Set `ZOHO_CLIENT_SECRET`
- Set `ZOHO_REFRESH_TOKEN`
- Set `ZOHO_BASE_URL` and `ZOHO_ACCOUNTS_BASE_URL` for your Zoho data center (`.com` or `.in`)
- Optional: set `ZOHO_PROJECTS_ENDPOINT` if your org uses a custom endpoint

4. Configure Microsoft Graph draft mode (recommended)

- Set `MS_GRAPH_ENABLED=true`
- Set `MS_GRAPH_TENANT_ID`
- Set `MS_GRAPH_CLIENT_ID`
- Set `MS_GRAPH_CLIENT_SECRET`
- Set `MS_GRAPH_MAILBOX_USER` (mailbox where drafts should be created)
- In Azure App Registration, add **Application** permission `Mail.ReadWrite` and grant admin consent.
- Keep `APP_BASE_URL` set to your deployed public URL so PDF links are valid externally.

5. Run app

```bash
npm run dev
```

6. Open `http://localhost:3000`

## Deployment Notes

- Keep all secrets in host environment variables (Render/GitHub/other host), never in Git.
- Generated PDFs are written to `generated-pdfs/`.
- In Graph mode, server creates real Outlook drafts and attaches generated PDF automatically.
- In fallback deeplink mode, browser security does not allow automatic file attachment; users attach PDF manually.
- For GoDaddy VPS auto-deploy with GitHub Actions, see:
  - `docs/DEPLOY_GODADDY_VPS.md`
  - `.github/workflows/deploy-vps.yml`

## Graph Diagnostics API

Use this endpoint to validate Graph configuration and pinpoint token/permission/mailbox issues without exposing secrets:

```bash
curl -sS "http://127.0.0.1:3000/api/graph/diagnostics"
```

Optional write probe (creates and deletes a test draft):

```bash
curl -sS "http://127.0.0.1:3000/api/graph/diagnostics?writeProbe=true"
```

The response includes:
- configuration readiness flags (`enabled`, `isConfigured`, required fields present)
- token status (`ok`, expiry, audience, roles/scopes, structured error)
- mailbox access check
- optional draft create/delete probe result
- actionable `hints` for fixes (for 401/403 and permission gaps)

## License

MIT (see `LICENSE`).
