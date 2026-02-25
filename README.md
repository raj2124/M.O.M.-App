# M.O.M App

M.O.M web app built from your provided M.O.M PDF format (`MOM-1 Page (1) 1.pdf`) with these flows:

- Click **New M.O.M Sheet**
- Fetch project data from Zoho Projects (project no, client name, project name)
- Fill remaining details manually
- Submit with options:
  - Mail M.O.M to client
  - Generate PDF
  - Print PDF
  - Any combination of the above

## What is implemented

- Sheet sections aligned to the PDF labels:
  - `1. GENERAL INFORMATION`
  - `2. DETAILS OF MEETING`
  - `3. MINUTES`
  - `4. AGENDA`
  - `Attendees`
- Zoho project search + select
- Form capture with dynamic rows for Agenda and Attendees
- Server-side PDF generation
- Optional SMTP email with PDF attachment
- Print via browser PDF print flow

## Tech stack

- Node.js + Express
- Vanilla HTML/CSS/JS frontend
- PDFKit for PDF generation
- Nodemailer for email

## Project structure

- `./src/server.js` - API and app server
- `./src/zohoClient.js` - Zoho integration
- `./src/pdfService.js` - PDF generation
- `./src/emailService.js` - SMTP email sending
- `./public/index.html` - App UI
- `./public/app.js` - Frontend logic
- `./public/styles.css` - Styling

## Setup

1. Install dependencies:

```bash
npm install
```

2. Copy environment file and fill values:

```bash
cp .env.example .env
```

Recommended local bind settings:
- `APP_HOST=127.0.0.1`
- `PORT=3000`

3. For Zoho live mode:

- Set `ZOHO_USE_MOCK=false`
- Set `ZOHO_PORTAL_ID`
- Set `ZOHO_ACCESS_TOKEN` (OAuth access token)
- Set `ZOHO_BASE_URL` and `ZOHO_ACCOUNTS_BASE_URL` for your Zoho data center (`.com` or `.in`)
- Set `ORG_USER_EMAIL_DOMAIN` (default `elegrow.com`) to restrict Elegrow attendee dropdown users
- Recommended for auto-renew:
  - `ZOHO_REFRESH_TOKEN`
  - `ZOHO_CLIENT_ID`
  - `ZOHO_CLIENT_SECRET`
- If your org uses a custom endpoint, set `ZOHO_PROJECTS_ENDPOINT`

4. For email sending:

- Set `SMTP_ENABLED=true`
- Fill SMTP settings (`SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`)

5. Run app:

```bash
npm run dev
```

6. Open:

- `http://localhost:3000`

## Notes

- Default mode is mock Zoho data (`ZOHO_USE_MOCK=true`) for local testing.
- Generated PDFs are stored in:
  - `./generated-pdfs`
- If the Zoho response schema in your portal differs, adjust field mapping in:
  - `./src/zohoClient.js`
- To show your official brand logo in the dashboard header, place files at:
  - `./public/assets/elegrow-logo-full.png`
  - optional symbol only: `./public/assets/elegrow-symbol.png`
