# Logictics Centre

Logictics Centre is a Neon-backed browser app for managing driver-separated pickup and delivery entries.

## User guide

- End-user walkthrough: [USER-WALKTHROUGH.md](USER-WALKTHROUGH.md)

## Current workflow

- The browser talks only to the local Node server.
- The local Node server executes the allowed PostgreSQL RPC functions against Neon.
- Admin and sales users can download the global list as CSV, send it to `admin3@giftwrap.co.za`, or run a test email.
- Microsoft Graph delivery sends from `artwork3@giftwrap.co.za` once `mail-config.js` or the `MAIL_*` environment variables are configured.
- When active driver-list items roll forward to a new day, the server emails a grouped carry-over breakdown to the configured `MAIL_TO` recipient.
- The UI now uses role-based dynamic page navigation instead of one long scrolling screen.

## Files

- `index.html` - app shell and page navigation mount
- `app.js` - UI, page navigation, CSV download, and email actions
- `styles.css` - layout and styling
- `serve.js` - static server, Neon RPC API, and email endpoints
- `neon-config.example.js` - local Neon connection string template
- `mail-config.example.js` - local Microsoft Graph config template
- `neon.sql` - tables, functions, sessions, and entry logic for Neon/Postgres

## Setup

1. Create a Neon project and database.
2. Run `npm install`.
3. Create `neon-config.js` from `neon-config.example.js`, or set `NEON_DATABASE_URL` / `DATABASE_URL`.
4. Create `mail-config.js` from `mail-config.example.js`, or set the `MAIL_*` environment variables.
5. Run `neon.sql` against your Neon database.
6. Start the local server with `node serve.js`.
7. Open `http://127.0.0.1:4173/`.

## Microsoft Graph mail

- The default provider is `microsoft-graph`.
- The app expects `MAIL_TENANT_ID`, `MAIL_CLIENT_ID`, and `MAIL_CLIENT_SECRET`, either in `mail-config.js` or environment variables.
- `MAIL_CLIENT_SECRET` must be the Azure client secret value, not the secret ID shown alongside it.
- The Azure app needs Microsoft Graph application permission to send mail as `artwork3@giftwrap.co.za`.
- `smtp` is still available as a fallback provider if you explicitly set `provider: "smtp"`.

## Entry fields

Each new entry now captures:

- Driver
- Pickup location
- Collection or delivery
- Quote number
- Sales order number
- Invoice number
- PO number
- Notice
- Created by

The created-by value comes from the signed-in user and is stored by the database.
