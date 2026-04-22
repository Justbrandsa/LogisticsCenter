# Logictics Centre

Logictics Centre is a local-database browser app for managing driver-separated pickup and delivery entries.

## User guide

- End-user walkthrough: [USER-WALKTHROUGH.md](USER-WALKTHROUGH.md)

## Current workflow

- The browser talks only to the local Node server.
- The local Node server reads and writes a SQLite database file.
- It prefers `data/route-ledger.sqlite` inside the project, but can fall back to a writable runtime folder when the deploy bundle is read-only.
- The `/api/status` response reports the active storage path and whether the runtime is using temporary fallback storage.
- Admin and sales users can download the global list as CSV, send it to `admin3@giftwrap.co.za`, or run a test email.
- Microsoft Graph delivery sends from `artwork3@giftwrap.co.za` once `mail-config.js` or the `MAIL_*` environment variables are configured.
- When active driver-list items roll forward to a new day, the server emails a grouped carry-over breakdown to the configured `MAIL_TO` recipient.
- Deleted entries are written to a server-side delete log before removal, and the server batches unsent delete-log emails every 10 minutes.
- The UI now uses role-based dynamic page navigation instead of one long scrolling screen.

## Files

- `index.html` - app shell and page navigation mount
- `app.js` - UI, page navigation, CSV download, and email actions
- `styles.css` - layout and styling
- `serve.js` - static server, local database API, and email endpoints
- `local-database.js` - SQLite-backed project-local persistence layer
- `mail-config.example.js` - local Microsoft Graph config template
- `data/route-ledger.sqlite` - preferred local database file path when the project folder is writable

## Setup

1. Run `npm install`.
2. Create `mail-config.js` from `mail-config.example.js`, or set the `MAIL_*` environment variables.
3. Start the local server with `node serve.js`.
4. Open `http://127.0.0.1:4173/`.
5. On first start, the app creates the SQLite file automatically and then asks for the first admin account.
6. Optional: set `LOGISTICS_DB_PATH` to choose an exact SQLite file path, or set `LOGISTICS_DATA_DIR` / `PERSISTENT_DATA_DIR` to choose the folder that should hold `route-ledger.sqlite`.
7. If you deploy on Vercel, add `CRON_SECRET` so the `/api/jobs/order-delete-log-email` cron endpoint can stay protected.
8. Set `LOGISTICS_REQUIRE_PERSISTENT_STORAGE=true` in production if you want the app to fail fast instead of silently falling back to temporary runtime storage.
9. Vercel-style serverless hosts do not offer durable project-local disk storage. If the app falls back to a temporary runtime folder, writes will work but the data can reset on restart, scale-out, or redeploy.

## Durable hosting for SQLite

- Vercel is not a long-term writable SQLite host for this project. It can read the bundled database file, and it can write to temporary runtime storage, but those writes are not durable across restarts or redeploys.
- For long-term writable SQLite, run `npm start` on a single-instance Node host with a persistent disk or volume.
- Railway: attach a Volume and mount it at `/app/data` for the simplest setup. The app will then write to `/app/data/route-ledger.sqlite`, and Railway also exposes `RAILWAY_VOLUME_MOUNT_PATH` automatically if you prefer a custom mount path.
- Render: attach a persistent disk and either mount it at `/opt/render/project/src/data` so the default project `data/` folder becomes durable, or mount it elsewhere and set `LOGISTICS_DATA_DIR` or `PERSISTENT_DATA_DIR` to that disk path. If you want to use `RENDER_DISK_PATH`, define it yourself as an app environment variable.
- Avoid multi-instance scaling when SQLite is the source of truth. This app expects one writable database file, one active writer host, and a persistent filesystem under that path.
- Keep using the built-in admin export or `node migrate-live-data.js` as a backup path before host moves or major deploy changes.

## Live data rescue

- Admins can export the full database through `POST /api/admin/data/export` with `{ "token": "<session-token>" }`.
- Admins can import a full exported dataset through `POST /api/admin/data/import` with `{ "token": "<session-token>", "data": { ... } }`.
- To pull data from a running live site into the current local SQLite file, run:
  `node migrate-live-data.js --source-url https://your-live-site.example --name "Admin Name" --password "admin-password"`
- Use `--token` instead of `--name` and `--password` if you already have an admin session token.
- The rescue command always writes a JSON backup under `data/` before it imports locally.

## Microsoft Graph mail

- The default provider is `microsoft-graph`.
- The app expects `MAIL_TENANT_ID`, `MAIL_CLIENT_ID`, and `MAIL_CLIENT_SECRET`, either in `mail-config.js` or environment variables.
- Optional mail-routing settings include `MAIL_FROM_NAME`, `MAIL_ADMIN_ACTION_TO`, and `MAIL_ROLLOVER_TEST_TO`, and the `maintenance` role can now override the live inbox routing from inside the app.
- `MAIL_CLIENT_SECRET` must be the Azure client secret value, not the secret ID shown alongside it.
- The Azure app needs Microsoft Graph application permission to send mail as `artwork3@giftwrap.co.za`.
- The app can request a sender label such as `Logistics Centre`, but Microsoft Graph may still show the mailbox profile name that is configured in Microsoft 365.
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
