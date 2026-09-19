# NMDC Coil Dispatch — transporter self-service + live dashboard

## Layout
- **GitHub Pages `dispatch.bharatsteels.in`** (repo BSC23609/weighbridge-data): add `nmdc.html` next to the weighbridge `index.html`, set `apiBase` in its CONFIG to the Vercel URL, add a nav link. Pages stays as it is.
- **Vercel** (this repo): APIs, sync, transporter page `/t/<token>`. Set `ALLOWED_ORIGINS=https://dispatch.bharatsteels.in`. The Vercel root shows a blank landing page only — the dashboard is not served from here, so transporters who trim their link see nothing.

## What's in here
| File | Purpose |
|---|---|
| `nmdc.html` | Dashboard for dispatch.bharatsteels.in (reads Neon via API, staff PIN). Yard marks a vehicle delivered with the Yard PIN. |
| `index.html` | Blank landing page for the Vercel root. |
| `t.html` | Transporter page, opened from the WhatsApp link `/t/<token>`. |
| `api/sync.js` | Every 5 min: Excel → Neon import, dispatch creation, WhatsApp via WATI, Neon → Excel mirror. |
| `api/tform.js` | Transporter form API (load / gate-out / delivered). |
| `api/arrive.js` | Yard marks delivered (PIN). |
| `api/dispatch.js` | Dashboard data. |
| `lib/` | Neon, Graph and WATI helpers. |
| `schema.sql` | Run once in Neon. |

## Two ways a dispatch gets in
1. **BSC pre-fills** the coil rows in Excel with Transporter Name (+ Mobile). The transporter is WhatsApped and completes truck details + status from their link.
2. **Transporter from scratch**: register the transporter once on the **Transporters** sheet (Name, Mobile, Active=Yes). The sync creates a standing link, writes it into the sheet and WhatsApps it (template `nmdc_transporter_link`). On that link they see all their dispatches (both modes) and can create new ones (DO, SO, destination, coils + weights, truck). Those rows are appended to `tbl_Dispatch` with **Entered By** = transporter name and are owned by the system (edit through the transporter, not in Excel).

## Ownership rule
Excel (BSC staff) owns: SO/DO, coil no, form, grade, size, weight, destination, transporter name + mobile, remarks.
Neon (transporter / yard) owns: status, vehicle no, LR no, dispatch date, expected & actual arrival. These are written back into the same Excel row — do not overtype them, and never sort/insert/delete rows in `tbl_Dispatch`.

## Environment variables (Vercel → Settings → Environment Variables)
| Name | Value |
|---|---|
| `DATABASE_URL` | Neon connection string (pooled) |
| `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` | Entra app registration (needs **Application** permission `Files.ReadWrite.All`, admin-consented) |
| `SHARE_LINK` | OneDrive sharing link of the workbook |
| `TABLE_NAME` | optional, default `tbl_Dispatch` |
| `CRON_SECRET` | long random string; cron-job.org sends it as `Authorization: Bearer …` |
| `PUBLIC_URL` | e.g. `https://dispatch.bharatsteels.in` (used in the WhatsApp link) |
| `YARD_PIN` | PIN for the "Mark delivered" button |
| `DASH_PIN` | required — PIN staff enter once to view the NMDC dashboard |
| `ALLOWED_ORIGINS` | `https://dispatch.bharatsteels.in` (comma-separated if more) |
| `WATI_API_URL` | WATI API endpoint, e.g. `https://live-mt-server.wati.io/<tenant-id>` |
| `WATI_API_KEY` | WATI access token |
| `WATI_TEMPLATE` | approved template name, default `nmdc_dispatch_link` |
| `WATI_TEMPLATE_TRANSPORTER` | approved template name, default `nmdc_transporter_link` |

## WATI template to create (Utility category)
Name: `nmdc_dispatch_link` — variables: name, coils, do_no, link
> Hello {{name}}, Bharat Steel (Chennai) has assigned {{coils}} coil(s) under DO {{do_no}} to you for pickup from NMDC.
> Please enter the vehicle & LR details and update the trip here: {{link}}
> — Bharat Steel logistics

Second template (Utility) — name `nmdc_transporter_link`, variables: name, link
> Hello {{name}}, this is your Bharat Steel (Chennai) dispatch link for NMDC pickups. Use it to log new dispatches and update vehicle, LR and delivery status: {{link}}
> Keep this message — the link is yours and does not expire.

## Scheduling
cron-job.org → URL `https://<your-domain>/api/sync`, every 5 minutes, GET, header `Authorization: Bearer <CRON_SECRET>`.
A healthy run returns `{"ok":true,"transporters":0,"imported":0,"updated":0,"dispatches":0,"notified":0,"mirrored":0,"appended":0,"errors":[]}`.

Existing databases: re-run `schema.sql` — the v2 block at the bottom adds the `transporters` table and new columns without touching data.
Until WATI is configured, sync still runs and links are created — `notify_error` in the `dispatches` table shows "WATI not configured"; you can copy the link from there and send it by hand.
