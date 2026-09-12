# AJCloseTheGap — Annie & Jay's Wedding

A small Node/Express site for Annie Queen C. Barbuco & Jay B. Josef's wedding —
February 21, 2027, 4:00 PM, Plaza Ibarra, Quezon City.

## Run locally

```
npm install
npm start
```

Visit http://localhost:3000

## RSVP storage

RSVPs are stored as JSON at `DATA_DIR/rsvps.json` (defaults to `./data`).
On Railway, `DATA_DIR` points at a mounted volume so responses survive
redeploys.

## Endpoints

- `POST /api/rsvp` — submit a response `{ name, attending: "yes"|"no", guests, message }`
- `GET /api/rsvps` — aggregate counts + the list of attending guests
- `GET /api/updates` — the wedding updates feed
- `GET /api/program` — the day's run-of-show
- `GET /api/table?name=` / `GET /api/table/:id` — find-your-table lookup
- `/admin` — unlisted admin page: RSVPs, seating list + Excel upload, per-guest QR codes, updates, and the program

## Seating list & QR codes

Upload an Excel file (`.xlsx`) from `/admin` with a "Name" column and a "Table" column —
guests are matched by name (case-insensitive) and updated in place; nothing is ever
deleted by an upload. Each guest then has a personal QR code (`/admin` → "View / Print QR")
that links straight to their table number and the program — no typing required when scanned.
