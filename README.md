# AJCloseTheGap — Anne & Jay's Wedding

A small Node/Express site for Anne Queen C. Barbuco & Jay B. Josef's wedding —
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
