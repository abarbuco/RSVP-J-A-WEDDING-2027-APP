const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "rsvps.json");

// --- tiny JSON-file "database" -------------------------------------------
function ensureStore() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, "[]", "utf8");
}

function readAll() {
  ensureStore();
  try {
    const raw = fs.readFileSync(DATA_FILE, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    console.error("Failed to read RSVP store, starting fresh:", err);
    return [];
  }
}

function writeAll(entries) {
  ensureStore();
  const tmp = DATA_FILE + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(entries, null, 2), "utf8");
  fs.renameSync(tmp, DATA_FILE);
}

ensureStore();

// --- middleware ------------------------------------------------------------
app.use(express.json({ limit: "10kb" }));
app.use(express.static(path.join(__dirname, "public")));

// --- API ---------------------------------------------------------------
app.post("/api/rsvp", (req, res) => {
  const body = req.body || {};
  const name = String(body.name || "").trim().slice(0, 80);
  const attending = body.attending === "yes" ? "yes" : body.attending === "no" ? "no" : null;
  let guests = parseInt(body.guests, 10);
  const message = String(body.message || "").trim().slice(0, 300);

  if (!name) return res.status(400).json({ error: "Please enter your name." });
  if (!attending) return res.status(400).json({ error: "Please let us know if you can make it." });
  if (attending === "yes") {
    if (!Number.isFinite(guests) || guests < 1) guests = 1;
    if (guests > 10) guests = 10;
  } else {
    guests = 0;
  }

  const entries = readAll();
  const entry = {
    id: crypto.randomUUID(),
    name,
    attending,
    guests,
    message,
    submittedAt: new Date().toISOString(),
  };
  entries.push(entry);
  writeAll(entries);

  res.status(201).json({ ok: true, entry });
});

app.get("/api/rsvps", (req, res) => {
  const entries = readAll();
  const attending = entries.filter((e) => e.attending === "yes");
  const declined = entries.filter((e) => e.attending === "no");
  const totalGuests = attending.reduce((sum, e) => sum + (e.guests || 1), 0);

  res.json({
    totalResponses: entries.length,
    attendingCount: attending.length,
    totalGuests,
    declinedCount: declined.length,
    guests: attending
      .slice()
      .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
      .map((e) => ({ name: e.name, guests: e.guests, submittedAt: e.submittedAt })),
  });
});

app.get("/healthz", (req, res) => res.json({ ok: true }));

app.listen(PORT, () => {
  console.log(`AJ Close The Gap wedding site listening on port ${PORT}`);
});
