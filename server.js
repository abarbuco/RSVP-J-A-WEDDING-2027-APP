const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const QRCode = require("qrcode");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "rsvps.json");
const UPDATES_FILE = path.join(DATA_DIR, "updates.json");
const PROGRAM_FILE = path.join(DATA_DIR, "program.json");
const SEATING_FILE = path.join(DATA_DIR, "seating.json");

// --- tiny JSON-file "database" -------------------------------------------
function ensureFile(file) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(file)) fs.writeFileSync(file, "[]", "utf8");
}

function readJson(file) {
  ensureFile(file);
  try {
    const raw = fs.readFileSync(file, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    console.error(`Failed to read ${file}, starting fresh:`, err);
    return [];
  }
}

function writeJson(file, entries) {
  ensureFile(file);
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(entries, null, 2), "utf8");
  fs.renameSync(tmp, file);
}

function readAll() { return readJson(DATA_FILE); }
function writeAll(entries) { writeJson(DATA_FILE, entries); }
function readUpdates() { return readJson(UPDATES_FILE); }
function writeUpdates(entries) { writeJson(UPDATES_FILE, entries); }
function readProgram() { return readJson(PROGRAM_FILE); }
function writeProgram(entries) { writeJson(PROGRAM_FILE, entries); }
function readSeating() { return readJson(SEATING_FILE); }
function writeSeating(entries) { writeJson(SEATING_FILE, entries); }

ensureFile(DATA_FILE);
ensureFile(UPDATES_FILE);
ensureFile(PROGRAM_FILE);
ensureFile(SEATING_FILE);

function normName(n) {
  return String(n || "").trim().toLowerCase();
}

// Add or update one seating entry by name (case-insensitive match).
function upsertSeating(list, rawName, rawTable) {
  const name = String(rawName || "").trim().slice(0, 80);
  const table = String(rawTable || "").trim().slice(0, 20);
  const key = normName(name);
  if (!key) return { entry: null, created: false };
  const idx = list.findIndex((s) => normName(s.name) === key);
  if (idx === -1) {
    const entry = { id: crypto.randomUUID(), name, table: table || null };
    list.push(entry);
    return { entry, created: true };
  }
  list[idx].table = table || null;
  if (name) list[idx].name = name;
  return { entry: list[idx], created: false };
}

// --- middleware ------------------------------------------------------------
// 8mb covers the small text payloads everywhere else plus a base64-encoded
// seating spreadsheet upload from /admin.
app.use(express.json({ limit: "8mb" }));

// The page lives at the repo root (index.html next to this file).
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

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

// Find-my-table — each guest's personal QR code (printed on their place
// card / invitation) links straight to /table/:id, so scanning it shows
// their table number and the day's program with no typing at all. A plain
// /table page (name search) is kept as a manual fallback.
app.get("/table", (req, res) => {
  res.sendFile(path.join(__dirname, "table.html"));
});
app.get("/table/:id", (req, res) => {
  res.sendFile(path.join(__dirname, "table.html"));
});

function findSeatingMatches(rawName) {
  const q = normName(rawName);
  if (!q) return { exact: null, candidates: [] };
  const list = readSeating();
  const exactMatches = list.filter((s) => normName(s.name) === q);
  if (exactMatches.length === 1) return { exact: exactMatches[0], candidates: [] };
  const partial = list.filter((s) => normName(s.name).includes(q));
  if (partial.length === 1) return { exact: partial[0], candidates: [] };
  return { exact: null, candidates: partial.map((s) => s.name) };
}

app.get("/api/table", (req, res) => {
  const { exact, candidates } = findSeatingMatches(req.query.name);
  if (exact) {
    return res.json({ found: true, name: exact.name, table: exact.table || null });
  }
  if (candidates.length > 1) {
    return res.json({ found: false, multiple: candidates.slice(0, 8) });
  }
  return res.json({ found: false, multiple: [] });
});

// Personal lookup by the id embedded in a guest's own QR code.
app.get("/api/table/:id", (req, res) => {
  const list = readSeating();
  const entry = list.find((s) => s.id === req.params.id);
  if (!entry) return res.status(404).json({ found: false });
  res.json({ found: true, name: entry.name, table: entry.table || null });
});

// The day's run-of-show — shown alongside the table lookup. Editable from
// /admin so Annie can update it herself without touching any code.
app.get("/api/program", (req, res) => {
  res.json({ program: readProgram() });
});

// Unlisted admin view — not linked from the public page. Shows RSVP data
// only (name, attending status, guest count, table, message, timestamp).
app.get("/admin", (req, res) => {
  res.sendFile(path.join(__dirname, "admin.html"));
});

app.get("/api/admin/rsvps", (req, res) => {
  const entries = readAll();
  const seating = readSeating();
  const attending = entries.filter((e) => e.attending === "yes");
  const declined = entries.filter((e) => e.attending === "no");
  const totalGuests = attending.reduce((sum, e) => sum + (e.guests || 1), 0);

  res.json({
    totalResponses: entries.length,
    attendingCount: attending.length,
    totalGuests,
    declinedCount: declined.length,
    rsvps: entries
      .slice()
      .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
      .map((e) => {
        const seat = seating.find((s) => normName(s.name) === normName(e.name));
        return {
          id: e.id,
          name: e.name,
          attending: e.attending,
          guests: e.guests,
          message: e.message,
          table: seat ? seat.table || null : null,
          submittedAt: e.submittedAt,
        };
      }),
  });
});

// --- Seating list (name -> table) -------------------------------------
// This is the source of truth for table numbers and per-guest QR codes.
// It's independent of the RSVP responses, since Annie's master seating
// list (uploaded from Excel) may include guests who haven't RSVP'd yet.
app.get("/api/admin/seating", (req, res) => {
  const list = readSeating()
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name));
  res.json({ seating: list });
});

app.post("/api/admin/seating", (req, res) => {
  const body = req.body || {};
  const name = String(body.name || "").trim();
  if (!name) return res.status(400).json({ error: "Please enter a name." });

  const list = readSeating();
  const { entry, created } = upsertSeating(list, name, body.table);
  writeSeating(list);
  res.status(created ? 201 : 200).json({ ok: true, entry, created });
});

app.delete("/api/admin/seating/:id", (req, res) => {
  const list = readSeating();
  const next = list.filter((s) => s.id !== req.params.id);
  if (next.length === list.length) return res.status(404).json({ error: "Not found." });
  writeSeating(next);
  res.json({ ok: true });
});

// Bulk-import from an Excel file: columns "Name" / "Full Name" and
// "Table" / "Table Number" (header matching is case-insensitive). Existing
// names are matched case-insensitively and updated in place; everyone
// else is added. Nothing is ever deleted by an upload.
app.post("/api/admin/seating/upload", (req, res) => {
  const body = req.body || {};
  const base64 = String(body.data || "");
  if (!base64) return res.status(400).json({ error: "No file received." });

  let rows;
  try {
    const buf = Buffer.from(base64, "base64");
    const workbook = XLSX.read(buf, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  } catch (err) {
    console.error("Seating upload parse failed:", err);
    return res.status(400).json({ error: "Could not read that file. Please upload a .xlsx or .xls file." });
  }

  if (!rows.length) return res.status(400).json({ error: "That file looks empty." });

  const headerKeys = Object.keys(rows[0]);
  const nameKey = headerKeys.find((k) => /^(full\s*name|guest\s*name|name)$/i.test(k.trim()))
    || headerKeys.find((k) => /name/i.test(k));
  const tableKey = headerKeys.find((k) => /^(table\s*(no\.?|number|#)?|seat(ing)?)$/i.test(k.trim()))
    || headerKeys.find((k) => /table|seat/i.test(k));

  if (!nameKey) {
    return res.status(400).json({ error: "Could not find a Name column. Please include a column named \"Name\"." });
  }

  const list = readSeating();
  let created = 0, updated = 0, skipped = 0;
  rows.forEach((row) => {
    const name = String(row[nameKey] || "").trim();
    const table = tableKey ? String(row[tableKey] || "").trim() : "";
    if (!name) { skipped++; return; }
    const { created: wasCreated } = upsertSeating(list, name, table);
    if (wasCreated) created++; else updated++;
  });
  writeSeating(list);

  res.json({ ok: true, created, updated, skipped, total: list.length });
});

// --- Wedding Updates ------------------------------------------------------
// Simple announcement board so guests (incl. less tech-savvy family) can
// check what's new without navigating anything complicated.
app.get("/api/updates", (req, res) => {
  const updates = readUpdates()
    .slice()
    .sort((a, b) => new Date(b.postedAt) - new Date(a.postedAt));
  res.json({ updates });
});

app.post("/api/admin/updates", (req, res) => {
  const body = req.body || {};
  const message = String(body.message || "").trim().slice(0, 500);
  if (!message) return res.status(400).json({ error: "Please write an update first." });

  const updates = readUpdates();
  const entry = {
    id: crypto.randomUUID(),
    message,
    postedAt: new Date().toISOString(),
  };
  updates.push(entry);
  writeUpdates(updates);
  res.status(201).json({ ok: true, entry });
});

app.delete("/api/admin/updates/:id", (req, res) => {
  const updates = readUpdates();
  const next = updates.filter((u) => u.id !== req.params.id);
  if (next.length === updates.length) return res.status(404).json({ error: "Update not found." });
  writeUpdates(next);
  res.json({ ok: true });
});

// --- Wedding Program (run-of-show) -----------------------------------------
app.post("/api/admin/program", (req, res) => {
  const body = req.body || {};
  const time = String(body.time || "").trim().slice(0, 40);
  const title = String(body.title || "").trim().slice(0, 120);
  if (!title) return res.status(400).json({ error: "Please add a program item first." });

  const program = readProgram();
  const entry = { id: crypto.randomUUID(), time, title };
  program.push(entry);
  writeProgram(program);
  res.status(201).json({ ok: true, entry });
});

app.delete("/api/admin/program/:id", (req, res) => {
  const program = readProgram();
  const next = program.filter((p) => p.id !== req.params.id);
  if (next.length === program.length) return res.status(404).json({ error: "Program item not found." });
  writeProgram(next);
  res.json({ ok: true });
});

// A guest's personal QR code — printed on their place card / invitation.
// Scanning it opens /table/:id directly, no typing required.
app.get("/api/admin/qr/:id", async (req, res) => {
  const list = readSeating();
  const entry = list.find((s) => s.id === req.params.id);
  if (!entry) return res.status(404).send("Guest not found.");

  const base = `${req.protocol}://${req.get("host")}`;
  const url = `${base}/table/${entry.id}`;
  try {
    const svg = await QRCode.toString(url, {
      type: "svg",
      margin: 1,
      color: { dark: "#2b211a", light: "#fffcf6" },
    });
    res.type("image/svg+xml").send(svg);
  } catch (err) {
    console.error("QR generation failed:", err);
    res.status(500).send("Could not generate QR code.");
  }
});

app.get("/healthz", (req, res) => res.json({ ok: true }));

app.listen(PORT, () => {
  console.log(`AJ Close The Gap wedding site listening on port ${PORT}`);
});
