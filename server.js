const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const QRCode = require("qrcode");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;
// One shared password for the admin area — set ADMIN_PASSWORD in Railway's
// Variables tab to change it later (no code change or GitHub upload needed).
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "AJCloseTheGap2027";
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
  return String(n || "").trim().toLowerCase().replace(/\s+/g, " ");
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

// Make sure a seating record (and therefore a stable id for a QR code)
// exists for this name, WITHOUT touching a table number that may already
// be there (e.g. from an Excel upload done ahead of time). Used right when
// a guest RSVPs "yes", so their personal QR code exists immediately.
function ensureSeating(list, rawName) {
  const name = String(rawName || "").trim().slice(0, 80);
  const key = normName(name);
  if (!key) return null;
  let entry = list.find((s) => normName(s.name) === key);
  if (!entry) {
    entry = { id: crypto.randomUUID(), name, table: null };
    list.push(entry);
  }
  return entry;
}

async function generateQrSvg(url) {
  return QRCode.toString(url, {
    type: "svg",
    margin: 1,
    color: { dark: "#2b211a", light: "#fffcf6" },
  });
}

// --- middleware ------------------------------------------------------------
// 8mb covers the small text payloads everywhere else plus a base64-encoded
// seating spreadsheet upload from /admin.
app.use(express.json({ limit: "8mb" }));

// One shared password gate for the admin page and everything under
// /api/admin — a plain on-page login form (not the browser's native Basic
// Auth prompt, which doesn't show up reliably in every browser/app). Guests
// never see this: /table, /api/rsvp, /api/updates, /api/program and
// /api/qr all stay open.
const ADMIN_TOKEN = crypto.createHash("sha256").update(ADMIN_PASSWORD + ":ajclosethegap-admin").digest("hex");

function parseCookies(req) {
  const header = req.headers.cookie || "";
  const out = {};
  header.split(";").forEach((pair) => {
    const idx = pair.indexOf("=");
    if (idx === -1) return;
    const k = pair.slice(0, idx).trim();
    const v = pair.slice(idx + 1).trim();
    if (k) out[k] = decodeURIComponent(v);
  });
  return out;
}

function isAdminAuthed(req) {
  return parseCookies(req).admin_token === ADMIN_TOKEN;
}

function requireAdminPassword(req, res, next) {
  if (isAdminAuthed(req)) return next();
  res.status(401).json({ error: "Please log in first." });
}

app.use("/api/admin", requireAdminPassword);

// Login/logout live outside the /api/admin prefix so they aren't gated by
// the middleware above.
app.post("/api/admin-login", (req, res) => {
  const body = req.body || {};
  const password = String(body.password || "");
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: "Incorrect password. Please try again." });
  }
  res.cookie("admin_token", ADMIN_TOKEN, {
    httpOnly: true,
    sameSite: "lax",
    maxAge: 1000 * 60 * 60 * 24 * 30,
  });
  res.json({ ok: true });
});

app.post("/api/admin-logout", (req, res) => {
  res.clearCookie("admin_token");
  res.json({ ok: true });
});

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

  // A guest who is coming gets a seating record (and so a stable id for
  // their personal QR code) right away — the table number itself may still
  // be blank until Annie assigns it, which the table page handles gracefully.
  let seatingId = null;
  if (attending === "yes") {
    const seatingList = readSeating();
    const seatEntry = ensureSeating(seatingList, name);
    writeSeating(seatingList);
    seatingId = seatEntry ? seatEntry.id : null;
  }

  const entry = {
    id: crypto.randomUUID(),
    name,
    attending,
    guests,
    message,
    seatingId,
    submittedAt: new Date().toISOString(),
  };
  entries.push(entry);
  writeAll(entries);

  res.status(201).json({ ok: true, entry, seatingId });
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

// Find-my-table — every guest who RSVPs "yes" gets a personal QR code right
// on the confirmation, so scanning it later shows their table number and
// the day's program with no typing at all. A plain /table page (name
// search) is kept as a manual fallback.
app.get("/table", (req, res) => {
  res.sendFile(path.join(__dirname, "table.html"));
});
app.get("/table/:id", (req, res) => {
  res.sendFile(path.join(__dirname, "table.html"));
});

// Public: a guest's own personal QR code, e.g. shown right after they RSVP.
app.get("/api/qr/:id", async (req, res) => {
  const list = readSeating();
  const entry = list.find((s) => s.id === req.params.id);
  if (!entry) return res.status(404).send("Guest not found.");

  const base = `${req.protocol}://${req.get("host")}`;
  const url = `${base}/table/${entry.id}`;
  try {
    const svg = await generateQrSvg(url);
    res.type("image/svg+xml").send(svg);
  } catch (err) {
    console.error("QR generation failed:", err);
    res.status(500).send("Could not generate QR code.");
  }
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

// Name suggestions for the RSVP form's autocomplete, sourced from the
// Excel seating list — helps guests type their name the same way it's
// spelled on the sheet, so their RSVP links to the right table. Purely a
// convenience: the field stays free-text, this is not a required match.
app.get("/api/seating-names", (req, res) => {
  const q = normName(req.query.q);
  if (!q || q.length < 2) return res.json({ names: [] });
  const list = readSeating();
  const startsWith = [];
  const contains = [];
  list.forEach((s) => {
    const n = normName(s.name);
    if (n.startsWith(q)) startsWith.push(s.name);
    else if (n.includes(q)) contains.push(s.name);
  });
  res.json({ names: startsWith.concat(contains).slice(0, 8) });
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

// Admin view, behind the shared password. Not logged in yet? Show the
// plain login page instead — same URL either way, nothing to remember.
app.get("/admin", (req, res) => {
  if (!isAdminAuthed(req)) {
    return res.sendFile(path.join(__dirname, "admin-login.html"));
  }
  res.sendFile(path.join(__dirname, "admin.html"));
});

// A simple, big-text, search-only view for ushers/coordinators at the door
// on the wedding day — same password as the rest of admin, no editing tools.
app.get("/admin/checkin", (req, res) => {
  if (!isAdminAuthed(req)) {
    return res.sendFile(path.join(__dirname, "admin-login.html"));
  }
  res.sendFile(path.join(__dirname, "admin-checkin.html"));
});

// A printable sheet of every guest's personal QR code (name + QR + table),
// for printing onto invitations or using as table/escort cards.
app.get("/admin/qr-print", (req, res) => {
  if (!isAdminAuthed(req)) {
    return res.sendFile(path.join(__dirname, "admin-login.html"));
  }
  res.sendFile(path.join(__dirname, "admin-qr-print.html"));
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
          seatingId: seat ? seat.id : null,
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

// Admin alias of /api/qr/:id (same QR, same no-auth model) — kept so
// existing "View / Print QR" links in /admin keep working.
app.get("/api/admin/qr/:id", async (req, res) => {
  const list = readSeating();
  const entry = list.find((s) => s.id === req.params.id);
  if (!entry) return res.status(404).send("Guest not found.");

  const base = `${req.protocol}://${req.get("host")}`;
  const url = `${base}/table/${entry.id}`;
  try {
    const svg = await generateQrSvg(url);
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
