const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const QRCode = require("qrcode");
const XLSX = require("xlsx");
const archiver = require("archiver");

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
const SETTINGS_FILE = path.join(DATA_DIR, "settings.json");
const UPLOADS_DIR = path.join(DATA_DIR, "uploads");

// --- tiny JSON-file "database" -------------------------------------------
function ensureFile(file, defaultContent) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(file)) fs.writeFileSync(file, defaultContent, "utf8");
}

function readJson(file) {
  ensureFile(file, "[]");
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
  ensureFile(file, "[]");
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(entries, null, 2), "utf8");
  fs.renameSync(tmp, file);
}

// Same idea as readJson/writeJson, but for a single settings object rather
// than a list — used for small site-wide config like the shared photo
// album link, which Annie edits from /admin without needing a redeploy.
function readSettingsObj() {
  ensureFile(SETTINGS_FILE, "{}");
  try {
    const raw = fs.readFileSync(SETTINGS_FILE, "utf8");
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch (err) {
    console.error(`Failed to read ${SETTINGS_FILE}, starting fresh:`, err);
    return {};
  }
}

function writeSettingsObj(obj) {
  ensureFile(SETTINGS_FILE, "{}");
  const tmp = SETTINGS_FILE + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(obj, null, 2), "utf8");
  fs.renameSync(tmp, SETTINGS_FILE);
}

function readAll() { return readJson(DATA_FILE); }
function writeAll(entries) { writeJson(DATA_FILE, entries); }
function readUpdates() { return readJson(UPDATES_FILE); }
function writeUpdates(entries) { writeJson(UPDATES_FILE, entries); }
function readProgram() { return readJson(PROGRAM_FILE); }
function writeProgram(entries) { writeJson(PROGRAM_FILE, entries); }
function readSeating() { return readJson(SEATING_FILE); }
function writeSeating(entries) { writeJson(SEATING_FILE, entries); }

ensureFile(DATA_FILE, "[]");
ensureFile(UPDATES_FILE, "[]");
ensureFile(PROGRAM_FILE, "[]");
ensureFile(SEATING_FILE, "[]");
ensureFile(SETTINGS_FILE, "{}");

function normName(n) {
  return String(n || "").trim().toLowerCase().replace(/\s+/g, " ");
}

// The wedding day itself, in the venue's own timezone — guests can only
// self check-in on this calendar date (Asia/Manila), regardless of what
// timezone their phone thinks it's in. Also the anchor for the rest of the
// site's lifecycle: photo uploads close a week later, and the whole guest
// side of the site quietly retires into a read-only "memory" page a month
// after that. All of these are calendar-date comparisons in Asia/Manila, not
// the visitor's own device clock, so they land on the same real-world day
// for every guest regardless of timezone.
const WEDDING_DATE_MANILA = "2027-02-21";
const PHOTO_UPLOAD_CUTOFF_MANILA = "2027-03-01"; // uploading new photos closes ON this date
const SITE_LOCK_DATE_MANILA = "2027-03-21"; // Find My Table etc. close the day AFTER this date

function todayInManila() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Manila",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}
function isWeddingDayInManila() {
  return todayInManila() === WEDDING_DATE_MANILA;
}
function isPhotoUploadClosed() {
  return todayInManila() >= PHOTO_UPLOAD_CUTOFF_MANILA;
}
function isSiteLocked() {
  return todayInManila() > SITE_LOCK_DATE_MANILA;
}
// The RSVP deadline is admin-editable (a plain YYYY-MM-DD in site settings),
// not hardcoded, since Annie may need to move it. Guests can still respond
// through the deadline date itself; it closes starting the day after. With
// no deadline set yet, RSVP just stays open.
function isRsvpClosed(settings) {
  const deadline = String((settings && settings.rsvpDeadline) || "").trim();
  if (!deadline) return false;
  return todayInManila() > deadline;
}

// Add or update one seating entry by name (case-insensitive match). This is
// the ONLY way a name gets onto the guest list — via Excel upload or an
// admin adding/editing a row here — never automatically from the public
// RSVP form, so the wedding stays invite-only.
//
// rawMaxGuests follows the same "blank means leave it alone" rule as an
// Excel re-upload that doesn't include a Guests column: pass undefined/""
// to preserve whatever is already set (defaulting a brand-new entry to 1),
// or a number to set it explicitly.
function upsertSeating(list, rawName, rawTable, rawMaxGuests, rawNickname) {
  const name = String(rawName || "").trim().slice(0, 80);
  const table = String(rawTable || "").trim().slice(0, 20);
  const key = normName(name);
  if (!key) return { entry: null, created: false };

  let maxGuests; // undefined = leave as-is / default to 1 for a new entry
  if (rawMaxGuests !== undefined && rawMaxGuests !== null && String(rawMaxGuests).trim() !== "") {
    const parsed = parseInt(rawMaxGuests, 10);
    if (Number.isFinite(parsed) && parsed >= 1) maxGuests = Math.min(parsed, 20);
  }

  // Nickname follows the same "blank means leave it alone" rule as
  // maxGuests — a re-upload or edit that doesn't mention a nickname
  // shouldn't erase one that's already saved.
  let nickname; // undefined = leave as-is
  if (rawNickname !== undefined && rawNickname !== null && String(rawNickname).trim() !== "") {
    nickname = String(rawNickname).trim().slice(0, 60);
  }

  const idx = list.findIndex((s) => normName(s.name) === key);
  if (idx === -1) {
    const entry = { id: crypto.randomUUID(), name, table: table || null, maxGuests: maxGuests || 1, nickname: nickname || null };
    list.push(entry);
    return { entry, created: true };
  }
  list[idx].table = table || null;
  if (name) list[idx].name = name;
  if (maxGuests !== undefined) list[idx].maxGuests = maxGuests;
  else if (!list[idx].maxGuests) list[idx].maxGuests = 1;
  if (nickname !== undefined) list[idx].nickname = nickname;
  return { entry: list[idx], created: false };
}

// Some guest lists mark a plus-one as its own row instead of using a
// "Guests" column — e.g. one row "Kyle Josef" and a second row
// "Kyle Josef +1" (nickname "Kyle Partner") meaning "Kyle Josef, plus a
// companion not named yet." This upserts BOTH: it bumps Kyle's own
// headcount (so his RSVP form offers "2 guests" and a second name field),
// AND keeps the "+1" row as its own distinct, findable guest-list entry —
// so that placeholder can actually be selected/typed as Kyle's second
// guest when he RSVPs. Kyle's existing table/nickname are left alone
// unless this row's own cells actually say otherwise. A row without a
// trailing "+N" passes straight through to upsertSeating unchanged.
function upsertSeatingRow(list, rawName, rawTable, rawMaxGuests, rawNickname) {
  const raw = String(rawName || "").trim();
  const plusOneMatch = raw.match(/^(.*?)\s*\(?\+\s*(\d+)\)?\s*$/);
  const baseName = plusOneMatch ? plusOneMatch[1].trim() : "";
  if (!plusOneMatch || !baseName) {
    return upsertSeating(list, rawName, rawTable, rawMaxGuests, rawNickname);
  }

  const extra = parseInt(plusOneMatch[2], 10);
  const existingBase = list.find((s) => normName(s.name) === normName(baseName));
  const desiredMax =
    rawMaxGuests !== undefined && rawMaxGuests !== null && String(rawMaxGuests).trim() !== ""
      ? rawMaxGuests
      : Math.max((existingBase && existingBase.maxGuests) || 1, 1 + extra);
  const tableTrimmed = String(rawTable || "").trim();
  const baseTable = tableTrimmed ? rawTable : existingBase ? existingBase.table : rawTable;

  upsertSeating(list, baseName, baseTable, desiredMax, undefined);
  // The placeholder keeps the "+1" in its own name so it never collides
  // with the named guest above, and always brings exactly itself (1).
  return upsertSeating(list, raw, rawTable, 1, rawNickname);
}

// A guest may type either their full name or their nickname — this looks
// a typed name up against both, so "Lodi" finds the same guest-list entry
// as "Lodivico Cruz Josef" would. Full-name matches win over nickname
// matches when (rarely) both could apply.
function findSeatByNameOrNickname(list, rawName) {
  const key = normName(rawName);
  if (!key) return null;
  return (
    list.find((s) => normName(s.name) === key) ||
    list.find((s) => s.nickname && normName(s.nickname) === key) ||
    null
  );
}

async function generateQrSvg(url) {
  return QRCode.toString(url, {
    type: "svg",
    // A wider quiet zone (the QR spec recommends 4 modules) makes the code
    // decode reliably from a photo or screenshot upload, not just a live
    // camera scan — jsQR (used by the "upload your QR code" feature) is
    // fussier about this than most phone camera scanners.
    margin: 4,
    color: { dark: "#2b211a", light: "#fffcf6" },
  });
}

// --- middleware ------------------------------------------------------------
// 15mb covers the small text payloads everywhere else plus a base64-encoded
// seating spreadsheet or wedding-update photo from /admin.
app.use(express.json({ limit: "15mb" }));

// Photos attached to Wedding Updates are saved as real files on the
// persistent data volume (not stuffed into the JSON file) and served back
// out from here.
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
app.use("/uploads", express.static(UPLOADS_DIR));

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

// Guest-facing pages get edited and redeployed often, and browsers (mobile
// Safari especially) have repeatedly kept showing a stale cached copy after
// a deploy, confusing "did my change actually go live?" testing. Sending
// these with no-store means every visit always fetches the current file.
function sendHtmlNoCache(res, filename) {
  res.set("Cache-Control", "no-store");
  res.sendFile(path.join(__dirname, filename));
}

// The page lives at the repo root (index.html next to this file).
app.get("/", (req, res) => {
  sendHtmlNoCache(res, "index.html");
});

// The RSVP form now lives on its own page, linked from a big "RSVP Now"
// button on the homepage, so the homepage itself can stay short and simple.
app.get("/rsvp", (req, res) => {
  sendHtmlNoCache(res, "rsvp.html");
});

// Vendored locally (not loaded from a CDN) so the "upload your QR code
// image" feature on /table keeps working even on a venue's flaky WiFi or a
// guest's phone that blocks third-party scripts.
app.get("/jsQR.js", (req, res) => {
  res.sendFile(path.join(__dirname, "jsQR.js"));
});

// --- API ---------------------------------------------------------------
app.post("/api/rsvp", (req, res) => {
  if (isRsvpClosed(readSettingsObj())) {
    return res.status(403).json({ error: "Thank you for responding! RSVP is closed — message us if anything's changed." });
  }

  const body = req.body || {};
  const name = String(body.name || "").trim().slice(0, 80);
  const attending = body.attending === "yes" ? "yes" : body.attending === "no" ? "no" : null;
  const message = String(body.message || "").trim().slice(0, 300);

  if (!name) return res.status(400).json({ error: "Please enter your name." });
  if (!attending) return res.status(400).json({ error: "Please let us know if you can make it." });

  // The wedding is invite-only: an RSVP can only be submitted for a name
  // that's already on Annie & Jay's guest list (uploaded via Excel, or
  // added by hand in the admin seating list) — nobody can add themselves,
  // or anyone else, to the list just by filling out this form.
  const seatingList = readSeating();
  const seatEntry = findSeatByNameOrNickname(seatingList, name);
  if (!seatEntry) {
    return res.status(404).json({
      error: `We couldn't find "${name}" on our guest list. Please check the spelling, or reach out to Annie & Jay if you think this is a mistake.`,
    });
  }

  const maxGuests = seatEntry.maxGuests || 1;
  let guests = parseInt(body.guests, 10);
  if (attending === "yes") {
    if (!Number.isFinite(guests) || guests < 1) guests = 1;
    if (guests > maxGuests) {
      return res.status(400).json({
        error: `Your invitation allows up to ${maxGuests} ${maxGuests === 1 ? "guest" : "guests"}. Please reach out to Annie & Jay if you need to bring more.`,
      });
    }
  } else {
    guests = 0;
  }

  // Optional named additional guests in the party (e.g. "and my husband and
  // two kids") — each one must ALSO already be on the guest list, same as
  // the primary respondent. Capped to the headcount minus the person
  // filling out the form, regardless of what the client sends.
  const rawPartyNames = Array.isArray(body.partyNames) ? body.partyNames : [];
  const maxParty = attending === "yes" ? Math.max(0, guests - 1) : 0;
  const candidateNames = rawPartyNames
    .map((n) => String(n || "").trim().slice(0, 80))
    .filter(Boolean)
    .slice(0, maxParty);

  const partyNames = [];
  for (const candidate of candidateNames) {
    const candidateKey = normName(candidate);
    if (candidateKey === normName(seatEntry.name)) continue; // skip an accidental self-duplicate
    if (seatEntry.nickname && candidateKey === normName(seatEntry.nickname)) continue;
    const found = findSeatByNameOrNickname(seatingList, candidate);
    if (!found) {
      return res.status(404).json({
        error: `We couldn't find "${candidate}" on our guest list. Please check the spelling, or reach out to Annie & Jay if you think this is a mistake.`,
      });
    }
    partyNames.push(found.name);
  }

  const entries = readAll();
  const entry = {
    id: crypto.randomUUID(),
    name: seatEntry.name, // canonical spelling from the guest list
    attending,
    guests,
    partyNames,
    message,
    seatingId: seatEntry.id,
    submittedAt: new Date().toISOString(),
  };
  entries.push(entry);
  writeAll(entries);

  res.status(201).json({ ok: true, entry, seatingId: seatEntry.id });
});

// Exact-match lookup used by the RSVP form: confirms a typed name is
// actually on the guest list, and reports how many guests that invitation
// is allowed to bring (so the form can cap the dropdown before submitting).
app.get("/api/guest-lookup", (req, res) => {
  const list = readSeating();
  const entry = findSeatByNameOrNickname(list, req.query.name);
  if (!entry) return res.json({ found: false });
  res.json({ found: true, name: entry.name, maxGuests: entry.maxGuests || 1 });
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
  sendHtmlNoCache(res, "table.html");
});
app.get("/table/:id", (req, res) => {
  sendHtmlNoCache(res, "table.html");
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
  const exactMatches = list.filter((s) => normName(s.name) === q || (s.nickname && normName(s.nickname) === q));
  if (exactMatches.length === 1) return { exact: exactMatches[0], candidates: [] };
  const partial = list.filter((s) => normName(s.name).includes(q) || (s.nickname && normName(s.nickname).includes(q)));
  if (partial.length === 1) return { exact: partial[0], candidates: [] };
  return { exact: null, candidates: partial.map((s) => s.name) };
}

// A guest's QR code (or name search) shows their WHOLE party together —
// themselves plus anyone they named as additional guests on their RSVP —
// each with their own table number and check-in state, all from one scan.
// If this person isn't part of any RSVP's party (e.g. an admin-added
// seating entry nobody has RSVP'd for yet), the "party" is just themselves.
function getPartyFor(seatingEntry) {
  const seatingList = readSeating();
  const rsvps = readAll();
  const key = normName(seatingEntry.name);

  const owningRsvp = rsvps.find((e) => {
    if (e.attending !== "yes") return false;
    if (normName(e.name) === key) return true;
    return Array.isArray(e.partyNames) && e.partyNames.some((n) => normName(n) === key);
  });

  const names = owningRsvp
    ? [owningRsvp.name].concat(Array.isArray(owningRsvp.partyNames) ? owningRsvp.partyNames : [])
    : [seatingEntry.name];

  const seen = {};
  const party = [];
  names.forEach((n) => {
    const k = normName(n);
    if (seen[k]) return;
    seen[k] = true;
    const seat = seatingList.find((s) => normName(s.name) === k);
    if (!seat) return;
    party.push({
      id: seat.id,
      name: seat.name,
      table: seat.table || null,
      checkedIn: !!seat.checkedIn,
      checkedInAt: seat.checkedInAt || null,
    });
  });
  return party;
}

app.get("/api/table", (req, res) => {
  if (isSiteLocked()) return res.json({ found: false, locked: true });
  const { exact, candidates } = findSeatingMatches(req.query.name);
  if (exact) {
    return res.json({
      found: true,
      id: exact.id,
      name: exact.name,
      table: exact.table || null,
      checkedIn: !!exact.checkedIn,
      checkedInAt: exact.checkedInAt || null,
      party: getPartyFor(exact),
    });
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
  // Matching also checks nickname, but the suggestion shown (and filled in)
  // is always the Full Name — so typing "Lodi" surfaces "Lodivico Cruz
  // Josef" in the dropdown, not the nickname itself.
  list.forEach((s) => {
    const n = normName(s.name);
    const nick = s.nickname ? normName(s.nickname) : "";
    const startsMatch = n.startsWith(q) || (nick && nick.startsWith(q));
    const containsMatch = n.includes(q) || (nick && nick.includes(q));
    if (startsMatch) startsWith.push(s.name);
    else if (containsMatch) contains.push(s.name);
  });
  res.json({ names: startsWith.concat(contains).slice(0, 8) });
});

// Personal lookup by the id embedded in a guest's own QR code. One QR code
// (the primary RSVP respondent's) is enough for the whole party — the
// response includes everyone in it via getPartyFor().
app.get("/api/table/:id", (req, res) => {
  if (isSiteLocked()) return res.json({ found: false, locked: true });
  const list = readSeating();
  const entry = list.find((s) => s.id === req.params.id);
  if (!entry) return res.status(404).json({ found: false });
  res.json({
    found: true,
    id: entry.id,
    name: entry.name,
    table: entry.table || null,
    checkedIn: !!entry.checkedIn,
    checkedInAt: entry.checkedInAt || null,
    party: getPartyFor(entry),
  });
});

// Self check-in: a guest taps this after scanning their QR code (or finding
// their table by name) to let Annie know they've arrived. Only allowed on
// the wedding day itself so it can't be tapped early by accident — checked
// server-side too, not just hidden in the UI, in case someone hits the API
// directly.
app.post("/api/checkin/:id", (req, res) => {
  if (!isWeddingDayInManila()) {
    return res.status(403).json({
      error: "Check-in opens on the day of the wedding — February 21, 2027.",
      checkedIn: false,
    });
  }
  const list = readSeating();
  const entry = list.find((s) => s.id === req.params.id);
  if (!entry) return res.status(404).json({ error: "Guest not found.", checkedIn: false });

  if (!entry.checkedIn) {
    entry.checkedIn = true;
    entry.checkedInAt = new Date().toISOString();
    writeSeating(list);
  }
  res.json({ ok: true, checkedIn: true, checkedInAt: entry.checkedInAt, name: entry.name });
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
    return sendHtmlNoCache(res, "admin-login.html");
  }
  sendHtmlNoCache(res, "admin.html");
});

// A simple, big-text, search-only view for ushers/coordinators at the door
// on the wedding day — same password as the rest of admin, no editing tools.
app.get("/admin/checkin", (req, res) => {
  if (!isAdminAuthed(req)) {
    return sendHtmlNoCache(res, "admin-login.html");
  }
  sendHtmlNoCache(res, "admin-checkin.html");
});

// A printable sheet of every guest's personal QR code (name + QR + table),
// for printing onto invitations or using as table/escort cards.
app.get("/admin/qr-print", (req, res) => {
  if (!isAdminAuthed(req)) {
    return sendHtmlNoCache(res, "admin-login.html");
  }
  sendHtmlNoCache(res, "admin-qr-print.html");
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
          partyNames: Array.isArray(e.partyNames) ? e.partyNames : [],
          message: e.message,
          seatingId: seat ? seat.id : null,
          table: seat ? seat.table || null : null,
          checkedIn: seat ? !!seat.checkedIn : false,
          checkedInAt: seat ? seat.checkedInAt || null : null,
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
  const { entry, created } = upsertSeatingRow(list, name, body.table, body.maxGuests, body.nickname);
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

// Lets Annie mark a batch of already-listed guests as "attending" in one
// click, instead of having to open the public RSVP page and submit it
// once per person herself (e.g. for family she's already confirmed with
// by phone). Each selected guest gets their OWN real RSVP record under
// their own name — this is not a single combined entry, so headcounts,
// the Guest List, and each person's boarding pass all stay accurate.
// A guest who already has an RSVP is left alone (never duplicated); one
// who previously declined is flipped to attending. "Guests Allowed" on
// the seating entry still caps what a guest could add for themselves —
// this only ever registers the ONE named person per row, at 1 guest each.
app.post("/api/admin/rsvp/bulk-attend", (req, res) => {
  const body = req.body || {};
  const ids = Array.isArray(body.ids) ? body.ids : [];
  if (!ids.length) return res.status(400).json({ error: "No guests selected." });

  const seatingList = readSeating();
  const rsvps = readAll();
  let attended = 0, alreadyRsvpd = 0, notFound = 0;
  const results = [];

  ids.forEach((id) => {
    const seat = seatingList.find((s) => s.id === id);
    if (!seat) { notFound++; return; }

    const key = normName(seat.name);
    const existing = rsvps.find((e) => {
      if (normName(e.name) === key) return true;
      return Array.isArray(e.partyNames) && e.partyNames.some((n) => normName(n) === key);
    });

    if (existing) {
      if (existing.attending !== "yes") existing.attending = "yes";
      alreadyRsvpd++;
      results.push({ id, name: seat.name, status: "already_rsvpd" });
      return;
    }

    rsvps.push({
      id: crypto.randomUUID(),
      name: seat.name,
      attending: "yes",
      guests: 1,
      partyNames: [],
      message: "",
      seatingId: seat.id,
      submittedAt: new Date().toISOString(),
      recordedByAdmin: true,
    });
    attended++;
    results.push({ id, name: seat.name, status: "attending" });
  });

  writeAll(rsvps);
  res.json({ ok: true, attended, alreadyRsvpd, notFound, results });
});

// Bulk-import from an Excel file: columns "Name" / "Full Name", "Table" /
// "Table Number", and an optional "Guests" / "Max Guests" / "Party Size"
// column (header matching is case-insensitive) that sets how many people
// that invitation may bring — this is the ONLY place that count comes
// from; if the column is left out for a row, an existing entry's number
// stays as it was, and a brand-new one defaults to 1. Existing names are
// matched case-insensitively and updated in place; everyone else is
// added. Nothing is ever deleted by an upload.
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
  const maxGuestsKey = headerKeys.find((k) => /^(max\s*guests?|guests?\s*(allowed|allotted)?|party\s*size|no\.?\s*of\s*guests|number\s*of\s*guests|total\s*guests?)$/i.test(k.trim()))
    || headerKeys.find((k) => /guest/i.test(k));
  const nicknameKey = headerKeys.find((k) => /^nick\s*name$/i.test(k.trim()))
    || headerKeys.find((k) => /nick/i.test(k));

  if (!nameKey) {
    return res.status(400).json({ error: "Could not find a Name column. Please include a column named \"Name\"." });
  }

  const list = readSeating();
  let created = 0, updated = 0, skipped = 0;
  rows.forEach((row) => {
    const name = String(row[nameKey] || "").trim();
    const table = tableKey ? String(row[tableKey] || "").trim() : "";
    const maxGuestsRaw = maxGuestsKey ? row[maxGuestsKey] : undefined;
    const nicknameRaw = nicknameKey ? row[nicknameKey] : undefined;
    if (!name) { skipped++; return; }
    const { created: wasCreated } = upsertSeatingRow(list, name, table, maxGuestsRaw, nicknameRaw);
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

const IMAGE_MIME_EXT = {
  "image/jpeg": "jpg",
  "image/png": "png",
  "image/webp": "webp",
  "image/gif": "gif",
};
const MAX_IMAGE_BYTES = 8 * 1024 * 1024; // 8MB decoded

// Accepts a data URL ("data:image/jpeg;base64,...") or a bare base64 string
// plus a separate mime type. Saves the decoded bytes as a real file on the
// data volume and returns its public /uploads/... URL, or null if nothing
// usable was sent. Throws a descriptive Error for a bad/oversized image so
// the route can turn it into a clean 400.
function saveUpdateImage(rawImage, rawMime) {
  if (!rawImage) return null;
  let mime = String(rawMime || "").trim();
  let base64 = String(rawImage);
  const dataUrlMatch = base64.match(/^data:([^;]+);base64,(.*)$/s);
  if (dataUrlMatch) {
    mime = dataUrlMatch[1];
    base64 = dataUrlMatch[2];
  }
  const ext = IMAGE_MIME_EXT[mime];
  if (!ext) throw new Error("Please attach a JPEG, PNG, WEBP, or GIF image.");

  const buf = Buffer.from(base64, "base64");
  if (!buf.length) throw new Error("That image looks empty. Please try a different file.");
  if (buf.length > MAX_IMAGE_BYTES) throw new Error("That image is too large — please use one under 8MB.");

  const filename = `${crypto.randomUUID()}.${ext}`;
  fs.writeFileSync(path.join(UPLOADS_DIR, filename), buf);
  return `/uploads/${filename}`;
}

function deleteUpdateImage(imageUrl) {
  if (!imageUrl || !imageUrl.startsWith("/uploads/")) return;
  const filename = path.basename(imageUrl);
  const filePath = path.join(UPLOADS_DIR, filename);
  fs.unlink(filePath, () => {}); // best-effort; nothing to do if it's already gone
}

app.post("/api/admin/updates", (req, res) => {
  const body = req.body || {};
  const message = String(body.message || "").trim().slice(0, 500);
  if (!message) return res.status(400).json({ error: "Please write an update first." });

  let imageUrl = null;
  if (body.image) {
    try {
      imageUrl = saveUpdateImage(body.image, body.imageType);
    } catch (err) {
      return res.status(400).json({ error: err.message });
    }
  }

  const updates = readUpdates();
  const entry = {
    id: crypto.randomUUID(),
    message,
    imageUrl,
    postedAt: new Date().toISOString(),
  };
  updates.push(entry);
  writeUpdates(updates);
  res.status(201).json({ ok: true, entry });
});

app.delete("/api/admin/updates/:id", (req, res) => {
  const updates = readUpdates();
  const target = updates.find((u) => u.id === req.params.id);
  if (!target) return res.status(404).json({ error: "Update not found." });
  deleteUpdateImage(target.imageUrl);
  writeUpdates(updates.filter((u) => u.id !== req.params.id));
  res.json({ ok: true });
});

// --- Guest Photos ------------------------------------------------------
// Guests upload straight from the homepage (no Google Drive detour needed)
// and everyone sees a live gallery of the latest photos right there. Kept
// deliberately simple — no accounts, no approval queue to wait on — with
// the only moderation being Annie/Jay's ability to remove a photo from the
// admin panel if something shouldn't be up.
const GUEST_PHOTOS_FILE = path.join(DATA_DIR, "guest-photos.json");
function readGuestPhotos() { return readJson(GUEST_PHOTOS_FILE); }
function writeGuestPhotos(entries) { writeJson(GUEST_PHOTOS_FILE, entries); }
ensureFile(GUEST_PHOTOS_FILE, "[]");

app.post("/api/guest-photos", (req, res) => {
  if (isPhotoUploadClosed()) {
    return res.status(403).json({ error: "Thank you for sharing! Uploads are closed." });
  }

  const body = req.body || {};
  if (!body.image) return res.status(400).json({ error: "Please choose a photo to upload." });

  let imageUrl;
  try {
    imageUrl = saveUpdateImage(body.image, body.imageType);
  } catch (err) {
    return res.status(400).json({ error: err.message });
  }

  const name = String(body.name || "").trim().slice(0, 60);
  const photos = readGuestPhotos();
  const entry = { id: crypto.randomUUID(), imageUrl, name, uploadedAt: new Date().toISOString() };
  photos.push(entry);
  writeGuestPhotos(photos);
  res.status(201).json({ ok: true, photo: entry });
});

app.get("/api/guest-photos", (req, res) => {
  const limit = Math.min(200, Math.max(1, parseInt(req.query.limit, 10) || 24));
  const photos = readGuestPhotos()
    .slice()
    .sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt))
    .slice(0, limit);
  res.json({ photos });
});

app.delete("/api/admin/guest-photos/:id", (req, res) => {
  const photos = readGuestPhotos();
  const target = photos.find((p) => p.id === req.params.id);
  if (!target) return res.status(404).json({ error: "Photo not found." });
  deleteUpdateImage(target.imageUrl);
  writeGuestPhotos(photos.filter((p) => p.id !== req.params.id));
  res.json({ ok: true });
});

// Builds a friendly, collision-free filename for one guest photo, e.g.
// "2026-09-12-Tita Baby.jpg" — shared by the single-file and ZIP downloads
// below so a photo has the same name either way.
function friendlyPhotoName(photo, usedNames) {
  const filePath = path.join(UPLOADS_DIR, path.basename(photo.imageUrl || ""));
  const ext = path.extname(filePath) || ".jpg";
  const dateStamp = (photo.uploadedAt || "").slice(0, 10) || "undated";
  const who = String(photo.name || "guest").trim().replace(/[^a-z0-9 _-]/gi, "").slice(0, 40) || "guest";
  const base = `${dateStamp}-${who}`;
  let name = `${base}${ext}`;
  if (usedNames) {
    let n = 2;
    while (usedNames.has(name.toLowerCase())) {
      name = `${base}-${n}${ext}`;
      n++;
    }
    usedNames.add(name.toLowerCase());
  }
  return { filePath, name };
}

// One-click download of every guest photo (or, with ?ids=a,b,c, just a
// chosen selection) as a single ZIP file, so Annie can grab them in bulk and
// drop them into Google Drive (or anywhere else) herself — no Google account
// setup or ongoing integration needed.
app.get("/api/admin/guest-photos/export", (req, res) => {
  let photos = readGuestPhotos();
  const rawIds = String(req.query.ids || "").trim();
  if (rawIds) {
    const wanted = new Set(rawIds.split(",").map((s) => s.trim()).filter(Boolean));
    photos = photos.filter((p) => wanted.has(p.id));
    if (!photos.length) {
      return res.status(404).json({ error: "Please select at least one photo to download." });
    }
  } else if (!photos.length) {
    return res.status(404).json({ error: "There are no guest photos to export yet." });
  }

  const zipName = rawIds
    ? `aj-guest-photos-selected-${new Date().toISOString().slice(0, 10)}.zip`
    : `aj-guest-photos-${new Date().toISOString().slice(0, 10)}.zip`;
  res.attachment(zipName);
  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.on("error", (err) => {
    console.error("Guest photo export failed:", err);
    if (!res.headersSent) res.status(500);
    res.end();
  });
  archive.pipe(res);

  const usedNames = new Set();
  photos.forEach((p) => {
    const { filePath, name } = friendlyPhotoName(p, usedNames);
    if (!fs.existsSync(filePath)) return; // skip any record whose file is missing
    archive.file(filePath, { name });
  });

  archive.finalize();
});

// Single-photo download for the admin panel — lets Annie grab just one
// guest's photo without downloading the whole ZIP.
app.get("/api/admin/guest-photos/:id/download", (req, res) => {
  const photos = readGuestPhotos();
  const photo = photos.find((p) => p.id === req.params.id);
  if (!photo) return res.status(404).json({ error: "Photo not found." });
  const { filePath, name } = friendlyPhotoName(photo);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "That photo's file is missing." });
  res.download(filePath, name);
});

// Same single-photo download, but public (no admin login) — lets a guest
// save someone's shared photo straight from the live gallery on the
// invitation page, one at a time, without needing to right-click/long-press.
app.get("/api/guest-photos/:id/download", (req, res) => {
  const photos = readGuestPhotos();
  const photo = photos.find((p) => p.id === req.params.id);
  if (!photo) return res.status(404).json({ error: "Photo not found." });
  const { filePath, name } = friendlyPhotoName(photo);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "That photo's file is missing." });
  res.download(filePath, name);
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

// --- Site settings (shared photo album link, RSVP deadline, etc.) ----------
// A small key/value config Annie can edit from /admin without a code change
// or redeploy. Public so the invitation page can show the album button and
// so every page can check the site's current lifecycle stage (open, RSVP
// closed, uploads closed, fully locked into a read-only memory page) off one
// server-computed source of truth, rather than trusting each visitor's own
// device clock.
app.get("/api/settings", (req, res) => {
  const settings = readSettingsObj();
  res.json({
    photoAlbumUrl: settings.photoAlbumUrl || null,
    rsvpDeadline: settings.rsvpDeadline || null,
    rsvpClosed: isRsvpClosed(settings),
    photoUploadClosed: isPhotoUploadClosed(),
    siteLocked: isSiteLocked(),
    weddingPassed: todayInManila() > WEDDING_DATE_MANILA,
  });
});

app.post("/api/admin/settings", (req, res) => {
  const body = req.body || {};
  const raw = String(body.photoAlbumUrl || "").trim().slice(0, 500);
  if (raw && !/^https?:\/\//i.test(raw)) {
    return res.status(400).json({ error: "Please enter a full link starting with https://" });
  }
  const rawDeadline = String(body.rsvpDeadline || "").trim();
  if (rawDeadline && !/^\d{4}-\d{2}-\d{2}$/.test(rawDeadline)) {
    return res.status(400).json({ error: "Please pick a valid RSVP deadline date." });
  }
  const settings = readSettingsObj();
  settings.photoAlbumUrl = raw || null;
  settings.rsvpDeadline = rawDeadline || null;
  writeSettingsObj(settings);
  res.json({ ok: true, photoAlbumUrl: settings.photoAlbumUrl, rsvpDeadline: settings.rsvpDeadline });
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

const server = app.listen(PORT, () => {
  console.log(`AJ Close The Gap wedding site listening on port ${PORT}`);
});

// Exit cleanly when Railway stops this container during a redeploy (it
// sends SIGTERM to the old version once the new one is up). Without this,
// the process just gets killed and the shutdown can look like a crash in
// logs/alerts, even though it's a normal, expected part of every deploy.
function shutdown(signal) {
  console.log(`Received ${signal}, shutting down gracefully.`);
  server.close(() => process.exit(0));
  // Safety net in case something keeps an open connection alive.
  setTimeout(() => process.exit(0), 5000).unref();
}
process.on("SIGTERM", () => shutdown("SIGTERM"));
process.on("SIGINT", () => shutdown("SIGINT"));
