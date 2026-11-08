/**
 * Contractors database: backend
 *
 * Lives inside the Google Sheet (Extensions > Apps Script). The Sheet is the
 * database; this script is the password gate between it and the website.
 *
 * Tabs:  People     (one row per person)
 *        Contracts  (one row per contract, linked to a person)
 *
 * Passwords are NOT stored in this file. Set them under
 * Project Settings > Script Properties:
 *   PASSWORD       (required) the team password: lets people view the list
 *   EDIT_PASSWORD  (optional) if set, only this password can add/edit/delete.
 *                  If left out, everyone with PASSWORD can edit.
 *
 * Email and phone columns stay in the Sheet. They are never sent to the website.
 *
 * In the Sheet, the "Contractors" menu has: Check data, Refresh the website's data, Set up.
 */

const DATE_FORMAT = "yyyy-mm-dd";   // how date cells are formatted (matches the workbook)

// [field key, heading, ...other headings that also work]. Columns are found by heading,
// so they can be reordered, and any extra columns you add are left alone.
const PEOPLE = {
  tab: "People", prefix: "p",
  cols: [
    ["id", "Person ID", "ID"], ["name", "Name"], ["roles", "Role", "Roles"],
    ["location", "City/State", "Location"], ["country", "Country"], ["languages", "Language", "Languages"],
    ["bio", "Bio"], ["cvLink", "CV link"], ["email", "Email"], ["phone", "Phone"], ["notes", "Notes"]
  ],
  list: ["roles", "languages"],
  dates: [],
  hidden: ["email", "phone"],                 // stay in the Sheet only, never sent to the website
  managed: ["id"]                             // set by the script
};

const CONTRACTS = {
  tab: "Contracts", prefix: "c",
  cols: [
    ["id", "Contract ID", "ID"], ["person", "Name", "Person"], ["projects", "Projects"],
    ["start", "Start date", "Start"], ["end", "End date", "End"],
    ["payTerms", "Payment terms", "Pay terms"], ["notes", "Notes"],
    ["personId", "Person ID"]                 // added at the end by the script: keeps the link to the person
  ],
  list: ["projects"],
  dates: ["start", "end"],
  hidden: [],
  managed: ["id", "person", "personId"]
};

const MAX_FAILS = 10;       // wrong passwords allowed...
const LOCK_MINUTES = 10;    // ...before everyone is paused for this long
const CACHE_SECONDS = 600;  // the website's data is kept ready for up to 10 minutes

// ---------- Sheet menu ----------

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Contractors")
    .addItem("Check data", "checkData")
    .addItem("Refresh the website's data", "refreshSite")
    .addItem("Set up (run once)", "setup")
    .addToUi();
}

/** Run once after pasting the script: adds IDs, the Person ID column and the name dropdown. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = sheet_(PEOPLE), cSheet = sheet_(CONTRACTS);
  const pL = layout_(pSheet, PEOPLE), cL = layout_(cSheet, CONTRACTS);
  withLock_(repair_);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(pSheet.getRange(2, pL.col.name + 1, 1000, 1), true).setAllowInvalid(true).build();
  cSheet.getRange(2, cL.col.person + 1, Math.max(cSheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  pSheet.setFrozenRows(1);
  cSheet.setFrozenRows(1);
  clearCache_();
  ss.toast("Setup finished", "Contractors", 5);
}

function refreshSite() {
  clearCache_();
  SpreadsheetApp.getActiveSpreadsheet().toast("The website will load fresh data on its next visit.", "Contractors", 5);
}

/** Edits made by hand in the Sheet: make the website pick them up. */
function onEdit(e) {
  try {
    const name = e.range.getSheet().getName();
    if (name === PEOPLE.tab || name === CONTRACTS.tab) clearCache_();
  } catch (err) { /* simple triggers have limited permissions: the 10-minute cache expiry is the backstop */ }
}

/** Lists things worth fixing. Run from the Contractors menu. */
function checkData() {
  const d = freshData_(false);
  const issues = [];
  const list = function (title, names) {
    if (names.length) issues.push(title + " (" + names.length + "): " + names.slice(0, 8).join(", ") + (names.length > 8 ? ", …" : ""));
  };
  list("Contracts whose Name doesn't match anyone in People", d.unlinkedNames);
  const seen = {}, dupes = [];
  d.people.forEach(function (p) {
    const k = p.name.toLowerCase();
    if (seen[k] && dupes.indexOf(p.name) < 0) dupes.push(p.name);
    seen[k] = true;
  });
  list("People listed more than once", dupes);
  list("People with no Role", d.people.filter(function (p) { return !p.roles.length; }).map(function (p) { return p.name; }));
  list("People with no Country", d.people.filter(function (p) { return !p.country; }).map(function (p) { return p.name; }));
  list("Contracts with an end date before the start date", d.contracts.filter(function (c) { return c.start && c.end && c.start > c.end; }).map(function (c) { return c.person; }));
  list("Contracts with no end date (counted as ongoing)", d.contracts.filter(function (c) { return !c.end; }).map(function (c) { return c.person; }));
  list("Contracts with a date the site can't read", d.contracts.filter(function (c) {
    return (c.start && !/^\d{4}-\d{2}-\d{2}$/.test(c.start)) || (c.end && !/^\d{4}-\d{2}-\d{2}$/.test(c.end));
  }).map(function (c) { return c.person; }));
  list("Contracts with no payment terms", d.contracts.filter(function (c) { return !c.payTerms; }).map(function (c) { return c.person; }));
  const withContract = {};
  d.contracts.forEach(function (c) { withContract[c.personId] = true; });
  const none = d.people.filter(function (p) { return !withContract[p.id]; }).map(function (p) { return p.name; });
  const text = (issues.length ? issues.join("\n\n") : "No problems found.") +
    "\n\n" + d.people.length + " people, " + d.contracts.length + " contracts" +
    (none.length ? " (" + none.length + " people have no contract yet and show as Prospect)." : ".");
  Logger.log(text);
  try { SpreadsheetApp.getUi().alert("Data check", text, SpreadsheetApp.getUi().ButtonSet.OK); } catch (err) { /* not run from the Sheet */ }
}

// ---------- web app ----------

function doGet() {
  return json_({ ok: true, message: "Contractors API is running." });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const auth = checkPassword_(req.password);
    if (!auth.ok) return json_({ ok: false, error: auth.error });

    const writes = ["savePerson", "deletePerson", "saveContract", "deleteContract"];
    if (req.action === "list") return json_(reply_(auth, getData_()));
    if (writes.indexOf(req.action) < 0) return json_({ ok: false, error: "unknown_action" });
    if (!auth.canWrite) return json_({ ok: false, error: "read_only" });

    return json_(withLock_(function () {
      const cur = freshData_(true);                 // read straight from the Sheet, not the cache
      let id;
      if (req.action === "savePerson") { checkVersion_(cur.people, req.item); id = savePerson_(req.item); }
      else if (req.action === "deletePerson") { checkVersion_(cur.people, { id: req.id, v: req.v }); deletePerson_(req.id, cur); }
      else if (req.action === "saveContract") { checkVersion_(cur.contracts, req.item); id = saveContract_(req.item); }
      else if (req.action === "deleteContract") { checkVersion_(cur.contracts, { id: req.id, v: req.v }); deleteRow_(CONTRACTS, req.id); }
      repair_();                                    // e.g. a renamed person's contracts follow the new name
      const after = readAll_();
      clearCache_();
      cachePut_(after);
      return reply_(auth, after, id);
    }));
  } catch (err) {
    const msg = String(err && err.message);
    if (msg === "busy") return json_({ ok: false, error: "busy" });
    if (msg === "conflict") {
      const latest = reply_({ canWrite: true }, getData_());   // the newest data rides along so the page can refresh
      latest.ok = false;
      latest.error = "conflict";
      return json_(latest);
    }
    return json_({ ok: false, error: "server_error", detail: msg });
  }
}

function reply_(auth, d, id) {
  return {
    ok: true, canWrite: !!auth.canWrite, id: id,
    people: d.people, contracts: d.contracts,
    unlinked: d.unlinkedNames.length,
    unlinkedNames: auth.canWrite ? d.unlinkedNames.slice(0, 5) : []
  };
}

/** Only one save at a time. If the lock can't be had in 20 seconds the website shows "busy". */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error("busy");
  try { return fn(); } finally { lock.releaseLock(); }
}

/** Refuses a save/delete if someone changed that row after this person's screen loaded it. */
function checkVersion_(items, item) {
  if (!item || !item.id || !item.v) return;
  const found = items.filter(function (x) { return x.id === item.id; })[0];
  if (found && found.v !== item.v) throw new Error("conflict");
}

// ---------- password ----------

function checkPassword_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("fails") || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: "locked" };

  const props = PropertiesService.getScriptProperties();
  const viewPw = props.getProperty("PASSWORD");
  const editPw = props.getProperty("EDIT_PASSWORD");
  if (!viewPw) return { ok: false, error: "not_configured" };

  const given = String(pw || "");
  if (editPw && given === editPw) return { ok: true, canWrite: true };
  if (given === viewPw) return { ok: true, canWrite: !editPw };

  cache.put("fails", String(fails + 1), LOCK_MINUTES * 60);
  Utilities.sleep(800);
  return { ok: false, error: "bad_password" };
}

// ---------- cache (the whole data set is kept ready, in chunks) ----------

const CACHE_KEY = "data";

function cacheGet_() {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(CACHE_KEY + ":n") || 0);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(CACHE_KEY + ":" + i);
  const got = c.getAll(keys);
  let s = "";
  for (let i = 0; i < n; i++) {
    const part = got[CACHE_KEY + ":" + i];
    if (part == null) return null;
    s += part;
  }
  try { return JSON.parse(s); } catch (err) { return null; }
}

function cachePut_(d) {
  const s = JSON.stringify(d);
  const obj = {};
  let n = 0;
  for (let i = 0; i < s.length; i += 30000) { obj[CACHE_KEY + ":" + n] = s.slice(i, i + 30000); n++; }
  obj[CACHE_KEY + ":n"] = String(n);
  try { CacheService.getScriptCache().putAll(obj, CACHE_SECONDS); } catch (err) { /* too big or busy: just skip caching */ }
}

function gen_() { return CacheService.getScriptCache().get("gen") || ""; }

function clearCache_() {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(CACHE_KEY + ":n") || 0);
  const keys = [CACHE_KEY + ":n"];
  for (let i = 0; i < n; i++) keys.push(CACHE_KEY + ":" + i);
  c.removeAll(keys);
  c.put("gen", String(Date.now()) + Math.random(), 21600);
}

/** Reading takes no lock: it uses the cache, or reads the Sheet (and only locks if IDs need adding). */
function getData_() {
  const cached = cacheGet_();
  if (cached) return cached;
  const g0 = gen_();
  const d = freshData_(false);
  if (gen_() === g0) cachePut_(d);        // skip if a save happened while we were reading
  return d;
}

function freshData_(haveLock) {
  let d = readAll_();
  if (d.dirty) {
    if (haveLock) repair_(); else withLock_(repair_);
    d = readAll_();
  }
  return d;
}

// ---------- generic table helpers ----------

function fields_(S) { return S.cols.map(function (c) { return c[0]; }); }
function label_(S, f) { return S.cols.filter(function (c) { return c[0] === f; })[0][1]; }
function publicFields_(S) { return fields_(S).filter(function (f) { return S.hidden.indexOf(f) < 0; }); }
function writable_(S) {
  return fields_(S).filter(function (f) { return S.hidden.indexOf(f) < 0 && S.managed.indexOf(f) < 0; });
}
function isBlank_(row) { return row.every(function (c) { return String(c).trim() === ""; }); }
function ssTz_() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
function hash_(s) { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0; return h.toString(36); }

function sheet_(S) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(S.tab);
  if (!sheet) { sheet = ss.insertSheet(S.tab); sheet.setFrozenRows(1); }
  return sheet;
}

/** Finds each field's column by its heading, adding any missing heading at the end. */
function layout_(sheet, S) {
  const lastCol = sheet.getLastColumn();
  const header = lastCol
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); })
    : [];
  const col = {};
  const missing = [];
  S.cols.forEach(function (c) {
    const names = c.slice(1).concat([c[0]]).map(function (n) { return n.toLowerCase(); });
    let i = -1;
    names.forEach(function (n) { if (i < 0) i = header.indexOf(n); });
    if (i < 0) missing.push(c); else col[c[0]] = i;
  });
  if (missing.length) {
    sheet.getRange(1, header.length + 1, 1, missing.length)
      .setValues([missing.map(function (c) { return c[1]; })]).setFontWeight("bold");
    missing.forEach(function (c, k) { col[c[0]] = header.length + k; });
  }
  return { col: col, width: header.length + missing.length };
}

function splitList_(s) {
  const seen = {}, out = [];
  String(s || "").split(/[;,|]/).forEach(function (p) {
    const t = p.trim(), k = t.toLowerCase();
    if (t && !seen[k]) { seen[k] = true; out.push(t); }
  });
  return out;
}

function isoDate_(s) {
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (m) {
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return y + "-" + ("0" + m[1]).slice(-2) + "-" + ("0" + m[2]).slice(-2);
  }
  return s;
}

/** Every data row, as typed values (for dates) and as displayed text (for everything else). */
function readRecs_(sheet, L) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rng = sheet.getRange(2, 1, lastRow - 1, L.width);
  const raw = rng.getValues(), shown = rng.getDisplayValues();
  return raw.map(function (r, i) { return { row: i + 2, raw: r, shown: shown[i] }; });
}

function toItem_(rec, S, L, tz) {
  const item = {};
  publicFields_(S).forEach(function (f) {
    const j = L.col[f];
    let v;
    if (S.dates.indexOf(f) >= 0) {
      const r = rec.raw[j];
      v = r instanceof Date ? Utilities.formatDate(r, tz, "yyyy-MM-dd") : isoDate_(String(rec.shown[j]).trim());
    } else {
      v = String(rec.shown[j] == null ? "" : rec.shown[j]).trim();
    }
    item[f] = S.list.indexOf(f) >= 0 ? splitList_(v) : v;
  });
  return item;
}

/** Reads both tabs (no writing). `dirty` means IDs or links need fixing by repair_(). */
function readAll_() {
  const tz = ssTz_();
  let dirty = false;

  const pSheet = sheet_(PEOPLE), pL = layout_(pSheet, PEOPLE);
  const people = [];
  readRecs_(pSheet, pL).forEach(function (rec) {
    if (isBlank_(rec.shown)) return;
    const item = toItem_(rec, PEOPLE, pL, tz);
    if (!item.id) { dirty = true; return; }
    people.push(item);
  });
  const byId = {}, byName = {};
  people.forEach(function (p) {
    byId[p.id] = p;
    const k = p.name.toLowerCase();
    (byName[k] = byName[k] || []).push(p);
  });
  people.forEach(function (p) { p.v = hash_(JSON.stringify(p)); });

  const cSheet = sheet_(CONTRACTS), cL = layout_(cSheet, CONTRACTS);
  const contracts = [], unlinkedNames = [];
  readRecs_(cSheet, cL).forEach(function (rec) {
    if (isBlank_(rec.shown)) return;
    const c = toItem_(rec, CONTRACTS, cL, tz);
    let person = c.personId && byId[c.personId] ? byId[c.personId] : null;
    if (!person && c.person) {
      const m = byName[c.person.toLowerCase()];
      if (m && m.length === 1) person = m[0];
    }
    if (person) {
      if (c.personId !== person.id || c.person !== person.name) dirty = true;
      c.personId = person.id; c.person = person.name;
    } else {
      c.personId = "";
      unlinkedNames.push(c.person || "(no name)");
    }
    if (!c.id) { dirty = true; return; }
    c.v = hash_(JSON.stringify(c));
    contracts.push(c);
  });
  return { people: people, contracts: contracts, unlinkedNames: unlinkedNames, dirty: dirty };
}

// ---------- repairs: IDs and links (only ever called while holding the lock) ----------

function newId_(prefix, used) {
  let id;
  do { id = prefix + Utilities.getUuid().replace(/-/g, "").slice(0, 8); } while (used[id]);
  used[id] = true;
  return id;
}

function valuesOf_(sheet, L) {
  const lastRow = sheet.getLastRow();
  return lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, L.width).getValues();
}

function writeColumn_(sheet, colIdx, vals) {
  sheet.getRange(2, colIdx + 1, vals.length, 1).setNumberFormat("@")
    .setValues(vals.map(function (r) { return [r[colIdx]]; }));
}

/** Gives rows an ID, and keeps each contract's Name and Person ID in step with the People tab. */
function repair_() {
  const pSheet = sheet_(PEOPLE), pL = layout_(pSheet, PEOPLE);
  const cSheet = sheet_(CONTRACTS), cL = layout_(cSheet, CONTRACTS);
  const pVals = valuesOf_(pSheet, pL), cVals = valuesOf_(cSheet, cL);

  const used = {};
  pVals.forEach(function (r) { const id = String(r[pL.col.id]).trim(); if (id) used[id] = true; });
  cVals.forEach(function (r) { const id = String(r[cL.col.id]).trim(); if (id) used[id] = true; });

  let pChanged = false;
  pVals.forEach(function (r) {
    if (!isBlank_(r) && !String(r[pL.col.id]).trim()) { r[pL.col.id] = newId_(PEOPLE.prefix, used); pChanged = true; }
  });
  if (pChanged) writeColumn_(pSheet, pL.col.id, pVals);

  const nameOf = {}, idsByName = {};
  pVals.forEach(function (r) {
    const id = String(r[pL.col.id]).trim(), name = String(r[pL.col.name]).trim();
    if (!id || !name) return;
    nameOf[id] = name;
    (idsByName[name.toLowerCase()] = idsByName[name.toLowerCase()] || []).push(id);
  });

  let idChanged = false, linkChanged = false;
  cVals.forEach(function (r) {
    if (isBlank_(r)) return;
    if (!String(r[cL.col.id]).trim()) { r[cL.col.id] = newId_(CONTRACTS.prefix, used); idChanged = true; }
    const pid = String(r[cL.col.personId]).trim(), nm = String(r[cL.col.person]).trim();
    let id = pid && nameOf[pid] ? pid : null;
    if (!id && nm) {
      const m = idsByName[nm.toLowerCase()];
      if (m && m.length === 1) id = m[0];
    }
    if (!id) return;
    if (r[cL.col.personId] !== id) { r[cL.col.personId] = id; linkChanged = true; }
    if (r[cL.col.person] !== nameOf[id]) { r[cL.col.person] = nameOf[id]; linkChanged = true; }
  });
  if (idChanged) writeColumn_(cSheet, cL.col.id, cVals);
  if (linkChanged) { writeColumn_(cSheet, cL.col.person, cVals); writeColumn_(cSheet, cL.col.personId, cVals); }
}

// ---------- writing a row (only the columns this script owns; nothing else is touched) ----------

function findRow_(sheet, L, id) {
  const lastRow = sheet.getLastRow();
  if (!id || lastRow < 2) return 0;
  const ids = sheet.getRange(2, L.col.id + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id)) return i + 2;
  }
  return 0;
}

function usedIds_() {
  const used = {};
  [PEOPLE, CONTRACTS].forEach(function (S) {
    const sh = sheet_(S), L = layout_(sh, S);
    valuesOf_(sh, L).forEach(function (r) { const id = String(r[L.col.id]).trim(); if (id) used[id] = true; });
  });
  return used;
}

/** Writes runs of neighbouring columns. Dates go in as real dates; everything else as plain text. */
function writeCells_(sheet, rowNum, L, S, values, tz) {
  const idx = Object.keys(values).map(function (f) { return { f: f, j: L.col[f] }; })
    .sort(function (a, b) { return a.j - b.j; });
  let i = 0;
  while (i < idx.length) {
    let k = i;
    while (k + 1 < idx.length && idx[k + 1].j === idx[k].j + 1) k++;
    const run = idx.slice(i, k + 1);
    const isDate = function (c) { return S.dates.indexOf(c.f) >= 0; };
    const fmts = run.map(function (c) { return isDate(c) ? DATE_FORMAT : "@"; });
    const vals = run.map(function (c) {
      const v = values[c.f];
      return isDate(c) ? (v ? Utilities.parseDate(v, tz, "yyyy-MM-dd") : "") : v;
    });
    sheet.getRange(rowNum, run[0].j + 1, 1, run.length).setNumberFormats([fmts]).setValues([vals]);
    i = k + 1;
  }
}

function saveRow_(S, item, managed) {
  const sheet = sheet_(S), L = layout_(sheet, S), tz = ssTz_();
  S.dates.forEach(function (f) {
    const v = String(item[f] || "").trim();
    if (v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error("Dates must look like 2026-12-31");
  });
  let rowNum = findRow_(sheet, L, item.id);
  const values = {};
  writable_(S).forEach(function (f) {
    const v = item[f];
    values[f] = Array.isArray(v) ? v.join("; ") : String(v == null ? "" : v).trim();
  });
  Object.keys(managed || {}).forEach(function (f) { values[f] = managed[f]; });
  let id = item.id;
  if (!rowNum) {
    id = newId_(S.prefix, usedIds_());
    values.id = id;
    rowNum = sheet.getLastRow() + 1;
  }
  writeCells_(sheet, rowNum, L, S, values, tz);
  return id;
}

function deleteRow_(S, id) {
  const sheet = sheet_(S), L = layout_(sheet, S);
  const rowNum = findRow_(sheet, L, id);
  if (rowNum) sheet.deleteRow(rowNum);
}

// ---------- people ----------

function savePerson_(item) {
  if (!item || !String(item.name || "").trim()) throw new Error("Name is required");
  return saveRow_(PEOPLE, item, {});
}

/** Deleting a person also deletes their contracts. */
function deletePerson_(id, cur) {
  if (!id) return;
  const cSheet = sheet_(CONTRACTS), cL = layout_(cSheet, CONTRACTS);
  const rows = cur.contracts.filter(function (c) { return c.personId === id; })
    .map(function (c) { return findRow_(cSheet, cL, c.id); })
    .filter(function (n) { return n > 0; })
    .sort(function (a, b) { return b - a; });
  rows.forEach(function (n) { cSheet.deleteRow(n); });
  deleteRow_(PEOPLE, id);
}

// ---------- contracts ----------

function peopleNames_() {
  const sh = sheet_(PEOPLE), L = layout_(sh, PEOPLE);
  const out = {};
  valuesOf_(sh, L).forEach(function (r) {
    const id = String(r[L.col.id]).trim();
    if (id) out[id] = String(r[L.col.name]).trim();
  });
  return out;
}

function saveContract_(item) {
  const pid = String((item && item.personId) || "").trim();
  const name = peopleNames_()[pid];
  if (!name) throw new Error("Choose a person");
  if (item.start && item.end && item.start > item.end) throw new Error("The start date is after the end date");
  return saveRow_(CONTRACTS, item, { person: name, personId: pid });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
