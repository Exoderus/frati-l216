/**
 * Spațiul fraților — serviciul R∴L∴ Transilvania nr. 216
 * Rulează pe contul lojii. Primește tokenul Google (doar email) al fratelui,
 * verifică în registrul „Frați activi” că are „Activ” la „Acces aplicație”
 * și abia apoi întoarce datele lojii.
 */
const CFG = {
  CLIENT_ID: "1081751748144-ch9rep5dusskif6jbh6galtkov55h8i6.apps.googleusercontent.com",
  CALENDAR_ID: "loja.transilvania@gmail.com",
  SHEETS: {
    roster:   "1irr1R_LVlaHtj9NJ5ONQw0xzfcFAgB6JxD5AXTih7ys",
    presence: "1wzgKkQ4UfjfTJy8vcQJg0b0hgIrok-x0S45ozmXTI7I",
    visits:   "1Q_oiZyeHmlJ8Fd5EWAP2LccEsPrPKVHkVBsqZ3lAciw",
    plates:   "1YqHYz7_W6Yjw8m1BTzgWd1uvK5-zrAgMz6IX3NQ2rPQ",
    lodges:   "1Etk_aEJ_3PKXdCgPt7T21v9jKtfAYchN_8BHHJK-4c0",
    matricol: "1wahZJvhNIMfh10GCw47xhRPvxUfVSDCX8y5O3Swuark",
    news:     "1Lo_7AdiMoOA6li6yo2e62JbIVjwzTJ3PausRyohLjBQ"
  },
  FOLDERS: {
    tinute:  "12BoNMqN_mRn1tMZ4FvrPAPq3PnNM6HeT",
    minutul: "1GGwFADloHeiV3jAc2EZANc_p-9f3jEHs",
    library: "1hn9fZq0MKpgHwunFmZnwe3dYaSvH_uND"
  }
};

function doGet() {
  return json_({ ok: true, service: "Spațiul fraților" });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const who = verify_(req.token);
    if (who.error) return json_(who);
    const me = member_(who.email);
    if (req.action === "bootstrap") logAccess_(who.email, me);
    if (!me) return json_({ error: "not_member", email: who.email });
    if (req.action === "bootstrap") return json_(bootstrap_(me, who.email));
    if (req.action === "file") return json_(file_(String(req.id || "")));
    return json_({ error: "bad_request" });
  } catch (err) {
    return json_({ error: "server", message: String(err && err.message || err) });
  }
}

/* ---------- autentificare ---------- */
function verify_(token) {
  if (!token) return { error: "unauthorized" };
  const cache = CacheService.getScriptCache();
  const ck = "tok_" + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const r = UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { error: "unauthorized" };
  const t = JSON.parse(r.getContentText());
  const okIss = t.iss === "accounts.google.com" || t.iss === "https://accounts.google.com";
  if (t.aud !== CFG.CLIENT_ID || !okIss || String(t.email_verified) !== "true" || Number(t.exp) * 1000 < Date.now()) return { error: "unauthorized" };
  const out = { email: String(t.email).toLowerCase() };
  const ttl = Math.max(30, Math.min(600, Number(t.exp) - Math.floor(Date.now() / 1000)));
  cache.put(ck, JSON.stringify(out), ttl);
  return out;
}

function rosterRows_() {
  return SpreadsheetApp.openById(CFG.SHEETS.roster).getSheetByName("Frați activi").getRange("A4:M200").getDisplayValues()
    .map((r, i) => { r.row = i + 4; return r; })
    .filter(r => /^\d+$/.test(String(r[0]).trim()) && r[1]);
}
const norm_ = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
const key_ = (n, p) => norm_(n) + "|" + norm_(String(p || "").split(/\s+/)[0]);

function member_(email) {
  const rows = rosterRows_();
  const r = rows.find(x => String(x[10]).trim().toLowerCase() === email) || rows.find(x => String(x[9]).trim().toLowerCase() === email);
  if (!r || norm_(r[12]) !== "activ") return null;
  return { key: key_(r[1], r[2]), row: r.row, name: r[2] + " " + r[1] };
}

/* ---------- jurnal de acces ---------- */
// Coloanele N (Ultima accesare) și O (Accesări) din „Frați activi”, plus foaia „Jurnal acces”.
function logAccess_(email, me) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const ss = SpreadsheetApp.openById(CFG.SHEETS.roster);
    const now = Utilities.formatDate(new Date(), "Europe/Bucharest", "dd.MM.yyyy HH:mm");
    if (me) {
      const sh = ss.getSheetByName("Frați activi");
      const cnt = Number(sh.getRange(me.row, 15).getValue()) || 0;
      sh.getRange(me.row, 14, 1, 2).setValues([[now, cnt + 1]]);
    }
    let log = ss.getSheetByName("Jurnal acces");
    if (!log) {
      log = ss.insertSheet("Jurnal acces");
      log.getRange(1, 1, 1, 4).setValues([["Data și ora", "Cont Google", "Frate", "Rezultat"]]).setFontWeight("bold");
      log.setFrozenRows(1);
    }
    log.insertRowAfter(1);
    log.getRange(2, 1, 1, 4).setValues([[now, email, me ? me.name : "—", me ? "Intrat" : "Fără acces"]]);
  } catch (e) {
    // jurnalul nu blochează niciodată accesul
  } finally {
    lock.releaseLock();
  }
}

/* ---------- date ---------- */
function values_(id, range, sheetName) {
  const ss = SpreadsheetApp.openById(id);
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  return sh.getRange(range).getDisplayValues().filter(r => r.some(c => c !== ""));
}

function list_(folderId) {
  const out = [];
  const f = DriveApp.getFolderById(folderId);
  const it = f.getFiles();
  while (it.hasNext()) {
    const x = it.next();
    let id = x.getId(), mime = x.getMimeType();
    if (mime === "application/vnd.google-apps.shortcut") {
      try { id = x.getTargetId(); mime = x.getTargetMimeType(); } catch (e) { continue; }
    }
    out.push({ id: id, name: x.getName(), mimeType: mime, parent: folderId, modifiedTime: x.getLastUpdated().toISOString() });
  }
  const fo = f.getFolders();
  while (fo.hasNext()) {
    const x = fo.next();
    out.push({ id: x.getId(), name: x.getName(), mimeType: "application/vnd.google-apps.folder", parent: folderId, modifiedTime: x.getLastUpdated().toISOString() });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name, "ro"));
}

const isTinutaFolder_ = f => f.mimeType === "application/vnd.google-apps.folder" && /^tinuta/.test(norm_(f.name));
// Minutul de regularitate se prezintă în ținută: devine vizibil din ziua ținutei, de la ora 21:00.
function released_(name) {
  const m = /(\d{1,2})\.(\d{1,2})\.(\d{4})/.exec(String(name || ""));
  if (!m) return true;
  return Date.now() >= new Date(+m[3], +m[2] - 1, +m[1], 21, 0, 0).getTime();
}
const isTinutaDoc_ = f => /minuta|convocator/i.test(f.name) || /^image\//.test(f.mimeType);
const HIDDEN_ = () => [CFG.SHEETS.news, CFG.SHEETS.lodges];

function catalog_() {
  const tinuteTop = list_(CFG.FOLDERS.tinute).filter(isTinutaFolder_);
  let tinuteFiles = [];
  tinuteTop.forEach(f => { tinuteFiles = tinuteFiles.concat(list_(f.id).filter(x => x.mimeType !== "application/vnd.google-apps.folder" && isTinutaDoc_(x))); });
  const minutul = list_(CFG.FOLDERS.minutul).filter(x => x.mimeType !== "application/vnd.google-apps.folder" && released_(x.name));
  const libTop = list_(CFG.FOLDERS.library).filter(x => HIDDEN_().indexOf(x.id) < 0);
  let libFiles = [];
  libTop.filter(x => x.mimeType === "application/vnd.google-apps.folder").forEach(f => { libFiles = libFiles.concat(list_(f.id).filter(x => x.mimeType !== "application/vnd.google-apps.folder")); });
  const allowed = [].concat(tinuteFiles, minutul, libTop, libFiles).filter(x => x.mimeType !== "application/vnd.google-apps.folder").map(x => x.id);
  CacheService.getScriptCache().put("allowed", JSON.stringify(allowed), 21600);
  return { tinuteTop: tinuteTop, tinuteFiles: tinuteFiles, minutul: minutul, libTop: libTop, libFiles: libFiles };
}

function events_() {
  const cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
  if (!cal) return [];
  const from = new Date(); from.setHours(0, 0, 0, 0);
  const to = new Date(from.getTime() + 60 * 864e5);
  const tz = "Europe/Bucharest";
  return cal.getEvents(from, to).map(ev => {
    const s = ev.getStartTime();
    return ev.isAllDayEvent()
      ? { summary: ev.getTitle(), location: ev.getLocation(), start: { date: Utilities.formatDate(s, tz, "yyyy-MM-dd") } }
      : { summary: ev.getTitle(), location: ev.getLocation(), start: { dateTime: s.toISOString() } };
  });
}

/* Fișa fratelui din „Registrul matricol L216” — doar a fratelui autentificat */
function fisa_(key) {
  try {
    const rows = SpreadsheetApp.openById(CFG.SHEETS.matricol).getSheetByName("Frați").getRange("A7:R100").getDisplayValues();
    const r = rows.find(x => x[1] && key_(x[1], x[2]) === key);
    if (!r) return null;
    return { id: r[0], status: r[4], grad: r[5], functie: r[6], initiere: r[8], calfa: r[9], maestru: r[10],
             afiliere: r[11], provenienta: r[12], nas: r[13], ani: r[15] };
  } catch (e) { return null; }
}

function bootstrap_(me, email) {
  // registrul fraților: doar ce afișează aplicația (fără telefoane, e-mailuri, anul nașterii)
  const roster = rosterRows_().map(r => {
    const b = /^(\d{1,2})\.(\d{1,2})\./.exec(r[3]);
    return [r[0], r[1], r[2], b ? b[1] + "." + b[2] + ".2000" : "", r[4], r[5], "", "", "", "", "", r[11]];
  });
  const out = {
    me: { email: email, key: me.key, fisa: fisa_(me.key) },
    roster: roster,
    presence: SpreadsheetApp.openById(CFG.SHEETS.presence).getSheetByName("Registru prezență").getRange("A5:AZ200").getDisplayValues(),
    visits: values_(CFG.SHEETS.visits, "A5:E1000", "Vizite"),
    plates: values_(CFG.SHEETS.plates, "A5:J1000", "Planșe"),
    lodges: values_(CFG.SHEETS.lodges, "A2:J100"),
    news: values_(CFG.SHEETS.news, "A2:E300"),
    events: [],
    errors: {}
  };
  try { out.events = events_(); } catch (e) { out.errors.events = String(e.message || e); }
  const cat = catalog_();
  Object.keys(cat).forEach(k => out[k] = cat[k]);
  return out;
}

/* ---------- fișiere ---------- */
function file_(id) {
  if (!/^[\w-]{10,}$/.test(id)) return { error: "bad_request" };
  let allowed = JSON.parse(CacheService.getScriptCache().get("allowed") || "null");
  if (!allowed || allowed.indexOf(id) < 0) { catalog_(); allowed = JSON.parse(CacheService.getScriptCache().get("allowed") || "[]"); }
  if (allowed.indexOf(id) < 0) return { error: "forbidden" };
  const f = DriveApp.getFileById(id);
  const mime = f.getMimeType();
  const blob = /^application\/vnd\.google-apps\./.test(mime) ? f.getAs("application/pdf") : f.getBlob();
  return { name: f.getName(), mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Rulați o dată din editor (▶ autorizeaza) ca să acordați permisiunile. */
function autorizeaza() {
  rosterRows_(); catalog_(); events_();
  Logger.log("Permisiuni acordate. Serviciul este gata.");
}
