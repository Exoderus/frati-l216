/**
 * Spațiul fraților — serviciul R∴L∴ Transilvania nr. 216
 * Rulează pe contul lojii. Primește tokenul Google (doar email) al fratelui,
 * verifică în registrul „Frați activi” că are „Activ” la „Acces aplicație”
 * și abia apoi întoarce datele lojii.
 */
const CFG = {
  CLIENT_ID: "1081751748144-ch9rep5dusskif6jbh6galtkov55h8i6.apps.googleusercontent.com",
  CALENDAR_ID: "loja.transilvania@gmail.com",
  NOTIFY: "loja.transilvania@gmail.com", // aici sosesc anunțările de vizite din aplicație
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
    if (req.action === "file") return json_(file_(String(req.id || ""), level_(me.grad)));
    if (req.action === "plan_add") return json_(planAdd_(me, who.email, req));
    if (req.action === "plan_cancel") return json_(planCancel_(who.email, String(req.id || "")));
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
  return { key: key_(r[1], r[2]), row: r.row, name: r[2] + " " + r[1], full: r[1] + " " + r[2], grad: r[4] };
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

/* Acces pe grade: un folder din Bibliotecă al cărui nume conține „Calfă” se vede doar de la gradul de Calfă în sus,
   unul cu „Maestru” doar de Maeștri. Restul folderelor se văd de toți frații. Gradul vine din „Frați activi”. */
function level_(grad) {
  const g = norm_(grad);
  if (/maestr/.test(g)) return 3;
  if (/calf/.test(g)) return 2;
  return 1;
}
const folderLevel_ = name => { const n = norm_(name); return /maestr/.test(n) ? 3 : /calf/.test(n) ? 2 : 1; };

function catalog_(level) {
  level = level || 1;
  const tinuteTop = list_(CFG.FOLDERS.tinute).filter(isTinutaFolder_);
  let tinuteFiles = [];
  tinuteTop.forEach(f => { tinuteFiles = tinuteFiles.concat(list_(f.id).filter(x => x.mimeType !== "application/vnd.google-apps.folder" && isTinutaDoc_(x))); });
  const minutul = list_(CFG.FOLDERS.minutul).filter(x => x.mimeType !== "application/vnd.google-apps.folder" && released_(x.name));
  const libTop = list_(CFG.FOLDERS.library).filter(x => HIDDEN_().indexOf(x.id) < 0);
  let libFiles = [];
  libTop.filter(x => x.mimeType === "application/vnd.google-apps.folder" && folderLevel_(x.name) <= level).forEach(f => { libFiles = libFiles.concat(list_(f.id).filter(x => x.mimeType !== "application/vnd.google-apps.folder")); });
  const allowed = [].concat(tinuteFiles, minutul, libTop, libFiles).filter(x => x.mimeType !== "application/vnd.google-apps.folder").map(x => x.id);
  // Autorizarea de scurta durata se construieste in catalogCached_.
  const libTopVisible = libTop.filter(x => x.mimeType !== "application/vnd.google-apps.folder" || folderLevel_(x.name) <= level);
  return { tinuteTop: tinuteTop, tinuteFiles: tinuteFiles, minutul: minutul, libTop: libTopVisible, libFiles: libFiles };
}

/* Cache doar pentru date comune. Continutul mare sau erorile cad inapoi la citirea directa. */
function sharedCached_(key, ttl, fn) {
  const cache = CacheService.getScriptCache();
  try {
    const hit = cache.get("sh1_" + key);
    if (hit !== null) return JSON.parse(hit);
  } catch (_) {}
  const value = fn();
  try {
    const json = JSON.stringify(value);
    if (json.length < 45000) cache.put("sh1_" + key, json, ttl);
  } catch (_) {}
  return value;
}
function catalogCached_(level) {
  level = level || 1;
  const cat = sharedCached_("catalog_" + level, 300, () => catalog_(level));
  return cat;
}
function catalogAllowedIds_(cat) {
  return [].concat(cat.tinuteFiles || [], cat.minutul || [], cat.libTop || [], cat.libFiles || [])
    .filter(f => f.mimeType !== "application/vnd.google-apps.folder").map(f => f.id);
}
function events_() {
  const cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
  if (!cal) return [];
  const from = new Date(); from.setHours(0, 0, 0, 0);
  const to = new Date(from.getTime() + 60 * 864e5);
  const tz = "Europe/Bucharest";
  return cal.getEvents(from, to).map(ev => {
    const s = ev.getStartTime(), allDay = ev.isAllDayEvent();
    let e = ev.getEndTime();
    if (allDay) e = new Date(e.getTime() - 864e5); // sfârșitul unei zile întregi e exclusiv
    if (e < s) e = s;
    const day = d => Utilities.formatDate(d, tz, "yyyy-MM-dd");
    const id = ev.getId() + "@" + day(s);
    return allDay
      ? { id: id, summary: ev.getTitle(), location: ev.getLocation(), description: ev.getDescription(), start: { date: day(s) }, end: { date: day(e) } }
      : { id: id, summary: ev.getTitle(), location: ev.getLocation(), description: ev.getDescription(), start: { dateTime: s.toISOString() }, end: { date: day(e) } };
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

/* Masuratori anonime, exclusiv in jurnalul Apps Script; schema raspunsului nu se schimba. */
function bootstrapTimed_(stage, fn, metrics) {
  const started = Date.now();
  try { return fn(); }
  finally { metrics.push([stage, Date.now() - started]); }
}

function bootstrap_(me, email) {
  const started = Date.now();
  const metrics = [];
  let status = "ok";
  try {
    // Aceleasi citiri si aceeasi ordine ca in versiunea functionala.
    const roster = bootstrapTimed_("roster", () => rosterRows_().map(r => {
      const b = /^(\d{1,2})\.(\d{1,2})\./.exec(r[3]);
      return [r[0], r[1], r[2], b ? b[1] + "." + b[2] + ".2000" : "", r[4], r[5], "", "", "", "", "", r[11]];
    }), metrics);
    const fisa = bootstrapTimed_("member_profile", () => fisa_(me.key), metrics);
    const presence = bootstrapTimed_("presence", () => sharedCached_("presence", 120, () => SpreadsheetApp.openById(CFG.SHEETS.presence).getSheetByName("Registru prezență").getRange("A5:AZ200").getDisplayValues()), metrics);
    const visits = bootstrapTimed_("visits", () => sharedCached_("visits", 120, () => values_(CFG.SHEETS.visits, "A5:E1000", "Vizite")), metrics);
    const plates = bootstrapTimed_("plates", () => sharedCached_("plates", 120, () => values_(CFG.SHEETS.plates, "A5:J1000", "Planșe")), metrics);
    const lodges = bootstrapTimed_("lodges", () => sharedCached_("lodges", 1800, () => values_(CFG.SHEETS.lodges, "A2:J100")), metrics);
    const news = bootstrapTimed_("news", () => sharedCached_("news", 300, () => values_(CFG.SHEETS.news, "A2:E300")), metrics);
    const out = {
      me: { email: email, key: me.key, fisa: fisa },
      roster: roster,
      presence: presence,
      visits: visits,
      plates: plates,
      lodges: lodges,
      news: news,
      events: [],
      plans: [],
      excursions: [],
      errors: {}
    };
    try { out.excursions = bootstrapTimed_("excursions", () => sharedCached_("excursions", 120, () => excursions_()), metrics); }
    catch (e) { out.errors.excursions = String(e.message || e); }
    try { out.plans = bootstrapTimed_("plans", () => plans_(email), metrics); }
    catch (e) { out.errors.plans = String(e.message || e); }
    try { out.events = bootstrapTimed_("calendar", () => sharedCached_("calendar", 300, () => events_()), metrics); }
    catch (e) { out.errors.events = String(e.message || e); }
    const cat = bootstrapTimed_("drive_catalog", () => catalogCached_(level_(me.grad)), metrics);
    Object.keys(cat).forEach(k => out[k] = cat[k]);
    return out;
  } catch (e) {
    status = "error";
    throw e;
  } finally {
    // Nu scrie e-mailuri, nume, tokenuri, ID-uri de documente sau date din registre.
    try {
      console.log(JSON.stringify({
        kind: "bootstrap_performance",
        status: status,
        total_ms: Date.now() - started,
        stages_ms: Object.fromEntries(metrics)
      }));
    } catch (_) { /* masurarea nu intrerupe aplicatia */ }
  }
}

/* ---------- vizite anunțate („Cine mai merge?”) ----------
   Foaia „Vizite anunțate” din Registrul călătoriilor. Frații pot doar adăuga
   propria intenție și o pot anula pe a lor; Secretariatul schimbă Starea în Drive. */
const PLAN_SHEET_ = "Vizite anunțate";
const PLAN_STARI_ = ["Anunțată", "Confirmată", "Efectuată", "Anulată"];
function planSheet_() {
  const sh = SpreadsheetApp.openById(CFG.SHEETS.visits).getSheetByName(PLAN_SHEET_);
  if (!sh) throw new Error("Lipsește foaia " + PLAN_SHEET_);
  return sh;
}
function planRows_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 14).getDisplayValues().map((r, i) => ({ r: r, row: i + 2 }));
}
const tz_ = "Europe/Bucharest";
const dmy_ = d => Utilities.formatDate(d, tz_, "dd.MM.yyyy");
function parseDate_(s) {
  s = String(s || "").trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/.exec(s);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}
const clean_ = (s, n) => String(s || "").replace(/[\u0000-\u001f]/g, " ").trim().slice(0, n);

function plans_(email) {
  const t0 = new Date(); t0.setHours(0, 0, 0, 0);
  return planRows_(planSheet_()).map(x => {
    const r = x.r, d = parseDate_(r[3]);
    return { id: r[11], name: r[1], date: d ? Utilities.formatDate(d, tz_, "yyyy-MM-dd") : "", loja: r[4], orient: r[5],
             escort: r[6], car: r[7], note: r[8], stare: r[9] || "Anunțată", nota: r[10], evt: r[12], persons: r[13], mine: String(r[2]).toLowerCase() === email, _d: d };
  }).filter(p => p.id && norm_(p.stare) !== "anulata" && (p._d ? p._d >= t0 : String(p.evt).indexOf("ex:") === 0))
    .map(p => { delete p._d; return p; });
}

/* Excursiile lojii: foaia „Excursii” din Registrul călătoriilor (un rând pe excursie) */
function excursions_() {
  const sh = SpreadsheetApp.openById(CFG.SHEETS.visits).getSheetByName("Excursii");
  if (!sh) return [];
  return sh.getRange("A2:L20").getDisplayValues()
    .filter(r => r[0] && r[1] && norm_(r[11]) !== "nu")
    .map(r => ({ id: String(r[0]).trim(), titlu: r[1], de: r[2], pana: r[3], loc: r[4], adresa: r[5], program: r[6],
                 cost: r[7], termen: r[8], organizator: r[9], info: r[10] }));
}

function planAdd_(me, email, req) {
  const t0 = new Date(); t0.setHours(0, 0, 0, 0);
  let d = null, lojaTxt, orient, evt = "";
  let persons = "";
  if (req.ex) {
    // excursie a lojii
    const X = excursions_().find(x => x.id === String(req.ex));
    if (!X) return { error: "bad_request", message: "Excursia nu mai este în listă." };
    const end = parseDate_(X.pana) || parseDate_(X.de);
    if (end && end < t0) return { error: "bad_request", message: "Excursia a avut deja loc." };
    d = parseDate_(X.de);
    lojaTxt = clean_(X.titlu, 150); orient = clean_(X.loc, 100); evt = "ex:" + X.id;
    persons = String(Math.max(1, Math.min(6, parseInt(req.persons, 10) || 1)));
  } else {
    d = parseDate_(req.date);
    if (!d || d < t0 || d - t0 > 400 * 864e5) return { error: "bad_request", message: "Data nu este validă." };
    if (req.ev) {
      // eveniment din calendarul lojii
      const E = events_().find(x => x.id === String(req.ev));
      if (!E) return { error: "bad_request", message: "Evenimentul nu mai este în calendar." };
      const s0 = parseDate_(E.start.date || Utilities.formatDate(new Date(E.start.dateTime), tz_, "yyyy-MM-dd"));
      const s1 = parseDate_(E.end.date) || s0;
      if (d < s0 || d > s1) return { error: "bad_request", message: "Data nu este în perioada evenimentului." };
      lojaTxt = clean_(E.summary, 150); orient = clean_(E.location, 100); evt = E.id;
      persons = String(Math.max(1, Math.min(6, parseInt(req.persons, 10) || 1)));
    } else {
      const lodges = values_(CFG.SHEETS.lodges, "A2:C100");
      const L = lodges.find(l => norm_(l[0]) === norm_(req.loja) && String(l[1]).trim() === String(req.nr || "").trim());
      if (!L) return { error: "bad_request", message: "Loja nu este în listă." };
      lojaTxt = L[0] + (L[1] ? " nr. " + L[1] : ""); orient = L[2] || "";
    }
  }
  const escort = req.escort ? "Caută Maestru însoțitor" : "";
  const car = req.car === "ofer" ? "Are locuri în mașină" : req.car === "caut" ? "Caută loc în mașină" : "";
  const note = clean_(req.note, 200);
  const dTxt = d ? dmy_(d) : "";
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { error: "server", message: "busy" };
  try {
    const sh = planSheet_();
    const rows = planRows_(sh);
    const active = x => { const xd = parseDate_(x.r[3]); return norm_(x.r[9]) !== "anulata" && (!xd ? String(x.r[12]).indexOf("ex:") === 0 : xd >= t0); };
    const mineActive = rows.filter(x => String(x.r[2]).toLowerCase() === email && active(x));
    const dup = mineActive.find(x => evt.indexOf("ex:") === 0 ? x.r[12] === evt
      : (evt ? x.r[12] === evt : norm_(x.r[4]) === norm_(lojaTxt)) && x.r[3] === dTxt);
    if (dup) {
      sh.getRange(dup.row, 7, 1, 3).setValues([[escort, car, note]]);
      if (persons) sh.getRange(dup.row, 14).setNumberFormat("@").setValue(persons);
    } else {
      if (mineActive.length >= 10) return { error: "bad_request", message: "Ai deja 10 vizite anunțate." };
      const row = sh.getLastRow() + 1;
      const now = Utilities.formatDate(new Date(), tz_, "dd.MM.yyyy HH:mm");
      sh.getRange(row, 1, 1, 14).setNumberFormat("@").setValues([[now, me.full, email, dTxt, lojaTxt, orient, escort, car, note, "Anunțată", "", Utilities.getUuid().slice(0, 8), evt, persons]]);
      notify_("Vizită anunțată: Fr∴ " + me.full + " — " + lojaTxt + (dTxt ? ", " + dTxt : ""), [
        "Fratele " + me.full + " a anunțat în aplicație:",
        "",
        "Unde: " + lojaTxt + (orient ? " (" + orient + ")" : ""),
        "Data: " + (dTxt || "de stabilit"),
        persons && persons !== "1" ? "Persoane: " + persons : "",
        escort ? "Însoțire: " + escort : "",
        car ? "Mașina: " + car : "",
        note ? "Mesajul fratelui: " + note : "",
        "",
        "Confirmarea se face din coloana Stare a foii „Vizite anunțate”:",
        "https://docs.google.com/spreadsheets/d/" + CFG.SHEETS.visits + "/edit"
      ]);
    }
  } finally { lock.releaseLock(); }
  return { ok: true, plans: plans_(email) };
}

function planCancel_(email, id) {
  if (!/^[\w-]{4,40}$/.test(id)) return { error: "bad_request" };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { error: "server", message: "busy" };
  try {
    const sh = planSheet_();
    const x = planRows_(sh).find(y => y.r[11] === id && String(y.r[2]).toLowerCase() === email);
    if (!x) return { error: "forbidden" };
    if (norm_(x.r[9]) === "efectuata") return { error: "bad_request", message: "Vizita este deja efectuată." };
    sh.getRange(x.row, 10).setValue("Anulată");
    notify_("Vizită anulată: Fr∴ " + x.r[1] + " — " + x.r[4] + (x.r[3] ? ", " + x.r[3] : ""), [
      "Fratele " + x.r[1] + " a anulat din aplicație:",
      "",
      "Unde: " + x.r[4] + (x.r[5] ? " (" + x.r[5] + ")" : ""),
      "Data: " + (x.r[3] || "de stabilit"),
      "",
      "https://docs.google.com/spreadsheets/d/" + CFG.SHEETS.visits + "/edit"
    ]);
  } finally { lock.releaseLock(); }
  return { ok: true, plans: plans_(email) };
}

/* ---------- autorizare live a fisierelor ----------
   Catalogul poate ramane in cache pentru navigare. La deschiderea unui document
   controlam locatia curenta din Drive si gradul fratelui, nu lista memorata.
   Structura acceptata este identica cu catalog_(): radacina bibliotecii, un
   singur subfolder al bibliotecii, Minutul si folderele de Tinute. */
function folderAllowedNow_(folderId, level, name, mime) {
  if (folderId === CFG.FOLDERS.library) return true;
  if (folderId === CFG.FOLDERS.minutul) return released_(name);
  let folder;
  try { folder = DriveApp.getFolderById(folderId); } catch (_) { return false; }
  const parents = folder.getParents();
  while (parents.hasNext()) {
    const p = parents.next().getId();
    if (p === CFG.FOLDERS.library) return folderLevel_(folder.getName()) <= level;
    if (p === CFG.FOLDERS.tinute) return isTinutaFolder_({ name: folder.getName(), mimeType: "application/vnd.google-apps.folder" })
      && isTinutaDoc_({ name: name, mimeType: mime });
  }
  return false;
}
function fileAllowedNow_(f, level) {
  const id = f.getId(), name = f.getName(), mime = f.getMimeType();
  if (HIDDEN_().indexOf(id) >= 0) return false;
  let direct = false, forbiddenPlacement = false;
  const parents = f.getParents();
  while (parents.hasNext()) {
    const pid = parents.next().getId();
    if (folderAllowedNow_(pid, level, name, mime)) direct = true;
    // Un document mutat intr-un folder de grad superior nu se livreaza
    // printr-un shortcut eventual ramas in alt folder.
    if (pid !== CFG.FOLDERS.library && pid !== CFG.FOLDERS.minutul) {
      try {
        const folder = DriveApp.getFolderById(pid);
        const anc = folder.getParents();
        while (anc.hasNext()) {
          if (anc.next().getId() === CFG.FOLDERS.library && folderLevel_(folder.getName()) > level)
            forbiddenPlacement = true;
        }
      } catch (_) {}
    }
    if (pid === CFG.FOLDERS.minutul && !released_(name)) forbiddenPlacement = true;
  }
  if (forbiddenPlacement) return false;
  if (direct) return true;
  // Shortcuturile sunt cazul exceptional ce necesita consultarea catalogului.
  const cat = catalogCached_(level);
  // Pentru fisierele accesate prin shortcuts, verificam shortcutul LIVE.
  // O intrare ramasa doar in catalogul cache nu ofera acces.
  const entries = [].concat(cat.tinuteFiles || [], cat.minutul || [], cat.libTop || [], cat.libFiles || []);
  const folders = [...new Set(entries.filter(x => x.id === id && x.parent).map(x => x.parent))];
  for (const pid of folders) {
    if (!folderAllowedNow_(pid, level, name, mime)) continue;
    try {
      const it = DriveApp.getFolderById(pid).getFiles();
      while (it.hasNext()) {
        const candidate = it.next();
        if (candidate.getMimeType() === "application/vnd.google-apps.shortcut"
            && candidate.getTargetId() === id) return true;
      }
    } catch (_) {}
  }
  return false;
}
function file_(id, level) {
  if (!/^[\w-]{10,}$/.test(id)) return { error: "bad_request" };
  const f = DriveApp.getFileById(id);
  if (!fileAllowedNow_(f, level || 1)) return { error: "forbidden" };
  const mime = f.getMimeType();
  const blob = /^application\/vnd\.google-apps\./.test(mime) ? f.getAs("application/pdf") : f.getBlob();
  return { name: f.getName(), mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* Notificare pe emailul lojii; nu blochează niciodată aplicația */
function notify_(subject, lines) {
  try {
    MailApp.sendEmail({ to: CFG.NOTIFY, subject: subject, name: "Spațiul fraților",
      body: lines.filter((l, i, a) => l !== "" || (i > 0 && a[i - 1] !== "")).join("\n") });
  } catch (e) { /* fără email, anunțarea rămâne în foaie */ }
}

/** Rulați o dată din editor (▶ autorizeaza) ca să acordați permisiunile. */
function autorizeaza() {
  rosterRows_(); catalog_(); events_();
  MailApp.getRemainingDailyQuota(); // cere permisiunea de a trimite emailuri
  Logger.log("Permisiuni acordate. Serviciul este gata.");
}
