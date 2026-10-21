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
    news:     "1Lo_7AdiMoOA6li6yo2e62JbIVjwzTJ3PausRyohLjBQ",
    widow:    "16TcMEn5ZY5x1qB75iM3ss4BUet44p-l38bDdpKAr1Lk",
    dues:     "1p5_VVZVtcsHXlzQsqXuo2Y5kxkCnCdZCvXMEEzJKU2Y"
  },
  FOLDERS: {
    tinute:  "12BoNMqN_mRn1tMZ4FvrPAPq3PnNM6HeT",
    minutul: "1GGwFADloHeiV3jAc2EZANc_p-9f3jEHs",
    library: "1hn9fZq0MKpgHwunFmZnwe3dYaSvH_uND"
  }
};

function doGet(e) {
  const token=String((e&&e.parameter&&e.parameter.card)||"");
  if (token) return publicCard_(token);
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
    if (req.action === "photo_self") return json_(photoSelf_(me));
    if (req.action === "card_profile_get") return json_(cardProfileGet_(who.email));
    if (req.action === "card_profile_set") return json_(cardProfileSet_(who.email, req.profile));
    if (req.action === "card_links") return json_(cardLinks_(who.email));
    if (req.action === "file") return json_(file_(String(req.id || ""), level_(me.grad)));
    if (req.action === "plan_add") return json_(planAdd_(me, who.email, req));
    if (req.action === "plan_cancel") return json_(planCancel_(who.email, String(req.id || "")));
    if (req.action === "admin_read") return json_(adminRead_(me));
    if (req.action === "attendance_set") return json_(attendanceSet_(me, who.email, req));
    if (req.action === "payment_add") return json_(paymentAdd_(me, who.email, req));
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
  const k=key_(r[1],r[2]);
  return { key: k, row: r.row, name: r[2] + " " + r[1], full: r[1] + " " + r[2], grad: r[4], functie: r[5], demnitar: r[6], uniqueIdentity: rows.filter(x=>norm_(x[12])==="activ"&&key_(x[1],x[2])===k).length===1 };
}

/* Date de vizita private, controlate de titular. Nu contin poza si nu sunt publicate. */
function cardProfileKey_(email) {
  const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(email).toLowerCase());
  return "card_v1_" + Utilities.base64EncodeWebSafe(hash).replace(/=+$/, "");
}
function cardProfileGet_(email) {
  const raw=PropertiesService.getScriptProperties().getProperty(cardProfileKey_(email));
  const saved=raw?JSON.parse(raw):{};
  return { profile: Object.assign({
    personalPhone:"", personalEmail:"", company:"", position:"", workPhone:"",
    workEmail:"", workWebsite:"", showMasonic:false, showProfane:false,
    shareMasonicPhoto:false, shareProfanePhoto:false
  }, saved) };
}
function cardProfileSet_(email, input) {
  if (!input || typeof input!=="object" || Array.isArray(input)) return { error:"bad_request" };
  const textFields=["personalPhone","personalEmail","company","position","workPhone","workEmail","workWebsite"];
  const boolFields=["showMasonic","showProfane","shareMasonicPhoto","shareProfanePhoto"];
  const existing=cardProfileGet_(email).profile, result={...existing};
  for (const k of Object.keys(input)) {
    if (textFields.includes(k)) {
      if (typeof input[k]!=="string" || input[k].length>240) return {error:"bad_request"};
      result[k]=input[k].trim();
    } else if (boolFields.includes(k)) {
      if (typeof input[k]!=="boolean") return {error:"bad_request"};
      result[k]=input[k];
    } else return {error:"bad_request"};
  }
  // Preferintele de partajare sunt salvate, dar nu activeaza singure un endpoint public.
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {PropertiesService.getScriptProperties().setProperty(cardProfileKey_(email),JSON.stringify(result))}
  finally {lock.releaseLock()}
  return {profile:result};
}

/* Legaturi permanente: identificatori aleatori, niciodata date din registrul matricol. */
function cardLinks_(email) {
  const props=PropertiesService.getScriptProperties();
  const key=cardProfileKey_(email);
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  let pair;
  try {
    const saved=props.getProperty("card_links_"+key);
    if(saved)pair=JSON.parse(saved);
    else {
      pair={m:Utilities.getUuid().replace(/-/g,"")+Utilities.getUuid().replace(/-/g,""),
            p:Utilities.getUuid().replace(/-/g,"")+Utilities.getUuid().replace(/-/g,"")};
      props.setProperties({["card_links_"+key]:JSON.stringify(pair),
        ["card_public_"+pair.m]:email+"|m",["card_public_"+pair.p]:email+"|p"});
    }
  } finally {lock.releaseLock()}
  const base=ScriptApp.getService().getUrl();
  const p=cardProfileGet_(email).profile;
  return {masonic:p.showMasonic?base+"?card="+pair.m:"",
          profan:p.showProfane?base+"?card="+pair.p:""};
}
function cardEsc_(v) {
  return String(v||"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}
function cardVcf_(value) {
  return String(value||"").replace(/[\r\n]+/g," ").replace(/\\/g,"\\\\")
    .replace(/;/g,"\\;").replace(/,/g,"\\,").trim();
}
function publicCard_(token) {
  const denied=()=>HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta name="robots" content="noindex,nofollow"></head>'+
    '<body style="font-family:sans-serif;padding:32px">Acest contact nu este disponibil.</body></html>')
    .setTitle("Contact indisponibil");
  if(!/^[a-f0-9]{64}$/.test(token))return denied();
  try {
    const props=PropertiesService.getScriptProperties();
    const link=props.getProperty("card_public_"+token);
    if(!link)return denied();
    const ix=link.lastIndexOf("|"),email=link.slice(0,ix),kind=link.slice(ix+1);
    if(!["m","p"].includes(kind))return denied();
    const me=member_(email);
    if(!me||!me.uniqueIdentity)return denied();
    const p=cardProfileGet_(email).profile;
    if(!(kind==="m"?p.showMasonic:p.showProfane))return denied();
    const parts=String(me.name).trim().split(/\s+/);
    const fn=cardVcf_(me.name);
    const lines=["BEGIN:VCARD","VERSION:3.0","FN:"+fn,
      "N:"+cardVcf_(parts[parts.length-1])+";"+cardVcf_(parts.slice(0,-1).join(" "))+";;;"];
    let rows=[];
    if(kind==="m"){
      rows=[["ORG","R.L. Transilvania nr. 216 - MLNR"],
        ["TITLE",me.functie||""],["TEL;TYPE=CELL",p.personalPhone],
        ["EMAIL",p.personalEmail],["URL","https://transylvanialodge216.org/"],
        ["NOTE","R.L. Transilvania nr. 216 | Or. Cluj | Obedienta: Marea Loja Nationala din Romania"+(me.functie?" | Functie in Loja: "+me.functie:"")+(me.demnitar?" | Demnitate MLNR: "+me.demnitar:"")]];
    }else{
      rows=[["ORG",p.company],["TITLE",p.position],
        ["TEL;TYPE=WORK",p.workPhone],["EMAIL;TYPE=WORK",p.workEmail],
        ["URL",p.workWebsite]];
    }
    rows.forEach(r=>{if(r[1])lines.push(r[0]+":"+cardVcf_(r[1]))});
    if(kind==="m"?p.shareMasonicPhoto:p.shareProfanePhoto){
      const photo=photoSelf_(me);
      if(photo.base64 && photo.mimeType==="image/png"){
        const chunks=photo.base64.match(/.{1,72}/g)||[];
        lines.push("PHOTO;ENCODING=b;TYPE=PNG:"+chunks.join("\r\n "));
      }
    }
    lines.push("END:VCARD");
    const vcf=lines.join("\r\n")+"\r\n";
    // Servim direct vCard-ul in raspunsul GET: fara HTML si fara buton intermediar.
    // iOS poate deschide previzualizarea contactului; unele browsere Android descarca .vcf.
    return ContentService.createTextOutput(vcf)
      .setMimeType(ContentService.MimeType.VCARD);
  }catch(e){return denied()}
}

/* Fotografie privata: numai pentru membrul autentificat. */
function photoSelf_(me) {
  if (!me.uniqueIdentity) return { error: "photo_unavailable" };
  const rows = SpreadsheetApp.openById(CFG.SHEETS.matricol).getSheetByName("Frați").getRange("A7:E100").getDisplayValues();
  const found = rows.filter(r => norm_(r[4]) === "activ" && key_(r[1], r[2]) === me.key);
  if (found.length !== 1) return { error: "photo_unavailable" };
  const manifest = JSON.parse(DriveApp.getFileById("1Hhi2f5w2mvELIH8XSVNe8ii4Hr4u0EIm").getBlob().getDataAsString("UTF-8").replace(/^\uFEFF/, ""));
  const items = manifest.entries.filter(e => e.member_id === found[0][0] && e.status === "Activ");
  if (items.length !== 1 || !/^[A-Za-z0-9_-]+$/.test(items[0].drive_file_id)) return { error: "photo_unavailable" };
  const blob = DriveApp.getFileById(items[0].drive_file_id).getBlob();
  if (blob.getContentType() !== "image/png" || blob.getBytes().length > 2097152) return { error: "photo_unavailable" };
  return { mimeType: "image/png", base64: Utilities.base64Encode(blob.getBytes()) };
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

/* Trunchiul vaduvei, informativ pana la adoptarea hotararii.
   Se transmite DOAR randul fratelui autentificat si totalurile generale.
   Datele nominale ale altor frati nu intra niciodata in raspunsul API. */
function widowFund_(me) {
  const sh = SpreadsheetApp.openById(CFG.SHEETS.widow).getSheetByName("Situatie");
  if (!sh) throw new Error("missing_widow_sheet");
  const rows = sh.getRange("A3:J200").getDisplayValues();
  const totals = rows[0] || [];
  // Registrul financiar foloseste "Nume Prenume"; cel al membrilor activi
  // poate include prenume suplimentare. Asociem doar pe cheia deja folosita
  // de aplicatie: nume + primul prenume, cu potrivire exacta si unica.
  const matches = rows.slice(2).filter(r => {
    if (!r[1]) return false;
    const words = norm_(r[1]).split(/\s+/).filter(Boolean);
    return words.length >= 2 && words[0] + "|" + words[1] === me.key;
  });
  const own = matches.length === 1 ? matches[0] : null;
  return {
    year: "6026–6027",
    adopted: false,
    general: {
      calculated: totals[6] || "0",
      contributed: totals[7] || "0",
      difference: totals[8] || "0"
    },
    mine: own ? {
      taxedMeetings: own[5] || "0",
      calculated: own[6] || "—",
      contributed: own[7] || "—",
      difference: own[8] || "—",
      status: own[9] || "—"
    } : null
  };
}

/* Cotizatii / Gala / sponsorizari, citire read-only.
   Public: exclusiv cheie membru -> doua statusuri, fara sume, solduri sau sponsorizari.
   Personal: numai randul identificat unic pentru contul autentificat.
   Nicio obligatie nu se afiseaza ca neachitata inainte de vot + validare. */
function finance_(me) {
  const ss = SpreadsheetApp.openById(CFG.SHEETS.dues);
  const config = ss.getSheetByName("Parametri").getRange("B5:D9").getDisplayValues();
  const voteDues = norm_(config[0][2]) === "adoptat";
  const voteGala = norm_(config[1][2]) === "adoptat";
  const validated = norm_(config[4][0]) === "validat";
  const rows = ss.getSheetByName("Situație frați").getRange("A5:K200").getDisplayValues()
    .filter(r => /^\d+$/.test(String(r[0] || "").trim()) && r[1]);
  const byKey = {};
  const ambiguous = {};
  rows.forEach(r => {
    const w = norm_(r[1]).split(/\s+/).filter(Boolean);
    if (w.length < 2) return;
    const key = w[0] + "|" + w[1];
    if (byKey[key]) ambiguous[key] = true;
    byKey[key] = r;
  });
  const normStatus = (value, adopted) => {
    if (!adopted) return "În așteptarea votului";
    if (!validated) return "Neconfirmat";
    const s = norm_(value);
    if (s === "achitat integral") return "Achitat integral";
    if (s === "achitat partial") return "Achitat parțial";
    if (s === "neachitat") return "Neachitat";
    return "Neconfirmat";
  };
  const publicStatuses = {};
  rows.forEach(r => {
    const w = norm_(r[1]).split(/\s+/).filter(Boolean);
    if (w.length < 2) return;
    const key = w[0] + "|" + w[1];
    if (ambiguous[key]) return;
    publicStatuses[key] = {
      dues: normStatus(r[5], voteDues),
      gala: normStatus(r[9], voteGala)
    };
  });
  const own = !ambiguous[me.key] ? byKey[me.key] : null;
  return {
    year: "6026–6027",
    adoption: { dues: voteDues, gala: voteGala },
    publicStatuses: publicStatuses,
    mine: own ? {
      dues: { target: own[2], paid: own[3], difference: own[4], status: normStatus(own[5], voteDues) },
      gala: { target: own[6], paid: own[7], difference: own[8], status: normStatus(own[9], voteGala) },
      sponsorships: own[10],
      payments: paymentHistory_(paymentRows_(),me.key)
    } : null
  };
}


/* ---------- administrare: roluri explicite pe identitatea din registru ----------
   Niciun rol nu este acceptat din request. Fiecare actiune verifica membrul
   activ si rolul in backend. Mircea Ursu are consultare, nu editare. */
function adminRole_(me) {
  const roles = {
    "coca|mihai": "secretar",
    "rosu|silviu": "mv",
    "ursu|mircea": "secretar_zonal",
    "beschiu|ovidiu": "trezorier",
    "simu|horia": "demnitar_mlnr",
    "mihalache|richard": "demnitar_mlnr",
    "hossu|andrei": "demnitar_mlnr",
    "bartos|adrian": "demnitar_mlnr",
    "paun|ciprian": "demnitar_mlnr",
    "savus|cristian": "demnitar_mlnr"
  };
  if(!me || !me.uniqueIdentity) return "";
  const role=roles[me.key]||"";
  const office=norm_(me.functie);
  if(role==="secretar" && office!=="secretar")return "";
  if(role==="trezorier" && office!=="trezorier")return "";
  if(role==="mv" && office!=="maestru venerabil")return "";
  // Calitatea de demnitar trebuie să fie înregistrată și în coloana G.
  if(role==="demnitar_mlnr"&&!String(me.demnitar||"").trim())return "";
  if(role==="secretar_zonal"&&norm_(me.demnitar)!=="secretar zonal")return "";
  return role;
}
function adminRequired_(me, role) {
  const current = adminRole_(me);
  if (!current || (role && current !== role)) return { error: "forbidden" };
  return null;
}
function memberKeyFromFull_(name) {
  const words = norm_(name).split(/\s+/).filter(Boolean);
  return words.length > 1 ? words[0] + "|" + words[1] : "";
}
function moneyNum_(text) {
  const raw = String(text == null ? "" : text).trim();
  if (!raw) return 0;
  const cleaned = raw.replace(/\s|lei/gi, "").replace(/\./g, "").replace(",", ".");
  const value = Number(cleaned);
  return Number.isFinite(value) ? value : 0;
}
function paymentRows_() {
  const sh = SpreadsheetApp.openById(CFG.SHEETS.dues).getSheetByName("Plăți");
  return sh.getRange(5, 1, Math.max(1, sh.getLastRow() - 4), 9).getDisplayValues()
    .filter(r => r[1] && r[2] && r[3]);
}
function paymentHistory_(rows, key) {
  return rows.filter(r => memberKeyFromFull_(r[1]) === key).map(r => ({
    date: r[0], component: r[2], amount: moneyNum_(r[3]),
    reference: r[4] || "", notes: r[5] || ""
  })).sort((a,b) => {
    const stamp=s=>{
      const m=/^(\d{2})\.(\d{2})\.(\d{4})/.exec(String(s||""));
      return m?Number(m[3]+m[2]+m[1]):0;
    };
    return stamp(b.date)-stamp(a.date);
  });
}
function adminRead_(me) {
  const denied = adminRequired_(me); if (denied) return denied;
  const sh = SpreadsheetApp.openById(CFG.SHEETS.presence).getSheetByName("Registru prezență");
  const grid = sh.getRange("A5:AZ200").getDisplayValues();
  const headers = grid[0] || [];
  const dates = headers.map((v,i) => ({name:v,index:i}))
    .filter(x => /^\d{2}\.\d{2}\.\d{2,4}$/.test(String(x.name).trim()));
  const members = grid.slice(1).filter(r => /^\d+$/.test(String(r[0]||"").trim()) && r[1])
    .map(r => ({key:key_(r[1],r[2]),name:r[1]+" "+r[2],marks:dates.map(d=>r[d.index]||"")}));
  const financeSheet = SpreadsheetApp.openById(CFG.SHEETS.dues).getSheetByName("Situație frați");
  const financeRows = financeSheet.getRange("A5:K200").getDisplayValues()
    .filter(r => /^\d+$/.test(String(r[0]||"").trim()) && r[1]);
  const payments = paymentRows_();
  const financial = financeRows.map(r => {
    const key = memberKeyFromFull_(r[1]);
    return { key:key,name:r[1],
      dues:{target:moneyNum_(r[2]),paid:moneyNum_(r[3]),remaining:moneyNum_(r[4]),status:r[5]},
      gala:{target:moneyNum_(r[6]),paid:moneyNum_(r[7]),remaining:moneyNum_(r[8]),status:r[9]},
      sponsorships:moneyNum_(r[10]),payments:paymentHistory_(payments,key) };
  });
  return {role:adminRole_(me),dates:dates.map(d=>d.name),attendance:members,financial:financial};
}
function attendanceSet_(me, email, req) {
  const denied = adminRequired_(me,"secretar");if (denied) return denied;
  const value = String(req.value||"").toUpperCase();
  if (["P","M","N",""].indexOf(value) < 0) return {error:"bad_request"};
  const key = String(req.memberKey||""), date = String(req.date||"");
  if (!key || !/^\d{2}\.\d{2}\.\d{2,4}$/.test(date)) return {error:"bad_request"};
  const lock=LockService.getScriptLock();if(!lock.tryLock(15000))return {error:"server",message:"Datele sunt actualizate de alt operator."};
  try {
    const ss=SpreadsheetApp.openById(CFG.SHEETS.presence);
    const sh=ss.getSheetByName("Registru prezență");
    const headers=sh.getRange("A5:AZ5").getDisplayValues()[0];
    const ix=headers.findIndex(x=>String(x).trim()===date);
    if(ix<7)return {error:"bad_request"};
    const rows=sh.getRange("A6:C200").getDisplayValues();
    const targets=rows.map((r,i)=>({r,n:i+6})).filter(x=>/^\d+$/.test(String(x.r[0]||"").trim())&&key_(x.r[1],x.r[2])===key);
    if(targets.length!==1)return {error:"bad_request"};
    const cell=sh.getRange(targets[0].n,ix+1);
    const before=String(cell.getDisplayValue()||"").trim().toUpperCase();
    if(!["P","M","N",""].includes(before))return {error:"bad_request"};
    if(before===value)return {ok:true,unchanged:true};
    let journal=ss.getSheetByName("Jurnal aplicație");
    if(!journal){journal=ss.insertSheet("Jurnal aplicație");journal.appendRow(["Moment","Operator","Frate","Ținută","Anterior","Nou"]);}
    // Jurnal mai intai; daca scrierea celulei esueaza se raporteaza eroare.
    journal.appendRow([new Date(),email,targets[0].r[1]+" "+targets[0].r[2],date,before,value]);
    cell.setValue(value);
    CacheService.getScriptCache().remove("sh1_presence");
    return {ok:true,memberKey:key,date:date,value:value};
  } finally {lock.releaseLock();}
}
function paymentAdd_(me,email,req) {
  const denied=adminRequired_(me,"trezorier");if(denied)return denied;
  const id=String(req.paymentId||"");
  if(!/^[a-zA-Z0-9_-]{20,100}$/.test(id))return {error:"bad_request"};
  const key=String(req.memberKey||""),category=String(req.category||"");
  if(["Cotizație Lojă","Gala Elitelor","Sponsorizare"].indexOf(category)<0)return {error:"bad_request"};
  const amount=Number(req.amount);
  if(!Number.isFinite(amount)||amount<=0||amount>1000000||Math.abs(Math.round(amount*100)-amount*100)>1e-6)return {error:"bad_request"};
  const method=String(req.method||"Numerar");
  if(["Numerar","Transfer bancar"].indexOf(method)<0)return {error:"bad_request"};
  // Plata se poate primi de la orice frate înscris în registru, chiar dacă
  // nu și-a activat încă accesul la aplicație; operatorul rămâne autorizat separat.
  const roster=rosterRows_().filter(r=>key_(r[1],r[2])===key);
  if(roster.length!==1)return {error:"bad_request"};
  const name=roster[0][1]+" "+String(roster[0][2]).trim();
  const lock=LockService.getScriptLock();if(!lock.tryLock(15000))return {error:"server",message:"Operațiune în curs, reîncearcă."};
  try{
    const sh=SpreadsheetApp.openById(CFG.SHEETS.dues).getSheetByName("Plăți");
    const n=sh.getLastRow();
    if(n>=5){
      const ids=sh.getRange(5,7,n-4,1).getDisplayValues().map(r=>r[0]);
      if(ids.indexOf(id)>=0)return {ok:true,duplicate:true,paymentId:id};
    }
    const next=n+1;
    sh.getRange(next,1,1,9).setValues([[new Date(),name,category,amount,method,"Înregistrat din aplicație",id,email,new Date()]]);
    sh.getRange(next,1).setNumberFormat("dd.MM.yyyy HH:mm");
    SpreadsheetApp.flush();
    return {ok:true,paymentId:id};
  }finally{lock.releaseLock();}
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
      me: { email: email, key: me.key, fisa: fisa, adminRole: adminRole_(me), functieLoja: me.functie || "", demnitateMLNR: me.demnitar || "" },
      roster: roster,
      presence: presence,
      visits: visits,
      plates: plates,
      lodges: lodges,
      news: news,
      widow: null,
      finance: null,
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
    try { out.widow = bootstrapTimed_("widow_fund", () => widowFund_(me), metrics); }
    catch (_) { out.errors.widow = "unavailable"; }
    try { out.finance = bootstrapTimed_("finance", () => finance_(me), metrics); }
    catch (_) { out.errors.finance = "unavailable"; }
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

/* ---------- fișiere ---------- */
function file_(id, level) {
  if (!/^[\w-]{10,}$/.test(id)) return { error: "bad_request" };
  // Controlul de grad ramane per solicitare in member_(). Catalogul are TTL 5 minute.
  const allowed = catalogAllowedIds_(catalogCached_(level));
  if (allowed.indexOf(id) < 0) return { error: "forbidden" };
  const f = DriveApp.getFileById(id);
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
