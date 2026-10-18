/**
 * Retragerea partajărilor directe — R∴L∴ Transilvania nr. 216
 * Script de unică folosință. Se rulează din contul lojii, într-un proiect separat
 * (NU în „Spațiul fraților – serviciu”), apoi proiectul se șterge.
 *
 * 1) Rulați „verifica”  -> în Execution log vedeți cine ar fi scos, fără să schimbe nimic.
 * 2) Rulați „retrage”   -> scoate cititorii/comentatorii și oprește „Anyone with the link”.
 *
 * Rămân neatinși: proprietarul (contul lojii), toți editorii și conturile din PASTREAZA.
 */
const ELEMENTE = {
  "10. Tinute 6026 (folder)":            "12BoNMqN_mRn1tMZ4FvrPAPq3PnNM6HeT",
  "11. Minutul masonic (folder)":        "1GGwFADloHeiV3jAc2EZANc_p-9f3jEHs",
  "13. Biblioteca fraților (folder)":    "1hn9fZq0MKpgHwunFmZnwe3dYaSvH_uND",
  "Frați activi (foaie)":                "1irr1R_LVlaHtj9NJ5ONQw0xzfcFAgB6JxD5AXTih7ys",
  "Registru de prezență (foaie)":        "1wzgKkQ4UfjfTJy8vcQJg0b0hgIrok-x0S45ozmXTI7I",
  "Registru călătorii (foaie)":          "1Q_oiZyeHmlJ8Fd5EWAP2LccEsPrPKVHkVBsqZ3lAciw",
  "Registrul planșelor (foaie)":         "1YqHYz7_W6Yjw8m1BTzgWd1uvK5-zrAgMz6IX3NQ2rPQ"
};
const PASTREAZA = ["mihai.coca@gmail.com"]; // Secretarul, cont personal

function verifica() { ruleaza_(false); }
function retrage()  { ruleaza_(true); }

function ruleaza_(aplica) {
  let total = 0;
  Object.keys(ELEMENTE).forEach(nume => {
    const id = ELEMENTE[nume];
    let item;
    try { item = DriveApp.getFolderById(id); } catch (e) { item = DriveApp.getFileById(id); }
    total += curata_(item, nume, aplica);
    if (item.getFolders) { // folder: și tot conținutul, recursiv
      total += curataInterior_(item, nume, aplica);
    }
  });
  Logger.log((aplica ? "GATA. Accese retrase: " : "VERIFICARE. Accese care ar fi retrase: ") + total);
}

function curataInterior_(folder, cale, aplica) {
  let n = 0;
  const fi = folder.getFiles();
  while (fi.hasNext()) { const f = fi.next(); n += curata_(f, cale + " / " + f.getName(), aplica); }
  const fo = folder.getFolders();
  while (fo.hasNext()) { const d = fo.next(); n += curata_(d, cale + " / " + d.getName(), aplica); n += curataInterior_(d, cale + " / " + d.getName(), aplica); }
  return n;
}

function curata_(item, nume, aplica) {
  let n = 0;
  const pastreaza = PASTREAZA.map(x => x.toLowerCase());
  item.getViewers().forEach(u => {           // cititori + comentatori (editorii nu sunt incluși)
    const em = String(u.getEmail() || "").toLowerCase();
    if (!em || pastreaza.indexOf(em) >= 0) return;
    Logger.log((aplica ? "Scos: " : "De scos: ") + em + "  ←  " + nume);
    if (aplica) { try { item.removeViewer(em); } catch (e) { Logger.log("  (nu s-a putut: " + e.message + ")"); } }
    n++;
  });
  if (item.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    Logger.log((aplica ? "Link public oprit: " : "Link public de oprit: ") + nume);
    if (aplica) { try { item.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) { Logger.log("  (nu s-a putut: " + e.message + ")"); } }
    n++;
  }
  return n;
}
