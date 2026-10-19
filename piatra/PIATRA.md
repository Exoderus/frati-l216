# Piatra – specificație de implementare (aprobată de Secretar, 19.10.2026)

Gamificare discretă pentru aplicația fraților: o **piatră personală** (privată) și o **piatră a Lojii** (comună), care se șlefuiesc de la piatra brută la piatra cubică, cu sigiliul Lojii ieșind în relief.

Macheta aprobată: artifactul „Piatra – machetă” (claude.ai), planșele Acasă, Lucrarea Lojii, Situația mea.

## Imagini
- `piatra/piatra-00.webp` … `piatra-24.webp` (480×480, fundal transparent, ~26 KB fiecare): cadrul i = i lovituri.
- 0–23 = șlefuire progresivă, sigiliul apare în relief bucată cu bucată; **24 = piatra cubică perfectă**.
- Se încarcă doar cadrul curent (și următorul la animație).

## Reguli (an masonic, începe la **1 septembrie** și se resetează la 1 septembrie următor; anul curent 6026–6027, 19 ținute regulare)
Se numără doar evenimentele cu data de azi sau mai veche (marcajele P puse dinainte la o ținută viitoare nu contează până în ziua ținutei).
Lovituri de ciocan pentru un frate:
- prezent (P) la o ținută a Lojii: **1** (din Registrul de prezență)
- vizită la o lojă soră: **3**, max. **3 vizite** numărate pe an (Registrul vizitelor)
- planșă, prezentată în Lojă SAU publicată pe site: **5**, max. **2 planșe** pe an (Registrul planșelor). **Minutul de regularitate NU e planșă** – se exclude (titlu/tip conține „Minutul”).

**Piatra perfectă** = **24 lovituri** ȘI trei condiții obligatorii:
1. prezență minimă 50% (10 din 19 ținute),
2. cel puțin o vizită,
3. cel puțin o planșă.
Peste minim, prezența, vizitele suplimentare și a doua planșă se compensează între ele.

Cadru personal: perfect → 24; altfel `min(lovituri, 23)`.
Trepte (după lovituri/24): <25% Piatră brută · <50% Piatră cioplită · <75% Piatră fasonată · altfel Piatră netezită · perfect = Piatră cubică.

**Piatra Lojii**: suma loviturilor tuturor fraților activi, fiecare frate contribuie cu **max. 24**. Perfectă la **450**. Cadru = 24 dacă ≥450, altfel `min(23, floor(lov/450*24))`. Aceleași nume de trepte.

## Interfață
### Acasă
- Salut deja făcut: „Bună dimineața/ziua/seara, Frate X”.
- Sub cardul următoarei ținute: card **Lucrarea Lojii · 6026–6027** care **înlocuiește secțiunea „Viața lojii”**:
  - stânga: piatra (~150 px); dreapta: numele treptei (Cormorant), „N / 450 lovituri”, bară aurie, „Ultima lovitură: …” (fără nume de frați, ex. „ținuta din 22 oct. · 24 de lovituri”, „vizita de la Oradea · 18 oct.”).
  - rând jos (3 coloane): **Ultima ținută – X% prezență** · **Vizita anuală – N / 36 frați** (link la Loji › Călătoriile fraților, ca actualul card) · **Planșe – N în acest an**.
  - atingerea pietrei = lovitură de ciocan (doar efect); atingerea restului cardului deschide pagina Lucrarea Lojii.
### Pagina Lucrarea Lojii
- piatra mare (~270 px, atingere = lovitură), treapta, N/450, bară, „Până la <treapta următoare>: încă N lovituri”.
- „De unde vin loviturile”: din ținute · din N vizite · din N planșe.
- „Cum se șlefuiește”: 1/3/5 lovituri + „Minutul de regularitate nu intră la planșe.”
- „Ultimele lovituri” (max. 10, fără nume): ținută (dată, nr. frați prezenți, +N), vizite (grupate pe dată+lojă, nr. frați, +3×n), planșe (+5).
- buton „Vezi piatra mea” → Situația mea.
### Situația mea
- Sus, card **Piatra mea · 6026–6027**: piatra (~230 px), treapta, „N / 24 lovituri de ciocan”, bară, mesaj:
  - perfect: „Piatra ta e desăvârșită pentru anul acesta.”
  - altfel: „Mai ai N lovituri de ciocan până la piatra perfectă.” (română: „o lovitură”, „N lovituri”, ≥20 „N de lovituri”); dacă lovituri ≥24 dar lipsesc condiții: „Ai adunat loviturile. Îți mai lipsește …”.
- Lista celor 3 condiții (bifă aurie / cerc gol): Prezență minimă (x din 10 / x din 19), Vizita anuală, Planșa anului (în Lojă sau pentru site).
- Înlocuiește cardurile actuale „Vizite la alte loji” și „Planșele mele”. Privată – o vede doar fratele.

## Animație (sobră, < 1 s)
- La fiecare încărcare a paginii (Acasă, Lucrarea Lojii, Situația mea) piatra primește o lovitură: tresărire ușoară (translate 1–2 px, rotate ±0.5°, 0.32 s), praf de piatră (cercuri difuze bej `rgba(222,208,178,.55)` care cresc și se sting) și câteva așchii (`#b9ad93`, clip-path, cu gravitație), ~0.9 s.
- Dacă cadrul s-a schimbat de la ultima vizită (salvat în localStorage cu try/catch), noul cadru apare peste cel vechi cu fade 0.75 s în timpul loviturii.
- Atingerea pietrei = aceeași lovitură, fără schimbare de cadru.
- La piatra cubică: o strălucire caldă scurtă (halo radial, 1.6 s).
- `prefers-reduced-motion`: fără animație, direct cadrul nou.

## Etape (publicate separat)
1. Piatra Lojii pe Acasă (calcul + card + animație), scoate „Viața lojii”.
2. Pagina Lucrarea Lojii.
3. Piatra personală în Situația mea.
La fiecare publicare: crește versiunea în `index.html` și `version.json`.
