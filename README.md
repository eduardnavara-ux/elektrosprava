# Spotřeba statku — PWA

Webová appka (PWA) pro výpočet nákladů elektřiny statku Dvůr pod Dubem.
Tarif D57d (dvoutarif, 20 h NT), lokality (Dvůr, Stodola, Špýchar, Roubenka),
sezónní přepínač, FVE s prioritami přebytků, napojení na Tapo měřené zásuvky.

---

## 1. Nasazení na Netlify (stejně jako vajíčka)

### A) Nahrání na GitHub
Repo už máš: `eduardnavara-ux/elektrosprava`. Nahraj do něj obsah tohoto balíku.

Z příkazové řádky (ve složce projektu):
```bash
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/eduardnavara-ux/elektrosprava.git
git push -u origin main
```
Nebo přes web: na stránce repa „uploading an existing file" → přetáhni všechny
soubory a složky → Commit.

### B) Propojení s Netlify
1. [app.netlify.com](https://app.netlify.com) → **Add new site → Import an existing project**
2. Vyber GitHub → repo `elektrosprava`
3. Build nastavení se načtou samy z `netlify.toml`:
   - Build command: `npm run build`
   - Publish directory: `dist`
4. **Deploy**. Za chvíli máš veřejnou adresu (např. `elektrosprava.netlify.app`).

### C) Instalace na plochu telefonu
Otevři adresu v telefonu → v menu prohlížeče **Přidat na plochu** / **Install**.
Appka pak běží jako samostatná aplikace (ikona, offline režim).

---

## 2. Napojení Tapo měřených zásuvek (skimmer + dešťovka)

Aby appka brala reálná data místo odhadů, potřebuje je číst z Google Sheetu,
který plní Apps Script (viz složka `apps-script/`).

1. Nastav Apps Script podle `apps-script/NAVOD-gmail-parser.md` — vytvoří
   v Sheetu listy **Skimmer** a **Gardena**.
2. Každý list publikuj jako CSV: v Sheetu **Soubor → Sdílet → Publikovat na
   webu → vyber list → CSV → Publikovat**. Zkopíruj odkaz (končí `output=csv`).
3. V appce (pravý sloupec, karta „Tapo měřené zásuvky") vlož oba odkazy a klikni
   **↻ Načíst data**. U skimmeru a dešťovky se objeví štítek „měřeno" a jejich
   spotřeba se přepočítá z reálných dat.

Odkazy i celé nastavení se ukládají v prohlížeči (localStorage), takže je
zadáš jen jednou.

---

## 3. Lokální vývoj (volitelné)

```bash
npm install
npm run dev      # spustí na http://localhost:5173
npm run build    # vytvoří dist/ k nasazení
```

---

## Struktura

```
elektrosprava/
├─ index.html
├─ package.json
├─ vite.config.js        # Vite + PWA (manifest, service worker)
├─ netlify.toml          # build nastavení pro Netlify
├─ public/
│  ├─ favicon.svg
│  ├─ icon-192.png
│  └─ icon-512.png
├─ src/
│  ├─ main.jsx
│  └─ App.jsx            # celá appka
└─ apps-script/
   ├─ Kod.gs             # Gmail → Sheet (skimmer + dešťovka)
   └─ NAVOD-gmail-parser.md
```

---

## Poznámky

- **Sazby** jsou předvyplněné z faktury PRE (D57d): VT 4,414 / NT 3,299 Kč/kWh.
  Při změně ceníku je přepiš v kartě „Sazby".
- **FVE** je zatím odhad (7 kWp, jih, sklon 44°). Až bude reálná, dolaď výnos,
  ceny a priority přebytků (dům → voda → auto → baterie → síť).
- **Data** se ukládají jen ve tvém prohlížeči. Když appku otevřeš na jiném
  zařízení, nastavení se nepřenáší (řešilo by se to až přihlášením/cloudem).
