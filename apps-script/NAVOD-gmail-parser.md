# Tapo měřené zásuvky z Gmailu → návod na nastavení

Tahle část zařídí, že se data z měřených zásuvek (skimmer, Gardena dešťovka)
dostanou do appky **sama**, bez nahrávání. Princip: skript běží v tvém Google
účtu, jednou denně si vezme z Gmailu poslední Tapo exporty, podle názvu zařízení
je roztřídí a zapíše do Google Sheetu (zvlášť list Skimmer, zvlášť Gardena).
Appka pak ty listy čte.

---

## Část A — Google Sheet + skript

**1. Vytvoř Sheet**
Jdi na [sheets.new](https://sheets.new), pojmenuj ho třeba „Skimmer data".

**2. Otevři editor skriptů**
V Sheetu: nahoře **Rozšíření → Apps Script**. Otevře se nové okno.

**3. Vlož kód**
Smaž ukázkový obsah, vlož celý obsah souboru `Kod.gs`. Ulož (ikona diskety / Ctrl+S).

**4. Zapni Drive API** (skript ho používá k převodu .xls)
V editoru skriptů vlevo: **Služby (Services) → +** → vyber **Drive API** → **Přidat**.
Necháš výchozí verzi.

**5. První spuštění + povolení**
Nahoře vyber funkci `importTapoSkimmer` a klikni **Spustit**.
Google se zeptá na oprávnění (Gmail + Drive + Sheets) — povol. U varování
„Google tuto aplikaci neověřil" klikni **Rozšířené → Přejít na (název) → Povolit**.
Je to tvůj vlastní skript, je to bezpečné.

**6. Ověř**
Pokud máš v Gmailu Tapo mail z posledních 2 dnů, v Sheetu přibude list **Skimmer**
s daty (datum, hodina, kWh). Když ne, pošli si nový export z Tapo a spusť znovu.

**7. Naplánuj denní běh**
V editoru vyber funkci `setupTrigger`, klikni **Spustit**. Hotovo — od teď to
běží každý den ráno kolem 6:00 samo. (Štítek `tapo-zpracovano` v Gmailu zabraňuje
dvojímu zpracování stejného mailu.)

---

## Část B — Propojení s appkou

Appka potřebuje data ze Sheetu číst. Nejjednodušší cesta bez API klíčů:

**1. Publikuj list Skimmer jako CSV**
V Sheetu: **Soubor → Sdílet → Publikovat na webu** →
vyber list **Skimmer**, formát **CSV** → **Publikovat**.
Zkopíruj vzniklý odkaz (končí na `output=csv`).

**2. Vlož odkaz do appky**
V appce přibude pole „CSV odkaz na skimmer data" — vložíš odkaz, appka si data
sama natáhne a nahradí jimi ručně zadaný skimmer. Tohle pole doplním do PWA verze.

---

## Časté potíže

- **Sheet zůstal prázdný** → nemáš v Gmailu žádný Tapo mail za poslední 2 dny.
  Pošli nový export z Tapo aplikace (Energy Monitoring → ikona kalendáře → export).
- **„Drive is not defined"** → nezapnul jsi Drive API (krok A4).
- **Chce to znovu oprávnění** → normální po úpravě skriptu, jen povol.
- **Duplicitní data** → skript dedupuje podle data+hodiny, takže opakované běhy
  stejná data přepíšou, ne znásobí.

---

## Co dál

Tohle je „dodavatel dat". Skript rozlišuje zásuvky podle **Device Name** v těle
mailu (Tapo ho tam píše, např. „Skimmer zásuvka"). Každé zařízení má svůj list:
**Skimmer**, **Gardena**. Až přidáš další měřenou zásuvku, stačí v `Kod.gs`
v sekci `DEVICE_MAP` přidat řádek (část názvu → název listu) a pojmenovat zásuvku
v Tapo tak, aby název seděl. Skript pak data sám zařadí.

Pro vizualizaci napříč více zařízeními by se hodil Home Assistant (Energy
Dashboard) na malém Pi/mini-PC vedle UCG Ultra, ale to je samostatný krok.
