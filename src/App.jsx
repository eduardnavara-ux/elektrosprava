import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Perzistence nastavení (localStorage) ──
const LS_KEY = "elektrosprava_v1";
function loadSaved() {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || {}; } catch { return {}; }
}
function saveState(obj) {
  try { localStorage.setItem(LS_KEY, JSON.stringify(obj)); } catch {}
}

// ── Načtení dat z publikovaného Google Sheetu (CSV) ──
// Sheet musí být publikován: Soubor → Sdílet → Publikovat na webu → list → CSV.
// Vrací { perDayKwh, perDayCost, nDays } z hodinových řádků datum,hodina,kWh.
async function fetchSheetCsv(url, isNT, rates) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status);
  const text = await res.text();
  const lines = text.trim().split(/\r?\n/);
  let vtKwh = 0, ntKwh = 0;
  const days = new Set();
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",");
    if (cols.length < 3) continue;
    const date = cols[0].trim().replace(/^"|"$/g, "");
    const hour = parseInt(cols[1], 10);
    const kwh = parseFloat(cols[2]);
    if (!date || isNaN(hour) || isNaN(kwh)) continue;
    days.add(date);
    if (isNT(hour)) ntKwh += kwh; else vtKwh += kwh;
  }
  const nDays = Math.max(days.size, 1);
  return {
    nDays,
    perDayKwh: (vtKwh + ntKwh) / nDays,
    perDayCost: (vtKwh * rates.vtPerKwh + ntKwh * rates.ntPerKwh) / nDays,
  };
}

// Sazby vč. DPH 21 % — z faktury PRE 186630247, část B (D57d).
const DEFAULT_RATES = { vtPerKwh: 4.414, ntPerKwh: 3.299, jisticPerMonth: 671.55, omPerMonth: 143.99 };

// NT harmonogram D57d — reálné spínací časy ČEZ (povel a3b7dp01).
// Liší se všední den vs víkend. Krátký blok 00:00–00:16 zanedbán (zaokrouhleno).
const NT_WEEKDAY = [
  { from: 1.17, to: 7.93 },   // 01:10–07:56
  { from: 8.92, to: 13.27 },  // 08:55–13:16
  { from: 14.17, to: 19.18 }, // 14:10–19:11
  { from: 20.17, to: 24 },    // 20:10–24:00
];
const NT_WEEKEND = [
  { from: 1.17, to: 8.25 },   // 01:10–08:15
  { from: 9.25, to: 13.27 },  // 09:15–13:16
  { from: 14.17, to: 19.18 }, // 14:10–19:11
  { from: 20.17, to: 24 },    // 20:10–24:00
];
const DEFAULT_NT = NT_WEEKDAY;

// Lokality statku
const PLACES = ["Dvůr & zahrada", "Stodola", "Špýchar", "Roubenka"];

// Inventář. place = lokalita, group = kategorie, type = cont|burst, season = all|summer|winter
const DEFAULT_DEVICES = [
  // ── Dvůr & zahrada ──
  { id: 1, name: "3× Tapo C530WS", place: "Dvůr & zahrada", group: "Kamery & síť", type: "cont", watts: 11.7, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 2, name: "2× Tapo C560WS", place: "Dvůr & zahrada", group: "Kamery & síť", type: "cont", watts: 5, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 3, name: "2× Tapo C325WB", place: "Dvůr & zahrada", group: "Kamery & síť", type: "cont", watts: 6, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 4, name: "UniFi (UCG + 2 sw + 3 AP)", place: "Dvůr & zahrada", group: "Kamery & síť", type: "cont", watts: 35, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 5, name: "Homee + pohon brány (klid)", place: "Dvůr & zahrada", group: "Kamery & síť", type: "cont", watts: 4, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 17, name: "ČOV Asio Monocomp 4", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 50, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 18, name: "Ohradník AKO Ni 7000", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 12, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 27, name: "6× venkovní LED s čidlem", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 60, hoursPerDay: 4, ntShare: 70, season: "all" },
  { id: 19, name: "Napáječka SH 30 RBH", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 80, hoursPerDay: 12, ntShare: 83, season: "winter" },
  { id: 28, name: "Čerpadlo Wilo (oběh TV)", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 40, hoursPerDay: 12, ntShare: 83, season: "winter" },
  { id: 30, name: "Lednice Smeg (léto)", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 25, hoursPerDay: 24, ntShare: 83, season: "summer" },
  { id: 31, name: "Fosi Audio (málo)", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 10, hoursPerDay: 1, ntShare: 50, season: "all" },
  { id: 20, name: "Dvířka kurník", place: "Dvůr & zahrada", group: "Provoz", type: "burst", watts: 10, minPerCycle: 1, cyclesPerMonth: 60, ntShare: 50, season: "all" },
  { id: 21, name: "Brána Hörmann (chod)", place: "Dvůr & zahrada", group: "Provoz", type: "burst", watts: 250, minPerCycle: 1, cyclesPerMonth: 120, ntShare: 50, season: "all" },
  { id: 22, name: "Čerpadlo vrt 30 m", place: "Dvůr & zahrada", group: "Provoz", type: "burst", watts: 1100, minPerCycle: 10, cyclesPerMonth: 90, ntShare: 50, season: "all" },
  { id: 23, name: "Gardena 6100 dešťovka (Tapo měřená)", place: "Dvůr & zahrada", group: "Provoz", type: "burst", watts: 1100, minPerCycle: 30, cyclesPerMonth: 20, ntShare: 60, season: "summer" },
  { id: 33, name: "Skimmer bazén (Tapo měřený)", place: "Dvůr & zahrada", group: "Provoz", type: "cont", watts: 49, hoursPerDay: 10, ntShare: 85, season: "summer" },
  // ── Stodola (hlavní domácnost) ──
  { id: 6, name: "Lednice Bosch KGN39", place: "Stodola", group: "Domácnost", type: "cont", watts: 25, hoursPerDay: 24, ntShare: 83, season: "all" },
  { id: 14, name: "Denon PMA-600NE", place: "Stodola", group: "Domácnost", type: "cont", watts: 15, hoursPerDay: 3, ntShare: 50, season: "all" },
  { id: 15, name: "Osvětlení (LED)", place: "Stodola", group: "Domácnost", type: "cont", watts: 80, hoursPerDay: 5, ntShare: 60, season: "all" },
  { id: 25, name: "HP EliteBook", place: "Stodola", group: "Domácnost", type: "cont", watts: 30, hoursPerDay: 6, ntShare: 40, season: "all" },
  { id: 26, name: "MacBook Air", place: "Stodola", group: "Domácnost", type: "cont", watts: 14, hoursPerDay: 5, ntShare: 50, season: "all" },
  { id: 7, name: "Pračka Hoover HDPD696", place: "Stodola", group: "Domácnost", type: "burst", watts: 1170, minPerCycle: 60, cyclesPerMonth: 18, ntShare: 60, season: "all" },
  { id: 8, name: "Myčka Gorenje ULTRA16", place: "Stodola", group: "Domácnost", type: "burst", watts: 1320, minPerCycle: 30, cyclesPerMonth: 30, ntShare: 60, season: "all" },
  { id: 9, name: "Bojler Dražice 125 l (léto)", place: "Stodola", group: "Domácnost", type: "burst", watts: 2000, minPerCycle: 60, cyclesPerMonth: 30, ntShare: 80, season: "summer" },
  { id: 10, name: "Konvice", place: "Stodola", group: "Domácnost", type: "burst", watts: 2000, minPerCycle: 3, cyclesPerMonth: 75, ntShare: 40, season: "all" },
  { id: 11, name: "Vařič Sencor dvouplotýnka", place: "Stodola", group: "Domácnost", type: "burst", watts: 2900, minPerCycle: 30, cyclesPerMonth: 15, ntShare: 50, season: "all" },
  { id: 12, name: "Robot Roborock (2×/týd)", place: "Stodola", group: "Domácnost", type: "burst", watts: 300, minPerCycle: 90, cyclesPerMonth: 9, ntShare: 70, season: "all" },
  { id: 13, name: "Vysavač ETA (2×/týd)", place: "Stodola", group: "Domácnost", type: "burst", watts: 600, minPerCycle: 20, cyclesPerMonth: 9, ntShare: 30, season: "all" },
  { id: 16, name: "Nabíjení 2× iPhone", place: "Stodola", group: "Domácnost", type: "burst", watts: 20, minPerCycle: 120, cyclesPerMonth: 30, ntShare: 70, season: "all" },
  { id: 24, name: "Práce – pila, nářadí", place: "Stodola", group: "Provoz", type: "burst", watts: 1500, minPerCycle: 30, cyclesPerMonth: 8, ntShare: 30, season: "all" },
  { id: 29, name: "Podlahové DEVI 20 m²", place: "Stodola", group: "Topení", type: "cont", watts: 2000, hoursPerDay: 10, ntShare: 85, season: "winter" },
  // ── Špýchar (babička, ve výstavbě) ──
  { id: 32, name: "Podlahové DEVI koupelna ~7 m²", place: "Špýchar", group: "Topení", type: "cont", watts: 700, hoursPerDay: 8, ntShare: 85, season: "winter" },
  // ── Roubenka — zatím prázdné ──
];

const GROUPS = ["Domácnost", "Provoz", "Topení", "Kamery & síť"];
const fmt = (n, d = 2) => (isFinite(n) ? n : 0).toLocaleString("cs-CZ", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtKc = (n) => fmt(n, 0) + " Kč";

function parseTapo(arrayBuffer, fileName) {
  let wb;
  try { wb = XLSX.read(arrayBuffer, { type: "array", cellDates: false }); }
  catch (e) { return { error: "Soubor se nepodařilo přečíst (" + fileName + ")." }; }
  const sheetName = wb.SheetNames.find((n) => /day/i.test(n)) || wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
  const out = []; const unit = /power/i.test(fileName) ? "W" : "kWh";
  for (const r of rows) {
    if (!r || r.length < 2) continue;
    // formát "2026/06/16 01:00:00" — vyžaduj sekundy, ať nechytneme řádek s rozsahem
    const m = String(r[0] ?? "").trim().match(/^(\d{4})\/(\d{2})\/(\d{2})\s+(\d{1,2}):(\d{2}):(\d{2})\s*$/);
    if (!m) continue;
    const val = parseFloat(String(r[1]).replace(",", "."));
    if (isNaN(val)) continue;
    out.push({ date: `${m[1]}-${m[2]}-${m[3]}`, hour: +m[4] + +m[5] / 60, val });
  }
  return { unit, data: out };
}

export default function App() {
  const saved = loadSaved();
  const [rates, setRates] = useState(saved.rates || DEFAULT_RATES);
  const [ntBlocks, setNtBlocks] = useState(saved.ntBlocks || DEFAULT_NT);
  const [dayType, setDayType] = useState("weekday"); // weekday | weekend
  const [tapoFiles, setTapoFiles] = useState([]);
  const [devices, setDevices] = useState(saved.devices || DEFAULT_DEVICES);
  const [season, setSeason] = useState(saved.season || "summer");
  const [activePlace, setActivePlace] = useState("Celkem");
  const [openGroups, setOpenGroups] = useState({ "Domácnost": false, "Provoz": false, "Topení": false, "Kamery & síť": false });
  const [fve, setFve] = useState(saved.fve || {
    enabled: false,
    kwp: 7,
    yieldPerKwp: 980,
    sellPrice: 2,
    installCost: 200000,
    prio: {
      house: true,
      water: { on: true, capKwh: 7, save: 0 },
      car: { on: false, capKwh: 15 },
      battery: { on: false, capKwh: 10 },
    },
  });
  // napojení na Google Sheety (publikované CSV) pro měřené zásuvky
  const [sheets, setSheets] = useState(saved.sheets || { skimmer: "", gardena: "" });
  const [measured, setMeasured] = useState({}); // {skimmer:{perDayKwh,...}, gardena:{...}}
  const [syncStatus, setSyncStatus] = useState("");
  const [err, setErr] = useState("");
  const fileRef = useRef();

  // ulož nastavení při změně
  useEffect(() => {
    saveState({ rates, ntBlocks, devices, season, fve, sheets });
  }, [rates, ntBlocks, devices, season, fve, sheets]);

  const isNT = (hour) => ntBlocks.some((b) => hour >= b.from && hour < b.to);
  const ntHours = useMemo(() => ntBlocks.reduce((s, b) => s + (b.to - b.from), 0), [ntBlocks]);
  const activeInSeason = (d) => d.season === "all" || d.season === season;

  const handleFiles = async (files) => {
    setErr("");
    for (const file of files) {
      const buf = await file.arrayBuffer();
      const res = parseTapo(new Uint8Array(buf), file.name);
      if (res.error) { setErr(res.error); continue; }
      if (!res.data.length) { setErr(`V ${file.name} jsem nenašel hodinová data.`); continue; }
      setTapoFiles((p) => [...p, { name: file.name.replace(/\.[^.]+$/, ""), unit: res.unit, data: res.data }]);
    }
  };

  // načti reálná data z obou Sheetů
  const syncSheets = async () => {
    setSyncStatus("Načítám…");
    const next = {};
    for (const key of ["skimmer", "gardena"]) {
      const url = sheets[key];
      if (!url) continue;
      try {
        next[key] = await fetchSheetCsv(url, isNT, rates);
      } catch (e) {
        setSyncStatus("Chyba u " + key + ": " + e.message);
        return;
      }
    }
    setMeasured(next);
    const parts = [];
    if (next.skimmer) parts.push(`skimmer ${fmt(next.skimmer.perDayKwh, 2)} kWh/den`);
    if (next.gardena) parts.push(`dešťovka ${fmt(next.gardena.perDayKwh, 2)} kWh/den`);
    setSyncStatus(parts.length ? "✓ " + parts.join(" · ") : "Žádný odkaz nevyplněn.");
  };

  // auto-sync při startu, pokud jsou odkazy
  useEffect(() => {
    if (sheets.skimmer || sheets.gardena) syncSheets();
    // eslint-disable-next-line
  }, []);

  const blendedPrice = (ntShare) => ((100 - ntShare) * rates.vtPerKwh + ntShare * rates.ntPerKwh) / 100;

  // mapování měřených dat na spotřebiče (podle id: 33 = skimmer, 23 = gardena)
  const measuredFor = (d) => {
    if (d.id === 33 && measured.skimmer) return measured.skimmer;
    if (d.id === 23 && measured.gardena) return measured.gardena;
    return null;
  };

  const calcDevice = (d) => {
    const meas = measuredFor(d);
    if (meas) {
      // reálná data nahradí odhad: použij naměřenou denní spotřebu a cenu
      const dayKwh = activeInSeason(d) ? meas.perDayKwh : 0;
      const cost = activeInSeason(d) ? meas.perDayCost : 0;
      return { dayKwh, cost, cycleKwh: dayKwh, cycleCost: cost, measured: true };
    }
    let dayKwh = d.type === "cont" ? (d.watts * d.hoursPerDay) / 1000 : (d.watts * (d.minPerCycle / 60) * (d.cyclesPerMonth / 30)) / 1000;
    if (!activeInSeason(d)) dayKwh = 0;
    const ntKwh = (dayKwh * d.ntShare) / 100;
    const cost = (dayKwh - ntKwh) * rates.vtPerKwh + ntKwh * rates.ntPerKwh;
    const price = blendedPrice(d.ntShare);
    const cycleKwh = d.type === "burst" ? (d.watts * (d.minPerCycle / 60)) / 1000 : dayKwh;
    return { dayKwh, cost, cycleKwh, cycleCost: cycleKwh * price };
  };

  const tapo = useMemo(() => {
    let vtKwh = 0, ntKwh = 0; const days = new Set();
    tapoFiles.forEach((d) => { if (d.unit !== "kWh") return;
      d.data.forEach((h) => { days.add(h.date); if (isNT(h.hour)) ntKwh += h.val; else vtKwh += h.val; }); });
    const nDays = Math.max(days.size, 1);
    return { totalKwh: vtKwh + ntKwh, nDays, perDayKwh: (vtKwh + ntKwh) / nDays,
      perDayCost: (vtKwh * rates.vtPerKwh + ntKwh * rates.ntPerKwh) / nDays };
  }, [tapoFiles, ntBlocks, rates]);

  // souhrn po lokalitách
  const placeTotals = useMemo(() => {
    const t = {}; PLACES.forEach((p) => { t[p] = { dayKwh: 0, cost: 0, count: 0 }; });
    devices.forEach((d) => {
      const c = calcDevice(d);
      if (!t[d.place]) t[d.place] = { dayKwh: 0, cost: 0, count: 0 };
      t[d.place].dayKwh += c.dayKwh; t[d.place].cost += c.cost; t[d.place].count += 1;
    });
    return t;
  }, [devices, rates, season]);

  const devTotals = useMemo(() => {
    let dayKwh = 0, cost = 0;
    devices.forEach((d) => { const c = calcDevice(d); dayKwh += c.dayKwh; cost += c.cost; });
    return { dayKwh, cost };
  }, [devices, rates, season]);

  const monthly = useMemo(() => {
    const dailyKwh = tapo.perDayKwh + devTotals.dayKwh;
    const dailyCost = tapo.perDayCost + devTotals.cost;
    const fixedM = rates.jisticPerMonth + rates.omPerMonth;
    return { kwh: dailyKwh * 30, variable: dailyCost * 30, fixed: fixedM, total: dailyCost * 30 + fixedM };
  }, [tapo, devTotals, rates]);

  // ── FVE: výroba → kaskáda priorit (dům → voda → auto → baterie → síť) ──
  const fveCalc = useMemo(() => {
    if (!fve.enabled) return null;
    const yearKwh = fve.kwp * fve.yieldPerKwp;
    const monthShare = season === "summer" ? 0.12 : 0.035;   // léto ~12 %, zima ~3,5 % roční výroby/měs
    const prodMonth = yearKwh * monthShare;
    const prodDay = prodMonth / 30;
    const buyPrice = (rates.vtPerKwh + rates.ntPerKwh) / 2;   // průměrná nákupní cena (co ušetřím)

    // denní spotřeba domu, kolik z ní padne do FVE výrobních hodin (~30 % spotřeby je v "denním okně")
    const consDay = monthly.kwh / 30;
    const dayWindowCons = consDay * 0.35;                     // spotřeba v době, kdy FVE vyrábí

    let remaining = prodDay;                                  // co zbývá z výroby rozdělit
    const steps = [];

    // 1. okamžitá spotřeba domu
    const houseUse = Math.min(remaining, dayWindowCons);
    remaining -= houseUse;
    steps.push({ key: "house", label: "Dům (přímá spotřeba)", kwh: houseUse, save: houseUse * buyPrice, on: true });

    // 2. ohřev vody (bojler jako tepelná baterie)
    if (fve.prio.water.on) {
      const use = Math.min(remaining, fve.prio.water.capKwh);
      remaining -= use;
      steps.push({ key: "water", label: "Ohřev vody (bojler)", kwh: use, save: use * buyPrice, on: true });
    }
    // 3. baterie auta
    if (fve.prio.car.on) {
      const use = Math.min(remaining, fve.prio.car.capKwh);
      remaining -= use;
      steps.push({ key: "car", label: "Baterie auta", kwh: use, save: use * buyPrice, on: true });
    }
    // 4. domácí baterie
    if (fve.prio.battery.on) {
      const use = Math.min(remaining, fve.prio.battery.capKwh);
      remaining -= use;
      steps.push({ key: "battery", label: "Domácí baterie", kwh: use, save: use * buyPrice, on: true });
    }
    // 5. přetok do sítě (zbytek)
    const surplusDay = Math.max(remaining, 0);
    steps.push({ key: "grid", label: "Přetok do sítě", kwh: surplusDay, save: surplusDay * fve.sellPrice, on: true });

    const selfDay = prodDay - surplusDay;                     // vše kromě přetoku = vlastní spotřeba
    const saveDay = steps.reduce((s, x) => s + x.save, 0);
    const saveMonth = saveDay * 30;
    const newBill = Math.max(monthly.total - saveMonth, rates.jisticPerMonth + rates.omPerMonth);
    const payback = fve.installCost > 0 && saveMonth > 0 ? fve.installCost / (saveMonth * 9) : 0; // hrubě: ~9 "plných" měsíců/rok

    return {
      prodMonth, selfMonth: selfDay * 30, surplusMonth: surplusDay * 30,
      selfRate: prodDay > 0 ? selfDay / prodDay : 0,
      steps: steps.map((s) => ({ ...s, kwh: s.kwh * 30, save: s.save * 30 })),
      saveMonth, newBill, payback,
    };
  }, [fve, season, monthly, rates]);

  const upd = (id, f, v) => setDevices((p) => p.map((d) => (d.id === id ? { ...d, [f]: v } : d)));

  // spotřebiče k zobrazení podle aktivní lokality
  const shownDevices = activePlace === "Celkem" ? devices : devices.filter((d) => d.place === activePlace);
  const shownByGroup = useMemo(() => {
    const g = {};
    GROUPS.forEach((name) => { g[name] = []; });
    shownDevices.forEach((d) => { (g[d.group] || (g[d.group] = [])).push({ ...d, ...calcDevice(d) }); });
    return g;
  }, [shownDevices, rates, season]);

  const tabs = ["Celkem", ...PLACES, "☀ FVE"];

  return (
    <div style={S.page}>
      <style>{CSS}</style>
      <header style={S.header}>
        <div style={S.headTop}>
          <div>
            <div style={S.eyebrow}>Dvůr pod Dubem · elektřina</div>
            <h1 style={S.h1}>Spotřeba statku</h1>
          </div>
          <div style={S.seasonTog}>
            <button style={{ ...S.seasonBtn, ...(season === "summer" ? S.seasonOn : {}) }} onClick={() => setSeason("summer")}>Léto</button>
            <button style={{ ...S.seasonBtn, ...(season === "winter" ? S.seasonOn : {}) }} onClick={() => setSeason("winter")}>Zima</button>
          </div>
        </div>
      </header>

      <section style={S.summaryRow}>
        <Stat label="Měsíční spotřeba" value={fmt(monthly.kwh, 0) + " kWh"} accent />
        <Stat label="Variabilní / měs" value={fmtKc(monthly.variable)} />
        <Stat label="Stálé platy / měs" value={fmtKc(monthly.fixed)} />
        {fveCalc ? (
          <Stat label="Účet s FVE / měs" value={fmtKc(fveCalc.newBill)} big />
        ) : (
          <Stat label="Predikce zálohy / měs" value={fmtKc(monthly.total)} big />
        )}
      </section>

      {fveCalc && (
        <section style={S.fveSaveBar}>
          <span style={S.fveSaveText}>☀ FVE ušetří <strong>{fmtKc(fveCalc.saveMonth)}/měs</strong> ({season === "summer" ? "léto" : "zima"})</span>
          <span style={S.fveSaveDetail}>vlastní spotřeba {fmt(fveCalc.selfRate * 100, 0)} % · přetok {fmt(fveCalc.surplusMonth, 0)} kWh/měs</span>
        </section>
      )}

      {/* záložky lokalit */}
      <div style={S.tabs}>
        {tabs.map((t) => {
          if (t === "☀ FVE") {
            return (
              <button key={t} style={{ ...S.tab, ...(activePlace === t ? S.tabOn : {}), ...(activePlace !== t ? S.tabFve : {}) }} onClick={() => setActivePlace(t)}>
                <span style={{ ...S.tabName, ...(activePlace === t ? { color: "#fff" } : { color: GOLD }) }}>☀ FVE</span>
                <span style={{ ...S.tabVal, ...(activePlace === t ? { color: "#cfe0d3" } : {}) }}>{fveCalc ? fmtKc(fveCalc.saveMonth) + "/měs" : "nastav"}</span>
              </button>
            );
          }
          const pt = t === "Celkem" ? { dayKwh: devTotals.dayKwh, cost: devTotals.cost } : placeTotals[t];
          return (
            <button key={t} style={{ ...S.tab, ...(activePlace === t ? S.tabOn : {}) }} onClick={() => setActivePlace(t)}>
              <span style={{ ...S.tabName, ...(activePlace === t ? { color: "#fff" } : {}) }}>{t}</span>
              <span style={{ ...S.tabVal, ...(activePlace === t ? { color: "#cfe0d3" } : {}) }}>{fmtKc((pt?.cost || 0) * 30)}/měs</span>
            </button>
          );
        })}
      </div>

      {activePlace === "☀ FVE" ? (
        <FveMode fve={fve} setFve={setFve} fveCalc={fveCalc} season={season} fmt={fmt} fmtKc={fmtKc} S={S} RateRow={RateRow} />
      ) : (
      <div style={S.grid}>
        <div>
          {GROUPS.map((gName) => {
            const list = shownByGroup[gName];
            if (!list || !list.length) return null;
            const gKwh = list.reduce((s, d) => s + d.dayKwh, 0);
            const gCost = list.reduce((s, d) => s + d.cost, 0);
            const open = openGroups[gName];
            return (
              <div key={gName} style={S.card}>
                <button style={S.collapseHead} onClick={() => setOpenGroups((p) => ({ ...p, [gName]: !p[gName] }))}>
                  <span style={S.collapseTitle}>
                    <span style={{ ...S.chevron, transform: open ? "rotate(90deg)" : "rotate(0deg)" }}>›</span>
                    {gName}
                    <span style={S.collapseCount}>{list.length}</span>
                  </span>
                  <span style={S.cardHint}>{fmt(gKwh, 2)} kWh/den · {fmtKc(gCost * 30)}/měs</span>
                </button>
                {open && (
                  <div style={{ ...S.devList, marginTop: 14 }}>
                    {list.map((d) => (
                      <div key={d.id} style={{ ...S.devCard, opacity: activeInSeason(d) ? 1 : 0.4 }}>
                        <div style={S.devTop}>
                          <input style={S.devName} value={d.name} onChange={(e) => upd(d.id, "name", e.target.value)} />
                          {activePlace === "Celkem" && <span style={S.placeTag}>{d.place.split(" ")[0]}</span>}
                          {d.measured && <span style={S.measuredTag}>měřeno</span>}
                          {d.season !== "all" && <span style={S.seasonTag}>{d.season === "summer" ? "léto" : "zima"}</span>}
                          <span style={S.devCost}>{fmtKc(d.cost * 30)}<span style={S.devCostUnit}>/měs</span></span>
                          <button style={S.x} onClick={() => setDevices((p) => p.filter((x) => x.id !== d.id))}>×</button>
                        </div>
                        <div style={S.devFields}>
                          <Field label="Příkon"><input style={S.inS} type="number" value={d.watts} onChange={(e) => upd(d.id, "watts", +e.target.value)} /><span style={S.u}>W</span></Field>
                          {d.type === "cont" ? (
                            <Field label="Provoz"><input style={S.inS} type="number" value={d.hoursPerDay} onChange={(e) => upd(d.id, "hoursPerDay", +e.target.value)} /><span style={S.u}>h/den</span></Field>
                          ) : (
                            <>
                              <Field label="Délka"><input style={S.inS} type="number" value={d.minPerCycle} onChange={(e) => upd(d.id, "minPerCycle", +e.target.value)} /><span style={S.u}>min</span></Field>
                              <Field label="Cykly"><input style={S.inS} type="number" value={d.cyclesPerMonth} onChange={(e) => upd(d.id, "cyclesPerMonth", +e.target.value)} /><span style={S.u}>/měs</span></Field>
                            </>
                          )}
                          <Field label="NT %"><input style={S.inS} type="number" value={d.ntShare} onChange={(e) => upd(d.id, "ntShare", +e.target.value)} /><span style={S.u}>%</span></Field>
                        </div>
                        <div style={S.devUnit}>
                          {d.type === "burst" ? `${fmt(d.cycleKwh, 3)} kWh / cyklus · ${fmt(d.cycleCost, 2)} Kč / cyklus` : `${fmt(d.cycleKwh, 3)} kWh / den · ${fmt(d.cycleCost, 2)} Kč / den`}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
          {shownDevices.length === 0 && <Card title={activePlace} hint="Zatím prázdné"><p style={S.empty}>V této lokalitě zatím nejsou žádné spotřebiče. Přidej je tlačítkem níže.</p></Card>}
          <div style={S.addRow}>
            <button style={S.addBtn} onClick={() => setDevices((p) => [...p, { id: Date.now(), name: "Trvalý spotřebič", place: activePlace === "Celkem" ? "Stodola" : activePlace, group: "Provoz", type: "cont", watts: 100, hoursPerDay: 24, ntShare: 83, season: "all" }])}>+ Trvalý</button>
            <button style={S.addBtn} onClick={() => setDevices((p) => [...p, { id: Date.now(), name: "Nárazový spotřebič", place: activePlace === "Celkem" ? "Stodola" : activePlace, group: "Provoz", type: "burst", watts: 1000, minPerCycle: 10, cyclesPerMonth: 30, ntShare: 50, season: "all" }])}>+ Nárazový</button>
          </div>
        </div>

        <div>
          <Card title="Tapo měřené zásuvky" hint="Skimmer + Gardena dešťovka — reálná data nahradí odhady">
            <div style={S.sheetBox}>
              <label style={S.sheetLabel}>Skimmer — CSV odkaz na Google Sheet</label>
              <input style={S.sheetInput} placeholder="https://…/pub?gid=…&output=csv" value={sheets.skimmer}
                onChange={(e) => setSheets((s) => ({ ...s, skimmer: e.target.value.trim() }))} />
              <label style={S.sheetLabel}>Gardena dešťovka — CSV odkaz</label>
              <input style={S.sheetInput} placeholder="https://…/pub?gid=…&output=csv" value={sheets.gardena}
                onChange={(e) => setSheets((s) => ({ ...s, gardena: e.target.value.trim() }))} />
              <div style={S.sheetActions}>
                <button style={S.syncBtn} onClick={syncSheets}>↻ Načíst data</button>
                {syncStatus && <span style={S.syncStatus}>{syncStatus}</span>}
              </div>
            </div>

            <div style={S.orDivider}>nebo nahraj soubor ručně</div>
            <div style={S.drop} onClick={() => fileRef.current?.click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => { e.preventDefault(); handleFiles(e.dataTransfer.files); }}>
              <input ref={fileRef} type="file" multiple accept=".xls,.xlsx,.csv" style={{ display: "none" }} onChange={(e) => handleFiles(e.target.files)} />
              <span style={S.dropPlus}>+</span><span>Nahraj .xls</span>
            </div>
            {err && <p style={S.errBox}>{err}</p>}
            {tapoFiles.length > 0 && (
              <div style={{ marginTop: 12 }}>
                {tapoFiles.map((d, i) => {
                  const total = d.data.reduce((s, h) => s + h.val, 0);
                  return (
                    <div key={i} style={S.fileRow}>
                      <span>{d.name}</span>
                      <span style={S.fileVal}>{d.unit === "kWh" ? fmt(total, 3) + " kWh" : fmt(total / d.data.length, 0) + " W ø"}</span>
                      <button style={S.x} onClick={() => setTapoFiles((p) => p.filter((_, idx) => idx !== i))}>×</button>
                    </div>
                  );
                })}
                {tapo.totalKwh > 0 && <div style={S.perDay}>{tapo.nDays} dní · ø {fmt(tapo.perDayKwh, 3)} kWh/den · {fmtKc(tapo.perDayCost * 30)}/měs</div>}
              </div>
            )}
          </Card>

          <Card title={`Nízký tarif — ${fmt(ntHours, 1)} h/den`} hint="D57d, reálné časy ČEZ">
            <div style={S.dayTypeTog}>
              <button style={{ ...S.dayTypeBtn, ...(dayType === "weekday" ? S.dayTypeOn : {}) }}
                onClick={() => { setDayType("weekday"); setNtBlocks(NT_WEEKDAY); }}>Po–Pá</button>
              <button style={{ ...S.dayTypeBtn, ...(dayType === "weekend" ? S.dayTypeOn : {}) }}
                onClick={() => { setDayType("weekend"); setNtBlocks(NT_WEEKEND); }}>Víkend</button>
            </div>
            {ntBlocks.map((b, i) => (
              <div key={i} style={S.ntRow}>
                <span style={S.ntDot} />
                <input style={S.inTime} type="number" step="0.25" value={b.from} onChange={(e) => setNtBlocks((p) => p.map((x, idx) => idx === i ? { ...x, from: +e.target.value } : x))} />
                <span style={S.dash}>–</span>
                <input style={S.inTime} type="number" step="0.25" value={b.to} onChange={(e) => setNtBlocks((p) => p.map((x, idx) => idx === i ? { ...x, to: +e.target.value } : x))} />
                <span style={S.u}>h</span>
                <button style={S.x} onClick={() => setNtBlocks((p) => p.filter((_, idx) => idx !== i))}>×</button>
              </div>
            ))}
            <button style={S.addBtn} onClick={() => setNtBlocks((p) => [...p, { from: 0, to: 1 }])}>+ NT blok</button>
            <div style={S.dayStrip}>
              {Array.from({ length: 24 }, (_, h) => (
                <div key={h} title={`${String(h).padStart(2, "0")}:00 — ${isNT(h + 0.5) ? "NT" : "VT"}`}
                  style={{ ...S.dayCell, background: isNT(h + 0.5) ? "#2d6a4f" : "#e8b04b" }} />
              ))}
            </div>
            <div style={S.legend}>
              <span style={S.legendItem}><i style={{ ...S.legendDot, background: "#e8b04b" }} /> VT (vysoký tarif)</span>
              <span style={S.legendItem}><i style={{ ...S.legendDot, background: "#2d6a4f" }} /> NT (nízký tarif)</span>
            </div>
          </Card>

          <Card title="Sazby" hint="Z faktury PRE">
            <RateRow label="VT — vysoký tarif" unit="Kč/kWh" value={rates.vtPerKwh} onChange={(v) => setRates((r) => ({ ...r, vtPerKwh: v }))} />
            <RateRow label="NT — nízký tarif" unit="Kč/kWh" value={rates.ntPerKwh} onChange={(v) => setRates((r) => ({ ...r, ntPerKwh: v }))} />
            <RateRow label="Jistič 3×25 A" unit="Kč/měs" value={rates.jisticPerMonth} onChange={(v) => setRates((r) => ({ ...r, jisticPerMonth: v }))} />
            <RateRow label="Odběrné místo" unit="Kč/měs" value={rates.omPerMonth} onChange={(v) => setRates((r) => ({ ...r, omPerMonth: v }))} />
          </Card>
        </div>
      </div>
      )}

      <footer style={S.footer}>Rozděleno po budovách (Stodola = hlavní domácnost, Špýchar ve výstavbě pro babičku). Přepínač Léto/Zima nahoře mění sezónní spotřebiče. DEVI: 20 m² stodola + ~7 m² koupelna špýchar.</footer>
    </div>
  );
}

function FveMode({ fve, setFve, fveCalc, season, fmt, fmtKc, S, RateRow }) {
  const set = (patch) => setFve((f) => ({ ...f, ...patch }));
  const setPrio = (key, patch) => setFve((f) => ({ ...f, prio: { ...f.prio, [key]: { ...f.prio[key], ...patch } } }));
  return (
    <div>
      {!fve.enabled ? (
        <div style={S.fveIntro}>
          <h2 style={S.fveIntroTitle}>☀ Fotovoltaika</h2>
          <p style={S.fveIntroText}>
            Statek má jižní střechu Špýcharu (sklon 44°) ideální pro panely. Zapni FVE
            a appka spočítá výrobu, rozdělení přebytků podle priorit (dům → voda → auto →
            baterie → síť) a úsporu. Zatím na odhadech — až bude instalace reálná, dolaď čísla.
          </p>
          <label style={S.fveBigToggle}>
            <input type="checkbox" checked={fve.enabled} onChange={(e) => set({ enabled: e.target.checked })} />
            <span>Mám / plánuji FVE</span>
          </label>
        </div>
      ) : (
        <>
          {fveCalc && (
            <section style={S.summaryRow}>
              <Stat label={`Výroba (${season === "summer" ? "léto" : "zima"})`} value={fmt(fveCalc.prodMonth, 0) + " kWh"} accent />
              <Stat label="Vlastní spotřeba" value={fmt(fveCalc.selfRate * 100, 0) + " %"} />
              <Stat label="Přetok do sítě" value={fmt(fveCalc.surplusMonth, 0) + " kWh"} />
              <Stat label="Úspora / měs" value={fmtKc(fveCalc.saveMonth)} big />
            </section>
          )}

          <div style={S.grid}>
            <div>
              <Card title="Parametry FVE" hint="Jih, sklon 44°">
                <label style={S.fveToggle}>
                  <input type="checkbox" checked={fve.enabled} onChange={(e) => set({ enabled: e.target.checked })} />
                  <span>FVE aktivní</span>
                </label>
                <div style={{ marginTop: 12 }}>
                  <RateRow label="Výkon panelů" unit="kWp" value={fve.kwp} onChange={(v) => set({ kwp: v })} />
                  <RateRow label="Roční výnos" unit="kWh/kWp" value={fve.yieldPerKwp} onChange={(v) => set({ yieldPerKwp: v })} />
                  <RateRow label="Výkupní cena přetoků" unit="Kč/kWh" value={fve.sellPrice} onChange={(v) => set({ sellPrice: v })} />
                  <RateRow label="Cena instalace" unit="Kč" value={fve.installCost} onChange={(v) => set({ installCost: v })} />
                </div>
              </Card>

              <Card title="Priority přebytků" hint="Kam teče vyrobená elektřina">
                <div style={S.prioNote}>Vyrobená elektřina se rozdělí v tomto pořadí. Co zbude, jde do sítě za výkupní cenu.</div>
                <div style={S.prioRow}>
                  <span style={S.prioNum}>1</span>
                  <span style={S.prioLabel}>Dům (přímá spotřeba)</span>
                  <span style={S.prioFixed}>vždy</span>
                </div>
                <div style={S.prioRow}>
                  <span style={S.prioNum}>2</span>
                  <label style={S.prioCheck}><input type="checkbox" checked={fve.prio.water.on} onChange={(e) => setPrio("water", { on: e.target.checked })} /><span>Ohřev vody (bojler)</span></label>
                  <input style={S.prioCap} type="number" value={fve.prio.water.capKwh} onChange={(e) => setPrio("water", { capKwh: +e.target.value })} /><span style={S.u}>kWh/d</span>
                </div>
                <div style={S.prioRow}>
                  <span style={S.prioNum}>3</span>
                  <label style={S.prioCheck}><input type="checkbox" checked={fve.prio.car.on} onChange={(e) => setPrio("car", { on: e.target.checked })} /><span>Baterie auta</span></label>
                  <input style={S.prioCap} type="number" value={fve.prio.car.capKwh} onChange={(e) => setPrio("car", { capKwh: +e.target.value })} /><span style={S.u}>kWh/d</span>
                </div>
                <div style={S.prioRow}>
                  <span style={S.prioNum}>4</span>
                  <label style={S.prioCheck}><input type="checkbox" checked={fve.prio.battery.on} onChange={(e) => setPrio("battery", { on: e.target.checked })} /><span>Domácí baterie</span></label>
                  <input style={S.prioCap} type="number" value={fve.prio.battery.capKwh} onChange={(e) => setPrio("battery", { capKwh: +e.target.value })} /><span style={S.u}>kWh/d</span>
                </div>
                <p style={S.note}>Bojler je nejlevnější „baterie". U tarifu D57d (20 h NT) má domácí baterie slabší návratnost — zváž až po bojleru a autě.</p>
              </Card>
            </div>

            <div>
              <Card title="Rozdělení výroby" hint={season === "summer" ? "léto" : "zima"}>
                {fveCalc ? (
                  <div style={S.fveResult}>
                    <div style={S.fveResRow}><span>Výroba celkem</span><strong>{fmt(fveCalc.prodMonth, 0)} kWh/měs</strong></div>
                    {fveCalc.steps.map((st) => (
                      <div key={st.key} style={{ ...S.fveResRow, ...(st.key === "grid" ? { color: "#9aa090" } : {}) }}>
                        <span>{st.label}</span>
                        <strong>{fmt(st.kwh, 0)} kWh · {fmtKc(st.save)}</strong>
                      </div>
                    ))}
                    <div style={{ ...S.fveResRow, ...S.fveResAccent }}><span>Úspora celkem</span><strong>{fmtKc(fveCalc.saveMonth)}/měs</strong></div>
                    <div style={S.fveResRow}><span>Nový účet / měs</span><strong>{fmtKc(fveCalc.newBill)}</strong></div>
                    <div style={S.fveResRow}><span>Hrubá návratnost</span><strong>~{fmt(fveCalc.payback, 1)} let</strong></div>
                  </div>
                ) : <p style={S.empty}>Zapni FVE a vyplň parametry.</p>}
              </Card>

              <Card title="Tip na úsporu" hint="Přesun spotřeby pod panely">
                <p style={S.note}>
                  FVE vyrábí přes den, špička v poledne. Největší úsporu má přesun velkých
                  spotřebičů (bojler, myčka, pračka, čerpadla, dobíjení auta, akumulace DEVI)
                  do polední výrobní špičky. U tebe je výhoda strmá střecha (44°) — lépe
                  vyrábí na jaře/podzim a v zimě, kdy je slunce níž a ty topíš.
                </p>
              </Card>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function Stat({ label, value, accent, big }) {
  return (<div style={{ ...S.stat, ...(big ? S.statBig : {}) }}><div style={{ ...S.statLabel, ...(big ? { color: "#b9c4ad" } : {}) }}>{label}</div><div style={{ ...S.statValue, ...(accent ? { color: "#2d6a4f" } : {}), ...(big ? { color: "#fff" } : {}) }}>{value}</div></div>);
}
function Card({ title, hint, children }) {
  return (<div style={S.card}><div style={S.cardHead}><h2 style={S.cardTitle}>{title}</h2>{hint && <span style={S.cardHint}>{hint}</span>}</div>{children}</div>);
}
function Field({ label, children }) {
  return (<div style={S.field}><span style={S.fieldLabel}>{label}</span><div style={S.fieldIn}>{children}</div></div>);
}
function RateRow({ label, unit, value, onChange }) {
  return (<div style={S.rateRow}><span style={S.rateLabel}>{label}</span><div style={S.rateInputWrap}><input style={S.rateInput} type="number" step="0.001" value={value} onChange={(e) => onChange(+e.target.value)} /><span style={S.rateUnit}>{unit}</span></div></div>);
}

const ACCENT = "#2d6a4f", INK = "#1c2419", PAPER = "#f6f4ec", GOLD = "#e8b04b";
const S = {
  page: { minHeight: "100vh", background: PAPER, color: INK, fontFamily: "'DM Sans', system-ui, sans-serif", padding: "32px 20px 60px", maxWidth: 1100, margin: "0 auto" },
  header: { marginBottom: 20 },
  headTop: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, flexWrap: "wrap" },
  eyebrow: { textTransform: "uppercase", letterSpacing: "0.18em", fontSize: 11, fontWeight: 700, color: ACCENT, marginBottom: 8 },
  h1: { fontFamily: "'Fraunces', Georgia, serif", fontSize: 44, fontWeight: 600, margin: 0, lineHeight: 1.05, letterSpacing: "-0.02em" },
  seasonTog: { display: "flex", background: "#fff", border: "1px solid #e3e0d4", borderRadius: 10, padding: 3, gap: 3 },
  seasonBtn: { border: "none", background: "none", padding: "8px 16px", borderRadius: 8, fontSize: 14, fontWeight: 600, color: "#8a9080", cursor: "pointer", fontFamily: "inherit" },
  seasonOn: { background: ACCENT, color: "#fff" },
  summaryRow: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 14, marginBottom: 20 },
  stat: { background: "#fff", border: "1px solid #e3e0d4", borderRadius: 14, padding: "16px 18px" },
  statBig: { background: INK, border: "1px solid " + INK },
  statLabel: { fontSize: 12, color: "#8a9080", fontWeight: 600, marginBottom: 6 },
  statValue: { fontFamily: "'Fraunces', serif", fontSize: 26, fontWeight: 600 },
  tabs: { display: "flex", gap: 8, marginBottom: 22, flexWrap: "wrap" },
  tab: { display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 2, background: "#fff", border: "1px solid #e3e0d4", borderRadius: 11, padding: "10px 14px", cursor: "pointer", fontFamily: "inherit", minWidth: 110 },
  tabOn: { background: ACCENT, borderColor: ACCENT },
  tabName: { fontSize: 13.5, fontWeight: 700, color: INK },
  tabVal: { fontSize: 11.5, color: "#8a9080", fontWeight: 600 },
  grid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 18, alignItems: "start" },
  card: { background: "#fff", border: "1px solid #e3e0d4", borderRadius: 16, padding: 20, marginBottom: 18 },
  cardHead: { marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10, flexWrap: "wrap" },
  cardTitle: { fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, margin: 0 },
  cardHint: { fontSize: 12.5, color: ACCENT, fontWeight: 600 },
  collapseHead: { width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, background: "none", border: "none", padding: 0, cursor: "pointer", fontFamily: "inherit", flexWrap: "wrap" },
  collapseTitle: { display: "flex", alignItems: "center", gap: 9, fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, color: INK },
  chevron: { fontSize: 22, color: ACCENT, transition: "transform .18s", lineHeight: 1, display: "inline-block" },
  collapseCount: { fontSize: 12, fontWeight: 700, color: "#9aa090", background: "#f0ede1", borderRadius: 10, padding: "1px 8px", fontFamily: "'DM Sans', sans-serif" },
  devList: { display: "flex", flexDirection: "column", gap: 10 },
  devCard: { border: "1px solid #eee9da", borderRadius: 12, padding: "12px 12px 10px", background: "#fbfaf5", transition: "opacity .2s" },
  devTop: { display: "grid", gridTemplateColumns: "1fr auto auto auto auto", alignItems: "center", gap: 6, marginBottom: 8 },
  devName: { border: "none", background: "none", fontSize: 14.5, fontWeight: 600, fontFamily: "inherit", color: INK, padding: 0, minWidth: 0 },
  placeTag: { fontSize: 9.5, fontWeight: 700, color: "#6a7060", background: "#eef0e9", padding: "2px 6px", borderRadius: 5, textTransform: "uppercase", whiteSpace: "nowrap" },
  measuredTag: { fontSize: 9.5, fontWeight: 700, color: "#fff", background: "#2d6a4f", padding: "2px 6px", borderRadius: 5, textTransform: "uppercase", whiteSpace: "nowrap" },
  sheetBox: { background: "#fbfaf5", border: "1px solid #eee9da", borderRadius: 10, padding: 12, marginBottom: 12 },
  sheetLabel: { fontSize: 11, fontWeight: 700, color: "#6a7060", display: "block", marginBottom: 4, marginTop: 8 },
  sheetInput: { width: "100%", border: "1px solid #ddd9cb", borderRadius: 7, padding: "7px 9px", fontSize: 12.5, fontFamily: "inherit", background: "#fff", boxSizing: "border-box" },
  sheetActions: { display: "flex", alignItems: "center", gap: 10, marginTop: 12, flexWrap: "wrap" },
  syncBtn: { background: ACCENT, color: "#fff", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  syncStatus: { fontSize: 12, color: "#5a6450", flex: 1 },
  orDivider: { textAlign: "center", fontSize: 11.5, color: "#9aa090", margin: "4px 0 10px" },
  seasonTag: { fontSize: 9.5, fontWeight: 700, color: GOLD, background: "#fdf6e9", padding: "2px 6px", borderRadius: 5, textTransform: "uppercase" },
  devCost: { fontSize: 13, fontWeight: 700, color: ACCENT, fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" },
  devCostUnit: { fontSize: 10, fontWeight: 600, color: "#9aa090", marginLeft: 2 },
  devFields: { display: "flex", flexWrap: "wrap", gap: 10 },
  devUnit: { fontSize: 11.5, color: "#8a9080", marginTop: 9, paddingTop: 8, borderTop: "1px dashed #e8e3d3", fontVariantNumeric: "tabular-nums" },
  field: { display: "flex", flexDirection: "column", gap: 3 },
  fieldLabel: { fontSize: 10.5, color: "#9aa090", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" },
  fieldIn: { display: "flex", alignItems: "center", gap: 3 },
  inS: { border: "1px solid #ddd9cb", borderRadius: 7, padding: "5px 6px", fontSize: 13, fontFamily: "inherit", textAlign: "right", background: "#fff", width: 52 },
  u: { fontSize: 11, color: "#9aa090" },
  empty: { fontSize: 13.5, color: "#9aa090", lineHeight: 1.5 },
  addRow: { display: "flex", gap: 8, marginBottom: 14 },
  addBtn: { background: "none", border: "1px dashed #c3c8b8", borderRadius: 8, padding: "8px 12px", color: ACCENT, fontWeight: 600, fontSize: 13, cursor: "pointer", flex: 1, fontFamily: "inherit" },
  x: { background: "none", border: "none", color: "#c0563a", fontSize: 19, cursor: "pointer", lineHeight: 1, padding: 0 },
  drop: { border: "1.5px dashed #c3c8b8", borderRadius: 12, padding: "22px 16px", textAlign: "center", cursor: "pointer", color: "#6a7060", fontSize: 14, display: "flex", flexDirection: "column", alignItems: "center", gap: 6, background: "#fbfaf5" },
  dropPlus: { fontSize: 24, color: ACCENT, fontWeight: 300 },
  errBox: { fontSize: 13, color: "#c0563a", marginTop: 10, background: "#fbeee9", padding: "8px 10px", borderRadius: 8 },
  fileRow: { display: "grid", gridTemplateColumns: "1fr auto auto", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: "1px solid #eee9da", fontSize: 14 },
  fileVal: { fontWeight: 600, fontVariantNumeric: "tabular-nums" },
  perDay: { fontSize: 12, color: "#8a9080", marginTop: 10, textAlign: "center" },
  ntRow: { display: "flex", alignItems: "center", gap: 7, marginBottom: 8 },
  ntDot: { width: 8, height: 8, borderRadius: 4, background: ACCENT, flexShrink: 0 },
  inTime: { border: "1px solid #ddd9cb", borderRadius: 8, padding: "7px 8px", fontSize: 13.5, fontFamily: "inherit", background: "#fbfaf5", width: 60, textAlign: "right" },
  dash: { color: "#9aa090" },
  dayStrip: { display: "grid", gridTemplateColumns: "repeat(24, 1fr)", gap: 2, marginTop: 16, height: 26, borderRadius: 6, overflow: "hidden" },
  dayCell: { width: "100%", height: "100%" },
  legend: { display: "flex", gap: 16, marginTop: 10, fontSize: 12, color: "#6a7060" },
  dayTypeTog: { display: "flex", gap: 4, marginBottom: 14, background: "#f0ede1", borderRadius: 9, padding: 3, width: "fit-content" },
  dayTypeBtn: { border: "none", background: "none", padding: "6px 14px", borderRadius: 7, fontSize: 13, fontWeight: 600, color: "#8a9080", cursor: "pointer", fontFamily: "inherit" },
  dayTypeOn: { background: ACCENT, color: "#fff" },
  tabFve: { background: "#fdf9ef", border: "1px solid #ecdcb8" },
  fveIntro: { background: "#fff", border: "1px solid #e3e0d4", borderRadius: 16, padding: "32px 28px", textAlign: "center", maxWidth: 560, margin: "0 auto" },
  fveIntroTitle: { fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 600, margin: "0 0 12px" },
  fveIntroText: { fontSize: 14.5, color: "#5a6450", lineHeight: 1.6, marginBottom: 20 },
  fveBigToggle: { display: "inline-flex", alignItems: "center", gap: 10, fontSize: 15, fontWeight: 600, cursor: "pointer", background: "#fbfaf5", border: "1px solid #e3e0d4", borderRadius: 10, padding: "12px 20px" },
  legendItem: { display: "flex", alignItems: "center", gap: 5 },
  legendDot: { display: "inline-block", width: 11, height: 11, borderRadius: 3 },
  rateRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 },
  rateLabel: { fontSize: 13.5 },
  rateInputWrap: { display: "flex", alignItems: "center", gap: 8 },
  rateInput: { border: "1px solid #ddd9cb", borderRadius: 8, padding: "7px 9px", fontSize: 13.5, width: 92, textAlign: "right", fontFamily: "inherit", background: "#fbfaf5" },
  rateUnit: { fontSize: 12, color: "#9aa090", width: 52 },
  footer: { marginTop: 30, fontSize: 12.5, color: "#9aa090", lineHeight: 1.6, borderTop: "1px solid #e3e0d4", paddingTop: 16 },
  fveSaveBar: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, background: "linear-gradient(90deg, #fdf6e9, #eef4ec)", border: "1px solid #e3e0d4", borderRadius: 12, padding: "12px 18px", marginBottom: 22, flexWrap: "wrap" },
  fveSaveText: { fontSize: 14.5, color: INK },
  fveSaveDetail: { fontSize: 12.5, color: "#8a9080" },
  fveToggle: { display: "flex", alignItems: "center", gap: 9, fontSize: 14, fontWeight: 600, cursor: "pointer" },
  fveResult: { marginTop: 14, padding: "12px 14px", background: "#fbfaf5", borderRadius: 10, border: "1px solid #eee9da" },
  fveResRow: { display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0", color: "#5a6450" },
  fveResAccent: { borderTop: "1px solid #e8e3d3", borderBottom: "1px solid #e8e3d3", margin: "4px 0", padding: "8px 0", color: ACCENT, fontSize: 14 },
  prioHead: { fontSize: 13, fontWeight: 700, color: INK, marginTop: 16, marginBottom: 2 },
  prioNote: { fontSize: 11.5, color: "#9aa090", marginBottom: 10, lineHeight: 1.4 },
  prioRow: { display: "flex", alignItems: "center", gap: 8, marginBottom: 7 },
  prioNum: { width: 20, height: 20, borderRadius: 10, background: ACCENT, color: "#fff", fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  prioLabel: { fontSize: 13, flex: 1 },
  prioFixed: { fontSize: 11, color: "#9aa090", fontStyle: "italic" },
  prioCheck: { display: "flex", alignItems: "center", gap: 6, fontSize: 13, flex: 1, cursor: "pointer" },
  prioCap: { border: "1px solid #ddd9cb", borderRadius: 7, padding: "5px 6px", fontSize: 13, fontFamily: "inherit", textAlign: "right", background: "#fff", width: 50 },
};
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=DM+Sans:wght@400;500;600;700&display=swap');
* { box-sizing: border-box; } body { margin: 0; }
/* iOS: nevynucuj systémovou modrou na tlačítkách (inline barvy mají přednost) */
button { -webkit-tap-highlight-color: transparent; -webkit-appearance: none; appearance: none; }
.legend i { display:inline-block; width:11px; height:11px; border-radius:3px; margin-right:5px; vertical-align:middle; }
input:focus, button:focus-visible { outline: 2px solid ${ACCENT}; outline-offset: 1px; }
@media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
`;
