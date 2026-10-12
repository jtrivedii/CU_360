import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  Search, Shield, AlertTriangle, TrendingUp, TrendingDown, Building2, Users,
  ChevronRight, FileText, Sparkles, X, ChevronDown, Printer, Info, Landmark,
  MapPin, ArrowUpRight, ArrowDownRight, Minus, Loader2, Check, GitCompare,
  Sun, Moon, Plus, Home, MessageCircle, Send, Bot, Database, Upload, RefreshCw,
  Settings, KeyRound, Link2, Clock3, CheckCircle2, AlertCircle, Download,
  Workflow, Server, LockKeyhole, SlidersHorizontal, Target, Layers3, LogIn, LogOut, UserCircle
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, Cell, LineChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";

/* ============================== DATA ============================== */
const DATASET = {"fields":["name","city","state","region","bucket","assetsM","members","branches","employees","loansM","depositsM","niM","equityM","roa","roe","ltd","lta","dr","nco","py","capRatio","niiM","ceo","website","grossIncomeM","operatingExpenseM","delinquentAmountM","delinquentCount","chargeOffM","recoveryM","efficiencyRatio","netMargin","recoveryRate","charter","quarter"],"bucketLabels":["<$50M","$50-100M","$100-250M","$250-500M","$500M-1B","$1B-5B","$5B-10B","$10B+"],"rows":[]};
const FIELDS = DATASET.fields;
const BUCKET_LABELS = DATASET.bucketLabels;
const FI = Object.fromEntries(FIELDS.map((f, i) => [f, i]));

function rowToCU(row, idx) {
  return {
    idx,
    name: row[FI.name],
    city: row[FI.city],
    state: row[FI.state],
    bucket: row[FI.bucket],
    assetsM: row[FI.assetsM],
    members: row[FI.members],
    branches: row[FI.branches],
    employees: row[FI.employees],
    loansM: row[FI.loansM],
    depositsM: row[FI.depositsM],
    niM: row[FI.niM],
    equityM: row[FI.equityM],
    roa: row[FI.roa],
    roe: row[FI.roe],
    ltd: row[FI.ltd],
    lta: row[FI.lta],
    dr: row[FI.dr],
    nco: row[FI.nco],
    py: row[FI.py],
    capRatio: row[FI.capRatio],
    niiM: row[FI.niiM],
    ceo: row[FI.ceo] || "",
    website: row[FI.website] || "",
    grossIncomeM: row[FI.grossIncomeM] || 0,
    operatingExpenseM: row[FI.operatingExpenseM] || 0,
    delinquentAmountM: row[FI.delinquentAmountM] || 0,
    delinquentCount: row[FI.delinquentCount] || 0,
    chargeOffM: row[FI.chargeOffM] || 0,
    recoveryM: row[FI.recoveryM] || 0,
    efficiencyRatio: row[FI.efficiencyRatio] || 0,
    netMargin: row[FI.netMargin] || 0,
    recoveryRate: row[FI.recoveryRate] || 0,
    charter: row[FI.charter] || "",
    quarter: row[FI.quarter] || "",
  };
}

let ALL_CUS = [];

/* ============================== MULTI-QUARTER DATA LAYER ============================== */
// The portal is now backed by the supplied 2026 Q1 ACUI dataset. Additional quarters
// can still be registered at runtime without changing the existing analytics components.
let ACTIVE_CUS = ALL_CUS;
const DEFAULT_QUARTER = "2026 Q1";
let ACTIVE_QUARTER = DEFAULT_QUARTER;
let DATA_VERSION = 0; // bumped whenever the active dataset changes, so memoised views recompute
function setActiveData(rows, quarter) { ACTIVE_CUS = rows; ACTIVE_QUARTER = quarter; DATA_VERSION += 1; }
function findQuarterCU(rows, target) {
  if (!target) return null;
  const charter = String(target.charter || "").trim();
  if (charter) {
    const match = rows.find((r) => String(r.charter || "").trim() === charter);
    if (match) return match;
  }
  return rows.find((r) => r.name?.trim().toLowerCase() === target.name?.trim().toLowerCase() && r.state === target.state) || null;
}
function institutionIdentity(cu) {
  return cu?.charter ? `charter:${String(cu.charter).trim()}` : `${cu?.name?.trim().toLowerCase() || ""}|${cu?.state || ""}`;
}

/* ---------- Import helpers: state names, column aliases, number cleaning ---------- */
const STATE_NAMES = { AL:"Alabama", AK:"Alaska", AZ:"Arizona", AR:"Arkansas", CA:"California", CO:"Colorado", CT:"Connecticut", DE:"Delaware", DC:"District of Columbia", FL:"Florida", GA:"Georgia", HI:"Hawaii", ID:"Idaho", IL:"Illinois", IN:"Indiana", IA:"Iowa", KS:"Kansas", KY:"Kentucky", LA:"Louisiana", ME:"Maine", MD:"Maryland", MA:"Massachusetts", MI:"Michigan", MN:"Minnesota", MS:"Mississippi", MO:"Missouri", MT:"Montana", NE:"Nebraska", NV:"Nevada", NH:"New Hampshire", NJ:"New Jersey", NM:"New Mexico", NY:"New York", NC:"North Carolina", ND:"North Dakota", OH:"Ohio", OK:"Oklahoma", OR:"Oregon", PA:"Pennsylvania", RI:"Rhode Island", SC:"South Carolina", SD:"South Dakota", TN:"Tennessee", TX:"Texas", UT:"Utah", VT:"Vermont", VA:"Virginia", WA:"Washington", WV:"West Virginia", WI:"Wisconsin", WY:"Wyoming", PR:"Puerto Rico", VI:"U.S. Virgin Islands", GU:"Guam" };
const STATE_NAME_TO_CODE = Object.fromEntries(Object.entries(STATE_NAMES).map(([code, n]) => [n.toLowerCase(), code]));
function toStateCode(v) {
  const s = String(v ?? "").trim();
  if (!s) return "";
  if (s.length === 2) return s.toUpperCase();
  return STATE_NAME_TO_CODE[s.toLowerCase()] || s.toUpperCase().slice(0, 2);
}

// [key, label, required, aliases]. Aliases are compared after lower-casing and stripping
// everything except letters and digits, so "Total Assets", "TOTAL_ASSETS" and "total-assets" all match.
const IMPORT_FIELD_DEFS = [
  ["name", "Credit union name", true, ["name", "creditunion", "creditunionname", "cuname", "institution", "institutionname"]],
  ["city", "City", false, ["city", "cucity", "headquarterscity", "hqcity"]],
  ["state", "State", true, ["state", "st", "statecode", "stateabbr", "stateabbreviation", "custate"]],
  ["ceo", "CEO", false, ["ceo", "ceoname", "chiefexecutiveofficer", "chiefexecutivename"]],
  ["website", "Website", false, ["website", "url", "web", "homepage"]],
  ["charter", "Charter / CU number", false, ["charter", "charternumber", "charterno", "cunumber", "cuid"]],
  ["region", "Region", false, ["region", "regioncode"]],
  ["bucket", "Asset bucket (0-7)", false, ["bucket", "assetbucket", "assettier", "assetgroup"]],
  ["assetsM", "Total assets", true, ["assetsm", "assets", "totalassets", "totalassetsm", "assetsmillions", "totalassetsmillions", "acct010"]],
  ["members", "Members", false, ["members", "totalmembers", "membership", "nummembers", "numberofmembers", "membercount", "acct083"]],
  ["branches", "Branches", false, ["branches", "numbranches", "numberofbranches", "branchcount", "totalbranches"]],
  ["employees", "Employees (FTE)", false, ["employees", "fte", "totalemployees", "numemployees", "numberofemployees", "fulltimeequivalent", "fulltimeequivalents"]],
  ["loansM", "Total loans", false, ["loansm", "loans", "amountoftotalloansleases", "totalloans", "totalloansandleases", "netloans", "loansandleases", "acct025b"]],
  ["depositsM", "Total deposits / shares", false, ["depositsm", "deposits", "totaldeposits", "totalsharesamount", "sharesamount", "totalshares", "totalsharesanddeposits", "sharesanddeposits", "shares", "acct018"]],
  ["niM", "Net income", false, ["nim", "netincome", "netincomem", "netincomeloss", "ni", "acct661a"]],
  ["equityM", "Equity / net worth", false, ["equitym", "equity", "networthamount", "networth", "totalnetworth", "totalequity", "acct998"]],
  ["roa", "Return on assets", false, ["roa", "returnonassets", "returnonaverageassets", "returnonavgassets"]],
  ["roe", "Return on equity", false, ["roe", "returnonequity", "returnonaverageequity"]],
  ["ltd", "Loan-to-deposit", false, ["ltd", "loantodeposit", "loantodeposits", "loantodepositratio", "loantoshare", "loantoshares", "loantoshareratio", "loansharesratio"]],
  ["lta", "Loan-to-assets", false, ["lta", "loantoasset", "loantoassets", "loantoassetsratio", "loanstoassets"]],
  ["dr", "Delinquency rate", false, ["dr", "delinquency", "delinquencyrate", "delinquencyratio", "delinquentloanratio", "delinquentloansratio", "delinquentloanrate"]],
  ["nco", "Net charge-off rate", false, ["nco", "netchargeoff", "netchargeoffs", "netchargeoffrate", "netchargeoffratio", "chargeoffratio"]],
  ["py", "Portfolio yield", false, ["py", "portfolioyield", "loanyield", "yieldonloans", "yield"]],
  ["capRatio", "Net worth / capital ratio", false, ["capratio", "networthratio", "capitalratio", "networthtoassets", "capitaltoassets"]],
  ["niiM", "Net interest income", false, ["niim", "netinterestincome", "nii"]],
  ["operatingExpenseM", "Operating expense", false, ["operatingexpensem", "operatingexpense", "operatingexpenses", "totaloperatingexpense", "totaloperatingexpenses", "noninterestexpense", "noninterestexpensem"]],
  ["personnelExpenseM", "Personnel expense", false, ["personnelexpensem", "personnelexpense", "personnelexpenses", "personnelm", "staffexpensem", "compensationandbenefits"]],
  ["marketingExpenseM", "Marketing expense", false, ["marketingexpensem", "marketingexpense", "marketingm", "promotionalexpense"]],
  ["technologyExpenseM", "Technology expense", false, ["technologyexpensem", "technologyexpense", "technologym", "itexpensem"]],
  ["occupancyExpenseM", "Occupancy expense", false, ["occupancyexpensem", "occupancyexpense", "occupancym", "officeoccupancyexpense"]],
  ["otherOperatingExpenseM", "Other operating expense", false, ["otheroperatingexpensem", "otheroperatingexpense", "otheroperatingexpenses", "otherexpensem", "otherexpense"]],
  ["grossIncomeM", "Gross income", false, ["grossincomem", "grossincome", "grossincomeamount"]],
  ["delinquentAmountM", "Delinquent loans amount", false, ["delinquentamountm", "delinquentamount", "totalamountofdelinquentloansleases", "delinquentloansamount"]],
  ["delinquentCount", "Delinquent loans count", false, ["delinquentcount", "numberofdelinquentloansleases", "numberofdelinquentloans"]],
  ["chargeOffM", "Charge-offs", false, ["chargeoffm", "chargeoffs", "totalamountofallloanschargedoffytd"]],
  ["recoveryM", "Recoveries", false, ["recoverym", "recoveries", "totalamountofallloansrecoveredytd"]],
];
const IMPORT_FIELDS = IMPORT_FIELD_DEFS.map(([key, label, required, aliases]) => ({ key, label, required, aliases }));
const NUMERIC_KEYS = IMPORT_FIELDS.map((f) => f.key).filter((k) => !["name", "city", "state", "charter", "region", "bucket", "ceo", "website"].includes(k));
const MONEY_KEYS = ["assetsM", "loansM", "depositsM", "niM", "equityM", "niiM", "operatingExpenseM", "personnelExpenseM", "marketingExpenseM", "technologyExpenseM", "occupancyExpenseM", "otherOperatingExpenseM", "grossIncomeM", "delinquentAmountM", "chargeOffM", "recoveryM"];
const RATIO_KEYS = ["roa", "roe", "ltd", "lta", "dr", "nco", "py", "capRatio"];

const normKey = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

// Excel cells arrive as numbers, but CSV / pasted data often has "1,234", "$5.2", "3.1%" or "(45)".
function toNum(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  if (v == null || v instanceof Date) return 0;
  let s = String(v).trim();
  if (!s || /^(n\/?a|null|none|-|—)$/i.test(s)) return 0;
  const negative = /^\(.*\)$/.test(s) || /^-/.test(s);
  s = s.replace(/[^0-9.eE]/g, "");
  const n = Number(s);
  if (!Number.isFinite(n)) return 0;
  return negative ? -Math.abs(n) : n;
}

function autoMapColumns(headers) {
  const normHeaders = headers.map(normKey);
  const mapping = {};
  const used = new Set();
  const aliasesOf = (f) => [normKey(f.key), ...f.aliases.map(normKey)];
  // Pass 1: exact (normalised) matches, in alias priority order.
  IMPORT_FIELDS.forEach((f) => {
    for (const alias of aliasesOf(f)) {
      const i = normHeaders.findIndex((h, idx) => !used.has(idx) && h === alias);
      if (i >= 0) { mapping[f.key] = headers[i]; used.add(i); return; }
    }
  });
  // Pass 2: long aliases contained in a header (e.g. "TotalAssetsUSD"), skipping prior-year / change columns.
  const noise = /(prior|prev|pct|percent|change|growth|rank|avg|average|yoy|qoq)/;
  IMPORT_FIELDS.forEach((f) => {
    if (mapping[f.key]) return;
    for (const alias of aliasesOf(f).filter((a) => a.length >= 9)) {
      const i = normHeaders.findIndex((h, idx) => !used.has(idx) && h.includes(alias) && !noise.test(h));
      if (i >= 0) { mapping[f.key] = headers[i]; used.add(i); return; }
    }
  });
  return mapping;
}

/* ---------- Quarter label detection ("2023-Q4.xlsx", cycle-date column, sheet name) ---------- */
function quarterFromDate(d) {
  const dt = new Date(d.getTime() + 12 * 3600e3);
  let y = dt.getUTCFullYear(); let m = dt.getUTCMonth();
  // A quarter-end date (e.g. Dec 31) can drift just past midnight when Excel converts it.
  if ([0, 3, 6, 9].includes(m) && dt.getUTCDate() <= 2) { m -= 1; if (m < 0) { m = 11; y -= 1; } }
  return `${y} Q${Math.floor(m / 3) + 1}`;
}
function labelFromAny(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return quarterFromDate(v);
  if (typeof v === "number") {
    if (v > 20000 && v < 80000) return quarterFromDate(new Date(Math.round((v - 25569) * 864e5)));
    v = String(v);
  }
  const s = String(v);
  let m = s.match(/(20\d{2})\s*[-_ ]?\s*Q\s*([1-4])/i); if (m) return `${m[1]} Q${m[2]}`;
  m = s.match(/Q\s*([1-4])\s*[-_ ]?\s*(20\d{2})/i); if (m) return `${m[2]} Q${m[1]}`;
  m = s.match(/(20\d{2})[-\/.]?(0[1-9]|1[0-2])[-\/.]?([0-3]\d)/); if (m) return `${m[1]} Q${Math.ceil(Number(m[2]) / 3)}`;
  m = s.match(/(\d{1,2})[\/-](\d{1,2})[\/-](20\d{2})/); if (m) return `${m[3]} Q${Math.ceil(Number(m[1]) / 3)}`;
  return "";
}
function detectQuarterLabel(sheet) {
  const fromName = labelFromAny(String(sheet.fileName || "").replace(/\.[^.]+$/, ""));
  if (fromName) return fromName;
  const fromSheet = labelFromAny(sheet.sheetName);
  if (fromSheet) return fromSheet;
  const col = (sheet.headers || []).find((h) => ["cycledate", "quarter", "period", "reportdate", "asofdate", "calldate", "reportingperiod", "datadate"].includes(normKey(h)));
  if (col) {
    const hit = (sheet.rows || []).find((r) => r[col] !== "" && r[col] != null);
    const l = hit ? labelFromAny(hit[col]) : "";
    if (l) return l;
  }
  return "";
}

/* ---------- Read .xlsx / .xls / .csv / .json into { headers, rows } ---------- */
async function readSpreadsheetFile(file) {
  const lower = file.name.toLowerCase();
  const fromObjects = (rows, sheetName = null) => ({ headers: rows.length ? Object.keys(rows[0]) : [], rows, sheetName });

  if (lower.endsWith(".json")) {
    let data;
    try { data = JSON.parse(await file.text()); } catch { throw new Error("That JSON file could not be parsed."); }
    // Supports this app's own { fields: [...], rows: [[...]] } layout as well as an array of objects.
    if (data && Array.isArray(data.fields) && Array.isArray(data.rows) && Array.isArray(data.rows[0])) {
      return fromObjects(data.rows.map((r) => Object.fromEntries(data.fields.map((f, i) => [f, r[i]]))));
    }
    const rows = Array.isArray(data) ? data : (data && Array.isArray(data.rows) ? data.rows : []);
    return fromObjects(rows);
  }

  if (/\.(csv|tsv|txt)$/.test(lower)) return fromObjects(parseDelimitedText(await file.text()));

  let wb;
  try { wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true }); }
  catch { throw new Error("That file could not be opened as a spreadsheet. Save it as .xlsx or .csv and try again."); }

  // Use the sheet with the most rows (call-report workbooks sometimes include a README / notes tab).
  let best = null;
  wb.SheetNames.forEach((sn) => {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: true, blankrows: false });
    if (!best || aoa.length > best.aoa.length) best = { sn, aoa };
  });
  if (!best || best.aoa.length < 2) throw new Error("No data rows were found in that workbook.");

  // Header row = first row (within the first 25) that is mostly text. Handles title rows above the table.
  const isText = (c) => typeof c === "string" && c.trim() !== "" && Number.isNaN(Number(c.replace(/[,$%]/g, "")));
  const scores = best.aoa.slice(0, 25).map((r) => r.filter(isText).length);
  const top = Math.max(...scores);
  let headerIdx = scores.findIndex((s) => s >= Math.max(3, Math.floor(top * 0.8)));
  if (headerIdx < 0) headerIdx = 0;

  const seen = {};
  const headers = best.aoa[headerIdx].map((h, i) => {
    const base = String(h ?? "").trim() || `Column ${i + 1}`;
    seen[base] = (seen[base] || 0) + 1;
    return seen[base] > 1 ? `${base} (${seen[base]})` : base;
  });
  const rows = best.aoa.slice(headerIdx + 1)
    .filter((r) => r.some((c) => c !== "" && c != null))
    .map((r) => { const o = {}; headers.forEach((h, i) => { o[h] = r[i] ?? ""; }); return o; });
  return { headers, rows, sheetName: best.sn, sheetCount: wb.SheetNames.length };
}

/* ---------- Normalise uploaded rows into CU records ---------- */
function normalizeQuarterRows(rows, quarter, mappingIn) {
  const list = Array.isArray(rows) ? rows : [];
  const mapping = mappingIn || autoMapColumns(list.length ? Object.keys(list[0]) : []);
  const has = (k) => !!mapping[k];
  const get = (r, key) => (mapping[key] ? r[mapping[key]] : undefined);
  const text = (v) => (v == null || v instanceof Date ? "" : String(v).trim());
  const median = (arr) => { const a = arr.filter(Number.isFinite).sort((x, y) => x - y); return a.length ? a[Math.floor(a.length / 2)] : 0; };

  const parsed = [];
  let skipped = 0;
  list.forEach((row) => {
    const r = row || {};
    const name = text(get(r, "name"));
    if (!name) { skipped += 1; return; }
    const rec = { name, city: text(get(r, "city")), state: toStateCode(get(r, "state")), charter: text(get(r, "charter")), region: toNum(get(r, "region")), ceo: text(get(r, "ceo")), website: text(get(r, "website")) };
    NUMERIC_KEYS.forEach((k) => { rec[k] = toNum(get(r, k)); });
    // The new ACUI extract carries net-worth components separately. Prefer their sum over
    // NET_WORTH_AMOUNT when that field is duplicated / non-ratio-like (as in the supplied file).
    const equityComponentKeys = Object.keys(r).filter((k) => /^equity/i.test(normKey(k)) && /(amount|earnings|reserve|interest|gain|loss|merger)/i.test(k));
    const equityComponents = equityComponentKeys.reduce((sum, k) => sum + toNum(r[k]), 0);
    if (equityComponents && (!has("equityM") || rec.equityM >= Math.max(rec.assetsM * 0.5, 1))) rec.equityM = equityComponents;
    rec._bucket = get(r, "bucket");
    parsed.push(rec);
  });

  // Unit detection: call-report extracts are often in dollars or thousands; this app works in $ millions.
  const notes = [];
  const medAssets = median(parsed.map((r) => r.assetsM).filter((v) => v > 0));
  let divisor = 1;
  if (medAssets > 1e7) { divisor = 1e6; notes.push("Amounts looked like whole dollars, so they were converted to $ millions."); }
  else if (medAssets > 5e3) { divisor = 1e3; notes.push("Amounts looked like $ thousands, so they were converted to $ millions."); }
  if (divisor !== 1) parsed.forEach((rec) => MONEY_KEYS.forEach((k) => { rec[k] = rec[k] / divisor; }));

  // Ratios exported from Excel as fractions (0.85) rather than percents (85).
  const medLtd = has("ltd") ? median(parsed.map((r) => r.ltd).filter((v) => v > 0)) : null;
  const medCap = has("capRatio") ? median(parsed.map((r) => r.capRatio).filter((v) => v > 0)) : null;
  const tests = [];
  if (medLtd != null) tests.push(medLtd < 3);
  if (medCap != null) tests.push(medCap < 0.5);
  if (tests.length && tests.every(Boolean)) {
    parsed.forEach((rec) => RATIO_KEYS.forEach((k) => { rec[k] = rec[k] * 100; }));
    notes.push("Ratios looked like fractions (0.85), so they were converted to percents (85).");
  }

  // Calculate simple ratios when the file has the inputs but not the ratio itself.
  const derivedKeys = [];
  if (!has("ltd") && has("loansM") && has("depositsM")) { derivedKeys.push("ltd"); parsed.forEach((r) => { r.ltd = r.depositsM > 0 ? (r.loansM / r.depositsM) * 100 : 0; }); }
  if (!has("lta") && has("loansM") && has("assetsM")) { derivedKeys.push("lta"); parsed.forEach((r) => { r.lta = r.assetsM > 0 ? (r.loansM / r.assetsM) * 100 : 0; }); }
  const hasEquityValues = parsed.some((r) => r.equityM > 0);
  if (!has("capRatio") && (has("equityM") || hasEquityValues) && has("assetsM")) { derivedKeys.push("capRatio"); parsed.forEach((r) => { r.capRatio = r.assetsM > 0 ? (r.equityM / r.assetsM) * 100 : 0; }); }
  if (!has("roa") && has("niM") && has("assetsM")) { derivedKeys.push("roa"); parsed.forEach((r) => { r.roa = r.assetsM > 0 ? (r.niM / r.assetsM) * 100 : 0; }); }
  if (!has("roe") && has("niM") && has("equityM")) { derivedKeys.push("roe"); parsed.forEach((r) => { r.roe = r.equityM > 0 ? (r.niM / r.equityM) * 100 : 0; }); }
  if (!has("dr") && has("delinquentAmountM") && has("loansM")) { derivedKeys.push("dr"); parsed.forEach((r) => { r.dr = r.loansM > 0 ? (r.delinquentAmountM / r.loansM) * 100 : 0; }); }
  if (!has("nco") && has("chargeOffM") && has("recoveryM") && has("loansM")) { derivedKeys.push("nco"); parsed.forEach((r) => { r.nco = r.loansM > 0 ? ((r.chargeOffM - r.recoveryM) / r.loansM) * 100 : 0; }); }
  if (!has("py") && has("grossIncomeM") && has("assetsM")) { derivedKeys.push("py"); parsed.forEach((r) => { r.py = r.assetsM > 0 ? (r.grossIncomeM / r.assetsM) * 100 : 0; }); }

  // New-data operating analytics are derived from fields explicitly present in the ACUI extract.
  if (has("operatingExpenseM") && has("grossIncomeM")) parsed.forEach((r) => { r.efficiencyRatio = r.grossIncomeM > 0 ? (r.operatingExpenseM / r.grossIncomeM) * 100 : 0; });
  if (has("niM") && has("grossIncomeM")) parsed.forEach((r) => { r.netMargin = r.grossIncomeM !== 0 ? (r.niM / r.grossIncomeM) * 100 : 0; });
  if (has("chargeOffM") && has("recoveryM")) parsed.forEach((r) => { r.recoveryRate = r.chargeOffM > 0 ? (r.recoveryM / r.chargeOffM) * 100 : 0; });

  // De-duplicate inside a quarter. Charter number when available, otherwise name + city + state
  // (name + state alone would merge different credit unions such as the two "CITY" CUs in Texas).
  const seen = new Map();
  let duplicates = 0;
  parsed.forEach((rec) => {
    const key = rec.charter ? `#${rec.charter}` : `${rec.name}|${rec.city}|${rec.state}`.toLowerCase();
    if (seen.has(key)) duplicates += 1;
    seen.set(key, rec);
  });
  const records = [...seen.values()].map((rec, i) => {
    const { _bucket, ...rest } = rec;
    const b = Number(_bucket);
    const bucket = has("bucket") && Number.isInteger(b) && b >= 0 && b <= 7 ? b : inferBucket(rec.assetsM);
    return { ...rest, idx: i, bucket, quarter };
  });

  const report = {
    total: list.length,
    skipped,
    duplicates,
    mapped: IMPORT_FIELDS.filter((f) => has(f.key)).length,
    totalFields: IMPORT_FIELDS.length,
    derived: derivedKeys.map((k) => IMPORT_FIELDS.find((f) => f.key === k).label),
    missing: IMPORT_FIELDS.filter((f) => NUMERIC_KEYS.includes(f.key) && !has(f.key) && !derivedKeys.includes(f.key)).map((f) => f.label),
    notes,
  };
  return { records, report };
}

function inferBucket(assetsM) {
  if (assetsM < 50) return 0;
  if (assetsM < 100) return 1;
  if (assetsM < 250) return 2;
  if (assetsM < 500) return 3;
  if (assetsM < 1000) return 4;
  if (assetsM < 5000) return 5;
  if (assetsM < 10000) return 6;
  return 7;
}

function parseDelimitedText(text) {
  const workbook = XLSX.read(text, { type: "string", raw: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, blankrows: false });
  if (!rows.length) return [];
  const seen = new Map();
  const headers = rows[0].map((h, i) => {
    const base = String(h ?? "").replace(/^\uFEFF/, "").trim() || `Column ${i + 1}`;
    const count = (seen.get(base) || 0) + 1;
    seen.set(base, count);
    return count > 1 ? `${base} (${count})` : base;
  });
  return rows.slice(1).map((row) => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));
}

function downloadQuarterTemplate() {
  const headers = ["name","city","state","bucket","ceo","website","assetsM","members","branches","employees","loansM","depositsM","niM","equityM","roa","roe","ltd","lta","dr","nco","py","capRatio","niiM","grossIncomeM","operatingExpenseM","delinquentAmountM","delinquentCount","chargeOffM","recoveryM"];
  const ws = XLSX.utils.aoa_to_sheet([headers]);
  ws["!cols"] = headers.map((h) => ({ wch: Math.max(10, h.length + 2) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "CU360 data");
  XLSX.writeFile(wb, "cu360-quarter-template.xlsx");
}


/* ============================== FORMAT HELPERS ============================== */
const fmtM = (v) => {
  if (v == null || isNaN(v)) return "—";
  if (Math.abs(v) >= 1000) return `$${(v / 1000).toFixed(2)}B`;
  return `$${v.toFixed(1)}M`;
};
const fmtNum = (v) => (v == null || isNaN(v) ? "—" : Math.round(v).toLocaleString("en-US"));
const fmtPct = (v, d = 2) => (v == null || isNaN(v) ? "—" : `${v.toFixed(d)}%`);
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

/* Metrics where a LOWER value is the better outcome */
const LOWER_IS_BETTER = new Set(["dr", "nco", "ltd"]);

function percentileRank(peerValues, value) {
  if (!peerValues.length) return null;
  const below = peerValues.filter((v) => v <= value).length;
  return (below / peerValues.length) * 100;
}

function stat(peerValues) {
  if (!peerValues.length) return { median: 0, p25: 0, p75: 0, min: 0, max: 0 };
  const s = [...peerValues].sort((a, b) => a - b);
  const q = (p) => s[clamp(Math.floor(p * (s.length - 1)), 0, s.length - 1)];
  return { median: q(0.5), p25: q(0.25), p75: q(0.75), min: s[0], max: s[s.length - 1] };
}

const METRIC_META = {
  roa: { label: "Return on Assets", suffix: "%", decimals: 2, desc: "Net income as a share of average assets." },
  roe: { label: "Return on Equity", suffix: "%", decimals: 2, desc: "Net income as a share of member equity." },
  capRatio: { label: "Net Worth Ratio", suffix: "%", decimals: 2, desc: "Equity capital as a share of total assets. NCUA well-capitalized threshold is 7.00%." },
  dr: { label: "Delinquency Rate", suffix: "%", decimals: 2, desc: "Share of the loan portfolio past due." },
  nco: { label: "Net Charge-Off Rate", suffix: "%", decimals: 2, desc: "Net loan losses as a share of average loans." },
  ltd: { label: "Loan-to-Share Ratio", suffix: "%", decimals: 1, desc: "Total loans relative to total member deposits." },
  lta: { label: "Loan-to-Asset Ratio", suffix: "%", decimals: 1, desc: "Total loans relative to total assets." },
  py: { label: "Portfolio Yield", suffix: "%", decimals: 2, desc: "Income / asset yield proxy derived from gross income in the supplied quarter." },
};

/* ============================== PEER GROUP DESCRIPTION ============================== */
function peerGroupText(cu, opts, peerCount) {
  if (!cu) return "";
  const lo = opts.wideTier ? Math.max(0, cu.bucket - 1) : cu.bucket;
  const hi = opts.wideTier ? Math.min(7, cu.bucket + 1) : cu.bucket;
  const sizeText = lo === hi ? BUCKET_LABELS[lo] : `${BUCKET_LABELS[lo].split(" ")[0]}–${BUCKET_LABELS[hi]}`;
  const geo = opts.sameState ? `in ${cu.state}` : "nationwide";
  return `${sizeText} in assets, ${geo} — ${peerCount.toLocaleString()} credit union${peerCount === 1 ? "" : "s"}`;
}

function PeerGroupBanner({ cu, opts, peerCount }) {
  return (
    <div className="peer-banner">
      <Users size={14} strokeWidth={1.8} />
      <span><strong>Peer group:</strong> {peerGroupText(cu, opts, peerCount)}</span>
    </div>
  );
}

/* ============================== PEER GROUP + RISK LOGIC ============================== */
function usePeerGroup(cu, opts) {
  return useMemo(() => {
    if (!cu) return [];
    const lo = opts.wideTier ? Math.max(0, cu.bucket - 1) : cu.bucket;
    const hi = opts.wideTier ? Math.min(7, cu.bucket + 1) : cu.bucket;
    return ACTIVE_CUS.filter((c) => {
      if (c.idx === cu.idx) return false;
      if (c.bucket < lo || c.bucket > hi) return false;
      if (opts.sameState && c.state !== cu.state) return false;
      return true;
    });
  }, [cu, opts.wideTier, opts.sameState, DATA_VERSION]);
}

function computeRiskFlags(cu, peers) {
  const flags = [];
  const drVals = peers.map((p) => p.dr);
  const ncoVals = peers.map((p) => p.nco);
  const drStat = stat(drVals);
  const ncoStat = stat(ncoVals);

  if (cu.capRatio < 2) {
    flags.push({ sev: "critical", title: "Critically undercapitalized", detail: `Net worth ratio of ${fmtPct(cu.capRatio)} is below the NCUA critically-undercapitalized threshold of 2.00%.` });
  } else if (cu.capRatio < 4) {
    flags.push({ sev: "high", title: "Significantly undercapitalized", detail: `Net worth ratio of ${fmtPct(cu.capRatio)} falls in NCUA's significantly-undercapitalized band (2.00%–3.99%).` });
  } else if (cu.capRatio < 6) {
    flags.push({ sev: "high", title: "Undercapitalized", detail: `Net worth ratio of ${fmtPct(cu.capRatio)} is below the 6.00% adequately-capitalized floor.` });
  } else if (cu.capRatio < 7) {
    flags.push({ sev: "medium", title: "Below well-capitalized threshold", detail: `Net worth ratio of ${fmtPct(cu.capRatio)} sits just under NCUA's 7.00% well-capitalized bar.` });
  }

  if (cu.roa < 0) {
    flags.push({ sev: "high", title: "Negative earnings", detail: `Return on assets of ${fmtPct(cu.roa)} indicates a net loss for the period.` });
  } else if (cu.roa < 0.2) {
    flags.push({ sev: "medium", title: "Thin earnings margin", detail: `ROA of ${fmtPct(cu.roa)} leaves little cushion to build net worth or absorb a downturn.` });
  }

  if (cu.dr > Math.max(drStat.p75, 1.5)) {
    flags.push({ sev: "high", title: "Elevated delinquency", detail: `Delinquency rate of ${fmtPct(cu.dr)} exceeds the peer 75th percentile of ${fmtPct(drStat.p75)}.` });
  }
  if (cu.nco > Math.max(ncoStat.p75, 0.5)) {
    flags.push({ sev: "medium", title: "Elevated charge-offs", detail: `Net charge-off rate of ${fmtPct(cu.nco)} exceeds the peer 75th percentile of ${fmtPct(ncoStat.p75)}.` });
  }

  if (cu.ltd > 100) {
    flags.push({ sev: "medium", title: "Loans exceed deposits", detail: `Loan-to-share ratio of ${fmtPct(cu.ltd, 1)} means lending is funded partly by borrowings or wholesale sources rather than member deposits.` });
  }
  if (cu.lta > 92) {
    flags.push({ sev: "low", title: "Limited balance-sheet flexibility", detail: `Loan-to-asset ratio of ${fmtPct(cu.lta, 1)} leaves a thin cushion of liquid assets.` });
  }
  if (cu.ltd < 35) {
    flags.push({ sev: "low", title: "Excess liquidity, low loan demand", detail: `Loan-to-share ratio of ${fmtPct(cu.ltd, 1)} suggests idle deposits that could be deployed into higher-yielding loans or investments.` });
  }
  return flags;
}

const SEV_ORDER = { critical: 0, high: 1, medium: 2, low: 3 };
const SEV_STYLE = {
  critical: { color: "#E5484D", bg: "rgba(229,72,77,0.12)", label: "Critical" },
  high: { color: "#D65A4A", bg: "rgba(214,90,74,0.12)", label: "High" },
  medium: { color: "#D6A94A", bg: "rgba(214,169,74,0.13)", label: "Medium" },
  low: { color: "#8FA0BD", bg: "rgba(143,160,189,0.12)", label: "Low" },
};

/* ============================== SMALL UI PRIMITIVES ============================== */
function PercentileBar({ value, lowerIsBetter }) {
  const displayed = value == null ? null : clamp(lowerIsBetter ? 100 - value : value, 0, 100);
  const good = displayed != null && displayed >= 55;
  const bad = displayed != null && displayed <= 30;
  const color = displayed == null ? "#5D6472" : bad ? "#D65A4A" : good ? "#4C9A6A" : "#D6A94A";
  return (
    <div className="pbar">
      <div className="pbar-track">
        <div className="pbar-fill" style={{ width: `${displayed ?? 0}%`, background: color }} />
        <div className="pbar-mid" />
      </div>
    </div>
  );
}

function Delta({ value, suffix = "", invert = false, decimals = 1 }) {
  if (value == null || isNaN(value)) return <span className="delta flat">—</span>;
  const good = invert ? value < 0 : value > 0;
  const flat = Math.abs(value) < 0.01;
  const Icon = flat ? Minus : value > 0 ? ArrowUpRight : ArrowDownRight;
  return (
    <span className={`delta ${flat ? "flat" : good ? "good" : "bad"}`}>
      <Icon size={13} strokeWidth={2.5} />
      {Math.abs(value).toFixed(decimals)}{suffix}
    </span>
  );
}

function SeverityDot({ sev }) {
  const s = SEV_STYLE[sev];
  return <span className="sev-dot" style={{ background: s.color }} />;
}

function EmptyState({ icon: Icon, title, body }) {
  return (
    <div className="empty-state">
      <Icon size={28} strokeWidth={1.5} />
      <h3>{title}</h3>
      <p>{body}</p>
    </div>
  );
}

/* ============================== CU SEARCH ============================== */
function CUSearch({ onSelect, placeholder }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const boxRef = useRef(null);

  const results = useMemo(() => {
    if (q.trim().length < 2) return [];
    const needle = q.trim().toLowerCase();
    const matches = ACTIVE_CUS.filter((c) => c.name.toLowerCase().includes(needle));
    return matches.sort((a, b) => b.assetsM - a.assetsM).slice(0, 8);
  }, [q, DATA_VERSION]);

  useEffect(() => {
    function onClick(e) {
      if (boxRef.current && !boxRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, []);

  return (
    <div className="cu-search" ref={boxRef}>
      <Search size={16} strokeWidth={2} className="cu-search-icon" />
      <input
        value={q}
        onChange={(e) => { setQ(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
        placeholder={placeholder || "Search your credit union by name…"}
      />
      {open && results.length > 0 && (
        <div className="cu-search-results">
          {results.map((c) => (
            <button key={c.idx} className="cu-search-row" onClick={() => { onSelect(c); setQ(""); setOpen(false); }}>
              <span className="cu-search-name">{titleCase(c.name)}</span>
              <span className="cu-search-meta">{c.state} · {fmtM(c.assetsM)}</span>
            </button>
          ))}
        </div>
      )}
      {open && q.trim().length >= 2 && results.length === 0 && (
        <div className="cu-search-results"><div className="cu-search-empty">No credit union matches "{q}".</div></div>
      )}
    </div>
  );
}

function titleCase(s) {
  if (!s) return s;
  const smallWords = new Set(["of", "and", "the", "for", "a", "in"]);
  return s.toLowerCase().split(" ").map((w, i) => {
    if (w.length === 0) return w;
    if (smallWords.has(w) && i !== 0) return w;
    if (/^\d/.test(w)) return w.toUpperCase();
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(" ").replace(/\bfcu\b/i, "FCU").replace(/\bncua\b/i, "NCUA");
}
/* ============================== NEW DATA ANALYTICS ============================== */
function NewDataAnalytics({ cu, peers }) {
  const peerEfficiency = stat(peers.map((p) => Number(p.efficiencyRatio)).filter(Number.isFinite)).median;
  const peerMargin = stat(peers.map((p) => Number(p.netMargin)).filter(Number.isFinite)).median;
  const peerRecovery = stat(peers.map((p) => Number(p.recoveryRate)).filter(Number.isFinite)).median;
  const delinquentShare = cu.loansM > 0 ? (cu.delinquentAmountM / cu.loansM) * 100 : null;
  return (
    <div className="panel">
      <div className="panel-head">
        <h3>New-data operating analytics</h3>
        <span className="panel-head-sub">Derived from the supplied {ACTIVE_QUARTER} ACUI fields</span>
      </div>
      <div className="kpi-grid">
        <KPI label="Efficiency Ratio" value={fmtPct(cu.efficiencyRatio)} sub={`Peer median ${fmtPct(peerEfficiency)}`} warn={peerEfficiency != null && cu.efficiencyRatio > peerEfficiency} />
        <KPI label="Net Income Margin" value={fmtPct(cu.netMargin)} sub={`Peer median ${fmtPct(peerMargin)}`} warn={cu.netMargin < 0} />
        <KPI label="Delinquent Loans" value={fmtM(cu.delinquentAmountM)} sub={`${fmtPct(delinquentShare)} of total loans`} warn={delinquentShare != null && delinquentShare > stat(peers.map((p) => p.dr)).p75} />
        <KPI label="Recovery Rate" value={fmtPct(cu.recoveryRate)} sub={`Peer median ${fmtPct(peerRecovery)}`} />
      </div>
      <div className="compare-list">
        <div className="compare-row"><div className="compare-label">Gross income</div><div className="compare-values"><span className="compare-you">{fmtM(cu.grossIncomeM)}</span><span className="compare-vs">Operating expense {fmtM(cu.operatingExpenseM)}</span></div></div>
        <div className="compare-row"><div className="compare-label">Charge-offs / recoveries</div><div className="compare-values"><span className="compare-you">{fmtM(cu.chargeOffM)}</span><span className="compare-vs">Recovered {fmtM(cu.recoveryM)}</span></div></div>
        <div className="compare-row"><div className="compare-label">Delinquent loan count</div><div className="compare-values"><span className="compare-you">{fmtNum(cu.delinquentCount)}</span><span className="compare-vs">Loan-to-asset {fmtPct(cu.lta, 1)}</span></div></div>
      </div>
    </div>
  );
}

/* ============================== OVERVIEW TAB ============================== */
function OverviewTab({ cu, peers, opts }) {
  const metrics = ["roa", "capRatio", "dr", "ltd"];
  return (
    <div className="tab-body">
      <PeerGroupBanner cu={cu} opts={opts} peerCount={peers.length} />
      <div className="hero-row">
        <div className="hero-block">
          <span className="eyebrow-tag">{BUCKET_LABELS[cu.bucket]} peer tier</span>
          <h1>{titleCase(cu.name)}</h1>
          <p className="hero-sub"><MapPin size={14} /> {cu.state} &nbsp;·&nbsp; {fmtNum(cu.branches)} branches &nbsp;·&nbsp; {cu.employees ? `${fmtNum(cu.employees)} FTEs` : "FTE data not provided"}</p>
        </div>
        <div className="hero-figure">
          <span className="hero-figure-label">Total assets</span>
          <span className="hero-figure-value">{fmtM(cu.assetsM)}</span>
          <span className="hero-figure-sub">{fmtNum(cu.members)} members · {(cu.assetsM * 1e6 / cu.members).toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 })} in assets per member</span>{cu.ceo ? <span className="hero-figure-sub">CEO: {titleCase(cu.ceo)}</span> : null}{cu.website ? <a className="hero-figure-sub" href={cu.website} target="_blank" rel="noreferrer">Visit website ↗</a> : null}
        </div>
      </div>

      <div className="kpi-grid">
        <KPI label="Net Income (period)" value={fmtM(cu.niM)} sub={`${fmtPct(cu.roa)} ROA`} />
        <KPI label="Net Worth Ratio" value={fmtPct(cu.capRatio)} sub={cu.capRatio >= 7 ? "Well capitalized" : "Below 7.00% threshold"} warn={cu.capRatio < 7} />
        <KPI label="Delinquency Rate" value={fmtPct(cu.dr)} sub={`Peer median ${fmtPct(stat(peers.map(p=>p.dr)).median)}`} warn={cu.dr > stat(peers.map(p=>p.dr)).p75} />
        <KPI label="Loan-to-Share" value={fmtPct(cu.ltd, 1)} sub={`Peer median ${fmtPct(stat(peers.map(p=>p.ltd)).median, 1)}`} />
      </div>

      <NewDataAnalytics cu={cu} peers={peers} />

      <div className="panel">
        <div className="panel-head"><h3>Leadership & digital presence</h3><span className="panel-head-sub">Institution details from the active quarter</span></div>
        <div className="demographic-profile wide">
          <div><span>CEO</span><strong>{cu.ceo ? titleCase(cu.ceo) : "Not provided"}</strong></div>
          <div><span>Website</span><strong>{cu.website && /^https?:\/\//i.test(cu.website) ? <a href={cu.website} target="_blank" rel="noreferrer">Open ↗</a> : "Not provided"}</strong></div>
          <div><span>Headquarters state</span><strong>{cu.state || "Not provided"}</strong></div>
          <div><span>Branches</span><strong>{fmtNum(cu.branches)}</strong></div>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head">
          <h3>Where {shortName(cu.name)} stands against {peers.length.toLocaleString()} peers</h3>
          <span className="panel-head-sub">{BUCKET_LABELS[cu.bucket]} in assets</span>
        </div>
        <div className="metric-rank-list">
          {metrics.map((m) => {
            const meta = METRIC_META[m];
            const peerVals = peers.map((p) => p[m]);
            const pct = percentileRank(peerVals, cu[m]);
            const displayPct = pct == null ? null : LOWER_IS_BETTER.has(m) ? 100 - pct : pct;
            return (
              <div className="metric-rank-row" key={m}>
                <div className="metric-rank-label">
                  <span>{meta.label}</span>
                  <span className="metric-rank-value">{fmtPct(cu[m], meta.decimals)}</span>
                </div>
                <PercentileBar value={pct} lowerIsBetter={LOWER_IS_BETTER.has(m)} />
                <span className="metric-rank-pct">{displayPct == null ? "—" : `${Math.round(displayPct)}th pct.`}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function shortName(n) {
  const t = titleCase(n);
  return t.length > 28 ? t.slice(0, 26) + "…" : t;
}

function KPI({ label, value, sub, warn }) {
  return (
    <div className={`kpi ${warn ? "kpi-warn" : ""}`}>
      <span className="kpi-label">{label}</span>
      <span className="kpi-value">{value}</span>
      <span className="kpi-sub">{sub}</span>
    </div>
  );
}

/* ============================== PEER BENCHMARKING TAB ============================== */
function BenchmarkTab({ cu, peers, opts, setOpts }) {
  const metrics = Object.keys(METRIC_META);
  const radarData = metrics.map((m) => {
    const peerVals = peers.map((p) => p[m]);
    const pct = percentileRank(peerVals, cu[m]);
    const displayPct = pct == null ? null : LOWER_IS_BETTER.has(m) ? 100 - pct : pct;
    return { metric: METRIC_META[m].label.replace("Return on", "Ret. on").replace(" Rate", ""), value: displayPct == null ? null : Math.round(displayPct) };
  });

  const barData = metrics.map((m) => {
    const peerVals = peers.map((p) => p[m]);
    const s = stat(peerVals);
    return { metric: METRIC_META[m].label, You: cu[m], Peers: s.median };
  });

  return (
    <div className="tab-body">
      <PeerGroupBanner cu={cu} opts={opts} peerCount={peers.length} />
      <PeerOptsBar cu={cu} opts={opts} setOpts={setOpts} peerCount={peers.length} />

      <div className="grid-2">
        <div className="panel">
          <div className="panel-head"><h3>Percentile profile</h3><span className="panel-head-sub">100 = strongest among peers</span></div>
          <ResponsiveContainer width="100%" height={300}>
            <RadarChart data={radarData} outerRadius={100}>
              <PolarGrid stroke="rgba(255,255,255,0.1)" />
              <PolarAngleAxis dataKey="metric" tick={{ fill: "#8FA0BD", fontSize: 11 }} />
              <PolarRadiusAxis angle={90} domain={[0, 100]} tick={{ fill: "#586580", fontSize: 9 }} axisLine={false} />
              <Radar dataKey="value" stroke="#C9A227" fill="#C9A227" fillOpacity={0.32} strokeWidth={2} />
              <Tooltip contentStyle={tooltipStyle} formatter={(v) => [`${v}th percentile`, ""]} />
            </RadarChart>
          </ResponsiveContainer>
        </div>

        <div className="panel">
          <div className="panel-head"><h3>You vs. peer median</h3><span className="panel-head-sub">{BUCKET_LABELS[cu.bucket]}{opts.sameState ? `, ${cu.state}` : ""}</span></div>
          <div className="compare-list">
            {metrics.map((m) => {
              const meta = METRIC_META[m];
              const peerVals = peers.map((p) => p[m]);
              const s = stat(peerVals);
              const diff = cu[m] - s.median;
              const invert = LOWER_IS_BETTER.has(m);
              return (
                <div className="compare-row" key={m}>
                  <div className="compare-label" title={meta.desc}>{meta.label}</div>
                  <div className="compare-values">
                    <span className="compare-you">{fmtPct(cu[m], meta.decimals)}</span>
                    <span className="compare-vs">vs {fmtPct(s.median, meta.decimals)} median</span>
                    <Delta value={diff} suffix=" pts" invert={invert} decimals={meta.decimals} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Peer group detail</h3><span className="panel-head-sub">{peers.length.toLocaleString()} credit unions</span></div>
        <PeerTable cu={cu} peers={peers} />
      </div>
    </div>
  );
}

const tooltipStyle = { background: "#171B22", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 6, color: "#E9E6DD", fontSize: 12 };

function PeerOptsBar({ opts, setOpts, peerCount }) {
  return (
    <div className="opts-bar">
      <button className={`chip ${opts.sameState ? "chip-on" : ""}`} onClick={() => setOpts((o) => ({ ...o, sameState: !o.sameState }))}>
        {opts.sameState ? <Check size={13} /> : null} Same state only
      </button>
      <button className={`chip ${opts.wideTier ? "chip-on" : ""}`} onClick={() => setOpts((o) => ({ ...o, wideTier: !o.wideTier }))}>
        {opts.wideTier ? <Check size={13} /> : null} Include adjacent size tiers
      </button>
      <span className="opts-count">{peerCount.toLocaleString()} peers in group</span>
    </div>
  );
}

function PeerTable({ cu, peers, limit = 25 }) {
  const [sortKey, setSortKey] = useState("assetsM");
  const [sortDir, setSortDir] = useState(-1);
  const rows = useMemo(() => {
    const all = [cu, ...peers];
    const sorted = [...all].sort((a, b) => (a[sortKey] - b[sortKey]) * sortDir);
    return sorted.slice(0, limit);
  }, [cu, peers, sortKey, sortDir, limit]);

  const cols = [
    { key: "name", label: "Credit union" },
    { key: "assetsM", label: "Assets" },
    { key: "members", label: "Members" },
    { key: "roa", label: "ROA" },
    { key: "capRatio", label: "Net Worth" },
    { key: "dr", label: "Delinquency" },
    { key: "ltd", label: "Loan/Share" },
  ];

  function toggleSort(k) {
    if (k === sortKey) setSortDir((d) => -d);
    else { setSortKey(k); setSortDir(-1); }
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {cols.map((c) => (
              <th key={c.key} onClick={() => toggleSort(c.key)} className={sortKey === c.key ? "th-active" : ""}>
                {c.label} {sortKey === c.key ? (sortDir === -1 ? "▾" : "▴") : ""}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.idx} className={r.idx === cu.idx ? "row-you" : ""}>
              <td className="td-name">{titleCase(r.name)}{r.idx === cu.idx && <span className="you-tag">YOU</span>}
                <div className="td-sub">{r.state}</div>
              </td>
              <td className="mono">{fmtM(r.assetsM)}</td>
              <td className="mono">{fmtNum(r.members)}</td>
              <td className="mono">{fmtPct(r.roa)}</td>
              <td className="mono">{fmtPct(r.capRatio)}</td>
              <td className="mono">{fmtPct(r.dr)}</td>
              <td className="mono">{fmtPct(r.ltd, 1)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
/* ============================== COMPETITIVE INTELLIGENCE TAB ============================== */
function CompetitiveTab({ cu }) {
  const [scope, setScope] = useState("state");
  const [rivalQuery, setRivalQuery] = useState("");
  const [selectedRivalIds, setSelectedRivalIds] = useState([]);
  useEffect(() => {
    setSelectedRivalIds([]);
    setRivalQuery("");
  }, [cu.charter, cu.name, cu.state, DATA_VERSION]);
  const marketRows = useMemo(() => ACTIVE_CUS.filter((c) => {
    if (scope === "state") return c.state === cu.state;
    if (scope === "peers") return Math.abs(c.bucket - cu.bucket) <= 1;
    return true;
  }), [cu, scope, DATA_VERSION]);
  const peerRows = useMemo(() => marketRows.filter((c) => c.idx !== cu.idx), [marketRows, cu.idx]);
  const candidates = useMemo(() => peerRows.filter((c) => !selectedRivalIds.includes(c.idx)), [peerRows, selectedRivalIds]);
  const suggestions = useMemo(() => rivalQuery.trim().length >= 2 && selectedRivalIds.length < 3
    ? candidates.filter((c) => c.name.toLowerCase().includes(rivalQuery.trim().toLowerCase())).slice(0, 6)
    : [], [candidates, rivalQuery, selectedRivalIds.length]);
  const selectedRivals = useMemo(() => selectedRivalIds.map((idx) => peerRows.find((c) => c.idx === idx)).filter(Boolean), [peerRows, selectedRivalIds]);
  const nearestRivals = useMemo(() => [...peerRows].sort((a, b) => distScore(a, cu) - distScore(b, cu)).slice(0, 3), [peerRows, cu]);
  const comparisonCUs = selectedRivals.length ? selectedRivals : nearestRivals;
  const marketTotals = useMemo(() => ({
    assets: marketRows.reduce((sum, c) => sum + (Number(c.assetsM) || 0), 0),
    members: marketRows.reduce((sum, c) => sum + (Number(c.members) || 0), 0),
    assetRank: 1 + marketRows.filter((c) => c.assetsM > cu.assetsM).length,
    assetPercentile: percentileRank(marketRows.map((c) => c.assetsM), cu.assetsM) ?? 0,
  }), [marketRows, cu.assetsM]);
  const marketAssets = marketTotals.assets;
  const marketMembers = marketTotals.members;
  const assetRank = marketTotals.assetRank;
  const assetPercentile = marketTotals.assetPercentile;
  const assetShare = marketAssets > 0 ? (cu.assetsM / marketAssets) * 100 : 0;
  const memberShare = marketMembers > 0 ? (cu.members / marketMembers) * 100 : 0;
  const peerBenchmarks = useMemo(() => peerRows.length ? {
    efficiency: stat(peerRows.map((c) => c.efficiencyRatio)).median,
    capital: stat(peerRows.map((c) => c.capRatio)).median,
    delinquencyP75: stat(peerRows.map((c) => c.dr)).p75,
  } : { efficiency: null, capital: null, delinquencyP75: null }, [peerRows]);
  const peerEfficiency = peerBenchmarks.efficiency;
  const peerCapital = peerBenchmarks.capital;
  const delinquencyP75 = peerBenchmarks.delinquencyP75;
  const marketChartRows = useMemo(() => {
    const largest = [...marketRows].sort((a, b) => b.assetsM - a.assetsM).slice(0, 6);
    return largest.some((c) => c.idx === cu.idx) ? largest : [...largest.slice(0, 5), cu].sort((a, b) => b.assetsM - a.assetsM);
  }, [marketRows, cu]);
  const selectedLabel = scope === "state" ? `${cu.state} headquarters market` : scope === "peers" ? `Asset tiers ${Math.max(0, cu.bucket - 1) + 1}-${Math.min(7, cu.bucket + 1) + 1} (${BUCKET_LABELS[Math.max(0, cu.bucket - 1)]} through ${BUCKET_LABELS[Math.min(7, cu.bucket + 1)]}), nationwide` : "Nationwide market";
  const mixRows = useMemo(() => [
    { name: "Loans / assets", You: cu.assetsM > 0 ? (cu.loansM / cu.assetsM) * 100 : 0, Peers: peerRows.length ? stat(peerRows.map((c) => c.assetsM > 0 ? (c.loansM / c.assetsM) * 100 : 0)).median : 0 },
    { name: "Deposits / assets", You: cu.assetsM > 0 ? (cu.depositsM / cu.assetsM) * 100 : 0, Peers: peerRows.length ? stat(peerRows.map((c) => c.assetsM > 0 ? (c.depositsM / c.assetsM) * 100 : 0)).median : 0 },
    { name: "Net worth / assets", You: cu.capRatio, Peers: peerCapital ?? 0 },
  ], [cu, peerRows, peerCapital]);
  const comparisonMetrics = useMemo(() => [
    { key: "assetsM", label: "Total assets", format: fmtM },
    { key: "members", label: "Members", format: fmtNum },
    { key: "efficiencyRatio", label: "Efficiency ratio", format: (v) => fmtPct(v) },
    { key: "capRatio", label: "Net worth ratio", format: (v) => fmtPct(v) },
    { key: "roa", label: "Return on assets", format: (v) => fmtPct(v) },
    { key: "dr", label: "Delinquency rate", format: (v) => fmtPct(v) },
    { key: "ltd", label: "Loan-to-share", format: (v) => fmtPct(v, 1) },
  ], []);
  const scorecardMedians = useMemo(() => Object.fromEntries(comparisonMetrics.map((metric) => [
    metric.key,
    peerRows.length ? stat(peerRows.map((c) => Number(c[metric.key])).filter(Number.isFinite)).median : null,
  ])), [comparisonMetrics, peerRows]);
  const signals = [
    { risk: cu.capRatio < 7, title: cu.capRatio < 7 ? "Capital below 7%" : "Capital above 7% floor", detail: `Net worth ratio ${fmtPct(cu.capRatio)}; peer median ${fmtPct(peerCapital)}.` },
    { risk: peerEfficiency != null && cu.efficiencyRatio > peerEfficiency, title: cu.efficiencyRatio > peerEfficiency ? "Efficiency cost pressure" : "Efficiency at or below peer median", detail: `Efficiency ratio ${fmtPct(cu.efficiencyRatio)}; peer median ${fmtPct(peerEfficiency)}. Lower is generally better.` },
    { risk: delinquencyP75 != null && cu.dr > delinquencyP75, title: cu.dr > delinquencyP75 ? "Delinquency above peer 75th percentile" : "Delinquency within peer range", detail: `Delinquency ${fmtPct(cu.dr)}; peer 75th percentile ${fmtPct(delinquencyP75)}.` },
  ];

  return (
    <div className="tab-body">
      <div className="trend-hero competitive-hero"><div><span className="eyebrow-tag"><MapPin size={12}/> Competitive position · {ACTIVE_QUARTER}</span><h1>Competitive intelligence</h1><p>Benchmark market scale, operating performance, capital, and credit risk against a clearly defined comparison universe.</p></div><div className="competitive-scope"><span>Comparison universe</span><div className="competitive-scope-buttons" role="group" aria-label="Comparison universe">
        {[{ id: "state", label: "State market" }, { id: "peers", label: "Similar-size peers" }, { id: "national", label: "Nationwide" }].map((option) => <button key={option.id} className={`chip ${scope === option.id ? "chip-on" : ""}`} onClick={() => { setScope(option.id); setSelectedRivalIds([]); }}>{option.label}</button>)}
      </div><small>{selectedLabel} · {marketRows.length.toLocaleString()} institutions</small></div></div>
      <div className="kpi-grid">
        <KPI label="Asset rank" value={`#${assetRank || "—"}`} sub={`of ${marketRows.length.toLocaleString()} · ${Math.round(assetPercentile)}th percentile`} />
        <KPI label={scope === "state" ? "In-state asset share" : "Asset share of selected universe"} value={fmtPct(assetShare, assetShare < 1 ? 3 : 1)} sub={`${fmtM(marketAssets)} combined assets`} />
        <KPI label="Member share of universe" value={fmtPct(memberShare, memberShare < 1 ? 3 : 1)} sub={`${fmtNum(marketMembers)} members combined`} />
        <KPI label="Efficiency vs peer median" value={peerEfficiency == null ? "—" : `${cu.efficiencyRatio - peerEfficiency >= 0 ? "+" : ""}${(cu.efficiencyRatio - peerEfficiency).toFixed(2)} pp`} sub={`CU ${fmtPct(cu.efficiencyRatio)} · peers ${fmtPct(peerEfficiency)} · lower is better`} warn={peerEfficiency != null && cu.efficiencyRatio > peerEfficiency} />
      </div>

      <div className="grid-2">
        <div className="panel">
          <div className="panel-head"><h3>Market scale by assets</h3><span className="panel-head-sub">Largest institutions in selected universe · your CU highlighted</span></div>
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={marketChartRows.map((c) => ({ name: shortName(c.name), Assets: Math.round(c.assetsM), you: c.idx === cu.idx }))} layout="vertical" margin={{ left: 8, right: 24 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.08)" horizontal={false} />
              <XAxis type="number" tick={{ fill: "#8FA0BD", fontSize: 11 }} tickFormatter={(v) => fmtM(v)} />
              <YAxis type="category" dataKey="name" width={140} tick={{ fill: "#8FA0BD", fontSize: 11 }} />
              <Tooltip contentStyle={tooltipStyle} formatter={(v) => [fmtM(v), "Assets"]} />
              <Bar dataKey="Assets" radius={[0, 3, 3, 0]}>
                {marketChartRows.map((c, i) => <Cell key={i} fill={c.idx === cu.idx ? "#C9A227" : "#3A4457"} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="panel">
          <div className="panel-head"><h3>Balance-sheet structure ratios</h3><span className="panel-head-sub">CU vs peer median · independent ratios, not a 100% mix</span></div>
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={mixRows} margin={{ left: 0, right: 12, bottom: 12 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.08)" vertical={false} />
              <XAxis dataKey="name" tick={{ fill: "#8FA0BD", fontSize: 10 }} interval={0} />
              <YAxis tick={{ fill: "#8FA0BD", fontSize: 10 }} tickFormatter={(v) => `${v}%`} />
              <Tooltip contentStyle={tooltipStyle} formatter={(v) => [`${Number(v).toFixed(1)}%`]} />
              <Legend />
              <Bar dataKey="You" fill="#C9A227" radius={[3, 3, 0, 0]} />
              <Bar dataKey="Peers" fill="#648E91" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid-2">
        <div className="panel">
          <div className="panel-head"><div><h3>Competitive scorecard</h3><span className="panel-head-sub">Select up to 3 institutions from the active universe; default view uses closest asset-size rivals.</span></div></div>
          <div className="competitive-search"><input value={rivalQuery} onChange={(e) => setRivalQuery(e.target.value)} placeholder="Find a credit union to compare…" aria-label="Find a credit union to compare" />{suggestions.length > 0 && <div className="competitive-suggestions">{suggestions.map((c) => <button key={c.idx} onClick={() => { setSelectedRivalIds((ids) => ids.length < 3 ? [...ids, c.idx] : ids); setRivalQuery(""); }}><span>{titleCase(c.name)}</span><small>{c.state} · {fmtM(c.assetsM)}</small></button>)}</div>}</div>
          {selectedRivals.length > 0 && <div className="selected-rivals">{selectedRivals.map((c) => <button key={c.idx} className="chip chip-on" onClick={() => setSelectedRivalIds((ids) => ids.filter((idx) => idx !== c.idx))} title={`Remove ${titleCase(c.name)}`}>{shortName(c.name)} <X size={12}/></button>)}</div>}
          <div className="table-wrap"><table className="competitive-table"><thead><tr><th>Metric</th><th>{shortName(cu.name)} · you</th>{comparisonCUs.map((c) => <th key={c.idx}>{shortName(c.name)}</th>)}<th>Peer median</th></tr></thead><tbody>{comparisonMetrics.map((metric) => <tr key={metric.key}><td>{metric.label}</td><td className="mono">{metric.format(cu[metric.key])}</td>{comparisonCUs.map((c) => <td key={c.idx} className="mono">{metric.format(c[metric.key])}</td>)}<td className="mono">{scorecardMedians[metric.key] == null ? "—" : metric.format(scorecardMedians[metric.key])}</td></tr>)}</tbody></table></div>
          {selectedRivals.length === 0 && <small className="competitive-footnote">Closest-size rivals are a convenience shortlist; add/remove institutions above to define the comparison explicitly.</small>}
        </div>
        <div className="panel">
          <div className="panel-head"><h3>Executive opportunity & risk signals</h3><span className="panel-head-sub">Screening indicators for discussion, not supervisory determinations</span></div>
          <div className="competitive-signals">{signals.map((signal) => <div className={`competitive-signal ${signal.risk ? "signal-risk" : "signal-good"}`} key={signal.title}><span>{signal.risk ? <AlertTriangle size={15}/> : <CheckCircle2 size={15}/>}</span><div><strong>{signal.title}</strong><small>{signal.detail}</small></div></div>)}</div>
        </div>
      </div>

      <div className="panel competitive-limit"><Info size={16}/><div><strong>Data coverage and interpretation</strong><span>The local API starts with one quarter; Q1-to-Q2 growth and gainers/decliners require both source CSVs to be loaded from Connect Admin. Those CSVs contain loan-category balances, but the current CU model does not retain those fields, so the chart above compares balance-sheet ratios rather than product mix. Geography is headquarters-state only: the source has no city or branch addresses, so it cannot establish service-area overlap.</span></div></div>
    </div>
  );
}

function distScore(a, cu) { return Math.abs(a.assetsM - cu.assetsM); }

function MiniStat({ label, v, beat }) {
  return (
    <div className="mini-stat">
      <span className="mini-stat-label">{label}</span>
      <span className={`mini-stat-value ${beat ? "good" : ""}`}>{v}</span>
    </div>
  );
}

/* ============================== COMPARE INSTITUTIONS TAB ============================== */
const COMPARE_METRICS = ["assetsM", "members", "branches", "roa", "roe", "capRatio", "dr", "nco", "ltd", "lta", "py", "niM"];
const COMPARE_LABELS = {
  assetsM: "Total Assets", members: "Members", branches: "Branches", niM: "Net Income",
  roa: "ROA", roe: "ROE", capRatio: "Net Worth Ratio", dr: "Delinquency Rate",
  nco: "Net Charge-Off Rate", ltd: "Loan-to-Share", lta: "Loan-to-Asset", py: "Portfolio Yield",
};
const COMPARE_FORMAT = {
  assetsM: (v) => fmtM(v), members: (v) => fmtNum(v), branches: (v) => fmtNum(v), niM: (v) => fmtM(v),
  roa: (v) => fmtPct(v), roe: (v) => fmtPct(v), capRatio: (v) => fmtPct(v), dr: (v) => fmtPct(v),
  nco: (v) => fmtPct(v), ltd: (v) => fmtPct(v, 1), lta: (v) => fmtPct(v, 1), py: (v) => fmtPct(v),
};

function CompareTab({ primary }) {
  const [extras, setExtras] = useState([]);
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);

  // Keep extras clean if the primary CU changes (e.g. the user searches a new one up top).
  useEffect(() => {
    setExtras((ex) => ex.filter((c) => c.idx !== primary?.idx));
    // eslint-disable-next-line
  }, [primary?.charter, primary?.name, primary?.state, DATA_VERSION]);

  const list = primary ? [primary, ...extras].slice(0, 4) : extras.slice(0, 4);

  const results = useMemo(() => {
    if (q.trim().length < 2) return [];
    const needle = q.trim().toLowerCase();
    return ACTIVE_CUS
      .filter((c) => !list.some((l) => l.idx === c.idx) && (c.name.toLowerCase().includes(needle)))
      .sort((a, b) => b.assetsM - a.assetsM)
      .slice(0, 8);
  }, [q, list, DATA_VERSION]);

  function addCU(c) {
    if (list.length >= 4) return;
    setExtras((ex) => [...ex, c]);
    setQ("");
    setOpen(false);
  }
  function removeCU(idx) {
    setExtras((ex) => ex.filter((c) => c.idx !== idx));
  }

  return (
    <div className="tab-body">
      <div className="panel compare-add-panel">
        <div className="compare-add-head">
          <h3>Side-by-side comparison</h3>
          <span className="panel-head-sub">{list.length} of 4 institutions selected</span>
        </div>
        {list.length < 4 && (
          <div className="cu-search compare-search">
            <Search size={16} strokeWidth={2} className="cu-search-icon" />
            <input
              value={q}
              onChange={(e) => { setQ(e.target.value); setOpen(true); }}
              onFocus={() => setOpen(true)}
              placeholder="Add a credit union to compare…"
            />
            {open && results.length > 0 && (
              <div className="cu-search-results">
                {results.map((c) => (
                  <button key={c.idx} className="cu-search-row" onClick={() => addCU(c)}>
                    <span className="cu-search-name">{titleCase(c.name)}</span>
                    <span className="cu-search-meta">{c.state} · {fmtM(c.assetsM)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}
      </div>

      {list.length === 0 ? (
        <div className="panel"><EmptyState icon={GitCompare} title="No institutions selected" body="Search above to add up to four credit unions and compare them side by side." /></div>
      ) : (
        <>
          <div className="compare-cards" style={{ gridTemplateColumns: `repeat(${list.length}, 1fr)` }}>
            {list.map((c) => (
              <div className="compare-card" key={c.idx}>
                {c.idx !== primary?.idx && (
                  <button className="compare-remove" onClick={() => removeCU(c.idx)}><X size={13} /></button>
                )}
                {c.idx === primary?.idx && <span className="you-tag compare-you-tag">YOU</span>}
                <h4>{titleCase(c.name)}</h4>
                <p className="td-sub">{c.state}</p>
                <span className="compare-card-assets">{fmtM(c.assetsM)}</span>
                <span className="td-sub">{BUCKET_LABELS[c.bucket]}</span>
              </div>
            ))}
          </div>

          <div className="panel">
            <div className="panel-head"><h3>Metric comparison</h3></div>
            <div className="table-wrap">
              <table className="compare-table">
                <thead>
                  <tr>
                    <th>Metric</th>
                    {list.map((c) => <th key={c.idx}>{shortName(c.name)}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {COMPARE_METRICS.map((m) => {
                    const vals = list.map((c) => c[m]);
                    const best = LOWER_IS_BETTER.has(m) ? Math.min(...vals) : Math.max(...vals);
                    return (
                      <tr key={m}>
                        <td className="compare-metric-label">{COMPARE_LABELS[m]}</td>
                        {list.map((c) => (
                          <td key={c.idx} className={`mono ${c[m] === best && list.length > 1 ? "compare-best" : ""}`}>
                            {COMPARE_FORMAT[m](c[m])}
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div className="panel">
            <div className="panel-head"><h3>Assets at a glance</h3></div>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={list.map((c) => ({ name: shortName(c.name), Assets: Math.round(c.assetsM), you: c.idx === primary?.idx }))} margin={{ left: 8, right: 16 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.08)" vertical={false} />
                <XAxis dataKey="name" tick={{ fill: "#8FA0BD", fontSize: 11 }} />
                <YAxis tick={{ fill: "#8FA0BD", fontSize: 11 }} tickFormatter={(v) => `$${v}M`} />
                <Tooltip contentStyle={tooltipStyle} formatter={(v) => [`$${v}M`, "Assets"]} />
                <Bar dataKey="Assets" radius={[3, 3, 0, 0]}>
                  {list.map((c, i) => <Cell key={i} fill={c.idx === primary?.idx ? "#C9A227" : "#3A4457"} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </>
      )}
    </div>
  );
}


/* ============================== TREND COMPARISON TAB ============================== */
const TREND_METRICS = [
  { key: "assetsM", label: "Total Assets", kind: "money", decimals: 0 },
  { key: "members", label: "Members", kind: "number", decimals: 0 },
  { key: "loansM", label: "Total Loans", kind: "money", decimals: 0 },
  { key: "depositsM", label: "Total Deposits", kind: "money", decimals: 0 },
  ...Object.entries(METRIC_META).map(([key, meta]) => ({ key, label: meta.label, kind: "pct", decimals: meta.decimals })),
];
function formatTrendValue(value, metric) {
  if (value == null || Number.isNaN(Number(value))) return "—";
  if (metric.kind === "money") return `$${Number(value).toLocaleString("en-US", { maximumFractionDigits: metric.decimals })}M`;
  if (metric.kind === "number") return Number(value).toLocaleString("en-US", { maximumFractionDigits: metric.decimals });
  return `${Number(value).toFixed(metric.decimals)}%`;
}
function TrendComparisonTab({ primary, quarterRegistry, opts }) {
  const [metricKey, setMetricKey] = useState("assetsM");
  const [compareIdentity, setCompareIdentity] = useState("");
  const quarters = Object.keys(quarterRegistry);
  const metric = TREND_METRICS.find((m) => m.key === metricKey) || TREND_METRICS[0];
  const currentRows = quarterRegistry[quarters[quarters.length - 1]] || [];
  const compareCU = currentRows.find((c) => institutionIdentity(c) === compareIdentity);
  const availableCompare = currentRows.filter((c) => institutionIdentity(c) !== institutionIdentity(primary)).sort((a, b) => b.assetsM - a.assetsM);
  const series = useMemo(() => {
    if (!primary) return [];
    return quarters.map((quarter) => {
      const rows = quarterRegistry[quarter] || [];
      const findSame = (target) => findQuarterCU(rows, target);
      const you = findSame(primary);
      const comparison = compareCU ? findSame(compareCU) : null;
      const bucket = you?.bucket ?? primary.bucket;
      const lo = opts.wideTier ? Math.max(0, bucket - 1) : bucket;
      const hi = opts.wideTier ? Math.min(7, bucket + 1) : bucket;
      let peerRows = rows.filter((r) => r.bucket >= lo && r.bucket <= hi);
      if (opts.sameState) peerRows = peerRows.filter((r) => r.state === primary.state);
      const peerVals = peerRows.map((r) => Number(r[metricKey])).filter(Number.isFinite);
      return { quarter, You: you ? Number(you[metricKey]) : null, Comparison: comparison ? Number(comparison[metricKey]) : null, "Peer Median": peerVals.length ? stat(peerVals).median : null };
    });
  }, [primary, compareCU, metricKey, quarterRegistry, opts.sameState, opts.wideTier]);
  const validPoints = series.filter((d) => d.You != null);
  const first = validPoints[0]?.You;
  const last = validPoints[validPoints.length - 1]?.You;
  const change = first != null && last != null ? last - first : null;
  return <div className="tab-body">
    <div className="trend-hero"><div><span className="eyebrow-tag">Quarterly trend intelligence</span><h1>Trend comparison</h1><p>Track one metric over time and compare {shortName(primary.name)} with another credit union and the peer median.</p></div><div className="trend-controls"><label>Metric<select value={metricKey} onChange={(e) => setMetricKey(e.target.value)}>{TREND_METRICS.map((m) => <option key={m.key} value={m.key}>{m.label}</option>)}</select></label><label>Compare with<select value={compareIdentity} onChange={(e) => setCompareIdentity(e.target.value)}><option value="">Peer median only</option>{availableCompare.map((c) => <option key={institutionIdentity(c)} value={institutionIdentity(c)}>{titleCase(c.name)} · {c.state}</option>)}</select></label></div></div>
    <div className="kpi-grid trend-kpis"><KPI label="Latest value" value={formatTrendValue(last, metric)} sub={quarters[quarters.length - 1] || "—"}/><KPI label="First available" value={formatTrendValue(first, metric)} sub={validPoints[0]?.quarter || "—"}/><KPI label="Change over period" value={change == null ? "—" : `${change >= 0 ? "+" : ""}${formatTrendValue(change, metric)}`} sub={validPoints.length > 1 ? `${validPoints.length} reported quarters` : "Add more quarters to unlock trend"}/><KPI label="Data coverage" value={`${validPoints.length}/${quarters.length}`} sub="quarters with matching CU data"/></div>
    <div className="panel"><div className="panel-head"><h3>{metric.label} over time</h3><span className="panel-head-sub">{opts.sameState ? "Same-state peers" : "Peer tier"}{opts.wideTier ? " · adjacent tiers included" : ""}</span></div>{quarters.length < 2 ? <div className="trend-empty"><TrendingUp size={22}/><strong>Trend view is ready, but only one quarter is loaded.</strong><span>Import another quarter from Connect Admin to populate the time series. No historical values are being fabricated.</span></div> : <ResponsiveContainer width="100%" height={360}><LineChart data={series} margin={{ left: 8, right: 20, top: 12, bottom: 12 }}><CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.08)" vertical={false}/><XAxis dataKey="quarter" tick={{ fill: "#8FA0BD", fontSize: 11 }}/><YAxis tick={{ fill: "#8FA0BD", fontSize: 11 }} tickFormatter={(v) => metric.kind === "money" ? `$${v}M` : metric.kind === "number" ? Number(v).toLocaleString() : `${v}%`}/><Tooltip contentStyle={tooltipStyle} formatter={(v, name) => [formatTrendValue(v, metric), name]}/><Legend wrapperStyle={{ color: "#8FA0BD", fontSize: 11 }}/><Line type="monotone" dataKey="You" name={shortName(primary.name)} stroke="#C9A227" strokeWidth={3} dot={{ r: 3 }} connectNulls={false}/>{compareCU && <Line type="monotone" dataKey="Comparison" name={shortName(compareCU.name)} stroke="#43B8A5" strokeWidth={2} dot={{ r: 3 }} connectNulls={false}/>}<Line type="monotone" dataKey="Peer Median" stroke="#8FA0BD" strokeWidth={2} strokeDasharray="6 5" dot={{ r: 2 }} connectNulls={false}/></LineChart></ResponsiveContainer>}</div>
    <div className="panel"><div className="panel-head"><h3>Quarter-by-quarter detail</h3><span className="panel-head-sub">PPT-inspired comparison table</span></div><div className="table-wrap"><table><thead><tr><th>Quarter</th><th>{shortName(primary.name)}</th>{compareCU && <th>{shortName(compareCU.name)}</th>}<th>Peer median</th><th>Vs peer</th></tr></thead><tbody>{series.map((r) => { const diff = r.You != null && r["Peer Median"] != null ? r.You - r["Peer Median"] : null; return <tr key={r.quarter}><td>{r.quarter}</td><td className="mono">{formatTrendValue(r.You, metric)}</td>{compareCU && <td className="mono">{formatTrendValue(r.Comparison, metric)}</td>}<td className="mono">{formatTrendValue(r["Peer Median"], metric)}</td><td className="mono">{diff == null ? "—" : `${diff >= 0 ? "+" : ""}${formatTrendValue(diff, metric)}`}</td></tr>; })}</tbody></table></div></div>
  </div>;
}

/* ============================== US STATE MAP ============================== */
// Tile-grid layout (one equal tile per state, roughly following geography) so small northeastern
// states are as easy to hover and click as Texas. Format: [state, column, row].
const US_STATE_TILES = [
  ["AK",0,0],["ME",10,0],
  ["VT",9,1],["NH",10,1],
  ["WA",0,2],["ID",1,2],["MT",2,2],["ND",3,2],["MN",4,2],["IL",5,2],["WI",6,2],["MI",7,2],["NY",8,2],["RI",9,2],["MA",10,2],
  ["OR",0,3],["NV",1,3],["WY",2,3],["SD",3,3],["IA",4,3],["IN",5,3],["OH",6,3],["PA",7,3],["NJ",8,3],["CT",9,3],
  ["CA",0,4],["UT",1,4],["CO",2,4],["NE",3,4],["MO",4,4],["KY",5,4],["WV",6,4],["VA",7,4],["MD",8,4],["DE",9,4],
  ["AZ",1,5],["NM",2,5],["KS",3,5],["AR",4,5],["TN",5,5],["NC",6,5],["SC",7,5],["DC",8,5],
  ["OK",3,6],["LA",4,6],["MS",5,6],["AL",6,6],["GA",7,6],
  ["HI",0,7],["GU",1,7],["TX",3,7],["FL",8,7],["VI",9,7],["PR",10,7],
];
const TILE_W = 54, TILE_H = 48, TILE_GX = 7, TILE_GY = 6;
const TILE_PITCH_X = TILE_W + TILE_GX, TILE_PITCH_Y = TILE_H + TILE_GY;
const MAP_VIEW_W = 11 * TILE_PITCH_X - TILE_GX, MAP_VIEW_H = 8 * TILE_PITCH_Y - TILE_GY;

const compactCount = (v) => (v >= 1e6 ? `${(v / 1e6).toFixed(1)}M` : v >= 1e4 ? `${Math.round(v / 1e3)}K` : v >= 1e3 ? `${(v / 1e3).toFixed(1)}K` : String(Math.round(v)));
const compactAssets = (v) => (v >= 1e6 ? `$${(v / 1e6).toFixed(1)}T` : v >= 1e4 ? `$${Math.round(v / 1e3)}B` : v >= 1e3 ? `$${(v / 1e3).toFixed(1)}B` : `$${Math.round(v)}M`);
const MAP_METRICS = [
  { key: "count", label: "Credit unions", noun: "credit unions", short: compactCount, full: (v) => Math.round(v).toLocaleString() },
  { key: "members", label: "Members", noun: "members", short: compactCount, full: (v) => Math.round(v).toLocaleString() },
  { key: "assetsM", label: "Assets", noun: "assets", short: compactAssets, full: (v) => fmtM(v) },
  { key: "branches", label: "Branches", noun: "branches", short: compactCount, full: (v) => Math.round(v).toLocaleString() },
];

function buildStateStats(cus) {
  const by = {};
  (cus || []).forEach((c) => {
    const s = c.state;
    if (!s) return;
    const e = by[s] || (by[s] = { state: s, count: 0, members: 0, assetsM: 0, branches: 0, cus: [] });
    e.count += 1;
    e.members += Number(c.members) || 0;
    e.assetsM += Number(c.assetsM) || 0;
    e.branches += Number(c.branches) || 0;
    e.cus.push(c);
  });
  Object.values(by).forEach((e) => e.cus.sort((a, b) => (Number(b.assetsM) || 0) - (Number(a.assetsM) || 0)));
  return by;
}

function StateMap({ stats, selectedCU, onSelectCU, quarter }) {
  const [metricKey, setMetricKey] = useState("count");
  const [hover, setHover] = useState(null);
  const [pinned, setPinned] = useState(null);
  const homeState = selectedCU?.state || null;
  useEffect(() => { setPinned(null); }, [homeState]);

  const metric = MAP_METRICS.find((m) => m.key === metricKey);
  const valueOf = (s) => stats[s]?.[metricKey] || 0;
  const max = Math.max(1, ...US_STATE_TILES.map(([s]) => valueOf(s)));
  const total = US_STATE_TILES.reduce((sum, [s]) => sum + valueOf(s), 0);

  const active = hover || pinned || homeState;
  const info = active ? stats[active] : null;
  const tag = hover ? "Previewing" : pinned ? "Pinned" : "Your credit union’s state";
  const rank = info ? US_STATE_TILES.filter(([s]) => valueOf(s) > valueOf(active)).length + 1 : null;
  const statesWithData = US_STATE_TILES.filter(([s]) => valueOf(s) > 0).length;
  const share = info && total ? (valueOf(active) / total) * 100 : 0;

  function toggle(s) { setPinned((p) => (p === s ? null : s)); }

  return (
    <div className="map-shell">
      <div className="map-toolbar">
        <div className="map-metrics" role="tablist" aria-label="Map metric">
          {MAP_METRICS.map((m) => (
            <button key={m.key} role="tab" aria-selected={metricKey === m.key} className={`chip ${metricKey === m.key ? "chip-on" : ""}`} onClick={() => setMetricKey(m.key)}>{m.label}</button>
          ))}
        </div>
        {(pinned || hover) && <button className="map-reset" onClick={() => { setPinned(null); setHover(null); }}>Reset</button>}
      </div>

      <div className="map-layout">
        <div className="map-stage">
          <svg className="state-map" viewBox={`0 0 ${MAP_VIEW_W} ${MAP_VIEW_H}`} role="group" aria-label={`U.S. map of ${metric.noun} by headquarters state`}>
            {US_STATE_TILES.map(([s, col, row]) => {
              const v = valueOf(s);
              const strength = v ? Math.sqrt(v / max) : 0;
              const opacity = v ? 0.16 + 0.76 * strength : 0.05;
              const cls = ["map-tile", v ? "" : "empty", strength > 0.62 ? "hot" : "", active === s ? "is-active" : "", pinned === s ? "is-pinned" : "", homeState === s ? "is-home" : ""].filter(Boolean).join(" ");
              return (
                <g key={s} className={cls} transform={`translate(${col * TILE_PITCH_X},${row * TILE_PITCH_Y})`} tabIndex={0} role="button"
                   aria-label={`${STATE_NAMES[s] || s}: ${v ? metric.full(v) : "none"} ${metric.noun}`} aria-pressed={pinned === s}
                   onMouseEnter={() => setHover(s)} onMouseLeave={() => setHover((h) => (h === s ? null : h))}
                   onFocus={() => setHover(s)} onBlur={() => setHover((h) => (h === s ? null : h))}
                   onClick={() => toggle(s)}
                   onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(s); } }}>
                  <rect width={TILE_W} height={TILE_H} rx="8" fill="var(--brass)" fillOpacity={opacity} />
                  <text className="tile-code" x={TILE_W / 2} y="21" textAnchor="middle">{s}</text>
                  <text className="tile-val" x={TILE_W / 2} y="37" textAnchor="middle">{v ? metric.short(v) : "–"}</text>
                  {homeState === s && <text className="tile-star" x={TILE_W - 8} y="13" textAnchor="middle">★</text>}
                </g>
              );
            })}
          </svg>
          <div className="map-legend">
            <span>0</span><i className="map-legend-bar" /><span>{metric.short(max)}</span>
            <small>{metric.label} per state · {statesWithData} states &amp; territories · ★ = selected credit union</small>
          </div>
        </div>

        <aside className="map-detail" aria-live="polite">
          {active ? (
            <>
              <div className="map-detail-head">
                <span className="map-detail-tag">{tag}</span>
                <h4>{STATE_NAMES[active] || active}</h4>
              </div>
              {info ? (
                <>
                  <div className="map-detail-grid">
                    <div className={metricKey === "count" ? "is-metric" : ""}><span>Credit unions</span><strong>{info.count.toLocaleString()}</strong></div>
                    <div className={metricKey === "members" ? "is-metric" : ""}><span>Members</span><strong>{fmtNum(info.members)}</strong></div>
                    <div className={metricKey === "assetsM" ? "is-metric" : ""}><span>Assets</span><strong>{fmtM(info.assetsM)}</strong></div>
                    <div className={metricKey === "branches" ? "is-metric" : ""}><span>Branches</span><strong>{fmtNum(info.branches)}</strong></div>
                  </div>
                  <p className="map-detail-rank">Ranks <b>#{rank}</b> of {statesWithData} by {metric.noun} · {share.toFixed(share < 10 ? 1 : 0)}% of the national total</p>
                  <h5>Largest credit unions</h5>
                  <div className="map-detail-list">
                    {info.cus.slice(0, 5).map((c) => (
                      <button key={c.idx} className={selectedCU && c.idx === selectedCU.idx ? "is-you" : ""} onClick={() => onSelectCU && onSelectCU(c)} title="Open this credit union">
                        <span>{shortName(c.name)}</span><em>{fmtM(c.assetsM)}</em>
                      </button>
                    ))}
                  </div>
                </>
              ) : (
                <p className="map-detail-empty">No credit unions are headquartered in {STATE_NAMES[active] || active} in {quarter}.</p>
              )}
            </>
          ) : (
            <p className="map-detail-empty">Hover or tap a state to see its credit unions.</p>
          )}
          <div className="map-detail-hint">Hover to preview · click a state to pin it · click a credit union to open it</div>
        </aside>
      </div>
    </div>
  );
}

/* ============================== GAP ANALYSIS ============================== */
const GAP_METRICS = [
  { key:"roa", label:"Return on Assets", unit:"%", higher:true, meaning:"How effectively the CU turns its asset base into net income. Higher than peers generally indicates stronger profitability." },
  { key:"roe", label:"Return on Equity", unit:"%", higher:true, meaning:"How effectively the CU generates earnings from its capital base. Higher than peers generally indicates stronger returns on equity." },
  { key:"capRatio", label:"Net Worth Ratio", unit:"%", higher:true, meaning:"The CU's capital cushion relative to assets. Higher generally means more balance-sheet capacity to absorb losses." },
  { key:"lta", label:"Loan-to-Asset Ratio", unit:"%", higher:true, meaning:"How much of the asset base is deployed into loans. Higher can indicate stronger loan deployment, while an unusually high level can also mean less liquidity flexibility." },
  { key:"ltd", label:"Loan-to-Share Ratio", unit:"%", higher:false, meaning:"Loans relative to member shares/deposits. A lower result is treated here as favorable under the existing model because it indicates less funding pressure." },
  { key:"dr", label:"Delinquency Rate", unit:"%", higher:false, meaning:"The share of loans that are delinquent. Lower is better because it indicates stronger credit performance." },
  { key:"nco", label:"Net Charge-Off Rate", unit:"%", higher:false, meaning:"Net loan losses after recoveries. Lower is better because it indicates fewer credit losses." },
  { key:"py", label:"Portfolio Yield", unit:"%", higher:true, meaning:"Income/asset yield proxy derived from gross income in the supplied quarter. Higher generally indicates stronger income generation from assets." },
];

function GapAnalysisTab({ cu, peers }) {
  const rows = useMemo(() => GAP_METRICS.map((m) => {
    const values = peers.map((p) => Number(p[m.key])).filter(Number.isFinite);
    const median = values.length ? stat(values).median : null;
    const actual = Number(cu[m.key]);
    const gap = median == null || !Number.isFinite(actual) ? null : actual - median;
    const favorable = gap == null ? null : (m.higher ? gap >= 0 : gap <= 0);
    return { ...m, actual, median, gap, favorable };
  }), [cu, peers]);
  const opportunities = rows.filter((r) => r.gap != null && !r.favorable).sort((a,b) => Math.abs(b.gap)-Math.abs(a.gap));
  const strengths = rows.filter((r) => r.gap != null && r.favorable).sort((a,b) => Math.abs(b.gap)-Math.abs(a.gap));
  const fmt = (v, unit) => v == null || !Number.isFinite(v) ? "—" : `${v.toFixed(2)}${unit}`;
  return <div className="tab-body">
    <div className="trend-hero"><div><span className="eyebrow-tag"><Target size={12}/> Executive performance diagnostic</span><h1>Gap analysis</h1><p>See where the selected credit union differs from its active peer median, understand what each difference means, and focus attention on the largest unfavorable gaps.</p></div><div className="gap-summary"><span>Peer group</span><strong>{peers.length.toLocaleString()} institutions</strong></div></div>
    <div className="panel gap-explainer"><div className="panel-head"><div><h3>How to read this page</h3><span className="panel-head-sub">The gap is calculated as CU value minus peer median.</span></div><span className="gap-badge good">Favorable = better under the existing metric rule</span></div><div className="gap-explainer-grid"><div><strong>Positive gap</strong><span>The CU is above the peer median. Whether that is favorable depends on the metric.</span></div><div><strong>Negative gap</strong><span>The CU is below the peer median. Whether that is favorable depends on the metric.</span></div><div><strong>Opportunity</strong><span>The current position is less favorable than the peer median under the app's existing benchmark rule.</span></div><div><strong>Strength</strong><span>The current position is at or better than the peer median under the app's existing benchmark rule.</span></div></div></div>
    <div className="kpi-grid"><KPI label="Gaps to address" value={opportunities.length.toString()} sub="metrics flagged as opportunities"/><KPI label="Relative strengths" value={strengths.length.toString()} sub="metrics at / better than peer median"/><KPI label="Largest opportunity" value={opportunities[0] ? `${Math.abs(opportunities[0].gap).toFixed(2)} pp` : "—"} sub={opportunities[0]?.label || "No unfavorable gap"}/><KPI label="Benchmark" value="Peer median" sub={BUCKET_LABELS[cu.bucket]}/></div>
    <div className="panel"><div className="panel-head"><div><h3>Executive scorecard</h3><span className="panel-head-sub">What the gap means for the CU</span></div><span className="panel-head-sub">Benchmark: {BUCKET_LABELS[cu.bucket]} peer median</span></div><div className="table-wrap"><table><thead><tr><th>Metric</th><th>{shortName(cu.name)}</th><th>Peer median</th><th>Gap</th><th>Assessment</th><th>What it means</th></tr></thead><tbody>{rows.map((r) => <tr key={r.key}><td><strong>{r.label}</strong></td><td className="mono">{fmt(r.actual,r.unit)}</td><td className="mono">{fmt(r.median,r.unit)}</td><td className={`mono ${r.favorable ? "gap-good" : "gap-bad"}`}>{r.gap == null ? "—" : `${r.gap >= 0 ? "+" : ""}${r.gap.toFixed(2)}${r.unit}`}</td><td><span className={`gap-badge ${r.favorable ? "good" : "bad"}`}>{r.favorable ? "Strength" : "Opportunity"}</span></td><td className="gap-meaning">{r.meaning}</td></tr>)}</tbody></table></div></div>
    <div className="grid-2"><div className="panel"><div className="panel-head"><h3>Priority opportunities</h3><span className="panel-head-sub">Where management may want to investigate first</span></div>{opportunities.length ? <div className="gap-list">{opportunities.slice(0,5).map((r) => <div className="gap-item" key={r.key}><div><strong>{r.label}</strong><span>{fmt(r.actual,r.unit)} vs {fmt(r.median,r.unit)} peer median · {r.meaning}</span></div><b>{Math.abs(r.gap).toFixed(2)} pp</b></div>)}</div> : <EmptyState icon={Check} title="No unfavorable gaps" body="The selected metrics are at or better than the peer median."/>}</div><div className="panel"><div className="panel-head"><h3>Relative strengths</h3><span className="panel-head-sub">Where the CU is performing favorably vs peers</span></div><div className="gap-list">{strengths.slice(0,5).map((r) => <div className="gap-item" key={r.key}><div><strong>{r.label}</strong><span>{fmt(r.actual,r.unit)} vs {fmt(r.median,r.unit)} peer median · {r.meaning}</span></div><b className="gap-good">{Math.abs(r.gap).toFixed(2)} pp</b></div>)}</div></div></div>
  </div>;
}

/* ============================== DEMOGRAPHIC VIEWS TAB ============================== */
function DemographicViewsTab({ cu, peers, onSelectCU }) {
  // State-level roll-up of the active quarter (headquarters state; branch-level locations are not in the data).
  const stateStats = useMemo(() => buildStateStats(ACTIVE_CUS), [DATA_VERSION]);

  const peerMedianMembers = stat(peers.map((p) => p.members)).median;
  const peerMedianBranches = stat(peers.map((p) => p.branches)).median;
  const peerMedianEmployees = stat(peers.map((p) => p.employees)).median;
  const branchDensity = cu.branches ? cu.members / cu.branches : null;
  const employeeDensity = cu.employees ? cu.members / cu.employees : null;
  const peerBranchDensity = peerMedianBranches ? peerMedianMembers / peerMedianBranches : null;
  const peerEmployeeDensity = peerMedianEmployees ? peerMedianMembers / peerMedianEmployees : null;

  // State market position uses the supplied institution-level dataset only.
  const stateCUs = ACTIVE_CUS.filter((r) => r.state === cu.state);
  const stateAssetRank = stateCUs.length
    ? [...stateCUs].sort((a, b) => (Number(b.assetsM) || 0) - (Number(a.assetsM) || 0)).findIndex((r) => r.idx === cu.idx) + 1
    : null;
  const stateMemberRank = stateCUs.length
    ? [...stateCUs].sort((a, b) => (Number(b.members) || 0) - (Number(a.members) || 0)).findIndex((r) => r.idx === cu.idx) + 1
    : null;
  const peerMemberPct = percentileRank(peers.map((p) => p.members), cu.members);
  const memberVsPeer = peerMedianMembers ? ((cu.members / peerMedianMembers) - 1) * 100 : null;
  const branchVsPeer = peerMedianBranches ? ((cu.branches / peerMedianBranches) - 1) * 100 : null;

  return <div className="tab-body">
    <div className="trend-hero"><div><span className="eyebrow-tag">Market & membership profile</span><h1>Demographic views</h1><p>Understand where the selected credit union operates, how large its membership base is, and how it compares with its market and peers.</p></div></div>

    <div className="kpi-grid">
      <KPI label="Market location" value={cu.state} sub={cu.region ? `Region ${cu.region}` : "Headquarters state"}/>
      <KPI label="Members" value={fmtNum(cu.members)} sub={`Peer median ${fmtNum(peerMedianMembers)}`}/>
      <KPI label="Branches" value={fmtNum(cu.branches)} sub={`Peer median ${fmtNum(peerMedianBranches)}`}/>
      <KPI label="Member percentile" value={peerMemberPct == null ? "—" : `${Math.round(peerMemberPct)}th`} sub="Within selected peer set"/>
    </div>

    <div className="grid-2">
      <div className="panel"><div className="panel-head"><h3>Market position</h3><span className="panel-head-sub">Selected CU within {cu.state}</span></div><div className="region-cards">
        <div className="region-card"><strong>Asset rank</strong><div><span>{stateAssetRank ? `#${stateAssetRank}` : "—"}</span><small>among {stateCUs.length} state CUs</small></div></div>
        <div className="region-card"><strong>Member rank</strong><div><span>{stateMemberRank ? `#${stateMemberRank}` : "—"}</span><small>among {stateCUs.length} state CUs</small></div></div>
        <div className="region-card"><strong>Member scale</strong><div><span>{memberVsPeer == null ? "—" : `${memberVsPeer >= 0 ? "+" : ""}${memberVsPeer.toFixed(0)}%`}</span><small>vs peer median</small></div></div>
        <div className="region-card"><strong>Branch scale</strong><div><span>{branchVsPeer == null ? "—" : `${branchVsPeer >= 0 ? "+" : ""}${branchVsPeer.toFixed(0)}%`}</span><small>vs peer median</small></div></div>
      </div></div>

      <div className="panel"><div className="panel-head"><h3>Membership reach</h3><span className="panel-head-sub">Service intensity vs peer median</span></div><div className="compare-list">
        <div className="compare-row"><div className="compare-label">Members / branch</div><div className="compare-values"><span className="compare-you">{branchDensity == null ? "—" : Math.round(branchDensity).toLocaleString()}</span><span className="compare-vs">vs {peerBranchDensity == null ? "—" : Math.round(peerBranchDensity).toLocaleString()} median</span></div></div>
        <div className="compare-row"><div className="compare-label">Members / employee</div><div className="compare-values"><span className="compare-you">{employeeDensity == null ? "—" : Math.round(employeeDensity).toLocaleString()}</span><span className="compare-vs">vs {peerEmployeeDensity == null ? "—" : Math.round(peerEmployeeDensity).toLocaleString()} median</span></div></div>
        <div className="compare-row"><div className="compare-label">Members</div><div className="compare-values"><span className="compare-you">{fmtNum(cu.members)}</span><span className="compare-vs">peer median {fmtNum(peerMedianMembers)}</span></div></div>
      </div></div>
    </div>

    <div className="panel">
      <div className="panel-head"><h3>Market map</h3><span className="panel-head-sub">{ACTIVE_QUARTER} · credit unions by headquarters state · {shortName(cu.name)} marked ★</span></div>
      <StateMap stats={stateStats} selectedCU={cu} onSelectCU={onSelectCU} quarter={ACTIVE_QUARTER} />
    </div>

    <div className="panel"><div className="panel-head"><h3>Institution profile</h3><span className="panel-head-sub">Selected CU at a glance</span></div><div className="demographic-profile wide">
        <div><span>Headquarters state</span><strong>{cu.state}</strong></div>
        <div><span>Region</span><strong>{cu.region || "Not provided"}</strong></div>
        <div><span>Members</span><strong>{fmtNum(cu.members)}</strong></div>
        <div><span>Branches</span><strong>{fmtNum(cu.branches)}</strong></div>
        <div><span>Employees</span><strong>{cu.employees ? fmtNum(cu.employees) : "Not provided"}</strong></div>
        <div><span>Assets</span><strong>{fmtM(cu.assetsM)}</strong></div>
    </div></div>

    <div className="panel demographic-note"><Info size={16}/><div><strong>Data boundary</strong><span>The current dataset supports institution-level geography and operating-footprint fields: state, members, branches, assets, loans, shares/deposits, CEO and website. Headquarters city and employee count are not present in the supplied dataset, and member-level age, income, ethnicity, household or field-of-membership demographics are not present, so this view does not invent them.</span></div></div>
  </div>;
}

/* ============================== RISK ALERTS TAB ============================== */
function RiskTab({ cu, peers, flags, opts }) {
  const sorted = [...flags].sort((a, b) => SEV_ORDER[a.sev] - SEV_ORDER[b.sev]);
  const peerFlagRate = useMemo(() => {
    if (!peers.length) return 0;
    const flagged = peers.filter((p) => computeRiskFlags(p, peers).length > 0).length;
    return (flagged / peers.length) * 100;
  }, [peers]);

  return (
    <div className="tab-body">
      <PeerGroupBanner cu={cu} opts={opts} peerCount={peers.length} />
      <div className="kpi-grid">
        <KPI label="Active risk flags" value={flags.length} sub={flags.length ? "requires board attention" : "no flags raised"} warn={flags.length > 0} />
        <KPI label="Highest severity" value={sorted[0] ? SEV_STYLE[sorted[0].sev].label : "None"} sub={sorted[0]?.title || "Balance sheet within normal range"} warn={!!sorted[0]} />
        <KPI label="Net worth ratio" value={fmtPct(cu.capRatio)} sub="NCUA well-capitalized floor: 7.00%" warn={cu.capRatio < 7} />
        <KPI label="Peers with active flags" value={fmtPct(peerFlagRate, 0)} sub="share of peer group flagged" />
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Flags for {shortName(cu.name)}</h3><span className="panel-head-sub">rule-based, derived from {ACTIVE_QUARTER} call report data</span></div>
        {sorted.length === 0 ? (
          <EmptyState icon={Shield} title="No risk flags raised" body="Capital, earnings, delinquency and liquidity ratios are all within normal thresholds for this peer group." />
        ) : (
          <div className="flag-list">
            {sorted.map((f, i) => (
              <div className="flag-row" key={i} style={{ borderLeftColor: SEV_STYLE[f.sev].color }}>
                <div className="flag-row-top">
                  <span className="flag-sev" style={{ color: SEV_STYLE[f.sev].color, background: SEV_STYLE[f.sev].bg }}>{SEV_STYLE[f.sev].label}</span>
                  <span className="flag-title">{f.title}</span>
                </div>
                <p className="flag-detail">{f.detail}</p>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Methodology</h3></div>
        <p className="method-note">
          Flags are computed directly from this quarter's call-report ratios against fixed, disclosed thresholds
          (for example NCUA's published net-worth tiers) and against the peer group's own distribution
          ({peerGroupText(cu, opts, peers.length)}) for delinquency and charge-off rates. This is an indicative
          screen, not an official PCA determination — treat it as a starting point for board discussion, not a
          substitute for supervisory review.
        </p>
      </div>
    </div>
  );
}

/* ============================== BOARD REPORT TAB ============================== */
function BoardReportTab({ cu, peers, opts, flags }) {
  const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const metrics = Object.keys(METRIC_META);

  return (
    <div className="tab-body">
      <div className="board-toolbar no-print">
        <p>A print-ready summary for board packets and committee meetings.</p>
        <button className="btn-primary" onClick={() => window.print()}><Printer size={15} /> Print / Save as PDF</button>
      </div>

      <div className="board-sheet">
        <div className="board-sheet-head">
          <div>
            <span className="board-kicker">Board Performance Summary — {ACTIVE_QUARTER} ACUI Call Report</span>
            <h2>{titleCase(cu.name)}</h2>
            <p>{cu.state} &nbsp;·&nbsp; Prepared {today}</p>
          </div>
          <Landmark size={30} strokeWidth={1.4} />
        </div>

        <div className="board-grid">
          <BoardStat label="Total Assets" value={fmtM(cu.assetsM)} />
          <BoardStat label="Members" value={fmtNum(cu.members)} />
          <BoardStat label="Net Income" value={fmtM(cu.niM)} />
          <BoardStat label="Branches" value={fmtNum(cu.branches)} />
        </div>

        <h4 className="board-section-title">Key ratios vs. peer group ({BUCKET_LABELS[cu.bucket]}{opts.sameState ? `, ${cu.state}` : ""}, {peers.length.toLocaleString()} institutions)</h4>
        <table className="board-table">
          <thead><tr><th>Metric</th><th>{shortName(cu.name)}</th><th>Peer median</th><th>Peer 25th–75th</th></tr></thead>
          <tbody>
            {metrics.map((m) => {
              const meta = METRIC_META[m];
              const s = stat(peers.map((p) => p[m]));
              return (
                <tr key={m}>
                  <td>{meta.label}</td>
                  <td className="mono">{fmtPct(cu[m], meta.decimals)}</td>
                  <td className="mono">{fmtPct(s.median, meta.decimals)}</td>
                  <td className="mono">{fmtPct(s.p25, meta.decimals)} – {fmtPct(s.p75, meta.decimals)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>

        <h4 className="board-section-title">Risk flags</h4>
        {flags.length === 0 ? (
          <p className="board-note">No risk flags were raised this quarter across capital, earnings, delinquency, charge-off, or liquidity ratios.</p>
        ) : (
          <ul className="board-flag-list">
            {flags.map((f, i) => <li key={i}><strong>{SEV_STYLE[f.sev].label}:</strong> {f.title} — {f.detail}</li>)}
          </ul>
        )}

        <p className="board-footer">Source: ACUI data, {ACTIVE_QUARTER} cycle. Figures are as supplied and unaudited. This summary is generated for internal discussion and does not constitute a regulatory filing.</p>
      </div>
    </div>
  );
}

function BoardStat({ label, value }) {
  return (
    <div className="board-stat">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

/* ============================== AI RECOMMENDATIONS TAB ============================== */
function AITab({ cu, peers, opts, flags }) {
  const [status, setStatus] = useState("idle"); // idle | loading | done | error
  const [recs, setRecs] = useState(null);
  const [errMsg, setErrMsg] = useState("");

  async function generate() {
    setStatus("loading");
    setErrMsg("");
    try {
      const metrics = Object.keys(METRIC_META).map((m) => {
        const s = stat(peers.map((p) => p[m]));
        return `${METRIC_META[m].label}: ${cu[m].toFixed(2)}% (peer median ${s.median.toFixed(2)}%, peer 25th-75th ${s.p25.toFixed(2)}%-${s.p75.toFixed(2)}%)`;
      }).join("\n");
      const flagText = flags.length ? flags.map((f) => `- [${SEV_STYLE[f.sev].label}] ${f.title}: ${f.detail}`).join("\n") : "None raised.";

      const prompt = `You are a credit union financial analyst preparing recommendations for a board of directors.

CREDIT UNION: ${cu.name}, ${cu.state}
PEER GROUP: ${BUCKET_LABELS[cu.bucket]} in assets${opts.sameState ? `, ${cu.state} only` : ", nationwide"} (${peers.length} institutions)
TOTAL ASSETS: $${cu.assetsM.toFixed(1)}M | MEMBERS: ${Math.round(cu.members)} | BRANCHES: ${cu.branches}

KEY RATIOS VS PEERS:
${metrics}

ACTIVE RISK FLAGS:
${flagText}

Based strictly on the data above, produce 4-6 concrete, board-ready recommendations. Each should tie directly to a ratio or flag given above — do not invent facts not implied by the data. Respond with ONLY a JSON array (no markdown, no prose, no code fences), where each element is:
{"priority":"high"|"medium"|"low","title":"short action title","rationale":"1-2 sentences tying it to the specific data point(s) above","action":"one concrete next step"}`;

      const response = await fetch(`${import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000"}/api/llm/recommendations`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || "Recommendation service unavailable");
      const text = (data.content || []).map((b) => b.text || "").join("\n");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      setRecs(parsed);
      setStatus("done");
    } catch (e) {
      setErrMsg("Couldn't generate recommendations right now. Please try again.");
      setStatus("error");
    }
  }

  return (
    <div className="tab-body">
      <div className="panel ai-intro">
        <div className="ai-intro-icon"><Sparkles size={20} strokeWidth={1.6} /></div>
        <div>
          <h3>AI-generated board recommendations</h3>
          <p>Claude reads {shortName(cu.name)}'s ratios against its peer group and drafts concrete, data-grounded recommendations. Review before presenting — this is a first draft, not final guidance.</p>
        </div>
        <button className="btn-primary" onClick={generate} disabled={status === "loading"}>
          {status === "loading" ? <><Loader2 size={15} className="spin" /> Generating…</> : <><Sparkles size={15} /> {recs ? "Regenerate" : "Generate recommendations"}</>}
        </button>
      </div>

      {status === "error" && <div className="panel"><EmptyState icon={AlertTriangle} title="Generation failed" body={errMsg} /></div>}

      {status === "idle" && !recs && (
        <div className="panel"><EmptyState icon={Sparkles} title="No recommendations yet" body="Click Generate recommendations to have Claude analyze this credit union's ratios and draft board-ready guidance." /></div>
      )}

      {recs && (
        <div className="rec-grid">
          {recs.map((r, i) => (
            <div className="rec-card" key={i}>
              <span className="rec-priority" style={{ color: SEV_STYLE[r.priority === "high" ? "high" : r.priority === "medium" ? "medium" : "low"].color, background: SEV_STYLE[r.priority === "high" ? "high" : r.priority === "medium" ? "medium" : "low"].bg }}>
                {r.priority} priority
              </span>
              <h4>{r.title}</h4>
              <p className="rec-rationale">{r.rationale}</p>
              <div className="rec-action"><ChevronRight size={14} /> {r.action}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
/* ============================== CHAT ASSISTANT ============================== */
function buildChatSystemPrompt(cu, peers, opts, flags) {
  const base = `You are the in-app assistant for CU360 Intelligence, a credit union analytics platform built on ${ACTIVE_QUARTER} NCUA call report data covering ${ACTIVE_CUS.length.toLocaleString()} credit unions nationwide. You help the user (a credit union analyst or board member) understand the data, navigate the platform's tabs (Overview, Peer Benchmarking, Competitive Intel, Compare Institutions, Risk Alerts, Board Report, AI Recommendations), and interpret NCUA call-report metrics like net worth ratio, ROA, delinquency, and loan-to-share. Keep answers concise (a few sentences unless asked for depth), grounded only in the data given below, and never invent figures.`;
  if (!cu) {
    return `${base}\n\nNo credit union is currently selected — the user is on the home/search screen. If they ask about a specific ratio, answer generally; encourage them to search for their institution to unlock personalized analysis.`;
  }
  const metrics = Object.keys(METRIC_META).map((m) => {
    const s = stat(peers.map((p) => p[m]));
    return `${METRIC_META[m].label}: ${cu[m].toFixed(2)}% (peer median ${s.median.toFixed(2)}%)`;
  }).join("; ");
  const flagText = flags.length ? flags.map((f) => `${f.title} (${f.sev})`).join("; ") : "none";
  return `${base}\n\nCURRENTLY SELECTED CREDIT UNION: ${cu.name}, ${cu.state}\nAssets: $${cu.assetsM.toFixed(1)}M | Members: ${Math.round(cu.members)} | Branches: ${cu.branches}\nPEER GROUP: ${peerGroupText(cu, opts, peers.length)}\nKEY RATIOS: ${metrics}\nACTIVE RISK FLAGS: ${flagText}`;
}

const CHAT_SUGGESTIONS_GENERIC = [
  "What does net worth ratio mean?",
  "How is my peer group defined?",
  "What can this platform do?",
];
const CHAT_SUGGESTIONS_CU = [
  "Summarize my risk flags",
  "How do I compare to peers?",
  "What should the board focus on?",
];

function ChatPanel({ open, onToggle, cu, peers, opts, flags }) {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const scrollRef = useRef(null);

  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [messages, loading, open]);

  async function send(text) {
    const content = (text ?? input).trim();
    if (!content || loading) return;
    const nextMessages = [...messages, { role: "user", content }];
    setMessages(nextMessages);
    setInput("");
    setLoading(true);
    try {
      const system = buildChatSystemPrompt(cu, peers, opts, flags);
      const response = await fetch(`${import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000"}/api/llm/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 700,
          system,
          messages: nextMessages.slice(-14).map((m) => ({ role: m.role, content: m.content })),
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || "Chat service unavailable");
      const text2 = (data.content || []).map((b) => b.text || "").join("\n").trim();
      setMessages((m) => [...m, { role: "assistant", content: text2 || "I couldn't generate a response — please try again." }]);
    } catch (e) {
      setMessages((m) => [...m, { role: "assistant", content: "Something went wrong reaching the assistant. Please try again." }]);
    } finally {
      setLoading(false);
    }
  }

  const suggestions = cu ? CHAT_SUGGESTIONS_CU : CHAT_SUGGESTIONS_GENERIC;

  return (
    <>
      <button className={`chat-fab no-print ${open ? "chat-fab-open" : ""}`} onClick={onToggle} title={open ? "Close assistant" : "Open assistant"}>
        {open ? <X size={20} strokeWidth={2} /> : <MessageCircle size={20} strokeWidth={1.9} />}
      </button>
      <aside className={`chat-panel no-print ${open ? "chat-panel-open" : ""}`}>
        <div className="chat-head">
          <div className="chat-head-title"><Bot size={16} strokeWidth={1.8} /> CU360 Assistant</div>
          <span className="chat-head-sub">{cu ? `Context: ${shortName(cu.name)}` : "Ask about the platform or the data"}</span>
        </div>
        <div className="chat-scroll" ref={scrollRef}>
          {messages.length === 0 && (
            <div className="chat-empty">
              <p>Ask me anything about {cu ? `${shortName(cu.name)}'s ratios, peers, or risk flags` : "credit union metrics or how to use this platform"}.</p>
              <div className="chat-suggestions">
                {suggestions.map((s) => (
                  <button key={s} className="chat-suggestion" onClick={() => send(s)}>{s}</button>
                ))}
              </div>
            </div>
          )}
          {messages.map((m, i) => (
            <div key={i} className={`chat-bubble-row ${m.role === "user" ? "chat-row-user" : ""}`}>
              <div className={`chat-bubble ${m.role === "user" ? "chat-bubble-user" : "chat-bubble-assistant"}`}>{m.content}</div>
            </div>
          ))}
          {loading && (
            <div className="chat-bubble-row">
              <div className="chat-bubble chat-bubble-assistant chat-typing"><Loader2 size={14} className="spin" /> thinking…</div>
            </div>
          )}
        </div>
        <div className="chat-input-row">
          <input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") send(); }}
            placeholder="Ask a question…"
          />
          <button className="chat-send" onClick={() => send()} disabled={loading || !input.trim()}><Send size={15} /></button>
        </div>
      </aside>
    </>
  );
}

/* ============================== CONNECT ADMIN / DATA ADMIN ============================== */
function ConnectAdminTab({ quarterRegistry, selectedQuarter, onApplyQuarter, onActivateQuarter, setConnectionStatus }) {
  const [source, setSource] = useState("snowflake");
  const [quarterLabel, setQuarterLabel] = useState("");
  const [fileStatus, setFileStatus] = useState(null); // { kind: "success" | "error" | "info", text }
  const [importing, setImporting] = useState(false);
  const [parsed, setParsed] = useState(null);          // last file read: { headers, rows, fileName, sheetName, label }
  const [mapping, setMapping] = useState({});          // field key -> column header
  const [showMapping, setShowMapping] = useState(false);
  const [report, setReport] = useState(null);
  const [snowflake, setSnowflake] = useState({
    account: "", username: "jtrivedi@trellance.com", password: "", warehouse: "", database: "", schema: "", role: "", apiEndpoint: "/api/snowflake/query"
  });
  const [testing, setTesting] = useState(false);
  const [testMessage, setTestMessage] = useState("");
  const [schedule, setSchedule] = useState("quarterly");
  const quarters = Object.keys(quarterRegistry);

  function updateField(key, value) { setSnowflake((x) => ({ ...x, [key]: value })); }

  async function testConnection() {
    setTesting(true); setTestMessage("");
    if (!snowflake.apiEndpoint) {
      setTestMessage("Add a secure server endpoint first. Snowflake credentials should not be sent directly from the browser to Snowflake.");
      setTesting(false); return;
    }
    try {
      const response = await fetch(snowflake.apiEndpoint, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "test", account: snowflake.account, username: snowflake.username, password: snowflake.password, warehouse: snowflake.warehouse, database: snowflake.database, schema: snowflake.schema, role: snowflake.role })
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.message || `Connection test returned ${response.status}`);
      setTestMessage(data.message || "Snowflake connection test succeeded.");
      setConnectionStatus("connected");
    } catch (e) {
      setTestMessage(e.message || "Connection test failed. Check the endpoint and Snowflake settings.");
      setConnectionStatus("error");
    } finally { setTesting(false); }
  }

  function runImport(sheet, map, label) {
    const { records, report: rep } = normalizeQuarterRows(sheet.rows, label, map);
    if (!records.length) throw new Error("No usable rows were found. Check that the credit union name column is matched correctly.");
    const replaced = !!quarterRegistry[label];
    onApplyQuarter(label, records);
    setReport(rep);
    setShowMapping(false);
    setFileStatus({
      kind: "success",
      text: `${label}: ${replaced ? "refreshed" : "loaded"} ${records.length.toLocaleString()} credit unions from ${sheet.fileName}${sheet.sheetName ? ` (sheet “${sheet.sheetName}”)` : ""}. All tabs now use this data — search for a credit union to continue.`,
    });
  }

  async function importQuarter(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImporting(true);
    setReport(null);
    setFileStatus({ kind: "info", text: `Reading ${file.name}…` });
    try {
      const sheet = await readSpreadsheetFile(file);
      sheet.fileName = file.name;
      if (!sheet.rows.length) throw new Error("No data rows were found in that file.");
      const label = quarterLabel.trim() || detectQuarterLabel(sheet) || "New Quarter";
      sheet.label = label;
      const auto = autoMapColumns(sheet.headers);
      setParsed(sheet);
      setMapping(auto);
      const missing = IMPORT_FIELDS.filter((f) => f.required && !auto[f.key]);
      if (missing.length) {
        setShowMapping(true);
        setFileStatus({ kind: "error", text: `Read ${sheet.rows.length.toLocaleString()} rows, but couldn’t find the ${missing.map((f) => f.label.toLowerCase()).join(", ")} column${missing.length > 1 ? "s" : ""}. Choose them below, then import.` });
      } else {
        runImport(sheet, auto, label);
      }
    } catch (err) {
      setFileStatus({ kind: "error", text: err?.message || "Could not read that file." });
    } finally {
      setImporting(false);
      if (e.target) e.target.value = "";
    }
  }

  function confirmMapping() {
    try { runImport(parsed, mapping, quarterLabel.trim() || parsed.label); }
    catch (err) { setFileStatus({ kind: "error", text: err.message }); }
  }

  function activateQuarter(q) { onActivateQuarter(q); }

  return (
    <div className="tab-body">
      <div className="admin-header">
        <div>
          <span className="eyebrow-tag"><Settings size={12} /> Connect Admin</span>
          <h2>Data sources & quarterly intelligence</h2>
          <p>Manage quarter-aware data, secure connections, refresh behavior, and the publish pipeline without changing the existing analytics.</p>
        </div>
        <div className="admin-quarter-pill"><Clock3 size={14} /> Active: <strong>{selectedQuarter}</strong></div>
      </div>

      <div className="admin-grid">
        <div className="panel admin-panel">
          <div className="panel-head"><h3>Quarter data</h3><span className="panel-head-sub">No duplicate institutions per quarter</span></div>
          <div className="quarter-list">
            {quarters.map((q) => (
              <div className={`quarter-row ${q === selectedQuarter ? "quarter-active" : ""}`} key={q}>
                <div><strong>{q}</strong><span>{quarterRegistry[q].length.toLocaleString()} institutions</span></div>
                <button className="btn-secondary" onClick={() => activateQuarter(q)}>{q === selectedQuarter ? "Active" : "Use quarter"}</button>
              </div>
            ))}
          </div>
          <div className="field-grid">
            <label><span>New quarter label</span><input value={quarterLabel} onChange={(e) => setQuarterLabel(e.target.value)} placeholder="2024 Q1" /></label>
          </div>
          <div className="admin-actions">
            <label className={`btn-primary file-button ${importing ? "is-busy" : ""}`}>
              {importing ? <Loader2 size={14} className="spin" /> : <Upload size={14} />} {importing ? "Reading…" : "Upload Excel / CSV / JSON"}
              <input type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.ods,.csv,.tsv,.json,.txt" onChange={importQuarter} disabled={importing} hidden />
            </label>
            <button className="btn-secondary" onClick={downloadQuarterTemplate}><Download size={14} /> Excel template</button>
            {parsed && !showMapping && <button className="btn-secondary" onClick={() => setShowMapping(true)}><SlidersHorizontal size={14} /> Adjust columns</button>}
          </div>
          <p className="admin-hint">Accepts .xlsx, .xls, .csv and .json. Column headers are matched automatically, and amounts in dollars or thousands are converted to $ millions. Leave the quarter label blank to read it from the file name or a cycle-date column; uploading a quarter that already exists refreshes it.</p>
          {fileStatus && (
            <div className={`admin-message ${fileStatus.kind}`}>
              {fileStatus.kind === "success" ? <CheckCircle2 size={14} /> : fileStatus.kind === "error" ? <AlertCircle size={14} /> : <Loader2 size={14} className="spin" />} {fileStatus.text}
            </div>
          )}
          {report && (
            <div className="import-report">
              <div><b>{report.mapped}</b> of {report.totalFields} fields matched · <b>{report.total.toLocaleString()}</b> rows read{report.duplicates ? ` · ${report.duplicates.toLocaleString()} duplicate rows collapsed` : ""}{report.skipped ? ` · ${report.skipped.toLocaleString()} rows without a name skipped` : ""}</div>
              {report.notes.map((n) => <div key={n}>{n}</div>)}
              {report.derived.length > 0 && <div>Calculated from other columns: {report.derived.join(", ")}.</div>}
              {report.missing.length > 0 && <div className="import-missing">Not in the file (shown as 0): {report.missing.join(", ")}. Use “Adjust columns” if one of these exists under a different name.</div>}
            </div>
          )}
          {showMapping && parsed && (
            <div className="mapping-panel">
              <div className="mapping-head"><strong>Match your columns</strong><span>{parsed.fileName} · {parsed.rows.length.toLocaleString()} rows · {parsed.headers.length} columns</span></div>
              <div className="mapping-grid">
                {IMPORT_FIELDS.map((f) => (
                  <label key={f.key}>
                    <span>{f.label}{f.required ? " *" : ""}</span>
                    <select className={f.required && !mapping[f.key] ? "is-required-missing" : ""} value={mapping[f.key] || ""} onChange={(e) => setMapping((m) => ({ ...m, [f.key]: e.target.value }))}>
                      <option value="">— not in file —</option>
                      {parsed.headers.map((h) => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </label>
                ))}
              </div>
              <div className="admin-actions">
                <button className="btn-primary" onClick={confirmMapping} disabled={!mapping.name || !mapping.state || !mapping.assetsM}>Import with these columns</button>
                <button className="btn-secondary" onClick={() => setShowMapping(false)}>Close</button>
              </div>
            </div>
          )}
        </div>

        <div className="panel admin-panel">
          <div className="panel-head"><h3>Snowflake connection</h3><span className="panel-head-sub"><LockKeyhole size={12} /> password never stored in portal state</span></div>
          <div className="connection-mode-row">
            {["snowflake", "oauth", "backend"].map((m) => <button key={m} className={`chip ${source === m ? "chip-on" : ""}`} onClick={() => setSource(m)}>{m === "snowflake" ? "Username / password" : m === "oauth" ? "Snowflake OAuth" : "Secure backend"}</button>)}
          </div>
          <div className="field-grid two">
            <label><span>Account identifier</span><input value={snowflake.account} onChange={(e) => updateField("account", e.target.value)} placeholder="org-account" /></label>
            <label><span>Username</span><input value={snowflake.username} onChange={(e) => updateField("username", e.target.value)} /></label>
            <label><span>Password</span><input type="password" value={snowflake.password} onChange={(e) => updateField("password", e.target.value)} placeholder="Enter only when testing" /></label>
            <label><span>Warehouse</span><input value={snowflake.warehouse} onChange={(e) => updateField("warehouse", e.target.value)} placeholder="CU360_WH" /></label>
            <label><span>Database</span><input value={snowflake.database} onChange={(e) => updateField("database", e.target.value)} placeholder="CU360" /></label>
            <label><span>Schema</span><input value={snowflake.schema} onChange={(e) => updateField("schema", e.target.value)} placeholder="PUBLIC" /></label>
            <label><span>Role</span><input value={snowflake.role} onChange={(e) => updateField("role", e.target.value)} placeholder="ANALYST" /></label>
            <label><span>Secure API endpoint</span><input value={snowflake.apiEndpoint} onChange={(e) => updateField("apiEndpoint", e.target.value)} /></label>
          </div>
          <div className="admin-actions">
            <button className="btn-primary" onClick={testConnection} disabled={testing}>{testing ? <><Loader2 size={14} className="spin" /> Testing…</> : <><Link2 size={14} /> Test connection</>}</button>
            <span className="security-note"><KeyRound size={13} /> Use a backend proxy / serverless function for production. Never embed Snowflake passwords in the JSX bundle.</span>
          </div>
          {testMessage && <div className={`admin-message ${testMessage.toLowerCase().includes("succeeded") ? "success" : "error"}`}><AlertCircle size={14} /> {testMessage}</div>}
        </div>
      </div>

      <div className="admin-grid">
        <div className="panel admin-panel">
          <div className="panel-head"><h3>Data processing pipeline</h3><span className="panel-head-sub">Source → validate → deduplicate → publish</span></div>
          <div className="pipeline"><div><Database size={17}/><strong>Source</strong><span>Snowflake / file</span></div><ChevronRight/><div><SlidersHorizontal size={17}/><strong>Validate</strong><span>schema + types</span></div><ChevronRight/><div><CheckCircle2 size={17}/><strong>Deduplicate</strong><span>name + state</span></div><ChevronRight/><div><Workflow size={17}/><strong>Publish</strong><span>quarter dataset</span></div></div>
        </div>
        <div className="panel admin-panel">
          <div className="panel-head"><h3>Refresh policy</h3><span className="panel-head-sub">Designed for recurring quarterly updates</span></div>
          <div className="refresh-row"><RefreshCw size={17}/><div><strong>Scheduled refresh</strong><span>Use the secure backend endpoint to pull the latest Snowflake quarter.</span></div><select value={schedule} onChange={(e) => setSchedule(e.target.value)}><option value="manual">Manual</option><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="quarterly">Quarterly</option></select></div>
          <div className="refresh-row"><Server size={17}/><div><strong>Live query mode</strong><span>Keep dashboards on published snapshots or query Snowflake on demand.</span></div><span className="status-badge">Configurable</span></div>
        </div>
      </div>
    </div>
  );
}

/* ============================== APP SHELL ============================== */
const ADMIN_EMAILS = ["jayt@trellance.com", "jtrivedi@trellance.com"];

const NAV = [
  { id: "overview", label: "Overview", icon: Building2 },
  { id: "benchmark", label: "Peer Benchmarking", icon: TrendingUp },
  { id: "competitive", label: "Competitive Intel", icon: MapPin },
  { id: "compare", label: "Compare Institutions", icon: GitCompare },
  { id: "trend", label: "Trend Comparison", icon: TrendingUp },
  { id: "demographics", label: "Demographic Views", icon: Users },
  { id: "gap", label: "Gap Analysis", icon: Target },
  { id: "risk", label: "Risk Alerts", icon: Shield },
  { id: "board", label: "Board Report", icon: FileText },
  { id: "ai", label: "AI Recommendations", icon: Sparkles },
  { id: "admin", label: "Connect Admin", icon: Database },
];

export default function App() {
  const [cu, setCU] = useState(null);
  const [tab, setTab] = useState("overview");
  const [opts, setOpts] = useState({ sameState: false, wideTier: false });
  const [theme, setTheme] = useState("dark");
  const [chatOpen, setChatOpen] = useState(false);
  const [quarterRegistry, setQuarterRegistry] = useState({ [DEFAULT_QUARTER]: ALL_CUS });
  const [selectedQuarter, setSelectedQuarter] = useState(DEFAULT_QUARTER);
  const [connectionStatus, setConnectionStatus] = useState("not-configured");
  const [, setDataVersion] = useState(0);
  const [authUser, setAuthUser] = useState(null);
  const [loginOpen, setLoginOpen] = useState(false);
  const [loginEmail, setLoginEmail] = useState("");
  const isAdmin = !!authUser && ADMIN_EMAILS.includes(authUser.email.trim().toLowerCase());
  const visibleNav = isAdmin ? NAV : NAV.filter((n) => n.id !== "admin");
  useEffect(() => { if (!isAdmin && tab === "admin") setTab("overview"); }, [isAdmin, tab]);

  const [dataStatus, setDataStatus] = useState("loading");
  useEffect(() => {
    let cancelled = false;
    const apiBase = import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000";
    fetch(`${apiBase}/api/institutions`)
      .then((r) => { if (!r.ok) throw new Error(`Data API returned ${r.status}`); return r.json(); })
      .then((payload) => {
        if (cancelled) return;
        const rows = (payload.rows || []).map((row, idx) => Array.isArray(row) ? rowToCU(row, idx) : row);
        ALL_CUS = rows;
        setActiveData(rows, payload.quarter || DEFAULT_QUARTER);
        setQuarterRegistry({ [payload.quarter || DEFAULT_QUARTER]: rows });
        setSelectedQuarter(payload.quarter || DEFAULT_QUARTER);
        setDataVersion((v) => v + 1);
        setDataStatus("ready");
      })
      .catch((error) => { console.error("CU360 data load failed", error); if (!cancelled) setDataStatus("error"); });
    return () => { cancelled = true; };
  }, []);

  // One place that swaps the active dataset, so every tab refreshes together.
  function activateQuarter(q, registry = quarterRegistry) {
    const rows = registry[q];
    if (!rows) return;
    const nextRows = rows;
    // Keep the current institution selected when the same CU exists in the target quarter.
    // The active tab and existing benchmark filters are intentionally left untouched.
    const nextCU = findQuarterCU(nextRows, cu);
    setActiveData(nextRows, q);
    setSelectedQuarter(q);
    if (cu) setCU(nextCU);
    setDataVersion((v) => v + 1);
  }
  function changeQuarter(q) { activateQuarter(q); }
  function applyQuarter(label, rows) {
    setQuarterRegistry((cur) => Object.fromEntries(Object.entries({ ...cur, [label]: rows }).sort(([a], [b]) => a.localeCompare(b))));
    setActiveData(rows, label);
    setSelectedQuarter(label);
    if (cu) setCU(findQuarterCU(rows, cu));
    setDataVersion((v) => v + 1);
  }

  const peers = usePeerGroup(cu, opts);
  const flags = useMemo(() => (cu ? computeRiskFlags(cu, peers) : []), [cu, peers]);

  return (
    <div className={`cu360 theme-${theme}`}>
      <style>{CSS}</style>
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">CU</span>
          <div className="brand-text">
            <span className="brand-name">CU360<sup>™</sup></span>
            <span className="brand-tag">Intelligence Platform</span>
          </div>
        </div>
        <nav>
          <button
            className={`nav-item home-item ${!cu ? "nav-active" : ""}`}
            onClick={() => { setCU(null); setTab("overview"); }}
          >
            <Home size={17} strokeWidth={1.7} />
            Home
          </button>
          <div className="nav-divider" />
          {visibleNav.map((n) => {
            const Icon = n.icon;
            const disabled = !cu && n.id !== "admin";
            return (
              <button
                key={n.id}
                className={`nav-item ${tab === n.id ? "nav-active" : ""} ${disabled ? "nav-disabled" : ""}`}
                onClick={() => (cu || n.id === "admin") && setTab(n.id)}
                disabled={disabled}
              >
                <Icon size={17} strokeWidth={1.7} />
                {n.label}
              </button>
            );
          })}
        </nav>
        <div className="sidebar-foot">
          <Info size={13} />
          <span>Data - NCUA Call Report.</span>
        </div>
      </aside>

      <div className="main">
        <header className="topbar no-print">
          <CUSearch onSelect={(c) => { setCU(c); setTab("overview"); }} placeholder={cu ? "Switch credit union…" : "Search your credit union by name…"} />
          <div className="topbar-right">
            <div className="quarter-selector" title="Active reporting quarter">
              <Clock3 size={14} />
              <select value={selectedQuarter} onChange={(e) => changeQuarter(e.target.value)}>
                {Object.keys(quarterRegistry).map((q) => <option key={q} value={q}>{q}</option>)}
              </select>
            </div>
            {cu && (
              <div className="topbar-cu">
                <span className="topbar-cu-name">{titleCase(cu.name)}</span>
                <span className="topbar-cu-meta">{cu.state} · {fmtM(cu.assetsM)}</span>
              </div>
            )}
            {authUser ? <button className="login-pill" onClick={() => { setAuthUser(null); setTab("overview"); setCU(null); }} title="Sign out"><UserCircle size={15}/><span>{authUser.email}</span>{isAdmin && <b>Admin</b>}<LogOut size={14}/></button> : <button className="login-pill" onClick={() => setLoginOpen(true)}><LogIn size={15}/><span>Login</span></button>}
            <button className="theme-toggle" onClick={() => setTheme((t) => (t === "dark" ? "light" : "dark"))} title="Toggle light / dark mode">
              {theme === "dark" ? <Sun size={16} strokeWidth={1.8} /> : <Moon size={16} strokeWidth={1.8} />}
            </button>
          </div>
        </header>

        <main className="content">
          {dataStatus === "loading" && <div className="panel" role="status">Loading CU360 data…</div>}
          {dataStatus === "error" && <div className="panel" role="alert">Unable to load dashboard data. Confirm the FastAPI backend is running, then refresh.</div>}
          {dataStatus === "ready" && (<>
          {!cu && tab !== "admin" ? (
            <div className="landing">
              <span className="eyebrow-tag">{selectedQuarter} · NCUA Call Report Data · {ACTIVE_CUS.length.toLocaleString()} institutions</span>
              <h1>Know exactly where your credit union stands.</h1>
              <p>Search your institution to unlock peer benchmarking, competitive intelligence, risk alerts, board-ready reporting, and AI-drafted recommendations — all from the active reporting quarter.</p>
              <div className="landing-search"><CUSearch onSelect={(c) => { setCU(c); setTab("overview"); }} /></div>
              <div className="landing-features">
                {visibleNav.map((n) => {
                  const Icon = n.icon;
                  return <div className="landing-feature" key={n.id}><Icon size={16} strokeWidth={1.7} /> {n.label}</div>;
                })}
              </div>
            </div>
          ) : (
            <>
              {tab === "overview" && <OverviewTab cu={cu} peers={peers} opts={opts} />}
              {tab === "benchmark" && <BenchmarkTab cu={cu} peers={peers} opts={opts} setOpts={setOpts} />}
              {tab === "competitive" && <CompetitiveTab cu={cu} />}
              {tab === "risk" && <RiskTab cu={cu} peers={peers} flags={flags} opts={opts} />}
              {tab === "board" && <BoardReportTab cu={cu} peers={peers} opts={opts} flags={flags} />}
              {tab === "ai" && <AITab cu={cu} peers={peers} opts={opts} flags={flags} />}
              {tab === "compare" && <CompareTab primary={cu} />}
              {tab === "trend" && <TrendComparisonTab primary={cu} quarterRegistry={quarterRegistry} opts={opts} />}
              {tab === "demographics" && <DemographicViewsTab cu={cu} peers={peers} onSelectCU={setCU} />}
              {tab === "gap" && <GapAnalysisTab cu={cu} peers={peers} />}
              {tab === "spend" && <SpendAnalysisTab cu={cu} peers={peers} />}
              {tab === "admin" && isAdmin && <ConnectAdminTab quarterRegistry={quarterRegistry} selectedQuarter={selectedQuarter} onApplyQuarter={applyQuarter} onActivateQuarter={activateQuarter} setConnectionStatus={setConnectionStatus} />}
            </>
          )}
          </>)}
        </main>
      </div>
      <ChatPanel open={chatOpen} onToggle={() => setChatOpen((o) => !o)} cu={cu} peers={peers} opts={opts} flags={flags} />
      {loginOpen && <div className="login-overlay" onMouseDown={() => setLoginOpen(false)}><div className="login-modal" onMouseDown={(e) => e.stopPropagation()}><div className="login-modal-head"><div><span className="eyebrow-tag">CU360 access</span><h2>Sign in</h2></div><button className="icon-btn" onClick={() => setLoginOpen(false)}><X size={17}/></button></div><p>Sign in with your work email. Administrative tools are shown only to the authorized CU360 administrator.</p><label className="login-label">Work email<input autoFocus type="email" value={loginEmail} onChange={(e) => setLoginEmail(e.target.value)} placeholder="you@company.com" onKeyDown={(e) => { if (e.key === "Enter" && loginEmail.trim()) { setAuthUser({email: loginEmail.trim()}); setLoginOpen(false); setLoginEmail(""); } }}/></label><button className="login-submit" disabled={!loginEmail.trim()} onClick={() => { setAuthUser({email: loginEmail.trim()}); setLoginOpen(false); setLoginEmail(""); }}>Continue <ChevronRight size={15}/></button><small className="login-security-note">UI access gate only. Production authentication should be backed by your identity provider.</small></div></div>}
    </div>
  );
}

/* ============================== STYLES ============================== */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

.cu360 * { box-sizing: border-box; }
.cu360 {
  --ink: #0D1015;
  --panel: #171B22;
  --panel-2: #1D222B;
  --line: rgba(255,255,255,0.08);
  --text: #E9E6DD;
  --muted: #8B92A0;
  --muted-2: #5D6472;
  --brass: #C9A227;
  --brass-soft: rgba(201,162,39,0.14);
  --risk: #D65A4A;
  --warn: #D6A94A;
  --good: #4C9A6A;
  --sidebar-bg: #0A0C10;
  --topbar-bg: rgba(13,16,21,0.7);
  --shadow: rgba(0,0,0,0.45);
  font-family: 'IBM Plex Sans', -apple-system, sans-serif;
  background: var(--ink);
  color: var(--text);
  display: flex;
  min-height: 100vh;
  width: 100%;
  font-size: 14px;
  transition: background .15s, color .15s;
}
.cu360.theme-light {
  --ink: #F7F5EE;
  --panel: #FFFFFF;
  --panel-2: #F1EEE3;
  --line: rgba(30,27,20,0.10);
  --text: #211E17;
  --muted: #6B6752;
  --muted-2: #948F7B;
  --brass: #9C6F13;
  --brass-soft: rgba(156,111,19,0.12);
  --risk: #C1453B;
  --warn: #A9791E;
  --good: #2F7A50;
  --sidebar-bg: #FFFFFF;
  --topbar-bg: rgba(247,245,238,0.85);
  --shadow: rgba(30,27,20,0.14);
}
.cu360 h1, .cu360 h2, .cu360 h3, .cu360 h4 { font-family: 'Fraunces', serif; margin: 0; font-weight: 500; }
.mono { font-family: 'IBM Plex Mono', monospace; }

/* Sidebar */
.sidebar { width: 216px; flex-shrink: 0; background: var(--sidebar-bg); border-right: 1px solid var(--line); display: flex; flex-direction: column; padding: 22px 14px; }
.brand { display: flex; align-items: center; gap: 10px; padding: 0 6px 22px 6px; }
.brand-mark { width: 34px; height: 34px; border: 1px solid var(--brass); border-radius: 3px; display: flex; align-items: center; justify-content: center; font-family: 'Fraunces', serif; font-weight: 600; color: var(--brass); font-size: 13px; }
.brand-text { display: flex; flex-direction: column; line-height: 1.25; }
.brand-name { font-family: 'Fraunces', serif; font-size: 16.5px; color: var(--text); }
.brand-name sup { font-size: 9px; }
.brand-tag { font-size: 10.5px; color: var(--muted-2); letter-spacing: 0.02em; }
.sidebar nav { display: flex; flex-direction: column; gap: 2px; margin-top: 4px; }
.nav-item { display: flex; align-items: center; gap: 10px; padding: 9px 10px; border-radius: 6px; background: none; border: none; color: var(--muted); font-size: 13.3px; text-align: left; cursor: pointer; font-family: inherit; transition: background .12s, color .12s; width: 100%; }
.nav-item:hover:not(.nav-disabled) { background: var(--panel-2); color: var(--text); }
.nav-active { background: var(--brass-soft); color: var(--brass) !important; }
.nav-disabled { opacity: 0.35; cursor: not-allowed; }
.nav-divider { height: 1px; background: var(--line); margin: 8px 4px 10px; }
.sidebar-foot { margin-top: auto; display: flex; gap: 6px; color: var(--muted-2); font-size: 10.8px; padding: 10px 6px 0; border-top: 1px solid var(--line); align-items: flex-start; line-height: 1.4; }

/* Main / topbar */
.main { flex: 1; display: flex; flex-direction: column; min-width: 0; }
.topbar { display: flex; align-items: center; justify-content: space-between; gap: 20px; padding: 14px 28px; border-bottom: 1px solid var(--line); background: var(--topbar-bg); backdrop-filter: blur(6px); position: sticky; top: 0; z-index: 5; }
.topbar-right { display: flex; align-items: center; gap: 14px; }
.topbar-cu { display: flex; flex-direction: column; align-items: flex-end; line-height: 1.3; }
.topbar-cu-name { font-size: 13.5px; font-weight: 500; }
.topbar-cu-meta { font-size: 11.5px; color: var(--muted); }
.theme-toggle { display: flex; align-items: center; justify-content: center; width: 32px; height: 32px; border-radius: 7px; border: 1px solid var(--line); background: var(--panel); color: var(--muted); cursor: pointer; flex-shrink: 0; }
.theme-toggle:hover { color: var(--brass); border-color: var(--brass); }
.content { flex: 1; overflow-y: auto; padding: 28px 32px 60px; }

/* Search */
.cu-search { position: relative; width: 380px; max-width: 100%; }
.cu-search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: var(--muted-2); }
.cu-search input { width: 100%; background: var(--panel); border: 1px solid var(--line); border-radius: 7px; padding: 9px 12px 9px 34px; color: var(--text); font-size: 13.3px; font-family: inherit; }
.cu-search input:focus { outline: none; border-color: var(--brass); }
.cu-search-results { position: absolute; top: calc(100% + 6px); left: 0; right: 0; background: var(--panel); border: 1px solid var(--line); border-radius: 8px; overflow: hidden; z-index: 20; box-shadow: 0 12px 32px var(--shadow); max-height: 340px; overflow-y: auto; }
.cu-search-row { display: flex; flex-direction: column; gap: 2px; width: 100%; text-align: left; padding: 9px 13px; background: none; border: none; border-bottom: 1px solid var(--line); cursor: pointer; font-family: inherit; }
.cu-search-row:last-child { border-bottom: none; }
.cu-search-row:hover { background: var(--panel-2); }
.cu-search-name { font-size: 13px; color: var(--text); }
.cu-search-meta { font-size: 11.3px; color: var(--muted); }
.cu-search-empty { padding: 12px; font-size: 12.5px; color: var(--muted); }

/* Landing */
.landing { max-width: 620px; margin: 60px auto; text-align: left; }
.landing h1 { font-size: 34px; line-height: 1.15; margin: 14px 0 12px; }
.landing p { color: var(--muted); font-size: 14.5px; line-height: 1.6; margin: 0 0 26px; }
.landing-search { margin-bottom: 34px; }
.landing-search .cu-search { width: 100%; }
.landing-features { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.landing-feature { display: flex; align-items: center; gap: 9px; padding: 10px 13px; background: var(--panel); border: 1px solid var(--line); border-radius: 7px; font-size: 12.8px; color: var(--muted); }

.eyebrow-tag { display: inline-block; font-size: 11.5px; color: var(--brass); letter-spacing: 0.01em; margin-bottom: 4px; }

/* Peer group banner */
.peer-banner { display: flex; align-items: center; gap: 8px; padding: 9px 14px; border: 1px solid var(--line); background: var(--brass-soft); border-radius: 7px; font-size: 12.3px; color: var(--muted); margin-bottom: 18px; }
.peer-banner strong { color: var(--brass); font-weight: 600; }
.peer-banner svg { color: var(--brass); flex-shrink: 0; }

/* Overview hero */
.hero-row { display: flex; justify-content: space-between; align-items: flex-end; gap: 20px; margin-bottom: 26px; flex-wrap: wrap; }
.hero-block h1 { font-size: 27px; margin: 4px 0 8px; }
.hero-sub { display: flex; align-items: center; gap: 5px; color: var(--muted); font-size: 13px; margin: 0; }
.hero-figure { text-align: right; }
.hero-figure-label { display: block; font-size: 11.5px; color: var(--muted); margin-bottom: 3px; }
.hero-figure-value { display: block; font-family: 'Fraunces', serif; font-size: 30px; color: var(--brass); }
.hero-figure-sub { display: block; font-size: 11.5px; color: var(--muted-2); margin-top: 3px; }

/* KPI grid */
.kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 22px; }
.kpi { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 14px 15px; display: flex; flex-direction: column; gap: 5px; }
.kpi-warn { border-color: rgba(214,90,74,0.4); }
.kpi-label { font-size: 11.3px; color: var(--muted); }
.kpi-value { font-family: 'IBM Plex Mono', monospace; font-size: 21px; }
.kpi-sub { font-size: 11px; color: var(--muted-2); }
.kpi-warn .kpi-sub { color: var(--risk); }

/* Panels */
.panel { background: var(--panel); border: 1px solid var(--line); border-radius: 9px; padding: 18px 20px; margin-bottom: 18px; }
.panel-head { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 14px; flex-wrap: wrap; gap: 4px; }
.panel-head h3 { font-size: 16px; }
.panel-head-sub { font-size: 11.5px; color: var(--muted-2); }
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }

/* Percentile rank list */
.metric-rank-list { display: flex; flex-direction: column; gap: 14px; }
.metric-rank-row { display: grid; grid-template-columns: 200px 1fr 70px; align-items: center; gap: 12px; }
.metric-rank-label { display: flex; flex-direction: column; font-size: 12.5px; color: var(--muted); }
.metric-rank-value { font-family: 'IBM Plex Mono', monospace; color: var(--text); font-size: 13.5px; }
.metric-rank-pct { font-size: 11px; color: var(--muted-2); text-align: right; }
.pbar-track { position: relative; height: 6px; background: rgba(255,255,255,0.07); border-radius: 4px; overflow: visible; }
.pbar-fill { height: 100%; border-radius: 4px; transition: width .3s; }
.pbar-mid { position: absolute; left: 50%; top: -2px; bottom: -2px; width: 1px; background: rgba(255,255,255,0.18); }

/* Compare list (benchmark tab) */
.compare-list { display: flex; flex-direction: column; gap: 12px; }
.compare-row { display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid var(--line); }
.compare-row:last-child { border-bottom: none; padding-bottom: 0; }
.compare-label { font-size: 12.6px; color: var(--muted); max-width: 150px; }
.compare-values { display: flex; align-items: center; gap: 9px; }
.compare-you { font-family: 'IBM Plex Mono', monospace; font-size: 13.5px; }
.compare-vs { font-size: 11px; color: var(--muted-2); }

.delta { display: inline-flex; align-items: center; gap: 2px; font-size: 11.5px; font-family: 'IBM Plex Mono', monospace; padding: 2px 6px; border-radius: 4px; }
.delta.good { color: var(--good); background: rgba(76,154,106,0.13); }
.delta.bad { color: var(--risk); background: rgba(214,90,74,0.13); }
.delta.flat { color: var(--muted-2); background: rgba(255,255,255,0.05); }

/* Opts bar */
.opts-bar { display: flex; align-items: center; gap: 10px; margin-bottom: 18px; }
.chip { display: flex; align-items: center; gap: 5px; padding: 7px 12px; border-radius: 20px; border: 1px solid var(--line); background: var(--panel); color: var(--muted); font-size: 12px; cursor: pointer; font-family: inherit; }
.chip-on { border-color: var(--brass); color: var(--brass); background: var(--brass-soft); }
.opts-count { margin-left: auto; font-size: 11.5px; color: var(--muted-2); }

/* Table */
.table-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 12.8px; }
th { text-align: left; padding: 8px 10px; color: var(--muted-2); font-weight: 500; font-size: 11px; border-bottom: 1px solid var(--line); cursor: pointer; white-space: nowrap; }
th.th-active { color: var(--brass); }
td { padding: 9px 10px; border-bottom: 1px solid var(--line); vertical-align: top; }
.td-name { min-width: 190px; }
.td-sub { font-size: 10.8px; color: var(--muted-2); margin-top: 2px; }
.you-tag { margin-left: 8px; font-size: 9.5px; color: var(--brass); border: 1px solid var(--brass); border-radius: 3px; padding: 1px 5px; }
.row-you { background: var(--brass-soft); }

/* Competitive tab */
.competitive-hero { align-items:flex-end; }
.competitive-scope { display:flex; flex-direction:column; align-items:flex-end; gap:7px; min-width:340px; }
.competitive-scope > span { color:var(--muted-2); font-size:10px; text-transform:uppercase; }
.competitive-scope small { color:var(--muted-2); font-size:10.5px; }
.competitive-scope-buttons { display:flex; justify-content:flex-end; gap:6px; flex-wrap:wrap; }
.competitive-scope-buttons .chip { border-radius:7px; padding:7px 10px; }
.competitive-search { position:relative; margin-bottom:10px; }
.competitive-search input { width:100%; border:1px solid var(--line); background:var(--panel-2); color:var(--text); border-radius:7px; padding:9px 11px; font:inherit; font-size:12px; }
.competitive-search input:focus { outline:none; border-color:var(--brass); }
.competitive-suggestions { position:absolute; top:calc(100% + 4px); left:0; right:0; z-index:5; max-height:230px; overflow:auto; background:var(--panel); border:1px solid var(--line); border-radius:7px; box-shadow:0 10px 26px var(--shadow); }
.competitive-suggestions button { width:100%; border:0; border-bottom:1px solid var(--line); background:transparent; color:var(--text); padding:9px 11px; display:flex; justify-content:space-between; align-items:center; gap:10px; text-align:left; cursor:pointer; font:inherit; font-size:11.5px; }
.competitive-suggestions button:hover { background:var(--brass-soft); }
.competitive-suggestions small { color:var(--muted-2); white-space:nowrap; }
.selected-rivals { display:flex; flex-wrap:wrap; gap:6px; margin:0 0 10px; }
.selected-rivals .chip { border-radius:6px; cursor:pointer; }
.competitive-table { min-width:660px; }
.competitive-table th, .competitive-table td { white-space:nowrap; }
.competitive-table th:first-child, .competitive-table td:first-child { position:sticky; left:0; z-index:1; background:var(--panel); }
.competitive-table th:first-child { z-index:2; }
.competitive-footnote { display:block; color:var(--muted-2); font-size:10.5px; line-height:1.45; margin-top:9px; }
.competitive-signals { display:flex; flex-direction:column; gap:8px; }
.competitive-signal { display:flex; align-items:flex-start; gap:9px; padding:10px 11px; border:1px solid var(--line); border-radius:7px; background:var(--panel-2); }
.competitive-signal > span { flex-shrink:0; margin-top:1px; }
.competitive-signal div { display:flex; flex-direction:column; gap:3px; }
.competitive-signal strong { font-size:11.5px; }
.competitive-signal small { color:var(--muted); font-size:10.5px; line-height:1.4; }
.signal-risk > span, .signal-risk strong { color:var(--risk); }
.signal-good > span { color:var(--good); }
.competitive-limit { display:flex; align-items:flex-start; gap:10px; color:var(--muted); }
.competitive-limit > svg { flex-shrink:0; color:var(--brass); margin-top:2px; }
.competitive-limit div { display:flex; flex-direction:column; gap:4px; }
.competitive-limit strong { color:var(--text); font-size:12px; }
.competitive-limit span { font-size:11px; line-height:1.5; }
.rival-list { display: flex; flex-direction: column; gap: 2px; }
.rival-row { display: flex; justify-content: space-between; align-items: center; padding: 11px 4px; border-bottom: 1px solid var(--line); }
.rival-row:last-child { border-bottom: none; }
.rival-name { font-size: 13px; }
.rival-metrics { display: flex; gap: 16px; }
.mini-stat { display: flex; flex-direction: column; align-items: flex-end; }
.mini-stat-label { font-size: 10px; color: var(--muted-2); }
.mini-stat-value { font-family: 'IBM Plex Mono', monospace; font-size: 12.5px; }
.mini-stat-value.good { color: var(--good); }

/* Compare tab */
.compare-add-panel { display: flex; flex-direction: column; gap: 12px; }
.compare-add-head { display: flex; justify-content: space-between; align-items: baseline; }
.compare-add-head h3 { font-size: 16px; }
.compare-search { width: 100%; }
.compare-search input { width: 100%; }
.compare-cards { display: grid; gap: 14px; margin-bottom: 18px; }
.compare-card { position: relative; background: var(--panel); border: 1px solid var(--line); border-radius: 9px; padding: 16px 16px 14px; display: flex; flex-direction: column; gap: 3px; }
.compare-card h4 { font-size: 14.5px; line-height: 1.3; margin-top: 4px; padding-right: 20px; }
.compare-card-assets { font-family: 'IBM Plex Mono', monospace; font-size: 19px; color: var(--brass); margin: 4px 0 2px; }
.compare-remove { position: absolute; top: 12px; right: 12px; background: none; border: none; color: var(--muted-2); cursor: pointer; padding: 2px; }
.compare-remove:hover { color: var(--risk); }
.compare-you-tag { position: absolute; top: 14px; right: 14px; }
.compare-table th:first-child, .compare-table td:first-child { min-width: 160px; color: var(--muted); }
.compare-metric-label { font-size: 12.5px; }
.compare-best { color: var(--good); font-weight: 500; }

/* Risk */
.flag-list { display: flex; flex-direction: column; gap: 10px; }
.flag-row { border-left: 3px solid; background: var(--panel-2); border-radius: 0 7px 7px 0; padding: 11px 14px; }
.flag-row-top { display: flex; align-items: center; gap: 9px; margin-bottom: 4px; }
.flag-sev { font-size: 10px; padding: 2px 7px; border-radius: 4px; font-weight: 600; letter-spacing: 0.02em; }
.flag-title { font-size: 13.5px; }
.flag-detail { margin: 0; font-size: 12.3px; color: var(--muted); line-height: 1.5; }
.method-note { font-size: 12.3px; color: var(--muted); line-height: 1.6; margin: 0; }
.sev-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }

/* Buttons */
.btn-primary { display: inline-flex; align-items: center; gap: 7px; background: var(--brass); color: #1A1503; border: none; padding: 9px 16px; border-radius: 7px; font-size: 13px; font-weight: 600; cursor: pointer; font-family: inherit; white-space: nowrap; }
.btn-primary:disabled { opacity: 0.6; cursor: default; }
.spin { animation: spin 1s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }

/* Empty state */
.empty-state { display: flex; flex-direction: column; align-items: center; text-align: center; gap: 8px; padding: 34px 20px; color: var(--muted); }
.empty-state h3 { font-size: 15px; color: var(--text); }
.empty-state p { margin: 0; font-size: 12.8px; max-width: 340px; }

/* Board report */
.board-toolbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; color: var(--muted); font-size: 13px; }
.board-sheet { background: #F6F3EA; color: #1E1B14; border-radius: 8px; padding: 34px 38px; }
.board-sheet-head { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #1E1B14; padding-bottom: 16px; margin-bottom: 20px; }
.board-sheet-head h2 { font-size: 24px; margin: 4px 0; color: #1E1B14; }
.board-sheet-head p { margin: 0; font-size: 12.5px; color: #5A5646; }
.board-kicker { font-size: 11px; letter-spacing: 0.02em; color: #8A6B14; }
.board-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 24px; }
.board-stat { display: flex; flex-direction: column; gap: 3px; border: 1px solid #D8D2BF; border-radius: 6px; padding: 10px 12px; }
.board-stat span { font-size: 10.8px; color: #6B6752; }
.board-stat strong { font-family: 'IBM Plex Mono', monospace; font-size: 17px; color: #1E1B14; font-weight: 500; }
.board-section-title { font-size: 14px; margin: 22px 0 10px; color: #1E1B14; }
.board-table { width: 100%; border-collapse: collapse; font-size: 12.3px; }
.board-table th { text-align: left; padding: 7px 9px; border-bottom: 1px solid #1E1B14; color: #4A4636; font-weight: 600; }
.board-table td { padding: 7px 9px; border-bottom: 1px solid #D8D2BF; }
.board-note { font-size: 12.5px; color: #4A4636; }
.board-flag-list { margin: 0; padding-left: 18px; font-size: 12.5px; color: #3A3626; line-height: 1.7; }
.board-footer { margin-top: 26px; padding-top: 14px; border-top: 1px solid #D8D2BF; font-size: 10.8px; color: #7A7660; }

@media print {
  .no-print, .sidebar { display: none !important; }
  .cu360 { background: white; }
  .content { padding: 0; overflow: visible; }
  .board-sheet { border-radius: 0; }
}

/* AI tab */
.ai-intro { display: flex; align-items: center; gap: 16px; }
.ai-intro-icon { width: 42px; height: 42px; border-radius: 8px; background: var(--brass-soft); color: var(--brass); display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.ai-intro h3 { font-size: 15.5px; margin-bottom: 4px; }
.ai-intro p { margin: 0; font-size: 12.6px; color: var(--muted); line-height: 1.5; }
.ai-intro .btn-primary { margin-left: auto; flex-shrink: 0; }
.rec-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
.rec-card { background: var(--panel); border: 1px solid var(--line); border-radius: 9px; padding: 16px 18px; display: flex; flex-direction: column; gap: 8px; }
.rec-priority { align-self: flex-start; font-size: 10px; text-transform: capitalize; padding: 3px 9px; border-radius: 4px; font-weight: 600; }
.rec-card h4 { font-size: 14.5px; line-height: 1.35; }
.rec-rationale { margin: 0; font-size: 12.5px; color: var(--muted); line-height: 1.55; }
.rec-action { display: flex; align-items: center; gap: 4px; font-size: 12px; color: var(--brass); margin-top: auto; padding-top: 6px; border-top: 1px solid var(--line); }


.region-cards { display:grid; grid-template-columns:repeat(2,1fr); gap:8px; }
.region-card { border:1px solid var(--line); background:var(--panel-2); border-radius:7px; padding:12px; display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; }
.region-card > strong { grid-column:1/-1; font-size:12px; color:var(--text); }
.region-card div { display:flex; flex-direction:column; gap:2px; }
.region-card div span { font-family:'IBM Plex Mono',monospace; font-size:13px; }
.region-card div small { color:var(--muted-2); font-size:9.5px; }
.network-scale { display:grid; grid-template-columns:repeat(4,1fr); gap:1px; background:var(--line); border:1px solid var(--line); border-radius:7px; overflow:hidden; }
.network-scale > div { background:var(--panel); padding:15px; display:flex; flex-direction:column; gap:5px; }
.network-scale span { color:var(--muted-2); font-size:10px; text-transform:uppercase; letter-spacing:.06em; }
.network-scale strong { font-family:'IBM Plex Mono',monospace; font-size:19px; }
.login-pill { display:flex; align-items:center; gap:6px; min-height:32px; padding:0 10px; border:1px solid var(--line); border-radius:7px; background:var(--panel); color:var(--muted); cursor:pointer; font:inherit; font-size:11.5px; max-width:220px; }
.login-pill:hover { color:var(--text); border-color:var(--brass); }
.login-pill span { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.login-pill b { color:var(--brass); font-size:9px; text-transform:uppercase; letter-spacing:.05em; }
.login-overlay { position:fixed; inset:0; background:rgba(0,0,0,.58); display:flex; align-items:center; justify-content:center; z-index:100; padding:20px; }
.login-modal { width:100%; max-width:420px; background:var(--panel); border:1px solid var(--line); border-radius:10px; box-shadow:0 22px 70px var(--shadow); padding:22px; }
.login-modal-head { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; }
.login-modal h2 { font-size:25px; margin-top:3px; }
.login-modal p { color:var(--muted); line-height:1.5; font-size:12.5px; margin:13px 0 18px; }
.login-label { display:flex; flex-direction:column; gap:7px; color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.06em; }
.login-label input { background:var(--panel-2); border:1px solid var(--line); color:var(--text); border-radius:7px; padding:10px 11px; font:inherit; font-size:13px; text-transform:none; letter-spacing:normal; }
.login-label input:focus { outline:none; border-color:var(--brass); }
.login-submit { width:100%; margin-top:14px; display:flex; justify-content:center; align-items:center; gap:6px; border:1px solid var(--brass); background:var(--brass); color:#17130A; border-radius:7px; padding:10px 12px; font-weight:600; cursor:pointer; }
.login-submit:disabled { opacity:.45; cursor:not-allowed; }
.login-security-note { display:block; color:var(--muted-2); font-size:10px; line-height:1.45; margin-top:12px; }
.gap-explainer { margin-bottom:18px; }
.gap-explainer-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; }
.gap-explainer-grid > div { border:1px solid var(--line); background:var(--panel-2); border-radius:7px; padding:11px 12px; display:flex; flex-direction:column; gap:6px; }
.gap-explainer-grid strong { font-size:12px; color:var(--text); }
.gap-explainer-grid span { font-size:11.5px; line-height:1.45; color:var(--muted); }

@media (max-width: 900px) {
  .kpi-grid { grid-template-columns: 1fr 1fr; }
  .grid-2 { grid-template-columns: 1fr; }
  .rec-grid { grid-template-columns: 1fr; }
  .board-grid { grid-template-columns: 1fr 1fr; }
  .metric-rank-row { grid-template-columns: 140px 1fr 55px; }
  .trend-hero { flex-direction:column; align-items:stretch; }
  .trend-controls { min-width:0; justify-content:flex-start; }
  .competitive-hero { align-items:stretch; }
  .competitive-scope { align-items:flex-start; min-width:0; }
  .competitive-scope-buttons { justify-content:flex-start; }
  .gap-explainer-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
.gap-meaning { min-width:280px; max-width:430px; color:var(--muted); font-size:11.5px; line-height:1.45; }
.gap-summary { align-items:flex-start; }
  .spend-row { grid-template-columns:105px 1fr 42px; }
}

@media (max-width: 560px) {
  .gap-explainer-grid { grid-template-columns:1fr; }
  .competitive-scope-buttons { display:grid; grid-template-columns:1fr; width:100%; }
}

/* Chat assistant */
.chat-fab { position: fixed; bottom: 24px; right: 24px; width: 50px; height: 50px; border-radius: 50%; background: var(--brass); color: #1A1503; border: none; display: flex; align-items: center; justify-content: center; cursor: pointer; box-shadow: 0 8px 22px var(--shadow); z-index: 40; transition: transform .15s; }
.chat-fab:hover { transform: scale(1.06); }
.chat-fab-open { background: var(--panel); color: var(--text); border: 1px solid var(--line); }
.chat-panel { position: fixed; top: 0; right: 0; height: 100vh; width: 380px; max-width: 90vw; background: var(--panel); border-left: 1px solid var(--line); display: flex; flex-direction: column; z-index: 39; box-shadow: -12px 0 32px var(--shadow); transform: translateX(100%); transition: transform .25s ease; }
.chat-panel-open { transform: translateX(0); }
.chat-head { padding: 18px 18px 12px; border-bottom: 1px solid var(--line); }
.chat-head-title { display: flex; align-items: center; gap: 7px; font-family: 'Fraunces', serif; font-size: 15.5px; color: var(--brass); }
.chat-head-sub { font-size: 11.3px; color: var(--muted-2); }
.chat-scroll { flex: 1; overflow-y: auto; padding: 16px 18px; display: flex; flex-direction: column; gap: 10px; }
.chat-empty { display: flex; flex-direction: column; gap: 12px; color: var(--muted); font-size: 13px; line-height: 1.5; padding-top: 6px; }
.chat-suggestions { display: flex; flex-direction: column; gap: 7px; }
.chat-suggestion { text-align: left; background: var(--panel-2); border: 1px solid var(--line); color: var(--text); border-radius: 7px; padding: 9px 11px; font-size: 12.5px; cursor: pointer; font-family: inherit; }
.chat-suggestion:hover { border-color: var(--brass); color: var(--brass); }
.chat-bubble-row { display: flex; }
.chat-row-user { justify-content: flex-end; }
.chat-bubble { max-width: 85%; padding: 9px 13px; border-radius: 12px; font-size: 13px; line-height: 1.5; white-space: pre-wrap; }
.chat-bubble-assistant { background: var(--panel-2); color: var(--text); border-bottom-left-radius: 3px; }
.chat-bubble-user { background: var(--brass-soft); color: var(--text); border: 1px solid var(--brass); border-bottom-right-radius: 3px; }
.chat-typing { display: flex; align-items: center; gap: 7px; color: var(--muted); }
.chat-input-row { display: flex; gap: 8px; padding: 14px 16px; border-top: 1px solid var(--line); }
.chat-input-row input { flex: 1; background: var(--panel-2); border: 1px solid var(--line); border-radius: 8px; padding: 9px 12px; color: var(--text); font-size: 13px; font-family: inherit; }
.chat-input-row input:focus { outline: none; border-color: var(--brass); }
.chat-send { width: 36px; height: 36px; flex-shrink: 0; border-radius: 8px; background: var(--brass); color: #1A1503; border: none; display: flex; align-items: center; justify-content: center; cursor: pointer; }
.chat-send:disabled { opacity: 0.4; cursor: default; }

@media (max-width: 640px) {
  .chat-panel { width: 100vw; max-width: 100vw; }
}


/* Gap / Spend / geographic additions */
/* Interactive state map */
.map-shell { display:flex; flex-direction:column; gap:14px; }
.map-toolbar { display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; }
.map-metrics { display:flex; gap:7px; flex-wrap:wrap; }
.map-reset { border:1px solid var(--line); background:transparent; color:var(--muted); border-radius:16px; padding:5px 12px; font:inherit; font-size:12px; cursor:pointer; }
.map-reset:hover { border-color:var(--brass); color:var(--brass); }
.map-layout { display:grid; grid-template-columns:minmax(0,1.85fr) minmax(270px,1fr); gap:20px; align-items:start; }
.map-stage { min-width:0; background:linear-gradient(180deg,var(--panel-2),var(--panel)); border:1px solid var(--line); border-radius:10px; padding:16px 16px 12px; }
.state-map { width:100%; height:auto; max-height:540px; display:block; }
.state-map text { font-family:'IBM Plex Sans',sans-serif; fill:var(--text); pointer-events:none; user-select:none; }
.state-map .tile-code { font-size:15px; font-weight:600; }
.state-map .tile-val { font-size:10.5px; font-weight:500; fill:var(--muted); }
.state-map .tile-star { font-size:11px; fill:var(--brass); }
.map-tile { cursor:pointer; outline:none; }
.map-tile rect { stroke:var(--line); stroke-width:1; transition:stroke .12s, stroke-width .12s, fill-opacity .12s; }
.map-tile:hover rect, .map-tile:focus rect { stroke:var(--brass); stroke-width:2; }
.map-tile.is-home rect { stroke:var(--brass); stroke-width:2; stroke-dasharray:5 3; }
.map-tile.is-active rect { stroke:var(--brass); stroke-width:3; stroke-dasharray:none; filter:drop-shadow(0 0 7px rgba(201,162,39,.45)); }
.map-tile.empty .tile-code { fill:var(--muted-2); font-weight:500; }
.theme-dark .state-map .hot .tile-code, .theme-dark .state-map .hot .tile-val { fill:#14110A; }
.theme-light .state-map .hot .tile-code, .theme-light .state-map .hot .tile-val { fill:#FFFFFF; }
.theme-dark .state-map .hot .tile-star { fill:#14110A; }
.theme-light .state-map .hot .tile-star { fill:#FFFFFF; }
.map-legend { display:flex; align-items:center; flex-wrap:wrap; gap:8px; margin-top:12px; font-size:11.5px; color:var(--muted); }
.map-legend-bar { display:block; width:140px; height:8px; border-radius:4px; background:linear-gradient(90deg, rgba(201,162,39,.12), var(--brass)); border:1px solid var(--line); }
.map-legend small { flex-basis:100%; font-size:11px; color:var(--muted-2); }
.map-detail { background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:16px 18px; min-width:0; }
.map-detail-head { margin-bottom:12px; }
.map-detail-tag { display:inline-block; font-size:10px; letter-spacing:.08em; text-transform:uppercase; color:var(--brass); margin-bottom:3px; }
.map-detail-head h4 { font-size:22px; margin:0; font-family:'Fraunces',serif; font-weight:500; }
.map-detail-grid { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
.map-detail-grid > div { background:var(--panel-2); border:1px solid var(--line); border-radius:8px; padding:9px 11px; }
.map-detail-grid > div.is-metric { border-color:var(--brass); }
.map-detail-grid span { display:block; font-size:10px; text-transform:uppercase; letter-spacing:.07em; color:var(--muted-2); margin-bottom:3px; }
.map-detail-grid strong { font-size:16px; font-weight:500; }
.map-detail-rank { font-size:12px; color:var(--muted); margin:12px 0 14px; line-height:1.5; }
.map-detail-rank b { color:var(--text); font-weight:600; }
.map-detail h5 { font-size:11px; letter-spacing:.07em; text-transform:uppercase; color:var(--muted-2); margin:0 0 6px; font-weight:500; }
.map-detail-list { display:flex; flex-direction:column; gap:4px; }
.map-detail-list button { display:flex; justify-content:space-between; gap:10px; align-items:center; width:100%; text-align:left; background:transparent; border:1px solid transparent; border-radius:7px; padding:7px 9px; color:var(--text); font:inherit; font-size:13px; cursor:pointer; }
.map-detail-list button:hover { background:var(--panel-2); border-color:var(--line); }
.map-detail-list button.is-you { border-color:var(--brass); background:var(--brass-soft); }
.map-detail-list em { font-style:normal; color:var(--muted); font-size:12px; white-space:nowrap; }
.map-detail-empty { color:var(--muted); font-size:13px; line-height:1.5; margin:4px 0 8px; }
.map-detail-hint { margin-top:14px; padding-top:10px; border-top:1px solid var(--line); font-size:11px; color:var(--muted-2); }
.demographic-profile.wide { grid-template-columns:repeat(3,1fr); }
.demographic-profile.wide > div:nth-child(even) { border-right:1px solid var(--line); }
.demographic-profile.wide > div:nth-child(3n) { border-right:0; }
@media (max-width: 980px) { .map-layout { grid-template-columns:1fr; } .demographic-profile.wide { grid-template-columns:1fr 1fr; } .demographic-profile.wide > div:nth-child(3n) { border-right:1px solid var(--line); } .demographic-profile.wide > div:nth-child(even) { border-right:0; } }
/* Excel / CSV import */
.admin-hint { font-size:11.5px; color:var(--muted-2); line-height:1.5; margin:10px 0 0; }
.file-button.is-busy { opacity:.7; pointer-events:none; }
.admin-message.info { border:1px solid var(--line); }
.import-report { margin-top:10px; padding:10px 12px; border:1px solid var(--line); border-radius:8px; background:var(--panel-2); font-size:12px; line-height:1.55; color:var(--muted); display:flex; flex-direction:column; gap:3px; }
.import-report b { color:var(--text); font-weight:600; }
.import-missing { color:var(--muted-2); }
.mapping-panel { margin-top:12px; padding:14px; border:1px solid var(--brass); border-radius:9px; background:var(--panel-2); }
.mapping-head { display:flex; justify-content:space-between; gap:8px; flex-wrap:wrap; margin-bottom:10px; font-size:12px; color:var(--muted); }
.mapping-head strong { color:var(--text); font-size:13px; }
.mapping-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(210px,1fr)); gap:9px; max-height:330px; overflow:auto; padding-right:4px; }
.mapping-grid label { display:flex; flex-direction:column; gap:4px; font-size:11px; color:var(--muted); }
.mapping-grid select { width:100%; background:var(--panel); border:1px solid var(--line); color:var(--text); border-radius:7px; padding:7px 8px; font:inherit; font-size:12px; }
.mapping-grid select.is-required-missing { border-color:#D65A4A; }
.gap-summary { display:flex; flex-direction:column; align-items:flex-end; gap:4px; padding:10px 13px; border:1px solid var(--line); background:var(--panel); border-radius:8px; }
.gap-summary span { font-size:10px; color:var(--muted-2); text-transform:uppercase; letter-spacing:.07em; }
.gap-summary strong { font-size:14px; color:var(--brass); }
.gap-good { color:var(--good)!important; }
.gap-bad { color:var(--risk)!important; }
.gap-badge { display:inline-flex; padding:3px 7px; border-radius:999px; font-size:10px; }
.gap-badge.good { color:var(--good); background:rgba(76,154,106,.12); }
.gap-badge.bad { color:var(--risk); background:rgba(214,90,74,.12); }
.gap-list { display:flex; flex-direction:column; }
.gap-item { display:grid; grid-template-columns:minmax(0,1fr) 82px; align-items:start; gap:12px; padding:11px 4px; border-bottom:1px solid var(--line); }
.gap-item:last-child { border-bottom:0; }
.gap-item > div { display:flex; flex-direction:column; gap:3px; min-width:0; }
.gap-item > b { justify-self:end; align-self:start; padding-top:1px; text-align:right; white-space:nowrap; }
.gap-item span { color:var(--muted); font-size:11px; }
.gap-item b { font-family:'IBM Plex Mono',monospace; color:var(--risk); }

/* Connect Admin / multi-quarter additions */
.admin-header { display:flex; justify-content:space-between; gap:20px; align-items:flex-start; margin-bottom:18px; }
.admin-header h2 { font-size:26px; margin:8px 0 5px; }
.admin-header p { color:var(--muted); max-width:760px; margin:0; line-height:1.55; }
.admin-quarter-pill { display:flex; align-items:center; gap:7px; padding:9px 12px; border:1px solid var(--line); background:var(--panel); border-radius:8px; white-space:nowrap; color:var(--muted); }
.admin-quarter-pill strong { color:var(--text); }
.admin-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-bottom:14px; }
.admin-panel { min-width:0; }
.quarter-list { display:flex; flex-direction:column; gap:7px; margin-bottom:14px; }
.quarter-row { display:flex; justify-content:space-between; align-items:center; padding:10px 11px; border:1px solid var(--line); border-radius:8px; background:var(--panel-2); }
.quarter-row > div { display:flex; flex-direction:column; gap:3px; }
.quarter-row span { font-size:11px; color:var(--muted); }
.quarter-active { border-color:var(--brass); }
.field-grid { display:grid; grid-template-columns:1fr; gap:9px; }
.field-grid.two { grid-template-columns:1fr 1fr; }
.field-grid label { display:flex; flex-direction:column; gap:5px; }
.field-grid label span { font-size:11px; color:var(--muted); }
.field-grid input, .field-grid select { width:100%; background:var(--panel-2); border:1px solid var(--line); color:var(--text); border-radius:7px; padding:9px 10px; font:inherit; }
.field-grid input:focus { outline:none; border-color:var(--brass); }
.admin-actions { display:flex; align-items:center; flex-wrap:wrap; gap:8px; margin-top:12px; }
.file-button { position:relative; overflow:hidden; }
.file-button input { position:absolute; inset:0; opacity:0; cursor:pointer; }
.btn-secondary { display:inline-flex; align-items:center; gap:6px; border:1px solid var(--line); background:var(--panel-2); color:var(--text); border-radius:7px; padding:8px 11px; cursor:pointer; font:inherit; }
.btn-secondary:hover { border-color:var(--brass); color:var(--brass); }
.connection-mode-row { display:flex; gap:7px; flex-wrap:wrap; margin-bottom:12px; }
.security-note { display:flex; align-items:center; gap:5px; font-size:11px; color:var(--muted); max-width:620px; line-height:1.4; }
.admin-message { display:flex; align-items:flex-start; gap:7px; margin-top:10px; padding:9px 10px; border-radius:7px; background:var(--brass-soft); color:var(--text); font-size:12px; line-height:1.45; }
.admin-message.success { border:1px solid rgba(76,154,106,.35); }
.admin-message.error { border:1px solid rgba(214,90,74,.35); }
.pipeline { display:grid; grid-template-columns:1fr auto 1fr auto 1fr auto 1fr; align-items:center; gap:8px; }
.pipeline > div { min-height:88px; border:1px solid var(--line); border-radius:8px; background:var(--panel-2); display:flex; flex-direction:column; justify-content:center; align-items:center; gap:5px; text-align:center; }
.pipeline > div svg { color:var(--brass); }
.pipeline span { font-size:10px; color:var(--muted); }
.refresh-row { display:flex; align-items:center; gap:11px; padding:10px 0; border-bottom:1px solid var(--line); }
.refresh-row:last-child { border-bottom:0; }
.refresh-row > svg { color:var(--brass); flex-shrink:0; }
.refresh-row div { flex:1; display:flex; flex-direction:column; gap:3px; }
.refresh-row span { color:var(--muted); font-size:11px; line-height:1.4; }
.refresh-row select { background:var(--panel-2); color:var(--text); border:1px solid var(--line); border-radius:7px; padding:7px 9px; }
.status-badge { padding:5px 8px; border-radius:999px; background:var(--brass-soft); color:var(--brass)!important; }
.quarter-selector { display:flex; align-items:center; gap:6px; border:1px solid var(--line); background:var(--panel); border-radius:7px; padding:6px 9px; color:var(--brass); }
.quarter-selector select { background:transparent; border:0; color:var(--text); font:inherit; outline:none; cursor:pointer; }
.quarter-selector option { background:var(--panel); color:var(--text); }
.eyebrow-tag { display:inline-flex; align-items:center; gap:5px; }

/* Trend / demographic views */
.trend-hero { display:flex; justify-content:space-between; align-items:flex-end; gap:24px; margin-bottom:22px; }
.trend-hero h1 { font-size:30px; margin-top:5px; }
.trend-hero p { color:var(--muted); max-width:760px; line-height:1.55; margin:7px 0 0; }
.trend-controls { display:flex; gap:10px; flex-wrap:wrap; min-width:360px; justify-content:flex-end; }
.trend-controls label { display:flex; flex-direction:column; gap:6px; color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.06em; }
.trend-controls select { min-width:175px; background:var(--panel); color:var(--text); border:1px solid var(--line); border-radius:6px; padding:9px 11px; font:inherit; text-transform:none; letter-spacing:0; }
.trend-empty { min-height:290px; display:flex; align-items:center; justify-content:center; flex-direction:column; gap:9px; color:var(--muted); text-align:center; padding:30px; }
.trend-empty strong { color:var(--text); font-size:14px; }
.trend-empty svg { color:var(--brass); }
.demographic-profile { display:grid; grid-template-columns:1fr 1fr; gap:0; }
.demographic-profile > div { padding:13px 14px; border-bottom:1px solid var(--line); border-right:1px solid var(--line); }
.demographic-profile > div:nth-child(even) { border-right:0; }
.demographic-profile span { display:block; color:var(--muted-2); font-size:10px; text-transform:uppercase; letter-spacing:.07em; margin-bottom:4px; }
.demographic-profile strong { font-size:14px; font-weight:500; }
.demographic-note { display:flex; gap:11px; color:var(--muted); line-height:1.55; }
.demographic-note svg { color:var(--brass); flex-shrink:0; margin-top:2px; }
.demographic-note strong, .demographic-note span { display:block; }
.demographic-note strong { color:var(--text); margin-bottom:3px; }


.region-cards { display:grid; grid-template-columns:repeat(2,1fr); gap:8px; }
.region-card { border:1px solid var(--line); background:var(--panel-2); border-radius:7px; padding:12px; display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; }
.region-card > strong { grid-column:1/-1; font-size:12px; color:var(--text); }
.region-card div { display:flex; flex-direction:column; gap:2px; }
.region-card div span { font-family:'IBM Plex Mono',monospace; font-size:13px; }
.region-card div small { color:var(--muted-2); font-size:9.5px; }
.network-scale { display:grid; grid-template-columns:repeat(4,1fr); gap:1px; background:var(--line); border:1px solid var(--line); border-radius:7px; overflow:hidden; }
.network-scale > div { background:var(--panel); padding:15px; display:flex; flex-direction:column; gap:5px; }
.network-scale span { color:var(--muted-2); font-size:10px; text-transform:uppercase; letter-spacing:.06em; }
.network-scale strong { font-family:'IBM Plex Mono',monospace; font-size:19px; }
.login-pill { display:flex; align-items:center; gap:6px; min-height:32px; padding:0 10px; border:1px solid var(--line); border-radius:7px; background:var(--panel); color:var(--muted); cursor:pointer; font:inherit; font-size:11.5px; max-width:220px; }
.login-pill:hover { color:var(--text); border-color:var(--brass); }
.login-pill span { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.login-pill b { color:var(--brass); font-size:9px; text-transform:uppercase; letter-spacing:.05em; }
.login-overlay { position:fixed; inset:0; background:rgba(0,0,0,.58); display:flex; align-items:center; justify-content:center; z-index:100; padding:20px; }
.login-modal { width:100%; max-width:420px; background:var(--panel); border:1px solid var(--line); border-radius:10px; box-shadow:0 22px 70px var(--shadow); padding:22px; }
.login-modal-head { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; }
.login-modal h2 { font-size:25px; margin-top:3px; }
.login-modal p { color:var(--muted); line-height:1.5; font-size:12.5px; margin:13px 0 18px; }
.login-label { display:flex; flex-direction:column; gap:7px; color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.06em; }
.login-label input { background:var(--panel-2); border:1px solid var(--line); color:var(--text); border-radius:7px; padding:10px 11px; font:inherit; font-size:13px; text-transform:none; letter-spacing:normal; }
.login-label input:focus { outline:none; border-color:var(--brass); }
.login-submit { width:100%; margin-top:14px; display:flex; justify-content:center; align-items:center; gap:6px; border:1px solid var(--brass); background:var(--brass); color:#17130A; border-radius:7px; padding:10px 12px; font-weight:600; cursor:pointer; }
.login-submit:disabled { opacity:.45; cursor:not-allowed; }
.login-security-note { display:block; color:var(--muted-2); font-size:10px; line-height:1.45; margin-top:12px; }

@media (max-width: 900px) {
  .admin-grid { grid-template-columns:1fr; }
  .field-grid.two { grid-template-columns:1fr; }
  .pipeline { grid-template-columns:1fr; }
  .pipeline > svg { transform:rotate(90deg); justify-self:center; }
  .admin-header { flex-direction:column; }
}
`;