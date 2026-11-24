
/* Examples used with the sample data only. List 1: which rack each meter belongs to. List 2: which company each rack belongs to (the two company names are placeholders). */
const DEFAULT_MAPPING = [
  "Meter name,Rack name",
  "PDU-S1-FN4-B01,IBM", "PDU-S2-FN4-B01,IBM",
  "PDU-S1-FN4-B12,NR1", "PDU-S1-FN4-R01,NR1", "PDU-S2-FN4-B12,NR1", "PDU-S2-FN4-R01,NR1",
  "PDU-S1-FN4-Y01,QUANTUM", "PDU-S2-FN4-Y01,QUANTUM",
  "PDU-S1-FN4-B02,R5", "PDU-S2-FN4-B02,R5",
  "PDU-S1-FN4-R02,VTL", "PDU-S1-FN4-Y02,VTL", "PDU-S2-FN4-R02,VTL", "PDU-S2-FN4-Y02,VTL"
].join("\n");
const DEFAULT_RACK_COMPANY = [
  "Rack name,Company",
  "IBM,Example Company A", "NR1,Example Company A",
  "QUANTUM,Example Company B", "R5,Example Company B", "VTL,Example Company B"
].join("\n");

const SAMPLE_METERS = [
  ["PDU-S1-FN4-B01", 137.21, 33513.99], ["PDU-S2-FN4-B01", 113.46, 33549.61],
  ["PDU-S1-FN4-B12", 138.99, 33661.28], ["PDU-S1-FN4-R01", 105.55, 33539.07],
  ["PDU-S2-FN4-B12", 121.15, 33575.00], ["PDU-S2-FN4-R01", 124.78, 33644.30],
  ["PDU-S1-FN4-Y01", 118.36, 33747.23], ["PDU-S2-FN4-Y01", 145.07, 33686.50],
  ["PDU-S1-FN4-B02", 123.26, 33552.15], ["PDU-S2-FN4-B02", 132.34, 33624.24],
  ["PDU-S1-FN4-R02", 133.14, 33719.79], ["PDU-S1-FN4-Y02", 149.83, 33710.80],
  ["PDU-S2-FN4-R02", 119.32, 33608.02], ["PDU-S2-FN4-Y02", 124.88, 33449.46]
];

function makeSampleCsv() {
  const steps = 6;
  const lines = ["Metric Name,Time,Value (kWh)"];
  for (let day = 0; day < steps; day++) {
    const d = new Date(2026, 9, 1 + day * 6, 0, 0, 0);
    const ts = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") + " 00:00";
    for (const [name, a, b] of SAMPLE_METERS) {
      const v = day === steps - 1 ? b : Math.round((a + (b - a) * day / (steps - 1)) * 100) / 100;
      lines.push(name + "," + ts + "," + v.toFixed(2));
    }
  }
  return lines.join("\n");
}

/* ---------- reading text ---------- */
function decodeBuffer(buf) {
  const u = new Uint8Array(buf);
  if (u.length >= 2 && u[0] === 0xFF && u[1] === 0xFE) return new TextDecoder("utf-16le").decode(u.subarray(2));
  if (u.length >= 2 && u[0] === 0xFE && u[1] === 0xFF) return new TextDecoder("utf-16be").decode(u.subarray(2));
  return new TextDecoder("utf-8").decode(u).replace(/^﻿/, "");
}

function detectDelimiter(text) {
  const lines = text.split(/\r\n|\n|\r/).filter(l => l.trim() !== "").slice(0, 15);
  let best = ",", bestScore = -1;
  for (const d of [",", ";", "\t", "|"]) {
    const counts = lines.map(l => l.split(d).length - 1);
    const nonzero = counts.filter(c => c > 0);
    if (!nonzero.length) continue;
    const mode = nonzero.sort((a, b) => a - b)[Math.floor(nonzero.length / 2)];
    const score = nonzero.length * 100 + mode;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

function parseCSV(text, delim) {
  const rows = [];
  let row = [], field = "", inQ = false;
  const n = text.length;
  for (let i = 0; i < n; i++) {
    const c = text.charCodeAt(i);
    if (inQ) {
      if (c === 34) {
        if (text.charCodeAt(i + 1) === 34) { field += '"'; i++; } else inQ = false;
      } else field += text[i];
    } else if (c === 34 && field === "") {
      inQ = true;
    } else if (text[i] === delim) {
      row.push(field); field = "";
    } else if (c === 10 || c === 13) {
      if (c === 13 && text.charCodeAt(i + 1) === 10) i++;
      row.push(field); rows.push(row); row = []; field = "";
    } else field += text[i];
  }
  if (field !== "" || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/* Excel: turn one sheet into an array of rows. XL is the SheetJS library. */
function fmtLocal(d) {
  const x = new Date(Math.round(d.getTime() / 1000) * 1000);
  const p = n => String(n).padStart(2, "0");
  return x.getFullYear() + "-" + p(x.getMonth() + 1) + "-" + p(x.getDate()) + " " + p(x.getHours()) + ":" + p(x.getMinutes()) + ":" + p(x.getSeconds());
}
/* Excel dates are read as plain serial numbers and turned into text by hand. The reader's own Date objects drift by seconds in some time zones (Malaysia included), which would move a 12:00 a.m. reading onto the previous date. */
function fmtSerial(XL, v) {
  const d = XL.SSF.parse_date_code(v);
  if (!d) return "";
  const p = n => String(n).padStart(2, "0");
  return d.y + "-" + p(d.m) + "-" + p(d.d) + " " + p(d.H) + ":" + p(d.M) + ":" + p(Math.round(d.S));
}
function sheetRows(XL, wb, name) {
  const ws = wb.Sheets[name];
  if (!ws || !ws["!ref"]) return [];
  const rg = XL.utils.decode_range(ws["!ref"]);
  const out = [];
  for (let R = rg.s.r; R <= rg.e.r; R++) {
    const row = [];
    let any = false;
    for (let C = rg.s.c; C <= rg.e.c; C++) {
      const cell = ws[XL.utils.encode_cell({ r: R, c: C })];
      let v = "";
      if (cell && cell.v !== undefined && cell.v !== null) {
        v = cell.v;
        if (v instanceof Date) v = fmtLocal(v);
        else if (cell.t === "n" && cell.z && XL.SSF.is_date(cell.z)) v = fmtSerial(XL, v);
        if (v !== "") any = true;
      }
      row.push(v);
    }
    if (any) out.push(row);
  }
  return out;
}
function bestSheet(XL, wb, skip) {
  let best = wb.SheetNames.find(n => !skip || !skip.has(n)) || wb.SheetNames[0], bestN = -1;
  for (const n of wb.SheetNames) {
    if (skip && skip.has(n)) continue;
    const ref = wb.Sheets[n]["!ref"];
    if (!ref) continue;
    const r = XL.utils.decode_range(ref);
    const size = (r.e.r - r.s.r + 1) * Math.min(r.e.c - r.s.c + 1, 50);
    if (size > bestN) { bestN = size; best = n; }
  }
  return best;
}

/* ---------- numbers and dates ---------- */
const HEADER_WORDS = /metric|name|device|meter|pdu|time|date|stamp|value|kwh|reading|energy|rack|point|sensor|usage|last|current|previous|start|end/i;
const PURE_NUMBER = /^[-+]?\d+([.,]\d+)?$/;

function parseNumber(s) {
  if (s === null || s === undefined) return NaN;
  if (typeof s === "number") return Number.isFinite(s) ? s : NaN;
  let t = String(s).trim();
  if (t === "") return NaN;
  /* allow a unit or bracket around the number, but nothing that looks like a date or a code */
  t = t.replace(/^[^\d\-+.,]+/, "").replace(/[^\d]+$/, "").replace(/[\s']/g, "");
  if (!/^[-+]?[\d.,]+([eE][-+]?\d+)?$/.test(t)) return NaN;
  const hasC = t.includes(","), hasD = t.includes(".");
  if (hasC && hasD) {
    if (t.lastIndexOf(",") > t.lastIndexOf(".")) t = t.replace(/\./g, "").replace(",", ".");
    else t = t.replace(/,/g, "");
  } else if (hasC) {
    if (/^[-+]?\d{1,3}(,\d{3})+$/.test(t)) t = t.replace(/,/g, "");
    else t = t.replace(",", ".");
  }
  const v = Number(t);
  return Number.isFinite(v) ? v : NaN;
}

const DMY = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?$/i;

function parseTime(s, order) {
  if (s === null || s === undefined) return NaN;
  if (s instanceof Date) return s.getTime();
  const t = String(s).trim();
  if (!t) return NaN;
  if (/^\d+(\.\d+)?$/.test(t)) {
    const n = Number(t);
    if (n > 30000 && n < 80000) return Math.round((n - 25569) * 86400000);
    if (n > 1e9 && n < 1e11) return n * 1000;
    if (n >= 1e11 && n < 1e14) return n;
    return NaN;
  }
  if (/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(t)) {
    const v = Date.parse(t.replace(/\//g, "-").replace(" ", "T"));
    return Number.isNaN(v) ? NaN : v;
  }
  const m = t.match(DMY);
  if (m) {
    const a = +m[1], b = +m[2];
    let y = +m[3];
    if (y < 100) y += 2000;
    let day, mon;
    if (order === "mdy") { mon = a; day = b; } else { day = a; mon = b; }
    if (mon > 12 && day <= 12) { const x = day; day = mon; mon = x; }
    let h = m[4] ? +m[4] : 0;
    const mi = m[5] ? +m[5] : 0, sec = m[6] ? +m[6] : 0;
    if (m[7]) { const pm = m[7].toUpperCase() === "PM"; if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
    const d = new Date(y, mon - 1, day, h, mi, sec);
    return Number.isNaN(d.getTime()) ? NaN : d.getTime();
  }
  const v = Date.parse(t);
  return Number.isNaN(v) ? NaN : v;
}

function isDateText(c, order) {
  if (c === null || c === undefined || typeof c === "number") return false;
  const s = String(c).trim();
  if (!s || PURE_NUMBER.test(s)) return false;
  return !Number.isNaN(parseTime(s, order));
}

/* Day-first or month-first? Look for a date whose first or second part is above 12. */
function detectDateOrder(headers, rows) {
  let a12 = false, b12 = false, seen = 0;
  const check = c => {
    const m = DMY.exec(String(c === undefined || c === null ? "" : c).trim());
    if (!m) return;
    seen++;
    if (+m[1] > 12) a12 = true;
    if (+m[2] > 12) b12 = true;
  };
  for (const h of headers) check(h);
  const ncols = Math.min(headers.length, 6);
  for (let i = 0; i < rows.length && !(a12 || b12); i++) {
    const r = rows[i];
    for (let j = 0; j < ncols; j++) check(r[j]);
  }
  if (a12 && !b12) return { order: "dmy", sure: true, seen };
  if (b12 && !a12) return { order: "mdy", sure: true, seen };
  return { order: "dmy", sure: seen === 0, seen };
}

function detectUnit(header) {
  const h = String(header || "");
  if (/MWh/.test(h)) return 1000;
  if (/(^|[^kKmM])Wh\b/.test(h) && !/kWh/i.test(h)) return 0.001;
  return 1;
}

/* ---------- finding the header and the layout ---------- */
function findHeader(rows) {
  let best = 0, bestScore = 0;
  const lim = Math.min(rows.length, 30);
  for (let i = 0; i < lim; i++) {
    const cells = rows[i].map(c => String(c === undefined || c === null ? "" : c).trim()).filter(Boolean);
    if (cells.length < 2) continue;
    const score = cells.filter(c => HEADER_WORDS.test(c) && Number.isNaN(parseNumber(c))).length;
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}

function sampleRows(rows, n) {
  if (rows.length <= n) return rows;
  const out = [], step = rows.length / n;
  for (let i = 0; i < n; i++) out.push(rows[Math.floor(i * step)]);
  return out;
}

function classify(H, rows, order) {
  const samp = sampleRows(rows, 300);
  return H.map((_, j) => {
    let n = 0, num = 0, tm = 0;
    for (const r of samp) {
      const c = r[j];
      if (c === undefined || c === null || String(c).trim() === "") continue;
      n++;
      if (!Number.isNaN(parseNumber(c))) num++;
      if (isDateText(c, order)) tm++;
    }
    return { n, num: n ? num / n : 0, time: n ? tm / n : 0 };
  });
}

const METER_HEADER = /^(metric|meter|device|point|sensor|channel|pdu)( name| id| no\.?)?$|^(name|id)$/i;

function detectLayout(headers, rows, order) {
  const H = headers.map(h => String(h === undefined || h === null ? "" : h).trim());
  const cls = classify(H, rows, order);
  const isTime = j => cls[j].n > 0 && cls[j].time > 0.9 && cls[j].num < 0.5;
  const isNum = j => cls[j].n > 0 && cls[j].num > 0.9;
  const T = [], N = [], S = [];
  H.forEach((h, j) => {
    if (isTime(j)) T.push(j);
    else if (isNum(j)) N.push(j);
    else if (cls[j].n > 0) S.push(j);
  });
  const dateHeaders = H.filter((h, j) => j > 0 && isDateText(h, order)).length;
  if (H.length >= 4 && dateHeaders >= Math.max(3, (H.length - 1) * 0.6)) return "transposed";
  const lastCol = N.find(j => /last|previous|prev|start|begin|opening|initial|old|from/i.test(H[j]));
  const curCol = N.find(j => /current|latest|end|closing|final|new|to$/i.test(H[j]));
  if (lastCol !== undefined && curCol !== undefined && lastCol !== curCol && T.length === 0) return "summary";
  const meterHeader = H.some(h => METER_HEADER.test(h));
  if (S.length === 0 && N.length >= 1 && T.length >= 1 && !meterHeader) return "wide";
  return "long";
}

function guessColumns(headers, rows) {
  const H = headers.map(h => String(h === undefined || h === null ? "" : h).trim());
  const used = new Set();
  const pick = (...res) => {
    for (const re of res) {
      const i = H.findIndex((h, k) => !used.has(k) && re.test(h));
      if (i >= 0) { used.add(i); return i; }
    }
    return -1;
  };
  const sample = rows.slice(0, 200);
  const ratio = (idx, fn) => {
    if (idx < 0 || !sample.length) return 0;
    let ok = 0, seen = 0;
    for (const r of sample) { const c = r[idx]; if (c === undefined || String(c).trim() === "") continue; seen++; if (fn(c)) ok++; }
    return seen ? ok / seen : 0;
  };
  const rack = pick(/rack/i);
  const company = pick(/company|client|customer|tenant/i);
  let time = pick(/time|date|stamp/i);
  let value = pick(/kwh/i, /value|reading|energy|consumption|usage|total/i);
  let metric = pick(/metric/i, /pdu|meter|device|point|sensor|channel|name|^id$/i);
  if (time < 0) {
    for (let i = 0; i < H.length; i++) if (!used.has(i) && ratio(i, c => isDateText(c)) > 0.9) { time = i; used.add(i); break; }
  }
  if (value < 0) {
    for (let i = 0; i < H.length; i++) if (!used.has(i) && ratio(i, c => !Number.isNaN(parseNumber(c))) > 0.9) { value = i; used.add(i); break; }
  }
  if (metric < 0) {
    for (let i = 0; i < H.length; i++) if (!used.has(i) && ratio(i, c => Number.isNaN(parseNumber(c))) > 0.9) { metric = i; used.add(i); break; }
  }
  return { metric, time, value, rack, company };
}

function columnsFor(layout, headers, rows, order) {
  const H = headers.map(h => String(h === undefined || h === null ? "" : h).trim());
  const base = { metric: -1, time: -1, value: -1, rack: -1, company: -1, last: -1, current: -1, meterCols: [], timeCols: [] };
  if (layout === "long") return Object.assign(base, guessColumns(headers, rows));
  const cls = classify(H, rows, order);
  const isTime = j => cls[j].n > 0 && cls[j].time > 0.9 && cls[j].num < 0.5;
  const isNum = j => cls[j].n > 0 && cls[j].num > 0.9;
  const rackIdx = H.findIndex(h => /rack/i.test(h));
  const coIdx = H.findIndex(h => /company|client|customer|tenant/i.test(h));
  if (layout === "wide") {
    let t = H.findIndex((h, j) => isTime(j));
    if (t < 0) t = H.findIndex(h => /time|date|stamp/i.test(h));
    if (t < 0) t = 0;
    base.time = t;
    base.meterCols = H.map((h, j) => j).filter(j => j !== t && H[j] && isNum(j));
    return base;
  }
  if (layout === "transposed") {
    base.timeCols = H.map((h, j) => j).filter(j => H[j] && isDateText(H[j], order));
    const isMeterish = j => !base.timeCols.includes(j) && j !== rackIdx && j !== coIdx && !isNum(j);
    let m = H.findIndex((h, j) => isMeterish(j) && METER_HEADER.test(h));
    if (m < 0) m = H.findIndex((h, j) => isMeterish(j));
    if (m < 0) m = 0;
    base.metric = m;
    base.rack = rackIdx >= 0 && !base.timeCols.includes(rackIdx) ? rackIdx : -1;
    base.company = coIdx >= 0 && !base.timeCols.includes(coIdx) ? coIdx : -1;
    return base;
  }
  if (layout === "summary") {
    const nums = H.map((h, j) => j).filter(j => isNum(j));
    base.last = nums.find(j => /last|previous|prev|start|begin|opening|initial|old|from/i.test(H[j]));
    if (base.last === undefined) base.last = -1;
    base.current = nums.find(j => j !== base.last && /current|latest|end|closing|final|new|to$/i.test(H[j]));
    if (base.current === undefined) base.current = -1;
    let m = H.findIndex((h, j) => j !== rackIdx && j !== coIdx && !isNum(j) && METER_HEADER.test(h));
    if (m < 0) m = H.findIndex((h, j) => j !== rackIdx && j !== coIdx && !isNum(j));
    base.metric = m < 0 ? 0 : m;
    base.rack = rackIdx;
    base.company = coIdx;
    return base;
  }
  return base;
}

/* ---------- mapping and calculation ---------- */
/*
  Mapping text: one meter per line, as  meter name, rack name, company.
  - the rack and the company can each be left empty (PDU-1,,Acme gives a company but no rack)
  - an optional first line with headers (Meter name, Rack name, Company) lets the columns come in any order
  - the meter name may use * (any characters) or ? (one character), so one line can cover a whole rack: PDU-S1-FN2-R27*
*/
const HDR_METER = /^(meter|metric|pdu|device|point|sensor|channel)( name| id)?$|^name$/i;
const HDR_RACK = /^rack( name| no\.?| number)?$/i;
const HDR_COMPANY = /^(company|client|customer|tenant)( name)?$/i;

function parseMappingRows(text) {
  const t = String(text === undefined || text === null ? "" : text);
  if (!t.trim()) return [];
  const raw = parseCSV(t, detectDelimiter(t)).map(r => r.map(c => String(c).trim())).filter(r => r.some(c => c !== ""));
  if (!raw.length) return [];
  let im = 0, ir = 1, ic = 2, start = 0;
  const first = raw[0];
  const hm = first.findIndex(c => HDR_METER.test(c)), hr = first.findIndex(c => HDR_RACK.test(c)), hc = first.findIndex(c => HDR_COMPANY.test(c));
  if (hm >= 0 && (hr >= 0 || hc >= 0)) { im = hm; ir = hr; ic = hc; start = 1; }
  const out = [];
  for (let i = start; i < raw.length; i++) {
    const r = raw[i];
    const meter = r[im] || "";
    if (!meter) continue;
    out.push({ meter, rack: ir >= 0 ? (r[ir] || "") : "", company: ic >= 0 ? (r[ic] || "") : "" });
  }
  return out;
}

function mappingToText(rows) {
  const q = s => /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  const withCompany = rows.some(r => r.company);
  return rows.map(r => (withCompany ? [r.meter, r.rack, r.company] : [r.meter, r.rack]).map(q).join(",")).join("\n");
}

function wildcardRe(pattern) {
  const body = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
  return new RegExp("^" + body + "$", "i");
}

/* lookup(name) gives { rack, company } or null. Exact names win over patterns; a longer pattern wins over a shorter one. */
function buildMapper(text) {
  const exact = new Map(), wild = [], companies = [], seen = new Map();
  let ord = 0;
  for (const r of parseMappingRows(text)) {
    let company = "";
    if (r.company) {
      const k = r.company.toLowerCase();
      if (!seen.has(k)) { seen.set(k, r.company); companies.push(r.company); }
      company = seen.get(k);
    }
    const entry = { rack: r.rack, company, ord: ord++ };
    if (/[*?]/.test(r.meter)) wild.push({ re: wildcardRe(r.meter), lit: r.meter.replace(/[*?]/g, "").length, entry });
    else {
      const k = r.meter.toLowerCase();
      if (!exact.has(k)) exact.set(k, []);
      exact.get(k).push(entry);
    }
  }
  wild.sort((a, b) => b.lit - a.lit);
  /* all(name) gives every line that names the meter: a meter listed under two racks or two companies is counted in each, as in the Excel template */
  const all = name => {
    const hit = exact.get(String(name).toLowerCase());
    if (hit) return hit;
    for (const w of wild) if (w.re.test(name)) return [w.entry];
    return [];
  };
  const lookup = name => all(name)[0] || null;
  return { lookup, all, companies, size: exact.size + wild.length };
}

/*
  Rack list: one rack per line as  rack name, company. A rack name may use * or ? (R2* covers R20, R21 ...).
  An optional first line with headers (Rack name, Company) lets the columns come in any order.
*/
function parseRackCompanyRows(text) {
  const t = String(text === undefined || text === null ? "" : text);
  if (!t.trim()) return [];
  const raw = parseCSV(t, detectDelimiter(t)).map(r => r.map(c => String(c).trim())).filter(r => r.some(c => c !== ""));
  if (!raw.length) return [];
  let ir = 0, ic = 1, start = 0;
  const hr = raw[0].findIndex(c => HDR_RACK.test(c)), hc = raw[0].findIndex(c => HDR_COMPANY.test(c));
  if (hr >= 0 && hc >= 0 && hr !== hc) { ir = hr; ic = hc; start = 1; }
  const out = [];
  for (let i = start; i < raw.length; i++) {
    const rack = raw[i][ir] || "", company = raw[i][ic] || "";
    if (rack && company) out.push({ rack, company });
  }
  return out;
}
function rackCompanyToText(rows) {
  const q = s => /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  return rows.map(r => [r.rack, r.company].map(q).join(",")).join("\n");
}
function buildRackCompany(text) {
  const exact = new Map(), wild = [], companies = [], seen = new Map();
  for (const r of parseRackCompanyRows(text)) {
    const k = r.company.toLowerCase();
    if (!seen.has(k)) { seen.set(k, r.company); companies.push(r.company); }
    const company = seen.get(k);
    if (/[*?]/.test(r.rack)) wild.push({ re: wildcardRe(r.rack), lit: r.rack.replace(/[*?]/g, "").length, company });
    else exact.set(r.rack.toLowerCase(), company);
  }
  wild.sort((a, b) => b.lit - a.lit);
  const lookup = rack => {
    const hit = exact.get(String(rack).toLowerCase());
    if (hit) return hit;
    for (const w of wild) if (w.re.test(rack)) return w.company;
    return "";
  };
  return { lookup, companies };
}

/* every company named in either list, in order of first appearance */
function listCompanies(mappingText, rackCompanyText) {
  const seen = new Map();
  for (const c of buildMapper(mappingText).companies.concat(buildRackCompany(rackCompanyText).companies)) if (!seen.has(c.toLowerCase())) seen.set(c.toLowerCase(), c);
  return [...seen.values()];
}

/* kept for older callers: meter (lower case) -> rack, exact names only */
function parseMapping(text) {
  const map = new Map();
  for (const r of parseMappingRows(text)) if (r.rack && !/[*?]/.test(r.meter)) map.set(r.meter.toLowerCase(), r.rack);
  return map;
}

/* ---------- finding companies and racks inside the meter names ---------- */
/*
  Meter names often carry a code: PDU-S1-FN2-R27 = feed S1, zone FN2, rack code R27.
  analyseNames cuts every name into parts at the separator that splits most names the same way,
  and lists the values each part takes, so a part can be chosen to stand for the company and for the rack.
*/
const NAME_SEPARATORS = ["-", "_", ".", "/", ":", " "];
function analyseNames(names) {
  const list = [...new Set((names || []).map(n => String(n).trim()).filter(Boolean))];
  let best = null;
  for (const sep of NAME_SEPARATORS) {
    const counts = new Map();
    for (const n of list) { const k = n.split(sep).length; counts.set(k, (counts.get(k) || 0) + 1); }
    let nParts = 0, regular = 0;
    for (const [k, v] of counts) if (k >= 2 && v > regular) { nParts = k; regular = v; }
    if (regular && (!best || regular > best.regular || (regular === best.regular && nParts > best.nParts))) best = { sep, nParts, regular };
  }
  if (!best || best.regular < 2) return { ok: false, names: list.length, sep: "", parts: [], regular: 0 };
  const names2 = list.filter(n => n.split(best.sep).length === best.nParts);
  const collator = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  const parts = [];
  for (let i = 0; i < best.nParts; i++) {
    const m = new Map();
    for (const n of names2) { const v = n.split(best.sep)[i]; m.set(v, (m.get(v) || 0) + 1); }
    const values = [...m].map(([value, count]) => ({ value, count })).sort((a, b) => collator.compare(a.value, b.value));
    parts.push({ index: i, distinct: values.length, values });
  }
  return { ok: true, sep: best.sep, nParts: best.nParts, names: list.length, regular: names2.length, parts };
}

/*
  companyParts and rackParts are lists of part positions (0 = first). Every meter gets the company and the rack made of
  those parts joined together. The result is one mapping line per distinct combination, as a pattern: the parts that are not
  used become *, so meters added later with the same code are picked up too.
  labels maps a company code (for example "FN2") to the name to show on its button.
*/
function buildNameMapping(names, sep, nParts, companyParts, rackParts, labels) {
  const used = new Set([...companyParts, ...rackParts]);
  const seen = new Map();
  /* a part that never changes (the leading PDU, say) stays as written instead of becoming * */
  const firstParts = new Map();
  for (const raw of names || []) {
    const p = String(raw).trim().split(sep);
    if (p.length !== nParts) continue;
    p.forEach((x, i) => { if (!firstParts.has(i)) firstParts.set(i, new Set()); firstParts.get(i).add(x); });
  }
  for (let i = 0; i < nParts; i++) if (firstParts.has(i) && firstParts.get(i).size === 1) used.add(i);
  const collator = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  for (const raw of names || []) {
    const n = String(raw).trim();
    const p = n.split(sep);
    if (p.length !== nParts) continue;
    const pattern = p.map((x, i) => used.has(i) ? x : "*").join(sep);
    if (seen.has(pattern)) continue;
    const code = companyParts.map(i => p[i]).join(sep);
    seen.set(pattern, {
      meter: pattern,
      rack: rackParts.map(i => p[i]).join(sep),
      company: companyParts.length ? ((labels && labels[code] && String(labels[code]).trim()) || code) : ""
    });
  }
  return [...seen.values()].sort((a, b) => collator.compare(a.company, b.company) || collator.compare(a.meter, b.meter));
}

/* ---------- reading racks and meters out of an Excel report ---------- */
/*
  The report has a header row (… Last Reading … Current Reading … Usage … Total Usage …). Under it, a rack row has a name and no
  readings, and the meter rows below it (name plus readings) belong to that rack, until the next rack row.
  rows = one sheet as an array of rows. Gives { ok, pairs: [{ meter, rack }], racks, duplicates }.
*/
function extractRackReport(rows) {
  const text = r => (r || []).map(c => String(c === undefined || c === null ? "" : c).trim());
  let hi = -1, cLast = -1, cCur = -1;
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const r = text(rows[i]);
    const l = r.findIndex(c => /last/i.test(c) && /read|kwh|\(a\)/i.test(c)), u = r.findIndex(c => /current/i.test(c) && /read|kwh|\(b\)/i.test(c));
    if (l >= 0 && u >= 0 && l !== u) { hi = i; cLast = l; cCur = u; break; }
  }
  if (hi < 0) return { ok: false, pairs: [], entries: [], racks: 0, duplicates: 0, company: "", location: "", from: "", to: "" };
  /* the lines above the table: Company Name : X, Location : Y, From Date : ..., To Date : ... */
  const meta = re => {
    for (let i = 0; i < hi; i++) {
      const r = text(rows[i]);
      const j = r.findIndex(c => re.test(c));
      if (j >= 0) for (let k = j + 1; k < r.length; k++) if (r[k] && r[k] !== ":") return r[k];
    }
    return "";
  };
  const info = { company: meta(/^company( name)?$/i), location: meta(/^location$/i), from: meta(/^from( date)?$/i), to: meta(/^to( date)?$/i) };
  const head = text(rows[hi]);
  let nameCol = head.findIndex((c, j) => j !== cLast && j !== cCur && /rack|meter|metric|pdu|device|name/i.test(c));
  if (nameCol < 0) nameCol = 0;
  const pairs = [], entries = [], seen = new Set();
  let rack = "", racks = 0, duplicates = 0;
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const label = String(r[nameCol] === undefined || r[nameCol] === null ? "" : r[nameCol]).trim();
    if (!label || /^total$/i.test(label)) continue;
    const hasReading = !Number.isNaN(parseNumber(r[cLast])) || !Number.isNaN(parseNumber(r[cCur]));
    if (hasReading) {
      entries.push({ meter: label, rack });
      const key = label.toLowerCase();
      if (seen.has(key)) { duplicates++; continue; }
      seen.add(key);
      pairs.push({ meter: label, rack });
    } else { rack = label; racks++; }
  }
  return { ok: true, pairs, entries, racks, duplicates, company: info.company, location: info.location, from: info.from, to: info.to };
}

/* The Custom Report workbook: one sheet per company (header with Last and Current Reading, a rack row, then its meters) plus a raw data sheet. */
function extractTemplate(XL, wb) {
  const sheets = [];
  for (const name of wb.SheetNames) {
    const ref = wb.Sheets[name]["!ref"];
    if (!ref) continue;
    const rg = XL.utils.decode_range(ref);
    if (rg.e.r - rg.s.r > 3000) continue;   /* the raw readings sheet is far longer than any report sheet */
    const r = extractRackReport(sheetRows(XL, wb, name));
    if (r.ok && r.entries.length) sheets.push({ name, company: r.company || name, location: r.location, racks: r.racks, entries: r.entries, duplicates: r.duplicates });
  }
  return sheets;
}
function templateToText(sheets) {
  const rows = [];
  for (const sh of sheets) for (const e of sh.entries) rows.push({ meter: e.meter, rack: e.rack, company: sh.company });
  return mappingToText(rows);
}

const r6 = x => Math.round(x * 1e6) / 1e6;
const cellText = c => c === undefined || c === null ? "" : String(c).trim();

/*
  ds   = { layout, headers, rows, cols }
  opts = { method, valueType, unit, dateOrder, from, to, mappingText, sort }
*/
function aggregate(ds, opts) {
  const { rows, cols, headers } = ds;
  const layout = ds.layout || "long";
  const order = opts.dateOrder || "dmy";
  const unit = opts.unit || 1;
  const mapper = buildMapper(opts.mappingText || "");
  const rackCo = buildRackCompany(opts.rackCompanyText || "");
  const wantCo = opts.company ? String(opts.company).toLowerCase() : "";
  const seenNames = new Set();
  const looked = new Map();
  const look = name => {
    let e = looked.get(name);
    if (e === undefined) { e = mapper.lookup(name) || null; looked.set(name, e); }
    return e;
  };
  /* A company comes from the file's own company column when it has one, otherwise from the mapping. One spelling is kept per company. */
  const spelling = new Map(listCompanies(opts.mappingText || "", opts.rackCompanyText || "").map(c => [c.toLowerCase(), c]));
  const fixSpelling = v => {
    if (!v) return "";
    const k = v.toLowerCase();
    if (!spelling.has(k)) spelling.set(k, v);
    return spelling.get(k);
  };
  /* meter -> every { rack, company } it belongs to. A rack or company written in the data file wins over the lists; a meter the lists name twice belongs to both. */
  const entriesOf = new Map();
  const entriesFor = (name, rackCell, coCell) => {
    let list = entriesOf.get(name);
    if (list !== undefined) return list;
    const dr = cellText(rackCell), dc = cellText(coCell);
    const mapped = mapper.all(name);
    list = [];
    if (dr || dc || !mapped.length) {
      const e = mapped[0] || null;
      const rack = dr || (e ? e.rack : "");
      let co = dc || (e ? e.company : "");
      if (!co && rack) co = rackCo.lookup(rack);
      list.push({ rack, company: fixSpelling(co), ord: e ? e.ord : null });
    } else {
      for (const e of mapped) list.push({ rack: e.rack, company: fixSpelling(e.company || (e.rack ? rackCo.lookup(e.rack) : "")), ord: e.ord });
    }
    entriesOf.set(name, list);
    return list;
  };
  /* with a company chosen, only that company's meters are used */
  const inCompany = (name, dataCell, rackCell) => {
    const list = entriesFor(name, rackCell, dataCell);
    return !wantCo || list.some(e => e.company !== "" && e.company.toLowerCase() === wantCo);
  };
  const fromMs = opts.from ? opts.from : null;
  const toMs = opts.to ? opts.to + 86400000 - 1 : null;
  const hasTime = layout === "wide" || layout === "transposed" || (layout === "long" && cols.time >= 0);
  const stats = new Map();
  const summary = new Map();
  let skippedValue = 0, skippedTime = 0, used = 0, tMin = Infinity, tMax = -Infinity;

  function add(name, rackData, v, t, idx, coData) {
    seenNames.add(name);
    if (!inCompany(name, coData, rackData)) return;
    if (Number.isNaN(v)) { skippedValue++; return; }
    v *= unit;
    if (hasTime) {
      if (Number.isNaN(t)) { skippedTime++; return; }
      if (t < tMin) tMin = t;
      if (t > tMax) tMax = t;
      if (fromMs !== null && t < fromMs) return;
      if (toMs !== null && t > toMs) return;
    }
    used++;
    let s = stats.get(name);
    if (!s) {
      s = { name, order: stats.size, min: v, max: v, minT: t, maxT: t, days: new Map(), sum: 0, fT: Infinity, fV: v, lT: -Infinity, lV: v, n: 0 };
      stats.set(name, s);
    }
    s.n++;
    s.sum += v;
    if (v < s.min) { s.min = v; s.minT = t; }
    if (v >= s.max) { s.max = v; s.maxT = t; }
    if (hasTime) {
      /* one bucket per calendar date: 12:00 a.m. to 11:59 p.m. */
      const dt = new Date(t);
      const dk = dt.getFullYear() * 10000 + (dt.getMonth() + 1) * 100 + dt.getDate();
      let b = s.days.get(dk);
      if (!b) { b = { min: v, max: v, minT: t, maxT: t, fT: t, fV: v, lT: t, lV: v, n: 0 }; s.days.set(dk, b); }
      b.n++;
      if (v < b.min) { b.min = v; b.minT = t; }
      if (v >= b.max) { b.max = v; b.maxT = t; }
      if (t < b.fT) { b.fT = t; b.fV = v; }
      if (t >= b.lT) { b.lT = t; b.lV = v; }
    }
    const tt = hasTime ? t : idx;
    if (tt < s.fT) { s.fT = tt; s.fV = v; }
    if (tt >= s.lT) { s.lT = tt; s.lV = v; }
  }

  if (layout === "long") {
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const name = cellText(r[cols.metric]);
      if (!name) continue;
      add(name, cols.rack >= 0 ? r[cols.rack] : "", parseNumber(r[cols.value]), hasTime ? parseTime(r[cols.time], order) : NaN, i, cols.company >= 0 ? r[cols.company] : "");
    }
  } else if (layout === "wide") {
    const names = cols.meterCols.map(j => cellText(headers[j]));
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const t = parseTime(r[cols.time], order);
      if (Number.isNaN(t)) { if (cellText(r[cols.time])) skippedTime++; continue; }
      for (let k = 0; k < names.length; k++) {
        const cell = r[cols.meterCols[k]];
        if (cellText(cell) === "") continue;
        add(names[k], "", parseNumber(cell), t, i);
      }
    }
  } else if (layout === "transposed") {
    const times = cols.timeCols.map(j => parseTime(headers[j], order));
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const name = cellText(r[cols.metric]);
      if (!name) continue;
      const rackData = cols.rack >= 0 ? r[cols.rack] : "";
      const coData = cols.company >= 0 ? r[cols.company] : "";
      for (let k = 0; k < times.length; k++) {
        const cell = r[cols.timeCols[k]];
        if (cellText(cell) === "") continue;
        add(name, rackData, parseNumber(cell), times[k], i, coData);
      }
    }
  } else if (layout === "summary") {
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const name = cellText(r[cols.metric]);
      if (!name || summary.has(name)) continue;
      seenNames.add(name);
      if (!inCompany(name, cols.company >= 0 ? r[cols.company] : "", cols.rack >= 0 ? r[cols.rack] : "")) continue;
      const a = parseNumber(r[cols.last]), b = parseNumber(r[cols.current]);
      if (Number.isNaN(a) || Number.isNaN(b)) { skippedValue++; continue; }
      used++;
      summary.set(name, { name, order: summary.size, last: a * unit, current: b * unit, count: 1 });
    }
  }

  /* counter (meter reading that keeps climbing) or interval (consumption per reading)? */
  let autoMode = "counter";
  if (layout !== "summary") {
    let eligible = 0, like = 0;
    for (const s of stats.values()) {
      if (s.n < 3) continue;
      eligible++;
      const tol = 1e-9 * Math.max(1, Math.abs(s.max));
      if (s.lV >= s.fV - tol && Math.abs(s.lV - s.max) <= tol && Math.abs(s.fV - s.min) <= tol) like++;
    }
    autoMode = eligible === 0 || like / eligible >= 0.7 ? "counter" : "interval";
  }
  let mode = layout === "summary" ? "summary" : autoMode;
  if (layout !== "summary" && (opts.valueType === "counter" || opts.valueType === "interval")) mode = opts.valueType;

  /* per date (max - min inside each date, then added up) or over the whole period? */
  let usageRule = "period", dayRatio = 0;
  if (layout !== "summary" && mode !== "interval" && hasTime) {
    let multi = 0, totalDays = 0;
    for (const s of stats.values()) for (const b of s.days.values()) { totalDays++; if (b.n >= 2) multi++; }
    dayRatio = totalDays ? multi / totalDays : 0;
    const choice = opts.usageRule || "auto";
    usageRule = choice === "auto" ? (dayRatio >= 0.5 ? "day" : "period") : choice;
  }

  const items = [];
  if (layout === "summary") {
    for (const s of summary.values()) items.push({ name: s.name, order: s.order, last: s.last, current: s.current, usage: r6(s.current - s.last), count: 1 });
  } else {
    for (const s of stats.values()) {
      if (mode === "interval") items.push({ name: s.name, order: s.order, last: null, current: null, usage: r6(s.sum), count: s.n });
      else {
        const byTime = opts.method === "time";
        if (usageRule === "day") {
          const keys = [...s.days.keys()].sort((a, b) => a - b);
          const first = s.days.get(keys[0]), lastDay = s.days.get(keys[keys.length - 1]);
          let usage = 0;
          for (const b of s.days.values()) usage += byTime ? b.lV - b.fV : b.max - b.min;
          items.push({
            name: s.name, order: s.order,
            last: byTime ? first.fV : first.min, current: byTime ? lastDay.lV : lastDay.max,
            lastAt: byTime ? first.fT : first.minT, currentAt: byTime ? lastDay.lT : lastDay.maxT,
            usage: r6(usage), count: s.n, days: s.days.size
          });
          continue;
        }
        const last = byTime ? s.fV : s.min;
        const current = byTime ? s.lV : s.max;
        const lastAt = hasTime ? (byTime ? s.fT : s.minT) : null;
        const currentAt = hasTime ? (byTime ? s.lT : s.maxT) : null;
        items.push({ name: s.name, order: s.order, last, current, lastAt, currentAt, usage: r6(current - last), count: s.n });
      }
    }
  }

  /* one line per rack a meter belongs to */
  const base = items.splice(0, items.length);
  for (const it of base) {
    const list = entriesFor(it.name, "", "");
    const use = wantCo ? list.filter(e => e.company !== "" && e.company.toLowerCase() === wantCo) : list;
    for (const e of use.length ? use : list) items.push(Object.assign({}, it, { rack: e.rack, company: e.company, ord: e.ord === null ? 1e9 + it.order : e.ord }));
  }
  const timesListed = new Map();
  for (const it of items) timesListed.set(it.name, (timesListed.get(it.name) || 0) + 1);
  const listedTwice = [...timesListed].filter(([, n]) => n > 1).map(([name]) => name);
  /* a rack name used by two companies (NR1 at both) gets the company beside it when several companies are shown */
  const rackCos = new Map();
  for (const it of items) {
    const k = (it.rack || "").toLowerCase();
    if (!rackCos.has(k)) rackCos.set(k, new Set());
    rackCos.get(k).add(it.company || "");
  }
  const byRack = new Map();
  const unassigned = [];
  for (const it of items) {
    let rackName = it.rack || "Unassigned";
    if (it.rack && rackCos.get(it.rack.toLowerCase()).size > 1 && it.company) rackName += " · " + it.company;
    if (!it.rack) unassigned.push(it.name);
    let rk = byRack.get(rackName);
    if (!rk) { rk = { name: rackName, meters: [], total: 0, order: it.order, ord: it.ord }; byRack.set(rackName, rk); }
    if (it.ord < rk.ord) rk.ord = it.ord;
    rk.meters.push(it);
  }
  const collator = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  const racks = [...byRack.values()];
  for (const rk of racks) {
    if (opts.sort === "list") rk.meters.sort((a, b) => a.ord - b.ord);
    else rk.meters.sort((a, b) => collator.compare(a.name, b.name));
    rk.total = r6(rk.meters.reduce((acc, m) => acc + m.usage, 0));
  }
  if (opts.sort === "usage") racks.sort((a, b) => b.total - a.total);
  else if (opts.sort === "order") racks.sort((a, b) => a.order - b.order);
  else if (opts.sort === "list") racks.sort((a, b) => (a.name === "Unassigned") - (b.name === "Unassigned") || a.ord - b.ord);
  else racks.sort((a, b) => (a.name === "Unassigned") - (b.name === "Unassigned") || collator.compare(a.name, b.name));
  const grand = r6(racks.reduce((acc, r) => acc + r.total, 0));
  const flat = items.length > 0 && items.every(it => !it.rack);
  const meters = items.slice();
  if (opts.sort === "usage") meters.sort((a, b) => b.usage - a.usage);
  else if (opts.sort === "order") meters.sort((a, b) => a.order - b.order);
  else if (opts.sort === "list") meters.sort((a, b) => a.ord - b.ord);
  else meters.sort((a, b) => collator.compare(a.name, b.name));
  /* how many meters in the file belong to each company (before the company filter) */
  const counts = new Map([...spelling.values()].map(c => [c, 0]));
  let noCompany = 0;
  const noCoRacks = new Set();
  for (const name of seenNames) {
    const list = entriesOf.get(name) || [];
    const cos = new Set(list.map(e => e.company).filter(Boolean));
    if (cos.size) for (const c of cos) counts.set(c, (counts.get(c) || 0) + 1);
    else { noCompany++; for (const e of list) if (e.rack) noCoRacks.add(e.rack); }
  }
  return {
    company: opts.company || "", companies: [...counts.keys()], companyCounts: [...counts].map(([name, meters]) => ({ name, meters })), noCompany, racksWithoutCompany: [...noCoRacks].sort((a, b) => a.localeCompare(b, "en", { numeric: true, sensitivity: "base" })), allNames: [...seenNames], listedTwice,
    racks, flat, meters, usageRule, dayRatio, grand, unassigned, skippedValue, skippedTime, used, mode, autoMode, layout,
    meterCount: items.length,
    dataFrom: Number.isFinite(tMin) ? tMin : null,
    dataTo: Number.isFinite(tMax) ? tMax : null
  };
}

function buildModel(result) {
  const out = [];
  if (result.flat) {
    for (const m of result.meters) out.push({ type: "meter", rack: "", name: m.name, last: m.last, current: m.current, lastAt: m.lastAt === undefined ? null : m.lastAt, currentAt: m.currentAt === undefined ? null : m.currentAt, usage: m.usage });
    out.push({ type: "total", total: result.grand });
    return out;
  }
  for (const rk of result.racks) {
    out.push({ type: "rack", name: rk.name, total: rk.total });
    for (const m of rk.meters) out.push({ type: "meter", rack: rk.name, name: m.name, last: m.last, current: m.current, lastAt: m.lastAt === undefined ? null : m.lastAt, currentAt: m.currentAt === undefined ? null : m.currentAt, usage: m.usage });
  }
  out.push({ type: "total", total: result.grand });
  return out;
}

const HEAD = ["Rack Name", "Last Reading (kWh) (A)", "Current Reading (kWh) (B)", "Usage (B-A=C)", "Total Usage (kWh) (E)"];
/* result gives the right column titles: no rack column when there are no racks, and a different Usage title when usage is added up per date */
function headFor(result) {
  const usage = result && result.usageRule === "day" ? "Usage (kWh) (C)" : "Usage (B-A=C)";
  if (result && result.flat) return ["Meter Name", "Last Reading (kWh) (A)", "Current Reading (kWh) (B)", usage];
  return ["Rack Name", "Last Reading (kWh) (A)", "Current Reading (kWh) (B)", usage, "Total Usage (kWh) (E)"];
}
function modelToRows(model, result) {
  const f = x => x === null || x === undefined ? "" : (Math.round(x * 100) / 100).toFixed(2);
  const flat = !!(result && result.flat);
  const rows = [headFor(result)];
  for (const r of model) {
    if (flat) {
      if (r.type === "meter") rows.push([r.name, f(r.last), f(r.current), f(r.usage)]);
      else rows.push(["Total", "", "", f(r.total)]);
    } else if (r.type === "rack") rows.push([r.name, "", "", "", f(r.total)]);
    else if (r.type === "meter") rows.push([r.name, f(r.last), f(r.current), f(r.usage), ""]);
    else rows.push(["Total", "", "", "", f(r.total)]);
  }
  return rows;
}
function toCsv(rows) {
  const q = s => /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  return rows.map(r => r.map(q).join(",")).join("\r\n") + "\r\n";
}
function toTsv(rows) { return rows.map(r => r.join("\t")).join("\n"); }

