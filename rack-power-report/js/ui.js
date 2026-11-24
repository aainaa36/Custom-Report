(function () {
  const $ = id => document.getElementById(id);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = x => (Math.round(x * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtDate = ms => ms === null || ms === undefined ? "" : new Date(ms).toLocaleDateString("en-GB");
  const LAYOUT_NAMES = {
    long: "one row per reading (meter, time, kWh)",
    wide: "one column per meter, one row per time",
    transposed: "one row per meter, dates across the top",
    summary: "Last and Current readings already in the file"
  };
  const LAYOUT_HINTS = {
    wide: "Every numeric column except the time column is treated as one meter.",
    transposed: "Every column whose header is a date is treated as one reading time.",
    summary: "Usage is Current minus Last for each meter in the file.",
    long: "", auto: ""
  };

  const state = {
    headers: [], rows: [], layoutChoice: "auto", layout: "long",
    cols: { metric: -1, time: -1, value: -1, rack: -1, company: -1, last: -1, current: -1, meterCols: [], timeCols: [] },
    dateChoice: "auto", dateOrder: "dmy", dateSure: true, dateSeen: 0,
    valueType: "auto", ruleChoice: "auto", unitChoice: "auto", unit: 1,
    files: [], isSample: true, notes: [], workbook: null, sheet: "",
    method: "minmax", from: 0, to: 0, sort: "name", location: "",
    mappingText: DEFAULT_MAPPING, mappingIsDefault: true, rackCoText: DEFAULT_RACK_COMPANY, rackCoIsDefault: true, collapsed: new Set(), result: null, model: [],
    company: "", companies: [], templateInfo: null, savedMapping: "", savedRackCo: "", xl: null, nt: { sig: "", company: new Set(), rack: new Set(), labels: {} }
  };
  const ALL = "__all__";   /* the "All meters" button */
  try {
    const saved = localStorage.getItem("rackReport.mapping");
    /* kept apart from the sample data shown at start-up, and used when a real file is loaded */
    if (saved && saved.trim()) state.savedMapping = saved;
    const savedRc = localStorage.getItem("rackReport.rackCompany");
    if (savedRc && savedRc.trim()) state.savedRackCo = savedRc;
  } catch (e) { /* storage may be unavailable */ }

  /* ---------- companies ---------- */
  /* The buttons come from the company column of the file when it has one, otherwise from the company column of the mapping.
     With no company anywhere the app works as a plain rack report. */
  function validCompany(list) {
    if (!list.length) return "";
    if (!state.company || state.company === ALL) return state.company;
    return list.find(c => c.toLowerCase() === state.company.toLowerCase()) || "";
  }
  const needChoice = () => state.companies.length > 0 && !state.company;
  const companyFilter = () => (state.company && state.company !== ALL ? state.company : "");
  const slug = t => String(t).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");

  /* ---------- loading ---------- */
  function shapeTable(all, label) {
    const hi = findHeader(all);
    const headers = (all[hi] || []).map(h => cellText(h));
    const data = [];
    for (let i = hi + 1; i < all.length; i++) {
      const r = all[i];
      let any = false;
      for (let j = 0; j < r.length; j++) if (cellText(r[j]) !== "") { any = true; break; }
      if (any) data.push(r);
    }
    return { headers, data, label };
  }

  function mergeSources(parts) {
    const headers = parts[0].headers;
    const norm = a => a.map(h => h.toLowerCase());
    const rows = [];
    const notes = [];
    for (let p = 0; p < parts.length; p++) {
      const part = parts[p];
      if (p === 0) { for (const r of part.data) rows.push(r); continue; }
      const a = norm(headers), b = norm(part.headers);
      if (a.length === b.length && a.every((h, i) => h === b[i])) { for (const r of part.data) rows.push(r); continue; }
      const map = a.map(h => b.indexOf(h));
      if (map.every(i => i >= 0)) {
        for (const r of part.data) rows.push(map.map(i => r[i]));
        notes.push({ kind: "info", text: part.label + ": columns were in a different order, matched by name." });
      } else {
        notes.push({ kind: "warn", text: part.label + " has different columns from the first file, so it was left out." });
      }
    }
    return { headers, rows, notes };
  }

  async function readOne(f) {
    const buf = await f.arrayBuffer();
    const ext = (f.name.split(".").pop() || "").toLowerCase();
    if (ext === "xlsx" || ext === "xlsm" || ext === "xls") {
      if (typeof XLSX === "undefined") return { failed: "excel-missing", label: f.name };
      const wb = XLSX.read(buf, { type: "array", cellNF: true });
      /* a Custom Report workbook has one report sheet per company next to the raw readings sheet: the report sheets give the racks and companies */
      const tpl = extractTemplate(XLSX, wb);
      const sheet = bestSheet(XLSX, wb, new Set(tpl.map(t => t.name)));
      const part = shapeTable(sheetRows(XLSX, wb, sheet), f.name);
      part.wb = wb; part.sheet = sheet; part.template = tpl.length ? tpl : null;
      return part;
    }
    return shapeTable(parseCSV(decodeBuffer(buf), detectDelimiter(decodeBuffer(buf))), f.name);
  }

  async function loadFiles(fileList) {
    const files = [...fileList];
    if (!files.length) return;
    setStatus("Reading " + files.length + (files.length === 1 ? " file…" : " files…"));
    const parts = [], notes = [];
    for (const f of files) {
      try {
        const p = await readOne(f);
        if (p.failed === "excel-missing") notes.push({ kind: "bad", text: f.name + " is an Excel file, but the Excel reader did not load here. Save it as CSV and load that instead." });
        else if (!p.headers.length || !p.data.length) notes.push({ kind: "warn", text: f.name + " has no readable rows." });
        else parts.push(p);
      } catch (e) {
        notes.push({ kind: "bad", text: f.name + " could not be read. Check that it is a CSV or Excel file." });
      }
    }
    if (!parts.length) {
      state.notes = notes.length ? notes : [{ kind: "bad", text: "Could not read any rows from the file." }];
      setStatus("");
      renderAll();
      return;
    }
    const merged = mergeSources(parts);
    state.headers = merged.headers;
    state.rows = merged.rows;
    state.files = files.map(f => f.name);
    state.isSample = false;
    state.notes = notes.concat(merged.notes);
    state.ruleChoice = "auto";
    const tpl = (parts.find(p => p.template) || {}).template || null;
    state.templateInfo = tpl;
    if (tpl) {
      state.mappingText = templateToText(tpl);
      state.mappingIsDefault = false;
            } else if (state.mappingIsDefault) {
      state.mappingText = state.savedMapping || BUILT_IN_MAPPING;
      if (state.savedMapping) state.mappingIsDefault = false;
    }
    if (state.rackCoIsDefault) {
      state.rackCoText = state.savedRackCo || BUILT_IN_RACK_COMPANY;
      if (state.savedRackCo) state.rackCoIsDefault = false;
    }
    /* with the template's racks the report follows the Excel template: readings are meter counters, Last is the first reading of the start date,
       Current is the last reading of the end date, and racks and meters keep the template's order */
    const likeTemplate = !!tpl || BUILT_IN_MAPPING.trim() !== "";
    state.valueType = likeTemplate ? "counter" : "auto";
    state.method = likeTemplate ? "time" : "minmax";
    state.sort = likeTemplate ? "list" : "name";
    if (!state.location) state.location = (tpl && tpl[0].location) || (likeTemplate ? BUILT_IN_LOCATION : "");
    state.workbook = files.length === 1 && parts[0].wb ? { wb: parts[0].wb, label: parts[0].label } : null;
    state.sheet = parts[0].sheet || "";
    state.from = 0; state.to = 0;
    state.collapsed.clear();
    if (likeTemplate) state.ruleChoice = "period";
    reanalyse();
    setStatus("");
    renderAll(true);
  }

  function loadSample() {
    const part = shapeTable(parseCSV(makeSampleCsv(), ","), "sample");
    state.headers = part.headers;
    state.rows = part.data;
    state.files = [];
    state.isSample = true;
    state.notes = [];
    state.workbook = null; state.sheet = "";
    state.layoutChoice = "auto"; state.dateChoice = "auto"; state.valueType = "auto"; state.unitChoice = "auto";
    state.ruleChoice = "period";
    state.templateInfo = null;
    state.from = 0; state.to = 0;
    state.collapsed.clear();
    state.mappingText = DEFAULT_MAPPING;
    state.mappingIsDefault = true;
    state.rackCoText = DEFAULT_RACK_COMPANY;
    state.rackCoIsDefault = true;
    state.company = listCompanies(DEFAULT_MAPPING, DEFAULT_RACK_COMPANY)[0] || "";
    reanalyse();
    renderAll(true);
  }

  /* Work out date order, layout, columns and unit from the loaded table. */
  function reanalyse() {
    const d = detectDateOrder(state.headers, state.rows);
    state.dateOrder = state.dateChoice === "auto" ? d.order : state.dateChoice;
    state.dateSure = state.dateChoice !== "auto" || d.sure;
    state.dateSeen = d.seen;
    state.layout = state.layoutChoice === "auto" ? detectLayout(state.headers, state.rows, state.dateOrder) : state.layoutChoice;
    state.cols = columnsFor(state.layout, state.headers, state.rows, state.dateOrder);
    updateUnit();
  }

  function updateUnit() {
    if (state.unitChoice !== "auto") { state.unit = { kWh: 1, Wh: 0.001, MWh: 1000 }[state.unitChoice]; return; }
    const c = state.cols, H = state.headers;
    let h = "";
    if (state.layout === "long") h = H[c.value];
    else if (state.layout === "wide") h = H[c.meterCols[0]];
    else if (state.layout === "summary") h = H[c.current];
    state.unit = detectUnit(h);
  }

  function colsReady() {
    const c = state.cols;
    switch (state.layout) {
      case "long": return c.metric >= 0 && c.value >= 0 && c.metric !== c.value;
      case "wide": return c.time >= 0 && c.meterCols.length > 0;
      case "transposed": return c.metric >= 0 && c.timeCols.length > 0;
      case "summary": return c.metric >= 0 && c.last >= 0 && c.current >= 0 && c.last !== c.current;
    }
    return false;
  }

  /* ---------- rendering ---------- */
  function setStatus(msg) { $("status").textContent = msg || ""; }

  function fillSelect(sel, value, allowNone) {
    let html = allowNone ? '<option value="-1">None</option>' : "";
    state.headers.forEach((h, i) => { html += '<option value="' + i + '">' + esc(h || ("Column " + (i + 1))) + "</option>"; });
    sel.innerHTML = html;
    sel.value = String(value);
  }

  function renderControls() {
    const c = state.cols;
    fillSelect($("colMetric"), c.metric, false);
    fillSelect($("colValue"), c.value, false);
    fillSelect($("colTime"), c.time, state.layout === "long");
    fillSelect($("colRack"), c.rack, true);
    fillSelect($("colCompany"), c.company, true);
    fillSelect($("colLast"), c.last, false);
    fillSelect($("colCurrent"), c.current, false);
    $("layoutSel").value = state.layoutChoice;
    $("dateSel").value = state.dateChoice;
    $("typeSel").value = state.valueType;
    $("ruleSel").value = state.ruleChoice;
    $("unitSel").value = state.unitChoice;
    $("mapping").value = state.mappingText;
    $("rackCo").value = state.rackCoText;
    $("location").value = state.location;
    $("sort").value = state.sort;
    document.querySelectorAll('input[name="method"]').forEach(r => { r.checked = r.value === state.method; });
    const sf = $("sheetField");
    if (state.workbook && state.workbook.wb.SheetNames.length > 1) {
      sf.hidden = false;
      $("sheetSel").innerHTML = state.workbook.wb.SheetNames.map(n => '<option value="' + esc(n) + '">' + esc(n) + "</option>").join("");
      $("sheetSel").value = state.sheet;
    } else sf.hidden = true;
    renderVisibility();
  }

  function renderVisibility() {
    document.querySelectorAll("[data-layouts]").forEach(el => { el.hidden = !el.dataset.layouts.split(" ").includes(state.layout); });
    $("layoutHint").textContent = state.layoutChoice === "auto" ? "Detected: " + LAYOUT_NAMES[state.layout] + ". " + LAYOUT_HINTS[state.layout] : LAYOUT_HINTS[state.layout];
  }

  function toDateInput(ms) {
    const d = new Date(ms);
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }
  function fromDateInput(v) {
    if (!v) return 0;
    const [y, m, d] = v.split("-").map(Number);
    return new Date(y, m - 1, d).getTime();
  }

  function compute() {
    if (!colsReady()) { state.result = null; state.model = []; state.companies = listCompanies(state.mappingText, state.rackCoText); state.company = validCompany(state.companies); return; }
    const run = co => aggregate({ layout: state.layout, headers: state.headers, rows: state.rows, cols: state.cols }, {
      method: state.method, valueType: state.valueType, usageRule: state.ruleChoice, unit: state.unit, dateOrder: state.dateOrder,
      from: state.from || null, to: state.to || null, mappingText: state.mappingText, rackCompanyText: state.rackCoText, sort: state.sort, company: co
    });
    let r = run(companyFilter());
    state.companies = r.companies;
    const ok = validCompany(state.companies);
    if (ok !== state.company) { state.company = ok; r = run(companyFilter()); }
    state.result = r;
    /* until a company is chosen no report is shown, but the result is kept for the layout notes and the meter counts on the buttons */
    state.model = needChoice() ? [] : buildModel(r);
  }

  /* the date range the client chooses: shown above the report, with a button per month in the file */
  function renderRange() {
    const card = $("rangeCard"), r = state.result;
    if (!r || r.dataFrom === null) { card.hidden = true; return; }
    card.hidden = false;
    const lo = toDateInput(r.dataFrom), hi = toDateInput(r.dataTo);
    $("from").min = lo; $("from").max = hi; $("to").min = lo; $("to").max = hi;
    $("from").value = state.from ? toDateInput(state.from) : lo;
    $("to").value = state.to ? toDateInput(state.to) : hi;
    const fromV = $("from").value, toV = $("to").value;
    const months = [];
    for (let d = new Date(new Date(r.dataFrom).getFullYear(), new Date(r.dataFrom).getMonth(), 1); d.getTime() <= r.dataTo; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
      const first = new Date(d.getFullYear(), d.getMonth(), 1), last = new Date(d.getFullYear(), d.getMonth() + 1, 0);
      months.push({ label: first.toLocaleDateString("en-GB", { month: "short", year: "numeric" }), a: toDateInput(Math.max(first.getTime(), r.dataFrom)), b: toDateInput(Math.min(last.getTime(), r.dataTo)) });
    }
    const btn = (label, a, b) => '<button type="button" class="chip alt" data-a="' + a + '" data-b="' + b + '" aria-pressed="' + (fromV === a && toV === b) + '">' + esc(label) + "</button>";
    $("rangePresets").innerHTML = btn("Whole file", lo, hi) + months.slice(-12).map(m => btn(m.label, m.a, m.b)).join("");
    let note;
    if (state.from && state.to && state.from > state.to) note = "The start date is after the end date. Choose a start date on or before the end date.";
    else if (r.mode === "interval") note = "Usage is the total of all readings from the start date to the end date, including both days.";
    else if (r.usageRule === "day") note = "Usage is worked out date by date inside the chosen dates, then added up. Both the start and end dates are included.";
    else note = "Last reading is the first reading inside the chosen dates and Current reading is the latest reading inside them, with the date and time of each shown in the report. Usage is Current minus Last. Both the start and end dates are included.";
    $("rangeNote").textContent = note;
    $("rangeHint").textContent = "The file has readings from " + fmtDate(r.dataFrom) + " to " + fmtDate(r.dataTo) + ".";
  }

  /* the list of meters, racks and companies, shown for copying, and the company names offered while typing */
  /* before a real file is loaded the page shows sample data, so the list to build on is the saved or built-in one */
  const baseList = () => state.isSample && state.mappingIsDefault && (state.savedMapping || BUILT_IN_MAPPING.trim()) ? (state.savedMapping || BUILT_IN_MAPPING) : state.mappingText;
  function renderAddList() {
    const base = baseList();
    const rows = parseMappingRows(base);
    const cos = [...new Set(rows.map(r => r.company).filter(Boolean))];
    $("addCoList").innerHTML = cos.map(c => '<option value="' + esc(c) + '"></option>').join("");
    if ($("xlCode").value !== base) $("xlCode").value = base;
    const racks = new Set(rows.filter(r => r.rack).map(r => r.rack.toLowerCase() + "|" + (r.company || "").toLowerCase())).size;
    $("xlCount").textContent = "It has " + rows.length + " lines, " + racks + " racks and " + cos.length + " companies.";
  }
  
  function renderCompanies() {
    const card = $("companyCard");
    if (!state.companies.length) { card.hidden = true; $("companyBtns").innerHTML = ""; return; }
    card.hidden = false;
    const r = state.result;
    const count = name => { const f = r && r.companyCounts.find(c => c.name === name); return f ? f.meters : null; };
    const chip = (value, label, n, alt) => {
      const on = state.company === value;
      const none = n === 0 && !state.isSample;
      return '<button type="button" class="chip' + (alt ? " alt" : "") + '" aria-pressed="' + on + '" data-co="' + esc(value) + '"' + (none ? ' disabled title="No meters for this company in the loaded file"' : "") + ">" + esc(label) + (n === null || n === undefined ? "" : '<span class="n" title="' + n + ' meters in the loaded file">' + n + "</span>") + "</button>";
    };
    let html = "";
    for (const name of state.companies) html += chip(name, name, count(name), false);
    html += chip(ALL, "All meters", r ? r.companyCounts.reduce((a, c) => a + c.meters, 0) + r.noCompany : null, true);
    $("companyBtns").innerHTML = html;
    $("companyHint").textContent = needChoice()
      ? "Pick a company to build its report. Only the meters that belong to that company are used."
      : state.company === ALL ? "Showing every meter in the file, including meters that belong to no company." : "";
  }

  function renderSource() {
    let html = "";
    if (state.isSample) html += '<span class="pill sample">Sample data</span><span class="muted">Showing example readings for 14 PDUs, split between two example companies. Load your own file to replace them.</span>';
    else html += '<span class="pill file">' + esc(state.files.length === 1 ? state.files[0] : state.files.length + " files") + '</span><span class="muted">' + state.rows.length.toLocaleString("en-US") + " data rows" + (state.sheet && state.workbook ? ", sheet “" + esc(state.sheet) + "”" : "") + "</span>";
    $("source").innerHTML = html;
  }

  function renderNotices() {
    const list = [];
    const c = state.cols, r = state.result, H = state.headers;
    const hn = i => i >= 0 && H[i] !== undefined ? "“" + H[i] + "”" : "none";
    if (!state.rows.length) list.push({ kind: "bad", text: "No data loaded." });
    else if (!colsReady()) list.push({ kind: "bad", text: "Could not tell which columns hold the meters and the kWh readings. Open “Adjust layout, columns, period and rack mapping” and choose the layout and columns." });
    if (r && !state.isSample) {
      let what = "";
      if (state.layout === "long") what = "meter = " + hn(c.metric) + ", reading = " + hn(c.value) + ", time = " + hn(c.time) + (c.rack >= 0 ? ", rack = " + hn(c.rack) : "") + (c.company >= 0 ? ", company = " + hn(c.company) : "");
      else if (state.layout === "wide") what = c.meterCols.length + " meter columns, time = " + hn(c.time);
      else if (state.layout === "transposed") what = "meter = " + hn(c.metric) + ", " + c.timeCols.length + " date columns" + (c.rack >= 0 ? ", rack = " + hn(c.rack) : "") + (c.company >= 0 ? ", company = " + hn(c.company) : "");
      else what = "meter = " + hn(c.metric) + ", last = " + hn(c.last) + ", current = " + hn(c.current) + (c.rack >= 0 ? ", rack = " + hn(c.rack) : "") + (c.company >= 0 ? ", company = " + hn(c.company) : "");
      list.push({ kind: "info", text: "Layout" + (state.layoutChoice === "auto" ? " detected" : "") + ": " + LAYOUT_NAMES[state.layout] + ". Using " + what + "." });
      if (r.mode === "interval") list.push({ kind: "info", text: "The readings look like consumption per interval, so Usage is the sum of all readings in the period and Last and Current are left blank." + (state.valueType === "auto" ? " If they are meter counters, change “What the readings are”." : "") });
      if (r.mode !== "interval" && r.mode !== "summary" && c.time >= 0 || state.layout === "wide" || state.layout === "transposed") {
        if (r.usageRule === "day") list.push({ kind: "info", text: "Usage is worked out date by date: for each date, the highest reading minus the lowest reading between 12:00 a.m. and 11:59 p.m., then the dates are added up." });
        else if (state.ruleChoice === "auto" && r.mode === "counter") list.push({ kind: "info", text: "Most dates in this file have only one reading, so Usage is the highest minus the lowest reading over the whole period. Change “How Usage is worked out” to use the per-date rule." });
      }
      if (state.unit !== 1) list.push({ kind: "info", text: "Readings were converted " + (state.unit === 0.001 ? "from Wh to kWh." : "from MWh to kWh.") });
      if (!state.dateSure && state.dateSeen > 0 && (state.layout !== "summary")) list.push({ kind: "warn", text: "Dates like 1/4/2026 could mean 1 April or 4 January. They were read as " + (state.dateOrder === "dmy" ? "day first" : "month first") + ". Change “Date order” if that is wrong." });
    }
    if (r && !state.companies.length && !state.isSample) list.push({ kind: "info", text: "No companies were found: this file has no company column and list 2 (rack → company) is empty. Fill in list 2 under “Adjust layout, columns, period, racks and companies”, or use “Find companies and racks in the meter names” below." });
    if (r && needChoice()) {
      list.push({ kind: "info", text: "Choose a company above to see its report." });
    } else if (r) {
      const co = companyFilter();
      const coMeters = co ? (r.companyCounts.find(c => c.name === co) || { meters: 0 }).meters : null;
      if (co && coMeters === 0) list.push({ kind: "warn", text: "No meters for “" + co + "” were found in this file. Check that the meter names in the mapping match the names in the file." });
      else if (r.used === 0) list.push({ kind: "warn", text: "No readings fall inside the chosen period or contain a numeric value." });
      if (state.companies.length && r.racksWithoutCompany.length && !state.isSample) list.push({ kind: "info", text: r.racksWithoutCompany.length + (r.racksWithoutCompany.length === 1 ? " rack has" : " racks have") + " no company in list 2: " + r.racksWithoutCompany.slice(0, 6).join(", ") + (r.racksWithoutCompany.length > 6 ? " and more" : "") + "." });
      if (state.companies.length && r.noCompany > 0 && !state.isSample) list.push({ kind: "info", text: r.noCompany + (r.noCompany === 1 ? " meter in the file belongs" : " meters in the file belong") + " to no company" + (co ? " and " + (r.noCompany === 1 ? "is" : "are") + " not included." : ".") });
      if (r.skippedValue) list.push({ kind: "warn", text: r.skippedValue.toLocaleString("en-US") + " cells were skipped because the reading is not a number." });
      if (r.skippedTime) list.push({ kind: "warn", text: r.skippedTime.toLocaleString("en-US") + " rows were skipped because their time could not be read." });
      if (r.flat) list.push({ kind: "info", text: "No rack names were found in the file or the mapping, so the report lists one row per meter. Add a rack mapping or choose a rack column to group the meters into racks." });
      else if (r.unassigned.length) list.push({ kind: "warn", text: r.unassigned.length + (r.unassigned.length === 1 ? " meter has" : " meters have") + " no rack and appear under “Unassigned”: " + r.unassigned.slice(0, 6).join(", ") + (r.unassigned.length > 6 ? " and more" : "") + ". Add them to the rack mapping." });
    }
    if (r && state.templateInfo && !state.isSample) list.push({ kind: "info", text: "Racks, meters and companies were read from " + state.templateInfo.length + " report sheets of the workbook (" + state.templateInfo.map(t => t.name).join(", ") + "). The readings come from the sheet “" + state.sheet + "”." });
    if (r && r.listedTwice.length && !needChoice()) list.push({ kind: "warn", text: r.listedTwice.length + (r.listedTwice.length === 1 ? " meter is" : " meters are") + " listed more than once in the racks and companies list (" + r.listedTwice.slice(0, 4).join(", ") + (r.listedTwice.length > 4 ? " and more" : "") + "). Each listing is counted, as in the Excel template, so a total can look too high. Check the template if that is not intended." });
    for (const n of state.notes) list.push(n);
    $("notices").innerHTML = list.map(n => '<div class="notice ' + n.kind + '">' + esc(n.text) + "</div>").join("");
    if (list.some(n => n.kind === "bad")) $("settings").open = true;
  }

  function renderStats() {
    const r = state.result;
    if (!r || needChoice()) { $("stats").innerHTML = ""; return; }
    const range = periodText() || "no dates in the file";
    $("stats").innerHTML =
      "<span>Total usage <b>" + fmt(r.grand) + " kWh</b></span>" +
      (r.flat ? "" : "<span>Racks <b>" + r.racks.length + "</b></span>") +
      "<span>Meters <b>" + r.meterCount + "</b></span>" +
      "<span>Readings <b>" + r.used.toLocaleString("en-US") + "</b></span>" +
      "<span>Period <b>" + esc(range) + "</b></span>";
  }

  function renderTable() {
    const wrap = $("tableWrap");
    if (needChoice()) {
      wrap.innerHTML = '<div class="empty">Choose a company above and its report appears here.</div>';
      return;
    }
    if (!state.result || !state.model.length || state.result.used === 0) {
      wrap.innerHTML = '<div class="empty">' + (companyFilter() && state.result ? "There are no readings for this company in the loaded file." : "The report appears here once a file with meters and kWh readings is loaded.") + "</div>";
      return;
    }
    const cell = (x, at) => x === null || x === undefined ? '<span class="muted">–</span>' : fmt(x);
    const head = headFor(state.result);
    const flat = state.result.flat;
    let html = "<table class=\"report\"><thead><tr>" + head.map(h => "<th scope=\"col\">" + esc(h) + "</th>").join("") + "</tr></thead><tbody>";
    let open = true;
    for (const r of state.model) {
      if (r.type === "rack") {
        open = !state.collapsed.has(r.name);
        html += '<tr class="rack"><th scope="row"><button type="button" class="tog" aria-expanded="' + open + '" data-rack="' + esc(r.name) + '"><span class="chev"></span>' + esc(r.name) + '</button></th><td></td><td></td><td></td><td class="tot num">' + fmt(r.total) + "</td></tr>";
      } else if (r.type === "meter") {
        html += '<tr class="meter' + (open ? "" : " hidden") + '"><th scope="row"' + (flat ? ' class="flat"' : "") + ">" + esc(r.name) + '</th><td class="num">' + cell(r.last, r.lastAt) + '</td><td class="num">' + cell(r.current, r.currentAt) + '</td><td class="num">' + fmt(r.usage) + "</td>" + (flat ? "" : '<td class="tot"></td>') + "</tr>";
      } else if (flat) {
        html += '<tr class="total"><th scope="row">Total</th><td></td><td></td><td class="num">' + fmt(r.total) + "</td></tr>";
      } else {
        html += '<tr class="total"><th scope="row">Total</th><td></td><td></td><td></td><td class="tot num">' + fmt(r.total) + "</td></tr>";
      }
    }
    html += "</tbody></table>";
    wrap.innerHTML = html;
  }

  /* ---------- finding companies and racks in the meter names ---------- */
  function codeCounts(a, parts) {
    const m = new Map();
    const names = state.result ? state.result.allNames : [];
    for (const n of names) {
      const p = n.split(a.sep);
      if (p.length !== a.nParts) continue;
      const code = parts.map(i => p[i]).join(a.sep);
      m.set(code, (m.get(code) || 0) + 1);
    }
    const collator = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
    return [...m].sort((x, y) => collator.compare(x[0], y[0]));
  }

  function renderNameTool() {
    return;
    const card = $("setupCard"), body = $("ntBody"), hint = $("ntHint");
    const r = state.result;
    if (!r || !r.allNames.length) { card.hidden = true; return; }
    card.hidden = false;
    const a = analyseNames(r.allNames);
    state.nt.analysis = a;
    if (!a.ok) {
      hint.textContent = "The meter names do not share a pattern that can be split into parts, so companies and racks have to come from a company column in the file or from the mapping.";
      body.innerHTML = "";
      return;
    }
    hint.textContent = "The meter names are split at “" + a.sep + "” into " + a.nParts + " parts (" + a.regular + " meters). Tick the part, or parts, that tell the company and the rack apart. For example the zone for the company, and the zone plus the rack code for the rack.";
    let html = '<table class="nt"><thead><tr><th scope="col">Part</th><th scope="col">Values found in the file</th><th scope="col" class="c">Company</th><th scope="col" class="c">Rack</th></tr></thead><tbody>';
    for (const p of a.parts) {
      const vals = p.values.slice(0, 6).map(v => esc(v.value)).join(", ") + (p.distinct > 6 ? ", … (" + p.distinct + " different)" : "");
      const box = kind => '<td class="c"><input type="checkbox" data-nt="' + kind + '" data-i="' + p.index + '"' + (state.nt[kind].has(p.index) ? " checked" : "") + ' aria-label="Part ' + (p.index + 1) + ' tells the ' + kind + '"></td>';
      html += "<tr><td>Part " + (p.index + 1) + "</td><td>" + vals + "</td>" + box("company") + box("rack") + "</tr>";
    }
    html += "</tbody></table>";
    const cp = [...state.nt.company].sort((x, y) => x - y), rp = [...state.nt.rack].sort((x, y) => x - y);
    if (cp.length) {
      const codes = codeCounts(a, cp);
      html += '<div><div class="muted" style="font-size:13px; margin-bottom:6px;">' + codes.length + (codes.length === 1 ? " company" : " companies") + " found. Type the real company name in each box; leave it as it is to use the code.</div><div class=\"nt-names\">";
      for (const [code, n] of codes) {
        html += '<div class="field"><label>' + esc(code) + " · " + n + " meters</label><input type=\"text\" data-label=\"" + esc(code) + "\" value=\"" + esc(state.nt.labels[code] !== undefined ? state.nt.labels[code] : code) + "\" maxlength=\"60\"></div>";
      }
      html += "</div></div>";
    }
    if (rp.length) html += '<div class="muted" style="font-size:13px;">' + codeCounts(a, rp).length + " racks found, named after the chosen part(s).</div>";
    html += '<div><button class="btn primary" id="ntApply" type="button"' + (cp.length || rp.length ? "" : " disabled") + ">Use these companies and racks</button> <span class=\"muted\" style=\"font-size:13px;\">This fills list 1 in the settings above (the company goes on each meter's own line), where you can still edit it.</span></div>";
    body.innerHTML = html;
  }

  function renderNote() {
    const r = state.result;
    let t = "";
    if (r && !needChoice()) {
      if (r.mode === "summary") t = "Last (A) and Current (B) are taken straight from the file.";
      else if (r.mode === "interval") t = "Usage (C) is the sum of every reading in the period, so Last and Current are not used.";
      else if (r.usageRule === "day") t = "Usage (C) is worked out one date at a time: for each date, the highest reading minus the lowest reading between 12:00 a.m. and 11:59 p.m. of that date, and the dates are added together. Last (A) is the lowest reading on the first date and Current (B) is the highest reading on the last date, so Usage can be a little lower than Current minus Last, because use between one date's last reading and the next date's first reading is not counted. Set Period from and Period to to the same date for a single-day report.";
      else if (state.method === "time") t = "Last (A) is each meter's first reading by time in the period and Current (B) is its last reading. The small grey text shows when each reading was taken.";
      else t = "Last (A) is each meter's lowest reading in the period and Current (B) is its highest. For a meter counter that keeps climbing, these are the first and the last readings. The small grey text shows when each was taken.";
    }
    $("howNote").textContent = t;
  }

  /* the chosen dates, or the whole file when none are chosen */
  function periodText() {
    const r = state.result;
    if (!r || r.dataFrom === null) return "";
    return fmtDate(state.from || r.dataFrom) + " to " + fmtDate(state.to || r.dataTo);
  }
  function renderLocation() {
    const parts = [];
    if (companyFilter()) parts.push("Company: " + companyFilter());
    if (state.location) parts.push("Location: " + state.location);
    if (state.result && state.result.dataFrom !== null) parts.push("Period: " + periodText());
    $("locLine").textContent = parts.join("   ·   ");
  }

  function renderAll(resetControls) {
    compute();
    if (resetControls) renderControls(); else renderVisibility();
    {
      const names = state.result ? state.result.allNames : [];
      const sig = names.length + "|" + (names[0] || "") + "|" + (names[names.length - 1] || "");
      if (sig !== state.nt.sig) {
        state.nt = { sig, company: new Set(), rack: new Set(), labels: {} };
        renderNameTool();
        $("ntDetails").open = !!state.result && !state.companies.length && !state.isSample;
      }
    }
    renderCompanies();
    renderRange();
    renderAddList();
    $("rep-h").textContent = "Usage by rack" + (companyFilter() ? " · " + companyFilter() : "");
    renderSource();
    renderNotices();
    renderStats();
    renderTable();
    renderNote();
    const flatNow = !!(state.result && state.result.flat) && !needChoice();
    $("expandAll").hidden = flatNow; $("collapseAll").hidden = flatNow;
    $("sortLabel").textContent = flatNow ? "Sort meters" : "Sort racks";
    renderLocation();
    $("dlPdf").disabled = !state.model.length;
    $("dlCsv").disabled = !state.model.length;
    $("copyTbl").disabled = !state.model.length;
  }

  /* ---------- events ---------- */
  const drop = $("drop");
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", e => { if (e.dataTransfer && e.dataTransfer.files.length) loadFiles(e.dataTransfer.files); });
  $("file").addEventListener("change", e => { loadFiles(e.target.files); e.target.value = ""; });

  const colChange = key => e => { state.cols[key] = Number(e.target.value); if (key === "value" || key === "current") updateUnit(); renderAll(); };
  $("colMetric").addEventListener("change", colChange("metric"));
  $("colValue").addEventListener("change", colChange("value"));
  $("colTime").addEventListener("change", colChange("time"));
  $("colRack").addEventListener("change", colChange("rack"));
  $("colCompany").addEventListener("change", colChange("company"));
  $("colLast").addEventListener("change", colChange("last"));
  $("colCurrent").addEventListener("change", colChange("current"));

  $("layoutSel").addEventListener("change", e => { state.layoutChoice = e.target.value; reanalyse(); renderAll(true); });
  $("dateSel").addEventListener("change", e => { state.dateChoice = e.target.value; reanalyse(); renderAll(true); });
  $("typeSel").addEventListener("change", e => { state.valueType = e.target.value; renderAll(); });
  $("ruleSel").addEventListener("change", e => { state.ruleChoice = e.target.value; renderAll(); });
  $("unitSel").addEventListener("change", e => { state.unitChoice = e.target.value; updateUnit(); renderAll(); });
  $("sheetSel").addEventListener("change", e => {
    if (!state.workbook || typeof XLSX === "undefined") return;
    const part = shapeTable(sheetRows(XLSX, state.workbook.wb, e.target.value), state.workbook.label);
    state.sheet = e.target.value;
    state.headers = part.headers; state.rows = part.data;
    state.from = 0; state.to = 0;
    reanalyse();
    renderAll(true);
  });

  $("rangePresets").addEventListener("click", e => {
    const b = e.target.closest("button[data-a]");
    if (!b) return;
    state.from = fromDateInput(b.dataset.a); state.to = fromDateInput(b.dataset.b);
    renderAll();
  });
  $("from").addEventListener("change", e => { state.from = fromDateInput(e.target.value); renderAll(); });
  $("to").addEventListener("change", e => { state.to = fromDateInput(e.target.value); renderAll(); });
  document.querySelectorAll('input[name="method"]').forEach(r => r.addEventListener("change", e => { if (e.target.checked) { state.method = e.target.value; renderAll(); } }));
  $("location").addEventListener("input", e => { state.location = e.target.value.trim(); renderLocation(); });
  $("sort").addEventListener("change", e => { state.sort = e.target.value; renderAll(); });

  let mapTimer = null;
  $("mapping").addEventListener("input", e => {
    state.mappingText = e.target.value;
    state.mappingIsDefault = false;
    clearTimeout(mapTimer);
    mapTimer = setTimeout(() => {
      try { localStorage.setItem("rackReport.mapping", state.mappingText); } catch (err) { /* ignore */ }
        state.collapsed.clear();
      renderAll();
    }, 250);
  });
  $("mapFile").addEventListener("change", async e => {
    const f = e.target.files[0];
    if (!f) return;
    const text = decodeBuffer(await f.arrayBuffer());
    const rows = parseMappingRows(text);
    const withRack = rows.filter(r => r.rack);
    const looksLikeReadings = withRack.length > 0 && withRack.filter(r => !Number.isNaN(parseNumber(r.rack)) || isDateText(r.rack)).length / withRack.length > 0.5;
    if (looksLikeReadings) {
      state.notes = state.notes.filter(n => !n.mapWarn).concat([{ kind: "bad", mapWarn: true, text: "“" + f.name + "” looks like a readings file (its second column holds dates or numbers), not list 1. Nothing was changed. List 1 has the columns: meter name, rack name (company is optional)." }]);
      e.target.value = "";
      renderAll();
      return;
    }
    state.notes = state.notes.filter(n => !n.mapWarn);
    if (!rows.length) {
      state.notes = state.notes.concat([{ kind: "bad", mapWarn: true, text: "“" + f.name + "” has no rows with a meter name, so list 1 was not changed." }]);
      e.target.value = "";
      renderAll();
      return;
    }
    state.mappingText = mappingToText(rows);
    state.mappingIsDefault = false;
    $("mapping").value = state.mappingText;
    try { localStorage.setItem("rackReport.mapping", state.mappingText); } catch (err) { /* ignore */ }
    state.collapsed.clear();
    e.target.value = "";
    renderAll();
  });
  $("mapReset").addEventListener("click", () => {
    state.mappingText = BUILT_IN_MAPPING;
    state.mappingIsDefault = true;
    state.savedMapping = "";
    $("mapping").value = state.mappingText;
    try { localStorage.removeItem("rackReport.mapping"); } catch (err) { /* ignore */ }
    renderAll();
  });

  /* ---------- importing racks and companies from the Excel report ---------- */
  function renderXl() {
    const body = $("xlBody"), x = state.xl;
    if (!x) { body.innerHTML = ""; return; }
    if (x.error) { body.innerHTML = '<div class="notice bad">' + esc(x.error) + "</div>"; return; }
    let html = '<table class="nt"><thead><tr><th scope="col">Sheet</th><th scope="col">Racks</th><th scope="col">Meters</th><th scope="col">Company name</th><th scope="col" class="c">Use</th></tr></thead><tbody>';
    x.sheets.forEach((sh, i) => {
      html += "<tr><td>" + esc(sh.name) + "</td><td>" + (sh.ok ? sh.racks : "–") + "</td><td>" + (sh.ok ? sh.pairs.length : "–") + "</td>" +
        (sh.ok
          ? '<td><input type="text" data-xl-name="' + i + '" value="' + esc(sh.company) + '" maxlength="60" style="width:100%"></td><td class="c"><input type="checkbox" data-xl-use="' + i + '"' + (sh.use ? " checked" : "") + ' aria-label="Use sheet ' + esc(sh.name) + '"></td>'
          : '<td colspan="2" class="muted">No header with Last Reading and Current Reading found</td>') + "</tr>";
    });
    html += "</tbody></table>";
    const usable = x.sheets.filter(sh => sh.ok && sh.use);
    html += '<div><button class="btn primary" id="xlApply" type="button"' + (usable.length ? "" : " disabled") + ">Use these racks and companies</button> <span class=\"muted\" style=\"font-size:13px;\">This fills list 1 in the settings above (meter, rack and company on one line), where you can still edit it.</span></div>";
    if (x.code) html += '<div><label for="xlCode" style="font-weight:600;font-size:14px;">Code line for the page (replaces line 2, <code>const BUILT_IN_MAPPING = "";</code>)</label><textarea id="xlCode" rows="3" readonly style="width:100%;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12px;margin-top:6px;">' + esc(x.code) + '</textarea><button class="btn small" id="xlCopy" type="button" style="margin-top:6px;">Copy code line</button> <span class="muted" style="font-size:13px;">One single line, so there are no quotes or line breaks to fix by hand.</span></div>';
    body.innerHTML = html;
  }

  $("xlFile").addEventListener("change", async e => {
    const f = e.target.files[0];
    if (!f) return;
    e.target.value = "";
    if (typeof XLSX === "undefined") { state.xl = { error: "The Excel reader did not load here, so the report cannot be read. Save each sheet as CSV and use “Load list 1 from CSV” instead." }; renderXl(); return; }
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellNF: true });
      const sheets = wb.SheetNames.map(name => {
        const r = extractRackReport(sheetRows(XLSX, wb, name));
        return { name, ok: r.ok && r.entries.length > 0, pairs: r.entries, racks: r.racks, duplicates: r.duplicates, company: r.company || name, use: true };
      });
      state.xl = sheets.some(sh => sh.ok) ? { file: f.name, sheets } : { error: "“" + f.name + "” has no sheet with a header that contains Last Reading and Current Reading, so no racks could be read." };
    } catch (err) {
      state.xl = { error: "“" + f.name + "” could not be read. Check that it is an Excel file." };
    }
    renderXl();
  });

    $("addBtn").addEventListener("click", () => {
    const say = (text, bad) => { $("addMsg").textContent = text; $("addMsg").style.color = bad ? "#b3261e" : ""; };
    const company = $("addCo").value.trim(), rack = $("addRack").value.trim();
    const meters = [...new Set($("addMeters").value.split(/[\n\r,;\t]+/).map(x => x.trim()).filter(Boolean))];
    if (!company) { say("Type the company name.", true); return; }
    if (!rack) { say("Type the rack name.", true); return; }
    if (!meters.length) { say("Type at least one meter name, one per line.", true); return; }
    const rows = parseMappingRows(baseList());
    const spelled = (rows.find(r => r.company && r.company.toLowerCase() === company.toLowerCase()) || {}).company || company;
    const key = (m, rk, co) => [m, rk, co || ""].join("\u0001").toLowerCase();
    const have = new Set(rows.map(r => key(r.meter, r.rack, r.company)));
    let added = 0, elsewhere = 0;
    for (const m of meters) {
      if (have.has(key(m, rack, spelled))) continue;
      if (rows.some(r => r.meter.toLowerCase() === m.toLowerCase())) elsewhere++;
      rows.push({ meter: m, rack, company: spelled });
      have.add(key(m, rack, spelled));
      added++;
    }
    if (!added) { say("Those meters are already listed under this rack and company.", true); return; }
    state.mappingText = mappingToText(rows);
    state.mappingIsDefault = false;
    $("mapping").value = state.mappingText;
    try { localStorage.setItem("rackReport.mapping", state.mappingText); } catch (err) { /* ignore */ }
    $("addRack").value = ""; $("addMeters").value = "";
    renderAll();
    say("Added " + added + (added === 1 ? " meter" : " meters") + " to rack " + rack + " of " + spelled + "." + (elsewhere ? " " + elsewhere + (elsewhere === 1 ? " was" : " were") + " already listed under another rack or company, so it is counted in both." : "") + " Saved in this browser. Copy the list below into js/lookup.js to make it permanent.");
  });
  $("listCopy").addEventListener("click", async () => {
    const ta = $("xlCode");
    ta.select();
    try { await navigator.clipboard.writeText(ta.value); setStatus("List copied."); }
    catch (err) { try { document.execCommand("copy"); setStatus("List copied."); } catch (e2) { /* the text is selected, so Ctrl+C works */ } }
  });
  
  $("xlBody").addEventListener("input", e => {
    const t = e.target;
    if (t.dataset && t.dataset.xlName !== undefined && state.xl) state.xl.sheets[Number(t.dataset.xlName)].company = t.value;
  });
  $("xlBody").addEventListener("change", e => {
    const t = e.target;
    if (t.dataset && t.dataset.xlUse !== undefined && state.xl) { state.xl.sheets[Number(t.dataset.xlUse)].use = t.checked; renderXl(); }
  });
  $("xlBody").addEventListener("click", e => {
    if (!e.target.closest("#xlApply") || !state.xl || !state.xl.sheets) return;
    const rows = [];
    for (const sh of state.xl.sheets) {
      if (!sh.ok || !sh.use) continue;
      const company = (sh.company || "").trim() || sh.name;
      for (const p of sh.pairs) rows.push({ meter: p.meter, rack: p.rack, company });
    }
    if (!rows.length) return;
    state.mappingText = mappingToText(rows);
    state.mappingIsDefault = false;
    $("mapping").value = state.mappingText;
    try { localStorage.setItem("rackReport.mapping", state.mappingText); } catch (err) { /* ignore */ }
    state.company = "";
    state.collapsed.clear();
    state.xl.code = "const BUILT_IN_MAPPING = " + JSON.stringify(state.mappingText) + ";";
    renderAll();
    setStatus("List filled in from the Excel report: " + rows.length + " meters, " + new Set(rows.map(r => r.company.toLowerCase())).size + " companies. To keep it for every client, copy the list from the box at the bottom of this card into js/lookup.js.");
  });

  $("ntBody").addEventListener("change", e => {
    const t = e.target;
    if (t.dataset && t.dataset.nt) {
      const set = state.nt[t.dataset.nt], i = Number(t.dataset.i);
      if (t.checked) set.add(i); else set.delete(i);
      renderNameTool();
    }
  });
  $("ntBody").addEventListener("input", e => {
    const t = e.target;
    if (t.dataset && t.dataset.label !== undefined) state.nt.labels[t.dataset.label] = t.value;
  });
  $("ntBody").addEventListener("click", e => {
    if (!e.target.closest("#ntApply")) return;
    const a = state.nt.analysis;
    if (!a || !a.ok || !state.result) return;
    const rows = buildNameMapping(state.result.allNames, a.sep, a.nParts, [...state.nt.company].sort((x, y) => x - y), [...state.nt.rack].sort((x, y) => x - y), state.nt.labels);
    if (!rows.length) return;
    state.mappingText = mappingToText(rows);
    state.mappingIsDefault = false;
    $("mapping").value = state.mappingText;
    try { localStorage.setItem("rackReport.mapping", state.mappingText); } catch (err) { /* ignore */ }
    state.company = "";
    state.collapsed.clear();
    renderAll();
    setStatus("List 1 filled in: " + rows.length + (rows.length === 1 ? " line." : " lines.") + (state.companies.length ? " Choose a company above." : ""));
  });

  $("companyBtns").addEventListener("click", e => {
    const b = e.target.closest(".chip");
    if (!b || b.disabled) return;
    state.company = b.dataset.co;
    state.collapsed.clear();
    renderAll();
  });

  /* list 2: which company each rack belongs to */
  let rcTimer = null;
  $("rackCo").addEventListener("input", e => {
    state.rackCoText = e.target.value;
    state.rackCoIsDefault = false;
    clearTimeout(rcTimer);
    rcTimer = setTimeout(() => {
      try { localStorage.setItem("rackReport.rackCompany", state.rackCoText); } catch (err) { /* ignore */ }
      state.collapsed.clear();
      renderAll();
    }, 250);
  });
  $("rcFile").addEventListener("change", async e => {
    const f = e.target.files[0];
    if (!f) return;
    const rows = parseRackCompanyRows(decodeBuffer(await f.arrayBuffer()));
    const readingsLike = rows.length > 0 && rows.filter(r => !Number.isNaN(parseNumber(r.company)) || isDateText(r.company)).length / rows.length > 0.5;
    state.notes = state.notes.filter(n => !n.rcWarn);
    if (readingsLike || !rows.length) {
      state.notes = state.notes.concat([{ kind: "bad", rcWarn: true, text: "“" + f.name + "” " + (readingsLike ? "looks like a readings file" : "has no rows with both a rack name and a company") + ", not list 2. Nothing was changed. List 2 has two columns: rack name, company." }]);
      e.target.value = "";
      renderAll();
      return;
    }
    state.rackCoText = rackCompanyToText(rows);
    state.rackCoIsDefault = false;
    $("rackCo").value = state.rackCoText;
    try { localStorage.setItem("rackReport.rackCompany", state.rackCoText); } catch (err) { /* ignore */ }
    state.collapsed.clear();
    e.target.value = "";
    renderAll();
  });
  $("rcReset").addEventListener("click", () => {
    state.rackCoText = BUILT_IN_RACK_COMPANY;
    state.rackCoIsDefault = true;
    state.savedRackCo = "";
    $("rackCo").value = state.rackCoText;
    try { localStorage.removeItem("rackReport.rackCompany"); } catch (err) { /* ignore */ }
    state.notes = state.notes.filter(n => !n.rcWarn);
    renderAll();
  });

  $("tableWrap").addEventListener("click", e => {
    const b = e.target.closest(".tog");
    if (!b) return;
    const name = b.dataset.rack;
    if (state.collapsed.has(name)) state.collapsed.delete(name); else state.collapsed.add(name);
    renderTable();
  });
  $("expandAll").addEventListener("click", () => { state.collapsed.clear(); renderTable(); });
  $("collapseAll").addEventListener("click", () => { if (state.result) state.result.racks.forEach(r => state.collapsed.add(r.name)); renderTable(); });

  function fileStamp() {
    const ms = state.result && state.result.dataTo ? state.result.dataTo : Date.now();
    return (companyFilter() ? slug(companyFilter()) + "-" : "") + toDateInput(ms);
  }

  async function saveFile(filename, data) {
    try {
      const dl = window.claude && window.claude.use ? await window.claude.use("downloads") : null;
      if (!dl) {
        /* a normal page (for example on GitHub Pages): use the browser's own download */
        const type = /\.pdf$/i.test(filename) ? "application/pdf" : "text/csv;charset=utf-8";
        const a = document.createElement("a");
        a.href = URL.createObjectURL(new Blob([data], { type }));
        a.download = filename;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 2000);
        setStatus("Saved " + filename + ".");
        return;
      }
      await dl.save({ filename, data });
      setStatus("Saved " + filename + ".");
    } catch (err) {
      if (err && err.code === "declined") setStatus("Save cancelled.");
      else setStatus("Could not save the file here. Use Copy table and paste into Excel.");
    }
  }

  $("dlCsv").addEventListener("click", () => {
    if (!state.model.length) return;
    saveFile("rack-power-report-" + fileStamp() + ".csv", toCsv(modelToRows(state.model, state.result)));
  });

  /* ---------- PDF report: the same table as the page, under a header like the Excel template ---------- */
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const stampDate = ms => {
    const d = new Date(ms);
    return String(d.getDate()).padStart(2, "0") + "-" + MONTHS[d.getMonth()] + "-" + String(d.getFullYear()).slice(2);
  };
  const stampTime = ms => new Date(ms).toLocaleDateString("en-GB") + " " + new Date(ms).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });
  /* an optional logo: put the file at assets/logo.png next to index.html and it appears at the top right of the PDF */
  function loadLogo() {
    return new Promise(resolve => {
      const img = new Image();
      img.onload = () => resolve(img.naturalWidth ? img : null);
      img.onerror = () => resolve(null);
      img.src = "assets/logo.png";
    });
  }

  $("dlPdf").addEventListener("click", async () => {
    if (!state.model.length || !state.result) return;
    const jsPDF = window.jspdf && window.jspdf.jsPDF;
    if (!jsPDF || !(new jsPDF().autoTable)) { setStatus("The PDF library did not load. Check the internet connection and try again."); return; }
    try {
      const r = state.result, flat = r.flat, co = companyFilter();
      const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
      const W = doc.internal.pageSize.getWidth(), M = 10;
      /* header */
      doc.setFont("helvetica", "bold"); doc.setFontSize(15); doc.setTextColor(0);
      doc.text((co ? co.toUpperCase() + " " : "") + "POWER CONSUMPTION REPORT", M, 16);
      const logo = await loadLogo();
      if (logo) {
        try {
          const h = 14, w = h * logo.naturalWidth / logo.naturalHeight;
          doc.addImage(logo, "PNG", W - M - Math.min(w, 60), 8, Math.min(w, 60), Math.min(w, 60) * logo.naturalHeight / logo.naturalWidth);
        } catch (err) { /* a logo that cannot be read is left out */ }
      }
      const fromMs = state.from || r.dataFrom, toMs = state.to || r.dataTo;
      const left = [["Company Name", co || "All companies"], ["Location", state.location || ""]];
      const right = [["Reading Date", stampDate(Date.now())], ["From Date", fromMs ? stampDate(fromMs) : ""], ["To Date", toMs ? stampDate(toMs) : ""]];
      doc.setFontSize(10);
      const info = (rows, x, colonX, valueX) => rows.forEach((row, i) => {
        const y = 30 + i * 7;
        doc.setFont("helvetica", "normal");
        doc.text(row[0], x, y); doc.text(":", colonX, y); doc.text(String(row[1]), valueX, y);
      });
      info(left, M, M + 30, M + 34);
      info(right, 112, 142, 146);
      /* table */
      const head = headFor(r);
      const num = x => x === null || x === undefined || !Number.isFinite(x) ? "-" : fmt(x);
      const body = [];
      for (const row of state.model) {
        if (row.type === "rack") body.push({ kind: "rack", cells: [row.name, "", "", "", fmt(row.total)] });
        else if (row.type === "meter") {
          const cells = flat ? [row.name, num(row.last), num(row.current), num(row.usage)] : [row.name, num(row.last), num(row.current), num(row.usage), ""];
          body.push({ kind: "meter", cells, at: [row.lastAt, row.currentAt] });
        } else body.push({ kind: "total", cells: flat ? ["Total", "", "", fmt(row.total)] : ["Total", "", "", "", fmt(row.total)] });
      }
      const widths = flat ? [70, 40, 40, 40] : [56, 36, 36, 31, 31];
      const columnStyles = {};
      widths.forEach((w, i) => { columnStyles[i] = { cellWidth: w, halign: i === 0 ? "left" : "right" }; });
      doc.autoTable({
        startY: 56,
        margin: { left: M, right: M, top: 14, bottom: 14 },
        head: [head],
        body: body.map(b => b.cells),
        theme: "grid",
        styles: { font: "helvetica", fontSize: 9, cellPadding: { top: 1.6, bottom: 1.6, left: 2, right: 2 }, lineColor: [150, 150, 150], lineWidth: 0.2, textColor: 0, valign: "middle" },
        headStyles: { fillColor: [191, 191, 191], textColor: 0, fontStyle: "bold", halign: "center" },
        columnStyles,
        rowPageBreak: "avoid",
        didParseCell: d => {
          if (d.section !== "body") return;
          const b = body[d.row.index];
          if (b.kind === "rack") { d.cell.styles.fillColor = [217, 217, 217]; d.cell.styles.fontStyle = "bold"; }
          else if (b.kind === "total") { d.cell.styles.fillColor = [191, 191, 191]; d.cell.styles.fontStyle = "bold"; }
          else {
            if (d.column.index === 0) d.cell.styles.cellPadding = { top: 1.6, bottom: 1.6, left: 6, right: 2 };
            /* keep room under a reading for the date and time of that reading */
            if ((d.column.index === 1 || d.column.index === 2) && b.at[d.column.index - 1]) { d.cell.text = [d.cell.text[0], " "]; }
          }
        },
        didDrawCell: d => {
          if (d.section !== "body") return;
          const b = body[d.row.index];
          if (b.kind !== "meter" || (d.column.index !== 1 && d.column.index !== 2)) return;
          const at = b.at[d.column.index - 1];
          if (!at || !Number.isFinite(at)) return;
          doc.setFont("helvetica", "normal"); doc.setFontSize(7); doc.setTextColor(110);
          doc.text(stampTime(at), d.cell.x + d.cell.width - 2, d.cell.y + d.cell.height - 1.8, { align: "right" });
          doc.setFontSize(9); doc.setTextColor(0);
        }
      });
      const pages = doc.internal.getNumberOfPages();
      for (let i = 1; i <= pages; i++) {
        doc.setPage(i);
        doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(110);
        doc.text("Page " + i + " of " + pages, W - M, doc.internal.pageSize.getHeight() - 7, { align: "right" });
      }
      await saveFile("rack-power-report-" + fileStamp() + ".pdf", new Blob([doc.output("arraybuffer")], { type: "application/pdf" }));
    } catch (err) {
      setStatus("Could not build the PDF. Use Download CSV instead.");
    }
  });

  $("copyTbl").addEventListener("click", async () => {
    if (!state.model.length) return;
    const text = toTsv(modelToRows(state.model, state.result));
    let ok = false;
    try { await navigator.clipboard.writeText(text); ok = true; } catch (e) { /* try fallback */ }
    if (!ok) {
      try {
        const ta = document.createElement("textarea");
        ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta); ta.select();
        ok = document.execCommand("copy");
        document.body.removeChild(ta);
      } catch (e) { ok = false; }
    }
    setStatus(ok ? "Table copied. Paste it into Excel." : "Copy was blocked here. Select the table and copy it by hand.");
  });

  /* start with sample data so the page opens as a working report */
  loadSample();
  window.addEventListener("load", () => renderAll());
})();
