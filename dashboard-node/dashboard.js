#!/usr/bin/env node
/*
 * Vanilla Node.js port of "Automation 2.py" (OneTrust TPRM assessment dashboard, D2).
 *
 * Functional parity with the Python script is the contract: same filters, same
 * pivot semantics (including pandas' count-of-non-null-IDs), same KPI rounding
 * (Python's banker's rounding), same chart geometry, same Excel styling, same
 * four output filenames. Where pandas has a quirk, the quirk is reproduced and
 * commented rather than "improved".
 *
 * Dependencies (the only two allowed): exceljs, canvas.
 * Run with:  npm start   (reads "PRP Sample Jun (2).xlsx" from this folder)
 */
"use strict";

const fs = require("fs");
const ExcelJS = require("exceljs");
const { createCanvas } = require("canvas");

// =========================
// INPUT / OUTPUT
// =========================
const input_file = "PRP Sample Jun (2).xlsx";
const sheet_name = "OneTrust Assessment";
const output_file = "output file D2.xlsx";

const chart1_file = "chart_2026_assessment_status.png";
const chart2_file = "chart_supplier_response_time.png";

// =====================================================
// pandas equivalents, hand-written
// =====================================================

// ExcelJS cell.value can be a string, number, Date, boolean, null, or an
// object (richText / hyperlink / formula-with-cached-result). Reduce all of
// them to the plain value pandas would have seen via openpyxl.
function cellVal(cell) {
  let v = cell ? cell.value : null;
  if (v === null || v === undefined) return null;
  if (v instanceof Date || typeof v === "number" || typeof v === "string" || typeof v === "boolean") return v;
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
    if ("formula" in v || "sharedFormula" in v || "result" in v) {
      // pandas (openpyxl data_only) reads the cached result. exceljs OMITS the
      // result key entirely when the cache is falsy ("", 0, FALSE), so those
      // three are unrecoverable through exceljs — blank is the closest
      // reachable equivalent, and matches pandas for the "" case at least.
      const r = v.result;
      if (r === null || r === undefined) return null;
      if (typeof r === "object" && r.error) return null;
      return r;
    }
    if (v.hyperlink !== undefined) {
      const t = v.text;
      if (t && Array.isArray(t.richText)) return t.richText.map((x) => x.text).join("");
      return t !== undefined && t !== null ? t : v.hyperlink;
    }
    if (v.error) return null;
    if (v.text !== undefined) return v.text;
  }
  return v;
}

// pandas read_excel's default NA sentinels (keep_default_na=True): these exact
// strings read as NaN, so an ID of "N/A" is NOT counted by the pivot and an
// Organization of "NULL" becomes "" via fillna. Exact match on the raw cell
// text, no trimming — that is how pandas applies them.
const NA_TOKENS = new Set([
  "", "#N/A", "#N/A N/A", "#NA", "-1.#IND", "-1.#QNAN", "-NaN", "-nan",
  "1.#IND", "1.#QNAN", "<NA>", "N/A", "NA", "NULL", "NaN", "None", "n/a", "nan", "null",
]);
function naSieve(v) {
  return typeof v === "string" && NA_TOKENS.has(v) ? null : v;
}

// df[c].fillna("").astype(str).str.strip() — NaN/None become "", everything
// else is stringified then trimmed. Python spells booleans True/False and
// prints a datetime as "2026-01-01 00:00:00", not JS Date's locale form.
// Not emulated (documented, exotic): pandas dtype-dependent float formatting —
// a fully numeric text column containing a blank becomes float64 and would
// stringify 123 as "123.0"; this port always gives "123".
function pyStr(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number" && Number.isNaN(v)) return "";
  if (v === true) return "True";
  if (v === false) return "False";
  if (v instanceof Date) {
    const p = (n) => String(n).padStart(2, "0");
    return `${v.getUTCFullYear()}-${p(v.getUTCMonth() + 1)}-${p(v.getUTCDate())}` +
      ` ${p(v.getUTCHours())}:${p(v.getUTCMinutes())}:${p(v.getUTCSeconds())}`;
  }
  return String(v).trim();
}

// pd.to_datetime(..., errors="coerce").dt.year equivalent, returning the year
// or null. Handles the three shapes exceljs can produce for a date column:
// a JS Date (exceljs builds these in UTC — use UTC getters, never local, or
// timezones west of UTC lose every 1 January), an Excel serial number, or
// text in the common export shapes. Unparseable -> null, which simply fails
// the year filter instead of crashing (that is what errors="coerce" does).
//
// Two deliberate departures from pandas, both mandated by the spec ("normalize
// all three ... treat unparseable dates as failing the filter"):
//  - a column mixing several TEXT date shapes: pandas >= 2.0 infers a format
//    from the first value and coerces the rest to NaT; this parses per value.
//  - a numeric column with no date format: pandas reads epoch-nanoseconds
//    (every year becomes 1970); this reads an Excel serial.
// On date-formatted cells and single-format text — every real export — the
// two agree exactly.
const MONTH_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
// Real calendar validation: pandas coerces "2026-02-31" and "31/04/2026" to
// NaT, so a d<=31 shortcut would let impossible dates through the year filter.
function ymdYear(y, mo, d) {
  if (mo < 1 || mo > 12 || d < 1) return null;
  const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const dim = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return d <= dim[mo - 1] ? y : null;
}
function dateYear(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.getUTCFullYear();
  if (typeof v === "number") {
    if (!Number.isFinite(v)) return null;
    // Excel 1900 date system: serial 25569 = 1970-01-01 UTC.
    return new Date(Math.round((v - 25569) * 86400000)).getUTCFullYear();
  }
  const s = String(v).trim();
  if (!s) return null;
  let m;
  // Every date pattern is end-guarded (lookahead for end, whitespace, or a
  // time part): pandas coerces "2026-12-14xyz" to NaT, so a bare prefix match
  // must not count. Two-digit years ("14-Dec-26") and compact digits
  // ("20261214") are NOT emulated — exotic, and ambiguous without pandas'
  // century pivot; those coerce to null here.
  // Year first: 2026-12-14, 2026/12/14, 2026.12.14, 2026-12-14T09:30:00
  m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?=$|[T\s])/);
  if (m) return ymdYear(+m[1], +m[2], +m[3]);
  // Day/month numeric: pandas defaults to dayfirst=False, so a leading value
  // that COULD be a month is read as one; above 12 it can only be the day.
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?=$|[T\s])/);
  if (m) {
    const a = +m[1], b = +m[2], y = +m[3];
    return a <= 12 ? ymdYear(y, a, b) : ymdYear(y, b, a);
  }
  // Named month: 14-Dec-2026, 14 Dec 2026, 14 December 2026
  m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s,]+(\d{4})(?=$|[T\s])/);
  if (m) {
    const mo = MONTH_NUM[m[2].slice(0, 3).toLowerCase()];
    return mo ? ymdYear(+m[3], mo, +m[1]) : null;
  }
  // Month first: Dec 14, 2026 / December 14 2026
  m = s.match(/^([A-Za-z]{3,})[-\s](\d{1,2})[-\s,]+(\d{4})(?=$|[T\s])/);
  if (m) {
    const mo = MONTH_NUM[m[1].slice(0, 3).toLowerCase()];
    return mo ? ymdYear(+m[3], mo, +m[2]) : null;
  }
  // Month + year: "Dec 2026" -> pandas reads 2026-12-01
  m = s.match(/^([A-Za-z]{3,})[-\s,]+(\d{4})$/);
  if (m) {
    const mo = MONTH_NUM[m[1].slice(0, 3).toLowerCase()];
    return mo ? +m[2] : null;
  }
  // NO bare-year branch, deliberately: a text cell holding just "2026" is
  // type-inferred to int64 by pandas.read_excel BEFORE to_datetime runs, so
  // the script reads it as epoch-nanoseconds (year 1970) and drops the row.
  // Parsing it here as a year would diverge from the script. (Verified against
  // pandas 3.0.3 — to_datetime alone parses "2026", read_excel's inference
  // means it never gets the chance.)
  return null;
}

// Python's round(x, 2) decides from the EXACT binary value of the double and
// resolves true ties half-to-even. Multiplying by 100 first is wrong twice
// over: the multiply itself rounds (1/40 = 0.025 stored as 0.02500000000000000138…
// becomes 2.4999999999999996, flipping Python's 0.03 to 0.02), and it turns
// near-ties into false ties. Instead read the double's correctly-rounded
// 20-place decimal expansion and compare the tail beyond 2 places against an
// exact "500…0". The only doubles that ARE exact 2-dp midpoints are the
// eighths (0.125, 0.375, 0.625, 0.875) — everything else sits measurably to
// one side, and 20 places is far finer than double spacing, so the string
// comparison is exact. Verified against CPython over every n/d, d <= 200.
function pyRound2(x) {
  const s = x.toFixed(20); // e.g. "0.02500000000000000139"
  const dot = s.indexOf(".");
  const intPart = s.slice(0, dot), dec = s.slice(dot + 1);
  let k = Number(intPart) * 100 + Number(dec.slice(0, 2)); // value in hundredths, truncated
  const tail = dec.slice(2); // 18 digits deciding the direction
  const mid = "5".padEnd(tail.length, "0");
  if (tail > mid) k += 1;
  else if (tail === mid) k += k % 2 === 0 ? 0 : 1; // true tie: half-to-even
  return k / 100;
}

// Python sorts strings by code POINT ("Zebra" < "apple"); localeCompare does
// not, and JS's own < compares UTF-16 code UNITS, which reverses the order
// when an astral-plane character (emoji lead surrogate 0xD8xx) meets a high
// BMP character. Compare true code points, like Python does.
function ordinal(a, b) {
  const A = Array.from(String(a)), B = Array.from(String(b));
  const n = Math.min(A.length, B.length);
  for (let i = 0; i < n; i++) {
    const d = A[i].codePointAt(0) - B[i].codePointAt(0);
    if (d !== 0) return d < 0 ? -1 : 1;
  }
  return A.length === B.length ? 0 : A.length < B.length ? -1 : 1;
}

async function main() {
  // =========================
  // CHECK FILE
  // =========================
  if (!fs.existsSync(input_file)) {
    throw new Error(`Input file not found: ${input_file}`);
  }

  // =========================
  // LOAD DATA
  // =========================
  const inWb = new ExcelJS.Workbook();
  await inWb.xlsx.readFile(input_file);
  const inWs = inWb.getWorksheet(sheet_name);
  if (!inWs) throw new Error(`Worksheet named '${sheet_name}' not found`);

  // df.columns.str.strip(); EXACT duplicate headers: pandas renames later ones
  // at read time ("Tags.1"), so df["Tags"] is always the FIRST occurrence —
  // same here. Headers that collide only AFTER stripping ("Tags" + "Tags ")
  // are different: pandas mangles nothing, both become "Tags", and the script
  // dies on df["Tags"].str (a DataFrame has no .str) before writing anything.
  // Match the outcome that matters — no output — with a readable error.
  const headerCol = {};
  const rawsByName = new Map();
  inWs.getRow(1).eachCell({ includeEmpty: false }, (cell, col) => {
    const raw = cellVal(cell);
    const name = pyStr(raw);
    if (!name) return;
    if (!rawsByName.has(name)) rawsByName.set(name, new Set());
    rawsByName.get(name).add(String(raw));
    if (!(name in headerCol)) headerCol[name] = col;
  });

  const required_columns = ["ID", "Organization", "Stage", "Date created", "Tags"];
  const missing = required_columns.filter((c) => !(c in headerCol));
  if (missing.length) {
    // Python: raise ValueError(f"Missing required columns: {missing}")
    throw new Error(`Missing required columns: [${missing.map((c) => `'${c}'`).join(", ")}]`);
  }
  for (const c of required_columns) {
    if (rawsByName.get(c) && rawsByName.get(c).size > 1) {
      throw new Error(`Duplicate column '${c}' after header strip — the Python script aborts here too (df['${c}'] becomes a DataFrame)`);
    }
  }

  const df = [];
  inWs.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // header
    const get = (name) => naSieve(cellVal(row.getCell(headerCol[name])));
    df.push({
      // ID stays raw: the pivot's aggfunc="count" counts NON-NULL IDs, so a
      // blank (or NA-sentinel) ID must remain distinguishable from a value.
      id: get("ID"),
      organization: pyStr(get("Organization")),
      stage: pyStr(get("Stage")),
      tags: pyStr(get("Tags")),
      year: dateYear(get("Date created")),
    });
  });

  // =========================
  // FILTERS
  // =========================
  // Tags.str.contains(r"(?i)^cyber$") — anchored both ends = whole cell,
  // case-insensitive, on the already-trimmed string. AND year == 2026.
  const filtered = df.filter((r) => /^cyber$/i.test(r.tags) && r.year === 2026);

  // =========================
  // STAGE LOGIC
  // =========================
  const closed_stages = new Set(["Completed", "Under review"]);
  for (const r of filtered) {
    r.finalStage = closed_stages.has(r.stage) ? "Closed" : "Open";
  }

  // =========================
  // ORG MAPPING
  // =========================
  const org_map = new Map([
    ["Africa", "AFR"],
    ["APAC", "APAC"],
    ["BEES", "GRO"],
    ["BEES | FINTECH", "GRO"],
    ["Europe", "Europe"],
    ["GHQ", "GHQ"],
    ["South America Zone", "SAZ"],
    ["North America Zone", "NAZ"],
    ["Middle America Zone", "MAZ"],
  ]);
  for (const r of filtered) {
    r.orgDisplay = org_map.has(r.organization) ? org_map.get(r.organization) : r.organization;
  }

  // =========================
  // PIVOT TABLE
  // =========================
  // pd.pivot_table(values="ID", aggfunc="count", fill_value=0): every org in
  // `filtered` gets a row, but each cell counts only rows whose ID is
  // non-null — a Cyber/2026 row with a blank ID contributes its org to the
  // index and 0 to the counts, while still counting toward the KPI totals
  // below (which use len(filtered)). That asymmetry is the script's.
  const counts = new Map();
  for (const r of filtered) {
    if (!counts.has(r.orgDisplay)) counts.set(r.orgDisplay, { Open: 0, Closed: 0 });
  }
  for (const r of filtered) {
    if (r.id === null || r.id === undefined || r.id === "") continue;
    counts.get(r.orgDisplay)[r.finalStage] += 1;
  }

  let pivot = [...counts.entries()].map(([org, c]) => ({
    rowLabels: org,
    open: c.Open,
    closed: c.Closed,
    grandTotal: c.Open + c.Closed,
  }));

  const order = ["AFR", "APAC", "GRO", "Europe", "GHQ", "SAZ", "MAZ", "NAZ"];
  const sortOrder = (label) => {
    const i = order.indexOf(label);
    return i === -1 ? 999 : i;
  };
  // sort_values(["sort_order", "Row Labels"]) — fixed order first, unknown
  // orgs last, alphabetical (code-point) among themselves.
  pivot.sort((a, b) => sortOrder(a.rowLabels) - sortOrder(b.rowLabels) || ordinal(a.rowLabels, b.rowLabels));

  const grand_total_row = {
    rowLabels: "Grand Total",
    open: pivot.reduce((s, r) => s + r.open, 0),
    closed: pivot.reduce((s, r) => s + r.closed, 0),
    grandTotal: pivot.reduce((s, r) => s + r.grandTotal, 0),
  };
  const pivot_display = [...pivot, grand_total_row];

  // =========================
  // KPI DATA
  // =========================
  const baseline_25 = 0.6;
  const q1_26 = 0.32;
  const target_26 = 0.65;

  const closed_total = filtered.filter((r) => r.finalStage === "Closed").length;
  const record_total = filtered.length;

  const q2_26 = record_total ? pyRound2(closed_total / record_total) : 0;

  const kpi_data = [
    ["Baseline '25", baseline_25, "static"],
    ["Q1 '26", q1_26, "static"],
    ["Q2 '26", q2_26, ""],
    ["Target '26", target_26, "static"],
  ];

  // =====================================================
  // CREATE CHART 1 — hand-drawn stacked bars on raw canvas
  // (matplotlib 7.6in x 4.4in @ 180 dpi -> 1368 x 792 px)
  // =====================================================
  drawChart1(pivot, closed_total, record_total);

  // =====================================================
  // CREATE CHART 2 — hand-drawn horizontal bars
  // (7.6in x 3.8in @ 180 dpi -> 1368 x 684 px)
  // =====================================================
  drawChart2(kpi_data);

  // =========================
  // CREATE EXCEL
  // =========================
  const outWb = new ExcelJS.Workbook();
  const ws = outWb.addWorksheet("Dashboard");

  // =========================
  // STYLES
  // =========================
  const header_fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB7DEE8" } };
  const subheader_fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAF7" } };
  const grand_fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB7DEE8" } };
  const bold = { bold: true };
  const thin = { style: "thin", color: { argb: "FF000000" } };
  const border = { left: thin, right: thin, top: thin, bottom: thin };
  const center = { horizontal: "center", vertical: "center" };

  // =========================
  // FILTER AREA
  // =========================
  ws.getCell("A1").value = "Tags";
  ws.getCell("B1").value = "Cyber";
  ws.getCell("A2").value = "Date created";
  ws.getCell("B2").value = "2026";
  for (const ref of ["A1", "A2"]) {
    ws.getCell(ref).fill = header_fill;
    ws.getCell(ref).font = bold;
    ws.getCell(ref).border = border;
  }
  for (const ref of ["B1", "B2"]) {
    ws.getCell(ref).fill = subheader_fill;
    ws.getCell(ref).border = border;
  }

  // =========================
  // PIVOT TABLE OUTPUT
  // =========================
  const start_row = 4;
  const headers = ["Row Labels", "Open", "Closed", "Grand Total"];
  headers.forEach((h, i) => {
    const c = ws.getCell(start_row, i + 1);
    c.value = h;
    c.fill = header_fill;
    c.font = bold;
    c.border = border;
    c.alignment = center;
  });

  pivot_display.forEach((row, i) => {
    const r = start_row + 1 + i;
    const values = [row.rowLabels, row.open, row.closed, row.grandTotal];
    values.forEach((value, j) => {
      const c = ws.getCell(r, j + 1);
      // openpyxl writes "" as an empty cell — a blank Organization on a kept
      // row becomes an empty Row Label there, so it must here too.
      c.value = value === "" ? null : value;
      c.border = border;
      if (j + 1 > 1) c.alignment = center; // numeric cells centered, labels not
    });
    if (row.rowLabels === "Grand Total") {
      for (let col = 1; col <= headers.length; col++) {
        ws.getCell(r, col).fill = grand_fill;
        ws.getCell(r, col).font = bold;
      }
    }
  });

  // =========================
  // KPI TABLE
  // =========================
  const kpi_start_row = start_row + pivot_display.length + 5;
  ["Metric", "Value", "Remark"].forEach((text, i) => {
    const c = ws.getCell(kpi_start_row, i + 1);
    c.value = text;
    c.fill = header_fill;
    c.font = bold;
    c.border = border;
    c.alignment = center;
  });

  kpi_data.forEach(([metric, value, remark], i) => {
    const r = kpi_start_row + 1 + i;
    ws.getCell(r, 1).value = metric;
    ws.getCell(r, 2).value = value;
    // openpyxl writes "" as an empty cell (reads back as None); exceljs would
    // keep the empty string. Null here so the Q2 remark cell round-trips the
    // same as the Python's.
    ws.getCell(r, 3).value = remark === "" ? null : remark;
    for (const col of [1, 2, 3]) ws.getCell(r, col).border = border;
    ws.getCell(r, 2).numFmt = "0%";
  });

  const note_row = kpi_start_row + kpi_data.length + 2;
  ws.getCell(note_row, 1).value = "Q2 formula";
  ws.getCell(note_row, 2).value = `Closed / Grand Total = ${closed_total} / ${record_total}`;
  ws.getCell(note_row, 1).font = bold;

  // =========================
  // INSERT CHART IMAGES
  // =========================
  // openpyxl anchors at "F2"/"F23" with explicit pixel sizes; exceljs takes
  // the same thing as a zero-based tl cell plus an ext in pixels.
  const img1 = outWb.addImage({ filename: chart1_file, extension: "png" });
  ws.addImage(img1, { tl: { col: 5, row: 1 }, ext: { width: 560, height: 360 }, editAs: "oneCell" });
  const img2 = outWb.addImage({ filename: chart2_file, extension: "png" });
  ws.addImage(img2, { tl: { col: 5, row: 22 }, ext: { width: 560, height: 300 }, editAs: "oneCell" });

  // =========================
  // FORMATTING
  // =========================
  const widths = { A: 18, B: 12, C: 12, D: 14, E: 4, F: 14, G: 14, H: 14, I: 14, J: 14, K: 14, L: 14, M: 14 };
  for (const [col, width] of Object.entries(widths)) {
    ws.getColumn(col).width = width;
  }
  for (let row = 1; row < 50; row++) {
    ws.getRow(row).height = 22;
  }
  ws.views = [{ state: "frozen", xSplit: 0, ySplit: 4, topLeftCell: "A5", activeCell: "A5" }];

  // =========================
  // SAVE FILE
  // =========================
  await outWb.xlsx.writeFile(output_file);

  console.log("SUCCESS");
  console.log(output_file);
}

// =====================================================
// CHART 1 — stacked vertical bars, black background
// =====================================================
const SANS = "sans-serif";
const open_color = "#00AEEF";
const closed_color = "#D4AF37";

function drawChart1(pivot, closed_total, record_total) {
  const W = 1368, H = 792;
  const canvas = createCanvas(W, H);
  const ctx = canvas.getContext("2d");

  // fig/ax facecolor black
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, W, H);

  // Title: "2026 Assessment\n(closed/total)", bold, white, centered
  ctx.fillStyle = "#ffffff";
  ctx.textAlign = "center";
  ctx.textBaseline = "alphabetic";
  ctx.font = `bold 36px ${SANS}`;
  ctx.fillText("2026 Assessment", W / 2, 66);
  ctx.fillText(`(${closed_total}/${record_total})`, W / 2, 112);

  // Legend: Closed first, Open second (legend_order = [1, 0]), frameless
  const legend = [["Closed", closed_color], ["Open", open_color]];
  ctx.font = `22px ${SANS}`;
  const sw = 26, sh = 15, gap = 10, spacing = 44;
  let legendW = 0;
  for (const [label] of legend) legendW += sw + gap + ctx.measureText(label).width;
  legendW += spacing * (legend.length - 1);
  let lx = (W - legendW) / 2;
  const ly = 152;
  for (const [label, color] of legend) {
    ctx.fillStyle = color;
    ctx.fillRect(lx, ly - sh + 3, sw, sh);
    lx += sw + gap;
    ctx.fillStyle = "#ffffff";
    ctx.textAlign = "left";
    ctx.fillText(label, lx, ly);
    lx += ctx.measureText(label).width + spacing;
  }

  // Plot area; ylim = 0 .. max_total * 1.35 headroom
  const left = 56, right = W - 44, top = 196, bottom = H - 148;
  const plotW = right - left;
  const n = pivot.length;
  const max_total = Math.max(1, ...pivot.map((r) => r.open + r.closed));
  const scaleMax = max_total * 1.35;
  const yOf = (v) => bottom - (v / scaleMax) * (bottom - top);

  const slot = plotW / Math.max(n, 1);
  const barW = slot * 0.48; // matplotlib width=0.48

  pivot.forEach((r, i) => {
    const cx = left + slot * i + slot / 2;
    const x = cx - barW / 2;

    // Open segment at the bottom (blue)
    if (r.open > 0) {
      const yTop = yOf(r.open);
      ctx.fillStyle = open_color;
      ctx.fillRect(x, yTop, barW, bottom - yTop);
    }
    // Closed segment stacked on top (gold) — bottom=open_values
    if (r.closed > 0) {
      const yBase = yOf(r.open);
      const yTop = yOf(r.open + r.closed);
      ctx.fillStyle = closed_color;
      ctx.fillRect(x, yTop, barW, yBase - yTop);
    }

    // Segment labels: white, bold, centered, only when value > 0
    ctx.fillStyle = "#ffffff";
    ctx.font = `bold 21px ${SANS}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    if (r.open > 0) ctx.fillText(String(r.open), cx, (bottom + yOf(r.open)) / 2);
    if (r.closed > 0) ctx.fillText(String(r.closed), cx, (yOf(r.open) + yOf(r.open + r.closed)) / 2);

    // X tick labels: rotation=45, ha="right", white
    ctx.save();
    ctx.translate(cx + 4, bottom + 18);
    ctx.rotate(-Math.PI / 4);
    ctx.font = `22px ${SANS}`;
    ctx.textAlign = "right";
    ctx.textBaseline = "top";
    ctx.fillText(r.rowLabels, 0, 0);
    ctx.restore();
    ctx.textBaseline = "alphabetic";
  });

  // Bottom spine / axhline(0), white; no other spines, no gridlines, no y axis
  ctx.strokeStyle = "#ffffff";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(left - 6, bottom);
  ctx.lineTo(right + 6, bottom);
  ctx.stroke();

  fs.writeFileSync(chart1_file, canvas.toBuffer("image/png"));
}

// =====================================================
// CHART 2 — horizontal KPI bars, white background
// =====================================================
function drawChart2(kpi_data) {
  const W = 1368, H = 684;
  const canvas = createCanvas(W, H);
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, W, H);

  // Title
  ctx.fillStyle = "#000000";
  ctx.textAlign = "center";
  ctx.textBaseline = "alphabetic";
  ctx.font = `bold 32px ${SANS}`;
  ctx.fillText("Improve in Supplier Response Time", W / 2, 58);

  const left = 214, right = W - 132, top = 104, bottom = H - 86;
  const plotW = right - left, plotH = bottom - top;
  const xOf = (v) => left + v * plotW;

  // Dashed x gridlines at 0/25/50/75/100%, drawn BEFORE the bars (behind)
  ctx.font = `20px ${SANS}`;
  for (let i = 0; i <= 4; i++) {
    const v = i * 0.25;
    const gx = xOf(v);
    ctx.save();
    ctx.strokeStyle = "rgba(0,0,0,0.35)";
    ctx.setLineDash([8, 8]);
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(gx, top);
    ctx.lineTo(gx, bottom);
    ctx.stroke();
    ctx.restore();

    ctx.fillStyle = "#333333";
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(`${Math.round(v * 100)}%`, gx, bottom + 12);
  }

  // Left + bottom spines only (top/right hidden)
  ctx.strokeStyle = "#333333";
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, bottom);
  ctx.lineTo(right, bottom);
  ctx.stroke();

  // barh, height=0.48 of slot, inverted y-axis (first metric on top)
  const n = kpi_data.length;
  const slot = plotH / Math.max(n, 1);
  const barH = slot * 0.48;

  kpi_data.forEach(([metric, value], i) => {
    const yc = top + slot * i + slot / 2;
    const y = yc - barH / 2;
    const wpx = Math.max(0, Math.min(1, value)) * plotW;

    ctx.fillStyle = "#4472C4";
    ctx.fillRect(left, y, wpx, barH);

    // Metric names as left-side labels
    ctx.fillStyle = "#000000";
    ctx.font = `22px ${SANS}`;
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    ctx.fillText(metric, left - 16, yc);

    // Value labels: f"{value:.0%}", bold, just right of the bar end
    ctx.font = `bold 22px ${SANS}`;
    ctx.textAlign = "left";
    ctx.fillText(`${Math.round(value * 100)}%`, left + wpx + 12, yc);
  });

  fs.writeFileSync(chart2_file, canvas.toBuffer("image/png"));
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : String(err));
  process.exit(1);
});
