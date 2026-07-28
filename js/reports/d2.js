// Port of "daigram 2 automation/automation.py" (id: d2), running in the app.
//
// Functional parity with the Python script is the contract: same filters
// (Tags whole-cell (?i)^cyber$ AND Date created year == 2026), same pivot
// semantics (pandas counts non-null IDs, not rows), same KPI rounding
// (Python's half-to-even round(x, 2)), same hand-drawn chart geometry
// embedded as PNGs at F2/F23, same Excel styling, same output filename.
// Where pandas has a quirk, the quirk is reproduced and commented rather
// than "improved". Verified cell-for-cell against the unmodified script.
window.Reports.d2 = async function d2(wb) {
  var E = window.ReportEngine;

  // ── pandas equivalents, hand-written ──────────────────────────────────────

  // pandas read_excel's default NA sentinels (keep_default_na=True): these
  // exact strings read as NaN, so an ID of "N/A" is NOT counted by the pivot
  // and an Organization of "NULL" becomes "" via fillna. Exact match on the
  // raw cell text, no trimming — that is how pandas applies them.
  var NA_TOKENS = {
    '': 1, '#N/A': 1, '#N/A N/A': 1, '#NA': 1, '-1.#IND': 1, '-1.#QNAN': 1,
    '-NaN': 1, '-nan': 1, '1.#IND': 1, '1.#QNAN': 1, '<NA>': 1, 'N/A': 1,
    'NA': 1, 'NULL': 1, 'NaN': 1, 'None': 1, 'n/a': 1, 'nan': 1, 'null': 1,
  };
  function naSieve(v) {
    return typeof v === 'string' && NA_TOKENS.hasOwnProperty(v) ? null : v;
  }

  // df[c].fillna("").astype(str).str.strip() — NaN/None become "", everything
  // else is stringified then trimmed. Python spells booleans True/False.
  function pyStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' && Number.isNaN(v)) return '';
    if (v === true) return 'True';
    if (v === false) return 'False';
    if (v instanceof Date) {
      var p = function (n) { return ('0' + n).slice(-2); };
      return v.getUTCFullYear() + '-' + p(v.getUTCMonth() + 1) + '-' + p(v.getUTCDate()) +
        ' ' + p(v.getUTCHours()) + ':' + p(v.getUTCMinutes()) + ':' + p(v.getUTCSeconds());
    }
    return String(v).trim();
  }

  // pd.to_datetime(..., errors="coerce").dt.year equivalent, returning the
  // year or null. Handles a JS Date (UTC getters — never local time), an
  // Excel serial number, or text in the common export shapes. Real calendar
  // validation: pandas coerces "2026-02-31" to NaT, so a d<=31 shortcut would
  // let impossible dates through the year filter. No bare-year branch: a text
  // cell holding just "2026" is type-inferred to int64 by pandas.read_excel
  // BEFORE to_datetime runs, so the script reads epoch-nanoseconds (year
  // 1970) and drops the row — parsing it here would diverge from the script.
  var MONTH_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function ymdYear(y, mo, d) {
    if (mo < 1 || mo > 12 || d < 1) return null;
    var leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
    var dim = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    return d <= dim[mo - 1] ? y : null;
  }
  function dateYear(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v.getUTCFullYear();
    if (typeof v === 'number') {
      if (!isFinite(v)) return null;
      // Excel 1900 date system: serial 25569 = 1970-01-01 UTC.
      return new Date(Math.round((v - 25569) * 86400000)).getUTCFullYear();
    }
    var s = String(v).trim();
    if (!s) return null;
    var m;
    // Every pattern is end-guarded (end, whitespace, or a time part): pandas
    // coerces "2026-12-14xyz" to NaT, so a bare prefix match must not count.
    m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?=$|[T\s])/);
    if (m) return ymdYear(+m[1], +m[2], +m[3]);
    // Day/month numeric: pandas defaults to dayfirst=False, so a leading
    // value that COULD be a month is read as one; above 12 it is the day.
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?=$|[T\s])/);
    if (m) { return +m[1] <= 12 ? ymdYear(+m[3], +m[1], +m[2]) : ymdYear(+m[3], +m[2], +m[1]); }
    // Named month: 14-Dec-2026, 14 Dec 2026, 14 December 2026
    m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s,]+(\d{4})(?=$|[T\s])/);
    if (m) { var mo1 = MONTH_NUM[m[2].slice(0, 3).toLowerCase()]; return mo1 ? ymdYear(+m[3], mo1, +m[1]) : null; }
    // Month first: Dec 14, 2026 / December 14 2026
    m = s.match(/^([A-Za-z]{3,})[-\s](\d{1,2})[-\s,]+(\d{4})(?=$|[T\s])/);
    if (m) { var mo2 = MONTH_NUM[m[1].slice(0, 3).toLowerCase()]; return mo2 ? ymdYear(+m[3], mo2, +m[2]) : null; }
    // Month + year: "Dec 2026" -> pandas reads 2026-12-01
    m = s.match(/^([A-Za-z]{3,})[-\s,]+(\d{4})$/);
    if (m) { return MONTH_NUM[m[1].slice(0, 3).toLowerCase()] ? +m[2] : null; }
    return null;
  }

  // Python's round(x, 2) decides from the EXACT binary value of the double
  // and resolves true ties half-to-even. Multiplying by 100 first is wrong —
  // the multiply itself rounds (1/40 = 0.025 stored as 0.02500000000000000138…
  // becomes 2.4999999999999996, flipping Python's 0.03 to 0.02). Read the
  // double's correctly-rounded 20-place decimal expansion instead and compare
  // the tail beyond 2 places against an exact "500…0". Verified against
  // CPython for every n/d with d <= 200 (20,310 values, zero mismatches).
  function pyRound2(x) {
    var s = x.toFixed(20);
    var dot = s.indexOf('.');
    var dec = s.slice(dot + 1);
    var k = Number(s.slice(0, dot)) * 100 + Number(dec.slice(0, 2));
    var tail = dec.slice(2);
    var mid = '5';
    while (mid.length < tail.length) mid += '0';
    if (tail > mid) k += 1;
    else if (tail === mid && k % 2 !== 0) k += 1; // true tie: half-to-even
    return k / 100;
  }

  // Python sorts strings by code POINT; JS's < compares UTF-16 code units,
  // which reverses the order when an astral-plane character meets a high BMP
  // character. Compare true code points, like Python does.
  function ordinal(a, b) {
    var A = Array.from(String(a)), B = Array.from(String(b));
    var n = Math.min(A.length, B.length);
    for (var i = 0; i < n; i++) {
      var d = A[i].codePointAt(0) - B[i].codePointAt(0);
      if (d !== 0) return d < 0 ? -1 : 1;
    }
    return A.length === B.length ? 0 : (A.length < B.length ? -1 : 1);
  }

  // ── LOAD DATA ─────────────────────────────────────────────────────────────
  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers;

  // sheet_to_json dedupes repeated keys by occurrence: the 2nd "Tags" becomes
  // "Tags_1". Rebuild that mapping so each header position has its row key.
  var seenKey = {}, rowKeys = [];
  rawHeaders.forEach(function (h) {
    var n = seenKey[h] || 0;
    seenKey[h] = n + 1;
    rowKeys.push(n === 0 ? h : h + '_' + n);
  });

  // df.columns.str.strip(). pandas mangles EXACT duplicate headers at read
  // time ("Tags.1"), so df["Tags"] is always the FIRST occurrence — same
  // here. Headers that collide only AFTER stripping ("Tags" + "Tags ") are
  // not mangled: both become "Tags", df["Tags"] returns a DataFrame, and the
  // script dies on .str before writing anything — matched with a readable
  // error instead of a silent wrong number.
  var colKey = {}, rawsByName = {};
  rawHeaders.forEach(function (h, i) {
    var name = pyStr(h);
    if (!name) return;
    if (!rawsByName[name]) rawsByName[name] = [];
    if (rawsByName[name].indexOf(String(h)) === -1) rawsByName[name].push(String(h));
    if (!(name in colKey)) colKey[name] = rowKeys[i];
  });

  var REQUIRED = ['ID', 'Organization', 'Stage', 'Date created', 'Tags'];
  var missing = REQUIRED.filter(function (c) { return !(c in colKey); });
  if (missing.length) {
    // Python: raise ValueError(f"Missing required columns: {missing}")
    throw new Error('Missing required columns: [' + missing.map(function (c) { return "'" + c + "'"; }).join(', ') + ']');
  }
  REQUIRED.forEach(function (c) {
    if (rawsByName[c] && rawsByName[c].length > 1) {
      throw new Error("Duplicate column '" + c + "' after header strip — the Python script aborts here too (df['" + c + "'] becomes a DataFrame)");
    }
  });

  var df = sheet.rows.map(function (row) {
    var get = function (name) { return naSieve(row[colKey[name]]); };
    return {
      // ID stays raw: the pivot's aggfunc="count" counts NON-NULL IDs, so a
      // blank (or NA-sentinel) ID must remain distinguishable from a value.
      id: get('ID'),
      organization: pyStr(get('Organization')),
      stage: pyStr(get('Stage')),
      tags: pyStr(get('Tags')),
      year: dateYear(get('Date created')),
    };
  });

  // ── FILTERS ───────────────────────────────────────────────────────────────
  // Tags.str.contains(r"(?i)^cyber$") — anchored both ends = whole cell,
  // case-insensitive, on the already-trimmed string. AND year == 2026.
  var tagMatched = df.filter(function (r) { return /^cyber$/i.test(r.tags); });
  var yearMatched = df.filter(function (r) { return r.year === 2026; });
  var filtered = tagMatched.filter(function (r) { return r.year === 2026; });

  // ── STAGE LOGIC ───────────────────────────────────────────────────────────
  filtered.forEach(function (r) {
    r.finalStage = (r.stage === 'Completed' || r.stage === 'Under review') ? 'Closed' : 'Open';
  });

  // ── ORG MAPPING ───────────────────────────────────────────────────────────
  var ORG_MAP = {
    'Africa': 'AFR', 'APAC': 'APAC', 'BEES': 'GRO', 'BEES | FINTECH': 'GRO',
    'Europe': 'Europe', 'GHQ': 'GHQ', 'South America Zone': 'SAZ',
    'North America Zone': 'NAZ', 'Middle America Zone': 'MAZ',
  };
  filtered.forEach(function (r) {
    r.orgDisplay = ORG_MAP.hasOwnProperty(r.organization) ? ORG_MAP[r.organization] : r.organization;
  });

  // ── PIVOT TABLE ───────────────────────────────────────────────────────────
  // pd.pivot_table(values="ID", aggfunc="count", fill_value=0): every org in
  // `filtered` gets a row, but each cell counts only rows whose ID is
  // non-null — a Cyber/2026 row with a blank ID contributes its org to the
  // index and 0 to the counts, while still counting toward the KPI totals
  // below (which use len(filtered)). That asymmetry is the script's.
  var counts = {}, orgSeen = [];
  filtered.forEach(function (r) {
    if (!counts.hasOwnProperty(r.orgDisplay)) { counts[r.orgDisplay] = { Open: 0, Closed: 0 }; orgSeen.push(r.orgDisplay); }
  });
  filtered.forEach(function (r) {
    if (r.id === null || r.id === undefined || r.id === '') return;
    counts[r.orgDisplay][r.finalStage] += 1;
  });

  var pivot = orgSeen.map(function (org) {
    var c = counts[org];
    return { rowLabels: org, open: c.Open, closed: c.Closed, grandTotal: c.Open + c.Closed };
  });

  var ORDER = ['AFR', 'APAC', 'GRO', 'Europe', 'GHQ', 'SAZ', 'MAZ', 'NAZ'];
  function sortOrder(label) { var i = ORDER.indexOf(label); return i === -1 ? 999 : i; }
  // sort_values(["sort_order", "Row Labels"]) — fixed order first, unknown
  // orgs last, alphabetical (code-point) among themselves.
  pivot.sort(function (a, b) {
    return sortOrder(a.rowLabels) - sortOrder(b.rowLabels) || ordinal(a.rowLabels, b.rowLabels);
  });

  var grandTotalRow = {
    rowLabels: 'Grand Total',
    open: pivot.reduce(function (s, r) { return s + r.open; }, 0),
    closed: pivot.reduce(function (s, r) { return s + r.closed; }, 0),
    grandTotal: pivot.reduce(function (s, r) { return s + r.grandTotal; }, 0),
  };
  var pivotDisplay = pivot.concat([grandTotalRow]);

  // ── KPI DATA ──────────────────────────────────────────────────────────────
  var closed_total = filtered.filter(function (r) { return r.finalStage === 'Closed'; }).length;
  var record_total = filtered.length;
  var q2_26 = record_total ? pyRound2(closed_total / record_total) : 0;

  var KPI = [
    ["Baseline '25", 0.60, 'static'],
    ["Q1 '26", 0.32, 'static'],
    ["Q2 '26", q2_26, ''],
    ["Target '26", 0.65, 'static'],
  ];

  // ── CHARTS — hand-drawn on canvas, exactly the script's matplotlib geometry
  // Chart 1: 7.6in x 4.4in @ 180 dpi -> 1368 x 792. Chart 2: 1368 x 684.
  function makeCanvas(w, h) {
    var c = document.createElement('canvas');
    c.width = w; c.height = h;
    return c;
  }
  var OPEN_COLOR = '#00AEEF', CLOSED_COLOR = '#D4AF37';

  function drawChart1() {
    var W = 1368, H = 792;
    var canvas = makeCanvas(W, H);
    var ctx = canvas.getContext('2d');
    ctx.fillStyle = '#000000';
    ctx.fillRect(0, 0, W, H);

    // Title: "2026 Assessment\n(closed/total)", bold, white, centered
    ctx.fillStyle = '#ffffff';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'alphabetic';
    ctx.font = 'bold 36px sans-serif';
    ctx.fillText('2026 Assessment', W / 2, 66);
    ctx.fillText('(' + closed_total + '/' + record_total + ')', W / 2, 112);

    // Legend: Closed first, Open second (legend_order = [1, 0]), frameless
    var legend = [['Closed', CLOSED_COLOR], ['Open', OPEN_COLOR]];
    ctx.font = '22px sans-serif';
    var sw = 26, sh = 15, gap = 10, spacing = 44;
    var legendW = 0;
    legend.forEach(function (l) { legendW += sw + gap + ctx.measureText(l[0]).width; });
    legendW += spacing * (legend.length - 1);
    var lx = (W - legendW) / 2, ly = 152;
    legend.forEach(function (l) {
      ctx.fillStyle = l[1];
      ctx.fillRect(lx, ly - sh + 3, sw, sh);
      lx += sw + gap;
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'left';
      ctx.fillText(l[0], lx, ly);
      lx += ctx.measureText(l[0]).width + spacing;
    });

    // Plot area; ylim = 0 .. max_total * 1.35 headroom
    var left = 56, right = W - 44, top = 196, bottom = H - 148;
    var n = pivot.length;
    var maxTotal = Math.max.apply(null, [1].concat(pivot.map(function (r) { return r.open + r.closed; })));
    var scaleMax = maxTotal * 1.35;
    var yOf = function (v) { return bottom - (v / scaleMax) * (bottom - top); };
    var slot = (right - left) / Math.max(n, 1);
    var barW = slot * 0.48; // matplotlib width=0.48

    pivot.forEach(function (r, i) {
      var cx = left + slot * i + slot / 2;
      var x = cx - barW / 2;
      if (r.open > 0) {
        ctx.fillStyle = OPEN_COLOR;
        ctx.fillRect(x, yOf(r.open), barW, bottom - yOf(r.open));
      }
      if (r.closed > 0) {
        ctx.fillStyle = CLOSED_COLOR;
        ctx.fillRect(x, yOf(r.open + r.closed), barW, yOf(r.open) - yOf(r.open + r.closed));
      }
      // Segment labels: white, bold, centered, only when value > 0
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 21px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      if (r.open > 0) ctx.fillText(String(r.open), cx, (bottom + yOf(r.open)) / 2);
      if (r.closed > 0) ctx.fillText(String(r.closed), cx, (yOf(r.open) + yOf(r.open + r.closed)) / 2);
      // X tick labels: rotation=45, ha="right", white
      ctx.save();
      ctx.translate(cx + 4, bottom + 18);
      ctx.rotate(-Math.PI / 4);
      ctx.font = '22px sans-serif';
      ctx.textAlign = 'right';
      ctx.textBaseline = 'top';
      ctx.fillText(r.rowLabels, 0, 0);
      ctx.restore();
      ctx.textBaseline = 'alphabetic';
    });

    // Bottom spine / axhline(0), white; no other spines, no y axis
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(left - 6, bottom);
    ctx.lineTo(right + 6, bottom);
    ctx.stroke();

    return canvas.toDataURL('image/png');
  }

  function drawChart2() {
    var W = 1368, H = 684;
    var canvas = makeCanvas(W, H);
    var ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    ctx.fillStyle = '#000000';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'alphabetic';
    ctx.font = 'bold 32px sans-serif';
    ctx.fillText('Improve in Supplier Response Time', W / 2, 58);

    var left = 214, right = W - 132, top = 104, bottom = H - 86;
    var plotW = right - left, plotH = bottom - top;
    var xOf = function (v) { return left + v * plotW; };

    // Dashed x gridlines at 0/25/50/75/100%, drawn BEFORE the bars (behind)
    ctx.font = '20px sans-serif';
    for (var i = 0; i <= 4; i++) {
      var gx = xOf(i * 0.25);
      ctx.save();
      ctx.strokeStyle = 'rgba(0,0,0,0.35)';
      ctx.setLineDash([8, 8]);
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(gx, top);
      ctx.lineTo(gx, bottom);
      ctx.stroke();
      ctx.restore();
      ctx.fillStyle = '#333333';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      ctx.fillText(Math.round(i * 25) + '%', gx, bottom + 12);
    }

    // Left + bottom spines only (top/right hidden)
    ctx.strokeStyle = '#333333';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(left, top);
    ctx.lineTo(left, bottom);
    ctx.lineTo(right, bottom);
    ctx.stroke();

    // barh, height=0.48 of slot, inverted y-axis (first metric on top)
    var slot = plotH / Math.max(KPI.length, 1);
    var barH = slot * 0.48;
    KPI.forEach(function (k, i) {
      var metric = k[0], value = k[1];
      var yc = top + slot * i + slot / 2;
      var wpx = Math.max(0, Math.min(1, value)) * plotW;
      ctx.fillStyle = '#4472C4';
      ctx.fillRect(left, yc - barH / 2, wpx, barH);
      // Metric names as left-side labels
      ctx.fillStyle = '#000000';
      ctx.font = '22px sans-serif';
      ctx.textAlign = 'right';
      ctx.textBaseline = 'middle';
      ctx.fillText(metric, left - 16, yc);
      // Value labels: f"{value:.0%}", bold, just right of the bar end
      ctx.font = 'bold 22px sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText(Math.round(value * 100) + '%', left + wpx + 12, yc);
    });

    return canvas.toDataURL('image/png');
  }

  var chart1Png = drawChart1();
  var chart2Png = drawChart2();

  // ── CREATE EXCEL — the openpyxl block, one for one ───────────────────────
  // The grid mirrors every written cell so the in-app preview shows exactly
  // what the workbook holds (row 0 is re-treated as headers downstream).
  function setCell(grid, row1, col1, value) {
    var r = row1 - 1, c = col1 - 1;
    while (grid.length <= r) grid.push([]);
    while (grid[r].length <= c) grid[r].push(null);
    grid[r][c] = value;
  }
  var grid = [];
  setCell(grid, 1, 1, 'Tags'); setCell(grid, 1, 2, 'Cyber');
  setCell(grid, 2, 1, 'Date created'); setCell(grid, 2, 2, '2026');

  var startRow = 4;
  ['Row Labels', 'Open', 'Closed', 'Grand Total'].forEach(function (h, i) { setCell(grid, startRow, i + 1, h); });
  pivotDisplay.forEach(function (r, i) {
    var rowNum = startRow + 1 + i;
    // openpyxl writes "" as an empty cell — a blank Organization on a kept
    // row becomes an empty Row Label there, so it must here too.
    setCell(grid, rowNum, 1, r.rowLabels === '' ? null : r.rowLabels);
    setCell(grid, rowNum, 2, r.open);
    setCell(grid, rowNum, 3, r.closed);
    setCell(grid, rowNum, 4, r.grandTotal);
  });

  var kpiStartRow = startRow + pivotDisplay.length + 5;
  ['Metric', 'Value', 'Remark'].forEach(function (h, i) { setCell(grid, kpiStartRow, i + 1, h); });
  KPI.forEach(function (k, i) {
    var rowNum = kpiStartRow + 1 + i;
    setCell(grid, rowNum, 1, k[0]);
    setCell(grid, rowNum, 2, k[1]);
    setCell(grid, rowNum, 3, k[2] === '' ? null : k[2]);
  });
  var noteRow = kpiStartRow + KPI.length + 2;
  setCell(grid, noteRow, 1, 'Q2 formula');
  setCell(grid, noteRow, 2, 'Closed / Grand Total = ' + closed_total + ' / ' + record_total);

  var workbook = new ExcelJS.Workbook();
  var ws = workbook.addWorksheet('Dashboard');
  grid.forEach(function (r) { ws.addRow(r); });

  var HEADER_FILL = 'FFB7DEE8', SUBHEADER_FILL = 'FFD9EAF7';
  var THIN = { style: 'thin', color: { argb: 'FF000000' } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  var CENTER = { horizontal: 'center', vertical: 'middle' };
  function style(row, col, opts) {
    var cell = ws.getCell(row, col);
    if (opts.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
    if (opts.bold) cell.font = { bold: true };
    if (opts.border) cell.border = BORDER;
    if (opts.center) cell.alignment = CENTER;
    if (opts.numFmt) cell.numFmt = opts.numFmt;
  }

  // Filter area
  [1, 2].forEach(function (r) {
    style(r, 1, { fill: HEADER_FILL, bold: true, border: true });
    style(r, 2, { fill: SUBHEADER_FILL, border: true });
  });
  // Pivot table: header row, bordered body (numeric cells centered), Grand
  // Total row re-filled and bolded on top.
  for (var pc = 1; pc <= 4; pc++) style(startRow, pc, { fill: HEADER_FILL, bold: true, border: true, center: true });
  pivotDisplay.forEach(function (r, i) {
    var rowNum = startRow + 1 + i;
    for (var c = 1; c <= 4; c++) {
      style(rowNum, c, { border: true, center: c > 1 });
      if (r.rowLabels === 'Grand Total') style(rowNum, c, { fill: HEADER_FILL, bold: true });
    }
  });
  // KPI table
  for (var kc = 1; kc <= 3; kc++) style(kpiStartRow, kc, { fill: HEADER_FILL, bold: true, border: true, center: true });
  KPI.forEach(function (_, i) {
    var rowNum = kpiStartRow + 1 + i;
    for (var c2 = 1; c2 <= 3; c2++) style(rowNum, c2, { border: true });
    style(rowNum, 2, { numFmt: '0%' });
  });
  style(noteRow, 1, { bold: true });

  // Chart images at F2 (560x360) and F23 (560x300), like the script's
  // openpyxl anchors — exceljs takes a zero-based tl cell + ext in pixels.
  var img1 = workbook.addImage({ base64: chart1Png.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
  ws.addImage(img1, { tl: { col: 5, row: 1 }, ext: { width: 560, height: 360 }, editAs: 'oneCell' });
  var img2 = workbook.addImage({ base64: chart2Png.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
  ws.addImage(img2, { tl: { col: 5, row: 22 }, ext: { width: 560, height: 300 }, editAs: 'oneCell' });

  // Column widths A-M, row heights 1..49 (range(1, 50) — not "all rows"),
  // freeze_panes = "A5".
  [18, 12, 12, 14, 4, 14, 14, 14, 14, 14, 14, 14, 14].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
  for (var rh = 1; rh < 50; rh++) ws.getRow(rh).height = 22;
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 4, topLeftCell: 'A5', activeCell: 'A5' }];

  var buf = await workbook.xlsx.writeBuffer();

  // ── Run summary (stdout — rendered in the Script log, never in the sheet) ─
  function distinct(values) {
    var seen = [], countsBy = {};
    values.forEach(function (v) {
      var k = v === '' ? '(blank)' : String(v);
      if (!countsBy[k]) { countsBy[k] = 0; seen.push(k); }
      countsBy[k]++;
    });
    if (!seen.length) return '(none)';
    seen.sort(function (a, b) { return countsBy[b] - countsBy[a]; });
    return seen.map(function (k) { return k + ' x' + countsBy[k]; }).join('  ·  ');
  }
  var lines = [
    'SUCCESS',
    'output file D2.xlsx',
    '',
    'Input: ' + ((wb && wb.__fileName) || 'workbook.xlsx') + '   ·   ' + df.length + ' data rows',
    'Filter: Tags = "cyber" (whole cell, case-insensitive)  AND  Date created year = 2026',
    '  Tags matched:  ' + tagMatched.length,
    '  Year matched:  ' + yearMatched.length,
    '  Kept (both):   ' + record_total,
    'Stage (exact case {Completed, Under review} -> Closed):',
    '  Closed: ' + closed_total + '   Open: ' + (record_total - closed_total),
    "Q2 '26 = Closed / Grand Total = " + closed_total + ' / ' + record_total + ' = ' + Math.round(q2_26 * 100) + '%',
  ];
  if (!record_total) {
    lines.push(
      '',
      '  ^ this 0% is the filter matching nothing, not a measurement.',
      '  Tags seen:  ' + distinct(df.map(function (r) { return r.tags; })),
      '  Years seen: ' + distinct(df.map(function (r) { return r.year === null ? 'unreadable' : r.year; }))
    );
  } else if (record_total !== grandTotalRow.grandTotal) {
    lines.push('  note: ' + (record_total - grandTotalRow.grandTotal) +
      ' kept row(s) have a blank ID and are missing from the pivot counts' +
      ' (the script divides by rows, the pivot counts IDs).');
  }

  return {
    ok: true,
    stdout: lines.join('\n'),
    files: [{ name: 'output file D2.xlsx', bytes: buf, sheets: [{ name: 'Dashboard', grid: grid }] }],
    chartImages: { 'Dashboard': chart1Png },
  };
};
