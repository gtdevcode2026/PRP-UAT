// Port of "diagram4/automation.py" (id: s4, "Risk Dashboard").
// Reads "OneTrust - Risk Export", computes Open/Closed/Total risk counts
// (Open = Evaluation+Identified+Treatment stages, Closed = Monitoring stage,
// year-filtered on Date closed when present), derives a period label from the
// latest ELAPSED 2026 date in Date created/Date closed, builds a
// Baseline'25 -> last-3-periods -> Target'26 matrix, and writes
// "Risk_Output.xlsx" (sheets "Dashboard" + "History Used") and "History.xlsx".
//
// History now carries forward between runs — see window.S4History below.

// ── History store ───────────────────────────────────────────────────────────
// The script treats History.xlsx as "the historical data source for future
// executions" (spec steps 15/16/20): it reads the file back on the next run so
// periods accumulate. This app has no filesystem to read from — output files
// are download blobs — so the browser's own storage stands in for the script's
// working directory.
//
// Everything here is best-effort by design. Opened from file://, localStorage
// can throw on access (opaque origin, private browsing, storage disabled by
// policy). Every call is guarded, and when storage is unavailable the report
// behaves exactly as it did before this existed — seed fresh every run — and
// says so in its execution summary rather than pretending it saved anything.
window.S4History = (function () {
  'use strict';
  // Versioned: bump when a change makes rows written by a previous build
  // untrustworthy. Fixing how a period is DERIVED does not unsay a bad label
  // that was already saved — a wrong month keeps being charted from storage
  // long after the code that produced it is gone, which is exactly how a "Dec
  // '26" and a "Jul '26" each outlived their fix. Nothing distinguishes a
  // poisoned row from a good one after the fact, so the whole generation is
  // retired rather than asking every user to find the Reset button.
  //   v1 — clock-derived period ("Jul '26" beside May data)
  //   v2 — max()-date period (one scheduled December date named the report)
  var KEY = 'prp.s4.history.v3';
  var STALE_KEYS = ['prp.s4.history.v1', 'prp.s4.history.v2'];

  function store() {
    try {
      var s = window.localStorage;
      if (!s) return null;
      // Private mode can expose localStorage but throw on write, so probe it.
      var probe = KEY + '.probe';
      s.setItem(probe, '1');
      s.removeItem(probe);
      return s;
    } catch (e) { return null; }
  }
  function load() {
    var s = store();
    if (!s) return null;
    try {
      // Retired generations are dead weight, not fallbacks — drop them on sight
      // so a later key bump can never resurrect one.
      STALE_KEYS.forEach(function (k) { s.removeItem(k); });
      var raw = s.getItem(KEY);
      if (!raw) return null;
      var rows = JSON.parse(raw);
      var valid = Array.isArray(rows) && rows.length && rows.every(function (r) {
        return r && typeof r === 'object' && r.Month;
      });
      return valid ? rows : null;   // corrupt payload falls back to the seed
    } catch (e) { return null; }
  }
  function save(rows) {
    var s = store();
    if (!s) return false;
    try { s.setItem(KEY, JSON.stringify(rows)); return true; } catch (e) { return false; }
  }
  function clear() {
    var s = store();
    if (!s) return false;
    try {
      s.removeItem(KEY);
      STALE_KEYS.forEach(function (k) { s.removeItem(k); });
      return true;
    } catch (e) { return false; }
  }
  function count() { var r = load(); return r ? r.length : 0; }
  return { KEY: KEY, load: load, save: save, clear: clear, count: count, available: function () { return !!store(); } };
})();

window.Reports.s4 = async function s4(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  // ── Configuration (spec step 2: every business constant in one place) ──
  var SHEET_NAME = 'OneTrust - Risk Export';
  var OUTPUT_FILE = 'Risk_Output.xlsx';
  var HISTORY_FILE = 'History.xlsx';
  var DASH_SHEET = 'Dashboard', HIST_USED_SHEET = 'History Used', HIST_SHEET = 'Sheet1';
  var YEAR = 2026;
  var OPEN_STAGES = { 'Evaluation': 1, 'Identified': 1, 'Treatment': 1 };
  var CLOSED_STAGE = 'Monitoring';
  var BASELINE_LABEL = "Baseline '25";
  var BASELINE_OPEN_RISK = 536, BASELINE_CLOSED_RISK = 42, BASELINE_TOTAL_RISK = 578;
  var TARGET_LABEL = "Target '26";
  var TARGET_PERCENT = 0.80;
  // A month must hold at least this share of the year's dated rows before it can
  // name the reporting period — see getPeriodLabel. Mirrors PERIOD_MIN_SHARE in
  // diagram4/automation.py; change both together.
  var PERIOD_MIN_SHARE = 0.05;
  var SEED_HISTORY_WHEN_MISSING = true;
  var SEED_ROWS = [
    { Month: "Q1 '26", 'Open risk as on date': 655, 'Closed Risk in 2026': 41, 'Total Risk': 696, 'Risk Created in 2026': 118 },
    { Month: "Apr '26", 'Open risk as on date': 694, 'Closed Risk in 2026': 42, 'Total Risk': 736, 'Risk Created in 2026': 158 },
  ];
  var MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // clean_column_name / normalize_stage — remove \r and \n, then strip. That is
  // all the script does, so that is all this does: a widened cleaner would make
  // this report count rows the reference output does not.
  function clean(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' && Number.isNaN(v)) return '';
    return String(v).replace(/[\r\n]/g, '').trim();
  }

  // Name the sheet AND what the workbook actually contains — "not found" on its
  // own gives the user nothing to act on, and the script lists them too.
  if (wb && Array.isArray(wb.SheetNames) && wb.SheetNames.indexOf(SHEET_NAME) === -1) {
    throw new Error("Sheet '" + SHEET_NAME + "' not found. Available sheets: " + wb.SheetNames.join(', '));
  }
  var sheet = E.readSheet(wb, SHEET_NAME);
  var cleanedHeaders = sheet.headers.map(clean);

  // Two columns can carry the same header — an export that gained a replacement
  // "Stage" keeps the original beside it, usually empty. sheet_to_json dedupes
  // the row KEYS by occurrence (Stage, Stage_1), but re-keying by the cleaned
  // header collapsed them so only the FIRST was ever readable; a populated
  // second column was invisible and every stage lookup missed, with no error.
  // Rebuild that key mapping and keep the first non-blank value per row.
  var seenHeader = {}, rowKeys = [], dupHeaders = [];
  sheet.headers.forEach(function (h) {
    var n = seenHeader[h] || 0;
    seenHeader[h] = n + 1;
    rowKeys.push(n === 0 ? h : h + '_' + n);
  });
  Object.keys(seenHeader).forEach(function (h) {
    if (seenHeader[h] > 1) dupHeaders.push(clean(h) + ' x' + seenHeader[h]);
  });

  function blankCell(v) { return v === undefined || v === null || v === ''; }
  var rows = sheet.rows.map(function (row) {
    var out = {};
    rowKeys.forEach(function (key, i) {
      var name = cleanedHeaders[i];
      var v = row[key];
      if (v === undefined) v = row[sheet.headers[i]];
      if (!(name in out) || (blankCell(out[name]) && !blankCell(v))) out[name] = v;
    });
    return out;
  });

  // find_required_column: {clean(col).lower(): col} and an exact lookup. Not a
  // fuzzy match — an earlier version matched on word tokens, so "Date Created
  // (UTC)" resolved where the script would have raised. Both date columns are
  // required, so an unresolved one stops the run instead of being worked around.
  var lowered = cleanedHeaders.map(function (h) { return h.toLowerCase(); });
  function findRequiredColumn(name) {
    var i = lowered.indexOf(clean(name).toLowerCase());
    if (i === -1) {
      throw new Error("Required column '" + name + "' was not found. Available columns: " + cleanedHeaders.join(', '));
    }
    return cleanedHeaders[i];
  }

  var stageCol = findRequiredColumn('Stage');
  var dateCreatedCol = findRequiredColumn('Date created');
  var dateClosedCol = findRequiredColumn('Date closed');

  rows.forEach(function (r) { r[stageCol] = clean(r[stageCol]); });

  var openRisk = rows.filter(function (r) { return OPEN_STAGES.hasOwnProperty(r[stageCol]); }).length;
  var closedRisk = rows.filter(function (r) { return r[stageCol] === CLOSED_STAGE && E.excelYear(r[dateClosedCol]) === YEAR; }).length;
  var totalRisk = openRisk + closedRisk;
  var riskCreated = rows.filter(function (r) { return E.excelYear(r[dateCreatedCol]) === YEAR; }).length;
  var targetValue = Math.round(totalRisk * TARGET_PERCENT);

  // get_period_label — the reporting period is the LAST month in the file that
  // carries real volume, and March maps to "Q1 '26". The clock is never
  // consulted, so the same workbook always produces the same label.
  //
  // Deliberately not max(): the latest single date let one row dated 2026-12-14
  // (a scheduled closure, a mistyped year) name a "Dec '26" period on a file
  // whose data is May. A month holding at least PERIOD_MIN_SHARE of the year's
  // dated rows is a reporting month; anything after it is noise, and it is
  // named in the run summary rather than dropped quietly. The Python does the
  // same — keep the two in step if either changes.
  var monthTally = [], skippedMonths = [], chosenMonth = null, periodFloor = 0;

  function getPeriodLabel() {
    var byMonth = {}, total = 0;
    [dateCreatedCol, dateClosedCol].forEach(function (col) {
      rows.forEach(function (r) {
        var d = E.excelDateInfo(r[col]);
        if (!d || d.year !== YEAR) return;
        byMonth[d.month] = (byMonth[d.month] || 0) + 1;
        total++;
      });
    });
    var months = Object.keys(byMonth).map(Number).sort(function (a, b) { return a - b; });
    if (!months.length) {
      throw new Error('No valid ' + YEAR + ' dates found in Date created/Date closed columns.');
    }
    periodFloor = Math.max(2, Math.ceil(total * PERIOD_MIN_SHARE));
    var solid = months.filter(function (m) { return byMonth[m] >= periodFloor; });
    var pool = solid.length ? solid : months;
    chosenMonth = pool[pool.length - 1];

    monthTally = months.map(function (m) { return MONTH_ABBR[m - 1] + ' x' + byMonth[m]; });
    skippedMonths = months.filter(function (m) { return m > chosenMonth; })
      .map(function (m) { return MONTH_ABBR[m - 1] + ' x' + byMonth[m]; });

    return chosenMonth === 3 ? "Q1 '26" : (MONTH_ABBR[chosenMonth - 1] + " '26");
  }

  var periodLabel = getPeriodLabel();

  var processedOn = (function () {
    var d = new Date();
    function pad(n) { return String(n).padStart(2, '0'); }
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  })();

  var metrics = {
    Month: periodLabel,
    'Open risk as on date': openRisk,
    'Closed Risk in 2026': closedRisk,
    'Total Risk': totalRisk,
    'Risk Created in 2026': riskCreated,
    'Target 80%': targetValue,
    'Closed %': totalRisk ? closedRisk / totalRisk : 0,
    // os.path.basename(input_file) — the uploaded file's own name, threaded
    // through by runReport; the placeholder only shows if it wasn't supplied.
    'Input File': (wb && wb.__fileName) || 'workbook.xlsx',
    'Processed On': processedOn,
  };

  var HIST_COLS = ['Month', 'Open risk as on date', 'Closed Risk in 2026', 'Total Risk', 'Risk Created in 2026', 'Target 80%', 'Closed %', 'Input File', 'Processed On'];

  // load_or_create_history (step 15): prefer the table carried over from the
  // previous run; otherwise create one, seeded when configured to be.
  var storageOk = !!(window.S4History && window.S4History.available());
  var stored = window.S4History ? window.S4History.load() : null;
  var history, historySource;
  if (stored) {
    // Normalise: a stored row must expose every history column, since older
    // payloads may predate a column being added.
    history = stored.map(function (r) {
      var row = {};
      HIST_COLS.forEach(function (c) { row[c] = r[c] === undefined ? '' : r[c]; });
      return row;
    });
    historySource = 'browser storage';
  } else if (SEED_HISTORY_WHEN_MISSING) {
    history = SEED_ROWS.map(function (r) {
      var row = Object.assign({}, r);
      row['Target 80%'] = Math.round(row['Total Risk'] * TARGET_PERCENT);
      row['Closed %'] = row['Closed Risk in 2026'] / row['Total Risk'];
      row['Input File'] = 'Seed from sample image';
      row['Processed On'] = processedOn;
      return row;
    });
    historySource = 'seed rows';
  } else {
    history = [];
    historySource = 'empty';
  }

  // update_history (step 16): append only if this month isn't already present.
  // Never overwrite — a re-run of the same month keeps the original figures.
  var alreadyPresent = history.some(function (r) { return String(r.Month) === String(metrics.Month); });
  var historyUpdated = !alreadyPresent;
  if (historyUpdated) history = history.concat([metrics]);

  // save_history (step 20): this is what makes the next run cumulative.
  var historySaved = window.S4History ? window.S4History.save(history) : false;

  // build_dashboard_frames
  function periodSortKey(label) {
    label = String(label);
    var q = label.match(/^Q([1-4])/);
    if (q) return +q[1] * 3; // a quarter sorts at its ending month: Q1=3, Q2=6, Q3=9, Q4=12
    var m = label.match(/^([A-Za-z]{3})/);
    var order = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
    return m ? (order[m[1]] || 99) : 99;
  }
  var sortedHistory = history.slice().sort(function (a, b) { return periodSortKey(a.Month) - periodSortKey(b.Month); });
  // This dashboard is "as on" the period THIS file describes, so a stored period
  // LATER than it cannot belong here. Left in, it takes one of the three chart
  // columns and pushes a real period off the left edge — which is how a May
  // report kept showing Apr/May/Dec long after the December row that caused it
  // was fixed: the bad label was already saved, and fixing the derivation does
  // not unsay it. The row stays in history; it is simply not part of this
  // report. Mirrored in diagram4/automation.py — change both together.
  var currentKey = periodSortKey(periodLabel);
  var laterPeriods = sortedHistory.filter(function (r) { return periodSortKey(r.Month) > currentKey; })
    .map(function (r) { return r.Month; });
  var eligibleHistory = sortedHistory.filter(function (r) { return periodSortKey(r.Month) <= currentKey; });

  // Quarter roll-up (reference workflow steps 17-20): history keeps MONTHLY
  // rows, but a completed quarter displays as ONE column. A quarter is
  // complete when its ENDING month's row exists (Q2 <- Jun, Q3 <- Sep,
  // Q4 <- Dec), and the quarter shows that ending month's numbers — a
  // snapshot, not a sum, because these are cumulative as-on-date values.
  // Months of an incomplete quarter stay individual, so the display runs
  // "Q1, Q2, Jul, Aug" until September collapses it to "Q1, Q2, Q3". Q1
  // needs no month handling: March is stored as "Q1 '26" at derivation
  // time (getPeriodLabel), as is the seed row.
  function rollupDisplay(rows) {
    var QUARTER_OF = { 4: 'Q2', 5: 'Q2', 6: 'Q2', 7: 'Q3', 8: 'Q3', 9: 'Q3', 10: 'Q4', 11: 'Q4', 12: 'Q4' };
    var QUARTER_END = { 6: 'Q2', 9: 'Q3', 12: 'Q4' };
    function isMonthLabel(l) { return /^[A-Za-z]{3}/.test(String(l)); }
    var endingRow = {};
    rows.forEach(function (r) {
      if (!isMonthLabel(r.Month)) return;
      var q = QUARTER_END[periodSortKey(r.Month)];
      if (q) endingRow[q] = r;
    });
    var out = [], emitted = {};
    rows.forEach(function (r) {
      var q = isMonthLabel(r.Month) ? QUARTER_OF[periodSortKey(r.Month)] : null;
      if (q && endingRow[q]) {
        if (!emitted[q]) { emitted[q] = 1; out.push({ label: q + " '26", row: endingRow[q] }); }
        return; // month absorbed into its completed quarter
      }
      out.push({ label: String(r.Month), row: r });
    });
    return out;
  }
  var displayHistory = rollupDisplay(eligibleHistory);
  var graphDisplay = displayHistory.slice(-3);
  var graphHistory = graphDisplay.map(function (d) { return d.row; });
  var periodLabels = graphDisplay.map(function (d) { return d.label; });
  var targetForTable = Math.round(totalRisk * TARGET_PERCENT);

  var wideColumns = [BASELINE_LABEL].concat(periodLabels, [TARGET_LABEL]);
  var rowOpen = [BASELINE_OPEN_RISK].concat(graphHistory.map(function (r) { return r['Open risk as on date']; }), ['']);
  var rowClosed = [BASELINE_CLOSED_RISK].concat(graphHistory.map(function (r) { return r['Closed Risk in 2026']; }), ['']);
  var rowTotal = [BASELINE_TOTAL_RISK].concat(graphHistory.map(function (r) { return r['Total Risk']; }), ['']);
  var rowBlank = wideColumns.map(function () { return ''; });
  var rowCreated = [0].concat(graphHistory.map(function (r) { return r['Risk Created in 2026']; }), ['']);
  var rowBaseline = [BASELINE_TOTAL_RISK].concat(periodLabels.map(function () { return BASELINE_TOTAL_RISK; }), ['']);

  var tableRows = [
    ['1. Open risk as on date'].concat(rowOpen),
    ['Closed Risk in 2026'].concat(rowClosed),
    ['Total Risk'].concat(rowTotal),
    [''].concat(rowBlank),
    ['Risk Created in 2026'].concat(rowCreated),
    [''].concat(rowBaseline),
  ];
  var tableHeaders = [''].concat(wideColumns);

  var calcLabels = [BASELINE_LABEL].concat(periodLabels, [TARGET_LABEL]);
  var calcValues = [BASELINE_TOTAL_RISK].concat(graphHistory.map(function (r) { return r['Open risk as on date']; }), [targetForTable]);
  var calcPercents = [''].concat(graphHistory.map(function (r) { return r['Closed %']; }), [TARGET_PERCENT]);

  // --- build the Dashboard grid (0-indexed rows/cols, matching xlsxwriter's ws.write(row,col,...)) ---
  function setCell0(grid, row, col, value) {
    while (grid.length <= row) grid.push([]);
    while (grid[row].length <= col) grid[row].push(null);
    grid[row][col] = value;
  }
  // Every cell below that holds a fraction meant to READ as a percentage.
  // The original xlsxwriter script wrote these with pct_fmt (num_format "0%");
  // the port dropped the format and left the bare float, which is why the
  // sheet showed 0.058908046 instead of 5.9%. Tracked as coordinates so the
  // Excel number format and the in-app preview stay in step.
  // Whole percent, as the original xlsxwriter script wrote it (pct_fmt =
  // num_format "0%"): Q1 6%, Apr 6%, May 12%.
  var PCT_FMT = '0%';
  var pctCells = [];
  function setPercent0(grid, row, col, value) {
    setCell0(grid, row, col, value === '' ? '' : value);
    if (value !== '') pctCells.push([row, col]);
  }

  var g = [];
  tableHeaders.forEach(function (h, c) { setCell0(g, 0, c, h); });
  tableRows.forEach(function (row, r) { row.forEach(function (v, c) { setCell0(g, r + 1, c, v); }); });

  var startRow = 10;
  setCell0(g, startRow, 0, 'Calculation used for chart');
  calcLabels.forEach(function (label, r) {
    var excelRow = startRow + 1 + r;
    setCell0(g, excelRow, 0, label);
    setCell0(g, excelRow, 1, calcValues[r]);
    setPercent0(g, excelRow, 2, calcPercents[r]);
  });
  // The script's wording, verbatim. The arithmetic behind it (which inputs made
  // this Target) is in the run summary instead, where it cannot alter the sheet.
  setCell0(g, startRow + 1 + calcLabels.length, 0, 'Note');
  setCell0(g, startRow + 1 + calcLabels.length, 1, 'Target = 80% of latest Total Risk');

  var chartStart = startRow + 10;
  setCell0(g, chartStart, 0, 'Chart Label');
  setCell0(g, chartStart, 1, 'Chart Value');
  setCell0(g, chartStart, 2, 'Percent Label');
  calcLabels.forEach(function (label, r) {
    var excelRow = chartStart + 1 + r;
    setCell0(g, excelRow, 0, label);
    setCell0(g, excelRow, 1, calcValues[r]);
    setPercent0(g, excelRow, 2, calcPercents[r]);
  });

  var historyHeaderRow = HIST_COLS;
  var historyDataRows = history.map(function (r) { return HIST_COLS.map(function (c) { return r[c]; }); });
  var historyGrid = [historyHeaderRow].concat(historyDataRows);
  // "Closed %" is the same fraction-as-a-percent column, one sheet over.
  var closedPctCol = HIST_COLS.indexOf('Closed %');
  var histPctCells = historyDataRows.map(function (_, i) { return [i + 1, closedPctCol]; });

  // Excel keeps the live number (still computable / chartable) and gets the
  // number format. The in-app preview renders grid values verbatim, so hand
  // it the already-rendered percentage instead of the raw fraction.
  //
  // Dashboard only. History's grid stays numeric because selectCharts('s4')
  // derives the preview chart from History.xlsx via chartableFromSheet, which
  // picks its series by "is this column numeric" — stringifying Closed % there
  // would silently drop a series from the chart.
  function renderPercents(grid, cells) {
    var out = grid.map(function (row) { return row.slice(); });
    cells.forEach(function (rc) {
      var v = out[rc[0]] && out[rc[0]][rc[1]];
      if (typeof v === 'number') out[rc[0]][rc[1]] = Math.round(v * 100) + '%';
    });
    return out;
  }
  var gPreview = renderPercents(g, pctCells);

  // Preview chart comes from the History table (matches today's app, which
  // derives it from History.xlsx, not from the Dashboard's own matrix — see
  // ReportEngine.selectCharts' 's4' branch). The embedded chart on the
  // Dashboard sheet reproduces the original xlsxwriter native column chart
  // instead (black chart/plot area, gray Baseline/Target end-bars, gold
  // bars in between, white centered value labels, no legend/gridlines).
  var barColors = calcLabels.map(function (_, i) {
    return (i === 0 || i === calcLabels.length - 1) ? '#BFBFBF' : '#FFC000';
  });
  var progressTraces = [
    { x: calcLabels, y: calcValues, type: 'bar', marker: { color: barColors },
      text: calcValues.map(String), textposition: 'inside', insidetextanchor: 'middle', textangle: 0,
      textfont: { color: '#ffffff', size: 12 } },
  ];
  var progressLayout = {
    paper_bgcolor: '#000000', plot_bgcolor: '#000000',
    title: { text: "<b>Cumulative Risk Treatment<br>Progress</b>", font: { color: '#ffffff', size: 18 } },
    xaxis: { tickfont: { color: '#ffffff' }, linecolor: '#ffffff', showline: true },
    yaxis: { visible: false },
    showlegend: false,
    margin: { t: 70, r: 20, b: 50, l: 20 },
    // Closed % above each period bar, 80% above Target (workflow step 24).
    // PNG only: the native Excel chart's single series can't carry a second
    // label set, so there the percents stay in the Percent Label column.
    annotations: calcLabels.map(function (label, i) {
      var p = calcPercents[i];
      if (p === '') return null;
      return { x: label, y: calcValues[i], text: Math.round(p * 100) + '%', showarrow: false,
        yanchor: 'bottom', yshift: 2, font: { color: '#ffffff', size: 12 } };
    }).filter(Boolean),
  };
  var progressChartPng = await B.renderStyledPng(progressTraces, progressLayout, 720, 430);

  var workbook = new ExcelJS.Workbook();
  var wsDash = workbook.addWorksheet(DASH_SHEET);
  g.forEach(function (r) { wsDash.addRow(r); });
  // grid is 0-indexed, ExcelJS is 1-indexed.
  pctCells.forEach(function (rc) { wsDash.getCell(rc[0] + 1, rc[1] + 1).numFmt = PCT_FMT; });

  // ── Cell formatting — the xlsxwriter format objects, one for one ──
  // header_fmt / left_fmt gold, peach_fmt on the calculation labels, thin
  // borders throughout, body numbers right-aligned. Number formats are left
  // alone: PCT_FMT above already covers every percent cell.
  var GOLD = 'FFFFC000', PEACH = 'FFF4B183';
  var THIN = { style: 'thin', color: { argb: 'FF000000' } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  var CENTER = { horizontal: 'center', vertical: 'middle' }, RIGHT = { horizontal: 'right' };
  function fmt(row, col, o) {
    var cell = wsDash.getCell(row, col);
    if (o.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.fill } };
    if (o.bold || o.size) cell.font = { bold: !!o.bold, size: o.size || 11 };
    if (o.border) cell.border = BORDER;
    if (o.align) cell.alignment = o.align;
  }

  // Matrix: gold header row, gold row labels down column A, bordered body —
  // applied to blank cells too, as the script does, so the grid stays unbroken.
  var nCols = tableHeaders.length;
  for (var hc = 1; hc <= nCols; hc++) fmt(1, hc, { fill: GOLD, bold: true, border: true, align: CENTER });
  tableRows.forEach(function (_, r) {
    var rowNum = r + 2;
    fmt(rowNum, 1, { fill: GOLD, bold: true, border: true });
    for (var c = 2; c <= nCols; c++) fmt(rowNum, c, { border: true, align: RIGHT });
  });

  // title_fmt, then the calculation block: peach labels, bordered values,
  // centred percents. The helper block below reuses the same percent styling.
  var titleRow = startRow + 1;
  fmt(titleRow, 1, { bold: true, size: 14 });
  calcLabels.forEach(function (_, r) {
    var rowNum = titleRow + 1 + r;
    fmt(rowNum, 1, { fill: PEACH, border: true, align: CENTER });
    fmt(rowNum, 2, { border: true });
    fmt(rowNum, 3, calcPercents[r] === '' ? { border: true } : { border: true, align: CENTER });
  });
  var noteRowNum = titleRow + 1 + calcLabels.length;
  fmt(noteRowNum, 1, { fill: GOLD, bold: true, border: true });
  fmt(noteRowNum, 2, { border: true });
  var chartStartRow = chartStart + 1;
  calcLabels.forEach(function (_, r) {
    if (calcPercents[r] !== '') fmt(chartStartRow + 1 + r, 3, { border: true, align: CENTER });
  });

  // set_column(0,0,28) / set_column(1,6,14). xlsxwriter stores a requested
  // width plus its own padding, so these are the values that land on disk —
  // a bare 28/14 here would render narrower than the script's output.
  wsDash.getColumn(1).width = 28.7109375;
  for (var wc = 2; wc <= 7; wc++) wsDash.getColumn(wc).width = 14.7109375;
  // Native, editable progress chart (data-linked to a hidden helper block)
  // replaces the baked PNG: single gold series with the first and last bars
  // recolored gray (Baseline/Target), white centered labels, no legend.
  var s4Placements = [];
  if (window.NativeChartInject && window.fflate && calcLabels.length) {
    // Reference the "Chart Label/Chart Value" block: header at Excel row
    // chartStart+1, data rows chartStart+2 .. The bar labels are editable in
    // column A of that block. First/last bars are recolored gray (per-point).
    var R = window.NativeChartInject.ref, s4First = chartStart + 2, s4Last = chartStart + 1 + calcLabels.length;
    var s4Points = calcLabels.map(function (_, i) {
      return (i === 0 || i === calcLabels.length - 1) ? { idx: i, color: 'BFBFBF' } : null;
    }).filter(Boolean);
    s4Placements.push({
      sheetName: DASH_SHEET, anchor: { fromCol: 6, fromRow: 1, toCol: 15, toRow: 20 }, // ~"G2"
      def: {
        grouping: 'clustered', legend: false, title: 'Cumulative Risk Treatment Progress',
        chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
        // set_y_axis({visible: False, major_gridlines: {visible: False}}) and
        // set_x_axis({label_position: "low"}) — the bars carry their own labels.
        hideValAx: true, catTickLblPos: 'low',
        dataLabels: { position: 'ctr', color: 'FFFFFF' },
        categories: { ref: R(DASH_SHEET, 1, s4First, s4Last), cache: calcLabels },
        series: [
          { name: { lit: 'Progress' },
            values: { ref: R(DASH_SHEET, 2, s4First, s4Last), cache: calcValues }, color: 'FFC000', points: s4Points },
        ],
      },
    });
  }

  var wsHistUsed = workbook.addWorksheet(HIST_USED_SHEET);
  historyGrid.forEach(function (r) { wsHistUsed.addRow(r); });
  wsHistUsed.columns.forEach(function (c) { c.width = 16; });
  histPctCells.forEach(function (rc) { wsHistUsed.getCell(rc[0] + 1, rc[1] + 1).numFmt = PCT_FMT; });

  // Display History (workflow step 23): the rolled-up period list the chart
  // and matrix actually used — quarters where complete, months elsewhere.
  var DISPLAY_SHEET = 'Display History';
  var displayGrid = [['Period', 'Open risk as on date', 'Closed Risk in 2026', 'Total Risk', 'Risk Created in 2026', 'Closed %']]
    .concat(displayHistory.map(function (d) {
      return [d.label, d.row['Open risk as on date'], d.row['Closed Risk in 2026'], d.row['Total Risk'], d.row['Risk Created in 2026'], d.row['Closed %']];
    }));
  var wsDisplay = workbook.addWorksheet(DISPLAY_SHEET);
  displayGrid.forEach(function (r) { wsDisplay.addRow(r); });
  wsDisplay.columns.forEach(function (c) { c.width = 16; });
  displayGrid.slice(1).forEach(function (_, i) { wsDisplay.getCell(i + 2, 6).numFmt = PCT_FMT; });
  // Preview copy renders Closed % as "14%"; chartableFromSheet then skips the
  // now-text column, so the preview chart plots the count series only.
  var displayGridPreview = renderPercents(displayGrid, displayGrid.slice(1).map(function (_, i) { return [i + 1, 5]; }));

  var riskOutputBuf = await workbook.xlsx.writeBuffer();
  if (s4Placements.length) {
    try { riskOutputBuf = window.NativeChartInject.inject(new Uint8Array(riskOutputBuf), s4Placements); }
    catch (e) { console.error('s4 native chart inject failed:', e); }
  }

  var historyWorkbook = new ExcelJS.Workbook();
  var wsHistOnly = historyWorkbook.addWorksheet(HIST_SHEET);
  historyGrid.forEach(function (r) { wsHistOnly.addRow(r); });
  wsHistOnly.columns.forEach(function (c) { c.width = 16; });
  histPctCells.forEach(function (rc) { wsHistOnly.getCell(rc[0] + 1, rc[1] + 1).numFmt = PCT_FMT; });
  var historyBuf = await historyWorkbook.xlsx.writeBuffer();

  // ── Execution summary (step 29) ──
  // The script's closing print() block (automation.py:353-362). There is no
  // console here, so it is returned as stdout: buildPayload passes that
  // straight through and the results view renders it in the "Script log"
  // panel, which until now always read "No console output."
  var storageNote = !storageOk
    ? 'unavailable (private mode or file:// restrictions) — this run will not be remembered'
    : historySaved
      ? 'saved to browser storage · ' + history.length + ' period' + (history.length === 1 ? '' : 's') + ' held'
      : 'could not be saved (storage full or blocked)';
  var lines = [
    'Done.',
    'Input file: ' + metrics['Input File'],
    'Month calculated: ' + metrics.Month,
    'Open risk as on date: ' + openRisk,
    'Closed Risk in ' + YEAR + ': ' + closedRisk,
    'Total Risk: ' + totalRisk,
    'Target 80%: ' + targetValue,
    'History updated with new month: ' + (historyUpdated ? 'Yes' : 'No - month already existed'),
    'Created/updated: ' + OUTPUT_FILE,
    'Created/updated: ' + HISTORY_FILE,
    '',
    // Everything below is browser-side context — it explains the label rather
    // than changing it. The tally matters most: it shows which month the period
    // came from and, when a later month was ignored, that the rows exist.
    'Duplicate columns: ' + (dupHeaders.length
      ? dupHeaders.join(', ') + '   (first non-blank value per row is used)' : 'none'),
    'Reporting period: ' + MONTH_ABBR[chosenMonth - 1] + ' — the last month in ' +
      dateCreatedCol + ' / ' + dateClosedCol + ' holding at least ' + periodFloor + ' rows',
    'Dates by month in file: ' + monthTally.join(', ') +
      (skippedMonths.length ? '   — ignored as noise: ' + skippedMonths.join(', ') : ''),
    'Target arithmetic: ' + Math.round(TARGET_PERCENT * 100) + '% of Total Risk (' + openRisk +
      ' open + ' + closedRisk + ' closed = ' + totalRisk + ') = ' + targetValue,
    'History loaded from: ' + historySource,
    'History store: ' + storageNote,
    'Periods held: ' + history.map(function (r) { return r.Month; }).join(', '),
    'Display periods (quarter roll-up): ' + displayHistory.map(function (d) { return d.label; }).join(', '),
    'Periods in chart: ' + periodLabels.join(', ') + '  (last 3 of ' + displayHistory.length + ' display periods)',
  ];
  if (laterPeriods.length) {
    // The "  ^ " prefix is what makes the shell open the Script log by itself —
    // a period silently dropped from the chart must not go unmentioned.
    lines.push('  ^ Held but not charted: ' + laterPeriods.join(', ') +
      '   (later than ' + periodLabel + ', this file\'s period — use Reset to clear them)');
  }
  var stdout = lines.join('\n');

  return {
    ok: true,
    stdout: stdout,
    files: [
      { name: OUTPUT_FILE, bytes: riskOutputBuf, sheets: [{ name: DASH_SHEET, grid: gPreview }, { name: DISPLAY_SHEET, grid: displayGridPreview }, { name: HIST_USED_SHEET, grid: historyGrid }] },
      { name: HISTORY_FILE, bytes: historyBuf, sheets: [{ name: HIST_SHEET, grid: historyGrid }] },
    ],
    chartImages: (function () { var m = {}; m[DASH_SHEET] = progressChartPng; return m; })(),
  };
};
