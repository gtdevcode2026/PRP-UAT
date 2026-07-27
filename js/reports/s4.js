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
  // Versioned: bump when a change makes rows written by the previous build
  // untrustworthy. v1 rows were stamped by the clock-derived period code, which
  // could invent a month no sheet contained (a "Jul '26" alongside May data)
  // and then carry it forward into every later run. Nothing distinguishes a
  // poisoned row from a good one after the fact, so the whole generation is
  // retired rather than asking every user to find the Reset button.
  var KEY = 'prp.s4.history.v2';
  var STALE_KEYS = ['prp.s4.history.v1'];

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
  var SEED_HISTORY_WHEN_MISSING = true;
  // A month must hold at least this share of the year's dated rows to be
  // treated as the reporting period. A real month carries roughly 1/N of the
  // file; a handful of future-dated or mistyped rows must not name the report.
  // Raise it if stray months still slip through, lower it if a genuine (very
  // partial) month is being skipped — the summary always says what it ignored.
  var PERIOD_MIN_SHARE = 0.05;
  var SEED_ROWS = [
    { Month: "Q1 '26", 'Open risk as on date': 655, 'Closed Risk in 2026': 41, 'Total Risk': 696, 'Risk Created in 2026': 118 },
    { Month: "Apr '26", 'Open risk as on date': 694, 'Closed Risk in 2026': 42, 'Total Risk': 736, 'Risk Created in 2026': 158 },
  ];
  var MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // Strips what a plain trim leaves behind. Exports carry zero-width spaces,
  // BOM and non-breaking spaces; Stage is matched by exact lookup against
  // OPEN_STAGES, so one invisible character silently drops a risk from the
  // count with nothing on screen to explain the discrepancy.
  function clean(v) {
    if (E.isBlank(v)) return '';
    return String(v)
      .replace(/[​-‍﻿]/g, '')
      .replace(/[   ]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Name the sheet AND what the workbook actually contains — "not found" on its
  // own gives the user nothing to act on (automation.py:100 lists them too).
  if (wb && Array.isArray(wb.SheetNames) && wb.SheetNames.indexOf(SHEET_NAME) === -1) {
    throw new Error("Sheet '" + SHEET_NAME + "' not found. Available sheets: " + wb.SheetNames.join(', '));
  }
  var sheet = E.readSheet(wb, SHEET_NAME);
  var cleanedHeaders = sheet.headers.map(clean);
  var rows = sheet.rows.map(function (row) {
    var out = {};
    sheet.headers.forEach(function (h, i) { out[cleanedHeaders[i]] = row[h]; });
    return out;
  });
  // Column resolution (spec step 7: "do not assume exact formatting").
  //
  // An exact name match is far too brittle for a real export: "Date created"
  // arrives as "Created Date", "Date Created (UTC)", "Date created on"... When
  // the date columns went unresolved the report fell back to the system clock
  // and labelled itself with the month it was RUN in, so a May workbook
  // reported "Jul '26" with nothing on screen to say the label was invented.
  // Match on the words instead: every token of the wanted name must appear in
  // the header, preferring an exact hit and then the shortest candidate.
  var norm = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
  var normalized = cleanedHeaders.map(norm);
  function findColumn(name) {
    var want = norm(name);
    var exact = normalized.indexOf(want);
    if (exact !== -1) return cleanedHeaders[exact];
    var tokens = want.split(' ');
    var hits = [];
    normalized.forEach(function (h, i) {
      var words = h.split(' ');
      if (tokens.every(function (t) { return words.indexOf(t) !== -1; })) hits.push(i);
    });
    if (!hits.length) return null;
    hits.sort(function (a, b) { return normalized[a].length - normalized[b].length; });
    return cleanedHeaders[hits[0]];
  }
  function findRequiredColumn(name) {
    var key = findColumn(name);
    if (!key) {
      throw new Error("Required column '" + name + "' was not found. Columns present: " + cleanedHeaders.join(', '));
    }
    return key;
  }
  function findOptionalColumn(name) { return findColumn(name); }

  var stageCol = findRequiredColumn('Stage');
  var dateCreatedCol = findOptionalColumn('Date created');
  var dateClosedCol = findOptionalColumn('Date closed');

  rows.forEach(function (r) { r[stageCol] = clean(r[stageCol]); });

  var openRisk = rows.filter(function (r) { return OPEN_STAGES.hasOwnProperty(r[stageCol]); }).length;
  var closedRisk;
  if (dateClosedCol) {
    closedRisk = rows.filter(function (r) { return r[stageCol] === CLOSED_STAGE && E.excelYear(r[dateClosedCol]) === YEAR; }).length;
  } else {
    closedRisk = rows.filter(function (r) { return r[stageCol] === CLOSED_STAGE; }).length;
  }
  var totalRisk = openRisk + closedRisk;
  var riskCreated = dateCreatedCol ? rows.filter(function (r) { return E.excelYear(r[dateCreatedCol]) === YEAR; }).length : 0;
  var targetValue = Math.round(totalRisk * TARGET_PERCENT);

  // ── Reporting period — derived from the DATA, never from the clock ──
  //
  // The system clock is not consulted anywhere in this derivation. It used to
  // be, twice, and both were wrong: unresolved date columns fell back to "the
  // month the report was run in" (a May workbook reported Jul '26), and the
  // future-date guard compared against today, so the same workbook could label
  // itself differently depending on when it was opened. A dashboard built from
  // a fixed export must produce the same period every time it is run.
  //
  // Months are tallied from the file, and the period is the LAST month that
  // carries real volume. A reporting month has substance; a lone row dated
  // 2026-12-14 — a scheduled closure date, a mistyped year — is noise, and
  // letting it name the period is what produced a phantom "Dec '26" column.
  // Anything skipped is named in the execution summary, never dropped quietly.
  var periodSource, monthTally = [], skippedMonths = [];

  // Month -> reporting label (spec step 4). Jan, Feb and Mar all fall inside
  // Q1, which is why step 17's display order reads "Q1, Apr, May ... Dec" with
  // no Jan or Feb slot — only April onward are reported as individual months.
  // The year suffix is derived from YEAR rather than written out, so changing
  // the reporting year in the config block carries through here.
  function periodLabelForMonth(m) {
    var yy = " '" + String(YEAR).slice(2);
    return m <= 3 ? ('Q1' + yy) : (MONTH_ABBR[m - 1] + yy);
  }

  function getPeriodLabel() {
    var byMonth = {}, total = 0;
    [dateCreatedCol, dateClosedCol].forEach(function (col) {
      if (!col) return;
      rows.forEach(function (r) {
        var d = E.excelDateInfo(r[col]);
        if (!d || d.year !== YEAR) return;
        total++;
        byMonth[d.month] = (byMonth[d.month] || 0) + 1;
      });
    });
    var months = Object.keys(byMonth).map(Number).sort(function (a, b) { return a - b; });
    if (!months.length) {
      throw new Error('No valid ' + YEAR + ' dates found in ' +
        [dateCreatedCol, dateClosedCol].filter(Boolean).join(' / ') + '.');
    }
    monthTally = months.map(function (m) { return MONTH_ABBR[m - 1] + ' ×' + byMonth[m]; });

    var floor = Math.max(2, Math.ceil(total * PERIOD_MIN_SHARE));
    var solid = months.filter(function (m) { return byMonth[m] >= floor; });
    var pool = solid.length ? solid : months;
    var chosen = pool[pool.length - 1];
    skippedMonths = months.filter(function (m) { return m > chosen; })
      .map(function (m) { return MONTH_ABBR[m - 1] + ' ×' + byMonth[m]; });

    periodSource = 'the file — ' + byMonth[chosen] + ' of ' + total + ' dated rows fall in ' + MONTH_ABBR[chosen - 1] +
      (skippedMonths.length ? '; ignored as noise: ' + skippedMonths.join(', ') + ' (under ' + floor + ' rows)' : '');
    return periodLabelForMonth(chosen);
  }

  if (!dateCreatedCol && !dateClosedCol) {
    // Guessing the period from the clock is how a May workbook came to report
    // July. With nothing in the data to read, refuse rather than invent.
    throw new Error('Cannot determine the reporting period: no "Date created" or "Date closed" column was found. ' +
      'Columns present: ' + cleanedHeaders.join(', '));
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
    if (label.indexOf('Q1') === 0) return 3;
    var m = label.match(/^([A-Za-z]{3})/);
    var order = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
    return m ? (order[m[1]] || 99) : 99;
  }
  var sortedHistory = history.slice().sort(function (a, b) { return periodSortKey(a.Month) - periodSortKey(b.Month); });
  // The dashboard is "as on" the period THIS file describes. A stored period
  // that falls after it came from some other export and cannot belong here —
  // left in, it takes one of the three chart columns and pushes a real period
  // (Q1) off the left edge, which is how a May report came to show Apr/May/Jul.
  // The row stays in storage; it is not wrong, just not this report. This makes
  // the chart a function of the file alone, whatever else storage happens to
  // hold, so the same workbook always renders the same three periods.
  var currentKey = periodSortKey(periodLabel);
  var laterPeriods = sortedHistory.filter(function (r) { return periodSortKey(r.Month) > currentKey; })
    .map(function (r) { return r.Month; });
  var graphHistory = sortedHistory.filter(function (r) { return periodSortKey(r.Month) <= currentKey; }).slice(-3);
  var periodLabels = graphHistory.map(function (r) { return r.Month; });
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
  // Show the arithmetic, not just the rule. "Target = 80% of latest Total Risk"
  // is unfalsifiable on its own; spelling out the inputs makes a wrong Target
  // immediately traceable to the number that produced it.
  setCell0(g, startRow + 1 + calcLabels.length, 0, 'Note');
  setCell0(g, startRow + 1 + calcLabels.length, 1,
    'Target = ' + Math.round(TARGET_PERCENT * 100) + '% of Total Risk (' + openRisk + ' open + ' +
    closedRisk + ' closed = ' + totalRisk + ') = ' + targetForTable);

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
  var stdout = [
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
    'Reporting period from: ' + periodSource,
    'Date columns used: ' + [dateCreatedCol, dateClosedCol].filter(Boolean).join(' / '),
    'Dates by month in file: ' + monthTally.join(', '),
    'History loaded from: ' + historySource,
    'History store: ' + storageNote,
    'Periods in chart: ' + periodLabels.join(', ') + '  (latest 3 of ' + history.length + ')',
  ].concat(laterPeriods.length
    ? ['Held but not charted: ' + laterPeriods.join(', ') + '  (later than ' + periodLabel + ', this file\'s period)']
    : []).join('\n');

  return {
    ok: true,
    stdout: stdout,
    files: [
      { name: OUTPUT_FILE, bytes: riskOutputBuf, sheets: [{ name: DASH_SHEET, grid: gPreview }, { name: HIST_USED_SHEET, grid: historyGrid }] },
      { name: HISTORY_FILE, bytes: historyBuf, sheets: [{ name: HIST_SHEET, grid: historyGrid }] },
    ],
    chartImages: (function () { var m = {}; m[DASH_SHEET] = progressChartPng; return m; })(),
  };
};
