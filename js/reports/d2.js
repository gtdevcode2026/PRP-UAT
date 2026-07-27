// Port of "daigram 2 automation/automation.py" (id: d2).
// Reads "OneTrust Assessment", filters Tags=="cyber" (case-insensitive
// exact match) & Date created year==2026, maps Stage to Closed/Open, maps
// Organization to a display zone (falling back to the organization name
// itself when unmapped — unlike s2a/s3d's zone_map, "Europe" stays
// "Europe" here, not "EUR"), pivots Org Display x Final Stage (count, no
// margins), sorts to a fixed zone order, appends a Grand Total row, and
// computes a small KPI panel (3 static values + 1 computed Q2 rate).
// Writes "output file D2.xlsx" as a single hand-built "Dashboard" sheet
// (filter area, then the pivot table, then the KPI table below it) — the
// exact row offsets matter because the existing preview pipeline re-treats
// row 0 as headers and sparsity-trims the rest, same as s2a.
window.Reports.d2 = async function d2(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers.map(function (h) { return String(h).trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    sheet.headers.forEach(function (h, i) { out[rawHeaders[i]] = row[h]; });
    return out;
  });
  // Report every missing column at once — naming them one run at a time turns
  // a single malformed export into five round trips.
  var missing = ['ID', 'Organization', 'Stage', 'Date created', 'Tags'].filter(function (c) {
    return rawHeaders.indexOf(c) === -1;
  });
  if (missing.length) throw new Error('Missing required columns: ' + missing.join(', '));

  // Exports carry invisible characters that survive a plain trim: zero-width
  // spaces/joiners, BOM, and non-breaking spaces between words. They make a
  // cell that reads as "Cyber" fail an equality test against "cyber" for no
  // visible reason, so strip them before anything compares these values.
  function clean(v) {
    if (E.isBlank(v)) return '';
    return String(v)
      .replace(/[​-‍﻿]/g, '')
      .replace(/[   ]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }
  rows.forEach(function (r) {
    ['Organization', 'Stage', 'Tags'].forEach(function (c) { r[c] = clean(r[c]); });
  });

  // Step 7: Tags equals "Cyber" (case-insensitive) AND Date created in 2026.
  // Counted per-predicate as well as combined, because when the result is zero
  // the only useful question is WHICH half rejected everything.
  var tagOnly = 0, yearOnly = 0;
  var filtered = rows.filter(function (r) {
    var tagMatch = /^cyber$/i.test(r.Tags);
    var yearMatch = E.excelYear(r['Date created']) === 2026;
    if (tagMatch) tagOnly++;
    if (yearMatch) yearOnly++;
    return tagMatch && yearMatch;
  });

  var CLOSED_STAGES = { 'Completed': 1, 'Under review': 1 };
  filtered.forEach(function (r) {
    r['Final Stage'] = CLOSED_STAGES.hasOwnProperty(r.Stage) ? 'Closed' : 'Open';
  });

  var ORG_MAP = {
    'Africa': 'AFR', 'APAC': 'APAC', 'BEES': 'GRO', 'BEES | FINTECH': 'GRO',
    'Europe': 'Europe', 'GHQ': 'GHQ', 'South America Zone': 'SAZ',
    'North America Zone': 'NAZ', 'Middle America Zone': 'MAZ',
  };
  filtered.forEach(function (r) { r['Org Display'] = E.mapZone(r.Organization, ORG_MAP); });

  // values="ID", aggfunc="count" — pandas counts non-null IDs, NOT rows, so a
  // Cyber/2026 row with a blank ID is excluded from the pivot even though it
  // still counts toward the KPI totals below (which use len(filtered)). That
  // asymmetry is the script's, and the Grand Total row may legitimately come
  // out lower than the "(closed/record)" figure in the chart title.
  var pivotRaw = E.pivotCount(filtered, function (r) { return r['Org Display']; }, function (r) { return r['Final Stage']; }, { margins: false, valueFn: function (r) { return r.ID; } });
  var headers = pivotRaw.headers.slice();
  var dataRows = pivotRaw.rows.map(function (r) { return r.slice(); });
  ['Open', 'Closed'].forEach(function (col) {
    if (headers.indexOf(col) === -1) { headers.push(col); dataRows.forEach(function (r) { r.push(0); }); }
  });
  var openIdx = headers.indexOf('Open'), closedIdx = headers.indexOf('Closed');
  var pivotRows = pivotRaw.indexVals.map(function (iv, i) {
    var open = dataRows[i][openIdx] || 0, closed = dataRows[i][closedIdx] || 0;
    return { label: iv, open: open, closed: closed, grandTotal: open + closed };
  });

  var ORDER = ['AFR', 'APAC', 'GRO', 'Europe', 'GHQ', 'SAZ', 'MAZ', 'NAZ'];
  pivotRows.forEach(function (r) { r.sortOrder = ORDER.indexOf(r.label); if (r.sortOrder === -1) r.sortOrder = 999; });
  pivotRows.sort(function (a, b) {
    if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
    return a.label < b.label ? -1 : (a.label > b.label ? 1 : 0);
  });

  var totalOpen = pivotRows.reduce(function (a, r) { return a + r.open; }, 0);
  var totalClosed = pivotRows.reduce(function (a, r) { return a + r.closed; }, 0);
  var totalGrand = pivotRows.reduce(function (a, r) { return a + r.grandTotal; }, 0);
  var pivotDisplay = pivotRows.concat([{ label: 'Grand Total', open: totalOpen, closed: totalClosed, grandTotal: totalGrand }]);

  // Python's round() is half-to-EVEN, so round(0.125, 2) is 0.12 where JS's
  // Math.round would give 0.13. Small filtered sets land on those exact halves
  // (1/8, 3/8, ...), and the KPI is published, so match the script's rounding.
  function pyRound2(x) {
    var scaled = x * 100, floor = Math.floor(scaled), diff = scaled - floor;
    var n = diff > 0.5 ? floor + 1 : (diff < 0.5 ? floor : (floor % 2 === 0 ? floor : floor + 1));
    return n / 100;
  }

  var closedTotal = filtered.filter(function (r) { return r['Final Stage'] === 'Closed'; }).length;
  var recordTotal = filtered.length;
  var q2_26 = recordTotal ? pyRound2(closedTotal / recordTotal) : 0;
  var KPI = [
    ["Baseline '25", 0.60, 'static'],
    ["Q1 '26", 0.32, 'static'],
    ["Q2 '26", q2_26, ''],
    ["Target '26", 0.65, 'static'],
  ];

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
  var pivotHeaderRow = ['Row Labels', 'Open', 'Closed', 'Grand Total'];
  pivotHeaderRow.forEach(function (h, i) { setCell(grid, startRow, i + 1, h); });
  pivotDisplay.forEach(function (r, i) {
    var rowNum = startRow + 1 + i;
    setCell(grid, rowNum, 1, r.label);
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
    setCell(grid, rowNum, 3, k[2]);
  });
  var noteRow = kpiStartRow + KPI.length + 2;
  setCell(grid, noteRow, 1, 'Q2 formula');
  setCell(grid, noteRow, 2, 'Closed / Grand Total = ' + closedTotal + ' / ' + recordTotal);

  // ── Diagnostics ──
  // Q2 is a ratio over the filtered set, so when the filter matches nothing it
  // reports a confident "0%" that is indistinguishable from a real measurement.
  // These four rows make the difference visible: they say how many rows each
  // half of the Step 7 filter accepted, and — decisively — what the rejected
  // values actually look like, so one client run identifies the cause instead
  // of requiring another round trip.
  //
  // Written across all four columns on purpose. ReportEngine.trimSparseCols
  // drops any column filled in fewer than half the surviving rows, and column D
  // ("Grand Total") sits close to that line; a block of 2-cell rows would push
  // it under and silently delete a column from the in-app preview table.
  function distinct(values, limit) {
    var seen = [], counts = {};
    values.forEach(function (v) {
      var k = v === '' ? '(blank)' : String(v);
      if (!counts[k]) { counts[k] = 0; seen.push(k); }
      counts[k]++;
    });
    seen.sort(function (a, b) { return counts[b] - counts[a]; });
    var shown = seen.slice(0, limit).map(function (k) { return k + ' ×' + counts[k]; });
    if (seen.length > limit) shown.push('+' + (seen.length - limit) + ' more');
    return shown.join('  ·  ') || '(none)';
  }

  var unparsedDates = 0;
  var yearsSeen = rows.map(function (r) {
    var y = E.excelYear(r['Date created']);
    if (y === null) { unparsedDates++; return 'unreadable'; }
    return String(y);
  });

  var diagRow = noteRow + 2;
  ['Diagnostics', 'Rows read', 'Tags = Cyber', 'Created in 2026'].forEach(function (h, i) {
    setCell(grid, diagRow, i + 1, h);
  });
  setCell(grid, diagRow + 1, 1, 'Rows matching');
  setCell(grid, diagRow + 1, 2, rows.length);
  setCell(grid, diagRow + 1, 3, tagOnly);
  setCell(grid, diagRow + 1, 4, yearOnly);
  setCell(grid, diagRow + 2, 1, 'Tags seen');
  setCell(grid, diagRow + 2, 2, distinct(rows.map(function (r) { return r.Tags; }), 6));
  setCell(grid, diagRow + 3, 1, 'Date created seen');
  setCell(grid, diagRow + 3, 2, distinct(yearsSeen, 6) +
    (unparsedDates ? '  — ' + unparsedDates + ' cell(s) could not be read as a date' : ''));

  var files = [{ name: 'output file D2.xlsx', sheets: [{ name: 'Dashboard', grid: grid }] }];

  // Both embedded charts reproduce the original matplotlib styles instead
  // of the light in-page preview theme. Both are embedded into the file
  // (matching today's 2-chart output), even though only the KPI chart ever
  // surfaces in the preview — confirmed with the user.
  var images = {};

  // Chart 1: org-level stacked bar — black bg, Open (cyan) bottom + Closed
  // (gold) top, white in-bar labels, rotated white x labels, hidden y axis,
  // top-center legend (Closed shown before Open, matching legend_order).
  var orgLabels = pivotRows.map(function (r) { return r.label; });
  var openVals = pivotRows.map(function (r) { return r.open; });
  var closedVals = pivotRows.map(function (r) { return r.closed; });
  var orgTraces = [
    { x: orgLabels, y: openVals, type: 'bar', name: 'Open', marker: { color: '#00AEEF' }, width: 0.48,
      text: openVals.map(function (v) { return v > 0 ? String(v) : ''; }),
      textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff', size: 10 } },
    { x: orgLabels, y: closedVals, type: 'bar', name: 'Closed', marker: { color: '#D4AF37' }, width: 0.48,
      text: closedVals.map(function (v) { return v > 0 ? String(v) : ''; }),
      textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff', size: 10 } },
  ];
  var orgLayout = {
    barmode: 'stack',
    paper_bgcolor: '#000000', plot_bgcolor: '#000000',
    title: { text: '<b>2026 Assessment<br>(' + closedTotal + '/' + recordTotal + ')</b>', font: { color: '#ffffff', size: 16 } },
    xaxis: { tickangle: -45, tickfont: { color: '#ffffff' }, showline: true, linecolor: '#ffffff' },
    yaxis: { visible: false },
    legend: { orientation: 'h', x: 0.5, xanchor: 'center', y: 0.98, yanchor: 'top', traceorder: 'reversed',
      font: { color: '#ffffff' }, bgcolor: 'rgba(0,0,0,0)' },
    margin: { t: 90, r: 20, b: 90, l: 20 },
  };
  images.org = await B.renderStyledPng(orgTraces, orgLayout, 560, 400);

  // Chart 2: KPI horizontal bar — white bg, single blue series, 0-100% axis,
  // percent value labels to the right of each bar, dashed x gridlines.
  var kpiMetrics = KPI.map(function (k) { return k[0]; });
  var kpiValues = KPI.map(function (k) { return k[1]; });
  var kpiTraces = [
    { x: kpiValues, y: kpiMetrics, type: 'bar', orientation: 'h', marker: { color: '#4472C4' },
      text: kpiValues.map(function (v) { return Math.round(v * 100) + '%'; }),
      textposition: 'outside', textfont: { color: '#000000', size: 11 } },
  ];
  var kpiLayout = {
    paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
    title: { text: '<b>Improve in Supplier Response Time</b>', font: { color: '#000000', size: 15 } },
    xaxis: { range: [0, 1], tickvals: [0, 0.25, 0.5, 0.75, 1], ticktext: ['0%', '25%', '50%', '75%', '100%'],
      gridcolor: 'rgba(0,0,0,0.35)', griddash: 'dash' },
    yaxis: { autorange: 'reversed' },
    showlegend: false,
    margin: { t: 50, r: 40, b: 40, l: 90 },
  };
  images.kpi = await B.renderStyledPng(kpiTraces, kpiLayout, 560, 300);

  var workbook = new ExcelJS.Workbook();
  var ws = workbook.addWorksheet('Dashboard');
  grid.forEach(function (r) { ws.addRow(r); });

  // ── Cell styling — the openpyxl block, one for one ──
  // header_fill/subheader_fill/grand_fill + bold + thin border + centering.
  // ExcelJS wants ARGB, openpyxl takes RGB, hence the FF alpha prefix.
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

  // Filter area: labels in A, values in B.
  [1, 2].forEach(function (r) {
    style(r, 1, { fill: HEADER_FILL, bold: true, border: true });
    style(r, 2, { fill: SUBHEADER_FILL, border: true });
  });

  // Pivot table: header row, then bordered body with everything but the row
  // label centred, then the Grand Total row re-filled and bolded on top.
  var PIVOT_COLS = 4;
  for (var pc = 1; pc <= PIVOT_COLS; pc++) {
    style(startRow, pc, { fill: HEADER_FILL, bold: true, border: true, center: true });
  }
  pivotDisplay.forEach(function (r, i) {
    var rowNum = startRow + 1 + i;
    for (var c = 1; c <= PIVOT_COLS; c++) {
      style(rowNum, c, { border: true, center: c > 1 });
      if (r.label === 'Grand Total') style(rowNum, c, { fill: HEADER_FILL, bold: true });
    }
  });

  // KPI table: header row, bordered rows, values shown as whole percents.
  for (var kc = 1; kc <= 3; kc++) {
    style(kpiStartRow, kc, { fill: HEADER_FILL, bold: true, border: true, center: true });
  }
  KPI.forEach(function (_, i) {
    var rowNum = kpiStartRow + 1 + i;
    for (var c = 1; c <= 3; c++) style(rowNum, c, { border: true });
    style(rowNum, 2, { numFmt: '0%' });
  });
  style(noteRow, 1, { bold: true });

  // Diagnostics block: same header treatment as the tables above it.
  for (var dc = 1; dc <= 4; dc++) {
    style(diagRow, dc, { fill: HEADER_FILL, bold: true, border: true, center: dc > 1 });
    style(diagRow + 1, dc, { border: true, center: dc > 1 });
  }
  style(diagRow + 2, 1, { bold: true });
  style(diagRow + 3, 1, { bold: true });

  // Column widths (A-M), uniform row heights, and freeze_panes = "A5" — which
  // is a 4-row vertical split, so the pivot header stays put while scrolling.
  var WIDTHS = [18, 12, 12, 14, 4, 14, 14, 14, 14, 14, 14, 14, 14];
  WIDTHS.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
  // Cover the whole sheet plus the chart anchors below it. Derived rather than
  // hardcoded: the grid grows with the number of zones in the client's data.
  var lastStyledRow = Math.max(49, grid.length);
  for (var rh = 1; rh <= lastStyledRow; rh++) ws.getRow(rh).height = 22;
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: startRow }];

  // Native, editable charts (data-linked to hidden helper blocks) replace the
  // two baked PNGs: org = stacked column, kpi = horizontal percent bar.
  var placements = [];
  if (window.NativeChartInject && window.fflate) {
    var R = window.NativeChartInject.ref;
    if (orgLabels.length) {
      // Org pivot: header row 4 (B=Open, C=Closed), data rows 5..4+n. The chart
      // excludes the appended Grand Total row (pivotRows only). Labels editable
      // in column A; legend names editable via B4/C4.
      var orgLast = 4 + pivotRows.length;
      placements.push({
        sheetName: 'Dashboard', anchor: { fromCol: 5, fromRow: 1, toCol: 14, toRow: 20 }, // ~"F2"
        def: {
          grouping: 'stacked', legend: true, title: '2026 Assessment (' + closedTotal + '/' + recordTotal + ')',
          chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
          dataLabels: { position: 'ctr', color: 'FFFFFF' },
          categories: { ref: R('Dashboard', 1, 5, orgLast), cache: orgLabels },
          series: [
            { name: { ref: R('Dashboard', 2, 4, 4), lit: 'Open' },
              values: { ref: R('Dashboard', 2, 5, orgLast), cache: openVals }, color: '00AEEF' },
            { name: { ref: R('Dashboard', 3, 4, 4), lit: 'Closed' },
              values: { ref: R('Dashboard', 3, 5, orgLast), cache: closedVals }, color: 'D4AF37' },
          ],
        },
      });
    }
    if (kpiMetrics.length) {
      // KPI block: header at kpiStartRow (A=Metric, B=Value), data rows below.
      // The Baseline/Q1/Q2/Target labels sit in column A and are editable there.
      var kpiFirst = kpiStartRow + 1, kpiLast = kpiStartRow + kpiMetrics.length;
      placements.push({
        sheetName: 'Dashboard', anchor: { fromCol: 5, fromRow: 22, toCol: 14, toRow: 40 }, // ~"F23"
        def: {
          grouping: 'clustered', barDir: 'bar', catReversed: true, legend: false,
          title: 'Improve in Supplier Response Time', valNumFmt: '0%',
          dataLabels: { position: 'outEnd', numFmt: '0%', color: '000000' },
          categories: { ref: R('Dashboard', 1, kpiFirst, kpiLast), cache: kpiMetrics },
          series: [
            { name: { ref: R('Dashboard', 2, kpiStartRow, kpiStartRow), lit: 'Improvement' },
              values: { ref: R('Dashboard', 2, kpiFirst, kpiLast), cache: kpiValues }, color: '4472C4' },
          ],
        },
      });
    }
  }

  var buf = await workbook.xlsx.writeBuffer();
  if (placements.length) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('d2 native chart inject failed:', e); }
  }

  return {
    ok: true,
    files: [{ name: 'output file D2.xlsx', bytes: buf, sheets: [{ name: 'Dashboard', grid: grid }] }],
    chartImages: { 'Dashboard': images.org },
  };
};
