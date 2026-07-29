// Port of "Slide 12 2nd daigram Automation/automation.py" (id: s2a).
// Reads "OneTrust Assessment", filters to fixed Stage/Organization allow
// lists, pivots Organization x Stage (count, reindexed to fixed orders),
// aggregates to a Zone-level Closed/Open summary (merging duplicate mapped
// zones), and builds a separate Ageing>=90 "overdue" version of the same
// zone summary. Writes "OneTrust_Report.xlsx" with 3 sheets, each a
// hand-built cell grid (title/rule rows above the real table) — mirrored
// exactly here, including the blank/title rows, since the existing preview
// pipeline's "row 0 = header" + sparsity-trim transform depends on that
// same messy raw structure for byte-for-byte parity with today's app.
window.Reports.s2a = async function s2a(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var STAGE_ORDER = ['Completed', 'In progress', 'Not started', 'Under review'];
  // BEES / BEES | FINTECH / Growth variants are sanitized to 'GRO' before
  // the allow-list filter runs (E.sanitizeZone below), so they enter as the
  // single 'GRO' org — one row in the pivot instead of two BEES rows.
  var ORG_ORDER = [
    'Africa', 'APAC', 'GRO', 'Europe', 'GHQ',
    'Middle America Zone', 'North America Zone', 'South America Zone',
  ];
  var ZONE_MAP = {
    'Africa': 'AFR', 'APAC': 'APAC', 'GRO': 'GRO',
    'Europe': 'EUR', 'GHQ': 'GHQ', 'Middle America Zone': 'MAZ',
    'North America Zone': 'NAZ', 'South America Zone': 'SAZ',
  };
  var ZONE_ORDER = ['AFR', 'APAC', 'GRO', 'EUR', 'GHQ', 'MAZ', 'NAZ', 'SAZ'];
  var OVERDUE_THRESHOLD = 90;

  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers.map(function (h) { return String(h).trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    sheet.headers.forEach(function (h, i) { out[rawHeaders[i]] = row[h]; });
    return out;
  });
  ['ID', 'Stage', 'Organization', 'Ageing'].forEach(function (c) {
    if (rawHeaders.indexOf(c) === -1) throw new Error("Missing required columns in 'OneTrust Assessment': " + c);
  });

  rows.forEach(function (r) {
    r.Stage = E.isBlank(r.Stage) ? '' : String(r.Stage).trim();
    r.Organization = E.isBlank(r.Organization) ? '' : E.sanitizeZone(String(r.Organization).trim());
    var n = parseFloat(r.Ageing);
    r.Ageing = isNaN(n) ? null : n;
  });

  var data = rows
    .filter(function (r) { return STAGE_ORDER.indexOf(r.Stage) !== -1; })
    .filter(function (r) { return ORG_ORDER.indexOf(r.Organization) !== -1; });
  data.forEach(function (r) {
    r.Zone = ZONE_MAP[r.Organization];
    r['Is Overdue'] = r.Ageing !== null && r.Ageing >= OVERDUE_THRESHOLD;
  });

  function reindexPivot(pivot, targetIndex, targetCols) {
    var lookup = {};
    pivot.indexVals.forEach(function (iv, i) {
      lookup[iv] = {};
      pivot.headers.forEach(function (h, j) { lookup[iv][h] = pivot.rows[i][j]; });
    });
    return {
      indexVals: targetIndex.slice(),
      headers: targetCols.slice(),
      rows: targetIndex.map(function (iv) {
        return targetCols.map(function (col) {
          return (lookup[iv] && lookup[iv][col] !== undefined) ? lookup[iv][col] : 0;
        });
      }),
    };
  }

  // --- pivot_df: Organization x Stage, reindexed, + Grand Total col & row ---
  var orgPivotRaw = E.pivotCount(data, function (r) { return r.Organization; }, function (r) { return r.Stage; }, { margins: false });
  var orgPivot = reindexPivot(orgPivotRaw, ORG_ORDER, STAGE_ORDER);
  var orgWithGT = orgPivot.rows.map(function (r) { return r.concat([r.reduce(function (a, b) { return a + b; }, 0)]); });
  var pivotHeaders = STAGE_ORDER.concat(['Grand Total']);
  var pivotTotalsRow = pivotHeaders.map(function (_, c) { return orgWithGT.reduce(function (acc, r) { return acc + r[c]; }, 0); });

  // --- zone_stage -> zone_summary (Closed/Open/Grand Total by zone) ---
  var zoneAgg = {};
  ORG_ORDER.forEach(function (org, i) {
    var zone = ZONE_MAP[org];
    if (!zoneAgg[zone]) zoneAgg[zone] = pivotHeaders.map(function () { return 0; });
    zoneAgg[zone] = zoneAgg[zone].map(function (v, c) { return v + orgWithGT[i][c]; });
  });
  var completedIdx = pivotHeaders.indexOf('Completed'), underReviewIdx = pivotHeaders.indexOf('Under review');
  var inProgressIdx = pivotHeaders.indexOf('In progress'), notStartedIdx = pivotHeaders.indexOf('Not started');
  var gtIdx = pivotHeaders.indexOf('Grand Total');
  var zoneSummary = ZONE_ORDER.map(function (z) {
    var v = zoneAgg[z] || pivotHeaders.map(function () { return 0; });
    var closed = v[completedIdx] + v[underReviewIdx];
    var open = v[inProgressIdx] + v[notStartedIdx];
    return { zone: z, closed: closed, open: open, grandTotal: v[gtIdx] };
  });
  var totalClosed = zoneSummary.reduce(function (a, z) { return a + z.closed; }, 0);
  var totalOpen = zoneSummary.reduce(function (a, z) { return a + z.open; }, 0);
  var totalAll = zoneSummary.reduce(function (a, z) { return a + z.grandTotal; }, 0);

  // --- overdue_zone: same shape, computed only from Is-Overdue rows ---
  var overdueRows = data.filter(function (r) { return r['Is Overdue']; });
  var overduePivotRaw = E.pivotCount(overdueRows, function (r) { return r.Zone; }, function (r) { return r.Stage; }, { margins: false });
  var overduePivot = reindexPivot(overduePivotRaw, ZONE_ORDER, STAGE_ORDER);
  var overdueZone = ZONE_ORDER.map(function (z, i) {
    var v = overduePivot.rows[i];
    var closed = v[completedIdx] + v[underReviewIdx];
    var open = v[inProgressIdx] + v[notStartedIdx];
    return { zone: z, closed: closed, open: open, grandTotal: closed + open };
  });
  var totalOverdueClosed = overdueZone.reduce(function (a, z) { return a + z.closed; }, 0);
  var totalOverdueOpen = overdueZone.reduce(function (a, z) { return a + z.open; }, 0);
  var totalOverdue = overdueZone.reduce(function (a, z) { return a + z.grandTotal; }, 0);

  // --- grid builder (1-indexed row/col, matching ws.cell(row,col,value)) ---
  function setCell(grid, row1, col1, value) {
    var r = row1 - 1, c = col1 - 1;
    while (grid.length <= r) grid.push([]);
    while (grid[r].length <= c) grid[r].push(null);
    grid[r][c] = value;
  }

  // Sheet 1: Auto Pivot Summary
  var g1 = [];
  setCell(g1, 1, 1, "Pivot-style summary from 'OneTrust Assessment'");
  setCell(g1, 3, 1, 'Fields used');
  setCell(g1, 4, 1, 'Rows'); setCell(g1, 4, 2, 'Organization');
  setCell(g1, 5, 1, 'Columns'); setCell(g1, 5, 2, 'Stage');
  setCell(g1, 6, 1, 'Values'); setCell(g1, 6, 2, 'Count of ID');
  var pivotHeaderRow = ['Organization'].concat(pivotHeaders);
  pivotHeaderRow.forEach(function (h, i) { setCell(g1, 8, i + 1, h); });
  ORG_ORDER.forEach(function (org, i) {
    setCell(g1, 9 + i, 1, org);
    orgWithGT[i].forEach(function (v, c) { setCell(g1, 9 + i, c + 2, v); });
  });
  setCell(g1, 9 + ORG_ORDER.length, 1, 'Grand Total');
  pivotTotalsRow.forEach(function (v, c) { setCell(g1, 9 + ORG_ORDER.length, c + 2, v); });

  // Sheet 2: Auto Open Closed
  var g2 = [];
  setCell(g2, 1, 1, 'Open vs Closed by Zone');
  setCell(g2, 3, 1, 'Business rule');
  setCell(g2, 4, 1, 'Closed = Under review + Completed');
  setCell(g2, 5, 1, 'Open = In progress + Not started');
  ['Zone', 'Closed', 'Open', 'Grand Total'].forEach(function (h, i) { setCell(g2, 8, i + 1, h); });
  zoneSummary.forEach(function (z, i) {
    setCell(g2, 9 + i, 1, z.zone);
    setCell(g2, 9 + i, 2, z.closed);
    setCell(g2, 9 + i, 3, z.open);
    setCell(g2, 9 + i, 4, z.grandTotal);
  });
  var totalRow2 = 9 + ZONE_ORDER.length + 1;
  setCell(g2, totalRow2, 1, 'Total');
  setCell(g2, totalRow2, 2, totalClosed);
  setCell(g2, totalRow2, 3, totalOpen);
  setCell(g2, totalRow2, 4, totalAll);

  // Sheet 3: Auto Overdue {N}D
  var g3 = [];
  setCell(g3, 1, 1, 'Overdue Assessment Summary - Threshold = ' + OVERDUE_THRESHOLD + ' days');
  setCell(g3, 3, 1, 'Rule');
  setCell(g3, 4, 1, 'Overdue = Ageing >= ' + OVERDUE_THRESHOLD + ' days');
  setCell(g3, 5, 1, 'Closed overdue = Completed + Under review');
  setCell(g3, 6, 1, 'Open overdue = In progress + Not started');
  ['Zone', 'Closed Overdue', 'Open Overdue', 'Overdue Total'].forEach(function (h, i) { setCell(g3, 9, i + 1, h); });
  overdueZone.forEach(function (z, i) {
    setCell(g3, 10 + i, 1, z.zone);
    setCell(g3, 10 + i, 2, z.closed);
    setCell(g3, 10 + i, 3, z.open);
    setCell(g3, 10 + i, 4, z.grandTotal);
  });
  var totalRow3 = 10 + ZONE_ORDER.length + 1;
  setCell(g3, totalRow3, 1, 'Total');
  setCell(g3, totalRow3, 2, totalOverdueClosed);
  setCell(g3, totalRow3, 3, totalOverdueOpen);
  setCell(g3, totalRow3, 4, totalOverdue);

  var sheetName3 = 'Auto Overdue ' + OVERDUE_THRESHOLD + 'D';
  var files = [{
    name: 'OneTrust_Report.xlsx',
    sheets: [
      { name: 'Auto Pivot Summary', grid: g1 },
      { name: 'Auto Open Closed', grid: g2 },
      { name: sheetName3, grid: g3 },
    ],
  }];

  // Embedded charts reproduce the original matplotlib style exactly (white
  // background, blue "Closed" + orange "Open" stacked bars, white in-bar
  // value labels, bold total label above each bar, bottom legend, dashed
  // gridlines) instead of the light in-page preview theme. Both charts are
  // embedded into the file (matching today's 2-chart-per-file output) even
  // though the preview only ever surfaces the first one — confirmed with
  // the user as the intended behavior, not a bug to fix.
  function s2aChartTraces(zoneArr) {
    var zones = zoneArr.map(function (z) { return z.zone; });
    var closedVals = zoneArr.map(function (z) { return z.closed; });
    var openVals = zoneArr.map(function (z) { return z.open; });
    var totals = zoneArr.map(function (z) { return z.closed + z.open; });
    return {
      zones: zones, totals: totals,
      traces: [
        { x: zones, y: closedVals, type: 'bar', name: 'Closed', marker: { color: '#2F75B5' },
          text: closedVals.map(function (v) { return v > 0 ? String(v) : ''; }),
          textposition: 'inside', insidetextanchor: 'middle', textangle: 0, constraintext: 'none', textfont: { color: '#ffffff', size: 14 } },
        { x: zones, y: openVals, type: 'bar', name: 'Open', marker: { color: '#ED7D31' },
          text: openVals.map(function (v) { return v > 0 ? String(v) : ''; }),
          textposition: 'inside', insidetextanchor: 'middle', textangle: 0, constraintext: 'none', textfont: { color: '#ffffff', size: 14 } },
      ],
    };
  }
  function s2aLayout(title, yTitle, zones, totals) {
    var maxTotal = Math.max.apply(null, totals.concat([1]));
    return {
      barmode: 'stack',
      paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
      title: { text: '<b>' + title + '</b>', font: { color: '#000000', size: 14 } },
      xaxis: { title: 'Zone', tickfont: { color: '#000000' } },
      yaxis: { title: yTitle, gridcolor: 'rgba(0,0,0,0.3)', griddash: 'dash', range: [0, maxTotal * 1.2] },
      legend: { orientation: 'h', x: 0.5, xanchor: 'center', y: -0.25, font: { color: '#000000' } },
      annotations: zones.map(function (z, i) {
        return totals[i] > 0
          ? { x: z, y: totals[i], text: String(totals[i]), showarrow: false, yshift: 12, font: { color: '#000000', size: 9 } }
          : null;
      }).filter(Boolean),
      margin: { t: 60, r: 20, b: 90, l: 60 },
    };
  }
  // Live Plotly preview of the same styled chart (no baked image). Only the
  // 'Auto Open Closed' sheet previews a chart — matching what the preview
  // always showed; the Overdue sheet's chart lives in the Excel file only.
  var chartConfigs = {};
  var oc = s2aChartTraces(zoneSummary);
  chartConfigs['Auto Open Closed'] = {
    traces: oc.traces,
    layout: s2aLayout(totalOpen + '/' + totalAll + ' Open Assessment', 'Assessment Count', oc.zones, oc.totals),
  };

  var workbook = new ExcelJS.Workbook();
  function writeSheet(name, grid) {
    var ws = workbook.addWorksheet(name);
    grid.forEach(function (r) { ws.addRow(r); });
    ws.columns.forEach(function (c) { c.width = 18; });
    return ws;
  }
  writeSheet('Auto Pivot Summary', g1);
  var ws2 = writeSheet('Auto Open Closed', g2);
  var ws3 = writeSheet(sheetName3, g3);

  // Native, editable stacked charts (data-linked to hidden helper blocks)
  // replace the two baked PNGs. Series: Closed (blue) + Open (orange).
  var placements = [];
  var R = window.NativeChartInject && window.NativeChartInject.ref;
  // headerRow: g2's zone table header is row 8 (data 9..), g3's is row 9 (data 10..).
  // Zone labels sit in column A (editable there); B=Closed, C=Open.
  function s2aPlacement(sheetName, zoneArr, title, headerRow) {
    var zones = zoneArr.map(function (z) { return z.zone; });
    if (!zones.length) return;
    var first = headerRow + 1, last = headerRow + zones.length;
    placements.push({
      sheetName: sheetName, anchor: { fromCol: 5, fromRow: 1, toCol: 15, toRow: 20 }, // ~"F2"
      def: {
        grouping: 'stacked', legend: true, title: title,
        axisColor: '000000', dataLabels: { position: 'ctr', color: 'FFFFFF' },
        categories: { ref: R(sheetName, 1, first, last), cache: zones },
        series: [
          { name: { lit: 'Closed' },
            values: { ref: R(sheetName, 2, first, last), cache: zoneArr.map(function (z) { return z.closed; }) }, color: '2F75B5' },
          { name: { lit: 'Open' },
            values: { ref: R(sheetName, 3, first, last), cache: zoneArr.map(function (z) { return z.open; }) }, color: 'ED7D31' },
        ],
      },
    });
  }
  if (window.NativeChartInject && window.fflate) {
    s2aPlacement('Auto Open Closed', zoneSummary, totalOpen + '/' + totalAll + ' Open Assessment', 8);
    s2aPlacement(sheetName3, overdueZone,
      totalOverdueOpen + '/' + totalOverdue + ' Overdue Open Assessment (' + OVERDUE_THRESHOLD + '+ days)', 9);
  }

  var buf = await workbook.xlsx.writeBuffer();
  if (placements.length) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('s2a native chart inject failed:', e); }
  }

  return {
    ok: true,
    files: [{
      name: 'OneTrust_Report.xlsx',
      bytes: buf,
      sheets: [
        { name: 'Auto Pivot Summary', grid: g1 },
        { name: 'Auto Open Closed', grid: g2 },
        { name: sheetName3, grid: g3 },
      ],
    }],
    chartConfigs: chartConfigs,
  };
};
