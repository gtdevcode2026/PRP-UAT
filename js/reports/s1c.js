// Port of "Slide 12 1st daigram Automation/automation3.py" (id: s1c).
// Reads "TPRM Web-Portal Export", pivots zone_assessing x assessment_status
// (count, margins), ensures ACTIVE/Active/Deprioritized/Duplicate columns
// exist (0-filled if missing), merges case-variant Active columns into
// Active_Total, and writes "PRP_Final_Output3.xlsx" with 3 sheets: Pivot,
// Summary (status overview chart), Active by Zone (per-zone chart).
window.Reports.s1c = async function s1c(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'TPRM Web-Portal Export');
  function cleanKey(k) { return String(k).trim().toLowerCase(); }
  var rows = sheet.rows.map(function (row) {
    var out = {};
    Object.keys(row).forEach(function (k) { out[cleanKey(k)] = row[k]; });
    return out;
  });

  var pivot = E.pivotCount(
    rows,
    // GROWTH / growth arrive as case-variant zones — fold them into one
    // 'Growth' bucket so the Pivot sheet and Active-by-Zone chart show a
    // single combined column. Compared on letters only, so invisible
    // characters (zero-width space, NBSP) or stray punctuation in the
    // export can't split the bucket. Other zones keep their raw casing.
    function (r) { var z = String(r.zone_assessing); return z.replace(/[^a-z]/gi, '').toLowerCase() === 'growth' ? 'Growth' : z; },
    function (r) { return String(r.assessment_status); },
    { margins: true, marginsName: 'Grand Total' }
  );

  var headers = pivot.headers.slice();
  var dataRows = pivot.rows.map(function (r) { return r.slice(); });

  ['ACTIVE', 'Active', 'Deprioritized', 'Duplicate'].forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      headers.push(col);
      dataRows.forEach(function (r) { r.push(0); });
    }
  });

  var activeUpperIdx = headers.indexOf('ACTIVE');
  var activeIdx = headers.indexOf('Active');
  headers.push('Active_Total');
  dataRows.forEach(function (r) {
    var a1 = activeUpperIdx >= 0 ? (r[activeUpperIdx] || 0) : 0;
    var a2 = activeIdx >= 0 ? (r[activeIdx] || 0) : 0;
    r.push(a1 + a2);
  });

  var pivotHeaderRow = ['Zone'].concat(headers);
  var pivotOutRows = pivot.indexVals.map(function (iv, i) { return [iv].concat(dataRows[i]); });
  var pivotOutGrid = [pivotHeaderRow].concat(pivotOutRows);

  var grandIdx = pivot.indexVals.indexOf('Grand Total');
  var grandRow = dataRows[grandIdx];
  var activeTotalIdx = headers.indexOf('Active_Total');
  var deprioritizedIdx = headers.indexOf('Deprioritized');
  var duplicateIdx = headers.indexOf('Duplicate');
  var summaryGrid = [
    ['Status', 'Count'],
    ['Active', grandRow[activeTotalIdx]],
    ['Deprioritized', grandRow[deprioritizedIdx]],
    ['Duplicate', grandRow[duplicateIdx]],
  ];

  var zoneActiveGrid = [['Zone', 'Active']].concat(
    pivot.indexVals
      .map(function (iv, i) { return [iv, dataRows[i][activeTotalIdx]]; })
      .filter(function (pair) { return pair[0] !== 'Grand Total'; })
  );

  var files = [{
    name: 'PRP_Final_Output3.xlsx',
    sheets: [
      { name: 'Pivot', grid: pivotOutGrid },
      { name: 'Summary', grid: summaryGrid },
      { name: 'Active by Zone', grid: zoneActiveGrid },
    ],
  }];

  // Embedded charts reproduce the original openpyxl native BarChart look
  // (white background, single blue series, value labels above each bar,
  // axis titles) instead of the light in-page preview theme.
  function excelBarChart(categories, values, title, xTitle, yTitle) {
    return {
      traces: [{
        x: categories, y: values, type: 'bar',
        marker: { color: '#4472C4' },
        text: values.map(String), textposition: 'outside', textfont: { color: '#000000', size: 10 },
      }],
      layout: {
        paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
        title: { text: title, font: { color: '#000000', size: 13 } },
        xaxis: { title: xTitle, tickfont: { color: '#000000' }, linecolor: '#808080', showline: true },
        yaxis: { title: yTitle, tickfont: { color: '#000000' }, gridcolor: '#D9D9D9', zerolinecolor: '#808080' },
        showlegend: false,
        margin: { t: 50, r: 20, b: 50, l: 50 },
      },
    };
  }
  var statusRows = summaryGrid.slice(1);
  var zoneRows = zoneActiveGrid.slice(1);
  var chart1 = excelBarChart(
    statusRows.map(function (r) { return r[0]; }), statusRows.map(function (r) { return r[1]; }),
    'Assessment Status Overview', 'Assessment Status', 'Count'
  );
  var chart2 = excelBarChart(
    zoneRows.map(function (r) { return r[0]; }), zoneRows.map(function (r) { return r[1]; }),
    'Active by Zone', 'Zone', 'Count'
  );
  var images = {
    'Summary': await B.renderStyledPng(chart1.traces, chart1.layout, 480, 280),
    'Active by Zone': await B.renderStyledPng(chart2.traces, chart2.layout, 480, 280),
  };

  var workbook = new ExcelJS.Workbook();
  function writeSheet(name, grid) {
    var ws = workbook.addWorksheet(name);
    grid.forEach(function (r) { ws.addRow(r); });
    ws.columns.forEach(function (c) { c.width = 16; });
    return ws;
  }
  writeSheet('Pivot', pivotOutGrid);
  var wsSummary = writeSheet('Summary', summaryGrid);
  var wsZone = writeSheet('Active by Zone', zoneActiveGrid);

  // Native, editable single-series charts (data-linked to hidden helper
  // blocks) replace the two baked PNGs.
  var placements = [];
  var R = window.NativeChartInject && window.NativeChartInject.ref;
  // Both tables: header row 1 (A=label, B=Count), data rows 2..1+n. The status
  // / zone labels sit in column A and are editable there.
  function s1cPlacement(sheetName, rows, title) {
    var cats = rows.map(function (r) { return r[0]; });
    if (!cats.length) return;
    var last = 1 + cats.length;
    placements.push({
      sheetName: sheetName, anchor: { fromCol: 3, fromRow: 1, toCol: 12, toRow: 18 }, // ~"D2"
      def: {
        grouping: 'clustered', legend: false, title: title,
        axisColor: '000000', dataLabels: { position: 'outEnd', color: '000000' },
        categories: { ref: R(sheetName, 1, 2, last), cache: cats },
        series: [
          { name: { ref: R(sheetName, 2, 1, 1), lit: 'Count' },
            values: { ref: R(sheetName, 2, 2, last), cache: rows.map(function (r) { return r[1]; }) }, color: '4472C4' },
        ],
      },
    });
  }
  if (window.NativeChartInject && window.fflate) {
    s1cPlacement('Summary', statusRows, 'Assessment Status Overview');
    s1cPlacement('Active by Zone', zoneRows, 'Active by Zone');
  }

  var buf = await workbook.xlsx.writeBuffer();
  if (placements.length) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('s1c native chart inject failed:', e); }
  }

  return {
    ok: true,
    files: [{
      name: 'PRP_Final_Output3.xlsx',
      bytes: buf,
      sheets: [
        { name: 'Pivot', grid: pivotOutGrid },
        { name: 'Summary', grid: summaryGrid },
        { name: 'Active by Zone', grid: zoneActiveGrid },
      ],
    }],
    chartImages: images,
  };
};
