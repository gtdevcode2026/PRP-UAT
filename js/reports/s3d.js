// Port of "Slide 12 3rd daigram Automation/automation4.py" (id: s3d).
// Reads "OneTrust - Risk Export", finds Organization/ID/Stage/Aging|Ageing
// columns by case-insensitive alias lookup, extracts a numeric day-count
// from the Aging column via regex (defaulting to 0 when no digits are
// found — this makes every row "Open" on the shipped sample data, since its
// Aging column holds text, not numbers; reproduced deliberately, not
// "fixed" — confirmed with the user), classifies Open/Overdue (>90 days),
// merges BEES + BEES|FINTECH into a single org named "Growth" before any
// grouping, pivots Organization x Risk_Status, maps orgs to zones, and
// builds a separate Zone Summary (Total/Open counts by org name). Writes
// "Risk_Output.xlsx" with 3 sheets; charts anchored ~E4.
window.Reports.s3d = async function s3d(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'OneTrust - Risk Export');
  var rawHeaders = sheet.headers;
  var cleanedHeaders = rawHeaders.map(function (h) { return String(h).replace(/\r/g, '').replace(/\n/g, '').trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    rawHeaders.forEach(function (h, i) { out[cleanedHeaders[i]] = row[h]; });
    return out;
  });

  function findColumn(possibleNames) {
    var lookup = {};
    cleanedHeaders.forEach(function (c) { lookup[String(c).trim().toLowerCase()] = c; });
    for (var i = 0; i < possibleNames.length; i++) {
      var key = possibleNames[i].toLowerCase();
      if (lookup.hasOwnProperty(key)) return lookup[key];
    }
    return null;
  }
  var organizationCol = findColumn(['Organization']);
  var idCol = findColumn(['ID']);
  var stageCol = findColumn(['Stage']);
  var agingCol = findColumn(['Aging', 'Ageing']);
  var missing = [];
  if (!organizationCol) missing.push('Organization');
  if (!idCol) missing.push('ID');
  if (!stageCol) missing.push('Stage');
  if (!agingCol) missing.push('Aging or Ageing');
  if (missing.length) throw new Error('Missing required columns: ' + missing.join(', '));

  rows.forEach(function (r) {
    r[organizationCol] = E.isBlank(r[organizationCol]) ? '' : String(r[organizationCol]).trim();
    // BEES and BEES | FINTECH report as one combined org named Growth —
    // normalized here, before any grouping, so the pivot, Zone Summary and
    // their charts all see a single merged row. Compared on letters only,
    // so case, spacing, punctuation, and invisible characters (zero-width
    // space, NBSP) in the export can't split the bucket.
    var orgLetters = r[organizationCol].replace(/[^a-z]/gi, '').toLowerCase();
    if (orgLetters === 'bees' || orgLetters === 'beesfintech') r[organizationCol] = 'Growth';
    r[stageCol] = E.isBlank(r[stageCol]) ? '' : String(r[stageCol]).trim();
    r.Stage_Clean = r[stageCol].toLowerCase().trim();
  });

  function extractNumber(value) {
    if (E.isBlank(value)) return null;
    var m = String(value).trim().match(/\d+(\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  }
  rows.forEach(function (r) {
    var n = extractNumber(r[agingCol]);
    r.Aging_Days = (n === null) ? 0 : n;
  });

  var SELECTED_STAGES = { evaluation: 1, identified: 1, treatment: 1 };
  var filtered = rows
    .filter(function (r) { return SELECTED_STAGES.hasOwnProperty(r.Stage_Clean); })
    .filter(function (r) { return r[organizationCol] !== ''; });
  filtered.forEach(function (r) { r.Risk_Status = r.Aging_Days > 90 ? 'Overdue' : 'Open'; });

  var pivot = E.pivotCount(
    filtered,
    function (r) { return String(r[organizationCol]); },
    function (r) { return r.Risk_Status; },
    { margins: true, marginsName: 'Grand Total' }
  );
  var headers = pivot.headers.slice();
  var dataRows = pivot.rows.map(function (r) { return r.slice(); });
  ['Open', 'Overdue'].forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      headers.push(col);
      dataRows.forEach(function (r) { r.push(0); });
    }
  });
  var openIdx = headers.indexOf('Open'), overdueIdx = headers.indexOf('Overdue');
  if (headers.indexOf('Grand Total') === -1) {
    headers.push('Grand Total');
    dataRows.forEach(function (r) { r.push((r[openIdx] || 0) + (r[overdueIdx] || 0)); });
  }
  var gtIdx = headers.indexOf('Grand Total');
  var openOverdueGrid = [['Organization', 'Open', 'Overdue', 'Grand Total']].concat(
    pivot.indexVals.map(function (iv, i) {
      var r = dataRows[i];
      return [iv, r[openIdx] || 0, r[overdueIdx] || 0, r[gtIdx] || 0];
    })
  );

  var ZONE_MAP = {
    'Africa': 'AFR', 'APAC': 'APAC', 'Growth': 'GRO',
    'Europe': 'EUR', 'GHQ': 'GHQ', 'Middle America Zone': 'MAZ',
    'North America Zone': 'NAZ', 'South America Zone': 'SAZ',
  };
  var zoneAgg = {};
  openOverdueGrid.slice(1).forEach(function (r) {
    if (r[0] === 'Grand Total') return;
    var zone = ZONE_MAP[r[0]];
    if (zone === undefined) return; // dropna(subset=["Zones"]) equivalent
    if (!zoneAgg[zone]) zoneAgg[zone] = { open: 0, overdue: 0 };
    zoneAgg[zone].open += r[1];
    zoneAgg[zone].overdue += r[2];
  });
  var ZONE_ORDER = ['AFR', 'APAC', 'GRO', 'EUR', 'GHQ', 'MAZ', 'NAZ', 'SAZ'];
  var zonesPresent = Object.keys(zoneAgg).sort(function (a, b) {
    var ia = ZONE_ORDER.indexOf(a); if (ia === -1) ia = 999;
    var ib = ZONE_ORDER.indexOf(b); if (ib === -1) ib = 999;
    return ia - ib;
  });
  var openVsOverdueGrid = [['Zones', 'Open', 'Overdue']].concat(
    zonesPresent.map(function (z) { return [z, zoneAgg[z].open, zoneAgg[z].overdue]; })
  );

  var TOTAL_STAGES = { evaluation: 1, identified: 1, treatment: 1, monitoring: 1 };
  var OPEN_STAGES = { evaluation: 1, identified: 1, treatment: 1 };
  function countByOrg(rowsSubset) {
    var m = {};
    rowsSubset.forEach(function (r) {
      if (E.isBlank(r[idCol])) return; // pandas .count() skips nulls
      var k = r[organizationCol];
      m[k] = (m[k] || 0) + 1;
    });
    return m;
  }
  var totalCounts = countByOrg(rows.filter(function (r) { return TOTAL_STAGES.hasOwnProperty(r.Stage_Clean); }));
  var openCounts = countByOrg(rows.filter(function (r) { return OPEN_STAGES.hasOwnProperty(r.Stage_Clean); }));
  var zoneSummaryGrid = [['Zones', 'Total Risks', 'Open Risks']].concat(
    Object.keys(totalCounts).sort() // groupby(..., sort default True)
      .filter(function (o) { return o !== ''; })
      .map(function (o) { return [o, totalCounts[o], openCounts[o] || 0]; })
  );

  var files = [{
    name: 'Risk_Output.xlsx',
    sheets: [
      { name: 'Open Overdue Pivot', grid: openOverdueGrid },
      { name: 'Open vs Overdue', grid: openVsOverdueGrid },
      { name: 'Zone Summary', grid: zoneSummaryGrid },
    ],
  }];

  // Embedded charts reproduce the original xlsxwriter native stacked column
  // charts (black chart/plot area, white text, dark gridlines, bottom
  // legend, white in-bar value labels) instead of the light preview theme.
  function s3dLayout(title) {
    return {
      barmode: 'stack',
      paper_bgcolor: '#000000', plot_bgcolor: '#000000',
      title: { text: '<b>' + title + '</b>', font: { color: '#ffffff', size: 14 } },
      xaxis: { tickfont: { color: '#ffffff' }, linecolor: '#ffffff', showline: true },
      yaxis: { tickfont: { color: '#ffffff' }, gridcolor: '#444444' },
      legend: { orientation: 'h', x: 0.5, xanchor: 'center', y: -0.2, font: { color: '#ffffff' } },
      margin: { t: 50, r: 20, b: 70, l: 50 },
    };
  }
  var images = {};
  var oovDataRows = openVsOverdueGrid.slice(1);
  if (oovDataRows.length) {
    var oovZones = oovDataRows.map(function (r) { return r[0]; });
    var oovOpen = oovDataRows.map(function (r) { return r[1]; });
    var oovOverdue = oovDataRows.map(function (r) { return r[2]; });
    var oovTraces = [
      { x: oovZones, y: oovOpen, type: 'bar', name: 'Open', marker: { color: '#156082' },
        text: oovOpen.map(String), textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff' } },
      { x: oovZones, y: oovOverdue, type: 'bar', name: 'Overdue', marker: { color: '#C00000' },
        text: oovOverdue.map(String), textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff' } },
    ];
    images['Open vs Overdue'] = await B.renderStyledPng(oovTraces, s3dLayout('Open vs Overdue Risks'), 720, 420);
  }
  var zsDataRows = zoneSummaryGrid.slice(1);
  if (zsDataRows.length) {
    var zsZones = zsDataRows.map(function (r) { return r[0]; });
    var zsTotal = zsDataRows.map(function (r) { return r[1]; });
    var zsOpen = zsDataRows.map(function (r) { return r[2]; });
    var zsTraces = [
      { x: zsZones, y: zsTotal, type: 'bar', name: 'Total Risks', marker: { color: '#156082' },
        text: zsTotal.map(String), textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff' } },
      { x: zsZones, y: zsOpen, type: 'bar', name: 'Open Risks', marker: { color: '#F26C23' },
        text: zsOpen.map(String), textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff' } },
    ];
    images['Zone Summary'] = await B.renderStyledPng(zsTraces, s3dLayout('Zone wise Risks'), 720, 420);
  }

  var workbook = new ExcelJS.Workbook();
  function writeSheet(name, grid) {
    var ws = workbook.addWorksheet(name);
    grid.forEach(function (r) { ws.addRow(r); });
    ws.columns.forEach(function (c) { c.width = 20; });
    return ws;
  }
  writeSheet('Open Overdue Pivot', openOverdueGrid);
  var wsOpenVsOverdue = writeSheet('Open vs Overdue', openVsOverdueGrid);
  var wsZoneSummary = writeSheet('Zone Summary', zoneSummaryGrid);

  // Charts are written as NATIVE, editable Excel charts (not baked PNGs) so
  // bar colors and legend/series names can be changed in Excel. The preview
  // still uses the renderStyledPng images above; only the embedded file chart
  // becomes native. Data-linked to the A/B/C columns already on each sheet.
  var buf = await workbook.xlsx.writeBuffer();

  var placements = [];
  if (images['Open vs Overdue']) {
    var oovLast = oovDataRows.length + 1;
    placements.push({
      sheetName: 'Open vs Overdue',
      anchor: { fromCol: 4, fromRow: 3, toCol: 13, toRow: 20 }, // ~"E4"
      def: {
        grouping: 'stacked', legend: true, title: 'Open vs Overdue Risks',
        chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
        dataLabels: { position: 'ctr', color: 'FFFFFF' },
        categories: { ref: "'Open vs Overdue'!$A$2:$A$" + oovLast, cache: oovZones },
        series: [
          { name: { ref: "'Open vs Overdue'!$B$1", lit: 'Open' },
            values: { ref: "'Open vs Overdue'!$B$2:$B$" + oovLast, cache: oovOpen }, color: '156082' },
          { name: { ref: "'Open vs Overdue'!$C$1", lit: 'Overdue' },
            values: { ref: "'Open vs Overdue'!$C$2:$C$" + oovLast, cache: oovOverdue }, color: 'C00000' },
        ],
      },
    });
  }
  if (images['Zone Summary']) {
    var zsLast = zsDataRows.length + 1;
    placements.push({
      sheetName: 'Zone Summary',
      anchor: { fromCol: 4, fromRow: 3, toCol: 13, toRow: 20 }, // ~"E4"
      def: {
        grouping: 'stacked', legend: true, title: 'Zone wise Risks',
        chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
        dataLabels: { position: 'ctr', color: 'FFFFFF' },
        categories: { ref: "'Zone Summary'!$A$2:$A$" + zsLast, cache: zsZones },
        series: [
          { name: { ref: "'Zone Summary'!$B$1", lit: 'Total Risks' },
            values: { ref: "'Zone Summary'!$B$2:$B$" + zsLast, cache: zsTotal }, color: '156082' },
          { name: { ref: "'Zone Summary'!$C$1", lit: 'Open Risks' },
            values: { ref: "'Zone Summary'!$C$2:$C$" + zsLast, cache: zsOpen }, color: 'F26C23' },
        ],
      },
    });
  }
  if (placements.length && window.NativeChartInject && window.fflate) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('s3d native chart inject failed, file keeps data without native chart:', e); }
  }

  return {
    ok: true,
    files: [{
      name: 'Risk_Output.xlsx',
      bytes: buf,
      sheets: [
        { name: 'Open Overdue Pivot', grid: openOverdueGrid },
        { name: 'Open vs Overdue', grid: openVsOverdueGrid },
        { name: 'Zone Summary', grid: zoneSummaryGrid },
      ],
    }],
    chartImages: images,
  };
};
