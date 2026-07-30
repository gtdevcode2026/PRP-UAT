// Port of "daigram 1 automation/automation.py" (id: d1).
// Reads "TPRM Web-Portal Export", filters request_date year==2026 &
// category=="TECHNOLOGY", counts suppliers added per zone, merges onto a
// static 7-zone Tier-1 table, appends a Total row, embeds a chart, writes
// "PRP_Output.xlsx" sheet "Final Table" (chart anchored ~E2 as in the
// original openpyxl `ws.add_image(img, "E2")`).
window.Reports.d1 = async function d1(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'TPRM Web-Portal Export');

  function cleanKey(k) { return String(k).trim().toLowerCase().replace(/ /g, '_'); }
  var rows = sheet.rows.map(function (row) {
    var out = {};
    Object.keys(row).forEach(function (k) { out[cleanKey(k)] = row[k]; });
    return out;
  });

  var filtered = rows.filter(function (r) {
    var year = E.excelYear(r.request_date);
    if (year === null) return false;
    var cat = E.isBlank(r.category) ? '' : String(r.category).trim().toUpperCase();
    return year === 2026 && cat === 'TECHNOLOGY';
  });

  // groupby("supplier_zone")["id"].count() - pandas .count() counts non-null values.
  var groups = E.groupBy(filtered, function (r) { return r.supplier_zone; });
  var addedByZone = {};
  groups.forEach(function (zoneRows, zone) {
    addedByZone[zone] = zoneRows.filter(function (r) { return !E.isBlank(r.id); }).length;
  });

  var TIER1 = [
    ['NAZ', 123], ['AFR', 115], ['GHQ', 79], ['EUR', 35], ['APAC', 17], ['SAZ', 13], ['MAZ', 13],
  ];
  var chartRows = TIER1.map(function (t) {
    var zone = t[0], tier1Supplier = t[1];
    var added = addedByZone.hasOwnProperty(zone) ? addedByZone[zone] : 0;
    return [zone, tier1Supplier, added];
  });
  var tier1Sum = TIER1.reduce(function (a, t) { return a + t[1]; }, 0);
  var addedSum = chartRows.reduce(function (a, r) { return a + r[2]; }, 0);
  var finalRows = chartRows.concat([['Total', tier1Sum, addedSum]]);

  var headers = ['Zone', 'Tier-1 Supplier', 'Supplier Added by Zone'];
  var grid = [headers].concat(finalRows);
  var files = [{ name: 'PRP_Output.xlsx', sheets: [{ name: 'Final Table', grid: grid }] }];

  // Styled chart reproduces the original matplotlib style exactly (black
  // figure, blue/orange stacked bars, white in-bar value labels, bold white
  // title with the running total). The SAME config drives the live Plotly
  // "Chart preview" (via chartConfigs below) and mirrors the native Excel
  // chart injected into the workbook - preview and file always match.
  var zones = chartRows.map(function (r) { return r[0]; });
  var tier1Vals = chartRows.map(function (r) { return r[1]; });
  var addedVals = chartRows.map(function (r) { return r[2]; });
  var d1Traces = [
    {
      x: zones, y: tier1Vals, type: 'bar', name: 'Tier-1 Supplier',
      marker: { color: '#1f77b4' },
      text: tier1Vals.map(String), textposition: 'inside', insidetextanchor: 'middle', textangle: 0,
      textfont: { color: '#ffffff', size: 10 },
    },
    {
      x: zones, y: addedVals, type: 'bar', name: 'Supplier Added by Zone',
      marker: { color: '#ff7f0e' },
      text: addedVals.map(function (v) { return v > 0 ? String(v) : ''; }),
      textposition: 'inside', insidetextanchor: 'middle', textangle: 0,
      textfont: { color: '#ffffff', size: 10 },
    },
  ];
  var d1Layout = {
    barmode: 'stack',
    paper_bgcolor: '#000000', plot_bgcolor: '#000000',
    title: { text: '<b>(' + tier1Sum + ') Zone wise Tier 1 Suppliers</b>', font: { color: '#ffffff', size: 14 } },
    xaxis: { tickfont: { color: '#ffffff' }, showline: true, linecolor: '#ffffff', mirror: true },
    yaxis: { tickfont: { color: '#ffffff' }, showline: true, linecolor: '#ffffff', mirror: true, showgrid: false },
    legend: { bgcolor: '#000000', font: { color: '#ffffff' } },
    margin: { t: 60, r: 20, b: 50, l: 50 },
  };
  var chartConfigs = { 'Final Table': { traces: d1Traces, layout: d1Layout } };

  var workbook = new ExcelJS.Workbook();
  var ws = workbook.addWorksheet('Final Table');
  ws.addRow(headers);
  finalRows.forEach(function (r) { ws.addRow(r); });
  ws.columns.forEach(function (col) { col.width = 20; });

  // Native, editable stacked chart (data-linked to a hidden helper block) in
  // place of the baked PNG. The preview still uses images['Final Table'].
  var placements = [];
  if (window.NativeChartInject && window.fflate && zones.length) {
    // Data-link to the visible Final Table cells (A=zone, B/C=series) so the
    // zone labels + legend names are editable directly in Excel. Rows 2..1+n
    // are the chart rows (the appended Total row is excluded).
    var R = window.NativeChartInject.ref, last = zones.length + 1;
    placements.push({
      sheetName: 'Final Table', anchor: { fromCol: 4, fromRow: 1, toCol: 13, toRow: 18 }, // ~"E2"
      def: {
        grouping: 'stacked', legend: true, title: '(' + tier1Sum + ') Zone wise Tier 1 Suppliers',
        chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
        dataLabels: { position: 'ctr', color: 'FFFFFF' },
        categories: { ref: R('Final Table', 1, 2, last), cache: zones },
        series: [
          { name: { ref: R('Final Table', 2, 1, 1), lit: 'Tier-1 Supplier' },
            values: { ref: R('Final Table', 2, 2, last), cache: tier1Vals }, color: '1F77B4' },
          { name: { ref: R('Final Table', 3, 1, 1), lit: 'Supplier Added by Zone' },
            values: { ref: R('Final Table', 3, 2, last), cache: addedVals }, color: 'FF7F0E' },
        ],
      },
    });
  }

  var buf = await workbook.xlsx.writeBuffer();
  if (placements.length) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('d1 native chart inject failed:', e); }
  }

  return {
    ok: true,
    files: [{ name: 'PRP_Output.xlsx', bytes: buf, sheets: [{ name: 'Final Table', grid: grid }] }],
    chartConfigs: chartConfigs,
  };
};
