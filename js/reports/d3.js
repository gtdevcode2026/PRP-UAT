// Port of "daigram 3 automation/automation.py" (id: d3; the embedded/shipped
// version uses relative paths — "PRP Sample Jun (2).xlsx" / "output.xlsx" —
// the standalone .py on disk has stale absolute paths from the original
// author's machine, not used here).
// Reads "OneTrust Assessment", auto-detects Stage/Organization/Working1/
// Working2 columns by substring, filters Working2=="Beyond 1 Year Overdue" &
// Working1 in [Pending, Completed in 2026], pivots Organization x Working1
// (count), appends Grand Total column then row, writes "output.xlsx" sheet
// "Summary" with a native stacked-bar-style chart anchored ~H5.
window.Reports.d3 = async function d3(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers;
  var trimmedHeaders = rawHeaders.map(function (h) { return String(h).trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    rawHeaders.forEach(function (h, i) { out[trimmedHeaders[i]] = row[h]; });
    return out;
  });

  function findCol(substr) {
    for (var i = 0; i < trimmedHeaders.length; i++) {
      if (trimmedHeaders[i].indexOf(substr) !== -1) return trimmedHeaders[i];
    }
    throw new Error('Column containing "' + substr + '" not found');
  }
  var orgCol = findCol('Organization');
  var w1Col = findCol('Working1');
  var w2Col = findCol('Working2');
  // stage_col/classify_stage only feed "Final Status", which is used purely
  // as the pivot's *value* column for counting — its actual text (Completed/
  // Open) never appears in the output, so it needn't be computed here.

  var filtered = rows.filter(function (r) {
    return r[w2Col] === 'Beyond 1 Year Overdue' && (r[w1Col] === 'Pending' || r[w1Col] === 'Completed in 2026');
  });

  var pivot = E.pivotCount(
    filtered,
    // BEES / Growth variants report as the single 'GRO' organization.
    function (r) { return E.sanitizeZone(r[orgCol]); },
    function (r) { return String(r[w1Col]); },
    { margins: false }
  );

  // The script charts a FIXED window of columns B:C — "ONLY Pending +
  // Completed in 2026, exclude Grand Total column" — and colours series[0] red
  // / series[1] orange, so it presumes both Working1 columns are present in
  // that order. pivot_table only emits the values that actually occur, so when
  // the filter yields just one of them (all-Pending data, say) that fixed
  // window slides onto Grand Total and charts the total as if it were a
  // status. Pin both columns, zero-filled, the way the d2 script pins
  // Open/Closed. The filter above admits no other value, so this is exhaustive.
  var W1_COLUMNS = ['Completed in 2026', 'Pending'];
  var baseRows = pivot.rows.map(function (r) {
    return W1_COLUMNS.map(function (name) {
      var i = pivot.headers.indexOf(name);
      return i === -1 ? 0 : r[i];
    });
  });

  // Manual two-step Grand Total (matches `pivot["Grand Total"]=sum(axis=1)`
  // then `pivot.loc["Grand Total"]=sum()` — column first, then row over the
  // now-wider matrix). The sheet keeps all three columns: Completed, Pending
  // and Grand Total; only the CHART leaves the total out.
  var headers = W1_COLUMNS.concat(['Grand Total']);
  var dataRows = baseRows.map(function (r) {
    var sum = r.reduce(function (a, b) { return a + b; }, 0);
    return r.concat([sum]);
  });
  var totalsRow = headers.map(function (_, c) {
    return dataRows.reduce(function (acc, r) { return acc + r[c]; }, 0);
  });
  var indexVals = pivot.indexVals.concat(['Grand Total']);
  dataRows.push(totalsRow);

  var headerRow = [orgCol].concat(headers);
  var gridRows = indexVals.map(function (iv, i) { return [iv].concat(dataRows[i]); });
  var grid = [headerRow].concat(gridRows);

  var files = [{ name: 'output.xlsx', sheets: [{ name: 'Summary', grid: grid }] }];

  // Embedded chart reproduces the original openpyxl native stacked BarChart
  // (white background, "Completed in 2026"=red / "Pending"=orange, axis
  // titles) — built from the pivot BEFORE the Grand Total column/row were
  // appended, matching the original's Reference(min_col=2, max_col=3,
  // max_row=ws.max_row-1) which excludes both. Deviation from the Python
  // chart: white in-bar value labels (zeros hidden), same as d2/s3d.
  var D3_COLORS = { 'Completed in 2026': '#FF0000', 'Pending': '#FFC000' };
  var chartConfigs = {};
  if (pivot.indexVals.length) {
    var d3Traces = W1_COLUMNS.map(function (colName, idx) {
      var vals = baseRows.map(function (r) { return r[idx]; });
      return {
        x: pivot.indexVals, y: vals,
        type: 'bar', name: colName, marker: { color: D3_COLORS[colName] || '#4472C4' },
        text: vals.map(function (v) { return v > 0 ? String(v) : ''; }),
        textposition: 'inside', insidetextanchor: 'middle', textfont: { color: '#ffffff', size: 14 },
        // Keep every value upright and full-size: without these Plotly
        // rotates/shrinks labels to fit the smaller bars.
        textangle: 0, constraintext: 'none',
      };
    });
    var d3Layout = {
      barmode: 'stack',
      paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
      title: { text: 'Assessments Completed vs Open', font: { color: '#000000', size: 13 } },
      xaxis: { title: 'Region', tickfont: { color: '#000000' } },
      yaxis: { title: 'Count', gridcolor: '#D9D9D9' },
      legend: { font: { color: '#000000' } },
      margin: { t: 50, r: 20, b: 60, l: 50 },
    };
    // Live Plotly preview of the same styled chart (no baked image).
    chartConfigs['Summary'] = { traces: d3Traces, layout: d3Layout };
  }

  var workbook = new ExcelJS.Workbook();
  var ws = workbook.addWorksheet('Summary');
  grid.forEach(function (r) { ws.addRow(r); });
  ws.columns.forEach(function (c) { c.width = 16; });

  // Native, editable stacked chart (data-linked to a hidden helper block)
  // replaces the baked PNG. Dynamic series come from the pivot headers.
  var placements = [];
  if (window.NativeChartInject && window.fflate && pivot.indexVals.length) {
    // Summary grid: header row 1 (A=Region, B..=pivot headers, last col=Grand
    // Total); data rows 2..1+N (the appended Grand Total row is excluded).
    // Region labels are editable in column A; legend names via each header cell.
    var R = window.NativeChartInject.ref, N = pivot.indexVals.length, last = 1 + N;
    var d3Series = pivot.headers.map(function (colName, idx) {
      var col = 2 + idx; // A=1 is Region; headers start at B
      return {
        name: { ref: R('Summary', col, 1, 1), lit: colName },
        values: { ref: R('Summary', col, 2, last), cache: pivot.rows.map(function (r) { return r[idx]; }) },
        color: (D3_COLORS[colName] || '#4472C4').replace('#', ''),
      };
    });
    placements.push({
      sheetName: 'Summary', anchor: { fromCol: 7, fromRow: 4, toCol: 16, toRow: 21 }, // ~"H5"
      def: {
        grouping: 'stacked', legend: true, title: 'Assessments Completed vs Open', axisColor: '000000',
        dataLabels: { position: 'ctr', color: 'FFFFFF' },
        categories: { ref: R('Summary', 1, 2, last), cache: pivot.indexVals },
        series: d3Series,
      },
    });
  }

  var buf = await workbook.xlsx.writeBuffer();
  if (placements.length) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('d3 native chart inject failed:', e); }
  }

  return {
    ok: true,
    files: [{ name: 'output.xlsx', bytes: buf, sheets: [{ name: 'Summary', grid: grid }] }],
    chartConfigs: chartConfigs,
  };
};
