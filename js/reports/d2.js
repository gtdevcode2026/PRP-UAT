// Port of "daigram 2 automation/automation.py" (id: d2).
// Reads "OneTrust Assessment", filters Tags=="cyber" (case-insensitive
// exact match) & Date created year==2026, maps Stage to Closed/Open, maps
// Organization to a display zone (falling back to the organization name
// itself when unmapped - unlike s2a/s3d's zone_map, "Europe" stays
// "Europe" here, not "EUR"), pivots Org Display x Final Stage (count, no
// margins), sorts to a fixed zone order, appends a Grand Total row, and
// computes a KPI panel that follows TODAY's calendar: every finished
// period is static (Q1 fixed in code; later quarters and past months
// lock for good on the first run after they end), only the current
// month's bar stays live, Baseline/Target stay static bookends.
// Writes "output file D2.xlsx" as a single hand-built "Dashboard" sheet
// (filter area, then the pivot table, then the KPI table below it) - the
// exact row offsets matter because the existing preview pipeline re-treats
// row 0 as headers and sparsity-trims the rest, same as s2a.
window.Reports.d2 = async function d2(wb, opts) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers.map(function (h) { return String(h).trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    sheet.headers.forEach(function (h, i) { out[rawHeaders[i]] = row[h]; });
    return out;
  });
  ['ID', 'Organization', 'Stage', 'Date created', 'Tags'].forEach(function (c) {
    if (rawHeaders.indexOf(c) === -1) throw new Error('Missing required columns: ' + c);
  });

  rows.forEach(function (r) {
    ['Organization', 'Stage', 'Tags'].forEach(function (c) {
      r[c] = E.isBlank(r[c]) ? '' : String(r[c]).trim();
    });
  });

  // Dynamic Finance filter - names typed in the Generate card (shown for
  // Diagram 2 only). When names are given, a 2026 row whose Respondents cell
  // mentions ANY of them (case-insensitive substring, so 'A; B' cells match)
  // is Finance and excluded; every other 2026 row counts as Cyber, whatever
  // its Tags say. With no names - box cleared - or no Respondents column to
  // match against, the Tags column written at workbook-creation time decides,
  // exactly as before.
  var financeNames = ((opts && opts.financeRespondents) || []).map(function (n) {
    return String(n).trim().toLowerCase();
  }).filter(Boolean);
  var loweredHeaders = rawHeaders.map(function (h) { return h.toLowerCase(); });
  var respCol = null;
  ['Respondents', 'Respondent', 'Respondent Name', 'Assignee'].forEach(function (h) {
    var i = loweredHeaders.indexOf(h.toLowerCase());
    if (respCol === null && i !== -1) respCol = rawHeaders[i];
  });
  var dynamicFilter = financeNames.length > 0 && respCol !== null;

  var excludedFinance = 0;
  var filtered = rows.filter(function (r) {
    if (E.excelYear(r['Date created']) !== 2026) return false;
    if (!dynamicFilter) return /^cyber$/i.test(r.Tags);
    var resp = E.isBlank(r[respCol]) ? '' : String(r[respCol]).toLowerCase();
    if (financeNames.some(function (n) { return resp.indexOf(n) !== -1; })) {
      excludedFinance++;
      return false;
    }
    return true;
  });

  var CLOSED_STAGES = { 'Completed': 1, 'Under review': 1 };
  filtered.forEach(function (r) {
    r['Final Stage'] = CLOSED_STAGES.hasOwnProperty(r.Stage) ? 'Closed' : 'Open';
  });

  // BEES / BEES | FINTECH / Growth variants are sanitized to 'GRO' before
  // the map runs (E.sanitizeZone), so the map itself no longer needs keys
  // for them; 'GRO' falls through mapZone unchanged.
  var ORG_MAP = {
    'Africa': 'AFR', 'APAC': 'APAC',
    'Europe': 'Europe', 'GHQ': 'GHQ', 'South America Zone': 'SAZ',
    'North America Zone': 'NAZ', 'Middle America Zone': 'MAZ',
  };
  filtered.forEach(function (r) { r['Org Display'] = E.mapZone(E.sanitizeZone(r.Organization), ORG_MAP); });

  var pivotRaw = E.pivotCount(filtered, function (r) { return r['Org Display']; }, function (r) { return r['Final Stage']; }, { margins: false });
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

  var closedTotal = filtered.filter(function (r) { return r['Final Stage'] === 'Closed'; }).length;
  var recordTotal = filtered.length;
  // Script log line: which Finance filter actually ran, so a changed number
  // is never a mystery.
  var filterLog = dynamicFilter
    ? 'Finance filter: dynamic - ' + excludedFinance + ' row(s) mentioning [' + financeNames.join(', ') +
      '] in "' + respCol + '" excluded; ' + recordTotal + ' remaining 2026 rows counted as Cyber.'
    : financeNames.length
      ? 'Finance filter: names were given but no Respondents column was found - fell back to Tags == "cyber" (' + recordTotal + ' rows).'
      : 'Finance filter: Tags column (Tags == "cyber") - ' + recordTotal + ' rows.';
  // KPI roll-up driven by TODAY's date, not by the data. Every month's rate
  // is that month's Closed / Total (Date created). Quarters that have already
  // ended show as ONE bar = the average of their months' rates (Q1 = avg of
  // Jan/Feb/Mar); the quarter we are in shows each elapsed month as its own
  // bar. So Feb renders Jan+Feb; May renders Q1, Apr, May; July renders
  // Q1, Q2, Jul. Months/quarters with no rows are skipped rather than drawn
  // as misleading 0% bars. Baseline and Target stay as static bookends.
  var MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var mStats = {}; // month 1-12 -> { closed, total }
  filtered.forEach(function (r) {
    var d = E.excelDateInfo(r['Date created']);
    if (!d) return;
    var s = mStats[d.month] || (mStats[d.month] = { closed: 0, total: 0 });
    s.total++;
    if (r['Final Stage'] === 'Closed') s.closed++;
  });
  var nowMonth = new Date().getMonth() + 1;
  var nowQuarter = Math.floor((nowMonth - 1) / 3) + 1;
  var KPI = [["Baseline '25", 0.60, 'Static']];
  var kpiNotes = [];
  // Every value locks as soon as its period is over. A month locks when the
  // NEXT month begins: its rate is computed once, saved to this browser's
  // localStorage, and reused untouched on every later run whatever workbook
  // or filter is loaded - only the CURRENT month stays live. Finished
  // quarters lock the same way as the average of their three locked months
  // (Q1 '26 fixed at 32% and Q2 '26 at 66% in code). Clearing site data - or e.g.
  // localStorage.removeItem('PRP_D2_KPI_Jul_26') in the console - re-arms
  // the one-time computation for that period.
  var STATIC_Q = { 1: 0.32, 2: 0.66 };
  function readLock(key) {
    try { var sv = localStorage.getItem(key); if (sv !== null && !isNaN(Number(sv))) return Number(sv); } catch (e) {}
    return null;
  }
  function writeLock(key, v) { try { localStorage.setItem(key, String(v)); } catch (e) {} }
  // Rate for a FINISHED month: its saved lock first; else computed once from
  // this run's data and saved for good; null when there is nothing to go on.
  function finishedMonthRate(m) {
    var key = 'PRP_D2_KPI_' + MONTH_ABBR[m - 1] + '_26';
    var v = readLock(key);
    if (v !== null) return v;
    if (!mStats[m]) return null;
    v = Math.round((mStats[m].closed / mStats[m].total) * 100) / 100;
    writeLock(key, v);
    return v;
  }
  for (var q = 1; q < nowQuarter; q++) { // finished quarters -> one static bar
    var lockKey = 'PRP_D2_KPI_Q' + q + '_26';
    var locked = STATIC_Q.hasOwnProperty(q) ? STATIC_Q[q] : readLock(lockKey);
    if (locked === null) {
      var parts = [], sum = 0, n = 0;
      [q * 3 - 2, q * 3 - 1, q * 3].forEach(function (m) {
        var v = finishedMonthRate(m);
        if (v === null) return;
        sum += v; n++;
        parts.push(MONTH_ABBR[m - 1] + ' ' + Math.round(v * 100) + '%');
      });
      if (!n) continue;
      locked = Math.round((sum / n) * 100) / 100;
      writeLock(lockKey, locked);
      KPI.push(['Q' + q + " '26", locked, 'Static']);
      kpiNotes.push('Q' + q + ' = avg(' + parts.join(', ') + ') - now locked');
      continue;
    }
    KPI.push(['Q' + q + " '26", locked, 'Static']);
    kpiNotes.push('Q' + q + ' locked at ' + Math.round(locked * 100) + '%');
  }
  for (var m = nowQuarter * 3 - 2; m <= nowMonth; m++) { // running quarter
    if (m < nowMonth) { // finished month -> locked static bar
      var mv = finishedMonthRate(m);
      if (mv === null) continue;
      KPI.push([MONTH_ABBR[m - 1] + " '26", mv, 'Static']);
      kpiNotes.push(MONTH_ABBR[m - 1] + ' locked at ' + Math.round(mv * 100) + '%');
    } else if (mStats[m]) { // current month -> live
      KPI.push([MONTH_ABBR[m - 1] + " '26", Math.round((mStats[m].closed / mStats[m].total) * 100) / 100, '']);
      kpiNotes.push(MONTH_ABBR[m - 1] + ' = ' + mStats[m].closed + '/' + mStats[m].total + ' (live)');
    }
  }
  KPI.push(["Target '26", 0.65, 'Static']);

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
    setCell(grid, rowNum, 2, Math.round(k[1] * 100) + '%');
    setCell(grid, rowNum, 3, k[2]);
  });
  var noteRow = kpiStartRow + KPI.length + 2;
  setCell(grid, noteRow, 1, 'KPI formula');
  setCell(grid, noteRow, 2, 'Closed / Total per month; finished quarters average their months: ' +
    (kpiNotes.length ? kpiNotes.join(', ') : 'no month data yet'));

  var files = [{ name: 'output file D2.xlsx', sheets: [{ name: 'Dashboard', grid: grid }] }];

  // Both charts reproduce the original matplotlib styles. The PNG renders
  // below are embedded into the .xlsx (matching today's 2-chart output);
  // the SAME traces/layouts also drive the live Plotly "Chart preview"
  // (chartConfigs in the return), so preview and file always match.
  var images = {};

  // Chart 1: org-level stacked bar - black bg, Open (cyan) bottom + Closed
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

  // Chart 2: KPI horizontal bar - white bg, single blue series, 0-100% axis,
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
  ws.columns.forEach(function (c) { c.width = 16; });

  if (images.org) {
    var id1 = workbook.addImage({ base64: images.org, extension: 'png' });
    ws.addImage(id1, { tl: { col: 5, row: 1 }, ext: { width: 480, height: 300 } }); // ~"F2"
  }
  if (images.kpi) {
    var id2 = workbook.addImage({ base64: images.kpi, extension: 'png' });
    ws.addImage(id2, { tl: { col: 5, row: 22 }, ext: { width: 480, height: 260 } }); // ~"F23"
  }

  var buf = await workbook.xlsx.writeBuffer();

  return {
    ok: true,
    stdout: filterLog,
    files: [{ name: 'output file D2.xlsx', bytes: buf, sheets: [{ name: 'Dashboard', grid: grid }] }],
    // Preview shows the KPI chart only - matching what the preview always
    // showed; the org chart is still embedded in the Excel Dashboard.
    chartConfigs: {
      'Dashboard': { traces: kpiTraces, layout: kpiLayout },
    },
  };
};
