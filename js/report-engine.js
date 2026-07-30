// Shared utilities replacing the old Python harness (formerly PY_SETUP,
// index.html) that ran inside Pyodide. Ported 1:1 from that harness's exact
// semantics (sparse trimming thresholds, JSON-safety coercion, the generic
// chartable-column heuristic, TSV clipboard formatting) so every report
// produces the same preview/table/chart output it did under Python.
window.Reports = window.Reports || {};

window.ReportEngine = (function () {
  'use strict';

  var NAN_STRINGS = { 'nan': 1, 'none': 1, 'nat': 1, '': 1 };

  // Reads a named sheet from a SheetJS workbook (XLSX.read result) into rows
  // keyed by literal header text (mirrors pd.read_excel's default header=0),
  // plus the raw header row for callers that need to clean/rename columns.
  function readSheet(wb, sheetName) {
    var ws = wb.Sheets[sheetName];
    if (!ws) throw new Error('Sheet not found: ' + sheetName);
    var headers = (XLSX.utils.sheet_to_json(ws, { header: 1, raw: true })[0] || []);
    var rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    return { headers: headers, rows: rows };
  }

  function isBlank(v) {
    if (v === null || v === undefined) return true;
    if (typeof v === 'number') return Number.isNaN(v);
    return NAN_STRINGS.hasOwnProperty(String(v).trim().toLowerCase());
  }

  // Extracts calendar components from an Excel date cell WITHOUT going
  // through a JS Date object's local-timezone getters. Deliberately avoids
  // XLSX's cellDates:true (which builds dates via `new Date(1899,11,30,...)`
  // - a local-constructor call whose UTC offset for 1899 can differ from a
  // modern date's offset in timezones that changed their official UTC
  // offset historically, e.g. India shifted from UTC+5:21 to UTC+5:30 in
  // 1906. That mismatch silently mis-derives the year for some rows,
  // depending entirely on the machine's timezone). Reading the raw numeric
  // serial and decoding it with XLSX.SSF.parse_date_code sidesteps the
  // whole class of bug - this app must behave identically regardless of
  // what timezone the client machine is set to.
  function excelDateInfo(v) {
    if (isBlank(v)) return null;
    if (typeof v === 'number') {
      var d = XLSX.SSF.parse_date_code(v);
      if (!d) return null;
      return { year: d.y, month: d.m, day: d.d, hour: d.H || 0, minute: d.M || 0, second: Math.floor(d.S || 0) };
    }
    if (v instanceof Date) {
      return { year: v.getUTCFullYear(), month: v.getUTCMonth() + 1, day: v.getUTCDate(), hour: v.getUTCHours(), minute: v.getUTCMinutes(), second: v.getUTCSeconds() };
    }
    // Text dates. pandas' to_datetime accepts far more shapes than
    // `new Date()` does, and a report that silently drops every row it cannot
    // parse reports a confident wrong number rather than an error - a
    // "Date created" column exported as 14/12/2026 made every 2026 filter in
    // this app match nothing. Each shape below is decomposed with a regex and
    // never round-tripped through a local-time Date, for the same reason the
    // numeric branch above decodes the serial by hand.
    var s = String(v).trim();
    if (!s) return null;
    var m;

    // Year first: 2026-12-14, 2026/12/14, 2026.12.14, 2026-12-14T09:30:00.
    m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (m) return ymdInfo(+m[1], +m[2], +m[3], s);

    // Day/month numeric: 14/12/2026, 12/14/2026, 14-12-2026, 14.12.2026.
    // pandas defaults to dayfirst=False, so a leading value that COULD be a
    // month is read as one; a leading value above 12 can only be the day.
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
    if (m) {
      var a = +m[1], b = +m[2], y = +m[3];
      return a <= 12 ? ymdInfo(y, a, b, s) : ymdInfo(y, b, a, s);
    }

    // Named month: 14-Dec-2026, 14 Dec 2026, 14 December 2026.
    m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s,]+(\d{4})/);
    if (m) {
      var mo = MONTH_NUM[m[2].slice(0, 3).toLowerCase()];
      if (mo) return ymdInfo(+m[3], mo, +m[1], s);
    }
    // Month first: Dec 14, 2026 / December 14 2026.
    m = s.match(/^([A-Za-z]{3,})[-\s](\d{1,2})[-\s,]+(\d{4})/);
    if (m) {
      var mo2 = MONTH_NUM[m[1].slice(0, 3).toLowerCase()];
      if (mo2) return ymdInfo(+m[3], mo2, +m[2], s);
    }

    // Anything exotic: let the engine try. Every string still reaching here
    // was parsed as LOCAL time - only date-only ISO strings are treated as
    // UTC, and those are handled above - so read the local components. The
    // previous code read the UTC ones, which shifted every text date back a
    // day anywhere east of Greenwich (12/14/2026 came out as the 13th, and
    // 01/01/2026 came out in 2025 and was dropped by a year filter).
    var parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return { year: parsed.getFullYear(), month: parsed.getMonth() + 1, day: parsed.getDate(), hour: parsed.getHours(), minute: parsed.getMinutes(), second: parsed.getSeconds() };
    }
    return null;
  }

  var MONTH_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

  // Builds the calendar-component object shared by every text-date branch,
  // rejecting impossible month/day combinations (25/13/2026) rather than
  // letting them roll over into the next month the way Date would.
  function ymdInfo(year, month, day, src) {
    if (!(month >= 1 && month <= 12) || !(day >= 1 && day <= 31)) return null;
    var t = String(src).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    return {
      year: year, month: month, day: day,
      hour: t ? +t[1] : 0, minute: t ? +t[2] : 0, second: t && t[3] ? +t[3] : 0,
    };
  }

  function excelYear(v) {
    var info = excelDateInfo(v);
    return info ? info.year : null;
  }

  // A column is "numeric" (pandas numeric dtype) if every non-blank cell in
  // it is a JS number and at least one such value exists.
  function isNumericColumn(rows, colIdx) {
    var sawNumber = false;
    for (var i = 0; i < rows.length; i++) {
      var v = rows[i][colIdx];
      if (isBlank(v)) continue;
      if (typeof v !== 'number') return false;
      sawNumber = true;
    }
    return sawNumber;
  }

  // grid: 2D array, row 0 = headers, remaining rows = data - mirrors
  // pd.read_excel(path, header=0) turning row 0 into df.columns.
  // A real worksheet is rectangular - its column count is the widest row
  // written anywhere in the sheet, not however many cells happen to be
  // non-blank in row 0. Hand-built grids (setCell-style, sparse title/rule
  // rows above a real table) are naturally ragged arrays; pad every row to
  // the grid's overall max width before treating row 0 as the header, or
  // trimSparseCols would silently truncate to row 0's length.
  function padGrid(grid) {
    var width = grid.reduce(function (w, row) { return Math.max(w, row ? row.length : 0); }, 0);
    return grid.map(function (row) {
      row = row || [];
      var padded = row.slice();
      while (padded.length < width) padded.push(null);
      return padded;
    });
  }

  function sheetFromGrid(grid) {
    var padded = padGrid(grid);
    var headers = (padded[0] || []).map(function (h) {
      return (h === null || h === undefined || String(h).trim() === '') ? '' : String(h);
    });
    return { headers: headers, rows: padded.slice(1) };
  }

  // --- sparse trimming (mirrors _trim_sparse_rows / _trim_sparse_cols) ---
  function trimSparseRows(sheet) {
    if (!sheet.rows.length) return sheet;
    var threshold = Math.max(2, Math.floor(sheet.headers.length / 2));
    var kept = sheet.rows.filter(function (row) {
      var filled = 0;
      for (var i = 0; i < row.length; i++) if (!isBlank(row[i])) filled++;
      return filled >= threshold;
    });
    return { headers: sheet.headers, rows: kept };
  }

  function trimSparseCols(sheet) {
    if (!sheet.rows.length) return sheet;
    var threshold = Math.max(2, Math.floor(sheet.rows.length / 2));
    var keepIdx = [];
    for (var c = 0; c < sheet.headers.length; c++) {
      var filled = 0;
      for (var r = 0; r < sheet.rows.length; r++) if (!isBlank(sheet.rows[r][c])) filled++;
      if (filled >= threshold) keepIdx.push(c);
    }
    return {
      headers: keepIdx.map(function (i) { return sheet.headers[i]; }),
      rows: sheet.rows.map(function (row) { return keepIdx.map(function (i) { return row[i]; }); }),
    };
  }

  function trimSparse(sheet) {
    return trimSparseCols(trimSparseRows(sheet));
  }

  // --- JSON-safety coercion (mirrors _safe_val) ---
  function safeVal(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') {
      if (Number.isNaN(v) || !Number.isFinite(v)) return '';
      if (Number.isInteger(v) && Math.abs(v) < 1e15) return v;
      return Math.round(v * 1e6) / 1e6;
    }
    var s = String(v);
    return NAN_STRINGS.hasOwnProperty(s.trim().toLowerCase()) ? '' : s;
  }

  function previewRows(sheet, limit) {
    var rows = limit ? sheet.rows.slice(0, limit) : sheet.rows;
    return rows.map(function (row) { return row.map(safeVal); });
  }

  // --- TSV clipboard formatting (mirrors _clean_for_tsv) ---
  function toTsv(sheet, limit) {
    var headers = sheet.headers;
    var rows = limit ? sheet.rows.slice(0, limit) : sheet.rows;
    var numericCol = headers.map(function (_, i) { return isNumericColumn(sheet.rows, i); });
    var lines = [headers.join('\t')];
    rows.forEach(function (row) {
      var cells = row.map(function (v, i) {
        if (isBlank(v)) return '';
        if (numericCol[i]) {
          if (typeof v === 'number' && Number.isInteger(v)) return String(v);
          return String(v);
        }
        return String(v).replace(/[\t\r\n]+/g, ' ').trim();
      });
      lines.push(cells.join('\t'));
    });
    return lines.join('\n') + '\n';
  }

  // --- generic chart extraction (mirrors _chartable + _chart_json) ---
  // Exact match only - s4's "Total Risk" is a real series, not a totals line.
  function isTotalLabel(v) {
    var s = isBlank(v) ? '' : String(v).trim().toLowerCase();
    return s === 'grand total' || s === 'total';
  }
  function chartableFromSheet(sheet) {
    if (!sheet || !sheet.rows.length) return null;
    var headers = sheet.headers;
    var numericIdx = [], labelIdx = [];
    headers.forEach(function (h, i) {
      (isNumericColumn(sheet.rows, i) ? numericIdx : labelIdx).push(i);
    });
    if (!numericIdx.length || !labelIdx.length) return null;
    // A "Grand Total" COLUMN is the sum of the other series, so charting it adds
    // a redundant bar that dwarfs its own parts (d3's Summary). Drop it, same as
    // the totals ROW below - unless it is the only numeric column, leaving
    // nothing to plot. The table keeps the column; this only affects the chart.
    var nonTotalIdx = numericIdx.filter(function (i) { return !isTotalLabel(headers[i]); });
    if (nonTotalIdx.length) numericIdx = nonTotalIdx;
    var labelI = labelIdx[0];
    var rows = sheet.rows.filter(function (row) {
      return !isBlank(row[labelI]) && !isTotalLabel(row[labelI]);
    });
    if (!rows.length) return null;
    var labels = rows.map(function (row) { return String(row[labelI]); });
    var series = numericIdx.map(function (i, idx) {
      var name = String(headers[i]).replace(/:/g, '_').replace(/\./g, '_');
      if (!name) name = 'col_' + idx;
      return {
        name: name,
        values: rows.map(function (row) {
          var v = row[i];
          return (typeof v === 'number' && !Number.isNaN(v)) ? v : 0;
        }),
      };
    });
    return { type: 'bar', labels: labels, series: series };
  }

  // --- d2's KPI-block chart (mirrors _d2_chart_json) ---
  // Operates on the RAW grid (header=None equivalent) of d2's "Dashboard" sheet.
  function d2ChartFromGrid(grid) {
    var kpiRowIdx = -1;
    for (var i = 0; i < grid.length; i++) {
      if ((grid[i] || []).some(function (v) { return v === 'Metric'; })) { kpiRowIdx = i; break; }
    }
    if (kpiRowIdx === -1) return null;
    var headerRow = grid[kpiRowIdx];
    var metricCol = headerRow.indexOf('Metric');
    var valueCol = headerRow.indexOf('Value');
    if (metricCol === -1 || valueCol === -1) return null;
    var labels = [], values = [];
    // Read only the contiguous KPI table, stopping at the blank row that ends
    // it. This used to scan to the bottom of the sheet, so any later block with
    // a label and a number became an extra bar - a "Rows matching 36" row below
    // the table charted as 3600%. The KPI chart is a fixed four-metric panel;
    // it has no business reading rows the table does not own.
    for (var r = kpiRowIdx + 1; r < grid.length; r++) {
      var row = grid[r] || [];
      var metric = row[metricCol];
      if (isBlank(metric) || String(metric).trim() === '') break;
      var val = row[valueCol];
      // Genuine numeric cells only: parseFloat('2026 x36') would yield 2026.
      // Exception: strict "NN%" strings - d2's KPI Value column renders its
      // fractions as percent text ("60%"), which must keep charting as 0.60.
      var num = typeof val === 'number' ? val
        : (typeof val === 'string' && /^-?\d+(\.\d+)?%$/.test(val.trim())) ? parseFloat(val) / 100 : NaN;
      if (isBlank(val) || Number.isNaN(num)) continue;
      if (/formula/i.test(String(metric))) continue;
      labels.push(String(metric));
      values.push(num);
    }
    if (!labels.length) return null;
    return { type: 'hbar', labels: labels, series: [{ name: 'Value', values: values }] };
  }

  // --- s2a's zone Open/Closed chart (mirrors _s2a_chart_json) ---
  // Operates on the RAW grid of s2a's "Auto Open Closed" sheet.
  function s2aChartFromGrid(grid) {
    var hdrIdx = -1;
    for (var i = 0; i < grid.length; i++) {
      var vals = (grid[i] || []).map(function (v) { return String(v).trim().toLowerCase(); });
      if (vals.indexOf('zone') !== -1 && vals.indexOf('closed') !== -1 && vals.indexOf('open') !== -1) {
        hdrIdx = i; break;
      }
    }
    if (hdrIdx === -1) return null;
    var headerRow = grid[hdrIdx].map(function (v) { return String(v).trim(); });
    var zoneCol = headerRow.indexOf('Zone');
    var closedCol = headerRow.indexOf('Closed');
    var openCol = headerRow.indexOf('Open');
    if (zoneCol === -1 || closedCol === -1 || openCol === -1) return null;
    var labels = [], closedVals = [], openVals = [];
    for (var r = hdrIdx + 1; r < grid.length; r++) {
      var row = grid[r] || [];
      var zone = row[zoneCol];
      if (isBlank(zone)) continue;
      var zs = String(zone).trim().toLowerCase();
      if (zs === 'total' || zs === 'grand total' || zs === '') continue;
      var c = row[closedCol], o = row[openCol];
      var cn = typeof c === 'number' ? c : parseFloat(c); if (Number.isNaN(cn)) cn = 0;
      var on = typeof o === 'number' ? o : parseFloat(o); if (Number.isNaN(on)) on = 0;
      labels.push(String(zone));
      closedVals.push(cn);
      openVals.push(on);
    }
    if (!labels.length) return null;
    return { type: 'bar', labels: labels, series: [{ name: 'Closed', values: closedVals }, { name: 'Open', values: openVals }] };
  }

  // --- generic groupBy (Map-based reducer) ---
  function groupBy(rows, keyFn) {
    var map = new Map();
    rows.forEach(function (row) {
      var k = keyFn(row);
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(row);
    });
    return map;
  }

  function mapZone(org, table, fallback) {
    if (table.hasOwnProperty(org)) return table[org];
    return fallback === undefined ? org : fallback;
  }

  // Growth/BEES zone sanitizer, shared by every report that groups on an
  // organization or zone column: GRO, Growth, GROWTH, BEES, BEES | FINTECH -
  // in any casing, spacing, punctuation, or with invisible characters
  // (zero-width space, NBSP) - all report as the single zone 'GRO'.
  // Compared on letters only so visually-identical export variants can
  // never split the bucket. Every other value passes through unchanged.
  function sanitizeZone(v) {
    var s = (v === null || v === undefined) ? '' : String(v);
    var letters = s.replace(/[^a-z]/gi, '').toLowerCase();
    return (letters === 'gro' || letters === 'growth' || letters === 'bees' || letters === 'beesfintech') ? 'GRO' : s;
  }

  // Count pivot with pandas-matching semantics: columns sorted lexicographically
  // (code-point order, matching Python's default string sort), missing
  // index/column combinations filled with 0. If margins: a margin COLUMN
  // (row sums) is appended first, then a margin ROW (column sums over the
  // now-wider matrix, so the corner cell = total count) is appended second -
  // matching pandas pivot_table(margins=True) / the manual two-step
  // Grand-Total-column-then-row pattern used by the non-margins scripts.
  //
  // opts.valueFn mirrors pandas `values=<col>, aggfunc="count"`, which counts
  // NON-NULL value cells rather than rows: a row whose value cell is blank is
  // skipped, but it still contributes its index/column LABEL, because pandas
  // keeps that group in the result (at 0). Omit valueFn to count rows.
  function pivotCount(rows, indexFn, columnsFn, opts) {
    opts = opts || {};
    var marginsName = opts.marginsName || 'Grand Total';
    var valueFn = opts.valueFn;
    var indexVals = [], indexSeen = new Set();
    var colVals = [], colSeen = new Set();
    rows.forEach(function (row) {
      var iv = indexFn(row), cv = columnsFn(row);
      if (!indexSeen.has(iv)) { indexSeen.add(iv); indexVals.push(iv); }
      if (!colSeen.has(cv)) { colSeen.add(cv); colVals.push(cv); }
    });
    colVals.sort();
    indexVals.sort();
    var counts = {};
    indexVals.forEach(function (iv) {
      counts[iv] = {};
      colVals.forEach(function (cv) { counts[iv][cv] = 0; });
    });
    rows.forEach(function (row) {
      if (valueFn && isBlank(valueFn(row))) return;
      var iv = indexFn(row), cv = columnsFn(row);
      counts[iv][cv] += 1;
    });
    var headers = colVals.slice();
    var dataRows = indexVals.map(function (iv) {
      return colVals.map(function (cv) { return counts[iv][cv]; });
    });
    if (opts.margins) {
      dataRows = dataRows.map(function (row) {
        var sum = row.reduce(function (a, b) { return a + b; }, 0);
        return row.concat([sum]);
      });
      headers = headers.concat([marginsName]);
      var totalsRow = headers.map(function (_, c) {
        return dataRows.reduce(function (acc, row) { return acc + row[c]; }, 0);
      });
      indexVals = indexVals.concat([marginsName]);
      dataRows = dataRows.concat([totalsRow]);
    }
    return { indexVals: indexVals, headers: headers, rows: dataRows };
  }

  function mime(name) {
    var n = name.toLowerCase();
    if (n.endsWith('.xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (n.endsWith('.png')) return 'image/png';
    return 'application/octet-stream';
  }

  // --- per-sid sheet skip rules (mirrors run_all's processed_sheets loop) ---
  function skipSheet(sid, filename, sheetName) {
    var fn = filename.toLowerCase();
    if (sid === 's4' && (fn === 'history.xlsx' || sheetName === 'History Used')) return true;
    if (sid === 's3d' && sheetName === 'Open Overdue Pivot') return true;
    if (sid === 's1c' && sheetName === 'Pivot') return true;
    if (sid === 's2a' && sheetName === 'Auto Pivot Summary') return true;
    return false;
  }

  // files: [{ name, sheets: [{ name, grid }] }] (pure data, pre-write).
  // Returns the trimmed, skip-filtered sheet list in file/sheet order -
  // mirrors run_all's `processed_sheets`. Shared by the preview payload
  // builder AND each report module's own chart-embedding step, so both
  // always agree on which sheets exist post-trim.
  function processSheets(sid, files) {
    var out = [];
    files.forEach(function (file) {
      file.sheets.forEach(function (sheet) {
        if (skipSheet(sid, file.name, sheet.name)) return;
        out.push({ file: file.name, name: sheet.name, sheet: trimSparse(sheetFromGrid(sheet.grid)) });
      });
    });
    return out;
  }

  function findRawGrid(files, sheetName) {
    for (var f = 0; f < files.length; f++) {
      var sheets = files[f].sheets;
      for (var s = 0; s < sheets.length; s++) {
        if (sheets[s].name === sheetName) return sheets[s].grid;
      }
    }
    return null;
  }

  function firstXlsxFile(files) {
    for (var i = 0; i < files.length; i++) {
      if (files[i].name.toLowerCase().endsWith('.xlsx')) return files[i];
    }
    return null;
  }

  // Mirrors run_all's per-sid chart-selection rules exactly (d2/s4 = only the
  // first processed sheet; s2a = only its "Auto Open Closed" sheet; everyone
  // else = every processed sheet tries its own chartable extraction).
  // Returns an array parallel to `processed` (from processSheets): chart
  // JSON or null per entry. Used both for the JSON preview payload and to
  // decide which sheets a report module should embed a chart image into.
  function selectCharts(sid, processed, files) {
    var special = null;
    if (sid === 'd2') {
      var dashGrid = findRawGrid(files, 'Dashboard');
      special = dashGrid ? d2ChartFromGrid(dashGrid) : null;
    } else if (sid === 's4') {
      // Chart the Dashboard's "Chart Label/Chart Value" helper block - the
      // same Baseline + display periods + Target series the embedded native
      // chart plots. Falls back to the rolled-up Display History sheet, then
      // to the full monthly history file.
      var s4Dash = findRawGrid(files, 'Dashboard');
      if (s4Dash) {
        var s4Labels = [], s4Values = [], s4Started = false;
        for (var dr = 0; dr < s4Dash.length; dr++) {
          var drow = s4Dash[dr] || [];
          if (!s4Started) { if (drow[0] === 'Chart Label') s4Started = true; continue; }
          if (isBlank(drow[0])) break;
          var dv = typeof drow[1] === 'number' ? drow[1] : parseFloat(drow[1]);
          s4Labels.push(String(drow[0]));
          s4Values.push(Number.isNaN(dv) ? 0 : dv);
        }
        if (s4Labels.length) special = { type: 'bar', labels: s4Labels, series: [{ name: 'Progress', values: s4Values }] };
      }
      if (!special) {
        var dispGrid = findRawGrid(files, 'Display History');
        if (dispGrid) special = chartableFromSheet(trimSparse(sheetFromGrid(dispGrid)));
      }
      if (!special) {
        var histFile = files.filter(function (f) { return f.name.toLowerCase() === 'history.xlsx'; })[0];
        if (histFile) {
          for (var i = 0; i < histFile.sheets.length; i++) {
            var cdf = chartableFromSheet(trimSparse(sheetFromGrid(histFile.sheets[i].grid)));
            if (cdf) { special = cdf; break; }
          }
        }
      }
    } else if (sid === 's2a') {
      var xf = firstXlsxFile(files);
      var aocGrid = xf ? (xf.sheets.filter(function (s) { return s.name === 'Auto Open Closed'; })[0] || {}).grid : null;
      special = aocGrid ? s2aChartFromGrid(aocGrid) : null;
    }

    var chartsDone = 0;
    return processed.map(function (p) {
      var showChart;
      if (sid === 'd2' || sid === 's4') showChart = chartsDone === 0;
      else if (sid === 's2a') showChart = p.name === 'Auto Open Closed';
      else showChart = true;
      if (!showChart) return null;
      if (sid === 'd2' || sid === 's4' || sid === 's2a') {
        if (special) { chartsDone++; return special; }
        return null;
      }
      var cdf = chartableFromSheet(p.sheet);
      // d1's native Excel chart is stacked; the preview mirrors it (green
      // Supplier-Added bars stacked over the gold Tier-1 bars).
      if (cdf && sid === 'd1') cdf.stacked = true;
      if (cdf) { chartsDone++; return cdf; }
      return null;
    });
  }

  return {
    readSheet: readSheet,
    excelDateInfo: excelDateInfo,
    excelYear: excelYear,
    isBlank: isBlank,
    isNumericColumn: isNumericColumn,
    padGrid: padGrid,
    sheetFromGrid: sheetFromGrid,
    trimSparseRows: trimSparseRows,
    trimSparseCols: trimSparseCols,
    trimSparse: trimSparse,
    safeVal: safeVal,
    previewRows: previewRows,
    toTsv: toTsv,
    chartableFromSheet: chartableFromSheet,
    d2ChartFromGrid: d2ChartFromGrid,
    s2aChartFromGrid: s2aChartFromGrid,
    groupBy: groupBy,
    mapZone: mapZone,
    sanitizeZone: sanitizeZone,
    pivotCount: pivotCount,
    mime: mime,
    processSheets: processSheets,
    selectCharts: selectCharts,
  };
})();
