// Replaces run_all(sid, code) + runWithPyodide() (formerly index.html, running
// inside Pyodide). Each Reports.<id>(wb) function does its own data
// transform, renders its embedded chart image(s) via renderStyledPng
// (reproducing that report's original matplotlib/openpyxl/xlsxwriter chart
// look), writes its own styled output workbook(s) with ExcelJS, and returns
// { ok, files: [{ name, bytes: ArrayBuffer, sheets: [{name, grid}] }] }.
// buildPayload then derives the JSON preview (table + a SEPARATE, lighter
// preview chart via ReportEngine.processSheets/selectCharts) from those same
// in-memory grids — no write-then-reread round trip needed.
window.ReportBridge = (function () {
  'use strict';
  var E = window.ReportEngine;

  function arrayBufferToBase64(buf) {
    var bytes = new Uint8Array(buf);
    var chunk = 0x8000;
    var parts = [];
    for (var i = 0; i < bytes.length; i += chunk) {
      parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
    }
    return btoa(parts.join(''));
  }

  function dataUrlToBase64(dataUrl) {
    return dataUrl.split(',')[1];
  }

  // Renders a fully custom Plotly traces/layout pair to a PNG data URL via a
  // hidden Plotly instance, for embedding into the downloaded workbook. Each
  // report module builds traces/layout that reproduce that report's original
  // matplotlib/openpyxl/xlsxwriter chart look (colors, dark/light background,
  // data labels, legend position, etc.) — the downloaded file's chart should
  // look like the original script's output, not the light in-page preview
  // theme (buildBar/buildHBar in index.html, used only for the JSON preview).
  function renderStyledPng(traces, layout, width, height) {
    width = width || 720; height = height || 420;
    var host = document.createElement('div');
    host.style.position = 'fixed';
    host.style.left = '-10000px';
    host.style.top = '0';
    host.style.width = width + 'px';
    host.style.height = height + 'px';
    document.body.appendChild(host);
    // Plotly's own default family is Open Sans/Verdana. Pin the app font so the
    // PNGs embedded in the workbook are in the same single family as the UI.
    var fullLayout = Object.assign({ width: width, height: height }, layout);
    fullLayout.font = Object.assign({ family: "Arial, 'Helvetica Neue', Helvetica, sans-serif" }, fullLayout.font);
    return Plotly.newPlot(host, traces, fullLayout, { staticPlot: true })
      .then(function () { return Plotly.toImage(host, { format: 'png', width: width, height: height }); })
      .then(function (dataUrl) {
        Plotly.purge(host);
        document.body.removeChild(host);
        return dataUrl;
      })
      .catch(function (err) {
        Plotly.purge(host);
        document.body.removeChild(host);
        throw err;
      });
  }

  // result: { ok, stdout, stderr, files: [{ name, bytes: ArrayBuffer, sheets: [{name, grid}] }] }
  function buildPayload(sid, result) {
    var processed = E.processSheets(sid, result.files);
    var charts = E.selectCharts(sid, processed, result.files);
    var chartImages = result.chartImages || {};

    var sheetsOut = processed.map(function (p, i) {
      return {
        file: p.file,
        name: p.name,
        rows: p.sheet.rows.length,
        cols: p.sheet.headers.length,
        headers: p.sheet.headers,
        preview: E.previewRows(p.sheet, 50),
        chart: charts[i],
        chartImageUrl: chartImages[p.name] || null,
      };
    });

    var copyTsv = processed.length ? E.toTsv(processed[0].sheet, 5) : '';

    return {
      ok: !!result.ok,
      returncode: result.ok ? 0 : 1,
      stdout: result.stdout || '',
      stderr: result.stderr || '',
      files: result.files.map(function (f) {
        return { name: f.name, data_b64: arrayBufferToBase64(f.bytes), mime: E.mime(f.name) };
      }),
      sheets: sheetsOut,
      copy_tsv: copyTsv,
    };
  }

  async function runReport(entry, file) {
    var buf = await file.arrayBuffer();
    // No cellDates: true — date cells come back as raw Excel serials, and
    // each report reads them via ReportEngine.excelDateInfo/excelYear
    // instead of JS Date getters. See excelDateInfo's comment for why.
    var wb = XLSX.read(buf, { type: 'array' });
    // s4 records the source workbook's name in its History table, the way the
    // script logs os.path.basename(input_file).
    wb.__fileName = file && file.name;
    var fn = window.Reports[entry.id];
    if (!fn) throw new Error('No JS report implementation for ' + entry.id);
    var result = await fn(wb);
    return buildPayload(entry.id, result);
  }

  return {
    runReport: runReport,
    buildPayload: buildPayload,
    renderStyledPng: renderStyledPng,
    dataUrlToBase64: dataUrlToBase64,
    arrayBufferToBase64: arrayBufferToBase64,
  };
})();
