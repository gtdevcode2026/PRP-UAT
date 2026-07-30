// Injects native OOXML charts into an ExcelJS-produced .xlsx (an OPC zip).
// Opens the zip with fflate, adds chart/drawing parts, patches
// [Content_Types].xml, the worksheet .rels, and the sheet's <drawing> element,
// then re-zips. Sheet display name -> worksheet part is resolved via
// workbook.xml + its rels (never guessed), so multi-sheet outputs map correctly.
// Dual-export; depends on global fflate + NativeChart.
(function (root) {
  'use strict';
  var NC = root.NativeChart || (typeof require !== 'undefined' ? require('./native-chart.js') : null);

  function dec(u8) { return new TextDecoder().decode(u8); }
  function enc(s) { return new TextEncoder().encode(s); }

  // Resolve sheet display name -> worksheet part path (xl/worksheets/sheetN.xml)
  function sheetPathMap(files) {
    var wbXml = dec(files['xl/workbook.xml']);
    var relsXml = dec(files['xl/_rels/workbook.xml.rels']);
    var relTarget = {};
    relsXml.replace(/<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/?>/g,
      function (m, id, target) { relTarget[id] = target; return m; });
    var map = {};
    wbXml.replace(/<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/?>/g,
      function (m, name, rid) {
        var t = relTarget[rid] || '';
        map[name] = 'xl/' + t.replace(/^\/?xl\//, '').replace(/^\/+/, '');
        return m;
      });
    return map;
  }

  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&'); }

  function nextIndex(files, prefix, suffix) {
    var n = 0;
    var re = new RegExp('^' + escRe(prefix) + '(\\d+)' + escRe(suffix) + '$');
    Object.keys(files).forEach(function (p) {
      var m = p.match(re);
      if (m) n = Math.max(n, +m[1]);
    });
    return n + 1;
  }

  function ensureSheetRels(files, sheetPath) {
    var relsPath = sheetPath.replace(/([^\/]+)$/, '_rels/$1.rels');
    if (!files[relsPath]) {
      files[relsPath] = enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
    }
    return relsPath;
  }

  function addRel(files, relsPath, type, target) {
    var xml = dec(files[relsPath]);
    var ids = [];
    xml.replace(/Id="rId(\d+)"/g, function (m, n) { ids.push(+n); return m; });
    var id = 'rId' + ((ids.length ? Math.max.apply(null, ids) : 0) + 1);
    var rel = '<Relationship Id="' + id + '" Type="' + type + '" Target="' + target + '"/>';
    files[relsPath] = enc(xml.replace('</Relationships>', rel + '</Relationships>'));
    return id;
  }

  function addContentTypeOverride(files, partName, contentType) {
    var ct = dec(files['[Content_Types].xml']);
    if (ct.indexOf('PartName="' + partName + '"') !== -1) return;
    var o = '<Override PartName="' + partName + '" ContentType="' + contentType + '"/>';
    files['[Content_Types].xml'] = enc(ct.replace('</Types>', o + '</Types>'));
  }

  function insertDrawingEl(sheetXml, relId) {
    var el = '<drawing r:id="' + relId + '"/>';
    if (!/<worksheet\b[^>]*xmlns:r=/.test(sheetXml)) {
      sheetXml = sheetXml.replace(/<worksheet\b/,
        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    }
    // <drawing> must precede these if present, else go before </worksheet>
    var before = ['<legacyDrawing', '<tableParts', '<extLst', '</worksheet>'];
    for (var i = 0; i < before.length; i++) {
      var idx = sheetXml.indexOf(before[i]);
      if (idx !== -1) return sheetXml.slice(0, idx) + el + sheetXml.slice(idx);
    }
    return sheetXml;
  }

  function inject(bytes, placements) {
    var u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    var files = root.fflate.unzipSync(u8);
    var paths = sheetPathMap(files);
    var drawingForSheet = {}; // sheetPath -> { drawPath, drawRels, anchors:[] }

    placements.forEach(function (pl) {
      var sheetPath = paths[pl.sheetName];
      if (!sheetPath) throw new Error('sheet not found: ' + pl.sheetName);

      var dg = drawingForSheet[sheetPath];
      if (!dg) {
        var di = nextIndex(files, 'xl/drawings/drawing', '.xml');
        var drawPath = 'xl/drawings/drawing' + di + '.xml';
        var drawRels = 'xl/drawings/_rels/drawing' + di + '.xml.rels';
        // Reserve the drawing part NOW (overwritten at the end) so nextIndex()
        // sees it when allocating the next sheet's drawing - otherwise a second
        // sheet reuses drawing1 and its chart rel is clobbered.
        files[drawPath] = enc('');
        files[drawRels] = enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
        addContentTypeOverride(files, '/' + drawPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
        var sheetRels = ensureSheetRels(files, sheetPath);
        var drawRelId = addRel(files, sheetRels,
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          '../drawings/drawing' + di + '.xml');
        files[sheetPath] = enc(insertDrawingEl(dec(files[sheetPath]), drawRelId));
        dg = drawingForSheet[sheetPath] = { drawPath: drawPath, drawRels: drawRels, anchors: [] };
      }

      var ci = nextIndex(files, 'xl/charts/chart', '.xml');
      var chartPath = 'xl/charts/chart' + ci + '.xml';
      files[chartPath] = enc(NC.chartXml(pl.def));
      addContentTypeOverride(files, '/' + chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      var chartRelId = addRel(files, dg.drawRels,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
        '../charts/chart' + ci + '.xml');
      dg.anchors.push(NC.drawingXml(pl.anchor, chartRelId));
    });

    // write each sheet's drawing xml with all its anchors
    Object.keys(drawingForSheet).forEach(function (sp) {
      var dg = drawingForSheet[sp];
      var body = dg.anchors.map(function (a) {
        return a.replace(/^<\?xml[^>]*\?>/, '').replace(/<\/?xdr:wsDr[^>]*>/g, '');
      }).join('');
      files[dg.drawPath] = enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' + body + '</xdr:wsDr>');
    });

    return root.fflate.zipSync(files);
  }

  // Writes a chart's categories + series values into hidden, far-right helper
  // columns on an ExcelJS worksheet and returns the { categories, series }
  // portion of a chartDef pointing at those cells. Lets a report data-link a
  // native chart without depending on the visible table's exact column layout.
  // Charts plot hidden cells because chartXml sets c:plotVisOnly val="0".
  function colLetter(n) { // 1-based -> A, B, ... AA
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m - 1) / 26); }
    return s;
  }
  // Builds an absolute A1 range string pointing at a report's VISIBLE table
  // cells, e.g. ref('Summary', 2, 3, 8) -> "'Summary'!$B$3:$B$8". Charts use
  // these so category/series labels are editable by editing the table cell.
  function ref(sheet, col1, row1, row2) {
    var c = colLetter(col1);
    return "'" + sheet + "'!$" + c + "$" + row1 + (row2 && row2 !== row1 ? ':$' + c + '$' + row2 : '');
  }
  var API = { inject: inject, ref: ref, _colLetter: colLetter };
  root.NativeChartInject = API;
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
