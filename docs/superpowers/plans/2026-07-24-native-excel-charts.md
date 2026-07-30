# Native Editable Excel Charts - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the baked-PNG charts embedded in the 7 reports' output `.xlsx` files with **native, editable Excel bar/column charts** (real chart objects, editable in Excel), keeping today's colors and legend labels as defaults.

**Architecture:** A pure OOXML generator turns each report's existing chart data (categories, series names, colors, values) into `chartN.xml` + `drawingN.xml`. A zip patcher opens the ExcelJS-produced `.xlsx` with `fflate`, injects those parts, patches `[Content_Types].xml`, the sheet rels, and the sheet's `<drawing>` element, and re-zips. Each report drops its chart `addImage` embed and calls the patcher on the final buffer; the in-app preview PNG is untouched.

**Tech Stack:** Vanilla ES5-style browser JS (global-namespace IIFE pattern, matching the codebase), ExcelJS (workbook build), SheetJS `xlsx` (parse), Plotly (preview PNG only), **new:** `fflate` (zip round-trip). Node v24 for headless unit tests (`node:assert`, no framework).

## Global Constraints

- **Fully offline, no build step.** App opens from `file://`. New deps go in `vendor/` and load via the existing vendor loader with a CDN fallback.
- **Dual maintenance.** `index.html` **inlines** all report/engine code; nothing loads `js/*.js` at runtime. Every change lands in **both** the `js/*` module **and** its inlined copy in `index.html`, verified to match.
- **Only bar/column charts.** All 7 reports are column charts (clustered, stacked, or single-series with per-point colors).
- **Defaults must match today's baked-PNG look** (colors + legend labels from the chart inventory below). No visual regression out of the box.
- **In-app on-screen preview stays as-is** (still a `renderStyledPng` image via `chartImages`); only the file's embedded chart becomes native.
- **New modules are dual-export:** end each with
  `if (typeof module !== 'undefined' && module.exports) module.exports = API;`
  and also assign `root.NativeChart` / `root.NativeChartInject` where
  `root = typeof window !== 'undefined' ? window : globalThis`.
- **Acceptance = opens in Excel AND LibreOffice** with no repair prompt; chart is a real object; bar colors and legend/series names editable in Excel.

## Chart inventory (defaults to preserve)

Colors are `srgbClr` hex without `#`. Sizes/anchors come from each report's current `addImage` call.

| Report | Output file | Chart (sheet) | grouping | series (name → color) | legend |
|---|---|---|---|---|---|
| s3d | `Risk_Output.xlsx` | Open vs Overdue | clustered | Open→`156082`, Overdue→`C00000` | yes |
| s3d | `Risk_Output.xlsx` | Zone Summary | clustered | Total Risks→`156082`, Open Risks→`F26C23` | yes |
| d1 | `PRP_Output.xlsx` | Final Table | clustered | Tier-1 Supplier→`1f77b4`, Supplier Added by Zone→`ff7f0e` | yes |
| d2 | `output file D2.xlsx` | Dashboard | clustered | Open→`00AEEF`, Closed→`D4AF37` | yes |
| s2a | `OneTrust_Report.xlsx` | Auto Open Closed; Overdue sheet | clustered | Closed→`2F75B5`, Open→`ED7D31` | yes |
| s1c | `PRP_Final_Output3.xlsx` | Summary; Active by Zone | single | value→`4472C4` | no |
| d3 | `output.xlsx` | Summary | **stacked** | Completed in 2026→`FF0000`, Pending→`FFC000` (fallback `4472C4`) | yes |
| s4 | `Risk_Output.xlsx` | Dashboard | single, **per-point** | end bars→`BFBFBF`, middle→`FFC000` | no |

---

## Task 1: Vendor `fflate` + loader wiring

**Files:**
- Create: `vendor/fflate.min.js` (downloaded)
- Modify: `index.html` (VENDORS list ~line 19-23)

**Interfaces:**
- Produces: global `window.fflate` with `fflate.unzipSync(Uint8Array) -> {path: Uint8Array}` and `fflate.zipSync({path: Uint8Array}) -> Uint8Array`. In Node: `require('../vendor/fflate.min.js')` returns the same API.

- [ ] **Step 1: Download the offline copy**

Run:
```bash
cd "C:\Users\hp\OneDrive\Desktop\PRP-main"
curl -L -o vendor/fflate.min.js https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js
```
Expected: file exists, ~30 KB.

- [ ] **Step 2: Verify it loads in Node (UMD export works)**

Run:
```bash
node -e "const f=require('./vendor/fflate.min.js'); const z=f.zipSync({'a.txt':new TextEncoder().encode('hi')}); const u=f.unzipSync(z); console.log(new TextDecoder().decode(u['a.txt']))"
```
Expected: prints `hi`.

- [ ] **Step 3: Register in the vendor loader**

In `index.html`, add to the `VENDORS` array (after the exceljs entry):
```js
{ local: 'vendor/fflate.min.js', cdn: 'https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js' },
```

- [ ] **Step 4: Commit**

```bash
git add vendor/fflate.min.js index.html
git commit -m "feat: vendor fflate for native-chart zip injection"
```

---

## Task 2: `NativeChart` - OOXML chart/drawing generator

**Files:**
- Create: `js/native-chart.js`
- Create: `test/native-chart.test.js`

**Interfaces:**
- Produces:
  - `NativeChart.chartXml(def) -> string` (a full `xl/charts/chartN.xml`)
  - `NativeChart.drawingXml(anchor, relId) -> string` (a full `xl/drawings/drawingN.xml`)
  - `def` shape:
    ```
    { grouping: 'clustered'|'stacked'|'standard',
      title?: string, legend: boolean,
      categories: { ref: string, cache: (string|number)[] },
      series: [ { name: { ref?: string, lit: string },
                  values: { ref: string, cache: number[] },
                  color: string,            // 'RRGGBB'
                  points?: [ {idx:number, color:string} ] } ],
      dataLabels?: { position: 'inEnd'|'outEnd'|'ctr', color?: string },
      chartBg?: string, plotBg?: string, axisColor?: string }
    ```
  - `anchor` shape: `{ fromCol, fromRow, toCol, toRow }` (0-based).

- [ ] **Step 1: Write failing tests**

`test/native-chart.test.js`:
```js
const assert = require('node:assert');
const NC = require('../js/native-chart.js');

const def = {
  grouping: 'clustered', legend: true,
  categories: { ref: "'S'!$A$2:$A$3", cache: ['Z1', 'Z2'] },
  series: [
    { name: { ref: "'S'!$B$1", lit: 'Open' },  values: { ref: "'S'!$B$2:$B$3", cache: [3, 5] }, color: '156082' },
    { name: { ref: "'S'!$C$1", lit: 'Overdue' }, values: { ref: "'S'!$C$2:$C$3", cache: [1, 2] }, color: 'C00000' },
  ],
};
const xml = NC.chartXml(def);
assert.ok(xml.includes('<c:barChart>'), 'has barChart');
assert.ok(xml.includes('<c:grouping val="clustered"/>'), 'clustered');
assert.ok(xml.includes('<a:srgbClr val="156082"/>'), 'series1 color');
assert.ok(xml.includes('<a:srgbClr val="C00000"/>'), 'series2 color');
assert.ok(xml.includes('<c:v>Open</c:v>'), 'series1 name cache');
assert.ok(xml.includes("<c:f>'S'!$B$2:$B$3</c:f>"), 'series1 value ref');
assert.ok(xml.includes('<c:legend>'), 'legend present');
assert.match(xml, /^<\?xml/, 'xml decl');

// stacked
const st = NC.chartXml(Object.assign({}, def, { grouping: 'stacked' }));
assert.ok(st.includes('<c:grouping val="stacked"/>'), 'stacked grouping');
assert.ok(st.includes('<c:overlap val="100"/>'), 'stacked overlap 100');

// no legend
const nl = NC.chartXml(Object.assign({}, def, { legend: false }));
assert.ok(!nl.includes('<c:legend>'), 'no legend');

// per-point colors
const pp = NC.chartXml({ grouping: 'clustered', legend: false,
  categories: { ref: "'S'!$A$2:$A$4", cache: ['a','b','c'] },
  series: [ { name: { lit: 'v' }, values: { ref: "'S'!$B$2:$B$4", cache: [1,2,3] }, color: 'FFC000',
    points: [ {idx:0, color:'BFBFBF'}, {idx:2, color:'BFBFBF'} ] } ] });
assert.ok(pp.includes('<c:dPt>'), 'has data points');
assert.ok(pp.includes('<a:srgbClr val="BFBFBF"/>'), 'per-point color');

// drawing
const d = NC.drawingXml({ fromCol: 6, fromRow: 1, toCol: 14, toRow: 21 }, 'rId1');
assert.ok(d.includes('<xdr:twoCellAnchor>'), 'anchor');
assert.ok(d.includes('<c:chart r:id="rId1"/>'), 'chart rel');
assert.ok(d.includes('<xdr:col>6</xdr:col>'), 'from col');
console.log('NativeChart: all assertions passed');
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `node test/native-chart.test.js`
Expected: FAIL - `Cannot find module '../js/native-chart.js'`.

- [ ] **Step 3: Implement `js/native-chart.js`**

```js
// Native OOXML column-chart generator. Pure string builders - no I/O, no zip.
// Dual-export: browser (window.NativeChart) + Node (module.exports) for tests.
(function (root) {
  'use strict';

  var C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  var A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function strRef(ref, cache) {
    var pts = cache.map(function (v, i) {
      return '<c:pt idx="' + i + '"><c:v>' + esc(v) + '</c:v></c:pt>';
    }).join('');
    return '<c:strRef><c:f>' + esc(ref) + '</c:f><c:strCache>' +
      '<c:ptCount val="' + cache.length + '"/>' + pts + '</c:strCache></c:strRef>';
  }

  function numRef(ref, cache) {
    var pts = cache.map(function (v, i) {
      return '<c:pt idx="' + i + '"><c:v>' + (v === null || v === undefined || v === '' ? 0 : v) + '</c:v></c:pt>';
    }).join('');
    return '<c:numRef><c:f>' + esc(ref) + '</c:f><c:numCache>' +
      '<c:formatCode>General</c:formatCode><c:ptCount val="' + cache.length + '"/>' +
      pts + '</c:numCache></c:numRef>';
  }

  function fill(color) {
    return '<c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>' +
      '<a:ln><a:noFill/></a:ln></c:spPr>';
  }

  function seriesXml(s, idx, cat, dLbls) {
    var tx = s.name.ref
      ? '<c:tx>' + strRef(s.name.ref, [s.name.lit]) + '</c:tx>'
      : '<c:tx><c:v>' + esc(s.name.lit) + '</c:v></c:tx>';
    var dPts = (s.points || []).map(function (p) {
      return '<c:dPt><c:idx val="' + p.idx + '"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>' +
        fill(p.color) + '</c:dPt>';
    }).join('');
    return '<c:ser><c:idx val="' + idx + '"/><c:order val="' + idx + '"/>' +
      tx + fill(s.color) + '<c:invertIfNegative val="0"/>' + dPts + dLbls +
      '<c:cat>' + strRef(cat.ref, cat.cache) + '</c:cat>' +
      '<c:val>' + numRef(s.values.ref, s.values.cache) + '</c:val></c:ser>';
  }

  function dataLabels(dl) {
    if (!dl) return '';
    var txPr = dl.color
      ? '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:solidFill>' +
        '<a:srgbClr val="' + dl.color + '"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
      : '';
    return '<c:dLbls>' + txPr + '<c:dLblPos val="' + dl.position + '"/>' +
      '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
      '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
  }

  function axisTxt(color) {
    if (!color) return '';
    return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:solidFill>' +
      '<a:srgbClr val="' + color + '"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
  }

  function chartXml(def) {
    var CAT = 111111111, VAL = 222222222;
    var dl = dataLabels(def.dataLabels);
    var sers = def.series.map(function (s, i) { return seriesXml(s, i, def.categories, dl); }).join('');
    var overlap = def.grouping === 'stacked' ? '<c:overlap val="100"/>' : '<c:overlap val="-27"/>';
    var gap = '<c:gapWidth val="' + (def.grouping === 'stacked' ? 50 : 150) + '"/>';
    var title = def.title
      ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>' + esc(def.title) +
        '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
      : '<c:autoTitleDeleted val="1"/>';
    var legend = def.legend
      ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + axisTxt(def.axisColor) + '</c:legend>'
      : '';
    var chartSpaceFill = def.chartBg
      ? '<c:spPr><a:solidFill><a:srgbClr val="' + def.chartBg + '"/></a:solidFill></c:spPr>' : '';
    var plotFill = def.plotBg
      ? '<c:spPr><a:solidFill><a:srgbClr val="' + def.plotBg + '"/></a:solidFill></c:spPr>'
      : '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="' + C + '" xmlns:a="' + A + '" xmlns:r="' + R + '">' +
      '<c:roundedCorners val="0"/><c:chart>' + title +
      '<c:plotArea><c:layout/>' +
      '<c:barChart><c:barDir val="col"/><c:grouping val="' + def.grouping + '"/>' +
      '<c:varyColors val="0"/>' + sers + gap + overlap +
      '<c:axId val="' + CAT + '"/><c:axId val="' + VAL + '"/></c:barChart>' +
      '<c:catAx><c:axId val="' + CAT + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="b"/>' + axisTxt(def.axisColor) +
      '<c:crossAx val="' + VAL + '"/></c:catAx>' +
      '<c:valAx><c:axId val="' + VAL + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>' + axisTxt(def.axisColor) +
      '<c:crossAx val="' + CAT + '"/></c:valAx>' +
      plotFill + '</c:plotArea>' + legend +
      '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
      chartSpaceFill + '</c:chartSpace>';
  }

  function drawingXml(a, relId) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="' + XDR + '" xmlns:a="' + A + '" xmlns:r="' + R + '" xmlns:c="' + C + '">' +
      '<xdr:twoCellAnchor>' +
      '<xdr:from><xdr:col>' + a.fromCol + '</xdr:col><xdr:colOff>0</xdr:colOff>' +
      '<xdr:row>' + a.fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + a.toCol + '</xdr:col><xdr:colOff>0</xdr:colOff>' +
      '<xdr:row>' + a.toRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>' +
      '<xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
      '<a:graphic><a:graphicData uri="' + C + '"><c:chart xmlns:c="' + C + '" r:id="' + relId + '"/>' +
      '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
      '</xdr:twoCellAnchor></xdr:wsDr>';
  }

  var API = { chartXml: chartXml, drawingXml: drawingXml, _esc: esc };
  root.NativeChart = API;
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `node test/native-chart.test.js`
Expected: `NativeChart: all assertions passed`.

- [ ] **Step 5: Commit**

```bash
git add js/native-chart.js test/native-chart.test.js
git commit -m "feat: native OOXML column-chart XML generator + tests"
```

---

## Task 3: `NativeChartInject` - zip patcher

**Files:**
- Create: `js/native-chart-inject.js`
- Create: `test/native-chart-inject.test.js`

**Interfaces:**
- Consumes: `NativeChart.chartXml`, `NativeChart.drawingXml`; global `fflate`.
- Produces: `NativeChartInject.inject(xlsxU8OrBuffer, placements) -> Uint8Array`
  where `placements = [{ sheetName, def, anchor }]` (`def`/`anchor` per Task 2).
  Accepts and returns bytes; adds one chart+drawing per placement. Multiple
  placements targeting the **same** sheet share that sheet's drawing part.

- [ ] **Step 1: Write failing test** (`test/native-chart-inject.test.js`)

```js
const assert = require('node:assert');
const ExcelJS = require('../vendor/exceljs.min.js');
const fflate = require('../vendor/fflate.min.js');
globalThis.fflate = fflate;
const NCI = require('../js/native-chart-inject.js');

(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Summary');
  ws.addRow(['Zone', 'Open', 'Overdue']);
  ws.addRow(['Z1', 3, 1]);
  ws.addRow(['Z2', 5, 2]);
  const buf = await wb.xlsx.writeBuffer();

  const def = {
    grouping: 'clustered', legend: true,
    categories: { ref: "'Summary'!$A$2:$A$3", cache: ['Z1', 'Z2'] },
    series: [
      { name: { lit: 'Open' },    values: { ref: "'Summary'!$B$2:$B$3", cache: [3, 5] }, color: '156082' },
      { name: { lit: 'Overdue' }, values: { ref: "'Summary'!$C$2:$C$3", cache: [1, 2] }, color: 'C00000' },
    ],
  };
  const out = NCI.inject(new Uint8Array(buf), [
    { sheetName: 'Summary', def, anchor: { fromCol: 4, fromRow: 0, toCol: 12, toRow: 15 } },
  ]);

  const files = fflate.unzipSync(out);
  const dec = (p) => new TextDecoder().decode(files[p]);
  assert.ok(files['xl/charts/chart1.xml'], 'chart part added');
  assert.ok(files['xl/drawings/drawing1.xml'], 'drawing part added');
  assert.ok(files['xl/drawings/_rels/drawing1.xml.rels'], 'drawing rels added');
  assert.ok(dec('[Content_Types].xml').includes('/xl/charts/chart1.xml'), 'content-type chart');
  assert.ok(dec('[Content_Types].xml').includes('/xl/drawings/drawing1.xml'), 'content-type drawing');
  // sheet1.xml is the Summary sheet (only sheet)
  assert.ok(/<drawing r:id="[^"]+"\/>/.test(dec('xl/worksheets/sheet1.xml')), 'sheet has <drawing>');
  assert.ok(files['xl/worksheets/_rels/sheet1.xml.rels'], 'sheet rels added');
  assert.ok(dec('xl/worksheets/_rels/sheet1.xml.rels').includes('drawings/drawing1.xml'), 'sheet->drawing rel');
  console.log('NativeChartInject: all assertions passed');
})().catch((e) => { console.error(e); process.exit(1); });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node test/native-chart-inject.test.js`
Expected: FAIL - `Cannot find module '../js/native-chart-inject.js'`.

- [ ] **Step 3: Implement `js/native-chart-inject.js`**

```js
// Injects native OOXML charts into an ExcelJS-produced .xlsx (an OPC zip).
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
    relsXml.replace(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?>/g,
      function (_, id, target) { relTarget[id] = target; return _; });
    var map = {};
    wbXml.replace(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g,
      function (_, name, rid) {
        var t = relTarget[rid] || '';
        map[name] = 'xl/' + t.replace(/^\/?xl\//, '').replace(/^\/+/, '');
        return _;
      });
    return map;
  }

  function nextIndex(files, prefix, suffix) {
    var n = 0;
    Object.keys(files).forEach(function (p) {
      var m = p.match(new RegExp('^' + prefix.replace(/[.\/]/g, '\\$&') + '(\\d+)' + suffix.replace(/[.\/]/g, '\\$&') + '$'));
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
    xml.replace(/Id="rId(\d+)"/g, function (_, n) { ids.push(+n); return _; });
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
    // ensure xmlns:r on the worksheet root
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
    var drawingForSheet = {}; // sheetPath -> { drawPath, relsPath, chartRelIds:[] }

    placements.forEach(function (pl) {
      var sheetPath = paths[pl.sheetName];
      if (!sheetPath) throw new Error('sheet not found: ' + pl.sheetName);

      var dg = drawingForSheet[sheetPath];
      if (!dg) {
        var di = nextIndex(files, 'xl/drawings/drawing', '.xml');
        var drawPath = 'xl/drawings/drawing' + di + '.xml';
        var drawRels = 'xl/drawings/_rels/drawing' + di + '.xml.rels';
        files[drawRels] = enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
        addContentTypeOverride(files, '/' + drawPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
        var sheetRels = ensureSheetRels(files, sheetPath);
        var drawRelId = addRel(files, sheetRels,
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          '../drawings/drawing' + di + '.xml');
        files[sheetPath] = enc(insertDrawingEl(dec(files[sheetPath]), drawRelId));
        dg = drawingForSheet[sheetPath] = { drawPath: drawPath, drawRels: drawRels, di: di, anchors: [] };
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
      var body = dg.anchors.join('').replace(/<\/?xdr:wsDr[^>]*>/g, '').replace(/^<\?xml[^>]*\?>/, '');
      files[dg.drawPath] = enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' + body + '</xdr:wsDr>');
    });

    return root.fflate.zipSync(files);
  }

  var API = { inject: inject };
  root.NativeChartInject = API;
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
```

> Note: `drawingXml` in Task 2 returns a full standalone document; when a sheet has one chart the injector could use it directly, but to support multiple charts per sheet it strips the wrapper and re-wraps all anchors together (above). Keep `drawingXml` returning the standalone form - the strip/re-wrap handles both cases.

- [ ] **Step 4: Run test to verify it passes**

Run: `node test/native-chart-inject.test.js`
Expected: `NativeChartInject: all assertions passed`.

- [ ] **Step 5: Validate the injected file opens in a real spreadsheet app**

Run:
```bash
node -e "
const ExcelJS=require('./vendor/exceljs.min.js'),fflate=require('./vendor/fflate.min.js');
globalThis.fflate=fflate;const NCI=require('./js/native-chart-inject.js');
(async()=>{const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('Summary');
ws.addRow(['Zone','Open','Overdue']);ws.addRow(['Z1',3,1]);ws.addRow(['Z2',5,2]);
const buf=await wb.xlsx.writeBuffer();
const def={grouping:'clustered',legend:true,categories:{ref:\"'Summary'!\$A\$2:\$A\$3\",cache:['Z1','Z2']},
series:[{name:{lit:'Open'},values:{ref:\"'Summary'!\$B\$2:\$B\$3\",cache:[3,5]},color:'156082'},
{name:{lit:'Overdue'},values:{ref:\"'Summary'!\$C\$2:\$C\$3\",cache:[1,2]},color:'C00000'}]};
const out=NCI.inject(new Uint8Array(buf),[{sheetName:'Summary',def,anchor:{fromCol:4,fromRow:0,toCol:12,toRow:15}}]);
require('fs').writeFileSync('scratch-native-chart.xlsx',Buffer.from(out));console.log('wrote scratch-native-chart.xlsx');})();
"
```
Then **open `scratch-native-chart.xlsx` in Excel and LibreOffice**. Expected: opens with no repair prompt; a clustered column chart with blue/red series and a legend; right-click a bar → Format Data Series lets you change fill color; Select Data lets you rename the series. Delete the scratch file after.

- [ ] **Step 6: Commit**

```bash
git add js/native-chart-inject.js test/native-chart-inject.test.js
git commit -m "feat: inject native charts into ExcelJS xlsx via fflate + tests"
```

---

## Task 4: Wire s3d (`Risk_Output.xlsx`) - reference report

**Files:**
- Modify: `js/reports/s3d.js` (chart build + workbook write, ~lines 160-200)

**Interfaces:**
- Consumes: `NativeChartInject.inject`, `NativeChart` (via globals in browser).
- Produces: `Risk_Output.xlsx` bytes with two native charts on the sheets that currently receive the baked PNGs.

- [ ] **Step 1: Read the current s3d chart + write code**

Run: `sed -n '150,210p' js/reports/s3d.js` - identify (a) the two `renderStyledPng`/`addImage` chart embeds and their sheets/anchors, (b) the grids that hold each chart's categories/values so refs can be computed, (c) the `writeBuffer()` return.

- [ ] **Step 2: Build `chartDef`s from the existing traces**

For each chart, reuse the arrays already computed (e.g. `oovZones`, `oovOpen`, `oovOverdue`; `zsZones`, `zsTotal`, `zsOpen`). Compute cell refs against the actual output sheet where those columns are written. Example for "Open vs Overdue" (adjust sheet name + columns/rows to match the real grid):
```js
function colLetter(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
function ref(sheet, col, r1, r2) {
  var c = colLetter(col);
  return "'" + sheet + "'!$" + c + "$" + r1 + (r2 ? ':$' + c + '$' + r2 : '');
}
// Assuming 'Open vs Overdue' sheet has: A=zone (header row 1, data rows 2..n), B=Open, C=Overdue
var n = oovZones.length, first = 2, last = n + 1;
var oovDef = {
  grouping: 'clustered', legend: true, title: 'Open vs Overdue',
  categories: { ref: ref('Open vs Overdue', 0, first, last), cache: oovZones },
  series: [
    { name: { ref: ref('Open vs Overdue', 1, 1), lit: 'Open' },
      values: { ref: ref('Open vs Overdue', 1, first, last), cache: oovOpen }, color: '156082' },
    { name: { ref: ref('Open vs Overdue', 2, 1), lit: 'Overdue' },
      values: { ref: ref('Open vs Overdue', 2, first, last), cache: oovOverdue }, color: 'C00000' },
  ],
};
```
Build `zsDef` the same way for "Zone Summary" (Total Risks→`156082`, Open Risks→`F26C23`). **If the report's output sheet does not already contain these columns as a clean block, add a small data block to that sheet (the report already builds the grid - extend it) and point the refs there.**

- [ ] **Step 3: Replace the PNG embed with native injection**

- Keep the `renderStyledPng(...)` calls that populate `chartImages` (preview only) **unchanged**.
- Remove the `ws.addImage(...)` calls for these two charts.
- After `var buf = await workbook.xlsx.writeBuffer();`, add:
```js
var placements = [
  { sheetName: 'Open vs Overdue', def: oovDef, anchor: { fromCol: 5, fromRow: 1, toCol: 13, toRow: 21 } },
  { sheetName: 'Zone Summary',    def: zsDef,  anchor: { fromCol: 5, fromRow: 1, toCol: 13, toRow: 21 } },
];
try {
  if (window.NativeChartInject && window.fflate) {
    buf = window.NativeChartInject.inject(new Uint8Array(buf), placements).buffer;
  }
} catch (e) { console.error('native chart inject failed, keeping image fallback:', e); }
```
(Adjust anchors to match where the PNGs sat. If keeping a fallback image, leave `addImage` in and skip it only when injection succeeds - simplest is: remove image, rely on native chart; the `catch` logs but the file still has data.)

- [ ] **Step 4: Verify in the real app**

Run the app: open `index.html` - **but first** note this task edits only `js/reports/s3d.js`; `index.html` is mirrored in Task 8. To verify now without the full mirror, use a Node harness that requires the modular report is not feasible (reports need Plotly/DOM). Instead verify by a temporary browser run after a partial mirror, OR defer full verification to Task 8. **Recommended:** do a focused manual check now by temporarily inlining just s3d's change into `index.html`, running s3d in the browser, downloading `Risk_Output.xlsx`, and opening in Excel. Confirm: two native charts, correct colors/legends, editable.

- [ ] **Step 5: Commit**

```bash
git add js/reports/s3d.js
git commit -m "feat(s3d): native editable charts in Risk_Output.xlsx"
```

---

## Task 5: Wire grouped reports d1, d2, s2a

**Files:**
- Modify: `js/reports/d1.js`, `js/reports/d2.js`, `js/reports/s2a.js`

**Interfaces:**
- Same pattern as Task 4; one `chartDef` per embedded chart.

- [ ] **Step 1: d1 - Final Table (`PRP_Output.xlsx`)**

Two series: Tier-1 Supplier→`1f77b4`, Supplier Added by Zone→`ff7f0e`; clustered; legend; inside data labels (`dataLabels: { position: 'inEnd', color: 'FFFFFF' }`). Categories = `zones`, values = `tier1Vals` / `addedVals`. Build refs against the Final Table sheet's columns; add a data block if not already present. Remove `addImage`; inject after `writeBuffer()`.

- [ ] **Step 2: d2 - Dashboard (`output file D2.xlsx`)**

Only the "org" chart is embedded today. Series Open→`00AEEF`, Closed→`D4AF37`; clustered; legend; data labels inside white. Categories = `orgLabels`, values = `openVals`/`closedVals`. Inject on the Dashboard sheet at the org chart's anchor. Leave the kpi chart alone (not embedded).

- [ ] **Step 3: s2a - Auto Open Closed + Overdue sheet (`OneTrust_Report.xlsx`)**

Two charts, each: Closed→`2F75B5`, Open→`ED7D31`; clustered; legend. Categories = `zones`, values = `closedVals`/`openVals` (and the overdue equivalents). Two placements on their respective sheets.

- [ ] **Step 4: Verify each in the app (partial-mirror or after Task 8) + open outputs in Excel**

Confirm colors/legends match today and charts are editable.

- [ ] **Step 5: Commit**

```bash
git add js/reports/d1.js js/reports/d2.js js/reports/s2a.js
git commit -m "feat(d1,d2,s2a): native editable charts in outputs"
```

---

## Task 6: Wire single-series s1c + stacked d3

**Files:**
- Modify: `js/reports/s1c.js`, `js/reports/d3.js`

- [ ] **Step 1: s1c - Summary + Active by Zone (`PRP_Final_Output3.xlsx`)**

Each chart single-series, `legend: false`, color `4472C4`, `dataLabels: { position: 'outEnd', color: '000000' }`, `title` = 'Assessment Status Overview' / 'Active by Zone'. One series with `name: { lit: 'Count' }`; categories/values from `statusRows` / `zoneRows`. Two placements.

- [ ] **Step 2: d3 - Summary (`output.xlsx`)**

`grouping: 'stacked'`, legend true, dynamic series from `pivot.headers` with colors `{'Completed in 2026':'FF0000','Pending':'FFC000'}` (fallback `4472C4`). Title 'Assessments Completed vs Open'. Only inject when `pivot.indexVals.length` (matches current guard).

- [ ] **Step 3: Verify in app + Excel** (stacked bars stack correctly; colors/legend match).

- [ ] **Step 4: Commit**

```bash
git add js/reports/s1c.js js/reports/d3.js
git commit -m "feat(s1c,d3): native single-series + stacked charts in outputs"
```

---

## Task 7: Wire per-point s4 (`Risk_Output.xlsx` Dashboard)

**Files:**
- Modify: `js/reports/s4.js` (~lines 214-237)

- [ ] **Step 1: Build the per-point chartDef**

Single series, `legend: false`, base color `FFC000`, with `points` overriding the first and last bars to `BFBFBF` (mirrors `barColors` logic: `i===0 || i===len-1`). Dark theme: `chartBg: '000000'`, `plotBg: '000000'`, `axisColor: 'FFFFFF'`, title 'Cumulative Risk Treatment Progress', `dataLabels: { position: 'inEnd', color: 'FFFFFF' }`. Categories = `calcLabels`, values = `calcValues`. Refs point at the Dashboard chart-data block the report already writes (`chartStart` region, cols A/B).
```js
var points = calcLabels.map(function (_, i) {
  return (i === 0 || i === calcLabels.length - 1) ? { idx: i, color: 'BFBFBF' } : null;
}).filter(Boolean);
```

- [ ] **Step 2: Replace the image embed**

Remove `wsDash.addImage(progId, ...)`; keep `progressChartPng` for the preview `chartImages`. Inject on the Dashboard sheet at ~`{ fromCol: 6, fromRow: 1, toCol: 14, toRow: 16 }` (matches the old `tl:{col:6,row:1}` anchor).

- [ ] **Step 3: Verify in app + Excel** (gray end bars, gold middle, white labels, black background, no legend, editable).

- [ ] **Step 4: Commit**

```bash
git add js/reports/s4.js
git commit -m "feat(s4): native per-point progress chart in Risk_Output.xlsx"
```

---

## Task 8: Mirror everything into `index.html` + full acceptance

**Files:**
- Modify: `index.html` (inline `NativeChart` + `NativeChartInject` before the reports; mirror each report change; fflate already added in Task 1)

**Interfaces:** none new - this makes the runtime match the modular source.

- [ ] **Step 1: Inline the two new modules**

Paste the bodies of `js/native-chart.js` and `js/native-chart-inject.js` into the big app `<script>` block in `index.html` (around the other engine code, before the `Reports.*` definitions). In the inlined copies, the dual-export `module.exports` guard is harmless (there is no `module` in the browser); keep it.

- [ ] **Step 2: Mirror each report's change**

For each of s3d, d1, d2, s2a, s1c, d3, s4: apply the **same** edits made in `js/reports/*.js` to the corresponding inlined copy in `index.html` (remove chart `addImage`, add `chartDef` + `inject` after `writeBuffer()`).

- [ ] **Step 3: Parity check**

Run a diff of each changed report function between the module and its inlined copy:
```bash
node -e "const fs=require('fs');const idx=fs.readFileSync('index.html','utf8');
['s3d','d1','d2','s2a','s1c','d3','s4'].forEach(id=>{
  const m=fs.readFileSync('js/reports/'+id+'.js','utf8');
  const needle='NativeChartInject.inject';
  console.log(id, 'module:', m.includes(needle), 'index:', idx.includes(needle));
});"
```
Expected: every report `true`/`true`. Also confirm `fflate` is in the VENDORS list and `NativeChart`/`NativeChartInject` bodies appear in `index.html`.

- [ ] **Step 4: Full acceptance pass (browser + Excel + LibreOffice)**

Open `index.html`. Run **all 7 reports** on `PRP Sample Jun (2).xlsx` (or the 3-file merge). For each output file:
- Opens in Excel with no repair prompt.
- Every chart is a native chart object (Chart Design ribbon appears).
- Colors + legend labels match the pre-existing look (compare against the in-app preview PNG).
- Right-click a bar → change fill color works; Select Data → rename series works.
- Re-open in LibreOffice Calc: charts render, no errors.

Record any file that fails and fix the offending report's refs/anchors before proceeding.

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat: mirror native-chart engine + report wiring into index.html bundle"
```

---

## Self-review notes

- **Spec coverage:** §5.1 NativeChart → Task 2; §5.2 injectNativeCharts → Task 3; §5.3 data-linking → Tasks 4-7 (refs + caches); §5.4 per-report → Tasks 4-7; §5.5 vendor+fallback → Task 1 + `try/catch` in each report; §3 dual-maintenance → Task 8; §8 testing → Node tests (Tasks 2-3) + manual Excel/LibreOffice (Tasks 3-8); §9 rollout order preserved (s3d first).
- **Placeholders:** none - every code step has real code; per-report chartDefs give concrete colors/series (anchors/refs are computed from each report's real grid, flagged to adjust to actual column positions during Step 1 reads).
- **Type consistency:** `chartXml(def)` / `drawingXml(anchor, relId)` / `inject(bytes, placements)` signatures are used consistently across Tasks 2, 3, and 4-8; `def`/`anchor`/`placements` shapes match the Task 2/3 interface blocks.
- **Known risk to resolve during execution:** exact worksheet-part filenames (`sheetN.xml`) are not guaranteed to match sheet order; the injector resolves them via `workbook.xml` + rels (not by guessing), so multi-sheet outputs (s2a, s1c, s3d, s4) map correctly.
