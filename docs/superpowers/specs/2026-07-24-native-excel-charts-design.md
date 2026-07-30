# Native, editable Excel charts in report outputs

**Date:** 2026-07-24
**Status:** Approved design - ready for implementation plan
**Owner:** PRP Automation Dashboard

---

## 1. Problem & goal

The PRP pipeline is: three raw exports are merged into one consolidated workbook
(`CreatePRP`), then the user runs one of **7 report steps** (`d1, d2, d3, s1c,
s2a, s3d, s4`), each of which produces output `.xlsx` files (e.g.
`Risk_Output.xlsx`, `OneTrust_Report.xlsx`) with charts.

Today every chart is **baked to a flat PNG image** (`renderStyledPng`) and pasted
into the workbook via ExcelJS `addImage`. The PNG is not editable: the user
cannot change bar colors or legend/series labels in the delivered file.

**Goal:** every chart in the 7 reports' output files must be a **native,
editable Excel chart** - a real chart object, pre-filled with the same colors and
legend labels used today, that the user recolors and relabels **in Excel using
Excel's own chart tools**. There is **no in-app editor**; the editing capability
is Excel itself.

This restores the original design intent: the code comments state the outputs
were meant to "reproduce the original openpyxl native BarChart" and "original
xlsxwriter native column chart" - the PNG-baking was only a workaround for
ExcelJS's lack of chart support.

### Non-goals (YAGNI)
- No in-app color picker / legend editor.
- No persistence of user choices (Excel owns the file after download).
- No new chart shapes beyond **bar/column** (all 7 reports are bar charts).
- No change to data, tables, table formatting, or the report-selection flow.
- The **in-app on-screen preview stays as-is** (still a `renderStyledPng` image).

## 2. Success criteria (acceptance)

For each of the 7 reports, opening the output `.xlsx` in **Microsoft Excel**:
1. Each chart is a genuine Excel **chart object** (selectable, "Chart Design" /
   "Format" ribbon appears), not a picture.
2. Bar **fill colors** are editable via Excel (Format Data Series → Fill).
3. **Legend / series names** are editable via Excel (Select Data → Edit series,
   or edit the referenced header cell).
4. The chart is **data-linked** to cells and renders immediately on open
   (cached values present; recalculates when source cells change).
5. Default colors and legend labels **match today's baked-PNG output** so nothing
   visually regresses out of the box.
6. The file opens without repair prompts in **Excel and LibreOffice Calc**.

## 3. Constraints & key facts

- **Fully offline app.** `index.html` opens from `file://` with no server/build
  step. Vendored libs: `plotly`, `xlsx` (SheetJS community), `exceljs`.
- **ExcelJS cannot create native charts**; SheetJS community cannot write charts.
  Native charts therefore require **hand-written OOXML** chart parts injected into
  the `.xlsx` (an OPC zip) **after** ExcelJS writes it.
- ExcelJS bundles JSZip/Pako internally but **does not expose** them. We add a
  small standalone zip library, **`fflate`** (~30 KB), to `vendor/` with a CDN
  fallback, matching the existing vendor-loader pattern.
- **Dual maintenance:** `index.html` **inlines** `report-engine.js`, `bridge.js`,
  and all `js/reports/*.js`. Nothing loads the modular files at runtime, but the
  user requires both kept in sync. Every code change lands in **both** the module
  file and its inlined copy in `index.html`, verified to match.
- **One report runs at a time:** `handleRun()` → `runReport(entry, file)` parses
  the file, runs `Reports[id](wb)`, stores `S.result`, renders. The output
  `.xlsx` bytes are built inside each report module.

## 4. Chart inventory (source of truth for defaults)

Defaults below must be preserved. Exact sizes and anchor cells are taken from each
report's **current `addImage` call** during implementation.

| Report | Output file | Chart (sheet) | Kind | Series (name → color) | Legend | Notes |
|---|---|---|---|---|---|---|
| d1 | `PRP_Output.xlsx` | Final Table | grouped column | Tier-1 Supplier → `#1f77b4`; Supplier Added by Zone → `#ff7f0e` | yes | inside data labels |
| d2 | `output file D2.xlsx` | Dashboard | grouped column | Open → `#00AEEF`; Closed → `#D4AF37` | yes | only the "org" chart is embedded today (kpi chart is built but discarded) |
| d3 | `output.xlsx` | Summary | **stacked** column | Completed in 2026 → `#FF0000`; Pending → `#FFC000`; (fallback `#4472C4`) | yes | series set is dynamic from the pivot headers |
| s1c | `PRP_Final_Output3.xlsx` | Summary; Active by Zone | single-series column | value → `#4472C4` | **no** | outside data labels, axis titles |
| s2a | `OneTrust_Report.xlsx` | Auto Open Closed; Overdue sheet | grouped column | Closed → `#2F75B5`; Open → `#ED7D31` | yes | |
| s3d | `Risk_Output.xlsx` | Open vs Overdue; Zone Summary | grouped column | Open → `#156082`, Overdue → `#C00000`; Total Risks → `#156082`, Open Risks → `#F26C23` | yes | |
| s4 | `Risk_Output.xlsx` | Dashboard | single-series, **per-point colors** | end bars → `#BFBFBF`, middle bars → `#FFC000` | **no** | black chart/plot area, white centered value labels |

The `chartDef` for each is derived from the **same trace data the report already
computes** (`x` = categories, `y` = values, `name`, `marker.color`), so no new
data calculation is introduced.

## 5. Architecture

Three units with clear boundaries:

### 5.1 `NativeChart` - OOXML chart-part generator (new module)
**Purpose:** turn a `chartDef` into OOXML strings. Pure, no I/O, no zip.

**Interface**
```
NativeChart.buildChartXml(chartDef) -> string   // xl/charts/chartN.xml
NativeChart.buildDrawingXml(anchor, chartRelId) -> string  // xl/drawings/drawingN.xml
```
**`chartDef` shape**
```
{
  barDir: 'col',                 // all reports are column charts
  grouping: 'clustered'|'stacked'|'standard',
  categories: { ref: "'Sheet'!$A$2:$A$6", cache: ['Z1','Z2',...] },
  series: [
    { name: { ref?: "'Sheet'!$B$1", lit?: 'Open' },
      values: { ref: "'Sheet'!$B$2:$B$6", cache: [3,5,...] },
      color: '156082',           // srgb hex, no '#'
      points?: [{ idx: 0, color: 'BFBFBF' }, ...]   // per-point (s4)
    }, ...
  ],
  legend: true|false,
  title?: 'Open vs Overdue',
  dataLabels?: { show: true, position: 'inEnd'|'outEnd'|'ctr', color?: 'FFFFFF' },
  plotBg?: '000000', chartBg?: '000000',   // s4 dark theme
  axisFontColor?: 'FFFFFF'
}
```
**Depends on:** nothing. **Testable** in isolation by asserting the XML contains
the expected `c:barChart`, `c:ser`, `a:solidFill`, `c:dPt`, `c:legend` nodes.

### 5.2 `injectNativeCharts` - zip patcher (new, in `bridge.js`)
**Purpose:** given the ExcelJS output buffer and a list of chart placements,
return a new buffer whose sheets carry native charts.

**Interface**
```
injectNativeCharts(xlsxArrayBuffer, placements) -> Promise<ArrayBuffer>
// placements: [{ sheetName, chartDef, anchor: {fromCol,fromRow,toCol,toRow} }, ...]
```
**Behavior (using `fflate`):**
1. `unzipSync` the buffer into `{ path: Uint8Array }`.
2. For each placement, allocate `chartN.xml` + `drawingN.xml`, write them and
   their `_rels`.
3. Patch `[Content_Types].xml`: add `Override`s for each chart part
   (`…drawingml.chart+xml`) and drawing part (`…drawing+xml`).
4. Patch the target `xl/worksheets/sheetX.xml` (map sheet name → file via
   `workbook.xml`/`workbook.xml.rels`): inject `<drawing r:id="…"/>` before
   `</worksheet>`, and add the drawing relationship to that sheet's `_rels`.
5. `zipSync` back to an `ArrayBuffer`.

**Depends on:** `fflate`, `NativeChart`. **Isolated:** takes bytes in, bytes out;
does not know about reports.

### 5.3 Data-linking rule
Charts reference **real cells** (`c:strRef`/`c:numRef` with `c:f`) plus cached
values (`c:strCache`/`c:numCache`) so they render before recalculation and stay
editable. Rule per chart:
- If the categories/values already exist as a contiguous block in an output
  sheet, reference those cells.
- Otherwise, write a small **dedicated data block** into the sheet (hidden columns
  or an out-of-the-way area) and reference it.

Implementation records the exact `c:f` addresses when the report writes its grid.

### 5.4 Per-report integration
Each report module changes from *"bake PNG + `addImage`"* to *"keep PNG for
preview only, build `chartDef`(s) + record anchor, then `injectNativeCharts` on
the final buffer."* Concretely:
- Keep `renderStyledPng(...)` and the `chartImages` return value **unchanged**
  (in-app preview is untouched).
- **Remove** the `ws.addImage(...)` embed(s) for charts that become native.
- After `workbook.xlsx.writeBuffer()`, call `injectNativeCharts(buf, placements)`
  and return the patched buffer as `files[].bytes`.
- `placements` reuse each chart's **existing anchor cell and size**.

### 5.5 Vendor & loader
Add `vendor/fflate.min.js` and register it in the vendor loader
(`VENDORS` list in `index.html`) with a CDN fallback
(`https://cdn.jsdelivr.net/npm/fflate/umd/index.js`). If fflate fails to load,
`injectNativeCharts` falls back to the current baked-PNG embed so the app still
produces a (non-native) file rather than failing.

## 6. Data flow

```
handleRun ─► runReport(entry, file)
              ├─ XLSX.read(file) ─► wb
              └─ Reports[id](wb):
                   compute grids + traces  (unchanged)
                   renderStyledPng(traces) ─► PNG  (preview only, unchanged)
                   ExcelJS build sheets     (unchanged, minus chart addImage)
                   writeBuffer()            ─► buf
                   build chartDef(s) from traces + record anchors
                   injectNativeCharts(buf, placements) ─► buf'   ◄── NEW
                   return { files:[{bytes: buf'}], chartImages }  // preview PNG kept
```

## 7. Error handling & fidelity
- **fflate missing / inject throws:** catch per report, log to the script log,
  fall back to the existing `addImage` PNG embed for that file. The app never
  hard-fails a run because of chart injection.
- **OOXML strictness:** primary risk. Mitigations: reference cells + caches;
  keep chart XML minimal and schema-valid; validate every output in Excel **and**
  LibreOffice; start integration with one report (s3d → `Risk_Output.xlsx`) as the
  reference implementation, confirm it opens clean, then apply the same generator
  to the rest.
- **Stacked (d3) and per-point (s4)** are the two special cases the generator must
  handle beyond plain clustered columns; both are covered by `grouping` and
  `points[]` in `chartDef`.

## 8. Testing strategy
- **Unit (NativeChart):** given a `chartDef`, assert the produced XML contains the
  right structural nodes and the expected `srgbClr` fills and series refs; verify
  stacked vs clustered `c:grouping`, per-point `c:dPt`, legend on/off.
- **Integration (injectNativeCharts):** round-trip a small ExcelJS buffer, assert
  the zip gains the chart/drawing/rels parts and patched content-types, and that
  it re-parses as a valid zip.
- **Manual acceptance:** run all 7 reports on the sample workbook; open each output
  in Excel + LibreOffice; confirm the 6 acceptance criteria in §2; confirm bar
  color and legend edits work in Excel.
- **Parity check:** diff the changed functions between `js/*` and their inlined
  copies in `index.html` to confirm they match.

## 9. Rollout order
1. Vendor `fflate` + loader wiring (with fallback).
2. `NativeChart` generator (clustered first) + unit tests.
3. `injectNativeCharts` zip patcher + integration test.
4. Wire **s3d** (`Risk_Output.xlsx`) as the reference report; validate in Excel.
5. Extend to grouped reports (d1, d2, s2a), single-series (s1c), stacked (d3),
   per-point (s4).
6. Mirror all changes into `index.html`; run the full parity + acceptance pass.
