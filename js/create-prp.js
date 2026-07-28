// ── CreatePRP — in-browser port of create_prp.py ──────────────────────────
// Merges the 3 raw exports (risk-export / assessment-export / tprm|supplier)
// into the consolidated PRP workbook the report pipeline consumes.
//
// Intentional deviations from create_prp.py:
//  · TPRM sheet is named 'TPRM Web-Portal Export' (the Python writes
//    'TPRM WebApp Portal', but reports d1/s1c and the shipped sample output
//    read 'TPRM Web-Portal Export' — the Python name appears to be drift).
//  · Every appended formula cell also carries its computed result
//    ({formula, result}) so the in-browser reports — which read cached
//    values, never recalculate — see real data. Excel still recalculates on
//    open (calcProperties.fullCalcOnLoad), keeping TODAY() fresh.
//  · Tags are resolved via findOrAddColumn (case-insensitive) rather than
//    addColumn, so a source header of 'tags' is overwritten in place instead
//    of spawning a second 'Tags' column. See applyTags().
window.CreatePRP = (function () {
  'use strict';

  var SHEET_RISK = 'OneTrust - Risk Export';
  var SHEET_TPRM = 'TPRM Web-Portal Export';   // deviation — see header comment
  var SHEET_ASSESSMENT = 'OneTrust Assessment';
  var OUTPUT_NAME = 'PRP Sample Jun (2).xlsx';

  // ── tagging rules ───────────────────────────────────────────────────────
  // Replaces the manual step of filtering Respondents to a category, typing
  // the tag over the visible rows, clearing the filter, then filling the
  // rest. Matching is lowercased + substring, so a multi-respondent cell
  // ('Sreenivas Kukkala; A. Other') still resolves to Finance.
  var FINANCE_RESPONDENTS = ['sreenivas kukkala'];
  var FINANCE_TAG = 'Finance';
  var DEFAULT_TAG = 'Cyber';
  var TAGS_HEADER = 'Tags';
  var RESPONDENT_HEADERS = ['Respondents', 'Respondent', 'Respondent Name', 'Assignee'];

  var MS_PER_DAY = 86400000;
  var EXCEL_EPOCH = Date.UTC(1899, 11, 30);    // Excel serial 0
  var SEP26_SERIAL = serialFromDate(new Date(Date.UTC(2026, 8, 30))); // DATE(2026,9,30)

  function _yield() { return new Promise(function (res) { setTimeout(res, 0); }); }

  // ── classify — mirrors find_source_files()'s filename checks ────────────
  function classify(fileName) {
    var lower = String(fileName || '').toLowerCase();
    if (!/\.xlsx$/.test(lower)) return null;
    if (lower.indexOf('risk-export') !== -1) return 'risk';
    if (lower.indexOf('assessment-export') !== -1) return 'assessment';
    if (lower.indexOf('tprm') !== -1 || lower.indexOf('supplier') !== -1) return 'tprm';
    return null;
  }

  // ── date parsing — the 5 strptime formats, in the Python's order ────────
  // %Y is exactly 4 digits; day/month accept 1-2; invalid calendar dates
  // (31/02/…) fall through to the next format, like strptime raising.
  var DATE_FORMATS = [
    { re: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, d: 1, m: 2, y: 3 },  // %d/%m/%Y
    { re: /^(\d{1,2})-(\d{1,2})-(\d{4})$/,  d: 1, m: 2, y: 3 },   // %d-%m-%Y
    { re: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, d: 2, m: 1, y: 3 },  // %m/%d/%Y
    { re: /^(\d{4})-(\d{1,2})-(\d{1,2})$/,  d: 3, m: 2, y: 1 },   // %Y-%m-%d
    { re: /^(\d{1,2})-(\d{1,2})-(\d{4})$/,  d: 2, m: 1, y: 3 },   // %m-%d-%Y
  ];
  // UTC construction is load-bearing: ExcelJS serializes Dates via UTC epoch
  // math, so a local-midnight Date in a UTC+ timezone saves as the previous day.
  function parseDateString(s) {
    var str = String(s).trim();
    for (var i = 0; i < DATE_FORMATS.length; i++) {
      var f = DATE_FORMATS[i], m = f.re.exec(str);
      if (!m) continue;
      var y = +m[f.y], mo = +m[f.m], d = +m[f.d];
      if (mo < 1 || mo > 12 || d < 1 || d > 31) continue;
      var dt = new Date(Date.UTC(y, mo - 1, d));
      if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) continue;
      return dt;
    }
    return null;
  }

  function serialFromDate(d) { return Math.round((d.getTime() - EXCEL_EPOCH) / MS_PER_DAY); }
  function todayUTC() {
    var now = new Date();
    return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  }

  // NETWORKDAYS(a,b): inclusive Mon–Fri count; negative when a > b.
  function networkdaysSerial(a, b) {
    var sign = 1;
    if (a > b) { var t = a; a = b; b = t; sign = -1; }
    var days = b - a + 1;
    var fullWeeks = Math.floor(days / 7);
    var count = fullWeeks * 5;
    var rem = days - fullWeeks * 7;
    var dow = new Date(EXCEL_EPOCH + a * MS_PER_DAY).getUTCDay(); // 0=Sun … 6=Sat
    for (var i = 0; i < rem; i++) {
      var d = (dow + i) % 7;
      if (d !== 0 && d !== 6) count++;
    }
    return sign * count;
  }

  // Excel-comparison view of a copied cell value:
  //  blank → serial 0 (Excel treats blank as 0 in arithmetic/comparisons),
  //  Date/number → its serial, anything else → text (arithmetic would #VALUE!).
  function serialOf(v) {
    if (v === null || v === undefined || v === '') return { kind: 'blank', serial: 0 };
    if (v instanceof Date) return { kind: 'num', serial: serialFromDate(v) };
    if (typeof v === 'number') return { kind: 'num', serial: v };
    if (typeof v === 'boolean') return { kind: 'num', serial: v ? 1 : 0 };
    return { kind: 'text', serial: null };
  }

  // Collapse an ExcelJS rich value to the plain value openpyxl would see.
  function plainVal(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
      if ('formula' in v || 'sharedFormula' in v) return v.result === undefined ? null : v.result;
      if ('text' in v) return v.text;   // hyperlink → display text
      if ('error' in v) return null;
      if ('result' in v) return v.result === undefined ? null : v.result;
      return null;
    }
    return v;
  }

  // Copy value preserving formulas ({formula, result}) like openpyxl copies
  // formula strings; everything else collapses to the plain value.
  function copyVal(v) {
    if (v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date) &&
        ('formula' in v || 'sharedFormula' in v)) {
      return { formula: v.formula || v.sharedFormula, result: v.result };
    }
    return plainVal(v);
  }

  // ── copy_sheet — coordinate-preserving value copy ───────────────────────
  async function copySheet(src, target, progress, label) {
    var rows = src.rowCount, cols = src.columnCount;
    for (var r = 1; r <= rows; r++) {
      var vals = src.getRow(r).values;   // sparse array indexed by column number
      for (var c = 1; c <= cols; c++) {
        var v = vals[c];
        if (v === undefined || v === null) continue;
        var nv = copyVal(v);
        if (nv === null) continue;
        var cell = target.getCell(r, c);
        cell.value = nv;
        // openpyxl auto-formats datetime cells; ExcelJS shows a raw serial
        // without a numFmt, so mirror openpyxl's default here.
        if (nv instanceof Date) cell.numFmt = 'yyyy-mm-dd h:mm:ss';
      }
      if ((r & 2047) === 0) {
        progress('Copying ' + label + ' — row ' + r.toLocaleString() + ' of ' + rows.toLocaleString() + '…');
        await _yield();
      }
    }
  }

  // ── find_column / add_column (the case-sensitivity asymmetry is Python's) ──
  function findColumn(ws, headerName) {
    var target = String(headerName).toLowerCase();
    for (var c = 1; c <= ws.columnCount; c++) {
      var v = plainVal(ws.getCell(1, c).value);
      if (v !== null && String(v).trim().toLowerCase() === target) return c;
    }
    return null;
  }
  function addColumn(ws, headerName) {
    for (var c = 1; c <= ws.columnCount; c++) {
      var v = plainVal(ws.getCell(1, c).value);
      if (v !== null && String(v).trim() === headerName) return c;
    }
    var newCol = ws.columnCount + 1;
    ws.getCell(1, newCol).value = headerName;
    return newCol;
  }
  // Case-insensitive reuse. Used for Tags only: unlike Ageing/Working1/… the
  // Tags column already exists in the export, and its casing is whatever
  // OneTrust emitted — an exact-match miss would append a duplicate column
  // and leave the real one stale.
  function findOrAddColumn(ws, headerName) {
    var existing = findColumn(ws, headerName);
    if (existing) return existing;
    var newCol = ws.columnCount + 1;
    ws.getCell(1, newCol).value = headerName;
    return newCol;
  }

  // True when the row holds nothing at all — interior blanks in the export
  // must not pick up a default tag just for occupying a row number.
  function isEmptyRow(ws, r) {
    var vals = ws.getRow(r).values;
    if (!vals) return true;
    for (var c = 1; c < vals.length; c++) {
      var v = plainVal(vals[c]);
      if (v !== null && v !== undefined && v !== '') return false;
    }
    return true;
  }

  async function convertDateColumn(ws, colNum) {
    if (!colNum) return;
    for (var r = 2; r <= ws.rowCount; r++) {
      var cell = ws.getCell(r, colNum);
      var v = cell.value;
      if (v === null || v === undefined) continue;
      if (v instanceof Date) { cell.numFmt = 'dd-mm-yyyy'; continue; }
      var parsed = parseDateString(String(v));
      if (parsed) { cell.value = parsed; cell.numFmt = 'dd-mm-yyyy'; }
      if ((r & 2047) === 0) await _yield();
    }
  }

  // ── tagging ─────────────────────────────────────────────────────────────
  // Pure: respondent cell value → tag. Exposed for tests.
  function tagFor(respondent) {
    var text = (respondent === null || respondent === undefined)
      ? '' : String(respondent).trim().toLowerCase();
    for (var i = 0; i < FINANCE_RESPONDENTS.length; i++) {
      if (text.indexOf(FINANCE_RESPONDENTS[i]) !== -1) return FINANCE_TAG;
    }
    return DEFAULT_TAG;
  }

  // Overwrites Tags for every populated row — the filter-and-fill, in code.
  // Runs before the calculated columns so a freshly created Tags column
  // lands ahead of Ageing/Working1/Working2/30sep26, matching the layout the
  // reports expect.
  async function applyTags(ws, warnings, progress) {
    var respondentCol = null, headerUsed = null;
    for (var i = 0; i < RESPONDENT_HEADERS.length && !respondentCol; i++) {
      respondentCol = findColumn(ws, RESPONDENT_HEADERS[i]);
      if (respondentCol) headerUsed = RESPONDENT_HEADERS[i];
    }
    // Hard failure, not a warning: a silently untagged workbook would flow
    // into the reports looking valid.
    if (!respondentCol) {
      throw new Error('Respondent column not found in "' + SHEET_ASSESSMENT +
        '". Expected one of: ' + RESPONDENT_HEADERS.join(', '));
    }

    var tagsCol = findOrAddColumn(ws, TAGS_HEADER);
    var counts = { finance: 0, cyber: 0, header: headerUsed };

    for (var r = 2; r <= ws.rowCount; r++) {
      if (isEmptyRow(ws, r)) continue;
      var tag = tagFor(plainVal(ws.getCell(r, respondentCol).value));
      ws.getCell(r, tagsCol).value = tag;
      if (tag === FINANCE_TAG) counts.finance++; else counts.cyber++;
      if ((r & 2047) === 0) {
        progress('Tagging respondents — row ' + r.toLocaleString() + '…');
        await _yield();
      }
    }

    if (counts.finance === 0) {
      warnings.push(SHEET_ASSESSMENT + ': no rows matched the Finance respondent (' +
        headerUsed + ' column) — everything tagged ' + DEFAULT_TAG);
    }
    return counts;
  }

  // ── process_risk_export ─────────────────────────────────────────────────
  async function processRiskExport(ws, warnings) {
    await convertDateColumn(ws, findColumn(ws, 'Date created'));
    var agingCol = addColumn(ws, 'Aging');
    var todaySerial = serialFromDate(todayUTC());
    for (var r = 2; r <= ws.rowCount; r++) {
      var n = serialOf(plainVal(ws.getCell(r, 14).value));  // column N — hardcoded like the Python
      var cellVal = { formula: 'NETWORKDAYS(N' + r + ',TODAY())' };
      if (n.kind !== 'text') cellVal.result = networkdaysSerial(n.serial, todaySerial);
      else warnings.push(SHEET_RISK + ' row ' + r + ': N is text — Aging left to Excel');
      ws.getCell(r, agingCol).value = cellVal;
      if ((r & 2047) === 0) await _yield();
    }
  }

  // ── process_assessment ──────────────────────────────────────────────────
  async function processAssessment(ws, warnings, onProgress) {
    var progress = onProgress || function () {};

    // Step 1 — tag every populated row before anything is appended.
    var tagCounts = await applyTags(ws, warnings, progress);

    // Step 2 — the existing workflow, unchanged.
    progress('Adding assessment formulas…');
    await convertDateColumn(ws, findColumn(ws, 'Date created'));
    await convertDateColumn(ws, findColumn(ws, 'Date submitted'));
    var ageingCol = addColumn(ws, 'Ageing');
    var working1Col = addColumn(ws, 'Working1');
    var working2Col = addColumn(ws, 'Working2');
    var sep26Col = addColumn(ws, '30sep26');
    var todaySerial = serialFromDate(todayUTC());
    var CUTOFF_2025 = serialFromDate(new Date(Date.UTC(2025, 11, 31))); // DATE(2025,12,31)

    for (var r = 2; r <= ws.rowCount; r++) {
      var cRaw = plainVal(ws.getCell(r, 3).value);   // column C
      var nRaw = plainVal(ws.getCell(r, 14).value);  // column N
      var oRaw = plainVal(ws.getCell(r, 15).value);  // column O
      var n = serialOf(nRaw), o = serialOf(oRaw);

      // Ageing
      var ageing = { formula: 'NETWORKDAYS(N' + r + ',TODAY())' };
      if (n.kind !== 'text') ageing.result = networkdaysSerial(n.serial, todaySerial);
      else warnings.push(SHEET_ASSESSMENT + ' row ' + r + ': N is text — Ageing left to Excel');
      ws.getCell(r, ageingCol).value = ageing;

      // Working1 — Excel semantics: text "=" is case-insensitive; text > number
      // ranks TRUE; blank compares as 0; blank cell equals "".
      var cTxt = (cRaw === null || cRaw === undefined) ? '' : String(cRaw).toLowerCase();
      var isDone = (cTxt === 'completed' || cTxt === 'under review');
      var oGT = (o.kind === 'num') ? o.serial > CUTOFF_2025
              : (o.kind === 'text');
      var oBlank = (o.kind === 'blank');
      ws.getCell(r, working1Col).value = {
        formula: 'IF(AND(OR(C' + r + '="Completed",C' + r + '="Under Review"),O' + r + '>DATE(2025,12,31)),' +
                 '"Completed in 2026",IF(O' + r + '="","Pending","Completed Before 2026"))',
        result: (isDone && oGT) ? 'Completed in 2026' : (oBlank ? 'Pending' : 'Completed Before 2026'),
      };

      // Working2
      var w2 = { formula: 'IF(DATE(2026,9,30)-N' + r + '>=365,"Beyond 1 Year Overdue","Current")' };
      if (n.kind !== 'text') w2.result = (SEP26_SERIAL - n.serial >= 365) ? 'Beyond 1 Year Overdue' : 'Current';
      else warnings.push(SHEET_ASSESSMENT + ' row ' + r + ': N is text — Working2 left to Excel');
      ws.getCell(r, working2Col).value = w2;

      // 30sep26
      var s26 = { formula: 'DATE(2026,9,30)-N' + r };
      if (n.kind !== 'text') s26.result = SEP26_SERIAL - n.serial;
      ws.getCell(r, sep26Col).value = s26;

      if ((r & 2047) === 0) await _yield();
    }

    return { tags: tagCounts };
  }

  async function loadFirstSheet(file) {
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    var ws = wb.worksheets[0];   // Python's wb.active
    if (!ws) throw new Error('"' + file.name + '" has no worksheets');
    return ws;
  }

  // ── build — main(). files = { risk: File, assessment: File, tprm: File } ──
  async function build(files, onProgress) {
    var progress = onProgress || function () {};
    var warnings = [];
    if (!files.risk) throw new Error('Risk export file not found.');
    if (!files.assessment) throw new Error('Assessment export file not found.');
    if (!files.tprm) throw new Error('TPRM export file not found.');

    var out = new ExcelJS.Workbook();
    out.calcProperties.fullCalcOnLoad = true;   // Excel refreshes TODAY() maths on open
    var wsTprm = out.addWorksheet(SHEET_TPRM);
    var wsAssessment = out.addWorksheet(SHEET_ASSESSMENT);
    var wsRisk = out.addWorksheet(SHEET_RISK);

    progress('Reading risk export…');
    var srcRisk = await loadFirstSheet(files.risk);
    progress('Reading TPRM export…');
    var srcTprm = await loadFirstSheet(files.tprm);
    progress('Reading assessment export…');
    var srcAssessment = await loadFirstSheet(files.assessment);

    await copySheet(srcRisk, wsRisk, progress, 'risk export');
    await copySheet(srcTprm, wsTprm, progress, 'TPRM export');
    await copySheet(srcAssessment, wsAssessment, progress, 'assessment export');

    progress('Adding risk ageing formulas…');
    await processRiskExport(wsRisk, warnings);
    progress('Tagging assessment respondents…');
    var assessmentResult = await processAssessment(wsAssessment, warnings, progress);

    progress('Writing workbook…');
    var bytes = await out.xlsx.writeBuffer();
    return {
      bytes: bytes,
      warnings: warnings,
      stats: {
        risk: { rows: wsRisk.rowCount, cols: wsRisk.columnCount },
        tprm: { rows: wsTprm.rowCount, cols: wsTprm.columnCount },
        assessment: {
          rows: wsAssessment.rowCount,
          cols: wsAssessment.columnCount,
          tags: assessmentResult.tags,
        },
      },
    };
  }

  return {
    classify: classify,
    build: build,
    OUTPUT_NAME: OUTPUT_NAME,
    // exposed for tests
    _internal: {
      parseDateString: parseDateString,
      networkdaysSerial: networkdaysSerial,
      serialFromDate: serialFromDate,
      serialOf: serialOf,
      plainVal: plainVal,
      tagFor: tagFor,
      FINANCE_RESPONDENTS: FINANCE_RESPONDENTS,
      RESPONDENT_HEADERS: RESPONDENT_HEADERS,
    },
  };
})();