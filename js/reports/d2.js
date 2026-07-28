// Port of "daigram 2 automation/automation.py" (id: d2).
// Reads "OneTrust Assessment", filters Tags=="cyber" (case-insensitive
// exact match) & Date created year==2026, maps Stage to Closed/Open, maps
// Organization to a display zone (falling back to the organization name
// itself when unmapped — unlike s2a/s3d's zone_map, "Europe" stays
// "Europe" here, not "EUR"), pivots Org Display x Final Stage (count, no
// margins), sorts to a fixed zone order, appends a Grand Total row, and
// computes a small KPI panel (3 static values + 1 computed Q2 rate).
// Writes "output file D2.xlsx" as a single hand-built "Dashboard" sheet
// (filter area, then the pivot table, then the KPI table below it) — the
// exact row offsets matter because the existing preview pipeline re-treats
// row 0 as headers and sparsity-trims the rest, same as s2a.
window.Reports.d2 = async function d2(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers.map(function (h) { return String(h).trim(); });

  // Two columns can carry the same header. An export that gained a replacement
  // "Tags" keeps the original beside it, usually empty. readSheet dedupes the
  // row KEYS (Tags, Tags_1) but this loop re-keyed by the raw header, so both
  // positions wrote out['Tags'] = row['Tags'] and the populated second column
  // was unreachable — tagOnly went to 0 under any matching rule, with no error.
  // Take the first position that actually holds a value for this row.
  //
  // Not a matching-rule change: on a sheet with one column per name this picks
  // that same column, so literal parity with the script is untouched. Duplicates
  // are counted and named in the run summary rather than quietly worked around.
  // sheet_to_json dedupes repeated keys by occurrence: the 2nd "Tags" becomes
  // "Tags_1", the 3rd "Tags_2". Rebuild that mapping so every position is
  // addressable, then keep the first non-blank one per name.
  var seenHeader = {}, rowKeys = [], dupHeaders = [];
  sheet.headers.forEach(function (h) {
    var n = seenHeader[h] || 0;
    seenHeader[h] = n + 1;
    rowKeys.push(n === 0 ? h : h + '_' + n);
  });
  Object.keys(seenHeader).forEach(function (h) {
    if (seenHeader[h] > 1) dupHeaders.push(String(h).trim() + ' x' + seenHeader[h]);
  });

  function blankCell(v) { return v === undefined || v === null || v === ''; }
  var rows = sheet.rows.map(function (row) {
    var out = {};
    rowKeys.forEach(function (key, i) {
      var name = rawHeaders[i];
      var v = row[key];
      if (v === undefined) v = row[sheet.headers[i]];
      if (!(name in out) || (blankCell(out[name]) && !blankCell(v))) out[name] = v;
    });
    return out;
  });
  // ── Column resolution ────────────────────────────────────────────────────
  //
  // The reference workbook uses ID / Organization / Stage / Date created / Tags.
  // A real supplier-assessment export of the same data uses snake_case and
  // different words for the same things (assessment_status for Stage,
  // request_date for Date created, zone_assessing for Organization), so the
  // report threw "Missing required columns" on a file it could otherwise read
  // perfectly. Each field therefore accepts a small list of known synonyms.
  //
  // The canonical name always wins, so a workbook shaped like the reference one
  // resolves exactly as before and parity is untouched. Which header was
  // actually used is printed in COLUMNS USED — a silent alias would be worse
  // than the error it replaces.
  var ALIASES = {
    'ID': ['ID', 'id', 'code', 'assessment_id'],
    'Organization': ['Organization', 'Organisation', 'zone_assessing', 'supplier_zone', 'zone'],
    'Stage': ['Stage', 'assessment_status', 'status'],
    'Date created': ['Date created', 'request_date', 'date_created', 'created_date', 'created'],
    // Tags is the risk-category dimension. In this export the category lives in
    // one of several *_category columns and only one of them holds "cyber", so
    // the choice is made by looking rather than by guessing — see below.
    'Tags': ['Tags', 'main_category', 'category', 'gpo_category', 'fin_category'],
  };
  var CANON = Object.keys(ALIASES);

  function headerIndex(name) {
    for (var i = 0; i < rawHeaders.length; i++) {
      if (rawHeaders[i].toLowerCase() === String(name).toLowerCase()) return i;
    }
    return -1;
  }
  // Every alias that is actually present, canonical first.
  function candidates(canon) {
    var out = [];
    ALIASES[canon].forEach(function (a) {
      var i = headerIndex(a);
      if (i !== -1 && out.indexOf(rawHeaders[i]) === -1) out.push(rawHeaders[i]);
    });
    return out;
  }

  var resolved = {}, resolvedHow = {};
  CANON.forEach(function (canon) {
    var cands = candidates(canon);
    if (!cands.length) return;
    var pick = cands[0];
    var how = pick === canon ? '' : 'matched by alias';
    // For the category column, prefer whichever candidate actually contains a
    // cyber value. Picking the first present one would silently measure
    // main_category when the split lives in category, and the two disagree.
    if (canon === 'Tags' && cands.length > 1) {
      var withCyber = cands.filter(function (h) {
        return sheet.rows.some(function (row) { return /cyber/i.test(String(row[h] === undefined ? '' : row[h])); });
      });
      if (withCyber.length) {
        pick = withCyber[0];
        how = (pick === canon ? '' : 'matched by alias; ') + 'chosen over ' +
          cands.filter(function (h) { return h !== pick; }).join(', ') + ' because it contains a "cyber" value';
      } else {
        how = (pick === canon ? '' : 'matched by alias; ') + 'no candidate contains "cyber" (also looked at ' +
          cands.filter(function (h) { return h !== pick; }).join(', ') + ')';
      }
    }
    resolved[canon] = pick;
    resolvedHow[canon] = how;
  });

  // Report every missing column at once — naming them one run at a time turns
  // a single malformed export into five round trips.
  var missing = CANON.filter(function (c) { return !resolved[c]; });
  if (missing.length) {
    throw new Error('Missing required columns: ' + missing.map(function (c) {
      return c + ' (also accepts ' + ALIASES[c].slice(1).join(', ') + ')';
    }).join('; ') + '.  Headers found: ' + rawHeaders.join(', '));
  }
  // Re-key each row under the canonical names the rest of the report uses.
  rows.forEach(function (r) {
    CANON.forEach(function (canon) {
      if (resolved[canon] !== canon) r[canon] = r[resolved[canon]];
    });
  });

  // df[c] = df[c].fillna("").astype(str).str.strip()
  //
  // Deliberately no more than that. An earlier version also stripped zero-width
  // characters and folded NBSP, which made this report accept rows the script
  // rejects; the requirement is literal parity, so the cleaning matches pandas
  // exactly. E.isBlank is NOT used here — it treats the literal string "nan" as
  // blank, which fillna does not.
  //
  // One residual difference, stated rather than hidden: JS trim() also strips
  // U+FEFF (BOM), while Python's str.strip() does not — BOM is a format
  // character, not whitespace. A BOM-prefixed "Cyber" cell therefore matches
  // here and would not in the script.
  function pyStrip(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' && Number.isNaN(v)) return '';
    return String(v).trim();
  }
  rows.forEach(function (r) {
    ['Organization', 'Stage', 'Tags'].forEach(function (c) { r[c] = pyStrip(r[c]); });
  });

  // ── Matching: tolerant by default, but it tells you when it was tolerant ──
  //
  // The script's rules are exact: a whole-cell (?i)^cyber$ on Tags and an
  // exact-case {Completed, Under review} on Stage. They are right for the
  // reference workbook and wrong for a real OneTrust export, where Tags is a
  // delimited list ("Cyber, Third Party") and Stage casing drifts ("Under
  // Review"). Held strictly the report publishes a confident 0%; relaxed
  // silently it stops being a reproduction of the reference output.
  //
  // Both fallbacks are semantically safe — they can only admit rows a human
  // would agree are matches. Per-tag equality still rejects "Cyber Security"
  // (a different tag, not a delimiter away); case-folding still rejects any
  // stage that is not one of the two closed ones.
  var TAG_STRICT = function (tags) { return /^cyber$/i.test(tags); };
  var TAG_LOOSE = function (tags) {
    return String(tags).split(/[,;|\/\n]+/).some(function (t) { return /^cyber$/i.test(t.trim()); });
  };
  var STAGE_STRICT = function (s) { return s === 'Completed' || s === 'Under review'; };
  var STAGE_LOOSE = function (s) { return /^(completed|under review)$/i.test(String(s).trim()); };

  var yearOf = function (r) { return E.excelYear(r['Date created']); };

  // Match with the looser rule, then check whether it actually admitted anything
  // the script's own rule would have rejected. When it did not — every file the
  // script handles — the result IS the script's result and is reported as such.
  // When it did, the numbers are real but no longer a like-for-like reproduction
  // of the reference output, so the run says which rule produced them.
  //
  // Checking "did the looser rule admit extra rows" rather than "did the strict
  // rule find nothing" is what catches the partial case: an export where
  // "Completed" still matches exactly but "Under Review" does not would
  // otherwise keep half its closed rows and publish a quietly wrong figure.
  var tagHits = rows.filter(function (r) { return TAG_LOOSE(r.Tags); });
  var tagLoosened = tagHits.some(function (r) { return !TAG_STRICT(r.Tags); });
  var cyberHits = tagHits.length;          // the honest "Tags = Cyber" count

  // A real export tags by zone and year ("MAZ 2024", "GHQ 2022,SRM 2023") and
  // carries no risk-category tag at all. The Cyber filter then selects nothing
  // and the whole report is 0% — a filter that matches none of 1249 rows is not
  // narrowing the population, it is deleting it. Drop it and measure every
  // assessment instead, exactly as the date filter is dropped when no date
  // parses. This is announced everywhere it could mislead: the sheet's own
  // "Tags" filter cell says "all", and the run summary leads with it.
  //
  // But before falling back to "everything", read what the tags DO encode. This
  // export tags each row with its assessment cycle — "APAC 2026", "GHQ2026",
  // "SAZ Re-assessment 2026" — so the tag column already names a population,
  // just not a risk category. The cycle is the intended cohort and is NOT the
  // same set as "Date created in that year": an assessment raised in December
  // 2025 belongs to the 2026 cycle, and one created in 2026 may be closing out
  // 2025 work. On the real file the two differ by 32 rows (45 vs 77), which is
  // the whole gap between the expected figure and the published one.
  var YEAR_TARGET = 2026;
  var tagFilterDropped = false, tagCohort = false;
  if (!tagHits.length && rows.length) {
    var cohort = rows.filter(function (r) { return String(r.Tags).indexOf(String(YEAR_TARGET)) !== -1; });
    if (cohort.length) {
      tagCohort = true;
      tagHits = cohort;
    } else {
      tagFilterDropped = true;
      tagHits = rows.slice();
    }
  }
  var tagOnly = tagHits.length;

  // ── The year, same tolerant-but-announced treatment ──
  //
  // The script hardcodes 2026. That is the last filter here that could still
  // silently empty the set: a file exported for a different year, or one whose
  // "Date created" column no longer parses as a date, keeps every Cyber row and
  // then discards all of them, publishing 0% with no explanation.
  //
  // So: use 2026 when it matches anything, otherwise fall back to whichever year
  // the Cyber rows actually carry, and if none of them carry a readable date at
  // all, drop the date filter rather than the data. Every step says what it did,
  // and the year in use is written into the sheet's own filter cell so the
  // Dashboard never claims to be a 2026 report when it is not.
  var byYear = {};
  tagHits.forEach(function (r) {
    var y = yearOf(r);
    if (y !== null) byYear[y] = (byYear[y] || 0) + 1;
  });

  var yearUsed = YEAR_TARGET, yearLoosened = false, dateFilterDropped = false;
  var filtered;
  // When the cohort came from the tag year, that IS the period selection —
  // intersecting it with "Date created in the same year" would apply the period
  // twice and quietly discard the rows the cohort exists to include.
  if (tagCohort) {
    dateFilterDropped = true;
    filtered = tagHits.slice();
  } else {
    filtered = tagHits.filter(function (r) { return yearOf(r) === YEAR_TARGET; });
  }
  if (!filtered.length) {
    // Most-populated year wins; ties break to the later year so a rolling export
    // reports the current period rather than the one it is replacing.
    var years = Object.keys(byYear).map(Number).sort(function (a, b) {
      return byYear[b] - byYear[a] || b - a;
    });
    if (years.length) {
      yearUsed = years[0];
      yearLoosened = true;
      filtered = tagHits.filter(function (r) { return yearOf(r) === yearUsed; });
    } else if (tagHits.length) {
      dateFilterDropped = true;
      filtered = tagHits.slice();
    }
  }
  var yearOnly = rows.filter(function (r) { return yearOf(r) === yearUsed; }).length;

  var stageLoosened = filtered.some(function (r) { return STAGE_LOOSE(r.Stage) && !STAGE_STRICT(r.Stage); });
  filtered.forEach(function (r) { r['Final Stage'] = STAGE_LOOSE(r.Stage) ? 'Closed' : 'Open'; });

  var TAG_RULE = tagFilterDropped
    ? 'none — no tag in this file is "cyber"   (DROPPED — every assessment is counted)'
    : (tagCohort
      ? 'tag contains "' + YEAR_TARGET + '"   (COHORT — no tag is "cyber", so the ' + YEAR_TARGET +
        ' assessment cycle is used as the population)'
      : (tagLoosened
        ? 'per tag after splitting on , ; | / and newlines   (LOOSENED — the script matches the whole cell)'
        : 'whole cell = "cyber", case-insensitive   (the script\'s rule)'));
  var STAGE_RULE = stageLoosened
    ? 'Completed / Under review, case-insensitive   (LOOSENED — the script is exact-case)'
    : 'exact case {Completed, Under review}   (the script\'s rule)';
  var DATE_RULE = tagCohort
    ? 'not applied — the tag cohort already selects the ' + YEAR_TARGET + ' period'
    : (dateFilterDropped
    ? 'none — no Cyber row has a readable "Date created"   (DROPPED — the script keeps only 2026)'
    : (yearLoosened
      ? 'Date created in ' + yearUsed + '   (LOOSENED — the script hardcodes 2026, which matched nothing)'
      : 'Date created in 2026   (the script\'s rule)'));
  // Both charts are titled "<year> Assessment"; a report that fell back to
  // another year must not still be captioned 2026.
  var chartYear = dateFilterDropped ? 'Undated' : String(yearUsed);

  var ORG_MAP = {
    'Africa': 'AFR', 'APAC': 'APAC', 'BEES': 'GRO', 'BEES | FINTECH': 'GRO',
    'Europe': 'Europe', 'GHQ': 'GHQ', 'South America Zone': 'SAZ',
    'North America Zone': 'NAZ', 'Middle America Zone': 'MAZ',
  };
  filtered.forEach(function (r) { r['Org Display'] = E.mapZone(r.Organization, ORG_MAP); });

  // values="ID", aggfunc="count" — pandas counts non-null IDs, NOT rows, so a
  // Cyber/2026 row with a blank ID never reaches the pivot. Q2 is defined as
  // Closed / Grand Total — read off the pivot's own total row — so it is
  // computed from these counts below, not from len(filtered). The script
  // divides by len(filtered) while labelling the result "Grand Total"; on a
  // sheet with a blank ID the two disagree and the published percentage
  // contradicts the table printed directly above it.
  var pivotRaw = E.pivotCount(filtered, function (r) { return r['Org Display']; }, function (r) { return r['Final Stage']; }, { margins: false, valueFn: function (r) { return r.ID; } });
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

  // Python's round() is half-to-EVEN, so round(0.125, 2) is 0.12 where JS's
  // Math.round would give 0.13. Small filtered sets land on those exact halves
  // (1/8, 3/8, ...), and the KPI is published, so match the script's rounding.
  function pyRound2(x) {
    var scaled = x * 100, floor = Math.floor(scaled), diff = scaled - floor;
    var n = diff > 0.5 ? floor + 1 : (diff < 0.5 ? floor : (floor % 2 === 0 ? floor : floor + 1));
    return n / 100;
  }

  // Q2 '26 = Closed / Grand Total, both taken from the pivot's total row so the
  // percentage always reconciles against the table above it. keptRows — how many
  // rows passed the filter — can be higher when IDs are blank; that gap is
  // reported in the run summary rather than silently moving the denominator.
  var keptRows = filtered.length;
  var closedTotal = totalClosed;
  var recordTotal = totalGrand;
  var q2_26 = recordTotal ? pyRound2(closedTotal / recordTotal) : 0;
  // kpi_data — the Q2 remark is the script's empty string. A 0% here is either
  // a real measurement or the filter finding nothing, and the two look
  // identical; that distinction is reported in the run summary at the end
  // instead, which is browser-only and cannot alter the workbook.
  var KPI = [
    ["Baseline '25", 0.60, 'static'],
    ["Q1 '26", 0.32, 'static'],
    ["Q2 '26", q2_26, ''],
    ["Target '26", 0.65, 'static'],
  ];

  function setCell(grid, row1, col1, value) {
    var r = row1 - 1, c = col1 - 1;
    while (grid.length <= r) grid.push([]);
    while (grid[r].length <= c) grid[r].push(null);
    grid[r][c] = value;
  }

  var grid = [];
  // Both filter cells state what was actually applied, so the Dashboard cannot
  // claim to be a 2026 Cyber report while showing every assessment in the file.
  setCell(grid, 1, 1, 'Tags'); setCell(grid, 1, 2,
    tagFilterDropped ? 'all' : (tagCohort ? '*' + YEAR_TARGET + '*' : 'Cyber'));
  // The filter cell states the year actually applied, so the Dashboard cannot
  // claim to be a 2026 report when the rows behind it are not.
  setCell(grid, 2, 1, 'Date created'); setCell(grid, 2, 2,
    tagCohort ? 'n/a' : (dateFilterDropped ? 'all' : String(yearUsed)));

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
    setCell(grid, rowNum, 2, k[1]);
    setCell(grid, rowNum, 3, k[2]);
  });
  var noteRow = kpiStartRow + KPI.length + 2;
  setCell(grid, noteRow, 1, 'Q2 formula');
  setCell(grid, noteRow, 2, 'Closed / Grand Total = ' + closedTotal + ' / ' + recordTotal);

  // ── Run summary (returned as stdout, NOT written to the sheet) ──
  // Q2 is a ratio over the filtered set, so when the filter matches nothing it
  // reports a confident "0%" indistinguishable from a real measurement. These
  // lines say how many rows each half of the filter accepted and what the
  // rejected values look like, so one client run identifies the cause.
  //
  // This used to be a block of cells below the note row. It moved here so the
  // Dashboard sheet is a literal reproduction of the openpyxl output — stdout
  // has no counterpart in the script and cannot change the workbook. The shell
  // renders it in the "Script log" panel; s4 already returns stdout the same way.
  // No truncation. "+53 more" is exactly where the answer hides — the value that
  // explained a 0% has been in that tail twice now. Every distinct value is
  // listed, most frequent first, wrapped at `indent` so a 60-value column stays
  // readable instead of running off as one line.
  function distinct(values) {
    var seen = [], counts = {};
    values.forEach(function (v) {
      var k = v === '' ? '(blank)' : String(v);
      if (!counts[k]) { counts[k] = 0; seen.push(k); }
      counts[k]++;
    });
    if (!seen.length) return '(none)';
    seen.sort(function (a, b) { return counts[b] - counts[a]; });
    return seen.map(function (k) { return k + ' x' + counts[k]; }).join('  ·  ') +
      '   (' + seen.length + ' distinct)';
  }

  var unparsedDates = 0;
  var yearsSeen = rows.map(function (r) {
    var y = E.excelYear(r['Date created']);
    if (y === null) { unparsedDates++; return 'unreadable'; }
    return String(y);
  });

  var q2Reason = !tagOnly ? 'the sheet has no data rows — the 0% is not a measurement'
    : (!keptRows ? 'Cyber rows exist but none survived the date filter — the 0% is not a measurement'
    : (!recordTotal ? keptRows + ' row(s) passed the filter but every ID is blank, so the pivot counted nothing — the 0% is not a measurement'
    : (!closedTotal ? 'rows passed the filter, none are at a Closed stage' : '')));
  // ── Debugging aids: where each column came from, and why rows were dropped ──
  //
  // The funnel says HOW MANY rows each rule rejected. What it never said is
  // WHICH rule rejected a given row, or what that row actually looked like —
  // so a wrong answer still meant asking the user to open the workbook. The
  // column map catches a misnamed or duplicated header, the drop tally splits
  // the rejections into mutually exclusive buckets, and the sample rows show
  // one real example of each outcome with the reason attached.
  function pad(s, n) {
    s = String(s);
    while (s.length < n) s += ' ';
    return s;
  }
  function colLetter(i) {
    var s = '', n = i + 1;
    while (n > 0) { var rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function showDate(v) {
    var d = E.excelDateInfo(v);
    if (d) return d.year + '-' + ('0' + d.month).slice(-2) + '-' + ('0' + d.day).slice(-2);
    var s = String(v === null || v === undefined ? '' : v).trim();
    return s ? s.slice(0, 18) + ' (unreadable)' : '(blank)';
  }
  function blankish(v) {
    return (v === null || v === undefined || String(v).trim() === '') ? '(blank)' : String(v);
  }
  // Row number as the user sees it in Excel: +2 for the header row and 1-basing.
  function sampleLine(r, i, why) {
    return '    row ' + pad(i + 2, 6) + pad('ID ' + blankish(r.ID), 12) +
      pad(blankish(r.Organization), 20) + pad(blankish(r.Stage), 14) +
      pad(showDate(r['Date created']), 14) + 'Tags ' + JSON.stringify(String(r.Tags)) +
      (why ? '   <- ' + why : '');
  }

  var dropTagOnly = 0, dropDateOnly = 0, dropBoth = 0;
  var keptSamples = [], dropSamples = {};
  rows.forEach(function (r, i) {
    // In cohort mode the tag test is cohort membership, not the cyber match —
    // using TAG_LOOSE here counted every row as tag-rejected and the buckets
    // stopped summing to rows read.
    var tagPass = tagFilterDropped ? true
      : (tagCohort ? String(r.Tags).indexOf(String(YEAR_TARGET)) !== -1 : TAG_LOOSE(r.Tags));
    var datePass = dateFilterDropped || yearOf(r) === yearUsed;
    var sawYear = yearOf(r);
    var dateWhy = 'date is ' + (sawYear === null ? 'unreadable' : sawYear) + ', not ' + yearUsed;
    var tagWhy = tagCohort ? 'tag does not name the ' + YEAR_TARGET + ' cycle' : 'tag is not cyber';
    if (tagPass && datePass) {
      if (keptSamples.length < 3) keptSamples.push(sampleLine(r, i, ''));
      return;
    }
    var key;
    if (!tagPass && !datePass) { dropBoth++; key = tagWhy + ', and ' + dateWhy; }
    else if (!tagPass) { dropTagOnly++; key = tagWhy; }
    else { dropDateOnly++; key = dateWhy; }
    if (!dropSamples[key]) dropSamples[key] = sampleLine(r, i, key);
  });

  var columnMap = CANON.map(function (name) {
    var actual = resolved[name];
    var at = [];
    rawHeaders.forEach(function (h, i) { if (h === actual) at.push('col ' + colLetter(i)); });
    var notes = [];
    if (at.length > 1) notes.push(at.length + ' columns share this header — first non-blank per row is used');
    if (resolvedHow[name]) notes.push(resolvedHow[name]);
    return '  ' + pad(name, 16) + pad(at.join(' + '), 16) +
      (actual === name ? '' : '"' + actual + '"') +
      (notes.length ? '   (' + notes.join('; ') + ')' : '');
  });

  var RULE = '─────────────────────────────────────────────────────────────────';
  function section(title) {
    var s = '── ' + title + ' ';
    while (s.length < RULE.length) s += '─';
    return s;
  }
  var scope = tagCohort
    ? 'the ' + YEAR_TARGET + ' assessment cycle (by tag, not by Date created)'
    : (tagFilterDropped ? 'EVERY assessment' : 'Cyber assessments') +
      (dateFilterDropped ? ', any date' : ' dated ' + yearUsed);

  var lines = [
    'SUCCESS',
    'output file D2.xlsx',
    '',
    RULE,
    ' RESULT   ' + "Q2 '26 = " + Math.round(q2_26 * 100) + '%' +
      '   (Closed ' + closedTotal + ' / Grand Total ' + recordTotal + ')',
    ' SCOPE    ' + scope,
    RULE,
    '',
    'Input file: ' + ((wb && wb.__fileName) || 'workbook.xlsx'),
    'Sheet: OneTrust Assessment   ·   ' + rows.length + ' data rows   ·   ' + rawHeaders.length + ' columns',
  ];
  if (dupHeaders.length) {
    lines.push('Duplicate columns: ' + dupHeaders.join(', ') + '   (first non-blank value per row is used)');
  }
  lines.push(
    '',
    section('COLUMNS USED')
  );
  columnMap.forEach(function (l) { lines.push(l); });
  // The unmapped columns matter as much as the mapped ones: when a count comes
  // out wrong, the answer is usually in a column this report never looks at, and
  // not listing them cost a round trip.
  var usedHeaders = CANON.map(function (c) { return resolved[c]; });
  var otherCols = [];
  rawHeaders.forEach(function (h, i) {
    if (h && usedHeaders.indexOf(h) === -1) otherCols.push(colLetter(i) + ' "' + h + '"');
  });
  if (otherCols.length) {
    lines.push('  ' + pad('not used', 16) + otherCols.join('  ·  '));
  }
  lines.push(
    '',
    section('RULES APPLIED'),
    'Tags rule:  ' + TAG_RULE,
    'Stage rule: ' + STAGE_RULE,
    'Date rule:  ' + DATE_RULE,
    '',
    section('FUNNEL'),
    'Filter (' + (tagFilterDropped ? 'no tag filter' : 'Tags') + ', and ' +
      (dateFilterDropped ? 'no date filter' : 'Date created year ' + yearUsed) + '):',
    '  Rows read:            ' + rows.length,
    '  Tags = Cyber:         ' + cyberHits + (tagFilterDropped ? '   <- filter dropped, all ' + tagOnly + ' rows used' : ''),
  );
  if (tagCohort) {
    lines.push('  ' + pad('Tags name ' + YEAR_TARGET + ':', 22) + tagOnly + '   <- used as the population instead');
  }
  lines.push(
    '  Date created in ' + yearUsed + ': ' + yearOnly + (tagCohort ? '   (not used — see Date rule)' : ''),
    '  Kept (both):          ' + keptRows,
    '',
    'Rows dropped, by which rule rejected them:',
    '  tag rule only:        ' + dropTagOnly + (tagFilterDropped ? '   (tag rule was not applied)' : ''),
    '  date rule only:       ' + dropDateOnly + (dateFilterDropped ? '   (date rule was not applied)' : ''),
    '  both rules:           ' + dropBoth,
    '',
    'Stage of the kept rows:',
    '  Closed: ' + closedTotal,
    '  Open:   ' + totalOpen,
    '',
    section('PIVOT')
  );
  pivotDisplay.forEach(function (r) {
    lines.push('  ' + pad(r.label, 16) + 'open ' + pad(r.open, 6) + 'closed ' + pad(r.closed, 6) + 'total ' + r.grandTotal);
  });
  lines.push(
    '',
    "Q2 '26 = Closed / Grand Total = " + closedTotal + ' / ' + recordTotal + ' = ' + Math.round(q2_26 * 100) + '%',
    ''
  );
  // Warnings are collected rather than pushed inline so they sit together under
  // one heading — scattered through the log they read as commentary, in a block
  // they read as a checklist. The "  ^ " prefix is what makes the shell open the
  // Script log by itself, so it stays on every one of them.
  var warn = [];
  // The pivot counts IDs, so a blank ID drops a row from the Grand Total while
  // it still passed the filter. Say so, otherwise the two numbers above look
  // like an arithmetic error.
  if (keptRows !== recordTotal) {
    warn.push((keptRows - recordTotal) + ' kept row(s) have a blank ID and are not counted in the' +
      ' Grand Total (the pivot counts IDs, not rows).');
  }
  if (q2Reason) warn.push(q2Reason);
  // A number produced by anything other than the script's own rule is still
  // correct, but it is no longer a like-for-like reproduction of the reference
  // output — so it is flagged, not slipped through.
  if (tagLoosened || stageLoosened || yearLoosened || dateFilterDropped || tagFilterDropped || tagCohort) {
    warn.push('The script\'s exact rule matched nothing here, so a looser one was used' +
      ' (see the rules above). The numbers are real; they will not match a' +
      ' run of the Python script on this same file.');
  }
  if (yearLoosened) {
    warn.push('No Cyber row is dated 2026, so this report covers ' + yearUsed + ' instead.' +
      ' Years present on the Cyber rows: ' +
      Object.keys(byYear).map(Number).sort().map(function (y) { return y + ' x' + byYear[y]; }).join('  ·  ') +
      '. The "Date created" filter cell on the sheet says ' + yearUsed + '.');
  }
  if (tagCohort) {
    warn.push('No tag in this file is "cyber", but the tags DO name assessment cycles, so the ' +
      YEAR_TARGET + ' cycle is the population: ' + tagOnly + ' row(s) whose tag contains "' + YEAR_TARGET +
      '". This is NOT the same set as "Date created in ' + YEAR_TARGET + '" (' + yearOnly + ' rows) — a' +
      ' cycle can hold work raised the previous year. See ALTERNATIVE POPULATIONS below for all three' +
      ' readings and tell me which one Q2 should measure.');
  }
  if (dateFilterDropped && !tagCohort) {
    warn.push('Not one Cyber row has a readable "Date created", so the date filter was dropped' +
      ' and all ' + tagOnly + ' of them are counted. Check the column\'s format — a date stored as' +
      ' text in an unusual shape is the usual cause.');
  }
  // Tags arrive as delimited lists ("GHQ 2022,SRM 2023"), so listing distinct
  // CELLS buries the individual tags under dozens of unique combinations — a
  // real export showed 59 distinct cells and truncated to "+51 more", hiding the
  // only fact that mattered (whether any cyber tag exists at all). Split them
  // the way the matcher does and count the tags themselves.
  var individualTags = [];
  rows.forEach(function (r) {
    String(r.Tags).split(/[,;|\/\n]+/).forEach(function (t) {
      t = t.trim();
      if (t) individualTags.push(t);
    });
  });
  // When nothing matched, say whether the file contains anything cyber-shaped.
  // "no tag equals cyber" and "the word cyber appears nowhere" are different
  // problems: the first is a naming mismatch, the second means this export
  // simply is not the one this report was built for.
  if (tagFilterDropped) {
    var near = [];
    individualTags.forEach(function (t) {
      if (/cyber/i.test(t) && near.indexOf(t) === -1) near.push(t);
    });
    // The scope warning goes on BOTH branches. A near miss is the likelier case
    // and the more dangerous one — "Cyber Security" existing makes the number
    // look Cyber-shaped when it is not.
    warn.push('No tag in this file is "cyber", so the tag filter was dropped and all ' +
      tagOnly + ' rows are counted. THIS FIGURE COVERS EVERY ASSESSMENT, NOT JUST CYBER ONES.');
    warn.push(near.length
      ? 'These tags DO contain the word: ' + near.join('  ·  ') +
        '. If one of them is the population you want, say which and it becomes the filter.'
      : 'The word "cyber" appears in no tag at all. This export is tagged by zone and year' +
        ' rather than by risk category, so tell me which column carries the Cyber/non-Cyber' +
        ' split and the filter can point at it.');
  }
  lines.push(warn.length
    ? section('WARNINGS (' + warn.length + ')')
    : section('WARNINGS') + '\n  none — every rule ran exactly as the script defines it.');
  warn.forEach(function (w, i) { lines.push('  ^ ' + (i + 1) + '. ' + w); });
  // One real row per outcome, with the reason attached. A count says a rule
  // rejected 1172 rows; this says what one of them looked like, which is the
  // difference between "the filter is wrong" and "the data is wrong".
  // Three defensible readings of "the 2026 population" exist once the tag column
  // stops being a risk category, and they give different percentages. Printing
  // all three with their arithmetic means the right one can be picked by looking
  // instead of by another round trip.
  lines.push('', section('ALTERNATIVE POPULATIONS'));
  lines.push('  The same file read three ways — confirm which one Q2 should measure:');
  var byTagYear = rows.filter(function (r) { return String(r.Tags).indexOf(String(YEAR_TARGET)) !== -1; });
  var byDateYear = rows.filter(function (r) { return yearOf(r) === YEAR_TARGET; });
  var both = byTagYear.filter(function (r) { return yearOf(r) === YEAR_TARGET; });
  [
    ['tag contains "' + YEAR_TARGET + '"', byTagYear, tagCohort],
    ['Date created in ' + YEAR_TARGET, byDateYear, !tagCohort && !dateFilterDropped],
    ['both of the above', both, false],
  ].forEach(function (opt) {
    var set = opt[1];
    var cl = set.filter(function (r) { return STAGE_LOOSE(r.Stage); }).length;
    var pct = set.length ? Math.round(pyRound2(cl / set.length) * 100) : 0;
    lines.push('  ' + pad(opt[0], 26) + pad(set.length + ' rows', 12) +
      pad('closed ' + cl, 13) + '-> ' + pad(pct + '%', 6) + (opt[2] ? '   <- CHOSEN' : ''));
  });

  // Same idea one level down: once the population is settled, "Closed" is the
  // remaining free variable. On the real file no subset of the Stage values can
  // produce the expected closed count, which means closure is recorded
  // somewhere other than Stage — so every plausible reading is counted here
  // rather than guessed at one per round trip.
  lines.push('', section('CLOSURE READINGS'));
  lines.push('  "Closed" counted several ways over the same ' + filtered.length + ' rows:');
  var closedDateCol = null;
  ['Date closed', 'date_closed', 'closed_date', 'Date Closed', 'completion_date', 'date_completed', 'closed_on']
    .forEach(function (a) {
      if (closedDateCol) return;
      var i = headerIndex(a);
      if (i !== -1) closedDateCol = rawHeaders[i];
    });
  var readings = [
    ['Stage is Completed or Under review', function (r) { return STAGE_LOOSE(r.Stage); }, true],
    ['Stage is Completed only', function (r) { return /^completed$/i.test(String(r.Stage).trim()); }, false],
  ];
  if (closedDateCol) {
    readings.push(['"' + closedDateCol + '" is filled', function (r) { return !blankCell(r[closedDateCol]) && String(r[closedDateCol]).trim() !== ''; }, false]);
    readings.push(['Stage closed OR "' + closedDateCol + '" filled', function (r) {
      return STAGE_LOOSE(r.Stage) || (!blankCell(r[closedDateCol]) && String(r[closedDateCol]).trim() !== '');
    }, false]);
  }
  readings.forEach(function (rd) {
    var n = filtered.filter(rd[1]).length;
    var pct = recordTotal ? Math.round(pyRound2(n / recordTotal) * 100) : 0;
    lines.push('  ' + pad(rd[0], 40) + pad('closed ' + n, 13) + '-> ' + pad(pct + '%', 6) + (rd[2] ? '   <- CHOSEN' : ''));
  });
  if (!closedDateCol) {
    lines.push('  No "Date closed"-style column found. If closure is recorded in one of the' +
      ' unused columns listed above, name it and it becomes the rule.');
  }

  lines.push('', section('SAMPLE ROWS'));
  if (keptSamples.length) {
    lines.push('  kept:');
    keptSamples.forEach(function (l) { lines.push(l); });
  } else {
    lines.push('  kept:   (none)');
  }
  var dropKeys = Object.keys(dropSamples);
  if (dropKeys.length) {
    lines.push('  dropped — one example per reason:');
    dropKeys.forEach(function (k) { lines.push(dropSamples[k]); });
  } else {
    lines.push('  dropped: (none — every row was kept)');
  }

  lines.push(
    '',
    section('VALUES SEEN'),
    '  Tags:         ' + distinct(individualTags) + '   (individual tags, split on , ; | /)',
    '  Tag cells:    ' + distinct(rows.map(function (r) { return r.Tags; })) + '   (as written)',
    '  Stage:        ' + distinct((keptRows ? filtered : rows).map(function (r) { return r.Stage; })) +
      (keptRows ? '   (kept rows)' : '   (whole sheet — nothing passed the filter)'),
    '  Date created: ' + distinct(yearsSeen) +
      (unparsedDates ? '   — ' + unparsedDates + ' cell(s) could not be read as a date' : '')
  );
  var stdout = lines.join('\n');


  var files = [{ name: 'output file D2.xlsx', sheets: [{ name: 'Dashboard', grid: grid }] }];

  // Both embedded charts reproduce the original matplotlib styles instead
  // of the light in-page preview theme. Both are embedded into the file
  // (matching today's 2-chart output), even though only the KPI chart ever
  // surfaces in the preview — confirmed with the user.
  var images = {};

  // Chart 1: org-level stacked bar — black bg, Open (cyan) bottom + Closed
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
    title: { text: '<b>' + chartYear + ' Assessment<br>(' + closedTotal + '/' + recordTotal + ')</b>', font: { color: '#ffffff', size: 16 } },
    xaxis: { tickangle: -45, tickfont: { color: '#ffffff' }, showline: true, linecolor: '#ffffff' },
    yaxis: { visible: false },
    legend: { orientation: 'h', x: 0.5, xanchor: 'center', y: 0.98, yanchor: 'top', traceorder: 'reversed',
      font: { color: '#ffffff' }, bgcolor: 'rgba(0,0,0,0)' },
    margin: { t: 90, r: 20, b: 90, l: 20 },
  };
  images.org = await B.renderStyledPng(orgTraces, orgLayout, 560, 400);

  // Chart 2: KPI horizontal bar — white bg, single blue series, 0-100% axis,
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

  // ── Cell styling — the openpyxl block, one for one ──
  // header_fill/subheader_fill/grand_fill + bold + thin border + centering.
  // ExcelJS wants ARGB, openpyxl takes RGB, hence the FF alpha prefix.
  var HEADER_FILL = 'FFB7DEE8', SUBHEADER_FILL = 'FFD9EAF7';
  var THIN = { style: 'thin', color: { argb: 'FF000000' } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  var CENTER = { horizontal: 'center', vertical: 'middle' };
  function style(row, col, opts) {
    var cell = ws.getCell(row, col);
    if (opts.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
    if (opts.bold) cell.font = { bold: true };
    if (opts.border) cell.border = BORDER;
    if (opts.center) cell.alignment = CENTER;
    if (opts.numFmt) cell.numFmt = opts.numFmt;
  }

  // Filter area: labels in A, values in B.
  [1, 2].forEach(function (r) {
    style(r, 1, { fill: HEADER_FILL, bold: true, border: true });
    style(r, 2, { fill: SUBHEADER_FILL, border: true });
  });

  // Pivot table: header row, then bordered body with everything but the row
  // label centred, then the Grand Total row re-filled and bolded on top.
  var PIVOT_COLS = 4;
  for (var pc = 1; pc <= PIVOT_COLS; pc++) {
    style(startRow, pc, { fill: HEADER_FILL, bold: true, border: true, center: true });
  }
  pivotDisplay.forEach(function (r, i) {
    var rowNum = startRow + 1 + i;
    for (var c = 1; c <= PIVOT_COLS; c++) {
      style(rowNum, c, { border: true, center: c > 1 });
      if (r.label === 'Grand Total') style(rowNum, c, { fill: HEADER_FILL, bold: true });
    }
  });

  // KPI table: header row, bordered rows, values shown as whole percents.
  for (var kc = 1; kc <= 3; kc++) {
    style(kpiStartRow, kc, { fill: HEADER_FILL, bold: true, border: true, center: true });
  }
  KPI.forEach(function (_, i) {
    var rowNum = kpiStartRow + 1 + i;
    for (var c = 1; c <= 3; c++) style(rowNum, c, { border: true });
    style(rowNum, 2, { numFmt: '0%' });
  });
  style(noteRow, 1, { bold: true });

  // Column widths (A-M), uniform row heights, and freeze_panes = "A5" — which
  // is a 4-row vertical split, so the pivot header stays put while scrolling.
  var WIDTHS = [18, 12, 12, 14, 4, 14, 14, 14, 14, 14, 14, 14, 14];
  WIDTHS.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
  // for row in range(1, 50) — rows 1..49, not "as many rows as the grid has".
  // With enough zones the table can now run past row 49 unstyled, exactly as it
  // does in the script.
  for (var rh = 1; rh <= 49; rh++) ws.getRow(rh).height = 22;
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: startRow }];

  // Native, editable charts (data-linked to hidden helper blocks) replace the
  // two baked PNGs: org = stacked column, kpi = horizontal percent bar.
  var placements = [];
  if (window.NativeChartInject && window.fflate) {
    var R = window.NativeChartInject.ref;
    if (orgLabels.length) {
      // Org pivot: header row 4 (B=Open, C=Closed), data rows 5..4+n. The chart
      // excludes the appended Grand Total row (pivotRows only). Labels editable
      // in column A; legend names editable via B4/C4.
      var orgLast = 4 + pivotRows.length;
      placements.push({
        sheetName: 'Dashboard', anchor: { fromCol: 5, fromRow: 1, toCol: 14, toRow: 20 }, // ~"F2"
        def: {
          grouping: 'stacked', legend: true, title: chartYear + ' Assessment (' + closedTotal + '/' + recordTotal + ')',
          chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
          dataLabels: { position: 'ctr', color: 'FFFFFF' },
          categories: { ref: R('Dashboard', 1, 5, orgLast), cache: orgLabels },
          series: [
            { name: { ref: R('Dashboard', 2, 4, 4), lit: 'Open' },
              values: { ref: R('Dashboard', 2, 5, orgLast), cache: openVals }, color: '00AEEF' },
            { name: { ref: R('Dashboard', 3, 4, 4), lit: 'Closed' },
              values: { ref: R('Dashboard', 3, 5, orgLast), cache: closedVals }, color: 'D4AF37' },
          ],
        },
      });
    }
    if (kpiMetrics.length) {
      // KPI block: header at kpiStartRow (A=Metric, B=Value), data rows below.
      // The Baseline/Q1/Q2/Target labels sit in column A and are editable there.
      var kpiFirst = kpiStartRow + 1, kpiLast = kpiStartRow + kpiMetrics.length;
      placements.push({
        sheetName: 'Dashboard', anchor: { fromCol: 5, fromRow: 22, toCol: 14, toRow: 40 }, // ~"F23"
        def: {
          grouping: 'clustered', barDir: 'bar', catReversed: true, legend: false,
          title: 'Improve in Supplier Response Time', valNumFmt: '0%',
          dataLabels: { position: 'outEnd', numFmt: '0%', color: '000000' },
          categories: { ref: R('Dashboard', 1, kpiFirst, kpiLast), cache: kpiMetrics },
          series: [
            { name: { ref: R('Dashboard', 2, kpiStartRow, kpiStartRow), lit: 'Improvement' },
              values: { ref: R('Dashboard', 2, kpiFirst, kpiLast), cache: kpiValues }, color: '4472C4' },
          ],
        },
      });
    }
  }

  var buf = await workbook.xlsx.writeBuffer();
  if (placements.length) {
    try { buf = window.NativeChartInject.inject(new Uint8Array(buf), placements); }
    catch (e) { console.error('d2 native chart inject failed:', e); }
  }

  return {
    ok: true,
    stdout: stdout,
    files: [{ name: 'output file D2.xlsx', bytes: buf, sheets: [{ name: 'Dashboard', grid: grid }] }],
    chartImages: { 'Dashboard': images.org },
  };
};
