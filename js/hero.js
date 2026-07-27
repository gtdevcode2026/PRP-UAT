// ── Hero — Create PRP builder UI + GSAP choreography ──────────────────────
// The builder drives window.CreatePRP and hands the generated workbook to the
// existing pipeline via handleFileChange(). Animations are strictly
// progressive enhancement: nothing is hidden in CSS, all entrances are GSAP
// .from() tweens, and the page is fully usable if GSAP never loads.
(function () {
  'use strict';

  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var KIND_LABEL = { risk: 'Risk export', assessment: 'Assessment export', tprm: 'TPRM / Supplier' };
  var SLOT_HINT = {
    risk: '*risk-export*.xlsx',
    assessment: '*assessment-export*.xlsx',
    tprm: '*tprm* / *supplier*.xlsx',
  };
  var ICON_DL = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/><path d="m7 11 5 5 5-5"/><path d="M5 21h14"/></svg>';

  var H = { files: { risk: null, assessment: null, tprm: null }, unassigned: [], running: false, url: null };

  function el(id) { return document.getElementById(id); }
  function escText(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function setStatus(msg, cls) {
    var s = el('prp-status');
    if (!s) return;
    s.textContent = msg;
    s.className = 'prp-status' + (cls ? ' ' + cls : '');
  }

  function missingKinds() {
    return ['risk', 'assessment', 'tprm'].filter(function (k) { return !H.files[k]; });
  }

  function updateGenerateBtn() {
    var btn = el('prp-generate-btn');
    var missing = missingKinds();
    btn.disabled = missing.length > 0 || H.running;
    btn.textContent = H.running ? 'Creating…' : 'Create PRP workbook';
    if (!H.running) {
      if (missing.length === 3) setStatus('Add all three exports to enable');
      else if (missing.length) setStatus('Still missing: ' + missing.map(function (k) { return KIND_LABEL[k]; }).join(' · '));
      else setStatus('Ready to create the consolidated workbook');
    }
  }

  function renderSlots() {
    ['risk', 'assessment', 'tprm'].forEach(function (kind) {
      var slot = el('prp-slot-' + kind);
      var f = H.files[kind];
      slot.classList.toggle('filled', !!f);
      var fileEl = slot.querySelector('.prp-slot-file');
      fileEl.textContent = f ? (f.name + ' · ' + (f.size / 1024).toFixed(1) + ' KB') : SLOT_HINT[kind];
    });
    renderUnassigned();
    updateGenerateBtn();
  }

  function renderUnassigned() {
    var tray = el('prp-unassigned');
    if (!H.unassigned.length) { tray.hidden = true; tray.innerHTML = ''; return; }
    tray.hidden = false;
    var html = '<div class="prp-unassigned-hint">Could not recognise these from their filenames. Assign manually:</div>';
    H.unassigned.forEach(function (f, i) {
      html += '<div class="prp-unassigned-row"><span class="prp-unassigned-name" title="' + escText(f.name) + '">' + escText(f.name) + '</span>' +
        ['risk', 'assessment', 'tprm'].map(function (k) {
          return '<button type="button" class="prp-assign-btn" data-idx="' + i + '" data-kind="' + k + '">' + KIND_LABEL[k] + '</button>';
        }).join('') + '</div>';
    });
    tray.innerHTML = html;
    tray.querySelectorAll('.prp-assign-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var f = H.unassigned.splice(+btn.dataset.idx, 1)[0];
        if (f) assign(btn.dataset.kind, f);
        renderSlots();
      });
    });
  }

  function assign(kind, file) {
    // Last file wins, like the Python folder scan — but say so.
    if (H.files[kind]) setStatus('Replaced ' + KIND_LABEL[kind] + ' with "' + file.name + '"');
    H.files[kind] = file;
  }

  function addFiles(fileList) {
    var skipped = 0;
    Array.prototype.slice.call(fileList).forEach(function (f) {
      if (!/\.xlsx$/i.test(f.name)) { skipped++; return; }
      var kind = window.CreatePRP.classify(f.name);
      if (kind) assign(kind, f);
      else H.unassigned.push(f);
    });
    renderSlots();
    if (skipped) setStatus(skipped + ' file' + (skipped > 1 ? 's' : '') + ' skipped. Only .xlsx exports are supported', 'err');
  }

  async function handleGenerate() {
    if (H.running) return;
    if (typeof ExcelJS === 'undefined' || !ExcelJS.Workbook) {
      setStatus('Excel engine still loading. Try again in a moment', 'err');
      return;
    }
    H.running = true;
    updateGenerateBtn();
    var out = el('prp-outputs');
    out.classList.remove('show');
    out.innerHTML = '';
    try {
      var res = await window.CreatePRP.build(H.files, function (msg) { setStatus(msg); });
      var name = window.CreatePRP.OUTPUT_NAME;
      if (H.url) URL.revokeObjectURL(H.url);
      H.url = URL.createObjectURL(new Blob([res.bytes], { type: XLSX_MIME }));
      out.innerHTML =
        '<div class="run-out-top">' +
        '<span class="run-out-label">Download &middot; consolidated workbook</span>' +
        '<span class="run-out-ok">&#9679; complete</span>' +
        '</div>' +
        '<div class="dl-grid"><a class="dl-btn" href="' + H.url + '" download="' + escText(name) + '">' + ICON_DL + escText(name) + '</a></div>';
      out.classList.add('show');

      // Auto-feed the generated workbook into the pipeline's Workbook step.
      var fed = false;
      try {
        if (typeof handleFileChange === 'function' && typeof File === 'function') {
          handleFileChange(new File([res.bytes], name, { type: XLSX_MIME }));
          fed = true;
        }
      } catch (e) { /* download remains the handoff */ }
      var okMsg = fed
        ? 'Workbook created. Download above; it is already loaded into Step 2 below'
        : 'Workbook created. Download it, then upload in Step 2 below';
      if (res.warnings.length) {
        okMsg += ' · ' + res.warnings.length + ' warning' + (res.warnings.length > 1 ? 's' : '') + ' (see console)';
        console.warn('CreatePRP warnings:', res.warnings);
      }
      setStatus(okMsg, 'ok');
    } catch (err) {
      setStatus('Failed: ' + (err && err.message ? err.message : String(err)), 'err');
    } finally {
      H.running = false;
      var btn = el('prp-generate-btn');
      btn.disabled = missingKinds().length > 0;
      btn.textContent = 'Create PRP workbook';
    }
  }

  window.setupHeroBuilder = function setupHeroBuilder() {
    var dz = el('prp-dropzone');
    if (!dz) return;

    el('prp-file-input').addEventListener('change', function (e) {
      if (e.target.files.length) addFiles(e.target.files);
      e.target.value = '';
    });
    dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', function () { dz.classList.remove('drag-over'); });
    dz.addEventListener('drop', function (e) {
      e.preventDefault();
      dz.classList.remove('drag-over');
      if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
    });

    ['risk', 'assessment', 'tprm'].forEach(function (kind) {
      var slot = el('prp-slot-' + kind);
      var input = slot.querySelector('input[type="file"]');
      input.addEventListener('click', function (e) { e.stopPropagation(); });
      slot.addEventListener('click', function () { input.click(); });
      slot.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
      });
      input.addEventListener('change', function (e) {
        if (e.target.files.length) {
          var f = e.target.files[0];
          if (/\.xlsx$/i.test(f.name)) { assign(kind, f); renderSlots(); }
          else setStatus('Only .xlsx files are supported', 'err');
        }
        e.target.value = '';
      });
    });

    el('prp-generate-btn').addEventListener('click', handleGenerate);

    renderSlots();
  };

  // ── Arrival motion — idempotent; callable from the vendor loader and from
  // init(), whichever lands last with everything in place wins.
  //
  // This is a tool, not a landing page, so there is no orchestrated load
  // sequence: one 0.3s settle on the hero block and nothing else. The old
  // version staggered five elements over ~1.4s and reveal-gated the pipeline
  // controls on a ScrollTrigger, which left the whole app blank on arrival
  // and invisible entirely in print / headless renders. Every control is now
  // painted by CSS first; GSAP only softens how the hero lands. ──
  var _heroAnimated = false;
  window.tryInitHeroAnimations = function tryInitHeroAnimations() {
    if (_heroAnimated) return;
    if (!window.gsap) return;
    var hero = document.querySelector('.hero-inner');
    if (!hero) return;
    // Wait for the boot overlay to clear — otherwise the entrance plays unseen.
    var overlay = document.getElementById('prp-load-overlay');
    if (overlay && overlay.style.display !== 'none') return;
    _heroAnimated = true;
    if (window.scrollY > 40) return;  // deep-scrolled reload: nothing to settle

    gsap.matchMedia().add('(prefers-reduced-motion: no-preference)', function () {
      gsap.from(hero, { autoAlpha: 0, y: 10, duration: 0.3, ease: 'power2.out' });
    });
  };
})();
