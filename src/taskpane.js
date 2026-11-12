/*
 * taskpane.js — Word parallel reader (对照阅读).
 *
 * The task pane shows a second .docx next to the open document and keeps the two
 * aligned. Alignment uses headings, section keywords (摘要 / Abstract …) and
 * figure/table captions as anchors, and paragraph lengths between anchors.
 *
 * Two sync modes, chosen from what the running Word supports:
 *   viewport  — Word for Mac 16.100+ / Windows 2508+ (WordApiDesktop 1.4):
 *               the pane polls Word's scroll position and follows it; scrolling
 *               the pane scrolls Word. The cursor in Word is never moved.
 *   selection — older Word or Word on the web: the pane follows the cursor
 *               when it moves; double-clicking in the pane moves Word's cursor.
 */
(function () {
  'use strict';

  var PR = window.PRAlign;

  var POLL_MS = 250;            // how often Word's scroll position is read
  var REF_TOP = 8;              // px below the pane top that corresponds to the top of Word's window
  var USER_IDLE_MS = 700;       // after the last wheel/touch/key input, the pane is no longer "user-driven"
  var PUSH_THROTTLE_MS = 120;   // minimum interval between scroll writes to Word
  var OWN_WRITE_MS = 1500;      // Word scroll changes this soon after our own write are treated as ours
  var RESCAN_DELAY_MS = 1500;   // wait after edits in Word before re-reading the main document
  var PERIODIC_RESCAN_MS = 30000;

  var STYLE_MAP = [
    "p[style-name='Title'] => h1.pr-title:fresh",
    "p[style-name='Subtitle'] => p.pr-subtitle:fresh",
    "p[style-name='Caption'] => p.pr-caption:fresh",
    "p.PRH1 => h1:fresh",
    "p.PRH2 => h2:fresh",
    "p.PRH3 => h3:fresh",
    "p.PRH4 => h4:fresh",
    "p.PRH5 => h5:fresh",
    "p.PRH6 => h6:fresh",
    "p.PRH7 => h6.pr-h7:fresh",
    "p.PRH8 => h6.pr-h8:fresh",
    "p.PRH9 => h6.pr-h9:fresh"
  ];

  var el = {};
  var state = {
    inWord: false,
    caps: {},
    mode: 'none',            // 'viewport' | 'selection' | 'none'
    syncOn: true,
    fontPx: 15,
    docKey: 'doc:untitled',
    fileName: null,
    main: null,              // model of the document open in Word
    pane: null,              // model of the imported document
    align: null,             // PRAlign.alignDocuments result
    mapper: null,            // main position (pages, or 0..1 without page data) <-> pane px
    paneAnchors: [],         // [{top, node}] of matched anchors, for the "current section" mark
    dEst: { n: 0, lo: 0, hi: 0 }, // bounds on Word's scroll range, in pages
    fractional: false,       // true once Word reports a non-integer scroll percent
    useBottom: true,         // whether the last-visible-page constraint is trusted
    bottomMisses: 0,
    selSnapAt: 0,            // when the pane last jumped to the cursor position
    lastSelBlock: -1,
    paneTops: [],
    lastS: null,             // last observed Word scroll percent
    forceFollow: true,
    wordSetValue: null,
    wordSetAt: 0,
    paneUserUntil: 0,
    progScrollUntil: 0,
    pushTimer: 0,
    loopRunning: false,
    pollErrors: 0,
    pagesInViewportBroken: false,
    scanning: false,
    rescanAgain: false,
    rescanTimer: 0,
    resizeTimer: 0,
    selTimer: 0
  };

  /* ================================================================== */
  /* Start-up                                                            */
  /* ================================================================== */

  document.addEventListener('DOMContentLoaded', function () {
    ['scroller', 'doc', 'empty', 'drop', 'status', 'file', 'details', 'details-body', 'menu'].forEach(function (id) {
      el[id.replace(/-(\w)/g, function (_, c) { return c.toUpperCase(); })] = document.getElementById(id);
    });
    el.importBtn = document.getElementById('btn-import');
    el.toggle = document.getElementById('sync-toggle');
    el.moreBtn = document.getElementById('btn-more');
    el.detailsClose = document.getElementById('details-close');
    loadPrefs();
    bindUI();
    if (typeof Office === 'undefined') {
      setStatusHtml('<span class="warn">无法加载 Office 组件。请确认网络连接后重新打开加载项。</span>');
      return;
    }
    Office.onReady(function (info) {
      state.inWord = !!(info && info.host === Office.HostType.Word);
      if (!state.inWord) {
        state.mode = 'none';
        updateStatus();
        return;
      }
      start().catch(function (e) { showError('初始化失败', e); });
    });
  });

  async function start() {
    detectCaps();
    try { state.docKey = 'doc:' + (Office.context.document.url || 'untitled'); } catch (e) { /* keep default */ }
    registerSelectionHandler();
    await rescanMain();
    await restoreLastFile();
    registerChangeEvents();
    updateStatus();
    if (state.syncOn) startSync();
  }

  function detectCaps() {
    var req = Office.context.requirements;
    var has = function (name, v) { try { return req.isSetSupported(name, v); } catch (e) { return false; } };
    state.caps = {
      api13: has('WordApi', '1.3'),
      api16: has('WordApi', '1.6'),
      d12: has('WordApiDesktop', '1.2'),
      d14: has('WordApiDesktop', '1.4')
    };
    state.mode = state.caps.d14 ? 'viewport' : 'selection';
  }

  /* ================================================================== */
  /* UI                                                                  */
  /* ================================================================== */

  function bindUI() {
    el.importBtn.addEventListener('click', function () { el.file.value = ''; el.file.click(); });
    el.file.addEventListener('change', function () {
      var f = el.file.files && el.file.files[0];
      if (f) importFile(f);
    });

    el.toggle.checked = state.syncOn;
    el.toggle.addEventListener('change', function () {
      state.syncOn = el.toggle.checked;
      savePref('pr.sync', state.syncOn ? '1' : '0');
      if (state.syncOn) {
        if (state.caps.d14 && state.mode === 'selection' && state.inWord) state.mode = 'viewport'; // retry after a fallback
        state.forceFollow = true;
        startSync();
        if (state.mode === 'selection') followSelection();
      }
      updateStatus();
    });

    el.moreBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      var open = el.menu.hidden;
      el.menu.hidden = !open;
      el.moreBtn.setAttribute('aria-expanded', String(open));
    });
    document.addEventListener('click', function (e) {
      if (!el.menu.hidden && !el.menu.contains(e.target)) { el.menu.hidden = true; el.moreBtn.setAttribute('aria-expanded', 'false'); }
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') { el.menu.hidden = true; el.details.hidden = true; }
    });
    el.menu.addEventListener('click', function (e) {
      var b = e.target.closest('button[data-act]');
      if (!b) return;
      var act = b.getAttribute('data-act');
      if (act === 'font-' || act === 'font+') { setFont(state.fontPx + (act === 'font+' ? 1 : -1)); return; }
      el.menu.hidden = true;
      if (act === 'rescan') rescanMain();
      else if (act === 'details') { renderDetails(); el.details.hidden = false; }
      else if (act === 'clear') clearImported();
    });
    el.detailsClose.addEventListener('click', function () { el.details.hidden = true; });

    // user intent on the pane
    ['wheel', 'touchstart', 'touchmove', 'pointerdown', 'keydown'].forEach(function (ev) {
      el.scroller.addEventListener(ev, markUser, { passive: true });
    });
    el.scroller.addEventListener('scroll', onPaneScroll, { passive: true });
    el.doc.addEventListener('dblclick', onDocDoubleClick);
    el.doc.addEventListener('click', onDocClick);

    // drag and drop
    el.scroller.addEventListener('dragover', function (e) {
      e.preventDefault();
      el.drop.classList.add('over');
    });
    el.scroller.addEventListener('dragleave', function () { el.drop.classList.remove('over'); });
    el.scroller.addEventListener('drop', function (e) {
      e.preventDefault();
      el.drop.classList.remove('over');
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) importFile(f);
    });

    if (window.ResizeObserver) {
      new ResizeObserver(function () { scheduleRelayout(); }).observe(el.doc);
    }
    window.addEventListener('resize', scheduleRelayout);
    setFont(state.fontPx, true);
  }

  function setFont(px, silent) {
    state.fontPx = Math.max(11, Math.min(24, px));
    document.documentElement.style.setProperty('--doc-size', state.fontPx + 'px');
    if (!silent) { savePref('pr.font', String(state.fontPx)); scheduleRelayout(); }
  }

  function loadPrefs() {
    var s = readPref('pr.sync');
    if (s !== null) state.syncOn = s === '1';
    var f = parseInt(readPref('pr.font'), 10);
    if (f) state.fontPx = f;
  }
  function readPref(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
  function savePref(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* ignore */ } }

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function setStatusHtml(html) { el.status.innerHTML = html; }

  function showError(prefix, e) {
    var msg = e && (e.message || e.code) ? (e.message || e.code) : String(e || '');
    console.error(prefix, e);
    setStatusHtml('<span class="warn">' + esc(prefix + (msg ? '：' + msg : '')) + '</span>');
  }

  function updateStatus() {
    var line1 = [];
    if (state.fileName) line1.push('<span class="file">' + esc(state.fileName) + '</span>');
    if (state.align) {
      var s = state.align.stats;
      line1.push('已对齐 ' + s.anchorsMatched + ' 处标题/题注');
      if (s.paraGroups) line1.push('段落一一对应 ' + Math.round(100 * s.paraOneToOne / s.paraGroups) + '%');
    }
    var line2 = '';
    if (!state.inWord) {
      line2 = '<span class="warn">当前不在 Word 中，只能预览，不能同步。</span>';
    } else if (!state.syncOn) {
      line2 = '同步滚动已关闭。';
    } else if (state.mode === 'selection') {
      line2 = '<span class="warn">光标跟随模式</span>：这个版本的 Word 不提供滚动位置，在正文中点击或移动光标时右侧会跟上。';
    } else if (state.pane && state.main && state.main.unit !== 'page') {
      line2 = '<span class="warn">没有读到分页信息</span>，同步按全文比例估算，可能偏差较大。';
    }
    var html = line1.join(' · ');
    if (line2) html += (html ? '<br>' : '') + line2;
    setStatusHtml(html);
  }

  /* ================================================================== */
  /* Word access                                                         */
  /* ================================================================== */

  var wordQueue = Promise.resolve();
  // Serialises Word.run calls so that polling, scanning and writes never overlap.
  function wordRun(fn) {
    var run = wordQueue.then(function () { return Word.run(fn); });
    wordQueue = run.catch(function () { /* keep the queue alive */ });
    return run;
  }

  function safe(obj, key, dflt) {
    try { var v = obj[key]; return v === undefined ? dflt : v; } catch (e) { return dflt; }
  }

  /* ---------------- main document model ---------------- */

  // Reads the document open in Word. Optional details (fonts, pictures, tables,
  // paragraph ids, pages) are read in separate steps, so that a property missing
  // from a particular Word build only removes that detail instead of the scan.
  async function scanMain() {
    var caps = state.caps;
    var core = await wordRun(async function (ctx) {
      var paras = ctx.document.body.paragraphs;
      var base = ['text', 'style'];
      if (caps.api13) base.push('styleBuiltIn', 'tableNestingLevel', 'isListItem');
      var extra = ['outlineLevel', 'lineSpacing', 'spaceAfter', 'spaceBefore', 'font/size'];
      var toLoad = function (list) { return list.map(function (p) { return 'items/' + p; }).join(','); };
      try {
        paras.load(toLoad(base.concat(extra)));
        await ctx.sync();
      } catch (e) {
        console.warn('layout properties unavailable', e);
        paras = ctx.document.body.paragraphs;
        paras.load(toLoad(base));
        await ctx.sync();
      }
      var items = paras.items;
      var boldReq = [], listReq = [], picReq = [];
      var bold = new Map(), lists = new Map(), pics = new Map(), tableValues = [];
      try {
        items.forEach(function (p, i) {
          var t = (p.text || '').trim();
          var inTable = caps.api13 && safe(p, 'tableNestingLevel', 0) > 0;
          if (t.length <= 2 && !inTable) {
            // a paragraph that holds a picture has (almost) no text
            var pc = p.inlinePictures;
            pc.load('items/height');
            picReq.push([i, pc]);
          }
          if (!t) return;
          if (t.length <= 80 && !inTable) { p.font.load('bold'); boldReq.push(i); }
          if (caps.api13 && safe(p, 'isListItem', false) && headingStyleLevel(p) !== null) {
            var li = p.listItemOrNullObject;
            li.load('listString');
            listReq.push([i, li]);
          }
        });
        var tables = null;
        if (caps.api13) {
          tables = ctx.document.body.tables;
          tables.load('items/values,items/nestingLevel');
        }
        await ctx.sync();
        boldReq.forEach(function (i) { bold.set(i, safe(items[i].font, 'bold', false) === true); });
        listReq.forEach(function (pair) {
          if (!pair[1].isNullObject) lists.set(pair[0], safe(pair[1], 'listString', ''));
        });
        picReq.forEach(function (pair) {
          var h = 0;
          (safe(pair[1], 'items', []) || []).forEach(function (pic) { h += +safe(pic, 'height', 0) || 0; });
          if (h > 0) pics.set(pair[0], h);
        });
        if (tables) {
          (safe(tables, 'items', []) || []).forEach(function (t) {
            if (safe(t, 'nestingLevel', 1) === 1) tableValues.push(safe(t, 'values', null));
          });
        }
      } catch (e) {
        console.warn('formatting details unavailable', e);
      }
      return { items: items, bold: bold, lists: lists, pics: pics, tableValues: tableValues };
    });

    // paragraph ids let the cursor be located cheaply (Microsoft 365 only)
    var ids = null;
    if (caps.api16) {
      try {
        ids = await wordRun(async function (ctx) {
          var paras = ctx.document.body.paragraphs;
          paras.load('items/uniqueLocalId');
          await ctx.sync();
          return paras.items.map(function (p) { return p.uniqueLocalId; });
        });
        if (ids.length !== core.items.length) ids = null;
      } catch (e) {
        ids = null;
      }
    }

    var pageInfo = null;
    if (caps.d12) {
      try {
        pageInfo = await wordRun(async function (ctx) {
          var pages = ctx.document.activeWindow.activePane.pages;
          pages.load('items/index,items/height,items/width');
          await ctx.sync();
          var ranges = pages.items.map(function (pg) { var r = pg.getRange(); r.load('text'); return r; });
          await ctx.sync();
          return {
            lens: ranges.map(function (r) { return (r.text || '').length; }),
            width: +safe(pages.items[0], 'width', 595) || 595,
            height: +safe(pages.items[0], 'height', 842) || 842
          };
        });
      } catch (e) {
        console.warn('page model unavailable', e);
        pageInfo = null;
      }
    }
    return buildMainModel(core.items, core.bold, core.lists, core.pics, ids, pageInfo, core.tableValues);
  }

  function headingStyleLevel(p) {
    var sb = safe(p, 'styleBuiltIn', null);
    var m;
    if (sb && (m = /^Heading(\d)$/i.exec(sb))) return +m[1];
    if (sb && /^Title$/i.test(sb)) return 0;
    var st = (safe(p, 'style', '') || '').trim();
    if ((m = /^(?:heading|标题)\s*(\d)$/i.exec(st))) return +m[1];
    if (/^(?:title|标题)$/i.test(st)) return 0;
    return null;
  }

  function num(v, dflt) { v = +v; return isFinite(v) && v > 0 ? v : dflt; }

  // Rough height of a paragraph in points, from its text width, font size, line
  // spacing, spacing before/after and picture height. Only ratios between
  // paragraphs on the same page matter, so the constants need not be exact.
  function estimateHeight(p, text, picH, usable) {
    var size = num(safe(p.font, 'size', null), 12);
    var ls = num(safe(p, 'lineSpacing', null), 12);
    var mult = ls >= 6 && ls <= 48 ? ls / 12 : 1;   // Word reports single spacing as 12
    var lineH = Math.max(size, mult * size * 1.2);
    var before = +safe(p, 'spaceBefore', 0) || 0, after = +safe(p, 'spaceAfter', 0) || 0;
    if (picH) return picH + before + after + 4;
    var inTable = safe(p, 'tableNestingLevel', 0) > 0;
    var w = 0;
    for (var k = 0; k < text.length; k++) {
      var ch = text.charCodeAt(k);
      w += ch === 32 ? 0.25 : (ch >= 0x2e80 ? 1 : 0.44); // em widths: CJK full, Latin ~Times average
    }
    var lines = Math.max(1, Math.ceil(w * size / (inTable ? usable / 3 : usable)));
    var h = lines * lineH + before + after;
    return inTable ? h / 3 : h; // cells of a row sit side by side
  }

  // Height of a table in points: each row is as tall as its tallest cell.
  function tableHeight(values, usable, lineH) {
    if (!values || !values.length) return null;
    var h = 0;
    values.forEach(function (row) {
      var cols = Math.max(1, row.length), lines = 1;
      row.forEach(function (cell) {
        var w = 0;
        for (var k = 0; k < cell.length; k++) {
          var ch = cell.charCodeAt(k);
          w += ch === 32 ? 0.25 : (ch >= 0x2e80 ? 1 : 0.44);
        }
        lines = Math.max(lines, Math.ceil(w * lineH / 1.2 / (usable / cols - 10)) || 1);
      });
      h += lines * lineH + 6;
    });
    return h;
  }

  function buildMainModel(items, bold, lists, pics, ids, pageInfo, tableValues) {
    // Outline level of body text: the most common value. Word reports either
    // 1..9 + 10 (body) or 0..8 + 9 (body); both cases are handled.
    var counts = {};
    items.forEach(function (p) { var o = safe(p, 'outlineLevel', null); if (o !== null) counts[o] = (counts[o] || 0) + 1; });
    var bodyLevel = null, best = -1;
    Object.keys(counts).forEach(function (k) { if (counts[k] > best) { best = counts[k]; bodyLevel = +k; } });
    var zeroBased = bodyLevel === 9;

    var n = items.length;
    var c0 = new Array(n), clen = new Array(n), hts = new Array(n);
    var usable = Math.max(200, (pageInfo ? pageInfo.width : 595) - 160);
    var blocks = [];
    var c = 0, table = null, tableNo = 0, groups = [];
    items.forEach(function (p, i) {
      var text = p.text || '';
      c0[i] = c;
      clen[i] = text.length + 1;
      c += clen[i];
      hts[i] = estimateHeight(p, text, pics.get(i), usable);
      var inTable = safe(p, 'tableNestingLevel', 0) > 0;
      if (inTable) {
        if (!table) {
          table = { text: '', paraIndex: i, inTable: true, styleLevel: null };
          blocks.push(table);
          groups.push({ first: i, last: i, values: tableValues ? tableValues[tableNo] : null, lineH: 0 });
          tableNo++;
        }
        groups[groups.length - 1].last = i;
        var size = num(safe(p.font, 'size', null), 12);
        groups[groups.length - 1].lineH = Math.max(groups[groups.length - 1].lineH, size * 1.2);
        table.text += text + ' ';
        return;
      }
      table = null;
      if (!text.trim()) return;
      var level = headingStyleLevel(p);
      if (level === null) {
        var ol = safe(p, 'outlineLevel', null);
        if (typeof ol === 'number' && ol !== bodyLevel) {
          var lv = zeroBased ? ol + 1 : ol;
          if (lv >= 1 && lv <= 9) level = lv;
        }
      }
      var sb = safe(p, 'styleBuiltIn', '') || '';
      var st = safe(p, 'style', '') || '';
      blocks.push({
        text: text,
        paraIndex: i,
        styleLevel: level,
        isBold: bold.get(i) === true,
        isCaptionStyle: /^caption$/i.test(sb) || /^(caption|题注)$/i.test(st.trim()),
        inTable: false,
        listString: lists.get(i) || null
      });
    });
    var total = Math.max(c, 1);
    // spread each table's height over its paragraphs
    groups.forEach(function (g) {
      var th = tableHeight(g.values, usable, g.lineH || 14.4);
      if (!th) return;
      var cnt = g.last - g.first + 1;
      for (var q = g.first; q <= g.last; q++) hts[q] = th / cnt;
    });

    var model = { blocks: blocks, totalChars: total, paraCount: n, c0: c0, clen: clen, unit: 'char', end: 1 };
    if (ids) {
      model.idIndex = new Map();
      ids.forEach(function (id, i) { if (id) model.idIndex.set(id, i); });
    }

    if (pageInfo && pageInfo.lens.length) {
      buildPageGeometry(model, hts, pageInfo.lens, pageInfo.height - 144);
    }
    model.posOfPara = function (i) {
      if (model.unit === 'page') return model.paraP[i];
      return c0[i] / total;
    };
    model.cToPos = function (C) {
      var i = paraAtChar(model, C);
      if (model.unit !== 'page') return C / total;
      var f = Math.max(0, Math.min(1, (C - c0[i]) / clen[i]));
      var parts = model.parts[i];
      for (var k = 0; k < parts.length; k++) {
        var q = parts[k];
        if (f <= q.f1 || k === parts.length - 1) {
          var g = q.f1 > q.f0 ? Math.max(0, Math.min(1, (f - q.f0) / (q.f1 - q.f0))) : 0;
          return q.p0 + g * (q.p1 - q.p0);
        }
      }
      return model.paraP[i];
    };
    blocks.forEach(function (b) { b.pos = model.posOfPara(b.paraIndex); });
    return model;
  }

  function paraAtChar(model, C) {
    var lo = 0, hi = model.c0.length - 1;
    while (lo < hi) {
      var mid = (lo + hi + 1) >> 1;
      if (model.c0[mid] <= C) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  // Positions every paragraph on a continuous page axis P (page k spans
  // [k-1, k)). Page boundaries come from the characters Word reports per page;
  // positions inside a page come from the estimated paragraph heights.
  // A page whose last paragraph continues on the next page is full. Its content
  // height calibrates the page capacity, so that pages left short (a figure
  // pushed to the next page, the last page) keep their text at the top.
  function buildPageGeometry(model, hts, lens, nominalCapacity) {
    var N = lens.length, n = model.c0.length, total = model.totalChars;
    var sumL = lens.reduce(function (a, b) { return a + b; }, 0) || 1;
    var scale = total / sumL;
    var pStart = [], acc = 0;
    lens.forEach(function (L) { pStart.push(acc); acc += L * scale; });
    var parts = new Array(n), pageH = new Array(N).fill(0);
    // Page and paragraph character counts can disagree by a few characters
    // (fields, table markers). Overlaps smaller than this are rounding, not a
    // paragraph that really continues on the next page.
    var tol = Math.max(2, 0.004 * total / N);
    var k = 0;
    for (var i = 0; i < n; i++) {
      var s0 = model.c0[i], s1 = s0 + model.clen[i];
      while (k < N - 1 && pStart[k + 1] <= s0) k++;
      var raw = [];
      for (var kk = k; kk < N; kk++) {
        var ps = pStart[kk], pe = kk + 1 < N ? pStart[kk + 1] : Infinity;
        var a = Math.max(s0, ps), b = Math.min(s1, pe);
        if (b > a) raw.push({ k: kk, a: a, b: b });
        if (s1 <= pe) break;
      }
      if (raw.length > 1) {
        var kept = raw.filter(function (r) { return r.b - r.a >= tol; });
        if (!kept.length) kept = [raw.reduce(function (x, y) { return y.b - y.a > x.b - x.a ? y : x; })];
        kept[0].a = s0;
        kept[kept.length - 1].b = s1;
        raw = kept;
      }
      parts[i] = raw.map(function (r) {
        var fr = (r.b - r.a) / model.clen[i];
        pageH[r.k] += hts[i] * fr;
        return { k: r.k, f0: (r.a - s0) / model.clen[i], f1: (r.b - s0) / model.clen[i], h: hts[i] * fr };
      });
      if (!parts[i].length) parts[i].push({ k: k, f0: 0, f1: 1, h: 0 });
    }
    var full = new Array(N).fill(false);
    for (i = 0; i < n; i++) {
      for (var t = 0; t + 1 < parts[i].length; t++) full[parts[i][t].k] = true;
    }
    var fullH = pageH.filter(function (h, k) { return full[k] && h > 0; }).sort(function (a, b) { return a - b; });
    var capacity = fullH.length ? fullH[fullH.length >> 1] : Math.max(nominalCapacity || 0, Math.max.apply(null, pageH));
    var denom = pageH.map(function (h, k) { return full[k] ? (h || 1) : Math.max(capacity, h || 1); });
    var used = new Array(N).fill(0);
    var paraP = new Array(n);
    for (i = 0; i < n; i++) {
      parts[i].forEach(function (q) {
        var H = denom[q.k];
        q.p0 = q.k + used[q.k] / H;
        used[q.k] += q.h;
        q.p1 = q.k + Math.min(1, used[q.k] / H);
      });
      paraP[i] = parts[i][0].p0;
    }
    model.unit = 'page';
    model.end = N;
    model.pageCount = N;
    model.parts = parts;
    model.paraP = paraP;
  }

  /* ---------------- reading Word's scroll position ---------------- */

  // Word reports the scroll position as a percentage s of a scroll range D
  // (in pages): top of window = s/100 * D. D is learned from the first visible
  // page, which brackets the top of the window: page i is visible => i-1 <= top < i.
  function observeScroll(s, vis) {
    var m = state.main;
    if (!m || m.unit !== 'page') return;
    if (s % 1 !== 0) state.fractional = true;
    var N = m.end;
    var est = state.dEst;
    if (est.n !== N) { est.n = N; est.lo = Math.max(0.05, N - 3); est.hi = N; }
    if (!vis || !vis.length) return;
    var sLo = Math.max(0, state.fractional ? s - 0.02 : s - 0.5);
    var sHi = Math.min(100, state.fractional ? s + 0.02 : s + 1);
    var i = Math.min.apply(null, vis), j = Math.max.apply(null, vis);
    var lo = Math.max(0.05, N - 3), hi = N;
    if (sHi > 0) lo = Math.max(lo, 100 * (i - 1) / sHi);
    if (sLo > 0) hi = Math.min(hi, 100 * i / sLo);
    if (lo > hi) return;
    // If the scroll range is the document minus the window (D = N - window height),
    // the last visible page j also brackets the bottom of the window:
    // j-1 < top + (N - D) <= j. Used only while it agrees with everything else.
    if (state.useBottom && j < N) {
      var a = 1 - sHi / 100, b = 1 - sLo / 100;
      var lo2 = a > 0 ? (N - j) / b : lo, hi2 = a > 0 ? (N - j + 1) / a : hi;
      var l3 = Math.max(lo, lo2), h3 = Math.min(hi, hi2);
      var el3 = Math.max(est.lo, l3), eh3 = Math.min(est.hi, h3);
      if (l3 <= h3 && el3 <= eh3) { lo = l3; hi = h3; }
      else if (++state.bottomMisses > 5) state.useBottom = false;
    }
    var nlo = Math.max(est.lo, lo), nhi = Math.min(est.hi, hi);
    if (nlo > nhi) { nlo = lo; nhi = hi; } // zoom or window size changed: start again
    est.lo = nlo;
    est.hi = nhi;
  }

  function dHat() {
    var m = state.main, est = state.dEst;
    var N = m.end;
    var lo = est.n === N ? est.lo : Math.max(0.05, N - 3);
    var hi = est.n === N ? est.hi : N;
    if (hi - lo < 0.6) return (lo + hi) / 2;   // well bracketed: take the middle
    return Math.max(lo, Math.min(hi, N - 0.8)); // otherwise: a typical window shows ~0.8 page
  }

  // Word scroll percent -> position (main-document units) at the top of Word's window
  function viewportTopPos(s, vis) {
    var m = state.main;
    if (m.unit !== 'page') return Math.max(0, Math.min(1, s / 100));
    var top = Math.max(0, Math.min(100, s)) / 100 * dHat();
    if (vis && vis.length) {
      var i = Math.min.apply(null, vis);
      top = Math.max(i - 1, Math.min(i - 0.002, top));
    }
    return Math.max(0, Math.min(m.end, top));
  }

  // position that should be at the top of Word's window -> scroll percent
  function posToPercent(pos) {
    var m = state.main;
    var s = m.unit === 'page' ? 100 * pos / Math.max(dHat(), 0.05) : 100 * pos;
    s = Math.max(0, Math.min(100, s));
    return state.fractional ? Math.round(s * 100) / 100 : Math.round(s);
  }

  /* ---------------- rescanning ---------------- */

  async function rescanMain() {
    if (!state.inWord) return;
    if (state.scanning) { state.rescanAgain = true; return; }
    state.scanning = true;
    try {
      state.main = await scanMain();
      if (state.pane) realign();
      state.forceFollow = true;
      updateStatus();
    } catch (e) {
      showError('读取正文失败', e);
    } finally {
      state.scanning = false;
      if (state.rescanAgain) { state.rescanAgain = false; rescanMain(); }
    }
  }

  function scheduleRescan() {
    clearTimeout(state.rescanTimer);
    state.rescanTimer = setTimeout(rescanMain, RESCAN_DELAY_MS);
  }

  async function registerChangeEvents() {
    var ok = false;
    if (state.caps.api16) {
      try {
        await Word.run(async function (ctx) {
          ctx.document.onParagraphChanged.add(onWordEdited);
          ctx.document.onParagraphAdded.add(onWordEdited);
          ctx.document.onParagraphDeleted.add(onWordEdited);
          await ctx.sync();
        });
        ok = true;
      } catch (e) {
        console.warn('paragraph events unavailable', e);
      }
    }
    if (!ok) {
      setInterval(function () { if (state.syncOn && state.pane) rescanMain(); }, PERIODIC_RESCAN_MS);
    }
  }

  function onWordEdited() {
    scheduleRescan();
    return Promise.resolve();
  }

  /* ================================================================== */
  /* Imported document                                                   */
  /* ================================================================== */

  async function importFile(file) {
    if (!/\.docx$/i.test(file.name)) {
      setStatusHtml('<span class="warn">只支持 .docx 文件。旧版 .doc 请先在 Word 里另存为 .docx。</span>');
      return;
    }
    try {
      var buf = await file.arrayBuffer();
      await renderDocx(buf, file.name);
      idbPut(state.docKey, { name: file.name, data: buf, savedAt: Date.now() }).catch(function () { /* ignore */ });
    } catch (e) {
      showError('无法打开 ' + file.name, e);
    }
  }

  async function restoreLastFile() {
    try {
      var rec = await idbGet(state.docKey);
      if (rec && rec.data) await renderDocx(rec.data, rec.name);
    } catch (e) {
      console.warn('restore failed', e);
    }
  }

  async function clearImported() {
    el.doc.innerHTML = '';
    el.doc.hidden = true;
    el.empty.hidden = false;
    state.pane = null; state.align = null; state.mapper = null; state.fileName = null; state.paneAnchors = [];
    idbDel(state.docKey).catch(function () { /* ignore */ });
    updateStatus();
  }

  async function renderDocx(buf, name) {
    setStatusHtml('正在解析 ' + esc(name) + ' …');
    var prepared = preprocessDocx(buf);
    var result = await mammoth.convertToHtml({ arrayBuffer: prepared }, { styleMap: STYLE_MAP, includeDefaultStyleMap: true });
    el.doc.innerHTML = result.value;
    postProcessHtml(el.doc);
    el.empty.hidden = true;
    el.doc.hidden = false;
    markProgrammatic();
    el.scroller.scrollTop = 0;
    state.fileName = name;
    buildPaneModel();
    realign();
    state.forceFollow = true;
    updateStatus();
    // images change the layout once decoded
    el.doc.querySelectorAll('img').forEach(function (img) {
      if (!img.complete) img.addEventListener('load', scheduleRelayout, { once: true });
    });
  }

  // Gives every paragraph with an outline level (from its style or set directly,
  // e.g. 大纲级别 in the paragraph dialog) a synthetic style PRH1..PRH9, which the
  // style map turns into <h1>..<h6>. Returns the original buffer if nothing changes.
  function preprocessDocx(buf) {
    try {
      var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      var files = fflate.unzipSync(new Uint8Array(buf));
      var docBytes = files['word/document.xml'];
      if (!docBytes) return buf;
      var dec = new TextDecoder('utf-8');
      var parser = new DOMParser();

      var styleLevel = {};
      if (files['word/styles.xml']) {
        var sdoc = parser.parseFromString(dec.decode(files['word/styles.xml']), 'application/xml');
        var info = {};
        var styles = sdoc.getElementsByTagNameNS(W, 'style');
        for (var i = 0; i < styles.length; i++) {
          var s = styles[i];
          if (s.getAttributeNS(W, 'type') !== 'paragraph') continue;
          var based = s.getElementsByTagNameNS(W, 'basedOn')[0];
          var ol = s.getElementsByTagNameNS(W, 'outlineLvl')[0];
          info[s.getAttributeNS(W, 'styleId')] = {
            based: based ? based.getAttributeNS(W, 'val') : null,
            ol: ol ? parseInt(ol.getAttributeNS(W, 'val'), 10) : null
          };
        }
        var resolve = function (id, depth) {
          if (!id || depth > 12 || !info[id]) return null;
          if (info[id].ol !== null && !isNaN(info[id].ol)) return info[id].ol;
          return resolve(info[id].based, depth + 1);
        };
        Object.keys(info).forEach(function (id) { styleLevel[id] = resolve(id, 0); });
      }

      var ddoc = parser.parseFromString(dec.decode(docBytes), 'application/xml');
      var ps = ddoc.getElementsByTagNameNS(W, 'p');
      var changed = 0;
      for (var k = 0; k < ps.length; k++) {
        var p = ps[k];
        var pPr = childNS(p, W, 'pPr');
        var pStyle = pPr ? childNS(pPr, W, 'pStyle') : null;
        var lvl = null;
        var olEl = pPr ? childNS(pPr, W, 'outlineLvl') : null;
        if (olEl) lvl = parseInt(olEl.getAttributeNS(W, 'val'), 10);
        if ((lvl === null || isNaN(lvl)) && pStyle) {
          var v = styleLevel[pStyle.getAttributeNS(W, 'val')];
          lvl = v === undefined ? null : v;
        }
        if (lvl === null || isNaN(lvl) || lvl < 0 || lvl > 8) continue;
        if (!pPr) { pPr = ddoc.createElementNS(W, 'w:pPr'); p.insertBefore(pPr, p.firstChild); }
        if (!pStyle) { pStyle = ddoc.createElementNS(W, 'w:pStyle'); pPr.insertBefore(pStyle, pPr.firstChild); }
        pStyle.setAttributeNS(W, 'w:val', 'PRH' + (lvl + 1));
        changed++;
      }
      if (!changed) return buf;
      files['word/document.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(ddoc));
      var out = fflate.zipSync(files);
      return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
    } catch (e) {
      console.warn('preprocess skipped', e);
      return buf;
    }
  }

  function childNS(node, ns, local) {
    for (var c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && c.namespaceURI === ns && c.localName === local) return c;
    }
    return null;
  }

  function postProcessHtml(root) {
    root.querySelectorAll('a[href]').forEach(function (a) {
      if (/^\s*javascript:/i.test(a.getAttribute('href') || '')) a.removeAttribute('href');
    });
    root.querySelectorAll('table').forEach(function (t) {
      var wrap = document.createElement('div');
      wrap.className = 'pr-table-wrap';
      t.parentNode.insertBefore(wrap, t);
      wrap.appendChild(t);
    });
    var firstNote = root.querySelector('li[id^="footnote-"], li[id^="endnote-"]');
    if (firstNote && firstNote.parentElement) {
      var list = firstNote.parentElement;
      list.classList.add('pr-notes');
      var hr = document.createElement('hr');
      hr.className = 'pr-notes-sep';
      list.parentNode.insertBefore(hr, list);
    }
  }

  function ownText(li) {
    var t = '';
    li.childNodes.forEach(function (c) {
      if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) return;
      t += c.textContent;
    });
    return t;
  }

  function isAllBold(node) {
    var total = node.textContent.replace(/\s+/g, '').length;
    if (!total) return false;
    var bold = 0;
    node.querySelectorAll('strong, b').forEach(function (b) {
      var outer = b.parentElement && b.parentElement.closest('strong, b');
      if (outer && node.contains(outer)) return; // nested; counted with the outer element
      bold += b.textContent.replace(/\s+/g, '').length;
    });
    return bold >= total * 0.9;
  }

  function buildPaneModel() {
    var root = el.doc;
    var notes = root.querySelector('.pr-notes');
    var nodes = root.querySelectorAll('h1, h2, h3, h4, h5, h6, p, li, table');
    var blocks = [];
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (notes && notes.contains(n)) continue;
      var tag = n.tagName;
      if (tag !== 'TABLE' && n.closest('table')) continue;
      if (tag === 'P' && n.parentElement && n.parentElement.tagName === 'LI') continue;
      var text = tag === 'LI' ? ownText(n) : n.textContent;
      if (!text || !text.trim()) continue;
      var level = null;
      var m = /^H([1-6])$/.exec(tag);
      if (m) {
        level = +m[1];
        if (n.classList.contains('pr-title')) level = 0;
        else if (n.classList.contains('pr-h7')) level = 7;
        else if (n.classList.contains('pr-h8')) level = 8;
        else if (n.classList.contains('pr-h9')) level = 9;
      }
      blocks.push({
        el: n,
        text: text,
        styleLevel: level,
        isBold: tag === 'P' || tag === 'LI' ? isAllBold(n) : false,
        isCaptionStyle: n.classList.contains('pr-caption'),
        inTable: tag === 'TABLE',
        listString: null
      });
    }
    state.pane = { blocks: blocks, notes: notes };
  }

  /* ================================================================== */
  /* Alignment                                                           */
  /* ================================================================== */

  function realign() {
    if (!state.main || !state.pane) return;
    state.align = PR.alignDocuments(state.main.blocks, state.pane.blocks);
    markAnchors();
    rebuildMapper();
    if (!el.details.hidden) renderDetails();
  }

  function paneTop(node) {
    var r = node.getBoundingClientRect();
    var s = el.scroller.getBoundingClientRect();
    return r.top - s.top + el.scroller.scrollTop;
  }

  function paneContentEnd() {
    var sep = el.doc.querySelector('.pr-notes-sep');
    if (sep) return paneTop(sep);
    var blocks = state.pane.blocks;
    if (!blocks.length) return 0;
    var last = blocks[blocks.length - 1].el;
    return paneTop(last) + last.offsetHeight;
  }

  function rebuildMapper() {
    if (!state.main || !state.pane || !state.align) { state.mapper = null; return; }
    var tops = state.pane.blocks.map(function (b) { return paneTop(b.el); });
    state.paneTops = tops;
    var pts = [{ a: 0, b: 0 }];
    state.align.blockPairs.forEach(function (p) {
      var mb = state.main.blocks[p.a];
      if (mb && tops[p.b] !== undefined) pts.push({ a: mb.pos, b: Math.max(0, tops[p.b]) });
    });
    pts.push({ a: state.main.end, b: paneContentEnd() });
    state.mapper = PR.makeMapper(pts);
    var matched = new Set(state.align.anchorPairs.map(function (p) { return p.b; }));
    state.paneAnchors = [];
    state.pane.blocks.forEach(function (b, i) {
      if (matched.has(i)) state.paneAnchors.push({ top: tops[i], node: b.el });
    });
  }

  function markAnchors() {
    el.doc.querySelectorAll('.pr-anchor, .pr-anchor-unmatched, .pr-current').forEach(function (n) {
      n.classList.remove('pr-anchor', 'pr-anchor-unmatched', 'pr-current');
    });
    state.currentNode = null;
    var matched = new Set(state.align.anchorPairs.map(function (p) { return p.b; }));
    state.pane.blocks.forEach(function (b, i) {
      if (!b.anchor) return;
      b.el.classList.add(matched.has(i) ? 'pr-anchor' : 'pr-anchor-unmatched');
    });
  }

  function markCurrent(px) {
    var cur = null;
    for (var i = 0; i < state.paneAnchors.length; i++) {
      if (state.paneAnchors[i].top <= px + 4) cur = state.paneAnchors[i].node; else break;
    }
    if (cur === state.currentNode) return;
    if (state.currentNode) state.currentNode.classList.remove('pr-current');
    if (cur) cur.classList.add('pr-current');
    state.currentNode = cur;
  }

  function scheduleRelayout() {
    clearTimeout(state.resizeTimer);
    state.resizeTimer = setTimeout(function () {
      if (!state.pane) return;
      rebuildMapper();
      state.forceFollow = true;
    }, 150);
  }

  function renderDetails() {
    var body = el.detailsBody;
    if (!state.align || !state.main || !state.pane) {
      body.innerHTML = '<div class="details-note">导入对照文档后，这里会列出两边匹配上的标题和题注。</div>';
      return;
    }
    var A = state.main.blocks, B = state.pane.blocks;
    var html = '<div class="pair" style="cursor:default"><span class="t"><b>正文（左）</b></span><span class="arrow"></span><span class="t"><b>对照（右）</b></span></div>';
    var usedA = new Set(), usedB = new Set();
    state.align.anchorPairs.forEach(function (p, k) {
      usedA.add(p.a); usedB.add(p.b);
      html += '<div class="pair" data-k="' + k + '"><span class="t" title="' + esc(A[p.a].text) + '">' + esc(A[p.a].text) +
        '</span><span class="arrow">↔</span><span class="t" title="' + esc(B[p.b].text) + '">' + esc(B[p.b].text) + '</span></div>';
    });
    var ua = A.filter(function (b, i) { return b.anchor && !usedA.has(i); });
    var ub = B.filter(function (b, i) { return b.anchor && !usedB.has(i); });
    if (ua.length) {
      html += '<div class="details-sub">正文中未匹配</div>' + ua.map(function (b) { return '<div class="unmatched">' + esc(b.text) + '</div>'; }).join('');
    }
    if (ub.length) {
      html += '<div class="details-sub">对照文档中未匹配</div>' + ub.map(function (b) { return '<div class="unmatched">' + esc(b.text) + '</div>'; }).join('');
    }
    html += '<div class="details-note">相邻两个已匹配的标题之间，按段落长度把段落配对。点击一行，两边会同时跳到对应位置。</div>';
    body.innerHTML = html;
    body.querySelectorAll('.pair[data-k]').forEach(function (row) {
      row.addEventListener('click', function () {
        var p = state.align.anchorPairs[+row.getAttribute('data-k')];
        jumpBothToPaneBlock(state.pane.blocks[p.b].el, state.main.blocks[p.a].pos);
      });
    });
  }

  /* ================================================================== */
  /* Synchronisation                                                     */
  /* ================================================================== */

  function startSync() {
    if (!state.inWord || !state.syncOn) return;
    if (state.mode === 'viewport') loop();
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  async function loop() {
    if (state.loopRunning) return;
    state.loopRunning = true;
    try {
      while (state.syncOn && state.mode === 'viewport') {
        var t0 = performance.now();
        if (!document.hidden && state.mapper) {
          try {
            var obs = await pollOnce();
            state.pollErrors = 0;
            handleWordObservation(obs);
          } catch (e) {
            state.pollErrors++;
            console.warn('poll failed', e);
            if (state.pollErrors >= 3) {
              state.mode = 'selection';
              updateStatus();
              break;
            }
          }
        }
        await sleep(Math.max(60, POLL_MS - (performance.now() - t0)));
      }
    } finally {
      state.loopRunning = false;
    }
  }

  async function pollOnce() {
    var wantPages = !!(state.main && state.main.unit === 'page') && !state.pagesInViewportBroken;
    try {
      return await wordRun(async function (ctx) {
        var w = ctx.document.activeWindow;
        w.load('verticalPercentScrolled');
        var vp = null;
        if (wantPages) { vp = w.activePane.pagesEnclosingViewport; vp.load('items/index'); }
        await ctx.sync();
        return { s: w.verticalPercentScrolled, pages: vp ? vp.items.map(function (p) { return p.index; }) : null };
      });
    } catch (e) {
      if (!wantPages) throw e;
      state.pagesInViewportBroken = true; // retry once without the page list
      return pollOnce();
    }
  }

  function handleWordObservation(obs) {
    var s = obs && obs.s;
    if (typeof s !== 'number' || isNaN(s)) return;
    var now = performance.now();
    observeScroll(s, obs.pages);
    var key = s + '|' + (obs.pages ? obs.pages[0] : '');
    var changed = state.lastKey !== key;
    state.lastKey = key;
    state.lastS = s;
    if (!changed && !state.forceFollow) return;
    if (now < state.paneUserUntil) return;
    if (now - state.selSnapAt < 1200) return; // the pane was just placed on the cursor
    if (!state.forceFollow && state.wordSetValue !== null && now - state.wordSetAt < OWN_WRITE_MS &&
        Math.abs(s - state.wordSetValue) <= 1) return;
    state.forceFollow = false;
    followPaneTo(viewportTopPos(s, obs.pages), REF_TOP);
  }

  function followPaneTo(c, offset) {
    if (!state.mapper) return;
    var px = state.mapper.toB(c);
    animatePaneTo(px - offset);
    markCurrent(px);
    return px;
  }

  /* ---------------- pane scrolling ---------------- */

  var anim = { target: 0, raf: 0, then: null };

  function animatePaneTo(target, then) {
    var max = Math.max(0, el.scroller.scrollHeight - el.scroller.clientHeight);
    anim.target = Math.max(0, Math.min(max, target));
    anim.then = then || null;
    if (!anim.raf) anim.raf = requestAnimationFrame(animStep);
  }

  function animStep() {
    anim.raf = 0;
    if (performance.now() < state.paneUserUntil) { anim.then = null; return; }
    var cur = el.scroller.scrollTop;
    var d = anim.target - cur;
    if (Math.abs(d) < 1) {
      markProgrammatic();
      el.scroller.scrollTop = anim.target;
      var t = anim.then; anim.then = null;
      if (t) t();
      return;
    }
    markProgrammatic();
    el.scroller.scrollTop = cur + (Math.abs(d) > 1500 ? d * 0.5 : d * 0.25);
    if (Math.abs(el.scroller.scrollTop - cur) < 0.5) { // could not move (bounds or rounding)
      var t2 = anim.then; anim.then = null;
      if (t2) t2();
      return;
    }
    anim.raf = requestAnimationFrame(animStep);
  }

  function markProgrammatic() { state.progScrollUntil = performance.now() + 150; }

  function markUser() {
    state.paneUserUntil = performance.now() + USER_IDLE_MS;
    if (anim.raf) { cancelAnimationFrame(anim.raf); anim.raf = 0; anim.then = null; }
  }

  function onPaneScroll() {
    if (!state.mapper) return;
    var now = performance.now();
    // A scroll we did not cause ourselves counts as the user's, even without a
    // wheel or key event (for example, dragging the scroll bar).
    if (now >= state.paneUserUntil) {
      if (now < state.progScrollUntil) return;
      state.paneUserUntil = now + USER_IDLE_MS;
    }
    markCurrent(el.scroller.scrollTop + REF_TOP);
    if (state.syncOn && state.mode === 'viewport') schedulePush();
  }

  function schedulePush() {
    if (state.pushTimer) return;
    state.pushTimer = setTimeout(function () {
      state.pushTimer = 0;
      pushToWord();
    }, PUSH_THROTTLE_MS);
  }

  async function pushToWord() {
    if (!state.syncOn || state.mode !== 'viewport' || !state.mapper || !state.main) return;
    var pos = state.mapper.toA(el.scroller.scrollTop + REF_TOP);
    var s = posToPercent(pos);
    if (state.wordSetValue === s && state.lastS === s) return;
    try {
      await wordRun(async function (ctx) {
        ctx.document.activeWindow.verticalPercentScrolled = s;
        await ctx.sync();
      });
      state.wordSetValue = s;
      state.wordSetAt = performance.now();
      state.lastS = s;
    } catch (e) {
      console.warn('scroll write failed', e);
    }
  }

  /* ---------------- selection mode ---------------- */

  function registerSelectionHandler() {
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
        if (!state.syncOn || state.mode === 'none') return;
        clearTimeout(state.selTimer);
        state.selTimer = setTimeout(followSelection, 250);
      });
    } catch (e) {
      console.warn('selection events unavailable', e);
    }
  }

  // Position (main-document units) of the cursor in Word.
  async function selectionPos() {
    var m = state.main;
    if (m.idIndex) {
      var id = await wordRun(async function (ctx) {
        var p = ctx.document.getSelection().paragraphs.getFirst();
        p.load('uniqueLocalId');
        await ctx.sync();
        return p.uniqueLocalId;
      });
      if (m.idIndex.has(id)) return m.posOfPara(m.idIndex.get(id));
    }
    var n = await wordRun(async function (ctx) {
      var sel = ctx.document.getSelection();
      var r = ctx.document.body.getRange('Start').expandTo(sel.getRange('Start'));
      r.load('text');
      await ctx.sync();
      return (r.text || '').length;
    });
    return m.cToPos(n);
  }

  // Selection mode: the pane always follows the cursor.
  // Viewport mode: when the cursor moves to another paragraph, the matching
  // paragraph in the pane is highlighted, and brought into view if it is not visible.
  async function followSelection() {
    if (!state.mapper || !state.main || !state.caps.api13) return;
    try {
      var pos = await selectionPos();
      var px = state.mapper.toB(pos);
      var idx = paneBlockAt(px);
      var viewport = state.mode === 'viewport';
      if (viewport && idx === state.lastSelBlock) return;
      state.lastSelBlock = idx;
      var h = el.scroller.clientHeight, top = el.scroller.scrollTop;
      var visible = px >= top + h * 0.08 && px <= top + h * 0.85;
      if (!viewport || !visible) {
        state.selSnapAt = performance.now();
        state.paneUserUntil = 0;
        animatePaneTo(px - h * 0.3);
      }
      markCurrent(px);
      flashBlock(idx);
    } catch (e) {
      console.warn('selection follow failed', e); // e.g. cursor in a footnote or header
    }
  }

  function paneBlockAt(px) {
    var tops = state.paneTops, lo = 0, hi = tops.length - 1, ans = 0;
    while (lo <= hi) {
      var mid = (lo + hi) >> 1;
      if (tops[mid] <= px + 2) { ans = mid; lo = mid + 1; } else hi = mid - 1;
    }
    return ans;
  }

  function flashBlock(idx) {
    var b = state.pane && state.pane.blocks[idx];
    if (!b) return;
    b.el.classList.remove('pr-flash');
    void b.el.offsetWidth;
    b.el.classList.add('pr-flash');
  }

  /* ---------------- clicks in the pane ---------------- */

  function onDocClick(e) {
    var a = e.target.closest && e.target.closest('a[href]');
    if (!a) return;
    var href = a.getAttribute('href') || '';
    if (href.charAt(0) === '#') {
      e.preventDefault();
      var id = href.slice(1);
      var t = el.doc.querySelector('[id="' + id.replace(/"/g, '') + '"]');
      if (t) { markUser(); el.scroller.scrollTop = paneTop(t) - 24; }
      return;
    }
    if (/^https?:/i.test(href)) {
      e.preventDefault();
      try { Office.context.ui.openBrowserWindow(href); } catch (err) { window.open(href, '_blank', 'noopener'); }
    }
  }

  function onDocDoubleClick(e) {
    if (!state.mapper || !state.pane) return;
    var node = e.target.closest('h1, h2, h3, h4, h5, h6, p, li, .pr-table-wrap');
    if (!node || !el.doc.contains(node)) return;
    var top = paneTop(node);
    jumpBothToPaneBlock(node, state.mapper.toA(top));
  }

  // Scrolls the pane so that `node` sits at the alignment line, and moves Word to
  // the corresponding place (scroll in viewport mode, cursor in selection mode).
  function jumpBothToPaneBlock(node, pos) {
    var top = paneTop(node);
    if (state.mode === 'viewport' && state.syncOn) {
      state.paneUserUntil = 0;
      animatePaneTo(top - REF_TOP, function () { pushToWord(); });
      state.wordSetValue = null;
      // write immediately as well, so Word does not wait for the animation
      var s = posToPercent(pos);
      wordRun(async function (ctx) { ctx.document.activeWindow.verticalPercentScrolled = s; await ctx.sync(); })
        .then(function () { state.wordSetValue = s; state.wordSetAt = performance.now(); state.lastS = s; })
        .catch(function () { /* ignore */ });
    } else {
      state.selSnapAt = performance.now();
      animatePaneTo(top - el.scroller.clientHeight * 0.3);
      jumpWordToPos(pos);
    }
  }

  async function jumpWordToPos(pos) {
    if (!state.main || !state.inWord) return;
    var blocks = state.main.blocks;
    var target = blocks[0];
    for (var i = 0; i < blocks.length; i++) { if (blocks[i].pos <= pos + 1e-9) target = blocks[i]; else break; }
    if (!target) return;
    try {
      await wordRun(async function (ctx) {
        var ps = ctx.document.body.paragraphs;
        ps.load('items/text');
        await ctx.sync();
        var p = ps.items[target.paraIndex];
        if (p) { p.select('Start'); await ctx.sync(); }
      });
    } catch (e) {
      console.warn('jump failed', e);
    }
  }

  /* ================================================================== */
  /* IndexedDB: remembers the imported file per main document            */
  /* ================================================================== */

  var IDB_NAME = 'parallel-reader', IDB_STORE = 'files';

  function idbOpen() {
    return new Promise(function (resolve, reject) {
      if (!window.indexedDB) { reject(new Error('IndexedDB unavailable')); return; }
      var r = indexedDB.open(IDB_NAME, 1);
      r.onupgradeneeded = function () { r.result.createObjectStore(IDB_STORE); };
      r.onsuccess = function () { resolve(r.result); };
      r.onerror = function () { reject(r.error); };
    });
  }

  function idbOp(mode, fn) {
    return idbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(IDB_STORE, mode);
        var req = fn(tx.objectStore(IDB_STORE));
        tx.oncomplete = function () { db.close(); resolve(req ? req.result : undefined); };
        tx.onerror = function () { db.close(); reject(tx.error); };
      });
    });
  }

  function idbGet(key) { return idbOp('readonly', function (st) { return st.get(key); }); }
  function idbPut(key, val) { return idbOp('readwrite', function (st) { return st.put(val, key); }); }
  function idbDel(key) { return idbOp('readwrite', function (st) { return st.delete(key); }); }

  // exposed for tests
  window.__PR = { state: state, rescanMain: rescanMain, realign: realign, viewportTopPos: viewportTopPos, posToPercent: posToPercent, dHat: function () { return dHat(); },
    paraPos: function (i) { return state.main.posOfPara(i); } };
})();
