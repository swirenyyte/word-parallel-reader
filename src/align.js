/*
 * align.js — anchor detection and alignment for the Word parallel reader.
 *
 * The module has no dependency on Office.js or the DOM, so the same code runs
 * in the task pane and in Node unit tests.
 *
 * Pipeline:
 *   1. classifyBlocks(): decide which blocks (paragraphs, tables) are anchors:
 *      headings, section keywords such as 摘要/Abstract, and figure/table captions.
 *   2. assignPaths(): turn heading numbers such as "一、", "（二）", "3.2" into
 *      comparable dotted paths, so that "（二）" under "三、" matches "3.2".
 *   3. alignAnchors(): monotonic dynamic programming that pairs anchors of the
 *      two documents by shared keywords, numbers, captions, heading level and
 *      relative position.
 *   4. alignSegments(): between two matched anchors, pair paragraphs by length
 *      (a simplified Gale–Church alignment), so scrolling also lines up inside
 *      long sections.
 *   5. Mapper: piecewise-linear mapping between a position in document A and a
 *      position in document B, in both directions.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PRAlign = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ */
  /* Text helpers                                                        */
  /* ------------------------------------------------------------------ */

  var CJK_RE = /[㐀-䶿一-鿿豈-﫿぀-ヿ가-힯]/;

  // Convert full-width digits, letters and common punctuation to half-width,
  // collapse whitespace. Used for matching only, never for display.
  function normalizeForMatch(s) {
    if (!s) return '';
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c >= 0xff01 && c <= 0xff5e) {
        var h = String.fromCharCode(c - 0xfee0);
        // keep full-width brackets as they are recognised separately below
        out += h;
      } else if (c === 0x3000 || c === 0x00a0 || c === 0x2002 || c === 0x2003 || c === 0x2009) {
        out += ' ';
      } else {
        out += s[i];
      }
    }
    return out.replace(/[​‌‍﻿]/g, '').replace(/\s+/g, ' ').trim();
  }

  function isCJK(ch) {
    return CJK_RE.test(ch);
  }

  // Rough "amount of content" of a text, comparable between Chinese and English
  // after each document is normalised by its own total. A CJK character counts
  // as 1; a Latin letter, digit or symbol counts as about one third.
  function textWeight(s) {
    if (!s) return 0;
    var w = 0;
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === ' ') continue;
      if (isCJK(ch)) w += 1;
      else w += 0.34;
    }
    return w;
  }

  function cnNumeral(s) {
    if (/^\d+$/.test(s)) return parseInt(s, 10);
    var map = { '零': 0, '〇': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
    if (s.indexOf('百') >= 0) return null; // headings never go that far
    var idx = s.indexOf('十');
    if (idx >= 0) {
      var tens = idx === 0 ? 1 : map[s.slice(0, idx)];
      var ones = idx === s.length - 1 ? 0 : map[s.slice(idx + 1)];
      if (tens === undefined || ones === undefined) return null;
      return tens * 10 + ones;
    }
    if (s.length === 1 && map[s] !== undefined) return map[s];
    return null;
  }

  function romanToInt(s) {
    var v = { I: 1, V: 5, X: 10, L: 50 };
    var total = 0;
    for (var i = 0; i < s.length; i++) {
      var cur = v[s[i]], next = v[s[i + 1]];
      if (cur === undefined) return null;
      total += next && next > cur ? -cur : cur;
    }
    return total || null;
  }

  /* ------------------------------------------------------------------ */
  /* Heading numbers                                                     */
  /* ------------------------------------------------------------------ */

  // Returns { scheme, parts, rest } or null.
  //   scheme: 'arabic' (1 / 1.2 / 1.2.3), 'zh' (一、), 'pzh' (（一）),
  //           'par' ((1)), 'chapter' (第一章), 'section' (第一节), 'roman' (II.)
  function parseNumbering(text) {
    var t = normalizeForMatch(text);
    var m, n;

    m = t.match(/^第\s*([一二三四五六七八九十零〇两\d]+)\s*([章节部分篇])\s*[:.、]?\s*/);
    if (m) {
      n = cnNumeral(m[1]);
      if (n !== null) {
        return { scheme: m[2] === '节' ? 'section' : 'chapter', parts: [n], rest: t.slice(m[0].length).trim() };
      }
    }
    m = t.match(/^([一二三四五六七八九十零〇两]{1,3})\s*[、.]\s*/);
    if (m) {
      n = cnNumeral(m[1]);
      if (n !== null) return { scheme: 'zh', parts: [n], rest: t.slice(m[0].length).trim() };
    }
    m = t.match(/^[(（]\s*([一二三四五六七八九十零〇两]{1,3})\s*[)）]\s*[、.]?\s*/);
    if (m) {
      n = cnNumeral(m[1]);
      if (n !== null) return { scheme: 'pzh', parts: [n], rest: t.slice(m[0].length).trim() };
    }
    m = t.match(/^[(（]\s*(\d{1,2})\s*[)）]\s*[、.]?\s*/);
    if (m) {
      return { scheme: 'par', parts: [parseInt(m[1], 10)], rest: t.slice(m[0].length).trim() };
    }
    // 1 / 1. / 1.2 / 1.2. / 1、  followed by text
    m = t.match(/^(\d{1,2}(?:\.\d{1,2}){0,4})(?:\.|、|\s)\s*(?=[^\d\s.%])/);
    if (m) {
      var parts = m[1].split('.').map(function (x) { return parseInt(x, 10); });
      var ok = parts.every(function (p) { return p > 0 && p < 60; });
      // "2019年" never reaches here (more than two digits); "3.5 km" is rejected
      // later by the short-heading test because it would need to look like a title.
      if (ok) return { scheme: 'arabic', parts: parts, rest: t.slice(m[0].length).trim() };
    }
    m = t.match(/^(X{0,3}(?:IX|IV|V?I{0,3}))[.)]\s+(?=\S)/);
    if (m && m[1]) {
      n = romanToInt(m[1]);
      if (n) return { scheme: 'roman', parts: [n], rest: t.slice(m[0].length).trim() };
    }
    return null;
  }

  /* ------------------------------------------------------------------ */
  /* Section keywords                                                    */
  /* ------------------------------------------------------------------ */

  var KEYWORDS = {
    abstract: ['abstract', 'summary', '摘要', '内容摘要', '内容提要', '提要', '中文摘要', '英文摘要'],
    keywords: ['keywords', 'keyword', 'key words', 'key word', '关键词', '关键字'],
    introduction: ['introduction', 'intro', '引言', '导言', '前言', '绪论', '导论', '引论', '绪言', '问题的提出', '问题提出'],
    background: ['background', 'research background', 'study background', '背景', '研究背景', '研究缘起'],
    literature: ['literature review', 'literature', 'review of the literature', 'review of literature', 'related work', 'related works',
      'related research', 'previous studies', 'previous research', 'prior research', 'state of the art',
      '文献综述', '文献回顾', '相关研究', '研究综述', '既有研究', '已有研究', '文献评述', '研究现状', '国内外研究现状', '研究进展'],
    theory: ['theoretical framework', 'theory', 'theories', 'conceptual framework', 'analytical framework', 'theoretical background',
      'theoretical perspective', 'framework', '理论框架', '理论基础', '分析框架', '概念框架', '理论视角'],
    methods: ['method', 'methods', 'methodology', 'methodologies', 'research method', 'research methods', 'research design',
      'materials and methods', 'data and methods', 'methods and data', 'data and methodology', 'approach',
      '研究方法', '方法', '方法论', '研究设计', '数据与方法', '研究数据与方法', '材料与方法', '研究方法与数据'],
    data: ['data', 'data sources', 'data source', 'study area', 'research area', 'case study', 'case studies', 'study site', 'the case',
      '数据', '数据来源', '研究区域', '研究区概况', '研究区域概况', '案例', '案例研究', '研究案例', '案例介绍', '研究对象'],
    results: ['results', 'result', 'findings', 'finding', 'empirical results', 'analysis', 'analyses', 'empirical analysis',
      '结果', '研究结果', '研究发现', '发现', '实证结果', '实证分析', '分析结果', '结果分析', '分析'],
    discussion: ['discussion', 'discussions', '讨论', '分析与讨论'],
    conclusion: ['conclusion', 'conclusions', 'concluding remarks', 'concluding', 'summary and conclusions',
      '结论', '结语', '总结', '结论与展望', '余论', '结束语'],
    acknowledgements: ['acknowledgements', 'acknowledgments', 'acknowledgement', 'acknowledgment', '致谢', '鸣谢', '谢辞'],
    references: ['references', 'reference', 'bibliography', 'works cited', 'literature cited', '参考文献', '参考资料', '引用文献'],
    appendix: ['appendix', 'appendices', 'supplementary material', 'supplementary materials', 'supplementary information', '附录', '附件', '补充材料'],
    notes: ['notes', 'endnotes', 'footnotes', '注释', '尾注'],
    funding: ['funding', 'funding information', '基金项目', '资助项目'],
    limitations: ['limitations', 'limitation', '研究局限', '局限性', '研究不足', '不足与展望'],
    implications: ['implications', 'policy implications', '政策启示', '启示', '政策建议']
  };

  var KW_LOOKUP = (function () {
    var map = Object.create(null);
    Object.keys(KEYWORDS).forEach(function (key) {
      KEYWORDS[key].forEach(function (phrase) { map[phrase] = key; });
    });
    return map;
  })();

  function cleanKeywordText(s) {
    return normalizeForMatch(s)
      .toLowerCase()
      .replace(/[:：.。,，;；!！?？"'“”‘’()（）\[\]【】]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Returns a Set of canonical section keys found in a heading text.
  function keywordKeys(text) {
    var keys = new Set();
    var t = cleanKeywordText(text);
    if (!t) return keys;
    if (KW_LOOKUP[t]) { keys.add(KW_LOOKUP[t]); return keys; }
    var compact = t.replace(/ /g, '');
    if (CJK_RE.test(compact) && KW_LOOKUP[compact]) { keys.add(KW_LOOKUP[compact]); return keys; }
    if (/^(appendix|appendices)\b/.test(t) || /^附录/.test(compact)) { keys.add('appendix'); return keys; }
    // "Results and discussion", "结论与讨论", "Data & methods"
    var parts = t.split(/\s+and\s+|\s*&\s*|\s*\/\s*|\s*,\s*|与|和|及|以及|、/);
    if (parts.length > 1) {
      parts.forEach(function (p) {
        var q = p.trim();
        var qc = q.replace(/ /g, '');
        if (KW_LOOKUP[q]) keys.add(KW_LOOKUP[q]);
        else if (KW_LOOKUP[qc]) keys.add(KW_LOOKUP[qc]);
      });
    }
    return keys;
  }

  // "摘要：本文……" or "Abstract: This paper ..." in a single paragraph.
  function inlineLabel(text) {
    var t = normalizeForMatch(text);
    var m = t.match(/^(摘\s*要|abstract|关键词|关键字|key\s*words?)\s*[:：]/i);
    if (!m) return null;
    var k = m[1].replace(/\s+/g, '').toLowerCase();
    return (k === '摘要' || k === 'abstract') ? 'abstract' : 'keywords';
  }

  /* ------------------------------------------------------------------ */
  /* Captions                                                            */
  /* ------------------------------------------------------------------ */

  function captionKey(text) {
    var t = normalizeForMatch(text);
    var m = t.match(/^(?:figure|fig\.?)\s*(\d+(?:[.\-–]\d+)?)/i) || t.match(/^图\s*(\d+(?:[.\-–]\d+)?)/);
    if (m) return 'fig:' + m[1].replace(/[\-–]/, '.');
    m = t.match(/^(?:table|tab\.)\s*(\d+(?:[.\-–]\d+)?)/i) || t.match(/^表\s*(\d+(?:[.\-–]\d+)?)/);
    if (m) return 'tab:' + m[1].replace(/[\-–]/, '.');
    return null;
  }

  /* ------------------------------------------------------------------ */
  /* Block classification                                                */
  /* ------------------------------------------------------------------ */

  function isShortHeading(text) {
    var t = normalizeForMatch(text);
    if (!t) return false;
    if (/[。.!?！？;；,，]$/.test(t)) return false;
    var cjk = 0;
    for (var i = 0; i < t.length; i++) if (isCJK(t[i])) cjk++;
    if (cjk > 0) return t.length <= 40;
    var words = t.split(/\s+/).length;
    return words <= 14 && t.length <= 120;
  }

  /*
   * Each input block:
   *   { text, styleLevel (0 = title, 1..9 heading level, null otherwise),
   *     isBold, isCaptionStyle, inTable, listString }
   * Sets block.anchor to an object or null.
   */
  function classifyBlock(b) {
    var t = normalizeForMatch(b.text || '');
    if (!t || b.inTable) return null;

    var cap = captionKey(t);
    if (cap && (b.isCaptionStyle || t.length <= 300)) {
      return { kind: 'caption', capKey: cap, level: null, kw: new Set(), num: null, strong: !!b.isCaptionStyle };
    }

    var hasStyle = b.styleLevel !== null && b.styleLevel !== undefined;
    var src = t;
    if (hasStyle && b.listString) src = normalizeForMatch(b.listString) + ' ' + t;
    var num = parseNumbering(src);
    var rest = num ? num.rest : t;
    var kw = keywordKeys(rest);

    if (hasStyle) {
      return { kind: b.styleLevel === 0 ? 'title' : 'heading', level: b.styleLevel, kw: kw, num: num, strong: true };
    }

    var label = inlineLabel(t);
    if (label) return { kind: 'label', level: 1, kw: new Set([label]), num: null, strong: false };

    if (!isShortHeading(rest) || (num && !rest)) return null;
    if (kw.size) return { kind: 'heading', level: null, kw: kw, num: num, strong: false };
    if (num) return { kind: 'heading', level: null, kw: kw, num: num, strong: false };
    if (b.isBold && rest.length <= 60) return { kind: 'heading', level: null, kw: kw, num: null, strong: false, weak: true };
    return null;
  }

  // Converts heading numbers to dotted paths and fills in missing levels.
  function assignPaths(blocks) {
    var anchors = blocks.filter(function (b) { return b.anchor && (b.anchor.kind === 'heading' || b.anchor.kind === 'title'); });
    var usesZh = anchors.some(function (b) {
      var n = b.anchor.num;
      return n && (n.scheme === 'zh' || n.scheme === 'chapter' || n.scheme === 'pzh');
    });
    var cur = [];
    anchors.forEach(function (b) {
      var a = b.anchor, n = a.num, path = null;
      if (n) {
        var v = n.parts[0];
        switch (n.scheme) {
          case 'arabic':
            if (n.parts.length > 1) path = n.parts.slice();
            else path = usesZh ? [cur[0] || 0, cur[1] || 0, v] : [v];
            break;
          case 'zh': case 'chapter': case 'roman':
            path = [v]; break;
          case 'section': case 'pzh':
            path = [cur[0] || 0, v]; break;
          case 'par':
            path = usesZh ? [cur[0] || 0, cur[1] || 0, cur[2] || 0, v] : [cur[0] || 0, v]; break;
        }
      }
      if (path) {
        cur = path.slice();
        a.path = path.join('.');
        if (a.level === null || a.level === undefined) a.level = path.length;
      } else {
        a.path = null;
      }
      if (a.level === null || a.level === undefined) {
        a.level = a.kw.size ? 1 : (a.weak ? 3 : 2);
      }
    });
  }

  function classifyBlocks(blocks) {
    blocks.forEach(function (b) { b.anchor = classifyBlock(b); });
    assignPaths(blocks);
    return blocks;
  }

  /* ------------------------------------------------------------------ */
  /* Positions                                                           */
  /* ------------------------------------------------------------------ */

  // Adds b.weight and b.wpos (0..1, weight-based position of the block start).
  function addWeights(blocks) {
    var total = 0;
    blocks.forEach(function (b) {
      b.weight = b.weight !== undefined ? b.weight : textWeight(b.text);
      total += b.weight;
    });
    var acc = 0;
    blocks.forEach(function (b) {
      b.wpos = total > 0 ? acc / total : 0;
      acc += b.weight;
    });
    return total;
  }

  /* ------------------------------------------------------------------ */
  /* Anchor alignment                                                    */
  /* ------------------------------------------------------------------ */

  var NEG = -Infinity;

  function shares(setA, setB) {
    var hit = false;
    setA.forEach(function (k) { if (setB.has(k)) hit = true; });
    return hit;
  }

  function anchorScore(A, B) {
    var a = A.anchor, b = B.anchor;
    var dpos = Math.abs(A.wpos - B.wpos);
    if (a.kind === 'caption' || b.kind === 'caption') {
      if (a.kind !== b.kind || a.capKey !== b.capKey) return NEG;
      var sc = 9 - 4 * dpos;
      return sc > 0 ? sc : NEG;
    }
    if (a.kind === 'title' || b.kind === 'title') {
      if (a.kind !== b.kind) return NEG;
      return 5 - 6 * dpos > 0 ? 5 - 6 * dpos : NEG;
    }
    var s = 0;
    var kwShare = a.kw.size && b.kw.size && shares(a.kw, b.kw);
    if (a.kw.size && b.kw.size && !kwShare) return NEG;
    if (kwShare) s += 8;
    if (a.path && b.path) {
      if (a.path === b.path) s += 6;
      else if (!kwShare) return NEG;
    }
    if (s === 0) {
      // Structure only: same level, and not contradicted by a keyword on one side.
      if (a.level !== b.level) return NEG;
      if (a.kw.size || b.kw.size) return NEG;
      s = a.strong && b.strong ? 2.2 : (a.weak || b.weak ? 0.9 : 1.4);
      s -= 8 * dpos;
    } else {
      s -= 4 * dpos;
    }
    return s > 0 ? s : NEG;
  }

  // Monotonic best-score pairing of anchor blocks. Returns [{a: idxA, b: idxB, score}]
  // with indices into the blocks arrays.
  function alignAnchors(blocksA, blocksB) {
    var A = [], B = [];
    blocksA.forEach(function (b, i) { if (b.anchor) A.push(i); });
    blocksB.forEach(function (b, i) { if (b.anchor) B.push(i); });
    var n = A.length, m = B.length;
    var W = m + 1;
    var dp = new Float64Array((n + 1) * W);
    var bt = new Uint8Array((n + 1) * W); // 0 none, 1 up, 2 left, 3 diag
    for (var i = 1; i <= n; i++) {
      for (var j = 1; j <= m; j++) {
        var best = dp[(i - 1) * W + j], move = 1;
        var left = dp[i * W + j - 1];
        if (left > best) { best = left; move = 2; }
        var sc = anchorScore(blocksA[A[i - 1]], blocksB[B[j - 1]]);
        if (sc !== NEG) {
          var d = dp[(i - 1) * W + j - 1] + sc;
          if (d > best) { best = d; move = 3; }
        }
        dp[i * W + j] = best;
        bt[i * W + j] = move;
      }
    }
    var pairs = [];
    i = n; j = m;
    while (i > 0 && j > 0) {
      var mv = bt[i * W + j];
      if (mv === 3) {
        pairs.push({ a: A[i - 1], b: B[j - 1], score: anchorScore(blocksA[A[i - 1]], blocksB[B[j - 1]]) });
        i--; j--;
      } else if (mv === 1) i--;
      else j--;
    }
    pairs.reverse();
    return pairs;
  }

  /* ------------------------------------------------------------------ */
  /* Paragraph alignment inside sections (simplified Gale–Church)        */
  /* ------------------------------------------------------------------ */

  var GC_VAR = 12;
  var PRIOR = { m11: 0.89, m10: 0.0099 / 2, m21: 0.089 / 2, m22: 0.011 };

  function phi(x) {
    // standard normal CDF via Abramowitz–Stegun erf approximation
    var t = 1 / (1 + 0.3275911 * Math.abs(x) / Math.SQRT2);
    var y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t *
      Math.exp(-(x * x) / 2);
    return x >= 0 ? (1 + y) / 2 : (1 - y) / 2;
  }

  function lenCost(la, lb) {
    if (la <= 0 && lb <= 0) return 0;
    var mean = (la + lb) / 2;
    var delta = (lb - la) / Math.sqrt(GC_VAR * Math.max(mean, 1));
    var p = 2 * (1 - phi(Math.abs(delta)));
    return -Math.log(Math.max(p, 1e-12));
  }

  // wa, wb: arrays of weights in comparable units. Returns groups
  // [{i, j, di, dj}] where (i, j) are start offsets and (di, dj) group sizes.
  function galeChurch(wa, wb) {
    var n = wa.length, m = wb.length;
    var W = m + 1;
    var INF = 1e18;
    var D = new Float64Array((n + 1) * W).fill(INF);
    var B = new Int8Array((n + 1) * W * 2);
    D[0] = 0;
    var moves = [[1, 1, PRIOR.m11], [1, 0, PRIOR.m10], [0, 1, PRIOR.m10], [2, 1, PRIOR.m21], [1, 2, PRIOR.m21], [2, 2, PRIOR.m22]];
    for (var i = 0; i <= n; i++) {
      for (var j = 0; j <= m; j++) {
        if (i === 0 && j === 0) continue;
        var best = INF, bi = 0, bj = 0;
        for (var k = 0; k < moves.length; k++) {
          var di = moves[k][0], dj = moves[k][1];
          if (i < di || j < dj) continue;
          var prev = D[(i - di) * W + (j - dj)];
          if (prev >= INF) continue;
          var la = 0, lb = 0, q;
          for (q = 1; q <= di; q++) la += wa[i - q];
          for (q = 1; q <= dj; q++) lb += wb[j - q];
          var c = prev - Math.log(moves[k][2]) + lenCost(la, lb);
          if (c < best) { best = c; bi = di; bj = dj; }
        }
        D[i * W + j] = best;
        B[(i * W + j) * 2] = bi;
        B[(i * W + j) * 2 + 1] = bj;
      }
    }
    var groups = [];
    i = n; j = m;
    while (i > 0 || j > 0) {
      var gi = B[(i * W + j) * 2], gj = B[(i * W + j) * 2 + 1];
      if (gi === 0 && gj === 0) break;
      groups.push({ i: i - gi, j: j - gj, di: gi, dj: gj });
      i -= gi; j -= gj;
    }
    groups.reverse();
    return groups;
  }

  // For every gap between consecutive anchor pairs, pair the paragraphs inside it.
  // Returns { pairs: [{a, b}], stats: {oneToOne, total} }.
  function alignSegments(blocksA, blocksB, anchorPairs, totalA, totalB) {
    var bounds = [{ a: -1, b: -1 }].concat(anchorPairs).concat([{ a: blocksA.length, b: blocksB.length }]);
    var out = [];
    var oneToOne = 0, groupsTotal = 0;
    for (var s = 0; s + 1 < bounds.length; s++) {
      var ia = [], ib = [];
      for (var x = bounds[s].a + 1; x < bounds[s + 1].a; x++) if (blocksA[x].weight > 0) ia.push(x);
      for (var y = bounds[s].b + 1; y < bounds[s + 1].b; y++) if (blocksB[y].weight > 0) ib.push(y);
      if (!ia.length || !ib.length) continue;
      if (ia.length * ib.length > 250000) continue; // pathological size; fall back to linear
      var wa = ia.map(function (k) { return blocksA[k].weight / totalA * 10000; });
      var wb = ib.map(function (k) { return blocksB[k].weight / totalB * 10000; });
      var sa = wa.reduce(function (p, c) { return p + c; }, 0);
      var sb = wb.reduce(function (p, c) { return p + c; }, 0);
      var S = (sa + sb) / 2;
      wa = wa.map(function (v) { return v * S / sa; });
      wb = wb.map(function (v) { return v * S / sb; });
      var groups = galeChurch(wa, wb);
      groups.forEach(function (g) {
        if (g.di > 0 && g.dj > 0) {
          groupsTotal++;
          if (g.di === 1 && g.dj === 1) oneToOne++;
          out.push({ a: ia[g.i], b: ib[g.j] });
        } else {
          groupsTotal++;
        }
      });
    }
    return { pairs: out, stats: { oneToOne: oneToOne, total: groupsTotal } };
  }

  /* ------------------------------------------------------------------ */
  /* Full alignment and mapping                                          */
  /* ------------------------------------------------------------------ */

  /*
   * blocksA / blocksB: arrays of blocks (see classifyBlock for fields).
   * Returns {
   *   anchorPairs: [{a, b, score}], blockPairs: [{a, b, kind}],
   *   stats: { anchorsA, anchorsB, anchorsMatched, paraOneToOne, paraGroups }
   * }
   */
  function alignDocuments(blocksA, blocksB) {
    classifyBlocks(blocksA);
    classifyBlocks(blocksB);
    var totalA = addWeights(blocksA);
    var totalB = addWeights(blocksB);
    var anchorPairs = alignAnchors(blocksA, blocksB);
    var seg = alignSegments(blocksA, blocksB, anchorPairs, totalA || 1, totalB || 1);
    var all = anchorPairs.map(function (p) { return { a: p.a, b: p.b, kind: 'anchor' }; })
      .concat(seg.pairs.map(function (p) { return { a: p.a, b: p.b, kind: 'para' }; }));
    all.sort(function (p, q) { return p.a - q.a || p.b - q.b; });
    // keep a strictly increasing chain in both indices
    var chain = [];
    all.forEach(function (p) {
      var last = chain[chain.length - 1];
      if (!last || (p.a > last.a && p.b > last.b)) chain.push(p);
    });
    return {
      anchorPairs: anchorPairs,
      blockPairs: chain,
      stats: {
        anchorsA: blocksA.filter(function (b) { return b.anchor; }).length,
        anchorsB: blocksB.filter(function (b) { return b.anchor; }).length,
        anchorsMatched: anchorPairs.length,
        paraOneToOne: seg.stats.oneToOne,
        paraGroups: seg.stats.total
      }
    };
  }

  /*
   * points: [{a, b}] in any order (a = position in document A, b = position in B).
   * Returns { toB(a), toA(b), points } using piecewise-linear interpolation over a
   * chain that is non-decreasing in both coordinates.
   */
  function makeMapper(points) {
    var pts = points.slice().sort(function (p, q) { return p.a - q.a || p.b - q.b; });
    var chain = [];
    pts.forEach(function (p) {
      if (!isFinite(p.a) || !isFinite(p.b)) return;
      var last = chain[chain.length - 1];
      if (!last) { chain.push(p); return; }
      if (p.a > last.a && p.b >= last.b) chain.push(p);
    });
    function interp(x, keyX, keyY) {
      var n = chain.length;
      if (!n) return x;
      if (n === 1) return chain[0][keyY];
      if (x <= chain[0][keyX]) return chain[0][keyY];
      if (x >= chain[n - 1][keyX]) return chain[n - 1][keyY];
      var lo = 0, hi = n - 1;
      while (hi - lo > 1) {
        var mid = (lo + hi) >> 1;
        if (chain[mid][keyX] <= x) lo = mid; else hi = mid;
      }
      var p = chain[lo], q = chain[hi];
      var span = q[keyX] - p[keyX];
      if (span <= 0) return p[keyY];
      return p[keyY] + (x - p[keyX]) / span * (q[keyY] - p[keyY]);
    }
    return {
      points: chain,
      toB: function (a) { return interp(a, 'a', 'b'); },
      // For the inverse, b is non-decreasing; ties are resolved to the first a.
      toA: function (b) { return interp(b, 'b', 'a'); }
    };
  }

  return {
    normalizeForMatch: normalizeForMatch,
    textWeight: textWeight,
    parseNumbering: parseNumbering,
    keywordKeys: keywordKeys,
    inlineLabel: inlineLabel,
    captionKey: captionKey,
    isShortHeading: isShortHeading,
    classifyBlock: classifyBlock,
    classifyBlocks: classifyBlocks,
    assignPaths: assignPaths,
    addWeights: addWeights,
    anchorScore: anchorScore,
    alignAnchors: alignAnchors,
    galeChurch: galeChurch,
    alignSegments: alignSegments,
    alignDocuments: alignDocuments,
    makeMapper: makeMapper,
    KEYWORDS: KEYWORDS
  };
});
