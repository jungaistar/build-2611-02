/* 엑셀·PDF 문서 정리 — 판정 엔진
 *
 * 화면(app.js)은 판정하지 않는다. 표 찾기·병합 채우기·머리글 합치기·열 짝짓기 제안·값 변환·
 * 결과 만들기는 전부 여기서 하고 test/logic.test.js 로 고정한다. 브라우저와 node 양쪽에서 읽힌다.
 *
 * 흐름
 *   시트(SheetJS 워크시트) 또는 PDF 글자 좌표  →  grid(2차원 배열) + merges
 *   →  extractTable : 병합 채우기 · 제목/정보 줄(meta) · 머리글(여러 줄 합치기) · 데이터 줄 · 버린 줄
 *   →  suggest      : 목표 열마다 원본 열 하나를 제안(이름 유사도 + 값 형태)
 *   →  사람이 화면에서 고치고 확정
 *   →  buildOutput  : 목표 표 구조로 옮기고 값마다 출처(문서·병합 채움·변환·계산·직접 입력)를 남긴다
 *
 * 값의 출처를 남기는 이유: 문서에 적혀 있던 값과 병합을 풀어 채운 값·형식을 바꾼 값·계산한 값은
 * 신뢰도가 다르다. 같은 모양으로 내보내면 사람이 구분하지 못한 채 그대로 믿게 된다.
 */
(function (root) {
  'use strict';
  var DAY = 86400000;

  /* ---------- 글자 다듬기 ---------- */
  function str(v) {
    if (v == null) return '';
    if (v instanceof Date) return fmtDate(v);
    return String(v).replace(/\s+/g, ' ').trim();
  }
  // 이름 비교용: 괄호 안(단위)·공백·구두점을 지우고 소문자로
  function norm(s) {
    return str(s).toLowerCase()
      .replace(/\([^)]*\)|\[[^\]]*\]|（[^）]*）/g, '')
      .replace(/[\s.,_\-\/'’"`#:：·*~]/g, '');
  }

  /* ---------- 같은 뜻 낱말 (열 이름이 달라도 읽게 하는 표) ----------
   * 입력 양식이 바뀌면 코드가 아니라 이 표를 늘린다. 첫 낱말이 묶음의 이름이다. */
  var ALIASES = [
    { name: '품번', type: 'code', words: ['품번', '부품번호', '품목코드', '품목번호', '자재코드', '자재번호', '도번', '도면번호', 'partno', 'partnumber', 'pn', 'itemcode', 'itemno', 'materialcode', '코드'] },
    { name: '품명', type: 'text', words: ['품명', '부품명', '품목명', '자재명', '제품명', '명칭', '내역', 'description', 'desc', 'partname', 'itemname', 'name'] },
    { name: '규격', type: 'text', words: ['규격', '사양', '치수', 'spec', 'specification', 'size', 'dimension'] },
    { name: '수량', type: 'number', words: ['수량', '발주수량', '소요량', '견적수량', '주문수량', 'qty', 'quantity'] },
    { name: '단위', type: 'text', words: ['단위', 'unit', 'uom'] },
    { name: '단가', type: 'number', words: ['단가', '견적단가', '공급단가', '발주단가', 'unitprice', 'price'] },
    { name: '금액', type: 'number', words: ['금액', '공급가액', '견적금액', '합계금액', 'amount', 'totalprice', 'total'] },
    { name: '공급업체', type: 'text', words: ['공급업체', '업체', '업체명', '협력사', '거래처', '제조사', 'vendor', 'supplier', 'maker'] },
    { name: '납기', type: 'date', words: ['납기', '납기일', '납품일', '입고예정일', 'duedate', 'deliverydate', 'delivery'] },
    { name: '비고', type: 'text', words: ['비고', '메모', '특기사항', 'remark', 'remarks', 'note', 'notes', 'comment'] }
  ];
  function aliasGroup(name) {
    var n = norm(name);
    if (!n) return null;
    for (var i = 0; i < ALIASES.length; i++) if (ALIASES[i].words.indexOf(n) >= 0 || norm(ALIASES[i].name) === n) return ALIASES[i];
    return null;
  }

  /* ---------- 날짜 ---------- */
  function fmtDate(d) {
    return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
  }
  function ymd(y, m, d) {
    if (y < 100) y += 2000;
    var t = new Date(Date.UTC(y, m - 1, d));
    return (t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d) ? t : null;
  }
  // 엑셀 일련번호 · Date · '2026-12-05' · '2026.12.5' · '26.12.05' · '2026년 12월 5일'
  function toDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    if (typeof v === 'number') {
      if (!isFinite(v) || v < 20000 || v > 80000) return null;
      return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * DAY);
    }
    var m = String(v).trim().match(/^(\d{4}|\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?\.?$/);
    return m ? ymd(+m[1], +m[2], +m[3]) : null;
  }

  /* ---------- 값 형태 ---------- */
  var NUM_JUNK = /[,\s₩$]|원|krw|ea|개|pcs|set/gi;
  function toNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = str(v);
    if (!s) return null;
    var neg = /^\(.*\)$/.test(s);                      // 회계식 음수 (1,200)
    s = s.replace(/^\(|\)$/g, '').replace(NUM_JUNK, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    var n = parseFloat(s);
    return neg ? -n : n;
  }
  function kindOf(v) {
    if (v == null || v === '') return 'empty';
    if (v instanceof Date) return 'date';
    if (typeof v === 'number') return 'number';
    var s = str(v);
    if (!s) return 'empty';
    if (toDate(s)) return 'date';
    if (toNumber(s) != null) return 'number';
    if (s.length <= 30 && !/\s/.test(s) && /[A-Za-z]/.test(s) && /\d/.test(s)) return 'code';
    return 'text';
  }
  function profile(values) {
    var p = { n: 0, number: 0, date: 0, code: 0, text: 0 };
    values.forEach(function (v) { var k = kindOf(v); if (k !== 'empty') { p.n++; p[k]++; } });
    if (p.n) ['number', 'date', 'code', 'text'].forEach(function (k) { p[k] = p[k] / p.n; });
    return p;
  }
  // 목표 형식에 값 모양이 얼마나 맞는가 (0~1)
  function typeFit(type, p) {
    if (!p || !p.n) return 0.5;                         // 값이 없으면 판단 보류
    if (type === 'number') return p.number;
    if (type === 'date') return p.date;
    if (type === 'code') return Math.min(1, p.code + 0.6 * p.number + 0.3 * p.text);
    return Math.min(1, p.text + 0.5 * p.code + 0.2 * p.number);
  }

  /* ---------- 이름 유사도 ---------- */
  function bigrams(s) { var a = []; for (var i = 0; i < s.length - 1; i++) a.push(s.slice(i, i + 2)); return a; }
  function dice(a, b) {
    if (!a || !b) return 0;
    if (a === b) return 1;
    var x = bigrams(a), y = bigrams(b).slice(), hit = 0;
    if (!x.length || !y.length) return 0;
    x.forEach(function (g) { var k = y.indexOf(g); if (k >= 0) { hit++; y.splice(k, 1); } });
    return 2 * hit / (x.length + bigrams(b).length);
  }
  // 원본 열 이름 → 목표 열 이름. { score, why }
  function nameScore(src, target) {
    var ns = norm(src), nt = norm(target);
    if (!ns || !nt) return { score: 0, why: '' };
    if (ns === nt) return { score: 1, why: '이름이 같음' };
    var g = aliasGroup(target);
    var words = g ? g.words.concat([norm(g.name)]) : [nt];
    if (words.indexOf(ns) >= 0) return { score: 0.92, why: '같은 뜻 낱말(' + (g ? g.name : target) + ')' };
    var part = words.filter(function (w) { return w.length >= 2 && (ns.indexOf(w) >= 0 || (ns.length >= 3 && w.indexOf(ns) >= 0)); })[0];
    if (part) return { score: 0.72, why: '이름 일부가 겹침(' + part + ')' };
    var d = Math.max.apply(null, words.map(function (w) { return dice(ns, w); }));
    return { score: 0.7 * d, why: d > 0 ? '글자 비슷함 ' + Math.round(d * 100) + '%' : '' };
  }

  /* ---------- 시트 → grid ---------- */
  function colIndex(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n - 1; }
  function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function decodeCell(a) { var m = /^([A-Z]+)(\d+)$/.exec(a); return m ? { c: colIndex(m[1]), r: +m[2] - 1 } : null; }
  // SheetJS 워크시트 객체를 직접 읽는다(라이브러리 호출 없이) — node 테스트에서도 같은 길로 돈다.
  function gridFromSheet(ws) {
    if (!ws || !ws['!ref']) return { grid: [], merges: [] };
    var ref = ws['!ref'].split(':'), s = decodeCell(ref[0]), e = decodeCell(ref[1] || ref[0]);
    var grid = [];
    for (var r = s.r; r <= e.r; r++) {
      var row = [];
      for (var c = s.c; c <= e.c; c++) {
        var cell = ws[colName(c) + (r + 1)];
        row.push(cell == null || cell.v == null ? '' : cell.v);
      }
      grid.push(row);
    }
    var merges = (ws['!merges'] || []).map(function (m) {
      return { s: { r: m.s.r - s.r, c: m.s.c - s.c }, e: { r: m.e.r - s.r, c: m.e.c - s.c } };
    });
    return { grid: grid, merges: merges };
  }
  function csvToGrid(text) {
    var rows = [], row = [], cell = '', q = false;
    text = String(text).replace(/^\uFEFF/, '');
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  /* ---------- 병합 채우기 ----------
   * 엑셀 병합 셀은 왼쪽 위 칸에만 값이 있고 나머지는 비어 있다. 그대로 읽으면
   * 「공급업체」가 첫 줄에만 남고 아래 줄은 빈칸이 된다. 범위 전체에 그 값을 채운다. */
  function fillMerges(grid, merges) {
    var g = grid.map(function (r) { return r.slice(); }), filled = {};
    (merges || []).forEach(function (m) {
      var v = (g[m.s.r] || [])[m.s.c];
      for (var r = m.s.r; r <= m.e.r; r++) {
        if (!g[r]) continue;
        for (var c = m.s.c; c <= m.e.c; c++) {
          if (r === m.s.r && c === m.s.c) continue;
          if (g[r][c] === '' || g[r][c] == null) { g[r][c] = v; if (v !== '' && v != null) filled[r + ',' + c] = true; }
        }
      }
    });
    return { grid: g, filled: filled };
  }

  /* ---------- 표 찾기 ---------- */
  function cellsOf(row) { return (row || []).map(str).filter(Boolean); }
  function distinct(a) { return a.filter(function (v, i) { return a.indexOf(v) === i; }); }
  // 병합 채움으로 옆 칸에 반복된 값은 한 번으로 센다
  function folded(row) { return cellsOf(row).filter(function (v, i, a) { return i === 0 || a[i - 1] !== v; }); }
  var PAIR_RE = /^[^:：]{1,20}\s*[:：]\s*\S/;
  function isHeaderLike(row) {
    var c = folded(row);
    if (c.length < 2 || distinct(c).length < 2) return false;
    if (c.filter(function (v) { return PAIR_RE.test(v); }).length * 2 >= c.length) return false;   // 「업체명 : 가람정밀」 같은 정보 줄
    var numeric = c.filter(function (v) { var k = kindOf(v); return k === 'number' || k === 'date'; }).length;
    return numeric / c.length <= 0.2;
  }
  function isUnitRow(row) {                             // (EA) (원) 같은 단위만 있는 줄
    var c = cellsOf(row);
    return c.length >= 1 && c.every(function (v) { return /^[([（].*[)\]）]$/.test(v); });
  }
  // 머리글 시작 줄과 줄 수를 찾는다. merges 는 병합 채우기 전 원본 기준.
  function detectHeader(grid, merges) {
    var lim = Math.min(grid.length, 30), widest = 0, i;
    for (i = 0; i < lim; i++) widest = Math.max(widest, folded(grid[i]).length);
    var start = -1;
    for (i = 0; i < Math.min(grid.length, 20); i++) {
      if (isHeaderLike(grid[i]) && folded(grid[i]).length >= Math.max(2, Math.ceil(widest * 0.6))) { start = i; break; }
    }
    if (start < 0) return { start: 0, count: 1, found: false };
    var count = 1;
    while (count < 3 && start + count < grid.length) {
      var r = start + count, linked = (merges || []).some(function (m) {
        // 머리글 안에서 시작한 병합이 이 줄까지 내려오거나, 바로 윗줄에 가로 병합이 있으면 아랫줄도 머리글이다
        var inHead = m.s.r >= start && m.s.r < r;
        return inHead && (m.e.r >= r || (m.s.r === r - 1 && m.e.c > m.s.c && m.e.r === m.s.r));
      });
      if (linked || isUnitRow(grid[r])) count++; else break;
    }
    return { start: start, count: count, found: true };
  }

  // 머리글 위 줄에서 제목과 「항목 : 값」 정보를 뽑는다
  function readMeta(rows) {
    var meta = [], title = '';
    rows.forEach(function (row) {
      var c = folded(row);
      var pairs = 0;
      for (var i = 0; i < c.length; i++) {
        var m = /^([^:：]{1,20})\s*[:：]\s*(.+)$/.exec(c[i]);
        if (m) { meta.push({ key: m[1].trim(), value: m[2].trim() }); pairs++; }
        else if (/[:：]$/.test(c[i]) && c[i + 1] != null) { meta.push({ key: c[i].replace(/[:：]$/, '').trim(), value: c[i + 1] }); i++; pairs++; }
      }
      if (!pairs && c.length === 1 && !title) title = c[0].replace(/\s+/g, ' ');
    });
    return { title: title, meta: meta };
  }

  var TOTAL_RE = /^(합\s*계|소\s*계|총\s*계|누\s*계|계|total|sub\s*total|grand\s*total)$/i;
  /* grid 한 장을 표로. opts: { merges, header: {start, count} (없으면 자동), keepTotals } */
  function extractTable(grid, opts) {
    opts = opts || {};
    var merges = opts.merges || [];
    var f = fillMerges(grid, merges), g = f.grid;
    var width = g.reduce(function (w, r) { for (var c = r.length - 1; c >= 0; c--) if (str(r[c])) return Math.max(w, c + 1); return w; }, 0);
    g = g.map(function (r) { var x = r.slice(0, width); while (x.length < width) x.push(''); return x; });
    var auto = detectHeader(g, merges);
    var h = opts.header ? { start: Math.max(0, opts.header.start), count: Math.max(1, Math.min(5, opts.header.count)), found: true, manual: true } : auto;
    var head = readMeta(g.slice(0, h.start));

    // 머리글 여러 줄 → 열 이름 하나 (위아래가 같으면 한 번만)
    var names = [], seen = {};
    for (var c = 0; c < width; c++) {
      var parts = [];
      for (var r = h.start; r < h.start + h.count && r < g.length; r++) {
        var v = str(g[r][c]);
        if (v && parts[parts.length - 1] !== v) parts.push(v);
      }
      var nm = parts.join(' ') || ('열 ' + colName(c));
      if (seen[nm]) { seen[nm]++; nm = nm + ' (' + seen[nm] + ')'; } else seen[nm] = 1;
      names.push(nm);
    }
    var headSet = names.map(norm);

    var rows = [], skipped = [];
    for (var ri = h.start + h.count; ri < g.length; ri++) {
      var row = g[ri], cells = cellsOf(row);
      if (!cells.length) continue;
      var text = folded(row).join(' | ').slice(0, 80);
      if (!opts.keepTotals && cells.some(function (v) { return TOTAL_RE.test(v); })) { skipped.push({ row: ri + 1, reason: '합계 줄', text: text }); continue; }
      var asHead = row.map(function (v) { return norm(v); });
      var same = asHead.filter(function (v, k) { return v && v === headSet[k]; }).length;
      if (same >= Math.max(2, Math.ceil(cells.length * 0.6))) { skipped.push({ row: ri + 1, reason: '머리글 반복', text: text }); continue; }
      if (width >= 3 && distinct(cells).length === 1) { skipped.push({ row: ri + 1, reason: '값이 한 칸뿐인 줄(제목·메모로 보임)', text: text }); continue; }
      rows.push({
        src: ri + 1,
        cells: row,
        filled: row.map(function (v, k) { return !!f.filled[ri + ',' + k]; })
      });
    }
    var columns = names.map(function (n, k) {
      var vals = rows.map(function (r) { return r.cells[k]; });
      return { key: 'c:' + n, name: n, index: k, letter: colName(k), profile: profile(vals),
        samples: vals.filter(function (v) { return str(v); }).slice(0, 3).map(str) };
    }).filter(function (col) { return col.profile.n || !/^열 [A-Z]+$/.test(col.name); });
    return { title: head.title, meta: head.meta, header: { start: h.start, count: h.count, found: h.found, manual: !!h.manual },
      autoHeader: { start: auto.start, count: auto.count }, columns: columns, rows: rows, skipped: skipped, width: width, grid: g, filled: f.filled };
  }

  /* ---------- PDF: 글자 좌표 → grid ----------
   * PDF 에는 표 구조가 없다. 글자와 그 좌표만 있다.
   * 1) y 가 비슷한 글자를 한 줄로 묶고  2) 줄 안에서 가까운 글자를 한 칸으로 붙이고
   * 3) 칸이 가장 많은 줄(대개 머리글)의 칸 가운데를 열 기준으로 삼아, 다른 줄의 칸을 가장 가까운 열에 넣는다.
   * 오른쪽 정렬 숫자는 시작 x 가 들쭉날쭉해서 시작점이 아니라 가운데로 견준다. */
  function pdfItemsToGrid(items, opts) {
    opts = opts || {};
    var yTol = opts.yTol || 3, gap = opts.gap || 4;
    var lines = [];
    (items || []).forEach(function (it) {
      var s = it.str == null ? '' : String(it.str);
      if (!s.trim()) return;
      var x = it.transform[4], y = it.transform[5], w = it.width != null ? it.width : s.length * Math.abs(it.transform[0] || 10) * 0.5;
      var line = null;
      for (var i = 0; i < lines.length; i++) if (Math.abs(lines[i].y - y) <= yTol) { line = lines[i]; break; }
      if (!line) { line = { y: y, parts: [] }; lines.push(line); }
      line.parts.push({ x: x, x2: x + w, s: s });
    });
    lines.sort(function (a, b) { return b.y - a.y; });                // 위에서 아래로
    lines.forEach(function (l) {
      l.parts.sort(function (a, b) { return a.x - b.x; });
      var cells = [];
      l.parts.forEach(function (p) {
        var last = cells[cells.length - 1];
        if (last && p.x - last.x2 < gap) { last.s += (p.x - last.x2 > 0.8 ? ' ' : '') + p.s; last.x2 = Math.max(last.x2, p.x2); }
        else cells.push({ x: p.x, x2: p.x2, s: p.s });
      });
      cells.forEach(function (c) { c.s = c.s.replace(/\s+/g, ' ').trim(); c.mid = (c.x + c.x2) / 2; });
      l.cells = cells;
    });
    if (!lines.length) return [];
    var anchor = lines.reduce(function (a, l) { return l.cells.length > a.cells.length ? l : a; }, lines[0]);
    var mids = anchor.cells.map(function (c) { return c.mid; });
    function colOf(cell) {
      // 기준 칸과 가로로 겹치면 그 열, 아니면 가운데가 가장 가까운 열
      for (var i = 0; i < anchor.cells.length; i++) {
        var a = anchor.cells[i];
        if (cell.x < a.x2 && cell.x2 > a.x && (cell.x2 - cell.x) < (a.x2 - a.x) * 3) return i;
      }
      var best = 0;
      mids.forEach(function (m, i) { if (Math.abs(cell.mid - m) < Math.abs(cell.mid - mids[best])) best = i; });
      return best;
    }
    return lines.map(function (l) {
      var row = mids.map(function () { return ''; });
      if (l.cells.length === 1 && l !== anchor) { row[0] = l.cells[0].s; return row; }   // 한 칸짜리 줄(제목)은 첫 칸에
      l.cells.forEach(function (c) { var k = colOf(c); row[k] = row[k] ? row[k] + ' ' + c.s : c.s; });
      return row;
    });
  }

  /* ---------- 목표 표 구조 ---------- */
  var TYPES = { text: '글자', number: '숫자', date: '날짜', code: '코드' };
  function defaultTargets() {
    return [
      { name: '품번', type: 'code', required: true }, { name: '품명', type: 'text', required: true },
      { name: '규격', type: 'text' }, { name: '수량', type: 'number', required: true },
      { name: '단위', type: 'text' }, { name: '단가', type: 'number' }, { name: '금액', type: 'number' },
      { name: '공급업체', type: 'text' }, { name: '납기', type: 'date' }, { name: '비고', type: 'text' }
    ];
  }
  var TYPE_WORDS = { text: ['글자', '텍스트', '문자', 'text', 'string'], number: ['숫자', '수', 'number', 'numeric'], date: ['날짜', '일자', 'date'], code: ['코드', 'code', 'id'] };
  function typeWord(v) {
    var n = norm(v);
    for (var k in TYPE_WORDS) if (TYPE_WORDS[k].indexOf(n) >= 0) return k;
    return null;
  }
  // 목표 양식 엑셀: 첫 줄(머리글) = 목표 열 이름. 이름 끝 * 또는 (필수) = 필수.
  // 바로 아랫줄에 형식 낱말(숫자·날짜·코드·글자)이 있으면 그것을, 없으면 예시 값 모양 → 같은 뜻 낱말 표 순으로 형식을 정한다.
  function targetsFromGrid(grid) {
    var h = detectHeader(grid, []), head = grid[h.start] || [];
    var next = grid[h.start + 1] || [], typeRow = cellsOf(next).length && cellsOf(next).every(function (v) { return typeWord(v); });
    var out = [];
    head.forEach(function (v, c) {
      var raw = str(v);
      if (!raw) return;
      var required = /\*$|\(필수\)$/.test(raw), name = raw.replace(/\s*(\*|\(필수\))$/, '').trim();
      var type = typeRow ? typeWord(next[c]) : null;
      if (!type && !typeRow) {
        var p = profile(grid.slice(h.start + 1, h.start + 6).map(function (r) { return r[c]; }));
        if (p.n) type = p.date >= 0.6 ? 'date' : p.number >= 0.6 ? 'number' : p.code >= 0.6 ? 'code' : null;
      }
      if (!type) { var g = aliasGroup(name); type = g ? g.type : 'text'; }
      out.push({ name: name, type: type, required: required });
    });
    return out;
  }

  /* ---------- 짝짓기 제안 ---------- */
  var THRESHOLD = 0.45;
  function confidence(score) { return score >= 0.8 ? '높음' : score >= 0.6 ? '보통' : '낮음'; }
  // 고를 수 있는 원본: 열 · 머리글 위 정보(meta) · 시트 이름 · 파일 이름
  function candidates(table) {
    var list = table.columns.map(function (c) { return { key: c.key, label: c.letter + '열 · ' + c.name, name: c.name, profile: c.profile, samples: c.samples }; });
    (table.meta || []).forEach(function (m) {
      list.push({ key: 'm:' + m.key, label: '문서 정보 · ' + m.key, name: m.key, profile: profile([m.value]), samples: [str(m.value)], meta: true });
    });
    return list;
  }
  function scorePair(cand, target) {
    var ns = nameScore(cand.name, target.name), fit = typeFit(target.type, cand.profile);
    var score, why = [];
    if (ns.score >= 0.3) { score = 0.75 * ns.score + 0.25 * fit; why.push(ns.why); }
    else { score = 0.3 * fit; }
    if ((target.type === 'number' || target.type === 'date') && cand.profile.n && fit < 0.3) { score *= 0.6; why.push('값이 ' + TYPES[target.type] + ' 모양이 아님'); }
    else if (cand.profile.n && fit >= 0.8 && ns.score >= 0.3) why.push('값 모양 ' + TYPES[target.type]);
    if (cand.meta) score *= 0.95;                        // 같은 점수면 열을 먼저
    return { score: Math.round(score * 100) / 100, why: why.filter(Boolean).join(' · ') };
  }
  // 목표 열마다 원본 하나 — 점수 높은 짝부터 차례로 묶고, 이미 쓴 원본은 다시 쓰지 않는다.
  function suggest(table, targets) {
    var cands = candidates(table), pairs = [];
    targets.forEach(function (t) {
      cands.forEach(function (c) { var s = scorePair(c, t); if (s.score >= THRESHOLD) pairs.push({ t: t.name, c: c.key, score: s.score, why: s.why }); });
    });
    pairs.sort(function (a, b) { return b.score - a.score; });
    var map = {}, used = {};
    pairs.forEach(function (p) {
      if (map[p.t] || used[p.c]) return;
      map[p.t] = { src: p.c, score: p.score, why: p.why, conf: confidence(p.score), by: 'rule' };
      used[p.c] = true;
    });
    targets.forEach(function (t) { if (!map[t.name]) map[t.name] = { src: '', score: 0, why: '맞는 열을 찾지 못함', conf: '', by: 'rule' }; });
    return map;
  }

  /* ---------- 저장해 두는 짝짓기 규칙 ----------
   * 같은 업체가 같은 양식으로 다시 보내면 지난번에 확정한 짝을 그대로 쓴다.
   * 열 위치가 아니라 열 이름으로 저장한다 — 열 하나가 끼어들어도 깨지지 않게. */
  function signature(table) { return table.columns.map(function (c) { return norm(c.name); }).filter(Boolean); }
  function makeRule(name, table, mapping, targets) {
    var map = {};
    targets.forEach(function (t) { var m = mapping[t.name]; if (m && m.src) map[t.name] = m.src === 'const' ? { src: 'const', value: m.value } : { src: m.src }; });
    return { id: 'r' + Date.now().toString(36) + Math.floor(Math.random() * 1e4), name: name, signature: signature(table),
      header: { start: table.header.start, count: table.header.count }, map: map, saved: new Date().toISOString() };
  }
  function jaccard(a, b) {
    var A = distinct(a), B = distinct(b), inter = A.filter(function (x) { return B.indexOf(x) >= 0; }).length;
    return (A.length + B.length - inter) ? inter / (A.length + B.length - inter) : 0;
  }
  function matchRule(rules, table) {
    var sig = signature(table), best = null;
    (rules || []).forEach(function (r) { var s = jaccard(sig, r.signature || []); if (s >= 0.6 && (!best || s > best.score)) best = { rule: r, score: Math.round(s * 100) / 100 }; });
    return best;
  }
  function applyRule(rule, table, targets) {
    var keys = candidates(table).map(function (c) { return c.key; }).concat(['sheet', 'file']);
    var map = {};
    targets.forEach(function (t) {
      var m = rule.map[t.name];
      if (m && (m.src === 'const' || keys.indexOf(m.src) >= 0)) map[t.name] = { src: m.src, value: m.value, score: 1, why: '저장한 규칙 「' + rule.name + '」', conf: '규칙', by: 'saved' };
      else map[t.name] = { src: '', score: 0, why: m ? '규칙의 열 「' + m.src.replace(/^[cm]:/, '') + '」이 이 표에 없음' : '규칙에 없음', conf: '', by: 'saved' };
    });
    return map;
  }
  function validRules(json) {
    var data = typeof json === 'string' ? JSON.parse(json) : json;
    var list = Array.isArray(data) ? data : (data && data.rules) || [];
    return list.filter(function (r) { return r && r.name && r.map && typeof r.map === 'object' && Array.isArray(r.signature); });
  }

  /* ---------- 값 변환 ---------- */
  // { v, mark: 'same'|'conv', ok } — mark 'conv' 는 문서 값의 모양을 바꿨다는 뜻이다.
  function convert(v, type) {
    if (v == null || v === '' || str(v) === '') return { v: '', mark: 'same', ok: true, empty: true };
    if (type === 'number') {
      if (typeof v === 'number') return { v: v, mark: 'same', ok: true };
      var n = toNumber(v);
      return n == null ? { v: str(v), mark: 'same', ok: false } : { v: n, mark: 'conv', ok: true };
    }
    if (type === 'date') {
      var d = toDate(v);
      if (!d) return { v: str(v), mark: 'same', ok: false };
      var s = fmtDate(d);
      return { v: s, mark: (typeof v === 'string' && v.trim() === s) ? 'same' : 'conv', ok: true };
    }
    return { v: str(v), mark: 'same', ok: true };
  }

  /* ---------- 결과 만들기 ----------
   * mark: doc(문서 값) · merge(병합 범위를 풀어 채움) · conv(형식 변환) · meta(문서 머리 정보) · ctx(시트·파일 이름) · const(직접 입력) · calc(계산) */
  var MARKS = { doc: '문서 값', merge: '병합 채움', conv: '형식 변환', meta: '문서 정보', ctx: '시트·파일 이름', const: '직접 입력', calc: '계산', '': '' };
  function groupName(t) { var g = aliasGroup(t.name); return g ? g.name : null; }
  function buildOutput(table, mapping, targets, ctx) {
    ctx = ctx || {};
    var colByKey = {}; table.columns.forEach(function (c) { colByKey[c.key] = c; });
    var metaByKey = {}; (table.meta || []).forEach(function (m) { metaByKey['m:' + m.key] = m.value; });
    var qty = null, price = null, amount = null;
    targets.forEach(function (t) { var g = groupName(t); if (g === '수량') qty = t.name; if (g === '단가') price = t.name; if (g === '금액') amount = t.name; });
    var rows = [], issues = [];
    function issue(row, target, msg) { issues.push({ file: ctx.file || '', sheet: ctx.sheet || '', row: row, target: target, msg: msg }); }
    table.rows.forEach(function (r) {
      var values = {}, marks = {};
      targets.forEach(function (t) {
        var m = mapping[t.name] || {}, raw = '', mark = '';
        if (!m.src) { values[t.name] = ''; marks[t.name] = ''; return; }
        if (m.src === 'const') { raw = m.value; mark = 'const'; }
        else if (m.src === 'sheet') { raw = ctx.sheet || ''; mark = 'ctx'; }
        else if (m.src === 'file') { raw = ctx.file || ''; mark = 'ctx'; }
        else if (m.src.indexOf('m:') === 0) { raw = metaByKey[m.src]; mark = 'meta'; }
        else if (colByKey[m.src]) { var k = colByKey[m.src].index; raw = r.cells[k]; mark = r.filled[k] ? 'merge' : 'doc'; }
        var cv = convert(raw, t.type);
        if (!cv.ok) issue(r.src, t.name, '「' + str(raw) + '」을 ' + TYPES[t.type] + '로 읽지 못함 — 원래 글자를 그대로 둡니다');
        if (cv.mark === 'conv' && (mark === 'doc')) mark = 'conv';
        values[t.name] = cv.v; marks[t.name] = cv.empty ? '' : mark;
      });
      // 금액이 비었고 수량·단가가 있으면 계산해 채운다. 둘 다 있으면 맞는지 본다.
      if (qty && price && amount && typeof values[qty] === 'number' && typeof values[price] === 'number') {
        var calc = Math.round(values[qty] * values[price] * 100) / 100;
        if (values[amount] === '') { values[amount] = calc; marks[amount] = 'calc'; }
        else if (typeof values[amount] === 'number' && Math.abs(values[amount] - calc) > 0.5) issue(r.src, amount, '금액 ' + values[amount] + ' ≠ 수량×단가 ' + calc);
      }
      targets.forEach(function (t) { if (t.required && values[t.name] === '') issue(r.src, t.name, '필수 값이 비어 있음'); });
      rows.push({ values: values, marks: marks, src: { file: ctx.file || '', sheet: ctx.sheet || '', row: r.src } });
    });
    return { rows: rows, issues: issues, file: ctx.file || '', sheet: ctx.sheet || '', mapping: mapping };
  }
  function countMarks(rows) {
    var n = { doc: 0, merge: 0, conv: 0, meta: 0, ctx: 0, const: 0, calc: 0 };
    rows.forEach(function (r) { for (var k in r.marks) if (n[r.marks[k]] != null) n[r.marks[k]]++; });
    return n;
  }

  /* ---------- 내보내기 ---------- */
  function exportSheets(outputs, targets) {
    var names = targets.map(function (t) { return t.name; });
    var head = names.concat(['출처 파일', '출처 시트', '원본 행']);
    var data = [head], marks = [head], issues = [['파일', '시트', '원본 행', '목표 열', '확인할 내용']], rules = [['파일', '시트', '목표 열', '원본', '근거', '제안']];
    outputs.forEach(function (o) {
      o.rows.forEach(function (r) {
        var tail = [r.src.file, r.src.sheet, r.src.row];
        data.push(names.map(function (n) { return r.values[n]; }).concat(tail));
        marks.push(names.map(function (n) { return MARKS[r.marks[n] || ''] || ''; }).concat(tail));
      });
      o.issues.forEach(function (i) { issues.push([i.file, i.sheet, i.row, i.target, i.msg]); });
      targets.forEach(function (t) {
        var m = (o.mapping || {})[t.name] || {};
        rules.push([o.file, o.sheet, t.name, srcLabel(m), m.why || '', ({ rule: '규칙', saved: '저장 규칙', ai: 'AI 초안', user: '사람이 고침' })[m.by] || '']);
      });
    });
    return { '정리결과': data, '값출처': marks, '확인할것': issues, '짝짓기': rules };
  }
  function srcLabel(m) {
    if (!m || !m.src) return '(비움)';
    if (m.src === 'const') return '직접 입력: ' + str(m.value);
    if (m.src === 'sheet') return '시트 이름';
    if (m.src === 'file') return '파일 이름';
    return m.src.replace(/^c:/, '').replace(/^m:/, '문서 정보 · ');
  }
  function toCSV(aoa) {
    return aoa.map(function (row) {
      return row.map(function (v) {
        var s = v == null ? '' : String(v);
        return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(',');
    }).join('\r\n');
  }

  /* ---------- AI 보조 (선택) ----------
   * 화면이 요청을 보내고, 여기서는 묻는 글을 만들고 답을 검사만 한다.
   * 원본 열 이름·예시 값 3개만 보낸다(표 전체를 보내지 않는다). 답에 없는 열 이름이 나오면 버린다. */
  function aiPrompt(table, targets) {
    var cands = candidates(table);
    var system = '너는 엑셀 표 열 짝짓기 도우미다. 원본 열을 목표 열에 하나씩 짝지어 JSON 으로만 답한다. 확신이 없으면 빈 문자열로 둔다.';
    var user = '목표 열:\n' + targets.map(function (t) { return '- ' + t.name + ' (' + TYPES[t.type] + (t.required ? ', 필수' : '') + ')'; }).join('\n') +
      '\n\n원본 열(키 · 이름 · 예시 값):\n' + cands.map(function (c) { return '- ' + c.key + ' · ' + c.name + ' · ' + c.samples.join(' / '); }).join('\n') +
      '\n\n답 형식: {"목표 열 이름": {"src": "원본 키 또는 빈 문자열", "why": "한 줄 이유"}}';
    return { system: system, user: user };
  }
  function parseAiReply(text, table, targets) {
    var s = String(text || ''), m = s.match(/```(?:json)?\s*([\s\S]*?)```/), body = m ? m[1] : s;
    var a = body.indexOf('{'), b = body.lastIndexOf('}');
    if (a < 0 || b < a) throw new Error('AI 답에서 JSON 을 찾지 못했습니다');
    var obj = JSON.parse(body.slice(a, b + 1));
    var keys = candidates(table).map(function (c) { return c.key; });
    var map = {}, dropped = [];
    targets.forEach(function (t) {
      var v = obj[t.name], src = v && typeof v === 'object' ? v.src : v;
      if (src && keys.indexOf(src) < 0) { dropped.push(t.name + '→' + src); src = ''; }
      map[t.name] = { src: src || '', score: 0, why: (v && v.why) ? String(v.why).slice(0, 120) : '', conf: 'AI', by: 'ai' };
    });
    return { map: map, dropped: dropped };
  }

  /* ---------- 예제 (가짜 데이터 — 실제 사내 데이터가 아닙니다) ----------
   * 양식이 서로 다른 세 업체 파일. 병합 셀 · 두 줄 머리글 · 단위 줄 · 합계 줄 · 여러 시트 · 글자로 적힌 숫자·날짜를 일부러 넣었다. */
  function sample() {
    var A = { name: '가람정밀_견적서.xlsx', sheets: [
      { name: '견적서', aoa: [
        ['견 적 서', '', '', '', '', '', '', ''],
        ['업체명 : 가람정밀(주)', '', '', '', '견적일 : 2026-10-28', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        ['No', '품번', '품명', '규격', '', '수량', '단가(원)', '금액(원)'],
        ['', '', '', '재질', '치수', '', '', ''],
        [1, 'HB-2201-A', '유압 호스 어셈블리', 'NBR', '3/4" x 1200', 20, 18500, 370000],
        [2, 'HB-2201-B', '유압 호스 어셈블리', 'NBR', '3/4" x 1500', 12, 21000, 252000],
        [3, 'FT-3310', '리턴 필터 엘리먼트', '여과지', '10μm', 40, 9200, 368000],
        [4, 'SL-0457', '실 키트(붐 실린더)', 'PU', 'Φ110', 8, 64000, 521000],
        [5, 'BK-7702', '브래킷 용접품', 'SS400', 't12', 15, 33500, 502500],
        ['합계', '', '', '', '', 95, '', 2013500],
        ['비고 : 부가세 별도, 납기 발주 후 3주', '', '', '', '', '', '', '']
      ], merges: [
        [0, 0, 0, 7], [1, 0, 1, 3], [1, 4, 1, 7],
        [3, 0, 4, 0], [3, 1, 4, 1], [3, 2, 4, 2], [3, 3, 3, 4], [3, 5, 4, 5], [3, 6, 4, 6], [3, 7, 4, 7],
        [10, 0, 10, 4], [11, 0, 11, 7]
      ] },
      { name: '추가품목', aoa: [
        ['품번', '품명', '규격', '수량', '단가'],
        ['PN-5520', '핀 부싱', 'Φ50 x 60', '30', '7,800'],
        ['PN-5521', '핀 부싱', 'Φ60 x 70', '24', '9,400'],
        ['GR-1001', '그리스 니플', 'M8', '200', '350']
      ], merges: [] }
    ] };
    var B = { name: '한빛유압_부품목록.xlsx', sheets: [
      { name: '유압부품', aoa: [
        ['Part No.', 'Description', "Q'ty", 'Unit', 'Unit Price (KRW)', 'Vendor', 'Lead Time', 'Remarks'],
        ['MCV-901', 'Main Control Valve', 2, 'EA', 3850000, '한빛유압', '6주', ''],
        ['PMP-220', 'Gear Pump', 4, 'EA', 1240000, '', '4주', '도면 Rev.C'],
        ['CYL-B110', 'Boom Cylinder', 2, 'SET', 2980000, '', '8주', ''],
        ['ORK-30', 'O-ring Kit', 50, 'EA', 4800, '동해실링', '1주', '샘플 포함'],
        ['BSK-12', 'Back-up Ring', 120, 'EA', 650, '', '1주', '']
      ], merges: [[1, 5, 3, 5], [4, 5, 5, 5]] },
      { name: '전장부품', aoa: [
        ['전장 부품 목록 (2026년 4분기)', '', '', '', '', '', '', ''],
        ['Part No.', 'Description', "Q'ty", 'Unit', 'Unit Price', 'Vendor', 'Lead Time', 'Remarks'],
        ['', '', '(EA)', '', '(원)', '', '', ''],
        ['ECU-77', 'Machine Controller', 3, 'EA', 1450000, '누리전장', '5주', ''],
        ['HRN-410', 'Main Wire Harness', 6, 'EA', 286000, '', '3주', ''],
        ['SW-08', 'Rocker Switch', 40, 'EA', 5200, '', '2주', '색상 흑색']
      ], merges: [[0, 0, 0, 7], [3, 5, 5, 5]] }
    ] };
    var C = { name: '새솔산업_자재리스트.xlsx', sheets: [
      { name: '자재', aoa: [
        ['자재코드', '자재명/사양', '발주수량', '단위', '단가(원)', '공급가액', '납기일', '협력사'],
        ['ST-PL-12', '열연강판 SS400 t12', '1,250', 'KG', '1,180', '1,475,000', '2026.12.05', '새솔산업'],
        ['ST-PL-16', '열연강판 SS400 t16', '980', 'KG', '1,210', '', '2026.12.05', '새솔산업'],
        ['WD-ER70', '용접봉 ER70S-6 1.2mm', '', 'KG', '4,300', '', '26.11.20', '새솔산업'],
        ['PT-PRM', '프라이머 도료', '40', 'L', '협의', '', '2026년 11월 27일', '새솔산업']
      ], merges: [] }
    ] };
    return [A, B, C].map(function (f) {
      f.sheets.forEach(function (s) { s.merges = s.merges.map(function (m) { return Array.isArray(m) ? { s: { r: m[0], c: m[1] }, e: { r: m[2], c: m[3] } } : m; }); });
      return f;
    });
  }

  var api = {
    str: str, norm: norm, ALIASES: ALIASES, aliasGroup: aliasGroup, TYPES: TYPES, MARKS: MARKS, THRESHOLD: THRESHOLD,
    toDate: toDate, fmtDate: fmtDate, toNumber: toNumber, kindOf: kindOf, profile: profile, typeFit: typeFit,
    dice: dice, nameScore: nameScore, colName: colName, gridFromSheet: gridFromSheet, csvToGrid: csvToGrid,
    fillMerges: fillMerges, detectHeader: detectHeader, readMeta: readMeta, extractTable: extractTable, pdfItemsToGrid: pdfItemsToGrid,
    defaultTargets: defaultTargets, targetsFromGrid: targetsFromGrid, candidates: candidates, scorePair: scorePair, suggest: suggest,
    confidence: confidence, signature: signature, makeRule: makeRule, matchRule: matchRule, applyRule: applyRule, validRules: validRules,
    convert: convert, buildOutput: buildOutput, countMarks: countMarks, exportSheets: exportSheets, srcLabel: srcLabel, toCSV: toCSV,
    aiPrompt: aiPrompt, parseAiReply: parseAiReply, sample: sample
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Tidy = api;
})(this);
