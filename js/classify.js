/* 폴더 분류·추출 — 판정 모듈 (v1.1)
 *
 * 폴더 안 파일을 「문서 종류」로 나누고, 종류마다 정한 항목을 뽑아 조건에 맞는 행만 남긴다.
 * 판정은 여기에만 둔다. 화면(js/folder.js)과 로컬 스크립트(tools/folder-run.mjs)는 이 모듈을 똑같이 부른다.
 * 그래서 웹 도구와 Codex 가 돌리는 스크립트가 같은 규칙 파일(rules/doc-types.json)로 같은 답을 낸다.
 *
 * 흐름
 *   파일(시트·쪽마다 grid)  →  Tidy.extractTable 로 표 풀기(v1.0 엔진 그대로)
 *   →  docOf      : 파일 이름 · 머리글 낱말 · 본문 낱말을 모은다
 *   →  classify   : 종류마다 점수(규칙마다 가중치) — 최소 점수 미만이면 미분류, 1·2위가 가까우면 확인 필요
 *   →  extractDoc : 종류의 「뽑을 항목」(v1.0 목표 열 형식)으로 짝짓기 → buildOutput(값 출처) → 추출 조건으로 거르기
 *   →  exportBook : 분류표 + 종류별 시트 + 뺀 행 + 확인할 것
 *
 * 점수의 근거(어떤 규칙이 몇 점)를 늘 함께 돌려준다. 근거 없이 「견적서」라고만 하면 사람이 틀린 분류를 알아채지 못한다.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory(require('./engine.js'));
  else root.Classify = factory(root.Tidy);
})(this, function (T) {
  'use strict';

  /* ---------- 기본 규칙 (rules/doc-types.json 과 같은 내용 — 테스트가 둘이 같은지 본다) ----------
   * 파일을 직접 열어(file://) 쓸 때는 규칙 파일을 읽을 수 없어서 같은 내용을 여기에도 둔다. */
  var DEFAULT_RULES = {
    app: 'build-2611-02',
    schema: 1,
    note: '문서 종류 규칙 — 웹 도구(folder.html)와 tools/folder-run.mjs 가 같은 파일을 씁니다. 예제 종류 3가지는 가짜입니다.',
    closeMargin: 2,
    types: [
      {
        id: 'quote', name: '견적서', minScore: 4,
        match: {
          fileName: [{ word: '견적', weight: 4 }, { word: 'quotation', weight: 4 }, { word: 'quote', weight: 3 }],
          header: [{ word: '견적금액', weight: 2 }, { word: '견적수량', weight: 2 }, { word: '단가', weight: 1 }],
          body: [{ word: '견적서', weight: 3 }, { word: 'quotation', weight: 3 }, { word: '견적일', weight: 2 }, { word: '유효기간', weight: 2 }, { word: '부가세별도', weight: 1 }]
        },
        fields: [
          { name: '품번', type: 'code', required: true },
          { name: '품명', type: 'text', required: true },
          { name: '수량', type: 'number', required: true },
          { name: '단가', type: 'number' },
          { name: '금액', type: 'number' },
          { name: '공급업체', type: 'text', aliases: ['상호', '공급자'] },
          { name: '견적일', type: 'date', aliases: ['date', '견적일자', 'quotationdate'] }
        ],
        conditions: [{ field: '금액', op: '>=', value: 300000 }]
      },
      {
        id: 'statement', name: '거래명세서', minScore: 4,
        match: {
          fileName: [{ word: '거래명세', weight: 4 }, { word: '명세', weight: 2 }, { word: '납품', weight: 2 }, { word: 'invoice', weight: 3 }],
          header: [{ word: '공급가액', weight: 2 }, { word: '세액', weight: 2 }, { word: '납품일', weight: 2 }, { word: '입고일', weight: 2 }],
          body: [{ word: '거래명세서', weight: 3 }, { word: '공급자', weight: 2 }, { word: '공급받는자', weight: 2 }, { word: '인수자', weight: 1 }, { word: '납품', weight: 1 }]
        },
        fields: [
          { name: '품번', type: 'code', required: true },
          { name: '품명', type: 'text', required: true },
          { name: '수량', type: 'number', required: true },
          { name: '단가', type: 'number' },
          { name: '금액', type: 'number' },
          { name: '납품일', type: 'date', aliases: ['입고일', '거래일', '납품일자'] },
          { name: '공급업체', type: 'text', aliases: ['공급자', '상호'] }
        ],
        conditions: [{ field: '납품일', op: 'between', value: ['2026-10-01', '2026-10-31'] }]
      },
      {
        id: 'inspection', name: '검사성적서', minScore: 4,
        match: {
          fileName: [{ word: '성적서', weight: 4 }, { word: '검사', weight: 3 }, { word: 'inspection', weight: 4 }],
          header: [{ word: '측정값', weight: 2 }, { word: '판정', weight: 2 }, { word: '검사항목', weight: 2 }, { word: 'measured', weight: 2 }, { word: 'result', weight: 1 }],
          body: [{ word: '검사성적서', weight: 3 }, { word: 'inspectionreport', weight: 3 }, { word: '불합격', weight: 1 }]
        },
        fields: [
          { name: '품번', type: 'code', required: true },
          { name: '품명', type: 'text' },
          { name: '검사항목', type: 'text', required: true, aliases: ['항목', 'checkitem', 'inspectionitem'] },
          { name: '규격', type: 'text' },
          { name: '측정값', type: 'text', aliases: ['측정', '측정치', '실측값', 'measured', 'measurement'] },
          { name: '판정', type: 'text', required: true, aliases: ['결과', '합부', '판정결과', 'result'] },
          { name: '검사일', type: 'date', aliases: ['date', '검사일자', 'inspectiondate'] }
        ],
        conditions: [{ field: '판정', op: 'contains', value: ['NG', '불합격'] }]
      }
    ]
  };
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function defaultRules() { return clone(DEFAULT_RULES); }

  /* ---------- 규칙 파일 검사 ----------
   * 사람이 화면에서, 또는 Codex 가 고친 규칙 파일을 받기 전에 모양을 본다. 틀린 곳을 모두 말해 준다. */
  var KINDS = { fileName: '파일 이름', header: '머리글', body: '본문' };
  var OPS = ['>=', '<=', '>', '<', '=', '!=', 'between', 'contains', 'notContains', 'notEmpty'];
  function checkRules(input) {
    var errors = [], r;
    try { r = typeof input === 'string' ? JSON.parse(input) : clone(input); }
    catch (e) { return { ok: false, errors: ['JSON 을 읽지 못했습니다: ' + e.message], rules: null }; }
    if (!r || typeof r !== 'object' || !Array.isArray(r.types)) return { ok: false, errors: ['types 목록이 없습니다'], rules: null };
    if (r.closeMargin == null) r.closeMargin = 2;
    if (typeof r.closeMargin !== 'number' || r.closeMargin < 0) errors.push('closeMargin 은 0 이상의 숫자여야 합니다');
    if (!r.types.length) errors.push('문서 종류가 하나도 없습니다');
    var ids = {};
    r.types.forEach(function (t, i) {
      var at = '종류 ' + (i + 1) + (t && t.name ? '(' + t.name + ')' : '');
      if (!t || typeof t !== 'object') { errors.push(at + ': 모양이 틀렸습니다'); return; }
      if (!t.id || typeof t.id !== 'string') errors.push(at + ': id 가 없습니다');
      else if (ids[t.id]) errors.push(at + ': id 「' + t.id + '」가 겹칩니다'); else ids[t.id] = true;
      if (!t.name) errors.push(at + ': name 이 없습니다');
      if (typeof t.minScore !== 'number') errors.push(at + ': minScore 는 숫자여야 합니다');
      var m = t.match || {};
      Object.keys(m).forEach(function (k) { if (!KINDS[k]) errors.push(at + ': match 에 모르는 칸 「' + k + '」(fileName·header·body 만)'); });
      Object.keys(KINDS).forEach(function (k) {
        (m[k] || []).forEach(function (w, j) {
          if (!w || (!w.word && !w.re)) errors.push(at + ': ' + KINDS[k] + ' ' + (j + 1) + '번째에 word 나 re 가 없습니다');
          if (w && typeof w.weight !== 'number') errors.push(at + ': ' + KINDS[k] + ' 「' + (w.word || w.re) + '」의 weight 가 숫자가 아닙니다');
          if (w && w.re) { try { new RegExp(w.re, 'i'); } catch (e) { errors.push(at + ': 정규식 「' + w.re + '」이 틀렸습니다'); } }
        });
      });
      if (!Array.isArray(t.fields) || !t.fields.length) errors.push(at + ': 뽑을 항목(fields)이 없습니다');
      var names = {};
      (t.fields || []).forEach(function (f) {
        if (!f || !f.name) { errors.push(at + ': 이름 없는 항목이 있습니다'); return; }
        if (names[f.name]) errors.push(at + ': 항목 「' + f.name + '」이 겹칩니다'); names[f.name] = true;
        if (f.type && !T.TYPES[f.type]) errors.push(at + ': 항목 「' + f.name + '」의 형식 「' + f.type + '」을 모릅니다(text·number·date·code)');
        if (!f.type) f.type = 'text';
        if (f.aliases && !Array.isArray(f.aliases)) errors.push(at + ': 항목 「' + f.name + '」의 aliases 는 목록이어야 합니다');
      });
      (t.conditions || []).forEach(function (c) {
        if (!c || !names[c.field]) errors.push(at + ': 추출 조건의 항목 「' + (c && c.field) + '」이 fields 에 없습니다');
        if (!c || OPS.indexOf(c.op) < 0) errors.push(at + ': 추출 조건의 비교 「' + (c && c.op) + '」를 모릅니다(' + OPS.join(' ') + ')');
        else if (c.op === 'between' && !(Array.isArray(c.value) && c.value.length === 2)) errors.push(at + ': between 은 [시작, 끝] 두 값이 필요합니다');
      });
      if (!t.match) t.match = {};
      if (!t.conditions) t.conditions = [];
    });
    return { ok: !errors.length, errors: errors, rules: errors.length ? null : r };
  }
  // 사람이 읽고 고치기 쉽게: 짧은 목록·짧은 묶음은 한 줄로 접는다(내용은 JSON.stringify 와 같다)
  function rulesToJSON(r) {
    var s = JSON.stringify(r, null, 2);
    function fold(re, open, close) {
      return s.replace(re, function (all, body) {
        var one = open + ' ' + body.split('\n').map(function (l) { return l.trim(); }).join(' ') + ' ' + close;
        return one.length <= 150 ? one : all;
      });
    }
    for (var i = 0; i < 2; i++) {
      s = fold(/\[\n([^{}\[\]]*?)\n\s*\]/g, '[', ']');
      s = fold(/\{\n([^{}\[\]]*?(?:\[[^\]\n]*\][^{}\[\]]*?)*)\n\s*\}/g, '{', '}');
    }
    return s + '\n';
  }

  /* ---------- 파일 한 개 → 문서 ---------- */
  var SUPPORTED = /\.(xlsx|xlsm|xls|csv|pdf)$/i;
  function isSupported(path) { var n = baseName(path); return SUPPORTED.test(n) && !/^~\$/.test(n); }
  function baseName(path) { return String(path).split(/[\\/]/).pop(); }
  function folderOf(path) { var a = String(path).split(/[\\/]/); a.pop(); return a.join('/'); }
  // 낱말 비교용: 소문자, 공백·구두점 제거(괄호 안은 지우지 않는다 — 본문 낱말이 괄호 안에 있을 수 있다)
  function flat(s) { return T.str(s).toLowerCase().replace(/[\s.,_\-\/'’"`#:：·*~()（）\[\]]/g, ''); }

  /* sheets: [{ name, grid, merges }] — 시트(엑셀) 또는 쪽(PDF)마다 하나 */
  function docOf(path, sheets, kind) {
    var tables = [], header = [], body = [];
    (sheets || []).forEach(function (s) {
      var tb = T.extractTable(s.grid || [], { merges: s.merges || [] });
      var empty = !tb.columns.length || !tb.rows.length;
      tables.push({ sheet: s.name, table: tb, empty: empty });
      if (!empty) tb.columns.forEach(function (c) { header.push(c.name); });
      (s.grid || []).forEach(function (row) { row.forEach(function (v) { var x = T.str(v); if (x) body.push(x); }); });
    });
    return { path: String(path), name: baseName(path), folder: folderOf(path), kind: kind || (/\.pdf$/i.test(path) ? 'pdf' : 'xlsx'),
      tables: tables, header: header, text: body.join(' ') };
  }

  /* ---------- 종류 점수 ----------
   * 규칙 낱말 하나는 한 번만 센다(본문에 「견적」이 열 번 나와도 한 번). 많이 나온다고 점수가 부풀면
   * 긴 문서가 늘 이기게 된다. */
  function hitOf(w, hay) {
    if (w.re) { try { return new RegExp(w.re, 'i').test(hay.raw); } catch (e) { return false; } }
    var k = flat(w.word);
    return !!k && hay.flat.indexOf(k) >= 0;
  }
  function scoreType(doc, type) {
    var hays = {
      fileName: { raw: doc.name, flat: flat(doc.name) },
      header: { raw: doc.header.join(' | '), flat: doc.header.map(flat).join('|') },
      body: { raw: doc.text, flat: flat(doc.text) }
    };
    var score = 0, hits = [], m = type.match || {};
    Object.keys(KINDS).forEach(function (k) {
      (m[k] || []).forEach(function (w) {
        if (hitOf(w, hays[k])) { score += w.weight; hits.push({ kind: k, word: w.word || w.re, weight: w.weight }); }
      });
    });
    return { id: type.id, name: type.name, score: Math.round(score * 100) / 100, min: type.minScore, hits: hits };
  }
  function whyText(hits) {
    return hits.map(function (h) { return KINDS[h.kind] + ' 「' + h.word + '」 ' + (h.weight >= 0 ? '+' : '') + h.weight; }).join(' · ');
  }
  /* 결과 status: ok(분류됨) · check(확인 필요 — 1·2위가 closeMargin 이내) · none(미분류 — 최소 점수에 못 미침) */
  function classify(doc, rules) {
    var ranking = rules.types.map(function (t) { return scoreType(doc, t); })
      .sort(function (a, b) { return b.score - a.score || (b.score - b.min) - (a.score - a.min); });
    var passed = ranking.filter(function (r) { return r.score >= r.min; });
    if (!passed.length) {
      var top = ranking[0];
      return { status: 'none', type: '', score: top ? top.score : 0, ranking: ranking, hits: top ? top.hits : [],
        why: top && top.hits.length ? '가장 가까운 ' + top.name + ' ' + top.score + '점 < 최소 ' + top.min + '점' : '맞는 규칙이 하나도 없음' };
    }
    var best = passed[0], second = passed[1];
    if (second && best.score - second.score <= rules.closeMargin) {
      return { status: 'check', type: best.id, score: best.score, ranking: ranking, hits: best.hits, second: second,
        why: best.name + ' ' + best.score + '점 · ' + second.name + ' ' + second.score + '점 — 차이 ' + Math.round((best.score - second.score) * 100) / 100 + '점(기준 ' + rules.closeMargin + '점 이내)' };
    }
    return { status: 'ok', type: best.id, score: best.score, ranking: ranking, hits: best.hits, why: whyText(best.hits) };
  }

  /* ---------- 뽑을 항목 짝짓기 ----------
   * v1.0 짝짓기(Tidy.scorePair)에 규칙 파일의 aliases(이 종류에서만 쓰는 같은 뜻 낱말)를 더한다.
   * 고르는 방식은 v1.0 suggest 와 같다 — 점수 높은 짝부터, 한 원본은 한 번만. */
  function aliasScore(cand, field) {
    var n = T.norm(cand.name), best = null;
    (field.aliases || []).forEach(function (a) {
      var k = T.norm(a);
      if (!k) return;
      var s = n === k ? 0.9 : (k.length >= 2 && n.indexOf(k) >= 0) ? 0.75 : 0;
      if (s && (!best || s > best.s)) best = { s: s, a: a };
    });
    return best;
  }
  function mapFields(table, fields) {
    var cands = T.candidates(table), pairs = [];
    fields.forEach(function (f) {
      cands.forEach(function (c) {
        var s = T.scorePair(c, f), a = aliasScore(c, f), fit = T.typeFit(f.type, c.profile);
        if (a) {
          var as = Math.round((0.75 * a.s + 0.25 * fit) * (c.meta ? 0.95 : 1) * 100) / 100;
          if (as > s.score) s = { score: as, why: '규칙 파일의 같은 뜻 낱말(' + a.a + ')' };
        }
        if (s.score >= T.THRESHOLD) pairs.push({ t: f.name, c: c.key, score: s.score, why: s.why });
      });
    });
    pairs.sort(function (a, b) { return b.score - a.score; });
    var map = {}, used = {};
    pairs.forEach(function (p) {
      if (map[p.t] || used[p.c]) return;
      map[p.t] = { src: p.c, score: p.score, why: p.why, conf: T.confidence(p.score), by: 'rule' };
      used[p.c] = true;
    });
    fields.forEach(function (f) { if (!map[f.name]) map[f.name] = { src: '', score: 0, why: '맞는 열을 찾지 못함', conf: '', by: 'rule' }; });
    return map;
  }

  /* ---------- 추출 조건 ----------
   * 조건은 모두 맞아야 남긴다(그리고). 값이 비었거나 모양이 달라 견줄 수 없으면 남기지 않고 사유를 적는다. */
  function isDay(s) { return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s); }
  function asComparable(v, ref) {
    if (typeof ref === 'number') return typeof v === 'number' ? v : T.toNumber(v);
    var d = T.toDate(ref);
    if (d) { var x = T.toDate(v); return x ? T.fmtDate(x) : null; }
    return T.str(v);
  }
  function show(v) { return typeof v === 'number' ? v.toLocaleString('ko-KR') : T.str(v); }
  function condText(c) {
    if (c.op === 'between') return c.field + ' ' + show(c.value[0]) + ' ~ ' + show(c.value[1]);
    if (c.op === 'contains') return c.field + '에 「' + [].concat(c.value).join('」 또는 「') + '」 포함';
    if (c.op === 'notContains') return c.field + '에 「' + [].concat(c.value).join('」·「') + '」 없음';
    if (c.op === 'notEmpty') return c.field + ' 값 있음';
    return c.field + ' ' + c.op + ' ' + show(c.value);
  }
  function checkCond(values, c) {
    var v = values[c.field];
    if (c.op === 'notEmpty') return T.str(v) ? null : c.field + ' 값이 비어 있음';
    if (c.op === 'contains' || c.op === 'notContains') {
      var s = T.str(v).toLowerCase();
      var any = [].concat(c.value).some(function (w) { return s.indexOf(String(w).toLowerCase()) >= 0; });
      if (c.op === 'contains') return any ? null : c.field + ' 「' + T.str(v) + '」에 ' + [].concat(c.value).join('·') + ' 없음';
      return any ? c.field + ' 「' + T.str(v) + '」에 ' + [].concat(c.value).join('·') + ' 있음' : null;
    }
    if (T.str(v) === '') return c.field + ' 값이 비어 있어 견줄 수 없음';
    if (c.op === 'between') {
      var lo = asComparable(c.value[0], c.value[0]), hi = asComparable(c.value[1], c.value[1]), x = asComparable(v, c.value[0]);
      if (x == null) return c.field + ' 「' + T.str(v) + '」을 견줄 수 없음';
      return (x >= lo && x <= hi) ? null : c.field + ' ' + show(x) + ' 이(가) ' + show(lo) + ' ~ ' + show(hi) + ' 밖';
    }
    var a = asComparable(v, c.value), b = asComparable(c.value, c.value);
    if (a == null || (typeof b === 'number' && typeof a !== 'number')) return c.field + ' 「' + T.str(v) + '」을 견줄 수 없음';
    var okk = c.op === '>=' ? a >= b : c.op === '<=' ? a <= b : c.op === '>' ? a > b : c.op === '<' ? a < b : c.op === '=' ? a === b : a !== b;
    return okk ? null : c.field + ' ' + show(a) + ' — 조건 ' + c.op + ' ' + show(b) + ' 에 맞지 않음';
  }
  function checkRow(values, conditions) {
    var reasons = [];
    (conditions || []).forEach(function (c) { var r = checkCond(values, c); if (r) reasons.push(r); });
    return { ok: !reasons.length, reasons: reasons };
  }

  /* ---------- 문서 한 개 추출 ---------- */
  function findType(rules, id) { for (var i = 0; i < rules.types.length; i++) if (rules.types[i].id === id) return rules.types[i]; return null; }
  function extractDoc(doc, type) {
    var out = { kept: [], dropped: [], issues: [], tables: [], skippedTables: [] };
    doc.tables.forEach(function (t) {
      if (t.empty) { out.skippedTables.push({ sheet: t.sheet, reason: '표를 찾지 못함' }); return; }
      var mapping = mapFields(t.table, type.fields);
      var need = type.fields.filter(function (f) { return f.required; });
      var found = need.filter(function (f) { return mapping[f.name].src; });
      if (need.length && !found.length) { out.skippedTables.push({ sheet: t.sheet, reason: '필수 항목(' + need.map(function (f) { return f.name; }).join('·') + ')에 맞는 열이 없음' }); return; }
      var o = T.buildOutput(t.table, mapping, type.fields, { file: doc.path, sheet: t.sheet });
      var kept = 0, dropped = 0;
      o.rows.forEach(function (r) {
        var c = checkRow(r.values, type.conditions);
        if (c.ok) { out.kept.push(r); kept++; }
        else { out.dropped.push({ row: r, reasons: c.reasons }); dropped++; }
      });
      o.issues.forEach(function (i) { out.issues.push(i); });
      out.tables.push({ sheet: t.sheet, mapping: mapping, rows: o.rows.length, kept: kept, dropped: dropped });
    });
    return out;
  }

  /* ---------- 폴더 전체 ----------
   * assign: { 경로: 종류 id | '' } — 사람이 고른 종류. 확인 필요·미분류는 사람이 고르기 전에는 뽑지 않는다
   * (틀린 종류로 뽑은 행이 결과에 섞이면 뒤에서 걸러 내기 더 어렵다). */
  function runFolder(docs, rules, assign) {
    assign = assign || {};
    var items = docs.map(function (doc) {
      var cls = classify(doc, rules), final = '', by = 'rule';
      if (Object.prototype.hasOwnProperty.call(assign, doc.path)) { final = assign[doc.path] || ''; by = 'user'; }
      else if (cls.status === 'ok') final = cls.type;
      var type = final ? findType(rules, final) : null;
      return { doc: doc, cls: cls, type: type ? type.id : '', by: by, ext: type ? extractDoc(doc, type) : null };
    });
    var groups = rules.types.map(function (t) {
      var its = items.filter(function (i) { return i.type === t.id; });
      return { id: t.id, name: t.name, items: its,
        kept: its.reduce(function (n, i) { return n + i.ext.kept.length; }, 0),
        dropped: its.reduce(function (n, i) { return n + i.ext.dropped.length; }, 0) };
    });
    var pend = items.filter(function (i) { return !i.type; });
    return {
      items: items, groups: groups,
      check: pend.filter(function (i) { return i.by !== 'user' && i.cls.status === 'check'; }),
      none: pend.filter(function (i) { return i.by === 'user' || i.cls.status !== 'check'; })
    };
  }
  function statusLabel(it) {
    if (it.by === 'user') return it.type ? '사람이 고름' : '사람이 뺌';
    return { ok: '분류됨', check: '확인 필요', none: '미분류' }[it.cls.status];
  }

  /* ---------- 엑셀로 ---------- */
  var BAD_SHEET = /[\[\]:*?\/\\]/g;
  function sheetName(name, used) {
    var base = String(name).replace(BAD_SHEET, ' ').slice(0, 28) || '종류', n = base, k = 2;
    while (used[n]) n = base + ' ' + (k++);
    used[n] = true;
    return n;
  }
  function marksText(r) {
    return Object.keys(r.marks).filter(function (k) { return r.marks[k] && r.marks[k] !== 'doc'; })
      .map(function (k) { return k + ': ' + T.MARKS[r.marks[k]]; }).join('; ');
  }
  // skipped: [{ path, reason }] — 읽지 않은 파일(형식이 다르거나 읽지 못함). 주면 「건너뛴 파일」 시트를 더한다.
  function exportBook(res, rules, skipped) {
    var used = { '분류표': true, '뺀 행': true, '확인할것': true, '건너뛴 파일': true }, book = {};
    var cls = [['경로', '파일', '종류', '상태', '점수', '근거', '2순위', '뽑은 행', '뺀 행']];
    res.items.forEach(function (i) {
      var t = i.type ? findType(rules, i.type) : null, r2 = i.cls.ranking[1];
      cls.push([i.doc.path, i.doc.name, t ? t.name : '', statusLabel(i), i.cls.score, i.cls.status === 'ok' ? i.cls.why : i.cls.why + (i.cls.hits.length ? ' / ' + whyText(i.cls.hits) : ''),
        r2 ? r2.name + ' ' + r2.score + '점' : '', i.ext ? i.ext.kept.length : '', i.ext ? i.ext.dropped.length : '']);
    });
    book['분류표'] = cls;
    var drop = [['경로', '종류', '시트·쪽', '원본 행', '뺀 사유']], iss = [['경로', '시트·쪽', '원본 행', '항목', '확인할 내용']];
    res.groups.forEach(function (g) {
      var t = findType(rules, g.id), names = t.fields.map(function (f) { return f.name; });
      var aoa = [names.concat(['출처 경로', '시트·쪽', '원본 행', '값 출처(문서 값 외)'])];
      g.items.forEach(function (i) {
        i.ext.kept.forEach(function (r) { aoa.push(names.map(function (n) { return r.values[n]; }).concat([r.src.file, r.src.sheet, r.src.row, marksText(r)])); });
        i.ext.dropped.forEach(function (d) { drop.push([i.doc.path, g.name, d.row.src.sheet, d.row.src.row, d.reasons.join(' · ')]); });
        i.ext.issues.forEach(function (x) { iss.push([x.file, x.sheet, x.row, x.target, x.msg]); });
      });
      book[sheetName(g.name, used)] = aoa;
    });
    book['뺀 행'] = drop;
    book['확인할것'] = iss;
    if (skipped) book['건너뛴 파일'] = [['경로', '사유']].concat(skipped.map(function (x) { return [x.path, x.reason]; }));
    return book;
  }

  /* ---------- 이미 처리한 파일 건너뛰기 (로컬 스크립트용) ----------
   * 경로 + 크기 + 고친 시각이 같으면 같은 파일로 본다. 내용이 바뀌면 시각이 바뀌어 다시 처리한다. */
  function fileStamp(st) { return { size: st.size, mtime: Math.round(+st.mtime || st.mtimeMs || 0) }; }
  function alreadyDone(processed, path, stamp) {
    var p = processed && processed.files && processed.files[path];
    return !!p && p.size === stamp.size && p.mtime === stamp.mtime;
  }
  function markDone(processed, res, stamps, when) {
    processed = processed && processed.files ? processed : { app: 'build-2611-02', files: {} };
    res.items.forEach(function (i) {
      var s = stamps[i.doc.path]; if (!s) return;
      processed.files[i.doc.path] = { size: s.size, mtime: s.mtime, type: i.type, status: statusLabel(i), at: when };
    });
    return processed;
  }

  /* ---------- 예제 폴더 (가짜 데이터 — 실제 사내 데이터가 아닙니다) ----------
   * 종류 3가지 × 2 + 확인 필요 1 + 미분류 1, 그리고 이 도구가 읽지 않는 메모 파일 1개.
   * samples/ 의 실제 파일은 tools/make-samples.mjs 가 이 정의로 만든다. */
  function M(a) { return { s: { r: a[0], c: a[1] }, e: { r: a[2], c: a[3] } }; }
  function sampleFolder() {
    var garam = T.sample()[0];
    var F = [
      { path: '2026-10/견적/가람정밀_견적서_1028.xlsx', sheets: garam.sheets.map(function (s) { return { name: s.name, aoa: s.aoa, merges: s.merges }; }) },
      { path: '2026-10/견적/누리전장_quotation_Q2611.xlsx', sheets: [{ name: 'Quotation', aoa: [
        ['QUOTATION', '', '', '', ''],
        ['Vendor : 누리전장', '', 'Date : 2026-10-30', '', ''],
        ['Part No.', 'Description', "Q'ty", 'Unit Price', 'Amount'],
        ['ECU-77', 'Machine Controller', 3, 1450000, 4350000],
        ['HRN-410', 'Main Wire Harness', 6, 286000, 1716000],
        ['SW-08', 'Rocker Switch', 40, 5200, 208000]
      ], merges: [M([0, 0, 0, 4]), M([1, 0, 1, 1]), M([1, 2, 1, 4])] }] },
      { path: '2026-10/납품/한빛유압_거래명세서_1029.xlsx', sheets: [{ name: '명세', aoa: [
        ['거 래 명 세 서', '', '', '', '', '', '', ''],
        ['공급자 : 한빛유압', '', '', '', '공급받는자 : 가나건기(주)', '', '', ''],
        ['품번', '품명', '규격', '수량', '단가', '공급가액', '세액', '납품일'],
        ['MCV-901', 'Main Control Valve', '4포트', 2, 3850000, 7700000, 770000, '2026-10-29'],
        ['PMP-220', 'Gear Pump', '45cc', 4, 1240000, 4960000, 496000, '2026-10-29'],
        ['ORK-30', 'O-ring Kit', 'NBR', 50, 4800, 240000, 24000, '2026-10-30'],
        ['BSK-12', 'Back-up Ring', 'PTFE', 120, 650, 78000, 7800, '2026-11-03'],
        ['합계', '', '', '', '', 12978000, 1297800, '']
      ], merges: [M([0, 0, 0, 7]), M([1, 0, 1, 3]), M([1, 4, 1, 7])] }] },
      { path: '2026-10/납품/새솔산업_명세_1030.xlsx', sheets: [{ name: 'Sheet1', aoa: [
        ['납품 명세 (인수자 확인용)', '', '', '', '', '', ''],
        ['자재코드', '자재명', '수량', '단위', '단가', '금액', '입고일'],
        ['ST-PL-12', '열연강판 SS400 t12', '1,250', 'KG', '1,180', '1,475,000', '2026.10.15'],
        ['ST-PL-16', '열연강판 SS400 t16', '980', 'KG', '1,210', '1,185,800', '2026.10.22'],
        ['WD-ER70', '용접봉 ER70S-6 1.2mm', '300', 'KG', '4,300', '1,290,000', '2026.09.28'],
        ['PT-PRM', '프라이머 도료', '40', 'L', '12,500', '500,000', '26.10.30']
      ], merges: [M([0, 0, 0, 6])] }] },
      { path: '2026-11/품질/동해실링_검사성적서.xlsx', sheets: [{ name: '성적서', aoa: [
        ['검 사 성 적 서', '', '', '', '', ''],
        ['업체명 : 동해실링', '', '', '검사일 : 2026-11-02', '', ''],
        ['품번', '품명', '검사항목', '규격', '측정값', '판정'],
        ['ORK-30', 'O-ring Kit', '외경', 'Φ30 ±0.2', 'Φ30.1', '합격'],
        ['', '', '경도', '70 ±5 HA', '77 HA', '불합격'],
        ['', '', '외관', '흠 없음', '양호', '합격'],
        ['BSK-12', 'Back-up Ring', '두께', '1.5 ±0.1', '1.52', '합격'],
        ['', '', '내경', 'Φ12 +0.1', 'Φ12.18', 'NG'],
        ['', '', '외관', '흠 없음', '양호', 'OK']
      ], merges: [M([0, 0, 0, 5]), M([1, 0, 1, 2]), M([1, 3, 1, 5]), M([3, 0, 5, 0]), M([3, 1, 5, 1]), M([6, 0, 8, 0]), M([6, 1, 8, 1])] }] },
      { path: '2026-11/품질/ECU_inspection_report.pdf', pdf: true },
      { path: '2026-11/기타/견적_명세_통합.xlsx', sheets: [{ name: '정리', aoa: [
        ['견적·납품 정리', '', '', '', '', ''],
        ['견적일 : 2026-10-20', '', '', '', '', ''],
        ['품번', '품명', '수량', '단가', '공급가액', '납품일'],
        ['HB-2201-A', '유압 호스 어셈블리', 20, 18500, 370000, '2026-10-27'],
        ['FT-3310', '리턴 필터 엘리먼트', 40, 9200, 368000, '2026-10-28']
      ], merges: [M([0, 0, 0, 5]), M([1, 0, 1, 5])] }] },
      { path: '2026-11/기타/회의록_정리.xlsx', sheets: [{ name: '회의록', aoa: [
        ['주간 회의 정리', '', '', ''],
        ['안건', '담당', '기한', '진행'],
        ['협력사 방문 일정', '구매팀', '2026-11-10', '진행 중'],
        ['도면 개정 공유', '설계팀', '2026-11-14', '예정']
      ], merges: [M([0, 0, 0, 3])] }] },
      { path: '2026-11/기타/메모.txt', text: '이 파일은 엑셀·PDF 가 아니라 건너뜁니다.\n' }
    ];
    return F;
  }
  // PDF 예제의 글자 줄 — x 좌표로 칸을 나눈다(영문 Helvetica: 한글 글꼴을 넣지 않아도 되게)
  var PDF_LINES = [
    [[200, 'INSPECTION REPORT']],
    [[50, 'Vendor : Nuri Electric'], [330, 'Date : 2026-11-01']],
    [[50, 'Part No'], [130, 'Part Name'], [250, 'Check Item'], [340, 'Spec'], [420, 'Measured'], [500, 'Result']],
    [[50, 'ECU-77'], [130, 'Controller'], [250, 'Voltage'], [340, '24V'], [420, '24.1V'], [500, 'OK']],
    [[50, 'ECU-77'], [130, 'Controller'], [250, 'Insulation'], [340, '100MO'], [420, '85MO'], [500, 'NG']],
    [[50, 'SW-08'], [130, 'Rocker Switch'], [250, 'Torque'], [340, '2.0Nm'], [420, '2.1Nm'], [500, 'OK']]
  ];
  function pdfBytes(lines) {
    var y = 760, ops = 'BT /F1 10 Tf\n';
    lines.forEach(function (ln, i) {
      ln.forEach(function (p) { ops += '1 0 0 1 ' + p[0] + ' ' + (y - i * 22) + ' Tm (' + String(p[1]).replace(/([()\\])/g, '\\$1') + ') Tj\n'; });
    });
    ops += 'ET';
    var objs = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
      '<< /Length ' + ops.length + ' >>\nstream\n' + ops + '\nendstream',
      '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    ];
    var out = '%PDF-1.4\n', off = [];
    objs.forEach(function (o, i) { off.push(out.length); out += (i + 1) + ' 0 obj\n' + o + '\nendobj\n'; });
    var x = out.length;
    out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n' + off.map(function (n) { return ('000000000' + n).slice(-10) + ' 00000 n \n'; }).join('');
    out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + x + '\n%%EOF\n';
    var u = new Uint8Array(out.length);
    for (var i = 0; i < out.length; i++) u[i] = out.charCodeAt(i) & 255;
    return u;
  }
  function samplePdf() { return pdfBytes(PDF_LINES); }

  return {
    DEFAULT_RULES: DEFAULT_RULES, defaultRules: defaultRules, checkRules: checkRules, rulesToJSON: rulesToJSON,
    KINDS: KINDS, OPS: OPS, SUPPORTED: SUPPORTED, isSupported: isSupported, baseName: baseName, folderOf: folderOf, flat: flat,
    docOf: docOf, scoreType: scoreType, whyText: whyText, classify: classify, mapFields: mapFields,
    condText: condText, checkRow: checkRow, findType: findType, extractDoc: extractDoc, runFolder: runFolder,
    statusLabel: statusLabel, exportBook: exportBook, sheetName: sheetName,
    fileStamp: fileStamp, alreadyDone: alreadyDone, markDone: markDone,
    sampleFolder: sampleFolder, samplePdf: samplePdf, pdfBytes: pdfBytes
  };
});
