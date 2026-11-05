/* 화면 — 판정은 engine.js(Tidy)가 한다. 여기서는 파일을 읽어 넘기고, 그리고, 사람이 고른 값을 넘길 뿐이다. */
(function () {
  'use strict';
  var T = window.Tidy;
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var LS = { targets: 'b261102.targets', rules: 'b261102.rules', ai: 'b261102.ai' };
  function lsGet(k, d) { try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 저장 공간이 없으면 이번 화면에서만 쓴다 */ } }

  var state = {
    targets: lsGet(LS.targets, null) || T.defaultTargets(),
    rules: T.validRules(lsGet(LS.rules, [])),
    tables: [], cur: -1, seq: 0
  };
  function today() { return T.fmtDate(new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()))); }
  function msg(id, cls, html) { $(id).innerHTML = html ? '<div class="' + cls + '">' + html + '</div>' : ''; }

  /* ================= 1. 목표 열 ================= */
  function renderTargets() {
    var types = Object.keys(T.TYPES);
    $('targets').innerHTML = state.targets.map(function (t, i) {
      return '<tr><td class="num">' + (i + 1) + '</td>' +
        '<td><input type="text" data-i="' + i + '" data-f="name" value="' + esc(t.name) + '" aria-label="열 이름"></td>' +
        '<td><select data-i="' + i + '" data-f="type" aria-label="형식">' + types.map(function (k) { return '<option value="' + k + '"' + (t.type === k ? ' selected' : '') + '>' + T.TYPES[k] + '</option>'; }).join('') + '</select></td>' +
        '<td><input type="checkbox" data-i="' + i + '" data-f="required"' + (t.required ? ' checked' : '') + ' aria-label="필수"></td>' +
        '<td><button class="x" data-del="' + i + '">빼기</button></td></tr>';
    }).join('');
  }
  function targetsChanged(note) {
    state.targets = state.targets.filter(function (t) { return T.str(t.name); });
    var seen = {};
    state.targets.forEach(function (t) { t.name = T.str(t.name); if (seen[t.name]) t.name += ' (' + (++seen[t.name]) + ')'; else seen[t.name] = 1; });
    lsSet(LS.targets, state.targets);
    renderTargets();
    var n = 0;
    state.tables.forEach(function (t) { if (!t.confirmed && !t.empty) { propose(t); n++; } else if (t.confirmed) fillMissing(t); });
    msg('target-status', 'ok', (note || '목표 열 ' + state.targets.length + '개를 저장했습니다.') + (n ? ' 확정 전 표 ' + n + '개의 제안을 다시 만들었습니다.' : ''));
    renderAll();
  }
  function fillMissing(t) { state.targets.forEach(function (x) { if (!t.mapping[x.name]) t.mapping[x.name] = { src: '', why: '목표 열이 새로 생김', by: 'user' }; }); }
  $('targets').addEventListener('change', function (e) {
    var el = e.target, i = +el.getAttribute('data-i'), f = el.getAttribute('data-f');
    if (!f) return;
    var t = state.targets[i], old = t.name;
    if (f === 'required') t.required = el.checked; else t[f] = el.value;
    if (f === 'name') state.tables.forEach(function (tb) { if (tb.mapping[old]) { tb.mapping[T.str(el.value)] = tb.mapping[old]; delete tb.mapping[old]; } });
    targetsChanged();
  });
  $('targets').addEventListener('click', function (e) {
    var d = e.target.getAttribute('data-del');
    if (d == null) return;
    state.targets.splice(+d, 1); targetsChanged();
  });
  $('btn-add-target').onclick = function () { state.targets.push({ name: '새 열 ' + (state.targets.length + 1), type: 'text', required: false }); targetsChanged(); };
  $('btn-reset-target').onclick = function () { state.targets = T.defaultTargets(); targetsChanged('기본 목표 열(부품·견적 10개)로 되돌렸습니다.'); };
  $('btn-tpl').onclick = function () {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      state.targets.map(function (t) { return t.name + (t.required ? ' *' : ''); }),
      state.targets.map(function (t) { return T.TYPES[t.type]; })
    ]), '목표양식');
    XLSX.writeFile(wb, '문서정리_목표양식.xlsx');
  };
  $('tpl-file').onchange = function (e) {
    var f = e.target.files[0]; if (!f) return;
    readBook(f, function (wb) {
      var g = T.gridFromSheet(wb.Sheets[wb.SheetNames[0]]);
      var tg = T.targetsFromGrid(g.grid);
      if (!tg.length) return msg('target-status', 'warn', '양식의 첫 시트에서 열 이름을 찾지 못했습니다.');
      state.targets = tg; targetsChanged('양식 「' + esc(f.name) + '」에서 목표 열 ' + tg.length + '개를 읽었습니다.');
    });
    e.target.value = '';
  };

  /* ================= 2. 파일 읽기 ================= */
  function readBook(f, done) {
    var rd = new FileReader();
    rd.onload = function (e) {
      try {
        var wb = /\.csv$/i.test(f.name)
          ? { SheetNames: ['CSV'], csv: T.csvToGrid(new TextDecoder('utf-8').decode(e.target.result)) }
          : XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        if (wb.csv) { wb.Sheets = { CSV: XLSX.utils.aoa_to_sheet(wb.csv) }; }
        done(wb);
      } catch (err) { msg('file-status', 'warn', esc(f.name) + ' 을 읽지 못했습니다: ' + esc(err.message)); }
    };
    rd.readAsArrayBuffer(f);
  }
  function addTable(file, sheet, grid, merges, kind) {
    var t = { id: ++state.seq, file: file, sheet: sheet, kind: kind, grid: grid, merges: merges || [], confirmed: false, include: true, header: null };
    extract(t);
    var hit = !t.empty && T.matchRule(state.rules, t.table);
    if (hit) {
      var rh = hit.rule.header;
      if (rh && (rh.start !== t.table.header.start || rh.count !== t.table.header.count)) { t.header = rh; extract(t); }
      t.mapping = T.applyRule(hit.rule, t.table, state.targets);
      t.ruleNote = '저장한 규칙 「' + hit.rule.name + '」을 썼습니다(열 이름 ' + Math.round(hit.score * 100) + '% 일치).';
    } else if (!t.empty) propose(t);
    else t.mapping = {};
    if (t.empty) t.include = false;
    state.tables.push(t);
    return t;
  }
  function extract(t) {
    t.table = T.extractTable(t.grid, { merges: t.merges, header: t.header });
    t.empty = !t.table.columns.length || !t.table.rows.length;
  }
  function propose(t) { t.mapping = T.suggest(t.table, state.targets); t.ruleNote = ''; }

  function readFiles(files) {
    var list = Array.prototype.slice.call(files), before = state.tables.length, pending = list.length;
    if (!pending) return;
    msg('file-status', 'note', '읽는 중입니다…');
    function one() { if (--pending === 0) afterLoad(before); }
    list.forEach(function (f) {
      if (/\.pdf$/i.test(f.name)) { readPdf(f).then(one, function (err) { msg('file-status', 'warn', esc(f.name) + ' PDF 를 읽지 못했습니다: ' + esc(err.message)); one(); }); return; }
      readBook(f, function (wb) {
        wb.SheetNames.forEach(function (s) { var g = T.gridFromSheet(wb.Sheets[s]); addTable(f.name, s, g.grid, g.merges, 'xlsx'); });
        one();
      });
    });
  }
  // pdf.js 는 PDF 를 올렸을 때만 불러온다(모듈이라 동적 import). 주소는 문서 기준으로 만든다.
  var pdfjsP = null;
  function pdfjs() {
    if (!pdfjsP) pdfjsP = import(new URL('lib/pdf.min.mjs', document.baseURI).href).then(function (m) {
      m.GlobalWorkerOptions.workerSrc = new URL('lib/pdf.worker.min.mjs', document.baseURI).href; return m;
    });
    return pdfjsP;
  }
  function readPdf(f) {
    return f.arrayBuffer().then(function (buf) {
      return pdfjs().then(function (lib) {
        // cMapUrl 을 주지 않으면 한글 글꼴로 그린 PDF 의 글자가 빈 문자열로 나온다
        return lib.getDocument({ data: new Uint8Array(buf), cMapUrl: new URL('lib/cmaps/', document.baseURI).href, cMapPacked: true }).promise;
      });
    }).then(function (doc) {
      var jobs = [];
      for (var p = 1; p <= doc.numPages; p++) jobs.push(doc.getPage(p).then(function (pg) { return pg.getTextContent(); }));
      return Promise.all(jobs).then(function (pages) {
        pages.forEach(function (tc, i) {
          var t = addTable(f.name, (i + 1) + '쪽', T.pdfItemsToGrid(tc.items), [], 'pdf');
          if (!tc.items.length) t.note = '글자가 없는 쪽입니다(스캔한 그림일 수 있습니다). 이번 판은 글자가 든 PDF 만 읽습니다.';
        });
        doc.destroy();
      });
    });
  }
  function afterLoad(before) {
    var added = state.tables.slice(before), ok = added.filter(function (t) { return !t.empty; });
    var reused = ok.filter(function (t) { return t.ruleNote; }).length;
    msg('file-status', ok.length ? 'ok' : 'warn',
      '표 ' + added.length + '개를 읽었습니다' + (added.length - ok.length ? ' (표를 찾지 못한 것 ' + (added.length - ok.length) + '개)' : '') + '.' +
      (reused ? ' 저장한 규칙을 ' + reused + '개 표에 그대로 썼습니다.' : '') + ' 표를 하나씩 눌러 짝짓기를 확인하고 확정하세요.');
    if (state.cur < 0 || !state.tables[state.cur]) state.cur = state.tables.indexOf(ok[0] || added[0]);
    renderAll();
  }

  var drop = $('drop');
  $('file').onchange = function (e) { readFiles(e.target.files); e.target.value = ''; };
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { if (e.dataTransfer.files.length) readFiles(e.dataTransfer.files); });

  // 예제: 엔진의 가짜 파일 정의를 SheetJS 워크북으로 만든 뒤, 올린 파일과 똑같은 길(xlsx 쓰기→읽기)로 읽는다.
  function sampleBook(f) {
    var wb = XLSX.utils.book_new();
    f.sheets.forEach(function (s) { var ws = XLSX.utils.aoa_to_sheet(s.aoa); ws['!merges'] = s.merges; XLSX.utils.book_append_sheet(wb, ws, s.name); });
    return wb;
  }
  $('btn-sample').onclick = function () {
    var before = state.tables.length, files = T.sample();
    files.forEach(function (f) {
      var wb = XLSX.read(XLSX.write(sampleBook(f), { type: 'array', bookType: 'xlsx' }), { type: 'array', cellDates: true });
      wb.SheetNames.forEach(function (s) { var g = T.gridFromSheet(wb.Sheets[s]); addTable(f.name, s, g.grid, g.merges, 'xlsx'); });
    });
    $('sample-links').innerHTML = '<span class="sub">예제 파일 내려받기:</span>' + files.map(function (f, i) { return '<button class="btn" data-sample="' + i + '">' + esc(f.name) + '</button>'; }).join('');
    afterLoad(before);
  };
  $('sample-links').onclick = function (e) {
    var i = e.target.getAttribute('data-sample'); if (i == null) return;
    var f = T.sample()[+i]; XLSX.writeFile(sampleBook(f), f.name);
  };
  $('btn-clear-files').onclick = function () { state.tables = []; state.cur = -1; msg('file-status', '', ''); renderAll(); };

  /* ================= 표 목록 ================= */
  function status(t) {
    if (t.empty) return ['off', '표를 찾지 못함'];
    if (!t.include) return ['off', '뺌'];
    if (t.confirmed) return ['ok', '확정'];
    return ['todo', '검토 중'];
  }
  function renderTables() {
    $('tables-box').hidden = !state.tables.length;
    $('tables').innerHTML = state.tables.map(function (t, i) {
      var s = status(t);
      return '<li><button data-t="' + i + '" class="' + (i === state.cur ? 'cur' : '') + '"><b>' + esc(t.sheet) + '</b>' +
        '<span class="meta">' + esc(t.file) + ' · ' + t.table.rows.length + '줄 · 열 ' + t.table.columns.length + '개</span>' +
        '<span class="chip ' + s[0] + '">' + s[1] + '</span>' + (t.ruleNote ? ' <span class="chip saved">저장 규칙</span>' : '') + '</button></li>';
    }).join('');
  }
  $('tables').onclick = function (e) {
    var b = e.target.closest('button[data-t]'); if (!b) return;
    state.cur = +b.getAttribute('data-t'); renderAll();
    $('s-map').scrollIntoView({ behavior: 'smooth', block: 'start' });
  };
  $('btn-confirm-all').onclick = function () {
    var n = 0;
    state.tables.forEach(function (t) { if (!t.empty && t.include && !t.confirmed) { confirmTable(t); n++; } });
    msg('file-status', 'ok', n + '개 표를 제안대로 확정했습니다. 결과의 「확인할 것」을 꼭 훑어보세요.');
    renderAll();
  };

  /* ================= 3. 짝짓기 ================= */
  function cur() { return state.tables[state.cur]; }
  function srcOptions(t, sel) {
    var o = ['<option value="">(비움)</option>'];
    T.candidates(t.table).forEach(function (c) { o.push('<option value="' + esc(c.key) + '"' + (sel === c.key ? ' selected' : '') + '>' + esc(c.label) + '</option>'); });
    [['sheet', '시트·쪽 이름'], ['file', '파일 이름'], ['const', '직접 입력…']].forEach(function (p) { o.push('<option value="' + p[0] + '"' + (sel === p[0] ? ' selected' : '') + '>' + p[1] + '</option>'); });
    return o.join('');
  }
  function byChip(m) {
    if (m.by === 'ai') return '<span class="chip ai">AI 초안 — 검토 필요</span>';
    if (m.by === 'saved') return m.src ? '<span class="chip saved">저장 규칙</span>' : '';
    if (m.by === 'user') return '<span class="chip">사람이 고름</span>';
    if (!m.src) return '';
    return '<span class="chip ' + ({ '높음': 'hi', '보통': 'mid', '낮음': 'lo' })[m.conf] + '">신뢰 ' + esc(m.conf) + ' ' + Math.round((m.score || 0) * 100) + '</span>';
  }
  function renderMap() {
    var t = cur();
    $('s-map').hidden = !t;
    if (!t) return;
    var tb = t.table, h = tb.header;
    var head = '<p><b>' + esc(t.file) + '</b> · ' + esc(t.sheet) + ' — ' + tb.rows.length + '줄 · 열 ' + tb.columns.length + '개' +
      (tb.title ? ' · 제목 「' + esc(tb.title) + '」' : '') + '</p>';
    if (tb.meta.length) head += '<p class="sub">문서 정보(머리글 위): ' + tb.meta.map(function (m) { return '<code>' + esc(m.key) + ' = ' + esc(T.str(m.value)) + '</code>'; }).join(' ') + '</p>';
    if (t.ruleNote) head += '<div class="note">' + esc(t.ruleNote) + '</div>';
    if (t.note) head += '<div class="warn">' + esc(t.note) + '</div>';
    if (t.empty) head += '<div class="warn">이 ' + (t.kind === 'pdf' ? '쪽' : '시트') + '에서 표를 찾지 못했습니다. 머리글 행을 직접 정해 다시 읽어 보세요.</div>';
    if (t.confirmed) head += '<div class="ok">확정한 표입니다. 고치면 확정이 풀립니다.</div>';
    $('map-head').innerHTML = head;
    $('h-start').value = h.start + 1; $('h-count').value = h.count;
    $('btn-confirm').textContent = t.confirmed ? '확정 풀기' : '이 표 확정';
    $('btn-exclude').textContent = t.include ? '이 표 빼기' : '이 표 다시 넣기';
    $('btn-confirm').disabled = t.empty || !t.include;
    if (!$('rule-name').value || $('rule-name').getAttribute('data-for') !== String(t.id)) {
      $('rule-name').value = t.file.replace(/\.[^.]+$/, '') + ' · ' + t.sheet; $('rule-name').setAttribute('data-for', t.id);
    }
    renderPreview(t);
    var out = t.empty ? { rows: [], issues: [] } : T.buildOutput(tb, t.mapping, state.targets, { file: t.file, sheet: t.sheet });
    $('maptable').innerHTML = state.targets.map(function (x, i) {
      var m = t.mapping[x.name] || { src: '' };
      var pv = out.rows.slice(0, 3).map(function (r) { return esc(T.str(r.values[x.name])) + mark(r.marks[x.name]); }).filter(function (s) { return s; }).join('<br>');
      return '<tr class="' + (!m.src && x.required ? 'miss' : '') + (m.by === 'ai' ? ' by-ai' : '') + '"><td>' + esc(x.name) + (x.required ? ' *' : '') + '<small>' + T.TYPES[x.type] + '</small></td>' +
        '<td><select data-map="' + i + '" aria-label="' + esc(x.name) + ' 원본">' + srcOptions(t, m.src) + '</select>' +
        (m.src === 'const' ? '<input type="text" class="const" data-const="' + i + '" value="' + esc(m.value || '') + '" placeholder="모든 줄에 넣을 값">' : '') + '</td>' +
        '<td class="why">' + byChip(m) + '<br>' + esc(m.why || '') + '</td><td class="pv">' + (pv || '<span class="sub">—</span>') + '</td></tr>';
    }).join('');
    var iss = out.issues;
    var html = '';
    if (iss.length) html += '<details open><summary>확인할 것 ' + iss.length + '건</summary><div class="warn issues"><ul>' +
      iss.slice(0, 50).map(function (i) { return '<li>' + i.row + '행 · ' + esc(i.target) + ' — ' + esc(i.msg) + '</li>'; }).join('') + (iss.length > 50 ? '<li>…</li>' : '') + '</ul></div></details>';
    if (tb.skipped.length) html += '<details><summary>버린 줄 ' + tb.skipped.length + '개</summary><div class="note"><ul>' +
      tb.skipped.map(function (s) { return '<li>' + s.row + '행 · ' + esc(s.reason) + ' — ' + esc(s.text) + '</li>'; }).join('') + '</ul></div></details>';
    $('map-issues').innerHTML = html;
  }
  function renderPreview(t) {
    var tb = t.table, g = tb.grid, h = tb.header, skip = {};
    tb.skipped.forEach(function (s) { skip[s.row - 1] = true; });
    var lim = Math.min(g.length, Math.max(h.start + h.count + 10, 14)), width = tb.width || 1;
    var html = '<table><thead><tr><th class="rn">행</th>';
    for (var c = 0; c < width; c++) html += '<th>' + T.colName(c) + '</th>';
    html += '</tr></thead><tbody>';
    for (var r = 0; r < lim; r++) {
      var cls = r >= h.start && r < h.start + h.count ? 'hd' : r < h.start ? 'above' : skip[r] ? 'skip' : '';
      html += '<tr class="' + cls + '"><td class="rn">' + (r + 1) + '</td>';
      for (c = 0; c < width; c++) html += '<td class="' + (tb.filled[r + ',' + c] ? 'mf' : '') + '" title="' + esc(T.str(g[r][c])) + '">' + esc(T.str(g[r][c])) + '</td>';
      html += '</tr>';
    }
    if (g.length > lim) html += '<tr><td class="rn">…</td><td colspan="' + width + '" class="sub">' + (g.length - lim) + '줄 더 있음</td></tr>';
    $('preview').innerHTML = html + '</tbody></table>';
  }
  function mark(k) { return k && k !== 'doc' ? '<span class="mk ' + k + '">' + T.MARKS[k] + '</span>' : ''; }
  function edited(t) { t.confirmed = false; }

  $('maptable').addEventListener('change', function (e) {
    var t = cur(), el = e.target, i = el.getAttribute('data-map'), ci = el.getAttribute('data-const');
    if (i != null) {
      var name = state.targets[+i].name, prev = t.mapping[name] || {};
      t.mapping[name] = { src: el.value, value: el.value === 'const' ? (prev.value || '') : undefined, why: '사람이 고름', by: 'user' };
    } else if (ci != null) { t.mapping[state.targets[+ci].name].value = el.value; }
    else return;
    edited(t); renderAll();
  });
  $('h-apply').onclick = function () {
    var t = cur(); if (!t) return;
    var s = Math.max(1, parseInt($('h-start').value, 10) || 1), n = Math.max(1, Math.min(5, parseInt($('h-count').value, 10) || 1));
    t.header = { start: s - 1, count: n }; extract(t); if (!t.empty) { propose(t); t.include = true; }
    edited(t); renderAll();
  };
  $('h-auto').onclick = function () { var t = cur(); if (!t) return; t.header = null; extract(t); if (!t.empty) propose(t); edited(t); renderAll(); };
  $('btn-resuggest').onclick = function () { var t = cur(); if (!t || t.empty) return; propose(t); edited(t); renderAll(); };
  $('btn-exclude').onclick = function () { var t = cur(); if (!t) return; t.include = !t.include; if (!t.include) t.confirmed = false; renderAll(); };
  function confirmTable(t) {
    t.confirmed = true;
    if (!$('save-rule').checked) return;
    var name = t.id === +$('rule-name').getAttribute('data-for') ? T.str($('rule-name').value) : '';
    name = name || (t.file.replace(/\.[^.]+$/, '') + ' · ' + t.sheet);
    var rule = T.makeRule(name, t.table, t.mapping, state.targets);
    state.rules = state.rules.filter(function (r) { return r.name !== name; }).concat([rule]);
    lsSet(LS.rules, state.rules);
  }
  $('btn-confirm').onclick = function () {
    var t = cur(); if (!t) return;
    if (t.confirmed) t.confirmed = false; else confirmTable(t);
    var next = state.tables.findIndex(function (x) { return !x.empty && x.include && !x.confirmed; });
    msg('map-status', 'ok', t.confirmed ? '확정했습니다.' + ($('save-rule').checked ? ' 짝짓기 규칙도 저장했습니다.' : '') + (next >= 0 ? ' 다음 검토할 표: ' + esc(state.tables[next].sheet) : ' 모든 표를 확정했습니다.') : '확정을 풀었습니다.');
    renderAll();
  };

  /* ================= AI 보조 (선택) ================= */
  var ai = lsGet(LS.ai, { provider: 'openai', key: '' });
  $('ai-provider').value = ai.provider || 'openai'; $('ai-key').value = ai.key || '';
  $('btn-ai-save').onclick = function () { ai = { provider: $('ai-provider').value, key: $('ai-key').value.trim() }; lsSet(LS.ai, ai); msg('ai-status', 'ok', '저장했습니다. 키는 이 브라우저에만 있습니다.'); };
  $('btn-ai-forget').onclick = function () { ai = { provider: $('ai-provider').value, key: '' }; $('ai-key').value = ''; lsSet(LS.ai, ai); msg('ai-status', 'ok', '키를 지웠습니다.'); };
  function callAi(p) {
    if (ai.provider === 'anthropic') {
      return fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'content-type': 'application/json', 'x-api-key': ai.key, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
        body: JSON.stringify({ model: 'claude-haiku-5-5', max_tokens: 1200, system: p.system, messages: [{ role: 'user', content: p.user }] })
      }).then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error((j.error && j.error.message) || r.status); return j.content.map(function (c) { return c.text || ''; }).join(''); }); });
    }
    return fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'content-type': 'application/json', authorization: 'Bearer ' + ai.key },
      body: JSON.stringify({ model: 'gpt-4o-mini', temperature: 0, messages: [{ role: 'system', content: p.system }, { role: 'user', content: p.user }] })
    }).then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error((j.error && j.error.message) || r.status); return j.choices[0].message.content; }); });
  }
  $('btn-ai').onclick = function () {
    var t = cur(); if (!t || t.empty) return;
    if (!ai.key) { msg('map-status', 'note', 'AI 키가 없습니다. 아래 「6. AI 연결」에서 키를 넣으면 쓸 수 있습니다. 키 없이도 규칙 기반 제안으로 정리할 수 있습니다.'); return; }
    var p = T.aiPrompt(t.table, state.targets);
    if (!window.confirm('이 표의 열 이름과 예시 값(열마다 3개)을 외부 AI(' + (ai.provider === 'anthropic' ? 'Anthropic' : 'OpenAI') + ')로 보냅니다.\n사내 자료라면 외부로 보내도 되는지 먼저 확인하세요.\n\n보낼까요?')) return;
    msg('map-status', 'note', 'AI 에 묻는 중입니다…');
    callAi(p).then(function (text) {
      var r = T.parseAiReply(text, t.table, state.targets);
      state.targets.forEach(function (x) { if (r.map[x.name].src || !t.mapping[x.name] || !t.mapping[x.name].src) t.mapping[x.name] = r.map[x.name]; });
      edited(t);
      msg('map-status', 'note', 'AI 초안을 반영했습니다. 보라색 줄을 검토한 뒤 확정하세요.' + (r.dropped.length ? ' 표에 없는 열을 말한 답 ' + r.dropped.length + '건은 버렸습니다.' : ''));
      renderAll();
    }, function (err) { msg('map-status', 'warn', 'AI 요청이 실패했습니다: ' + esc(err.message) + ' — 규칙 기반 제안은 그대로입니다.'); });
  };

  /* ================= 4. 결과 ================= */
  function outputs() {
    return state.tables.filter(function (t) { return t.confirmed && t.include && !t.empty; })
      .map(function (t) { return T.buildOutput(t.table, t.mapping, state.targets, { file: t.file, sheet: t.sheet }); });
  }
  function renderResult() {
    var outs = outputs();
    $('s-result').hidden = !outs.length;
    if (!outs.length) return;
    var rows = [].concat.apply([], outs.map(function (o) { return o.rows; }));
    var issues = [].concat.apply([], outs.map(function (o) { return o.issues; }));
    var n = T.countMarks(rows);
    $('result-stats').innerHTML = '<div class="stats">' +
      stat('확정한 표', outs.length) + stat('정리된 줄', rows.length) + stat('문서 값', n.doc) +
      stat('병합 채움', n.merge) + stat('형식 변환', n.conv) + stat('계산', n.calc) + stat('문서 정보·이름·직접 입력', n.meta + n.ctx + n.const) +
      '<div class="stat' + (issues.length ? ' bad' : '') + '"><b>' + issues.length + '</b>확인할 것</div></div>';
    $('marks-legend').innerHTML = '값 출처: 표시 없음 = 문서 값 ' + ['merge', 'conv', 'meta', 'ctx', 'const', 'calc'].map(mark).join(' ');
    var lim = 200;
    var html = '<table><thead><tr>' + state.targets.map(function (x) { return '<th>' + esc(x.name) + '</th>'; }).join('') + '<th>출처</th></tr></thead><tbody>' +
      rows.slice(0, lim).map(function (r) {
        return '<tr>' + state.targets.map(function (x) {
          var v = r.values[x.name];
          return '<td class="' + (typeof v === 'number' ? 'num' : '') + '">' + esc(typeof v === 'number' ? v.toLocaleString('ko-KR') : v) + mark(r.marks[x.name]) + '</td>';
        }).join('') + '<td class="sub">' + esc(r.src.file) + ' · ' + esc(r.src.sheet) + ' · ' + r.src.row + '행</td></tr>';
      }).join('') + '</tbody></table>';
    if (rows.length > lim) html += '<p class="sub" style="padding:8px 11px">앞 ' + lim + '줄만 보입니다. 내려받으면 ' + rows.length + '줄 전부 들어갑니다.</p>';
    $('result-table').innerHTML = html;
    $('result-issues').innerHTML = issues.length ? '<details open><summary>확인할 것 ' + issues.length + '건</summary><div class="warn issues"><ul>' +
      issues.slice(0, 100).map(function (i) { return '<li>' + esc(i.file) + ' · ' + esc(i.sheet) + ' · ' + i.row + '행 · ' + esc(i.target) + ' — ' + esc(i.msg) + '</li>'; }).join('') + '</ul></div></details>' : '<div class="ok">확인할 것이 없습니다.</div>';
  }
  function stat(k, v) { return '<div class="stat"><b>' + v + '</b>' + k + '</div>'; }
  $('btn-xlsx').onclick = function () {
    var sh = T.exportSheets(outputs(), state.targets), wb = XLSX.utils.book_new();
    Object.keys(sh).forEach(function (k) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh[k]), k); });
    XLSX.writeFile(wb, '문서정리_결과_' + today() + '.xlsx');
  };
  $('btn-csv').onclick = function () {
    var sh = T.exportSheets(outputs(), state.targets);
    var blob = new Blob(['﻿' + T.toCSV(sh['정리결과'])], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = '문서정리_결과_' + today() + '.csv';
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  };

  /* ================= 5. 규칙 ================= */
  function renderRules() {
    $('rules').innerHTML = state.rules.length ? state.rules.map(function (r, i) {
      return '<li><span><b>' + esc(r.name) + '</b><p class="sub">원본 열 ' + r.signature.length + '개 · 짝 ' + Object.keys(r.map).length + '개 · ' + esc((r.saved || '').slice(0, 10)) + '</p></span>' +
        '<button class="x" data-rdel="' + i + '">지우기</button></li>';
    }).join('') : '<li class="sub">아직 저장한 규칙이 없습니다. 표를 확정하면 여기에 쌓입니다.</li>';
  }
  $('rules').onclick = function (e) {
    var d = e.target.getAttribute('data-rdel'); if (d == null) return;
    state.rules.splice(+d, 1); lsSet(LS.rules, state.rules); renderRules();
  };
  $('btn-rules-export').onclick = function () {
    var blob = new Blob([JSON.stringify({ app: 'build-2611-02', targets: state.targets, rules: state.rules }, null, 2)], { type: 'application/json' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = '문서정리_짝짓기규칙_' + today() + '.json';
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  };
  $('rules-file').onchange = function (e) {
    var f = e.target.files[0]; if (!f) return;
    f.text().then(function (txt) {
      var list = T.validRules(txt), names = {};
      state.rules.concat(list).forEach(function (r) { names[r.name] = r; });
      state.rules = Object.keys(names).map(function (k) { return names[k]; });
      lsSet(LS.rules, state.rules); renderRules();
      msg('rules-status', list.length ? 'ok' : 'warn', list.length ? '규칙 ' + list.length + '개를 불러왔습니다(같은 이름은 덮어씀). 다음에 올리는 파일부터 씁니다.' : '불러올 규칙이 없습니다.');
    }).catch(function (err) { msg('rules-status', 'warn', 'JSON 을 읽지 못했습니다: ' + esc(err.message)); });
    e.target.value = '';
  };

  function renderAll() { renderTables(); renderMap(); renderResult(); renderRules(); }
  renderTargets(); renderAll();
})();
