/* 폴더 분류 화면 — 판정은 classify.js(Classify)·engine.js(Tidy)가 한다.
 * 여기서는 폴더의 파일을 읽어 grid 로 넘기고, 결과를 그리고, 사람이 고른 종류·고친 규칙을 넘길 뿐이다. */
(function () {
  'use strict';
  var T = window.Tidy, C = window.Classify;
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var LS = 'b261102.doctypes';
  function msg(id, cls, html) { $(id).innerHTML = html ? '<div class="' + cls + '">' + html + '</div>' : ''; }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function download(name, blob) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  /* ---------- 규칙 ---------- */
  function loadRules() {
    try {
      var saved = localStorage.getItem(LS);
      if (saved) { var c = C.checkRules(saved); if (c.ok) return c.rules; }
    } catch (e) { /* 저장 공간을 못 쓰면 기본 규칙 */ }
    return C.defaultRules();
  }
  var state = { rules: loadRules(), docs: [], skipped: [], assign: {}, res: null, sample: false };
  function saveRules() { try { localStorage.setItem(LS, C.rulesToJSON(state.rules)); } catch (e) { /* 이번 화면에서만 */ } }

  /* ---------- 파일 읽기 ---------- */
  var pdfjsP = null;
  function pdfjs() {
    if (!pdfjsP) pdfjsP = import(new URL('lib/pdf.min.mjs', document.baseURI).href).then(function (m) {
      m.GlobalWorkerOptions.workerSrc = new URL('lib/pdf.worker.min.mjs', document.baseURI).href; return m;
    });
    return pdfjsP;
  }
  function pdfSheets(bytes) {
    return pdfjs().then(function (lib) {
      return lib.getDocument({ data: bytes, cMapUrl: new URL('lib/cmaps/', document.baseURI).href, cMapPacked: true }).promise;
    }).then(function (doc) {
      var jobs = [];
      for (var p = 1; p <= doc.numPages; p++) jobs.push(doc.getPage(p).then(function (pg) { return pg.getTextContent(); }));
      return Promise.all(jobs).then(function (pages) {
        doc.destroy();
        return pages.map(function (tc, i) { return { name: (i + 1) + '쪽', grid: T.pdfItemsToGrid(tc.items), merges: [] }; });
      });
    });
  }
  function bookSheets(wb) {
    return wb.SheetNames.map(function (s) { var g = T.gridFromSheet(wb.Sheets[s]); return { name: s, grid: g.grid, merges: g.merges }; });
  }
  function readOne(file, path) {
    return file.arrayBuffer().then(function (buf) {
      if (/\.pdf$/i.test(path)) return pdfSheets(new Uint8Array(buf));
      if (/\.csv$/i.test(path)) return [{ name: 'CSV', grid: T.csvToGrid(new TextDecoder('utf-8').decode(buf)), merges: [] }];
      return bookSheets(XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true }));
    });
  }
  // list: [{ file, path }]
  function readList(list, note) {
    state.docs = []; state.skipped = []; state.assign = {};
    msg('fc-status', 'note', '읽는 중입니다…');
    var jobs = list.map(function (x) {
      if (!C.isSupported(x.path)) { state.skipped.push({ path: x.path, reason: '엑셀·CSV·PDF 가 아님' }); return null; }
      return readOne(x.file, x.path).then(function (sheets) { return C.docOf(x.path, sheets); },
        function (err) { state.skipped.push({ path: x.path, reason: '읽지 못함: ' + err.message }); return null; });
    });
    return Promise.all(jobs).then(function (docs) {
      state.docs = docs.filter(Boolean).sort(function (a, b) { return a.path < b.path ? -1 : 1; });
      rerun();
      msg('fc-status', state.docs.length ? 'ok' : 'warn', (note || '') + '파일 ' + state.docs.length + '개를 읽어 분류했습니다' +
        (state.skipped.length ? ' (건너뜀 ' + state.skipped.length + '개)' : '') + '.');
      if (state.docs.length) $('s-groups').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }
  function relPath(f) { return f.webkitRelativePath || f.relativePath || f.name; }
  $('fc-folder').onchange = function (e) {
    var list = Array.prototype.map.call(e.target.files, function (f) { return { file: f, path: relPath(f) }; });
    state.sample = false; readList(list); e.target.value = '';
  };
  $('fc-files').onchange = function (e) {
    var list = Array.prototype.map.call(e.target.files, function (f) { return { file: f, path: f.name }; });
    state.sample = false; readList(list); e.target.value = '';
  };
  // 끌어다 놓은 폴더는 안을 차례로 펼친다(webkitGetAsEntry). 경로는 폴더 이름부터 적는다.
  function walkEntry(entry, out) {
    if (entry.isFile) return new Promise(function (res) { entry.file(function (f) { out.push({ file: f, path: entry.fullPath.replace(/^\//, '') }); res(); }, function () { res(); }); });
    if (!entry.isDirectory) return Promise.resolve();
    var reader = entry.createReader();
    return new Promise(function (res) {
      var all = [];
      (function more() {
        reader.readEntries(function (batch) {
          if (!batch.length) return Promise.all(all.map(function (en) { return walkEntry(en, out); })).then(res);
          all = all.concat(Array.prototype.slice.call(batch)); more();
        }, function () { res(); });
      })();
    });
  }
  var drop = $('fc-drop');
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) {
    var items = e.dataTransfer.items, out = [], jobs = [];
    if (items && items.length && items[0].webkitGetAsEntry) {
      for (var i = 0; i < items.length; i++) { var en = items[i].webkitGetAsEntry(); if (en) jobs.push(walkEntry(en, out)); }
      Promise.all(jobs).then(function () { state.sample = false; readList(out); });
    } else {
      state.sample = false;
      readList(Array.prototype.map.call(e.dataTransfer.files, function (f) { return { file: f, path: f.name }; }));
    }
  });

  // 예제 폴더: classify.js 의 가짜 파일 정의를 실제 xlsx·pdf 바이트로 만든 뒤, 올린 파일과 같은 길로 읽는다.
  function sampleFiles() {
    return C.sampleFolder().map(function (f) {
      var bytes;
      if (f.pdf) bytes = C.samplePdf();
      else if (f.text != null) bytes = new TextEncoder().encode(f.text);
      else {
        var wb = XLSX.utils.book_new();
        f.sheets.forEach(function (s) { var ws = XLSX.utils.aoa_to_sheet(s.aoa); ws['!merges'] = s.merges; XLSX.utils.book_append_sheet(wb, ws, s.name); });
        bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
      }
      return { file: new File([bytes], C.baseName(f.path)), path: '예제폴더/' + f.path };
    });
  }
  function runSample() {
    state.sample = true;
    return readList(sampleFiles(), '예제 폴더(가짜 데이터)의 ');
  }
  $('btn-fc-sample').onclick = runSample;
  $('btn-fc-clear').onclick = function () { state.docs = []; state.skipped = []; state.assign = {}; state.res = null; msg('fc-status', '', ''); render(); };

  /* ---------- 분류·추출 다시 ---------- */
  function rerun() { state.res = state.docs.length ? C.runFolder(state.docs, state.rules, state.assign) : null; render(); }

  /* ---------- 그리기 ---------- */
  function stat(k, v, bad) { return '<div class="stat' + (bad ? ' bad' : '') + '"><b>' + v + '</b>' + esc(k) + '</div>'; }
  function chip(it) {
    var lb = C.statusLabel(it), cls = { '분류됨': 'ok', '확인 필요': 'todo', '미분류': 'off', '사람이 고름': 'saved', '사람이 뺌': 'off' }[lb];
    return '<span class="chip ' + cls + '">' + esc(lb) + '</span>';
  }
  function typeSelect(it) {
    var o = ['<option value="__auto"' + (it.by !== 'user' ? ' selected' : '') + '>자동(' + esc(it.cls.status === 'ok' ? (C.findType(state.rules, it.cls.type) || {}).name : C.statusLabel({ by: 'rule', cls: it.cls })) + ')</option>'];
    state.rules.types.forEach(function (t) { o.push('<option value="' + esc(t.id) + '"' + (it.by === 'user' && it.type === t.id ? ' selected' : '') + '>' + esc(t.name) + '</option>'); });
    o.push('<option value=""' + (it.by === 'user' && !it.type ? ' selected' : '') + '>뽑지 않음</option>');
    return '<select data-path="' + esc(it.doc.path) + '" aria-label="' + esc(it.doc.name) + ' 종류 고르기">' + o.join('') + '</select>';
  }
  function fileRow(it) {
    var rk = it.cls.ranking.slice(0, 2).map(function (r) { return esc(r.name) + ' ' + r.score + '점'; }).join(' · ');
    var why = it.cls.status === 'ok' ? it.cls.why : it.cls.why + (it.cls.hits.length ? ' — ' + C.whyText(it.cls.hits) : '');
    return '<li class="fc-file" data-path="' + esc(it.doc.path) + '"><div class="fc-name">' +
      (it.doc.folder ? '<span class="fc-dir">' + esc(it.doc.folder) + '/</span>' : '') + '<b>' + esc(it.doc.name) + '</b> ' + chip(it) +
      (it.ext ? ' <span class="sub">뽑은 행 ' + it.ext.kept.length + ' · 뺀 행 ' + it.ext.dropped.length + '</span>' : '') +
      '</div><div class="fc-why"><span class="fc-score">' + it.cls.score + '점</span> ' + esc(why) + '<br><span class="sub">순위: ' + rk + '</span></div>' +
      '<div class="fc-pick">' + typeSelect(it) + '</div></li>';
  }
  function groupBox(id, title, items, extra) {
    return '<div class="fc-group" data-group="' + esc(id) + '"><h3>' + esc(title) + ' <span class="sub">파일 ' + items.length + '개' + (extra || '') + '</span></h3>' +
      (items.length ? '<ul class="fc-files">' + items.map(fileRow).join('') + '</ul>' : '<p class="sub">없음</p>') + '</div>';
  }
  function render() {
    var r = state.res;
    $('s-groups').hidden = !r; $('s-extract').hidden = !r;
    renderRulesSum();
    if (!r) return;
    $('fc-stats').innerHTML = '<div class="stats">' + stat('읽은 파일', r.items.length) +
      r.groups.map(function (g) { return stat(g.name, g.items.length); }).join('') +
      stat('확인 필요', r.check.length, r.check.length) + stat('미분류·뺌', r.none.length, r.none.length) + stat('건너뜀', state.skipped.length) + '</div>';
    var html = r.groups.map(function (g) { return groupBox(g.id, g.name, g.items, ' · 뽑은 행 ' + g.kept + ' · 뺀 행 ' + g.dropped); }).join('');
    if (r.check.length) html += groupBox('check', '확인 필요', r.check, ' · 종류를 골라 주세요');
    if (r.none.length) html += groupBox('none', '미분류', r.none, ' · 어느 종류에도 못 미침 — 필요하면 종류를 고르세요');
    if (state.skipped.length) html += '<div class="fc-group" data-group="skipped"><h3>건너뜀 <span class="sub">파일 ' + state.skipped.length + '개</span></h3><ul class="fc-files">' +
      state.skipped.map(function (s) { return '<li class="fc-file"><div class="fc-name"><b>' + esc(s.path) + '</b></div><div class="fc-why sub">' + esc(s.reason) + '</div></li>'; }).join('') + '</ul></div>';
    if (state.sample) html += '<p class="sub">예제 파일은 리포의 <code>samples/</code> 폴더에도 실제 파일로 있습니다(로컬 스크립트 시연용): ' +
      C.sampleFolder().filter(function (f) { return !f.text; }).map(function (f) { return '<a href="samples/' + f.path.split('/').map(encodeURIComponent).join('/') + '" download>' + esc(C.baseName(f.path)) + '</a>'; }).join(' · ') + '</p>';
    $('fc-groups').innerHTML = html;
    renderExtract();
  }
  function mark(k) { return k && k !== 'doc' ? '<span class="mk ' + k + '">' + T.MARKS[k] + '</span>' : ''; }
  function renderExtract() {
    var r = state.res, html = '';
    r.groups.forEach(function (g) {
      var t = C.findType(state.rules, g.id);
      var conds = t.conditions.length ? t.conditions.map(C.condText).join(' 그리고 ') : '없음(모든 행)';
      html += '<div class="fc-ext" data-type="' + esc(g.id) + '"><h3>' + esc(g.name) + '</h3>' +
        '<p class="sub">추출 조건: ' + esc(conds) + '</p>';
      if (!g.items.length) { html += '<p class="sub">이 종류로 분류된 파일이 없습니다.</p></div>'; return; }
      html += '<div class="stats">' + stat('파일', g.items.length) + stat('조건에 맞은 행', g.kept) + stat('뺀 행', g.dropped) + '</div>';
      var rows = [], drops = [], iss = [], skippedT = [];
      g.items.forEach(function (it) {
        rows = rows.concat(it.ext.kept); drops = drops.concat(it.ext.dropped.map(function (d) { return { it: it, d: d }; }));
        iss = iss.concat(it.ext.issues);
        it.ext.skippedTables.forEach(function (s) { skippedT.push(it.doc.path + ' · ' + s.sheet + ' — ' + s.reason); });
      });
      html += '<div class="tablewrap fc-table"><table><thead><tr>' + t.fields.map(function (f) { return '<th>' + esc(f.name) + '</th>'; }).join('') + '<th>출처</th></tr></thead><tbody>' +
        rows.map(function (row) {
          return '<tr>' + t.fields.map(function (f) {
            var v = row.values[f.name];
            return '<td class="' + (typeof v === 'number' ? 'num' : '') + '">' + esc(typeof v === 'number' ? v.toLocaleString('ko-KR') : v) + mark(row.marks[f.name]) + '</td>';
          }).join('') + '<td class="sub">' + esc(row.src.file) + ' · ' + esc(row.src.sheet) + ' · ' + row.src.row + '행</td></tr>';
        }).join('') + '</tbody></table></div>';
      if (drops.length) html += '<details><summary>뺀 행 ' + drops.length + '개</summary><div class="note"><ul>' +
        drops.map(function (x) { return '<li>' + esc(x.it.doc.name) + ' · ' + esc(x.d.row.src.sheet) + ' · ' + x.d.row.src.row + '행 — ' + esc(x.d.reasons.join(' · ')) + '</li>'; }).join('') + '</ul></div></details>';
      if (iss.length) html += '<details><summary>확인할 것 ' + iss.length + '건</summary><div class="warn issues"><ul>' +
        iss.map(function (i) { return '<li>' + esc(C.baseName(i.file)) + ' · ' + esc(i.sheet) + ' · ' + i.row + '행 · ' + esc(i.target) + ' — ' + esc(i.msg) + '</li>'; }).join('') + '</ul></div></details>';
      if (skippedT.length) html += '<details><summary>읽지 않은 시트 ' + skippedT.length + '개</summary><div class="note"><ul>' + skippedT.map(function (s) { return '<li>' + esc(s) + '</li>'; }).join('') + '</ul></div></details>';
      html += '</div>';
    });
    $('fc-extract').innerHTML = html;
  }
  $('fc-groups').addEventListener('change', function (e) {
    var p = e.target.getAttribute('data-path'); if (p == null) return;
    if (e.target.value === '__auto') delete state.assign[p]; else state.assign[p] = e.target.value;
    rerun();
  });
  $('btn-fc-xlsx').onclick = function () {
    if (!state.res) return;
    var book = C.exportBook(state.res, state.rules, state.skipped), wb = XLSX.utils.book_new();
    Object.keys(book).forEach(function (k) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(book[k]), k); });
    XLSX.writeFile(wb, '폴더분류_결과_' + today() + '.xlsx');
  };

  /* ---------- 규칙 편집 ---------- */
  function renderRulesSum() {
    $('fc-rules-sum').innerHTML = state.rules.types.map(function (t) {
      var m = t.match || {}, n = Object.keys(C.KINDS).map(function (k) { return (m[k] || []).length ? C.KINDS[k] + ' ' + m[k].length : ''; }).filter(Boolean).join(' · ');
      return '<tr><td><b>' + esc(t.name) + '</b><br><span class="sub">' + esc(t.id) + '</span></td><td class="num">' + t.minScore + '</td><td>' + esc(n) + '</td>' +
        '<td class="wrapcell">' + esc(t.fields.map(function (f) { return f.name + (f.required ? '*' : ''); }).join(' · ')) + '</td>' +
        '<td class="wrapcell">' + esc(t.conditions.length ? t.conditions.map(C.condText).join(' 그리고 ') : '없음') + '</td></tr>';
    }).join('') + '<tr><td colspan="5" class="sub">확인 필요 기준: 1·2위 점수 차이 ' + state.rules.closeMargin + '점 이내</td></tr>';
    var ta = $('fc-rules-json');
    if (document.activeElement !== ta) ta.value = C.rulesToJSON(state.rules);
  }
  function applyRules(text, note) {
    var c = C.checkRules(text);
    if (!c.ok) { msg('fc-rules-status', 'warn', '규칙을 적용하지 않았습니다. 고칠 곳:<ul>' + c.errors.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul>'); return false; }
    state.rules = c.rules; saveRules();
    // 사라진 종류를 고른 파일은 자동으로 되돌린다
    Object.keys(state.assign).forEach(function (p) { if (state.assign[p] && !C.findType(state.rules, state.assign[p])) delete state.assign[p]; });
    $('fc-rules-json').value = C.rulesToJSON(state.rules);
    rerun();
    msg('fc-rules-status', 'ok', (note || '규칙을 적용했습니다.') + ' 문서 종류 ' + state.rules.types.length + '개' + (state.docs.length ? ' — 파일 ' + state.docs.length + '개를 다시 분류했습니다.' : '.'));
    return true;
  }
  $('btn-fc-rules-apply').onclick = function () { applyRules($('fc-rules-json').value); };
  $('btn-fc-rules-reset').onclick = function () { applyRules(C.defaultRules(), '기본 규칙(예제 종류 3가지)으로 되돌렸습니다.'); };
  $('btn-fc-rules-dl').onclick = function () { download('doc-types.json', new Blob([C.rulesToJSON(state.rules)], { type: 'application/json' })); };
  $('fc-rules-up').onchange = function (e) {
    var f = e.target.files[0]; if (!f) return;
    f.text().then(function (txt) { applyRules(txt, '「' + esc(f.name) + '」 규칙을 적용했습니다.'); });
    e.target.value = '';
  };

  render();
  // index.html 의 「예제 폴더로 바로 보기」 링크(folder.html#example)로 오면 예제를 바로 돌린다
  if (location.hash === '#example') runSample();
})();
