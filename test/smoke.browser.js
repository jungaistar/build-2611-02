/* 브라우저에 실제로 띄워 화면이 그려지는지 본다 — node test/smoke.browser.js
 * 규칙 테스트가 다 통과해도 app.js 오타 하나면 빈 화면이 된다. playwright 가 없으면 건너뛴다.
 * PDF 는 pdf.js 를 동적으로 불러오는 길(워커 주소·cmaps)이 깨지기 쉬워서, 작은 PDF 를 즉석에서 만들어 올려 본다. */
'use strict';
var http = require('http'), fs = require('fs'), path = require('path');
var chromium;
try { chromium = require('playwright').chromium; }
catch (e) { console.log('playwright 가 없어 화면 연기 테스트를 건너뜁니다.'); process.exit(0); }
var ROOT = path.join(__dirname, '..');
var MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.bcmap': 'application/octet-stream' };
var srv = http.createServer(function (q, s) {
  var p = path.join(ROOT, decodeURIComponent(q.url.split('?')[0]).replace(/\/$/, '/index.html'));
  fs.readFile(p, function (e, b) { if (e) { s.writeHead(404); return s.end(); } s.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' }); s.end(b); });
});

// 글자 좌표만 있는 표 한 장짜리 PDF (Helvetica, 영문) — 가짜 데이터
function makePdf() {
  var rows = [['Part No', 'Description', 'Qty', 'Unit Price'], ['AX-100', 'Gear Pump', '12', '15000'], ['AX-200', 'Cylinder', '3', '128000']];
  var xs = [50, 160, 330, 420], y = 700, ops = 'BT /F1 11 Tf\n';
  rows.forEach(function (r, i) { r.forEach(function (s, k) { ops += '1 0 0 1 ' + xs[k] + ' ' + (y - i * 20) + ' Tm (' + s + ') Tj\n'; }); });
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
  out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + x + '\n%%EOF';
  return Buffer.from(out, 'latin1');
}

var passed = 0, failed = 0, v10 = 0;
function ok(c, label, d) { if (c) passed++; else { failed++; console.log('  X ' + label + (d ? '  ' + d : '')); } }
srv.listen(0, async function () {
  var base = 'http://127.0.0.1:' + srv.address().port + '/';
  var br = await chromium.launch(); var pg = await br.newPage({ viewport: { width: 1280, height: 900 } });
  var errs = []; pg.on('pageerror', function (e) { errs.push(e.message); });
  try {
    await pg.goto(base);
    ok(await pg.$$eval('#targets tr', function (t) { return t.length; }) === 10, '기본 목표 열 10개');
    await pg.click('#btn-sample');
    await pg.waitForSelector('#tables button', { timeout: 5000 });
    ok(await pg.$$eval('#tables button', function (t) { return t.length; }) === 5, '예제 표 5개(시트 5장)');
    ok(await pg.isVisible('#s-map'), '짝짓기 영역이 열린다');
    ok(await pg.$$eval('#preview tr.hd', function (t) { return t.length; }) === 2, '견적서 머리글 두 줄 표시');
    ok(await pg.$$eval('#preview td.mf', function (t) { return t.length; }) > 0, '병합 채움 칸 표시');
    var vendor = await pg.$eval('#maptable select[data-map="7"]', function (s) { return s.value; });
    ok(vendor === 'm:업체명', '공급업체 ← 문서 정보 업체명 제안', vendor);
    ok(/가람정밀/.test(await pg.textContent('#maptable')), '미리보기에 변환된 값');
    await pg.click('#btn-confirm');
    ok(await pg.isVisible('#s-result'), '확정하면 결과가 열린다');
    ok(/9행/.test(await pg.textContent('#result-issues')), '금액 불일치(9행)가 확인할 것에 뜬다');
    // 두 번째 표에서 사람이 하나 고친다
    await pg.click('#tables button[data-t="2"]');
    await pg.selectOption('#maptable select[data-map="9"]', "c:Lead Time");
    ok(/사람이 고름/.test(await pg.textContent('#maptable')), '사람이 고친 짝 표시');
    await pg.click('#btn-confirm-all');
    var rows = await pg.$$eval('#result-table tbody tr', function (t) { return t.length; });
    ok(rows === 20, '모두 확정 → 5개 표 20줄', rows);
    ok(await pg.$$eval('#result-table .mk.merge', function (t) { return t.length; }) > 0, '결과에 병합 채움 표시');
    ok(await pg.$$eval('#result-table .mk.calc', function (t) { return t.length; }) > 0, '결과에 계산 표시');
    ok(await pg.$$eval('#rules li .x', function (t) { return t.length; }) === 5, '확정한 표마다 규칙 저장');
    var dl = await Promise.all([pg.waitForEvent('download'), pg.click('#btn-xlsx')]);
    ok(/\.xlsx$/.test(dl[0].suggestedFilename()), '엑셀 내려받기', dl[0].suggestedFilename());
    var dl2 = await Promise.all([pg.waitForEvent('download'), pg.click('#btn-csv')]);
    ok(/\.csv$/.test(dl2[0].suggestedFilename()), 'CSV 내려받기');
    // 같은 예제를 다시 올리면 저장한 규칙을 쓴다
    await pg.click('#btn-clear-files');
    await pg.click('#btn-sample');
    ok(/저장한 규칙을 5개 표/.test(await pg.textContent('#file-status')), '다시 올리면 저장 규칙 재사용');
    await pg.click('#tables button[data-t="2"]');
    ok(await pg.$eval('#maptable select[data-map="9"]', function (s) { return s.value; }) === 'c:Lead Time', '사람이 고친 짝이 그대로 돌아온다');
    // AI 키 없이 누르면 안내만
    await pg.click('#btn-ai');
    ok(/키가 없습니다/.test(await pg.textContent('#map-status')), 'AI 키가 없으면 안내만 하고 규칙 기반 유지');
    // PDF
    await pg.click('#btn-clear-files');
    await pg.setInputFiles('#file', { name: 'parts.pdf', mimeType: 'application/pdf', buffer: makePdf() });
    await pg.waitForSelector('#tables button', { timeout: 10000 });
    var price = await pg.$eval('#maptable select[data-map="5"]', function (s) { return s.value; });
    ok(price === 'c:Unit Price', 'PDF 표에서 단가 열을 찾는다', price);
    ok(/Gear Pump/.test(await pg.textContent('#maptable')), 'PDF 글자가 칸으로 들어온다');
    // 좁은 화면 가로 넘침
    await pg.setViewportSize({ width: 360, height: 800 });
    var over = await pg.evaluate(function () { return document.documentElement.scrollWidth - document.documentElement.clientWidth; });
    ok(over <= 0, '360px 에서 가로 넘침 없음', over);
    await pg.goto(base + 'guide.html');
    ok(/짝짓기/.test(await pg.textContent('main')), '사용법 페이지');
    v10 = passed + failed + 1;   // v1.0 화면 테스트(맨 끝 「페이지 오류 없음」 포함)
    ok(/폴더 분류·추출/.test(await pg.textContent('main')), '사용법에 폴더 분류·Codex 절');

    /* ===== v1.1 폴더 분류 ===== */
    await pg.setViewportSize({ width: 1280, height: 900 });
    await pg.goto(base + 'folder.html');
    ok(await pg.$$eval('#fc-rules-sum tr', function (t) { return t.length; }) === 4, '폴더 분류: 기본 규칙 3종 요약');
    await pg.click('#btn-fc-sample');
    await pg.waitForSelector('.fc-group[data-group="inspection"] .fc-file', { timeout: 15000 });
    function cnt(sel) { return pg.$$eval(sel, function (t) { return t.length; }); }
    var gq = await cnt('.fc-group[data-group="quote"] .fc-file'), gs = await cnt('.fc-group[data-group="statement"] .fc-file'), gi = await cnt('.fc-group[data-group="inspection"] .fc-file');
    ok(gq === 2 && gs === 2 && gi === 2, '예제 폴더 → 견적서 2 · 거래명세서 2 · 검사성적서 2(PDF 포함)', [gq, gs, gi].join('/'));
    ok(await cnt('.fc-group[data-group="check"] .fc-file') === 1, '확인 필요 1');
    ok(await cnt('.fc-group[data-group="none"] .fc-file') === 1, '미분류 1');
    ok(/메모\.txt/.test(await pg.textContent('.fc-group[data-group="skipped"]')), '엑셀·PDF 가 아닌 파일은 건너뜀');
    ok(/파일 이름 「견적」 \+4/.test(await pg.textContent('.fc-group[data-group="quote"]')), '분류 근거(어떤 규칙이 몇 점)');
    ok(/예제폴더\/2026-10\/견적\//.test(await pg.textContent('.fc-group[data-group="quote"] .fc-dir')), '하위 폴더 경로 표시');
    ok(await cnt('.fc-ext[data-type="quote"] tbody tr') === 6, '견적서 추출: 금액 ≥ 30만 6행');
    ok(await cnt('.fc-ext[data-type="inspection"] tbody tr') === 3, '검사성적서 추출: 불합격·NG 3행(PDF 1행 포함)');
    ok(/조건에 맞은 행/.test(await pg.textContent('.fc-ext[data-type="quote"]')) && /뺀 행 5개/.test(await pg.textContent('.fc-ext[data-type="quote"]')), '조건에 맞은 행 · 뺀 행 수');
    ok(await cnt('.fc-ext[data-type="inspection"] .mk.merge') > 0, '추출 결과에 값 출처(병합 채움) 표시');
    // 확인 필요 파일을 사람이 견적서로 고른다
    await pg.selectOption('.fc-group[data-group="check"] select', 'quote');
    ok(await cnt('.fc-group[data-group="quote"] .fc-file') === 3 && await cnt('.fc-ext[data-type="quote"] tbody tr') === 8, '확인 필요를 견적서로 고르면 묶음·추출에 들어간다');
    ok(!(await pg.$('.fc-group[data-group="check"]')), '고른 뒤 확인 필요 묶음이 사라진다');
    var dl3 = await Promise.all([pg.waitForEvent('download'), pg.click('#btn-fc-xlsx')]);
    var XLSX = require('../lib/xlsx.full.min.js'), wbx = XLSX.read(fs.readFileSync(await dl3[0].path()), { type: 'buffer' });
    ok(wbx.SheetNames.join(',') === '분류표,견적서,거래명세서,검사성적서,뺀 행,확인할것,건너뛴 파일', '엑셀: 분류표 + 종류별 시트', wbx.SheetNames.join(','));
    ok(XLSX.utils.sheet_to_json(wbx.Sheets['견적서'], { header: 1 }).length === 9, '엑셀 견적서 시트 8행 + 머리글');
    // 규칙 고치기
    await pg.click('#s-fc-rules summary');
    var rj = JSON.parse(await pg.inputValue('#fc-rules-json'));
    rj.types[0].minScore = 20;
    await pg.fill('#fc-rules-json', JSON.stringify(rj));
    await pg.click('#btn-fc-rules-apply');
    ok(/규칙을 적용했습니다/.test(await pg.textContent('#fc-rules-status')), '규칙 JSON 고쳐 적용');
    ok(await cnt('.fc-group[data-group="quote"] .fc-file') === 1, '최소 점수를 올리면 다시 분류(사람이 고른 1개만 남음)');
    await pg.fill('#fc-rules-json', '{ "types": [ { "id": "a" } ] }');
    await pg.click('#btn-fc-rules-apply');
    ok(/고칠 곳/.test(await pg.textContent('#fc-rules-status')), '틀린 규칙은 적용하지 않고 고칠 곳을 보여 준다');
    await pg.click('#btn-fc-rules-reset');
    ok(await cnt('.fc-group[data-group="quote"] .fc-file') === 3, '기본 규칙으로 되돌리기');
    var dl4 = await Promise.all([pg.waitForEvent('download'), pg.click('#btn-fc-rules-dl')]);
    ok(dl4[0].suggestedFilename() === 'doc-types.json', '규칙 JSON 내려받기');
    // 실제 폴더 올리기 (samples/)
    await pg.setInputFiles('#fc-folder', path.join(ROOT, 'samples'));
    await pg.waitForFunction(function () { return /파일 8개/.test(document.getElementById('fc-status').textContent); }, null, { timeout: 15000 });
    ok(/samples\/2026-10\/견적\//.test(await pg.textContent('.fc-group[data-group="quote"] .fc-dir')), '폴더 올리기: samples/ 하위 폴더 경로');
    ok(await cnt('.fc-group[data-group="inspection"] .fc-file') === 2 && await cnt('.fc-group[data-group="check"] .fc-file') === 1, '폴더 올리기: 예제와 같은 분류(로컬 스크립트와 같은 규칙)');
    // 첫 화면에서 바로 가기 링크
    await pg.goto(base);
    await pg.click('#go-folder-sample');
    await pg.waitForSelector('.fc-group[data-group="quote"] .fc-file', { timeout: 15000 });
    ok(/folder\.html#example$/.test(pg.url()), '첫 화면 「예제 폴더로 바로 보기」 → 폴더 분류 예제');
    await pg.setViewportSize({ width: 360, height: 800 });
    var over2 = await pg.evaluate(function () { return document.documentElement.scrollWidth - document.documentElement.clientWidth; });
    ok(over2 <= 0, '폴더 분류 360px 에서 가로 넘침 없음', over2);
  } catch (e) { failed++; console.log('  X 예외: ' + e.message); }
  ok(!errs.length, '페이지 오류 없음', errs.join(' | '));
  await br.close(); srv.close();
  console.log('\nv1.0 화면 ' + v10 + '개 · v1.1 폴더 분류 ' + (passed + failed - v10) + '개(사용법 절 포함)');
  console.log('\n화면 통과 ' + passed + ' · 실패 ' + failed);
  process.exit(failed ? 1 : 0);
});
