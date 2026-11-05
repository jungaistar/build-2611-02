/* 판정 규칙 고정 — node test/logic.test.js
 * 화면이 아니라 「어느 열이 어디로 가는가」와 「값이 어디서 왔는가」가 틀리면 안 되는 도구다.
 * 예제 파일은 SheetJS 로 실제 xlsx 를 쓰고 다시 읽어서 — 올린 파일과 같은 길로 — 검사한다. */
'use strict';
var T = require('../js/engine.js');
var XLSX = require('../lib/xlsx.full.min.js');
var passed = 0, failed = 0;
function ok(c, label, detail) { if (c) passed++; else { failed++; console.log('  X ' + label + (detail ? '\n      ' + detail : '')); } }
function eq(g, w, label) { ok(JSON.stringify(g) === JSON.stringify(w), label, '기대: ' + JSON.stringify(w) + '  실제: ' + JSON.stringify(g)); }

// 실제 xlsx 로 쓰고 다시 읽는다 — 병합 정보가 파일을 거쳐도 살아 있는지까지 본다
function roundtrip(sheet) {
  var ws = XLSX.utils.aoa_to_sheet(sheet.aoa); ws['!merges'] = sheet.merges;
  var wb = { SheetNames: [sheet.name], Sheets: {} }; wb.Sheets[sheet.name] = ws;
  var back = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellDates: true });
  return T.gridFromSheet(back.Sheets[sheet.name]);
}
var S = T.sample(), tg = T.defaultTargets();
function table(fi, si) { var g = roundtrip(S[fi].sheets[si]); return T.extractTable(g.grid, { merges: g.merges }); }
function out(fi, si, map) { var t = table(fi, si); return T.buildOutput(t, map || T.suggest(t, tg), tg, { file: S[fi].name, sheet: S[fi].sheets[si].name }); }

/* ---- 값 모양 ---- */
eq(T.toNumber('1,250'), 1250, '쉼표 숫자');
eq(T.toNumber('7,800원'), 7800, '원 붙은 숫자');
eq(T.toNumber('(1,200)'), -1200, '회계식 음수');
eq(T.toNumber('협의'), null, '글자는 숫자가 아님');
eq(T.fmtDate(T.toDate('2026.12.05')), '2026-12-05', '점 날짜');
eq(T.fmtDate(T.toDate('26.11.20')), '2026-11-20', '두 자리 연도');
eq(T.fmtDate(T.toDate('2026년 11월 27일')), '2026-11-27', '한글 날짜');
eq(T.fmtDate(T.toDate(46361)), '2026-12-05', '엑셀 일련번호');
eq(T.toDate('2026-02-30'), null, '없는 날짜는 거른다');
eq(T.kindOf('HB-2201-A'), 'code', '영문+숫자 = 코드');
eq(T.kindOf('유압 호스'), 'text', '글자');

/* ---- 이름 유사도 ---- */
eq(T.norm('단가(원)'), '단가', '괄호 안 단위는 이름 비교에서 뺀다');
eq(T.norm("Q'ty"), 'qty', '구두점 제거');
eq(T.nameScore('Part No.', '품번').score, 0.92, '같은 뜻 낱말 표: Part No. = 품번');
ok(T.nameScore('자재명/사양', '품명').score >= 0.7, '이름 일부가 겹치면 0.7 이상');
eq(T.nameScore('No', '비고').score < 0.45, true, '짧은 이름(No)이 note 안에 들었다고 비고로 잡지 않는다');

/* ---- 병합 채우기 ---- */
var f = T.fillMerges([['A', ''], ['', '']], [{ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }]);
eq(f.grid, [['A', 'A'], ['A', 'A']], '병합 범위 전체에 왼쪽 위 값을 채운다');
eq(Object.keys(f.filled).length, 3, '채운 칸 3개를 기록한다(원래 칸은 제외)');

/* ---- 가람정밀 견적서: 제목·정보 줄·두 줄 머리글·합계 줄 ---- */
var A = table(0, 0);
eq(A.header, { start: 3, count: 2, found: true, manual: false }, '머리글은 4~5행 두 줄(정보 줄을 머리글로 오인하지 않음)');
eq(A.title, '견 적 서', '제목');
eq(A.meta, [{ key: '업체명', value: '가람정밀(주)' }, { key: '견적일', value: '2026-10-28' }], '머리글 위 「항목 : 값」');
eq(A.columns.map(function (c) { return c.name; }), ['No', '품번', '품명', '규격 재질', '규격 치수', '수량', '단가(원)', '금액(원)'], '가로 병합 머리글은 「규격 재질」처럼 합친다');
eq(A.rows.length, 5, '데이터 5줄');
eq(A.skipped.map(function (s) { return s.reason; }), ['합계 줄', '값이 한 칸뿐인 줄(제목·메모로 보임)'], '합계·메모 줄은 버리고 사유를 남긴다');
var mA = T.suggest(A, tg);
eq(mA['공급업체'].src, 'm:업체명', '공급업체는 문서 정보 「업체명」에서');
eq(mA['단가'].src, 'c:단가(원)', '단가(원) → 단가');
eq(mA['비고'].src, '', '맞는 열이 없으면 비운다');
var oA = T.buildOutput(A, mA, tg, { file: S[0].name, sheet: '견적서' });
eq(oA.rows[0].marks['공급업체'], 'meta', '문서 정보에서 온 값은 meta 표시');
eq(oA.issues.map(function (i) { return i.row + ':' + i.target; }), ['9:금액'], '금액 ≠ 수량×단가 인 줄(9행)을 짚는다');

/* ---- 한빛유압: 세로 병합(Vendor) · 영문 머리글 ---- */
var oB = out(1, 0);
eq(oB.rows.map(function (r) { return r.values['공급업체']; }), ['한빛유압', '한빛유압', '한빛유압', '동해실링', '동해실링'], '세로 병합된 공급업체를 모든 줄에 채운다');
eq(oB.rows.map(function (r) { return r.marks['공급업체']; }), ['doc', 'merge', 'merge', 'doc', 'merge'], '채운 값은 merge, 원래 칸은 doc');
eq(oB.rows[0].values['금액'], 7700000, '금액 열이 없으면 수량×단가로 계산');
eq(oB.rows[0].marks['금액'], 'calc', '계산한 값은 calc 표시');
var B2 = table(1, 1);
eq(B2.header.count, 2, '단위 줄 (EA)(원) 은 머리글에 붙인다');
eq(B2.columns[2].name, "Q'ty (EA)", '단위가 이름 뒤에 붙는다');
eq(T.suggest(B2, tg)['수량'].src, "c:Q'ty (EA)", '단위가 붙어도 수량으로 짝짓는다');

/* ---- 새솔산업: 글자로 적힌 숫자·날짜, 빈 필수 값, 읽을 수 없는 값 ---- */
var oC = out(2, 0);
eq(oC.rows[0].values['수량'], 1250, '「1,250」 → 1250');
eq(oC.rows[0].marks['수량'], 'conv', '형식을 바꾼 값은 conv 표시');
eq(oC.rows.map(function (r) { return r.values['납기']; }), ['2026-12-05', '2026-12-05', '2026-11-20', '2026-11-27'], '날짜 표기 세 가지를 한 모양으로');
eq(oC.issues.map(function (i) { return i.row + ':' + i.target; }), ['4:수량', '5:단가'], '빈 필수 수량 · 「협의」 단가를 짚는다');
eq(oC.rows[3].values['단가'], '협의', '읽지 못한 값은 지우지 않고 원래 글자로 둔다');

/* ---- 다섯 시트 모두 핵심 열 짝짓기 ---- */
var all = [[0, 0], [0, 1], [1, 0], [1, 1], [2, 0]].map(function (p) { return T.suggest(table(p[0], p[1]), tg); });
ok(all.every(function (m) { return m['품번'].src && m['품명'].src && m['수량'].src; }), '예제 5개 시트 모두 필수 열(품번·품명·수량)을 찾는다');

/* ---- 머리글을 사람이 정하면 그대로 ---- */
var g0 = roundtrip(S[0].sheets[0]);
var man = T.extractTable(g0.grid, { merges: g0.merges, header: { start: 3, count: 1 } });
eq(man.header.manual, true, '사람이 정한 머리글 표시');
eq([man.rows.length, man.skipped[0].reason], [5, '머리글 반복'], '머리글을 한 줄로 정하면 둘째 머리글 줄은 「머리글 반복」으로 버린다');

/* ---- 목표 양식 엑셀 ---- */
eq(T.targetsFromGrid([['품번 *', '수량', '입고일'], ['코드', '숫자', '날짜']]),
  [{ name: '품번', type: 'code', required: true }, { name: '수량', type: 'number', required: false }, { name: '입고일', type: 'date', required: false }], '양식: 이름·필수(*)·형식 줄');
eq(T.targetsFromGrid([['품명', '단가'], ['볼트', 120]]).map(function (t) { return t.type; }), ['text', 'number'], '형식 줄이 없으면 예시 값 모양으로');

/* ---- 규칙 저장·재사용 ---- */
var B1 = table(1, 0), mB = T.suggest(B1, tg);
mB['비고'] = { src: 'c:Lead Time', by: 'user' };              // 사람이 고친 짝
var rule = T.makeRule('한빛유압 목록', B1, mB, tg);
var hit = T.matchRule([rule], B2);
ok(hit && hit.rule === rule, '같은 업체의 다른 시트(열 이름 같음)에 규칙이 맞는다');
eq(T.applyRule(rule, B2, tg)['비고'].src, 'c:Lead Time', '사람이 고친 짝이 규칙으로 다시 쓰인다');
eq(T.matchRule([rule], table(2, 0)), null, '양식이 다른 파일에는 규칙을 쓰지 않는다');
eq(T.validRules(JSON.stringify({ rules: [rule, { name: 'x' }] })).length, 1, '불러온 JSON 에서 모양이 맞는 규칙만 받는다');

/* ---- PDF 좌표 → 표 ---- */
function it(s, x, y, w) { return { str: s, transform: [10, 0, 0, 10, x, y], width: w }; }
var items = [
  it('부품 단가표', 200, 780, 60),
  it('품번', 50, 740, 20), it('품명', 150, 740, 20), it('수량', 300, 740, 20), it('단가', 400, 740, 20),
  it('AX-100', 50, 720, 34), it('기어', 150, 720, 20), it('펌프', 172, 720, 20), it('12', 305, 720, 10), it('15,000', 392, 720, 30),
  it('AX-200', 50, 701, 34), it('실린더', 150, 701, 30), it('3', 310, 701.5, 5), it('128,000', 390, 701, 35)
];
var pg = T.pdfItemsToGrid(items);
eq(pg[1], ['품번', '품명', '수량', '단가'], 'PDF 머리글 줄');
eq(pg[2], ['AX-100', '기어 펌프', '12', '15,000'], '가까운 글자는 한 칸으로 붙이고, 오른쪽 정렬 숫자도 제 열에');
eq(pg[3][2], '3', 'y 가 조금 어긋나도 같은 줄');
var pt = T.extractTable(pg, {});
eq([pt.title, pt.rows.length, T.suggest(pt, tg)['단가'].src], ['부품 단가표', 2, 'c:단가'], 'PDF 표도 같은 짝짓기로');

/* ---- 내보내기 ---- */
var sh = T.exportSheets([oB], tg);
eq(Object.keys(sh), ['정리결과', '값출처', '확인할것', '짝짓기'], '엑셀 시트 4장');
eq(sh['정리결과'].length, 6, '머리글 + 5줄');
eq(sh['값출처'][2][7], '병합 채움', '값출처 시트에 출처 이름');
eq(T.toCSV([['a', 'b,c'], ['따옴표"', '줄\n바꿈']]), 'a,"b,c"\r\n"따옴표""","줄\n바꿈"', 'CSV 따옴표 처리');
eq(T.csvToGrid('품번,수량\r\n"A,1",3\n'), [['품번', '수량'], ['A,1', '3']], 'CSV 읽기');

/* ---- AI 답 검사 ---- */
var ar = T.parseAiReply('설명\n```json\n{"품번": {"src": "c:Part No.", "why": "번호"}, "규격": {"src": "c:없는열"}}\n```', B1, tg);
eq(ar.map['품번'].src, 'c:Part No.', 'AI 답의 짝을 읽는다');
eq([ar.map['규격'].src, ar.dropped.length], ['', 1], '표에 없는 열을 말하면 버린다');
eq(ar.map['품번'].by, 'ai', 'AI 제안은 ai 로 표시(검토 필요)');
ok(T.aiPrompt(B1, tg).user.indexOf('3850000') >= 0 && T.aiPrompt(B1, tg).user.indexOf('BSK-12') < 0, 'AI 에는 열마다 예시 3개만 보낸다');


/* ================= v1.1 폴더 분류·추출 (js/classify.js) ================= */
var C = require('../js/classify.js'), fs = require('fs'), path = require('path'), cp = require('child_process'), os = require('os');
var ROOT = path.join(__dirname, '..');
var v10 = passed + failed;

/* ---- 규칙 파일 ---- */
var fileRules = JSON.parse(fs.readFileSync(path.join(ROOT, 'rules/doc-types.json'), 'utf8'));
eq(fileRules, C.DEFAULT_RULES, 'rules/doc-types.json 과 화면 기본 규칙이 같다(웹·스크립트가 같은 규칙)');
ok(C.checkRules(fileRules).ok, '기본 규칙 파일은 검사를 통과한다');
eq(JSON.parse(C.rulesToJSON(C.DEFAULT_RULES)), C.DEFAULT_RULES, '규칙 JSON 내려받기 → 올리기 왕복이 같다');
var bad = C.defaultRules();
bad.types[1].id = 'quote';
bad.types[0].conditions.push({ field: '없는항목', op: '>=', value: 1 });
bad.types[2].conditions[0].op = 'like';
bad.types[2].match.fileName.push({ re: '([', weight: 1 });
var bc = C.checkRules(bad);
eq([bc.ok, bc.errors.length], [false, 4], '틀린 곳(겹친 id·없는 항목·모르는 비교·틀린 정규식)을 모두 알려 준다');
ok(/겹칩니다/.test(bc.errors[1] + bc.errors[0] + bc.errors[2] + bc.errors[3]), '겹친 id 를 짚는다');
eq(C.checkRules('{ "types": [').ok, false, 'JSON 이 깨지면 받지 않는다');
var b2 = C.defaultRules(); b2.types[0].conditions = [{ field: '견적일', op: 'between', value: ['2026-10-01'] }];
ok(!C.checkRules(b2).ok, 'between 은 두 값이 필요하다');

/* ---- 예제 폴더를 실제 파일 모양으로 ---- */
var SF = C.sampleFolder();
function sampleDoc(f) { return C.docOf(f.path, f.sheets.map(function (s) { var g = roundtrip(s); return { name: s.name, grid: g.grid, merges: g.merges }; })); }
var docs = SF.filter(function (f) { return f.sheets; }).map(sampleDoc);
function byName(n) { return docs.filter(function (d) { return d.name.indexOf(n) >= 0; })[0]; }
var R = C.defaultRules();
eq(SF.length, 9, '예제 폴더 9개(xlsx 7 · pdf 1 · 메모 1)');
eq(SF.map(function (f) { return C.isSupported(f.path); }).filter(Boolean).length, 8, '메모.txt 는 읽는 대상이 아니다');
ok(!C.isSupported('a/~$견적.xlsx') && C.isSupported('a/b.PDF'), '엑셀 임시 파일(~$)은 건너뛰고 대문자 확장자는 받는다');
eq([C.folderOf('2026-10/견적/a.xlsx'), C.baseName('2026-10/견적/a.xlsx')], ['2026-10/견적', 'a.xlsx'], '하위 폴더 경로와 파일 이름을 나눈다');

/* ---- 분류 판정 ---- */
var cG = C.classify(byName('가람정밀'), R);
eq([cG.status, cG.type, cG.score], ['ok', 'quote', 11], '가람정밀 견적서 → 견적서 11점');
eq(cG.hits.map(function (h) { return h.kind + ':' + h.word; }), ['fileName:견적', 'header:단가', 'body:견적서', 'body:견적일', 'body:부가세별도'], '근거: 어떤 규칙이 맞았는지 남긴다');
ok(cG.hits.some(function (h) { return h.word === '견적서'; }), '본문 「견 적 서」(띄어 쓴 제목)도 「견적서」로 찾는다');
eq(C.classify(byName('quotation'), R).type, 'quote', '영문 견적서(quotation) → 견적서');
eq(C.classify(byName('한빛유압'), R).type, 'statement', '거래명세서 → 거래명세서');
eq([C.classify(byName('새솔'), R).type, C.classify(byName('새솔'), R).score], ['statement', 6], '이름에 「명세」만 있어도 머리글 「입고일」·본문 낱말로 거래명세서');
eq(C.classify(byName('동해실링'), R).type, 'inspection', '검사성적서 → 검사성적서');
var cM = C.classify(byName('회의록'), R);
eq([cM.status, cM.type], ['none', ''], '회의록 → 미분류(어느 종류에도 못 미침)');
var cX = C.classify(byName('통합'), R);
eq([cX.status, cX.score, cX.second.score], ['check', 7, 7], '견적·명세 통합 → 확인 필요(7점 대 7점)');
ok(/차이 0점/.test(cX.why), '확인 필요 근거에 점수 차이를 적는다');
var R2 = C.defaultRules(); R2.closeMargin = -1;
eq(C.classify(byName('통합'), R2).status, 'ok', 'closeMargin 을 줄이면 확인 필요가 풀린다(규칙 파일로 조정)');
var R5 = C.defaultRules(); R5.closeMargin = 0;
eq(C.classify(byName('통합'), R5).status, 'check', '차이가 기준과 같으면(0점 · 기준 0점) 확인 필요 — 경계 포함');
var R3 = C.defaultRules(); R3.types[1].minScore = 7;
eq(C.classify(byName('새솔'), R3).status, 'none', '최소 점수를 올리면 6점 문서는 미분류');
var R4 = C.defaultRules(); R4.types[0].match.fileName.push({ re: '_10\\d\\d\\.xlsx$', weight: 1 });
eq(C.scoreType(byName('가람정밀'), R4.types[0]).score, 12, '정규식 규칙도 점수에 더한다');
var once = C.docOf('견적.xlsx', [{ name: 's', grid: [['견적서 견적서 견적서', 'a'], ['품번', '품명']], merges: [] }]);
eq(C.scoreType(once, R.types[0]).score, 7, '낱말 하나는 여러 번 나와도 한 번만 센다');

/* ---- 뽑을 항목 짝짓기 (규칙 파일 aliases) ---- */
var tS = byName('새솔').tables[0].table;
eq(C.mapFields(tS, R.types[1].fields)['납품일'].src, 'c:입고일', 'aliases: 입고일 → 납품일');
var tQ = byName('quotation').tables[0].table, mQ = C.mapFields(tQ, R.types[0].fields);
eq([mQ['견적일'].src, mQ['공급업체'].src], ['m:Date', 'm:Vendor'], '머리글 위 「Date :」·「Vendor :」를 견적일·공급업체로');
var mI = C.mapFields(byName('동해실링').tables[0].table, R.types[2].fields);
eq(['검사항목', '측정값', '판정', '검사일'].map(function (k) { return mI[k].src; }), ['c:검사항목', 'c:측정값', 'c:판정', 'm:검사일'], '검사성적서 항목 짝짓기');

/* ---- 추출 조건 ---- */
eq(C.checkRow({ 금액: 300000 }, [{ field: '금액', op: '>=', value: 300000 }]).ok, true, '금액 ≥ 값: 경계값은 남긴다');
eq(C.checkRow({ 금액: 299999 }, [{ field: '금액', op: '>=', value: 300000 }]).reasons.length, 1, '금액 ≥ 값: 모자라면 뺀다');
eq(C.checkRow({ 금액: '1,250,000' }, [{ field: '금액', op: '>', value: 1000000 }]).ok, true, '글자로 적힌 숫자도 견준다');
eq(C.checkRow({ 금액: '협의' }, [{ field: '금액', op: '>=', value: 1 }]).reasons[0], '금액 「협의」을 견줄 수 없음', '견줄 수 없는 값은 사유를 남기고 뺀다');
var between = [{ field: '납품일', op: 'between', value: ['2026-10-01', '2026-10-31'] }];
eq(['2026-10-01', '2026-10-31', '2026-09-30', '2026-11-01'].map(function (d) { return C.checkRow({ 납품일: d }, between).ok; }), [true, true, false, false], '날짜 기간: 양 끝 포함');
eq(C.checkRow({ 판정: 'ng' }, [{ field: '판정', op: 'contains', value: ['NG', '불합격'] }]).ok, true, '낱말 포함(대소문자 무시, 목록 중 하나)');
eq(C.checkRow({ 판정: '합격' }, [{ field: '판정', op: 'contains', value: ['NG', '불합격'] }]).ok, false, '「합격」은 「불합격」 조건에 걸리지 않는다');
eq(C.checkRow({ 비고: '샘플' }, [{ field: '비고', op: 'notContains', value: '샘플' }]).ok, false, '낱말 없음 조건');
eq(C.checkRow({ 품번: '' }, [{ field: '품번', op: 'notEmpty' }]).ok, false, '값 있음 조건');
eq(C.checkRow({ 금액: 5, 품번: '' }, [{ field: '금액', op: '>', value: 1 }, { field: '품번', op: 'notEmpty' }]).reasons.length, 1, '조건은 모두 맞아야 남긴다(그리고)');
eq(C.condText(between[0]), '납품일 2026-10-01 ~ 2026-10-31', '조건을 사람이 읽는 글로');

/* ---- 문서 하나 추출 ---- */
function counts(n, ti) { var e = C.extractDoc(byName(n), R.types[ti]); return [e.kept.length, e.dropped.length]; }
eq(counts('가람정밀', 0), [4, 4], '가람정밀: 금액 ≥ 30만 4행 남김 · 4행 뺌(추가품목 시트 계산 금액 포함)');
eq(counts('quotation', 0), [2, 1], '누리전장 견적: 2행 남김 · 1행 뺌');
eq(counts('한빛유압', 1), [3, 1], '한빛유압 명세: 10월 납품 3행 · 11월 1행 뺌(합계 줄은 표에서 이미 버림)');
eq(counts('새솔', 1), [3, 1], '새솔산업 명세: 9월 입고 1행 뺌');
var eD = C.extractDoc(byName('동해실링'), R.types[2]);
eq(eD.kept.map(function (r) { return r.values['품번'] + ':' + r.values['판정']; }), ['ORK-30:불합격', 'BSK-12:NG'], '검사성적서: 불합격·NG 행만 뽑는다');
eq(eD.kept.map(function (r) { return r.marks['품번']; }), ['merge', 'merge'], '세로 병합된 품번은 「병합 채움」 표시(v1.0 값 출처 재사용)');
var eG = C.extractDoc(byName('가람정밀'), R.types[0]);
eq(eG.dropped.map(function (d) { return d.row.marks['금액']; }).filter(function (m) { return m === 'calc'; }).length, 3, '금액 열이 없는 시트는 수량×단가로 계산한 금액으로 조건을 본다');
var eM = C.extractDoc(byName('회의록'), R.types[0]);
eq([eM.kept.length, ((eM.skippedTables[0] || {}).reason || '').indexOf('필수 항목') === 0], [0, true], '필수 항목이 하나도 없는 표는 뽑지 않고 사유를 남긴다');

/* ---- 폴더 전체 · 사람이 고르기 ---- */
var RF = C.runFolder(docs, R);
eq(RF.groups.map(function (g) { return g.items.length; }), [2, 2, 1], '종류별 묶음(xlsx 만: 견적 2 · 명세 2 · 성적서 1)');
eq([RF.check.length, RF.none.length], [1, 1], '확인 필요 1 · 미분류 1 은 뽑지 않는다');
var asg = {}; asg[byName('통합').path] = 'quote'; asg[byName('가람정밀').path] = '';
var RA = C.runFolder(docs, R, asg);
eq([RA.groups[0].items.length, RA.groups[0].kept], [2, 4], '확인 필요를 견적서로 고르면 뽑고(2행), 사람이 뺀 파일은 빠진다');
eq(C.statusLabel(RA.none.filter(function (i) { return i.doc.name.indexOf('가람') === 0; })[0]), '사람이 뺌', '사람이 뺀 파일 표시');
eq(RA.check.length, 0, '사람이 고른 파일은 확인 필요에서 빠진다');

/* ---- 엑셀 ---- */
var bk = C.exportBook(RF, R, [{ path: 'x/메모.txt', reason: '엑셀·CSV·PDF 가 아님' }]);
eq(Object.keys(bk), ['분류표', '견적서', '거래명세서', '검사성적서', '뺀 행', '확인할것', '건너뛴 파일'], '엑셀: 분류표 + 종류별 시트 + 뺀 행 · 확인할것 · 건너뛴 파일');
eq([bk['분류표'].length, bk['견적서'].length, bk['뺀 행'].length], [8, 7, 12], '분류표 7파일 · 견적서 6행 · 뺀 행 11개(+머리글)');
eq(bk['분류표'][0], ['경로', '파일', '종류', '상태', '점수', '근거', '2순위', '뽑은 행', '뺀 행'], '분류표 칸: 파일·종류·점수·근거');
var used = {}; eq([C.sheetName('a/b:c', used), C.sheetName('a/b:c', used)], ['a b c', 'a b c 2'], '시트 이름: 못 쓰는 글자를 바꾸고 겹치면 번호');

/* ---- 이미 처리한 파일 ---- */
var pr = C.markDone(null, RF, (function () { var m = {}; docs.forEach(function (d) { m[d.path] = { size: 10, mtime: 5 }; }); return m; })(), 'now');
ok(C.alreadyDone(pr, docs[0].path, { size: 10, mtime: 5 }), '같은 경로·크기·시각이면 이미 처리함');
ok(!C.alreadyDone(pr, docs[0].path, { size: 10, mtime: 6 }), '고친 시각이 바뀌면 다시 처리');

/* ---- samples/ 의 실제 파일이 예제 정의와 같다 ---- */
var same = SF.filter(function (f) { return f.sheets; }).every(function (f) {
  var wb = XLSX.read(fs.readFileSync(path.join(ROOT, 'samples', f.path)), { type: 'buffer' });
  return f.sheets.every(function (s) { return JSON.stringify(T.gridFromSheet(wb.Sheets[s.name]).grid) === JSON.stringify(s.aoa); });
});
ok(same, 'samples/ 엑셀 7개가 classify.js 예제 정의와 같다(tools/make-samples.mjs)');
ok(Buffer.compare(fs.readFileSync(path.join(ROOT, 'samples/2026-11/품질/ECU_inspection_report.pdf')), Buffer.from(C.samplePdf())) === 0, 'samples/ PDF 가 예제 정의와 같다');

/* ---- 로컬 스크립트 tools/folder-run.mjs 가 samples 를 처리한다 (PDF 포함) ---- */
var tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'b261102-'));
function run(extra) { return cp.execFileSync(process.execPath, [path.join(ROOT, 'tools/folder-run.mjs'), path.join(ROOT, 'samples')].concat(extra), { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }); }
var sum = JSON.parse(run(['--out', path.join(tmp, 'out.xlsx'), '--processed', path.join(tmp, 'processed.json'), '--json']));
eq([sum.files, sum.skipped], [8, 1], '스크립트: 8개 분류 · 메모 1개 건너뜀');
eq(sum.groups.map(function (g) { return g.name + ' ' + g.files + '/' + g.kept + '/' + g.dropped; }), ['견적서 2/6/5', '거래명세서 2/6/2', '검사성적서 2/3/6'], '스크립트: 종류별 파일·뽑은 행·뺀 행');
eq([sum.check, sum.unclassified], [['2026-11/기타/견적_명세_통합.xlsx'], ['2026-11/기타/회의록_정리.xlsx']], '스크립트: 확인 필요·미분류 경로');
var pdfItem = sum.items.filter(function (i) { return /\.pdf$/.test(i.path); })[0];
eq([pdfItem.type, pdfItem.kept, pdfItem.dropped], ['inspection', 1, 2], '스크립트: PDF 검사성적서도 읽어 NG 1행을 뽑는다(node 에서 pdf.js)');
var owb = XLSX.read(fs.readFileSync(path.join(tmp, 'out.xlsx')), { type: 'buffer' });
eq(owb.SheetNames, ['분류표', '견적서', '거래명세서', '검사성적서', '뺀 행', '확인할것', '건너뛴 파일'], '스크립트 결과 엑셀 시트가 웹과 같다');
eq(XLSX.utils.sheet_to_json(owb.Sheets['검사성적서'], { header: 1 }).length, 4, '스크립트 결과: 검사성적서 3행 + 머리글');
var again = JSON.parse(run(['--out', path.join(tmp, 'out2.xlsx'), '--processed', path.join(tmp, 'processed.json'), '--json']));
eq([again.files, again.skipped], [0, 9], '--processed: 두 번째 실행은 이미 처리한 8개를 건너뛴다');
var txt = run(['--out', path.join(tmp, 'out3.xlsx')]);
ok(/확인 필요\t-\t7/.test(txt) && /건너뜀/.test(txt), '스크립트 콘솔 분류표(상태·종류·점수·근거)');
var badR = path.join(tmp, 'bad.json'); fs.writeFileSync(badR, JSON.stringify(bad));
var code = 0; try { run(['--rules', badR, '--out', path.join(tmp, 'o4.xlsx')]); } catch (e) { code = e.status; }
eq(code, 2, '틀린 규칙 파일이면 스크립트가 멈춘다(결과를 만들지 않음)');
fs.rmSync(tmp, { recursive: true, force: true });

console.log('\nv1.0 규칙 ' + v10 + '개 · v1.1 폴더 분류 ' + (passed + failed - v10) + '개');
console.log('\n규칙 통과 ' + passed + ' · 실패 ' + failed);
process.exit(failed ? 1 : 0);
