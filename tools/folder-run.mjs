#!/usr/bin/env node
// 폴더 분류·추출 (로컬 스크립트) — 웹 도구(folder.html)와 같은 엔진·같은 규칙 파일을 쓴다.
//
//   node tools/folder-run.mjs <폴더> [--rules rules/doc-types.json] [--out 결과.xlsx] [--processed processed.json] [--json]
//
//   --rules      분류 규칙 파일(기본: 리포의 rules/doc-types.json)
//   --out        결과 엑셀(기본: 지금 폴더의 폴더분류_결과_<날짜>.xlsx). 읽는 폴더 안에는 쓰지 않는다.
//   --processed  이미 처리한 파일 기록. 주면 같은 파일(경로·크기·고친 시각)은 건너뛰고, 처리한 파일을 덧붙여 저장한다.
//   --json       분류표 대신 요약 JSON 을 출력(다른 도구·Codex 가 읽기 쉽게)
//
// 엑셀·CSV·PDF(글자가 든 PDF) 를 읽는다. 스캔 그림 PDF 는 글자가 없어 미분류가 된다.
import { readdirSync, statSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join, relative, resolve, sep } from 'node:path';
import { XLSX, Classify, ROOT, readSheets } from './read-node.mjs';

function arg(name, def) { const i = process.argv.indexOf(name); return i > 0 && process.argv[i + 1] ? process.argv[i + 1] : def; }
const pos = process.argv.slice(2).filter((a, i, all) => !a.startsWith('--') && !(i > 0 && /^--(rules|out|processed)$/.test(all[i - 1])));
if (!pos[0]) { console.error('사용: node tools/folder-run.mjs <폴더> [--rules 규칙.json] [--out 결과.xlsx] [--processed processed.json] [--json]'); process.exit(2); }
const folder = resolve(pos[0]);
const today = new Date().toISOString().slice(0, 10);
const rulesPath = resolve(arg('--rules', join(ROOT, 'rules/doc-types.json')));
const outPath = resolve(arg('--out', '폴더분류_결과_' + today + '.xlsx'));
const procPath = arg('--processed', '');
const asJson = process.argv.includes('--json');

const chk = Classify.checkRules(readFileSync(rulesPath, 'utf8'));
if (!chk.ok) { console.error('규칙 파일에 틀린 곳이 있습니다 (' + rulesPath + '):\n- ' + chk.errors.join('\n- ')); process.exit(2); }
const rules = chk.rules;

function walk(d, acc) {
  for (const n of readdirSync(d).sort()) {
    if (n.startsWith('.')) continue;
    const p = join(d, n), st = statSync(p);
    if (st.isDirectory()) walk(p, acc); else acc.push({ abs: p, rel: relative(folder, p).split(sep).join('/'), st });
  }
  return acc;
}
const all = walk(folder, []);
const outAbs = resolve(outPath);
let processed = procPath && existsSync(procPath) ? JSON.parse(readFileSync(procPath, 'utf8')) : null;
const docs = [], skipped = [], stamps = {};
for (const f of all) {
  if (f.abs === outAbs) continue;
  if (!Classify.isSupported(f.rel)) { skipped.push({ path: f.rel, reason: '엑셀·CSV·PDF 가 아님' }); continue; }
  const stamp = Classify.fileStamp(f.st);
  if (procPath && Classify.alreadyDone(processed, f.rel, stamp)) { skipped.push({ path: f.rel, reason: '이미 처리함(processed)' }); continue; }
  try { docs.push(Classify.docOf(f.rel, await readSheets(f.abs))); stamps[f.rel] = stamp; }
  catch (e) { skipped.push({ path: f.rel, reason: '읽지 못함: ' + e.message }); }
}
const res = Classify.runFolder(docs, rules);
const book = Classify.exportBook(res, rules, skipped);
const wb = XLSX.utils.book_new();
for (const k of Object.keys(book)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(book[k]), k);
if (outAbs.startsWith(folder + sep)) { console.error('결과 파일을 읽는 폴더 안에 쓰지 않습니다 — --out 을 폴더 밖으로 주세요.'); process.exit(2); }
writeFileSync(outAbs, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
if (procPath) { processed = Classify.markDone(processed, res, stamps, new Date().toISOString()); writeFileSync(procPath, JSON.stringify(processed, null, 2) + '\n'); }

const summary = {
  folder, rules: rulesPath, out: outAbs, files: docs.length, skipped: skipped.length,
  groups: res.groups.map(g => ({ id: g.id, name: g.name, files: g.items.length, kept: g.kept, dropped: g.dropped })),
  check: res.check.map(i => i.doc.path), unclassified: res.none.map(i => i.doc.path),
  items: res.items.map(i => ({ path: i.doc.path, type: i.type, status: Classify.statusLabel(i), score: i.cls.score, why: i.cls.why, kept: i.ext ? i.ext.kept.length : 0, dropped: i.ext ? i.ext.dropped.length : 0 }))
};
if (asJson) { console.log(JSON.stringify(summary, null, 2)); process.exit(0); }

const nameOf = id => (Classify.findType(rules, id) || {}).name || '-';
console.log('폴더 분류 — ' + folder);
console.log('규칙 ' + rulesPath + ' · 문서 종류 ' + rules.types.length + '개\n');
console.log(['상태', '종류', '점수', '뽑음/뺌', '경로', '근거'].join('\t'));
for (const i of summary.items) console.log([i.status, nameOf(i.type), i.score, i.type ? i.kept + '/' + i.dropped : '-', i.path, i.why].join('\t'));
for (const s of skipped) console.log(['건너뜀', '-', '-', '-', s.path, s.reason].join('\t'));
console.log('\n종류별: ' + summary.groups.map(g => g.name + ' ' + g.files + '개(' + g.kept + '행)').join(' · ') +
  ' · 확인 필요 ' + summary.check.length + ' · 미분류 ' + summary.unclassified.length + ' · 건너뜀 ' + skipped.length);
if (summary.check.length || summary.unclassified.length) console.log('확인 필요·미분류 파일은 뽑지 않았습니다. 웹 도구(folder.html)에서 종류를 고르거나 규칙 파일을 고친 뒤 다시 돌리세요.');
console.log('결과: ' + outAbs);
