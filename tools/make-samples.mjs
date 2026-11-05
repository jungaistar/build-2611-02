// 예제 폴더 만들기 — node tools/make-samples.mjs [폴더=samples]
// js/classify.js 의 sampleFolder() 정의로 실제 xlsx · pdf · txt 파일을 쓴다. 가짜 데이터입니다(실제 사내 데이터가 아닙니다).
import { mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { XLSX, Classify, ROOT } from './read-node.mjs';

const dir = process.argv[2] || join(ROOT, 'samples');
for (const f of Classify.sampleFolder()) {
  const p = join(dir, f.path);
  mkdirSync(dirname(p), { recursive: true });
  if (f.pdf) writeFileSync(p, Classify.samplePdf());
  else if (f.text != null) writeFileSync(p, f.text);
  else {
    const wb = XLSX.utils.book_new();
    for (const s of f.sheets) { const ws = XLSX.utils.aoa_to_sheet(s.aoa); ws['!merges'] = s.merges; XLSX.utils.book_append_sheet(wb, ws, s.name); }
    // 고친 시각이 아니라 내용으로 비교되도록 만든 날짜를 고정한다(다시 만들어도 같은 파일)
    wb.Props = { Title: '예제(가짜 데이터)', CreatedDate: new Date(Date.UTC(2026, 10, 3)) };
    writeFileSync(p, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true }));
  }
  console.log('썼습니다:', f.path);
}
