// node 에서 파일 하나를 읽어 시트·쪽마다 grid 로 — tools/folder-run.mjs 와 테스트가 함께 쓴다.
// 판정은 하지 않는다(판정은 js/classify.js). 라이브러리는 lib/ 에 동봉한 것만 쓴다(npm 의존성 없음).
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join } from 'node:path';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
export const XLSX = require(join(ROOT, 'lib/xlsx.full.min.js'));
export const Tidy = require(join(ROOT, 'js/engine.js'));
export const Classify = require(join(ROOT, 'js/classify.js'));
export { ROOT };

let pdfjsP = null;
function pdfjs() {
  if (!pdfjsP) pdfjsP = import(pathToFileURL(join(ROOT, 'lib/pdf.min.mjs')).href).then(m => {
    m.GlobalWorkerOptions.workerSrc = pathToFileURL(join(ROOT, 'lib/pdf.worker.min.mjs')).href;
    return m;
  });
  return pdfjsP;
}

// → [{ name, grid, merges }]
export async function readSheets(file) {
  if (/\.pdf$/i.test(file)) {
    const lib = await pdfjs();
    const doc = await lib.getDocument({ data: new Uint8Array(readFileSync(file)), cMapUrl: join(ROOT, 'lib/cmaps/'), cMapPacked: true, isEvalSupported: false, verbosity: 0 }).promise;
    const out = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const tc = await (await doc.getPage(p)).getTextContent();
      out.push({ name: p + '쪽', grid: Tidy.pdfItemsToGrid(tc.items), merges: [], noText: !tc.items.length });
    }
    await doc.destroy();
    return out;
  }
  if (/\.csv$/i.test(file)) return [{ name: 'CSV', grid: Tidy.csvToGrid(readFileSync(file, 'utf8')), merges: [] }];
  const wb = XLSX.read(readFileSync(file), { type: 'buffer', cellDates: true });
  return wb.SheetNames.map(s => { const g = Tidy.gridFromSheet(wb.Sheets[s]); return { name: s, grid: g.grid, merges: g.merges }; });
}
