# build-2611-02 · 엑셀·PDF 문서 정리

생성형 AI 업무자동화 전문가과정 **2차수** 수강생 프로젝트 02 — 기획 Wang Jeon.

**바로 쓰기: https://build-2611-02.jobability.co.kr/**

양식이 제각각인 엑셀·PDF 표를 올리면 정해 둔 최종 표 구조(목표 열)로 원본 열을 짝지어 제안합니다.
여러 시트·병합 셀·두 줄 머리글을 풀어 읽고, 사람이 확인·확정한 표만 엑셀·CSV 로 내보냅니다.
확정한 짝짓기는 규칙으로 저장해 같은 양식의 다음 파일에 그대로 씁니다.
파일은 브라우저 안에서만 처리하며 서버로 올라가지 않습니다. AI 연결은 선택입니다.

| 파일 | 역할 |
|---|---|
| `js/engine.js` | 판정 엔진(병합 채우기·머리글·PDF 좌표·짝짓기·변환·결과·규칙) — 판정은 전부 여기 |
| `js/app.js` | 화면(표 정리) |
| `folder.html` · `js/folder.js` | 화면(폴더 분류, v1.1) |
| `js/classify.js` | 폴더 분류·추출 판정(종류 점수·미분류·확인 필요·조건 추출·엑셀) — v1.1 |
| `rules/doc-types.json` | 분류 규칙 파일 — 웹 화면과 로컬 스크립트가 같이 씀 |
| `tools/folder-run.mjs` | 폴더 분류 로컬 스크립트(node, npm 설치 없음, PDF 포함) |
| `tools/read-node.mjs` · `tools/make-samples.mjs` | node 파일 읽기 · 예제 폴더 만들기 |
| `samples/` | 예제 폴더(가짜 업체 9개 파일) |
| `AGENTS.md` | Codex 작업 지시서 |
| `test/logic.test.js` | 판정 규칙 테스트 `node test/logic.test.js` |
| `test/smoke.browser.js` | 화면 연기 테스트(playwright, PDF 포함) |
| `lib/` | SheetJS · pdf.js(+cmaps) 동봉 |
| `CLAUDE.md` | 작업 지시서 + 기획서 원문 |

PDF 읽기는 pdf.js 모듈을 불러오므로 웹 주소(배포 사이트나 로컬 서버)로 열어야 합니다. 파일을 더블클릭해 열면 엑셀·CSV 만 됩니다.

## 폴더 분류·추출 (v1.1)

**https://build-2611-02.jobability.co.kr/folder.html** — 폴더를 통째로 올리면 안의 엑셀·PDF 를 규칙(파일 이름·머리글·본문 낱말, 가중치)으로 문서 종류별로 나누고, 종류마다 정한 항목 중 추출 조건에 맞는 행만 뽑아 엑셀(분류표 + 종류별 시트)로 저장합니다.
근거(어떤 규칙이 몇 점)를 함께 보이고, 어느 종류에도 못 미치면 「미분류」, 두 종류가 가까우면 「확인 필요」로 남겨 사람이 고릅니다.

Codex 는 웹페이지 안에서 돌지 않으므로, 같은 엔진·같은 규칙 파일을 쓰는 로컬 스크립트와 작업 지시서(`AGENTS.md`)를 함께 둡니다.

```sh
node tools/folder-run.mjs samples                       # 예제 폴더
node tools/folder-run.mjs <폴더> --rules rules/doc-types.json --out 결과.xlsx --processed processed.json
```

예제 데이터는 실제 사내 데이터가 아닙니다.

## 기획서 · 버전

- 사이트 「기획서」 페이지: https://build-2611-02.jobability.co.kr/plan.html — 워드 판 내려받기
- 워드 판: `docs/기획서/build-2611-02_WangJeon_엑셀PDF문서정리_기획서_v<버전>.docx` (버전마다 남김)
- 원본: `docs/plan.json` → 생성기 `build-2611-00/tools/plandoc/make_plan.py` 로 워드·그림·plan.html·CHANGELOG.md 를 다시 만든다
- 버전을 올릴 때: `plan.json` 의 `version` 을 바꾸고 `versions` 맨 앞에 바뀐 점을 적은 뒤 생성기를 다시 돌린다
