/**
 * 태양계 탐사대 — 결과 수집용 Google Apps Script
 * 사용법은 README.md를 참고하세요.
 * 수업 코드는 아래 CLASS_CODE를 바꾸세요. (이 파일을 GitHub에 올릴 때는 실제 코드를 넣지 마세요.)
 */
const CLASS_CODE = '여기에-수업코드';
const SHEET_NAME = '탐사 결과';
const HEADERS = ['제출 시각', '학번', '이름', '행성 탐사 (한 번에 맞힌 수)', '분류 실험실 (한 번에 맞힌 수)', '퀴즈 최고 점수', '퀴즈 최고 정답 수', '퀴즈 도전 횟수', '행성별 결과', '분류 시도 횟수·표 실수'];

// 학생이 '제출'을 누르면 실행돼요.
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'format' }); }
  if (String(d.code || '').trim() !== CLASS_CODE) return out({ ok: false, error: 'code' });
  addRow(d);
  return out({ ok: true });
}

// 웹 앱 주소를 브라우저 주소창에 넣으면 연결 확인 문구가 보여요.
function doGet() {
  return ContentService.createTextOutput('✅ 태양계 탐사대 결과 수집기가 작동 중이에요. 이 주소를 index.html의 SUBMIT_URL에 넣으세요.');
}

// 편집기에서 이 함수를 골라 ▶실행 하면 시트에 시험용 줄이 하나 생겨요. (권한 승인도 이때 해요)
function testWrite() {
  addRow({ num: '00000', name: '시험', explore: '0/0', lab: '0/0', gameBest: 0, gameRight: '0/15', gamePlays: 0, planets: '시험 줄입니다. 지워도 돼요.', critTries: '' });
}

function addRow(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  sh.appendRow([
    new Date(), clean(d.num, 6), clean(d.name, 12), clean(d.explore, 10), clean(d.lab, 10),
    Number(d.gameBest) || 0, clean(d.gameRight, 10), Number(d.gamePlays) || 0,
    clean(d.planets, 200), clean(d.critTries, 120)
  ]);
}

// 길이를 자르고, 시트에서 수식으로 실행되지 않도록 막습니다. (학번 앞의 0도 지켜져요)
function clean(v, max) {
  const s = String(v == null ? '' : v).slice(0, max);
  return /^[=+\-@0]/.test(s) ? "'" + s : s;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
