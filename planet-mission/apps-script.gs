/**
 * 태양계 탐사대 — 결과 수집용 Google Apps Script
 * 사용법은 README.md를 참고하세요.
 * 수업 코드와 교사 코드는 아래에서 바꾸세요. (이 파일을 GitHub에 올릴 때는 실제 코드를 넣지 마세요.)
 */
const CLASS_CODE = '여기에-수업코드';
const TEACHER_CODE = '여기에-교사코드';
const SHEET_NAME = '탐사 현황';
const HEADERS = ['마지막 활동', '학번', '이름', '임무 1 진행', '임무 1 통과', '임무 1 도전 횟수', '임무 2 완료', '퀴즈 도전 횟수', '보너스 클리어',
  '행성 탐사 (한 번에 맞힌 수)', '분류 실험실 (한 번에 맞힌 수)', '퀴즈 최고 점수', '퀴즈 최고 정답 수', '행성별 결과', '분류 시도 횟수·표 실수'];
const KEYS = ['time', 'num', 'name', 'm1', 'm1pass', 'm1try', 'm2', 'gamePlays', 'bonus', 'explore', 'lab', 'gameBest', 'gameRight', 'planets', 'critTries'];

// 학생 기기가 진행 상황을 보내거나, 교사 화면이 결과를 불러올 때 실행돼요.
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'format' }); }
  if (d.action === 'list') {
    if (String(d.teacher || '').trim() !== TEACHER_CODE) return out({ ok: false, error: 'teacher' });
    return out({ ok: true, rows: listRows() });
  }
  if (String(d.code || '').trim() !== CLASS_CODE) return out({ ok: false, error: 'code' });
  if (d.action === 'check') return out({ ok: true });
  saveRow(d);
  return out({ ok: true });
}

// 웹 앱 주소를 브라우저 주소창에 넣으면 연결 확인 문구가 보여요.
function doGet() {
  return ContentService.createTextOutput('✅ 태양계 탐사대 결과 수집기가 작동 중이에요. 이 주소를 index.html의 SUBMIT_URL에 넣으세요.');
}

// 편집기에서 이 함수를 골라 ▶실행 하면 시트에 시험용 줄이 하나 생겨요. (권한 승인도 이때 해요)
function testWrite() {
  saveRow({ num: '00000', name: '시험', m1: '0/8', m1try: 1, explore: '0/0', lab: '0/0', gameRight: '0/15', planets: '시험 줄입니다. 지워도 돼요.' });
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

// 학생 한 명당 한 줄: 같은 학번+이름이 있으면 그 줄을 고쳐 쓰고, 없으면 새 줄을 만들어요.
function saveRow(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const num = String(d.num || '').slice(0, 6), name = String(d.name || '').slice(0, 12);
    const row = [
      new Date(), clean(num, 6), clean(name, 12), clean(d.m1, 5), d.m1pass ? 'O' : 'X', Number(d.m1try) || 1, d.m2 ? 'O' : 'X',
      Number(d.gamePlays) || 0, Number(d.bonus) || 0, clean(d.explore, 10), clean(d.lab, 10),
      Number(d.gameBest) || 0, clean(d.gameRight, 10), clean(d.planets, 200), clean(d.critTries, 120)
    ];
    const n = sh.getLastRow();
    let at = 0;
    if (n > 1) {
      const keys = sh.getRange(2, 2, n - 1, 2).getDisplayValues();
      const i = keys.findIndex(k => k[0] === num && k[1] === name);
      if (i >= 0) at = i + 2;
    }
    if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

function listRows() {
  const sh = sheet(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, KEYS.length).getDisplayValues()
    .map(r => Object.fromEntries(KEYS.map((k, i) => [k, r[i]])));
}

// 길이를 자르고, 시트에서 수식으로 실행되지 않도록 막습니다. (학번 앞의 0도 지켜져요)
function clean(v, max) {
  const s = String(v == null ? '' : v).slice(0, max);
  return /^[=+\-@0]/.test(s) ? "'" + s : s;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
