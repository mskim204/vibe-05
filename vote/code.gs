// 투표앱 - 구글 Apps Script 웹앱
// 시트 구성 (스프레드시트에 연결된 스크립트로 사용)
//   코드 탭: 1행 = 코드            / 2행부터 투표 코드(숫자 4자리)
//   투표 탭: 1행 = 시각 | 코드 | 고른 것 / 2행부터 투표 기록이 쌓여요

const CODE_SHEET = '코드';
const VOTE_SHEET = '투표';
const CHOICES = ['떡볶이', '치킨', '마라탕', '붕어빵'];

// 주소 끝에 ?page=result 를 붙이면 결과 화면
function doGet(e) {
  const p = (e && e.parameter) || {};
  const page = p.page || p.pape || '';
  const t = HtmlService.createTemplateFromFile('index');
  t.page = page === 'result' ? 'result' : 'vote';
  return t.evaluate()
    .setTitle(t.page === 'result' ? '투표 결과' : '투표하기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 화면에는 보기 목록만 보내요 (코드 목록은 절대 보내지 않음)
function getChoices() {
  return CHOICES;
}

// 투표: 결과 메시지만 돌려줘요
function vote(code, choice) {
  code = String(code || '').trim();
  if (CHOICES.indexOf(choice) === -1) {
    return { ok: false, message: '보기를 하나 골라 주세요' };
  }
  if (!/^\d{4}$/.test(code)) {
    return { ok: false, message: '없는 코드예요' };
  }

  // 여러 명이 동시에 눌러도 같은 코드가 두 번 저장되지 않게 잠금
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const codeSheet = ss.getSheetByName(CODE_SHEET);
    const voteSheet = ss.getSheetByName(VOTE_SHEET);

    if (columnValues_(codeSheet, 1).indexOf(code) === -1) {
      return { ok: false, message: '없는 코드예요' };
    }
    if (columnValues_(voteSheet, 2).indexOf(code) !== -1) {
      return { ok: false, message: '이미 투표했어요' };
    }

    // 앞의 ' 는 0으로 시작하는 코드(예: 0123)가 숫자로 바뀌지 않게 해줘요
    voteSheet.appendRow([new Date(), "'" + code, choice]);
    SpreadsheetApp.flush();
    return { ok: true, message: choice + '에 투표했어요!' };
  } finally {
    lock.releaseLock();
  }
}

// 보기별 표 수만 돌려줘요
function getResults() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET);
  const counts = {};
  CHOICES.forEach(function (c) { counts[c] = 0; });
  columnValues_(sheet, 3).forEach(function (c) {
    if (counts.hasOwnProperty(c)) counts[c]++;
  });
  return CHOICES.map(function (c) { return { name: c, count: counts[c] }; });
}

// 2행부터 한 열의 값을 화면에 보이는 글자 그대로 읽기
function columnValues_(sheet, col) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, col, last - 1, 1).getDisplayValues()
    .map(function (r) { return String(r[0]).trim(); });
}
