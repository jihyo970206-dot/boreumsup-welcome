/**
 * 보름숲 외국인 리뷰이벤트 현황 — 랜딩페이지 기록 수집용 Apps Script
 *
 * 처음 한 번: 위쪽 함수 선택 칸에서 setup 을 고르고 ▶실행 (권한 허용)
 * 그다음: 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자) → 웹 앱 URL 복사
 */

const LOG_SHEET = '로그';
const SUMMARY_SHEET = '요약';
const FILE_TITLE = '보름숲 외국인 리뷰이벤트 현황';
const TZ = 'Asia/Seoul';

const EVENTS = {
  visit: '접속',
  wine_redeem: '와인 제공 완료',
  review_click: '후기 버튼 클릭'
};
const PLATFORMS = { naver: '네이버', google: '구글', dianping: '따종디엔핑' };
const LANGS = { zh: '중국어', en: '영어' };

/** 랜딩페이지에서 보내는 기록을 받는 곳 */
function doPost(e) { return record_(e); }
function doGet(e) { return record_(e); }

function record_(e) {
  const p = (e && e.parameter) || {};
  const label = EVENTS[p.event];
  if (!label) return reply_('ignored');

  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET) || setup();
    const now = new Date();
    sheet.appendRow([
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(now, TZ, 'HH:mm:ss'),
      label,
      PLATFORMS[p.platform] || '',
      LANGS[p.lang] || clean_(p.lang, 10),
      clean_(p.loc, 20),
      clean_(p.sid, 40),
      clean_(p.device, 20)
    ]);
  } finally {
    lock.releaseLock();
  }
  return reply_('ok');
}

/** 시트 만들기 (처음 한 번 실행) */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename(FILE_TITLE);
  ss.setSpreadsheetTimeZone(TZ);

  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.getSheets()[0].getName() === 'Sheet1' || ss.getSheets()[0].getName() === '시트1'
      ? ss.getSheets()[0].setName(LOG_SHEET)
      : ss.insertSheet(LOG_SHEET);
  }
  log.getRange(1, 1, 1, 8).setValues([['날짜', '시간', '이벤트', '플랫폼', '언어', '위치', '방문 ID', '기기']])
    .setFontWeight('bold').setBackground('#1F3A33').setFontColor('#FFFFFF');
  log.setFrozenRows(1);
  log.setColumnWidths(1, 8, 110);

  let sum = ss.getSheetByName(SUMMARY_SHEET);
  if (!sum) sum = ss.insertSheet(SUMMARY_SHEET, 0);
  sum.clear();
  const L = "'" + LOG_SHEET + "'!";
  const rows = [
    ['보름숲 외국인 리뷰이벤트 현황', ''],
    ['', ''],
    ['전체', '값'],
    ['접속 수 (페이지 열림)', '=COUNTIF(' + L + 'C:C,"접속")'],
    ['방문자 수 (중복 제외)', '=IFERROR(COUNTUNIQUE(FILTER(' + L + 'G:G,' + L + 'C:C="접속")),0)'],
    ['와인 제공 완료', '=COUNTIF(' + L + 'C:C,"와인 제공 완료")'],
    ['후기 버튼 클릭 (전체)', '=COUNTIF(' + L + 'C:C,"후기 버튼 클릭")'],
    ['  · 따종디엔핑', '=COUNTIFS(' + L + 'C:C,"후기 버튼 클릭",' + L + 'D:D,"따종디엔핑")'],
    ['  · 구글', '=COUNTIFS(' + L + 'C:C,"후기 버튼 클릭",' + L + 'D:D,"구글")'],
    ['  · 네이버', '=COUNTIFS(' + L + 'C:C,"후기 버튼 클릭",' + L + 'D:D,"네이버")'],
    ['후기 버튼을 누른 방문자 수', '=IFERROR(COUNTUNIQUE(FILTER(' + L + 'G:G,' + L + 'C:C="후기 버튼 클릭")),0)'],
    ['', ''],
    ['전환율', ''],
    ['방문자 → 와인 제공', '=IFERROR(B6/B5,0)'],
    ['방문자 → 후기 버튼 클릭', '=IFERROR(B11/B5,0)'],
    ['', ''],
    ['일별 현황', '']
  ];
  sum.getRange(1, 1, rows.length, 2).setValues(rows);
  sum.getRange('A19').setFormula('=IFERROR(QUERY(' + L + 'A:C,"select A, count(C) where C is not null group by A pivot C label A \'날짜\'",1),"아직 기록이 없습니다")');
  sum.getRange('E3').setValue('언어별 현황').setFontWeight('bold');
  sum.getRange('E4').setFormula('=IFERROR(QUERY(' + L + 'C:E,"select E, count(C) where E is not null group by E pivot C label E \'언어\'",1),"아직 기록이 없습니다")');

  sum.getRange('A1').setFontSize(14).setFontWeight('bold');
  ['A3:B3', 'A13', 'A17'].forEach(function (r) { sum.getRange(r).setFontWeight('bold'); });
  sum.getRange('B14:B15').setNumberFormat('0.0%');
  sum.setColumnWidth(1, 220);
  sum.setColumnWidth(2, 90);
  return log;
}

function clean_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // 수식으로 해석되지 않게
  return s;
}

function reply_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT);
}
