/**
 * G·M코드 퀴즈 기록·판정 스크립트 (구글 시트 하나로 끝)
 *
 * 붙이는 곳: 구글 시트 > 확장 프로그램 > Apps Script, 또는 Apps Script 사이트의 새 프로젝트
 *          (Apps Script 사이트에서 만들면 설정하기가 'GM코드 퀴즈 기록' 시트를 새로 만들고, 실행 로그에 그 주소를 보여 준다)
 * 처음 한 번: 위 함수 목록에서 '설정하기'를 골라 실행(권한 허용) → 배포 > 새 배포 > 웹 앱
 *            (실행: 나, 액세스 권한: 모든 사용자) → 나온 웹 앱 주소를 퀴즈에 넣는다.
 *
 * 하는 일
 *   1) 퀴즈가 제출마다 보내는 기록(이름, 점수, 브라우저)을 '응시 기록' 탭에 한 줄씩 쌓는다.
 *      시각과 합격 여부는 이 스크립트가 직접 정한다(학생이 조작할 수 없음).
 *   2) 평일 오전 9시마다 '전 평일 9시 ~ 오늘 9시' 기록만 보고 이름별 최종 결과
 *      (합격 / 불합격 / 미응시)를 판정한다. 그 기간에 한 번이라도 90점 이상이면 합격.
 *
 * 탭
 *   응시 기록   : 자동으로 쌓임(고치지 않는다)
 *   명단        : 선생님이 A열에 학생 이름을 한 줄에 한 명씩 붙여 넣는다
 *   최근 판정   : 가장 최근 판정의 이름별 결과, 최고 점수, 응시 횟수
 *   판정 기록   : 이름별로 날짜마다 결과를 한 칸씩 쌓은 표
 *   명단 외 이름: 명단에 없는 이름으로 응시한 기록(오타, 장난 이름 확인용)
 */

const PASS_SCORE = 90;
const RECORD_SHEET = '응시 기록';
const ROSTER_SHEET = '명단';
const TODAY_SHEET = '최근 판정';
const LOG_SHEET = '판정 기록';
const OTHER_SHEET = '명단 외 이름';
const DAYS = ['일', '월', '화', '수', '목', '금', '토'];
const COLORS = { '합격': '#d9f2e3', '불합격': '#fde2e1', '미응시': '#eeeeee' };

/* ---------- 처음 설정 ---------- */

// 기록할 시트: 설정하기 때 기억해 둔 시트. 시트에 붙인 스크립트면 그 시트,
// Apps Script 사이트에서 따로 만든 스크립트면 'GM코드 퀴즈 기록' 시트를 새로 만든다.
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  throw new Error('먼저 편집기에서 설정하기를 실행하세요.');
}

// 알림: 시트를 열어 둔 상태가 아니면 띄울 수 없으므로 실행 로그에만 남긴다
function toast_(ss, msg) {
  Logger.log(msg);
  try { ss.toast(msg); } catch (e) {}
}

function 설정하기() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : (SpreadsheetApp.getActive() || SpreadsheetApp.create('GM코드 퀴즈 기록'));
  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('기록 시트 주소: ' + ss.getUrl());
  recordSheet_(ss);
  if (!ss.getSheetByName(ROSTER_SHEET)) ss.insertSheet(ROSTER_SHEET).getRange(1, 1).setValue('이름').setFontWeight('bold');
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '판정하기')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('판정하기').timeBased().everyDays(1).atHour(9).create(); // 매일 9~10시 사이 실행, 주말은 건너뜀
  toast_(ss, '설정 완료: 명단 탭에 이름을 넣고, 웹 앱으로 배포하세요.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('GM코드 퀴즈')
    .addItem('지금 판정하기', '판정하기')
    .addItem('처음 설정하기', '설정하기')
    .addToUi();
}

function recordSheet_(ss) {
  let sh = ss.getSheetByName(RECORD_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RECORD_SHEET, 0);
    sh.getRange(1, 1, 1, 5).setValues([['시각', '이름', '점수', '결과', '브라우저']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- 1) 기록 받기 (퀴즈가 POST로 보냄) ---------- */

function doPost(e) {
  const out = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 20);
    const score = Number(d.score);
    const browser = String(d.browser || '').slice(0, 60);
    if (!name || !Number.isInteger(score) || score < 0 || score > 100) return out({ ok: false, error: 'bad data' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      recordSheet_(ss_())
        .appendRow([new Date(), "'" + name, score, score >= PASS_SCORE ? '합격' : '불합격', browser]);
    } finally {
      lock.releaseLock();
    }
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() {
  return ContentService.createTextOutput('GM코드 퀴즈 기록 서버가 동작 중입니다.');
}

/* ---------- 2) 평일 9시 판정 ---------- */

// 이름 비교는 띄어쓰기를 무시한다(홍 길동 = 홍길동).
function norm_(s) { return String(s == null ? '' : s).replace(/^'/, '').replace(/\s+/g, ''); }
// =, +, -, @로 시작하는 이름이 수식으로 읽히지 않게 한다
function safe_(s) { s = String(s); return /^[=+@-]/.test(s) ? "'" + s : s; }
function isWeekday_(d) { const w = d.getDay(); return w >= 1 && w <= 5; }
function at9_(d) { const x = new Date(d); x.setHours(9, 0, 0, 0); return x; }

// 지금 시각 이전의 가장 최근 평일 오전 9시
function latestCutoff_(now) {
  const d = at9_(now);
  if (d > now) d.setDate(d.getDate() - 1);
  while (!isWeekday_(d)) d.setDate(d.getDate() - 1);
  return d;
}
// 그 직전 평일 오전 9시(월요일 판정이면 금요일 9시)
function prevCutoff_(cutoff) {
  const d = new Date(cutoff);
  do { d.setDate(d.getDate() - 1); } while (!isWeekday_(d));
  return d;
}
function label_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d') + '(' + DAYS[d.getDay()] + ')';
}

function readRoster_(ss) {
  let sh = ss.getSheetByName(ROSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ROSTER_SHEET);
    sh.getRange(1, 1).setValue('이름');
  }
  const names = [];
  const seen = {};
  if (sh.getLastRow() >= 1) {
    sh.getRange(1, 1, sh.getLastRow(), 1).getValues().forEach(r => {
      const n = String(r[0]).trim();
      if (!n || n === '이름' || seen[norm_(n)]) return;
      seen[norm_(n)] = true;
      names.push(n);
    });
  }
  return names;
}

function 판정하기(e) {
  const now = new Date();
  if (e && e.triggerUid && !isWeekday_(now)) return; // 자동 실행은 주말에 건너뜀

  const ss = ss_();
  const sh = recordSheet_(ss);
  const cutoff = latestCutoff_(now);
  const start = prevCutoff_(cutoff);

  // 기간 안의 기록을 이름별로 모은다
  const stats = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
      const t = r[0];
      if (!(t instanceof Date) || t < start || t >= cutoff) return;
      const key = norm_(r[1]);
      if (!key) return;
      const score = Number(r[2]) || 0;
      const passed = String(r[3]) === '합격' || score >= PASS_SCORE;
      const s = stats[key] || (stats[key] = { name: String(r[1]).replace(/^'/, '').trim(), tries: 0, best: 0, passed: false });
      s.tries++;
      s.best = Math.max(s.best, score);
      s.passed = s.passed || passed;
    });
  }

  let roster = readRoster_(ss);
  const rosterKeys = {};
  roster.forEach(n => { rosterKeys[norm_(n)] = true; });
  const others = Object.keys(stats).filter(k => !rosterKeys[k]).map(k => stats[k]);
  if (roster.length === 0) roster = others.map(s => s.name); // 명단이 비었으면 응시자만 판정

  const result = n => {
    const s = stats[norm_(n)];
    if (!s) return { r: '미응시', best: '', tries: 0 };
    return { r: s.passed ? '합격' : '불합격', best: s.best, tries: s.tries };
  };
  const period = '판정 기간: ' + label_(start) + ' 9:00 ~ ' + label_(cutoff) + ' 9:00';

  // 최근 판정
  const today = ss.getSheetByName(TODAY_SHEET) || ss.insertSheet(TODAY_SHEET);
  today.clear();
  const rows = roster.map(n => { const x = result(n); return [safe_(n), x.r, x.best, x.tries]; });
  const count = r => rows.filter(x => x[1] === r).length;
  today.getRange(1, 1).setValue(period).setFontWeight('bold');
  today.getRange(2, 1).setValue('합격 ' + count('합격') + '명 · 불합격 ' + count('불합격') + '명 · 미응시 ' + count('미응시') + '명');
  today.getRange(4, 1, 1, 4).setValues([['이름', '결과', '최고 점수', '응시 횟수']]).setFontWeight('bold');
  if (rows.length) {
    today.getRange(5, 1, rows.length, 4).setValues(rows);
    today.getRange(5, 2, rows.length, 1).setBackgrounds(rows.map(x => [COLORS[x[1]]]));
  }

  // 판정 기록: 날짜 열을 하나씩 쌓는다(같은 날짜는 덮어씀)
  const log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  if (log.getLastRow() < 1) log.getRange(1, 1).setValue('이름').setFontWeight('bold');
  const lastCol = Math.max(1, log.getLastColumn());
  const heads = log.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  let col = heads.indexOf(label_(cutoff)) + 1;
  if (col === 0) {
    col = lastCol + 1;
    log.getRange(1, col).setValue(label_(cutoff)).setFontWeight('bold');
  }
  const logNames = log.getLastRow() >= 2 ? log.getRange(2, 1, log.getLastRow() - 1, 1).getValues().map(r => norm_(r[0])) : [];
  roster.forEach(n => {
    let i = logNames.indexOf(norm_(n));
    if (i < 0) {
      logNames.push(norm_(n));
      i = logNames.length - 1;
      log.getRange(i + 2, 1).setValue(safe_(n));
    }
    const r = result(n).r;
    log.getRange(i + 2, col).setValue(r).setBackground(COLORS[r]);
  });

  // 명단 외 이름
  const other = ss.getSheetByName(OTHER_SHEET) || ss.insertSheet(OTHER_SHEET);
  other.clear();
  other.getRange(1, 1).setValue(period + ' (명단에 없는 이름으로 응시한 기록)').setFontWeight('bold');
  other.getRange(3, 1, 1, 4).setValues([['적은 이름', '결과', '최고 점수', '응시 횟수']]).setFontWeight('bold');
  const orows = others.map(s => [safe_(s.name), s.passed ? '합격' : '불합격', s.best, s.tries]);
  if (orows.length) other.getRange(4, 1, orows.length, 4).setValues(orows);

  toast_(ss, period + ' 판정을 마쳤습니다.');
}
