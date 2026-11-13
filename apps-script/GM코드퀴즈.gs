/**
 * G·M코드 퀴즈 판정 스크립트 (구글 시트 하나, 기록을 쌓지 않는 방식)
 *
 * 붙이는 곳: Apps Script 프로젝트(또는 구글 시트 > 확장 프로그램 > Apps Script)
 * 처음 한 번: 함수 목록에서 '설정하기' 실행 → 웹 앱으로 배포(실행: 나, 액세스: 모든 사용자)
 *
 * '명단' 탭 하나로 본다.
 *   A 이름      : 선생님이 2행부터 학생 이름을 한 줄에 한 명씩 넣는다
 *   B 최종 결과 : 평일 오전 9시에 확정된 결과(합격/불합격). 아침에 이 칸만 보면 된다
 *   C 진행 중   : 지난 9시 이후 응시 결과. 응시할 때마다 바로 바뀐다
 *                 (한 번이라도 90점 이상이면 합격, 응시했지만 못 넘으면 불합격, 괄호 안은 그 기간 최고 점수)
 * 평일 9시가 지나면 C를 B로 옮기고 C를 비운다(B는 매번 덮어씀). 그 기간에 응시하지 않았으면 불합격 (미응시).
 * 주말에는 넘기지 않으므로 월요일 9시 결과는 '금요일 9시 ~ 월요일 9시' 기준이다.
 *
 * '명단 외 이름' 탭: 명단에 없는 이름(오타, 장난 이름)으로 응시한 경우를 같은 방식으로 보여 준다.
 *                    다음 기간에 다시 응시하지 않으면 9시에 지워진다.
 */

const PASS_SCORE = 90;
const ROSTER_SHEET = '명단';
const OTHER_SHEET = '명단 외 이름';
const OLD_SHEETS = ['응시 기록', '최근 판정', '판정 기록']; // 예전 방식 탭(설정하기 때 지움)
const DAYS = ['일', '월', '화', '수', '목', '금', '토'];
const COLORS = { '합격': '#d9f2e3', '불합격': '#fde2e1', '': null };

/* ---------- 공통 ---------- */

// 기록할 시트: 설정하기 때 기억해 둔 시트. 시트에 붙인 스크립트면 그 시트,
// Apps Script 사이트에서 따로 만든 스크립트면 'GM코드 퀴즈 기록' 시트를 새로 만든다.
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  throw new Error('먼저 편집기에서 설정하기를 실행하세요.');
}

// 알림: 시트를 열어 둔 상태가 아니면 띄울 수 없으므로 실행 로그에도 남긴다
function toast_(ss, msg) {
  Logger.log(msg);
  try { ss.toast(msg); } catch (e) {}
}

// 이름 비교는 띄어쓰기를 무시한다(홍 길동 = 홍길동).
function norm_(s) { return String(s == null ? '' : s).replace(/^'/, '').replace(/\s+/g, ''); }
// =, +, -, @로 시작하는 이름이 수식으로 읽히지 않게 한다
function safe_(s) { s = String(s); return /^[=+@-]/.test(s) ? "'" + s : s; }
function isWeekday_(d) { const w = d.getDay(); return w >= 1 && w <= 5; }

// 지금 시각 이전의 가장 최근 평일 오전 9시
function latestCutoff_(now) {
  const d = new Date(now);
  d.setHours(9, 0, 0, 0);
  if (d > now) d.setDate(d.getDate() - 1);
  while (!isWeekday_(d)) d.setDate(d.getDate() - 1);
  return d;
}
// 그 직전 평일 오전 9시(월요일이면 금요일 9시)
function prevCutoff_(cutoff) {
  const d = new Date(cutoff);
  do { d.setDate(d.getDate() - 1); } while (!isWeekday_(d));
  return d;
}
function label_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d') + '(' + DAYS[d.getDay()] + ') 9시';
}

function sheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 3).setValues([[name === ROSTER_SHEET ? '이름' : '적은 이름', '최종 결과', '진행 중']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- 처음 설정 ---------- */

function 설정하기() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : (SpreadsheetApp.getActive() || SpreadsheetApp.create('GM코드 퀴즈 기록'));
  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('기록 시트 주소: ' + ss.getUrl());

  const ro = sheet_(ss, ROSTER_SHEET);
  ro.getRange(1, 1, 1, 3).setValues([['이름', '최종 결과', '진행 중']]).setFontWeight('bold');
  ro.setFrozenRows(1);
  sheet_(ss, OTHER_SHEET);
  OLD_SHEETS.forEach(n => { const sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });

  // 지금부터 다음 평일 9시까지가 첫 기간
  if (!props.getProperty('LAST_CUTOFF')) props.setProperty('LAST_CUTOFF', String(latestCutoff_(new Date()).getTime()));

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '판정하기')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('판정하기').timeBased().everyDays(1).atHour(9).create(); // 매일 9~10시 사이 실행
  toast_(ss, '설정 완료: 명단 탭 A열 2행부터 학생 이름을 넣으세요.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('GM코드 퀴즈')
    .addItem('지금 9시 판정 반영하기', '판정하기')
    .addToUi();
}

/* ---------- 응시 결과 받기 (퀴즈가 POST로 보냄) ---------- */

function doPost(e) {
  const out = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 20);
    const score = Number(d.score);
    if (!name || !Number.isInteger(score) || score < 0 || score > 100) return out({ ok: false, error: 'bad data' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = ss_();
      closeIfDue_(ss); // 9시가 지났는데 아직 넘기지 않았으면 먼저 넘기고 기록한다
      mark_(ss, name, score);
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

// 칸 글자 '합격 (100)'에서 결과와 점수를 읽는다
function parse_(v) {
  const m = String(v).match(/^(합격|불합격)(?:\s*\((\d+)\))?/);
  return m ? { passed: m[1] === '합격', best: m[2] ? Number(m[2]) : -1 } : null;
}
function color_(v) { return String(v).indexOf('불합격') === 0 ? COLORS['불합격'] : String(v).indexOf('합격') === 0 ? COLORS['합격'] : null; }

// 학생 줄의 '진행 중' 칸을 바꾼다. 한 번 합격하면 그 기간에는 합격으로 남고, 점수는 그 기간 최고 점수.
function mark_(ss, name, score) {
  const key = norm_(name);
  const ro = sheet_(ss, ROSTER_SHEET);
  let sh = ro, row = findRow_(ro, key);
  if (row < 0) {
    sh = sheet_(ss, OTHER_SHEET);
    row = findRow_(sh, key);
    if (row < 0) {
      row = Math.max(2, sh.getLastRow() + 1);
      sh.getRange(row, 1).setValue(safe_(name));
    }
  }
  const cell = sh.getRange(row, 3);
  const prev = parse_(cell.getValue()) || { passed: false, best: -1 };
  const best = Math.max(prev.best, score);
  const v = (prev.passed || best >= PASS_SCORE ? '합격' : '불합격') + ' (' + best + ')';
  cell.setValue(v).setBackground(color_(v));
}

function findRow_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const names = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) if (norm_(names[i][0]) === key) return i + 2;
  return -1;
}

/* ---------- 평일 9시 판정 ---------- */

function 판정하기() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    if (closeIfDue_(ss)) toast_(ss, '9시 판정을 반영했습니다.');
    else toast_(ss, '이미 가장 최근 9시 판정이 반영되어 있습니다.');
  } finally {
    lock.releaseLock();
  }
}

// 마지막으로 넘긴 뒤 평일 9시가 지났으면 '진행 중'을 '최종 결과'로 옮긴다.
function closeIfDue_(ss) {
  const props = PropertiesService.getScriptProperties();
  const cutoff = latestCutoff_(new Date());
  const last = Number(props.getProperty('LAST_CUTOFF') || 0);
  if (cutoff.getTime() <= last) return false;

  const head = '최종 결과 (' + label_(prevCutoff_(cutoff)) + ' ~ ' + label_(cutoff) + ')';

  // 명단: 응시하지 않았으면 불합격
  const ro = sheet_(ss, ROSTER_SHEET);
  ro.getRange(1, 2).setValue(head);
  const n = ro.getLastRow();
  if (n >= 2) {
    const vals = ro.getRange(2, 1, n - 1, 3).getValues();
    const out = vals.map(r => (String(r[0]).trim() ? [String(r[2]) || '불합격 (미응시)', ''] : ['', '']));
    ro.getRange(2, 2, out.length, 2).setValues(out);
    ro.getRange(2, 2, out.length, 2).setBackgrounds(out.map(r => [color_(r[0]), null]));
  }

  // 명단 외 이름: 이번 기간에 응시한 이름만 남긴다
  const ot = sheet_(ss, OTHER_SHEET);
  ot.getRange(1, 2).setValue(head);
  const m = ot.getLastRow();
  if (m >= 2) {
    const keep = ot.getRange(2, 1, m - 1, 3).getValues()
      .filter(r => String(r[0]).trim() && String(r[2]))
      .map(r => [safe_(String(r[0]).replace(/^'/, '')), String(r[2]), '']);
    ot.getRange(2, 1, m - 1, 3).clearContent().setBackground(null);
    if (keep.length) {
      ot.getRange(2, 1, keep.length, 3).setValues(keep);
      ot.getRange(2, 2, keep.length, 1).setBackgrounds(keep.map(r => [color_(r[1])]));
    }
  }

  props.setProperty('LAST_CUTOFF', String(cutoff.getTime()));
  return true;
}
