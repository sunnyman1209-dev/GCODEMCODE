/**
 * G·M코드 퀴즈 판정 스크립트
 * 붙이는 곳: 설문 응답이 쌓이는 구글 시트 > 확장 프로그램 > Apps Script
 *
 * 평일 오전 9시마다 '전 평일 9시 ~ 오늘 9시' 사이의 응시 기록만 보고
 * 이름별 최종 결과(합격 / 불합격 / 미응시)를 판정한다.
 * 그 기간에 한 번이라도 90점 이상이면 합격이다. 날마다 새로 판정한다.
 *
 * 만드는 탭
 *   명단        : 선생님이 A열에 학생 이름을 한 줄에 한 명씩 붙여 넣는다(처음 실행 때 빈 탭을 만들어 줌).
 *   최근 판정   : 가장 최근 판정의 이름별 결과, 최고 점수, 응시 횟수
 *   판정 기록   : 이름별로 날짜마다 결과를 한 칸씩 쌓은 표
 *   명단 외 이름: 명단에 없는 이름으로 응시한 기록(오타, 장난 이름 확인용)
 */

const PASS_SCORE = 90;
const ROSTER_SHEET = '명단';
const TODAY_SHEET = '최근 판정';
const LOG_SHEET = '판정 기록';
const OTHER_SHEET = '명단 외 이름';
const DAYS = ['일', '월', '화', '수', '목', '금', '토'];
const COLORS = { '합격': '#d9f2e3', '불합격': '#fde2e1', '미응시': '#eeeeee' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('GM코드 퀴즈')
    .addItem('지금 판정하기', '판정하기')
    .addItem('평일 9시 자동 판정 켜기', '자동판정켜기')
    .addToUi();
}

// 매일 오전 9시~10시 사이에 판정하기를 실행한다(주말에는 판정하기가 스스로 건너뜀).
function 자동판정켜기() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '판정하기')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('판정하기').timeBased().everyDays(1).atHour(9).create();
  SpreadsheetApp.getActive().toast('평일 오전 9시~10시 사이에 자동으로 판정합니다.');
}

// 이름 비교는 띄어쓰기를 무시한다(홍 길동 = 홍길동).
function norm_(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }
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

// 설문 응답 탭: 첫 줄에 '이름'과 '점수'가 있는 탭
function findResponseSheet_(ss) {
  for (const sh of ss.getSheets()) {
    if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) continue;
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (head.indexOf('이름') >= 0 && head.indexOf('점수') >= 0) return { sh, head };
  }
  return null;
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

  const ss = SpreadsheetApp.getActive();
  const found = findResponseSheet_(ss);
  if (!found) throw new Error('첫 줄에 "이름"과 "점수"가 있는 응답 탭을 찾지 못했습니다.');
  const { sh, head } = found;
  const cTime = 0;
  const cName = head.indexOf('이름');
  const cScore = head.indexOf('점수');
  const cResult = head.indexOf('결과');

  const cutoff = latestCutoff_(now);
  const start = prevCutoff_(cutoff);

  // 기간 안의 기록을 이름별로 모은다
  const stats = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(r => {
      const t = r[cTime];
      if (!(t instanceof Date) || t < start || t >= cutoff) return;
      const key = norm_(r[cName]);
      if (!key) return;
      const score = Number(r[cScore]) || 0;
      const passed = (cResult >= 0 && String(r[cResult]) === '합격') || score >= PASS_SCORE;
      const s = stats[key] || (stats[key] = { name: String(r[cName]).trim(), tries: 0, best: 0, passed: false });
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
  const rows = roster.map(n => { const x = result(n); return [n, x.r, x.best, x.tries]; });
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
      log.getRange(i + 2, 1).setValue(n);
    }
    const r = result(n).r;
    log.getRange(i + 2, col).setValue(r).setBackground(COLORS[r]);
  });

  // 명단 외 이름
  const other = ss.getSheetByName(OTHER_SHEET) || ss.insertSheet(OTHER_SHEET);
  other.clear();
  other.getRange(1, 1).setValue(period + ' (명단에 없는 이름으로 응시한 기록)').setFontWeight('bold');
  other.getRange(3, 1, 1, 4).setValues([['적은 이름', '결과', '최고 점수', '응시 횟수']]).setFontWeight('bold');
  const orows = others.map(s => [s.name, s.passed ? '합격' : '불합격', s.best, s.tries]);
  if (orows.length) other.getRange(4, 1, orows.length, 4).setValues(orows);

  ss.toast(period + ' 판정을 마쳤습니다.');
}
