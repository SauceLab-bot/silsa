/**
 * 재고 실사앱 (구글시트 연동)
 * - 이 스크립트는 '재고표' 시트에 붙어 있고, 웹앱(Index.html)으로 실사를 입력합니다.
 * - 입력은 웹앱을 배포한 본인 계정만 가능합니다 (배포 설정: 실행 = 나, 액세스 = 나만).
 * - 시트는 다른 사람에게 '뷰어'로만 공유하면 보기만 가능합니다.
 */

var SHEET_NAME = '재고표';
var HIST_NAME = '이력';
var CONF_NAME = '설정';
var FIRST_ROW = 6;          // 품목이 시작되는 행
var NCOL = 22;              // A ~ V
var C = { NAME: 1, EXP: 2, OUT0: 3, OUT_END: 18, BAL: 19, TOT: 20, IN_D: 21, IN_Q: 22 };

// 품목별 1파렛트 박스 수 (예전 실사앱 기준). 설정 시트에서 바꿀 수 있습니다.
var DEFAULT_PLT = {
  '영풍)요뽀끼 매콤달콤컵떡볶이140g': 48, '영풍)요뽀끼 짜장컵떡볶이120g': 48, '영풍)요뽀끼 치즈컵떡볶이120g': 48,
  '예스미미)라면조리기 미미네 매콤라볶이290g': 30, '성지)황금가 찹쌀누룽지700g': 48,
  '한울)Yellow)쫀득한고구마스틱60g': 100, '한울)촉촉한꿀고구마말랭이60g': 48,
  '에스엠이지)유니콘 폼 마쉬멜로42g': 114, '에스엠이지)스위티 하트 마쉬멜로45g': 114,
  '에프투에스)크리켓라이터미니(10*50)': 48, '에프투에스)크리켓라이터퓨전(10g*50)': 48,
  '셀피알)개빼로(닭고기&치즈껌)': 54, '셀피알)방수밴드': 40, '셀피알)방수우의(블랙)': 30, '셀피알)초박형 콘돔': 32,
  '에스엠이지)대만여주차590ml': 48, '에스엠이지)메이플스토리 노래나오는캔디12g': 32,
  '에스엠이지)브레드노래나오는캔디12G': 32, '에스엠이지)비이)킨더카드2p': 384,
  '에스엠이지)씨케이 스트롱브렌드500g': 160, '에스엠이지)어라운지 브라운 너티500g(10입)': 60,
  '에스엠이지)터키힐메이플크림쿠키200g': 60, '칠갑)매운맛해물수제비': 42,
  '한울)소담미고구마바23g*10': 30, '한울)소담미단호박바23g*10': 30, '한울)이집트골든데이츠': 40
};

/* ---------- 웹앱 ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('재고 실사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 편집기에서 한 번 실행: 권한 승인 + 설정/이력 시트 준비 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  ensureConf_(ss);
  ensureHist_(ss);
  return '준비 완료: ' + ss.getName();
}

/* ---------- 시트 접근 ---------- */
function ss_() {
  var s = null;
  try { s = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (s) return s;
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
}
function sheet_(ss) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error("'" + SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  return sh;
}
function toNum_(v) { var n = Number(v); return (v === '' || v === null || isNaN(n)) ? 0 : n; }
function serial_(s) { var p = s.split('-'); return Math.round((Date.UTC(+p[0], +p[1] - 1, +p[2]) - Date.UTC(1899, 11, 30)) / 864e5); }
function cellDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  if (typeof v === 'number' && v > 20000) { var d = new Date(Date.UTC(1899, 11, 30) + v * 864e5); return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2); }
  var t = String(v || '').replace(/\D/g, '');
  if (t.length === 8) return t.slice(0, 4) + '-' + t.slice(4, 6) + '-' + t.slice(6, 8);
  return '';
}
function today_(tz) { return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); }

/** 품목 블록: 이름이 있는 행에서 다음 이름 전까지 */
function readBlocks_(sh) {
  var last = sh.getLastRow();
  if (last < FIRST_ROW) return [];
  var n = last - FIRST_ROW + 1;
  var vals = sh.getRange(FIRST_ROW, 1, n, NCOL).getValues();
  var blocks = [], cur = null;
  for (var i = 0; i < n; i++) {
    var name = String(vals[i][0]).trim();
    if (name) { cur = { name: name, start: FIRST_ROW + i, end: FIRST_ROW + i, rows: [] }; blocks.push(cur); }
    if (cur) { cur.end = FIRST_ROW + i; cur.rows.push(vals[i]); }
  }
  return blocks;
}
function lotsOf_(b, tz) {
  var out = [];
  b.rows.forEach(function (r) {
    var bal = toNum_(r[C.BAL - 1]);
    if (bal > 0) out.push({ d: cellDate_(r[C.EXP - 1], tz), n: bal });
  });
  out.sort(function (a, b2) { return a.d < b2.d ? -1 : a.d > b2.d ? 1 : 0; });
  return out;
}

/* ---------- 설정/이력 시트 ---------- */
function ensureConf_(ss) {
  var sh = ss.getSheetByName(CONF_NAME);
  var blocks = readBlocks_(sheet_(ss));
  if (!sh) {
    sh = ss.insertSheet(CONF_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['품목', '박스/PLT', '최근 실사']]);
    sh.getRange('C:C').setNumberFormat('@');
    try { sh.hideSheet(); } catch (e) {}
  }
  var have = {};
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
  rows.forEach(function (r, i) { have[String(r[0])] = i + 2; });
  blocks.forEach(function (b) {
    if (!have[b.name]) {
      sh.appendRow([b.name, DEFAULT_PLT[b.name] || 0, '']);
      have[b.name] = sh.getLastRow();
    }
  });
  return { sh: sh, rows: have };
}
function ensureHist_(ss) {
  var sh = ss.getSheetByName(HIST_NAME);
  if (!sh) {
    sh = ss.insertSheet(HIST_NAME);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange(1, 1, 1, 9).setValues([['시각', '품목', '구분', '유통기한', 'PLT', '입고수량', '출고합계', '잔량', '낱BOX']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- 웹앱에서 부르는 함수 ---------- */
function ping() { return 'pong ' + new Date().toISOString(); }
function getInitJson() { return JSON.stringify(getInit()); }

function getInit() {
  var ss = ss_(), sh = sheet_(ss), tz = ss.getSpreadsheetTimeZone();
  var conf = ensureConf_(ss);
  var confVals = conf.sh.getRange(2, 1, Math.max(conf.sh.getLastRow() - 1, 1), 3).getValues();
  var info = {};
  confVals.forEach(function (r) { info[String(r[0])] = { plt: toNum_(r[1]), last: String(r[2] || '') }; });
  var items = readBlocks_(sh).map(function (b) {
    var c = info[b.name] || { plt: 0, last: '' };
    return { name: b.name, slots: b.end - b.start + 1, plt: c.plt, last: c.last, lots: lotsOf_(b, tz) };
  });
  return { items: items, today: today_(tz), url: ss.getUrl() };
}

function setPlt(name, v) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), conf = ensureConf_(ss);
    var row = conf.rows[name];
    if (!row) throw new Error('품목을 찾을 수 없습니다.');
    v = Math.max(0, Math.floor(Number(v) || 0));
    conf.sh.getRange(row, 2).setValue(v);
    return v;
  } finally { lock.releaseLock(); }
}

/** p = { name, lots: [{d:'yyyy-mm-dd', p:PLT, b:낱BOX, n:총BOX}] }  → 그 품목의 유통기한·입고 수량을 실사값으로 교체 */
function saveItem(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), sh = sheet_(ss), tz = ss.getSpreadsheetTimeZone();
    var b = null;
    readBlocks_(sh).forEach(function (x) { if (x.name === p.name) b = x; });
    if (!b) throw new Error('재고표에서 품목을 찾을 수 없습니다: ' + p.name);
    var slots = b.end - b.start + 1;

    // 같은 유통기한은 합산
    var m = {};
    (p.lots || []).forEach(function (l) {
      var n = Number(l.n);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(l.d))) throw new Error('유통기한 형식이 올바르지 않습니다.');
      if (!(n >= 0) || isNaN(n)) throw new Error('수량이 올바르지 않습니다.');
      if (n === 0) return;
      var t = m[l.d] || (m[l.d] = { d: l.d, p: 0, b: 0, n: 0 });
      t.p += toNum_(l.p); t.b += toNum_(l.b); t.n += n;
    });
    var lots = Object.keys(m).sort().map(function (k) { return m[k]; });
    if (lots.length > slots) {
      throw new Error('이 품목은 유통기한을 ' + slots + '개까지 적을 수 있어요. (지금 ' + lots.length + '개) 재고표에서 줄을 늘리거나 유통기한을 정리해 주세요.');
    }

    // 이력 보관 (변경 전 / 실사)
    var hist = ensureHist_(ss), now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
    var hrows = [];
    b.rows.forEach(function (r) {
      var inq = toNum_(r[C.IN_Q - 1]), exp = cellDate_(r[C.EXP - 1], tz), outs = 0;
      for (var c = C.OUT0 + 1; c <= C.OUT_END; c += 2) outs += toNum_(r[c - 1]);
      if (inq || exp || outs) hrows.push([now, p.name, '변경 전', exp, '', inq, outs, toNum_(r[C.BAL - 1]), '']);
    });
    if (lots.length) lots.forEach(function (l) { hrows.push([now, p.name, '실사', l.d, l.p, l.n, 0, l.n, l.b]); });
    else hrows.push([now, p.name, '실사', '', '', 0, 0, 0, '']);
    hist.getRange(hist.getLastRow() + 1, 1, hrows.length, 9).setValues(hrows);

    // 시트 갱신: 유통기한·입고 날짜/수량 = 실사값, 출고 기록 비움 (잔량·합계 수식은 그대로)
    var today = today_(tz);
    for (var k = 0; k < slots; k++) {
      var r = b.start + k, l = lots[k];
      sh.getRange(r, C.OUT0, 1, C.OUT_END - C.OUT0 + 1).clearContent();
      if (l) {
        sh.getRange(r, C.EXP).setValue(serial_(l.d));
        sh.getRange(r, C.IN_D).setValue(serial_(today));
        sh.getRange(r, C.IN_Q).setValue(l.n);
      } else {
        sh.getRange(r, C.EXP).clearContent();
        sh.getRange(r, C.IN_D, 1, 2).clearContent();
      }
    }
    sh.getRange(b.start, C.EXP, slots, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(b.start, C.IN_D, slots, 1).setNumberFormat('yyyy-mm-dd');

    var conf = ensureConf_(ss);
    conf.sh.getRange(conf.rows[p.name], 3).setValue(now);
    SpreadsheetApp.flush();
    return { saved: true, at: now };
  } finally { lock.releaseLock(); }
}

function getHistory() {
  var ss = ss_(), sh = ensureHist_(ss);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - 400);
  var vals = sh.getRange(from, 1, last - from + 1, 9).getValues();
  var groups = [], idx = {};
  vals.forEach(function (r) {
    var key = r[0] + '|' + r[1];
    var g = idx[key];
    if (!g) { g = idx[key] = { ts: String(r[0]), name: String(r[1]), before: [], after: [] }; groups.push(g); }
    var row = { d: String(r[3] instanceof Date ? Utilities.formatDate(r[3], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : r[3] || ''), p: r[4], q: toNum_(r[5]), bal: toNum_(r[7]), b: r[8] };
    (r[2] === '실사' ? g.after : g.before).push(row);
  });
  return groups.reverse().slice(0, 60);
}
