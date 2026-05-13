// ============================================================
// 고3 학습 플래너 — Google Apps Script Backend
// GitHub Pages CORS 우회를 위해 JSONP 지원 포함
// ============================================================

const SHEET_LOG  = '학습기록';
const START_DATE = '2026-05-16';
const END_DATE   = '2026-11-19';
const WD_SUBJ = ['수학','국어','영어단어','사탐'];
const WE_SUBJ = ['영어','역사'];

function doGet(e) {
  const p  = e.parameter || {};
  const cb = p.callback;
  let result;
  try { result = dispatch(p.action, p, {}); }
  catch(err) { result = { error: err.message }; }
  const json = JSON.stringify(result);
  const out  = cb ? `${cb}(${json})` : json;
  return ContentService.createTextOutput(out)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}
  let result;
  try { result = dispatch(body.action, {}, body); }
  catch(err) { result = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatch(action, g, p) {
  const a = Object.assign({}, g, p);
  switch(action) {
    case 'getDayData':    return getDayData(a.date);
    case 'saveDayData':   return saveDayData(a);
    case 'getWeekStats':  return getWeekStats(a.weekStart);
    case 'getCalendar':
    case 'getMonthStats': return getMonthStats(a.month);
    case 'getAllStats':    return getAllStats();
    default: return { error: 'Unknown action: ' + action };
  }
}

function getSubs(dateStr) {
  return (new Date(dateStr).getDay() % 6 === 0) ? WE_SUBJ : WD_SUBJ;
}
function dayLbl(dateStr) {
  return ['일','월','화','수','목','금','토'][new Date(dateStr).getDay()];
}
function inR(ds) { return ds >= START_DATE && ds <= END_DATE; }
function fmtD(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(['날짜','요일','과목','완료','메모','저장시각']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getDayData(dateStr) {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  const subjs = getSubs(dateStr);
  const result = {};
  subjs.forEach(s => { result[s] = { done: false, memo: '' }; });
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]) === dateStr && subjs.includes(r[2])) {
      result[r[2]] = { done: r[3] === true || String(r[3]).toUpperCase() === 'TRUE', memo: r[4] || '' };
    }
  }
  return { date: dateStr, day: dayLbl(dateStr), subjects: result };
}

function saveDayData(p) {
  const dateStr = p.date, data = p.data || {};
  if (!inR(dateStr)) return { error: '범위 밖 날짜' };
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  const subjs = getSubs(dateStr);
  const now = new Date().toISOString();
  subjs.forEach(subj => {
    const info = data[subj] || { done: false, memo: '' };
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === dateStr && rows[i][2] === subj) {
        sh.getRange(i+1, 4, 1, 3).setValues([[info.done, info.memo, now]]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([dateStr, dayLbl(dateStr), subj, info.done, info.memo, now]);
  });
  return { ok: true, date: dateStr };
}

function getWeekStats(weekStartStr) {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  const days = [];
  for (let d = 0; d < 7; d++) {
    const dt = new Date(weekStartStr); dt.setDate(dt.getDate() + d);
    const dateStr = fmtD(dt);
    if (!inR(dateStr)) { days.push(null); continue; }
    const subjs = getSubs(dateStr);
    let done = 0; const memos = [];
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === dateStr) {
        if (rows[i][3] === true || String(rows[i][3]).toUpperCase() === 'TRUE') done++;
        if (rows[i][4]) memos.push({ subj: rows[i][2], memo: rows[i][4] });
      }
    }
    days.push({ date: dateStr, day: dayLbl(dateStr), done, total: subjs.length, memos });
  }
  return { weekStart: weekStartStr, days };
}

function getMonthStats(month) {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  const byDate = {};
  for (let i = 1; i < rows.length; i++) {
    const ds = String(rows[i][0]);
    if (!ds.startsWith(month) || !inR(ds)) continue;
    if (!byDate[ds]) byDate[ds] = { done: 0, total: getSubs(ds).length };
    if (rows[i][3] === true || String(rows[i][3]).toUpperCase() === 'TRUE') byDate[ds].done++;
  }
  const [y, m] = month.split('-').map(Number);
  const dim = new Date(y, m, 0).getDate();
  for (let d = 1; d <= dim; d++) {
    const ds = `${month}-${String(d).padStart(2,'0')}`;
    if (inR(ds) && !byDate[ds]) byDate[ds] = { done: 0, total: getSubs(ds).length };
  }
  const dayArr = Object.keys(byDate).sort().map(k => ({ date: k, ...byDate[k] }));
  const tD = dayArr.reduce((s,x) => s+x.done, 0);
  const tG = dayArr.reduce((s,x) => s+x.total, 0);
  return { month, days: dayArr, totalDone: tD, totalGoal: tG, rate: tG ? Math.round(tD/tG*100) : 0 };
}

function getAllStats() {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  const byMonth = {}, bySubj = {}, doneMap = {};

  for (let i = 1; i < rows.length; i++) {
    const ds = String(rows[i][0]), subj = rows[i][2];
    const done = rows[i][3] === true || String(rows[i][3]).toUpperCase() === 'TRUE';
    if (!inR(ds)) continue;
    const mo = ds.substring(0,7);
    if (!byMonth[mo]) byMonth[mo] = { done:0, total:0 };
    byMonth[mo].total++; if (done) byMonth[mo].done++;
    if (!bySubj[subj]) bySubj[subj] = { done:0, total:0 };
    bySubj[subj].total++; if (done) bySubj[subj].done++;
    doneMap[`${ds}__${subj}`] = done;
  }

  const problematic = [];
  [...WD_SUBJ, ...WE_SUBJ].forEach(subj => {
    let max = 0, streak = 0;
    for (let dt = new Date(START_DATE); dt <= new Date(END_DATE); dt.setDate(dt.getDate()+1)) {
      const ds = fmtD(dt);
      if (!getSubs(ds).includes(subj)) { streak = 0; continue; }
      if (!doneMap[`${ds}__${subj}`]) { streak++; if (streak > max) max = streak; }
      else streak = 0;
    }
    if (max >= 3) problematic.push({ subj, maxStreak: max });
  });

  return {
    byMonth:    Object.keys(byMonth).sort().map(m => ({ month:m, ...byMonth[m], rate: byMonth[m].total ? Math.round(byMonth[m].done/byMonth[m].total*100) : 0 })),
    bySubject:  Object.keys(bySubj).map(s => ({ subj:s, ...bySubj[s], rate: bySubj[s].total ? Math.round(bySubj[s].done/bySubj[s].total*100) : 0 })),
    problematic
  };
}
