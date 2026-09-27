// 入力記録を小さく保つ（1ファイル1000万セルまで）
//  1. 入力記録シートの使っていない列（S列〜）を削る
//  2. 年が変わったら、終わったロット（完了・削除済み）の記録を別のスプレッドシート「LPG容器 入力記録_2026」へ移す
//     入力中のロットの記録は残す。ロット単位でまとめて移すので、1つのロットの記録が分かれることはない
//     移した記録は、再開・再送信のときにそこから読み戻す

const ARCHIVE_PREFIX = 'LPG容器 入力記録_';
const ARCHIVE_CHUNK = 8000;   // 1回に調べる行数（入力の待ち時間を短くするため小分けにする）

// 入力画面の読み込み後に裏で呼ばれる（画面は待たない）
function runMaintenance() {
  const out = { trimmed: false, archived: 0, archiveDone: true };
  try { out.trimmed = trimLogColumns_(); } catch (e) { console.error('列の削除に失敗: ' + e.message); }
  try { Object.assign(out, archiveLogStep_()); } catch (e) { console.error('入力記録の移動に失敗: ' + e.message); }
  try {
    // 前日までの採番表PDF。3日分ずつ作り、残りがあれば続けて呼ばせる
    if (saibanStep_() >= 3) out.archiveDone = false;
  } catch (e) { console.error('採番表の作成に失敗: ' + e.message); }
  return out;
}

function trimLogColumns_() {
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log) return false;
  moveLogColumns_(log);
  const extra = log.getMaxColumns() - LOG_HEADERS.length;
  if (extra <= 0 || log.getLastColumn() > LOG_HEADERS.length) return false;
  log.deleteColumns(LOG_HEADERS.length + 1, extra);
  return true;
}

function tokyoYear_(d) { return Number(Utilities.formatDate(d || new Date(), 'Asia/Tokyo', 'yyyy')); }

function archiveIds_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('LOG_ARCHIVE_IDS') || '{}'); } catch (e) { return {}; }
}

// 移し先のスプレッドシート（年ごと）。PDFと同じ「LPG容器 検査成績表PDF」フォルダに作る
function archiveSheet_(year) {
  const props = PropertiesService.getScriptProperties();
  const ids = archiveIds_();
  if (ids[year]) {
    try { return SpreadsheetApp.openById(ids[year]).getSheets()[0]; } catch (e) { /* 消されていたら作り直す */ }
  }
  const ss = SpreadsheetApp.create(ARCHIVE_PREFIX + year);
  try { DriveApp.getFileById(ss.getId()).moveTo(getPdfFolder_()); } catch (e) { /* 移せなくてもマイドライブに残る */ }
  const sh = ss.getSheets()[0].setName(SHEET_LOG);
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > LOG_HEADERS.length) sh.deleteColumns(LOG_HEADERS.length + 1, sh.getMaxColumns() - LOG_HEADERS.length);
  ids[year] = ss.getId();
  props.setProperty('LOG_ARCHIVE_IDS', JSON.stringify(ids));
  return sh;
}

// 文字として入れている列（先頭の0を残す）
const LOG_TEXT_COLS = ['容器番号', '表示番号', '入力番号(NG時)'];
function asWritable_(row) {
  return row.map(function(v, i) {
    return LOG_TEXT_COLS.indexOf(LOG_HEADERS[i]) >= 0 && v !== '' && v !== null ? "'" + v : v;
  });
}

// 年が変わっていたら、終わったロットの記録を少しずつ移す。1回に ARCHIVE_CHUNK 行まで
function archiveLogStep_() {
  const props = PropertiesService.getScriptProperties();
  const year = tokyoYear_() - 1;               // 移す年（去年）
  const done = props.getProperty('LOG_ARCHIVE_DONE');
  if (!done) { props.setProperty('LOG_ARCHIVE_DONE', String(year)); return { archived: 0, archiveDone: true }; }
  if (Number(done) >= year) return { archived: 0, archiveDone: true };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { archived: 0, archiveDone: false };
  try {
    const log = getSheet_(SHEET_LOG);
    moveLogColumns_(log, true);
    const active = {};
    readLots_().forEach(function(l) { if (l.status === STATUS_ACTIVE) active[l.lotId] = true; });
    const from = Number(props.getProperty('LOG_ARCHIVE_CURSOR') || 2);
    const last = log.getLastRow();
    if (last < from) {
      props.setProperty('LOG_ARCHIVE_DONE', String(year));
      props.deleteProperty('LOG_ARCHIVE_CURSOR');
      return { archived: 0, archiveDone: true };
    }
    const n = Math.min(ARCHIVE_CHUNK, last - from + 1);
    const rows = log.getRange(from, 1, n, LOG_HEADERS.length).getValues();
    const keep = [], move = [];
    rows.forEach(function(r) { (active[String(r[LC_LOT])] ? keep : move).push(r); });
    if (move.length) {
      const dest = archiveSheet_(year);
      dest.getRange(dest.getLastRow() + 1, 1, move.length, LOG_HEADERS.length).setValues(move.map(asWritable_));
      formatLogMass_(dest, dest.getLastRow() - move.length + 1, move.length);
      if (keep.length) log.getRange(from, 1, keep.length, LOG_HEADERS.length).setValues(keep.map(asWritable_));
      log.deleteRows(from + keep.length, move.length);
    }
    const next = from + keep.length;
    if (from + n - 1 >= last) {
      props.setProperty('LOG_ARCHIVE_DONE', String(year));
      props.deleteProperty('LOG_ARCHIVE_CURSOR');
      return { archived: move.length, archiveDone: true };
    }
    props.setProperty('LOG_ARCHIVE_CURSOR', String(next));
    return { archived: move.length, archiveDone: false };
  } finally {
    lock.releaseLock();
  }
}

// 移したロットの記録を、入力記録シートの行と同じ形で返す（新しい年のものから探す）
function archivedRowsForLot_(lotId) {
  const ids = archiveIds_();
  const years = Object.keys(ids).sort().reverse();
  for (let i = 0; i < years.length; i++) {
    let sh;
    try { sh = SpreadsheetApp.openById(ids[years[i]]).getSheets()[0]; } catch (e) { continue; }
    if (sh.getLastRow() < 2) continue;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEADERS.length).getValues()
      .filter(function(r) { return String(r[LC_LOT]) === lotId; });
    if (rows.length) return rows;
  }
  return [];
}

// 移したロットを再開するとき、記録を入力記録シートへ戻す（移し先の行は状態「戻し済み」にする）
function restoreArchivedLot_(lot) {
  const log = getSheet_(SHEET_LOG);
  if (log.getLastRow() >= 2) {
    const has = log.getRange(2, LC_LOT + 1, log.getLastRow() - 1, 1).getValues()
      .some(function(r) { return String(r[0]) === lot.lotId; });
    if (has) return 0;
  }
  const ids = archiveIds_();
  let moved = 0;
  Object.keys(ids).sort().forEach(function(y) {
    let sh;
    try { sh = SpreadsheetApp.openById(ids[y]).getSheets()[0]; } catch (e) { return; }
    if (sh.getLastRow() < 2) return;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEADERS.length).getValues();
    const rows = [];
    v.forEach(function(r, i) {
      if (String(r[LC_LOT]) !== lot.lotId || r[LC_STATUS] === '戻し済み') return;
      rows.push(r.slice());
      sh.getRange(2 + i, LC_STATUS + 1).setValue('戻し済み');
    });
    if (rows.length) {
      log.getRange(log.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows.map(asWritable_));
      formatLogMass_(log, log.getLastRow() - rows.length + 1, rows.length);
      moved += rows.length;
    }
  });
  return moved;
}
