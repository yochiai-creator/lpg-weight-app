// 入力記録を小さく保つ（1ファイル1000万セルまで）
//  1. 入力記録シートの使っていない列（S列〜）を削る
//  2. 年が変わったら、終わったロット（完了・削除済み）の記録を別のスプレッドシート「LPG容器 入力記録_2026」へ移す
//     入力中のロットの記録は残す。ロット単位でまとめて移すので、1つのロットの記録が分かれることはない
//     移した記録は、再開・再送信のときにそこから読み戻す

const ARCHIVE_PREFIX = 'LPG容器 入力記録_';
const ARCHIVE_CHUNK = 8000;   // 1回に調べる行数（入力の待ち時間を短くするため小分けにする）

// 入力画面の読み込み後に裏で呼ばれる（画面は待たない）
// opts.scheduled: 朝・夜の自動処理から。画面からの呼び出しは、その日の朝の自動処理が済んでいれば何もしない（起動を軽くする）
function runMaintenance(opts) {
  const out = { trimmed: false, archived: 0, archiveDone: true };
  if (!(opts && opts.scheduled)) {
    try { if (sharedProps_().getProperty('MAINT_DAY') === tokyoDate_(new Date())) return out; } catch (e) { console.error(e.message); }
  }
  try { ensureMorningTrigger_(); } catch (e) { console.error('朝の自動処理の設定に失敗: ' + e.message); }
  try { mergeRootFolders_(); ensureFolderLayout_(); ensurePdfKindFolders_(); } catch (e) { console.error('フォルダの整理に失敗: ' + e.message); }
  try { fixPaddedLots_(); } catch (e) { console.error('ロット番号の修正に失敗: ' + e.message); }
  // 2026-09-29: 完了済みなのに「最近完了したロット」の削除ボタンで消えてしまった HEP59701 を1回だけ戻す
  try {
    const props = sharedProps_();
    if (props.getProperty('RESTORED_HEP59701') !== '1') {
      try { undeleteLot('HEP59701'); } catch (e) { console.log('HEP59701: ' + e.message); }
      props.setProperty('RESTORED_HEP59701', '1');
    }
  } catch (e) { console.error('ロットの復元に失敗: ' + e.message); }
  try { sokoSaibanRedo_(); } catch (e) { console.error('底黒を外した採番表の作り直しに失敗: ' + e.message); }
  try { retime20261021_(); } catch (e) { console.error('10/21の時刻の付け直しに失敗: ' + e.message); }
  try { out.trimmed = trimLogColumns_(); } catch (e) { console.error('列の削除に失敗: ' + e.message); }
  try { Object.assign(out, archiveLogStep_()); } catch (e) { console.error('入力記録の移動に失敗: ' + e.message); }
  try {
    // 前日までの採番表PDF。3日分ずつ作り、残りがあれば続けて呼ばせる
    if (saibanStep_() >= 3) out.archiveDone = false;
    // 時刻の列がない採番表を作り直す（1回だけ・3日分ずつ）
    if (out.archiveDone && saibanRedoStep_() > 0) out.archiveDone = false;
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
  try { return JSON.parse(sharedProps_().getProperty('LOG_ARCHIVE_IDS') || '{}'); } catch (e) { return {}; }
}

// 移し先のスプレッドシート（年ごと）。「LPG容器 質量入力 / 入力記録（過去分）」に作る
function archiveSheet_(year) {
  const props = sharedProps_();
  const ids = archiveIds_();
  if (ids[year]) {
    try { return SpreadsheetApp.openById(ids[year]).getSheets()[0]; } catch (e) { /* 消されていたら作り直す */ }
  }
  // 別のプロジェクト（旧URL）が作ったものがあればそれを使う
  const found = findSpreadsheetIn_(childFolder_(getPdfFolder_(), ARCHIVE_FOLDER), ARCHIVE_PREFIX + year);
  if (found) { ids[year] = found.getId(); props.setProperty('LOG_ARCHIVE_IDS', JSON.stringify(ids)); return found.getSheets()[0]; }
  const ss = SpreadsheetApp.create(ARCHIVE_PREFIX + year);
  try { DriveApp.getFileById(ss.getId()).moveTo(childFolder_(getPdfFolder_(), ARCHIVE_FOLDER)); } catch (e) { /* 移せなくてもマイドライブに残る */ }
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
  const props = sharedProps_();
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

// 2026-10-21: v80のiPadで送信が止まり、夜（18:44〜）にまとめて送り直したため、入力日時が届いた時刻になった。
// 対象は 18:44〜20:46 に届いた行だけ。v80はiPadで入れた時刻を送っていないので、作業日報の作業時間（8:36〜12:50・13:40〜15:45）に流れた順番のまま均等に割り振る（推定）。
// 流れていない9ロット（備考「採番なし」）は動かさない。1回だけ。終わったらその日の採番表を作り直す
const RETIME_1021_SKIP = ['HEP75301', 'HEP75801', 'HEP75901', 'HEP76001', 'HEP76101', 'HEP76201', 'HEP76301', 'HEP76701', 'HEP76901'];
function retime20261021_() {
  const props = sharedProps_();
  if (props.getProperty('RETIMED_20261021') === '1') return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
    if (!log || log.getLastRow() < 2) return;
    const last = log.getLastRow(), from = Math.max(2, last - 30000);
    const vals = log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getValues();
    // 10/21の送り直しがあったスプレッドシートだけ（9ロットの記録がなければ何もしない）
    if (!vals.some(function(r) { return RETIME_1021_SKIP.indexOf(String(r[LC_LOT])) >= 0; })) { props.setProperty('RETIMED_20261021', '1'); return; }
    const cut = new Date('2026-10-21T18:44:00+09:00').getTime(), end = new Date('2026-10-21T20:46:00+09:00').getTime();   // 送り直しは 20:44:57 まで
    const idx = [];
    vals.forEach(function(r, i) {
      const t = r[LC_TIME];
      if (t instanceof Date && t.getTime() >= cut && t.getTime() < end && RETIME_1021_SKIP.indexOf(String(r[LC_LOT])) < 0) idx.push(i);
    });
    if (idx.length) {
      idx.sort(function(a, b) { return (vals[a][LC_TIME] - vals[b][LC_TIME]) || (a - b); });
      const spans = [['08:36', '12:50'], ['13:40', '15:45']].map(function(p) {
        return [new Date('2026-10-21T' + p[0] + ':00+09:00').getTime(), new Date('2026-10-21T' + p[1] + ':00+09:00').getTime()];
      });
      const total = spans.reduce(function(s, p) { return s + p[1] - p[0]; }, 0);
      const step = total / idx.length;
      idx.forEach(function(i, k) {
        let off = Math.round(k * step);
        let t = spans[0][0] + off;
        if (off >= spans[0][1] - spans[0][0]) t = spans[1][0] + off - (spans[0][1] - spans[0][0]);
        vals[i][LC_TIME] = new Date(t);
      });
      const lo = Math.min.apply(null, idx), hi = Math.max.apply(null, idx);
      log.getRange(from + lo, LC_TIME + 1, hi - lo + 1, 1).setValues(vals.slice(lo, hi + 1).map(function(r) { return [r[LC_TIME]]; }));
      SpreadsheetApp.flush();
    }
    props.setProperty('RETIMED_20261021', '1');
    console.log('10/21の時刻を付け直した: ' + idx.length + '行');
  } finally {
    lock.releaseLock();
  }
  makeSaibanPdf({ date: '2026-10-21' });
}

// v109: 採番表から底黒を外したので、底黒の入力があった前日までの日を1回だけ作り直す（saibanStep_ が SAIBAN_STALE を作り直す）
function sokoSaibanRedo_() {
  const props = sharedProps_();
  if (props.getProperty('SOKO_SAIBAN_REDO') === '1') return;
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  const days = [];
  if (log && log.getLastRow() >= 2) {
    const last = log.getLastRow(), from = Math.max(2, last - 30000);
    const soko = {};
    readLots_().forEach(function(l) { if (l.spec === SPEC_SOKO) soko[l.lotId] = true; });
    const today = tokyoDate_(new Date());
    log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getValues().forEach(function(r) {
      const t = r[LC_TIME];
      if (!(t instanceof Date) || r[LC_STATUS] !== '有効') return;
      if (r[LC_KIND] !== SPEC_SOKO && !soko[String(r[LC_LOT])]) return;
      const d = tokyoDate_(t);
      if (d < today && days.indexOf(d) < 0) days.push(d);
    });
  }
  const list = (props.getProperty('SAIBAN_STALE') || '').split(',').filter(String);
  days.forEach(function(d) { if (list.indexOf(d) < 0) list.push(d); });
  if (list.length) props.setProperty('SAIBAN_STALE', list.join(','));
  props.setProperty('SOKO_SAIBAN_REDO', '1');
}
