// LPG容器 質量転記アプリ
// 検査スタッフがiPad（外付けテンキー）で「容器番号」と「質量」を入力すると、
// 組容器番号から記録先の検査成績表シートを判定して自動転記する。

const SHEET_LOTS = 'ロット';
const SHEET_LOG = '入力記録';
const SHEET_TEMPLATE = '書式_成績表';
const REPORT_PREFIX = '成績表_';

const LOT_HEADERS = ['ロットID', '登録日時', '記号', '開始番号', '終了番号', '容器区分', '状態',
  '成績表シート', '代表容器番号', '耐圧試験日', '全増加(cm3)', '恒久増加(cm3)', '恒久増加率(%)',
  'PDF', '登録者', '送信日時', '送信先', '仕様'];
const LOG_HEADERS = ['記録ID', '入力日時', 'ロットID', 'グループNo', '容器番号', '容器区分', '表示番号', '質量(kg)', '区分',
  '一致結果', '入力番号(NG時)', '上書き前', 'Googleアカウント', '状態', '入力者', '端末', '取消日時', '備考'];
// 入力記録の列の位置（0始まり）。列の並びを変えても処理はこの表で追従する
const LC = {};
LOG_HEADERS.forEach(function(h, i) { LC[h] = i; });
const LC_ID = LC['記録ID'], LC_TIME = LC['入力日時'], LC_LOT = LC['ロットID'], LC_GROUP = LC['グループNo'],
  LC_SERIAL = LC['容器番号'], LC_KINDSIZE = LC['容器区分'], LC_DISP = LC['表示番号'], LC_MASS = LC['質量(kg)'],
  LC_KIND = LC['区分'], LC_MATCH = LC['一致結果'], LC_NG = LC['入力番号(NG時)'], LC_PREV = LC['上書き前'],
  LC_STATUS = LC['状態'], LC_WORKER = LC['入力者'], LC_UNDO = LC['取消日時'], LC_NOTE = LC['備考'];

const STATUS_ACTIVE = '入力中';
const STATUS_DONE = '完了';
const MISSING = '欠番';
const REPAIR = '修正';        // ラインで品質不良のため修正に回した容器（欠番と同じく質量なしで記録）
const KIND_FIX = '訂正';      // 入力ミスを直した上書き（印や件数は出さない）
const KIND_DUP = 'ダブり';    // 同じ容器が2回流れてきた上書き
const NOTE_SEAL = 'シール違い';
const SPEC_SOKO = '底黒';      // 再搬入（底黒仕様）: 質量は入れず、流れた容器の質量欄に〇を付ける
const CIRCLE = '〇';

// 成績表の配置（元Excel様式）: 5ブロック × 20行、1ブロック6列
// 容器番号 | ☑ | 質量10の位 | 1の位 | "," | 小数1位
const GRID_FIRST_ROW = 5;
const GRID_ROWS = 20;
const GRID_FIRST_COL = 2; // B列
const BLOCK_WIDTH = 6;
const BLOCKS = 5;
const LOT_MAX = GRID_ROWS * BLOCKS; // 100本

// 成績表の見出し（A1）の文字サイズ: 1行目「高圧ガス容器検査成績表」と、「※容器の製造年月を…」の行を分けて設定
const TITLE_FONT_SIZE = 15;   // 高圧ガス容器検査成績表（元Excelと同じ）
const NOTE_FONT_SIZE = 11;    // ※容器の製造年月を目視確認の上、容器番号前にチェックを入れる

const MASS_MIN = 0.1;
const MASS_MAX = 99.9;

// ---------- Web app ----------

// 本番のプロジェクト（スプレッドシートにバインド）と入力画面のURL。
// 旧URL（別プロジェクト 1WRdps6z…）で開かれたら、入力はさせずに新しいURLへ案内する
const MAIN_SCRIPT_ID = '17mOYxTzPwzIsrILrYL7pH0Rn-jYOjQ8f52n5KZm0-0p-ojhxBG246vu4';
const MAIN_WEBAPP_URL = 'https://script.google.com/a/macros/nodagumi40.com/s/AKfycbyDM2IQ8NY6LgQpLb0gpQoK32fAQtWVBrdBQduZ6FOBWREFn61Qb55R9Q20Y4u9IfQ5/exec';

function isOldProject_() {
  try { return ScriptApp.getScriptId() !== MAIN_SCRIPT_ID; } catch (e) { return false; }
}

function movedPageHtml_() {
  return '<!doctype html><html><head><meta charset="utf-8">' +
    '<style>body{font-family:sans-serif;background:#f3f5fa;margin:0;padding:40px 20px;text-align:center;color:#1f2a5c}' +
    '.box{max-width:560px;margin:0 auto;background:#fff;border-radius:16px;padding:32px 24px;box-shadow:0 6px 24px rgba(20,26,60,.12)}' +
    'h1{font-size:22px;margin:0 0 12px}p{font-size:16px;line-height:1.7;margin:0 0 20px}' +
    'a.btn{display:inline-block;background:#1f2a5c;color:#fff;text-decoration:none;font-size:20px;font-weight:700;padding:16px 28px;border-radius:12px}' +
    'small{display:block;margin-top:20px;color:#6b7390;font-size:13px;line-height:1.6;word-break:break-all}</style></head><body>' +
    '<div class="box"><h1>このURLは使えなくなりました</h1>' +
    '<p>LPG容器 質量入力は新しいURLに移りました。<br>下のボタンで開いて、ホーム画面に追加し直してください。</p>' +
    '<a class="btn" href="' + MAIN_WEBAPP_URL + '" target="_top">新しい入力画面を開く</a>' +
    '<small>新しいURL: ' + MAIN_WEBAPP_URL + '</small></div></body></html>';
}

function doGet() {
  if (isOldProject_()) {
    return HtmlService.createHtmlOutput(movedPageHtml_()).setTitle('LPG容器 質量入力（URLが変わりました）')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('LPG容器 質量入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// シートの整備（タイムゾーン・列見出し・シート名・見出しの文字サイズ）。1つが失敗しても残りは必ず行う
function ensureSheets_() {
  [ensureTokyoTime_, ensureLogHeaders_, ensureTitleFont_, ensureSettingRows_, ensureTypicalDefault_, ensureVolumes_, ensureVolumes3_, ensureTrim_, cleanDoneSheets_, ensureLogDayLines_].forEach(function(fn) {
    try { fn(); } catch (e) { console.error('シートの整備に失敗: ' + (fn.name || '') + ' ' + e.message); }
  });
}

// 後から増えた設定項目（内容積など）を既存の設定シートに足す
function ensureSettingRows_() {
  const sh = getSpreadsheet_().getSheetByName(SHEET_SETTINGS);
  if (!sh) return;
  const have = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]).trim(); }) : [];
  SETTING_ROWS.forEach(function(r) { if (have.indexOf(r[0]) < 0) sh.appendRow(r); });
}

// ---------- シートを小さく保つ ----------
// 成績表は A1:AF32 に収まる。外側の空の行・列を削ってセル数（1ファイル1000万セルまで）を節約する
const TRIM_ROWS = 34, TRIM_COLS = 34;
function trimSheet_(sh) {
  const keepR = Math.max(sh.getLastRow(), TRIM_ROWS), keepC = Math.max(sh.getLastColumn(), TRIM_COLS);
  if (sh.getMaxRows() > keepR) sh.deleteRows(keepR + 1, sh.getMaxRows() - keepR);
  if (sh.getMaxColumns() > keepC) sh.deleteColumns(keepC + 1, sh.getMaxColumns() - keepC);
}
// 書式シートと今ある成績表シートを1度だけ小さくする
function ensureTrim_() {
  const props = sharedProps_();
  if (props.getProperty('SHEETS_TRIMMED') === '1') return;
  getSpreadsheet_().getSheets().forEach(function(sh) {
    const n = sh.getName();
    if (n === SHEET_TEMPLATE || n.indexOf(REPORT_PREFIX) === 0) trimSheet_(sh);
  });
  props.setProperty('SHEETS_TRIMMED', '1');
}
// 完了済み（PDF保存済み）のロットに残っている成績表シートを消す。1回に30枚まで（多いときは次に開いたときに続き）
function cleanDoneSheets_() {
  if (!readSettings_().removeDoneSheet) return;
  const ss = getSpreadsheet_();
  const names = {};
  ss.getSheets().forEach(function(sh) { names[sh.getName()] = sh; });
  let n = 0;
  readLots_().forEach(function(lot) {
    if (n >= 30 || lot.status !== STATUS_DONE || !lot.pdf || !names[lot.sheetName]) return;
    ss.deleteSheet(names[lot.sheetName]);
    n++;
  });
}

// 成績表の内容積（N3、「内容積：」と「lit」の間）
const VOLUME_CELL = 'N3';
function volumeFor_(kind, settings) {
  const v = (settings || readSettings_()).volume[String(kind || '').trim()];
  return v ? String(v) : '';
}
function setVolume_(sh, kind, settings) {
  const v = volumeFor_(kind, settings);
  const cell = sh.getRange(VOLUME_CELL);
  cell.setNumberFormat('@').setValue(v);
}
// v56までの標準質量の初期値
const TYPICAL_OLD_DEFAULTS = ['5kg=6.8, 8kg=9.6, 20kg=16.7, 30kg=24.0, 50kg=34.8, 50kg S付=36.3',
  '5kg=6.8, 8kg=9.6, 10kg=11.3, 20kg=16.7, 30kg=23.5, 30kg把手=24.0, 50kg=34.8, 50kg S付=36.3'];
// v34の仮の内容積（設定シートに入っていたら正式な値に置き換える）
const VOLUME_OLD_DEFAULT = '5kg=11.8, 8kg=18.8, 20kg=47.0, 30kg=70.5, 50kg=117.5, 50kg S付=117.5';
// 設定シートの標準質量が以前の初期値のままなら、今の初期値に置き換える（10kg・30kg把手・20kg三部制を足す。30kgは把手なしの23.5）
function ensureTypicalDefault_() {
  const conf = getSpreadsheet_().getSheetByName(SHEET_SETTINGS);
  if (!conf || conf.getLastRow() < 2) return;
  conf.getRange(2, 1, conf.getLastRow() - 1, 2).getValues().forEach(function(r, i) {
    if (String(r[0]).trim() === '標準質量' && TYPICAL_OLD_DEFAULTS.indexOf(String(r[1]).trim()) >= 0) {
      conf.getRange(2 + i, 2).setValue(SETTING_ROWS.filter(function(x) { return x[0] === '標準質量'; })[0][1]);
    }
  });
}
// 入力中ロットの成績表に内容積を1度だけ入れ直す（v34より前に作ったロット・仮の値が入ったロット用）
function ensureVolumes_() {
  const props = sharedProps_();
  if (props.getProperty('VOLUME_FILLED') === '2') return;
  const ss = getSpreadsheet_();
  const conf = ss.getSheetByName(SHEET_SETTINGS);
  if (conf && conf.getLastRow() >= 2) {
    conf.getRange(2, 1, conf.getLastRow() - 1, 2).getValues().forEach(function(r, i) {
      if (String(r[0]).trim() === '内容積' && String(r[1]).trim() === VOLUME_OLD_DEFAULT) {
        conf.getRange(2 + i, 2).setValue(SETTING_ROWS.filter(function(x) { return x[0] === '内容積'; })[0][1]);
      }
    });
  }
  const settings = readSettings_();
  const old = parseTypical_(VOLUME_OLD_DEFAULT);
  readLots_().forEach(function(lot) {
    if (lot.status !== STATUS_ACTIVE) return;
    const sh = ss.getSheetByName(lot.sheetName);
    if (!sh) return;
    const cur = String(sh.getRange(VOLUME_CELL).getValue());
    const wasOld = Object.keys(old).some(function(k) { return old[k].toFixed(1) === cur; });
    if (cur === '' || wasOld) setVolume_(sh, lot.kind, settings);
  });
  props.setProperty('VOLUME_FILLED', '2');
}

// 内容積が以前の初期値のままなら今の初期値（10kg・20kg三部制・30kg把手あり）に置き換え、入力中ロットの空欄の内容積を埋める（1回だけ）
const VOLUME_OLD_DEFAULTS = ['5kg=12, 8kg=19, 20kg=47, 30kg=71, 50kg=118, 50kg S付=118',
  '5kg=12, 8kg=19, 10kg=24, 20kg=47, 30kg=71, 30kg把手=71, 50kg=118, 50kg S付=118'];
function ensureVolumes3_() {
  const props = sharedProps_();
  if (props.getProperty('VOLUME_FILLED') === '4') return;
  const ss = getSpreadsheet_();
  const conf = ss.getSheetByName(SHEET_SETTINGS);
  if (conf && conf.getLastRow() >= 2) {
    conf.getRange(2, 1, conf.getLastRow() - 1, 2).getValues().forEach(function(r, i) {
      if (String(r[0]).trim() === '内容積' && VOLUME_OLD_DEFAULTS.indexOf(String(r[1]).trim()) >= 0) {
        conf.getRange(2 + i, 2).setValue(SETTING_ROWS.filter(function(x) { return x[0] === '内容積'; })[0][1]);
      }
    });
  }
  const settings = readSettings_();
  readLots_().forEach(function(lot) {
    if (lot.status !== STATUS_ACTIVE) return;
    const sh = ss.getSheetByName(lot.sheetName);
    if (sh && String(sh.getRange(VOLUME_CELL).getValue()) === '') setVolume_(sh, lot.kind, settings);
  });
  props.setProperty('VOLUME_FILLED', '4');
}

function onOpen() {
  ensureSheets_();
  SpreadsheetApp.getUi().createMenu('LPG容器検査')
    .addItem('初期セットアップ', 'setup')
    .addItem('入力画面のURLを表示', 'showWebAppUrl')
    .addSeparator()
    .addItem('表示中の成績表をPDF保存', 'exportActiveSheetPdf')
    .addSeparator()
    .addItem('毎晩の自動処理を設定（採番表PDFなど）', 'installNightlyTrigger')
    .addToUi();
}

// 記録先のスプレッドシート「LPG容器_質量入力」。ウェブアプリからは「開いているスプレッドシート」が
// 取れないため、IDで開く。スクリプトプロパティ SPREADSHEET_ID があればそちらを優先する
const DEFAULT_SPREADSHEET_ID = '1h0VM9ECv1NnSuwuo2jxTaiPVC6o5oYwWi1lZ9Qx1hms';

// 1回の実行の中では同じものを使い回す（openById は毎回だと遅い）
let SS_CACHE_ = null;
function getSpreadsheet_() {
  if (SS_CACHE_) return SS_CACHE_;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('記録先のスプレッドシートが見つかりません。Apps Script の「プロジェクトの設定 > スクリプト プロパティ」に ' +
      'SPREADSHEET_ID（スプレッドシートURLの /d/ と /edit の間の文字列）を追加してください。');
  }
  SS_CACHE_ = ss;
  return ss;
}

// スプレッドシートのタイムゾーンが日本時間でないと入力日時がずれて表示されるので、日本時間にそろえる
function ensureTokyoTime_() {
  const ss = getSpreadsheet_();
  if (ss.getSpreadsheetTimeZone() !== 'Asia/Tokyo') ss.setSpreadsheetTimeZone('Asia/Tokyo');
}

// 見出しセルを「タイトル15pt・※の行11pt（どちらも太字）」にする
function applyTitleStyle_(range) {
  const text = String(range.getValue());
  const cut = text.indexOf('※');
  if (cut < 0) { range.setFontSize(TITLE_FONT_SIZE); return; }
  const style = function(size) { return SpreadsheetApp.newTextStyle().setFontSize(size).setBold(true).build(); };
  range.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text)
    .setTextStyle(0, cut, style(TITLE_FONT_SIZE))
    .setTextStyle(cut, text.length, style(NOTE_FONT_SIZE))
    .build());
}

// 書式シートと作成済みの成績表の見出しをそろえる（サイズを変えたときに1回だけ全シートに適用）
function ensureTitleFont_() {
  const props = sharedProps_();
  const key = 'TITLE_FONT_SIZE_APPLIED';
  const want = TITLE_FONT_SIZE + '/' + NOTE_FONT_SIZE;
  if (props.getProperty(key) === want) return;
  getSpreadsheet_().getSheets().forEach(function(sh) {
    const n = sh.getName();
    if (n === SHEET_TEMPLATE || n.indexOf(REPORT_PREFIX) === 0) applyTitleStyle_(sh.getRange('A1'));
  });
  props.setProperty(key, want);
}

// 入力記録の質量(kg)・上書き前を小数1桁で表示（35 → 35.0）
function formatLogMass_(log, fromRow, rows) {
  log.getRange(fromRow, LC_MASS + 1, rows, 1).setNumberFormat('0.0');
  log.getRange(fromRow, LC_PREV + 1, rows, 1).setNumberFormat('0.0');
}

// v43までの並び（…容器番号, 表示番号, … 備考, グループNo, 容器区分）なら、
// グループNoを容器番号の前、容器区分を容器番号の後ろへ列ごと移す（書式もそのまま動く）
// held: 呼び出し元がすでにスクリプトロックを持っている（記録の書き込み中など）
function moveLogColumns_(log, held) {
  if (log.getRange(1, 4).getValue() !== '容器番号') return;
  const lock = held ? null : LockService.getScriptLock();
  if (lock) lock.waitLock(20000);
  try {
    // 同時に開いた端末と二重に動かさないよう、ロックを取ってから確かめ直す
    const head = log.getRange(1, 1, 1, Math.max(log.getLastColumn(), 18)).getValues()[0];
    if (head[3] !== '容器番号') return;
    if (head[16] !== 'グループNo') log.getRange(1, 17).setValue('グループNo');
    if (head[17] !== '容器区分') log.getRange(1, 18).setValue('容器区分');
    log.moveColumns(log.getRange(1, 17), 4);   // グループNo → D列（容器番号は E列へ）
    log.moveColumns(log.getRange(1, 18), 6);   // 容器区分 → F列（容器番号の後ろ）
    SpreadsheetApp.flush();
  } finally {
    if (lock) lock.releaseLock();
  }
}

// 入力記録の列見出しを最新にそろえる（列の追加・名前の変更。データの行はそのまま）
// あわせて、旧名「担当者」シートを「入力者」に名前を変える
function ensureLogHeaders_() {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(SHEET_LOG);
  if (log) {
    moveLogColumns_(log);
    const cur = log.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
    LOG_HEADERS.forEach(function(h, i) { if (cur[i] !== h) log.getRange(1, i + 1).setValue(h); });
    // これまでの記録の質量も小数1桁表示に（1回だけ）
    const props = sharedProps_();
    if (props.getProperty('LOG_MASS_FORMAT') !== '0.0' && log.getMaxRows() > 1) {
      formatLogMass_(log, 2, log.getMaxRows() - 1);
      props.setProperty('LOG_MASS_FORMAT', '0.0');
    }
    // これまでの記録にもグループNo・容器区分を入れる（1回だけ。v44で早見表のルールに合わせて入れ直し）
    if (props.getProperty('LOG_GROUP_FILLED') !== '4' && log.getLastRow() > 1) {
      const kinds = {};
      readLots_().forEach(function(l) { kinds[l.lotId] = l; });
      const n = log.getLastRow() - 1;
      const v = log.getRange(2, 1, n, LOG_HEADERS.length).getDisplayValues();
      const grp = log.getRange(2, LC_GROUP + 1, n, 1).getValues();
      const ks = log.getRange(2, LC_KINDSIZE + 1, n, 1).getValues();
      v.forEach(function(r, i) {
        const l = kinds[r[LC_LOT]];
        if (!l) return;
        grp[i] = [groupNoOf_(l.kind, String(r[LC_SERIAL]).slice(l.prefix.length))];
        ks[i] = [l.kind];
      });
      log.getRange(2, LC_GROUP + 1, n, 1).setValues(grp);
      log.getRange(2, LC_KINDSIZE + 1, n, 1).setValues(ks);
      props.setProperty('LOG_GROUP_FILLED', '4');
    }
  }
  const lotsSh = ss.getSheetByName(SHEET_LOTS);
  if (lotsSh) {
    const lc = lotsSh.getRange(1, 1, 1, LOT_HEADERS.length).getValues()[0];
    LOT_HEADERS.forEach(function(h, i) { if (lc[i] !== h) lotsSh.getRange(1, i + 1).setValue(h); });
  }
  const oldWorkers = ss.getSheetByName(SHEET_WORKERS_OLD);
  if (oldWorkers && !ss.getSheetByName(SHEET_WORKERS)) {
    oldWorkers.setName(SHEET_WORKERS);
    if (oldWorkers.getRange('A1').getValue() === '担当者名') oldWorkers.getRange('A1').setValue('入力者名');
  }
}

function getSheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。メニュー「LPG容器検査 > 初期セットアップ」を実行してください。');
  return sh;
}

function userEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// ---------- 純粋ロジック（テスト可能） ----------

// グループNo（容器1本ごと）: 5kg・8kg・10kg・20kg は50本ごと（番号取り早見表: 1〜50 → 1 … 49951〜50000 → 0、50001〜50050 → 1）
// 30kg・50kg系は100本ごと（59701〜59800 → 598）
function groupNoOf_(kind, serial) {
  const k = String(kind || '').trim(), n = Number(serial);
  if (!(n > 0)) return '';
  if (/^(30kg|50kg)/.test(k)) return Math.floor((n - 1) / 100) + 1;
  if (/^(5kg|8kg|10kg|20kg)/.test(k)) return (Math.floor((n - 1) / 50) + 1) % 1000;
  return '';
}

// 容器番号の下3桁（1000番目は "000"）。成績表の「容器番号」欄に表示する値。
function displayNumber_(serial) {
  return ('00' + (Number(serial) % 1000)).slice(-3);
}

function padSerial_(n, width) {
  let s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

// ロット内の位置(0..99) → 成績表のセル位置
function slotPosition_(index) {
  const block = Math.floor(index / GRID_ROWS);
  return {
    row: GRID_FIRST_ROW + (index % GRID_ROWS),
    numberCol: GRID_FIRST_COL + block * BLOCK_WIDTH
  };
}

// 質量 → [☑, 10の位, 1の位, ",", 小数1位]
function massCells_(mass) {
  if (mass === MISSING) return [false, '欠', '番', '', ''];
  if (mass === REPAIR) return [false, '修', '正', '', ''];
  if (mass === CIRCLE) return [true, '', CIRCLE, '', ''];
  if (mass === null || mass === '' || mass === undefined) return [false, '', '', ',', ''];
  const tenths = Math.round(Number(mass) * 10);
  const tens = Math.floor(tenths / 100);
  const ones = Math.floor(tenths / 10) % 10;
  return [true, tens === 0 ? '' : tens, ones, ',', tenths % 10];
}

// [☑, 10の位, 1の位, ",", 小数1位] → 質量 / '欠番' / null
function parseMassCells_(cells) {
  const d = cells[1], e = cells[2], g = cells[4];
  if (d === '欠' || e === '番') return MISSING;
  if (d === '修' || e === '正') return REPAIR;
  if (e === CIRCLE) return CIRCLE;
  if (e === '' || e === null) return null;
  const tenths = (Number(d) || 0) * 100 + Number(e) * 10 + (Number(g) || 0);
  return Math.round(tenths) / 10;
}

// テンキー入力の質量文字列を数値に。小数点なしの「68」は 6.8、「363」は 36.3 とみなす。
function parseMassInput_(text) {
  const s = String(text).trim();
  if (!/^\d+(\.\d?)?$|^\.\d$/.test(s)) return NaN;
  if (s.indexOf('.') >= 0) return Math.round(parseFloat(s) * 10) / 10;
  if (s.length === 1) return NaN; // 「7」だけは小数点の打ち忘れとして弾く
  return Number(s) / 10;
}

function isValidMass_(m) {
  return typeof m === 'number' && isFinite(m) && m >= MASS_MIN && m <= MASS_MAX &&
    Math.abs(Math.round(m * 10) - m * 10) < 1e-6;
}

// ---------- ロット ----------

function lotRowToObj_(r) {
  const o = {};
  LOT_HEADERS.forEach(function(h, i) { o[h] = r[i]; });
  const start = String(o['開始番号']);
  const end = String(o['終了番号']);
  return {
    lotId: String(o['ロットID']),
    prefix: String(o['記号'] || ''),
    start: start,
    end: end,
    kind: String(o['容器区分'] || ''),
    status: String(o['状態']),
    sheetName: String(o['成績表シート']),
    pressure: {
      repNumber: o['代表容器番号'] === '' ? '' : String(o['代表容器番号']),
      testDate: toIsoDate_(o['耐圧試験日']),
      totalExp: o['全増加(cm3)'],
      permExp: o['恒久増加(cm3)'],
      permRate: o['恒久増加率(%)']
    },
    pdf: String(o['PDF'] || ''),
    spec: String(o['仕様'] || ''),
    sentAt: String(o['送信日時'] || ''),
    sentTo: String(o['送信先'] || '')
  };
}

// 表示値「2026/07/22」「2026年7月22日」→「2026-07-22」
function toIsoDate_(text) {
  const m = String(text || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function readLots_() {
  const sh = getSheet_(SHEET_LOTS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, LOT_HEADERS.length).getDisplayValues()
    .map(function(r, i) {
      const lot = lotRowToObj_(r);
      lot.row = i + 2;
      return lot;
    })
    .filter(function(l) { return l.lotId; });
}

// ロットを1件探す。ロットシートが大きくなっても速いよう、A列を検索して見つかった行だけ読む
function findLot_(lotId) {
  const sh = getSheet_(SHEET_LOTS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const col = sh.getRange(2, 1, last - 1, 1);
    if (typeof col.createTextFinder === 'function') {
      const cell = col.createTextFinder(String(lotId)).matchEntireCell(true).findNext();
      if (cell) {
        const lot = lotRowToObj_(sh.getRange(cell.getRow(), 1, 1, LOT_HEADERS.length).getDisplayValues()[0]);
        lot.row = cell.getRow();
        if (lot.lotId === lotId) return lot;
      }
    }
  }
  const lots = readLots_();
  for (let i = 0; i < lots.length; i++) if (lots[i].lotId === lotId) return lots[i];
  throw new Error('ロットが見つかりません: ' + lotId);
}

function lotSize_(lot) {
  return Number(lot.end) - Number(lot.start) + 1;
}

function getBootstrap() {
  ensureSheets_();
  const lots = readLots_().filter(function(l) { return l.status === STATUS_ACTIVE; });
  const settings = readSettings_();
  const marks = readMarksAll_();
  return {
    user: userEmail_(),
    workers: readWorkers_(),
    mail: { to: settings.to, cc: settings.cc, auto: settings.auto },
    typical: settings.typical,
    lots: lots.map(function(l) { return withProgress_(l, marks); }),
    recentDone: readLots_().filter(function(l) { return l.status === STATUS_DONE; })
      .slice(-10).reverse().map(function(l) { return withProgress_(l, marks); })
  };
}

function withProgress_(lot, marksMap) {
  lot.entries = readProgress_(lot);
  const m = (marksMap || readMarksAll_())[lot.lotId] || {};
  lot.dups = m.dups || {};
  lot.seals = m.seals || {};
  delete lot.row;
  return lot;
}

// 成績表シートから現在の入力状況を読む: { "052023": 6.8, ... }
function readProgress_(lot) {
  const sh = getSpreadsheet_().getSheetByName(lot.sheetName);
  if (!sh) return progressFromLog_(lot);
  const values = sh.getRange(GRID_FIRST_ROW, GRID_FIRST_COL, GRID_ROWS, BLOCK_WIDTH * BLOCKS).getValues();
  const result = {};
  const size = lotSize_(lot);
  const width = lot.start.length;
  for (let i = 0; i < size; i++) {
    const r = i % GRID_ROWS;
    const c = Math.floor(i / GRID_ROWS) * BLOCK_WIDTH;
    const mass = parseMassCells_(values[r].slice(c + 1, c + 6));
    if (mass !== null) result[padSerial_(Number(lot.start) + i, width)] = mass;
  }
  return result;
}

// 成績表シートがない（完了して消した）ロットの入力状況を、入力記録の有効な行から組み立てる
// full: 入力記録を全部読む（再開・再送信）。省略時は直近3万行を1回の実行で使い回す（一覧表示用）
let LOG_TAIL_CACHE_ = null;
function progressFromLog_(lot, full) {
  const out = {};
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log) return out;
  let v;
  if (log.getLastRow() < 2) v = [];
  else if (full) v = log.getRange(2, 1, log.getLastRow() - 1, LOG_HEADERS.length).getValues();
  else {
    if (!LOG_TAIL_CACHE_) {
      const last = log.getLastRow(), from = Math.max(2, last - 30000);
      LOG_TAIL_CACHE_ = log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getValues();
    }
    v = LOG_TAIL_CACHE_;
  }
  let mine = v.filter(function(r) { return String(r[LC_LOT]) === lot.lotId; });
  // 年替わりで別のスプレッドシートへ移したロットは、そこから読む
  if (full && !mine.length) mine = archivedRowsForLot_(lot.lotId);
  mine.forEach(function(r) {
    if (String(r[LC_LOT]) !== lot.lotId || r[LC_STATUS] !== '有効') return;
    const serial = String(r[LC_SERIAL]).slice(lot.prefix.length);
    const kind = r[LC_KIND];
    if (kind === MISSING) out[serial] = MISSING;
    else if (kind === REPAIR) out[serial] = REPAIR;
    else if (kind === SPEC_SOKO) out[serial] = CIRCLE;
    else if (r[LC_MASS] !== '') out[serial] = Number(r[LC_MASS]);
  });
  return out;
}

// 消した成績表シートを、書式シートから作り直して入力記録の値を書き戻す（再開・再送信用）
function rebuildReportSheet_(lot) {
  const entries = progressFromLog_(lot, true);
  const sh = createReportSheet_(lot);
  Object.keys(entries).forEach(function(serial) {
    const pos = slotPosition_(Number(serial) - Number(lot.start));
    sh.getRange(pos.row, pos.numberCol + 1, 1, 5).setValues([massCells_(entries[serial])]);
  });
  return sh;
}

// 範囲をまとめて登録: 開始〜終了を 001〜100 区切り（紙の成績表と同じ100本単位）に分けてロットを作る
// input: { prefix, start, end, kind } → { created: [ロットID], skipped: [既にあったロットID] }
function createLots(input) {
  const start = String(input.start || '').trim();
  const end = String(input.end || '').trim();
  if (!/^\d{3,8}$/.test(start) || !/^\d{3,8}$/.test(end)) throw new Error('開始・終了は数字で入力してください');
  const s = Number(start), e = Number(end);
  if (e < s) throw new Error('終了が開始より小さくなっています');
  const width = Math.max(start.length, end.length);
  const blocks = lotBlocks_(s, e);
  if (blocks.length > 50) throw new Error('一度に作れるのは50ロット（5000本）までです（' + blocks.length + 'ロットになります）');
  const prefix = String(input.prefix || '').trim().toUpperCase();
  const created = [], skipped = [];
  blocks.forEach(function(b) {
    try {
      created.push(createLot({ prefix: prefix, start: padSerial_(b, width), end: padSerial_(b + LOT_MAX - 1, width), kind: input.kind, spec: input.spec }).lotId);
    } catch (err) {
      if (/登録済み/.test(err.message)) skipped.push(prefix + padSerial_(b, width));
      else throw new Error(err.message + '（' + created.length + 'ロット作成済み）');
    }
  });
  return { created: created, skipped: skipped };
}

// 範囲 s〜e を含む100本単位のロットの開始番号（xx01始まり）
function lotBlocks_(s, e) {
  const out = [];
  for (let b = s - ((s - 1) % LOT_MAX + LOT_MAX) % LOT_MAX; b <= e; b += LOT_MAX) out.push(b);
  return out;
}

function createLot(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const prefix = String(input.prefix || '').trim().toUpperCase();
    const start = String(input.start || '').trim();
    let end = String(input.end || '').trim();
    if (!/^\d{3,8}$/.test(start)) throw new Error('組容器番号（開始）は数字で入力してください');
    if (!end) end = padSerial_(Number(start) + LOT_MAX - 1, start.length);
    if (!/^\d{3,8}$/.test(end)) throw new Error('組容器番号（終了）は数字で入力してください');
    const size = Number(end) - Number(start) + 1;
    if (size < 1 || size > LOT_MAX) throw new Error('1枚の成績表は最大100本です（開始〜終了を確認してください）');
    if (!/^[A-Z0-9]*$/.test(prefix)) throw new Error('記号は英数字で入力してください');

    const spec = input.spec === SPEC_SOKO ? SPEC_SOKO : '';
    const lotId = prefix + start + (spec ? '-' + SPEC_SOKO : '');
    const lots = readLots_();
    lots.forEach(function(l) {
      if (l.lotId === lotId) throw new Error('この組容器番号は登録済みです: ' + lotId);
    });
    // 入力中のロットと番号が重なると、番号から記録先を決められない（通常と底黒で同じ番号など）
    lots.filter(function(l) { return l.status === STATUS_ACTIVE && l.prefix === prefix; }).forEach(function(l) {
      if (Number(start) <= Number(l.end) && Number(l.start) <= Number(end)) {
        throw new Error('入力中の「' + (l.spec ? l.spec + ' ' : '') + l.prefix + l.start + '〜' + l.end + '」と番号が重なります。先にそちらを完了してください');
      }
    });
    const sheetName = REPORT_PREFIX + prefix + start + '-' + end.slice(-Math.min(end.length, 5)) + (spec ? '_' + SPEC_SOKO : '');
    const lot = {
      lotId: lotId, prefix: prefix, start: start, end: end,
      kind: String(input.kind || ''), status: STATUS_ACTIVE, sheetName: sheetName, spec: spec,
      pressure: { repNumber: start, testDate: '', totalExp: '', permExp: '', permRate: '' }, pdf: ''
    };
    createReportSheet_(lot);
    getSheet_(SHEET_LOTS).appendRow([lotId, new Date(), prefix, "'" + start, "'" + end, lot.kind,
      STATUS_ACTIVE, sheetName, "'" + start, '', '', '', '', '', userEmail_(), '', '', spec]);
    return withProgress_(lot);
  } finally {
    lock.releaseLock();
  }
}

function createReportSheet_(lot) {
  const ss = getSpreadsheet_();
  const template = ss.getSheetByName(SHEET_TEMPLATE);
  if (!template) throw new Error('シート「' + SHEET_TEMPLATE + '」がありません。初期セットアップを実行してください。');
  if (ss.getSheetByName(lot.sheetName)) throw new Error('シート「' + lot.sheetName + '」が既にあります');
  const sh = template.copyTo(ss).setName(lot.sheetName);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(ss.getNumSheets());

  const width = lot.start.length;
  const endLabel = lot.prefix + lot.end;
  applyTitleStyle_(sh.getRange('A1'));
  sh.getRange('V2').setValue(lot.prefix + lot.start + ' ～ ' + endLabel + (lot.spec ? '（' + lot.spec + '）' : ''));
  sh.getRange('H3').setValue(lot.prefix);
  setVolume_(sh, lot.kind);

  const size = lotSize_(lot);
  const grid = [];
  for (let r = 0; r < GRID_ROWS; r++) {
    const row = [];
    for (let b = 0; b < BLOCKS; b++) {
      const i = b * GRID_ROWS + r;
      const label = i < size ? displayNumber_(padSerial_(Number(lot.start) + i, width)) : '';
      row.push(label);
      Array.prototype.push.apply(row, i < size ? massCells_(null) : ['', '', '', '', '']);
    }
    grid.push(row);
  }
  // 元Excelの質量欄は0〜9のプルダウン入力規則付き。「欠番」を書けるよう外してからチェックボックスを付け直す
  sh.getRange(GRID_FIRST_ROW, GRID_FIRST_COL, GRID_ROWS, BLOCK_WIDTH * BLOCKS).clearDataValidations();
  for (let b = 0; b < BLOCKS; b++) {
    const col = GRID_FIRST_COL + b * BLOCK_WIDTH;
    sh.getRange(GRID_FIRST_ROW, col, GRID_ROWS, 1).setNumberFormat('@');
    sh.getRange(GRID_FIRST_ROW, col + 1, GRID_ROWS, 1).insertCheckboxes();
  }
  sh.getRange(GRID_FIRST_ROW, GRID_FIRST_COL, GRID_ROWS, BLOCK_WIDTH * BLOCKS).setValues(grid);

  // 耐圧試験欄（代表容器番号は開始番号を初期値に）
  sh.getRange('H26').setNumberFormat('@').setValue(lot.start);
  sh.getRange('H27').clearContent();
  sh.getRange('T26:T28').clearContent();
  return sh;
}

// ---------- 質量入力 ----------

// payload: { lotId, serial, mass (number) | missing: true, ngInput, clientId }
function recordEntry(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    LAST_LOG_DAY_ = null;
    return recordOne_(payload);
  } finally {
    lock.releaseLock();
  }
}

// 送信待ちをまとめて保存する（通信1回で複数本）。途中でエラーになったらそこで止め、それまでの結果とエラーを返す
// → { results: [recordEntry と同じ結果…], error: 'メッセージ' | '', errorClientId }
function recordEntries(list) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const out = { results: [], error: '', errorClientId: '' };
  try {
    LAST_LOG_DAY_ = null;
    for (let i = 0; i < (list || []).length; i++) {
      try {
        out.results.push(recordOne_(list[i]));
      } catch (e) {
        out.error = e.message;
        out.errorClientId = list[i].clientId || '';
        break;
      }
    }
    return out;
  } finally {
    lock.releaseLock();
  }
}

let LOG_LAYOUT_CHECKED_ = false;

// ---------- 日付が変わる所に線（入力記録・採番表の一覧） ----------
const DAY_LINE_COLOR = '#1f2a5c';
function dayLine_(sh, row, width) {
  sh.getRange(row, 1, 1, width).setBorder(true, null, null, null, null, null, DAY_LINE_COLOR, SpreadsheetApp.BorderStyle.SOLID_THICK);
}
// 入力記録の最後の行の日付（1回の実行の中では覚えておく）
let LAST_LOG_DAY_ = null;
function lastLogDay_(log) {
  if (LAST_LOG_DAY_) return LAST_LOG_DAY_;
  const last = log.getLastRow();
  if (last < 2) return null;
  const t = log.getRange(last, LC_TIME + 1).getValue();
  LAST_LOG_DAY_ = t instanceof Date ? Utilities.formatDate(t, 'Asia/Tokyo', 'yyyy-MM-dd') : null;
  return LAST_LOG_DAY_;
}
// これまでの入力記録にも日付の変わり目の線を1回だけ引く
function ensureLogDayLines_() {
  const props = sharedProps_();
  if (props.getProperty('LOG_DAY_LINES') === '1') return;
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log) return;
  if (log.getLastRow() >= 3) {
    const t = log.getRange(2, LC_TIME + 1, log.getLastRow() - 1, 1).getValues();
    let prev = null;
    t.forEach(function(r, i) {
      if (!(r[0] instanceof Date)) return;
      const d = Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd');
      if (prev && d !== prev) dayLine_(log, 2 + i, LOG_HEADERS.length);
      prev = d;
    });
  }
  props.setProperty('LOG_DAY_LINES', '1');
}
function recordOne_(payload) {
  {
    const lot = findLot_(payload.lotId);
    if (lot.status !== STATUS_ACTIVE) throw new Error('このロットは完了済みです: ' + lot.lotId);
    const serial = String(payload.serial);
    const index = Number(serial) - Number(lot.start);
    if (!(index >= 0 && index < lotSize_(lot))) throw new Error('容器番号がロットの範囲外です: ' + serial);

    let mass;
    if (lot.spec === SPEC_SOKO) {
      mass = CIRCLE;
    } else if (payload.missing) {
      mass = MISSING;
    } else if (payload.repair) {
      mass = REPAIR;
    } else {
      mass = Math.round(Number(payload.mass) * 10) / 10;
      if (!isValidMass_(mass)) throw new Error('質量は' + MASS_MIN + '〜' + MASS_MAX + 'の範囲で入力してください');
    }

    const sh = getSheet_(lot.sheetName);
    const pos = slotPosition_(index);
    const range = sh.getRange(pos.row, pos.numberCol + 1, 1, 5);
    const prev = parseMassCells_(range.getValues()[0]);
    // 底黒: 〇が付いている容器がもう一度流れても二重には記録しない
    if (mass === CIRCLE && prev === CIRCLE) return { recordId: null, clientId: payload.clientId, serial: serial, mass: mass, prev: prev, already: true };
    range.setValues([massCells_(mass)]);

    const recordId = 'R' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMddHHmmss') +
      Math.floor(Math.random() * 1000);
    // 区分: 新規=通常 / 欠番 / 入力済みの上書き=修正（入力ミスを直した）かダブり（2回流れた）
    // 別の端末が先に同じ容器を入れていた場合など、理由の指定がない上書きはダブりとする
    let kind = '通常';
    if (prev !== null) kind = payload.overwrite === 'fix' ? KIND_FIX : KIND_DUP;
    else if (mass === MISSING) kind = MISSING;
    else if (mass === REPAIR) kind = REPAIR;
    if (mass === CIRCLE) kind = SPEC_SOKO;
    const note = payload.seal ? NOTE_SEAL : '';
    const logSh = getSheet_(SHEET_LOG);
    if (!LOG_LAYOUT_CHECKED_) { moveLogColumns_(logSh, true); LOG_LAYOUT_CHECKED_ = true; }   // 古い並びのまま新しい並びの行を書かないように
    const row = {
      '記録ID': recordId, '入力日時': new Date(), 'ロットID': lot.lotId, 'グループNo': groupNoOf_(lot.kind, serial),
      '容器番号': "'" + lot.prefix + serial, '容器区分': lot.kind || '', '表示番号': "'" + displayNumber_(serial),
      '質量(kg)': typeof mass === 'number' ? mass : '', '区分': kind,
      '一致結果': payload.ngInput ? 'NG' : 'OK', '入力番号(NG時)': payload.ngInput ? "'" + payload.ngInput : '',
      '上書き前': prev === null ? '' : prev, 'Googleアカウント': userEmail_(), '状態': '有効',
      '入力者': String(payload.worker || ''), '端末': String(payload.device || ''), '取消日時': '', '備考': note
    };
    const prevDay = lastLogDay_(logSh);
    logSh.appendRow(LOG_HEADERS.map(function(h) { return row[h]; }));
    formatLogMass_(logSh, logSh.getLastRow(), 1);
    const today = Utilities.formatDate(row['入力日時'], 'Asia/Tokyo', 'yyyy-MM-dd');
    if (prevDay && prevDay !== today) dayLine_(logSh, logSh.getLastRow(), LOG_HEADERS.length);
    LAST_LOG_DAY_ = today;
    const result = { recordId: recordId, clientId: payload.clientId, serial: serial, mass: mass, prev: prev, kind: kind, note: note };

    // 全数（欠番を含む）そろったら自動で完了・PDF作成・送信する
    if (prev === null && countFilled_(lot) >= lotSize_(lot)) {
      result.full = true;
      // 自動で完了・送信するのは、送付先が入っていて「自動送信=する」のときだけ（送付先が空なら完了は手動で）
      const st = readSettings_();
      if (!lot.sentAt && st.auto && st.to) {
        try {
          const done = finishLot_(lot, true);
          result.completed = true;
          result.sentTo = done.sentTo;
          result.mailError = done.mailError;
        } catch (e) {
          result.mailError = e.message;
        }
      }
    }
    return result;
  }
}

// 直前の入力を取り消し、上書き前の値に戻す
function undoEntry(recordId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = getSheet_(SHEET_LOG);
    moveLogColumns_(log, true);
    const last = log.getLastRow();
    const from = Math.max(2, last - 500);
    if (last < 2) throw new Error('取り消す記録がありません');
    const values = log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getDisplayValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][LC_ID] !== recordId) continue;
      if (values[i][LC_STATUS] !== '有効') throw new Error('既に取り消し済みです');
      const lot = findLot_(values[i][LC_LOT]);
      const serial = values[i][LC_SERIAL].slice(lot.prefix.length);
      const prevText = values[i][LC_PREV];
      const prev = prevText === '' ? null : (prevText === MISSING || prevText === REPAIR || prevText === CIRCLE ? prevText : Number(prevText));
      const pos = slotPosition_(Number(serial) - Number(lot.start));
      getSheet_(lot.sheetName).getRange(pos.row, pos.numberCol + 1, 1, 5).setValues([massCells_(prev)]);
      log.getRange(from + i, LC_STATUS + 1).setValue('取消');
      log.getRange(from + i, LC_UNDO + 1).setValue(new Date());
      return { lotId: lot.lotId, serial: serial, mass: prev };
    }
    throw new Error('記録が見つかりません: ' + recordId);
  } finally {
    lock.releaseLock();
  }
}

// 入力ミスを消す: その容器のマスを空に戻し、入力記録の有効な行をすべて「取消」にする（行は履歴として残す）
function clearEntry(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lot = findLot_(input.lotId);
    const serial = String(input.serial);
    const idx = Number(serial) - Number(lot.start);
    if (!(idx >= 0 && idx <= Number(lot.end) - Number(lot.start))) throw new Error('このロットの番号ではありません: ' + serial);
    const pos = slotPosition_(idx);
    getSheet_(lot.sheetName).getRange(pos.row, pos.numberCol + 1, 1, 5).setValues([massCells_(null)]);
    const log = getSheet_(SHEET_LOG);
    moveLogColumns_(log, true);
    const label = lot.prefix + serial;
    let marked = 0;
    if (log.getLastRow() >= 2) {
      const last = log.getLastRow(), from = Math.max(2, last - 5000);
      const v = log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getDisplayValues();
      const now = new Date();
      v.forEach(function(r, i) {
        if (r[LC_LOT] === lot.lotId && r[LC_SERIAL] === label && r[LC_STATUS] === '有効') {
          log.getRange(from + i, LC_STATUS + 1).setValue('取消');
          log.getRange(from + i, LC_UNDO + 1).setValue(now);
          marked++;
        }
      });
    }
    return { lotId: lot.lotId, serial: serial, marked: marked };
  } finally {
    lock.releaseLock();
  }
}

// 印（ダブり・シール違い）の数: { ロットID: { dups: {容器番号: 回数}, seals: {...} } }
// 入力記録の直近5000行の有効な行から数える（区分が「通常」のままの上書きは以前の記録なのでダブり扱い）
function readMarksAll_() {
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) return {};
  const last = log.getLastRow(), from = Math.max(2, last - 5000);
  const width = Math.min(log.getLastColumn(), LOG_HEADERS.length);
  const out = {};
  const bump = function(lotId, key, serial) {
    const m = out[lotId] = out[lotId] || { dups: {}, seals: {} };
    m[key][serial] = (m[key][serial] || 0) + 1;
  };
  log.getRange(from, 1, last - from + 1, width).getDisplayValues().forEach(function(r) {
    if (r[LC_STATUS] !== '有効') return;
    if (r[LC_KIND] === KIND_DUP || (r[LC_PREV] !== '' && r[LC_KIND] === '通常')) bump(r[LC_LOT], 'dups', r[LC_SERIAL]);
    if (r[LC_NOTE] === NOTE_SEAL) bump(r[LC_LOT], 'seals', r[LC_SERIAL]);
  });
  return out;
}

function getLot(lotId) {
  return withProgress_(findLot_(lotId));
}

// ---------- 完了・送信 ----------

function countFilled_(lot) {
  return Object.keys(readProgress_(lot)).length;
}

// 「完了にする」ボタン（未入力が残っていても完了にする）。input: { lotId, send }
function completeLot(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lot = findLot_(input.lotId);
    return finishLot_(lot, !!input.send);
  } finally {
    lock.releaseLock();
  }
}

// 完了にしてPDFを作り、送信設定があればメールで納品する
function finishLot_(lot, send) {
  const lots = getSheet_(SHEET_LOTS);
  lots.getRange(lot.row, 7).setValue(STATUS_DONE);
  lot.status = STATUS_DONE;
  const sh = getSpreadsheet_().getSheetByName(lot.sheetName) || rebuildReportSheet_(lot);
  const pdf = exportLotPdf_(lot, sh);
  const result = { lotId: lot.lotId, pdfUrl: pdf.url, sentTo: '', mailError: '', sheetRemoved: false };
  if (send) {
    try {
      result.sentTo = sendLotMail_(lot, pdf.blob);
    } catch (e) {
      result.mailError = e.message;
      lots.getRange(lot.row, 17).setValue('送信失敗: ' + e.message);
    }
  }
  // PDFをドライブに保存できたら成績表シートは消す（再開すると入力記録から作り直す）
  if (pdf.url && readSettings_().removeDoneSheet) {
    getSpreadsheet_().deleteSheet(sh);
    result.sheetRemoved = true;
  }
  return result;
}

// 完了済みロットをもう一度送る
function resendLot(lotId) {
  const lot = findLot_(lotId);
  const sh = getSpreadsheet_().getSheetByName(lot.sheetName);
  if (sh) {
    const pdf = exportLotPdf_(lot, sh);
    return { lotId: lot.lotId, pdfUrl: pdf.url, sentTo: sendLotMail_(lot, pdf.blob) };
  }
  // 成績表シートを消したロットは、保存済みのPDFをそのまま送る（なければ作り直してPDFにし、また消す）
  const saved = savedPdfBlob_(lot);
  if (saved) return { lotId: lot.lotId, pdfUrl: lot.pdf, sentTo: sendLotMail_(lot, saved) };
  const tmp = rebuildReportSheet_(lot);
  const pdf = exportLotPdf_(lot, tmp);
  if (readSettings_().removeDoneSheet) getSpreadsheet_().deleteSheet(tmp);
  return { lotId: lot.lotId, pdfUrl: pdf.url, sentTo: sendLotMail_(lot, pdf.blob) };
}

function savedPdfBlob_(lot) {
  const m = String(lot.pdf || '').match(/\/d\/([\w-]{10,})/) || String(lot.pdf || '').match(/[?&]id=([\w-]{10,})/);
  if (!m) return null;
  try { return DriveApp.getFileById(m[1]).getBlob().setName(lot.sheetName + '.pdf'); } catch (e) { return null; }
}

// ロットの容器区分（機種）を変える
function changeLotKind(input) {
  const kind = String(input.kind || '').trim();
  if (!kind) throw new Error('容器区分を選んでください');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lot = findLot_(input.lotId);
    getSheet_(SHEET_LOTS).getRange(lot.row, 6).setValue(kind);
    const sh = getSpreadsheet_().getSheetByName(lot.sheetName);
    if (sh) setVolume_(sh, kind);
    return { lotId: lot.lotId, kind: kind };
  } finally {
    lock.releaseLock();
  }
}

// 登録間違いのロットを削除: ロットの行と成績表シートを消す。入力記録は消さずに状態を「削除」にして残す
function deleteLot(lotId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lot = findLot_(lotId);
    if (lot.status === STATUS_DONE) throw new Error('完了したロットは削除できません（成績表を保存済み）: ' + lot.lotId);
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(lot.sheetName);
    if (sh) ss.deleteSheet(sh);
    getSheet_(SHEET_LOTS).deleteRow(lot.row);
    let marked = 0;
    const log = ss.getSheetByName(SHEET_LOG);
    if (log) moveLogColumns_(log, true);
    if (log && log.getLastRow() >= 2) {
      const last = log.getLastRow(), from = Math.max(2, last - 5000);
      const v = log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getValues();
      v.forEach(function(r, i) {
        if (r[LC_LOT] === lot.lotId && r[LC_STATUS] === '有効') { log.getRange(from + i, LC_STATUS + 1).setValue('削除'); marked++; }
      });
    }
    return { lotId: lot.lotId, marked: marked };
  } finally {
    lock.releaseLock();
  }
}

// 削除してしまったロットを入力記録から戻す（ロット行を作り直し、入力記録の「削除」を「有効」に戻す）
// 保存済みの成績表PDFがドライブにあれば「完了」で、なければ「入力中」で戻して成績表シートを作り直す
function undeleteLot(lotId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    try { findLot_(lotId); return { lotId: lotId, already: true }; } catch (e) { /* 無いので戻す */ }
    const log = getSheet_(SHEET_LOG);
    if (log.getLastRow() < 2) throw new Error('入力記録にありません: ' + lotId);
    const v = log.getRange(2, 1, log.getLastRow() - 1, LOG_HEADERS.length).getValues();
    const idx = [];
    v.forEach(function(r, i) { if (String(r[LC_LOT]) === lotId && r[LC_STATUS] === '削除') idx.push(i); });
    if (!idx.length) throw new Error('削除した記録が見つかりません: ' + lotId);
    const m = String(lotId).match(/^(.*?)(\d+)(-底黒)?$/);
    if (!m) throw new Error('ロットIDの形が違います: ' + lotId);
    const prefix = m[1], start = m[2], spec = m[3] ? SPEC_SOKO : '';
    const end = padSerial_(Number(start) + LOT_MAX - 1, start.length);
    const sheetName = REPORT_PREFIX + prefix + start + '-' + end + (spec ? '_' + SPEC_SOKO : '');
    const kind = String(v[idx[0]][LC_KINDSIZE] || '');
    let first = v[idx[0]][LC_TIME];
    idx.forEach(function(i) { if (v[i][LC_TIME] instanceof Date && v[i][LC_TIME] < first) first = v[i][LC_TIME]; });
    let pdfUrl = '', newest = null;
    const files = DriveApp.getFilesByName(sheetName + '.pdf');
    while (files.hasNext()) {
      const f = files.next();
      if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
    }
    if (newest) pdfUrl = newest.getUrl();
    const status = pdfUrl ? STATUS_DONE : STATUS_ACTIVE;
    getSheet_(SHEET_LOTS).appendRow([lotId, first instanceof Date ? first : new Date(), prefix, "'" + start, "'" + end, kind,
      status, sheetName, "'" + start, '', '', '', '', pdfUrl, userEmail_(), '', '', spec]);
    idx.forEach(function(i) { log.getRange(2 + i, LC_STATUS + 1).setValue('有効'); });
    const lot = findLot_(lotId);
    if (status === STATUS_ACTIVE && !getSpreadsheet_().getSheetByName(sheetName)) rebuildReportSheet_(lot);
    return { lotId: lotId, status: status, restored: idx.length, pdfUrl: pdfUrl };
  } finally {
    lock.releaseLock();
  }
}

function reopenLot(lotId) {
  const lot = findLot_(lotId);
  getSheet_(SHEET_LOTS).getRange(lot.row, 7).setValue(STATUS_ACTIVE);
  lot.status = STATUS_ACTIVE;
  restoreArchivedLot_(lot);   // 年替わりで移した記録があれば入力記録へ戻す
  if (!getSpreadsheet_().getSheetByName(lot.sheetName)) rebuildReportSheet_(lot);
  return withProgress_(lot);
}

// ---------- PDF ----------

function exportLotPdf_(lot, sh) {
  SpreadsheetApp.flush();
  const ss = getSpreadsheet_();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
    '&gid=' + sh.getSheetId() +
    '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false' +
    '&pagenum=UNDEFINED&horizontal_alignment=CENTER' +
    '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4' +
    '&r1=0&c1=0&r2=30&c2=31';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  const blob = res.getBlob().setName(lot.sheetName + '.pdf');
  const file = pdfKindFolder_(getPdfMonthFolder_(new Date()), lot.kind).createFile(blob);
  if (lot.row) getSheet_(SHEET_LOTS).getRange(lot.row, 14).setValue(file.getUrl());
  return { url: file.getUrl(), blob: blob };
}

// ---------- 共有の設定値（今のURLと古いURLのプロジェクトで同じ値を使う） ----------
// スクリプトプロパティはプロジェクトごとに別なので、保存先フォルダのIDや「1回だけの処理」の印は
// 記録用スプレッドシートの隠しシート「_システム」（A列=キー、B列=値）に持つ。
// まだ無いキーは、そのプロジェクトのスクリプトプロパティの値を引き継ぐ
const SHEET_SYSTEM = '_システム';
let SHARED_CACHE_ = null;
function sharedProps_() {
  const load = function() {
    if (SHARED_CACHE_) return SHARED_CACHE_;
    const ss = getSpreadsheet_();
    let sh = ss.getSheetByName(SHEET_SYSTEM);
    if (!sh) {
      sh = ss.insertSheet(SHEET_SYSTEM);
      sh.getRange(1, 1, 1, 2).setValues([['キー', '値']]);
      try { sh.hideSheet(); } catch (e) { /* 隠せなくても動く */ }
    }
    const map = {}, rows = {};
    if (sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function(r, i) {
        if (r[0] !== '') { map[String(r[0])] = String(r[1]); rows[String(r[0])] = i + 2; }
      });
    }
    SHARED_CACHE_ = { sh: sh, map: map, rows: rows };
    return SHARED_CACHE_;
  };
  const legacy = PropertiesService.getScriptProperties();
  return {
    getProperty: function(k) {
      const c = load();
      if (k in c.map) return c.map[k] === '' ? null : c.map[k];
      return legacy.getProperty(k);
    },
    setProperty: function(k, v) {
      const c = load();
      v = String(v);
      if (c.rows[k]) c.sh.getRange(c.rows[k], 2).setValue("'" + v);
      else { c.sh.appendRow([k, "'" + v]); c.rows[k] = c.sh.getLastRow(); }
      c.map[k] = v;
    },
    deleteProperty: function(k) {
      const c = load();
      if (c.rows[k]) c.sh.getRange(c.rows[k], 2).setValue('');
      c.map[k] = '';
      try { legacy.deleteProperty(k); } catch (e) { /* 無視 */ }
    }
  };
}

// ---------- ドライブのフォルダ（全部ここにまとめる） ----------
// LPG容器 質量入力/
//   LPG容器_質量入力（このスプレッドシート）
//   成績表PDF/2026年/09月/50kg/   完了したロットの成績表PDF（月の中を機種ごとに分ける）
//   採番表/2026年/09月/      採番表PDF
//   採番表/LPG容器 採番表_2026   採番表の一覧（スプレッドシート、年ごと）
//   入力記録（過去分）/LPG容器 入力記録_2026   年替わりに移した入力記録
const ROOT_FOLDER_NAME = 'LPG容器 質量入力';
const REPORT_PDF_FOLDER = '成績表PDF';
const ARCHIVE_FOLDER = '入力記録（過去分）';

// まとめ先のフォルダ（共有の PDF_FOLDER_ID）。未設定なら、同じ名前のフォルダのうち一番古いものを使う
function getPdfFolder_() {
  const props = sharedProps_();
  const id = props.getProperty('PDF_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 削除されていたら探し直す */ }
  }
  const found = rootCandidates_()[0];
  if (found) { props.setProperty('PDF_FOLDER_ID', found.getId()); return found; }
  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('PDF_FOLDER_ID', folder.getId());
  props.setProperty('FOLDER_LAYOUT', '1');
  try { DriveApp.getFileById(getSpreadsheet_().getId()).moveTo(folder); } catch (e) { /* 移せなくても動く */ }
  return folder;
}

// 「LPG容器 質量入力」（と旧名「LPG容器 検査成績表PDF」）のフォルダを古い順に
function rootCandidates_() {
  const out = [];
  [ROOT_FOLDER_NAME, 'LPG容器 検査成績表PDF'].forEach(function(n) {
    const it = DriveApp.getFoldersByName(n);
    while (it.hasNext()) out.push(it.next());
  });
  return out.sort(function(a, b) { return a.getDateCreated() - b.getDateCreated(); });
}

// 同じ名前のまとめ先フォルダが2つ以上できていたら、一番古いものに中身を集めて、空になった方はゴミ箱へ（1回だけ）
function mergeRootFolders_() {
  const props = sharedProps_();
  if (props.getProperty('ROOTS_MERGED') === '1') return 0;
  const cands = rootCandidates_();
  const root = cands[0] || getPdfFolder_();   // 一番古いものに集める（どちらのプロジェクトで動いても同じ結果）
  props.setProperty('PDF_FOLDER_ID', root.getId());
  let merged = 0;
  cands.forEach(function(f) {
    if (f.getId() === root.getId()) return;
    mergeFolderInto_(f, root);
    f.setTrashed(true);
    merged++;
  });
  try { DriveApp.getFileById(getSpreadsheet_().getId()).moveTo(root); } catch (e) { /* 移せなくても動く */ }
  props.setProperty('ROOTS_MERGED', '1');
  return merged;
}

// from の中身を into へ移す。同じ名前のフォルダは中身を合わせ、同じ名前のファイルは新しい方を残す
function mergeFolderInto_(from, into) {
  const subs = from.getFolders(), folders = [];
  while (subs.hasNext()) folders.push(subs.next());
  folders.forEach(function(f) {
    const same = into.getFoldersByName(f.getName());
    if (same.hasNext()) { mergeFolderInto_(f, same.next()); f.setTrashed(true); }
    else f.moveTo(into);
  });
  const it = from.getFiles(), files = [];
  while (it.hasNext()) files.push(it.next());
  files.forEach(function(file) {
    const same = into.getFilesByName(file.getName());
    if (same.hasNext()) {
      const other = same.next();
      if (other.getLastUpdated() >= file.getLastUpdated()) { file.setTrashed(true); return; }
      other.setTrashed(true);
    }
    file.moveTo(into);
  });
}

// v52までの並び（成績表PDFの年フォルダが直下）を、上の並びに1回だけ整理する
function ensureFolderLayout_() {
  const props = sharedProps_();
  if (props.getProperty('FOLDER_LAYOUT') === '1') return false;
  if (!props.getProperty('PDF_FOLDER_ID')) { getPdfFolder_(); return true; }
  const root = getPdfFolder_();
  if (root.getName() !== ROOT_FOLDER_NAME) root.setName(ROOT_FOLDER_NAME);
  const pdfRoot = childFolder_(root, REPORT_PDF_FOLDER);
  const saiban = childFolder_(root, SAIBAN_FOLDER);
  const archive = childFolder_(root, ARCHIVE_FOLDER);
  const folders = root.getFolders();
  const years = [];
  while (folders.hasNext()) { const f = folders.next(); if (/^\d{4}年$/.test(f.getName())) years.push(f); }
  years.forEach(function(f) {
    const same = pdfRoot.getFoldersByName(f.getName());
    if (!same.hasNext()) { f.moveTo(pdfRoot); return; }
    // 同じ年のフォルダがもうあれば中身を移す
    const dest = same.next(), subs = f.getFolders();
    while (subs.hasNext()) {
      const m = subs.next(), d2 = dest.getFoldersByName(m.getName());
      if (!d2.hasNext()) { m.moveTo(dest); continue; }
      const into = d2.next(), fs = m.getFiles();
      while (fs.hasNext()) fs.next().moveTo(into);
    }
  });
  const files = root.getFiles(), loose = [];
  while (files.hasNext()) loose.push(files.next());
  loose.forEach(function(f) {
    const n = f.getName();
    if (n.indexOf(ARCHIVE_PREFIX) === 0) f.moveTo(archive);
    else if (n.indexOf(SAIBAN_SS_PREFIX) === 0) f.moveTo(saiban);
    else if (/\.pdf$/i.test(n)) f.moveTo(pdfRoot);
  });
  try { DriveApp.getFileById(getSpreadsheet_().getId()).moveTo(root); } catch (e) { /* 移せなくても動く */ }
  props.setProperty('FOLDER_LAYOUT', '1');
  return true;
}

// 保存先: 「LPG容器 質量入力 / 成績表PDF / 2026年 / 09月」のように年・月のフォルダを自動で作って入れる
function getPdfMonthFolder_(date) {
  const tz = 'Asia/Tokyo';
  const year = childFolder_(childFolder_(getPdfFolder_(), REPORT_PDF_FOLDER), Utilities.formatDate(date, tz, 'yyyy') + '年');
  return childFolder_(year, Utilities.formatDate(date, tz, 'MM') + '月');
}
// 月フォルダの中を機種（容器区分）ごとに分ける: 成績表PDF/2026年/09月/50kg/
function pdfKindFolder_(monthFolder, kind) {
  return childFolder_(monthFolder, String(kind || '').trim() || 'その他');
}

// v53までに月フォルダの直下に保存した成績表PDFを、機種のフォルダへ1回だけ移す
function ensurePdfKindFolders_() {
  const props = sharedProps_();
  if (props.getProperty('PDF_KIND_FOLDERS') === '1') return;
  readLots_().forEach(function(lot) {
    const m = String(lot.pdf || '').match(/\/d\/([\w-]{10,})/);
    if (!m) return;
    try {
      const file = DriveApp.getFileById(m[1]);
      const parents = file.getParents();
      if (!parents.hasNext()) return;
      const parent = parents.next();
      if (/^\d{2}月$/.test(parent.getName())) file.moveTo(pdfKindFolder_(parent, lot.kind));
    } catch (e) { /* 消されたPDFは飛ばす */ }
  });
  props.setProperty('PDF_KIND_FOLDERS', '1');
}

function findSpreadsheetIn_(folder, name) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    try { return SpreadsheetApp.openById(it.next().getId()); } catch (e) { /* 開けないものは飛ばす */ }
  }
  return null;
}

function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function exportActiveSheetPdf() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const lot = readLots_().filter(function(l) { return l.sheetName === sh.getName(); })[0];
  if (!lot) {
    ui.alert('成績表シート（' + REPORT_PREFIX + '〜）を開いてから実行してください');
    return;
  }
  ui.alert('PDFを保存しました:\n' + exportLotPdf_(lot, sh).url);
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? '入力画面のURL:\n' + url
    : 'まだウェブアプリとしてデプロイされていません。README の手順でデプロイしてください。');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
