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
const LOG_HEADERS = ['記録ID', '入力日時', 'ロットID', '容器番号', '表示番号', '質量(kg)', '区分',
  '一致結果', '入力番号(NG時)', '上書き前', 'Googleアカウント', '状態', '入力者', '端末', '取消日時', '備考',
  'グループNo', '容器区分'];

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

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('LPG容器 質量入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// シートの整備（タイムゾーン・列見出し・シート名・見出しの文字サイズ）。1つが失敗しても残りは必ず行う
function ensureSheets_() {
  [ensureTokyoTime_, ensureLogHeaders_, ensureTitleFont_, ensureSettingRows_, ensureVolumes_].forEach(function(fn) {
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
// v34の仮の初期値（設定シートに入っていたら正式な値に置き換える）
const VOLUME_OLD_DEFAULT = '5kg=11.8, 8kg=18.8, 20kg=47.0, 30kg=70.5, 50kg=117.5, 50kg S付=117.5';
// 入力中ロットの成績表に内容積を1度だけ入れ直す（v34より前に作ったロット・仮の値が入ったロット用）
function ensureVolumes_() {
  const props = PropertiesService.getScriptProperties();
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

function onOpen() {
  ensureSheets_();
  SpreadsheetApp.getUi().createMenu('LPG容器検査')
    .addItem('初期セットアップ', 'setup')
    .addItem('入力画面のURLを表示', 'showWebAppUrl')
    .addSeparator()
    .addItem('表示中の成績表をPDF保存', 'exportActiveSheetPdf')
    .addToUi();
}

// 記録先のスプレッドシート「LPG容器_質量入力」。ウェブアプリからは「開いているスプレッドシート」が
// 取れないため、IDで開く。スクリプトプロパティ SPREADSHEET_ID があればそちらを優先する
const DEFAULT_SPREADSHEET_ID = '1h0VM9ECv1NnSuwuo2jxTaiPVC6o5oYwWi1lZ9Qx1hms';

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('記録先のスプレッドシートが見つかりません。Apps Script の「プロジェクトの設定 > スクリプト プロパティ」に ' +
      'SPREADSHEET_ID（スプレッドシートURLの /d/ と /edit の間の文字列）を追加してください。');
  }
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
  const props = PropertiesService.getScriptProperties();
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
  log.getRange(fromRow, 6, rows, 1).setNumberFormat('0.0');
  log.getRange(fromRow, 10, rows, 1).setNumberFormat('0.0');
}

// 入力記録の列見出しを最新にそろえる（列の追加・名前の変更。データの行はそのまま）
// あわせて、旧名「担当者」シートを「入力者」に名前を変える
function ensureLogHeaders_() {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(SHEET_LOG);
  if (log) {
    const cur = log.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
    LOG_HEADERS.forEach(function(h, i) { if (cur[i] !== h) log.getRange(1, i + 1).setValue(h); });
    // これまでの記録の質量も小数1桁表示に（1回だけ）
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('LOG_MASS_FORMAT') !== '0.0' && log.getMaxRows() > 1) {
      formatLogMass_(log, 2, log.getMaxRows() - 1);
      props.setProperty('LOG_MASS_FORMAT', '0.0');
    }
    // これまでの記録にもグループNo・容器区分を入れる（1回だけ）
    if (props.getProperty('LOG_GROUP_FILLED') !== '1' && log.getLastRow() > 1) {
      const kinds = {};
      readLots_().forEach(function(l) { kinds[l.lotId] = l; });
      const n = log.getLastRow() - 1;
      const ids = log.getRange(2, 3, n, 2).getDisplayValues();
      const out = log.getRange(2, 17, n, 2).getValues();
      ids.forEach(function(r, i) {
        const l = kinds[r[0]];
        if (!l || out[i][1] !== '') return;
        out[i] = [groupNoOf_(l.kind, String(r[1]).slice(l.prefix.length)), l.kind];
      });
      log.getRange(2, 17, n, 2).setValues(out);
      props.setProperty('LOG_GROUP_FILLED', '1');
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

// グループNo（容器1本ごと）: 5kg・8kg・20kg は50本ごと（50001〜50050 → 1）、30kg・50kg系は100本ごと（59701〜59800 → 598）
function groupNoOf_(kind, serial) {
  const k = String(kind || '').trim(), n = Number(serial);
  if (!(n > 0)) return '';
  if (/^(30kg|50kg)/.test(k)) return Math.floor((n - 1) / 100) + 1;
  if (/^(5kg|8kg|20kg)/.test(k)) return n >= 50001 ? Math.floor((n - 50001) / 50) + 1 : '';
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

function findLot_(lotId) {
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
  if (!sh) return {};
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
    logSh.appendRow([recordId, new Date(), lot.lotId, "'" + lot.prefix + serial,
      "'" + displayNumber_(serial), typeof mass === 'number' ? mass : '', kind,
      payload.ngInput ? 'NG' : 'OK', payload.ngInput ? "'" + payload.ngInput : '',
      prev === null ? '' : prev, userEmail_(), '有効', String(payload.worker || ''), String(payload.device || ''), '', note,
      groupNoOf_(lot.kind, serial), lot.kind || '']);
    formatLogMass_(logSh, logSh.getLastRow(), 1);
    const result = { recordId: recordId, clientId: payload.clientId, serial: serial, mass: mass, prev: prev, kind: kind, note: note };

    // 全数（欠番を含む）そろったら自動で完了・PDF作成・送信する
    if (prev === null && countFilled_(lot) >= lotSize_(lot)) {
      result.full = true;
      if (!lot.sentAt && readSettings_().auto) {
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
  } finally {
    lock.releaseLock();
  }
}

// 直前の入力を取り消し、上書き前の値に戻す
function undoEntry(recordId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = getSheet_(SHEET_LOG);
    const last = log.getLastRow();
    const from = Math.max(2, last - 500);
    if (last < 2) throw new Error('取り消す記録がありません');
    const values = log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getDisplayValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][0] !== recordId) continue;
      if (values[i][11] !== '有効') throw new Error('既に取り消し済みです');
      const lot = findLot_(values[i][2]);
      const serial = values[i][3].slice(lot.prefix.length);
      const prevText = values[i][9];
      const prev = prevText === '' ? null : (prevText === MISSING || prevText === REPAIR || prevText === CIRCLE ? prevText : Number(prevText));
      const pos = slotPosition_(Number(serial) - Number(lot.start));
      getSheet_(lot.sheetName).getRange(pos.row, pos.numberCol + 1, 1, 5).setValues([massCells_(prev)]);
      log.getRange(from + i, 12).setValue('取消');
      log.getRange(from + i, 15).setValue(new Date());
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
    const label = lot.prefix + serial;
    let marked = 0;
    if (log.getLastRow() >= 2) {
      const last = log.getLastRow(), from = Math.max(2, last - 5000);
      const v = log.getRange(from, 1, last - from + 1, 12).getDisplayValues();
      const now = new Date();
      v.forEach(function(r, i) {
        if (r[2] === lot.lotId && r[3] === label && r[11] === '有効') {
          log.getRange(from + i, 12).setValue('取消');
          log.getRange(from + i, 15).setValue(now);
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
    if (r[11] !== '有効') return;
    if (r[6] === KIND_DUP || (r[9] !== '' && r[6] === '通常')) bump(r[2], 'dups', r[3]);
    if (r[15] === NOTE_SEAL) bump(r[2], 'seals', r[3]);
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
  const sh = getSheet_(lot.sheetName);
  const pdf = exportLotPdf_(lot, sh);
  const result = { lotId: lot.lotId, pdfUrl: pdf.url, sentTo: '', mailError: '' };
  if (send) {
    try {
      result.sentTo = sendLotMail_(lot, pdf.blob);
    } catch (e) {
      result.mailError = e.message;
      lots.getRange(lot.row, 17).setValue('送信失敗: ' + e.message);
    }
  }
  return result;
}

// 完了済みロットをもう一度送る
function resendLot(lotId) {
  const lot = findLot_(lotId);
  const sh = getSheet_(lot.sheetName);
  const pdf = exportLotPdf_(lot, sh);
  return { lotId: lot.lotId, pdfUrl: pdf.url, sentTo: sendLotMail_(lot, pdf.blob) };
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
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(lot.sheetName);
    if (sh) ss.deleteSheet(sh);
    getSheet_(SHEET_LOTS).deleteRow(lot.row);
    let marked = 0;
    const log = ss.getSheetByName(SHEET_LOG);
    if (log && log.getLastRow() >= 2) {
      const last = log.getLastRow(), from = Math.max(2, last - 5000);
      const v = log.getRange(from, 1, last - from + 1, 12).getValues();
      v.forEach(function(r, i) {
        if (r[2] === lot.lotId && r[11] === '有効') { log.getRange(from + i, 12).setValue('削除'); marked++; }
      });
    }
    return { lotId: lot.lotId, marked: marked };
  } finally {
    lock.releaseLock();
  }
}

function reopenLot(lotId) {
  const lot = findLot_(lotId);
  getSheet_(SHEET_LOTS).getRange(lot.row, 7).setValue(STATUS_ACTIVE);
  lot.status = STATUS_ACTIVE;
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
  const file = getPdfFolder_().createFile(blob);
  if (lot.row) getSheet_(SHEET_LOTS).getRange(lot.row, 14).setValue(file.getUrl());
  return { url: file.getUrl(), blob: blob };
}

function getPdfFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PDF_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 削除されていたら作り直す */ }
  }
  const folder = DriveApp.createFolder('LPG容器 検査成績表PDF');
  props.setProperty('PDF_FOLDER_ID', folder.getId());
  return folder;
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
