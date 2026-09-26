// LPG容器 質量転記アプリ
// 検査スタッフがiPad（外付けテンキー）で「容器番号」と「質量」を入力すると、
// 組容器番号から記録先の検査成績表シートを判定して自動転記する。

const SHEET_LOTS = 'ロット';
const SHEET_LOG = '入力記録';
const SHEET_TEMPLATE = '書式_成績表';
const REPORT_PREFIX = '成績表_';

const LOT_HEADERS = ['ロットID', '登録日時', '記号', '開始番号', '終了番号', '容器区分', '状態',
  '成績表シート', '代表容器番号', '耐圧試験日', '全増加(cm3)', '恒久増加(cm3)', '恒久増加率(%)',
  'PDF', '登録者', '送信日時', '送信先'];
const LOG_HEADERS = ['記録ID', '入力日時', 'ロットID', '容器番号', '表示番号', '質量(kg)', '区分',
  '一致結果', '入力番号(NG時)', '上書き前', '入力者', '状態', '担当者', '端末', '取消日時'];

const STATUS_ACTIVE = '入力中';
const STATUS_DONE = '完了';
const MISSING = '欠番';

// 成績表の配置（元Excel様式）: 5ブロック × 20行、1ブロック6列
// 容器番号 | ☑ | 質量10の位 | 1の位 | "," | 小数1位
const GRID_FIRST_ROW = 5;
const GRID_ROWS = 20;
const GRID_FIRST_COL = 2; // B列
const BLOCK_WIDTH = 6;
const BLOCKS = 5;
const LOT_MAX = GRID_ROWS * BLOCKS; // 100本

const MASS_MIN = 0.1;
const MASS_MAX = 99.9;

// ---------- Web app ----------

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('LPG容器 質量入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
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

function getSheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。メニュー「LPG容器検査 > 初期セットアップ」を実行してください。');
  return sh;
}

function userEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// ---------- 純粋ロジック（テスト可能） ----------

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
  try { ensureTokyoTime_(); } catch (e) { /* 権限がない場合も画面は開く */ }
  const lots = readLots_().filter(function(l) { return l.status === STATUS_ACTIVE; });
  const settings = readSettings_();
  const dups = readDupsAll_();
  return {
    user: userEmail_(),
    workers: readWorkers_(),
    mail: { to: settings.to, cc: settings.cc, auto: settings.auto },
    typical: settings.typical,
    lots: lots.map(function(l) { return withProgress_(l, dups); }),
    recentDone: readLots_().filter(function(l) { return l.status === STATUS_DONE; })
      .slice(-10).reverse().map(function(l) { return withProgress_(l, dups); })
  };
}

function withProgress_(lot, dupMap) {
  lot.entries = readProgress_(lot);
  lot.dups = (dupMap || readDupsAll_())[lot.lotId] || {};
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
      created.push(createLot({ prefix: prefix, start: padSerial_(b, width), end: padSerial_(b + LOT_MAX - 1, width), kind: input.kind }).lotId);
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

    const lotId = prefix + start;
    const lots = readLots_();
    lots.forEach(function(l) {
      if (l.lotId === lotId) throw new Error('この組容器番号は登録済みです: ' + lotId);
    });
    const sheetName = REPORT_PREFIX + lotId + '-' + end.slice(-Math.min(end.length, 5));
    const lot = {
      lotId: lotId, prefix: prefix, start: start, end: end,
      kind: String(input.kind || ''), status: STATUS_ACTIVE, sheetName: sheetName,
      pressure: { repNumber: start, testDate: '', totalExp: '', permExp: '', permRate: '' }, pdf: ''
    };
    createReportSheet_(lot);
    getSheet_(SHEET_LOTS).appendRow([lotId, new Date(), prefix, "'" + start, "'" + end, lot.kind,
      STATUS_ACTIVE, sheetName, "'" + start, '', '', '', '', '', userEmail_()]);
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
  sh.getRange('V2').setValue(lot.prefix + lot.start + ' ～ ' + endLabel);
  sh.getRange('H3').setValue(lot.prefix);

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
    if (payload.missing) {
      mass = MISSING;
    } else {
      mass = Math.round(Number(payload.mass) * 10) / 10;
      if (!isValidMass_(mass)) throw new Error('質量は' + MASS_MIN + '〜' + MASS_MAX + 'の範囲で入力してください');
    }

    const sh = getSheet_(lot.sheetName);
    const pos = slotPosition_(index);
    const range = sh.getRange(pos.row, pos.numberCol + 1, 1, 5);
    const prev = parseMassCells_(range.getValues()[0]);
    range.setValues([massCells_(mass)]);

    const recordId = 'R' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMddHHmmss') +
      Math.floor(Math.random() * 1000);
    getSheet_(SHEET_LOG).appendRow([recordId, new Date(), lot.lotId, "'" + lot.prefix + serial,
      "'" + displayNumber_(serial), mass === MISSING ? '' : mass, mass === MISSING ? MISSING : '通常',
      payload.ngInput ? 'NG' : 'OK', payload.ngInput ? "'" + payload.ngInput : '',
      prev === null ? '' : prev, userEmail_(), '有効', String(payload.worker || ''), String(payload.device || ''), '']);
    const result = { recordId: recordId, clientId: payload.clientId, serial: serial, mass: mass, prev: prev };

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
      const prev = prevText === '' ? null : (prevText === MISSING ? MISSING : Number(prevText));
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

// ダブり（入力済みの容器を上書きした記録）の数: { ロットID: { 容器番号: 回数 } }。入力記録の直近5000行から数える
function readDupsAll_() {
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) return {};
  const last = log.getLastRow(), from = Math.max(2, last - 5000);
  const out = {};
  log.getRange(from, 1, last - from + 1, 12).getDisplayValues().forEach(function(r) {
    if (r[11] === '有効' && r[9] !== '') {
      out[r[2]] = out[r[2]] || {};
      out[r[2]][r[3]] = (out[r[2]][r[3]] || 0) + 1;
    }
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
    return { lotId: lot.lotId, kind: kind };
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
