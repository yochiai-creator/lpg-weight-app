// 納品メール（全数そろったら検査成績表PDF＋入力データCSVを自動送信）と、設定・入力者シートの読み込み

const SHEET_SETTINGS = '設定';
const SHEET_WORKERS = '入力者';
const SHEET_WORKERS_OLD = '担当者';   // 以前の名前（画面を開いたときに「入力者」へ名前を変える）

// 設定シート: A列=項目 / B列=値 / C列=説明
const SETTING_ROWS = [
  ['送付先', '', '納品メールの宛先。複数ならカンマ区切り（空欄のあいだは送信しない）'],
  ['CC', '', '控えを送る宛先（任意）'],
  ['自動送信', 'する', '「する」: 全数（欠番を含む）そろった時点で自動送信 ／「しない」: 完了ボタンで送信'],
  ['差出人名', '野田組', 'メールの差出人として表示する名前'],
  ['件名', '【高圧ガス容器検査成績表】{組容器番号}（{容器区分}）', '{組容器番号} {容器区分} {本数} {欠番} {修正} が使えます'],
  ['本文',
    'ご担当者様\n\nいつもお世話になっております。\n高圧ガス容器検査成績表をお送りします。\n\n' +
    '組容器番号：{組容器番号}\n容器区分：{容器区分}\n本数：{本数}（欠番 {欠番}・修正 {修正}）\n\n' +
    '添付：検査成績表（PDF）、質量データ（CSV）\n\nよろしくお願いいたします。',
    '改行はそのまま使えます'],
  ['標準質量', '5kg=6.8, 8kg=9.6, 20kg=16.7, 30kg=24.0, 50kg=34.8, 50kg S付=36.3',
    'ロットの1本目から候補ボタンを出すときの真ん中の値（容器区分=kg）。5本以上入るとそのロットの実際の中央値を使う']
];

// 「5kg=6.8, 50kg S付=36.3」→ { '5kg': 6.8, '50kg S付': 36.3 }
function parseTypical_(text) {
  const out = {};
  String(text || '').split(/[,、\n]/).forEach(function(part) {
    const m = part.match(/^\s*(.+?)\s*[=＝]\s*([\d.]+)\s*$/);
    if (m) out[m[1]] = Number(m[2]);
  });
  return out;
}

function readSettings_() {
  const sh = getSpreadsheet_().getSheetByName(SHEET_SETTINGS);
  const map = {};
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function(r) {
      if (r[0] !== '') map[String(r[0]).trim()] = String(r[1]);
    });
  }
  const def = {};
  SETTING_ROWS.forEach(function(r) { def[r[0]] = r[1]; });
  const get = function(k) { return k in map ? map[k] : def[k]; };
  return {
    to: get('送付先').trim(),
    cc: get('CC').trim(),
    auto: get('自動送信').trim() !== 'しない',
    senderName: get('差出人名'),
    subject: get('件名'),
    body: get('本文'),
    typical: parseTypical_(get('標準質量'))
  };
}

function readWorkers_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_WORKERS) || ss.getSheetByName(SHEET_WORKERS_OLD);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues()
    .map(function(r) { return r[0].trim(); })
    .filter(function(n) { return n; });
}

// 送信して宛先を返す。宛先が空なら送らずにエラー
function sendLotMail_(lot, pdfBlob) {
  const st = readSettings_();
  if (!st.to) throw new Error('設定シートの「送付先」が空です');
  const progress = readProgress_(lot);
  let missing = 0, repair = 0;
  Object.keys(progress).forEach(function(k) { if (progress[k] === MISSING) missing++; if (progress[k] === REPAIR) repair++; });
  const vars = {
    '組容器番号': lot.prefix + lot.start + '～' + lot.prefix + lot.end,
    '容器区分': lot.kind || '',
    '本数': String(Object.keys(progress).length - missing - repair) + '本',
    '欠番': String(missing),
    '修正': String(repair),
    '耐圧試験日': lot.pressure && lot.pressure.testDate ? lot.pressure.testDate.replace(/-/g, '/') : ''
  };
  // 値が空になる差し込み（例: 耐圧試験日）を含む行は、行ごと省く
  const fill = function(t) {
    return String(t).split('\n').filter(function(line) {
      const keys = line.match(/\{([^}]+)\}/g) || [];
      return !keys.some(function(m) { const k = m.slice(1, -1); return k in vars && vars[k] === ''; });
    }).join('\n').replace(/\{([^}]+)\}/g, function(m, k) { return k in vars ? vars[k] : m; });
  };
  const options = {
    name: st.senderName,
    attachments: [pdfBlob, buildLotCsv_(lot)]
  };
  if (st.cc) options.cc = st.cc;
  MailApp.sendEmail(st.to, fill(st.subject), fill(st.body), options);

  const lots = getSheet_(SHEET_LOTS);
  if (lot.row) lots.getRange(lot.row, 16, 1, 2).setValues([[new Date(), st.to + (st.cc ? ' / CC: ' + st.cc : '')]]);
  return st.to;
}

// ロットの最新の有効な入力（容器ごと）をCSVに。Excelで開けるようBOM付きUTF-8
function buildLotCsv_(lot) {
  const log = getSheet_(SHEET_LOG);
  const latest = {};
  if (log.getLastRow() >= 2) {
    log.getRange(2, 1, log.getLastRow() - 1, Math.min(log.getLastColumn(), LOG_HEADERS.length)).getDisplayValues().forEach(function(r) {
      if (r[2] === lot.lotId && r[11] === '有効') latest[r[3]] = r;
    });
  }
  const rows = [['容器番号', '表示番号', '質量(kg)', '区分', '一致結果', '入力日時', '入力者', '備考']];
  const width = lot.start.length;
  for (let n = Number(lot.start); n <= Number(lot.end); n++) {
    const serial = lot.prefix + padSerial_(n, width);
    const r = latest[serial];
    rows.push(r ? [serial, r[4], r[5], r[6], r[7], r[1], r[12], r[15] || ''] : [serial, displayNumber_(n), '', '未入力', '', '', '', '']);
  }
  const csv = rows.map(function(r) {
    return r.map(function(v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(',');
  }).join('\r\n');
  return Utilities.newBlob('﻿' + csv, 'text/csv', lot.sheetName + '.csv');
}
