// 納品メール（全数そろったら検査成績表PDF＋入力データCSVを自動送信）と、設定・入力者シートの読み込み

const SHEET_SETTINGS = '設定';
const SHEET_WORKERS = '入力者';
const SHEET_WORKERS_OLD = '担当者';   // 以前の名前（画面を開いたときに「入力者」へ名前を変える）

// 設定シート: A列=項目 / B列=値 / C列=説明
const SETTING_ROWS = [
  ['送付先', '', '納品メールの宛先。複数ならカンマ区切り（空欄のあいだは送信しない）'],
  ['CC', '', '控えを送る宛先（任意）'],
  ['自動送信', 'する', '「する」: 送付先が入っていれば、全数（欠番を含む）そろった時点で自動で完了・送信 ／「しない」: 完了ボタンで完了・送信（送付先が空のときも自動では完了しない）'],
  ['差出人名', '野田組', 'メールの差出人として表示する名前'],
  ['件名', '【高圧ガス容器検査成績表】{組容器番号}（{容器区分}）', '{組容器番号} {容器区分} {本数} {欠番} {修正} が使えます'],
  ['本文',
    'ご担当者様\n\nいつもお世話になっております。\n高圧ガス容器検査成績表をお送りします。\n\n' +
    '組容器番号：{組容器番号}\n容器区分：{容器区分}\n本数：{本数}（欠番 {欠番}・修正 {修正}）\n\n' +
    '添付：検査成績表（PDF）、質量データ（CSV）\n\nよろしくお願いいたします。',
    '改行はそのまま使えます'],
  ['標準質量', '5kg=6.8, 8kg=9.6, 10kg=11.3, 20kg=16.7, 20kg三部制=17.5, 30kg=23.5, 30kg把手=24.0, 50kg=34.8, 50kg S付=36.3',
    'ロットの1本目から候補ボタンを出すときの真ん中の値（容器区分=kg）。5本以上入るとそのロットの実際の中央値を使う'],
  ['完了後の成績表シート', '削除する',
    '「削除する」: 完了してPDFを保存したら成績表シートを消す（PDFはドライブに残る。再開すると入力記録から作り直す）／「残す」: 消さない'],
  ['内容積', '5kg=12, 8kg=19, 10kg=24, 20kg=47, 20kg三部制=47, 30kg=71, 30kg把手=71, 50kg=118, 50kg S付=118',
    '成績表の「内容積：　lit」に自動で入れる値（容器区分=リットル）。ロット登録時・機種変更時に入る。空欄の区分は入れない'],
  ['件名（まとめて送信）', '【高圧ガス容器検査成績表】{ロット数}ロット分（{組容器番号}）',
    'ホーム画面の「まとめてメール送信」の件名。{ロット数} {本数} {組容器番号}（最初～最後）が使えます'],
  ['本文（まとめて送信）',
    'ご担当者様\n\nいつもお世話になっております。\n高圧ガス容器検査成績表を{ロット数}ロット分まとめてお送りします。\n\n' +
    '{一覧}\n\n添付：検査成績表（PDF）、質量データ（CSV）各{ロット数}件\n採番表（PDF）：{採番表}\n\nよろしくお願いいたします。',
    '{一覧} にロットごとの「組容器番号（容器区分）本数」が1行ずつ入ります。{採番表} は付けた採番表の日付（付けないときはその行ごと省く）']
];

// v71の「本文（まとめて送信）」の初期値（{採番表} の行がない）。このままなら今の初期値に置き換える
const BULK_BODY_OLD_DEFAULT = 'ご担当者様\n\nいつもお世話になっております。\n高圧ガス容器検査成績表を{ロット数}ロット分まとめてお送りします。\n\n' +
  '{一覧}\n\n添付：検査成績表（PDF）、質量データ（CSV）各{ロット数}件\n\nよろしくお願いいたします。';
const BULK_MAIL_MAX_LOTS = 20;   // 1通に添付するロット数。超えたら複数のメールに分ける（添付は1通25MBまで）

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
    bulkSubject: get('件名（まとめて送信）'),
    bulkBody: get('本文（まとめて送信）'),
    typical: parseTypical_(get('標準質量')),
    volume: parseTypical_(get('内容積')),
    removeDoneSheet: String(get('完了後の成績表シート')).trim() !== '残す'
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

// メールに差し込む値（本数・欠番・修正など）
function lotMailVars_(lot) {
  const progress = readProgress_(lot);
  let missing = 0, repair = 0;
  Object.keys(progress).forEach(function(k) { if (progress[k] === MISSING) missing++; if (progress[k] === REPAIR) repair++; });
  return {
    '組容器番号': lot.prefix + lot.start + '～' + lot.prefix + lot.end,
    '容器区分': lot.kind || '',
    '本数': String(Object.keys(progress).length - missing - repair) + '本',
    '欠番': String(missing),
    '修正': String(repair),
    '耐圧試験日': lot.pressure && lot.pressure.testDate ? lot.pressure.testDate.replace(/-/g, '/') : ''
  };
}

// 値が空になる差し込み（例: 耐圧試験日）を含む行は、行ごと省く
function fillMail_(t, vars) {
  return String(t).split('\n').filter(function(line) {
    const keys = line.match(/\{([^}]+)\}/g) || [];
    return !keys.some(function(m) { const k = m.slice(1, -1); return k in vars && vars[k] === ''; });
  }).join('\n').replace(/\{([^}]+)\}/g, function(m, k) { return k in vars ? vars[k] : m; });
}

function markSent_(lot, st) {
  if (lot.row) getSheet_(SHEET_LOTS).getRange(lot.row, 16, 1, 2).setValues([[new Date(), st.to + (st.cc ? ' / CC: ' + st.cc : '')]]);
}

// 送信して宛先を返す。宛先が空なら送らずにエラー
function sendLotMail_(lot, pdfBlob) {
  const st = readSettings_();
  if (!st.to) throw new Error('設定シートの「送付先」が空です');
  const vars = lotMailVars_(lot);
  const options = {
    name: st.senderName,
    attachments: [pdfBlob, buildLotCsv_(lot)]
  };
  if (st.cc) options.cc = st.cc;
  MailApp.sendEmail(st.to, fillMail_(st.subject, vars), fillMail_(st.body, vars), options);
  markSent_(lot, st);
  return st.to;
}

// 「送らずに消す」: 完了したロットをメールで送らないことにし、最近完了したロットの一覧から外す（成績表PDFは残す）
// 送信先の列に「送らない」と書く。あとでまとめて送信・再送信すると送信日時・送信先で上書きされる
const NOT_SENT_MARK = '送らない';
function isSkipped_(lot) { return !lot.sentAt && String(lot.sentTo || '').indexOf(NOT_SENT_MARK) === 0; }
function skipLotMail(input) {
  const lotId = typeof input === 'object' && input ? input.lotId : input;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lot = findLot_(lotId);
    if (lot.status !== STATUS_DONE) throw new Error('完了していないロットです: ' + lot.lotId);
    if (lot.sentAt) throw new Error('送信済みのロットです: ' + lot.lotId);
    getSheet_(SHEET_LOTS).getRange(lot.row, 17).setValue(NOT_SENT_MARK + '（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '）');
    return { lotId: lot.lotId };
  } finally {
    lock.releaseLock();
  }
}

// まとめて送信の候補（完了したロット。新しい順）
function listDoneLots() {
  return readLots_().filter(function(l) { return l.status === STATUS_DONE; }).reverse().map(function(l) {
    return { lotId: l.lotId, prefix: l.prefix, start: l.start, end: l.end, kind: l.kind, sheetName: l.sheetName,
      sentAt: l.sentAt, pdf: l.pdf, spec: l.spec, skipped: isSkipped_(l) };
  });
}

// 選んだ完了ロットの成績表PDF＋CSVを、1通（多いときは BULK_MAIL_MAX_LOTS ロットずつ）にまとめて送る
// input: [lotId, ...] または { lotIds: [...], saiban: true }（saiban: そのロットの容器が流れた日の採番表PDFも付ける）
function sendLotsMail(input) {
  const st = readSettings_();
  if (!st.to) throw new Error('設定シートの「送付先」が空です');
  const lotIds = Array.isArray(input) ? input : (input && input.lotIds);
  const withSaiban = !Array.isArray(input) && !!(input && input.saiban);
  const ids = (lotIds || []).map(String).filter(function(id, i, a) { return id && a.indexOf(id) === i; });
  if (!ids.length) throw new Error('送るロットを選んでください');
  const lots = ids.map(function(id) {
    const lot = findLot_(id);
    if (lot.status !== STATUS_DONE) throw new Error('完了していないロットは送れません: ' + lot.lotId);
    return lot;
  });
  // 番号順にそろえる
  lots.sort(function(a, b) { return a.prefix.localeCompare(b.prefix) || Number(a.start) - Number(b.start); });
  const logRows = logDisplayRows_();
  const lotDays = withSaiban ? saibanDatesForLots_(lots.map(function(l) { return l.lotId; })) : {};
  const saibanBlobs = {};   // 日付 → PDF（分けて送るときに同じ日を作り直さない）
  let mails = 0, saibanCount = 0;
  for (let i = 0; i < lots.length; i += BULK_MAIL_MAX_LOTS) {
    const part = lots.slice(i, i + BULK_MAIL_MAX_LOTS);
    const attachments = [], lines = [];
    let total = 0;
    part.forEach(function(lot) {
      attachments.push(lotPdf_(lot).blob, buildLotCsv_(lot, logRows));
      const v = lotMailVars_(lot);
      total += parseInt(v['本数'], 10) || 0;
      lines.push('・' + v['組容器番号'] + '（' + v['容器区分'] + '）' + v['本数'] +
        (v['欠番'] !== '0' ? '　欠番 ' + v['欠番'] : '') + (v['修正'] !== '0' ? '　修正 ' + v['修正'] : ''));
    });
    const days = [];
    part.forEach(function(lot) { (lotDays[lot.lotId] || []).forEach(function(d) { if (days.indexOf(d) < 0) days.push(d); }); });
    days.sort();
    const sentDays = [];
    days.forEach(function(d) {
      if (!(d in saibanBlobs)) saibanBlobs[d] = saibanPdfBlob_(d);
      if (saibanBlobs[d]) { attachments.push(saibanBlobs[d]); sentDays.push(d); }
    });
    saibanCount += sentDays.length;
    const first = part[0], last = part[part.length - 1];
    const vars = {
      'ロット数': String(part.length),
      '本数': total + '本',
      '組容器番号': first.prefix + first.start + (part.length > 1 ? '～' + last.prefix + last.end : '～' + first.prefix + first.end),
      '一覧': lines.join('\n'),
      '採番表': sentDays.map(function(d) { return d.slice(5).replace('-', '/'); }).join('・')
    };
    const options = { name: st.senderName, attachments: attachments };
    if (st.cc) options.cc = st.cc;
    MailApp.sendEmail(st.to, fillMail_(st.bulkSubject, vars), fillMail_(st.bulkBody, vars), options);
    part.forEach(function(lot) { markSent_(lot, st); });
    mails++;
  }
  return { sentTo: st.to, lots: lots.length, mails: mails, saiban: saibanCount };
}

// 入力記録（表示どおりの文字）
function logDisplayRows_() {
  const log = getSheet_(SHEET_LOG);
  return log.getLastRow() < 2 ? [] :
    log.getRange(2, 1, log.getLastRow() - 1, Math.min(log.getLastColumn(), LOG_HEADERS.length)).getDisplayValues();
}

// ロットの最新の有効な入力（容器ごと）をCSVに。Excelで開けるようBOM付きUTF-8
// logRows: まとめて送信のとき、1度読んだ入力記録を使い回す
function buildLotCsv_(lot, logRows) {
  const latest = {};
  (logRows || logDisplayRows_()).forEach(function(r) {
    if (r[LC_LOT] === lot.lotId && r[LC_STATUS] === '有効') latest[r[LC_SERIAL]] = r;
  });
  // 年替わりで別のスプレッドシートへ移したロットは、そこから読む（表示と同じ形の文字にそろえる）
  if (!Object.keys(latest).length) {
    archivedRowsForLot_(lot.lotId).forEach(function(r) {
      if (r[LC_STATUS] !== '有効') return;
      latest[String(r[LC_SERIAL])] = r.map(function(v, i) {
        if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        if ((i === LC_MASS || i === LC_PREV) && typeof v === 'number') return v.toFixed(1);
        return String(v);
      });
    });
  }
  const rows = [['容器番号', '表示番号', '質量(kg)', '区分', '一致結果', '入力日時', '入力者', '備考']];
  const width = lot.start.length;
  for (let n = Number(lot.start); n <= Number(lot.end); n++) {
    const serial = lot.prefix + padSerial_(n, width);
    const r = latest[serial];
    rows.push(r ? [serial, r[LC_DISP], r[LC_MASS], r[LC_KIND], r[LC_MATCH], r[LC_TIME], r[LC_WORKER], r[LC_NOTE] || ''] : [serial, displayNumber_(n), '', '未入力', '', '', '', '']);
  }
  const csv = rows.map(function(r) {
    return r.map(function(v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(',');
  }).join('\r\n');
  return Utilities.newBlob('﻿' + csv, 'text/csv', lot.sheetName + '.csv');
}
