// 採番表: その日に流れた容器を、入力記録の入力順に「流れた順番・容器番号」で並べる
//  1日分を容器記号もまとめて1つの通し番号にして、
//   ・スプレッドシート「LPG容器 質量入力 / 採番表 / 2026年 / LPG容器 採番表_2026-10」（月ごと）の「採番表」シートに書き足す
//   ・PDF（紙の採番表と同じ形、1ページ200本＝50行×4列）を「LPG容器 質量入力 / 採番表 / 2026年 / 07月」に保存
//  毎晩22時台に自動（nightlyJob）。取りこぼした日は入力画面を開いたときに裏で作る。ホームのボタンでいつでも作り直せる

const SAIBAN_FOLDER = '採番表';
const SAIBAN_PER_PAGE = 200, SAIBAN_ROWS = 50;

function tokyoDate_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }

// その日に流れた容器（入力順）。欠番（流れていない）・訂正（同じ容器の打ち直し）・ダブり・取消した入力は入れない
function saibanRows_(date) {
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) return [];
  const last = log.getLastRow(), from = Math.max(2, last - 30000);
  const lots = {};
  readLots_().forEach(function(l) { lots[l.lotId] = l; });
  return log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getValues().filter(function(r) {
    const t = r[LC_TIME];
    return t instanceof Date && tokyoDate_(t) === date && r[LC_STATUS] !== '取消' &&
      r[LC_KIND] !== MISSING && r[LC_KIND] !== KIND_FIX && r[LC_KIND] !== KIND_DUP;
  }).sort(function(a, b) { return a[LC_TIME] - b[LC_TIME]; }).map(function(r) {
    const lot = lots[String(r[LC_LOT])];
    const prefix = lot ? lot.prefix : '';
    const full = String(r[LC_SERIAL]);
    return {
      group: prefix + (lot && lot.spec === SPEC_SOKO ? '（底黒）' : ''),
      prefix: prefix,
      full: full,
      number: full.indexOf(prefix) === 0 ? full.slice(prefix.length) : full,
      time: r[LC_TIME],
      groupNo: r[LC_GROUP],
      lotId: String(r[LC_LOT]),
      kind: String(r[LC_KINDSIZE] || (lot && lot.kind) || ''),
      mark: r[LC_KIND] === REPAIR ? '修正' : '',
      worker: String(r[LC_WORKER] || '')
    };
  });
}

function saibanEsc_(s) { return String(s).replace(/[&<>"]/g, function(c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

// 1日分のHTML（A4縦、1ページ200本）。容器記号が2つ以上ある日は番号に記号を付ける
function saibanHtml_(date, group, rows) {
  const multi = rows.some(function(r) { return r.group !== (rows[0] && rows[0].group); });
  const kinds = [], workers = [];
  rows.forEach(function(r) {
    if (r.kind && kinds.indexOf(r.kind) < 0) kinds.push(r.kind);
    if (r.worker && workers.indexOf(r.worker) < 0) workers.push(r.worker);
  });
  const kg = kinds.length ? kinds.map(function(k) { return k.replace(/kg.*$/, ''); })
    .filter(function(k, i, a) { return a.indexOf(k) === i; }).join('・') + 'kg容器' : '容器';
  const d = date.split('-');
  const head = '<div class="head"><div class="title">' + saibanEsc_(kg) + '　採番表</div>' +
    '<div class="meta"><span>容器記号：<b>' + saibanEsc_(group || '（なし）') + '</b></span>' +
    '<span>作業日：<b>' + Number(d[0]) + '年' + Number(d[1]) + '月' + Number(d[2]) + '日</b></span>' +
    '<span>作業者：<b>' + saibanEsc_(workers.join('・')) + '</b></span></div></div>';
  const pages = [];
  for (let p = 0; p * SAIBAN_PER_PAGE < rows.length; p++) {
    let t = '<table><tr>';
    for (let c = 0; c < 4; c++) t += '<th class="no">流れた<br>順番</th><th class="num">容器番号　' + ((rows[0] && rows[0].number.length) || 5) + '桁</th><th class="tm">時刻</th>';
    t += '</tr>';
    for (let i = 0; i < SAIBAN_ROWS; i++) {
      t += '<tr>';
      for (let c = 0; c < 4; c++) {
        const n = p * SAIBAN_PER_PAGE + c * SAIBAN_ROWS + i;   // 0始まり
        const r = rows[n];
        t += '<td class="no">' + (n + 1) + '</td><td class="num">' +
          (r ? saibanEsc_(multi ? r.full + (r.group !== r.prefix ? '底黒' : '') : r.number) + (r.mark ? '<span class="mk">' + saibanEsc_(r.mark) + '</span>' : '') : '') + '</td>' +
          '<td class="tm">' + (r && r.time instanceof Date ? Utilities.formatDate(r.time, 'Asia/Tokyo', 'HH:mm') : '') + '</td>';
      }
      t += '</tr>';
    }
    t += '</table>';
    pages.push('<div class="page">' + head + t + '<div class="foot">' + (p + 1) + ' / ' +
      Math.ceil(rows.length / SAIBAN_PER_PAGE) + '　（全 ' + rows.length + ' 本）</div></div>');
  }
  return '<html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 8mm; }' +
    'body { font-family: sans-serif; margin: 0; }' +
    '.page { page-break-after: always; } .page:last-child { page-break-after: auto; }' +
    '.head { margin-bottom: 2mm; } .title { font-size: 14pt; font-weight: bold; }' +
    '.meta { font-size: 10pt; margin-top: 1mm; } .meta span { margin-right: 10mm; }' +
    'table { width: 100%; border-collapse: collapse; table-layout: fixed; }' +
    'th, td { border: 0.6pt solid #000; font-size: 8.5pt; height: 4.9mm; padding: 0 1mm; }' +
    'th { font-size: 7pt; line-height: 1.1; background: #eee; } td.no { text-align: center; }' +
    'th.no, td.no { width: 6.5%; } th.tm, td.tm { width: 7%; } td.tm { font-size: 8pt; text-align: center; }' +
    'td.num { font-size: ' + (multi ? '8.5pt' : '10pt') + '; letter-spacing: .5pt; white-space: nowrap; overflow: hidden; }' +
    '.mk { font-size: 6.5pt; margin-left: 1.5mm; color: #b00; }' +
    '.foot { font-size: 8pt; text-align: right; margin-top: 1mm; }' +
    '</style></head><body>' + pages.join('') + '</body></html>';
}

function saibanFolder_(date) {
  const d = date.split('-');
  return childFolder_(childFolder_(childFolder_(getPdfFolder_(), SAIBAN_FOLDER), d[0] + '年'), d[1] + '月');
}

// ---------- スプレッドシート（年ごと） ----------
const SAIBAN_SS_PREFIX = 'LPG容器 採番表_';
const SAIBAN_HEADERS = ['作業日', '流れた順番', '容器記号', '容器番号', '容器区分', 'グループNo', '入力時刻', '入力者', '備考', 'ロットID'];

// 一覧のスプレッドシートは月ごと（key = 'yyyy-MM'）。「採番表 / 2026年 / LPG容器 採番表_2026-10」
// （v69より前は年ごと 'yyyy' で「採番表」直下。そのファイルはそのまま残る）
function saibanSheet_(key) {
  const props = sharedProps_();
  let ids = {};
  try { ids = JSON.parse(props.getProperty('SAIBAN_SS_IDS') || '{}'); } catch (e) { ids = {}; }
  if (ids[key]) {
    try { return SpreadsheetApp.openById(ids[key]).getSheets()[0]; } catch (e) { /* 消されていたら作り直す */ }
  }
  const base = childFolder_(getPdfFolder_(), SAIBAN_FOLDER);
  const folder = /^\d{4}-\d{2}$/.test(key) ? childFolder_(base, key.slice(0, 4) + '年') : base;
  const found = findSpreadsheetIn_(folder, SAIBAN_SS_PREFIX + key);
  if (found) { ids[key] = found.getId(); props.setProperty('SAIBAN_SS_IDS', JSON.stringify(ids)); return found.getSheets()[0]; }
  const ss = SpreadsheetApp.create(SAIBAN_SS_PREFIX + key);
  try { ss.setSpreadsheetTimeZone('Asia/Tokyo'); } catch (e) { /* 無視 */ }
  try { DriveApp.getFileById(ss.getId()).moveTo(folder); } catch (e) { /* マイドライブに残る */ }
  const sh = ss.getSheets()[0].setName('採番表');
  sh.getRange(1, 1, 1, SAIBAN_HEADERS.length).setValues([SAIBAN_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > SAIBAN_HEADERS.length) sh.deleteColumns(SAIBAN_HEADERS.length + 1, sh.getMaxColumns() - SAIBAN_HEADERS.length);
  ids[key] = ss.getId();
  props.setProperty('SAIBAN_SS_IDS', JSON.stringify(ids));
  return sh;
}

// その日の行を書き直す（前に書いた同じ日の行は消してから書き足し、日付・順番で並べ直す）
function writeSaibanSheet_(date, rows) {
  const sh = saibanSheet_(date.slice(0, 7));   // 月ごとのファイル
  const last = sh.getLastRow();
  if (last >= 2) {
    const days = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (let i = days.length - 1; i >= 0; i--) {
      if (days[i][0] !== date) continue;
      let j = i;
      while (j > 0 && days[j - 1][0] === date) j--;
      sh.deleteRows(2 + j, i - j + 1);
      i = j;
    }
  }
  if (rows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, SAIBAN_HEADERS.length).setValues(rows.map(function(r, i) {
      return ["'" + date, i + 1, r.group, "'" + r.number, r.kind, r.groupNo === '' ? '' : r.groupNo, r.time, r.worker, r.mark, r.lotId];
    }));
    sh.getRange(start, 7, rows.length, 1).setNumberFormat('HH:mm:ss');
    if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, SAIBAN_HEADERS.length).sort([{ column: 1 }, { column: 2 }]);
    // 日付が変わる所に線（並べ直したので引き直す）
    const n = sh.getLastRow() - 1;
    sh.getRange(2, 1, n, SAIBAN_HEADERS.length).setBorder(false, null, null, null, null, false);
    const days = sh.getRange(2, 1, n, 1).getDisplayValues();
    for (let i = 1; i < days.length; i++) if (days[i][0] !== days[i - 1][0]) dayLine_(sh, 2 + i, SAIBAN_HEADERS.length);
  }
  return sh.getParent().getUrl();
}

// 指定日の採番表（スプレッドシートへの書き足し＋PDF）を作る。同じ日の古い行・PDFは置き換える
// input: { date: 'yyyy-MM-dd' }（省略時は今日） → { date, count, groups: [記号], pdfUrl, sheetUrl }
function makeSaibanPdf(input) {
  const date = (input && input.date) || tokyoDate_(new Date());
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付は yyyy-MM-dd で指定してください');
  const rows = saibanRows_(date);
  const out = { date: date, count: rows.length, groups: [], pdfUrl: '', sheetUrl: '' };
  if (!rows.length) return out;
  rows.forEach(function(r) { if (out.groups.indexOf(r.group) < 0) out.groups.push(r.group); });
  out.sheetUrl = writeSaibanSheet_(date, rows);
  const folder = saibanFolder_(date);
  const name = '採番表_' + date + '.pdf';
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  const label = out.groups.map(function(g) { return g || '（なし）'; }).join('・');
  const blob = HtmlService.createHtmlOutput(saibanHtml_(date, label, rows)).getAs('application/pdf').setName(name);
  out.pdfUrl = folder.createFile(blob).getUrl();
  return out;
}

// メールに付ける採番表PDF。前日までは保存済みのPDFを使い（なければ作る）、今日の分はその時点までで作り直す
function saibanPdfBlob_(date) {
  const name = '採番表_' + date + '.pdf';
  if (date < tokyoDate_(new Date())) {
    const it = saibanFolder_(date).getFilesByName(name);
    if (it.hasNext()) return it.next().getBlob().setName(name);
  }
  if (!makeSaibanPdf({ date: date }).count) return null;
  const it2 = saibanFolder_(date).getFilesByName(name);
  return it2.hasNext() ? it2.next().getBlob().setName(name) : null;
}

// ロットの容器が流れた日（採番表に載る入力がある日）。lotIds → { lotId: ['yyyy-MM-dd', ...] }
function saibanDatesForLots_(lotIds) {
  const out = {};
  lotIds.forEach(function(id) { out[id] = []; });
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) return out;
  log.getRange(2, 1, log.getLastRow() - 1, LOG_HEADERS.length).getValues().forEach(function(r) {
    const days = out[String(r[LC_LOT])];
    const t = r[LC_TIME];
    if (!days || !(t instanceof Date) || r[LC_STATUS] === '取消' ||
      r[LC_KIND] === MISSING || r[LC_KIND] === KIND_FIX || r[LC_KIND] === KIND_DUP) return;
    const d = tokyoDate_(t);
    if (days.indexOf(d) < 0) days.push(d);
  });
  return out;
}

// 時刻の列を足す前（v75まで）に作った採番表を、1回だけ作り直す（1回に3日分。残りは次の呼び出しで）
// 対象は入力記録にある日のうち、自動で作り終えた日まで（今日の分は夜に作る）
function saibanRedoStep_() {
  const props = sharedProps_();
  if (props.getProperty('SAIBAN_TIME_REDONE') === '1') return 0;
  let todo = props.getProperty('SAIBAN_TIME_TODO');
  if (todo === null) {
    const until = props.getProperty('SAIBAN_DONE_UNTIL') || '';
    const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
    const days = [];
    if (log && log.getLastRow() >= 2) {
      const last = log.getLastRow(), from = Math.max(2, last - 30000);
      log.getRange(from, LC_TIME + 1, last - from + 1, 1).getValues().forEach(function(r) {
        if (!(r[0] instanceof Date)) return;
        const d = tokyoDate_(r[0]);
        if (d <= until && days.indexOf(d) < 0) days.push(d);
      });
    }
    todo = days.sort().join(',');
  }
  const list = todo ? todo.split(',') : [];
  let made = 0;
  while (list.length && made < 3) { makeSaibanPdf({ date: list.shift() }); made++; }
  if (list.length) props.setProperty('SAIBAN_TIME_TODO', list.join(','));
  else { props.setProperty('SAIBAN_TIME_TODO', ''); props.setProperty('SAIBAN_TIME_REDONE', '1'); }
  return list.length ? made : 0;
}

// 前日まででまだ作っていない日の採番表を作る（1回に3日分まで）。初回は昨日の分から
function saibanStep_() {
  const props = sharedProps_();
  const today = tokyoDate_(new Date());
  const day = 24 * 3600 * 1000;
  let doneUntil = props.getProperty('SAIBAN_DONE_UNTIL');
  if (!doneUntil) doneUntil = tokyoDate_(new Date(Date.now() - 2 * day));
  let made = 0;
  for (let k = 0; k < 3; k++) {
    const next = tokyoDate_(new Date(new Date(doneUntil + 'T12:00:00+09:00').getTime() + day));
    if (next >= today) break;
    makeSaibanPdf({ date: next });
    props.setProperty('SAIBAN_DONE_UNTIL', next);
    doneUntil = next;
    made++;
  }
  return made;
}

// ---------- 毎晩の自動処理（夜は作業しないので、その日の分を夜にまとめて作る） ----------
const NIGHTLY_HOUR = 22;   // 22時台に動く

// 時間指定の自動処理から呼ばれる: 今日の採番表PDF、入力記録の整理、完了済みシートの片付け
function nightlyJob() {
  if (isOldProject_()) return;                     // 旧URLのプロジェクトでは動かさない（二重に作らない）
  const today = tokyoDate_(new Date());
  try { mergeRootFolders_(); } catch (e) { console.error('フォルダの統合に失敗: ' + e.message); }
  saibanStep_();                                   // 取りこぼした前日までの分
  makeSaibanPdf({ date: today });
  sharedProps_().setProperty('SAIBAN_DONE_UNTIL', today);
  const until = Date.now() + 4 * 60 * 1000;        // 実行時間の上限（6分）に余裕を持たせる
  let r;
  do { r = runMaintenance(); } while (r.archiveDone === false && Date.now() < until);
  try { cleanDoneSheets_(); } catch (e) { console.error('完了済みシートの片付けに失敗: ' + e.message); }
}

// メニュー「毎晩の自動処理を設定」: 毎日22時台に nightlyJob を動かす（何度押しても1つだけ）
function installNightlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'nightlyJob') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlyJob').timeBased().everyDays(1).atHour(NIGHTLY_HOUR).inTimezone('Asia/Tokyo').create();
  const msg = '毎晩 ' + NIGHTLY_HOUR + '時台に、その日の採番表PDFを作ります（入力記録の整理・完了済みシートの片付けも一緒に行います）。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
  return msg;
}
