// 採番表: その日に流れた容器を、入力記録の入力順に「流れた順番・容器番号」で並べたPDF（紙の採番表と同じ形）
//  1ページ200本（50行×4列）。容器記号（底黒は別）ごとに1つのPDF
//  前日までの分は入力画面を開いたときに裏で自動作成（runMaintenance）。今日の分はホームのボタンでいつでも作れる
//  保存先: 「LPG容器 検査成績表PDF / 採番表 / 2026年 / 07月」

const SAIBAN_FOLDER = '採番表';
const SAIBAN_PER_PAGE = 200, SAIBAN_ROWS = 50;

function tokyoDate_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }

// その日に流れた容器（入力順）。欠番（流れていない）・訂正（同じ容器の打ち直し）・取消した入力は入れない
function saibanRows_(date) {
  const log = getSpreadsheet_().getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) return [];
  const last = log.getLastRow(), from = Math.max(2, last - 30000);
  const lots = {};
  readLots_().forEach(function(l) { lots[l.lotId] = l; });
  return log.getRange(from, 1, last - from + 1, LOG_HEADERS.length).getValues().filter(function(r) {
    const t = r[LC_TIME];
    return t instanceof Date && tokyoDate_(t) === date && r[LC_STATUS] !== '取消' &&
      r[LC_KIND] !== MISSING && r[LC_KIND] !== KIND_FIX;
  }).sort(function(a, b) { return a[LC_TIME] - b[LC_TIME]; }).map(function(r) {
    const lot = lots[String(r[LC_LOT])];
    const prefix = lot ? lot.prefix : '';
    const full = String(r[LC_SERIAL]);
    return {
      group: prefix + (lot && lot.spec === SPEC_SOKO ? '（底黒）' : ''),
      prefix: prefix,
      number: full.indexOf(prefix) === 0 ? full.slice(prefix.length) : full,
      kind: String(r[LC_KINDSIZE] || (lot && lot.kind) || ''),
      mark: r[LC_KIND] === KIND_DUP ? 'W' : r[LC_KIND] === REPAIR ? '修正' : '',
      worker: String(r[LC_WORKER] || '')
    };
  });
}

function saibanEsc_(s) { return String(s).replace(/[&<>"]/g, function(c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

// 1グループ分のHTML（A4縦、1ページ200本）
function saibanHtml_(date, group, rows) {
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
    for (let c = 0; c < 4; c++) t += '<th class="no">流れた<br>順番</th><th class="num">容器番号　' + ((rows[0] && rows[0].number.length) || 5) + '桁</th>';
    t += '</tr>';
    for (let i = 0; i < SAIBAN_ROWS; i++) {
      t += '<tr>';
      for (let c = 0; c < 4; c++) {
        const n = p * SAIBAN_PER_PAGE + c * SAIBAN_ROWS + i;   // 0始まり
        const r = rows[n];
        t += '<td class="no">' + (n + 1) + '</td><td class="num">' +
          (r ? saibanEsc_(r.number) + (r.mark ? '<span class="mk">' + saibanEsc_(r.mark) + '</span>' : '') : '') + '</td>';
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
    'th.no, td.no { width: 9%; } td.num { font-size: 10pt; letter-spacing: .5pt; }' +
    '.mk { font-size: 6.5pt; margin-left: 1.5mm; color: #b00; }' +
    '.foot { font-size: 8pt; text-align: right; margin-top: 1mm; }' +
    '</style></head><body>' + pages.join('') + '</body></html>';
}

function saibanFolder_(date) {
  const d = date.split('-');
  return childFolder_(childFolder_(childFolder_(getPdfFolder_(), SAIBAN_FOLDER), d[0] + '年'), d[1] + '月');
}

// 指定日の採番表PDFを作る（同じ名前の古いPDFはゴミ箱へ）。input: { date: 'yyyy-MM-dd' }（省略時は今日）
// → { date, files: [{ group, count, url }] }
function makeSaibanPdf(input) {
  const date = (input && input.date) || tokyoDate_(new Date());
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付は yyyy-MM-dd で指定してください');
  const rows = saibanRows_(date);
  const groups = {};
  rows.forEach(function(r) { (groups[r.group] = groups[r.group] || []).push(r); });
  const out = { date: date, files: [] };
  const names = Object.keys(groups).sort();
  if (!names.length) return out;
  const folder = saibanFolder_(date);
  names.forEach(function(g) {
    const name = '採番表_' + (g || '記号なし') + '_' + date + '.pdf';
    const old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);
    const blob = HtmlService.createHtmlOutput(saibanHtml_(date, g, groups[g])).getAs('application/pdf').setName(name);
    const file = folder.createFile(blob);
    out.files.push({ group: g, count: groups[g].length, url: file.getUrl() });
  });
  return out;
}

// 前日まででまだ作っていない日の採番表を作る（1回に3日分まで）。初回は昨日の分から
function saibanStep_() {
  const props = PropertiesService.getScriptProperties();
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
