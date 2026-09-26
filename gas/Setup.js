// 初期セットアップ（スプレッドシートのメニューから1回実行）
// 1. スプレッドシートIDをスクリプトプロパティに保存（ウェブアプリから開くため）
// 2. 「ロット」「入力記録」シートを作成
// 3. 元Excelから取り込んだ「検査成績表 (001-100)」を複製して「書式_成績表」を作成

const TEMPLATE_SOURCE_CANDIDATES = ['検査成績表 (001-100)', '検査成績表(001-100)', '検査成績表 ※例'];

function setup() {
  const ss = getSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  if (ss.getSpreadsheetTimeZone() !== 'Asia/Tokyo') ss.setSpreadsheetTimeZone('Asia/Tokyo');

  ensureSheet_(ss, SHEET_LOTS, LOT_HEADERS);
  ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
  const settings = ss.getSheetByName(SHEET_SETTINGS) || ensureSheet_(ss, SHEET_SETTINGS, ['項目', '値', '説明']);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, SETTING_ROWS.length, 3).setValues(SETTING_ROWS);
    settings.setColumnWidth(1, 110).setColumnWidth(2, 420).setColumnWidth(3, 420);
    settings.getRange(2, 2, SETTING_ROWS.length, 2).setWrap(true);
    settings.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['する', 'しない'], true).build());
  }
  // 後から増えた設定項目（標準質量など）を既存の設定シートに足す
  const have = settings.getLastRow() >= 2 ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]); }) : [];
  SETTING_ROWS.forEach(function(r) { if (have.indexOf(r[0]) < 0) settings.appendRow(r); });
  const workers = ss.getSheetByName(SHEET_WORKERS) || ensureSheet_(ss, SHEET_WORKERS, ['担当者名']);
  if (workers.getLastRow() < 2) workers.getRange('B1').setValue('← A2から下に検査スタッフの名前を1行1人で入力（入力画面で選べるようになります）');

  const msg = [];
  if (!ss.getSheetByName(SHEET_TEMPLATE)) {
    const source = TEMPLATE_SOURCE_CANDIDATES.map(function(n) { return ss.getSheetByName(n); })
      .filter(function(s) { return s; })[0];
    if (source) {
      source.copyTo(ss).setName(SHEET_TEMPLATE);
      msg.push('「' + source.getName() + '」から「' + SHEET_TEMPLATE + '」を作成しました。');
    } else {
      msg.push('【要対応】成績表の書式シートが見つかりません。元Excelの「検査成績表 (001-100)」シートを' +
        'このスプレッドシートにコピーし、名前を「' + SHEET_TEMPLATE + '」にしてください。');
    }
  } else {
    msg.push('「' + SHEET_TEMPLATE + '」は作成済みです。');
  }
  msg.push('セットアップが完了しました。');
  msg.push('「設定」シートの送付先と、「担当者」シートのスタッフ名を入力してください。');
  // エディタから実行したときは画面にダイアログを出せないので、実行ログに出す
  try { SpreadsheetApp.getUi().alert(msg.join('\n')); } catch (e) { console.log(msg.join('\n')); }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, 0);
  // 新しい列が増えた場合も見出しをそろえる（既存データはそのまま）
  const current = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (current.join('|') !== headers.join('|') && current.length <= headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1f2a5c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
