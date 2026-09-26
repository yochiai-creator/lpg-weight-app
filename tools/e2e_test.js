// E2Eテスト: 本物の gas/*.js を模擬スプレッドシート上で動かし、gas/Index.html をChromiumで開いて
// テンキー操作をシミュレートする。  実行: npm test
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { load } = require('./gas_mock.js');

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok: !!ok });
  console.log((ok ? 'PASS' : 'FAIL') + '  ' + name + (ok || detail === undefined ? '' : '  → ' + detail));
}

(async () => {
  const { ctx, ss } = load();
  let html = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Index.html'), 'utf8');
  // google.script.run をNode側のGAS関数に橋渡しする
  const stub = `<script>window.google={script:{run:new Proxy({},{get(t,p){ if(p==='withSuccessHandler') return function(ok){ return {withFailureHandler(fail){ return new Proxy({},{get(_,fn){return function(arg){ window.__gas(fn,JSON.stringify(arg===undefined?null:arg)).then(r=>{const o=JSON.parse(r); if(o.err) fail(new Error(o.err)); else setTimeout(()=>ok(o.v),10);});}}});}}};}})}};</script>`;
  html = html.replace('<head>', '<head>' + stub);

  const exe = '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
  const browser = await chromium.launch(fs.existsSync(exe) ? { executablePath: exe } : {});
  const page = await browser.newPage({ viewport: { width: 1180, height: 820 } });
  page.on('pageerror', e => check('ページでJSエラーなし', false, e.message));
  page.on('dialog', d => d.accept());
  await page.exposeFunction('__gas', (fn, arg) => {
    try { const v = ctx[fn](JSON.parse(arg)); return JSON.stringify({ v: JSON.parse(JSON.stringify(v === undefined ? null : v)) }); }
    catch (e) { return JSON.stringify({ err: e.message }); }
  });
  await page.setContent(html);
  await page.waitForTimeout(300);

  const KEY = { '.': 'NumpadDecimal', '+': 'NumpadAdd', '-': 'NumpadSubtract', '*': 'NumpadMultiply', '/': 'NumpadDivide', E: 'NumpadEnter' };
  const keys = async (seq, wait = 150) => {
    for (const c of seq) await page.keyboard.press(KEY[c] || 'Numpad' + c);
    await page.waitForTimeout(wait);
    return page.textContent('#notice');
  };
  const create = async (pre, st) => {
    await page.fill('#fPrefix', pre); await page.fill('#fStart', st); await page.selectOption('#fKind', '50kg');
    await page.click('#btnCreate'); await page.waitForTimeout(300);
    return page.textContent('#lotMsg');
  };

  // ---- ロット登録
  check('ロット登録', (await create('HEP', '37001')).includes('登録しました'));
  await create('HEP', '36001');
  check('同じ組容器番号は登録できない', (await create('HEP', '37001')).includes('登録済み'));
  check('担当者を選ぶまで入力開始できない', await page.isDisabled('#btnStart'));
  await page.selectOption('#fWorker', '山田'); await page.fill('#fDevice', 'iPad-1');
  await page.click('#btnStart'); await page.waitForTimeout(100);

  const sh = () => ss.getSheetByName('成績表_HEP37001-37100');
  const slot = i => sh().getRange(5 + i % 20, 2 + Math.floor(i / 20) * 6, 1, 6).getValues()[0];

  // ---- 容器番号の判定
  check('下3桁が2ロットに当たると4桁以上を要求', (await keys('23E')).includes('4桁以上'));
  await page.keyboard.press('Escape');
  check('入力桁数の表示は下4桁', (await page.textContent('#numHint')).includes('下4桁'));
  check('1本目から候補ボタン（50kgのいつもの値34.8が真ん中）', (await keys('7001'), (await page.locator('#cands button').count()) === 10 && (await page.getAttribute('#cands button:nth-child(6)', 'data-mass')) === '34.8'));
  check('1本目でも前回±ボタン（いつもの値基準）', (await page.getAttribute('#steps button[data-delta="0"]', 'data-mass')) === '34.8');
  check('5本入るまでは1桁入力しない（範囲表示なし）', !(await page.textContent('#massHint')).includes('1桁で保存'));
  await page.keyboard.press('Escape');
  check('4桁でEnterなしに記録先が決まる', (await keys('7023')).includes('質量を入力'));
  check('質量「348」を3桁でEnterなしに34.8として保存', (await keys('348', 400)).includes('34.8 kg'));
  check('成績表に 023 ☑ 3 4 , 8 が入る', slot(22).join('|') === '023|true|3|4|,|8', slot(22).join('|'));
  for (const n of ['7024', '7025', '7026', '7027']) await keys(n + '+');
  check('＋で前回と同じ質量', (await page.textContent('#notice')).includes('HEP37027　34.8'));
  check('標準から外れた質量は再確認', (await keys('7030.340')).includes('離れています'));
  check('再度Enterで確定', (await keys('E')).includes('34.0 kg'));
  check('標準が決まると1桁入力の範囲を表示', (await keys('7061'), (await page.textContent('#massHint')).includes('34.3〜35.2')));
  check('1桁「7」で34.7を保存', (await keys('7', 400)).includes('HEP37061　34.7'));
  check('1桁「0」は範囲内の35.0', (await keys('70620', 400)).includes('HEP37062　35.0'));
  check('「.」のあと3桁で範囲外も入力', (await keys('7063.353', 400)).includes('HEP37063　35.3'));
  await keys('7064');
  fs.mkdirSync(path.join(__dirname, 'out'), { recursive: true });
  await page.screenshot({ path: path.join(__dirname, 'out', 'candidates.png') });
  check('候補ボタンが10個出る', (await page.locator('#cands button').count()) === 10);
  check('一番多い値のボタンが濃い色', (await page.getAttribute('#cands button.freq1', 'data-mass')) === '34.8');
  await page.click('#cands button[data-mass="34.9"]'); await page.waitForTimeout(400);
  check('候補ボタンのタップで保存', (await page.textContent('#notice')).includes('HEP37064　34.9'));
  check('保存後は候補ボタンが消える', (await page.locator('#cands button').count()) === 0);
  await page.click('#grid .cell[data-serial="37065"]'); await page.waitForTimeout(100);
  check('ロットのタブは1行の横スクロール', await page.evaluate(() => { const t = getComputedStyle(document.getElementById('tabs')); return t.flexWrap === 'nowrap' && t.overflowX === 'auto'; }));
  check('マスのタップで容器を選べる', (await page.textContent('#resolved')).includes('HEP37065'));
  await page.click('#cands button[data-mass="34.8"]'); await page.waitForTimeout(400);
  check('マス→候補ボタンのタップ2回で保存', (await page.textContent('#notice')).includes('HEP37065　34.8'));
  await page.click('#grid .cell[data-serial="37065"]'); await page.waitForTimeout(100);
  check('入力済みのマスは上書き確認', (await page.textContent('#notice')).includes('もう一度タップ'));
  await page.click('#grid .cell[data-serial="37065"]'); await page.waitForTimeout(100);
  check('もう一度タップで上書きの質量入力へ', (await page.locator('#cands button').count()) === 10);
  await page.keyboard.press('Escape');
  await keys('7066');
  check('前回±0.1のボタン（前回34.8）', (await page.getAttribute('#steps button[data-delta="-1"]', 'data-mass')) === '34.7' && (await page.getAttribute('#steps button[data-delta="1"]', 'data-mass')) === '34.9');
  await page.click('#steps button[data-delta="1"]'); await page.waitForTimeout(400);
  check('前回＋0.1のタップで保存', (await page.textContent('#notice')).includes('HEP37066　34.9'));
  await keys('7067');
  check('次の「前回」はいま入れた34.9', (await page.getAttribute('#steps button[data-delta="0"]', 'data-mass')) === '34.9');
  await page.click('#steps button[data-delta="-1"]'); await page.waitForTimeout(400);
  check('前回−0.1のタップで保存', (await page.textContent('#notice')).includes('HEP37067　34.8'));
  check('入力済みは上書き確認', (await keys('7023')).includes('入力済み'));
  await keys('E.349');
  check('欠番', (await keys('7031-E', 400)).includes('欠番'));
  check('成績表に「欠 番」', slot(30).join('|') === '031|false|欠|番||', slot(30).join('|'));
  check('範囲外の番号はエラー', (await keys('9999')).includes('どのロットの範囲にもありません'));
  await page.keyboard.press('Escape');
  check('NG→実容器番号で保存', (await keys('7040/7041.347', 400)).includes('HEP37041'));
  check('質量の範囲外はエラー', (await keys('7050.0.0')).includes('0.1〜99.9'));
  await page.keyboard.press('Escape');
  // Enterで確定するモード
  await page.uncheck('#autoMode');
  check('自動確定オフでは4桁でも進まない', (await keys('7060')).includes('容器番号を入力') || !(await page.textContent('#notice')).includes('質量'));
  check('自動確定オフでもEnterで保存', (await keys('E348E', 400)).includes('HEP37060　34.8'));
  await page.check('#autoMode');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(500);
  await keys('*');
  check('直前の取消', (await keys('E', 500)).includes('取り消しました'));
  check('入力中は桁数が減らない（ロット完了前）', (await page.textContent('#numHint')).includes('下4桁'));
  check('取消でセルが空に戻る', slot(59)[2] === '' && slot(59)[1] === false, slot(59).join('|'));
  check('NGで入れたHEP37041は残る', slot(40)[1] === true, slot(40).join('|'));

  const log = ss.getSheetByName('入力記録');
  const rows = log.getRange(2, 1, log.getLastRow() - 1, 15).getValues();
  const ng = rows.find(r => r[3] === 'HEP37041');
  check('入力記録にNGと最初の番号が残る', ng && ng[7] === 'NG' && ng[8] === 'HEP37040');
  const undone = rows.find(r => r[3] === 'HEP37060');
  check('取消した行は「取消」で残る', undone && undone[11] === '取消' && undone[14] !== '');
  check('入力記録に担当者・端末', rows.every(r => r[12] === '山田' && r[13] === 'iPad-1'));
  const over = rows.filter(r => r[3] === 'HEP37023');
  check('上書き前の値を記録', over.length === 2 && over[1][9] === 34.8);

  // ---- 手動完了（送信なし）
  await page.click('#btnHome'); await page.waitForTimeout(300);
  await page.locator('#activeLots .card').first().locator('button').click();
  check('完了ダイアログに未入力本数', (await page.textContent('#dlgMissing')).includes('未入力'));
  await page.uncheck('#dSend'); await page.click('#dSave'); await page.waitForTimeout(400);
  const lots = () => ss.getSheetByName('ロット').getRange(2, 1, 2, 17).getDisplayValues();
  check('手動完了で状態=完了・PDF保存', lots()[0][6] === '完了' && lots()[0][13] !== '');
  check('送信しない指定ではメールなし', global.MAILS.length === 0);

  // ---- 全数そろったら自動完了・送信
  await page.click('#btnStart'); await page.waitForTimeout(100);
  check('ロットが1つなら入力桁数は下3桁', (await page.textContent('#numHint')).includes('下3桁'));
  for (let i = 1; i <= 100; i++) {
    for (const c of String(36000 + i).slice(-3)) await page.keyboard.press('Numpad' + c);
    if (i === 50) { await page.keyboard.press('NumpadSubtract'); await page.keyboard.press('NumpadEnter'); }
    else for (const c of (i <= 5 ? '347' : '7')) await page.keyboard.press('Numpad' + c);
  }
  await page.waitForTimeout(3000);
  check('全数そろうと画面に送信完了', (await page.textContent('#notice')).includes('全数そろいました'));
  const mail = global.MAILS[0] || {};
  check('メールを1通送信', global.MAILS.length === 1);
  check('宛先・件名', mail.to === 'nouhin@example.com' && (mail.sub || '').includes('HEP36001～HEP36100'));
  check('本文に本数と欠番', (mail.body || '').includes('本数：99本（欠番 1）'));
  check('本文に空の差し込み行が残らない', !(mail.body || '').includes('耐圧試験日'));
  check('PDFとCSVを添付', JSON.stringify(mail.att) === JSON.stringify(['成績表_HEP36001-36100.pdf', '成績表_HEP36001-36100.csv']));
  check('ロットに送信日時・送信先', lots()[1][6] === '完了' && lots()[1][15] !== '' && lots()[1][16] === 'nouhin@example.com');

  // ---- 範囲を入れてまとめて登録
  await page.click('#btnHome'); await page.waitForTimeout(300);
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '38050'); await page.fill('#fEnd', '38350');
  check('まとめ登録のプレビュー', (await page.textContent('#lotMsg')).includes('HEP38001〜HEP38400') && (await page.textContent('#btnCreate')) === '4ロット登録');
  await page.click('#btnCreate'); await page.waitForTimeout(1500);
  check('入力中ロット一覧は枠の中でスクロール', await page.evaluate(() => getComputedStyle(document.getElementById('activeLots')).overflowY === 'auto'));
  check('100本ずつ4ロット作成', (await page.textContent('#lotMsg')).includes('4ロット登録しました'));
  check('成績表シートが4枚できる', ['38001-38100', '38101-38200', '38201-38300', '38301-38400'].every(r => ss.getSheetByName('成績表_HEP' + r)));
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '38201'); await page.fill('#fEnd', '38500');
  await page.click('#btnCreate'); await page.waitForTimeout(1500);
  const m2 = await page.textContent('#lotMsg');
  check('登録済みは飛ばして残りだけ作成', m2.includes('1ロット登録しました（HEP38401）') && m2.includes('HEP38201') && m2.includes('HEP38301'));

  fs.mkdirSync(path.join(__dirname, 'out'), { recursive: true });
  await page.screenshot({ path: path.join(__dirname, 'out', 'home.png'), fullPage: true });
  await browser.close();

  const failed = results.filter(r => !r.ok).length;
  console.log('\n' + (results.length - failed) + ' / ' + results.length + ' PASS');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
