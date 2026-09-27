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
  check('スプレッドシートを日本時間にそろえる', ss.getSpreadsheetTimeZone() === 'Asia/Tokyo');
  check('旧「担当者」シートは「入力者」に名前が変わる', !!ss.getSheetByName('入力者') && !ss.getSheetByName('担当者') && ss.getSheetByName('入力者').getRange('A1').getValue() === '入力者名');
  const lh = ss.getSheetByName('入力記録').getRange(1, 1, 1, 16).getValues()[0];
  check('入力記録の見出し: Googleアカウント・入力者', lh[10] === 'Googleアカウント' && lh[12] === '入力者');
  check('ロット登録', (await create('HEP', '37001')).includes('登録しました'));
  await create('HEP', '36001');
  await page.fill('#fPrefix', ''); await page.fill('#fStart', '73301'); await page.selectOption('#fKind', '20kg');
  await page.click('#btnCreate'); await page.waitForTimeout(300);
  check('機種ボタンはロット0の機種も表示', (await page.textContent('#kindFilter')).includes('5kg（0）'));
  check('機種の切り替えボタン（すべて・20kg・50kg）', (await page.textContent('#kindFilter')).includes('すべて（3）') && (await page.textContent('#kindFilter')).includes('20kg（1）') && (await page.textContent('#kindFilter')).includes('50kg（2）'));
  check('すべてのときは機種ごとの見出し', (await page.locator('#activeLots .kind-head').count()) === 2);
  await page.click('#kindFilter button:has-text("20kg")');
  check('20kgを選ぶと20kgのロットだけ', (await page.locator('#activeLots .card').count()) === 1);
  await page.click('#kindFilter button:has-text("すべて")');
  const c73 = page.locator('#activeLots .card', { hasText: '73301' });
  await c73.locator('button:has-text("機種を変更")').click();
  await c73.locator('select').selectOption('30kg');
  await c73.locator('button:has-text("変更")').first().click(); await page.waitForTimeout(400);
  check('内容積が成績表に自動で入る（50kg=117.5）', String(ss.getSheetByName('成績表_HEP37001-37100').getRange('N3').getValue()) === '117.5');
  const sh73 = ss.getSheets().find(x => x.getName().startsWith('成績表_73301'));
  check('機種を変更すると内容積も変わる（30kg=70.5）', sh73 && String(sh73.getRange('N3').getValue()) === '70.5');
  check('機種を変更できる', (await page.textContent('#kindFilter')).includes('30kg（1）') && (await page.textContent('#kindFilter')).includes('20kg（0）'));
  check('ロットシートの容器区分も変わる', ss.getSheetByName('ロット').getRange(2, 1, 3, 6).getValues().some(r => r[0] === '73301' && r[5] === '30kg'));
  check('同じ組容器番号は登録できない', (await create('HEP', '37001')).includes('登録済み'));
  check('入力者を選ぶまで入力開始できない', await page.isDisabled('#btnStart'));
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
  check('質量「348」を3桁でEnterなしに34.8として保存', (await keys('348', 400)).includes('W34.8'));
  check('成績表に 023 ☑ 3 4 , 8 が入る', slot(22).join('|') === '023|true|3|4|,|8', slot(22).join('|'));
  for (const n of ['7024', '7025', '7026', '7027']) await keys(n + '+');
  check('＋で前回と同じ質量', (await page.textContent('#notice')).includes('HEP37027　W34.8'));
  check('標準から外れた質量は再確認', (await keys('7030.340')).includes('離れています'));
  check('再度Enterで確定', (await keys('E')).includes('W34.0'));
  check('標準が決まると1桁入力の範囲を表示', (await keys('7061'), (await page.textContent('#massHint')).includes('34.3〜35.2')));
  check('1桁「7」で34.7を保存', (await keys('7', 400)).includes('HEP37061　W34.7'));
  check('1桁「0」は範囲内の35.0', (await keys('70620', 400)).includes('HEP37062　W35.0'));
  check('「.」のあと3桁で範囲外も入力', (await keys('7063.353', 400)).includes('HEP37063　W35.3'));
  await keys('7064');
  fs.mkdirSync(path.join(__dirname, 'out'), { recursive: true });
  await page.screenshot({ path: path.join(__dirname, 'out', 'candidates.png') });
  check('候補ボタンが10個出る', (await page.locator('#cands button').count()) === 10);
  check('一番多い値のボタンが濃い色', (await page.getAttribute('#cands button.freq1', 'data-mass')) === '34.8');
  await page.click('#cands button[data-mass="34.9"]'); await page.waitForTimeout(400);
  check('候補ボタンのタップで保存', (await page.textContent('#notice')).includes('HEP37064　W34.9'));
  check('保存後は候補ボタンが消える', (await page.locator('#cands button').count()) === 0);
  await page.click('#grid .cell[data-serial="37065"]'); await page.waitForTimeout(100);
  check('ロットのタブは2段の横スクロール', await page.evaluate(() => { const t = getComputedStyle(document.getElementById('tabs')); return t.gridAutoFlow.startsWith('column') && t.gridTemplateRows.split(' ').length === 2 && t.overflowX === 'auto'; }));
  check('マスのタップで容器を選べる', (await page.textContent('#resolved')).includes('HEP37065'));
  await page.click('#cands button[data-mass="34.8"]'); await page.waitForTimeout(400);
  check('マス→候補ボタンのタップ2回で保存', (await page.textContent('#notice')).includes('HEP37065　W34.8'));
  await page.click('#grid .cell[data-serial="37065"]'); await page.waitForTimeout(100);
  check('入力済みのマスは上書きの理由を選ぶ', (await page.textContent('#notice')).includes('理由') && !(await page.getAttribute('#dupChoice', 'class')).includes('hidden'));
  await page.click('#grid .cell[data-serial="37065"]'); await page.waitForTimeout(100);
  check('もう一度タップで上書きの質量入力へ', (await page.locator('#cands button').count()) === 10);
  await page.keyboard.press('Escape');
  await keys('7066');
  check('前回±0.1のボタン（前回34.8）', (await page.getAttribute('#steps button[data-delta="-1"]', 'data-mass')) === '34.7' && (await page.getAttribute('#steps button[data-delta="1"]', 'data-mass')) === '34.9');
  await page.click('#steps button[data-delta="1"]'); await page.waitForTimeout(400);
  check('前回＋0.1のタップで保存', (await page.textContent('#notice')).includes('HEP37066　W34.9'));
  await keys('7067');
  check('次の「前回」はいま入れた34.9', (await page.getAttribute('#steps button[data-delta="0"]', 'data-mass')) === '34.9');
  await page.click('#steps button[data-delta="-1"]'); await page.waitForTimeout(400);
  check('前回−0.1のタップで保存', (await page.textContent('#notice')).includes('HEP37067　W34.8'));
  check('入力済みは上書き確認', (await keys('7023')).includes('入力済み'));
  check('ダブりは赤い点滅表示', (await page.getAttribute('#notice', 'class')).includes('dup') && (await page.textContent('#notice')).includes('ダブり'));
  await keys('+.349', 400);
  check('ダブったマスは赤枠＋「W」', (await page.getAttribute('#grid .cell[data-serial="37023"]', 'class')).includes('dup') && (await page.textContent('#grid .cell[data-serial="37023"]')).includes('W'));
  check('表の見出しにダブり件数', (await page.textContent('#gridSub')).includes('ダブり 1件'));
  check('開き直してもダブりが残る（サーバー記録）', ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').dups['HEP37023'] === 1);
  // 修正（入力ミスを直す）
  check('訂正: Enterで訂正として上書き', (await keys('7024E', 150)).includes('訂正'));
  await keys('0', 400);
  check('訂正は印を付けない', !(await page.textContent('#grid .cell[data-serial="37024"]')).includes('修') && !(await page.getAttribute('#grid .cell[data-serial="37024"]', 'class')).includes('dup'));
  // 修正（品質不良）
  await keys('7069');
  await page.click('#btnRepair'); await page.waitForTimeout(400);
  check('修正ボタンで保存', (await page.textContent('#notice')).includes('HEP37069　修正'));
  check('表のマスに「修正」', (await page.textContent('#grid .cell[data-serial="37069"]')).includes('修正') && (await page.getAttribute('#grid .cell[data-serial="37069"]', 'class')).includes('repair'));
  check('表の見出しに修正本数', (await page.textContent('#gridSub')).includes('修正 1本'));
  check('成績表のマスに「修 正」', slot(68).join('|') === '069|false|修|正||', slot(68).join('|'));
  // シール違い
  await keys('7068');
  await page.click('#btnSeal');
  check('シール違いボタンで印が付く', (await page.getAttribute('#btnSeal', 'class')).includes('on'));
  await page.click('#cands button[data-mass="34.8"]'); await page.waitForTimeout(400);
  check('保存メッセージにシール違い', (await page.textContent('#notice')).includes('【シール違い】'));
  check('シール違いは「シ」の印', (await page.textContent('#grid .cell[data-serial="37068"]')).includes('シ'));
  const logRows = () => { const lg = ss.getSheetByName('入力記録'); return lg.getRange(2, 1, lg.getLastRow() - 1, 16).getValues(); };
  const r69 = logRows().filter(r => r[3] === 'HEP37069').pop();
  check('入力記録: 修正の区分（質量なし）', r69[6] === '修正' && r69[5] === '');
  const r24 = logRows().filter(r => r[3] === 'HEP37024').pop(), r23 = logRows().filter(r => r[3] === 'HEP37023').pop(), r68 = logRows().filter(r => r[3] === 'HEP37068').pop();
  check('入力記録: 訂正の区分と上書き前', r24[6] === '訂正' && r24[9] === 34.8 && r24[5] === 35);
  check('入力記録: ダブりの区分', r23[6] === 'ダブり');
  check('入力記録: 備考にシール違い', r68[15] === 'シール違い' && r68[6] === '通常');
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
  check('自動確定オフでもEnterで保存', (await keys('E348E', 400)).includes('HEP37060　W34.8'));
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
  check('入力記録に入力者・端末', rows.every(r => r[12] === '山田' && r[13] === 'iPad-1'));
  const over = rows.filter(r => r[3] === 'HEP37023');
  check('上書き前の値を記録', over.length === 2 && over[1][9] === 34.8 && over[1][6] === 'ダブり');

  // ---- 入力ミスを消す
  await page.click('#grid .cell[data-serial="37090"]'); await page.waitForTimeout(150);
  check('空のマスには「この入力を消す」を出さない', !(await page.isVisible('#btnClear')));
  await page.click('#cands button >> nth=4'); await page.waitForTimeout(400);
  const hasBefore = ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').entries['37090'] !== undefined;
  await page.click('#grid .cell[data-serial="37090"]'); await page.waitForTimeout(150);
  await page.screenshot({ path: path.join(__dirname, 'out', 'clear.png') });
  check('入力済みのマスをタップすると「入力を消す」が出る', hasBefore && await page.isVisible('#btnDel'));
  await page.click('#btnDel'); await page.waitForTimeout(400);
  check('消すとマスが空に戻る', ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').entries['37090'] === undefined && !(await page.textContent('#grid .cell[data-serial="37090"]')).includes('W'));
  const hepSh = ss.getSheetByName('成績表_HEP37001-37100');
  check('消すと成績表の質量欄も空', hepSh.getRange(5 + 89 % 20, 2 + Math.floor(89 / 20) * 6, 1, 6).getValues()[0].slice(2).join('') === ',');
  const cl = ss.getSheetByName('入力記録').getRange(2, 1, ss.getSheetByName('入力記録').getLastRow() - 1, 16).getValues().filter(r => r[3] === 'HEP37090');
  check('消した記録は「取消」で残る', cl.length >= 1 && cl.every(r => r[11] === '取消'));

  // ---- タブレット縦向き: 上に表・下に入力
  await page.setViewportSize({ width: 820, height: 1180 }); await page.waitForTimeout(200);
  await keys('7070');
  const gridBox = await page.locator('#grid').boundingBox(), padBox = await page.locator('#pad').boundingBox(), numBox = await page.locator('#fldNum').boundingBox();
  check('縦向きは表が上・入力が下', gridBox.y + gridBox.height <= numBox.y + 1);
  check('入力欄は横に並ぶ（番号欄の右にテンキー）', padBox.x > numBox.x + numBox.width && Math.abs(padBox.y - numBox.y) < 80);
  const sh1 = await page.evaluate(() => document.documentElement.scrollHeight); console.log('   scrollHeight', sh1);
  const sheetBox = await page.locator('#massSheet').boundingBox(), inBox = await page.locator('#inPanel').boundingBox();
  check('番号が決まると入力バーが出る', (await page.getAttribute('#massSheet', 'class')).includes('open'));
  check('入力バーは入力欄に重ならない', sheetBox.y + sheetBox.height <= inBox.y + 2);
  check('縦向きでも1画面に収まる', sh1 <= 1180 + 40);
  await page.screenshot({ path: path.join(__dirname, 'out', 'portrait.png') });
  const xBox = await page.locator('#btnSheetClose').boundingBox();
  check('入力バーの右上に×がある', xBox && xBox.x + xBox.width > sheetBox.x + sheetBox.width - 60 && xBox.y < sheetBox.y + 50);
  await page.click('#btnSheetClose'); await page.waitForTimeout(100);
  check('×で入力バーが閉じる', !(await page.getAttribute('#massSheet', 'class') || '').includes('open'));
  await page.keyboard.press('Escape');
  check('番号入力中は入力バーを出さない', !(await page.getAttribute('#massSheet', 'class') || '').includes('open'));
  await page.setViewportSize({ width: 1180, height: 820 }); await page.waitForTimeout(200);

  // ---- 手動完了（送信なし）
  await page.click('#btnHome'); await page.waitForTimeout(300);
  await page.locator('#activeLots .card', { hasText: 'HEP37001' }).locator('button:has-text("完了にする")').click();
  check('完了ダイアログに未入力本数', (await page.textContent('#dlgMissing')).includes('未入力'));
  await page.uncheck('#dSend'); await page.click('#dSave'); await page.waitForTimeout(400);
  const lots = () => ss.getSheetByName('ロット').getRange(2, 1, 2, 17).getDisplayValues();
  check('手動完了で状態=完了・PDF保存', lots()[0][6] === '完了' && lots()[0][13] !== '');
  check('送信しない指定ではメールなし', global.MAILS.length === 0);

  // ---- 全数そろったら自動完了・送信
  await page.click('#btnStart'); await page.waitForTimeout(100);
  check('番号が重ならなければ入力桁数は下3桁', (await page.textContent('#numHint')).includes('下3桁'));
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
  check('本文に本数と欠番', (mail.body || '').includes('本数：99本（欠番 1・修正 0）'));
  check('本文に空の差し込み行が残らない', !(mail.body || '').includes('耐圧試験日'));
  check('PDFとCSVを添付', JSON.stringify(mail.att) === JSON.stringify(['成績表_HEP36001-36100.pdf', '成績表_HEP36001-36100.csv']));
  check('ロットに送信日時・送信先', lots()[1][6] === '完了' && lots()[1][15] !== '' && lots()[1][16] === 'nouhin@example.com');

  // ---- 範囲を入れてまとめて登録
  await page.click('#btnHome'); await page.waitForTimeout(300);
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '38050'); await page.fill('#fEnd', '38350');
  check('まとめ登録のプレビュー', (await page.textContent('#lotMsg')).includes('HEP38001〜HEP38400') && (await page.textContent('#btnCreate')) === '4ロット登録');
  await page.click('#btnCreate'); await page.waitForTimeout(1500);
  check('入力中ロット一覧は枠の中でスクロール', await page.evaluate(() => getComputedStyle(document.getElementById('activeLots')).overflowY === 'auto'));
  const titleOk = (sheet) => { const rt = ss.getSheetByName(sheet).getRange('A1').getRichText(); if (!rt) return false;
    const cut = rt.text.indexOf('※'); return rt.runs.some(r => r.a === 0 && r.z === cut && r.size === 15 && r.bold) && rt.runs.some(r => r.a === cut && r.z === rt.text.length && r.size === 11 && r.bold); };
  check('見出し: タイトル15pt・※の行11pt（新しい成績表）', titleOk('成績表_HEP38001-38100'));
  check('見出し: 書式シートも同じ', titleOk('書式_成績表'));
  check('100本ずつ4ロット作成', (await page.textContent('#lotMsg')).includes('4ロット登録しました'));
  check('成績表シートが4枚できる', ['38001-38100', '38101-38200', '38201-38300', '38301-38400'].every(r => ss.getSheetByName('成績表_HEP' + r)));
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '38201'); await page.fill('#fEnd', '38500');
  await page.click('#btnCreate'); await page.waitForTimeout(1500);
  const m2 = await page.textContent('#lotMsg');
  check('登録済みは飛ばして残りだけ作成', m2.includes('1ロット登録しました（HEP38401）') && m2.includes('HEP38201') && m2.includes('HEP38301'));

  // ---- 登録間違いの削除
  await page.fill('#fPrefix', 'ZZ'); await page.fill('#fStart', '99901'); await page.fill('#fEnd', ''); await page.selectOption('#fSpec', '');
  await page.click('#btnCreate'); await page.waitForTimeout(500);
  const lotsBefore = ss.getSheetByName('ロット').getLastRow();
  await page.locator('#activeLots .card', { hasText: 'ZZ99901' }).locator('button:has-text("削除")').click(); await page.waitForTimeout(500);
  check('削除でロットが一覧から消える', !(await page.textContent('#activeLots')).includes('ZZ99901'));
  check('削除でロットの行と成績表シートが消える', ss.getSheetByName('ロット').getLastRow() === lotsBefore - 1 && !ss.getSheetByName('成績表_ZZ99901-99999'));

  // ---- 底黒（再搬入）: 登録欄で「底黒」を選んで登録し、流れた容器の質量欄に〇
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '38101'); await page.fill('#fEnd', ''); await page.selectOption('#fSpec', '底黒');
  await page.click('#btnCreate'); await page.waitForTimeout(500);
  check('入力中の通常ロットと番号が重なる底黒は登録できない', (await page.textContent('#lotMsg')).includes('番号が重なります'));
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '37001'); await page.fill('#fEnd', ''); await page.selectOption('#fSpec', '底黒');
  await page.click('#btnCreate'); await page.waitForTimeout(800);
  await page.selectOption('#fSpec', '');
  check('底黒ロットができる（入力中に【底黒】）', (await page.textContent('#activeLots')).includes('【底黒】HEP37001'));
  const sokoSh = ss.getSheetByName('成績表_HEP37001-37100_底黒');
  check('底黒の成績表は別シート（組容器番号に（底黒））', !!sokoSh && String(sokoSh.get(2, 22)).includes('（底黒）'));
  await page.click('#btnStart'); await page.waitForTimeout(200);
  check('底黒: 番号を入れるとその場で〇', (await keys('7023', 400)).includes('〇: HEP37023（底黒）'));
  check('底黒: 質量の入力バーは出ない', !(await page.getAttribute('#massSheet', 'class') || '').includes('open'));
  const sokoSlot = sokoSh.getRange(5 + 22 % 20, 2 + Math.floor(22 / 20) * 6, 1, 6).getValues()[0].join('|');
  check('底黒: 成績表の質量欄に〇', sokoSlot === '023|true||〇||', sokoSlot);
  check('底黒: もう一度流れても二重に記録しない', (await keys('7023', 400)).includes('〇済み'));
  const sokoRows = ss.getSheetByName('入力記録').getRange(2, 1, ss.getSheetByName('入力記録').getLastRow() - 1, 16).getValues().filter(r => r[2] === 'HEP37001-底黒');
  check('底黒: 入力記録は区分「底黒」で1行', sokoRows.length === 1 && sokoRows[0][6] === '底黒' && sokoRows[0][5] === '');
  check('元の成績表の質量はそのまま', slot(22).join('|') === '023|true|3|4|,|9', slot(22).join('|'));
  await page.click('#btnHome'); await page.waitForTimeout(300);

  fs.mkdirSync(path.join(__dirname, 'out'), { recursive: true });
  await page.screenshot({ path: path.join(__dirname, 'out', 'home.png'), fullPage: true });
  // ---- 起動高速化: 2回目の起動は前回の一覧をすぐ出し、裏で最新に更新する
  {
    const p2 = await browser.newPage({ viewport: { width: 1180, height: 820 } });
    p2.on('pageerror', e => check('ページでJSエラーなし(再起動)', false, e.message));
    let slow = 0;
    await p2.exposeFunction('__gas', async (fn, arg) => {
      if (fn === 'getBootstrap' && slow) await new Promise(r => setTimeout(r, slow));
      try { const v = ctx[fn](JSON.parse(arg)); return JSON.stringify({ v: JSON.parse(JSON.stringify(v === undefined ? null : v)) }); }
      catch (e) { return JSON.stringify({ err: e.message }); }
    });
    await p2.route('http://lpg.test/', r => r.fulfill({ contentType: 'text/html', body: html }));
    await p2.goto('http://lpg.test/'); await p2.waitForTimeout(400);
    const n1 = await p2.locator('#activeLots .card').count();
    slow = 1500;
    await p2.reload(); await p2.waitForTimeout(300);
    check('再起動: 前回の一覧がすぐ出る', n1 > 0 && (await p2.locator('#activeLots .card').count()) === n1);
    check('再起動: 更新中の表示が出る', await p2.isVisible('#staleMsg'));
    await p2.waitForTimeout(1600);
    check('再起動: 最新に更新されると表示が消える', !(await p2.isVisible('#staleMsg')) && (await p2.locator('#activeLots .card').count()) === n1);
    await p2.close();
  }
  await browser.close();

  const failed = results.filter(r => !r.ok).length;
  console.log('\n' + (results.length - failed) + ' / ' + results.length + ' PASS');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
