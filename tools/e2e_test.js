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
  page.on('dialog', d => { check('ダイアログを使わない', false, d.message()); d.accept(); });
  await page.exposeFunction('__gas', (fn, arg) => {
    try { const v = ctx[fn](JSON.parse(arg)); return JSON.stringify({ v: JSON.parse(JSON.stringify(v === undefined ? null : v)) }); }
    catch (e) { return JSON.stringify({ err: e.message }); }
  });
  await page.setContent(html);
  await page.waitForTimeout(300);
  // 製造年月OKの必須は専用のテストで確かめる（ほかのテストは質量をすぐ入れる）
  const noDateOk = async (pg) => pg.evaluate(() => { const c = document.getElementById('dateOkMode'); c.checked = false; c.dispatchEvent(new Event('change')); });
  await noDateOk(page);

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
  const lh = ss.getSheetByName('入力記録').getRange(1, 1, 1, 18).getValues()[0];
  check('入力記録の見出し: Googleアカウント・入力者', lh[12] === 'Googleアカウント' && lh[14] === '入力者');
  check('ロット登録', (await create('HEP', '37001')).includes('登録しました'));
  await create('HEP', '36001');
  await page.fill('#fPrefix', ''); await page.fill('#fStart', '73301'); await page.selectOption('#fKind', '20kg');
  await page.click('#btnCreate'); await page.waitForTimeout(300);
  check('機種ボタンはロット0の機種も表示', (await page.textContent('#kindFilter')).includes('5kg（0）'));
  check('機種の切り替えボタン（すべて・20kg・50kg）', (await page.textContent('#kindFilter')).includes('すべて（3）') && (await page.textContent('#kindFilter')).includes('20kg（1）') && (await page.textContent('#kindFilter')).includes('50kg（2）'));
  check('すべてのときは機種ごとの見出し', (await page.locator('#activeLots .kind-head').count()) === 2);
  await page.click('#kindFilter button:has-text("20kg")');
  check('20kgを選ぶと20kgのロットだけ', (await page.locator('#activeLots .card').count()) === 1);
  check('20kgのグループNoは50本ごと（73301〜73400 → 467・468）', (await page.textContent('#activeLots .card .ln-grp')) === '467・468');
  await page.click('#kindFilter button:has-text("すべて")');
  const c73 = page.locator('#activeLots .card', { hasText: '73301' });
  await c73.locator('button:has-text("機種を変更")').click();
  await c73.locator('select').selectOption('30kg');
  await c73.locator('button:has-text("変更")').first().click(); await page.waitForTimeout(400);
  check('内容積が成績表に自動で入る（50kg=118）', String(ss.getSheetByName('成績表_HEP37001-37100').getRange('N3').getValue()) === '118');
  const sh73 = ss.getSheets().find(x => x.getName().startsWith('成績表_73301'));
  check('機種を変更すると内容積も変わる（30kg=71）', sh73 && String(sh73.getRange('N3').getValue()) === '71');
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
  check('入力済みのマスは警告なしでそのまま入力バーを開く', (await page.textContent('#notice')).includes('入力済み') && !(await page.getAttribute('#notice', 'class')).includes('dup') && (await page.locator('#cands button').count()) === 10);
  await page.keyboard.press('Escape');
  await keys('7066');
  check('前回±0.1のボタン（前回34.8）', (await page.getAttribute('#steps button[data-delta="-1"]', 'data-mass')) === '34.7' && (await page.getAttribute('#steps button[data-delta="1"]', 'data-mass')) === '34.9');
  await page.click('#steps button[data-delta="1"]'); await page.waitForTimeout(400);
  check('前回＋0.1のタップで保存', (await page.textContent('#notice')).includes('HEP37066　W34.9'));
  await keys('7067');
  check('次の「前回」はいま入れた34.9', (await page.getAttribute('#steps button[data-delta="0"]', 'data-mass')) === '34.9');
  await page.click('#steps button[data-delta="-1"]'); await page.waitForTimeout(400);
  check('前回−0.1のタップで保存', (await page.textContent('#notice')).includes('HEP37067　W34.8'));
  check('入力済みの番号は警告なし（入力済みと前のWを表示）', (await keys('7023')).includes('入力済み') && !(await page.getAttribute('#notice', 'class')).includes('dup'));
  await page.click('#btnDupMass'); await page.waitForTimeout(500);
  check('入力バーの「ダブり」で前と同じWのまま記録', (await page.textContent('#notice')).includes('【ダブり】HEP37023　W34.8'));
  check('ダブったマスは赤枠＋「W」', (await page.getAttribute('#grid .cell[data-serial="37023"]', 'class')).includes('dup') && (await page.textContent('#grid .cell[data-serial="37023"]')).includes('W'));
  check('表の見出しにダブり件数', (await page.textContent('#gridSub')).includes('ダブり 1件'));
  check('開き直してもダブりが残る（サーバー記録）', ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').dups['HEP37023'] === 1);
  // 修正（入力ミスを直す）
  check('訂正: 入力済みの番号に質量を入れ直すと訂正', (await keys('7024', 150)).includes('入力済み'));
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
  const logRows = () => { const lg = ss.getSheetByName('入力記録'); return lg.getRange(2, 1, lg.getLastRow() - 1, 18).getValues(); };
  { const r = logRows().find(x => x[4] === 'HEP37024'); check('入力記録にグループNoと容器区分（HEP37024・50kg → 371）', r && r[3] === 371 && r[5] === '50kg', r && [r[3], r[5]].join('|')); }
  { const g = ctx.groupNoOf_; check('グループNoの計算（20kg 59701→195・3001→61・95001→901 ／ 50kg 59800→598）', g('20kg', '59701') === 195 && g('20kg', '62050') === 241 && g('8kg', '50051') === 2 && g('50kg S付', '59800') === 598 && g('20kg', '3001') === 61 && g('20kg', '48001') === 961 && g('20kg', '95001') === 901 && g('5kg', '1') === 1 && g('20kg', '49951') === 0 && g('20kg', '50000') === 0 && g('20kg', '99999') === 0); }
  const r69 = logRows().filter(r => r[4] === 'HEP37069').pop();
  check('入力記録: 修正の区分（質量なし）', r69[8] === '修正' && r69[7] === '');
  const r24 = logRows().filter(r => r[4] === 'HEP37024').pop(), r23 = logRows().filter(r => r[4] === 'HEP37023').pop(), r68 = logRows().filter(r => r[4] === 'HEP37068').pop();
  check('入力記録: 訂正の区分と上書き前', r24[8] === '訂正' && r24[11] === 34.8 && r24[7] === 35);
  check('入力記録: ダブりの区分', r23[8] === 'ダブり');
  check('入力記録: 備考にシール違い', r68[17] === 'シール違い' && r68[8] === '通常');
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
  check('直前の取消（＊）は確認なしですぐ取り消す', (await keys('*', 500)).includes('取り消しました'));
  check('入力中は桁数が減らない（ロット完了前）', (await page.textContent('#numHint')).includes('下4桁'));
  check('取消でセルが空に戻る', slot(59)[2] === '' && slot(59)[1] === false, slot(59).join('|'));
  check('NGで入れたHEP37041は残る', slot(40)[1] === true, slot(40).join('|'));

  const log = ss.getSheetByName('入力記録');
  const rows = log.getRange(2, 1, log.getLastRow() - 1, 18).getValues();
  const ng = rows.find(r => r[4] === 'HEP37041');
  check('入力記録にNGと最初の番号が残る', ng && ng[9] === 'NG' && ng[10] === 'HEP37040');
  const undone = rows.find(r => r[4] === 'HEP37060');
  check('取消した行は「取消」で残る', undone && undone[13] === '取消' && undone[16] !== '');
  check('入力記録に入力者・端末', rows.every(r => r[14] === '山田' && r[15] === 'iPad-1'));
  const over = rows.filter(r => r[4] === 'HEP37023');
  check('上書き前の値を記録', over.length === 2 && over[1][11] === 34.8 && over[1][8] === 'ダブり');

  check('グループNoは100本ごとの組の番号（HEP37001〜37100 → 371）', (await page.textContent('#gridTitle')).startsWith('HEP371') || (await page.locator('#tabs button', { hasText: 'HEP371' }).count()) === 1);
  // ---- 表を左右スワイプ／‹ › でロット切替
  const title0 = await page.textContent('#gridTitle');
  await page.click('#lotNext'); await page.waitForTimeout(150);
  const title1 = await page.textContent('#gridTitle');
  check('› で次のロットの表に切り替わる', title1 !== title0);
  await page.click('#lotPrev'); await page.waitForTimeout(150);
  check('‹ で前のロットに戻る', (await page.textContent('#gridTitle')) === title0);
  await page.evaluate(() => {
    const g = document.getElementById('grid');
    const mk = (type, x) => { const t = new Touch({ identifier: 1, target: g, clientX: x, clientY: 300 }); g.dispatchEvent(new TouchEvent(type, { touches: type === 'touchend' ? [] : [t], changedTouches: [t], bubbles: true, cancelable: true })); };
    mk('touchstart', 700); mk('touchend', 500);
  });
  await page.waitForTimeout(150);
  check('表を左にスワイプすると次のロット', (await page.textContent('#gridTitle')) === title1);
  await page.click('#lotPrev'); await page.waitForTimeout(150);

  // ---- 入力ミスを消す
  await page.click('#grid .cell[data-serial="37090"]'); await page.waitForTimeout(150);
  check('空のマスには「この入力を消す」を出さない', !(await page.isVisible('#btnClear')));
  await page.click('#cands button >> nth=4'); await page.waitForTimeout(400);
  const hasBefore = ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').entries['37090'] !== undefined;
  await page.click('#grid .cell[data-serial="37090"]'); await page.waitForTimeout(150);
  await page.screenshot({ path: path.join(__dirname, 'out', 'clear.png') });
  check('入力済みのマスをタップすると「この入力を消す」が出る', hasBefore && await page.isVisible('#btnClear'));
  await page.click('#btnClear'); await page.waitForTimeout(400);
  check('消すとマスが空に戻る', ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').entries['37090'] === undefined && !(await page.textContent('#grid .cell[data-serial="37090"]')).includes('W'));
  const hepSh = ss.getSheetByName('成績表_HEP37001-37100');
  check('消すと成績表の質量欄も空', hepSh.getRange(5 + 89 % 20, 2 + Math.floor(89 / 20) * 6, 1, 6).getValues()[0].slice(2).join('') === ',');
  const cl = ss.getSheetByName('入力記録').getRange(2, 1, ss.getSheetByName('入力記録').getLastRow() - 1, 18).getValues().filter(r => r[4] === 'HEP37090');
  check('消した記録は「取消」で残る', cl.length >= 1 && cl.every(r => r[13] === '取消'));

  await page.click('#grid .cell[data-serial="37069"]'); await page.waitForTimeout(150);
  await page.click('#btnClear'); await page.waitForTimeout(400);
  check('修正（品質不良）も消せる', ctx.getBootstrap().lots.find(l => l.lotId === 'HEP37001').entries['37069'] === undefined);
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
  await page.uncheck('#dSend'); await page.click('#dSave'); await page.waitForTimeout(150);
  check('未入力があると完了は2回押し', (await page.textContent('#dlgMsg')).includes('もう一度押して'));
  await page.click('#dSave'); await page.waitForTimeout(400);
  const lots = () => ss.getSheetByName('ロット').getRange(2, 1, 2, 17).getDisplayValues();
  check('手動完了で状態=完了・PDF保存', lots()[0][6] === '完了' && lots()[0][13] !== '');
  check('送信しない指定ではメールなし', global.MAILS.length === 0);
  check('完了してPDFを保存したら成績表シートを消す', !ss.getSheetByName('成績表_HEP37001-37100'));
  { const y = String(new Date().getFullYear()), mo = ('0' + (new Date().getMonth() + 1)).slice(-2);
    check('PDFは年・月・機種のフォルダに保存', global.PDFS.some(p => p === '/LPG容器 質量入力/成績表PDF/' + y + '年/' + mo + '月/50kg/成績表_HEP37001-37100.pdf'), global.PDFS.join(',')); }
  check('シートを消した完了ロットも入力状況は入力記録から出る', Object.keys(ctx.getBootstrap().recentDone.find(l => l.lotId === 'HEP37001').entries).length >= 15);

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
  const zz = page.locator('#activeLots .card', { hasText: 'ZZ99901' });
  await zz.locator('button:has-text("削除")').click(); await page.waitForTimeout(200);
  check('ロット削除: 1回目はボタンが「もう一度押すと削除」になるだけ', (await page.textContent('#activeLots')).includes('ZZ99901') && (await zz.locator('button:has-text("もう一度押すと削除")').count()) === 1);
  await zz.locator('button:has-text("もう一度押すと削除")').click(); await page.waitForTimeout(500);
  check('削除でロットが一覧から消える', !(await page.textContent('#activeLots')).includes('ZZ99901'));
  check('削除でロットの行と成績表シートが消える', ss.getSheetByName('ロット').getLastRow() === lotsBefore - 1 && !ss.getSheetByName('成績表_ZZ99901-99999'));

  // ---- 底黒（再搬入）: 登録欄で「底黒」を選んで登録し、流れた容器の質量欄に〇
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '38101'); await page.fill('#fEnd', ''); await page.selectOption('#fSpec', '底黒');
  await page.click('#btnCreate'); await page.waitForTimeout(500);
  check('入力中の通常ロットと番号が重なる底黒は登録できない', (await page.textContent('#lotMsg')).includes('番号が重なります'));
  await page.fill('#fPrefix', 'HEP'); await page.fill('#fStart', '37001'); await page.fill('#fEnd', ''); await page.selectOption('#fSpec', '底黒');
  await page.click('#btnCreate'); await page.waitForTimeout(800);
  await page.selectOption('#fSpec', '');
  check('底黒ロットができる（入力中に【底黒】）', (await page.locator('#activeLots .card', { hasText: 'HEP37001-底黒' }).count() + await page.locator('#activeLots .card:has(.ln-soko)', { hasText: 'HEP37001' }).count()) >= 1);
  const sokoSh = ss.getSheetByName('成績表_HEP37001-37100_底黒');
  check('底黒の成績表は別シート（組容器番号に（底黒））', !!sokoSh && String(sokoSh.get(2, 22)).includes('（底黒）'));
  await page.click('#btnStart'); await page.waitForTimeout(200);
  check('底黒: 番号を入れるとその場で〇', (await keys('7023', 400)).includes('〇: HEP37023（底黒）'));
  check('底黒: 質量の入力バーは出ない', !(await page.getAttribute('#massSheet', 'class') || '').includes('open'));
  const sokoSlot = sokoSh.getRange(5 + 22 % 20, 2 + Math.floor(22 / 20) * 6, 1, 6).getValues()[0].join('|');
  check('底黒: 成績表の質量欄に〇', sokoSlot === '023|true||〇||', sokoSlot);
  check('底黒: もう一度流れても二重に記録しない', (await keys('7023', 400)).includes('〇済み'));
  const sokoRows = ss.getSheetByName('入力記録').getRange(2, 1, ss.getSheetByName('入力記録').getLastRow() - 1, 18).getValues().filter(r => r[2] === 'HEP37001-底黒');
  check('底黒: 入力記録は区分「底黒」で1行', sokoRows.length === 1 && sokoRows[0][8] === '底黒' && sokoRows[0][7] === '');
  ctx.reopenLot('HEP37001');
  check('再開すると成績表シートを入力記録から作り直す', !!sh() && slot(22).join('|') === '023|true|3|4|,|8' && slot(30).join('|') === '031|false|欠|番||', slot(22).join('|') + ' / ' + (sh() ? slot(30).join('|') : ''));
  { const before = global.MAILS.length; const r = ctx.resendLot('HEP36001');
    check('シートを消したロットの再送信は保存済みPDFを送る', global.MAILS.length === before + 1 && r.pdfUrl.includes('drive.google.com') && !ss.getSheetByName('成績表_HEP36001-36100') && global.MAILS[before].att[0] === '成績表_HEP36001-36100.pdf'); }
  { const t = ss.getSheetByName('書式_成績表'); check('書式シートの余った行・列を削る', t.getMaxRows() <= 34 && t.getMaxColumns() <= 34, t.getMaxRows() + 'x' + t.getMaxColumns()); }
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
  // ---- v43までの並びの入力記録を、グループNo→容器番号→容器区分の並びに移す
  {
    const m = load(), lg = m.ss.getSheetByName('入力記録') || m.ss.insertSheet('入力記録');
    const old = ['記録ID','入力日時','ロットID','容器番号','表示番号','質量(kg)','区分','一致結果','入力番号(NG時)','上書き前','Googleアカウント','状態','入力者','端末','取消日時','備考','グループNo','容器区分'];
    lg.getRange(1, 1, 1, 18).setValues([old]);
    lg.getRange(2, 1, 1, 18).setValues([['R1', 't', 'HEP59701', 'HEP59723', '023', 34.8, '通常', 'OK', '', '', 'a@b', '有効', '山田', 'iPad', '', '', 598, '50kg']]);
    m.ctx.ensureLogHeaders_();
    const h = lg.getRange(1, 1, 1, 18).getValues()[0], d = lg.getRange(2, 1, 1, 18).getValues()[0];
    check('入力記録の並び替え: グループNo・容器番号・容器区分', h.slice(2, 7).join('|') === 'ロットID|グループNo|容器番号|容器区分|表示番号' && d.slice(2, 8).join('|') === 'HEP59701|598|HEP59723|50kg|023|34.8', h.join('|') + ' / ' + d.join('|'));
    m.ctx.ensureLogHeaders_();
    check('並び替えは1回だけ', lg.getRange(1, 4).getValue() === 'グループNo' && lg.getRange(2, 5).getValue() === 'HEP59723');
  }
  // ---- 入力記録の整理: 使わない列を削る・年が変わったら終わったロットの記録を別のスプレッドシートへ移す
  {
    const m = load(), c = m.ctx;
    c.setup();
    const A = c.createLot({ prefix: 'AR', start: '00101', kind: '20kg' }), B = c.createLot({ prefix: 'AR', start: '00201', kind: '20kg' });
    c.recordEntry({ lotId: A.lotId, serial: '00101', mass: 16.7, worker: '山田' });
    c.recordEntry({ lotId: A.lotId, serial: '00102', mass: null, missing: true, worker: '山田' });
    c.recordEntry({ lotId: B.lotId, serial: '00201', mass: 16.9, worker: '山田' });
    c.completeLot({ lotId: A.lotId, send: false });
    const lg = m.ss.getSheetByName('入力記録');
    c.runMaintenance();
    check('整理: 入力記録の使わない列を削る', lg.getMaxColumns() === 18, lg.getMaxColumns());
    check('整理: 年が変わるまでは移さない', lg.getLastRow() === 4 && global.ARCHIVES.length === 0);
    const P = c.sharedProps_(), y = new Date().getFullYear();
    P.setProperty('LOG_ARCHIVE_DONE', String(y - 2));   // 去年の分がまだ移っていない状態にする
    let r, k = 0; do { r = c.runMaintenance(); } while (r.archiveDone === false && ++k < 10);
    const rest = lg.getRange(2, 1, lg.getLastRow() - 1, 18).getValues();
    check('整理: 入力中のロットの記録は残す', rest.length === 1 && rest[0][2] === B.lotId && rest[0][4] === 'AR00201');
    const ar = global.ARCHIVES[0], ash = ar && ar.getSheets()[0];
    check('整理: 終わったロットは「LPG容器 入力記録_去年」へ移す', ar && ar.name === 'LPG容器 入力記録_' + (y - 1) && ash.getLastRow() === 3 && ash.getRange(2, 5).getValue() === 'AR00101' && ash.getRange(2, 7).getValue() === '101');
    check('整理: 移し終えたら次の年まで動かない', P.getProperty('LOG_ARCHIVE_DONE') === String(y - 1) && c.runMaintenance().archived === 0);
    const csv = c.buildLotCsv_(c.findLot_(A.lotId)).content || '';
    check('整理: 移したロットの再送信用CSVも作れる', String(csv).includes('"AR00101","101","16.7"'), String(csv).slice(0, 200));
    c.reopenLot(A.lotId);
    const back = lg.getRange(2, 1, lg.getLastRow() - 1, 18).getValues().filter(x => x[2] === A.lotId);
    const shA = m.ss.getSheetByName(c.findLot_(A.lotId).sheetName);
    check('整理: 移したロットを再開すると記録を戻して成績表を作り直す', back.length === 2 && !!shA && shA.getRange(5, 2, 2, 6).getValues().map(x => x.join('|')).join(' / ') === '101|true|1|6|,|7 / 102|false|欠|番||', shA && shA.getRange(5, 2, 2, 6).getValues().map(x => x.join('|')).join(' / '));
    check('整理: 移し先の行は「戻し済み」', ash.getRange(2, 14).getValue() === '戻し済み');
  }
  // ---- 採番表: その日に流れた順番のPDF
  {
    const m = load(), c = m.ctx;
    c.setup();
    ['74101', '74201', '74301'].forEach(st => c.createLot({ prefix: 'HXF', start: st, kind: '20kg' }));
    const order = [];
    for (let i = 0; i < 250; i++) { const n = 74101 + ((i * 37) % 300); order.push(String(n)); }
    order.forEach((n, i) => c.recordEntry({ lotId: c.readLots_().find(l => Number(l.start) <= +n && +n <= Number(l.end)).lotId, serial: n, mass: 16.7, worker: i < 100 ? '山崎' : '田中' }));
    const lot1 = c.readLots_()[0].lotId;
    c.recordEntry({ lotId: lot1, serial: '74102', mass: 16.8, overwrite: 'fix', worker: '山崎' });   // 訂正（数えない）
    c.recordEntry({ lotId: lot1, serial: '74103', mass: 16.6, overwrite: 'dup', worker: '山崎' });   // ダブり（数えない）
    // 別の記号のロットも同じ日に流れる（1日1つの通し番号にまとめる）
    const o = c.createLot({ prefix: 'HEP', start: '59701', kind: '50kg' });
    c.recordEntry({ lotId: o.lotId, serial: '59723', mass: 34.8, worker: '田中' });
    let freeN = 74400; while (order.includes(String(freeN))) freeN--;   // まだ流れていない番号を修正にする
    const freeS = String(freeN);
    c.recordEntry({ lotId: c.readLots_().find(l => +l.start <= freeN && freeN <= +l.end).lotId, serial: freeS, mass: null, repair: true, worker: '田中' });
    const today = c.tokyoDate_(new Date());
    const r = c.makeSaibanPdf({ date: today });
    const key = Object.keys(global.PDFBLOBS).find(k => k.endsWith('/採番表_' + today + '.pdf'));
    const html = key ? global.PDFBLOBS[key].html : '';
    check('採番表: 1日1つにまとめる（訂正・ダブりは数えない）', r.count === 252 && r.groups.join(',') === 'HXF,HEP' && !!r.pdfUrl && !!r.sheetUrl, JSON.stringify(r));
    const nums = [...html.matchAll(/<td class="no">(\d+)<\/td><td class="num">([A-Z]*\d*)/g)].map(x => [Number(x[1]), x[2]]).filter(x => x[1]).sort((a, b) => a[0] - b[0]).map(x => x[1]);
    check('採番表PDF: 入力した順に通し番号（記号が混ざる日は記号付き）', nums.length === 252 && nums.slice(0, 250).join(',') === order.map(n => 'HXF' + n).join(',') && nums[250] === 'HEP59723' && nums[251] === 'HXF' + freeS, nums.slice(248).join(','));
    check('採番表PDF: 1ページ200本・見出しに記号・作業者', (html.match(/class="page"/g) || []).length === 2 && html.includes('容器記号：<b>HXF・HEP</b>') && html.includes('作業者：<b>山崎・田中</b>') && html.includes('20・50kg容器　採番表'));
    check('採番表PDF: 修正に印・ダブりは載せない', html.includes('HXF' + freeS + '<span class="mk">修正</span>') && !html.includes('>W<'));
    const d = today.split('-');
    check('採番表PDF: 保存先は 採番表/年/月', !!key && key.startsWith('/LPG容器 質量入力/採番表/' + d[0] + '年/' + d[1] + '月/'), key);
    const sss = global.ARCHIVES.find(x => x.name === 'LPG容器 採番表_' + d[0]), ssh = sss && sss.getSheets()[0];
    const srows = ssh ? ssh.getRange(2, 1, ssh.getLastRow() - 1, 10).getValues() : [];
    check('採番表シート: 年ごとのスプレッドシートに1行1本で書き足す', ssh && ssh.getName() === '採番表' && srows.length === 252 && srows[0][0] === today && srows[0][1] === 1 && srows[0][2] === 'HXF' && srows[0][3] === order[0] && srows[250][2] === 'HEP' && srows[250][5] === 598 && srows[251][8] === '修正', JSON.stringify(srows[0]) + JSON.stringify(srows[250]));
    c.makeSaibanPdf({ date: today });
    check('採番表: 作り直すと同じ日の行・PDFは置き換える', ssh.getLastRow() - 1 === 252 && global.PDFS.filter(p => p.endsWith('/採番表_' + today + '.pdf')).length === 1);
    check('採番表: 流れていない日は作らない', c.makeSaibanPdf({ date: '2000-01-01' }).count === 0);
    c.installNightlyTrigger(); c.installNightlyTrigger();
    check('毎晩の自動処理: 22時台の時間指定が1つだけ', global.TRIGGERS.length === 1 && global.TRIGGERS[0].o.hour === 22 && global.TRIGGERS[0].o.tz === 'Asia/Tokyo');
    const beforeBlob = global.PDFBLOBS[key];
    c.nightlyJob();
    check('毎晩の自動処理: その日の採番表PDFを作る', global.PDFBLOBS[key] !== beforeBlob && global.PDFS.filter(p => p.endsWith('/採番表_' + today + '.pdf')).length === 1 && c.sharedProps_().getProperty('SAIBAN_DONE_UNTIL') === today);
  }
  // ---- ドライブのフォルダを「LPG容器 質量入力」に全部まとめる（v52までの並びからの整理）
  {
    const m = load(), c = m.ctx, P = c.sharedProps_();
    const old = c.DriveApp.createFolder('LPG容器 検査成績表PDF');
    P.setProperty('PDF_FOLDER_ID', old.getId());
    old.createFolder('2026年').createFolder('09月').createFile({ name: '成績表_HEP59701-59800.pdf' });
    old.createFile({ name: '成績表_HRH00001-00100.pdf' });
    old.createFolder('採番表').createFolder('2026年').createFolder('09月').createFile({ name: '採番表_2026-09-27.pdf' });
    const sa = c.SpreadsheetApp.create('LPG容器 採番表_2026'); c.DriveApp.getFileById(sa.getId());
    c.ensureFolderLayout_();
    const names = f => f.kids.map(k => k.name).sort().join(',');
    check('フォルダ整理: 名前を「LPG容器 質量入力」に', old.getName() === 'LPG容器 質量入力' && names(old) === '入力記録（過去分）,成績表PDF,採番表', names(old));
    const paths = global.PDFS.filter(p => p.endsWith('.pdf')).sort().join(' | ');
    check('フォルダ整理: 成績表PDFは「成績表PDF/年/月」、採番表はそのまま', paths === '/LPG容器 質量入力/成績表PDF/2026年/09月/成績表_HEP59701-59800.pdf | /LPG容器 質量入力/成績表PDF/成績表_HRH00001-00100.pdf | /LPG容器 質量入力/採番表/2026年/09月/採番表_2026-09-27.pdf', paths);
    check('フォルダ整理: 1回だけ', P.getProperty('FOLDER_LAYOUT') === '1' && c.ensureFolderLayout_() === false);
    // 月フォルダ直下の古いPDFを機種のフォルダへ
    const mf = old.kids.find(k => k.name === '成績表PDF').kids.find(k => k.name === '2026年').kids.find(k => k.name === '09月');
    const fid = mf.createFile({ name: '成績表_HXF74101-74200.pdf' }).getId();
    c.createLot({ prefix: 'HXF', start: '74101', kind: '20kg' });
    const lr = c.findLot_('HXF74101').row; m.ss.getSheetByName('ロット').getRange(lr, 14).setValue('https://drive.google.com/file/d/' + fid + '/view');
    c.ensurePdfKindFolders_();
    check('機種フォルダ: 月フォルダ直下の成績表PDFを機種のフォルダへ移す', global.PDFS.includes('/LPG容器 質量入力/成績表PDF/2026年/09月/20kg/成績表_HXF74101-74200.pdf'), global.PDFS.join(' | '));
    const r = c.getPdfMonthFolder_(new Date());
    check('フォルダ整理のあと: 成績表PDFは 成績表PDF/年/月 に入る', r.parent.parent.name === '成績表PDF');
  }
  // ---- 同じ名前のまとめ先フォルダが2つできたときの統合（旧URLと今のURLでプロジェクトが別だったため）
  {
    const m = load(), c = m.ctx;
    const a = c.DriveApp.createFolder('LPG容器 質量入力');                 // 古い方（本物）
    a.createFolder('成績表PDF').createFolder('2026年').createFolder('09月').createFolder('50kg').createFile({ name: '成績表_HEP59701-59800.pdf' });
    a.createFolder('採番表').createFolder('2026年').createFolder('09月').createFile({ name: '採番表_2026-09-27.pdf' });
    const b = c.DriveApp.createFolder('LPG容器 質量入力');                 // 22時に新しくできた方
    const bs = b.createFolder('採番表');
    const b9 = bs.createFolder('2026年').createFolder('09月');
    b9.createFile({ name: '採番表_2026-09-27.pdf' }); b9.createFile({ name: '採番表_2026-09-28.pdf' });
    const sa = c.SpreadsheetApp.create('LPG容器 採番表_2026'); c.DriveApp.getFileById(sa.getId()).moveTo(bs);
    c.DriveApp.getFileById(m.ss.getId()).moveTo(b);
    check('統合前: 保存先が決まっていなければ一番古いフォルダを使う', c.getPdfFolder_().getId() === a.getId());
    c.mergeRootFolders_();
    const live = global.DRIVE_ROOT.kids.filter(k => k.name === 'LPG容器 質量入力');
    const paths = global.PDFS.sort().join(' | ');
    check('統合: 古い方の1つにまとまる', live.length === 1 && live[0] === a && b.trashed === true);
    check('統合: 同じ名前のPDFは新しい方を残し、無い方はそのまま移す', paths.includes('/LPG容器 質量入力/採番表/2026年/09月/採番表_2026-09-28.pdf') &&
      global.PDFS.filter(p => p.endsWith('採番表_2026-09-27.pdf')).length === 1 && paths.includes('/成績表PDF/2026年/09月/50kg/成績表_HEP59701-59800.pdf'), paths);
    check('統合: 採番表のスプレッドシートと記録用スプレッドシートも移る', paths.includes('/LPG容器 質量入力/採番表/LPG容器 採番表_2026') && sa.fo.parent.parent === a && m.ss.fo.parent === a);
    check('統合: 保存先は記録用スプレッドシートの隠しシートに入る（旧URLのプロジェクトとも共有）', m.ss.getSheetByName('_システム').getRange(2, 1, m.ss.getSheetByName('_システム').getLastRow() - 1, 2).getValues().some(r => r[0] === 'PDF_FOLDER_ID' && r[1] === a.getId()));
    check('統合のあと: 採番表のスプレッドシートは名前で見つけて使う（作り直さない）', c.saibanSheet_('2026').getParent() === sa && global.ARCHIVES.length === 1);
  }
  // ---- 旧URL（別プロジェクト）は新しいURLへ案内するだけ
  {
    const m = load(), c = m.ctx;
    check('今のプロジェクトでは入力画面を出す', c.isOldProject_() === false);
    global.SCRIPT_ID = '1WRdps6zQgFVtqMeW9OhRvLeZGafuOCXFZkv-89skSv6WJ0atLhBZnSfG';
    check('旧プロジェクトでは案内画面にする', c.isOldProject_() === true && c.movedPageHtml_().includes('AKfycbyDM2IQ8NY6LgQpLb0gpQoK32fAQtWVBrdBQduZ6FOBWREFn61Qb55R9Q20Y4u9IfQ5/exec'));
    delete global.SCRIPT_ID;
  }
  // ---- 標準質量に 10kg・30kg把手（初期値のままの設定は置き換える）
  {
    const m = load(), c = m.ctx; c.setup();
    const conf = m.ss.getSheetByName('設定'), rows = conf.getRange(2, 1, conf.getLastRow() - 1, 2).getValues();
    const i = rows.findIndex(r => r[0] === '標準質量');
    conf.getRange(2 + i, 2).setValue('5kg=6.8, 8kg=9.6, 20kg=16.7, 30kg=24.0, 50kg=34.8, 50kg S付=36.3');
    c.ensureTypicalDefault_();
    const t = c.readSettings_().typical;
    check('標準質量: 10kg=11.3・30kg=23.5・30kg把手=24.0', t['10kg'] === 11.3 && t['30kg'] === 23.5 && t['30kg把手'] === 24.0, JSON.stringify(t));
    check('標準質量: 20kg三部制=17.5', t['20kg三部制'] === 17.5);
    check('20kg三部制のグループNoは20kgと同じ50本ごと', c.groupNoOf_('20kg三部制', '59701') === 195);
    conf.getRange(2 + i, 2).setValue('5kg=7.0');
    c.ensureTypicalDefault_();
    check('標準質量: 自分で書き換えた値はそのまま', c.readSettings_().typical['5kg'] === 7.0 && !c.readSettings_().typical['10kg']);
    check('30kg把手のグループNoは30kgと同じ100本ごと', c.groupNoOf_('30kg把手', '59701') === 598);
  }
  // ---- まとめて保存（通信1回で複数本）
  {
    const m = load(), c = m.ctx; c.setup();
    const L = c.createLot({ prefix: 'BT', start: '00101', kind: '20kg' });
    c.createLot({ prefix: 'BT', start: '00201', kind: '20kg' });
    const before = global.TEXTFINDS || 0;
    const r = c.recordEntries([
      { lotId: L.lotId, serial: '00101', mass: 16.7, clientId: 'c1' },
      { lotId: L.lotId, serial: '00102', mass: 16.8, clientId: 'c2' },
      { lotId: L.lotId, serial: '00999', mass: 16.8, clientId: 'c3' },   // 範囲外 → ここで止まる
      { lotId: L.lotId, serial: '00103', mass: 16.9, clientId: 'c4' }
    ]);
    const e = c.getLot(L.lotId).entries;
    check('まとめて保存: 順に保存し、エラーの所で止める', r.results.length === 2 && r.errorClientId === 'c3' && r.error.includes('範囲外') && e['00101'] === 16.7 && e['00102'] === 16.8 && e['00103'] === undefined, JSON.stringify(r));
    check('まとめて保存: ロットはA列の検索で探す（ロットシートを全部読まない）', (global.TEXTFINDS || 0) > before);
    check('ロット検索: 見つからないロットはエラー', (() => { try { c.findLot_('NOPE'); return false; } catch (x) { return x.message.includes('見つかりません'); } })());
  }
  // ---- 完了したロットは削除できない・削除してしまったロットを戻す
  {
    const m = load(), c = m.ctx; c.setup();
    const L = c.createLot({ prefix: 'HEP', start: '59701', kind: '50kg' });
    for (let i = 1; i <= 3; i++) c.recordEntry({ lotId: L.lotId, serial: String(59700 + i), mass: 34.8, worker: '山田' });
    c.completeLot({ lotId: L.lotId, send: false });
    check('完了したロットは削除できない', (() => { try { c.deleteLot(L.lotId); return false; } catch (e) { return e.message.includes('完了したロットは削除できません'); } })());
    // 以前の作りで消えてしまった状態を再現（ロット行を消し、入力記録を「削除」に）
    const lotsSh = m.ss.getSheetByName('ロット'); lotsSh.deleteRow(c.findLot_(L.lotId).row);
    const lg = m.ss.getSheetByName('入力記録');
    for (let r = 2; r <= lg.getLastRow(); r++) if (lg.getRange(r, 3).getValue() === 'HEP59701') lg.getRange(r, 14).setValue('削除');
    c.runMaintenance();
    const back = c.findLot_('HEP59701');
    const st = lg.getRange(2, 1, lg.getLastRow() - 1, 18).getValues().filter(r => r[2] === 'HEP59701').map(r => r[13]);
    check('削除してしまった完了ロットを戻す（PDFがあれば完了で）', back.status === '完了' && back.end === '59800' && back.kind === '50kg' && back.pdf.includes('drive.google.com') && st.length === 3 && st.every(x => x === '有効'), JSON.stringify(back) + st.join(','));
    check('戻したロットの入力状況は入力記録から出る', Object.keys(c.getLot('HEP59701').entries).length === 3);
    check('戻すのは1回だけ', c.undeleteLot('HEP59701').already === true);
  }
  // ---- 入力記録・採番表の一覧で、日付が変わる所に線
  {
    const m = load(), c = m.ctx; c.setup();
    const L = c.createLot({ prefix: 'DL', start: '00101', kind: '20kg' });
    const lg = m.ss.getSheetByName('入力記録');
    c.recordEntry({ lotId: L.lotId, serial: '00101', mass: 16.7 });
    c.recordEntry({ lotId: L.lotId, serial: '00102', mass: 16.7 });
    lg.getRange(2, 2).setValue(new Date(Date.now() - 86400000)); lg.getRange(3, 2).setValue(new Date(Date.now() - 86400000));   // 昨日の2本にする
    c.recordEntry({ lotId: L.lotId, serial: '00103', mass: 16.7 });
    c.recordEntry({ lotId: L.lotId, serial: '00104', mass: 16.7 });
    const b = lg.borders || {};
    check('入力記録: 日付が変わった行の上に線', b[4] === true && !b[3] && !b[5], JSON.stringify(b));
    lg.borders = {};
    c.sharedProps_().deleteProperty('LOG_DAY_LINES'); c.ensureLogDayLines_();
    check('入力記録: これまでの記録にも1回だけ線を引く', (lg.borders || {})[4] === true && !(lg.borders || {})[3]);
  }
  // ---- 「＋ 前回と同じ」はロットごと（50kgのあとに20kgで押しても50kgの値を入れない）・標準から大きく外れた値は確認
  {
    const m = load(), c = m.ctx; c.setup();
    const A = c.createLot({ prefix: 'HEP', start: '64201', kind: '50kg' });
    for (let i = 1; i <= 6; i++) c.recordEntry({ lotId: A.lotId, serial: String(64200 + i), mass: 34.8 });
    const B = c.createLot({ prefix: 'HXP', start: '97101', kind: '20kg' });
    const p3 = await browser.newPage({ viewport: { width: 1180, height: 820 } });
    p3.on('pageerror', e => check('ページでJSエラーなし(+)', false, e.message));
    await p3.exposeFunction('__gas', (fn, arg) => {
      try { const v = c[fn](JSON.parse(arg)); return JSON.stringify({ v: JSON.parse(JSON.stringify(v === undefined ? null : v)) }); }
      catch (e) { return JSON.stringify({ err: e.message }); }
    });
    await p3.setContent(html); await p3.waitForTimeout(500);
    await noDateOk(p3);
    await p3.selectOption('#fWorker', '山田'); await p3.fill('#fDevice', 'x');
    await p3.click('#btnStart'); await p3.waitForTimeout(200);
    const k3 = async (seq, w = 250) => { for (const ch of seq) await p3.keyboard.press(KEY[ch] || 'Numpad' + ch); await p3.waitForTimeout(w); return p3.textContent('#notice'); };
    const tabsTxt = (await p3.textContent('#tabs')).replace(/\s+/g, '');
    check('上のバー: 20kg以下は容器番号、50kgはグループNo', tabsTxt.includes('HXP97101〜') && tabsTxt.includes('HEP643'), tabsTxt);
    await k3('207'); await k3('8', 400);
    const n1 = await k3('105'); const n2 = await k3('+', 600);
    check('「＋」は同じロットの値（20kgのロットに50kgの値を入れない）', n2.includes('HXP97105') && n2.includes('W16.7') && c.getLot(B.lotId).entries['97105'] === 16.7, n2);
    await k3('106'); const n3 = await k3('348', 400);
    check('20kgのロットに34.8は確認してから', n3.includes('大きく離れています') && c.getLot(B.lotId).entries['97106'] === undefined, n3);
    const n4 = await k3('E', 600);
    check('確認してEnterなら保存できる', c.getLot(B.lotId).entries['97106'] === 34.8, n4);
    await p3.close();
  }
  // ---- 製造年月OKを押すまで質量は入れられない
  {
    const m = load(), c = m.ctx; c.setup();
    const L = c.createLot({ prefix: 'DT', start: '00101', kind: '20kg' });
    const p4 = await browser.newPage({ viewport: { width: 820, height: 1180 } });
    p4.on('pageerror', e => check('ページでJSエラーなし(製造年月)', false, e.message));
    await p4.exposeFunction('__gas', (fn, arg) => {
      try { const v = c[fn](JSON.parse(arg)); return JSON.stringify({ v: JSON.parse(JSON.stringify(v === undefined ? null : v)) }); }
      catch (e) { return JSON.stringify({ err: e.message }); }
    });
    await p4.setContent(html); await p4.waitForTimeout(500);
    await p4.selectOption('#fWorker', '山田'); await p4.fill('#fDevice', 'x');
    await p4.click('#btnStart'); await p4.waitForTimeout(200);
    const k4 = async (seq, w = 250) => { for (const ch of seq) await p4.keyboard.press(KEY[ch] || 'Numpad' + ch); await p4.waitForTimeout(w); return p4.textContent('#notice'); };
    await k4('105');
    check('製造年月OKのボタンが入力バーの上に出る', await p4.isVisible('#btnDateOk') && (await p4.getAttribute('#massSheet', 'class')).includes('locked'));
    const n1 = await k4('167', 400);
    check('OKの前は質量を打っても入らない', n1.includes('製造年月') && c.getLot(L.lotId).entries['00105'] === undefined, n1);
    await p4.click('#cands button >> nth=5', { force: true }); await p4.waitForTimeout(300);
    check('OKの前は候補ボタンをタップしても入らない', c.getLot(L.lotId).entries['00105'] === undefined);
    await k4('E');
    check('Enterで製造年月OK', (await p4.textContent('#btnDateOk')).includes('✓') && !(await p4.getAttribute('#massSheet', 'class')).includes('locked'));
    await k4('167', 500);
    check('OKのあとは質量が入る', c.getLot(L.lotId).entries['00105'] === 16.7);
    await k4('106');
    check('次の容器ではまたOKが必要', (await p4.getAttribute('#massSheet', 'class')).includes('locked'));
    await p4.screenshot({ path: path.join(__dirname, 'out', 'dateok.png') });
    await p4.click('#btnDateOk'); await p4.waitForTimeout(100);
    await p4.click('#cands button >> nth=5'); await p4.waitForTimeout(400);
    check('ボタンでOKしてから候補タップで保存', c.getLot(L.lotId).entries['00106'] === 16.7);
    await p4.close();
  }
  // ---- 10kg（50本ごと・内容積24）・30kg把手（内容積71）
  {
    const m = load(), c = m.ctx; c.setup();
    check('10kgのグループNoは50本ごと', c.groupNoOf_('10kg', '59701') === 195 && c.groupNoOf_('10kg', '50050') === 1);
    const conf = m.ss.getSheetByName('設定'), rows = conf.getRange(2, 1, conf.getLastRow() - 1, 2).getValues();
    conf.getRange(2 + rows.findIndex(r => r[0] === '内容積'), 2).setValue('5kg=12, 8kg=19, 20kg=47, 30kg=71, 50kg=118, 50kg S付=118');
    const L = c.createLot({ prefix: 'TK', start: '00101', kind: '10kg' });
    const sh = m.ss.getSheetByName(c.findLot_(L.lotId).sheetName);
    check('（前の設定のままだと）10kgの内容積は空', String(sh.getRange('N3').getValue()) === '');
    c.ensureVolumes3_();
    check('設定の内容積に10kg=24・30kg把手=71を足し、入力中ロットの空欄も埋める', c.readSettings_().volume['10kg'] === 24 && c.readSettings_().volume['30kg把手'] === 71 && String(sh.getRange('N3').getValue()) === '24');
  }
  await browser.close();

  const failed = results.filter(r => !r.ok).length;
  console.log('\n' + (results.length - failed) + ' / ' + results.length + ' PASS');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
