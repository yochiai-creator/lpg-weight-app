// gas/*.js の構文チェック（clasp push 前に実行）
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..', 'gas');
let ng = 0;
for (const f of fs.readdirSync(dir).filter(f => f.endsWith('.js'))) {
  try { execFileSync(process.execPath, ['--check', path.join(dir, f)], { stdio: 'pipe' }); console.log('OK  ' + f); }
  catch (e) { ng++; console.log('NG  ' + f + '\n' + e.stderr); }
}
process.exit(ng ? 1 : 0);
