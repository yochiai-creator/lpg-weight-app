// 使用中のロット: iPadが入力したロットはそのiPadの「使用中」になり、ほかのiPadは見るだけ（入力・消す・完了・削除・機種変更はできない）
// ホームに戻る・完了する・10分入力がないと使用中は外れる。記録は CacheService（全員で共通、すぐ読める）
const CLAIM_KEY = 'LOT_CLAIMS';
const CLAIM_IDLE_MS = 10 * 60 * 1000;

function claimCache_() {
  try { return CacheService.getScriptCache(); } catch (e) { return null; }
}

// { ロットID: { devId, worker, device, at } }（期限切れは除く）
function readClaims_() {
  const c = claimCache_();
  if (!c) return {};
  let m = {};
  try { m = JSON.parse(c.get(CLAIM_KEY) || '{}'); } catch (e) { m = {}; }
  const now = Date.now();
  Object.keys(m).forEach(function(k) { if (!m[k] || now - m[k].at > CLAIM_IDLE_MS) delete m[k]; });
  return m;
}

function writeClaims_(m) {
  const c = claimCache_();
  if (c) c.put(CLAIM_KEY, JSON.stringify(m), 6 * 3600);
}

// 同じiPadか（印が消えて作り直された場合も、入力者と端末名が同じなら同じiPadとみなす）
function sameDevice_(cl, who) {
  if (cl.devId === String(who.devId)) return true;
  return !!cl.worker && !!cl.device && cl.worker === String(who.worker || '') && cl.device === String(who.device || '');
}

function claimText_(cl) {
  return [cl.worker, cl.device].filter(function(x) { return x; }).join('・') || '別のiPad';
}

// 別のiPadが使用中ならエラー。who: { devId, worker, device }（devId のない古い画面は確かめない）
// take: true ならこのiPadの使用中にする（入力したとき）
// 保存1回（まとめて保存なら全部の入力）の中では1度だけ読み、書くのは使用中が新しくなったときか1分ごと（保存を遅くしない）
let CLAIMS_EXEC_ = null;
const CLAIM_REFRESH_MS = 60 * 1000;
function checkClaim_(lotId, who, take) {
  if (!who || !who.devId) return;
  // 使い回すのは保存（take）のときだけ。消す・完了などはその都度読む
  const m = take ? (CLAIMS_EXEC_ || (CLAIMS_EXEC_ = readClaims_())) : readClaims_();
  const cl = m[lotId];
  if (cl && !sameDevice_(cl, who)) throw new Error(claimText_(cl) + ' が入力中のため、このロットは見るだけです（' + lotId + '）');
  if (take) {
    const now = Date.now();
    if (cl && cl.devId === String(who.devId) && now - cl.at < CLAIM_REFRESH_MS) return;
    m[lotId] = { devId: String(who.devId), worker: String(who.worker || ''), device: String(who.device || ''), at: now };
    writeClaims_(m);
  }
}

function releaseClaim_(lotId) {
  CLAIMS_EXEC_ = null;
  const m = readClaims_();
  if (m[lotId]) { delete m[lotId]; writeClaims_(m); }
}

// 画面用: { ロットID: { worker, device, mine } }。who: devId か { devId, worker, device }
function getClaims(who) {
  if (typeof who !== 'object' || !who) who = { devId: who };
  const m = readClaims_(), out = {};
  Object.keys(m).forEach(function(k) { out[k] = { worker: m[k].worker, device: m[k].device, mine: !!who.devId && sameDevice_(m[k], who) }; });
  return out;
}

// ホームに戻ったとき、このiPadの使用中を外す
function releaseLots(who) {
  if (typeof who !== 'object' || !who) who = { devId: who };
  if (!who.devId) return 0;
  const m = readClaims_();
  let n = 0;
  Object.keys(m).forEach(function(k) { if (sameDevice_(m[k], who)) { delete m[k]; n++; } });
  if (n) writeClaims_(m);
  return n;
}
