// 受付打刻（processScan）のテスト。監査レポート v2 の T01〜T06 に対応。
// 実行: TZ=Asia/Tokyo node --test tests/
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./gas-mock');

const KEY = 'test-kiosk-key-0123456789';

function setup() {
  const env = createEnv({ now: '2026-11-07T15:00:00+09:00', props: { KIOSK_KEY: KEY, GCHAT_WEBHOOK_URL: 'https://chat.example/webhook' } });
  env.addSheet('公開プロフィール', [
    ['ID', '生徒メール', '保護者メール', '氏名', 'ニックネーム', '目標', 'トークン', '模試名', '模試日', '', 'カレンダー'],
    ['S001', 's1@example.com', 'p1@example.com', '山田', '', 15, 'tok1', '', '', '', ''],
    ['S002', 's2@example.com', 'p2@example.com', '佐藤', '', 15, 'tok2', '', '', '', ''],
  ]);
  env.addSheet('管理シート', [['日付', 'ID', '', '入室', '退室', 'エール']]);
  return env;
}
const scan = (env, id, key = KEY, req) => env.ctx.processScan(id, key, req);
const logRows = (env) => env.sheets['管理シート'].rows.slice(1);

test('T01 受付キーなし・誤り・サーバー未設定は拒否し、何も変えない', () => {
  const env = setup();
  let r = scan(env, 'S001', '');
  assert.equal(r.ok, false); assert.equal(r.code, 'KIOSK_UNAUTHORIZED');
  r = scan(env, 'S001', 'wrong-key-wrong-key-wrong');
  assert.equal(r.ok, false); assert.equal(r.code, 'KIOSK_UNAUTHORIZED');
  delete env.props.KIOSK_KEY;
  r = scan(env, 'S001', KEY);
  assert.equal(r.ok, false); assert.equal(r.code, 'KIOSK_NOT_CONFIGURED');
  assert.equal(logRows(env).length, 0);
  assert.equal(env.sent.mail.length + env.sent.chat.length, 0);
});

test('T02 未登録ID・空IDは拒否し、ログ行が増えない', () => {
  const env = setup();
  assert.equal(scan(env, 'NOPE').code, 'UNKNOWN_ID');
  assert.equal(scan(env, '   ').code, 'EMPTY_ID');
  assert.equal(scan(env, 'NOPE').ok, false);
  assert.equal(logRows(env).length, 0);
  assert.equal(env.sent.mail.length + env.sent.chat.length, 0);
});

test('正常：入室 → 退室で1行、通知は各1回ずつ', () => {
  const env = setup();
  let r = scan(env, 'S001');
  assert.equal(r.ok, true); assert.equal(r.code, 'CHECK_IN');
  assert.match(r.message, /山田 さんが 入室/);
  env.advance(60 * 60 * 1000);
  r = scan(env, 'S001');
  assert.equal(r.code, 'CHECK_OUT');
  assert.equal(logRows(env).length, 1);
  assert.ok(logRows(env)[0][4] instanceof env.ctx.Date);
  assert.equal(env.sent.mail.length, 2);
  assert.equal(env.sent.chat.length, 2);
});

test('T03 同じ生徒を3秒間隔で2回 → 2回目は無視・通知なし', () => {
  const env = setup();
  scan(env, 'S001');
  env.advance(3000);
  const r = scan(env, 'S001');
  assert.equal(r.ok, true); assert.equal(r.code, 'DUPLICATE');
  assert.equal(logRows(env).length, 1);
  assert.equal(logRows(env)[0][4], undefined); // 退室になっていない
  assert.equal(env.sent.mail.length, 1);
  assert.equal(env.sent.chat.length, 1);
});

test('退室直後（5秒以内）の再スキャンで再入室にならない', () => {
  const env = setup();
  scan(env, 'S001');
  env.advance(60 * 60 * 1000);
  scan(env, 'S001');
  env.advance(2000);
  assert.equal(scan(env, 'S001').code, 'DUPLICATE');
  assert.equal(logRows(env).length, 1);
});

test('T04 同じ生徒を6秒間隔で2回 → 入室・退室の両方が有効', () => {
  const env = setup();
  assert.equal(scan(env, 'S001').code, 'CHECK_IN');
  env.advance(6000);
  assert.equal(scan(env, 'S001').code, 'CHECK_OUT');
  assert.equal(env.sent.chat.length, 2);
});

test('5秒ちょうどは受け付ける（5秒未満だけ無視）', () => {
  const env = setup();
  scan(env, 'S001');
  env.advance(5000);
  assert.equal(scan(env, 'S001').code, 'CHECK_OUT');
});

test('T05 別の生徒が1秒差で打刻 → 両方受理', () => {
  const env = setup();
  assert.equal(scan(env, 'S001').code, 'CHECK_IN');
  env.advance(1000);
  assert.equal(scan(env, 'S002').code, 'CHECK_IN');
  assert.equal(logRows(env).length, 2);
});

test('T06 同じrequestIdの再送 → 1回分だけ記録・通知', () => {
  const env = setup();
  const first = scan(env, 'S001', KEY, 'req-1');
  for (let i = 0; i < 20; i++) {
    env.advance(10000); // 5秒ガードを越えても、同じrequestIdなら記録しない
    const r = scan(env, 'S001', KEY, 'req-1');
    assert.equal(r.code, 'ALREADY_PROCESSED');
    assert.equal(r.message, first.message);
  }
  assert.equal(logRows(env).length, 1);
  assert.equal(logRows(env)[0][4], undefined);
  assert.equal(env.sent.chat.length, 1);
});

test('ロックが取れないときは失敗として返し、何も変えない', () => {
  const env = setup();
  env.holdLock();
  const r = scan(env, 'S001');
  assert.equal(r.ok, false); assert.equal(r.code, 'BUSY');
  assert.equal(logRows(env).length, 0);
});

test('前日の入室中の行は今日の打刻に影響しない', () => {
  const env = setup();
  env.sheets['管理シート'].rows.push([env.D('2026-11-06T10:00:00+09:00'), 'S001', '', env.D('2026-11-06T10:00:00+09:00'), '']);
  assert.equal(scan(env, 'S001').code, 'CHECK_IN');
  assert.equal(logRows(env).length, 2);
});

test('doPost 経由でも ok:false がそのまま返る', () => {
  const env = setup();
  const out = env.ctx.doPost({ postData: { contents: JSON.stringify({ action: 'scan', id: 'NOPE', kioskKey: KEY }) } });
  const json = JSON.parse(out.body);
  assert.equal(json.ok, false);
  assert.equal(json.code, 'UNKNOWN_ID');
  assert.ok(json.error);
});
