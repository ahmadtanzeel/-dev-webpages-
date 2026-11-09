// 自動欠席メールの廃止（T11）と催促メールトリガー（R-04）のテスト
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./gas-mock');

function trig(fn) { return { getHandlerFunction: () => fn }; }

test('T11 checkAbsenceAndSendMail は呼ばれてもメールを送らない', () => {
  const env = createEnv({ now: '2026-11-07T15:00:00+09:00' });
  // シートが無くても、先頭で止まるので例外にならない
  env.ctx.checkAbsenceAndSendMail();
  assert.equal(env.sent.gmail.length, 0);
});

test('removeAbsenceTrigger は欠席メールのトリガーだけを消す', () => {
  const env = createEnv({ triggers: [trig('checkAbsenceAndSendMail'), trig('checkAbsenceAndSendMail'), trig('sendReminderVersionA')] });
  assert.equal(env.ctx.removeAbsenceTrigger(), 2);
  assert.deepEqual(env.ctx.listTriggers(), ['sendReminderVersionA']);
});

test('旧 setupTrigger を実行してもトリガーを作らない', () => {
  const env = createEnv({ triggers: [trig('checkAbsenceAndSendMail')] });
  env.ctx.setupTrigger();
  assert.deepEqual(env.ctx.listTriggers(), []);
});

test('催促トリガーは A だけを登録する', () => {
  const env = createEnv({ triggers: [trig('sendReminderVersionA'), trig('sendReminderVersionB'), trig('checkAbsenceAndSendMail')] });
  env.ctx.setupReminderTrigger();
  assert.deepEqual(env.ctx.listTriggers().sort(), ['checkAbsenceAndSendMail', 'sendReminderVersionA']);
});
