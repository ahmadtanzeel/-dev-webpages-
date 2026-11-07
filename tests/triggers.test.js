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

test('REMINDER_VERSION 未指定なら催促トリガーを登録せず、既存も触らない', () => {
  const env = createEnv({ triggers: [trig('sendReminderVersionA'), trig('sendReminderVersionB')] });
  assert.throws(() => env.ctx.setupReminderTrigger(), /REMINDER_VERSION/);
  assert.equal(env.ctx.listTriggers().length, 2);
});
