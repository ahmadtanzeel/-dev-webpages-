// 入力チェック・式の混入対策（S-04）・予定登録（S-05）・設定保存（R-08）のテスト
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./gas-mock');

function fakeCalendar(env) {
  const events = [];
  return {
    events,
    getEvents: (a, b) => events.filter(e => e.start < b && e.end > a).map(e => ({
      getTitle: () => e.title, getStartTime: () => e.start, getEndTime: () => e.end,
    })),
    createEvent: (title, start, end) => { events.push({ title, start, end }); },
  };
}

function setup() {
  const cal = {};
  const env = createEnv({ now: '2026-11-07T15:00:00+09:00', calendars: cal });
  cal['cal-1'] = fakeCalendar(env);
  env.addSheet('公開プロフィール', [
    ['ID', '生徒メール', '保護者メール', '氏名', 'ニックネーム', '目標', 'トークン', '模試名', '模試日', '', 'カレンダー'],
    ['S001', 's1@example.com', 'p1@example.com', '山田', '', 15, 'tok1', '', '', '', 'cal-1'],
  ]);
  return { env, cal: cal['cal-1'] };
}

test('safeCellText：先頭が = + - @ のときだけ文字列扱いにする', () => {
  const { env } = setup();
  const f = env.ctx.safeCellText;
  assert.equal(f('=HYPERLINK("x")'), '\'=HYPERLINK("x")');
  assert.equal(f('+1'), "'+1");
  assert.equal(f('-test'), "'-test");
  assert.equal(f('@a'), "'@a");
  assert.equal(f('青チャート'), '青チャート');
  assert.equal(f('a=b'), 'a=b');
});

test('T09 = から始まる投稿は数式として保存されない', () => {
  const { env } = setup();
  env.ctx.postBookPost('tok1', '=IMPORTXML("http://x")', 5, '=1+1');
  const row = env.sheets['参考書投稿'].rows[1];
  assert.equal(row[3], '\'=IMPORTXML("http://x")');
  assert.equal(row[5], "'=1+1");
});

test('設定保存：無効トークンは例外（ok:true で返さない）', () => {
  const { env } = setup();
  assert.throws(() => env.ctx.saveStudentSettings('bad', '模試', '2027-01-16'), /無効なURL/);
  const out = JSON.parse(env.ctx.doPost({ postData: { contents: JSON.stringify({ action: 'saveSettings', token: 'bad', examName: 'x', examDate: '2027-01-16' }) } }).body);
  assert.equal(out.ok, false);
});

test('設定保存：日付・長さのチェックと式の混入対策', () => {
  const { env } = setup();
  assert.throws(() => env.ctx.saveStudentSettings('tok1', '模試', '2027-02-30'), /日付/);
  assert.throws(() => env.ctx.saveStudentSettings('tok1', '模試', '2099-01-01'), /範囲外/);
  assert.throws(() => env.ctx.saveStudentSettings('tok1', 'あ'.repeat(51), '2027-01-16'), /50文字/);
  assert.throws(() => env.ctx.saveStudentSettings('tok1', '', '2027-01-16'), /模試名/);
  env.ctx.saveStudentSettings('tok1', '=全統模試', '2027-01-16');
  const row = env.sheets['生徒設定'].rows.find(r => r[0] === 'S001');
  assert.equal(row[1], "'=全統模試");
  assert.equal(row[2], '2027-01-16');
});

test('予定登録：正常に登録できる', () => {
  const { env, cal } = setup();
  const msg = env.ctx.addCalendarEventsForStudent('tok1', [
    { date: '2026-11-08', start: '10:00', end: '12:00', memo: '数学' },
    { date: '2026-11-09', start: '18:00', end: '20:00', memo: '' },
  ]);
  assert.match(msg, /2件の予定を登録/);
  assert.equal(cal.events.length, 2);
  assert.equal(cal.events[0].title, '自習室 - 数学');
  assert.equal(cal.events[0].start.getHours(), 10);
});

test('T10 予定登録：不正な入力は1件も登録しない', () => {
  const { env, cal } = setup();
  const ok = { date: '2026-11-08', start: '10:00', end: '12:00' };
  const bad = [
    [{ date: '2026-13-01', start: '10:00', end: '12:00' }, /日付/],
    [{ date: '2026-11-08', start: '25:00', end: '26:00' }, /時刻/],
    [{ date: '2026-11-08', start: '12:00', end: '10:00' }, /終了時刻/],
    [{ date: '2026-11-01', start: '10:00', end: '12:00' }, /過去/],
    [{ date: '2027-06-01', start: '10:00', end: '12:00' }, /90日/],
    [{ date: '2026-11-08', start: '10:00', end: '12:00', memo: 'x'.repeat(51) }, /50文字/],
  ];
  for (const [ev, re] of bad) {
    assert.throws(() => env.ctx.addCalendarEventsForStudent('tok1', [ok, ev]), re);
  }
  const many = Array.from({ length: 32 }, () => ok);
  assert.throws(() => env.ctx.addCalendarEventsForStudent('tok1', many), /31件/);
  assert.equal(cal.events.length, 0);
});

test('予定登録：同じ時間の予定は二重に作らない（再送・二重押し）', () => {
  const { env, cal } = setup();
  const evs = [{ date: '2026-11-08', start: '10:00', end: '12:00' }, { date: '2026-11-08', start: '10:00', end: '12:00' }];
  env.ctx.addCalendarEventsForStudent('tok1', evs);
  assert.equal(cal.events.length, 1);
  const msg = env.ctx.addCalendarEventsForStudent('tok1', evs);
  assert.match(msg, /0件の予定を登録.*1件は登録済み/);
  assert.equal(cal.events.length, 1);
});

test('予定登録：無効トークンは拒否', () => {
  const { env, cal } = setup();
  assert.throws(() => env.ctx.addCalendarEventsForStudent('bad', [{ date: '2026-11-08', start: '10:00', end: '12:00' }]), /unauthorized/);
  assert.equal(cal.events.length, 0);
});
