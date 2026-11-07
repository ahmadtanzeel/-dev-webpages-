// GAS のサービスをローカルで真似する最小限のモック。
// コード.js / AbsenceCheck.js を vm で読み込み、関数を直接呼んでテストする。
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

function pad(n) { return String(n).padStart(2, '0'); }
function formatDate(d, _tz, fmt) {
  // テストはTZ=Asia/Tokyoで動かす前提。必要な書式だけ対応
  return fmt
    .replace('yyyy', d.getFullYear())
    .replace('MM', pad(d.getMonth() + 1))
    .replace('dd', pad(d.getDate()))
    .replace('HH', pad(d.getHours()))
    .replace('mm', pad(d.getMinutes()))
    .replace('ss', pad(d.getSeconds()));
}

class Sheet {
  constructor(name, rows) { this.name = name; this.rows = rows || []; }
  getLastRow() { return this.rows.length; }
  getDataRange() {
    const self = this;
    return { getValues: () => self.rows.map(r => r.slice()) };
  }
  getRange(row, col, numRows, numCols) {
    const self = this;
    numRows = numRows || 1; numCols = numCols || 1;
    const ensure = () => {
      while (self.rows.length < row + numRows - 1) self.rows.push([]);
    };
    return {
      getValues() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const src = self.rows[row - 1 + r] || [];
          const line = [];
          for (let c = 0; c < numCols; c++) {
            const v = src[col - 1 + c];
            line.push(v === undefined ? '' : v);
          }
          out.push(line);
        }
        return out;
      },
      setValues(vals) {
        ensure();
        vals.forEach((line, r) => line.forEach((v, c) => { self.rows[row - 1 + r][col - 1 + c] = v; }));
        return this;
      },
      setValue(v) { ensure(); self.rows[row - 1][col - 1] = v; return this; },
      setFontWeight() { return this; },
      setNumberFormat() { return this; },
    };
  }
  appendRow(r) { this.rows.push(r.slice()); }
  deleteRow(i) { this.rows.splice(i - 1, 1); }
  setFrozenRows() {}
}

function createEnv(opts) {
  opts = opts || {};
  const sheets = {};
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = new Sheet(n, [])),
  };
  const props = Object.assign({}, opts.props || {});
  const cacheStore = {};
  const sent = { mail: [], chat: [], gmail: [], calendarCreated: [] };
  let clock = opts.now ? new Date(opts.now) : new Date();
  let lockAvailable = true;
  const calendars = opts.calendars || {};

  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    Date: class extends Date {
      constructor(...a) { if (a.length === 0) super(clock.getTime()); else super(...a); }
      static now() { return clock.getTime(); }
    },
    JSON, Math, String, Number, Array, Object, Set, Map, parseInt, parseFloat, isNaN, RegExp, Error,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      openById: () => ss,
      flush() {},
    },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null) }),
    },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (k in cacheStore ? cacheStore[k] : null),
        put: (k, v) => { cacheStore[k] = v; },
        remove: (k) => { delete cacheStore[k]; },
      }),
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => { if (!lockAvailable) return false; lockAvailable = false; return true; },
        waitLock: () => { if (!lockAvailable) throw new Error('lock timeout'); lockAvailable = false; },
        releaseLock: () => { lockAvailable = true; },
      }),
    },
    Utilities: { formatDate },
    MailApp: { sendEmail: (m) => sent.mail.push(m) },
    GmailApp: { sendEmail: (...a) => sent.gmail.push(a), createDraft() {} },
    UrlFetchApp: { fetch: (url, o) => sent.chat.push(JSON.parse(o.payload)) },
    CalendarApp: {
      getCalendarById: (id) => calendars[id] || null,
    },
    ScriptApp: {
      _triggers: opts.triggers || [],
      getProjectTriggers() { return this._triggers; },
      deleteTrigger(t) { this._triggers = this._triggers.filter(x => x !== t); },
      newTrigger(fn) {
        const self = this;
        const b = { timeBased: () => b, everyMinutes: () => b, onWeekDay: () => b, atHour: () => b,
          create: () => { const t = { getHandlerFunction: () => fn }; self._triggers.push(t); return t; } };
        return b;
      },
      WeekDay: { MONDAY: 'MONDAY' },
    },
    ContentService: {
      createTextOutput: (s) => ({ setMimeType() { return { body: s }; } }),
      MimeType: { JSON: 'json' },
    },
    HtmlService: {},
  };
  vm.createContext(ctx);
  for (const f of ['コード.js', 'AbsenceCheck.js']) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
  }

  return {
    ctx, sheets, sent, props, cacheStore,
    addSheet: (n, rows) => (sheets[n] = new Sheet(n, rows)),
    // コード側の instanceof Date を通すため、テストデータの日時は必ずこれで作る
    D: (...a) => new ctx.Date(...a),
    setNow: (d) => { clock = new Date(d); },
    advance: (ms) => { clock = new Date(clock.getTime() + ms); },
    holdLock: () => { lockAvailable = false; },
    releaseLock: () => { lockAvailable = true; },
  };
}

module.exports = { createEnv, Sheet };
