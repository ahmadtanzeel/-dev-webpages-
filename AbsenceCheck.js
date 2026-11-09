// ============================================================
// SSS Education 自習室管理システム
// 欠席確認メール 自動送信モジュール
// ============================================================
// 【前提】
//   - 公開プロフィールシートに「カレンダーID」「メールアドレス」「保護者メール」列を追加済み
//   - 送信済みログシートが存在する（なければ自動作成）
//   - このスクリプトを既存GASプロジェクトに追加してトリガー設定する
// ============================================================

// 公開プロフィールシートの列定義・スプレッドシートIDは
// コード.gs の IDX_PROFILE / CALENDAR_CONFIG に一本化（同一GASプロジェクトのため直接参照可能）

// ========== 設定値（環境に合わせて変更） ==========
const CONFIG = {
  SHEET_LOG:      '管理シート',                    // 入退室ログシート名
  SHEET_SENT:     '送信済みログ',                  // 送信済みログシート名（新設）
  ABSENCE_MINUTES: 5,                             // 予定開始から何分後に欠席判定するか

  // 管理シートのカラム番号（1始まり）
  LOG_COL_DATE:     1,   // 日付
  LOG_COL_ID:       2,   // 生徒ID
  LOG_COL_CHECKIN:  4,   // 入室時刻
};

// ============================================================
// 【廃止】自動欠席判定メール（2026-11-07 監査レポート v2 R-02）
// 欠席の確認は教室での確認に一本化したため、この機能は使わない。
// 古いトリガーが残っていても送信されないよう、関数の先頭で止めている。
// 再開する場合は運用の合意を取った上で ABSENCE_MAIL_ENABLED を true にすること。
// ============================================================
const ABSENCE_MAIL_ENABLED = false;

// ============================================================
// メイン関数 — （廃止前は）時間トリガーで5分ごとに実行
// ============================================================
function checkAbsenceAndSendMail() {
  if (!ABSENCE_MAIL_ENABLED) {
    console.warn('自動欠席メールは廃止済みのため送信しません。removeAbsenceTrigger() を実行してトリガーを削除してください');
    return;
  }
  const ss     = SpreadsheetApp.openById(CALENDAR_CONFIG.SPREADSHEET_ID);
  const today  = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 各シートを取得
  const profileSheet = ss.getSheetByName(CALENDAR_CONFIG.SHEET_PROFILE);
  const logSheet     = ss.getSheetByName(CONFIG.SHEET_LOG);
  const sentSheet    = ensureSheet(CONFIG.SHEET_SENT, ['日付', '送信キー', '生徒ID', '生徒名', '予定開始', '送信先']);

  // 生徒マスタを全取得（1行目はヘッダーなのでスキップ）
  const profiles = profileSheet.getDataRange().getValues().slice(1);

  // 本日の入室済み生徒IDセットを作成
  const checkedInIds = getCheckedInIds(logSheet, todayStr);

  // 本日の送信済みキーセットを作成（重複送信防止）
  const sentKeys = getSentKeys(sentSheet, todayStr);

  // 各生徒のカレンダーをチェック
  profiles.forEach(row => {
    const studentId   = String(row[IDX_PROFILE.ID]).trim();
    const studentName = String(row[IDX_PROFILE.NAME]).trim();
    const studentMail = String(row[IDX_PROFILE.STUDENT_EMAIL]).trim();
    const parentMail  = String(row[IDX_PROFILE.PARENT_EMAIL]).trim();
    const calendarId  = String(row[IDX_PROFILE.CALENDAR_ID]).trim();

    // カレンダーIDまたはメールが未設定の生徒はスキップ
    if (!calendarId || calendarId === '' || !studentMail || studentMail === '') return;

    // 生徒のカレンダーから本日の予定を取得
    let calendar;
    try {
      calendar = CalendarApp.getCalendarById(calendarId);
    } catch(e) {
      console.log(`カレンダー取得失敗: ${studentName} (${calendarId})`);
      return;
    }
    if (!calendar) return;

    const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
    const endOfDay   = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    const events     = calendar.getEvents(startOfDay, endOfDay);

    events.forEach(event => {
      const eventStart = event.getStartTime();
      const minutesElapsed = (today - eventStart) / 1000 / 60;

      // 予定開始からCONFIG.ABSENCE_MINUTES分以上経過しているか
      if (minutesElapsed < CONFIG.ABSENCE_MINUTES) return;

      // 送信済みキー（生徒ID + 予定開始時刻）で重複チェック
      const eventStartStr = Utilities.formatDate(eventStart, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const sentKey = `${studentId}_${eventStartStr}`;
      if (sentKeys.has(sentKey)) return;

      // 入室済みであればスキップ
      if (checkedInIds.has(studentId)) return;

      // ===== 欠席判定 → メール送信 =====
      const recipients = buildRecipients(studentMail, parentMail);
      const subject    = buildSubject(studentName, eventStart);
      const body       = buildBody(studentName, eventStart);

      sendMailSafely(recipients, subject, body, `${studentName} (${eventStartStr})`, () => {
        logSent(sentSheet, todayStr, sentKey, studentId, studentName, eventStartStr, recipients);
      });
    });
  });
}

// ============================================================
// ヘルパー関数群
// ============================================================

/** 本日入室済みの生徒IDセットを返す */
function getCheckedInIds(logSheet, todayStr) {
  const ids = new Set();
  const data = logSheet.getDataRange().getValues();
  data.slice(1).forEach(row => {
    const rowDate = Utilities.formatDate(new Date(row[CONFIG.LOG_COL_DATE - 1]), 'Asia/Tokyo', 'yyyy/MM/dd');
    const checkin = row[CONFIG.LOG_COL_CHECKIN - 1];
    if (rowDate === todayStr && checkin !== '') {
      ids.add(String(row[CONFIG.LOG_COL_ID - 1]).trim());
    }
  });
  return ids;
}

/** 本日送信済みのキーセットを返す */
function getSentKeys(sentSheet, todayStr) {
  const keys = new Set();
  const data = sentSheet.getDataRange().getValues();
  data.slice(1).forEach(row => {
    const rowDateStr = Utilities.formatDate(row[0],'Asia/Tokyo','yyyy/MM/dd');
    if (rowDateStr === todayStr) keys.add(String(row[1]));
  });
  return keys;
}

/** 送信先メールアドレスを組み立てる（カンマ区切り） */
function buildRecipients(studentMail, parentMail) {
  const list = [studentMail];
  if (parentMail && parentMail !== '' && parentMail !== 'undefined') {
    list.push(parentMail);
  }
  return list.join(',');
}

/** メール送信を安全に行う共通ヘルパー（欠席確認・催促メール共通、成功時のみ onSuccess でログ記録） */
function sendMailSafely(recipients, subject, body, label, onSuccess) {
  try {
    GmailApp.sendEmail(recipients, subject, body);
    onSuccess();
    console.log(`${label} 送信完了`);
  } catch (e) {
    console.error(`${label} 送信失敗 — ${e.message}`);
  }
}

/** メール件名を組み立てる */
function buildSubject(studentName, eventStart) {
  const timeStr = Utilities.formatDate(eventStart, 'Asia/Tokyo', 'HH:mm');
  return `【自習室】${studentName}さんの来室が確認できません（${timeStr}〜）`;
}

/** メール本文を組み立てる */
function buildBody(studentName, eventStart) {
  const timeStr = Utilities.formatDate(eventStart, 'Asia/Tokyo', 'HH:mm');
  return `${studentName} さん（および保護者の方）

本日 ${timeStr} からの自習室利用予定になっておりますが、
まだ来室が確認できておりません。

欠席の場合はご連絡いただけますと幸いです。
体調など問題ないことをお祈りしております。

──────────────────────
SSS Education
──────────────────────`;
}

/** 送信済みログに1行追記 */
function logSent(sentSheet, todayStr, sentKey, studentId, studentName, eventStartStr, recipients) {
  sentSheet.appendRow([todayStr, sentKey, studentId, studentName, eventStartStr, recipients]);
}

// ============================================================
// 自動欠席メールのトリガー削除（GASエディタで手動実行）
// ※ インストール型トリガーは作成したアカウントごとに管理されるため、
//   過去にトリガーを設定したことがある各アカウントで1回ずつ実行すること。
//   送信済みログシートは削除しない（過去の記録として残す）。
// ============================================================
function removeAbsenceTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAbsenceAndSendMail') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  console.log(`自動欠席メールのトリガーを ${removed} 件削除しました`);
  return removed;
}

// 旧セットアップ関数。誤って実行してもトリガーを作らず、削除だけ行う
function setupTrigger() {
  console.warn('自動欠席メールは廃止済みです。トリガーは作成せず、残っているものを削除します');
  return removeAbsenceTrigger();
}

// ============================================================
// SSS Education 自習室管理システム
// カレンダー未入力 催促メール 自動送信モジュール
// ============================================================
// 【前提】
//   - absence_check.gs と同じGASプロジェクト内に追加
//   - 公開プロフィールシートに「カレンダーID」「メールアドレス」列が設定済み
//   - 月曜日の朝（例：8:00）に時間トリガーで実行
// ============================================================

// ========== 設定値（absence_check.gsのCONFIGと共通化推奨） ==========
const REMINDER_CONFIG = {
  SHEET_SENT:       '催促送信済みログ',       // 催促用の送信済みログ（別シート）
  SHEET_SENT_HEADERS: ['送信キー', '生徒ID', '生徒名', 'バージョン', '対象週', '送信先', '送信日時'],

  // VERSION_B用：直近何日間に予定がなければ送るか
  VERSION_B_DAYS_THIS_WEEK: 7,   // 今週（7日間）
  VERSION_B_DAYS_NEXT_WEEK: 14,  // 翌週（14日間）
};

// ============================================================
// Version A：翌週分（月〜日）が1件も入力されていない生徒に送る
// ============================================================
function sendReminderVersionA() {
  const ss = SpreadsheetApp.openById(CALENDAR_CONFIG.SPREADSHEET_ID);
  const profiles  = ss.getSheetByName(CALENDAR_CONFIG.SHEET_PROFILE)
                      .getDataRange().getValues().slice(1);
  const sentSheet = ensureSheet(REMINDER_CONFIG.SHEET_SENT, REMINDER_CONFIG.SHEET_SENT_HEADERS);

  // 翌週の月曜〜日曜を算出
  const today     = new Date();
  const nextMonday = getNextMonday(today);
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  nextSunday.setHours(23, 59, 59);

  const weekKey = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'yyyy/MM/dd'); // 重複防止キー
  const sentKeys = getReminderSentKeys(sentSheet);

  profiles.forEach(row => {
    const studentId  = String(row[IDX_PROFILE.ID]).trim();
    const name       = String(row[IDX_PROFILE.NAME]).trim();
    const email      = String(row[IDX_PROFILE.STUDENT_EMAIL]).trim();
    const parentMail = String(row[IDX_PROFILE.PARENT_EMAIL]).trim();
    const calId      = String(row[IDX_PROFILE.CALENDAR_ID]).trim();

    if (!calId || !email) return;

    const sentKey = `A_${studentId}_${weekKey}`;
    if (sentKeys.has(sentKey)) return; // 送信済みはスキップ

    let calendar;
    try { calendar = CalendarApp.getCalendarById(calId); } catch(e) { return; }
    if (!calendar) return;

    const events = calendar.getEvents(nextMonday, nextSunday);

    if (events.length === 0) {
      // 翌週に予定が1件もない → 催促メール送信
      const recipients = buildRecipients(email, parentMail);
      const subject    = buildReminderSubjectA(name, nextMonday, nextSunday);
      const body       = buildReminderBodyA(name, nextMonday, nextSunday);
      sendMailSafely(recipients, subject, body, `[Version A] ${name}`, () => {
        logReminderSent(sentSheet, sentKey, studentId, name, 'A', weekKey, recipients);
      });
    }
  });
}

// ============================================================
// Version B：今週 OR 翌週のどちらかが空の生徒に送る
// ============================================================
function sendReminderVersionB() {
  const ss = SpreadsheetApp.openById(CALENDAR_CONFIG.SPREADSHEET_ID);
  const profiles  = ss.getSheetByName(CALENDAR_CONFIG.SHEET_PROFILE)
                      .getDataRange().getValues().slice(1);
  const sentSheet = ensureSheet(REMINDER_CONFIG.SHEET_SENT, REMINDER_CONFIG.SHEET_SENT_HEADERS);

  const today      = new Date();
  const thisMonday = getThisMonday(today);
  const thisSunday = new Date(thisMonday);
  thisSunday.setDate(thisMonday.getDate() + 6);
  thisSunday.setHours(23, 59, 59);

  const nextMonday = new Date(thisMonday);
  nextMonday.setDate(thisMonday.getDate() + 7);
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  nextSunday.setHours(23, 59, 59);

  const weekKey  = Utilities.formatDate(thisMonday, 'Asia/Tokyo', 'yyyy/MM/dd');
  const sentKeys = getReminderSentKeys(sentSheet);

  profiles.forEach(row => {
    const studentId  = String(row[IDX_PROFILE.ID]).trim();
    const name       = String(row[IDX_PROFILE.NAME]).trim();
    const email      = String(row[IDX_PROFILE.STUDENT_EMAIL]).trim();
    const parentMail = String(row[IDX_PROFILE.PARENT_EMAIL]).trim();
    const calId      = String(row[IDX_PROFILE.CALENDAR_ID]).trim();

    if (!calId || !email) return;

    const sentKey = `B_${studentId}_${weekKey}`;
    if (sentKeys.has(sentKey)) return;

    let calendar;
    try { calendar = CalendarApp.getCalendarById(calId); } catch(e) { return; }
    if (!calendar) return;

    const thisWeekEvents = calendar.getEvents(thisMonday, thisSunday);
    const nextWeekEvents = calendar.getEvents(nextMonday, nextSunday);

    const thisWeekEmpty = thisWeekEvents.length === 0;
    const nextWeekEmpty = nextWeekEvents.length === 0;

    if (thisWeekEmpty || nextWeekEmpty) {
      const recipients = buildRecipients(email, parentMail);
      const subject    = buildReminderSubjectB(name, thisWeekEmpty, nextWeekEmpty);
      const body       = buildReminderBodyB(name, thisMonday, thisSunday, nextMonday, nextSunday, thisWeekEmpty, nextWeekEmpty);
      sendMailSafely(recipients, subject, body, `[Version B] ${name} (今週空:${thisWeekEmpty}, 翌週空:${nextWeekEmpty})`, () => {
        logReminderSent(sentSheet, sentKey, studentId, name, 'B', weekKey, recipients);
      });
    }
  });
}

// ============================================================
// 日付ユーティリティ
// ============================================================

/** 今週の月曜日（0:00:00）を返す（実装は コード.gs の getMonday に一本化） */
function getThisMonday(date) {
  return getMonday(date);
}

/** 翌週の月曜日（0:00:00）を返す */
function getNextMonday(date) {
  const thisMonday = getThisMonday(date);
  const next = new Date(thisMonday);
  next.setDate(thisMonday.getDate() + 7);
  return next;
}

// ============================================================
// メール文面
// ============================================================

function buildReminderSubjectA(name, nextMonday, nextSunday) {
  const from = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'M/d');
  const to   = Utilities.formatDate(nextSunday,  'Asia/Tokyo', 'M/d');
  return `【自習室】来週（${from}〜${to}）の自習予定を入力してください`;
}

function buildReminderBodyA(name, nextMonday, nextSunday) {
  const from = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'M月d日');
  const to   = Utilities.formatDate(nextSunday,  'Asia/Tokyo', 'M月d日');
  return `${name} さん

来週（${from}〜${to}）の自習室利用予定がまだ入力されていません。

Googleカレンダーに来週の予定を入力をお願いします。
自習の習慣化のために、先生も一緒に応援しています！

──────────────────────
SSS Education
──────────────────────`;
}

function buildReminderSubjectB(name, thisWeekEmpty, nextWeekEmpty) {
  const parts = [];
  if (thisWeekEmpty) parts.push('今週');
  if (nextWeekEmpty) parts.push('来週');
  return `【自習室】${parts.join('・')}の自習予定を入力してください`;
}

function buildReminderBodyB(name, thisMonday, thisSunday, nextMonday, nextSunday, thisWeekEmpty, nextWeekEmpty) {
  const lines = [`${name} さん\n`];
  if (thisWeekEmpty) {
    const from = Utilities.formatDate(thisMonday, 'Asia/Tokyo', 'M月d日');
    const to   = Utilities.formatDate(thisSunday,  'Asia/Tokyo', 'M月d日');
    lines.push(`・今週（${from}〜${to}）の予定が入力されていません`);
  }
  if (nextWeekEmpty) {
    const from = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'M月d日');
    const to   = Utilities.formatDate(nextSunday,  'Asia/Tokyo', 'M月d日');
    lines.push(`・来週（${from}〜${to}）の予定が入力されていません`);
  }
  lines.push('\nGoogleカレンダーに予定を入力をお願いします。');
  lines.push('自習の習慣化のために、先生も一緒に応援しています！');
  lines.push('\n──────────────────────\nSSS Education\n──────────────────────');
  return lines.join('\n');
}

// ============================================================
// 送信済みログ管理
// ============================================================

function getReminderSentKeys(sentSheet) {
  const keys = new Set();
  sentSheet.getDataRange().getValues().slice(1).forEach(row => {
    keys.add(String(row[0]));
  });
  return keys;
}

function logReminderSent(sentSheet, sentKey, studentId, name, version, weekKey, recipients) {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sentSheet.appendRow([sentKey, studentId, name, version, weekKey, recipients, now]);
}

// ============================================================
// トリガー登録：月曜日 8:00 に実行（初回のみ手動実行）
//
// 【注意】Version A・Bは「どちらを採用するか検討中」の2案であり、
// このまま両方登録すると毎週月曜、対象条件に一致した生徒には
// 催促メールが2通（A・B）届く。GASの時間トリガーは分単位を
// 指定できないため実際はほぼ同時刻に実行される
// （下のコメントの「8:05にずらす」は現状コード上は未実装）。
// 本運用に入れる際はどちらか一方だけを登録すること。
// ============================================================
// 使う方を 'A' か 'B' で指定する（監査レポート v2 R-04）。
// 未指定（null）のままでは setupReminderTrigger() は何も登録しない。
//   A：来週の予定が1件もない生徒に送る
//   B：今週か来週のどちらかが空の生徒に送る
const REMINDER_VERSION = 'A';

function setupReminderTrigger() {
  if (REMINDER_VERSION !== 'A' && REMINDER_VERSION !== 'B') {
    throw new Error('REMINDER_VERSION に A か B を指定してから実行してください（両方の同時登録はしません）');
  }

  // 既存トリガー（A・B両方）を削除してから、選んだ方だけを登録する
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'sendReminderVersionA' || fn === 'sendReminderVersionB') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('sendReminderVersion' + REMINDER_VERSION)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  console.log(`催促メールトリガーを登録しました（Version ${REMINDER_VERSION}・毎週月曜 8時台）`);
}

// 現在登録されているトリガーの一覧をログに出す（停止確認用）
function listTriggers() {
  const list = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  console.log('このアカウントで登録中のトリガー: ' + (list.length ? list.join(', ') : '（なし）'));
  return list;
}