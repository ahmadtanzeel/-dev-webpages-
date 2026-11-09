// ====================================================================
// 【列構成 最新版】インデックス定義（公開プロフィール / 個人情報マスタ共通）
// A:ID / B:生徒メアド / C:保護者メアド / D:氏名 / E:ニックネーム
// F:目標時間 / G:URLトークン / H:模試名 / I:模試日 / K:カレンダーID
// ====================================================================
const IDX_PROFILE = {
  ID: 0,             // A列
  STUDENT_EMAIL: 1,  // B列
  PARENT_EMAIL: 2,   // C列
  NAME: 3,           // D列
  NICKNAME: 4,       // E列
  GOAL_HOURS: 5,     // F列
  TOKEN: 6,          // G列
  EXAM_NAME: 7,      // H列
  EXAM_DATE: 8,      // I列
  CALENDAR_ID: 10    // K列
};
const IDX_PERSONAL = IDX_PROFILE;

// ====================================================================
// 🔔 Google Chat Webhook URL
// ====================================================================
const GCHAT_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('GCHAT_WEBHOOK_URL');

// ====================================================================
// 🔗 生徒マイページのベースURL
// ====================================================================
const STUDENT_DASHBOARD_BASE_URL = PropertiesService.getScriptProperties().getProperty('STUDENT_DASHBOARD_BASE_URL');

// ====================================================================
// 📘 小テスト用スプレッドシートの設定
// ====================================================================
const TEST_SS_ID = PropertiesService.getScriptProperties().getProperty('TEST_SS_ID');
const TEST_TEMPLATE_NAME = 'テンプレート';

// ====================================================================
// 管理シート（打刻ログ）の列インデックス
// A列:日付 / B列:ID / D列:入室時刻 / E列:退室時刻 / F列:エール数
// ====================================================================
const IDX_LOG = {
  DATE: 0,    // A列
  ID: 1,      // B列
  IN: 3,      // D列
  OUT: 4,     // E列
  CHEERS: 5   // F列
};

// ====================================================================
// 📚 参考書おすすめ機能用シート
// ====================================================================
const SHEET_BOOK_POSTS = '参考書投稿';
const SHEET_BOOK_REACTIONS = '参考書リアクション';
const BOOK_POSTS_HEADERS = ['ID', '投稿日時', '生徒ID', '書名', '評価', 'コメント'];
const BOOK_REACTIONS_HEADERS = ['投稿ID', '生徒ID', '絵文字', '日時'];

// ====================================================================
// カレンダー取得用シート（列定義は IDX_PROFILE に一本化）
// ====================================================================
const CALENDAR_CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('CALENDAR_SPREADSHEET_ID'),
  SHEET_PROFILE: '公開プロフィール',
};

// ====================================================================
// ① アクセス振り分け処理
// ====================================================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token = params.token;

  try {
    if (token) {
      let data;
      switch (action) {
        case 'stats': {
          data = getStudentStats(token);
          break;
        }
        case 'heatmap': {
          data = getHeatmap(token);
          break;
        }
        case 'bookPosts': {
          data = getBookPosts(token);
          break;
        }
        case 'getCalendarEvents': {
          data = getCalendarEventsForStudent(token, params.year, params.month);
          break;
        }
        //以下テスト環境でのダッシュボード表示用
        /*case 'dashboard': {
          const dashboad = HtmlService.createTemplateFromFile('dashboard');
          dashboad.token = token;
          return dashboad.evaluate();
        }*/
        //ここまで
      }
      return jsonResponse({ ok: true, data: data });
      /*if (action === 'stats' && token) {
        const data = getStudentStats(token);
        return jsonResponse({ ok: true, data: data });
      }
      if (action === 'heatmap' && token) {
        const data = getHeatmap(token);
        return jsonResponse({ ok: true, data: data });
      }
      if (action === 'bookPosts' && token) {
        const data = getBookPosts(token);
        return jsonResponse({ ok: true, data: data });
      }*/
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SSS Education 自習室受付')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;

    if (action === 'saveSettings') {
      const msg = saveStudentSettings(body.token, body.examName, body.examDate);
      return jsonResponse({ ok: true, message: msg });
    }
    if (action === 'scan') {
      // 受付打刻は { ok, code, message } をそのまま返す（失敗時は ok:false）
      return jsonResponse(processScan(body.id, body.kioskKey, body.requestId));
    }
    // 📚 参考書おすすめ投稿
    if (action === 'postBook') {
      const msg = postBookPost(body.token, body.bookName, body.rating, body.comment);
      return jsonResponse({ ok: true, message: msg });
    }
    // 📚 参考書おすすめへのリアクション
    if (action === 'reactBook') {
      const msg = reactBookPost(body.token, body.postId, body.emoji);
      return jsonResponse({ ok: true, message: msg });
    }
    // 📅 自習室予定のカレンダー登録
    if (action === 'addCalendarEvents') {
      const msg = addCalendarEventsForStudent(body.token, body.events);
      return jsonResponse({ ok: true, message: msg });
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====================================================================
// 🛠 ユーティリティ：トークン→生徒情報変換
// ====================================================================
function findStudentByToken(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  for (let i = 1; i < profileData.length; i++) {
    if (String(profileData[i][IDX_PROFILE.TOKEN]).trim() === token) {
      return {
        id: String(profileData[i][IDX_PROFILE.ID]).trim(),
        name: profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || ''
      };
    }
  }
  return null;
}

// ====================================================================
// 🛠 ユーティリティ：シート確保（無ければ作成）
// ====================================================================
function ensureSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ====================================================================
// 🛠 ユーティリティ：自由入力の文字列をシートに安全に書くための変換
//   先頭が = + - @（とタブ・改行）の文字列は、数式として解釈されないよう
//   先頭に ' を付けて「文字列」として保存する（' はセルの表示・取得値には含まれない）
// ====================================================================
function safeCellText(value) {
  const s = String(value === null || value === undefined ? '' : value);
  return /^[=+\-@\t\r\n]/.test(s) ? "'" + s : s;
}

// 'yyyy-MM-dd' 形式で、実在する日付かどうか
function isValidDateString(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) return false;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return d.getFullYear() === Number(m[1]) && d.getMonth() === Number(m[2]) - 1 && d.getDate() === Number(m[3]);
}

// ====================================================================
// ② ダッシュボード用データ集計
// ====================================================================
function getStudentStats(studentToken) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(studentToken);
  if (cachedData) return JSON.parse(cachedData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName('公開プロフィール');
  const logSheet = ss.getSheetByName('管理シート');
  const settingsSheet = ss.getSheetByName('生徒設定') || ss.insertSheet('生徒設定');

  const profileData = profileSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  const settingsData = settingsSheet.getDataRange().getValues();

  let targetId = null;
  let targetNickname = "学習者";
  let weeklyGoalHours = 15;
  let customExam = { name: "", date: "" };
  let nickMap = {};

  for (let i = 1; i < profileData.length; i++) {
    let pId = String(profileData[i][IDX_PROFILE.ID]).trim();
    let pToken = String(profileData[i][IDX_PROFILE.TOKEN]).trim();
    let pName = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || pId;
    nickMap[pId] = pName;

    if (pToken === studentToken) {
      targetId = pId;
      targetNickname = pName;
      let goalRaw = profileData[i][IDX_PROFILE.GOAL_HOURS];
      if (goalRaw && !isNaN(parseFloat(goalRaw))) weeklyGoalHours = parseFloat(goalRaw);
    }
  }

  if (!targetId) throw new Error("無効なURLです。管理者に連絡してください。");

  for (let i = 1; i < settingsData.length; i++) {
    if (String(settingsData[i][0]).trim() === targetId) {
      customExam.name = settingsData[i][1];
      let sDate = settingsData[i][2];
      customExam.date = sDate instanceof Date ? Utilities.formatDate(sDate, "JST", "yyyy/MM/dd") : sDate;
      break;
    }
  }

  const toDate = (val) => {
    if (!val) return null;
    if (val instanceof Date) return val;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  };

  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const thisMondayMs = getMonday(new Date()).getTime();
  const lastMondayMs = thisMondayMs - 7 * 24 * 3600000;

  let totalMs = 0, todayMs = 0, weekMs = 0, lastWeekMs = 0, todayCheers = 0;
  let dailyMap = {}, weeklyMap = {}, monthlyMap = {}, weeklyRankingMap = {};
  let recentActions = [], historyList = [];
  let uniqueDates = new Set();

  // 過去30日 vs その前の30日（傾向比較）
  let last30dMs = 0;
  let prev30dMs = 0;
  const ms30d = 30 * 24 * 3600000;

  for (let i = logData.length - 1; i >= 1; i--) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = toDate(logData[i][IDX_LOG.IN]);
    let outTime = toDate(logData[i][IDX_LOG.OUT]);

    if (inTime) {
      let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
      let diffMs = (outTime) ? outTime.getTime() - inTime.getTime() : 0;
      if (diffMs < 0) diffMs = 0;

      if (rowId === targetId) {
        if (diffMs > 0) {
          totalMs += diffMs;
          if (logDateMs === todayStart) todayMs += diffMs;
          if (logDateMs >= (todayStart - 7 * 24 * 60 * 60 * 1000)) weekMs += diffMs;

          if (inTime.getTime() >= now.getTime() - ms30d) {
            last30dMs += diffMs;
          } else if (inTime.getTime() >= now.getTime() - 2 * ms30d) {
            prev30dMs += diffMs;
          }

          if (logDateMs >= lastMondayMs && logDateMs < thisMondayMs) {
            lastWeekMs += diffMs;
          }

          let dayKey = Utilities.formatDate(inTime, "JST", "MM/dd");
          dailyMap[dayKey] = (dailyMap[dayKey] || 0) + diffMs;

          let monday = getMonday(inTime);
          let weekKey = Utilities.formatDate(monday, "JST", "MM/dd") + "週";
          weeklyMap[weekKey] = (weeklyMap[weekKey] || 0) + diffMs;

          let monthKey = Utilities.formatDate(inTime, "JST", "yyyy/MM");
          monthlyMap[monthKey] = (monthlyMap[monthKey] || 0) + diffMs;

          uniqueDates.add(Utilities.formatDate(inTime, "JST", "yyyy/MM/dd"));

          historyList.push({
            date: Utilities.formatDate(inTime, "JST", "MM/dd"),
            in: Utilities.formatDate(inTime, "JST", "HH:mm"),
            out: outTime ? Utilities.formatDate(outTime, "JST", "HH:mm") : "---",
            time: formatTime(diffMs)
          });
        }
        if (logDateMs === todayStart) {
          todayCheers += parseInt(logData[i][IDX_LOG.CHEERS] || 0);
        }
      }

      // 今週ランキング集計
      if (logDateMs >= thisMondayMs && diffMs > 0) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + diffMs;
      }

      if (recentActions.length < 5) {
        recentActions.push({
          id: rowId,
          name: nickMap[rowId] || "学習者",
          action: (outTime) ? "退室" : "入室",
          time: Utilities.formatDate((outTime ? outTime : inTime), "JST", "HH:mm")
        });
      }
    }
  }

  // 入室中（未退室）を暫定加算
  for (let i = logData.length - 1; i >= 1; i--) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = toDate(logData[i][IDX_LOG.IN]);
    let outTime = toDate(logData[i][IDX_LOG.OUT]);
    if (!inTime) continue;
    let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
    if (logDateMs < thisMondayMs) break;

    if (!outTime) {
      let liveMs = now.getTime() - inTime.getTime();
      if (liveMs > 0 && liveMs < 12 * 3600000) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + liveMs;
      }
    }
  }

  let top5Ranking = Object.keys(weeklyRankingMap).map(id => {
    return { name: nickMap[id] || "学習者", hours: (weeklyRankingMap[id] / 3600000).toFixed(1) };
  }).filter(r => parseFloat(r.hours) > 0)
    .sort((a, b) => parseFloat(b.hours) - parseFloat(a.hours))
    .slice(0, 5);

  // 比較データ（先週比・前月比）
  const compareData = {
    thisWeekHours: (weekMs / 3600000).toFixed(1),
    lastWeekHours: (lastWeekMs / 3600000).toFixed(1),
    last30dHours: (last30dMs / 3600000).toFixed(1),
    prev30dHours: (prev30dMs / 3600000).toFixed(1)
  };

  // 小テスト
  const testProgress = getOrCreateTestProgress(targetNickname, targetId);

  const formatChartData = (map) => {
    let labels = Object.keys(map).sort();
    let values = labels.map(k => (map[k] / 3600000).toFixed(1));
    return { labels, values };
  };

  const totalHours = totalMs / 3600000;

  const resultData = {
    name: targetNickname,
    id: targetId,
    today: formatTime(todayMs),
    week: formatTime(weekMs),
    total: formatTime(totalMs),
    totalHours: totalHours,
    daily: formatChartData(dailyMap),
    weekly: formatChartData(weeklyMap),
    monthly: formatChartData(monthlyMap),
    rank: calculateRank(totalHours),
    community: { ranking: top5Ranking, feed: recentActions },
    weeklyGoal: { currentHours: (weekMs / 3600000).toFixed(1), targetHours: weeklyGoalHours, percent: Math.min(100, ((weekMs / 3600000) / weeklyGoalHours) * 100).toFixed(1) },
    history: historyList.slice(0, 30),
    streak: { totalDays: uniqueDates.size, active: calculateStreak(uniqueDates) },
    tests: testProgress,
    exam: customExam,
    cheers: todayCheers,
    compare: compareData,
    level: calculateLevel(totalHours)
  };

  cache.put(studentToken, JSON.stringify(resultData), 900);
  return resultData;
}

// ====================================================================
// 🔥 ヒートマップ：過去7日 × 24時間（各時間帯ごとの在室分数）
// ====================================================================
function getHeatmap(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logData = ss.getSheetByName('管理シート').getDataRange().getValues();
  const targetId = student.id;

  const now = new Date();
  // 過去7日分の日付配列（今日を最後にする）
  const days = [];
  const dowJa = ['日', '月', '火', '水', '木', '金', '土'];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - i);
    days.push({
      date: Utilities.formatDate(d, "JST", "yyyy-MM-dd"),
      label: Utilities.formatDate(d, "JST", "MM/dd"),
      dow: dowJa[d.getDay()],
      isToday: i === 0,
      hours: new Array(24).fill(0)  // 各時間帯の在室分数
    });
  }

  // 開始（7日前の0時）と終了（今）の範囲
  const rangeStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 6).getTime();

  for (let i = 1; i < logData.length; i++) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    if (rowId !== targetId) continue;

    let inT = logData[i][IDX_LOG.IN];
    let outT = logData[i][IDX_LOG.OUT];
    if (!(inT instanceof Date)) continue;

    // 未退室はとりあえず今をendとする（上限12hで安全策）
    let endT;
    if (outT instanceof Date) {
      endT = outT;
    } else {
      const provisional = new Date(inT.getTime() + 12 * 3600000);
      endT = provisional < now ? provisional : now;
    }
    if (endT <= inT) continue;
    if (endT.getTime() < rangeStart) continue;

    // 範囲外（古すぎる）はスキップ
    if (inT.getTime() < rangeStart - 24 * 3600000) continue;

    // 1時間ブロックごとに、該当日のhoursに分数を加算
    let cur = new Date(inT.getTime());
    while (cur < endT) {
      // 過去7日範囲外ならスキップして次の時間へ
      const dayKey = Utilities.formatDate(cur, "JST", "yyyy-MM-dd");
      const dayObj = days.find(d => d.date === dayKey);
      if (!dayObj) {
        // 範囲外
        const nextHour = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate(), cur.getHours() + 1, 0, 0);
        cur = nextHour;
        continue;
      }

      const h = cur.getHours();
      const nextHour = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate(), h + 1, 0, 0);
      const sliceEnd = nextHour < endT ? nextHour : endT;
      const minutes = (sliceEnd.getTime() - cur.getTime()) / 60000;
      if (minutes > 0) dayObj.hours[h] += minutes;
      cur = sliceEnd;
    }
  }

  // 各dayObj.hoursを整数（分）に丸める
  days.forEach(d => {
    d.hours = d.hours.map(m => Math.round(m));
  });

  return { days };
}

// ====================================================================
// 📚 参考書おすすめ：取得（最新50件）
// ====================================================================
function getBookPosts(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  ensureSheet(SHEET_BOOK_POSTS, BOOK_POSTS_HEADERS);
  ensureSheet(SHEET_BOOK_REACTIONS, BOOK_REACTIONS_HEADERS);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postSheet = ss.getSheetByName(SHEET_BOOK_POSTS);
  const reactionSheet = ss.getSheetByName(SHEET_BOOK_REACTIONS);

  // 名前マップ
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  const nickMap = {};
  for (let i = 1; i < profileData.length; i++) {
    nickMap[String(profileData[i][IDX_PROFILE.ID]).trim()] =
      profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || '学習者';
  }

  const postData = postSheet.getDataRange().getValues();
  const reactionData = reactionSheet.getDataRange().getValues();

  // リアクション集計
  const reactionMap = {}; // postId → { emoji → count }
  const myReactions = {}; // postId → [emojis]
  for (let i = 1; i < reactionData.length; i++) {
    const postId = String(reactionData[i][0]).trim();
    const userId = String(reactionData[i][1]).trim();
    const emoji = String(reactionData[i][2]).trim();
    if (!postId || !emoji) continue;
    if (!reactionMap[postId]) reactionMap[postId] = {};
    reactionMap[postId][emoji] = (reactionMap[postId][emoji] || 0) + 1;
    if (userId === student.id) {
      if (!myReactions[postId]) myReactions[postId] = [];
      myReactions[postId].push(emoji);
    }
  }

  const posts = [];
  for (let i = postData.length - 1; i >= 1 && posts.length < 50; i--) {
    const postId = String(postData[i][0]).trim();
    const postedAt = postData[i][1];
    const studentId = String(postData[i][2]).trim();
    const bookName = String(postData[i][3] || '').trim();
    const rating = parseInt(postData[i][4]) || 0;
    const comment = String(postData[i][5] || '').trim();

    if (!postId || !postedAt || !bookName) continue;

    posts.push({
      id: postId,
      time: postedAt instanceof Date ? Utilities.formatDate(postedAt, "JST", "MM/dd HH:mm") : String(postedAt),
      name: nickMap[studentId] || '学習者',
      isMine: studentId === student.id,
      bookName: bookName,
      rating: Math.max(1, Math.min(5, rating)),
      comment: comment,
      reactions: reactionMap[postId] || {},
      myReactions: myReactions[postId] || []
    });
  }

  return { posts };
}

// ====================================================================
// 📚 参考書おすすめ：投稿
// ====================================================================
function postBookPost(token, bookName, rating, comment) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  bookName = String(bookName || '').trim();
  comment = String(comment || '').trim();
  const r = parseInt(rating);

  if (!bookName) throw new Error("参考書名を入力してください");
  if (bookName.length > 80) throw new Error("参考書名は80文字以内にしてください");
  if (comment.length > 300) throw new Error("コメントは300文字以内にしてください");
  if (isNaN(r) || r < 1 || r > 5) throw new Error("おすすめ度は★1〜★5で選んでください");

  // 簡易NGワード
  const ngWords = ['死ね', 'バカ', 'うざい', 'きもい'];
  const checkTarget = bookName + ' ' + comment;
  for (let w of ngWords) {
    if (checkTarget.indexOf(w) >= 0) throw new Error("不適切な表現が含まれています");
  }

  const sheet = ensureSheet(SHEET_BOOK_POSTS, BOOK_POSTS_HEADERS);
  const postId = 'B' + Date.now() + '_' + Math.floor(Math.random() * 1000);
  sheet.appendRow([postId, new Date(), student.id, safeCellText(bookName), r, safeCellText(comment)]);

  return "投稿しました";
}

// ====================================================================
// 📚 参考書おすすめ：リアクション（トグル）
// ====================================================================
function reactBookPost(token, postId, emoji) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  postId = String(postId || '').trim();
  emoji = String(emoji || '').trim();
  if (!postId || !emoji) throw new Error("パラメータが不正です");

  const allowed = ['👍', '🔥', '💡', '👀', '✨'];
  if (allowed.indexOf(emoji) < 0) throw new Error("使えない絵文字です");

  const sheet = ensureSheet(SHEET_BOOK_REACTIONS, BOOK_REACTIONS_HEADERS);
  const data = sheet.getDataRange().getValues();

  // 既存チェック（トグル削除）
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === postId &&
      String(data[i][1]).trim() === student.id &&
      String(data[i][2]).trim() === emoji) {
      sheet.deleteRow(i + 1);
      return "removed";
    }
  }

  sheet.appendRow([postId, student.id, emoji, new Date()]);
  return "added";
}

// ====================================================================
// 🆕 レベル＆経験値計算
//   レベルNから次レベルまでの必要時間 = 5 + N * 2 時間
//   Lv1: 累計0-7h / Lv2: 7-16h / Lv3: 16-27h ...
// ====================================================================
function calculateLevel(totalHours) {
  let level = 1;
  let needed = 5 + level * 2;
  let acc = 0;

  while (totalHours - acc >= needed) {
    acc += needed;
    level++;
    needed = 5 + level * 2;
  }

  const currentExp = totalHours - acc;
  return {
    level: level,
    currentExp: +currentExp.toFixed(1),
    expForNext: +needed.toFixed(1),
    percent: Math.min(100, +(currentExp / needed * 100).toFixed(1))
  };
}

// ====================================================================
// ③ 設定保存
// ====================================================================
function saveStudentSettings(token, examName, examDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  let studentId = null;

  for (let i = 1; i < profileData.length; i++) {
    if (String(profileData[i][IDX_PROFILE.TOKEN]).trim() === token) {
      studentId = String(profileData[i][IDX_PROFILE.ID]).trim();
      break;
    }
  }
  // 以前は「認証エラー」を成功扱い（ok:true）で返していたため、例外にして ok:false で返す
  if (!studentId) throw new Error("無効なURLです。管理者に連絡してください。");

  examName = String(examName || '').trim();
  examDate = String(examDate || '').trim();
  if (!examName) throw new Error("模試名を入力してください");
  if (examName.length > 50) throw new Error("模試名は50文字以内にしてください");
  if (!isValidDateString(examDate)) throw new Error("模試の日付が正しくありません");
  const examYear = Number(examDate.slice(0, 4));
  const thisYear = new Date().getFullYear();
  if (examYear < thisYear - 1 || examYear > thisYear + 3) throw new Error("模試の日付の年が範囲外です");

  const settingsSheet = ensureSheet("生徒設定", ['生徒ID', '模試名', '模試日程']);
  const settingsData = settingsSheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < settingsData.length; i++) {
    if (String(settingsData[i][0]).trim() === studentId) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) targetRow = settingsSheet.getLastRow() + 1;

  settingsSheet.getRange(targetRow, 1).setValue(studentId);
  settingsSheet.getRange(targetRow, 2).setValue(safeCellText(examName));
  settingsSheet.getRange(targetRow, 3).setValue(examDate);

  CacheService.getScriptCache().remove(token);

  return "設定を保存しました！";
}

// ====================================================================
// ④ 受付打刻処理
// ====================================================================
// 同じ生徒の打刻を無視する間隔（ミリ秒）。直前の成功打刻からこの時間内は状態を切り替えない
const SCAN_DUPLICATE_GUARD_MS = 5000;
// 打刻処理の順番待ちの最大時間（ミリ秒）。5秒ガードとは別物
const SCAN_LOCK_WAIT_MS = 10000;
// 同じ requestId の再送を同じ結果で返す期間（秒）
const SCAN_REQUEST_CACHE_SEC = 600;

/**
 * 受付端末の確認。スクリプトプロパティ KIOSK_KEY と一致するキーを持つ端末だけが打刻できる。
 * KIOSK_KEY が未設定の場合は、安全側に倒して全て拒否する。
 */
function checkKioskKey(kioskKey) {
  const expected = PropertiesService.getScriptProperties().getProperty('KIOSK_KEY');
  if (!expected) return 'KIOSK_NOT_CONFIGURED';
  const given = String(kioskKey || '');
  // 長さが違えば不一致。同じ長さなら全文字を比べる（比較時間で中身を推測されにくくする）
  if (given.length !== expected.length) return 'KIOSK_UNAUTHORIZED';
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0 ? null : 'KIOSK_UNAUTHORIZED';
}

function scanResult(ok, code, message) {
  return { ok: ok, code: code, message: message, error: ok ? undefined : message };
}

/**
 * 受付打刻の入口。
 * 1. 受付キー確認 → 2. 生徒IDが名簿にあるか確認 → 3. ロックを取って打刻
 * → 4. ロック解放後に通知。いずれかで失敗した場合はシート・通知・キャッシュを変えない。
 */
function processScan(studentId, kioskKey, requestId) {
  const kioskError = checkKioskKey(kioskKey);
  if (kioskError === 'KIOSK_NOT_CONFIGURED') {
    return scanResult(false, kioskError, 'サーバー側で受付キーが設定されていません。管理者に連絡してください');
  }
  if (kioskError) {
    return scanResult(false, kioskError, '受付キーが正しくありません');
  }

  const cleanTargetId = String(studentId || '').trim();
  if (!cleanTargetId) return scanResult(false, 'EMPTY_ID', 'IDが読み込めませんでした');
  if (cleanTargetId.length > 64) return scanResult(false, 'UNKNOWN_ID', '登録されていないIDです');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 生徒名簿の確認（シートを書き換える前に行う）
  const pData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  let profile = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][IDX_PROFILE.ID]).trim() === cleanTargetId) {
      const rawParentEmail = pData[i][IDX_PROFILE.PARENT_EMAIL];
      profile = {
        name: pData[i][IDX_PROFILE.NAME] || pData[i][IDX_PROFILE.NICKNAME] || "学習者",
        token: String(pData[i][IDX_PROFILE.TOKEN] || '').trim() || null,
        parentEmail: (rawParentEmail && String(rawParentEmail).indexOf('@') > 0) ? String(rawParentEmail).trim() : null
      };
      break;
    }
  }
  if (!profile) return scanResult(false, 'UNKNOWN_ID', '登録されていないIDです');

  const cache = CacheService.getScriptCache();
  const reqKey = requestId ? 'scanreq_' + String(requestId).slice(0, 80) : null;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SCAN_LOCK_WAIT_MS)) {
    return scanResult(false, 'BUSY', '混み合っています。もう一度かざしてください');
  }

  let record;
  try {
    // 同じリクエストの再送なら、前回の結果をそのまま返す（通知もしない）
    if (reqKey) {
      const prev = cache.get(reqKey);
      if (prev) {
        const p = JSON.parse(prev);
        return scanResult(true, 'ALREADY_PROCESSED', p.message);
      }
    }

    record = writeScanRecord(ss, cleanTargetId, new Date());
    if (record.duplicate) {
      return scanResult(true, 'DUPLICATE', `${profile.name} さんは直前に打刻済みです`);
    }

    // 打刻結果を確定させてからロックを外す
    SpreadsheetApp.flush();
    record.message = `${profile.name} さんが ${record.actionType} しました！`;
    if (reqKey) cache.put(reqKey, JSON.stringify({ message: record.message }), SCAN_REQUEST_CACHE_SEC);
  } catch (e) {
    console.error('打刻処理で例外:', e);
    return scanResult(false, 'WRITE_FAILED', '記録に失敗しました。もう一度かざしてください');
  } finally {
    lock.releaseLock();
  }

  // ここから先は打刻が確定した後の処理（失敗しても打刻自体は成功扱い）
  if (profile.token) cache.remove(profile.token);

  if (profile.parentEmail) {
    try {
      notifyParent(profile.name, profile.parentEmail, record.actionType, record.time, record.studyMs);
    } catch (e) {
      console.error('[優先度2] 保護者メール処理で例外:', e);
    }
  }

  try {
    notifyGoogleChat(profile.name, record.actionType, record.time, record.studyMs);
  } catch (e) {
    console.error('[優先度3] Webhook処理で例外:', e);
  }

  return scanResult(true, record.actionType === '入室' ? 'CHECK_IN' : 'CHECK_OUT', record.message);
}

/**
 * 管理シートへの入退室の書き込み（ロック内で呼ぶこと）。
 * 同じ生徒の直前の打刻から SCAN_DUPLICATE_GUARD_MS 以内なら何もせず duplicate:true を返す。
 */
function writeScanRecord(ss, cleanTargetId, now) {
  const sheet = ss.getSheetByName("管理シート");
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");

  const sheetMaxRow = sheet.getLastRow();

  let actualLastDataRow = 1;
  if (sheetMaxRow >= 2) {
    const bColumn = sheet.getRange(2, 2, sheetMaxRow - 1, 1).getValues();
    for (let i = bColumn.length - 1; i >= 0; i--) {
      const v = bColumn[i][0];
      if (v !== "" && v !== null && v !== undefined) {
        actualLastDataRow = i + 2;
        break;
      }
    }
  }

  let existingRowIndex = -1;
  let existingInTime = null;
  let lastEventMs = null; // 本日この生徒が最後に打刻した時刻（入室・退室のどちらか新しい方）
  let searchSize = Math.min(100, Math.max(0, actualLastDataRow - 1));
  let attempt = 0;
  let maxAttempts = 3;

  while (existingRowIndex === -1 && attempt < maxAttempts && searchSize > 0) {
    let startRow = Math.max(2, actualLastDataRow - searchSize + 1);
    let numRows = actualLastDataRow - startRow + 1;
    if (numRows <= 0) break;

    const data = sheet.getRange(startRow, 1, numRows, 5).getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      const rowDate = data[i][IDX_LOG.DATE];
      const rowDateStr = rowDate instanceof Date ? Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") : "";
      if (rowDateStr !== todayStr) continue;
      if (String(data[i][IDX_LOG.ID]).trim() !== cleanTargetId) continue;
      if (lastEventMs === null) {
        const inT = data[i][IDX_LOG.IN], outT = data[i][IDX_LOG.OUT];
        const times = [inT, outT].filter(t => t instanceof Date).map(t => t.getTime());
        if (times.length > 0) lastEventMs = Math.max.apply(null, times);
      }
      if (data[i][IDX_LOG.OUT] === "" || data[i][IDX_LOG.OUT] === null) {
        existingRowIndex = startRow + i;
        existingInTime = data[i][IDX_LOG.IN];
        break;
      }
    }

    attempt++;
    searchSize *= 3;
  }

  // 5秒ガード：直前の打刻から5秒未満なら、入室/退室を切り替えない
  if (lastEventMs !== null && now.getTime() - lastEventMs < SCAN_DUPLICATE_GUARD_MS) {
    return { duplicate: true };
  }

  let actionType = "";
  let studyMs = 0;
  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, IDX_LOG.OUT + 1).setValue(now);
    actionType = "退室";
    if (existingInTime instanceof Date) {
      studyMs = now.getTime() - existingInTime.getTime();
      if (studyMs < 0) studyMs = 0;
    }
  } else {
    const targetRow = actualLastDataRow + 1;
    sheet.getRange(targetRow, 1, 1, 2).setValues([[now, cleanTargetId]]);
    sheet.getRange(targetRow, 4).setValue(now);
    actionType = "入室";
  }

  return { duplicate: false, actionType: actionType, studyMs: studyMs, time: now };
}

// ====================================================================
// ④-2 保護者向けメール送信
// ====================================================================
function notifyParent(studentName, parentEmail, actionType, datetime, studyMs) {
  const timeStr = Utilities.formatDate(datetime, "JST", "yyyy年MM月dd日 HH時mm分");
  const subject = `【SSS Education】${studentName}さんが${actionType}しました`;

  let body =
    `${studentName} さんの保護者様\n\n` +
    `いつもSSS Educationをご利用いただき、ありがとうございます。\n` +
    `お子様が以下の通り自習室に${actionType}されました。\n\n` +
    `  日時：${timeStr}\n` +
    `  行動：${actionType}\n`;

  if (actionType === "退室" && studyMs > 0) {
    body += `  本日の学習時間：${formatTime(studyMs)}\n`;
  }

  body +=
    `\n` +
    `引き続き、お子様の学習を見守ってまいります。\n` +
    `何かございましたら、お気軽に塾までお問い合わせください。\n\n` +
    `──────────────────\n` +
    `SSS Education 自習室管理システム\n` +
    `※このメールは自動送信されています。返信はできませんのでご了承ください。\n` +
    `──────────────────\n`;

  try {
    MailApp.sendEmail({
      to: parentEmail,
      subject: subject,
      body: body
    });
  } catch (err) {
    console.error(`保護者メール送信失敗 (${parentEmail}):`, err);
  }
}

// ====================================================================
// ④-3 Google Chat通知
// ====================================================================
function notifyGoogleChat(studentName, actionType, datetime, studyMs) {
  if (!GCHAT_WEBHOOK_URL) return;

  const timeStr = Utilities.formatDate(datetime, "JST", "HH:mm");
  let message;

  if (actionType === "入室") {
    message = `🟢 *${studentName}* さんが入室しました（${timeStr}）`;
  } else if (actionType === "退室") {
    const studyTime = studyMs > 0 ? formatTime(studyMs) : "—";
    message = `🔴 *${studentName}* さんが退室しました（${timeStr} / 学習時間: ${studyTime}）`;
  } else {
    message = `${studentName} さんが${actionType}しました（${timeStr}）`;
  }

  try {
    UrlFetchApp.fetch(GCHAT_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
  } catch (err) {
    console.error('Google Chat通知失敗:', err);
  }
}

// ====================================================================
// ⑥ 生徒マイページURLのメール下書き一括作成
// ====================================================================
function createDraftsByIds(idsString) {
  if (!idsString || !String(idsString).trim()) {
    return { success: false, message: "⚠ IDが入力されていません" };
  }

  const ids = String(idsString)
    .split(/[,、\s]+/)
    .map(s => s.trim())
    .filter(s => s);

  if (ids.length === 0) {
    return { success: false, message: "⚠ 有効なIDがありません" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();

  const profileMap = {};
  for (let i = 1; i < profileData.length; i++) {
    const id = String(profileData[i][IDX_PROFILE.ID]).trim();
    if (!id) continue;
    profileMap[id] = {
      name: profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || id,
      parentEmail: String(profileData[i][IDX_PROFILE.PARENT_EMAIL] || '').trim(),
      studentEmail: String(profileData[i][IDX_PROFILE.STUDENT_EMAIL] || '').trim(),
      token: String(profileData[i][IDX_PROFILE.TOKEN] || '').trim()
    };
  }

  let successList = [];
  let errorList = [];

  ids.forEach(id => {
    const profile = profileMap[id];
    if (!profile) {
      errorList.push(`✗ ${id}：プロフィールが見つかりません`);
      return;
    }
    if (!profile.token) {
      errorList.push(`✗ ${id} (${profile.name})：URLトークン未設定`);
      return;
    }

    const to = profile.studentEmail;
    if (!to || to.indexOf('@') < 0) {
      errorList.push(`✗ ${id} (${profile.name})：生徒メアド(B列)が未設定`);
      return;
    }

    const url = STUDENT_DASHBOARD_BASE_URL + encodeURIComponent(profile.token);
    const subject = `【SSS Education】${profile.name}さん専用 学習ダッシュボードのご案内`;
    const body =
      `${profile.name} さん\n\n` +
      `いつもSSS Educationでの学習、お疲れさまです。\n` +
      `あなた専用の学習ダッシュボードをご用意しました。\n\n` +
      `▼ あなたのマイページURL\n${url}\n\n` +
      `※ このURLはあなた個人専用です。他の方には共有しないようご注意ください。\n` +
      `※ スマートフォンの場合、ホーム画面に追加すると、アプリのように使えます。\n\n` +
      `【マイページでできること】\n` +
      `・自分の自習室入退室履歴の確認\n` +
      `・週間/月間の学習時間グラフ\n` +
      `・小テストの進捗状況\n` +
      `・受付用のQRコード（毎回これをかざせばOK）\n\n` +
      `不明な点があれば、講師までお気軽にお声がけください。\n\n` +
      `──────────────────\n` +
      `SSS Education\n` +
      `──────────────────\n`;

    try {
      GmailApp.createDraft(to, subject, body);
      successList.push(`✓ ${profile.name} (${id}) → ${to}`);
    } catch (e) {
      errorList.push(`✗ ${id} (${profile.name})：下書き作成失敗 - ${e.message}`);
    }
  });

  let message = '';
  if (successList.length > 0) {
    message += `✅ ${successList.length}件の下書きを作成しました\n\n`;
    message += successList.join('\n');
  }
  if (errorList.length > 0) {
    if (message) message += '\n\n';
    message += `⚠ ${errorList.length}件のエラー\n\n`;
    message += errorList.join('\n');
  }
  if (successList.length > 0) {
    message += `\n\n👉 Gmailの「下書き」フォルダで内容を確認してから送信してください。`;
  }

  return {
    success: successList.length > 0,
    message: message
  };
}

// ====================================================================
// ⑦ 小テスト進捗の取得
// ====================================================================
function getOrCreateTestProgress(studentName, studentId) {
  const testProgress = [];

  try {
    const testSs = SpreadsheetApp.openById(TEST_SS_ID);

    let testSheet = testSs.getSheetByName(studentName);
    if (!testSheet) {
      testSheet = testSs.getSheetByName(studentId);
    }

    if (!testSheet) {
      const templateSheet = testSs.getSheetByName(TEST_TEMPLATE_NAME);
      if (!templateSheet) {
        console.error(`テンプレートシート「${TEST_TEMPLATE_NAME}」が見つかりません`);
        return testProgress;
      }

      const lock = LockService.getScriptLock();
      const acquired = lock.tryLock(10000);
      if (!acquired) {
        console.warn('LockService取得失敗。シート自動生成をスキップします。');
        return testProgress;
      }

      try {
        testSheet = testSs.getSheetByName(studentName);
        if (!testSheet) {
          const copied = templateSheet.copyTo(testSs);
          copied.setName(studentName);
          testSheet = copied;
          console.log(`新規シート作成: ${studentName}`);
        }
      } catch (copyErr) {
        console.error('シート作成失敗:', copyErr.message);
        return testProgress;
      } finally {
        lock.releaseLock();
      }
    }

    if (testSheet) {
      const testData = testSheet.getDataRange().getValues();
      for (let i = 1; i < testData.length; i++) {
        const bookName = testData[i][0];
        if (!bookName) continue;
        let totalCount = 0, completedCount = 0;
        const blocks = [];
        for (let j = 1; j <= 30; j++) {
          const cellValue = testData[i][j];
          if (cellValue === true || cellValue === false) {
            totalCount++;
            if (cellValue === true) { completedCount++; blocks.push(true); }
            else { blocks.push(false); }
          }
        }
        if (totalCount > 0) {
          testProgress.push({
            name: bookName,
            total: totalCount,
            completed: completedCount,
            blocks: blocks
          });
        }
      }
    }
  } catch (e) {
    console.error("小テスト取得失敗:", e.message);
  }

  return testProgress;
}

// ====================================================================
// 補助関数群
// ====================================================================
function formatTime(ms) { return `${Math.floor(ms / 3600000)}時間${Math.floor((ms % 3600000) / 60000)}分`; }

/** 指定日を含む週の月曜日（0:00:00）を返す（コード.gs / AbsenceCheck.gs 共通） */
function getMonday(d) {
  let day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
  return new Date(d.getFullYear(), d.getMonth(), diff);
}

function calculateRank(hours) {
  if (hours >= 300) return { current: "SSS MASTER", remainHours: 0 };
  if (hours >= 150) return { current: "PLATINUM", remainHours: (300 - hours).toFixed(1) };
  if (hours >= 50) return { current: "GOLD", remainHours: (150 - hours).toFixed(1) };
  if (hours >= 10) return { current: "SILVER", remainHours: (50 - hours).toFixed(1) };
  return { current: "BRONZE", remainHours: (10 - hours).toFixed(1) };
}

function calculateStreak(uniqueDates) {
  let sorted = Array.from(uniqueDates).sort().reverse();
  let streak = 0, check = new Date();
  check.setHours(0, 0, 0, 0);
  for (let i = 0; i < sorted.length; i++) {
    let d = new Date(sorted[i]); d.setHours(0, 0, 0, 0);
    if ((check - d) / (1000 * 60 * 60 * 24) <= 1) { streak++; check = d; } else break;
  }
  return streak;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('★管理メニュー').addItem('QRメールの下書き作成', 'showDraftDialog').addToUi();
}

function showDraftDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('draftDialog').setWidth(400).setHeight(320),
    'QRコード送付'
  );
}

function sendCheer(targetStudentId) {
  return "エール機能は廃止されました";
}

// ====================================================================
// 🔧 デバッグ用：今週のランキング状況を出力
// ====================================================================
function debugWeeklyRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('管理シート');
  const profileSheet = ss.getSheetByName('公開プロフィール');

  const logData = logSheet.getDataRange().getValues();
  const profileData = profileSheet.getDataRange().getValues();

  let nickMap = {};
  for (let i = 1; i < profileData.length; i++) {
    let pId = String(profileData[i][IDX_PROFILE.ID]).trim();
    let pName = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || pId;
    nickMap[pId] = pName;
  }

  const thisMondayMs = getMonday(new Date()).getTime();
  console.log('今週月曜:', new Date(thisMondayMs));

  let weeklyMap = {};
  let weekRecords = [];

  for (let i = logData.length - 1; i >= 1; i--) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = logData[i][IDX_LOG.IN];
    let outTime = logData[i][IDX_LOG.OUT];

    if (!(inTime instanceof Date)) continue;
    let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
    if (logDateMs < thisMondayMs) continue;

    let diffMs = (outTime instanceof Date) ? outTime.getTime() - inTime.getTime() : 0;
    weekRecords.push({
      id: rowId,
      name: nickMap[rowId] || '不明',
      in: inTime,
      out: outTime || '(未退室)',
      hours: diffMs > 0 ? (diffMs / 3600000).toFixed(2) : 0
    });
    if (diffMs > 0) {
      weeklyMap[rowId] = (weeklyMap[rowId] || 0) + diffMs;
    }
  }

  console.log('今週のレコード件数:', weekRecords.length);
  weekRecords.forEach(r => console.log(`  ${r.name}(${r.id}): ${r.in} → ${r.out} (${r.hours}h)`));

  console.log('\n今週ランキング集計結果:');
  Object.keys(weeklyMap).forEach(id => {
    console.log(`  ${nickMap[id]}: ${(weeklyMap[id] / 3600000).toFixed(2)}h`);
  });
}

function debugProfileColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('公開プロフィール');
  if (!sheet) { console.log('シート「公開プロフィール」が見つかりません'); return; }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  console.log('=== シート上の列構成 ===');
  headers.forEach((h, i) => {
    const col = String.fromCharCode(65 + i);
    console.log(`  ${col}列(${i}): ${h}`);
  });

  console.log('\n=== コード側の認識（IDX_PROFILE） ===');
  Object.keys(IDX_PROFILE).forEach(key => {
    const idx = IDX_PROFILE[key];
    const col = String.fromCharCode(65 + idx);
    const headerVal = idx < headers.length ? headers[idx] : '(範囲外)';
    console.log(`  ${key.padEnd(15)} → ${col}列(${idx})：${headerVal}`);
  });
}

// ====================================================================
// 🔧 参考書機能用シートの初期化（手動実行用）
// ====================================================================
function initBookFeatureSheets() {
  ensureSheet(SHEET_BOOK_POSTS, BOOK_POSTS_HEADERS);
  ensureSheet(SHEET_BOOK_REACTIONS, BOOK_REACTIONS_HEADERS);
  console.log('✓ 参考書機能用シートを初期化しました');
}

// ====================================================================
// カレンダー参照
// ====================================================================
function getCalendarEventsForStudent(token, year, month) {
  const ss = SpreadsheetApp.openById(CALENDAR_CONFIG.SPREADSHEET_ID);
  const profiles = ss.getSheetByName(CALENDAR_CONFIG.SHEET_PROFILE)
    .getDataRange().getValues().slice(1);

  const student = profiles.find(row => String(row[IDX_PROFILE.TOKEN]).trim() === token);
  if (!student) return { error: 'unauthorized' };

  const calId = String(student[IDX_PROFILE.CALENDAR_ID]).trim();
  if (!calId) return { events: [] };

  const y = parseInt(year);
  const m = parseInt(month) - 1;
  const startDate = new Date(y, m, 1);
  const endDate = new Date(y, m + 1, 0, 23, 59, 59);

  let calendar;
  try { calendar = CalendarApp.getCalendarById(calId); } catch (e) {
    return { events: [] };
  }
  if (!calendar) return { events: [] };

  const events = calendar.getEvents(startDate, endDate).map(ev => ({
    title: ev.getTitle(),
    start: Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss"),
    end: Utilities.formatDate(ev.getEndTime(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss"),
    allDay: ev.isAllDayEvent(),
  }));

  return { events };
}

// ====================================================================
// カレンダー登録
// ====================================================================
function addCalendarEventsForStudent(token, events) {
  const ss = SpreadsheetApp.openById(CALENDAR_CONFIG.SPREADSHEET_ID);
  const profiles = ss.getSheetByName(CALENDAR_CONFIG.SHEET_PROFILE)
    .getDataRange().getValues().slice(1);

  const student = profiles.find(row => String(row[IDX_PROFILE.TOKEN]).trim() === token);
  if (!student) throw new Error('unauthorized');

  const calId = String(student[IDX_PROFILE.CALENDAR_ID]).trim();
  if (!calId) throw new Error('カレンダーIDが設定されていません');

  let calendar;
  try { calendar = CalendarApp.getCalendarById(calId); } catch (e) {
    throw new Error('カレンダーの取得に失敗しました');
  }
  if (!calendar) throw new Error('カレンダーが見つかりません');

  // 先に全件をチェックし、1件でも不正なら何も登録しない
  const items = validateCalendarEvents(events);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('混み合っています。少し待ってからもう一度送信してください');

  let count = 0, skipped = 0;
  try {
    // 同じ時間帯の「自習室」予定が既にあれば作らない（送信の二重押し・再送対策）
    const rangeStart = new Date(Math.min.apply(null, items.map(i => i.start.getTime())));
    const rangeEnd   = new Date(Math.max.apply(null, items.map(i => i.end.getTime())));
    const existing = new Set(calendar.getEvents(rangeStart, rangeEnd)
      .filter(ev => String(ev.getTitle()).indexOf('自習室') === 0)
      .map(ev => ev.getStartTime().getTime() + '_' + ev.getEndTime().getTime()));

    items.forEach(it => {
      const key = it.start.getTime() + '_' + it.end.getTime();
      if (existing.has(key)) { skipped++; return; }
      calendar.createEvent(it.title, it.start, it.end);
      existing.add(key);
      count++;
    });
  } finally {
    lock.releaseLock();
  }

  return skipped > 0
    ? `${count}件の予定を登録しました（同じ時間の予定${skipped}件は登録済みのためスキップ）`
    : `${count}件の予定を登録しました`;
}

// 予定登録の上限・範囲
const CAL_MAX_EVENTS_PER_REQUEST = 31;
const CAL_MAX_DAYS_AHEAD = 90;
const CAL_MAX_MEMO_LENGTH = 50;

/**
 * 生徒から送られた予定の一覧をチェックし、登録用の { title, start, end } に変換する。
 * 不正なものが1件でもあれば例外（何も登録しない）。
 */
function validateCalendarEvents(events) {
  if (!Array.isArray(events) || events.length === 0) throw new Error('予定データがありません');
  if (events.length > CAL_MAX_EVENTS_PER_REQUEST) {
    throw new Error(`一度に登録できる予定は${CAL_MAX_EVENTS_PER_REQUEST}件までです`);
  }

  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const lastDay = new Date(todayStart.getFullYear(), todayStart.getMonth(), todayStart.getDate() + CAL_MAX_DAYS_AHEAD + 1);
  const timeRe = /^([01]\d|2[0-3]):([0-5]\d)$/;
  const seen = new Set();
  const items = [];

  events.forEach((ev, idx) => {
    const n = idx + 1;
    if (!ev || typeof ev !== 'object') throw new Error(`${n}件目の予定が不正です`);
    const date = String(ev.date || '');
    const st = timeRe.exec(String(ev.start || ''));
    const en = timeRe.exec(String(ev.end || ''));
    if (!isValidDateString(date)) throw new Error(`${n}件目の日付が正しくありません`);
    if (!st || !en) throw new Error(`${n}件目の時刻が正しくありません`);

    const y = Number(date.slice(0, 4)), mo = Number(date.slice(5, 7)) - 1, d = Number(date.slice(8, 10));
    const start = new Date(y, mo, d, Number(st[1]), Number(st[2]));
    const end   = new Date(y, mo, d, Number(en[1]), Number(en[2]));
    if (end <= start) throw new Error(`${n}件目：終了時刻は開始時刻より後にしてください`);
    if (start < todayStart) throw new Error(`${n}件目：過去の日付には登録できません`);
    if (start >= lastDay) throw new Error(`${n}件目：${CAL_MAX_DAYS_AHEAD}日より先の予定は登録できません`);

    const memo = String(ev.memo || '').replace(/[\r\n\t]+/g, ' ').trim();
    if (memo.length > CAL_MAX_MEMO_LENGTH) throw new Error(`${n}件目：メモは${CAL_MAX_MEMO_LENGTH}文字以内にしてください`);

    const key = start.getTime() + '_' + end.getTime();
    if (seen.has(key)) return; // 同じ送信内の重複は1件にまとめる
    seen.add(key);

    items.push({ title: memo ? `自習室 - ${memo}` : '自習室', start: start, end: end });
  });

  return items;
}