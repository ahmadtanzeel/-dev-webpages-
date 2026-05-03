// ====================================================================
// 【列構成 最新版】インデックス定義（個人情報・公開プロフィール共通）
// A列:ID, B列:氏名, C列:ニックネーム, D列:目標時間, E列:URL, F列:トークン, G列:模試名, H列:模試日
// ====================================================================
// ====================================================================
// 【列構成 最新版】インデックス定義（個人情報・公開プロフィール共通）
// A列:ID, B列:氏名, C列:ニックネーム, D列:目標時間, E列:URLトークン, F列:模試名, G列:模試日
// ====================================================================
const IDX_PROFILE = {
  ID: 0,          // A列: 学習者ID
  NAME: 1,        // B列: 学習者氏名（実名）
  NICKNAME: 2,    // C列: ニックネーム
  GOAL_HOURS: 3,  // D列: 目標勉強時間
  TOKEN: 4,       // E列: URLトークン ★ここを4に修正！
  EXAM_NAME: 5,   // F列: 目標模試名  ★上に詰める
  EXAM_DATE: 6    // G列: 模試日付    ★上に詰める
};
const IDX_PERSONAL = IDX_PROFILE;

// ====================================================================
// ① アクセス振り分け処理（Webアプリの入り口・必須！）
// ====================================================================
// ====================================================================
// ① アクセス振り分け処理（API化対応版）
// ====================================================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token  = params.token;

  // --- API: 学習統計データ取得 ---
  if (action === 'stats' && token) {
    try {
      const data = getStudentStats(token);
      return jsonResponse({ ok: true, data: data });
    } catch (err) {
      return jsonResponse({ ok: false, error: String(err && err.message || err) });
    }
  }

  // --- 受付画面（index.html、トークンなし）：従来通り ---
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SSS Education 自習室受付')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ====================================================================
// ①-2 POSTリクエスト受け口（設定保存などの書き込み処理）
// ====================================================================
function doPost(e) {
  try {
    // フロントは Content-Type: text/plain でJSON文字列を送ってくる（CORSプリフライト回避のため）
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;

    if (action === 'saveSettings') {
      const msg = saveStudentSettings(body.token, body.examName, body.examDate);
      return jsonResponse({ ok: true, message: msg });
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }
}

// ====================================================================
// ①-3 JSONレスポンス共通ヘルパー
// ====================================================================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
// ====================================================================
// ② ダッシュボード用データ集計（実名対応・最適化版）
// ====================================================================
function getStudentStats(studentToken) {
  // ★ キャッシュ確認（あれば爆速表示）
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

  // プロフィールから情報を抽出（実名を取得）
  for (let i = 1; i < profileData.length; i++) {
    let pId = String(profileData[i][IDX_PROFILE.ID]).trim();
    let pToken = String(profileData[i][IDX_PROFILE.TOKEN]).trim();
    
    // ★ B列の実名を取得（空欄ならC列のニックネーム、それもなければIDを表示）
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

  // 生徒設定（模試など）の取得
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
  const getMonday = (d) => {
    let day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff);
  };
  const thisMondayMs = getMonday(new Date()).getTime();

  let totalMs = 0, todayMs = 0, weekMs = 0, todayCheers = 0;
  let dailyMap = {}, weeklyMap = {}, monthlyMap = {}, weeklyRankingMap = {};
  let recentActions = [], historyList = [];
  let uniqueDates = new Set();

  for(let i = logData.length - 1; i >= 1; i--){
    let rowId = String(logData[i][1]).trim();
    let inTime = toDate(logData[i][3]);  
    let outTime = toDate(logData[i][4]); 

    if(inTime) {
      let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
      let diffMs = (outTime) ? outTime.getTime() - inTime.getTime() : 0;
      if (diffMs < 0) diffMs = 0;

      if(rowId === targetId) {
        if(diffMs > 0) {
          totalMs += diffMs;
          if (logDateMs === todayStart) todayMs += diffMs;
          if (logDateMs >= (todayStart - 7*24*60*60*1000)) weekMs += diffMs;
          
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
          todayCheers += parseInt(logData[i][5] || 0); // F列(6列目)の応援数を取得
        }
      }
      
      if(logDateMs >= thisMondayMs && diffMs > 0) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + diffMs;
      }
      
      if(recentActions.length < 5) {
        recentActions.push({
          id: rowId,
          name: nickMap[rowId] || "学習者", // ★実名が表示される
          action: (outTime) ? "退室" : "入室",
          time: Utilities.formatDate((outTime ? outTime : inTime), "JST", "HH:mm")
        });
      }
    }
  }

  let top5Ranking = Object.keys(weeklyRankingMap).map(id => {
    return { name: nickMap[id] || "学習者", hours: (weeklyRankingMap[id] / 3600000).toFixed(1) };
  }).sort((a, b) => b.hours - a.hours).slice(0, 5);

  const TEST_SS_ID = '1uKXnpKeGCuyPpRAryFP7Ou4K4t3SgTNksZFQAhvj45E'; 
  let testProgress = [];
  try {
    const testSs = SpreadsheetApp.openById(TEST_SS_ID);
    const testSheet = testSs.getSheetByName(targetId);
    if (testSheet) {
      const testData = testSheet.getDataRange().getValues();
      for (let i = 1; i < testData.length; i++) {
        let bookName = testData[i][0];
        if (!bookName) continue;
        let totalCount = 0, completedCount = 0, blocks = [];
        for (let j = 1; j <= 30; j++) {
          let cellValue = testData[i][j];
          if (cellValue === true || cellValue === false) {
            totalCount++;
            if (cellValue === true) { completedCount++; blocks.push(true); }
            else { blocks.push(false); }
          }
        }
        if (totalCount > 0) testProgress.push({ name: bookName, total: totalCount, completed: completedCount, blocks: blocks });
      }
    }
  } catch (e) { console.error("小テスト取得失敗:", e.message); }

  const formatChartData = (map) => {
    let labels = Object.keys(map).sort();
    let values = labels.map(k => (map[k]/3600000).toFixed(1));
    return { labels, values };
  };

  const resultData = {
    name: targetNickname,
    id: targetId, 
    today: formatTime(todayMs),
    week: formatTime(weekMs),
    total: formatTime(totalMs),
    daily: formatChartData(dailyMap),
    weekly: formatChartData(weeklyMap),
    monthly: formatChartData(monthlyMap),
    rank: calculateRank(totalMs / 3600000),
    community: { ranking: top5Ranking, feed: recentActions },
    weeklyGoal: { currentHours: (weekMs / 3600000).toFixed(1), targetHours: weeklyGoalHours, percent: Math.min(100, ((weekMs/3600000)/weeklyGoalHours)*100).toFixed(1) },
    history: historyList.slice(0, 30),
    streak: { totalDays: uniqueDates.size, active: calculateStreak(uniqueDates) },
    tests: testProgress,
    exam: customExam,
    cheers: todayCheers
    
  };
  

  cache.put(studentToken, JSON.stringify(resultData), 900); // 15分間キャッシュ
  return resultData;
}

// ====================================================================
// ③ エール（応援）送信処理
// ====================================================================
function sendCheer(targetStudentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("管理シート");
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");

  for (let i = data.length - 1; i >= 1; i--) {
    const rowDateStr = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : "";
    if (String(data[i][1]).trim() === targetStudentId && rowDateStr === todayStr) {
      const currentCheer = parseInt(data[i][5] || 0); // F列
      sheet.getRange(i + 1, 6).setValue(currentCheer + 1);
      return "応援を届けました！";
    }
  }
  return "現在、応援できる記録が見つかりませんでした。";
}

// ====================================================================
// ④ 生徒個別の模試・目標設定保存
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
  if (!studentId) return "認証エラー";

  const settingsSheet = ss.getSheetByName("生徒設定");
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
  settingsSheet.getRange(targetRow, 2).setValue(examName);
  settingsSheet.getRange(targetRow, 3).setValue(examDate);
  
  CacheService.getScriptCache().remove(token); // 設定を変えたらキャッシュを破棄
  
  return "設定を保存しました！";
}

// ====================================================================
// ⑤ 受付打刻処理（キャッシュ自動クリア機能付き）
// ====================================================================
function processScan(studentId) {
  if (!studentId) return "エラー：IDが読み込めませんでした";
  const cleanTargetId = String(studentId).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("管理シート");
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");
  
  const bValues = sheet.getRange("B:B").getValues();
  const data = sheet.getDataRange().getValues();
  let existingRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const rowDateStr = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : "";
    if (String(data[i][1]).trim() === cleanTargetId && rowDateStr === todayStr && (data[i][4] === "" || data[i][4] === null)) {
      existingRowIndex = i + 1; break;
    }
  }

  let actionType = "";
  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, 5).setValue(now);
    actionType = "退室";
  } else {
    let targetRow = Math.max(2, sheet.getLastRow() + 1);
    for (let j = 1; j < bValues.length; j++) {
      if (bValues[j][0] === "" || bValues[j][0] === null) { targetRow = j + 1; break; }
    }
    sheet.getRange(targetRow, 1).setValue(now);
    sheet.getRange(targetRow, 2).setValue(cleanTargetId);
    sheet.getRange(targetRow, 4).setValue(now);
    actionType = "入室";
  }
  
  const pData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  let studentName = "学習者";
  let studentToken = null; 

  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][IDX_PROFILE.ID]).trim() === cleanTargetId) { 
      studentName = pData[i][IDX_PROFILE.NAME] || pData[i][IDX_PROFILE.NICKNAME] || "学習者"; 
      studentToken = String(pData[i][IDX_PROFILE.TOKEN]).trim();
      break; 
    }
  }

  if (studentToken) CacheService.getScriptCache().remove(studentToken);

  return `${studentName} さんが ${actionType} しました！`;
}

// ====================================================================
// 補助関数群
// ====================================================================
function formatTime(ms) { return `${Math.floor(ms/3600000)}時間${Math.floor((ms%3600000)/60000)}分`; }
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
  check.setHours(0,0,0,0);
  for(let i=0; i<sorted.length; i++) {
    let d = new Date(sorted[i]); d.setHours(0,0,0,0);
    if((check - d) / (1000*60*60*24) <= 1) { streak++; check = d; } else break;
  }
  return streak;
}
function onOpen() { SpreadsheetApp.getUi().createMenu('★管理メニュー').addItem('QRメールの下書き作成', 'showDraftDialog').addToUi(); }
function showDraftDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('draftDialog').setWidth(400).setHeight(320), 'QRコード送付'); }