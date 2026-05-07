// ====================================================================
// 【列構成 最新版】インデックス定義（公開プロフィール / 個人情報マスタ共通）
// A:ID / B:氏名 / C:ニックネーム / D:生徒メアド / E:保護者メアド
// F:目標時間 / G:URLトークン / H:模試名 / I:模試日
// ====================================================================
const IDX_PROFILE = {
  ID: 0,             // A列
  NAME: 1,           // B列
  NICKNAME: 2,       // C列
  STUDENT_EMAIL: 3,  // D列 ★新規
  PARENT_EMAIL: 4,   // E列 ★新規
  GOAL_HOURS: 5,     // F列（旧D列）
  TOKEN: 6,          // G列（旧E列）
  EXAM_NAME: 7,      // H列（旧F列）
  EXAM_DATE: 8       // I列（旧G列）
};
const IDX_PERSONAL = IDX_PROFILE;

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
// ① アクセス振り分け処理
// ====================================================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token  = params.token;

  if (action === 'stats' && token) {
    try {
      const data = getStudentStats(token);
      return jsonResponse({ ok: true, data: data });
    } catch (err) {
      return jsonResponse({ ok: false, error: String(err && err.message || err) });
    }
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
      const msg = processScan(body.id);
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
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = toDate(logData[i][IDX_LOG.IN]);
    let outTime = toDate(logData[i][IDX_LOG.OUT]);

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
          todayCheers += parseInt(logData[i][IDX_LOG.CHEERS] || 0);
        }
      }

      // 今週ランキング集計（全員、退室済みのみ）
      if(logDateMs >= thisMondayMs && diffMs > 0) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + diffMs;
      }

      if(recentActions.length < 5) {
        recentActions.push({
          id: rowId,
          name: nickMap[rowId] || "学習者",
          action: (outTime) ? "退室" : "入室",
          time: Utilities.formatDate((outTime ? outTime : inTime), "JST", "HH:mm")
        });
      }
    }
  }

  // ★ ランキング修正：未退室（diffMs=0）も「現在進行中の入室」として含めて、データがある人を全員カウント
  // 上記のままだと「入室中だがまだ退室していない人」がランキングに入らないため
  // 改善版：ranking用のmapに、退室時刻がない場合は入室時刻から現在時刻までを暫定加算
  for(let i = logData.length - 1; i >= 1; i--){
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = toDate(logData[i][IDX_LOG.IN]);
    let outTime = toDate(logData[i][IDX_LOG.OUT]);
    if (!inTime) continue;
    let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
    if (logDateMs < thisMondayMs) break;  // 今週分のみで打ち切り（高速化）

    if (!outTime) {
      // 入室中：暫定で現在時刻までを加算
      let liveMs = now.getTime() - inTime.getTime();
      if (liveMs > 0 && liveMs < 12 * 3600000) {  // 12時間超は異常値とみなしスキップ
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + liveMs;
      }
    }
  }

  let top5Ranking = Object.keys(weeklyRankingMap).map(id => {
    return { name: nickMap[id] || "学習者", hours: (weeklyRankingMap[id] / 3600000).toFixed(1) };
  }).filter(r => parseFloat(r.hours) > 0)
    .sort((a, b) => parseFloat(b.hours) - parseFloat(a.hours))
    .slice(0, 5);

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

  cache.put(studentToken, JSON.stringify(resultData), 900);
  return resultData;
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

  CacheService.getScriptCache().remove(token);

  return "設定を保存しました！";
}

// ====================================================================
// ④ 受付打刻処理（★ 保護者メール通知機能つき ★）
// ====================================================================
function processScan(studentId) {
  if (!studentId) return "エラー：IDが読み込めませんでした";
  const cleanTargetId = String(studentId).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("管理シート");
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return "エラー：シートが空です";

  // ★ 直近100件だけを後ろから取得（typically 1日以内の打刻）
  // 100件で見つからなければ300件、900件と段階的に拡張する保険つき
  let existingRowIndex = -1;
  let existingInTime = null;  // 退室時にメールへ含める「本日の学習時間」計算用
  let searchSize = Math.min(100, lastRow - 1);
  let attempt = 0;
  let maxAttempts = 3; // 100, 300, 900 件まで遡る

  while (existingRowIndex === -1 && attempt < maxAttempts && searchSize > 0) {
    let startRow = Math.max(2, lastRow - searchSize + 1);
    let numRows = lastRow - startRow + 1;
    if (numRows <= 0) break;

    // A〜E列のみ取得（F列のエール数は不要）
    const data = sheet.getRange(startRow, 1, numRows, 5).getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      const rowDate = data[i][IDX_LOG.DATE];
      const rowDateStr = rowDate instanceof Date ? Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") : "";
      if (rowDateStr !== todayStr) continue;
      if (String(data[i][IDX_LOG.ID]).trim() !== cleanTargetId) continue;
      if (data[i][IDX_LOG.OUT] === "" || data[i][IDX_LOG.OUT] === null) {
        existingRowIndex = startRow + i; // 実際の行番号
        existingInTime = data[i][IDX_LOG.IN]; // 入室時刻を保持
        break;
      }
    }

    attempt++;
    searchSize *= 3; // 100 → 300 → 900件
  }

  let actionType = "";
  let studyMs = 0;
  if (existingRowIndex > 0) {
    // 退室処理
    sheet.getRange(existingRowIndex, IDX_LOG.OUT + 1).setValue(now);
    actionType = "退室";
    if (existingInTime instanceof Date) {
      studyMs = now.getTime() - existingInTime.getTime();
      if (studyMs < 0) studyMs = 0;
    }
  } else {
    // 入室処理（最終行の次に追記）
    const targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1, 1, 4).setValues([[now, cleanTargetId, "", now]]);
    actionType = "入室";
  }

  // プロフィール検索：見つかったら即break
  const pData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  let studentName = "学習者";
  let studentToken = null;
  let parentEmail = null;

  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][IDX_PROFILE.ID]).trim() === cleanTargetId) {
      studentName = pData[i][IDX_PROFILE.NAME] || pData[i][IDX_PROFILE.NICKNAME] || "学習者";
      studentToken = String(pData[i][IDX_PROFILE.TOKEN]).trim();

      // 保護者メアドを取得（@マークを含むかで簡易バリデーション）
      const rawParentEmail = pData[i][IDX_PROFILE.PARENT_EMAIL];
      if (rawParentEmail && String(rawParentEmail).indexOf('@') > 0) {
        parentEmail = String(rawParentEmail).trim();
      }
      break;
    }
  }

  // 🔔 保護者にメール通知（E列が入力されていれば送信）
  if (parentEmail) {
    notifyParent(studentName, parentEmail, actionType, now, studyMs);
  }

  if (studentToken) CacheService.getScriptCache().remove(studentToken);

  return `${studentName} さんが ${actionType} しました！`;
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

  // 退室時のみ「本日の学習時間」を追記
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
// 補助関数群
// ====================================================================
function formatTime(ms) { return `${Math.floor(ms/3600000)}時間${Math.floor((ms%3600000)/60000)}分`; }

function calculateRank(hours) {
  if (hours >= 300) return { current: "SSS MASTER", remainHours: 0 };
  if (hours >= 150) return { current: "PLATINUM", remainHours: (300 - hours).toFixed(1) };
  if (hours >= 50)  return { current: "GOLD", remainHours: (150 - hours).toFixed(1) };
  if (hours >= 10)  return { current: "SILVER", remainHours: (50 - hours).toFixed(1) };
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

function onOpen() {
  SpreadsheetApp.getUi().createMenu('★管理メニュー').addItem('QRメールの下書き作成', 'showDraftDialog').addToUi();
}

function showDraftDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('draftDialog').setWidth(400).setHeight(320),
    'QRコード送付'
  );
}

// ====================================================================
// ⑤ 旧エール送信（廃止済み・互換用に残置）
// ====================================================================
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

  const getMonday = (d) => {
    let day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff);
  };
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
  console.log('今週のレコード詳細:');
  weekRecords.forEach(r => console.log(`  ${r.name}(${r.id}): ${r.in} → ${r.out} (${r.hours}h)`));

  console.log('\n今週ランキング集計結果:');
  Object.keys(weeklyMap).forEach(id => {
    console.log(`  ${nickMap[id]}: ${(weeklyMap[id]/3600000).toFixed(2)}h`);
  });

  if (weekRecords.length === 0) {
    console.log('⚠ 今週のレコードがありません。日付の判定がズレている可能性があります。');
  }
  if (Object.keys(weeklyMap).length === 0) {
    console.log('⚠ ランキングが空です。退室済みの記録が今週分にない可能性があります。');
  }
}

// ====================================================================
// 🔧 デバッグ用：processScanの速度テスト
// ====================================================================
function testProcessScanSpeed() {
  const start = new Date().getTime();
  const result = processScan('テスト用の実在する学習者ID');  // ここを実IDに書き換えて実行
  const elapsed = new Date().getTime() - start;
  console.log(`処理時間: ${elapsed}ms / 結果: ${result}`);
}

// ====================================================================
// 🔧 デバッグ用：「公開プロフィール」シートの列構成を確認
//    シート列とコード側のIDX_PROFILE定数が一致しているか確認用
// ====================================================================
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
// 🔧 デバッグ用：保護者メール送信テスト
//    test@example.com の部分を自分のメアドに変更して実行
// ====================================================================
function testEmailSend() {
  const testEmail = 'test@example.com';  // ★自分のメアドに変更
  notifyParent('テスト 太郎', testEmail, '退室', new Date(), 3 * 3600000 + 25 * 60000);
  console.log(`テストメールを ${testEmail} に送信しました`);
}