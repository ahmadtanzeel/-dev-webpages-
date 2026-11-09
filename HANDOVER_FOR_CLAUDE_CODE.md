# SSS Education 自習室管理システム 引継ぎ資料

最終更新：2026-11-07（監査レポート v2 の対応：受付打刻の修正、自動欠席メール停止、入力チェック。12章参照）

※ 12章の変更で行番号がずれているため、1〜10章の行番号は目安として読むこと。

このドキュメントは、リポジトリ内の各ファイルを実際に読み込んで確認した「現状の仕様」をまとめたものです。
開発者が複数世代にわたっているため、意図が完全には分からない箇所もそのまま正直に記載しています。

---

## 1. システム全体像

塾（SSS Education）の自習室運営システム。3つの実行環境にまたがる。

```
[受付iPad] index.html ──POST(scan)──┐
                                     │
[生徒スマホ] dashboard.html ──GET/POST─┤→ GAS Web App (コード.js / AbsenceCheck.js)
                                     │      ├─ スプレッドシート（生徒マスタ・ログ・投稿等）
[スプレッドシートUI] draftDialog.html ┘      ├─ Googleカレンダー（生徒ごとの個別カレンダー）
   (google.script.run で直接呼び出し)         └─ Gmail / Google Chat 通知
```

- **index.html / dashboard.html** は GitHub Pages で静的ホスティングされている（`https://ahmadtanzeel.github.io/-dev-webpages-/` 配下、コード.js 内の `STUDENT_DASHBOARD_BASE_URL` 参照）。
- **コード.js / AbsenceCheck.js / draftDialog.html / appsscript.json** は同一の Google Apps Script プロジェクトにバインドされ、スプレッドシートから直接デプロイされる。GAS は同一プロジェクト内ならスクリプトファイルをまたいでもグローバルスコープが共有されるため、`コード.js` の定数・関数を `AbsenceCheck.js` からそのまま参照できる（`IDX_PROFILE`, `CALENDAR_CONFIG`, `ensureSheet` など）。
- フロントエンドとGASの通信は `fetch` ベース。POSTは CORS プリフリcourt（OPTIONS）を避けるため、JSONを送るにもかかわらず意図的に `Content-Type: text/plain;charset=utf-8` を使っている（`apiPost` / `sseSubmitAll` / index.htmlの`executeScan`）。

---

## 2. スプレッドシートのデータ構造

コード.js 側で `SpreadsheetApp.getActiveSpreadsheet()`（バインド先＝メインの管理用スプレッドシート）と、`CALENDAR_CONFIG.SPREADSHEET_ID`（`1HWZDOIJaQB0S3K4a9hXomJZFqmdznmKywQ2Sc_ON1Fo`）で明示的に開く別スプレッドシートの2系統がある。**実際には同一のスプレッドシートIDを指すよう運用されている前提**（`getActiveSpreadsheet()` を使う関数と `CALENDAR_CONFIG.SPREADSHEET_ID` を使う関数が同じ「公開プロフィール」シートを参照しているため）。

### 2.1 「公開プロフィール」シート（生徒マスタ）
`IDX_PROFILE`（コード.js 1-17行目）で一元管理：

| 列 | 意味 | 定数 |
|---|---|---|
| A (0) | 生徒ID | `ID` |
| B (1) | 生徒メールアドレス | `STUDENT_EMAIL` |
| C (2) | 保護者メールアドレス | `PARENT_EMAIL` |
| D (3) | 氏名 | `NAME` |
| E (4) | ニックネーム | `NICKNAME` |
| F (5) | 週間目標時間 | `GOAL_HOURS` |
| G (6) | URLトークン（dashboard.htmlの`?token=`と照合） | `TOKEN` |
| H (7) | 目標模試名 | `EXAM_NAME` |
| I (8) | 目標模試日 | `EXAM_DATE` |
| J (9) | （未使用・欠番） | ― |
| K (10) | 生徒個別のGoogleカレンダーID | `CALENDAR_ID` |

**これが唯一の正のマッピング。** 以前は `コード.js` と `AbsenceCheck.js` にそれぞれ別の列定義（インデックス基準もバラバラ）が重複しており、それが過去のバグ（トークン列のズレ、カレンダーID列の不一致）の直接原因だった。現在は `AbsenceCheck.js` 側は独自定義を持たず、`IDX_PROFILE` / `CALENDAR_CONFIG` を直接参照している。

### 2.2 「管理シート」（入退室ログ）
`IDX_LOG`（コード.js 40-46行目）：A=日付, B=生徒ID, D=入室時刻, E=退室時刻, F=エール数（`CHEERS`）。
- C列は使われていない（IDX_LOGにC相当の定義なし）。
- F列「エール数」は`getStudentStats`が集計して`cheers`としてAPIレスポンスに含めているが、**dashboard.html側にこの値を表示する箇所は現存しない**（`sendCheer()`関数もコード.js側で「エール機能は廃止されました」という文字列を返すだけの空実装になっている）。旧機能の残骸としてバックエンド側にだけ集計ロジックが残っている状態。

### 2.3 「生徒設定」シート
列：[生徒ID, 模試名, 模試日程]（0,1,2列、`IDX_PROFILE`とは無関係の独自レイアウト）。`getStudentStats`・`saveStudentSettings`が直接ハードコードされた列番号で読み書きする。生徒が dashboard.html の「⚙ 目標模試を自分専用に設定する」から保存すると、ここに1行追記/更新される。シートが存在しない場合は `getStudentStats` 内で自動作成される（`ss.getSheetByName('生徒設定') || ss.insertSheet('生徒設定')`）。

### 2.4 「送信済みログ」シート（欠席確認メール用・AbsenceCheck.js）
列：[日付, 送信キー, 生徒ID, 生徒名, 予定開始, 送信先]。`ensureSheet`で自動作成。送信キーは `${studentId}_${予定開始日時}` で、同日中の重複送信防止に使う。

### 2.5 「催促送信済みログ」シート（AbsenceCheck.js）
列：[送信キー, 生徒ID, 生徒名, バージョン, 対象週, 送信先, 送信日時]。Version A/B共通の1シートに両方のログが記録される。送信キーは `A_${studentId}_${週}` / `B_${studentId}_${週}` でバージョンごとに独立。

### 2.6 参考書おすすめ機能用シート
- 「参考書投稿」：[ID, 投稿日時, 生徒ID, 書名, 評価, コメント]
- 「参考書リアクション」：[投稿ID, 生徒ID, 絵文字, 日時]（同一生徒×同一投稿×同一絵文字の組み合わせはトグル式で行ごと追加/削除）

いずれも初回アクセス時に `ensureSheet` で自動作成される（`initBookFeatureSheets()` で手動初期化も可能）。

### 2.7 小テスト進捗用スプレッドシート（別ファイル）
`TEST_SS_ID`（`1uKXnpKeGCuyPpRAryFP7Ou4K4t3SgTNksZFQAhvj45E`）という**別のスプレッドシート**。生徒名（またはID）と同名のシートが生徒ごとに存在し、A列＝参考書名、B〜AD列（1〜30列目）＝チェックボックス（true/false）で各コマの完了状況を表す。生徒のシートが存在しない場合、テンプレートシート（`テンプレート`という名前で存在する前提）を `LockService` で排他制御しながら複製して新規作成する。

---

## 3. コード.js 関数リファレンス

### エントリポイント
- **`doGet(e)`**：`?token=...&action=...` で呼ばれる。`action` は `stats` / `heatmap` / `bookPosts` / `getCalendarEvents` のいずれか。tokenが無い場合は `index.html`（受付画面）をHTMLとして返す（GASにバインドされたHTML出力。ただし実運用ではGitHub Pages側の`index.html`が使われており、この分岐は事実上使われていない可能性が高い＝旧アーキテクチャの名残）。
- **`doPost(e)`**：JSON body の `action` で分岐。`saveSettings` / `scan` / `postBook` / `reactBook` / `addCalendarEvents` の5種類。

### データ集計
- **`getStudentStats(studentToken)`**：ダッシュボードのメインデータ。`CacheService`で **15分（900秒）キャッシュ**。ログ全行を毎回スキャンする力技実装（生徒数・ログ行数が今後大きく増えると重くなる可能性がある）。
  - 「入室中（未退室）」の生徒については、今週ランキング集計に限り**暫定的に現在時刻までの経過時間を加算**する。ただし未退室が12時間を超えている場合は異常値とみなして加算しない（`liveMs < 12 * 3600000` のガード）。同様の12時間キャップは `getHeatmap` にもある。
  - `calculateStreak`の連続日数判定は「直近の記録日から見て1日以内（＝前日 or 当日）なら連続とみなす」というルール。
  - 週の開始は月曜日固定（`getMonday`）。
- **`getHeatmap(token)`**：直近7日×24時間の在室分数。未退室の場合は「入室時刻+12時間」と「現在時刻」の早い方を終了時刻とみなす。
- **`calculateLevel(totalHours)`**：レベルNからN+1に必要な時間は `5 + N*2` 時間（Lv1→2は7h、Lv2→3は9h...）。
- **`calculateRank(hours)`**：BRONZE(0-10h) → SILVER(10-50h) → GOLD(50-150h) → PLATINUM(150-300h) → SSS MASTER(300h~) の5段階。

### 受付処理
- **`processScan(studentId)`**：QRスキャン（実体はID文字列の読み取り）→入退室トグル。
  - その日の最終データ行を**手動で後方走査**して特定するロジックがある（GASの`getLastRow()`は末尾に空行があると誤動作しやすいため、B列を末尾から遡って実データ行を探している）。
  - 同日中の未退室レコードを探すために、直近100行→300行→900行と**探索範囲を3倍ずつ広げながら最大3回試行**する独自のバックオフ実装（`searchSize *= 3`）。データが少ない日は1回で見つかる想定。
  - 退室時、保護者メール（`notifyParent`）とGoogle Chat通知（`notifyGoogleChat`）を送るが、**両方ともtry/catchで握りつぶしており、失敗してもスキャン自体は成功として返す**（入退室の記録を優先し、通知の失敗でユーザー体験を止めない設計）。

### メール・通知
- **`notifyParent`**：`MailApp.sendEmail`で保護者に入退室を通知（無条件・毎回送信）。
- **`notifyGoogleChat`**：`GCHAT_WEBHOOK_URL`（コード.js内に直書きされたWebhook URL＋トークン）にPOST。
- **`createDraftsByIds(idsString)`**：スプレッドシートのメニュー（`onOpen`→`showDraftDialog`→`draftDialog.html`）から呼ばれる。カンマ/読点/空白区切りのID一覧を受け取り、各生徒宛にマイページURL案内メールの**下書き**をGmailに作成する（実際の送信はしない＝人間の確認を挟む設計）。

### カレンダー連携
- **`getCalendarEventsForStudent(token, year, month)`**：`doGet`の`getCalendarEvents`アクションから呼ばれる。生徒のトークンから`CALENDAR_ID`列を引き、その生徒個人のGoogleカレンダーから指定月の予定を取得して返す。カレンダーID未設定・取得失敗時は空配列を返す（エラーにしない）。
- **`addCalendarEventsForStudent(token, events)`**：`doPost`の`addCalendarEvents`アクションから呼ばれる。dashboard.htmlの予定入力モーダルから送られた `{date, start, end, memo}[]` を、生徒個人カレンダーに `calendar.createEvent()` で一括登録する。イベントタイトルは `memo` があれば `自習室 - ${memo}`、無ければ `自習室` 固定。

### その他ユーティリティ
- **`ensureSheet(sheetName, headers)`**：シートが無ければヘッダー付きで作成し、`setFrozenRows(1)`する共通ヘルパー。コード.js / AbsenceCheck.js 双方から使われる。
- **`findStudentByToken(token)`**：トークン→`{id, name}`変換。`getHeatmap`が使用。
- **`debugWeeklyRanking()` / `debugProfileColumns()`**：手動実行専用のデバッグ用関数（本番フローからは呼ばれない）。`debugProfileColumns()`はシート上の実際のヘッダーと`IDX_PROFILE`の対応を突き合わせてログ出力するので、列がズレた時の調査に便利。

---

## 4. AbsenceCheck.js 関数リファレンス

同一GASプロジェクト内の別ファイル。**大きく2つの独立した自動送信バッチ処理**が1ファイルにまとまっている。

### 4.1 欠席確認メール（`checkAbsenceAndSendMail`）
- 想定運用：**5分ごとの時間トリガー**（`setupTrigger()`を初回に手動実行して登録）。
- 全生徒をループし、各生徒のGoogleカレンダーから「本日の予定」を取得。予定開始時刻から `CONFIG.ABSENCE_MINUTES`（5分）以上経過していて、かつ当日まだ入室記録がなく、かつ同じキーで未送信なら、生徒＋保護者宛に「来室確認できません」メールを送る。
- 重複送信防止キーは `${studentId}_${予定開始のyyyy/MM/dd HH:mm}`。同じ予定に対しては1日1回しか送らない。

### 4.2 カレンダー未入力の催促メール（`sendReminderVersionA` / `sendReminderVersionB`）
- 想定運用：**毎週月曜8:00の週次トリガー**（`setupReminderTrigger()`で登録）。
- **Version A**：来週（月〜日）の予定が1件も入っていない生徒に送る。
- **Version B**：今週 or 来週のどちらかが空の生徒に送る（Aより広い条件でヒットしやすい）。
- ⚠️ **重要な既知の問題**（コード内に注意コメントあり、AbsenceCheck.js 405-414行目）：`setupReminderTrigger()` は Version A と B の両方のトリガーを同時に登録してしまう。両方とも「毎週月曜8:00」指定で、GASの時間トリガーは分単位を指定できないためほぼ同時に実行される。**このまま本運用に入れると、条件に一致した生徒には催促メールが2通（A・B）届く。** ユーザー（現在の保守担当者）自身が「どちらを採用するか検討中」と認識しており、本運用前にどちらか一方だけを登録し直す必要がある。**これは開発者本人が把握している既知の未決事項であり、Claude側の推測で処理を変更すべきではない。**

### 4.3 共通ヘルパー
- **`buildRecipients(studentMail, parentMail)`**：生徒＋保護者メールをカンマ区切りにまとめる（保護者メールが空文字/'undefined'なら生徒のみ）。欠席確認・催促メール共通。
- **`sendMailSafely(recipients, subject, body, label, onSuccess)`**：`GmailApp.sendEmail`を try/catch で包み、成功時のみ `onSuccess`コールバック（シートへのログ記録）を呼ぶ共通ヘルパー。3箇所（欠席確認、催促A、催促B）から呼ばれるが、**各呼び出し元固有の「週選択・対象判定」ロジックはあえて統合せず、共通化したのは送信＋ログ記録の定型処理のみ**（統合し過ぎるとVersion A/Bの差分の意味が失われるため、意図的に分離を保っている）。
- **`getSentKeys` / `getReminderSentKeys`**：当日 or 全期間の送信済みキー集合を作るヘルパー。前者は「本日分のみ」、後者は「全期間分」を見る点に違いがある（催促メールは週1回粒度なので当日フィルタが不要なため）。

---

## 5. dashboard.html（生徒用マイページ）仕様

### 5.1 全体構成
1つのHTMLファイルにCSS（`<style>`）とJS（末尾の`<script>`、行3193〜4448）がすべてインラインで書かれた単一ページアプリ。フレームワーク不使用（Vanilla JS）。外部ライブラリはChart.js（CDN）とQRコード生成用の`qrcode.min.js`（同梱ファイル）のみ。

URLの `?token=xxx` パラメータ（`studentToken`、3202行目）で生徒を識別する。**このURLは生徒ごとに個別発行され、他人に共有されると誰でもその生徒のデータを閲覧・操作できてしまう**（認証はトークンの知識のみに依存する簡易方式）。

タブは4つ：`HOME` / `TRENDS` / `BOOKS` / `LOG`（`data-tab`属性で切替、`switchTab()`）。TRENDSとBOOKSタブは初回表示時に遅延ロードされる（`heatmapLoaded` / `booksLoaded`フラグ）。

### 5.2 API通信層
- `apiGet(params)`：`GET ${API_ENDPOINT}?...`
- `apiPost(payload)`：`POST ${API_ENDPOINT}`、`Content-Type: text/plain;charset=utf-8`（CORSプリフライト回避のため意図的）
- 上記2つをラップした `fetchStats` / `fetchHeatmap` / `fetchBooks` / `saveStudentSettings` / `apiPostBook` / `apiReactBook` がある。
- **カレンダー関連（`fetchEvents` / `sseSubmitAll`）だけは上記のラッパーを使わず、個別に生の`fetch`を書いている**（実装時期が別だったための不統一。動作に問題はないが、様式は揃っていない）。

### 5.3 ローカルストレージ利用一覧（すべてブラウザ単位・生徒間で共有されない）
| キー | 用途 |
|---|---|
| `sss_ach_unlocked` | 解除済み実績IDの配列（実績の「初回解除」演出を1回だけ出すためのフラグ管理） |
| `sss_missions_YYYY-M-D` | その日の3件のデイリーミッション内容＋手動チェック状態 |
| `sss_mission_all_YYYY-M-D` | その日のミッション全達成演出を出したかどうか |
| `sss_yest_percent` | 週次目標の達成率（ミッション「goal」の前回値比較に使用） |
| `sss_book_today_YYYY-M-D` | その日に参考書投稿済みかどうか（ミッション「book」判定用） |
| `sss_streak_flag` | 連続日数マイルストーン（7/30/100日）の演出を出したかどうか |
| `sss_goal_week_YYYY-M-D` | その週の目標達成演出を出したかどうか（週の月曜日をキーにする） |
| `qr_cache_${studentId}` | QRコードのdata URLキャッシュ（再生成コスト削減） |

### 5.4 実績（ACHIEVEMENTS）とミッション（MISSION_POOL）
- `ACHIEVEMENTS`（3499行目〜）：15種類の固定実績。条件はすべて`getStudentStats`のレスポンスから計算できるクライアント側判定（サーバー側に実績テーブルは無い）。
- `MISSION_POOL`（3553行目〜）：7種類のミッション定義から、日替わりで3件を**日付ベースの疑似乱数シャッフル**（`hashStr`によるシード付きソート、`localStorage`にその日の抽選結果をキャッシュ）で選ぶ。
  - ⚠️ `id: 'test'`（「小テストを1つ進める」）のミッションは `check: d => false` と**常にfalseがハードコードされており、自動達成することはない**（手動チェックのみ可能）。実装が未完のまま残っている可能性が高いが、意図の確認が取れていないため未修正。

### 5.5 予定入力モーダル（Schedule entry / `sse-*` 接頭辞）
`+ 予定を追加` ボタン（`#sse-open-btn`）→ `openModal('sse-modal')`。
- 1件ずつ「日付・開始・終了・メモ」を入力し `+ この予定を追加する`（`sseAddSchedule`）でローカルの`schedules`配列に追加。**この時点ではまだサーバーに送信されない**（複数件まとめて後で一括送信する設計）。
- 繰り返し設定：`なし` / `毎週(4週)` / `曜日指定(2週)` の3モード。「毎週」は同じ曜日で4週分、「曜日指定」は選択した曜日について2週間（14日）分を機械的に展開する。
- `Googleカレンダーに登録する`（`sseSubmitAll`）で `schedules` 配列全体を1回のPOSTで `addCalendarEvents` アクションに送信。成功したら `eventsCache` をクリアして `render()` を呼び、カレンダー表示（週/月ビュー）を即座に再取得・再描画する。

### 5.6 カレンダー表示（週/月ビュー）
- `fetchEvents(year, month)` は月単位でGASから取得し、`eventsCache`（メモリ内、ページ再読み込みで消える）にキャッシュ。
- 週表示・月表示とも、1日あたり最大2件までチップを表示し、それ以上は `.sss-cal-more` で `+N` の集約表示を出す（`renderWeek`/`renderMonth`共通の仕様。以前は月表示のみこの制限があったが、現在は週表示も同じ2件制限＋`+N`に揃えている）。

**セルの高さ（固定）**：`.sss-week-cell`は`height:140px`、`.sss-month-cell`は`height:131px`で固定。「日付＋チップ2枠（各チップがタイトル2行まで伸びた最大ケース）＋`+N`表示」が収まる最大値を手計算して割り当てている（内訳はdashboard.htmlの当該CSSコメント参照）。日によってセルの高さがバラつかないよう、意図的に`min-height`ではなく`height`固定＋`overflow:hidden`にしている。

**チップの高さ（可変・2〜3行）**：`.sss-event-chip`自体は高さ固定にしていない。中身は「時刻（`chip-time`、常に1行）」＋「タイトル（`chip-title`）」の2段構成で、`chip-title`に`-webkit-line-clamp:2`を指定し**最大2行まで折り返し表示**（省略はその2行に収まらない場合のみ）。結果としてチップ全体は「時刻1行＋タイトル1〜2行」＝2〜3行分の高さになる。タイトルを省略せず見せたいという要望と、長いタイトルでグリッドが横に崩れないようにしたいという要望（`min-width:auto`のflexboxデフォルト挙動が原因だった）の両方を満たす形。

**タイトルの改行位置制御**：予定タイトルは`addCalendarEventsForStudent`が生成する`自習室 - ${memo}`という形式が基本（2.1章・3章参照）。`chipTitleHtml(title)`関数（dashboard.html、`escapeHtml`の直後に定義）が、タイトル中の最初の`- `の直後に`<wbr>`（改行候補タグ）を挿入し、「自習室 -」と実際のメモ内容を分けて表示できるようにしている。ただし`<wbr>`は日本語（CJK）テキストでは`word-break: keep-all` + `overflow-wrap: break-word`の組み合わせが無いと無視されることがあるため（[参考記事](https://chaika.hatenablog.com/entry/2024/01/09/083000)）、`.chip-title`のCSSも`word-break: keep-all;`にしている（以前は`break-word`で任意の文字位置から折り返す設定だったが、`<wbr>`を効かせるために変更）。

**モバイル幅対応**：`@media (max-width: 480px)`で、
  - `chip-time`内の終了時刻部分（`chipTimeHtml(e)`が生成する`<span class="t-end">-終了時刻</span>`）を`display:none`にし、**開始時刻のみ**を表示。
  - `.sss-event-chip`/`.sss-cal-more`の`font-size`を9px→7pxに縮小、`letter-spacing`を0にし、`padding`も詰めて、狭い画面でも全角3文字程度（例：「自習室」）＋開始時刻が収まるようにしている。
  - セル自体の固定高さ（140px/131px）はモバイルでも変更していない（文字が小さくなる分、縦方向に余白ができるだけで崩れない）。

### 5.7 モーダルの開閉
`openModal(modalId)` / `closeAllModals()` の2関数に統一されている（`settings-modal`・`sse-modal`共通）。以前は予定入力モーダル専用の別実装（独自のoverlay要素・別関数）が並存していたが、重複のため汎用実装に統合済み。

### 5.8 イベント配線
すべて `addEventListener` ベース（`onclick`属性は現在コード全体に1つも残っていない）。静的要素は `setupUIEvents()`（3313行目）に集約。動的に`innerHTML`で再生成される要素（ミッション一覧・参考書リアクションボタン・予定入力の削除ボタン）は、親要素へのイベント委譲（`closest()`＋`data-*`属性読み取り）で処理している。

---

## 6. index.html（受付iPad用スキャン画面）仕様

シンプルな全画面フォーム。タップでオーバーレイを閉じてテキスト入力にフォーカスし、USB/Bluetoothバーコードスキャナ（キーボード入力として認識されるタイプ）でIDを読み取って**Enterキー相当の入力**で送信する想定（`keydown`の`Enter`判定）。

- 入力中に全角文字が混入したら即座に削除する正規表現フィルタがある（スキャナの誤読・IME混入対策と思われる）。
- `setInterval`で1秒ごとに、入力欄からフォーカスが外れていたら強制的に戻す（iPadでの操作中にフォーカスが外れて次のスキャンを取りこぼす事故を防ぐため）。
- APIは `doPost`の`scan`アクションを叩き、`processScan`の結果メッセージをそのまま表示する。

### API_ENDPOINTについて（2026-09-14 修正済み）
`dashboard.html`の`API_ENDPOINT`（3197行目）は最新デプロイURLに更新済み。`index.html`の`API_ENDPOINT`（212行目）も同日中に同じURLへ修正済み。両ファイルとも以下のURLで統一されている：

```
https://script.google.com/macros/s/AKfycbyXbt6myq6hiFODjR4wI9LB_1jUz-JJvsqRsmHlL5VQ4L62KuVgT3q8b9lVarwMGxf0uw/exec
```

**今後GASを再デプロイした際は、`dashboard.html`と`index.html`の両方の`API_ENDPOINT`を新URLに更新すること**（片方だけ更新すると、受付とマイページで異なるバックエンドを向いてしまう）。

---

## 7. draftDialog.html 仕様

GASエディタのスプレッドシート側メニュー（`onOpen()`→「★管理メニュー」→「QRメールの下書き作成」）から `showDraftDialog()` で開かれるモーダルダイアログ。`google.script.run`（GASのHTMLService専用API、`fetch`は使わない）で直接 `createDraftsByIds()` を呼び出す。生徒ID一覧をカンマ区切りで入力すると、各生徒宛のマイページURL案内メールをGmailに**下書きとして**作成する（誤送信防止のため、実送信は人間が下書きを確認してから行う設計）。

---

## 8. 既知の問題・要確認事項（優先度付き）

1. **【既知・運用注意】`setupReminderTrigger()`がVersion A/Bを同時登録する**（4.2章参照）。本運用開始前にどちらか一方に絞る必要あり。コード内にコメント済み。
2. **【設計メモ】`エール（cheers）`機能はフロントエンドから撤去済みだが、バックエンド（`getStudentStats`の`cheers`集計、`管理シート`F列読み取り）は残っている。** 表示先が無いだけでAPIレスポンスには今も含まれる。実害はないが、将来的に「使われていないフィールド」として整理対象になりうる。
3. **【運用メモ】`TARGET_COMMON`（共通テスト日）・`TARGET_NATIONAL`（国立二次試験日）はdashboard.html内に西暦日付でハードコーディングされている**（現在値：2027/01/16, 2027/02/25）。年度が変わるたびに手動更新が必要。
4. **【設計メモ】`MISSION_POOL`の`id:'test'`ミッションは条件が常に`false`で、自動達成することがない**（手動チェックのみ）。未完成実装の可能性があるが、意図未確認のため放置中。
5. **【設計メモ】`window._sssInitDone`フラグ**（初期データ読み込み後1.5秒で`true`になる）は、ページ読み込み直後に実績・ストリークの「解除演出」が誤って再生されるのを防ぐためのタイミング制御と推測されるが、実装者の意図は未確認。
6. **【秘匿情報】`GCHAT_WEBHOOK_URL`（Google Chat Webhook・トークン付きURL）、`CALENDAR_CONFIG.SPREADSHEET_ID`、`TEST_SS_ID`などがコード.js内に平文でハードコーディングされている。** GASの性質上一般的ではあるが、リポジトリがpublicになった場合は要注意（現状GitHubリポジトリの公開設定は本ドキュメントの範囲外につき未確認）。

---

## 9. 操作手順（半年後の自分向け・そのまま実行できるチェックリスト）

このセクションだけは「取説」として使えるよう、コードを読み返さなくても実行できる手順に絞って書く。

### 9.1 GAS Web Appを再デプロイする手順
1. GASエディタで コード.js / AbsenceCheck.js を編集・保存する。
2. 右上「デプロイ」→「デプロイを管理」→ 既存デプロイの鉛筆アイコン →「バージョン」で「新しいバージョン」を選択 →「デプロイ」。
   - 既存デプロイを「新しいバージョン」で更新する限り、Web App のURLは変わらない。URLが変わるのは「新しいデプロイ」を新規作成した場合のみ（基本的には新規作成しなくてよい）。
3. もしURLが変わった場合（新規デプロイを作った場合）は、`dashboard.html`（3197行目）と`index.html`（212行目）の**両方**の`API_ENDPOINT`を新URLに更新する（片方だけ更新すると受付とマイページが別バックエンドを向く。9.4章のトラブル表も参照）。
4. `git add` → `commit` → `push`（SSHエイリアス `github.com.arbeiten` 経由）でGitHub Pagesに反映する。
5. GitHub Pagesへの反映には数分ラグが出ることがある。反映後、実際にdashboard.html/index.htmlをブラウザで開いて動作確認する。

### 9.2 新入生を追加する手順
1. 「公開プロフィール」シートに1行追加する。
2. 各列を埋める（列の意味は2.1章の表を参照）：ID／生徒メール／保護者メール／氏名／ニックネーム／週間目標時間／URLトークン（他の生徒と重複しないランダム文字列を自分で発行して入力）／目標模試名・日（空欄でも可）／J列は空欄のまま／カレンダーID（その生徒用に用意したGoogleカレンダーのIDを控えて入力）。
3. GASエディタのスプレッドシート側メニュー「★管理メニュー」→「QRメールの下書き作成」を開き、新入生のIDを入力して実行する（`createDraftsByIds`）。→ Gmailの下書きフォルダにマイページURL案内メールが作成される。
4. 下書きの内容を確認してから手動で送信する（自動送信はしない設計）。

### 9.3 トリガーの初回登録（新しいGASプロジェクトに移行した場合など）
1. GASエディタで関数`setupTrigger`を選択して実行（▶ボタン）。→ 欠席確認メールが5分ごとに走るようになる。
2. `setupReminderTrigger`を実行する前に、催促メールをVersion A・Bどちらで運用するか決めること（8章-1参照、現時点では未決定）。決めずに実行すると両方登録されて二重送信になる。実行してしまった場合は、GASエディタ左メニュー「トリガー」から片方を手動削除する。

### 9.4 よくある不具合の切り分け表
| 症状 | まず確認する箇所 |
|---|---|
| 生徒のマイページで「URLが正しくありません」と出る | URLの`?token=`の値と「公開プロフィール」G列が一致しているか |
| 生徒のマイページで「無効なURLです。管理者に連絡してください」と出る | 同上（サーバー側`getStudentStats`の`targetId`未検出エラー） |
| 受付iPadでスキャンしてもエラー・無反応 | `dashboard.html`と`index.html`の`API_ENDPOINT`が同じURLか（9.1章）、GAS側のデプロイが有効か |
| カレンダーの予定が生徒のページに反映されない | 「公開プロフィール」K列（カレンダーID）が正しいか、そのGoogleカレンダーへGASのアカウントがアクセスできるか |
| 催促・欠席確認メールが二重に届く | `setupReminderTrigger()`でVersion A/B両方のトリガーが有効になっていないか（GASエディタ「トリガー」画面で確認、8章-1参照） |
| データの列がズレて表示される・変な値が出る | GASエディタで`debugProfileColumns()`を手動実行し、ログでシートの実際の列と`IDX_PROFILE`の対応を突き合わせる |

### 9.5 年次メンテナンス
- `TARGET_COMMON` / `TARGET_NATIONAL`（dashboard.html 3198-3199行目）：共通テスト・国立二次試験の目標日。年度が変わったら西暦日付を手動で書き換える。

### 9.6 その他の運用情報
- **GitHub Pages**：`dashboard.html` / `index.html` を含むこのリポジトリを静的ホスティング。SSH経由でのpush用に、`~/.ssh/config`に `github.com.arbeiten` というホストエイリアスと専用鍵を設定済み（他アカウント用のデフォルト`github.com`設定と分離）。手順は9.1章参照。

### 9.7 コミット前にローカルで動作確認する方法

**フロントエンド（dashboard.html / index.html）**：commit/push不要。`API_ENDPOINT`は本番デプロイのGAS URLを直接叩く作りなので、ローカルでファイルをそのままブラウザで開く（または`python -m http.server`等の簡易サーバーで配信する）だけで、GitHub Pagesに反映する前に動作確認できる。ただし実際に通信するバックエンド（スプレッドシート・カレンダー・Gmail）は現状テスト環境用になっているため、本番データへの影響は無い前提で運用している。

**バックエンド（コード.js / AbsenceCheck.js）**：GAS用の公式CLI「**clasp**」を使うと、GASエディタに毎回貼り直さずローカル編集→反映ができる。このリポジトリでは既に`clasp clone`済みで、`.clasp.json`（スクリプトIDなど）と`.claspignore`（後述）がリポジトリ直下にある。

1. 事前準備：Node.js/npmが使えること、対象GASプロジェクトにアクセス権のあるGoogleアカウントを持っていること。
2. インストール：`npm install -g @google/clasp`
3. ログイン：`clasp login`（ブラウザが開くので認可する）
4. 初回のみ、https://script.google.com/home/usersettings で「Google Apps Script API」をONにする。
5. ローカルで編集する（`コード.js` / `AbsenceCheck.js` / `appsscript.json` / `draftDialog.html`）。
6. GASプロジェクトへ反映：`clasp push`（確認が出たら`y`）。**ローカルの内容でGAS側を上書きするため、GASエディタ側で直接編集していた変更があると消える**。以後は「ローカル編集→push」に一本化するのが安全。
   - `.claspignore`により、`qrcode.min.js`（サードパーティのブラウザ用ライブラリ）や`README.md`/`HANDOVER_FOR_CLAUDE_CODE.md`（ドキュメント類）はpush対象から除外している。**`dashboard.html`/`index.html`は意図的に除外しておらず、`clasp push`でGASプロジェクト側にも反映される**（コード.jsの`doGet`が`HtmlService.createHtmlOutputFromFile('index')`を参照する箇所があるため、GASプロジェクト内の同名ファイルもこの内容で更新されることになる）。
7. ブラウザを開かずに関数を直接実行して確認：`clasp run 関数名`（引数が必要なら`-p '["引数1","引数2"]'`でJSON配列を渡す）。`console.log`の出力もターミナルに表示される。
8. GASエディタで見た目を確認したい時：`clasp open`。
9. Web Appとして本番同様に確認したい場合、Web Appの「デプロイ」自体は従来通りGASエディタで「新しいバージョン」を作成する（`clasp deploy`でCLIから行うことも可能だが、まずはエディタでの操作で十分）。

---

## 10. 次回以降の積み残しタスク

未着手のまま次回に持ち越すタスク（2026-09-16時点）。

- ~~カレンダーの各セルの高さ固定~~ → 2026-09-16に対応完了（5.6章参照）。

1. **通知機能の実装（要確認）**：着手前に、既存で動いている通知系（保護者への入退室メール`notifyParent`、Google Chat通知`notifyGoogleChat`、欠席確認メール`checkAbsenceAndSendMail`、カレンダー未入力の催促メール`sendReminderVersionA/B`）と何が違うのか・何を新規に作るのかを整理してから着手する（3章・4章参照）。既存の通知経路と重複/競合しないか要確認。

---

## 11. このドキュメントの作成方針について

このリポジトリは複数の開発者が関わってきており、現在の保守担当者（このドキュメントの依頼者）は「自分が書いていないコードの意図を推測でコメント化することはしたくない」という方針を明確にしている。そのためコード本体へのコメント追加は必要最小限（`AbsenceCheck.js`の`setupReminderTrigger()`のみ、本人が意図を把握している箇所）に留めており、本ドキュメントはコードを変更せずに**外部から見た仕様・挙動・既知の問題を客観的に記述する**形でその方針を補っている。今後このリポジトリを引き継ぐ人（人間・AIエージェント問わず）は、まずこのドキュメントで全体像を掴んだ上で、個々のロジックの「なぜそうなっているか」を変更する前には保守担当者に確認することを推奨する。

---

## 12. 2026-11-07 監査レポート v2 対応で変わったこと

### 12.1 変更内容
| 対象 | 変更 |
|---|---|
| 受付打刻（`processScan`） | 受付キー確認 → 名簿確認 → ロックして記録 → 確定後に通知、の順に変更。名簿にないID・空IDは**シートを書き換える前に** `ok:false` で断る |
| 二重打刻 | 同じ生徒の直前の打刻（入室・退室の新しい方）から**5秒未満**は記録も通知もしない（`SCAN_DUPLICATE_GUARD_MS`）。別の生徒は待たない |
| 再送 | 受付画面が打刻ごとに `requestId` を付ける。同じ `requestId` は10分間、前回の結果を返すだけ |
| 受付画面（index.html） | 最初に「受付キー」を入力し、端末に保存。結果ごとに「登録されていません」「打刻済み」などを表示 |
| 自動欠席メール | **廃止**。`checkAbsenceAndSendMail` は先頭で止まる（`ABSENCE_MAIL_ENABLED = false`）。`setupTrigger` はトリガーを作らず削除だけ行う |
| 催促メール A/B | `REMINDER_VERSION` に A か B を指定したときだけ、その片方を登録。未指定なら登録しない |
| 式の混入対策 | 参考書投稿・模試名は、先頭が `= + - @` なら文字列として保存（`safeCellText`） |
| 設定保存 | 無効なURLは `ok:false` で返す（以前は「認証エラー」を成功扱いで返していた）。模試名50文字・日付の形式と年をチェック |
| 予定登録 | 1回31件まで、今日〜90日先、時刻の形式、終了＞開始、メモ50文字をチェック。1件でも不正なら全件登録しない。同じ時間の「自習室」予定は作らない |
| マイページ | 予定一覧のメモを、画面に出す前にエスケープ |
| テスト | `npm test` で GAS を真似たモック上のテストが動く（`tests/`、GASには送らない） |

### 12.2 本番に反映するときの手順（この順で）
1. **受付キーを決めて登録**：GASエディタ →「プロジェクトの設定」→「スクリプト プロパティ」に `KIOSK_KEY` を追加。値は推測されにくい長い文字列（例：パスワード生成ツールで作った32文字以上）。**チャット・メール・資料には貼らない。**
2. コードを GAS に反映（`clasp push` か GitHub の main への push）→ 9.1 の手順で**既存デプロイを新しいバージョンで更新**。※ 受付キー未登録のまま更新すると、打刻が全て「設定エラー」になる。
3. GitHub Pages の index.html / dashboard.html を更新。
4. 受付iPadで受付画面を開き直し、「TAP TO START」→ 受付キーを入力。テスト用IDで入室・退室・5秒以内の再スキャンを確認。
5. **欠席メールのトリガー削除**：GASエディタで `removeAbsenceTrigger` を実行。過去にトリガーを設定したことのある**各Googleアカウントで**実行し、`listTriggers` で `checkAbsenceAndSendMail` が無いことを確認。「送信済みログ」シートは消さない。
6. 催促メールを続ける場合は `REMINDER_VERSION` を `'A'` か `'B'` にして `setupReminderTrigger` を実行。

### 12.3 まだ対応していないもの（監査レポート v2 の番号）
- S-01：Web App の公開設定（`ANYONE_ANONYMOUS`）そのもの、管理者用処理の権限分け
- S-08：操作回数の上限・監査ログ
- 「退塾済み」の扱い：名簿に退塾の列がないため、退塾者は「公開プロフィール」から行を外す（またはIDを変える）運用が必要
- P2（R-05〜R-14、S-09）、D-01〜D-07
