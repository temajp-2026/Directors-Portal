# CLAUDE.md — TeMA 理事ポータル

NPO法人 日本繊維商品めんてなんす研究会（TeMA）の理事向け業務システム。
このファイルは Claude Code が作業するときの前提・ルールをまとめたもの。

## 応答・ドキュメントのルール
- ユーザーとのやり取り、画面の文言、コメント、ドキュメントはすべて日本語。
- 利用者（理事・講師）はITに詳しくない人が多い。画面の文言は平易に、専門用語を避ける。
- 変更したら、利用者向けの `docs/理事ポータル_機能説明書_*.md`（NotebookLMのヘルプデスクの元資料）も更新する。
  内部の仕組み（シートの列、API名など）は機能説明書に書かない（NotebookLMは公開設定のため）。

## 構成
| 層 | 内容 |
|---|---|
| 画面 | このリポジトリ直下の HTML/JS（GitHub Pages：https://temajp-2026.github.io/Directors-Portal/ ） |
| サーバー | `gas/Code.gs`（Google Apps Script のウェブアプリ。`doPost` のみ） |
| データ | GASに紐づいたスプレッドシート「理事ポータル管理DB」 |
| ファイル | Googleドライブ（理事ポータル_講座PDF／報告書添付／講座資料庫） |

### 画面ファイル
| ファイル | 画面 | ログイン |
|---|---|---|
| index.html | トップ（ダッシュボード・メニュー・講座一覧・講座進捗ページ） | 要 |
| course-progress.html | 講座準備進捗確認（8段階） | 要 |
| invoice-ledger.html | 請求書管理簿（※ `請求書管理簿.html` は同内容の旧名。**必ず同じ内容にそろえる**） | 要＋請求書アクセス |
| invoice-lecturer.html | 講師用請求書（※ `請求書_講師用.html` も同内容。**必ずそろえる**） | 不要 |
| reports.html | 月次報告書（ベータ版） | 要 |
| materials.html / materials-lecturer.html | 資料庫（理事用／講師用） | 要／不要 |
| sns.html | SNS文章作成＆履歴（ベータ版） | 要 |
| prep-phases.js | 8段階の判定（index と course-progress で共用） | － |
| portal-nav.js | 画面切り替えバーと講座ごとのリンク（ログインが必要な6画面で共用） | － |
| 講師依頼_請書兼請求書.html | 旧試作。未使用 | － |

## 共通の実装ルール（画面）
- 各HTMLは1ファイル完結（CSS・JSを内包）。外部ライブラリは極力使わない（pdf.js はSRI付きでjsDelivrから遅延読み込み）。
- API呼び出しは `callApi(action, payload)`：`POST` / `Content-Type: text/plain`（CORSプリフライト回避）、`{action, idToken, payload}`。
  `list`/`get` 始まりは通信失敗時に1回だけ自動再試行。`unauthorized` でログイン画面へ戻す。
- ログイン情報：`localStorage['riji_portal_auth']`（`{idToken, user, savedAt}`、6時間で失効。全タブ共有）。
- 文字サイズ：CSS変数 `--fs-scale`、`localStorage['temaFontSizePref']`。フォントサイズは `calc(var(--fs-scale, 1) * Npx)` で書く。
- 配色：CSS変数 `--gw-*`（Google Workspace 風）。メニューの色：請求書=青 / 準備進捗=緑 / 月次報告書=オレンジ / 資料庫=紫 / SNS=ピンク。
- 利用者が入力した文字は必ず `escapeHtml` してから表示する（innerHTMLに生で入れない）。URLは `safeUrl` を通す。
- スマホ表示必須（表はスマホでカード表示 or 横スクロール）。印刷が絡む画面は `@media print` も確認する。
- 他画面への講座指定リンク：`index.html?openCourse=ID`、`course-progress.html?course=ID`、`invoice-ledger.html?course=ID`（`?open=請求書ID` `?newForCourse=講座ID` も可）、`materials.html?course=ID`、`sns.html?course=ID`。

## 共通の実装ルール（GAS：gas/Code.gs）
- 変更内容は、ファイル先頭のコメント「▼ このコードを更新して反映するときは」に日付付きで追記する（管理者が読む更新手順になっている）。
- シートの列は `*_HEADERS` 配列で定義。列を追加したら `initializeSpreadsheet` で移行される（既存データを壊さない）。列の追加がある更新は「initializeSpreadsheet の再実行が必要」と明記する。
- シートへの書き込みは必ず `safeCell_` / `safeRow_` / `safeRows_` を通す（数式インジェクション対策）。URL欄は `safeHttpUrl_`。
- 権限：`PUBLIC_ACTIONS`（講師用・未ログイン）、`INVOICE_ACTIONS`（請求書アクセス必須）、それ以外は名簿の理事。新しいアクションを追加したら、どれに属するか必ず決める。
- 書き込み系は `LockService` で直列化（時間のかかる処理は `LOCK_EXEMPT`）。
- 講座一覧は `readCourseRowsForList_` で必要な列だけ読む（`pdfOcrText` やサムネイルの大きなセルを読まない）。行検索はID列だけ読む。全列 `getDataRange()` は避ける。
- 5秒以上の処理は ActivityLog に `slow_request` として記録される。ログイン拒否は `login_denied` / `login_token_rejected`。
- **秘密情報（AIのAPIキー等）はコードに書かない。** スクリプト プロパティ（`PropertiesService.getScriptProperties()`）に保存する。このリポジトリは公開されている前提で扱う。

## 主な業務ルール
- 8段階：企画決定→講師関係→募集準備→募集開始→申込管理→開催準備→開催→終了処理（目安日数は Settings シート）。
- 「不要」設定（Courses.phaseJson）：naForm / naLecturerPdf / naDirectorMaterials / naCoursePdf / naLecturer（講師なし＝講師・請書・支払いの項目を全部完了扱い）。
- 請求書ステータス：作成済み →（講師確定）確定済み →（理事）支払い済み／取消。確定済みは「差し戻し」（reopenInvoice）で作成済みに戻せる。
- 謝礼0円の請書（交通費等も無し）は「支払いなし」：講師用ページで振込先・インボイス欄を隠し、講師確定で完了扱い。
- 支払方法「現金」＝当日持参（見学会など。研修形式「見学」で新規作成すると初期値が現金）。講師用ページで振込先を求めない。
- 金額：消費税10%、源泉徴収10.21%（法人は0）、小数点以下切り捨て。
- 支払い済みから3日たつと、講師用リンクでは振込先を伏せる（請求書アクセスのある理事がログイン中なら全表示）。

## 反映（デプロイ）手順
1. 画面：`main` ブランチに push すると GitHub Pages に自動反映。
2. GAS：`gas/` で `clasp push` → `clasp deploy -i <既存のデプロイID>`（**新しいデプロイを作らない**。URLが変わり画面とつながらなくなる）。
   列の追加がある場合は、Apps Script エディタで `initializeSpreadsheet` を実行。
3. 本番の確認：ログイン → 各画面 → 講師用リンク（シークレットウィンドウ）。

## テスト
- 画面の確認は Playwright で、`script.google.com` への通信をモックしてローカルの HTML を開く（`localStorage['riji_portal_auth']` に `{idToken:'t', user:{...}, savedAt: Date.now()}` を入れるとログイン済みで開ける）。
- GAS は Node で `new Function(src)` による構文チェックと、SpreadsheetApp 等をモックした単体確認を行う。
