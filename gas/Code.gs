/**
 * NPO法人 日本繊維商品めんてなんす研究会　理事ポータル - バックエンド (Google Apps Script)
 *
 * ▼ セットアップ手順
 * 1. 管理用に使うGoogleスプレッドシートを新規作成する(空のままでOK)。
 * 2. スプレッドシートのメニュー「拡張機能」→「Apps Script」を開き、
 *    表示されたエディタの中身をこのファイルの内容に置き換える。
 * 3. 下記 CLIENT_ID を、Google Cloud Consoleで発行したOAuthクライアントIDに置き換える。
 *    (Cloud Console → APIとサービス → 認証情報 → OAuthクライアントID → ウェブアプリケーション
 *     承認済みのJavaScript生成元に、GitHub PagesのURL 例) https://<ユーザー名>.github.io を登録)
 * 4. エディタ上部の関数選択プルダウンで「initializeSpreadsheet」を選び、▷実行 を一度だけ押す。
 *    (必要なシートとヘッダー行、サンプルデータが自動作成されます。既存シートがある場合は
 *     安全に新しい列へ移行されます。再実行しても壊れません。)
 * 5. 作成された「AllowedAccounts」シートを開き、サンプル行を削除して、
 *    実際にログインを許可する理事のGmailアドレスを1行ずつ登録する。
 *    E列「請求書アクセス」に何か印(例：○)を入れた人だけが、請求書管理簿を利用できる。
 * 6. 「Departments」シートに、担当科(部署・分野)の一覧を登録する。
 *    B列「スキルアップコース区分」に印がある行が、講師依頼書のアーカイブ配信同意欄の
 *    表示判定に使われる。
 * 7. 「Settings」シートで、請求書関連アラートの基準日数を調整できる(コード変更不要)。
 *    「PulldownOptions」シートで、研修形式・講座ステータスなどのプルダウン選択肢を
 *    自由に追加・編集できる(コード変更不要)。
 * 8. 「デプロイ」→「新しいデプロイ」→ 種類の選択で「ウェブアプリ」を選択。
 *      実行するユーザー: 自分
 *      アクセスできるユーザー: 全員
 *    デプロイ後に表示される /exec で終わるURLを、フロントエンドの GAS_API_URL に設定する。
 * 9. スプレッドシートの内容を更新した場合や関数を修正した場合は、
 *    「新しいデプロイ」ではなく既存デプロイの「編集」→バージョンを「新バージョン」にして
 *    再デプロイすると、URLを変えずに更新できる。
 *
 * ▼ このコードを更新して反映するときは
 * ・コードを貼り替えたら、必ずもう一度「initializeSpreadsheet」を実行してください
 *   (開催場所のプルダウン選択肢や、交通費・宿泊費の列など、後から追加された項目が補われます。
 *    既存のデータが消えることはありません)。
 * ・その後、上記8の手順で「新バージョン」として再デプロイしてください。
 * ・もし開始時刻・終了時刻がスプレッドシート上で時刻の値(1899/12/30 9:00:00 のような表示)に
 *   自動変換されてしまっている場合は、「fixTimeColumns」を一度だけ実行すると、
 *   見た目も"09:00"のようなテキストに直せます(必須ではありません。アプリ上の表示は自動で補正されます)。
 * ・今回、講座PDFの保存(Googleドライブ)とダッシュボードのカレンダー表示を追加したため、
 *   再デプロイ時にGoogleから権限の再承認(ドライブ・カレンダーへのアクセス許可)を求められます。
 *   内容を確認のうえ「許可」してください。
 * ・ダッシュボードに表示するGoogleカレンダー(下記 CALENDAR_ID)は、このスクリプトを実行する
 *   Googleアカウント(通常は自分)に、Googleカレンダー側で「予定の表示(閲覧)」以上の権限で
 *   共有されている必要があります。共有されていない場合、ダッシュボードには
 *   「カレンダーを表示できません」という案内が出ます。
 * ・「PDFから自動入力」機能(ワードプレス用HTML作成画面)を使うには、Apps Scriptエディタの
 *   左メニュー「サービス」の＋ボタンから「Drive API」(高度なサービス)を追加し、
 *   案内されるGoogle CloudプロジェクトでもDrive APIを有効化してください。
 *   設定していない場合、「PDFから自動入力」を押すとエラーになりますが、他の機能には影響しません。
 * ・講座カードの「開催概要」→「準備進捗」画面(講師打ち合わせメモ・講座準備チェックリスト・
 *   講師資料PDF・参加者募集フォーム集計)を追加したため、コードを貼り替えたら必ず
 *   「initializeSpreadsheet」を再実行してください。
 *   Coursesシートに新しい列が追加され、PulldownOptionsシートに「講座準備チェック項目」という
 *   区分(チェックリストの項目一覧)が自動で追加されます。
 *   このチェックリストの項目は、PulldownOptionsシートの「講座準備チェック項目」の区分の行を
 *   追加・編集・削除するだけで、コード変更なしで自由に増減できます。
 *   また、この機能でも講師資料PDFをGoogleドライブに保存するため、再デプロイ時に
 *   ドライブへのアクセス許可の再承認を求められることがあります。
 * ・「準備進捗」画面の「参加者募集フォーム」欄に、別のGoogleアカウントで作成した参加者募集
 *   フォームの「回答」用スプレッドシートのURLを貼り付けると、講座ごとの申込件数(会場・ZOOMの
 *   内訳)を自動集計して表示します。この機能を使うには、そのスプレッドシートを、この
 *   Apps Scriptを実行しているGoogleアカウント(このコードを貼り付けたスプレッドシートの
 *   オーナー)宛てに「閲覧者」以上の権限で共有しておいてください。
 *   会場/ZOOMの内訳は、回答シートの見出し行から「参加方法」等それらしい列名を自動検出して
 *   判定しますが、うまく判定できない場合は「準備進捗」画面の詳細設定欄に、実際の列見出し
 *   文字列を入力すると、その列を使って集計するよう指定できます。
 * ・講座を追加・編集・削除すると、下記 CALENDAR_ID のGoogleカレンダーに予定を自動で
 *   作成・更新・削除するようにしました。この機能を使うには、そのカレンダーが、この
 *   スクリプトを実行するGoogleアカウントに「予定の変更」権限(閲覧だけでなく編集も可能)で
 *   共有されている必要があります。閲覧権限のみの場合、ダッシュボードの予定表示は今まで通り
 *   動きますが、自動登録・更新・削除はできず、講座の保存自体は失敗しません(エラーにはならず、
 *   静かにスキップされます)。コードを貼り替えたら「initializeSpreadsheet」を再実行し、
 *   Coursesシートに「calendarEventId」列が追加されたことを確認してください。
 * ・講座案内PDFのサムネイルが表示されない(📄アイコンのままになる)不具合を修正しました。
 *   新規アップロード分は自動的に修正済みの方式になりますが、既に登録済みの講座については、
 *   コードを貼り替えて再デプロイした後、エディタの関数選択プルダウンで「refreshAllCourseThumbnails」
 *   を選んで一度だけ実行すると、既存のサムネイルもまとめて直ります(必須の下準備はなく、
 *   この関数も上記のDrive API有効化さえ済んでいれば動作します。何度実行しても安全です)。
 * ・「資料庫」機能(各講座の資料を理事がアップロードして管理する画面materials.html、および
 *   講師が専用リンクからログイン不要で資料を提出できるmaterials-lecturer.html)を追加しました。
 *   コードを貼り替えたら、必ず「initializeSpreadsheet」を再実行してください(Materialsシートと
 *   PulldownOptionsシートの「資料区分」が自動で作成されます。実行しなくても、資料庫シート自体は
 *   初回アクセス時に自動作成されますが、区分のプルダウン選択肢を確認したい場合は実行してください)。
 *   講師用リンクは、資料庫で講座の「資料を管理」を開き、「講師用リンク」欄の「コピー」から取得できます
 *   (請求書の講師用リンクと同じ考え方で、講座IDが推測困難なため認証なしでアクセスできる仕組みです)。
 * ・「講座準備進捗確認」画面(course-progress.html)を追加しました。全講座の準備状況を
 *   企画決定→講師関係→募集準備→募集開始→申込管理→開催準備→開催→終了処理の8段階で表示・更新します。
 *   コードを貼り替えたら「initializeSpreadsheet」を再実行してください。Settingsシートに
 *   8段階の目安時期の行(phasePlanOffsetDays 等。開催日からの日数、マイナス＝開催前)が追加され、
 *   Coursesシートに「phaseJson」列が追加されます。目安の日数はSettingsシートの値を書き換えるだけで変更できます。
 * ・「SNS文章作成＆履歴」画面(sns.html)を追加しました。講座ごとにInstagramの投稿文(募集告知・開催前リマインド・
 *   開催報告)を雛形から作成し、下書き→投稿予定→投稿済みの履歴を管理します。コードを貼り替えたら
 *   「initializeSpreadsheet」を再実行してください。SnsPostsシート(履歴)とSnsTemplatesシート(雛形)が作成されます。
 *   雛形の文章・ハッシュタグは、SnsTemplatesシートの「本文」「ハッシュタグ」列を書き換えるだけで変更できます
 *   (本文中の{講座名}などの差し込み項目は、画面で投稿文を作るときに講座の情報に置き換わります)。
 * ・講座に「講師依頼時間」(講師の方にお願いする時間。講座の時間とは別)を追加し、SNS投稿文に講座案内PDFの
 *   読み取り(OCR)結果を差し込めるようにしました。コードを貼り替えたら「initializeSpreadsheet」を再実行して
 *   ください(Coursesシートに「講師依頼開始時刻」「講師依頼終了時刻」「pdfOcrText」「pdfOcrFileId」列、
 *   Invoicesシートに「courseStartTime」「courseEndTime」列が追加されます。既存のデータは消えません)。
 *   SNSの雛形も講座案内PDFの項目入りの新しい内容にしました。すでにSnsTemplatesシートがある場合、雛形は
 *   自動では置き換わりません。新しい雛形を使う場合は、SnsTemplatesシートの「募集告知」「開催前リマインド」
 *   「開催報告」の行を削除してから、SNS画面を開き直してください(足りない種類の行だけ新しい雛形で作り直されます)。
 * ・【2026-09-27 動作の安定化と高速化】新しいシートや列の追加はありません(initializeSpreadsheetの再実行は不要)。
 *   コードを貼り替えて「新しいバージョン」で再デプロイするだけで反映されます。主な変更点：
 *   - ログイン確認を最大6時間キャッシュ(以前は毎回Googleに問い合わせていた)。ログインの有効期限も
 *     1時間から6時間に延び、作業中に突然エラーになることが減ります。
 *   - AllowedAccountsシートの読み込みを1回にまとめ、5分間キャッシュ。登録を外した・請求書アクセスの印を
 *     変えた場合は最大5分で反映されます。すぐに反映したいときは、エディタで「clearAccountCache」を実行してください。
 *   - 申込集計(参加者募集フォームの回答シート)を3分間キャッシュし、複数講座をまとめて取得できるようにしました。
 *     「準備進捗」画面の「再集計」ボタンは、キャッシュを使わず最新を数えます。
 *   - Googleカレンダーの予定一覧を5分間キャッシュ(講座・役員会を保存したときは即時に最新化)。
 *   - 講座一覧からサムネイル画像を省いた軽量版を追加(トップページ以外の画面で使用し、通信量を削減)。
 *   - データを書き換える操作は1件ずつ順番に処理し、同時保存で片方の変更が消えるのを防ぎます。
 * ・【2026-09-27 セキュリティ点検】新しいシートや列の追加はありません(initializeSpreadsheetの再実行は不要)。
 *   - 講師用リンクからの請求書確定は「作成済み」のときだけ受け付ける(確定後・支払後・取消後の書き換えを防止)。
 *     交通費・宿泊費は「別途支給」のときだけ・100万円まで。入力文字数にも上限を設定。
 *   - 講師用の取得APIは必要な項目だけを返す(担当理事のメール・取消理由などは返さない)。
 *   - 講師の資料提出はファイルの種類(拡張子)・20MB・1講座1日30件に制限。中止講座は受け付けない。
 *   - すべてのシート書き込みで「=」「+」「-」「@」始まりの文字列を文字として保存(数式インジェクション対策)。
 *   - 月次報告書PDFに埋め込む画像は、報告書の添付フォルダ内の画像に限定。
 *   - URL欄(投稿URL・回答用スプレッドシートURL)は https:// のみ受け付け。GETアクセスはデータを返さない。
 * ・【2026-09-27 ファイルの閲覧範囲・振込先の表示】コードを貼り替えたら、必ず「initializeSpreadsheet」を1回実行してください。
 *   (実行時に「トリガーの作成」などの新しい権限の承認を求められます。承認してください)
 *   - 保存フォルダ(講座PDF・報告書添付・講座資料庫)の共有を、名簿(AllowedAccounts)の理事だけに変更し、
 *     既存ファイルの「リンクを知っている全員」共有も解除します(syncDriveSharing)。毎日午前4時ごろ自動で名簿と同期します。
 *   - 講師用リンクで、支払い済みから3日経過した請求書は口座番号を下4桁のみ表示し、その他の振込先を伏せます。
 * ・【2026-10-09(2) 表示の高速化その2】新しいシートや列の追加はありません(initializeSpreadsheetの再実行は不要)。
 *   - 月次報告書・資料庫・役員会・SNSの一覧を読むたびに、シートの準備(見出しの確認・列の書式の固定)を
 *     書き込みで行っていたため、準備進捗・月次報告書・SNSの画面が遅くなっていた。準備は6時間に1回にした
 *     (保存・削除のときは今まで通り毎回行う)。
 *   - SNS画面のデータ取得で、講座のサムネイル画像を読まないようにした。
 * ・【2026-10-09 表示の高速化と講座の条件】新しいシートや列の追加はありません(initializeSpreadsheetの再実行は不要)。
 *   - 講座一覧の読み込みで、PDFの読み取り結果(1件最大4万文字)とサムネイル画像の列を読まないようにし、
 *     講座の行を探す処理もID列だけを読むようにした(講座が増えるほど遅くなっていた原因)。
 *   - 5秒以上かかった処理を ActivityLog に「slow_request」として記録する(原因調査用)。
 *   - 講座準備進捗確認に「講師なし(勉強会など)」「講座案内PDFなし」を追加(phaseJson の naLecturer / naCoursePdf)。
 *   - 謝礼0円の請書(支払いなし)を判定して、準備進捗の支払いの項目を完了扱いにする(invoiceStates.noPayment)。
 * ・【2026-10-07 ログイン拒否の記録】新しいシートや列の追加はありません(initializeSpreadsheetの再実行は不要)。
 *   名簿にないアカウントでログインしようとしたとき、ActivityLog に「login_denied」として日時・メールアドレス・
 *   Googleアカウント名を記録します(同じ人は10分に1回)。請求書アクセスのない人が請求書機能を開いた場合も
 *   「invoice_access_denied」として記録します。Googleのログイン情報の確認で断った場合(メールアドレス未確認の
 *   Googleアカウントなど。期限切れは除く)は「login_token_rejected」として理由と一緒に記録します。
 * ・【2026-09-27 請求書の差し戻し】新しいシートや列の追加はありません。請求書管理簿の「その他の操作」に
 *   「講師に差し戻す（修正してもらう）」を追加しました(reopenInvoice。確定済み→作成済み。講師の入力内容は残ります)。
 * ・【2026-09-27 講座準備進捗の「不要」】新しいシートや列の追加はありません(initializeSpreadsheetの再実行は不要)。
 *   講座準備進捗確認で、参加者募集フォーム・講師資料PDF・理事資料を「不要」に設定できるようにしました
 *   (Coursesシートの phaseJson に naForm / naLecturerPdf / naDirectorMaterials として保存)。
 *   コードを貼り替えて「新しいバージョン」で再デプロイしてください。
 */

var CLIENT_ID = '199375582652-12dgtvao2mo6070n5gup3mrqrpeq9lif.apps.googleusercontent.com';

// ダッシュボードに表示するGoogleカレンダーのID(このスクリプトを実行するGoogleアカウントと
// 共有〈閲覧可〉されている必要があります)。
var CALENDAR_ID = 'tema.jp.analytics@gmail.com';

var SHEET_ALLOWED = 'AllowedAccounts';
var SHEET_COURSES = 'Courses';
var SHEET_LOG = 'ActivityLog';
var SHEET_INVOICES = 'Invoices';
var SHEET_DEPARTMENTS = 'Departments';
var SHEET_SETTINGS = 'Settings';
var SHEET_OPTIONS = 'PulldownOptions';
var SHEET_COMMENTS = 'Comments';
var SHEET_REPORTS = 'Reports';
var SHEET_BOARD_MEETINGS = 'BoardMeetings';
var SHEET_MATERIALS = 'Materials';
var SHEET_SNS_POSTS = 'SnsPosts';
var SHEET_SNS_TEMPLATES = 'SnsTemplates';

var COMMENT_HEADERS = ['id', 'courseId', 'authorName', 'authorEmail', 'body', 'createdAt'];

// 月次報告書(担当科ごとに月1件)。構造化フォーム形式：
// 報告者・直近の講座開催(タイトル/開催日/参加人数)を複数件・その他報告内容・課題改善点等。
// coursesJsonは [{ title, eventDate, venueCount, zoomCount }, ...] のJSON文字列。
var REPORT_HEADERS = [
  'id', 'department', 'month', 'reporterName', 'coursesJson', 'otherContent', 'issuesContent',
  'attachmentsJson', 'status', 'updatedBy', 'updatedAt', 'submittedAt'
];

// 役員会の開催日程(月次報告書の「未提出」アラート・提出期限の基準に使用)。
// 年度(4月始まり～翌3月)ごとに、fyMonthIndex(1=4月～12=翌3月)を1件ずつ、まとめて設定する。
var BOARD_MEETING_HEADERS = ['id', 'fiscalYear', 'fyMonthIndex', 'meetingDate', 'note', 'calendarEventId', 'updatedBy', 'updatedAt'];

var COURSE_HEADERS = [
  'id', '名称', '開催日', '開始時刻', '終了時刻', '開催場所', '研修形式',
  '担当理事', '担当科', 'ステータス', '作成日時', '更新日時',
  'pdfFileId', 'pdfFileName', 'pdfUrl', 'pdfThumbnailUrl',
  'lecturerMeetingDone', 'prepMaterialsDone',
  'lecturerMeetingNote', 'prepChecklistJson',
  'lecturerMaterialFileId', 'lecturerMaterialFileName', 'lecturerMaterialUrl',
  'formSheetUrl', 'formVenueColumnOverride', 'calendarEventId',
  'phaseJson',
  // 講師依頼時間(講師の方にお願いする時間。講座の時間＝開始時刻/終了時刻とは別。空欄なら講座の時間と同じ扱い)
  '講師依頼開始時刻', '講師依頼終了時刻',
  // 講座案内PDFの読み取り結果(OCR)のキャッシュ。pdfOcrFileIdが現在のpdfFileIdと一致する間だけ使う
  'pdfOcrText', 'pdfOcrFileId'
];

var INVOICE_HEADERS = [
  'id', 'status', 'courseId', 'department',
  'eventDate', 'startTime', 'endTime', 'venue', 'topic', 'format',
  'directorName', 'feeExTax', 'travelHandling', 'paymentMethod', 'paymentDueDate', 'equipmentJson',
  'invoiceRegistered', 'invoiceNumber', 'withholdingType',
  'lecturerName', 'lecturerAddress', 'lecturerContact',
  'lecturerCompanyName', 'lecturerCompanyAddress', 'lecturerCompanyContact', 'lecturerCompanyEmail',
  'bankName', 'branchName', 'accountType', 'accountNumber', 'accountHolderKana',
  'photoConsent', 'archiveConsent', 'archiveConsentCondition',
  'submitDate', 'signerName',
  'createdBy', 'createdAt', 'updatedBy', 'updatedAt',
  'paidAt', 'paidBy', 'voidReason', 'voidedBy', 'voidedAt',
  'travelFee', 'lodgingFee', 'directorMessage',
  // 講座の時間(請書兼請求書に「講座の時間」として表示する)。startTime/endTimeは「講師依頼時間」として扱う。
  'courseStartTime', 'courseEndTime'
];

// 各講座の資料庫(理事がアップロードする資料＋講師が専用リンクから提出する資料をまとめて管理)。
// 1資料＝1行。削除は請求書の「取消」と同じ考え方で論理削除とし(status='削除済み')、行自体は記録として残す。
var MATERIAL_HEADERS = [
  'id', 'courseId', 'department', 'courseName',
  'category', 'uploaderRole', 'uploaderName', 'note',
  'fileId', 'fileName', 'fileUrl', 'mimeType',
  'status', 'createdAt', 'updatedAt',
  'deletedBy', 'deletedAt'
];

// SNS投稿文の履歴(1投稿文＝1行)。status: 下書き / 投稿予定 / 投稿済み / 削除済み(論理削除)。
// 日付列(scheduledDate/postedDate)は、スプレッドシートの自動日付変換を避けるため書式なしテキストで保存する。
var SNS_POST_HEADERS = [
  'id', 'courseId', 'courseName', 'platform', 'postType', 'caption',
  'status', 'scheduledDate', 'postedDate', 'postUrl',
  'createdBy', 'createdAt', 'updatedBy', 'updatedAt', 'deletedBy', 'deletedAt'
];
// SNS投稿文の雛形(種類ごとに1行)。スプレッドシート上で自由に書き換えられる。
var SNS_TEMPLATE_HEADERS = ['種類', '本文', 'ハッシュタグ', '使える差し込み項目'];
var SNS_POST_TYPES = ['募集告知', '開催前リマインド', '開催報告'];
var SNS_PLACEHOLDER_HELP = '{講座名} {開催日} {開催日時} {時間} {会場} {担当科} {研修形式} {開催日までの日数} {申込数} {参加人数} ／ 講座案内PDFから：{講座内容} {講師} {対象} {定員} {参加費} {申込締切} {案内PDFの内容}（PDFから読み取れなかった項目は、その行ごと省かれます）';
var SNS_DEFAULT_HASHTAGS = '#日本繊維商品めんてなんす研究会 #TeMA #クリーニング #クリーニング店 #繊維製品 #講習会 #セミナー';
var SNS_DEFAULT_TEMPLATES = [
  ['募集告知',
   '【講座のご案内】\n{講座名}\n\n日本繊維商品めんてなんす研究会では、下記の講座を開催します。\nクリーニング・繊維製品のメンテナンスに携わる皆さまのご参加をお待ちしております。\n\n{講座内容}\n\n📅 日時：{開催日時}\n📍 会場：{会場}\n📝 形式：{研修形式}\n👤 講師：{講師}\n🎯 対象：{対象}\n👥 定員：{定員}\n💴 参加費：{参加費}\n⏰ 申込締切：{申込締切}\n\nお申し込み・詳細は、プロフィールのリンク（tema.jp）からご確認ください。',
   SNS_DEFAULT_HASHTAGS, SNS_PLACEHOLDER_HELP],
  ['開催前リマインド',
   '【まもなく開催】\n{講座名}\n\n開催まで、あと{開催日までの日数}日となりました。\nお申し込みがまだの方は、お早めにお申し込みください。\n\n📅 日時：{開催日時}\n📍 会場：{会場}\n👤 講師：{講師}\n⏰ 申込締切：{申込締切}\n\nお申し込みは、プロフィールのリンク（tema.jp）から。',
   SNS_DEFAULT_HASHTAGS, SNS_PLACEHOLDER_HELP],
  ['開催報告',
   '【開催報告】\n{講座名}\n\n{開催日}に「{講座名}」を開催しました。\n👤 講師：{講師}\n当日は{参加人数}の皆さまにご参加いただきました。ご参加いただいた皆さま、ありがとうございました。\n\n{講座内容}\n\n今後も講座を開催してまいります。最新の講座情報は、プロフィールのリンク（tema.jp）からご覧ください。',
   SNS_DEFAULT_HASHTAGS, SNS_PLACEHOLDER_HELP]
];

// 講座準備進捗確認の8段階。offsetは開催日からの日数(マイナス＝開催前、プラス＝開催後)で、その日を過ぎても
// 未完了なら「滞留」とみなす。目安の日数はSettingsシートの settingKey の行で変更できる(コード変更不要)。
var PREP_PHASES = [
  { key: 'plan',         label: '企画決定', settingKey: 'phasePlanOffsetDays',         defaultOffset: -90 },
  { key: 'lecturer',     label: '講師関係', settingKey: 'phaseLecturerOffsetDays',     defaultOffset: -60 },
  { key: 'recruitPrep',  label: '募集準備', settingKey: 'phaseRecruitPrepOffsetDays',  defaultOffset: -45 },
  { key: 'recruitStart', label: '募集開始', settingKey: 'phaseRecruitStartOffsetDays', defaultOffset: -40 },
  { key: 'applications', label: '申込管理', settingKey: 'phaseApplicationsOffsetDays', defaultOffset: -7 },
  { key: 'eventPrep',    label: '開催準備', settingKey: 'phaseEventPrepOffsetDays',    defaultOffset: -3 },
  { key: 'event',        label: '開催',     settingKey: 'phaseEventOffsetDays',        defaultOffset: 0 },
  { key: 'finish',       label: '終了処理', settingKey: 'phaseFinishOffsetDays',       defaultOffset: 30 }
];

/** 初回に一度だけ手動実行してください(シートの初期化。再実行しても安全) */
function initializeSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var allowed = ss.getSheetByName(SHEET_ALLOWED) || ss.insertSheet(SHEET_ALLOWED);
  if (allowed.getLastRow() === 0) {
    allowed.appendRow(safeRow_(['メールアドレス', '氏名', '役割', '登録日', '請求書アクセス']));
    allowed.appendRow(safeRow_(['example@gmail.com', '(サンプル・削除してください)', '理事', new Date(), '']));
  } else {
    var headerRow = allowed.getRange(1, 1, 1, 5).getValues()[0];
    if (headerRow[4] !== '請求書アクセス') {
      allowed.getRange(1, 5).setValue(safeCell_('請求書アクセス'));
    }
  }

  var courses = ss.getSheetByName(SHEET_COURSES) || ss.insertSheet(SHEET_COURSES);
  // 列を追加する前に、シートの列数を足りるだけ広げておく(足りないまま見出しを書き込むとエラーで止まるため)
  ensureColumnCapacity(courses, COURSE_HEADERS.length);
  if (courses.getLastRow() === 0) {
    courses.appendRow(safeRow_(COURSE_HEADERS));
  } else {
    migrateCoursesSheetIfNeeded(courses);
  }

  var log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(safeRow_(['日時', '操作者', '操作', '内容']));
  }

  var invoices = ss.getSheetByName(SHEET_INVOICES) || ss.insertSheet(SHEET_INVOICES);
  ensureColumnCapacity(invoices, INVOICE_HEADERS.length);
  if (invoices.getLastRow() === 0) {
    invoices.appendRow(safeRow_(INVOICE_HEADERS));
  } else {
    migrateInvoicesSheetIfNeeded(invoices);
  }

  var departments = ss.getSheetByName(SHEET_DEPARTMENTS) || ss.insertSheet(SHEET_DEPARTMENTS);
  if (departments.getLastRow() === 0) {
    departments.appendRow(safeRow_(['担当科名', 'スキルアップコース区分']));
    departments.appendRow(safeRow_(['スキルアップコース', '○']));
    departments.appendRow(safeRow_(['講師派遣', '']));
  }

  var settings = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  if (settings.getLastRow() === 0) {
    settings.appendRow(safeRow_(['項目', '値', '説明']));
    settings.appendRow(safeRow_(['createAlertDays', 40, '開催日の何日前までに請求書が未作成ならアラートを出すか']));
    settings.appendRow(safeRow_(['unrespondedAlertDays', 14, '開催日の何日前までに講師側の返信(確定)がなければアラートを出すか']));
    settings.appendRow(safeRow_(['paymentWarningDays', 2, '支払予定日の何日前から注意(黄色)表示にするか']));
  }
  // 講座準備進捗確認(8段階)の各段階の目安時期。開催日からの日数(マイナス＝開催前、プラス＝開催後)。
  // 既存のSettingsシートには、足りない行だけ追加する(既に値を変更していれば上書きしない)。
  PREP_PHASES.forEach(function (ph) {
    ensureSettingRow_(settings, ph.settingKey, ph.defaultOffset, '講座準備進捗確認「' + ph.label + '」の目安時期(開催日からの日数。マイナス＝開催前、プラス＝開催後)');
  });

  var options = ss.getSheetByName(SHEET_OPTIONS) || ss.insertSheet(SHEET_OPTIONS);
  if (options.getLastRow() === 0) {
    options.appendRow(safeRow_(['区分', '値']));
    ['座学', '見学', '実技', 'その他'].forEach(function (v) { options.appendRow(safeRow_(['研修形式', v])); });
    ['準備中', '実施予定', '実施済み', '中止'].forEach(function (v) { options.appendRow(safeRow_(['講座ステータス', v])); });
  }
  // 開催場所の選択肢(既存シートには後から追加分だけ足りない場合があるため、既存有無を見て追加する)
  ensurePulldownCategory(options, '開催場所', ['白王ビル２Fホール', '白王ビル３Fセミナールーム', 'その他（自由記入）']);
  // 講座準備チェック項目(「準備進捗」画面の③講座準備チェックリストに表示される項目。
  // ここに行を追加・編集・削除すれば、コード変更なしでチェックリストの項目を増減できる)
  ensurePulldownCategory(options, '講座準備チェック項目', [
    '受付名簿', '会場看板', '資料印刷', '筆記用具', '茶菓・飲み物', 'マイク・スピーカー確認', '座席配置'
  ]);
  // 資料庫の資料区分(理事が資料をアップロードする際の分類。行の追加・編集・削除で自由に増減できる)
  ensurePulldownCategory(options, '資料区分', [
    '教材・スライド', '配布資料(レジュメ等)', '準備資料', '写真・記録', 'その他'
  ]);

  var comments = ss.getSheetByName(SHEET_COMMENTS) || ss.insertSheet(SHEET_COMMENTS);
  if (comments.getLastRow() === 0) {
    comments.appendRow(safeRow_(COMMENT_HEADERS));
  }

  // Reports/BoardMeetingsは今回の機能追加で列の並び順・項目が変わっているため、
  // 単純な末尾追加(appendMissingHeaders)ではなく、列名でデータを対応付けし直す
  // migrateSheetHeaderOrder_()で安全に再構築する(以前のバージョンで一度でも
  // initializeSpreadsheetを実行したことがある場合の「見出しとデータの列がズレて
  // 提出済みにならない」不具合の再発防止)。
  var reports = ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);
  migrateSheetHeaderOrder_(reports, REPORT_HEADERS);

  var boardMeetings = ss.getSheetByName(SHEET_BOARD_MEETINGS) || ss.insertSheet(SHEET_BOARD_MEETINGS);
  migrateSheetHeaderOrder_(boardMeetings, BOARD_MEETING_HEADERS);

  var materials = ss.getSheetByName(SHEET_MATERIALS) || ss.insertSheet(SHEET_MATERIALS);
  migrateSheetHeaderOrder_(materials, MATERIAL_HEADERS);

  getSnsPostSheet_();
  getSnsTemplateSheet_();

  // 開始時刻・終了時刻の列は、GoogleスプレッドシートがHH:MM形式の文字列を時刻値へ自動変換してしまい、
  // 読み書きの際にズレが生じることがあるため、列の書式を「書式なしテキスト」に固定しておく。
  courses.getRange('D2:E').setNumberFormat('@');
  invoices.getRange('F2:G').setNumberFormat('@');
  courses.getRange(2, courseCol('講師依頼開始時刻'), courses.getMaxRows() - 1, 2).setNumberFormat('@');
  invoices.getRange(2, invoiceCol('courseStartTime'), invoices.getMaxRows() - 1, 2).setNumberFormat('@');

  Logger.log('初期化が完了しました。AllowedAccounts / Departments / Settings / PulldownOptions シートの内容をご確認ください。');

  // ドライブの共有を「ログインできる理事のみ」に合わせ、毎日の自動実行を設定する
  try { syncDriveSharing(); } catch (eShare) { Logger.log('ドライブの共有設定でエラー: ' + eShare); }
  try { installDriveSharingTrigger_(); } catch (eTrig) { Logger.log('自動実行の設定でエラー: ' + eTrig); }
}

/**
 * 【手動実行用・一度だけでOK】既存データの開始時刻・終了時刻が、
 * Googleスプレッドシートによって時刻の値(例：1899/12/30 9:00:00)へ自動変換されてしまっている場合に、
 * スプレッドシート上の見た目も"09:00"のようなテキストに戻す一括修正。
 * アプリ側の表示は既に自動で補正されるため必須ではないが、スプレッドシートを直接開いたときの
 * 見た目も揃えたい場合に、エディタの関数選択プルダウンから選んで実行してください。
 * 何度実行しても安全です。
 */
function fixTimeColumns() {
  fixTimeColumnsInSheet(getCourseSheet(), [courseCol('開始時刻'), courseCol('終了時刻'), courseCol('講師依頼開始時刻'), courseCol('講師依頼終了時刻')]);
  fixTimeColumnsInSheet(getInvoiceSheet(), [invoiceCol('startTime'), invoiceCol('endTime'), invoiceCol('courseStartTime'), invoiceCol('courseEndTime')]);
  Logger.log('開始時刻・終了時刻の表示を修正しました。');
}

function fixTimeColumnsInSheet(sheet, cols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  cols.forEach(function (col) {
    var range = sheet.getRange(2, col, lastRow - 1, 1);
    var values = range.getValues();
    var changed = false;
    var fixed = values.map(function (row) {
      var v = row[0];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        changed = true;
        return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm')];
      }
      return [v];
    });
    range.setNumberFormat('@');
    if (changed) range.setValues(safeRows_(fixed));
  });
}

/**
 * 【手動実行用・一度だけでOK】既に登録済みの講座案内PDFのサムネイルを、外部ホットリンク方式(URL)から
 * base64データURI方式へ一括で作り直す。drive.google.com/thumbnail?id=... の直接埋め込みが環境によって
 * 読み込めず、一部の講座カードで📄プレースホルダーのままになる不具合への対応(新規アップロード分は
 * uploadCoursePdf側で既に対応済み。この関数は、それより前にアップロード済みの講座を対象とする)。
 * pdfFileIdが設定されている行のみが対象。何度実行しても安全(常に最新のサムネイルで上書きする)。
 * 講座数が多い場合、実行に数分かかることがあります。
 */
function refreshAllCourseThumbnails() {
  var sheet = getCourseSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('講座データがありません。'); return; }
  var nameCol = courseCol('名称');
  var fileIdCol = courseCol('pdfFileId');
  var thumbCol = courseCol('pdfThumbnailUrl');
  var names = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues();
  var fileIds = sheet.getRange(2, fileIdCol, lastRow - 1, 1).getValues();
  var updated = 0, noPdf = 0, fetchFailed = 0, errored = 0;
  for (var i = 0; i < fileIds.length; i++) {
    var fileId = fileIds[i][0];
    if (!fileId) { noPdf++; continue; }
    // 1行の取得・保存に失敗しても、そこで処理全体が止まらないようにする
    // (以前のバージョンでセル文字数制限エラーが発生し、途中で処理が中断していた不具合の再発防止)。
    try {
      var dataUri = fetchDriveThumbnailDataUri_(fileId);
      sheet.getRange(2 + i, thumbCol).setValue(safeCell_(dataUri));
      if (dataUri) {
        updated++;
      } else {
        fetchFailed++;
        Logger.log('サムネイル取得できず(PDF未生成・非対応形式等)：行' + (2 + i) + '「' + names[i][0] + '」');
      }
    } catch (e) {
      errored++;
      Logger.log('行' + (2 + i) + '「' + names[i][0] + '」のサムネイル更新でエラー：' + (e && e.message ? e.message : e));
    }
  }
  Logger.log('サムネイルを更新しました：成功 ' + updated + '件 / PDF未添付 ' + noPdf + '件 / 取得できず(上のログ参照) ' + fetchFailed + '件 / エラー ' + errored + '件');
}

/** PulldownOptionsシートに、指定した区分の行が1つも無ければデフォルト値を追加する(既存シートへの追加移行用) */
function ensurePulldownCategory(sheet, category, defaults) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === category) return; // 既に存在する
  }
  defaults.forEach(function (v) { sheet.appendRow(safeRow_([category, v])); });
}

/** 旧形式(7列)のCoursesシートを新形式(12列)へ移行する。すでに新形式なら何もしない。 */
function migrateCoursesSheetIfNeeded(sheet) {
  var lastCol = sheet.getLastColumn();
  var header = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (header[3] === '開始時刻') {
    appendMissingHeaders(sheet, header, COURSE_HEADERS);
    return;
  }

  var data = sheet.getDataRange().getValues();
  var newRows = [COURSE_HEADERS];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    // 旧形式: id, 名称, 開催日, 担当理事, ステータス, 作成日時, 更新日時
    newRows.push([
      r[0], r[1] || '', r[2] || '', '', '', '', '座学',
      r[3] || '', '', r[4] || '準備中', r[5] || new Date(), r[6] || new Date()
    ]);
  }
  sheet.clear();
  sheet.getRange(1, 1, newRows.length, COURSE_HEADERS.length).setValues(safeRows_(newRows));
}

/** 旧形式(28列)のInvoicesシートを新形式へ移行する。すでに新形式なら、後から追加された列だけ補う。 */
function migrateInvoicesSheetIfNeeded(sheet) {
  var lastCol = sheet.getLastColumn();
  var header = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (header[2] === 'courseId') {
    appendMissingHeaders(sheet, header, INVOICE_HEADERS);
    return;
  }

  var data = sheet.getDataRange().getValues();
  var newRows = [INVOICE_HEADERS];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    // 旧形式: id,status,eventDate,startTime,endTime,venue,topic,format,directorName,feeExTax,
    //         travelHandling,paymentMethod,paymentDueDate,equipmentJson,invoiceRegistered,invoiceNumber,
    //         withholdingType,bankName,branchName,accountType,accountNumber,accountHolderKana,photoConsent,
    //         submitDate,signerName,createdBy,createdAt,updatedAt
    var obj = {
      id: r[0], status: r[1] || '作成済み', courseId: '', department: '',
      eventDate: r[2] || '', startTime: r[3] || '', endTime: r[4] || '',
      venue: r[5] || '', topic: r[6] || '', format: r[7] || '座学',
      directorName: r[8] || '', feeExTax: r[9] || 0,
      travelHandling: r[10] || '別途支給', paymentMethod: r[11] || '銀行振込',
      paymentDueDate: r[12] || '', equipmentJson: r[13] || '{}',
      invoiceRegistered: r[14] || '', invoiceNumber: r[15] || '', withholdingType: r[16] || '個人',
      lecturerName: '', lecturerAddress: '', lecturerContact: '',
      lecturerCompanyName: '', lecturerCompanyAddress: '', lecturerCompanyContact: '', lecturerCompanyEmail: '',
      bankName: r[17] || '', branchName: r[18] || '', accountType: r[19] || '普通',
      accountNumber: r[20] || '', accountHolderKana: r[21] || '',
      photoConsent: r[22] || '', archiveConsent: '', archiveConsentCondition: '',
      submitDate: r[23] || '', signerName: r[24] || '',
      createdBy: r[25] || '', createdAt: r[26] || new Date(),
      updatedBy: r[25] || '', updatedAt: r[27] || new Date(),
      paidAt: '', paidBy: '', voidReason: '', voidedBy: '', voidedAt: '',
      travelFee: 0, lodgingFee: 0, directorMessage: ''
    };
    newRows.push(INVOICE_HEADERS.map(function (h) { return obj[h]; }));
  }
  sheet.clear();
  sheet.getRange(1, 1, newRows.length, INVOICE_HEADERS.length).setValues(safeRows_(newRows));
}

/** 既存シートの見出し行に、現在のHEADERS配列にはあるが足りない列があれば、末尾に追加する(既存データは保持)。 */
function appendMissingHeaders(sheet, currentHeader, allHeaders) {
  var missing = allHeaders.filter(function (h) { return currentHeader.indexOf(h) === -1; });
  if (!missing.length) return;
  var startCol = currentHeader.length + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues(safeRows_([missing]));
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var defaultRow = missing.map(function () { return ''; });
    var defaultRows = [];
    for (var i = 0; i < lastRow - 1; i++) defaultRows.push(defaultRow.slice());
    sheet.getRange(2, startCol, lastRow - 1, missing.length).setValues(safeRows_(defaultRows));
  }
}

/**
 * 見出し行が想定の並び(targetHeaders)と完全に一致しない場合に、既存データを列名で
 * 対応付けし直しながらシートを安全に再構築する。appendMissingHeaders()は「末尾に追加」
 * しかしないため、項目の並び順や意味が変わった場合(列の削除・入れ替えを伴う仕様変更時)に
 * 既存データがズレた列に残ってしまう(reportCol()等が計算する列位置と実データの列がズレる)
 * 問題があった。ReportsシートやBoardMeetingsシートのように構造が変わりうる設定系シート向け。
 */
function migrateSheetHeaderOrder_(sheet, targetHeaders) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues(safeRows_([targetHeaders]));
    return;
  }

  var currentHeader = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var sameOrder = currentHeader.length === targetHeaders.length
    && targetHeaders.every(function (h, i) { return currentHeader[i] === h; });
  if (sameOrder) return; // 既に想定通りの並びなら何もしない

  var dataRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  var oldIndexByHeader = {};
  currentHeader.forEach(function (h, i) { if (h && !(h in oldIndexByHeader)) oldIndexByHeader[h] = i; });

  var rebuilt = dataRows
    .filter(function (row) { return row[0]; }) // idが空の行(空行)は除外
    .map(function (row) {
      return targetHeaders.map(function (h) {
        var oldIdx = oldIndexByHeader[h];
        return (oldIdx === undefined) ? '' : row[oldIdx];
      });
    });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, targetHeaders.length).setValues(safeRows_([targetHeaders]));
  if (rebuilt.length) {
    sheet.getRange(2, 1, rebuilt.length, targetHeaders.length).setValues(safeRows_(rebuilt));
  }
}

// ブラウザでURLを直接開いた場合(GET)は動作確認の応答だけを返す。データの取得・更新はPOSTのみ受け付ける
// (以前はGETでも全機能が動き、URLにログイン情報を含めて呼び出せてしまったため)。
function doGet(e) { return jsonOut({ ok: true, data: 'pong' }); }
function doPost(e) { return handleRequest(e); }

var SLOW_REQUEST_LOG_MS = 5000; // これより時間のかかった処理は ActivityLog に「slow_request」として記録する(原因調査用)

function handleRequest(e) {
  var lock = null;
  var startedAt = Date.now();
  try {
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      body = e.parameter;
    }
    var action = body.action;
    if (body.payload !== undefined && (body.payload === null || typeof body.payload !== 'object' || Array.isArray(body.payload))) {
      body.payload = {};
    }

    if (action === 'ping') {
      return jsonOut({ ok: true, data: 'pong' });
    }

    // 講師側(未ログイン)からもアクセスされるアクション。idの推測しづらさ(UUID)で保護している。
    var PUBLIC_ACTIONS = ['getInvoicePublic', 'submitInvoiceLecturer', 'getCourseMaterialsPublic', 'uploadMaterialLecturer'];
    var isPublicAction = PUBLIC_ACTIONS.indexOf(action) !== -1;

    // 請求書管理簿の機能。AllowedAccountsシートの「請求書アクセス」列に印がある理事のみ許可。
    var INVOICE_ACTIONS = ['listInvoices', 'createInvoice', 'updateInvoiceDirector', 'voidInvoice', 'markInvoicePaid', 'reopenInvoice', 'getSettings'];

    var user = null;
    if (!isPublicAction) {
      if (action === 'logout') { forgetIdToken_(body.idToken); return jsonOut({ ok: true, data: true }); }
      user = verifyIdToken(body.idToken);
      if (!user) {
        var rej = LAST_TOKEN_REJECT_ || {};
        // 期限切れ(通常のログイン切れ)以外の理由で断ったときは記録する(メールアドレス未確認のGoogleアカウントなど)
        if (rej.reason && rej.reason !== 'expired') {
          logDeniedAccess_({ email: rej.email || '(不明)', name: rej.name || '' }, 'login_token_rejected', TOKEN_REJECT_LABELS_[rej.reason] || rej.reason);
        }
        return jsonOut({ ok: false, error: 'unauthorized', reason: rej.reason || '' });
      }
      var account = getAccountInfo_(user.email); // AllowedAccountsシートは1回だけ読む(5分キャッシュ)
      if (!account.allowed) {
        logDeniedAccess_(user, 'login_denied', '名簿(AllowedAccounts)に登録されていないためログインを拒否');
        return jsonOut({ ok: false, error: 'not_allowed' });
      }
      if (INVOICE_ACTIONS.indexOf(action) !== -1 && !account.invoice) {
        logDeniedAccess_(user, 'invoice_access_denied', '請求書アクセスの印がないため請求書機能を拒否');
        return jsonOut({ ok: false, error: 'invoice_access_denied' });
      }
      // Googleアカウントのプロフィール名は未設定/表示名が無い場合があり、その際はコメント等の
      // 表示名が空になったりメールアドレスになったりしてしまう。AllowedAccountsシートに
      // 登録された氏名があれば、そちらを優先して使う(理事が自分で確実に登録した名前のため)。
      if (account.name) user.name = account.name;
      ensureDriveAccessForUser_(user.email); // 保存フォルダを閲覧できるようにする(6時間に1回だけ確認)
    }

    // 同時保存の取りこぼし防止：データを書き換える操作は1件ずつ順番に処理する(読み取り系は並列のまま)。
    // 例：2人の理事が同じ講座のチェック項目を同時に付けた場合に、片方の変更が消えるのを防ぐ。
    // 時間のかかる処理(PDF一本化・ファイルのアップロード)は、他の人の保存を待たせないよう対象外にする。
    var LOCK_EXEMPT = ['compileMonthlyReportsPdf', 'ocrCoursePdf', 'uploadCoursePdf', 'uploadLecturerMaterial', 'uploadReportAttachment',
      'uploadReportInlineImage', 'uploadMaterialDirector', 'uploadMaterialLecturer'];
    if (!/^(list|get)/.test(action) && LOCK_EXEMPT.indexOf(action) === -1) {
      lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) {
        lock = null;
        return jsonOut({ ok: false, error: '他の方の保存処理と重なり混み合っています。少し待ってからもう一度お試しください。' });
      }
    }

    var result;
    switch (action) {
      case 'listCourses':
        result = listCourses(body.payload);
        break;
      case 'createCourse':
        result = createCourse(body.payload, user);
        break;
      case 'updateCourse':
        result = updateCourse(body.payload, user);
        break;
      case 'deleteCourse':
        result = deleteCourse(body.payload, user);
        break;
      case 'uploadCoursePdf':
        result = uploadCoursePdf(body.payload, user);
        break;
      case 'deleteCoursePdf':
        result = deleteCoursePdf(body.payload, user);
        break;
      case 'updatePrepProgress':
        result = updatePrepProgress(body.payload, user);
        break;
      case 'uploadLecturerMaterial':
        result = uploadLecturerMaterial(body.payload, user);
        break;
      case 'deleteLecturerMaterial':
        result = deleteLecturerMaterial(body.payload, user);
        break;
      case 'getFormSignupSummary':
        result = getFormSignupSummary(body.payload);
        break;
      case 'getFormSignupSummaries':
        result = getFormSignupSummaries(body.payload);
        break;
      case 'listUpcomingEvents':
        result = listUpcomingEvents(body.payload);
        break;
      case 'listComments':
        result = listComments(body.payload);
        break;
      case 'addComment':
        result = addComment(body.payload, user);
        break;
      case 'listAnnouncements':
        result = listAnnouncements(body.payload);
        break;
      case 'ocrCoursePdf':
        result = ocrCoursePdf(body.payload, user);
        break;
      case 'listDepartments':
        result = listDepartments();
        break;
      case 'listOptions':
        result = listOptions();
        break;
      case 'getSettings':
        result = getSettings();
        break;
      case 'listInvoices':
        result = listInvoices();
        break;
      case 'createInvoice':
        result = createInvoice(body.payload, user);
        break;
      case 'updateInvoiceDirector':
        result = updateInvoiceDirector(body.payload, user);
        break;
      case 'voidInvoice':
        result = voidInvoice(body.payload, user);
        break;
      case 'reopenInvoice':
        result = reopenInvoice(body.payload, user);
        break;
      case 'markInvoicePaid':
        result = markInvoicePaid(body.payload, user);
        break;
      case 'getInvoicePublic':
        result = getInvoicePublic(body.payload, body.idToken);
        break;
      case 'submitInvoiceLecturer':
        result = submitInvoiceLecturer(body.payload);
        break;
      case 'listReports':
        result = listReports(body.payload);
        break;
      case 'saveReport':
        result = saveReport(body.payload, user);
        break;
      case 'uploadReportAttachment':
        result = uploadReportAttachment(body.payload, user);
        break;
      case 'uploadReportInlineImage':
        result = uploadReportInlineImage(body.payload, user);
        break;
      case 'deleteReportAttachment':
        result = deleteReportAttachment(body.payload, user);
        break;
      case 'compileMonthlyReportsPdf':
        result = compileMonthlyReportsPdf(body.payload, user);
        break;
      case 'getReportSubmissionStatus':
        result = getReportSubmissionStatus(body.payload);
        break;
      case 'listBoardMeetings':
        result = listBoardMeetings();
        break;
      case 'getBoardMeetingsForYear':
        result = getBoardMeetingsForYear(body.payload);
        break;
      case 'saveBoardMeetingsForYear':
        result = saveBoardMeetingsForYear(body.payload, user);
        break;
      case 'getCoursePdfFile':
        result = getCoursePdfFile(body.payload);
        break;
      case 'getCoursePdfText':
        result = getCoursePdfText(body.payload, user);
        break;
      case 'listSnsData':
        result = listSnsData(body.payload);
        break;
      case 'saveSnsPost':
        result = saveSnsPost(body.payload, user);
        break;
      case 'deleteSnsPost':
        result = deleteSnsPost(body.payload, user);
        break;
      case 'getPrepOverview':
        result = getPrepOverview(body.payload);
        break;
      case 'updatePrepItem':
        result = updatePrepItem(body.payload, user);
        break;
      case 'listMaterials':
        result = listMaterials(body.payload);
        break;
      case 'uploadMaterialDirector':
        result = uploadMaterialDirector(body.payload, user);
        break;
      case 'deleteMaterial':
        result = deleteMaterial(body.payload, user);
        break;
      case 'getCourseMaterialsPublic':
        result = getCourseMaterialsPublic(body.payload);
        break;
      case 'uploadMaterialLecturer':
        result = uploadMaterialLecturer(body.payload);
        break;
      default:
        return jsonOut({ ok: false, error: 'unknown_action' });
    }
    var out = jsonOut({ ok: true, data: result, user: user });
    var elapsed = Date.now() - startedAt;
    if (elapsed >= SLOW_REQUEST_LOG_MS) {
      try { appendLog(user ? user.email : '(講師・未ログイン)', 'slow_request', action + '　' + (elapsed / 1000).toFixed(1) + '秒'); } catch (eSlow) {}
    }
    return out;
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

/** ===== ログイン確認(高速化のためキャッシュを使う) =====
 * 以前はAPIを呼ぶたびに毎回Googleのtokeninfoへ問い合わせ(1回0.2〜0.5秒)、さらに
 * AllowedAccountsシートを3回読んでいたため、ダッシュボードの表示が遅くなっていた。
 * 一度確認できたログインは最大6時間キャッシュし(=ポータル側のログイン有効時間)、
 * 登録アカウントの情報は5分間キャッシュする(シートで登録を外した場合も5分以内に反映される)。 */
var SESSION_CACHE_SECONDS = 21600;   // 6時間(CacheServiceの上限)
var ACCOUNT_CACHE_SECONDS = 300;     // 5分

function tokenCacheKey_(idToken) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(idToken), Utilities.Charset.UTF_8);
  return 'sess_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
}

/** IDトークンを確認し、{ email, name } を返す(確認できなければnull)。確認済みのものはキャッシュから返す。 */
function verifyIdToken(idToken) {
  if (!idToken) return null;
  var cache = null, key = '';
  try { cache = CacheService.getScriptCache(); key = tokenCacheKey_(idToken); } catch (e) { cache = null; }
  if (cache) {
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e2) {} }
  }
  var user = verifyIdTokenRemote_(idToken);
  if (user && cache) { try { cache.put(key, JSON.stringify(user), SESSION_CACHE_SECONDS); } catch (e3) {} }
  return user;
}

/** 直近にIDトークンを断った理由(記録用)。reason: expired / invalid / wrong_client / email_unverified */
var LAST_TOKEN_REJECT_ = null;
var TOKEN_REJECT_LABELS_ = {
  invalid: 'Googleのログイン情報を確認できなかったためログインを拒否',
  wrong_client: '別のアプリ用のログイン情報だったためログインを拒否(CLIENT_IDの設定違いの可能性)',
  email_unverified: 'Googleアカウントのメールアドレスが未確認(確認メールのリンクを押していない)のためログインを拒否'
};

/** 記録用に、IDトークンの中身(メールアドレス・名前・期限)を検証せずに読み取る。判定には使わない */
function peekIdToken_(idToken) {
  try {
    var part = String(idToken || '').split('.')[1] || '';
    part = part.replace(/-/g, '+').replace(/_/g, '/');
    while (part.length % 4) part += '=';
    return JSON.parse(Utilities.newBlob(Utilities.base64Decode(part)).getDataAsString());
  } catch (e) { return {}; }
}

/** GoogleのtokeninfoエンドポイントでIDトークンを検証し、メールアドレス等を取り出す */
function verifyIdTokenRemote_(idToken) {
  var peek = peekIdToken_(idToken);
  function reject(reason) {
    LAST_TOKEN_REJECT_ = { reason: reason, email: String(peek.email || '').slice(0, 200), name: String(peek.name || '').slice(0, 100) };
    return null;
  }
  var res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) {
    var expired = peek.exp && Number(peek.exp) * 1000 < Date.now();
    return reject(expired ? 'expired' : 'invalid');
  }
  var info = JSON.parse(res.getContentText());
  if (info.aud !== CLIENT_ID) return reject('wrong_client'); // 自分のアプリ宛てのトークンでなければ拒否
  if (info.email_verified !== 'true' && info.email_verified !== true) return reject('email_unverified');
  return { email: info.email, name: info.name || info.email };
}

/** ログアウト時に、そのトークンのキャッシュを消す */
function forgetIdToken_(idToken) {
  if (!idToken) return;
  try { CacheService.getScriptCache().remove(tokenCacheKey_(idToken)); } catch (e) {}
}

/** AllowedAccountsシートを1回だけ読み、{ allowed, name(B列の登録氏名), invoice(E列の請求書アクセス) } を返す。5分キャッシュ。 */
function getAccountInfo_(email) {
  var em = String(email || '').trim().toLowerCase();
  var cache = null, key = '';
  try {
    cache = CacheService.getScriptCache();
    key = 'acct_' + (cache.get('acct_ver') || '0') + '_' + em; // clearAccountCache()で世代を変えると一斉に無効になる
  } catch (e) { cache = null; }
  if (cache) {
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e2) {} }
  }
  var info = { allowed: false, name: '', invoice: false };
  var values = getSheet(SHEET_ALLOWED).getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === em) {
      info = { allowed: true, name: String(values[i][1] || '').trim(), invoice: String(values[i][4] || '').trim() !== '' };
      break;
    }
  }
  if (cache) { try { cache.put(key, JSON.stringify(info), ACCOUNT_CACHE_SECONDS); } catch (e3) {} }
  return info;
}

/** AllowedAccountsシートに登録されたメールアドレスかどうか */
function isAllowed(email) { return getAccountInfo_(email).allowed; }

/** AllowedAccountsシートに登録された氏名(B列)。登録が無い/空欄なら空文字。 */
function getAllowedAccountName(email) { return getAccountInfo_(email).name; }

/** AllowedAccountsシートの「請求書アクセス」列(E列)に印がある人だけ、請求書管理簿の機能を許可する */
function hasInvoiceAccess(email) { return getAccountInfo_(email).invoice; }

/** AllowedAccountsシートを編集した直後に、すぐ反映させたい場合にエディタから実行する(通常は5分で自動反映) */
function clearAccountCache() {
  CacheService.getScriptCache().put('acct_ver', Utilities.getUuid(), SESSION_CACHE_SECONDS);
  Logger.log('登録アカウントのキャッシュを消去しました(シートの変更がすぐに反映されます)');
  syncDriveSharing(); // ドライブの閲覧権限も名簿に合わせる
}

/** ===== 設定・マスタ ===== */

function listDepartments() {
  var sheet = getSheet(SHEET_DEPARTMENTS);
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    rows.push({
      name: values[i][0],
      isSkillUpCourse: String(values[i][1] || '').trim() !== ''
    });
  }
  return rows;
}

/** 担当科名 → スキルアップコース区分(true/false) のマップ。講師依頼書のアーカイブ同意欄の表示判定に使う。 */
function departmentSkillUpMap() {
  var sheet = getSheet(SHEET_DEPARTMENTS);
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    map[values[i][0]] = String(values[i][1] || '').trim() !== '';
  }
  return map;
}

/** PulldownOptionsシートから、区分ごとのプルダウン選択肢一覧を取得する(SS側で自由に編集可能) */
function listOptions() {
  var sheet = getSheet(SHEET_OPTIONS);
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var category = values[i][0];
    var value = values[i][1];
    if (!category || !value) continue;
    if (!map[category]) map[category] = [];
    map[category].push(value);
  }
  return map;
}

function getSettings() {
  var sheet = getSheet(SHEET_SETTINGS);
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    map[values[i][0]] = values[i][1];
  }
  return {
    createAlertDays: Number(map.createAlertDays) || 40,
    unrespondedAlertDays: Number(map.unrespondedAlertDays) || 14,
    paymentWarningDays: Number(map.paymentWarningDays) || 2
  };
}

/** ===== 講座マスタ (Courses) ===== */

function courseCol(fieldName) {
  var idx = COURSE_HEADERS.indexOf(fieldName);
  if (idx === -1) throw new Error('invalid_course_field:' + fieldName);
  return idx + 1;
}

function courseRowToObject(values) {
  return {
    id: values[0], name: values[1], eventDate: formatDate(values[2]),
    startTime: formatTimeValue(values[3]), endTime: formatTimeValue(values[4]), venue: values[5], format: values[6],
    directorName: values[7], department: values[8], status: values[9],
    createdAt: formatDate(values[10]), updatedAt: formatDate(values[11]),
    pdfFileId: values[courseCol('pdfFileId') - 1] || '',
    pdfFileName: values[courseCol('pdfFileName') - 1] || '',
    pdfUrl: values[courseCol('pdfUrl') - 1] || '',
    pdfThumbnailUrl: values[courseCol('pdfThumbnailUrl') - 1] || '',
    lecturerMeetingDone: values[courseCol('lecturerMeetingDone') - 1] === true,
    prepMaterialsDone: values[courseCol('prepMaterialsDone') - 1] === true,
    lecturerMeetingNote: values[courseCol('lecturerMeetingNote') - 1] || '',
    prepChecklistJson: values[courseCol('prepChecklistJson') - 1] || '{}',
    lecturerMaterialFileId: values[courseCol('lecturerMaterialFileId') - 1] || '',
    lecturerMaterialFileName: values[courseCol('lecturerMaterialFileName') - 1] || '',
    lecturerMaterialUrl: values[courseCol('lecturerMaterialUrl') - 1] || '',
    formSheetUrl: values[courseCol('formSheetUrl') - 1] || '',
    formVenueColumnOverride: values[courseCol('formVenueColumnOverride') - 1] || '',
    calendarEventId: values[courseCol('calendarEventId') - 1] || '',
    // 講座準備進捗確認(8段階)のうち、手動でチェックする段階の記録。{ recruitStarted:'YYYY-MM-DD', applicationsClosed:'YYYY-MM-DD' }
    phaseJson: values[courseCol('phaseJson') - 1] || '{}',
    lecturerStartTime: formatTimeValue(values[courseCol('講師依頼開始時刻') - 1]),
    lecturerEndTime: formatTimeValue(values[courseCol('講師依頼終了時刻') - 1])
    // pdfOcrText は大きくなりうるため講座一覧には含めない(getCoursePdfTextで個別に取得する)
  };
}

/** 講座一覧。payload.lite=true のときはサムネイル画像(1件あたり最大約45KB)を省く。
 *  サムネイルを表示するのはトップページだけなので、他の画面は軽量版を使い通信量を大きく減らす。 */
function listCourses(payload) {
  var lite = !!(payload && payload.lite);
  var sheet = getCourseSheet();
  var values = readCourseRowsForList_(sheet, lite);
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) continue;
    var obj = courseRowToObject(values[i]);
    if (lite) obj.pdfThumbnailUrl = '';
    rows.push(obj);
  }
  rows.sort(function (a, b) { return String(b.eventDate).localeCompare(String(a.eventDate)); });
  return rows;
}

function findCourseRowIndex(sheet, id) {
  if (!id) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  // 【高速化】ID列(A列)だけを読む(以前はシート全体を読んでおり、PDFの読み取り結果やサムネイル画像の大きなセルまで毎回読み込んでいた)
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return -1;
}

/**
 * 【高速化】講座一覧用に、Coursesシートから必要な列だけを読む(見出し行を除く)。
 * 「pdfOcrText」(PDFの読み取り結果。1件最大4万文字)は一覧では使わないので読まない。
 * lite のときは「pdfThumbnailUrl」(サムネイル画像。1件最大4.5万文字)も読まない。
 * 講座が増えるほどこの大きなセルの読み込みが重くなり、各画面の表示が遅くなっていた。
 */
function readCourseRowsForList_(sheet, lite) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var n = lastRow - 1;
  var width = courseCol('pdfOcrText') - 1; // pdfOcrText より前の列まで
  var thumbCol = courseCol('pdfThumbnailUrl');
  if (!lite) return sheet.getRange(2, 1, n, width).getValues();
  // サムネイルの列は読まずに飛ばし、その位置には空欄を入れて列の並びを保つ
  var left = sheet.getRange(2, 1, n, thumbCol - 1).getValues();
  var right = sheet.getRange(2, thumbCol + 1, n, width - thumbCol).getValues();
  var out = [];
  for (var i = 0; i < n; i++) out.push(left[i].concat(['']).concat(right[i]));
  return out;
}

/* ===== 講座 → Googleカレンダー 自動連携 ===== */

/** "HH:mm"形式の文字列を { h, m } に分解する。解釈できなければnull */
function parseHm_(s) {
  if (!s) return null;
  var m = String(s).match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return { h: Number(m[1]), m: Number(m[2]) };
}

/**
 * 講座の開催日・開始時刻・終了時刻から、カレンダー予定の開始/終了日時を組み立てる。
 * 開催日が無ければnull。時刻が無ければ終日予定として扱う。
 */
function buildCourseEventTimes_(payload) {
  if (!payload || !payload.eventDate) return null;
  var m = String(payload.eventDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  var y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);

  var startTime = parseHm_(payload.startTime);
  if (!startTime) {
    return { allDay: true, start: new Date(y, mo, d) };
  }
  var start = new Date(y, mo, d, startTime.h, startTime.m, 0);
  var endTime = parseHm_(payload.endTime);
  var end;
  if (endTime) {
    end = new Date(y, mo, d, endTime.h, endTime.m, 0);
    if (end.getTime() <= start.getTime()) end = new Date(start.getTime() + 60 * 60000);
  } else {
    end = new Date(start.getTime() + 60 * 60000);
  }
  return { allDay: false, start: start, end: end };
}

/**
 * 講座の内容をもとに、CALENDAR_ID のGoogleカレンダーへ予定を作成・更新・削除する。
 * カレンダーへの書き込み権限が無い、共有されていない等の理由で失敗しても、
 * 講座そのものの保存を妨げないよう、例外を投げずに元のcalendarEventIdをそのまま返す。
 * 戻り値：新しい(または変わらない)calendarEventId。開催日が未入力になった場合は''を返す。
 */
function syncCourseCalendarEvent_(payload, existingEventId) {
  clearCalendarCache_();
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) return existingEventId || '';

    var times = buildCourseEventTimes_(payload);
    if (!times) {
      if (existingEventId) {
        try {
          var evToRemove = cal.getEventById(existingEventId);
          if (evToRemove) evToRemove.deleteEvent();
        } catch (eRemove) {}
      }
      return '';
    }

    var title = payload.name || '(講座名未設定)';
    var location = payload.venue || '';
    var descLines = [];
    if (payload.directorName) descLines.push('担当理事：' + payload.directorName);
    if (payload.department) descLines.push('担当科：' + payload.department);
    if (payload.format) descLines.push('研修形式：' + payload.format);
    if (payload.lecturerStartTime || payload.lecturerEndTime) {
      descLines.push('講師依頼時間：' + (payload.lecturerStartTime || '') + '〜' + (payload.lecturerEndTime || ''));
    }
    descLines.push('（理事ポータルの講座管理より自動登録・更新されます。内容の変更は理事ポータルで行ってください）');
    var description = descLines.join('\n');

    var event = null;
    if (existingEventId) {
      try { event = cal.getEventById(existingEventId); } catch (eGet) { event = null; }
    }

    if (event) {
      if (times.allDay) {
        event.setAllDayDate(times.start);
      } else {
        event.setTime(times.start, times.end);
      }
      event.setTitle(title);
      event.setLocation(location);
      event.setDescription(description);
      return event.getId();
    }

    var created = times.allDay
      ? cal.createAllDayEvent(title, times.start, { location: location, description: description })
      : cal.createEvent(title, times.start, times.end, { location: location, description: description });
    return created.getId();
  } catch (err) {
    // カレンダーが共有されていない・編集権限が無い等。講座の保存は継続させる。
    return existingEventId || '';
  }
}

/** 講師依頼時間(開始・終了)を書き込む。時刻の自動変換を避けるため、書き込む前にセルを書式なしテキストにする */
function setCourseLecturerTimes_(sheet, rowIndex, payload) {
  var startCol = courseCol('講師依頼開始時刻');
  var headerCell = sheet.getRange(1, startCol, 1, 2);
  var header = headerCell.getValues()[0];
  if (!header[0] || !header[1]) headerCell.setValues(safeRows_([['講師依頼開始時刻', '講師依頼終了時刻']])); // initializeSpreadsheet未実行でも見出しを入れる
  var range = sheet.getRange(rowIndex, startCol, 1, 2);
  range.setNumberFormat('@');
  range.setValues(safeRows_([[payload.lecturerStartTime || '', payload.lecturerEndTime || '']]));
}

function createCourse(payload, user) {
  var sheet = getCourseSheet();
  var id = Utilities.getUuid();
  var now = new Date();
  var calendarEventId = syncCourseCalendarEvent_(payload, '');
  sheet.appendRow(safeRow_([
    id, payload.name || '', payload.eventDate || '',
    payload.startTime || '', payload.endTime || '', payload.venue || '', payload.format || '座学',
    payload.directorName || '', payload.department || '', payload.status || '準備中',
    now, now
  ]));
  var newRowIndex = sheet.getLastRow();
  if (calendarEventId) {
    sheet.getRange(newRowIndex, courseCol('calendarEventId')).setValue(safeCell_(calendarEventId));
  }
  setCourseLecturerTimes_(sheet, newRowIndex, payload);
  appendLog(user.email, 'createCourse', payload.name);
  return { id: id, calendarEventId: calendarEventId };
}

// 注意：lecturerMeetingDone・prepMaterialsDone・lecturerMeetingNote・prepChecklistJson は
// 「準備進捗」画面専用のupdatePrepProgress()で更新する。この関数(講座の基本情報の編集)では触れない。
function updateCourse(payload, user) {
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.id);
  if (rowIndex === -1) throw new Error('course_not_found');
  var existingEventId = sheet.getRange(rowIndex, courseCol('calendarEventId')).getValue() || '';
  sheet.getRange(rowIndex, courseCol('名称'), 1, 8).setValues(safeRows_([[
    payload.name || '', payload.eventDate || '',
    payload.startTime || '', payload.endTime || '', payload.venue || '', payload.format || '座学',
    payload.directorName || '', payload.department || ''
  ]]));
  sheet.getRange(rowIndex, courseCol('ステータス')).setValue(safeCell_(payload.status || '準備中'));
  setCourseLecturerTimes_(sheet, rowIndex, payload);
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  var newEventId = syncCourseCalendarEvent_(payload, existingEventId);
  if (newEventId !== existingEventId) {
    sheet.getRange(rowIndex, courseCol('calendarEventId')).setValue(safeCell_(newEventId));
  }
  appendLog(user.email, 'updateCourse', payload.name);
  return { id: payload.id, calendarEventId: newEventId };
}

function deleteCourse(payload, user) {
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.id);
  if (rowIndex === -1) throw new Error('course_not_found');

  // 取消以外の請求書が紐付いている場合は、記録の整合性のため削除を拒否する(不正防止の方針に合わせる)
  var invoiceSheet = getInvoiceSheet();
  var invoiceValues = invoiceSheet.getDataRange().getValues();
  var courseIdCol = invoiceCol('courseId') - 1;
  var statusCol = invoiceCol('status') - 1;
  for (var i = 1; i < invoiceValues.length; i++) {
    if (invoiceValues[i][courseIdCol] === payload.id && invoiceValues[i][statusCol] !== '取消') {
      throw new Error('course_has_invoice');
    }
  }

  var pdfFileId = sheet.getRange(rowIndex, courseCol('pdfFileId')).getValue();
  if (pdfFileId) { try { DriveApp.getFileById(pdfFileId).setTrashed(true); } catch (e) {} }
  var materialFileId = sheet.getRange(rowIndex, courseCol('lecturerMaterialFileId')).getValue();
  if (materialFileId) { try { DriveApp.getFileById(materialFileId).setTrashed(true); } catch (e) {} }

  var calendarEventId = sheet.getRange(rowIndex, courseCol('calendarEventId')).getValue();
  if (calendarEventId) {
    clearCalendarCache_();
    try {
      var calForDelete = CalendarApp.getCalendarById(CALENDAR_ID);
      if (calForDelete) {
        var evForDelete = calForDelete.getEventById(calendarEventId);
        if (evForDelete) evForDelete.deleteEvent();
      }
    } catch (eCalDel) { /* カレンダー権限が無い等。講座の削除は継続する */ }
  }

  sheet.deleteRow(rowIndex);
  appendLog(user.email, 'deleteCourse', payload.id);
  return { id: payload.id };
}

/** ===== ドライブの共有範囲(ログインできる理事のみ) =====
 * 【2026-09-27 変更】以前は保存したファイルを「リンクを知っている全員が閲覧可」にしていたため、リンクが
 * 漏れると誰でも見られた。現在はファイル個別のリンク共有をやめ、保存先の3つのフォルダを、名簿シート
 * (AllowedAccounts)に登録された理事のGoogleアカウントにだけ「閲覧者」として共有する(フォルダ内のファイルは
 * フォルダの共有設定を引き継ぐ)。講師など名簿にない人は、リンクを知っていても開けない。
 * ・理事が初めてポータルを使ったとき、その理事を自動でフォルダの閲覧者に追加する(共有のお知らせメールは送らない)。
 * ・名簿から外した理事の閲覧権限は、毎日の自動実行(syncDriveSharing)で外れる。すぐに外したい場合は
 *   エディタで「syncDriveSharing」を実行する(clearAccountCacheを実行した場合も同時に行う)。
 */
var DRIVE_SHARE_CACHE_SECONDS = 21600; // 6時間

function portalFolders_() {
  return [getOrCreateCoursePdfFolder(), getOrCreateReportFolder_(), getOrCreateMaterialsFolder_()];
}

/** 名簿シートに登録されたメールアドレス(小文字)の一覧。サンプル行は除く */
function rosterEmails_() {
  var values = getSheet(SHEET_ALLOWED).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var em = String(values[i][0] || '').trim().toLowerCase();
    if (em && em.indexOf('@') > 0 && em !== 'example@gmail.com' && out.indexOf(em) === -1) out.push(em);
  }
  return out;
}

function folderUserEmails_(folder) {
  var out = [];
  folder.getViewers().concat(folder.getEditors()).forEach(function (u) {
    var em = String(u.getEmail() || '').toLowerCase();
    if (em) out.push(em);
  });
  try { var owner = folder.getOwner(); if (owner) out.push(String(owner.getEmail() || '').toLowerCase()); } catch (e) {}
  return out;
}

/** フォルダに閲覧者を追加する(共有のお知らせメールは送らない。Drive APIが使えない場合は通常の方法で追加) */
function addFolderViewer_(folder, email) {
  try {
    Drive.Permissions.insert({ role: 'reader', type: 'user', value: email }, folder.getId(), { sendNotificationEmails: false });
  } catch (e) {
    folder.addViewer(email);
  }
}

/** ログインした理事が、保存フォルダを閲覧できるようにする(6時間に1回だけ確認。失敗しても本来の処理は続ける) */
function ensureDriveAccessForUser_(email) {
  var em = String(email || '').trim().toLowerCase();
  if (!em) return;
  var cache = null, key = '';
  try {
    cache = CacheService.getScriptCache();
    key = 'drv_' + (cache.get('acct_ver') || '0') + '_' + em;
    if (cache.get(key)) return;
    cache.put(key, '1', DRIVE_SHARE_CACHE_SECONDS); // 同時に届いた他の通信で重複して追加しないよう先に記録する
  } catch (e) { return; }
  try {
    portalFolders_().forEach(function (folder) {
      if (folderUserEmails_(folder).indexOf(em) === -1) addFolderViewer_(folder, em);
    });
  } catch (e2) {
    try { cache.remove(key); } catch (e3) {}
  }
}

/**
 * 保存フォルダの共有を名簿に合わせる(管理者向け。毎日自動実行＋エディタから手動実行も可)。
 * ・フォルダとフォルダ内のファイルの「リンクを知っている全員」共有を解除する(以前のファイルも含む)
 * ・名簿の理事を閲覧者に追加し、名簿にない閲覧者を外す(編集者・オーナーはそのまま)
 */
function syncDriveSharing() {
  var roster = rosterEmails_();
  var added = 0, removed = 0, privatized = 0;
  portalFolders_().forEach(function (folder) {
    try { if (folder.getSharingAccess() !== DriveApp.Access.PRIVATE) { folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); privatized++; } } catch (e) {}
    var viewers = folder.getViewers().map(function (u) { return String(u.getEmail() || '').toLowerCase(); });
    var present = folderUserEmails_(folder);
    roster.forEach(function (em) {
      if (present.indexOf(em) === -1) { try { addFolderViewer_(folder, em); added++; } catch (e) { Logger.log('共有できませんでした: ' + em + ' ' + e); } }
    });
    viewers.forEach(function (em) {
      if (em && roster.indexOf(em) === -1) { try { folder.removeViewer(em); removed++; } catch (e) {} }
    });
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      try { if (f.getSharingAccess() !== DriveApp.Access.PRIVATE) { f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); privatized++; } } catch (e) {}
    }
  });
  try { var c = CacheService.getScriptCache(); c.put('acct_ver', Utilities.getUuid(), SESSION_CACHE_SECONDS); } catch (e) {}
  var msg = 'ドライブの共有を名簿に合わせました：追加 ' + added + '件／解除 ' + removed + '件／リンク共有を解除したファイル・フォルダ ' + privatized + '件';
  Logger.log(msg);
  return msg;
}

/** syncDriveSharing を毎日自動で実行する設定を作る(すでにあれば何もしない) */
function installDriveSharingTrigger_() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'syncDriveSharing'; });
  if (!exists) ScriptApp.newTrigger('syncDriveSharing').timeBased().everyDays(1).atHour(4).create();
}

/** 講座案内PDFを保存するGoogleドライブのフォルダを取得(無ければ作成)する */
function getOrCreateCoursePdfFolder() {
  var name = '理事ポータル_講座PDF';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

/**
 * 【サムネイル不具合対応】講座案内PDFのサムネイルを、外部への直接リンク(drive.google.com/thumbnail?id=...)
 * ではなく、GASの実行権限(OAuth)でGoogleドライブから画像そのものを取得し、base64データURIとして返す。
 * drive.google.com/thumbnail?id=... の直接埋め込みは「Anyone with the link」共有でも、GitHub Pages等の
 * 外部サイトからの読み込みが環境によって失敗することがある(実際に一部の講座カードで📄プレースホルダーの
 * ままになる不具合を確認)。base64データURIとして埋め込めば、表示時に外部へ画像リクエストが発生しないため
 * この問題が起きない。
 * アップロード直後はGoogleドライブ側のサムネイル生成が間に合っていないことがあるため、
 * 空だった場合は1.2秒待って1回だけ再試行する。取得できなければ空文字を返し、
 * フロントエンド側の既存の📄プレースホルダー表示にフォールバックさせる。
 * 【重要】スプレッドシートの1セルには50,000文字までしか入らないため、Driveのサムネイル取得URL末尾の
 * サイズ指定を小さく(幅200px相当)書き換えたうえで取得し、それでもbase64化した結果が安全マージンを
 * 超える場合は保存自体をあきらめて空文字を返す(この場合も📄プレースホルダー表示にフォールバックする
 * だけで、エラーにはならない)。
 */
function fetchDriveThumbnailDataUri_(fileId) {
  var MAX_DATA_URI_LENGTH = 45000; // シートの1セル上限(50,000文字)に対する安全マージン
  for (var attempt = 0; attempt < 2; attempt++) {
    try {
      var meta = Drive.Files.get(fileId);
      var thumbnailLink = meta && meta.thumbnailLink;
      if (thumbnailLink) {
        // 末尾の "=s220" や "=w220-h220" のようなサイズ指定を外し、幅200px相当を明示的に指定し直す
        // (指定が無い/一致しない場合はそのまま末尾に追加する)。
        var smallUrl = thumbnailLink.replace(/=[sw]\d+(-[a-z0-9]+)*$/i, '') + '=s200';
        var res = UrlFetchApp.fetch(smallUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        });
        if (res.getResponseCode() === 200) {
          var blob = res.getBlob();
          var mime = blob.getContentType() || 'image/jpeg';
          var dataUri = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
          if (dataUri.length <= MAX_DATA_URI_LENGTH) {
            return dataUri;
          }
          // 縮小指定してもなお大きすぎる場合は、セルの文字数制限を超えるため保存しない
        }
      }
    } catch (e) {
      // 取得失敗時は再試行、それでもダメなら呼び出し元で空文字として扱う
    }
    if (attempt === 0) Utilities.sleep(1200);
  }
  return '';
}

/** 講座に案内PDFを添付(既に添付済みの場合は差し替え)する。payload: { courseId, fileName, mimeType, base64Data } */
function uploadCoursePdf(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  if (!payload.base64Data) throw new Error('no_file');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');

  var oldFileId = sheet.getRange(rowIndex, courseCol('pdfFileId')).getValue();
  if (oldFileId) { try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {} }

  var bytes = Utilities.base64Decode(payload.base64Data);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/pdf', payload.fileName || '講座案内.pdf');
  var folder = getOrCreateCoursePdfFolder();
  var file = folder.createFile(blob);

  var fileId = file.getId();
  var viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
  // 【サムネイル不具合対応】外部ホットリンクURLではなく、base64データURIとして取得・保存する。
  var thumbnailUrl = fetchDriveThumbnailDataUri_(fileId);

  sheet.getRange(rowIndex, courseCol('pdfFileId'), 1, 4).setValues(safeRows_([[
    fileId, payload.fileName || '講座案内.pdf', viewUrl, thumbnailUrl
  ]]));
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'uploadCoursePdf', payload.fileName || '');
  return { pdfFileId: fileId, pdfFileName: payload.fileName || '講座案内.pdf', pdfUrl: viewUrl, pdfThumbnailUrl: thumbnailUrl };
}

/** 講座に添付されたPDFを削除する。payload: { courseId } */
function deleteCoursePdf(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');

  var oldFileId = sheet.getRange(rowIndex, courseCol('pdfFileId')).getValue();
  if (oldFileId) { try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {} }

  sheet.getRange(rowIndex, courseCol('pdfFileId'), 1, 4).setValues(safeRows_([['', '', '', '']]));
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'deleteCoursePdf', payload.courseId);
  return { id: payload.courseId };
}

/**
 * 準備進捗(講師打ち合わせメモ・チェック済みフラグ・備品チェックリスト・参加者募集フォームの
 * 集計元スプレッドシート情報)を更新する。講座の基本情報(名称・日時など)には触れない専用API。
 * payload: { id, lecturerMeetingDone, lecturerMeetingNote, prepChecklistJson, prepMaterialsDone,
 *            formSheetUrl, formVenueColumnOverride }
 */
function updatePrepProgress(payload, user) {
  if (!payload || !payload.id) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.id);
  if (rowIndex === -1) throw new Error('course_not_found');
  sheet.getRange(rowIndex, courseCol('lecturerMeetingDone'), 1, 2).setValues(safeRows_([[
    !!payload.lecturerMeetingDone, !!payload.prepMaterialsDone
  ]]));
  sheet.getRange(rowIndex, courseCol('lecturerMeetingNote')).setValue(safeCell_(payload.lecturerMeetingNote || ''));
  sheet.getRange(rowIndex, courseCol('prepChecklistJson')).setValue(safeCell_(payload.prepChecklistJson || '{}'));
  sheet.getRange(rowIndex, courseCol('formSheetUrl')).setValue(safeCell_(/^[A-Za-z0-9_-]{20,}$/.test(String(payload.formSheetUrl || '').trim()) ? String(payload.formSheetUrl).trim() : safeHttpUrl_(payload.formSheetUrl))); // URLまたはスプレッドシートIDのみ
  sheet.getRange(rowIndex, courseCol('formVenueColumnOverride')).setValue(safeCell_(payload.formVenueColumnOverride || ''));
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'updatePrepProgress', payload.id);
  return { id: payload.id };
}

/** ===== 講座進捗コメント (Comments) ===== */

function getCommentSheet() {
  var sheet = getSheet(SHEET_COMMENTS);
  ensureColumnCapacity(sheet, COMMENT_HEADERS.length);
  return sheet;
}

function commentRowToObject(values) {
  return {
    id: values[0], courseId: values[1], authorName: values[2], authorEmail: values[3],
    body: values[4], createdAt: formatDateTime(values[5])
  };
}

/**
 * 講座進捗のコメント一覧を取得する。
 * payload.courseId を指定すればその講座のコメントのみ、省略すれば全講座から新しい順に取得する
 * (ダッシュボードの「最近のコメント」表示に使用)。payload.limit で件数を絞り込める。
 */
function listComments(payload) {
  var sheet = getCommentSheet();
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var obj = commentRowToObject(values[i]);
    if (payload && payload.courseId && obj.courseId !== payload.courseId) continue;
    rows.push(obj);
  }
  rows.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  if (payload && payload.limit) rows = rows.slice(0, payload.limit);
  return rows;
}

/** 講座進捗にコメントを投稿する。payload: { courseId, body } */
function addComment(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var body = (payload.body || '').trim();
  if (!body) throw new Error('empty_comment');
  var sheet = getCommentSheet();
  var id = Utilities.getUuid();
  var now = new Date();
  sheet.appendRow(safeRow_([id, payload.courseId, user.name || user.email, user.email, body, now]));
  appendLog(user.email, 'addComment', payload.courseId);
  return { id: id, createdAt: formatDateTime(now) };
}

/* ===== ホームページのお知らせ (tema.jp) ===== */

// tema.jpトップページの「TeMAの最新情報はこちら」欄をそのまま再現するため、取得元は
// トップページ自体にする(以前は一覧ページ/new_course_list/を見ていたが、そちらは
// トップページの「最新情報」欄とは別の全件一覧ページで、表示内容が一致しなかったため)。
var ANNOUNCEMENTS_SOURCE_URL = 'https://tema.jp/';
// トップページの「TeMAの最新情報はこちら」欄の直後に「新着一覧」の記事リンクが並んでいるが、
// トップページには他にも(下部の「ベーシックコース」案内カード等)/new-courses/配下と紛らわしい
// リンクが無いとは限らないため、この見出し文字列(トップページ内に一意)より後ろの範囲だけを
// 対象にスクレイピングする。
var ANNOUNCEMENTS_SECTION_HEADING = 'TeMAの最新情報はこちら';
// フォールバック表示時・「もっと見る」の遷移先は、お知らせの全件一覧ページのほうが親切なのでこちら。
var ANNOUNCEMENTS_MORE_URL = 'https://tema.jp/new_course_list/';
var ANNOUNCEMENTS_CACHE_KEY = 'tema_announcements_v2';
var ANNOUNCEMENTS_CACHE_SECONDS = 3600; // 1時間キャッシュ(サイトへの負荷軽減・表示速度向上のため)

/**
 * tema.jpの「新着情報」ページからお知らせ一覧を取得する(スクレイピング)。
 * 取得に成功した場合のみScriptCacheに1時間キャッシュする(失敗結果をキャッシュすると
 * 復旧後も最大1時間フォールバック表示のままになってしまうため)。取得やサイト構造の
 * 変化で失敗した場合は、ページへのリンク1件のみを返すフォールバックにする(エラーにはしない)。
 */
function listAnnouncements(payload) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(ANNOUNCEMENTS_CACHE_KEY);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        if (parsed && parsed.length) return parsed;
      } catch (e2) { /* キャッシュが壊れていれば再取得 */ }
    }
    var items = fetchAnnouncementsFromSite_(); // 取得失敗時はnull
    if (items && items.length) {
      try { cache.put(ANNOUNCEMENTS_CACHE_KEY, JSON.stringify(items), ANNOUNCEMENTS_CACHE_SECONDS); } catch (e3) {}
      return items;
    }
    return announcementsFallback_();
  } catch (err) {
    return announcementsFallback_();
  }
}

function announcementsFallback_() {
  return [{ title: '最新のお知らせを tema.jp で見る', url: ANNOUNCEMENTS_MORE_URL }];
}

/** tema.jpのHTMLを取得し、お知らせ記事(/new-courses/)へのリンクを抽出する。失敗時はnullを返す。 */
function fetchAnnouncementsFromSite_() {
  var res;
  try {
    res = UrlFetchApp.fetch(ANNOUNCEMENTS_SOURCE_URL, {
      muteHttpExceptions: true,
      followRedirects: true,
      // GASのデフォルトUser-Agentだとセキュリティプラグイン等に弾かれ空HTMLが
      // 返ってくる場合があるため、通常のブラウザに近いヘッダーを付与する。
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'
      }
    });
  } catch (fetchErr) {
    return null;
  }
  if (!res || res.getResponseCode() !== 200) return null;

  var html = res.getContentText();

  // トップページの「TeMAの最新情報はこちら」という見出し(トップページ内に一意に1箇所だけ
  // 存在する)より後ろの範囲だけを対象にする。以前は「お知らせ」「ベーシックコース」といった
  // 単語で範囲を切り出していたが、ナビゲーションメニュー等にも同じ単語が含まれておりズレる
  // 場合があったため、より一意性の高い見出し文言に変更した。見出しが見つからない場合は
  // サイト構造が変わった可能性が高いため、誤った内容を拾うよりフォールバック扱いにする。
  var sectionStart = html.indexOf(ANNOUNCEMENTS_SECTION_HEADING);
  if (sectionStart === -1) return null;
  var sectionHtml = html.slice(sectionStart);

  // 各記事は「タイトルへのリンク」と「詳細はこちら等の汎用リンク」の2本のリンクを
  // 持つ場合があり、記事によってどちらが先に出現するか分からないため、同じURLに対して
  // 複数のリンクが見つかった場合は汎用的な文言(詳細はこちら 等)よりも具体的なタイトルを
  // 優先して採用する。
  var GENERIC_LINK_TEXTS = ['詳細はこちら', '詳細を見る', '続きを見る', '続きを読む', 'もっと見る', 'read more', 'more'];
  var urlToEntry = {};
  var urlOrder = [];
  var linkRegex = /<a\s+[^>]*href="([^"]*\/new-courses\/[^"]*)"[^>]*>([\s\S]*?)<\/a>/gi;
  var match;
  var scanned = 0;
  while ((match = linkRegex.exec(sectionHtml)) !== null && scanned < 60 && urlOrder.length < 12) {
    scanned++;
    var url = match[1];
    var rawTitle = decodeHtmlEntities_(stripHtmlTags_(match[2]).trim());
    if (!rawTitle) continue;
    if (!/^https?:\/\//.test(url)) {
      url = 'https://tema.jp' + (url.charAt(0) === '/' ? url : '/' + url);
    }
    var isGeneric = GENERIC_LINK_TEXTS.indexOf(rawTitle.toLowerCase()) !== -1;
    if (!urlToEntry[url]) {
      urlToEntry[url] = { title: rawTitle, isGeneric: isGeneric };
      urlOrder.push(url);
    } else if (urlToEntry[url].isGeneric && !isGeneric) {
      urlToEntry[url].title = rawTitle;
      urlToEntry[url].isGeneric = false;
    }
  }
  var items = urlOrder.slice(0, 6).map(function (u) {
    return { title: urlToEntry[u].title, url: u };
  });
  return items.length ? items : null;
}

function stripHtmlTags_(s) {
  return String(s || '').replace(/<[^>]+>/g, '');
}

function decodeHtmlEntities_(s) {
  return String(s || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, n) { return String.fromCharCode(parseInt(n, 10)); });
}

/** 講座に講師資料PDFを添付(既に添付済みの場合は差し替え)する。payload: { courseId, fileName, mimeType, base64Data } */
function uploadLecturerMaterial(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  if (!payload.base64Data) throw new Error('no_file');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');

  var oldFileId = sheet.getRange(rowIndex, courseCol('lecturerMaterialFileId')).getValue();
  if (oldFileId) { try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {} }

  var bytes = Utilities.base64Decode(payload.base64Data);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/pdf', payload.fileName || '講師資料.pdf');
  var folder = getOrCreateCoursePdfFolder();
  var file = folder.createFile(blob);

  var fileId = file.getId();
  var viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

  sheet.getRange(rowIndex, courseCol('lecturerMaterialFileId'), 1, 3).setValues(safeRows_([[
    fileId, payload.fileName || '講師資料.pdf', viewUrl
  ]]));
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'uploadLecturerMaterial', payload.fileName || '');
  return { lecturerMaterialFileId: fileId, lecturerMaterialFileName: payload.fileName || '講師資料.pdf', lecturerMaterialUrl: viewUrl };
}

/** 講座に添付された講師資料PDFを削除する。payload: { courseId } */
function deleteLecturerMaterial(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');

  var oldFileId = sheet.getRange(rowIndex, courseCol('lecturerMaterialFileId')).getValue();
  if (oldFileId) { try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {} }

  sheet.getRange(rowIndex, courseCol('lecturerMaterialFileId'), 1, 3).setValues(safeRows_([['', '', '']]));
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'deleteLecturerMaterial', payload.courseId);
  return { id: payload.courseId };
}

/**
 * 別のGoogleアカウントで作成された参加者募集フォームの回答集計を取得する。
 * 回答用スプレッドシートを、このスクリプトを実行するGoogleアカウントに「閲覧者」以上で
 * 共有しておく必要がある(共有していない場合は form_sheet_access_failed エラーになる)。
 *
 * 「会場」参加・「ZOOM」参加の内訳は、回答シートの見出し行(1行目)から、
 * 列名に「参加方法」「参加形態」「参加形式」「会場」「ZOOM」のいずれかを含む列を自動検出し、
 * その列の値に「会場」という文字が含まれていれば会場参加、"zoom"(大文字小文字問わず)が
 * 含まれていればZOOM参加として集計する。列名の付け方がフォームごとに違って自動検出できない
 * 場合は、payload.columnOverride に実際の列見出し文字列を渡すと、その列を使って集計する
 * (「準備進捗」画面の詳細設定欄から指定可能)。
 *
 * payload: { sheetUrl, columnOverride }
 * 戻り値: { total, venueCount, zoomCount, otherCount, columnUsed, autoDetected }
 */
var FORM_SUMMARY_CACHE_SECONDS = 180; // 申込集計は3分キャッシュ(外部スプレッドシートを開く処理が重いため)

function getFormSignupSummary(payload) {
  var rawUrl = ((payload && payload.sheetUrl) || '').trim();
  if (!rawUrl) throw new Error('form_sheet_url_missing');
  var idMatch = rawUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  var sheetId = idMatch ? idMatch[1] : rawUrl;
  var override = ((payload && payload.columnOverride) || '').trim();

  var cache = null, key = '';
  try {
    cache = CacheService.getScriptCache();
    var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sheetId + '|' + override, Utilities.Charset.UTF_8);
    key = 'form_' + Utilities.base64EncodeWebSafe(d);
  } catch (e) { cache = null; }
  if (cache && !(payload && payload.force)) {
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e2) {} }
  }
  var result = computeFormSignupSummary_(sheetId, override);
  if (cache) { try { cache.put(key, JSON.stringify(result), FORM_SUMMARY_CACHE_SECONDS); } catch (e3) {} }
  return result;
}

/**
 * 複数講座の申込集計をまとめて取得する(画面から講座ごとに何十回もAPIを呼ばないようにするため)。
 * payload: { items: [{ courseId, sheetUrl, columnOverride }], force }
 * 戻り値: { courseId: 集計結果 | { error } }
 */
function getFormSignupSummaries(payload) {
  var items = (payload && payload.items) || [];
  var out = {};
  items.slice(0, 20).forEach(function (it) {
    if (!it || !it.courseId) return;
    try {
      out[it.courseId] = getFormSignupSummary({ sheetUrl: it.sheetUrl, columnOverride: it.columnOverride, force: payload.force });
    } catch (e) {
      out[it.courseId] = { error: String(e && e.message ? e.message : e) };
    }
  });
  return out;
}

function computeFormSignupSummary_(sheetId, override) {
  var ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    throw new Error('form_sheet_access_failed:' + (err && err.message ? err.message : String(err)));
  }

  var sheets = ss.getSheets();
  var target = sheets[0];
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('回答') !== -1) { target = sheets[i]; break; }
  }

  var values = target.getDataRange().getValues();
  if (values.length < 1) {
    return { total: 0, venueCount: 0, zoomCount: 0, otherCount: 0, columnUsed: '', autoDetected: false };
  }

  var headers = values[0];
  var colIndex = -1;
  if (override) {
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]) === override) { colIndex = h; break; }
    }
  }

  var autoDetected = false;
  if (colIndex === -1) {
    var keywords = ['参加方法', '参加形態', '参加形式', '会場', 'ZOOM', 'Zoom', 'zoom'];
    for (var k = 0; k < keywords.length && colIndex === -1; k++) {
      for (var h2 = 0; h2 < headers.length; h2++) {
        if (String(headers[h2]).indexOf(keywords[k]) !== -1) { colIndex = h2; autoDetected = true; break; }
      }
    }
  }

  var total = 0, venueCount = 0, zoomCount = 0, otherCount = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var isEmpty = row.every(function (c) { return c === '' || c === null; });
    if (isEmpty) continue;
    total++;
    if (colIndex === -1) continue;
    var cell = String(row[colIndex] || '');
    if (cell.indexOf('会場') !== -1) venueCount++;
    else if (/zoom/i.test(cell)) zoomCount++;
    else otherCount++;
  }

  return {
    total: total, venueCount: venueCount, zoomCount: zoomCount, otherCount: otherCount,
    columnUsed: colIndex !== -1 ? String(headers[colIndex]) : '', autoDetected: autoDetected
  };
}

/** ダッシュボード用：指定日数以内のGoogleカレンダーの予定一覧を取得する。payload: { days } */
var CALENDAR_CACHE_SECONDS = 300; // 予定一覧は5分キャッシュ(講座・役員会を保存したときは即時に消去する)
function clearCalendarCache_() {
  try { CacheService.getScriptCache().removeAll(['cal_30', 'cal_60', 'cal_90']); } catch (e) {}
}

function listUpcomingEvents(payload) {
  var days = (payload && Number(payload.days)) || 60;
  var cache = null, key = 'cal_' + days;
  try { cache = CacheService.getScriptCache(); } catch (e0) { cache = null; }
  if (cache) {
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e1) {} }
  }
  var list = listUpcomingEventsRaw_(days);
  if (cache) { try { cache.put(key, JSON.stringify(list), CALENDAR_CACHE_SECONDS); } catch (e2) {} }
  return list;
}

function listUpcomingEventsRaw_(days) {
  var cal;
  try {
    cal = CalendarApp.getCalendarById(CALENDAR_ID);
  } catch (e) {
    throw new Error('calendar_not_accessible');
  }
  if (!cal) throw new Error('calendar_not_accessible');

  var now = new Date();
  var until = new Date(now.getTime() + days * 86400000);
  var events = cal.getEvents(now, until);
  return events.slice(0, 100).map(function (e) {
    return {
      title: e.getTitle(),
      start: Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
      end: Utilities.formatDate(e.getEndTime(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
      allDay: e.isAllDayEvent(),
      location: e.getLocation() || ''
    };
  });
}

/**
 * 講座に添付されたPDFをOCRでテキスト化する(ワードプレス用HTML作成画面の自動入力に利用)。
 * 【要設定】Apps Scriptエディタの「サービス」から「Drive API」(高度なサービス)を追加し、
 * 紐づくGoogle CloudプロジェクトでもDrive APIを有効化しておく必要があります。
 * payload: { courseId }
 */
function ocrCoursePdf(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');
  var pdfFileId = sheet.getRange(rowIndex, courseCol('pdfFileId')).getValue();
  if (!pdfFileId) throw new Error('pdf_not_found');

  var text = ocrPdfFileText_(pdfFileId);
  appendLog(user.email, 'ocrCoursePdf', payload.courseId);
  return { text: text };
}

/**
 * 講座案内PDFのファイルそのものを返す(要ログイン)。SNS画面で、PDFの各ページをブラウザ上で
 * Instagram用の正方形画像に書き出すために使う(Googleドライブから直接は画面に読み込めないため)。
 * payload: { courseId } 戻り値: { base64, fileName, mimeType }
 */
function getCoursePdfFile(payload) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');
  var pdfFileId = sheet.getRange(rowIndex, courseCol('pdfFileId')).getValue();
  if (!pdfFileId) throw new Error('pdf_not_found');
  var file = DriveApp.getFileById(pdfFileId);
  var blob = file.getBlob();
  var bytes = blob.getBytes();
  if (bytes.length > 15 * 1024 * 1024) throw new Error('pdf_too_large');
  return { base64: Utilities.base64Encode(bytes), fileName: file.getName(), mimeType: blob.getContentType() || 'application/pdf' };
}

/** Drive上のPDFをOCRしてテキストを返す(要：Drive API 高度なサービス)。失敗時は ocr_failed:<詳細> を投げる */
function ocrPdfFileText_(pdfFileId) {
  var tempDocId = null;
  try {
    var blob = DriveApp.getFileById(pdfFileId).getBlob();
    // resourceにmimeTypeを(Googleドキュメント形式などへ)指定してしまうと、
    // 「すでに変換済みのファイル」として扱われOCRが行えないため、ここでは指定しない。
    // convert:true と ocr:true を指定することで、PDFがOCR付きでGoogleドキュメントへ変換される。
    var resource = { title: 'OCR_TEMP_' + Utilities.getUuid() };
    var converted = Drive.Files.insert(resource, blob, { convert: true, ocr: true, ocrLanguage: 'ja' });
    tempDocId = converted.id;
    return DocumentApp.openById(tempDocId).getBody().getText();
  } catch (err) {
    // 原因究明のため、詳しいエラー内容をそのまま画面側に返す(ocr_failed: の後ろが実際のエラーメッセージ)
    throw new Error('ocr_failed:' + (err && err.message ? err.message : String(err)));
  } finally {
    if (tempDocId) { try { Drive.Files.remove(tempDocId); } catch (e) {} }
  }
}

/**
 * 講座案内PDFの読み取り結果(OCRテキスト)を返す(要ログイン。SNS投稿文の作成に使う)。
 * 一度読み取った結果はCoursesシート(pdfOcrText)に保存し、PDFが差し替えられるまで使い回す
 * (OCRは10〜20秒かかるため)。payload: { courseId, force(trueで読み直す) }
 * 戻り値: { text, noPdf(PDF未登録), cached(保存済みの結果を返した) }
 */
function getCoursePdfText(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');
  var pdfFileId = sheet.getRange(rowIndex, courseCol('pdfFileId')).getValue();
  if (!pdfFileId) return { text: '', noPdf: true, cached: false };

  var cacheRange = sheet.getRange(rowIndex, courseCol('pdfOcrText'), 1, 2);
  var cache = cacheRange.getValues()[0];
  if (!payload.force && cache[1] && cache[1] === pdfFileId) {
    return { text: String(cache[0] || ''), noPdf: false, cached: true };
  }
  var text = ocrPdfFileText_(pdfFileId);
  var MAX_CACHE_LENGTH = 40000; // スプレッドシートの1セル上限(50,000文字)に対する安全マージン
  var header = sheet.getRange(1, courseCol('pdfOcrText'), 1, 2);
  var hv = header.getValues()[0];
  if (!hv[0] || !hv[1]) header.setValues(safeRows_([['pdfOcrText', 'pdfOcrFileId']])); // initializeSpreadsheet未実行でも見出しを入れる
  cacheRange.setValues(safeRows_([[text.slice(0, MAX_CACHE_LENGTH), pdfFileId]]));
  appendLog(user.email, 'getCoursePdfText', payload.courseId);
  return { text: text, noPdf: false, cached: false };
}

/** ===== 請書兼請求書 (Invoices) ===== */

function invoiceCol(fieldName) {
  var idx = INVOICE_HEADERS.indexOf(fieldName);
  if (idx === -1) throw new Error('invalid_invoice_field:' + fieldName);
  return idx + 1;
}

function invoiceRowToObject(values) {
  var obj = {};
  INVOICE_HEADERS.forEach(function (h, i) { obj[h] = values[i]; });
  obj.equipment = safeParseJson(obj.equipmentJson);
  delete obj.equipmentJson;
  obj.startTime = formatTimeValue(obj.startTime);
  obj.endTime = formatTimeValue(obj.endTime);
  obj.courseStartTime = formatTimeValue(obj.courseStartTime);
  obj.courseEndTime = formatTimeValue(obj.courseEndTime);
  obj.eventDate = formatDate(obj.eventDate);
  obj.paymentDueDate = formatDate(obj.paymentDueDate);
  obj.submitDate = formatDate(obj.submitDate);
  obj.createdAt = formatDate(obj.createdAt);
  obj.updatedAt = formatDate(obj.updatedAt);
  obj.paidAt = formatDate(obj.paidAt);
  obj.voidedAt = formatDate(obj.voidedAt);
  return obj;
}

function safeParseJson(s) {
  try { return JSON.parse(s || '{}'); } catch (e) { return {}; }
}

function listInvoices() {
  var sheet = getInvoiceSheet();
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    rows.push(invoiceRowToObject(values[i]));
  }
  rows.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return rows;
}

function findInvoiceRowIndex(sheet, id) {
  if (!id) return -1;
  // コピー&ペースト時に前後へ紛れ込む空白・改行を吸収するため、両辺ともtrimしてから比較する
  // (「リンクが無効」と誤って表示されることがある不具合の再発防止策)。
  var target = String(id).trim();
  if (!target) return -1;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target) return i + 1; // 1始まりの行番号
  }
  return -1;
}

function createInvoice(payload, user) {
  var sheet = getInvoiceSheet();
  var id = Utilities.getUuid();
  var now = new Date();
  var obj = {
    id: id, status: '作成済み',
    courseId: payload.courseId || '', department: payload.department || '',
    eventDate: payload.eventDate || '', startTime: payload.startTime || '', endTime: payload.endTime || '',
    venue: payload.venue || '', topic: payload.topic || '', format: payload.format || '座学',
    directorName: payload.directorName || '', feeExTax: payload.feeExTax || 0,
    travelHandling: payload.travelHandling || '別途支給', paymentMethod: payload.paymentMethod || '銀行振込',
    paymentDueDate: payload.paymentDueDate || '', equipmentJson: JSON.stringify(payload.equipment || {}),
    invoiceRegistered: '', invoiceNumber: '', withholdingType: '個人',
    lecturerName: '', lecturerAddress: '', lecturerContact: '',
    lecturerCompanyName: '', lecturerCompanyAddress: '', lecturerCompanyContact: '', lecturerCompanyEmail: '',
    bankName: '', branchName: '', accountType: '普通', accountNumber: '', accountHolderKana: '',
    photoConsent: '', archiveConsent: '', archiveConsentCondition: '',
    submitDate: '', signerName: '',
    createdBy: user.email, createdAt: now, updatedBy: user.email, updatedAt: now,
    paidAt: '', paidBy: '', voidReason: '', voidedBy: '', voidedAt: '',
    travelFee: 0, lodgingFee: 0, directorMessage: payload.directorMessage || '',
    courseStartTime: payload.courseStartTime || '', courseEndTime: payload.courseEndTime || ''
  };
  sheet.appendRow(safeRow_(INVOICE_HEADERS.map(function (h) { return obj[h]; })));
  appendLog(user.email, 'createInvoice', payload.topic || payload.directorName || '');
  return { id: id };
}

function updateInvoiceDirector(payload, user) {
  var sheet = getInvoiceSheet();
  var rowIndex = findInvoiceRowIndex(sheet, payload.id);
  if (rowIndex === -1) throw new Error('invoice_not_found');
  if (sheet.getRange(rowIndex, invoiceCol('status')).getValue() === '取消') throw new Error('invoice_voided'); // 取消済みは編集不可(画面と同じ規則をサーバーでも確認)
  sheet.getRange(rowIndex, invoiceCol('courseId'), 1, 14).setValues(safeRows_([[
    payload.courseId || '', payload.department || '',
    payload.eventDate || '', payload.startTime || '', payload.endTime || '',
    payload.venue || '', payload.topic || '', payload.format || '座学',
    payload.directorName || '', payload.feeExTax || 0,
    payload.travelHandling || '別途支給', payload.paymentMethod || '銀行振込',
    payload.paymentDueDate || '', JSON.stringify(payload.equipment || {})
  ]]));
  sheet.getRange(rowIndex, invoiceCol('directorMessage')).setValue(safeCell_(payload.directorMessage || ''));
  var courseTimeRange = sheet.getRange(rowIndex, invoiceCol('courseStartTime'), 1, 2);
  courseTimeRange.setNumberFormat('@');
  courseTimeRange.setValues(safeRows_([[payload.courseStartTime || '', payload.courseEndTime || '']]));
  sheet.getRange(rowIndex, invoiceCol('updatedBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, invoiceCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'updateInvoiceDirector', payload.id);
  return { id: payload.id };
}

/** 取消(論理削除)。行自体は消さず、一覧には「取消」として残す(不正防止) */
function voidInvoice(payload, user) {
  var sheet = getInvoiceSheet();
  var rowIndex = findInvoiceRowIndex(sheet, payload.id);
  if (rowIndex === -1) throw new Error('invoice_not_found');
  if (sheet.getRange(rowIndex, invoiceCol('status')).getValue() === '取消') throw new Error('invoice_voided');
  sheet.getRange(rowIndex, invoiceCol('status')).setValue(safeCell_('取消'));
  sheet.getRange(rowIndex, invoiceCol('voidReason')).setValue(safeCell_(payload.reason || ''));
  sheet.getRange(rowIndex, invoiceCol('voidedBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, invoiceCol('voidedAt')).setValue(safeCell_(new Date()));
  sheet.getRange(rowIndex, invoiceCol('updatedBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, invoiceCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'voidInvoice', payload.id + (payload.reason ? (' 理由:' + payload.reason) : ''));
  return { id: payload.id };
}

/**
 * 講師の確定を取り消して「作成済み」に戻す(差し戻し)。講師が入力した内容はそのまま残るため、講師は同じリンクを開くと
 * 前回の内容が入った状態から修正して、もう一度確定できる。差し戻せるのは「確定済み」だけ(支払い済み・取消は不可)。
 * 誰がいつ・どんな理由で差し戻したかは ActivityLog に残す。
 */
function reopenInvoice(payload, user) {
  var sheet = getInvoiceSheet();
  var rowIndex = findInvoiceRowIndex(sheet, payload && payload.id);
  if (rowIndex === -1) throw new Error('invoice_not_found');
  if (sheet.getRange(rowIndex, invoiceCol('status')).getValue() !== '確定済み') throw new Error('invoice_not_confirmed');
  sheet.getRange(rowIndex, invoiceCol('status')).setValue(safeCell_('作成済み'));
  sheet.getRange(rowIndex, invoiceCol('updatedBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, invoiceCol('updatedAt')).setValue(safeCell_(new Date()));
  var reason = String(payload.reason || '').slice(0, 500);
  appendLog(user.email, 'reopenInvoice', payload.id + (reason ? (' 理由:' + reason) : ''));
  return { id: payload.id };
}

function markInvoicePaid(payload, user) {
  var sheet = getInvoiceSheet();
  var rowIndex = findInvoiceRowIndex(sheet, payload.id);
  if (rowIndex === -1) throw new Error('invoice_not_found');
  // 支払い済みにできるのは、講師が内容を確定した(確定済み)請求書だけ(画面と同じ規則をサーバーでも確認)
  if (sheet.getRange(rowIndex, invoiceCol('status')).getValue() !== '確定済み') throw new Error('invoice_not_confirmed');
  sheet.getRange(rowIndex, invoiceCol('status')).setValue(safeCell_('支払い済み'));
  sheet.getRange(rowIndex, invoiceCol('paidAt')).setValue(safeCell_(new Date()));
  sheet.getRange(rowIndex, invoiceCol('paidBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, invoiceCol('updatedBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, invoiceCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'markInvoicePaid', payload.id);
  return { id: payload.id };
}

/** 講師側(未ログイン)からの取得。idはUUIDで推測困難なため認証なしで許可している。 */
var BANK_MASK_DAYS_AFTER_PAID = 3; // 支払い済みから何日後に、講師用リンクの振込先を伏せるか

/** 口座番号の下4桁だけを残して伏せる(例：1234567 → ***4567) */
function maskAccountNumber_(v) {
  var s = String(v || '').replace(/\s/g, '');
  if (!s) return '';
  return (s.length > 4 ? new Array(s.length - 4 + 1).join('*') : '') + s.slice(-4);
}

/** 講師用リンクを、請求書アクセスのある理事がログイン中のブラウザで開いたかどうか(その場合は伏せない) */
function isInvoiceDirector_(idToken) {
  if (!idToken) return false;
  try {
    var u = verifyIdToken(idToken);
    if (!u) return false;
    var a = getAccountInfo_(u.email);
    return !!(a.allowed && a.invoice);
  } catch (e) { return false; }
}

function getInvoicePublic(payload, idToken) {
  var sheet = getInvoiceSheet();
  var requestedId = payload && payload.id;
  var rowIndex = findInvoiceRowIndex(sheet, requestedId);
  if (rowIndex === -1) {
    // 「リンクが無効」表示の原因調査用に、見つからなかったIDをログへ残す
    // (実在しないIDなのか、表記ゆれで一致しなかっただけなのかを後から確認できるようにする)。
    try { appendLog('(講師・未ログイン)', 'getInvoicePublic_not_found', String(requestedId || '')); } catch (eLog) {}
    throw new Error('invoice_not_found');
  }
  var values = sheet.getRange(rowIndex, 1, 1, INVOICE_HEADERS.length).getValues()[0];
  var obj = invoiceRowToObject(values);
  obj.isSkillUpCourse = !!departmentSkillUpMap()[obj.department];
  // 【セキュリティ】支払い済みから3日たったら、講師用リンクでは口座番号を下4桁だけにし、その他の振込先は伏せる
  // (請求書アクセスのある理事がログイン中に開いた場合は、確認のため全て表示する)
  var paidAtRaw = values[invoiceCol('paidAt') - 1];
  var paidAt = paidAtRaw ? new Date(paidAtRaw) : null;
  if (obj.status === '支払い済み' && paidAt && !isNaN(paidAt.getTime()) &&
      (new Date().getTime() - paidAt.getTime()) >= BANK_MASK_DAYS_AFTER_PAID * 86400000 && !isInvoiceDirector_(idToken)) {
    obj.accountNumber = maskAccountNumber_(obj.accountNumber);
    obj.bankName = obj.bankName ? '＊＊＊' : '';
    obj.branchName = obj.branchName ? '＊＊＊' : '';
    obj.accountHolderKana = obj.accountHolderKana ? '＊＊＊' : '';
    obj.bankMasked = true;
  }
  // 【セキュリティ】講師用ページで使わない内部情報(担当理事のメールアドレス・取消理由・支払処理者など)は返さない
  ['createdBy', 'updatedBy', 'paidBy', 'paidAt', 'voidReason', 'voidedBy', 'voidedAt'].forEach(function (k) { delete obj[k]; });
  // 取消済みの請求書では、講師の個人情報・振込先も返さない(リンクが第三者に渡っても見えないようにする)
  if (obj.status === '取消') {
    ['invoiceNumber', 'lecturerName', 'lecturerAddress', 'lecturerContact', 'lecturerCompanyName', 'lecturerCompanyAddress',
      'lecturerCompanyContact', 'lecturerCompanyEmail', 'bankName', 'branchName', 'accountNumber', 'accountHolderKana', 'signerName'
    ].forEach(function (k) { obj[k] = ''; });
  }
  return obj;
}

/** 未ログインの利用者から受け取った文字列を、最大文字数で切り詰めて返す(文字列以外は空文字) */
function publicText_(v, maxLen) {
  if (v === null || v === undefined) return '';
  if (typeof v !== 'string' && typeof v !== 'number') return '';
  return String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, maxLen || 200);
}

/** 講師側(未ログイン)からの確定送信 */
function submitInvoiceLecturer(payload) {
  var sheet = getInvoiceSheet();
  var rowIndex = findInvoiceRowIndex(sheet, payload && payload.id);
  if (rowIndex === -1) throw new Error('invoice_not_found');
  // 【セキュリティ】確定できるのは「作成済み」の請求書だけ。以前はサーバー側で状態を確認しておらず、
  // リンクを知っている人なら確定後・支払い後・取消後でも振込先口座などを書き換えられてしまった。
  var currentStatus = sheet.getRange(rowIndex, invoiceCol('status')).getValue();
  if (currentStatus !== '作成済み') throw new Error('invoice_already_submitted');
  var travelHandling = sheet.getRange(rowIndex, invoiceCol('travelHandling')).getValue() || '別途支給';
  var allowTravel = travelHandling === '別途支給'; // 交通費・宿泊費を講師が記入できるのは「別途支給」の場合だけ
  function fee(v) {
    var n = Math.round(Number(v) || 0);
    if (n < 0) n = 0;
    if (n > 1000000) throw new Error('invalid_amount'); // 桁違いの入力を防ぐ(100万円まで)
    return allowTravel ? n : 0;
  }
  var p = payload;
  sheet.getRange(rowIndex, invoiceCol('invoiceRegistered'), 1, 20).setValues(safeRows_([[
    publicText_(p.invoiceRegistered, 20), publicText_(p.invoiceNumber, 20), publicText_(p.withholdingType, 20) || '個人',
    publicText_(p.lecturerName, 100), publicText_(p.lecturerAddress, 300), publicText_(p.lecturerContact, 100),
    publicText_(p.lecturerCompanyName, 200), publicText_(p.lecturerCompanyAddress, 300), publicText_(p.lecturerCompanyContact, 100), publicText_(p.lecturerCompanyEmail, 200),
    publicText_(p.bankName, 100), publicText_(p.branchName, 100), publicText_(p.accountType, 20) || '普通',
    publicText_(p.accountNumber, 20), publicText_(p.accountHolderKana, 100),
    publicText_(p.photoConsent, 20), publicText_(p.archiveConsent, 20), publicText_(p.archiveConsentCondition, 1000),
    publicText_(p.submitDate, 20), publicText_(p.signerName, 100)
  ]]));
  sheet.getRange(rowIndex, invoiceCol('travelFee')).setValue(safeCell_(fee(p.travelFee)));
  sheet.getRange(rowIndex, invoiceCol('lodgingFee')).setValue(safeCell_(fee(p.lodgingFee)));
  sheet.getRange(rowIndex, invoiceCol('status')).setValue(safeCell_('確定済み'));
  sheet.getRange(rowIndex, invoiceCol('updatedBy')).setValue(safeCell_('(講師・未ログイン)'));
  sheet.getRange(rowIndex, invoiceCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog('(講師・未ログイン)', 'submitInvoiceLecturer', payload.id);
  return { id: payload.id };
}

/**
 * ===== SNS文章作成＆履歴 (Instagram) =====
 * 講座ごとに「募集告知」「開催前リマインド」「開催報告」のInstagram投稿文を、雛形(SnsTemplatesシート)に
 * 講座の情報を差し込んで作成し(差し込みは画面側で行う)、下書き→投稿予定→投稿済みの履歴をSnsPostsシートで管理する。
 * ログインした理事なら誰でも使える。削除は論理削除(status='削除済み')。
 */

function snsPostCol_(fieldName) {
  var idx = SNS_POST_HEADERS.indexOf(fieldName);
  if (idx === -1) throw new Error('invalid_sns_field:' + fieldName);
  return idx + 1;
}

function getSnsPostSheet_() {
  var sheet = getOrCreateManagedSheet_(SHEET_SNS_POSTS, SNS_POST_HEADERS);
  ensureColumnCapacity(sheet, SNS_POST_HEADERS.length);
  // 投稿予定日・投稿日は"2026-10-01"のような文字列のまま保存したいので、列を書式なしテキストに固定する
  sheet.getRange(1, snsPostCol_('scheduledDate'), sheet.getMaxRows(), 2).setNumberFormat('@');
  return sheet;
}

/** 雛形シート。初回作成時(見出し以外に行が無いとき)は、3種類の既定の雛形を入れておく */
function getSnsTemplateSheet_() {
  var sheet = getOrCreateManagedSheet_(SHEET_SNS_TEMPLATES, SNS_TEMPLATE_HEADERS);
  var values = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < values.length; i++) { if (values[i][0]) existing[values[i][0]] = true; }
  SNS_DEFAULT_TEMPLATES.forEach(function (row) {
    if (!existing[row[0]]) sheet.appendRow(safeRow_(row.slice()));
  });
  return sheet;
}

function snsPostRowToObject_(values) {
  var obj = {};
  SNS_POST_HEADERS.forEach(function (h, i) { obj[h] = values[i]; });
  obj.scheduledDate = formatDate(obj.scheduledDate);
  obj.postedDate = formatDate(obj.postedDate);
  obj.createdAt = formatDateTime(obj.createdAt);
  obj.updatedAt = formatDateTime(obj.updatedAt);
  obj.deletedAt = formatDateTime(obj.deletedAt);
  return obj;
}

function listSnsPosts_() {
  var values = getSheetForRead_(SHEET_SNS_POSTS, SNS_POST_HEADERS, getSnsPostSheet_).getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var obj = snsPostRowToObject_(values[i]);
    if (obj.status === '削除済み') continue;
    rows.push(obj);
  }
  return rows;
}

function listSnsTemplates_() {
  var values = getSheetForRead_(SHEET_SNS_TEMPLATES, SNS_TEMPLATE_HEADERS, getSnsTemplateSheet_).getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var type = values[i][0];
    if (!type || map[type]) continue;
    map[type] = { postType: type, body: String(values[i][1] || ''), hashtags: String(values[i][2] || '') };
  }
  return map;
}

/**
 * 提出済みの月次報告書から、講座ごとの参加人数(会場・ZOOM)を拾う(開催報告の「{参加人数}」の差し込み用)。
 * 照合は講座準備進捗確認と同じく講座名(スペース無視)＋開催日(双方にある場合)。
 */
function attendanceFromReports_(courses) {
  var byTitle = {};
  courses.forEach(function (c) {
    var k = normalizeCourseTitle_(c.name);
    if (!k) return;
    if (!byTitle[k]) byTitle[k] = [];
    byTitle[k].push(c);
  });
  var result = {};
  listReports({}).forEach(function (r) {
    if (r.status !== '提出済み') return;
    (r.courses || []).forEach(function (entry) {
      (byTitle[normalizeCourseTitle_(entry.title)] || []).forEach(function (c) {
        if (entry.eventDate && c.eventDate && entry.eventDate !== c.eventDate) return;
        result[c.id] = { venue: Number(entry.venueCount || 0), zoom: Number(entry.zoomCount || 0) };
      });
    });
  });
  return result;
}

/** SNS画面に必要なデータをまとめて返す(要ログイン)。{ posts, templates, postTypes, attendance } */
function listSnsData(payload) {
  var courses = listCourses({ lite: true }); // 出席人数の照合には講座名・開催日・IDだけ使うので、サムネイル画像は読まない
  return {
    posts: listSnsPosts_(),
    templates: listSnsTemplates_(),
    postTypes: SNS_POST_TYPES,
    attendance: attendanceFromReports_(courses)
  };
}

function normalizeYmd_(s) {
  var m = String(s || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

/**
 * 投稿文を保存する(新規・更新)。要ログイン。
 * payload: { id(更新時), courseId, postType, caption, status('下書き'|'投稿予定'|'投稿済み'), scheduledDate, postedDate, postUrl }
 * 投稿済みで投稿日が空なら今日の日付を入れる。戻り値: 保存後の投稿オブジェクト
 */
function saveSnsPost(payload, user) {
  if (!payload) throw new Error('invalid_sns_post');
  var postType = String(payload.postType || '');
  if (SNS_POST_TYPES.indexOf(postType) === -1) throw new Error('invalid_sns_post_type');
  var status = String(payload.status || '下書き');
  if (['下書き', '投稿予定', '投稿済み'].indexOf(status) === -1) throw new Error('invalid_sns_status');
  var caption = String(payload.caption || '');
  if (!caption.trim()) throw new Error('empty_caption');
  var scheduledDate = normalizeYmd_(payload.scheduledDate);
  if (status === '投稿予定' && !scheduledDate) throw new Error('scheduled_date_required');
  var postedDate = normalizeYmd_(payload.postedDate);
  if (status === '投稿済み' && !postedDate) postedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (status !== '投稿済み') postedDate = '';

  var courseName = '';
  if (payload.courseId) {
    var courseSheet = getCourseSheet();
    var cRow = findCourseRowIndex(courseSheet, payload.courseId);
    if (cRow === -1) throw new Error('course_not_found');
    courseName = courseSheet.getRange(cRow, courseCol('名称')).getValue() || '';
  }

  var sheet = getSnsPostSheet_();
  var now = new Date();
  var who = user.name || user.email;
  if (payload.id) {
    var rowIndex = findMaterialRowIndex_(sheet, payload.id); // id(1列目)で行を探す汎用処理として流用
    if (rowIndex === -1) throw new Error('sns_post_not_found');
    var current = snsPostRowToObject_(sheet.getRange(rowIndex, 1, 1, SNS_POST_HEADERS.length).getValues()[0]);
    if (current.status === '削除済み') throw new Error('sns_post_not_found');
    sheet.getRange(rowIndex, snsPostCol_('courseId'), 1, 9).setValues(safeRows_([[
      payload.courseId || '', courseName, 'Instagram', postType, caption,
      status, scheduledDate, postedDate, safeHttpUrl_(payload.postUrl)
    ]]));
    sheet.getRange(rowIndex, snsPostCol_('updatedBy'), 1, 2).setValues(safeRows_([[who, now]]));
    appendLog(user.email, 'saveSnsPost', courseName + '：' + postType + '(' + status + ')');
    return snsPostRowToObject_(sheet.getRange(rowIndex, 1, 1, SNS_POST_HEADERS.length).getValues()[0]);
  }
  var obj = {
    id: Utilities.getUuid(), courseId: payload.courseId || '', courseName: courseName, platform: 'Instagram',
    postType: postType, caption: caption, status: status, scheduledDate: scheduledDate, postedDate: postedDate,
    postUrl: safeHttpUrl_(payload.postUrl),
    createdBy: who, createdAt: now, updatedBy: who, updatedAt: now, deletedBy: '', deletedAt: ''
  };
  sheet.appendRow(safeRow_(SNS_POST_HEADERS.map(function (h) { return obj[h]; })));
  appendLog(user.email, 'saveSnsPost', courseName + '：' + postType + '(' + status + '・新規)');
  return snsPostRowToObject_(sheet.getRange(sheet.getLastRow(), 1, 1, SNS_POST_HEADERS.length).getValues()[0]);
}

/** 投稿文を削除する(論理削除。要ログイン)。payload: { id } */
function deleteSnsPost(payload, user) {
  if (!payload || !payload.id) throw new Error('sns_post_not_found');
  var sheet = getSnsPostSheet_();
  var rowIndex = findMaterialRowIndex_(sheet, payload.id);
  if (rowIndex === -1) throw new Error('sns_post_not_found');
  sheet.getRange(rowIndex, snsPostCol_('status')).setValue(safeCell_('削除済み'));
  sheet.getRange(rowIndex, snsPostCol_('deletedBy'), 1, 2).setValues(safeRows_([[user.email, new Date()]]));
  sheet.getRange(rowIndex, snsPostCol_('updatedBy'), 1, 2).setValues(safeRows_([[user.name || user.email, new Date()]]));
  appendLog(user.email, 'deleteSnsPost', payload.id);
  return { id: payload.id };
}

/**
 * ===== 講座準備進捗確認 (8段階の進捗) =====
 * 全講座の準備状況を「企画決定→講師関係→募集準備→募集開始→申込管理→開催準備→開催→終了処理」の
 * 8段階で横断的に確認・更新する画面(course-progress.html)用。ログインした理事なら誰でも使える。
 * 請求書アクセス権限のない理事にも②③⑥の判定に必要な「請求書のステータスと支払予定日」だけを返し、
 * 金額・講師の口座情報などの請求書の中身は返さない。
 */

/** Settingsシートに指定キーの行が無ければ追加する(既存の値は上書きしない) */
function ensureSettingRow_(sheet, key, value, description) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) return;
  }
  sheet.appendRow(safeRow_([key, value, description]));
}

/** 8段階それぞれの目安(開催日からの日数)を、Settingsシートの値(無ければ既定値)で返す */
function getPrepPhaseSettings_() {
  var map = {};
  try {
    var values = getSheet(SHEET_SETTINGS).getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0]) map[values[i][0]] = values[i][1];
    }
  } catch (e) { /* Settingsシートが無くても既定値で動かす */ }
  return PREP_PHASES.map(function (ph) {
    var v = map[ph.settingKey];
    var n = (v === '' || v === null || v === undefined || isNaN(Number(v))) ? ph.defaultOffset : Math.round(Number(v));
    return { key: ph.key, label: ph.label, offsetDays: n };
  });
}

/** 講座に紐付く請求書のうち、取消以外で最新のもの(取消しか無ければ最新の取消)を返す。ダッシュボードと同じ判定。 */
function findInvoiceForCourse_(invoices, courseId) {
  var linked = invoices.filter(function (inv) { return inv.courseId === courseId; });
  if (!linked.length) return null;
  var active = linked.filter(function (inv) { return inv.status !== '取消'; });
  return (active.length ? active : linked).sort(function (a, b) {
    return String(b.updatedAt || '').localeCompare(String(a.updatedAt || ''));
  })[0];
}

/** 講座名の照合用(前後・途中の半角/全角スペースを無視する) */
function normalizeCourseTitle_(s) {
  return String(s || '').replace(/[\s　]+/g, '');
}

/**
 * 講座準備進捗確認の画面に必要な情報をまとめて返す(要ログイン)。
 * 戻り値: {
 *   courses, invoiceStates: { courseId: { status, paymentDueDate, eventDate } },
 *   materialCounts: { courseId: { director, lecturer } },
 *   reportedCourses: { courseId: 'YYYY-MM' (その講座が記載された提出済み月次報告書の月度) },
 *   checklistItems, statusOptions, phases: [{ key, label, offsetDays }], alertSettings
 * }
 */
/**
 * 支払いが発生しない請書(謝礼0円で、交通費・宿泊費も無い)かどうか。
 * このような請書は、講師が確定した時点で「支払いまで完了」とみなす(振込は不要のため)。
 * ※金額などの中身は返さず、真偽だけを返す(請求書アクセスのない理事にも安全に渡せる)
 */
function invoiceHasNoPayment_(inv) {
  if (!inv) return false;
  if (Math.round(Number(inv.feeExTax) || 0) !== 0) return false;
  if (inv.travelHandling !== '別途支給') return true;
  var confirmed = inv.status === '確定済み' || inv.status === '支払い済み';
  return confirmed && (Number(inv.travelFee) || 0) + (Number(inv.lodgingFee) || 0) === 0;
}

function getPrepOverview(payload) {
  var courses = listCourses({ lite: true });
  var invoices = listInvoices();
  var invoiceStates = {};
  courses.forEach(function (c) {
    var inv = findInvoiceForCourse_(invoices, c.id);
    if (inv) invoiceStates[c.id] = {
      status: inv.status || '', paymentDueDate: inv.paymentDueDate || '', eventDate: inv.eventDate || '',
      paymentMethod: inv.paymentMethod || '', noPayment: invoiceHasNoPayment_(inv)
    };
  });

  var materialCounts = {};
  listMaterials({}).forEach(function (m) {
    if (!materialCounts[m.courseId]) materialCounts[m.courseId] = { director: 0, lecturer: 0 };
    if (m.uploaderRole === '講師') materialCounts[m.courseId].lecturer++; else materialCounts[m.courseId].director++;
  });

  // 月次報告書の「直近の講座開催」は講座IDではなく講座名・開催日で記録されているため、
  // 講座名(スペースを無視)が一致し、双方に開催日があれば開催日も一致するものを「記載済み」とみなす。
  var coursesByTitle = {};
  courses.forEach(function (c) {
    var k = normalizeCourseTitle_(c.name);
    if (!k) return;
    if (!coursesByTitle[k]) coursesByTitle[k] = [];
    coursesByTitle[k].push(c);
  });
  var reportedCourses = {};
  listReports({}).forEach(function (r) {
    if (r.status !== '提出済み') return;
    (r.courses || []).forEach(function (entry) {
      (coursesByTitle[normalizeCourseTitle_(entry.title)] || []).forEach(function (c) {
        if (entry.eventDate && c.eventDate && entry.eventDate !== c.eventDate) return;
        reportedCourses[c.id] = r.month;
      });
    });
  });

  var options = listOptions();
  return {
    // トップページは講座一覧を別に読み込んでいるため、payload.noCourses=true なら講座一覧を返さない(通信量削減)
    courses: (payload && payload.noCourses) ? [] : courses,
    invoiceStates: invoiceStates,
    materialCounts: materialCounts,
    reportedCourses: reportedCourses,
    checklistItems: options['講座準備チェック項目'] || [],
    statusOptions: options['講座ステータス'] || [],
    phases: getPrepPhaseSettings_(),
    alertSettings: getSettings()
  };
}

/**
 * 講座準備進捗確認の画面から、チェック1項目だけを更新する(要ログイン)。
 * payload: { courseId, kind, key, value(true/false) }
 *   kind='lecturerMeetingDone' … ①講師打ち合わせ済み
 *   kind='checklist'           … 備品チェックリストの項目(keyに項目名)
 *   kind='phase'               … 手動でチェックする段階(key='recruitStarted' 募集開始 / 'applicationsClosed' 申込締切・参加者確定)
 *   kind='held'                … ⑤開催済み(ステータスを「実施済み」にする。外すと「実施予定」(選択肢に無ければ「準備中」)に戻す)
 * 更新後、講座進捗ページと同じ基準で prepMaterialsDone(④講座準備資料作成済み)も再計算して保存する。
 * 戻り値: 更新後の講座オブジェクト
 */
function updatePrepItem(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  var sheet = getCourseSheet();
  // phaseJson列は今回追加した列のため、initializeSpreadsheet未実行でも見出しが入るようにしておく
  var phaseHeaderCell = sheet.getRange(1, courseCol('phaseJson'));
  if (!phaseHeaderCell.getValue()) phaseHeaderCell.setValue(safeCell_('phaseJson'));

  var rowIndex = findCourseRowIndex(sheet, payload.courseId);
  if (rowIndex === -1) throw new Error('course_not_found');
  var course = courseRowToObject(sheet.getRange(rowIndex, 1, 1, COURSE_HEADERS.length).getValues()[0]);
  var on = !!payload.value;
  var kind = payload.kind;
  var key = payload.key ? String(payload.key) : '';
  var options = listOptions();

  var checklistState = safeParseJson(course.prepChecklistJson);
  if (kind === 'lecturerMeetingDone') {
    sheet.getRange(rowIndex, courseCol('lecturerMeetingDone')).setValue(safeCell_(on));
  } else if (kind === 'checklist') {
    if (!key) throw new Error('invalid_prep_item');
    checklistState[key] = on;
    sheet.getRange(rowIndex, courseCol('prepChecklistJson')).setValue(safeCell_(JSON.stringify(checklistState)));
  } else if (kind === 'phase') {
    // naForm/naLecturerPdf/naDirectorMaterials：参加者募集フォーム・講師資料PDF・理事資料を「不要」にする設定
    if (['recruitStarted', 'applicationsClosed', 'naForm', 'naLecturerPdf', 'naDirectorMaterials', 'naLecturer', 'naCoursePdf'].indexOf(key) === -1) throw new Error('invalid_prep_item');
    var phaseState = safeParseJson(course.phaseJson);
    phaseState[key] = on ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    sheet.getRange(rowIndex, courseCol('phaseJson')).setValue(safeCell_(JSON.stringify(phaseState)));
  } else if (kind === 'held') {
    var statusOptions = options['講座ステータス'] || [];
    var revertStatus = statusOptions.indexOf('実施予定') !== -1 ? '実施予定' : '準備中';
    sheet.getRange(rowIndex, courseCol('ステータス')).setValue(safeCell_(on ? '実施済み' : revertStatus));
  } else {
    throw new Error('invalid_prep_item');
  }

  var items = options['講座準備チェック項目'] || [];
  var allChecked = items.length > 0 && items.every(function (k) { return !!checklistState[k]; });
  var phaseNow = kind === 'phase' ? phaseState : safeParseJson(course.phaseJson);
  var naPdf = !!(phaseNow.naLecturerPdf || phaseNow.naLecturer);
  sheet.getRange(rowIndex, courseCol('prepMaterialsDone')).setValue(safeCell_(allChecked && (!!course.lecturerMaterialFileId || naPdf)));
  sheet.getRange(rowIndex, courseCol('更新日時')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'updatePrepItem', (course.name || '') + '：' + kind + (key ? '/' + key : '') + '=' + (on ? 'ON' : 'OFF'));
  return courseRowToObject(sheet.getRange(rowIndex, 1, 1, COURSE_HEADERS.length).getValues()[0]);
}

/**
 * ===== 資料庫 (Materials) =====
 * 各講座の資料(教材・配布資料・準備資料など)を、理事側でアップロードして管理する機能。
 * 加えて、請求書の講師用リンクと同じ考え方(講座IDが推測困難なUUIDであることを利用)で、
 * 講師にも専用リンクを発行し、ログイン不要で資料を提出してもらえる。
 * 講師用リンクの画面には、理事がアップロードした資料は表示せず、その講座で講師自身が
 * 提出した資料のみを表示する(理事内部向けの資料を外部の講師に見せてしまわないため)。
 * 削除は請求書の「取消」と同じ考え方で論理削除とし(status='削除済み')、行自体は記録として残す。
 */

function materialCol(fieldName) {
  var idx = MATERIAL_HEADERS.indexOf(fieldName);
  if (idx === -1) throw new Error('invalid_material_field:' + fieldName);
  return idx + 1;
}

function materialRowToObject(values) {
  var obj = {};
  MATERIAL_HEADERS.forEach(function (h, i) { obj[h] = values[i]; });
  obj.createdAt = formatDate(obj.createdAt);
  obj.updatedAt = formatDate(obj.updatedAt);
  obj.deletedAt = formatDate(obj.deletedAt);
  return obj;
}

function getMaterialSheet() {
  var sheet = getOrCreateManagedSheet_(SHEET_MATERIALS, MATERIAL_HEADERS);
  ensureColumnCapacity(sheet, MATERIAL_HEADERS.length);
  return sheet;
}

function findMaterialRowIndex_(sheet, id) {
  if (!id) return -1;
  var target = String(id).trim();
  if (!target) return -1;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target) return i + 1; // 1始まりの行番号
  }
  return -1;
}

function getOrCreateMaterialsFolder_() {
  var name = '理事ポータル_講座資料庫';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

/**
 * 資料の一覧を取得する(理事側・要ログイン)。削除済み(status='削除済み')は含まない。
 * payload: { courseId(任意。指定時はその講座の資料のみ), department(任意) }
 */
function listMaterials(payload) {
  var sheet = getSheetForRead_(SHEET_MATERIALS, MATERIAL_HEADERS, getMaterialSheet);
  var values = sheet.getDataRange().getValues();
  var courseId = payload && payload.courseId;
  var department = payload && payload.department;
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var obj = materialRowToObject(values[i]);
    if (obj.status === '削除済み') continue;
    if (courseId && obj.courseId !== courseId) continue;
    if (department && obj.department !== department) continue;
    rows.push(obj);
  }
  rows.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return rows;
}

/** 理事側から資料をアップロードする(要ログイン)。payload: { courseId, category, note, fileName, mimeType, base64Data } */
function uploadMaterialDirector(payload, user) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  if (!payload.base64Data) throw new Error('no_file');
  var courseSheet = getCourseSheet();
  var courseRowIndex = findCourseRowIndex(courseSheet, payload.courseId);
  if (courseRowIndex === -1) throw new Error('course_not_found');
  var course = courseRowToObject(courseSheet.getRange(courseRowIndex, 1, 1, COURSE_HEADERS.length).getValues()[0]);

  var bytes = Utilities.base64Decode(payload.base64Data);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', payload.fileName || '資料');
  var folder = getOrCreateMaterialsFolder_();
  var file = folder.createFile(blob);

  var sheet = getMaterialSheet();
  var id = Utilities.getUuid();
  var now = new Date();
  var obj = {
    id: id, courseId: payload.courseId, department: course.department || '', courseName: course.name || '',
    category: payload.category || 'その他', uploaderRole: '理事', uploaderName: user.name || user.email, note: payload.note || '',
    fileId: file.getId(), fileName: payload.fileName || file.getName(),
    fileUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view', mimeType: payload.mimeType || '',
    status: '有効', createdAt: now, updatedAt: now, deletedBy: '', deletedAt: ''
  };
  sheet.appendRow(safeRow_(MATERIAL_HEADERS.map(function (h) { return obj[h]; })));
  appendLog(user.email, 'uploadMaterialDirector', (course.name || '') + '：' + (payload.fileName || ''));
  return listMaterials({ courseId: payload.courseId });
}

/** 資料を削除する(論理削除。要ログイン)。payload: { id } */
function deleteMaterial(payload, user) {
  if (!payload || !payload.id) throw new Error('material_not_found');
  var sheet = getMaterialSheet();
  var rowIndex = findMaterialRowIndex_(sheet, payload.id);
  if (rowIndex === -1) throw new Error('material_not_found');

  var fileId = sheet.getRange(rowIndex, materialCol('fileId')).getValue();
  if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }

  var courseId = sheet.getRange(rowIndex, materialCol('courseId')).getValue();
  sheet.getRange(rowIndex, materialCol('status')).setValue(safeCell_('削除済み'));
  sheet.getRange(rowIndex, materialCol('deletedBy')).setValue(safeCell_(user.email));
  sheet.getRange(rowIndex, materialCol('deletedAt')).setValue(safeCell_(new Date()));
  sheet.getRange(rowIndex, materialCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'deleteMaterial', payload.id);
  return listMaterials({ courseId: courseId });
}

/**
 * 講師側(未ログイン)が、資料アップロード用リンクから講座情報と提出済み資料一覧を取得する。
 * courseIdはUUIDで推測困難なため認証なしで許可している(請求書の講師用リンクと同じ考え方)。
 */
function getCourseMaterialsPublic(payload) {
  var courseSheet = getCourseSheet();
  var requestedId = payload && payload.courseId;
  var trimmedId = requestedId ? String(requestedId).trim() : requestedId;
  var courseRowIndex = findCourseRowIndex(courseSheet, trimmedId);
  if (courseRowIndex === -1) {
    try { appendLog('(講師・未ログイン)', 'getCourseMaterialsPublic_not_found', String(requestedId || '')); } catch (eLog) {}
    throw new Error('course_not_found');
  }
  var course = courseRowToObject(courseSheet.getRange(courseRowIndex, 1, 1, COURSE_HEADERS.length).getValues()[0]);
  // 【セキュリティ】講師用ページには、講師が提出した資料の表示に必要な項目だけを返す
  var materials = listMaterials({ courseId: course.id }).filter(function (m) { return m.uploaderRole === '講師'; }).map(function (m) {
    // ファイルはログインできる理事だけが開けるため、講師用ページにはファイルのリンクを返さない
    return { id: m.id, category: m.category, uploaderName: m.uploaderName, note: m.note, fileName: m.fileName,
      mimeType: m.mimeType, createdAt: m.createdAt };
  });
  return {
    course: {
      id: course.id, name: course.name, eventDate: course.eventDate,
      startTime: course.startTime, endTime: course.endTime, venue: course.venue, department: course.department,
      lecturerStartTime: course.lecturerStartTime, lecturerEndTime: course.lecturerEndTime
    },
    materials: materials
  };
}

/** 講師側(未ログイン)からの資料アップロード。payload: { courseId, fileName, mimeType, base64Data, lecturerName, note } */
// 講師がログインなしで提出できるファイルの種類(拡張子)と上限。これ以外は受け付けない
// (リンクが第三者に渡った場合に、プログラム等の不審なファイルを置かれるのを防ぐため)。
var LECTURER_UPLOAD_EXTENSIONS = ['pdf', 'ppt', 'pptx', 'doc', 'docx', 'xls', 'xlsx', 'key', 'pages', 'numbers', 'txt', 'csv',
  'jpg', 'jpeg', 'png', 'gif', 'heic', 'heif', 'webp', 'mp4', 'mov', 'm4v', 'mp3', 'm4a', 'wav', 'zip'];
var LECTURER_UPLOAD_MAX_BYTES = 20 * 1024 * 1024;   // 20MB(画面側の上限と同じ)
var LECTURER_UPLOAD_DAILY_LIMIT = 30;               // 1講座あたり1日30件まで

/** ファイル名から、フォルダ区切りや制御文字など不要な文字を除き、長さを制限する */
function safeFileName_(name, fallback) {
  var n = String(name || '').replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_').replace(/^\.+/, '').trim().slice(0, 150);
  return n || fallback || '資料';
}

function uploadMaterialLecturer(payload) {
  if (!payload || !payload.courseId) throw new Error('course_not_found');
  if (!payload.base64Data || typeof payload.base64Data !== 'string') throw new Error('no_file');
  var courseId = String(payload.courseId).trim();
  var courseSheet = getCourseSheet();
  var courseRowIndex = findCourseRowIndex(courseSheet, courseId);
  if (courseRowIndex === -1) throw new Error('course_not_found');
  var course = courseRowToObject(courseSheet.getRange(courseRowIndex, 1, 1, COURSE_HEADERS.length).getValues()[0]);
  if (course.status === '中止') throw new Error('中止になった講座のため、資料を受け付けていません。');

  var fileName = safeFileName_(payload.fileName, '資料');
  var ext = (fileName.match(/\.([A-Za-z0-9]+)$/) || [])[1];
  if (!ext || LECTURER_UPLOAD_EXTENSIONS.indexOf(ext.toLowerCase()) === -1) {
    throw new Error('このファイルの種類は提出できません。PDF・Word・Excel・PowerPoint・画像・動画などのファイルでお送りください。');
  }
  if (payload.base64Data.length > Math.ceil(LECTURER_UPLOAD_MAX_BYTES / 3) * 4 + 4) {
    throw new Error('ファイルが大きすぎます(20MBまで)。');
  }
  // 1講座あたりの1日の提出件数を制限する(いたずらによる大量アップロード対策)
  var cache = CacheService.getScriptCache();
  var countKey = 'lecup_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '_' + courseId;
  var count = Number(cache.get(countKey) || 0);
  if (count >= LECTURER_UPLOAD_DAILY_LIMIT) throw new Error('本日の提出件数の上限に達しました。明日以降に再度お試しいただくか、担当理事にご連絡ください。');
  cache.put(countKey, String(count + 1), 86400);

  var bytes = Utilities.base64Decode(payload.base64Data);
  if (bytes.length > LECTURER_UPLOAD_MAX_BYTES) throw new Error('ファイルが大きすぎます(20MBまで)。');
  var blob = Utilities.newBlob(bytes, publicText_(payload.mimeType, 100) || 'application/octet-stream', fileName);
  var folder = getOrCreateMaterialsFolder_();
  var file = folder.createFile(blob);

  var sheet = getMaterialSheet();
  var id = Utilities.getUuid();
  var now = new Date();
  var obj = {
    id: id, courseId: courseId, department: course.department || '', courseName: course.name || '',
    category: '講師提出資料', uploaderRole: '講師', uploaderName: publicText_(payload.lecturerName, 100).trim() || '(未記入)', note: publicText_(payload.note, 1000),
    fileId: file.getId(), fileName: fileName,
    fileUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view', mimeType: publicText_(payload.mimeType, 100),
    status: '有効', createdAt: now, updatedAt: now, deletedBy: '', deletedAt: ''
  };
  sheet.appendRow(safeRow_(MATERIAL_HEADERS.map(function (h) { return obj[h]; })));
  appendLog('(講師・未ログイン)', 'uploadMaterialLecturer', (course.name || '') + '：' + fileName);
  return getCourseMaterialsPublic({ courseId: courseId });
}

/** ===== 月次報告書 (Reports) ===== */
// 担当科×月度で1件。ログインしていれば誰でも編集可(担当科ごとの権限制限はしない)。
// エクセルで作成したものを各理事が集めてPDF１本化していた作業を、構造化フォームへの入力＋
// 手動ボタンでの月次PDF一本化に置き換える。
// フォーマット：報告者／直近の講座開催(タイトル・開催日・参加人数：会場◎名・ZOOM○名を複数件)／
//              その他報告内容／課題・改善点等。

function reportCol(fieldName) {
  var idx = REPORT_HEADERS.indexOf(fieldName);
  if (idx === -1) throw new Error('invalid_report_field:' + fieldName);
  return idx + 1;
}

/**
 * name という名前のシートを取得する。存在しなければheaders付きで自動作成し、既に存在する
 * 場合も見出しの並びがheadersと異なれば自動的に再構築する(migrateSheetHeaderOrder_)。
 * これまでReports/BoardMeetingsは「initializeSpreadsheetを手動で再実行し忘れる」と
 * シート自体が存在せず保存操作がsheet_not_foundエラーで失敗してしまい、「報告書が
 * 保存されない」という分かりにくい不具合として現れていたため、通常のアクセス経路でも
 * 自己修復するようにしている。
 */
function getOrCreateManagedSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues(safeRows_([headers]));
    return sheet;
  }
  migrateSheetHeaderOrder_(sheet, headers);
  return sheet;
}

/**
 * 【高速化】読み取り専用の処理で使うシートの取得。
 * シートの準備(見出しの並びの確認・列の書式の固定・既定の行の追加)は、書き込みを伴うため1回あたり0.5〜1秒以上かかる。
 * 以前は一覧を読むたびに毎回行っていたので、準備が済んだら6時間は省略する(保存・削除などの書き込み処理では今まで通り毎回行う)。
 * キャッシュの名前に見出しの内容を含めるので、コードで列を追加した場合は自動で準備し直す。
 */
var SHEET_SETUP_CACHE_SECONDS = 21600; // 6時間
function getSheetForRead_(name, headers, setupFn) {
  var cache = null, key = '';
  try {
    cache = CacheService.getScriptCache();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, name + '\n' + headers.join('\t'), Utilities.Charset.UTF_8);
    key = 'setup_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
  } catch (e) { cache = null; }
  if (cache && cache.get(key)) {
    var existing = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (existing) return existing;
  }
  var sheet = setupFn();
  if (cache) { try { cache.put(key, '1', SHEET_SETUP_CACHE_SECONDS); } catch (e2) {} }
  return sheet;
}

function getReportSheet() {
  var sheet = getOrCreateManagedSheet_(SHEET_REPORTS, REPORT_HEADERS);
  ensureColumnCapacity(sheet, REPORT_HEADERS.length);
  // 月度列("2026-09"等)がスプレッドシートによって日付として自動変換されるのを防ぐため、
  // 書式なしテキストに固定しておく(開始時刻・終了時刻の列と同じ対策)。
  sheet.getRange(1, reportCol('month'), sheet.getMaxRows(), 1).setNumberFormat('@');
  return sheet;
}

function safeParseJsonArray_(s) {
  try {
    var v = JSON.parse(s || '[]');
    return Array.isArray(v) ? v : [];
  } catch (e) {
    return [];
  }
}

/** 構造化フォームの「直近の講座開催」1件分を、安全な形に正規化する(HTMLタグ除去・数値化) */
function sanitizeReportCourseEntry_(c) {
  var title = stripHtmlTags_(String((c && c.title) || '')).trim();
  var eventDate = stripHtmlTags_(String((c && c.eventDate) || '')).trim();
  var venueCount = Math.max(0, Math.round(Number(c && c.venueCount) || 0));
  var zoomCount = Math.max(0, Math.round(Number(c && c.zoomCount) || 0));
  return { title: title, eventDate: eventDate, venueCount: venueCount, zoomCount: zoomCount };
}

/** タイトルが空の行を除いた、安全な講座リストを作る */
function sanitizeReportCourses_(courses) {
  return (Array.isArray(courses) ? courses : [])
    .map(sanitizeReportCourseEntry_)
    .filter(function (c) { return c.title; });
}

function reportRowToObject(values) {
  var obj = {};
  REPORT_HEADERS.forEach(function (h, i) { obj[h] = values[i]; });
  obj.month = formatMonthValue_(obj.month);
  obj.courses = safeParseJsonArray_(obj.coursesJson);
  delete obj.coursesJson;
  obj.attachments = safeParseJsonArray_(obj.attachmentsJson);
  delete obj.attachmentsJson;
  obj.otherContent = obj.otherContent || '';
  obj.issuesContent = obj.issuesContent || '';
  obj.reporterName = obj.reporterName || '';
  obj.updatedAt = formatDateTime(obj.updatedAt);
  obj.submittedAt = formatDateTime(obj.submittedAt);
  return obj;
}

function findReportRowIndex_(sheet, department, month) {
  var values = sheet.getDataRange().getValues();
  var depCol = reportCol('department') - 1, monthCol = reportCol('month') - 1;
  for (var i = 1; i < values.length; i++) {
    if (values[i][depCol] === department && formatMonthValue_(values[i][monthCol]) === String(month)) return i + 1;
  }
  return -1;
}

function findReportRowIndexById_(sheet, id) {
  if (!id) return -1;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === id) return i + 1;
  }
  return -1;
}

/** 報告書一覧を取得する。payload.month(YYYY-MM)を指定すればその月度のみ、省略すれば全件。 */
function listReports(payload) {
  var sheet = getSheetForRead_(SHEET_REPORTS, REPORT_HEADERS, getReportSheet);
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var obj = reportRowToObject(values[i]);
    if (payload && payload.month && obj.month !== payload.month) continue;
    rows.push(obj);
  }
  rows.sort(function (a, b) { return String(a.department).localeCompare(String(b.department), 'ja'); });
  return rows;
}

/**
 * 報告書を保存する(担当科×月度で1件。既存があれば上書き、無ければ新規作成)。
 * payload: {
 *   department, month(YYYY-MM), reporterName,
 *   courses: [{ title, eventDate(YYYY-MM-DD), venueCount, zoomCount }, ...],
 *   otherContent, issuesContent, status('下書き'|'提出済み'、省略可)
 * }
 * 「提出済み」で保存する場合、本文(講座・その他報告内容・課題改善点)が空でも添付ファイルが
 * 1つ以上あれば提出可(ファイルのみでの提出に対応)。すべて空の状態での「提出済み」保存はエラーにする。
 */
function saveReport(payload, user) {
  if (!payload || !payload.department) throw new Error('department_required');
  if (!payload.month) throw new Error('month_required');
  var sheet = getReportSheet();
  var rowIndex = findReportRowIndex_(sheet, payload.department, payload.month);
  var now = new Date();

  var reporterName = stripHtmlTags_(String(payload.reporterName || user.name || user.email)).trim();
  var courses = sanitizeReportCourses_(payload.courses);
  var otherContent = stripHtmlTags_(String(payload.otherContent || '')).trim();
  var issuesContent = stripHtmlTags_(String(payload.issuesContent || '')).trim();
  var requestedStatus = payload.status === '提出済み' ? '提出済み' : (payload.status === '下書き' ? '下書き' : null);

  function isContentEmpty(existingAttachmentsCount) {
    return !courses.length && !otherContent && !issuesContent && !existingAttachmentsCount;
  }

  if (rowIndex === -1) {
    var finalStatusNew = requestedStatus || '下書き';
    if (finalStatusNew === '提出済み' && isContentEmpty(0)) throw new Error('empty_report');
    var id = Utilities.getUuid();
    sheet.appendRow(safeRow_([
      id, payload.department, payload.month, reporterName, JSON.stringify(courses), otherContent, issuesContent,
      '[]', finalStatusNew, user.name || user.email, now, finalStatusNew === '提出済み' ? now : ''
    ]));
    appendLog(user.email, 'saveReport', payload.department + ' ' + payload.month);
    return { id: id, status: finalStatusNew };
  }

  var existing = sheet.getRange(rowIndex, 1, 1, REPORT_HEADERS.length).getValues()[0];
  var finalStatus = requestedStatus || existing[reportCol('status') - 1] || '下書き';
  if (finalStatus === '提出済み') {
    var existingAttachments = safeParseJsonArray_(existing[reportCol('attachmentsJson') - 1]);
    if (isContentEmpty(existingAttachments.length)) throw new Error('empty_report');
  }
  sheet.getRange(rowIndex, reportCol('reporterName')).setValue(safeCell_(reporterName));
  sheet.getRange(rowIndex, reportCol('coursesJson')).setValue(safeCell_(JSON.stringify(courses)));
  sheet.getRange(rowIndex, reportCol('otherContent')).setValue(safeCell_(otherContent));
  sheet.getRange(rowIndex, reportCol('issuesContent')).setValue(safeCell_(issuesContent));
  sheet.getRange(rowIndex, reportCol('status')).setValue(safeCell_(finalStatus));
  sheet.getRange(rowIndex, reportCol('updatedBy')).setValue(safeCell_(user.name || user.email));
  sheet.getRange(rowIndex, reportCol('updatedAt')).setValue(safeCell_(now));
  if (finalStatus === '提出済み' && !existing[reportCol('submittedAt') - 1]) {
    sheet.getRange(rowIndex, reportCol('submittedAt')).setValue(safeCell_(now));
  } else if (finalStatus === '下書き') {
    sheet.getRange(rowIndex, reportCol('submittedAt')).setValue(safeCell_(''));
  }
  appendLog(user.email, 'saveReport', payload.department + ' ' + payload.month);
  return { id: existing[0], status: finalStatus };
}

/** 報告書添付ファイルを保存するGoogleドライブのフォルダを取得(無ければ作成)する */
function getOrCreateReportFolder_() {
  var name = '理事ポータル_報告書添付';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

/** 報告書にファイルを添付する。payload: { reportId, fileName, mimeType, base64Data } */
function uploadReportAttachment(payload, user) {
  if (!payload || !payload.reportId) throw new Error('report_not_found');
  if (!payload.base64Data) throw new Error('no_file');
  var sheet = getReportSheet();
  var rowIndex = findReportRowIndexById_(sheet, payload.reportId);
  if (rowIndex === -1) throw new Error('report_not_found');

  var bytes = Utilities.base64Decode(payload.base64Data);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', payload.fileName || '添付ファイル');
  var folder = getOrCreateReportFolder_();
  var file = folder.createFile(blob);

  var attachments = safeParseJsonArray_(sheet.getRange(rowIndex, reportCol('attachmentsJson')).getValue());
  attachments.push({
    fileId: file.getId(),
    fileName: payload.fileName || file.getName(),
    url: 'https://drive.google.com/file/d/' + file.getId() + '/view'
  });
  sheet.getRange(rowIndex, reportCol('attachmentsJson')).setValue(safeCell_(JSON.stringify(attachments)));
  sheet.getRange(rowIndex, reportCol('updatedBy')).setValue(safeCell_(user.name || user.email));
  sheet.getRange(rowIndex, reportCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'uploadReportAttachment', payload.fileName || '');
  return { attachments: attachments };
}

/**
 * 報告書本文(その他報告内容／課題・改善点等)の文中・文末に挿入する画像をDriveへアップロードする。
 * 添付ファイル一覧(attachmentsJson)とは別枠で扱い、特定の報告書IDには紐付けない
 * (本文中に挿入される [image:ファイルID] というプレースホルダーだけで参照される)。
 * payload: { fileName, mimeType, base64Data }
 */
function uploadReportInlineImage(payload, user) {
  if (!payload || !payload.base64Data) throw new Error('no_file');
  var bytes = Utilities.base64Decode(payload.base64Data);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'image/png', payload.fileName || '画像.png');
  var folder = getOrCreateReportFolder_();
  var file = folder.createFile(blob);
  appendLog(user.email, 'uploadReportInlineImage', payload.fileName || '');
  return { fileId: file.getId(), fileName: file.getName(), url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
}

/** 報告書の添付ファイルを削除する。payload: { reportId, fileId } */
function deleteReportAttachment(payload, user) {
  if (!payload || !payload.reportId || !payload.fileId) throw new Error('report_not_found');
  var sheet = getReportSheet();
  var rowIndex = findReportRowIndexById_(sheet, payload.reportId);
  if (rowIndex === -1) throw new Error('report_not_found');

  var attachments = safeParseJsonArray_(sheet.getRange(rowIndex, reportCol('attachmentsJson')).getValue());
  var remaining = attachments.filter(function (a) { return a.fileId !== payload.fileId; });
  if (remaining.length !== attachments.length) {
    try { DriveApp.getFileById(payload.fileId).setTrashed(true); } catch (e) {}
  }
  sheet.getRange(rowIndex, reportCol('attachmentsJson')).setValue(safeCell_(JSON.stringify(remaining)));
  sheet.getRange(rowIndex, reportCol('updatedBy')).setValue(safeCell_(user.name || user.email));
  sheet.getRange(rowIndex, reportCol('updatedAt')).setValue(safeCell_(new Date()));
  appendLog(user.email, 'deleteReportAttachment', payload.fileId);
  return { attachments: remaining };
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** ファイルが指定フォルダの直下にあるかどうか */
function isInFolder_(file, folder) {
  var parents = file.getParents();
  var target = folder.getId();
  while (parents.hasNext()) { if (parents.next().getId() === target) return true; }
  return false;
}

function isImageFileName_(name) {
  return /\.(png|jpe?g|gif|webp|bmp)$/i.test(String(name || ''));
}

/** DriveのファイルIDから、PDF化(HTML→Googleドキュメント変換)の中に埋め込める
 * base64データURI形式の<img>タグを組み立てる(Drive上のURLをそのまま<img src>に
 * 指定すると変換時にアクセスできず失敗することがあるため、base64で直接埋め込む)。 */
function imageTagForFileId_(fileId, styleAttr) {
  try {
    var file = DriveApp.getFileById(fileId);
    // 【セキュリティ】報告書の添付フォルダにある画像だけを埋め込む。以前は本文に任意のファイルIDを書くと、
    // スクリプト所有者のドライブにある無関係なファイルまでPDFに埋め込めてしまった。
    if (!isInFolder_(file, getOrCreateReportFolder_()) || !/^image\//.test(file.getMimeType())) {
      return '<span style="color:#c5221f;">［画像を読み込めませんでした］</span>';
    }
    var blob = file.getBlob();
    var mime = blob.getContentType() || 'image/png';
    var b64 = Utilities.base64Encode(blob.getBytes());
    return '<img src="data:' + mime + ';base64,' + b64 + '" style="' + (styleAttr || 'max-width:100%;') + '">';
  } catch (e) {
    return '<span style="color:#c5221f;">［画像を読み込めませんでした］</span>';
  }
}

/**
 * 報告書本文(その他報告内容／課題・改善点等)のテキストを、改行を<br>に変換しつつHTMLエスケープし、
 * 本文中に埋め込まれた [image:ファイルID] というプレースホルダーを、実際の画像(base64埋め込み)に
 * 置き換える(reports.html側の「画像を挿入」ボタンで挿入されるプレースホルダー形式)。
 */
function replaceInlineImagePlaceholders_(text) {
  var raw = String(text || '');
  var IMG_TOKEN_RE = /\[image:([a-zA-Z0-9_-]+)\]/g;
  var htmlParts = [];
  var lastIndex = 0;
  var match;
  while ((match = IMG_TOKEN_RE.exec(raw)) !== null) {
    htmlParts.push(escapeHtml_(raw.slice(lastIndex, match.index)).replace(/\n/g, '<br>'));
    // 文章→画像→文章の縦並びレイアウト。本文幅を超えない・縦横比維持(height:auto)・中央配置・
    // 前後14pt程度の余白。幅を明示しないため、小さい画像を不必要に拡大することもない。
    // 【2026-09 最終調整】画像はあくまで報告内容を補足する資料であり主役ではないため、
    // 本文幅の68%程度までさらに縮小する(前後の余白バランスは変更しない)。
    htmlParts.push('<div style="margin:14pt 0; text-align:center;">' + imageTagForFileId_(match[1], 'max-width:68%; max-height:300pt; height:auto;') + '</div>');
    lastIndex = IMG_TOKEN_RE.lastIndex;
  }
  htmlParts.push(escapeHtml_(raw.slice(lastIndex)).replace(/\n/g, '<br>'));
  return htmlParts.join('');
}

/**
 * 指定した月度の「提出済み」報告書をすべてまとめて1つのPDFにする(手動ボタンで実行)。
 * HTML→Googleドキュメント変換(convert:true)→PDFエクスポートの流れで作成する
 * (ocrCoursePdfで使っているHTML→ドキュメント変換の仕組みを流用。OCRは行わない)。
 * payload: { month(YYYY-MM) }
 */
function compileMonthlyReportsPdf(payload, user) {
  var month = payload && payload.month;
  if (!month) throw new Error('month_required');
  var reports = listReports({ month: month }).filter(function (r) { return r.status === '提出済み'; });
  if (!reports.length) throw new Error('no_submitted_reports');
  reports.sort(function (a, b) { return String(a.department).localeCompare(String(b.department), 'ja'); });

  var monthLabelMatch = String(month).match(/^(\d{4})-(\d{2})/);
  var monthLabel = monthLabelMatch ? (Number(monthLabelMatch[1]) + '年' + Number(monthLabelMatch[2]) + '月') : escapeHtml_(month);
  var monthDotLabel = monthLabelMatch ? (Number(monthLabelMatch[1]) + '.' + monthLabelMatch[2]) : escapeHtml_(month);

  // 【2026-09 レイアウト再設計】理事会で紙/PDFとして読みやすい、情報密度の高い月次資料を目指す。
  // 「Web画面をPDF化した資料」ではなく「印刷物として設計された月次理事会資料」とするため、
  // (1)担当科ごとの強制改ページを廃止し複数科を同一ページに連続掲載する、(2)ページ上部・各見出し・
  // 各セクションの余白を大幅に圧縮する、(3)配色は本文・補助文字・罫線・薄い背景・アクセントの5色に
  // 限定し、アクセント色(青)は罫線・見出し番号など最小限の面積にとどめる、という方針に基づく。
  // Googleドキュメントへの変換(HTML→convert:true)を経てPDF化されるため、flexbox/grid/box-shadow/
  // border-radius/absolute配置/擬似要素などは使わず、table・border・padding・margin・
  // background-color・font-size・line-heightなど再現性の高い表現のみを用いる。
  // 不自然な改ページ(担当科名だけ/見出しだけ/表ヘッダーだけが取り残される)を避けるため、
  // 各セクション(マストヘッド・講座表・その他報告・課題ブロック・添付一覧)単位でのみ
  // page-break-inside:avoidを使う(セクションの途中では使わず、段落の分割自体は妨げない)。
  var INK_PRIMARY = '#222222';   // 見出し・担当科名などの濃い文字
  var INK_SECONDARY = '#666666'; // 補助文字(報告者名・日付など)
  var INK_MUTED = '#777777';     // さらに弱い補助文字(該当講座なし等の案内文)
  var BODY_TEXT = '#333333';     // 本文(その他報告・課題改善点の段落)の可読性向上用の文字色
  var HAIRLINE = '#DADDE1';      // 罫線
  var TABLE_HEAD_BG = '#F4F5F6'; // 講座表ヘッダー背景
  var ISSUE_BG = '#F8F9FA';      // 課題・改善点の背景(講座表ヘッダーよりさらに軽い階調)
  var ACCENT = '#2A78D6';        // アクセント(見出し番号のみに使用)
  var DEPT_DIVIDER = '#8AB7EC';  // 担当科の区切り線(ACCENTより淡く、「目立たせる」より「自然に区切る」ため)

  /** 「01　直近の講座開催」のような番号付きセクション見出し(番号のみアクセント色、下線1本) */
  function sectionHeading_(num, label) {
    return '<table width="100%" style="border-collapse:collapse; margin:14pt 0 6pt;"><tr>' +
      '<td style="font-size:11.5pt; font-weight:bold; color:' + INK_PRIMARY + '; border-bottom:1px solid ' + HAIRLINE + '; padding-bottom:3pt;">' +
        '<span style="color:' + ACCENT + ';">' + num + '</span>　' + label +
      '</td>' +
    '</tr></table>';
  }

  var htmlParts = [];
  // ページ上部：法人名(8.5pt)／MONTHLY REPORT＋年月(8pt)／タイトル(19pt)／細いアクセント罫線、の4段のみ。
  htmlParts.push(
    '<div style="margin-bottom:14pt;">' +
      '<div style="font-size:8.5pt; color:' + INK_SECONDARY + ';">NPO法人 日本繊維商品めんてなんす研究会</div>' +
      '<table width="100%" style="border-collapse:collapse; margin:3pt 0 1pt;"><tr>' +
        '<td style="font-size:8pt; letter-spacing:1.5px; color:' + INK_SECONDARY + '; font-weight:bold;">MONTHLY REPORT</td>' +
        '<td align="right" style="font-size:8.5pt; color:' + INK_SECONDARY + ';">' + monthDotLabel + '</td>' +
      '</tr></table>' +
      '<div style="font-size:19pt; font-weight:bold; color:' + INK_PRIMARY + '; margin-top:1pt;">' + monthLabel + '度 月次報告書</div>' +
      '<div style="border-top:1.5px solid ' + ACCENT + '; margin-top:7pt;"></div>' +
    '</div>'
  );

  reports.forEach(function (r, idx) {
    // 強制改ページは原則廃止。担当科の切れ目は「上下余白+2pxアクセント罫線」のみで表現し、
    // 複数の担当科が同一ページに連続して掲載されるようにする(自然にページが尽きた場合のみ
    // ブラウザ/Googleドキュメント側の通常改ページに委ねる)。
    var deptTopMargin = idx === 0 ? '8pt' : '28pt';
    var deptRule = idx === 0 ? '' : ('<div style="border-top:1px solid ' + DEPT_DIVIDER + '; margin-bottom:8pt;"></div>');

    // 担当科名・報告者(マストヘッド)。見出しだけが取り残されないよう次セクションとの分離を避ける。
    htmlParts.push(
      '<div style="page-break-inside:avoid; page-break-after:avoid; margin-top:' + deptTopMargin + ';">' +
        deptRule +
        '<div style="font-size:16pt; font-weight:bold; color:' + INK_PRIMARY + ';">' + escapeHtml_(r.department) + '</div>' +
        '<div style="font-size:9.5pt; color:' + INK_SECONDARY + '; margin-top:2pt;">報告者　' + escapeHtml_(r.reporterName || '(未記入)') + '</div>' +
      '</div>'
    );

    // 01 直近の講座開催(開催日／講座名／会場／ZOOM／合計の5列。数字は右揃え、色分けバッジは廃止)
    htmlParts.push('<div style="page-break-inside:avoid;">');
    htmlParts.push(sectionHeading_('01', '直近の講座開催'));
    if (r.courses && r.courses.length) {
      // 【2026-09 耐久テスト対応】長い講座名が入っても「会場/ZOOM/合計」のヘッダーが縦に
      // 折り返されないよう、列幅をtable-layout:fixed + width指定(%とwidth属性の併用。
      // Google Docs変換での再現性を優先しwidth属性も併記)で固定する。講座名列のみ
      // 可変幅の中で折り返しを許可し、他4列の幅は講座名の長さに影響されない。
      htmlParts.push('<table width="100%" style="border-collapse:collapse; margin-bottom:4pt; table-layout:fixed;">');
      htmlParts.push(
        '<tr>' +
          '<td width="15%" style="width:15%; background:' + TABLE_HEAD_BG + '; border:1px solid ' + HAIRLINE + '; padding:6pt 7pt; font-size:9.5pt; font-weight:bold; white-space:nowrap;">開催日</td>' +
          '<td width="55%" style="width:55%; background:' + TABLE_HEAD_BG + '; border:1px solid ' + HAIRLINE + '; padding:6pt 7pt; font-size:9.5pt; font-weight:bold;">講座名</td>' +
          '<td width="10%" style="width:10%; background:' + TABLE_HEAD_BG + '; border:1px solid ' + HAIRLINE + '; padding:6pt 7pt; font-size:9.5pt; font-weight:bold; text-align:right; white-space:nowrap;">会場</td>' +
          '<td width="10%" style="width:10%; background:' + TABLE_HEAD_BG + '; border:1px solid ' + HAIRLINE + '; padding:6pt 7pt; font-size:9.5pt; font-weight:bold; text-align:right; white-space:nowrap;">ZOOM</td>' +
          '<td width="10%" style="width:10%; background:' + TABLE_HEAD_BG + '; border:1px solid ' + HAIRLINE + '; padding:6pt 7pt; font-size:9.5pt; font-weight:bold; text-align:right; white-space:nowrap;">合計</td>' +
        '</tr>'
      );
      r.courses.forEach(function (c) {
        var venue = Number(c.venueCount || 0), zoom = Number(c.zoomCount || 0);
        htmlParts.push(
          '<tr>' +
            '<td width="15%" style="width:15%; border:1px solid ' + HAIRLINE + '; padding:7pt; font-size:10pt; white-space:nowrap; color:' + INK_SECONDARY + ';">' + escapeHtml_(c.eventDate || '-') + '</td>' +
            '<td width="55%" style="width:55%; border:1px solid ' + HAIRLINE + '; padding:7pt; font-size:10pt; color:' + INK_PRIMARY + '; word-break:break-word; overflow-wrap:anywhere;">' + escapeHtml_(c.title) + '</td>' +
            '<td width="10%" style="width:10%; border:1px solid ' + HAIRLINE + '; padding:7pt; font-size:10pt; text-align:right;">' + venue + '</td>' +
            '<td width="10%" style="width:10%; border:1px solid ' + HAIRLINE + '; padding:7pt; font-size:10pt; text-align:right;">' + zoom + '</td>' +
            '<td width="10%" style="width:10%; border:1px solid ' + HAIRLINE + '; padding:7pt; font-size:10pt; text-align:right; font-weight:bold;">' + (venue + zoom) + '</td>' +
          '</tr>'
        );
      });
      htmlParts.push('</table>');
    } else {
      htmlParts.push('<div style="font-size:9.5pt; color:' + INK_MUTED + '; margin-bottom:4pt;">該当する講座開催はありません。</div>');
    }
    htmlParts.push('</div>');

    // 02 その他報告(独立した情報ブロックとして、段落・改行を保ったまま表示)
    htmlParts.push('<div style="page-break-inside:avoid;">');
    htmlParts.push(sectionHeading_('02', 'その他報告'));
    htmlParts.push(
      '<div style="font-size:10.5pt; line-height:1.6; color:' + BODY_TEXT + '; margin:0 0 4pt;">' +
        (r.otherContent ? replaceInlineImagePlaceholders_(r.otherContent) : '<span style="color:' + INK_SECONDARY + ';">(記載なし)</span>') +
      '</div>'
    );
    htmlParts.push('</div>');

    // 03 課題・改善点(警告色は使わず、薄い背景+左3pxアクセント罫線の通常項目として扱う)。
    // 【2026-09 最終調整】内容が空の場合はボックス自体を出さず、「特記事項なし」という
    // 控えめな案内文のみを表示する(背景・左罫線は内容がある場合のみ使用)。
    htmlParts.push('<div style="page-break-inside:avoid;">');
    htmlParts.push(sectionHeading_('03', '課題・改善点'));
    if (r.issuesContent) {
      htmlParts.push(
        '<table width="100%" style="border-collapse:collapse; margin:0 0 4pt;"><tr>' +
          '<td style="background:' + ISSUE_BG + '; border-left:3px solid ' + ACCENT + '; padding:8pt 10pt; font-size:10.5pt; line-height:1.6; color:' + BODY_TEXT + ';">' +
            replaceInlineImagePlaceholders_(r.issuesContent) +
          '</td>' +
        '</tr></table>'
      );
    } else {
      htmlParts.push('<div style="font-size:9.5pt; color:' + INK_MUTED + '; margin:0 0 4pt;">特記事項なし</div>');
    }
    htmlParts.push('</div>');

    // 04 添付資料。画像は本文の後に改ページしてそのまま1ページの画像として掲載し、
    // 画像以外(PDF等)はリンクとして一覧に載せる(添付ファイルの実ページをそのまま
    // 結合するには本来のバイナリ結合が必要でApps Scriptの標準機能だけでは難しいため、
    // 画像は掲載・それ以外はリンク、という扱いにしている。この方針自体は今回変更しない)。
    var imageAttachments = (r.attachments || []).filter(function (a) { return isImageFileName_(a.fileName); });
    var otherAttachments = (r.attachments || []).filter(function (a) { return !isImageFileName_(a.fileName); });
    if (otherAttachments.length) {
      htmlParts.push('<div style="page-break-inside:avoid;">');
      htmlParts.push(sectionHeading_('04', '添付資料'));
      otherAttachments.forEach(function (a) {
        htmlParts.push(
          '<div style="padding:6pt 0; border-bottom:1px solid ' + HAIRLINE + '; font-size:10pt;">' +
            '<span style="color:' + INK_PRIMARY + '; word-break:break-word; overflow-wrap:anywhere;">' + escapeHtml_(a.fileName) + '</span>　' +
            '<a href="' + escapeHtml_(a.url) + '" style="color:' + ACCENT + '; font-size:9.5pt; text-decoration:underline; white-space:nowrap;">ファイルを開く</a>' +
          '</div>'
        );
      });
      htmlParts.push('</div>');
    }
    imageAttachments.forEach(function (a) {
      htmlParts.push('<div style="page-break-before:always;"></div>');
      htmlParts.push(
        '<div style="font-size:8.5pt; color:' + INK_SECONDARY + '; margin-bottom:10pt;">' +
          '添付資料　' + escapeHtml_(r.department) + '　' + escapeHtml_(a.fileName) +
        '</div>'
      );
      htmlParts.push('<div style="text-align:center;">' + imageTagForFileId_(a.fileId, 'max-width:100%; max-height:23cm; height:auto;') + '</div>');
    });
  });
  var combinedHtml = '<html><body style="font-family: \'Noto Sans JP\', Arial, sans-serif; color:' + INK_PRIMARY + ';">' + htmlParts.join('\n') + '</body></html>';

  var tempDocId = null;
  try {
    var blob = Utilities.newBlob(combinedHtml, 'text/html', month + '_報告書_temp.html');
    var resource = { title: month + '月度報告書' };
    var converted = Drive.Files.insert(resource, blob, { convert: true });
    tempDocId = converted.id;
    var pdfBlob = DriveApp.getFileById(tempDocId).getAs('application/pdf');
    var folder = getOrCreateReportFolder_();
    var pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName(month + '月度報告書.pdf');
    appendLog(user.email, 'compileMonthlyReportsPdf', month);
    return { fileId: pdfFile.getId(), url: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view', departmentCount: reports.length };
  } catch (err) {
    // 【要設定】Drive API(高度なサービス)が未設定の場合もここでエラーになる(ocrCoursePdfと同様)。
    throw new Error('compile_failed:' + (err && err.message ? err.message : String(err)));
  } finally {
    if (tempDocId) { try { Drive.Files.remove(tempDocId); } catch (e) {} }
  }
}

/** "YYYY-MM-DD..."形式の文字列(またはDate)をDateへ変換する。解釈できなければnull */
function parseDateStr_(s) {
  if (!s) return null;
  if (Object.prototype.toString.call(s) === '[object Date]') return s;
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/**
 * 指定した月度の報告書提出期限を求める。
 * 「その月度の月初以降で最も早い役員会の開催日」を、その月度の報告をレビューする役員会とみなし、
 * その2日前を提出期限とする(役員会の開催日は「役員会の設定」画面であらかじめ登録しておく)。
 * 該当する役員会が登録されていない場合はnullを返す。
 */
function findReportDeadline_(month) {
  var m = String(month || '').match(/^(\d{4})-(\d{2})/);
  if (!m) return null;
  var monthStart = new Date(Number(m[1]), Number(m[2]) - 1, 1);
  var meetings = listBoardMeetings();
  var best = null;
  meetings.forEach(function (mtg) {
    var d = parseDateStr_(mtg.meetingDate);
    if (!d) return;
    if (d.getTime() >= monthStart.getTime()) {
      if (!best || d.getTime() < best.getTime()) best = d;
    }
  });
  if (!best) return null;
  var deadline = new Date(best.getTime() - 2 * 86400000);
  return {
    meetingDate: Utilities.formatDate(best, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    deadline: Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

/**
 * ダッシュボード用：指定した月度について、各担当科の報告書提出状況を返す。
 * payload: { month(YYYY-MM) }
 * 戻り値: { month, meetingDate, deadline, pastDeadline, items:[{department,submitted,reportId}], missingDepartments }
 */
function getReportSubmissionStatus(payload) {
  var month = payload && payload.month;
  if (!month) throw new Error('month_required');
  var departments = listDepartments();
  var reports = listReports({ month: month });
  var submittedByDept = {};
  reports.forEach(function (r) { if (r.status === '提出済み') submittedByDept[r.department] = r; });
  var deadlineInfo = findReportDeadline_(month);
  var pastDeadline = false;
  if (deadlineInfo) {
    var deadlineDate = parseDateStr_(deadlineInfo.deadline);
    if (deadlineDate) pastDeadline = new Date().getTime() > (deadlineDate.getTime() + 86400000 - 1);
  }
  var items = departments.map(function (d) {
    var r = submittedByDept[d.name];
    return { department: d.name, submitted: !!r, reportId: r ? r.id : '' };
  });
  return {
    month: month,
    meetingDate: deadlineInfo ? deadlineInfo.meetingDate : '',
    deadline: deadlineInfo ? deadlineInfo.deadline : '',
    pastDeadline: pastDeadline,
    items: items,
    missingDepartments: items.filter(function (it) { return !it.submitted; }).map(function (it) { return it.department; })
  };
}

/** ===== 役員会の設定 (BoardMeetings) ===== */
// 役員会の開催日程を、年度(4月～翌3月)単位でまとめて登録しておく設定画面用。
// 月次報告書の提出期限・未提出アラートの基準に使うほか、各開催日をGoogleカレンダー(CALENDAR_ID)へ
// 自動登録する(共有されていない等の理由で失敗しても、設定の保存自体は継続する)。

/** 年度内の12ヶ月分の定義(fyMonthIndex 1=4月 ～ 12=翌3月)を返す */
function fiscalYearMonths_() {
  var labels = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月(翌年)', '2月(翌年)', '3月(翌年)'];
  var calMonths = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
  var yearOffsets = [0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1];
  var months = [];
  for (var i = 0; i < 12; i++) {
    months.push({ fyMonthIndex: i + 1, label: labels[i], calMonth: calMonths[i], yearOffset: yearOffsets[i] });
  }
  return months;
}

function boardMeetingCol(fieldName) {
  var idx = BOARD_MEETING_HEADERS.indexOf(fieldName);
  if (idx === -1) throw new Error('invalid_board_meeting_field:' + fieldName);
  return idx + 1;
}

function boardMeetingRowToObject(values) {
  var obj = {};
  BOARD_MEETING_HEADERS.forEach(function (h, i) { obj[h] = values[i]; });
  obj.meetingDate = formatDate(obj.meetingDate);
  obj.updatedAt = formatDateTime(obj.updatedAt);
  return obj;
}

function getBoardMeetingSheet() {
  var sheet = getOrCreateManagedSheet_(SHEET_BOARD_MEETINGS, BOARD_MEETING_HEADERS);
  ensureColumnCapacity(sheet, BOARD_MEETING_HEADERS.length);
  return sheet;
}

/** 役員会の開催日を、登録済みの全年度分まとめて一覧取得する(報告書の提出期限計算に使用) */
function listBoardMeetings() {
  var sheet = getSheetForRead_(SHEET_BOARD_MEETINGS, BOARD_MEETING_HEADERS, getBoardMeetingSheet);
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    rows.push(boardMeetingRowToObject(values[i]));
  }
  rows.sort(function (a, b) { return String(a.meetingDate).localeCompare(String(b.meetingDate)); });
  return rows;
}

/** 指定した年度(fiscalYear、例："2026"＝2026年4月～2027年3月)の12ヶ月分の設定状況を取得する */
function getBoardMeetingsForYear(payload) {
  var fiscalYear = String((payload && payload.fiscalYear) || '').trim();
  if (!fiscalYear) throw new Error('fiscal_year_required');
  var sheet = getSheetForRead_(SHEET_BOARD_MEETINGS, BOARD_MEETING_HEADERS, getBoardMeetingSheet);
  var values = sheet.getDataRange().getValues();
  var byMonth = {};
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var obj = boardMeetingRowToObject(values[i]);
    if (String(obj.fiscalYear) === fiscalYear) byMonth[Number(obj.fyMonthIndex)] = obj;
  }
  var months = fiscalYearMonths_().map(function (def) {
    var existing = byMonth[def.fyMonthIndex];
    return {
      fyMonthIndex: def.fyMonthIndex, label: def.label,
      meetingDate: existing ? existing.meetingDate : '',
      note: existing ? existing.note : ''
    };
  });
  return { fiscalYear: fiscalYear, months: months };
}

/**
 * 役員会の開催日をGoogleカレンダーへ自動登録・更新・削除する(終日予定)。
 * カレンダーへの書き込み権限が無い等で失敗しても、例外を投げずに元のeventIdをそのまま返す
 * (講座のカレンダー連携と同じ方針)。meetingDateが空の場合は既存の予定を削除して''を返す。
 */
function syncBoardMeetingCalendarEvent_(meetingDate, title, description, existingEventId) {
  clearCalendarCache_();
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) return existingEventId || '';

    if (!meetingDate) {
      if (existingEventId) {
        try {
          var evToRemove = cal.getEventById(existingEventId);
          if (evToRemove) evToRemove.deleteEvent();
        } catch (eRemove) {}
      }
      return '';
    }

    var m = String(meetingDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return existingEventId || '';
    var dateObj = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

    var event = null;
    if (existingEventId) {
      try { event = cal.getEventById(existingEventId); } catch (eGet) { event = null; }
    }
    if (event) {
      event.setAllDayDate(dateObj);
      event.setTitle(title);
      event.setDescription(description);
      return event.getId();
    }
    var created = cal.createAllDayEvent(title, dateObj, { description: description });
    return created.getId();
  } catch (err) {
    return existingEventId || '';
  }
}

/**
 * 指定した年度の役員会開催日を、12ヶ月分まとめて保存する(空欄の月は削除扱い)。
 * payload: { fiscalYear, months: [ { fyMonthIndex, meetingDate(YYYY-MM-DD or ''), note }, ... ] }
 * 保存と同時に、各開催日をGoogleカレンダー(CALENDAR_ID)へ自動登録・更新・削除する。
 */
function saveBoardMeetingsForYear(payload, user) {
  var fiscalYear = String((payload && payload.fiscalYear) || '').trim();
  if (!fiscalYear) throw new Error('fiscal_year_required');
  var monthsInput = (payload && payload.months) || [];

  var sheet = getBoardMeetingSheet();
  var values = sheet.getDataRange().getValues();

  // 既存行を fiscalYear + fyMonthIndex でインデックス化(行番号を保持しておく)
  var existingByMonth = {};
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var fy = String(values[i][boardMeetingCol('fiscalYear') - 1]);
    if (fy !== fiscalYear) continue;
    var mi = Number(values[i][boardMeetingCol('fyMonthIndex') - 1]);
    existingByMonth[mi] = { rowIndex: i + 1, values: values[i] };
  }

  var resultMonths = [];
  fiscalYearMonths_().forEach(function (def) {
    var input = monthsInput.filter(function (m) { return Number(m.fyMonthIndex) === def.fyMonthIndex; })[0];
    var meetingDate = (input && input.meetingDate) ? String(input.meetingDate).trim() : '';
    var note = (input && input.note) || '';
    var existing = existingByMonth[def.fyMonthIndex];
    var title = fiscalYear + '年度 役員会（' + def.label + '）';

    if (!meetingDate) {
      // 空欄で保存された月は「開催日なし」として扱い、既存行があれば削除する(カレンダーの予定も削除)
      if (existing) {
        var existingEventId = existing.values[boardMeetingCol('calendarEventId') - 1];
        syncBoardMeetingCalendarEvent_('', title, note, existingEventId);
        sheet.deleteRow(existing.rowIndex);
        // deleteRowで後続行が1行ずつ繰り上がるため、この年度内で保持している他月の行番号を補正する
        Object.keys(existingByMonth).forEach(function (k) {
          if (existingByMonth[k].rowIndex > existing.rowIndex) existingByMonth[k].rowIndex -= 1;
        });
      }
      resultMonths.push({ fyMonthIndex: def.fyMonthIndex, meetingDate: '', note: '' });
      return;
    }

    if (existing) {
      var oldEventId = existing.values[boardMeetingCol('calendarEventId') - 1];
      var newEventId = syncBoardMeetingCalendarEvent_(meetingDate, title, note, oldEventId);
      sheet.getRange(existing.rowIndex, boardMeetingCol('meetingDate')).setValue(safeCell_(meetingDate));
      sheet.getRange(existing.rowIndex, boardMeetingCol('note')).setValue(safeCell_(note));
      sheet.getRange(existing.rowIndex, boardMeetingCol('calendarEventId')).setValue(safeCell_(newEventId));
      sheet.getRange(existing.rowIndex, boardMeetingCol('updatedBy')).setValue(safeCell_(user.email));
      sheet.getRange(existing.rowIndex, boardMeetingCol('updatedAt')).setValue(safeCell_(new Date()));
    } else {
      var newEventId2 = syncBoardMeetingCalendarEvent_(meetingDate, title, note, '');
      var id = Utilities.getUuid();
      sheet.appendRow(safeRow_([id, fiscalYear, def.fyMonthIndex, meetingDate, note, newEventId2, user.email, new Date()]));
    }
    resultMonths.push({ fyMonthIndex: def.fyMonthIndex, meetingDate: meetingDate, note: note });
  });

  appendLog(user.email, 'saveBoardMeetingsForYear', fiscalYear);
  return { fiscalYear: fiscalYear, months: resultMonths };
}

/**
 * ログイン(または請求書機能)を断ったときに、日時・ログインしようとしたメールアドレス・Googleアカウント名を
 * ActivityLog に残す。1回の画面表示で何度も通信が来るため、同じ人・同じ種類は10分に1回だけ記録する。
 * 「ログインできない」と言われたら、ActivityLog の操作欄が login_denied の行のメールアドレスと
 * AllowedAccounts を見比べると、アカウント違いや打ち間違いがすぐ分かる。
 */
var DENIED_LOG_INTERVAL_SECONDS = 600;
function logDeniedAccess_(user, kind, reason) {
  try {
    var email = String((user && user.email) || '').trim().toLowerCase();
    if (!email) return;
    var cache = CacheService.getScriptCache();
    var key = 'deny_' + kind + '_' + Utilities.base64EncodeWebSafe(email).slice(0, 200);
    if (cache.get(key)) return;
    cache.put(key, '1', DENIED_LOG_INTERVAL_SECONDS);
    var name = (user && user.name && user.name !== user.email) ? ('　Googleアカウント名：' + user.name) : '';
    appendLog(email, kind, reason + name);
  } catch (e) { /* 記録に失敗しても、本来の応答(拒否)はそのまま返す */ }
}

function appendLog(email, action, detail) {
  var sheet = getSheet(SHEET_LOG);
  // 未ログインの講師側から送られた文字列も記録されるため、長さを制限する(ログの肥大化・いたずら対策)
  sheet.appendRow(safeRow_([new Date(), String(email || '').slice(0, 200), action, String(detail == null ? '' : detail).slice(0, 500)]));
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('sheet_not_found:' + name);
  return sheet;
}

/** シートの列数が足りない場合は自動で広げる(コード側の項目追加にスプレッドシートが追いついていない事故を防ぐ) */
function ensureColumnCapacity(sheet, neededCols) {
  var have = sheet.getMaxColumns();
  if (have < neededCols) sheet.insertColumnsAfter(have, neededCols - have);
}

function getCourseSheet() {
  var sheet = getSheet(SHEET_COURSES);
  ensureColumnCapacity(sheet, COURSE_HEADERS.length);
  return sheet;
}

function getInvoiceSheet() {
  var sheet = getSheet(SHEET_INVOICES);
  ensureColumnCapacity(sheet, INVOICE_HEADERS.length);
  return sheet;
}

function formatDate(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

/** コメントの投稿日時など、日付と時刻をまとめて表示するための書式化 */
function formatDateTime(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  return v;
}

/** 開始時刻・終了時刻の読み取り正規化。
 * スプレッドシートが"09:00"のような文字列を時刻のDate値へ自動変換してしまうことがあるため、
 * その場合はHH:mm形式の文字列に戻す(そうしないと<input type="time">に正しく表示されない)。 */
function formatTimeValue(v) {
  if (!v && v !== 0) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }
  return v;
}

/** 報告書の月度("2026-09"のようなYYYY-MM)の読み取り正規化。
 * 開始時刻・終了時刻と同様に、スプレッドシートが"2026-09"のような文字列を日付のDate値へ
 * 自動変換してしまうことがあり、その場合は月度の絞り込み(getReportSubmissionStatus等)や
 * 既存行の検出(findReportRowIndex_)が一致せず、正しく保存したはずの報告書が「まだ作成
 * されていない」ように見えてしまう不具合の原因になっていたため、その場合はYYYY-MM形式の
 * 文字列に戻す。 */
function formatMonthValue_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return v;
}

/** ===== スプレッドシートへの書き込みの安全対策(数式インジェクション対策) =====
 * 「=」「+」「-」「@」で始まる文字列をそのままシートに書き込むと、数式として実行されてしまう
 * (講師用ページなどから「=IMPORTXML(...)」のような文字列を送られると、シートを開いた人の環境で
 * 外部へのデータ送信などが起きうる。また電話番号「+81-90-…」が計算されて壊れる不具合も起きる)。
 * すべての書き込みでこの関数を通し、先頭に「'」を付けて必ず文字列として保存する
 * (「'」はシート上では表示されず、読み出した値にも含まれない)。 */
function safeCell_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v) && !/^[+-]?\d+(\.\d+)?$/.test(v)) return "'" + v;
  return v;
}
/** http(s)のURLだけを受け付ける(「javascript:」などのリンクが保存され、他の理事が押して動いてしまうのを防ぐ) */
function safeHttpUrl_(v) {
  var u = String(v || '').trim();
  if (!u) return '';
  if (!/^https?:\/\/[^\s]+$/i.test(u)) throw new Error('URLは https:// で始まる形で入力してください。');
  return u.slice(0, 2000);
}
function safeRow_(row) { return (row || []).map(safeCell_); }
function safeRows_(rows) { return (rows || []).map(safeRow_); }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
