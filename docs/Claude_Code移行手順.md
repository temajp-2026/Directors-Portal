# 理事ポータル：開発環境を Claude Code に移す手順

これまでの「ファイルを受け取って手でアップロード・貼り替え」から、Claude Code が GitHub と Apps Script に直接反映する形に移します。所要時間の目安は30〜60分です。

## 0. 先に済ませておくこと
- 2026-10-09 にお渡しした最新ファイル（HTML 11個・GAS_Code.gs）を、いつも通り反映しておく。
  （移行後は GitHub と Apps Script の内容が「正」になるため、最新の状態から始めます）

## 1. どの Claude Code を使うか決める

| 使い方 | 向いていること | 注意 |
|---|---|---|
| **自分のパソコンの Claude Code（デスクトップアプリ／ターミナル）**（おすすめ） | 画面（GitHub）も GAS（clasp）も直接反映できる | パソコンに Node.js が必要 |
| Claude Code on the web（クラウド） | パソコンの準備が不要 | GAS の反映（clasp）のログインが難しいため、GAS は今まで通り手で貼り替えになる |

以下は「自分のパソコンの Claude Code」の手順です。

## 2. GitHub の準備
1. Claude Code が使う GitHub アカウントに、`temajp-2026/Directors-Portal` への**書き込み権限**があることを確認する。
   - リポジトリの持ち主（temajp-2026）のアカウントでログインしていれば問題なし。
   - 別のアカウントを使う場合は、リポジトリの Settings → Collaborators でそのアカウントを追加する。
2. パソコンにリポジトリを取得する（Claude Code に「temajp-2026/Directors-Portal をクローンして」と頼めば実行してくれます）。

## 3. 移行キットをリポジトリに入れる
この移行キットの中身を、リポジトリの一番上の階層にそのまま置きます。

```
Directors-Portal/
├── CLAUDE.md            ← Claude Code が毎回読む「前提とルール」
├── .gitignore           ← ログイン情報をコミットしないための設定
├── gas/
│   ├── Code.gs          ← GAS のコード（2026-10-09 版）
│   └── .clasp.json.example
├── docs/
│   ├── 理事ポータル_機能説明書_2026-10-09.md
│   ├── 理事ポータル_操作ガイド補足資料_2026-10-09.md
│   ├── 引き継ぎ_2026-10-09.md
│   └── Claude_Code移行手順.md（この資料）
└── （今ある HTML・JS はそのまま）
```

**注意：このリポジトリは公開されています**（GitHub Pages の無料プランは公開リポジトリが前提）。
GAS のコードや資料も誰でも読める状態になります。コードにパスワードやAPIキーは入っていないので問題はありませんが、
今後 AI の API キーなどは**絶対にコードに書かず**、Apps Script の「スクリプト プロパティ」に保存します（CLAUDE.md にもルールとして記載済み）。
GAS のコードを非公開にしたい場合は、`gas/` だけ別の非公開リポジトリに置く方法もあります。

## 4. GAS を Claude Code から反映できるようにする（clasp）
clasp は、Apps Script のコードをパソコンから反映するための Google の公式ツールです。

1. https://script.google.com/home/usersettings を開き、「Google Apps Script API」を**オン**にする。
   （スクリプトを持っている Google アカウント＝ tema.jp.analytics@gmail.com でログインして操作）
2. Claude Code に次のように頼む：
   > 「clasp をインストールして、tema.jp.analytics@gmail.com でログインできるようにして」
   - ブラウザが開くので、tema.jp.analytics@gmail.com でログインして許可します。
   - ログイン情報はパソコンの中（.clasprc.json）に保存され、GitHub には上がりません。
3. スクリプトIDを確認する：Apps Script エディタ → 左の ⚙「プロジェクトの設定」→「スクリプト ID」をコピー。
4. Claude Code に次のように頼む：
   > 「gas/.clasp.json.example を参考に gas/.clasp.json を作って。スクリプトIDは ○○○。そのあと clasp pull して、gas/Code.gs と Apps Script の内容に差がないか確認して」
   - `appsscript.json`（Apps Script の設定ファイル）も取得されます。
5. デプロイIDを確認する：Apps Script エディタ →「デプロイ」→「デプロイを管理」→ 今のウェブアプリのデプロイID（AKfycbxOkLQh… で始まるもの）。
   これを Claude Code に伝えておくと、以後は「既存のデプロイを新しいバージョンで更新」までやってくれます。
   **新しいデプロイを作ると URL が変わり、ポータルとつながらなくなります。** 必ず既存のデプロイを更新します。

## 5. 最初に Claude Code に頼むこと（例）
> 「docs/引き継ぎ_2026-10-09.md と CLAUDE.md を読んで、現状を把握して。まだ何も変更しないで、次にやるべきことを整理して」

## 6. 移行後の流れ
1. やりたいことを Claude Code に伝える。
2. Claude Code がコードを直し、ブラウザで確認する（Playwright）。
3. 画面の変更：`main` に push → 数分で GitHub Pages に反映。
4. GAS の変更：`clasp push` → 既存デプロイを新バージョンで更新。列の追加がある場合は、案内に従い initializeSpreadsheet を実行。
5. 機能説明書（docs/）も更新 → NotebookLM のソースを入れ替える。

## 7. これまでのチャット（claude.ai のプロジェクト）について
- 機能説明書・補足資料・議事録・開発仕様書はプロジェクト「TeMA理事ポータル」にも残っています。
- 移行後は、リポジトリの `docs/` を最新版として扱ってください。
