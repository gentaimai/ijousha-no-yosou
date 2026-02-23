# 日本選手権 予想アプリ（Google Apps Script / 無料運用）

仲間内で「各種目の上位3位予想」を入力・編集し、管理人が公開するまで他人の入力/集計を見せないための Web アプリです。

この実装は以下の2構成で運用できます（どちらも無料枠で運用可能）。

- `GitHub Pages（画面） + Google Apps Script + Google Spreadsheet`（推奨 / スマホ表示向け）
- `Google Apps Script + Google Spreadsheet` のみ

## できること

- 締切日時（例: `03/13`）まで何度でも予想を編集
- 管理人が公開するまで他人の予想は閲覧不可
- 管理人が公開するまで集計ページは閲覧不可
- 管理人が大会結果（1〜3位）を入力すると自動採点
- 2025年日本選手権のエントリー（種目・選手）を初期セットアップ時に自動投入

## 構成

- `/Users/imaigenta/Documents/RankMaker/gas/Code.gs` : サーバー処理（保存・権限・集計）
- `/Users/imaigenta/Documents/RankMaker/gas/Bridge.html` : GitHub Pages からの呼び出し用 bridge（hidden iframe 用）
- `/Users/imaigenta/Documents/RankMaker/gas/OfficialMasterData.gs` : 2025年公式PDFを解析して生成したエントリーマスタ
- `/Users/imaigenta/Documents/RankMaker/gas/Index.html` : 画面
- `/Users/imaigenta/Documents/RankMaker/gas/Client.html` : フロント JS
- `/Users/imaigenta/Documents/RankMaker/gas/Styles.html` : スタイル
- `/Users/imaigenta/Documents/RankMaker/frontend/index.html` : GitHub Pages 用画面
- `/Users/imaigenta/Documents/RankMaker/frontend/app.js` : GitHub Pages 用フロント JS
- `/Users/imaigenta/Documents/RankMaker/frontend/styles.css` : GitHub Pages 用スタイル
- `/Users/imaigenta/Documents/RankMaker/data/master-data.official-2025.json` : 2025年公式PDF解析結果（JSON）
- `/Users/imaigenta/Documents/RankMaker/scripts/parse_official_entries_pdf.py` : PDF解析スクリプト（再生成用）

## 前提

- Google アカウント
- Google Spreadsheet 1つ
- Google Apps Script（スプレッドシートに紐づける「コンテナバインド」を推奨）

## セットアップ手順

### 1. Google Apps Script（バックエンド）を用意

1. Google スプレッドシートを新規作成
1. `拡張機能 > Apps Script` を開く
1. `.gs` ファイルを作成して貼り付け
   - `Code.gs` ← `/Users/imaigenta/Documents/RankMaker/gas/Code.gs`
   - `OfficialMasterData.gs` ← `/Users/imaigenta/Documents/RankMaker/gas/OfficialMasterData.gs`
1. HTML ファイルを4つ作成して貼り付け
   - `Index`
   - `Client`
   - `Styles`
1. HTML ファイル `Bridge` も追加して貼り付け
   - `Bridge` ← `/Users/imaigenta/Documents/RankMaker/gas/Bridge.html`
1. `/Users/imaigenta/Documents/RankMaker/gas/appsscript.json` 相当の設定にする（タイムゾーン `Asia/Tokyo`）
1. Apps Script エディタで `setupApp()` を1回実行（権限許可が必要）
   - シート初期化に加え、同梱済みの2025年公式エントリーが自動投入されます
1. `デプロイ > 新しいデプロイ > ウェブアプリ`
   - 実行ユーザー: 自分
   - アクセスできるユーザー: リンクを知っている全員（仲間内のみ）
1. 発行された URL を共有
1. 管理用ページは `?page=admin` を付けたURLを自分だけ控える
   - 例: `https://script.google.com/.../exec?page=admin`

### 2. GitHub Pages（フロント）を用意（推奨）

1. GitHub リポジトリに `frontend/` の内容を配置（Pages 公開対象）
1. GitHub Pages を有効化
1. 初回アクセス時に URL パラメータ `gas` で GAS の `/exec` URL を指定
   - 例: `https://<yourname>.github.io/<repo>/?gas=https://script.google.com/macros/s/XXXX/exec`
1. 管理ページは `?page=admin&gas=...` でアクセス
   - 例: `https://<yourname>.github.io/<repo>/?page=admin&gas=https://script.google.com/macros/s/XXXX/exec`

以降は `gas` パラメータを指定するとブラウザに保存されるため、同じ端末では省略可能です。

## 初期運用フロー（管理者）

1. 管理ページ（`?page=admin`）を開く
1. 締切日時を設定（例: `2026-03-13T23:59:59+09:00`）
1. 大会終了後に結果を入力
1. 必要なタイミングで「予想公開」「集計公開」を ON

参加者は事前登録不要です。参加者ページで名前を入力すると、その名前で予想が保存されます。

## 2025年エントリーについて

- このリポジトリには、指定された 2025年日本選手権の PDF を解析して生成したエントリーを同梱しています
- `setupApp()` 実行時にそのデータが自動で `events` / `entries` シートへ投入されます
- 再生成したい場合は `/Users/imaigenta/Documents/RankMaker/scripts/parse_official_entries_pdf.py` を使って `gas/OfficialMasterData.gs` / `data/master-data.official-2025.json` を更新してください

## 採点ルール（現在の実装）

- 順位完全一致: `3点`
- 上位3位内的中（順位違い）: `1点`

`/Users/imaigenta/Documents/RankMaker/gas/Code.gs` の `scorePrediction_()` を変更すれば調整可能です。

## 注意点

- Apps Script の同時編集は小規模用途（5人程度）なら問題になりにくいですが、保存衝突が気になる場合は `LockService` を追加してください
