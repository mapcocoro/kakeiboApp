# kakeibo-app 開発メモ（Claude向け）

## プロジェクト概要

家族共有の家計簿Webアプリ。純粋なHTML/CSS/JS（ビルドツールなし）。
GitHub Pages でホスティング、Firebase Firestore でリアルタイム同期。

## 公開URL

- **アプリ**: https://mapcocoro.github.io/kakeiboApp/
- **リポジトリ**: https://github.com/mapcocoro/kakeiboApp.git

## Firebase

- **プロジェクトID**: kakeibo-8ad7c
- **リージョン**: asia-northeast1（東京）
- **Firestoreコレクション**: `kakeiboData`
- **ドキュメント構成**:
  - `expenses_meta` — チャンク数（chunkCount）を記録
  - `expenses_chunk_0`, `expenses_chunk_1`, ... — 支出データ（400KB単位で分割）
  - `savedReports` — 保存済みレポート
  - `monthlyMemos` — 月別メモ・予定
  - `yearMemos` — 年別メモ
  - `furusatoTaxData` — ふるさと納税データ

## ファイル構成

```
index.html          # メインHTML（同期インジケーター含む）
style.css           # スタイル
app.js              # アプリロジック（4300行超）
firebase-config.js  # Firebaseプロジェクト設定（APIキーあり・公開リポジトリ）
firebase-sync.js    # Firestore同期モジュール
```

## 重要な設計

### Firebase同期の仕組み
- `firebase-sync.js` がlocalStorageをラップして同期
- 保存時: localStorage → Firestore（500msデバウンス）
- 起動時: Firestoreからlocalストレージを上書き
- リアルタイム: `onSnapshot` で他ユーザーの変更を検知してUI更新
- 自分の変更は `sessionId` で識別してループを防止
- expensesのみチャンク分割（他キーは単一ドキュメント）

### エクスポート/インポートの注意
- エクスポートはCSV形式
- **一覧タブ** でエクスポート → フィルター中のデータのみ
- **入力/推移/ふるさと納税タブ** でエクスポート → 全データ
- インポート時「キャンセル」= 既存データを削除して置き換え（推奨）
- インポート時「OK」= 既存データに追加（重複に注意）

### APIキーのセキュリティ
- `firebase-config.js` にAPIキーがGitHubに公開されている
- Google Cloud ConsoleでHTTPリファラー制限済み（`https://mapcocoro.github.io/*` のみ許可）
- Firestoreルール: `kakeiboData` コレクションのみ読み書き可

## デプロイ方法

コードを修正したら:
```bash
git add <ファイル>
git commit -m "メッセージ"
git push origin main
```
GitHub Pagesが自動でデプロイ（1〜2分）

## ローカルのindex.htmlについて

Firebase同期が有効になっているため、ローカルで `index.html` を開くとFirestoreのデータでlocalStorageが上書きされる。
**基本的にローカルのindex.htmlは使わず、GitHub PagesのURLを使うこと。**

バックアップ用途でローカルを使う場合は、Firestoreにデータが入っている状態で開くと上書きされることに注意。
