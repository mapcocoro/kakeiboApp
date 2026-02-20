# 💰 家計簿Webアプリ

家族で共有できる家計簿Webアプリです。Firebase Firestoreによりリアルタイムで同期されます。

## 🌐 アプリURL

**→ https://mapcocoro.github.io/kakeiboApp/**

- 家族全員がこのURLを使う
- スマホ・タブレット・PCどこからでもOK
- データはクラウド（Firebase）に保存・同期される

---

## 📱 使い方

### 日常の使い方
1. https://mapcocoro.github.io/kakeiboApp/ を開く
2. 「入力」タブで支出を記録
3. 自動でFirebaseに保存・家族全員に同期される

### バックアップ（月1回推奨）
1. **「入力」タブ**（または推移・ふるさと納税タブ）を開く
   - ⚠️ 一覧・分析タブでエクスポートするとフィルター中のデータのみになるので注意
2. 「データエクスポート」ボタンをクリック
3. CSVファイルをGoogleドライブ等に保存

### CSVからの復元
1. 「データインポート」をクリック
2. CSVファイルを選択
3. 確認ダイアログで **「キャンセル」** を選ぶ（既存データを削除して新規インポート）
4. 「削除してもよいですか？」→ OK

---

## ⚠️ 注意事項

### ローカルのindex.htmlは使わないこと
MacにあるHTMLファイルを直接ブラウザで開くと、Firebaseからデータが読み込まれlocalStorageが上書きされる場合があります。**必ずGitHub PagesのURLを使ってください。**

### データのバックアップについて
- データはFirebase（クラウド）に保存されるため、ブラウザのキャッシュ削除では消えません
- ただし万が一に備えて、月1回はCSVエクスポートでバックアップを取ってください
- バックアップ先: GoogleドライブやDropboxなど

### 同期エラーが出た場合
ヘッダーに「⚠ 同期エラー」と表示された場合:
- インターネット接続を確認
- ページを再読み込み（Cmd+R）

---

## 📁 ファイル構成

```
kakeibo-app/
├── index.html           # メインHTML
├── style.css            # スタイルシート
├── app.js               # アプリロジック
├── firebase-config.js   # Firebase設定（変更不要）
├── firebase-sync.js     # Firebase同期モジュール
└── CLAUDE.md            # 開発メモ
```

---

## 💾 データ保存先

| データ | 保存場所 |
|--------|---------|
| 支出データ | Firebase Firestore（クラウド） |
| ふるさと納税 | Firebase Firestore（クラウド） |
| 月別・年別メモ | Firebase Firestore（クラウド） |
| レポート | Firebase Firestore（クラウド） |

---

## 🔄 バージョン履歴

### v3.0 (2026-02-20)
- Firebase Firestoreによるリアルタイム同期を追加
- 家族複数人での同時利用・データ共有が可能に
- 大容量データのチャンク分割保存に対応
- ヘッダーに同期ステータスインジケーター追加

### v2.2 (2025-11-14)
- Topix完全復元ツール追加

### v2.1 (2025-11-14)
- GitHub Pages公開

### v1.0 (2025-11-13)
- 初回リリース

---

**Web版URL**: https://mapcocoro.github.io/kakeiboApp/
**リポジトリ**: https://github.com/mapcocoro/kakeiboApp.git
