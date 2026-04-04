// ========================================
// Firebase設定ファイル
// ========================================
//
// 【セットアップ手順】
//
// 1. https://console.firebase.google.com/ にアクセス（Googleアカウントでログイン）
//
// 2. 「プロジェクトを追加」→ プロジェクト名を入力（例: kakeibo-family）
//    → Googleアナリティクスは「無効」でOK → 「プロジェクトを作成」
//
// 3. プロジェクト作成後、左サイドバー「構築」→「Firestore Database」
//    → 「データベースの作成」→「本番環境モード」→ ロケーションは「asia-northeast1（東京）」を選択
//
// 4. 左サイドバー「プロジェクトの概要（ホームアイコン）」→「</> ウェブ」アイコンをクリック
//    → アプリのニックネームを入力（例: kakeibo-web）→「アプリを登録」
//    → 表示されるfirebaseConfigの値を下記にコピー
//
// 5. Firestoreのセキュリティルールを設定（「Firestore Database」→「ルール」タブ）:
//    以下のルールを貼り付けて「公開」をクリック:
//
//    rules_version = '2';
//    service cloud.firestore {
//      match /databases/{database}/documents {
//        match /kakeiboData/{document} {
//          allow read, write: if request.auth != null;
//        }
//      }
//    }
//
// 6. Authentication を有効化（左サイドバー「構築」→「Authentication」）:
//    →「始める」→「Google」プロバイダを有効化 → プロジェクトの公開名を入力 → 保存
//
// ========================================

const firebaseConfig = {
    apiKey: "AIzaSyDL_PQR-8NDSDiMhV7UwT_s87lHU1d-d1g",
    authDomain: "kakeibo-8ad7c.firebaseapp.com",
    projectId: "kakeibo-8ad7c",
    storageBucket: "kakeibo-8ad7c.firebasestorage.app",
    messagingSenderId: "330842348952",
    appId: "1:330842348952:web:242466e6f6759a7e911fed"
};

// Firebase初期化
if (typeof firebase !== 'undefined') {
    firebase.initializeApp(firebaseConfig);
}
