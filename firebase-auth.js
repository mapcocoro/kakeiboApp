// ========================================
// Firebase認証モジュール
// ========================================
// Googleログインで認証し、認証済みユーザーのみアプリを使用可能にする

const firebaseAuth = (() => {
    const loginScreen = document.getElementById('loginScreen');
    const appContainer = document.querySelector('.container');
    const loginBtn = document.getElementById('googleLoginBtn');
    const loginError = document.getElementById('loginError');
    const loginLoading = document.getElementById('loginLoading');
    const logoutBtn = document.getElementById('logoutBtn');
    const userNameEl = document.getElementById('userName');

    let authInitialized = false;

    function showApp(user) {
        loginScreen.style.display = 'none';
        appContainer.style.display = '';
        userNameEl.textContent = user.displayName || user.email;

        // 初回のみアプリ初期化イベントを発火（ループ防止）
        if (!authInitialized) {
            authInitialized = true;
            window.dispatchEvent(new CustomEvent('authReady', { detail: { user } }));
        }
    }

    function showLogin() {
        loginScreen.style.display = 'flex';
        appContainer.style.display = 'none';
        loginBtn.style.display = 'inline-flex';
        loginLoading.style.display = 'none';
    }

    async function init() {
        if (typeof firebase === 'undefined') return;

        // ログイン中の表示
        loginLoading.style.display = 'block';
        loginBtn.style.display = 'none';

        // リダイレクト結果を処理（ページ復帰時）
        try {
            const result = await firebase.auth().getRedirectResult();
            if (result.user) {
                showApp(result.user);
                return;
            }
        } catch (error) {
            console.error('リダイレクト結果エラー:', error);
            loginError.textContent = 'ログインに失敗しました: ' + error.message;
            loginError.style.display = 'block';
        }

        // 既存セッションの確認
        firebase.auth().onAuthStateChanged((user) => {
            if (user) {
                showApp(user);
            } else {
                showLogin();
            }
        });

        // Googleログインボタン
        loginBtn.addEventListener('click', () => {
            loginBtn.disabled = true;
            loginError.style.display = 'none';
            loginLoading.style.display = 'block';

            const provider = new firebase.auth.GoogleAuthProvider();
            firebase.auth().signInWithRedirect(provider);
        });

        // ログアウトボタン
        logoutBtn.addEventListener('click', async () => {
            if (confirm('ログアウトしますか？')) {
                authInitialized = false;
                await firebase.auth().signOut();
            }
        });
    }

    // DOM読み込み後に初期化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    return { init };
})();

window.firebaseAuth = firebaseAuth;
