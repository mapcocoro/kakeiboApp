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

    // 認証状態の変化を監視
    function init() {
        if (typeof firebase === 'undefined') return;

        // ログイン中の表示
        loginLoading.style.display = 'block';
        loginBtn.style.display = 'none';

        firebase.auth().onAuthStateChanged((user) => {
            loginLoading.style.display = 'none';

            if (user) {
                // ログイン成功 → アプリ表示
                loginScreen.style.display = 'none';
                appContainer.style.display = '';
                userNameEl.textContent = user.displayName || user.email;

                // アプリ初期化イベントを発火
                window.dispatchEvent(new CustomEvent('authReady', { detail: { user } }));
            } else {
                // 未ログイン → ログイン画面表示
                loginScreen.style.display = 'flex';
                appContainer.style.display = 'none';
                loginBtn.style.display = 'inline-flex';
            }
        });

        // Googleログインボタン
        loginBtn.addEventListener('click', async () => {
            loginBtn.disabled = true;
            loginError.style.display = 'none';

            try {
                const provider = new firebase.auth.GoogleAuthProvider();
                await firebase.auth().signInWithRedirect(provider);
            } catch (error) {
                console.error('ログインエラー:', error);
                loginError.textContent = 'ログインに失敗しました: ' + error.message;
                loginError.style.display = 'block';
                loginBtn.disabled = false;
            }
        });

        // ログアウトボタン
        logoutBtn.addEventListener('click', async () => {
            if (confirm('ログアウトしますか？')) {
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
