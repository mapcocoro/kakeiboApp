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

        // 既存セッションの確認
        firebase.auth().onAuthStateChanged((user) => {
            if (user) {
                showApp(user);
            } else {
                showLogin();
            }
        });

        // Googleログインボタン（ポップアップ方式）
        loginBtn.addEventListener('click', async () => {
            loginBtn.disabled = true;
            loginError.style.display = 'none';

            try {
                const provider = new firebase.auth.GoogleAuthProvider();
                await firebase.auth().signInWithPopup(provider);
            } catch (error) {
                console.error('ログインエラー:', error);
                loginError.textContent = 'エラー: ' + error.code + ' - ' + error.message;
                loginError.style.display = 'block';
                loginBtn.disabled = false;
            }
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
