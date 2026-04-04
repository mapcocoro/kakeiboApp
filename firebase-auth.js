// ========================================
// Firebase認証モジュール（メール/パスワード方式）
// ========================================

const firebaseAuth = (() => {
    const loginScreen = document.getElementById('loginScreen');
    const appContainer = document.querySelector('.container');
    const loginError = document.getElementById('loginError');
    const loginLoading = document.getElementById('loginLoading');
    const logoutBtn = document.getElementById('logoutBtn');
    const userNameEl = document.getElementById('userName');

    let authInitialized = false;

    function showApp(user) {
        loginScreen.style.display = 'none';
        appContainer.style.display = '';
        userNameEl.textContent = user.email;

        if (!authInitialized) {
            authInitialized = true;
            window.dispatchEvent(new CustomEvent('authReady', { detail: { user } }));
        }
    }

    function showLogin() {
        loginScreen.style.display = 'flex';
        appContainer.style.display = 'none';
        loginLoading.style.display = 'none';
    }

    function init() {
        if (typeof firebase === 'undefined') return;

        loginLoading.style.display = 'block';

        firebase.auth().onAuthStateChanged((user) => {
            if (user) {
                showApp(user);
            } else {
                showLogin();
            }
        });

        // ログインボタン
        document.getElementById('loginSubmitBtn').addEventListener('click', async () => {
            const email = document.getElementById('loginEmail').value.trim();
            const password = document.getElementById('loginPassword').value;
            const btn = document.getElementById('loginSubmitBtn');

            if (!email || !password) {
                loginError.textContent = 'メールアドレスとパスワードを入力してください';
                loginError.style.display = 'block';
                return;
            }

            btn.disabled = true;
            loginError.style.display = 'none';

            try {
                await firebase.auth().signInWithEmailAndPassword(email, password);
            } catch (error) {
                console.error('ログインエラー:', error);
                const messages = {
                    'auth/user-not-found': 'アカウントが見つかりません',
                    'auth/wrong-password': 'パスワードが違います',
                    'auth/invalid-email': 'メールアドレスの形式が正しくありません',
                    'auth/invalid-credential': 'メールアドレスまたはパスワードが違います',
                    'auth/too-many-requests': 'ログイン試行回数が多すぎます。しばらく待ってください'
                };
                loginError.textContent = messages[error.code] || 'ログインエラー: ' + error.message;
                loginError.style.display = 'block';
                btn.disabled = false;
            }
        });

        // Enterキーでログイン
        document.getElementById('loginPassword').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') document.getElementById('loginSubmitBtn').click();
        });

        // ログアウトボタン
        logoutBtn.addEventListener('click', async () => {
            if (confirm('ログアウトしますか？')) {
                authInitialized = false;
                await firebase.auth().signOut();
            }
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    return { init };
})();

window.firebaseAuth = firebaseAuth;
