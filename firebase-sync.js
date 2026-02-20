// ========================================
// Firebase同期モジュール
// ========================================

const firebaseSync = (() => {
    let db = null;
    const sessionId = Date.now().toString() + '_' + Math.random().toString(36).substr(2, 9);
    const COLLECTION = 'kakeiboData';

    // 保存デバウンス用
    let pendingSaves = {};
    let saveTimer = null;

    // 同期ステータス表示タイマー
    let statusTimer = null;

    // Firestoreからデータを読み込み、localStorageを更新する
    async function init(firebaseDb) {
        db = firebaseDb;

        const keys = ['expenses', 'savedReports', 'monthlyMemos', 'yearMemos', 'furusatoTaxData'];
        let loadedAny = false;

        showSyncStatus('loading');

        for (const key of keys) {
            try {
                const doc = await db.collection(COLLECTION).doc(key).get();
                if (doc.exists) {
                    const value = doc.data().value;
                    localStorage.setItem(key, JSON.stringify(value));
                    loadedAny = true;
                }
            } catch (error) {
                console.warn(`Firebase読み込みエラー (${key}):`, error);
            }
        }

        if (loadedAny) {
            showSyncStatus('loaded');
        } else {
            showSyncStatus('idle');
        }

        return loadedAny;
    }

    // リアルタイムリスナーを設定（他のユーザーの変更を検知）
    function startListening(callbacks) {
        if (!db) return;

        for (const [key, callback] of Object.entries(callbacks)) {
            db.collection(COLLECTION).doc(key).onSnapshot(snapshot => {
                if (!snapshot.exists) return;
                const data = snapshot.data();

                // 自分自身が保存したデータは無視
                if (data.sessionId === sessionId) return;

                const value = data.value;
                localStorage.setItem(key, JSON.stringify(value));

                showSyncStatus('synced');
                callback(value);
            }, error => {
                console.warn(`Firebaseリスナーエラー (${key}):`, error);
            });
        }
    }

    // データを保存（500msデバウンス）
    function save(key, data) {
        if (!db) return;

        pendingSaves[key] = data;

        if (saveTimer) clearTimeout(saveTimer);
        saveTimer = setTimeout(flushSaves, 500);
    }

    // 保留中の保存をまとめてFirestoreに書き込む
    async function flushSaves() {
        if (!db || Object.keys(pendingSaves).length === 0) return;

        const batch = db.batch();
        const serverTimestamp = firebase.firestore.FieldValue.serverTimestamp();

        for (const [key, data] of Object.entries(pendingSaves)) {
            const ref = db.collection(COLLECTION).doc(key);
            batch.set(ref, { value: data, updatedAt: serverTimestamp, sessionId: sessionId });
        }
        pendingSaves = {};

        try {
            showSyncStatus('saving');
            await batch.commit();
            showSyncStatus('saved');
        } catch (error) {
            console.error('Firebase保存エラー:', error);
            if (error.code === 'invalid-argument' && error.message.includes('exceed')) {
                showSyncStatus('error', 'データが大きすぎます。エクスポートで分割してください。');
            } else {
                showSyncStatus('error', '同期エラー: ' + error.message);
            }
        }
    }

    // 同期ステータスインジケーターを更新
    function showSyncStatus(status, message) {
        const indicator = document.getElementById('syncIndicator');
        if (!indicator) return;

        const config = {
            loading:  { text: '読み込み中...', color: '#888' },
            loaded:   { text: '✓ 読み込み完了', color: '#27ae60' },
            saving:   { text: '☁ 保存中...', color: '#888' },
            saved:    { text: '✓ 保存完了', color: '#27ae60' },
            synced:   { text: '↻ データ更新', color: '#2980b9' },
            idle:     { text: '', color: '#888' },
            error:    { text: '⚠ ' + (message || '同期エラー'), color: '#e74c3c' },
        };

        const c = config[status] || config.idle;
        indicator.textContent = c.text;
        indicator.style.color = c.color;

        if (statusTimer) clearTimeout(statusTimer);
        if (status === 'saved' || status === 'loaded' || status === 'synced') {
            statusTimer = setTimeout(() => {
                indicator.textContent = '';
            }, 3000);
        }
    }

    return { init, save, startListening };
})();

window.firebaseSync = firebaseSync;
