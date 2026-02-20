// ========================================
// Firebase同期モジュール（チャンク分割対応）
// ========================================

const firebaseSync = (() => {
    let db = null;
    const sessionId = Date.now().toString() + '_' + Math.random().toString(36).substr(2, 9);
    const COLLECTION = 'kakeiboData';
    const CHUNK_BYTES = 400 * 1024; // 400KB per chunk

    // 保存デバウンス用
    let pendingSaves = {};
    let saveTimer = null;

    // 同期ステータス表示タイマー
    let statusTimer = null;

    // expenses を400KB単位に分割
    function chunkExpenses(expenses) {
        const chunks = [];
        let current = [];
        let currentSize = 0;
        for (const item of expenses) {
            const size = JSON.stringify(item).length;
            if (currentSize + size > CHUNK_BYTES && current.length > 0) {
                chunks.push(current);
                current = [];
                currentSize = 0;
            }
            current.push(item);
            currentSize += size;
        }
        if (current.length > 0) chunks.push(current);
        return chunks.length > 0 ? chunks : [[]];
    }

    // expenses をチャンク分割して保存
    async function saveExpensesChunked(expenses) {
        const chunks = chunkExpenses(expenses);
        const ts = firebase.firestore.FieldValue.serverTimestamp();
        const batch = db.batch();

        // メタデータ（チャンク数を記録）
        batch.set(db.collection(COLLECTION).doc('expenses_meta'), {
            chunkCount: chunks.length, updatedAt: ts, sessionId
        });

        // 各チャンク
        chunks.forEach((chunk, i) => {
            batch.set(db.collection(COLLECTION).doc(`expenses_chunk_${i}`), {
                value: chunk, updatedAt: ts, sessionId
            });
        });

        await batch.commit();
    }

    // expenses を全チャンク結合して読み込み
    async function loadExpensesChunked() {
        const metaDoc = await db.collection(COLLECTION).doc('expenses_meta').get();

        if (!metaDoc.exists) {
            // 旧形式（単一ドキュメント）も試みる
            const doc = await db.collection(COLLECTION).doc('expenses').get();
            return doc.exists ? doc.data().value : null;
        }

        const { chunkCount } = metaDoc.data();
        const docs = await Promise.all(
            Array.from({ length: chunkCount }, (_, i) =>
                db.collection(COLLECTION).doc(`expenses_chunk_${i}`).get()
            )
        );
        return docs.flatMap(d => d.exists ? d.data().value : []);
    }

    // Firestoreからデータを読み込み、localStorageを更新する
    async function init(firebaseDb) {
        db = firebaseDb;
        showSyncStatus('loading');
        let loadedAny = false;

        try {
            // expenses（チャンク対応）
            const expenses = await loadExpensesChunked();
            if (expenses !== null) {
                localStorage.setItem('expenses', JSON.stringify(expenses));
                loadedAny = true;
            }

            // その他のキー（単一ドキュメント）
            const otherKeys = ['savedReports', 'monthlyMemos', 'yearMemos', 'furusatoTaxData'];
            for (const key of otherKeys) {
                const doc = await db.collection(COLLECTION).doc(key).get();
                if (doc.exists) {
                    localStorage.setItem(key, JSON.stringify(doc.data().value));
                    loadedAny = true;
                }
            }
        } catch (error) {
            console.warn('Firebase読み込みエラー:', error);
        }

        showSyncStatus(loadedAny ? 'loaded' : 'idle');
        return loadedAny;
    }

    // リアルタイムリスナーを設定（他のユーザーの変更を検知）
    function startListening(callbacks) {
        if (!db) return;

        // expenses はメタデータの変更を監視 → 全チャンクを再読み込み
        if (callbacks['expenses']) {
            db.collection(COLLECTION).doc('expenses_meta').onSnapshot(async snapshot => {
                if (!snapshot.exists) return;
                if (snapshot.data().sessionId === sessionId) return;

                const expenses = await loadExpensesChunked();
                if (expenses !== null) {
                    localStorage.setItem('expenses', JSON.stringify(expenses));
                    showSyncStatus('synced');
                    callbacks['expenses'](expenses);
                }
            }, err => console.warn('expensesリスナーエラー:', err));
        }

        // その他のキー
        const otherKeys = ['savedReports', 'monthlyMemos', 'yearMemos', 'furusatoTaxData'];
        for (const key of otherKeys) {
            if (!callbacks[key]) continue;
            db.collection(COLLECTION).doc(key).onSnapshot(snapshot => {
                if (!snapshot.exists) return;
                const data = snapshot.data();
                if (data.sessionId === sessionId) return;
                localStorage.setItem(key, JSON.stringify(data.value));
                showSyncStatus('synced');
                callbacks[key](data.value);
            }, err => console.warn(`${key}リスナーエラー:`, err));
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

        const toSave = { ...pendingSaves };
        pendingSaves = {};

        try {
            showSyncStatus('saving');

            // expenses はチャンク分割して保存
            if (toSave['expenses']) {
                await saveExpensesChunked(toSave['expenses']);
                delete toSave['expenses'];
            }

            // 残りのキーを一括保存
            if (Object.keys(toSave).length > 0) {
                const batch = db.batch();
                const ts = firebase.firestore.FieldValue.serverTimestamp();
                for (const [key, data] of Object.entries(toSave)) {
                    batch.set(db.collection(COLLECTION).doc(key), { value: data, updatedAt: ts, sessionId });
                }
                await batch.commit();
            }

            showSyncStatus('saved');
        } catch (error) {
            console.error('Firebase保存エラー:', error);
            showSyncStatus('error', '同期エラー: ' + error.message);
        }
    }

    // 同期ステータスインジケーターを更新
    function showSyncStatus(status, message) {
        const indicator = document.getElementById('syncIndicator');
        if (!indicator) return;

        const config = {
            loading: { text: '読み込み中...', color: '#888' },
            loaded:  { text: '✓ 読み込み完了', color: '#27ae60' },
            saving:  { text: '☁ 保存中...', color: '#888' },
            saved:   { text: '✓ 保存完了', color: '#27ae60' },
            synced:  { text: '↻ データ更新', color: '#2980b9' },
            idle:    { text: '', color: '#888' },
            error:   { text: '⚠ ' + (message || '同期エラー'), color: '#e74c3c' },
        };

        const c = config[status] || config.idle;
        indicator.textContent = c.text;
        indicator.style.color = c.color;

        if (statusTimer) clearTimeout(statusTimer);
        if (['saved', 'loaded', 'synced'].includes(status)) {
            statusTimer = setTimeout(() => { indicator.textContent = ''; }, 3000);
        }
    }

    return { init, save, startListening };
})();

window.firebaseSync = firebaseSync;
