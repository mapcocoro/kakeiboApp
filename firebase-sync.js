// ========================================
// Firebase同期モジュール（チャンク分割対応）
// ========================================

const firebaseSync = (() => {
    let db = null;
    const sessionId = Date.now().toString() + '_' + Math.random().toString(36).substr(2, 9);
    const COLLECTION = 'kakeiboData';
    const CHUNK_BYTES = 400 * 1024; // 400KB per chunk

    // バックアップ設定
    const BACKUP_INDEX_DOC = 'backup_index'; // 利用可能なバックアップ日付の一覧
    const MAX_BACKUPS = 30;                  // 保持する日数
    const OTHER_KEYS = ['savedReports', 'monthlyMemos', 'yearMemos', 'furusatoTaxData'];

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

        // バックアップ UI を有効化し、必要なら当日分を自動バックアップ
        setupBackupUI();
        if (loadedAny) {
            createDailyBackupIfNeeded(); // 非同期・完了待ちしない
        }

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

    // ====================================
    // 自動バックアップ機能
    // ====================================
    // バックアップは同じ kakeiboData コレクション内に
    // `backup_<日付>_<キー>` という名前で保存する（Firestoreルール変更不要）。
    // 利用可能な日付は `backup_index` ドキュメントで管理する。

    // ローカル日付を YYYY-MM-DD で返す
    function todayStr() {
        const d = new Date();
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    }

    // localStorage から現在のデータを取り出す
    function readLocal(key, fallback) {
        try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : fallback;
        } catch (e) {
            return fallback;
        }
    }

    // 指定日付のバックアップ用ドキュメントを全削除
    async function deleteBackupForDate(date) {
        const metaRef = db.collection(COLLECTION).doc(`backup_${date}_expenses_meta`);
        const metaDoc = await metaRef.get();
        const batch = db.batch();
        if (metaDoc.exists) {
            const { chunkCount } = metaDoc.data();
            for (let i = 0; i < chunkCount; i++) {
                batch.delete(db.collection(COLLECTION).doc(`backup_${date}_expenses_chunk_${i}`));
            }
            batch.delete(metaRef);
        }
        for (const key of OTHER_KEYS) {
            batch.delete(db.collection(COLLECTION).doc(`backup_${date}_${key}`));
        }
        await batch.commit();
    }

    // 1日1回、現在のデータをバックアップ（既に当日分があればスキップ）
    async function createDailyBackupIfNeeded() {
        if (!db) return;
        try {
            const indexRef = db.collection(COLLECTION).doc(BACKUP_INDEX_DOC);
            const indexDoc = await indexRef.get();
            let dates = indexDoc.exists ? (indexDoc.data().dates || []) : [];

            const today = todayStr();
            if (dates.includes(today)) return; // 当日分は作成済み

            // 現在のデータを取得（空なら読み込み失敗の可能性 → バックアップしない）
            const expenses = readLocal('expenses', []);
            const hasOther = OTHER_KEYS.some(k => localStorage.getItem(k));
            if ((!expenses || expenses.length === 0) && !hasOther) return;

            const chunks = chunkExpenses(expenses);
            const ts = firebase.firestore.FieldValue.serverTimestamp();
            const batch = db.batch();

            batch.set(db.collection(COLLECTION).doc(`backup_${today}_expenses_meta`), {
                chunkCount: chunks.length, createdAt: ts
            });
            chunks.forEach((chunk, i) => {
                batch.set(db.collection(COLLECTION).doc(`backup_${today}_expenses_chunk_${i}`), {
                    value: chunk, createdAt: ts
                });
            });
            for (const key of OTHER_KEYS) {
                const val = readLocal(key, null);
                if (val !== null) {
                    batch.set(db.collection(COLLECTION).doc(`backup_${today}_${key}`), {
                        value: val, createdAt: ts
                    });
                }
            }

            // 新しい日付を先頭に追加し、30日を超えたら古いものを削除
            dates = [today, ...dates];
            const toDelete = dates.slice(MAX_BACKUPS);
            dates = dates.slice(0, MAX_BACKUPS);

            batch.set(indexRef, { dates, updatedAt: ts });
            await batch.commit();

            for (const old of toDelete) {
                await deleteBackupForDate(old);
            }
            console.log(`自動バックアップ完了: ${today}（保持 ${dates.length} 件）`);
        } catch (error) {
            console.warn('自動バックアップ失敗:', error);
        }
    }

    // 利用可能なバックアップ日付の一覧を返す（新しい順）
    async function listBackups() {
        if (!db) return [];
        const indexDoc = await db.collection(COLLECTION).doc(BACKUP_INDEX_DOC).get();
        return indexDoc.exists ? (indexDoc.data().dates || []) : [];
    }

    // 指定日付のバックアップを現在のデータに復元する
    async function restoreBackup(date) {
        if (!db) throw new Error('未接続');

        // バックアップから読み込み
        const metaDoc = await db.collection(COLLECTION).doc(`backup_${date}_expenses_meta`).get();
        if (!metaDoc.exists) throw new Error('バックアップが見つかりません');
        const { chunkCount } = metaDoc.data();
        const chunkDocs = await Promise.all(
            Array.from({ length: chunkCount }, (_, i) =>
                db.collection(COLLECTION).doc(`backup_${date}_expenses_chunk_${i}`).get()
            )
        );
        const expenses = chunkDocs.flatMap(d => d.exists ? d.data().value : []);

        const others = {};
        for (const key of OTHER_KEYS) {
            const doc = await db.collection(COLLECTION).doc(`backup_${date}_${key}`).get();
            if (doc.exists) others[key] = doc.data().value;
        }

        // 現在のデータ（live）へ書き戻す
        showSyncStatus('saving');
        await saveExpensesChunked(expenses);
        const ts = firebase.firestore.FieldValue.serverTimestamp();
        const batch = db.batch();
        for (const [key, value] of Object.entries(others)) {
            batch.set(db.collection(COLLECTION).doc(key), { value, updatedAt: ts, sessionId });
        }
        await batch.commit();

        // localStorage も更新
        localStorage.setItem('expenses', JSON.stringify(expenses));
        for (const [key, value] of Object.entries(others)) {
            localStorage.setItem(key, JSON.stringify(value));
        }
        showSyncStatus('saved');
    }

    // ====================================
    // バックアップ UI（ヘッダーのボタン → モーダル）
    // ====================================
    function setupBackupUI() {
        const btn = document.getElementById('backupBtn');
        if (!btn || btn.dataset.ready) return;
        btn.dataset.ready = '1';

        btn.addEventListener('click', async () => {
            const overlay = document.createElement('div');
            overlay.style.cssText = 'position:fixed; inset:0; background:rgba(0,0,0,0.4); display:flex; justify-content:center; align-items:center; z-index:9999;';

            const box = document.createElement('div');
            box.style.cssText = 'background:#fff; border-radius:10px; padding:24px; width:90%; max-width:380px; max-height:80vh; overflow:auto; box-shadow:0 8px 30px rgba(0,0,0,0.2);';
            box.innerHTML = '<h3 style="margin:0 0 6px;">🛡 バックアップから復元</h3>' +
                '<p style="font-size:12px; color:#5f6368; margin:0 0 16px;">毎日アプリを開いた時に自動保存されています（直近30日分）。復元すると現在のデータが選んだ日の状態に置き換わります。</p>' +
                '<div id="backupList" style="font-size:13px; color:#888;">読み込み中...</div>' +
                '<div style="text-align:right; margin-top:16px;"><button id="backupCloseBtn" class="btn btn-secondary">閉じる</button></div>';

            overlay.appendChild(box);
            document.body.appendChild(overlay);
            const close = () => document.body.removeChild(overlay);
            overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
            box.querySelector('#backupCloseBtn').addEventListener('click', close);

            const listEl = box.querySelector('#backupList');
            try {
                const dates = await listBackups();
                if (dates.length === 0) {
                    listEl.textContent = 'まだバックアップがありません（明日以降、自動で作成されます）。';
                    return;
                }
                listEl.innerHTML = '';
                dates.forEach(date => {
                    const row = document.createElement('div');
                    row.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px solid #eee;';
                    const label = document.createElement('span');
                    label.textContent = date;
                    label.style.color = '#202124';
                    const rbtn = document.createElement('button');
                    rbtn.textContent = 'この日に戻す';
                    rbtn.className = 'btn btn-secondary';
                    rbtn.style.cssText = 'font-size:11px; padding:4px 10px;';
                    rbtn.addEventListener('click', async () => {
                        if (!confirm(`${date} の状態に戻します。\n現在のデータは上書きされます（今日の自動バックアップは残ります）。\nよろしいですか？`)) return;
                        rbtn.disabled = true;
                        rbtn.textContent = '復元中...';
                        try {
                            await restoreBackup(date);
                            alert('復元しました。画面を更新します。');
                            location.reload();
                        } catch (err) {
                            alert('復元に失敗しました: ' + err.message);
                            rbtn.disabled = false;
                            rbtn.textContent = 'この日に戻す';
                        }
                    });
                    row.appendChild(label);
                    row.appendChild(rbtn);
                    listEl.appendChild(row);
                });
            } catch (err) {
                listEl.textContent = '一覧の読み込みに失敗しました: ' + err.message;
            }
        });
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

    return { init, save, startListening, listBackups, restoreBackup, createDailyBackupIfNeeded };
})();

window.firebaseSync = firebaseSync;
