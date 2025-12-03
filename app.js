// ========================================
// 状態管理
// ========================================

class StateManager {
    constructor() {
        this.stateKey = 'appState';
    }

    // 状態を保存
    saveState(state) {
        localStorage.setItem(this.stateKey, JSON.stringify(state));
    }

    // 状態を読み込み
    loadState() {
        const data = localStorage.getItem(this.stateKey);
        return data ? JSON.parse(data) : null;
    }

    // 特定のキーの状態を更新
    updateState(key, value) {
        const state = this.loadState() || {};
        state[key] = value;
        this.saveState(state);
    }
}

// ========================================
// データ管理
// ========================================

class ExpenseManager {
    constructor() {
        this.expenses = this.loadExpenses();
        this.reports = this.loadReports();
        this.memos = this.loadMemos();
        this.yearMemos = this.loadYearMemos();
    }

    // localStorageからデータを読み込み
    loadExpenses() {
        const data = localStorage.getItem('expenses');
        return data ? JSON.parse(data) : [];
    }

    // レポートを読み込み
    loadReports() {
        const data = localStorage.getItem('savedReports');
        return data ? JSON.parse(data) : [];
    }

    // レポートを保存
    saveReports() {
        localStorage.setItem('savedReports', JSON.stringify(this.reports));
    }

    // レポートを追加
    addReport(report) {
        report.id = Date.now().toString();
        report.createdAt = new Date().toISOString();
        this.reports.push(report);
        this.saveReports();
        return report;
    }

    // レポートを削除
    deleteReport(id) {
        this.reports = this.reports.filter(r => r.id !== id);
        this.saveReports();
    }

    // すべてのレポートを取得
    getAllReports() {
        return this.reports;
    }

    // メモを読み込み
    loadMemos() {
        const data = localStorage.getItem('monthlyMemos');
        return data ? JSON.parse(data) : {};
    }

    // メモを保存
    saveMemos() {
        localStorage.setItem('monthlyMemos', JSON.stringify(this.memos));
    }

    // 月のメモを取得
    getMemo(yearMonth) {
        // 毎回localStorageから最新データを読み込む
        const memos = this.loadMemos();
        return memos[yearMonth] || { events: '', plans: '' };
    }

    // 月のメモを保存
    saveMemo(yearMonth, events, plans) {
        this.memos[yearMonth] = { events, plans };
        this.saveMemos();
    }

    // 年度メモを読み込み
    loadYearMemos() {
        const data = localStorage.getItem('yearMemos');
        return data ? JSON.parse(data) : {};
    }

    // 年度メモを保存
    saveYearMemosToStorage() {
        localStorage.setItem('yearMemos', JSON.stringify(this.yearMemos));
    }

    // 年度メモを取得
    getYearMemo(year) {
        const memos = this.loadYearMemos();
        return memos[year] || '';
    }

    // 年度メモを保存
    saveYearMemo(year, memo) {
        this.yearMemos[year] = memo;
        this.saveYearMemosToStorage();
    }

    // localStorageにデータを保存
    saveExpenses() {
        try {
            const data = JSON.stringify(this.expenses);
            localStorage.setItem('expenses', data);

            // データサイズをログ出力
            const sizeInMB = (data.length / 1024 / 1024).toFixed(2);
            console.log(`データ保存: ${this.expenses.length}件 (${sizeInMB}MB)`);
        } catch (error) {
            if (error.name === 'QuotaExceededError') {
                alert('データ容量が上限を超えました。古いデータを削除するか、データを分割してインポートしてください。');
                throw new Error('LocalStorage容量超過');
            } else {
                throw error;
            }
        }
    }

    // 支出を追加
    addExpense(expense) {
        expense.id = Date.now().toString() + '_' + Math.random().toString(36).substr(2, 9);
        this.expenses.push(expense);
        this.saveExpenses();

        // ふるさと納税への自動同期
        this.syncToFurusato(expense);

        return expense;
    }

    // 複数の支出を一括追加（インポート用・保存は最後に1回だけ）
    addExpensesBatch(expenses) {
        expenses.forEach((expense, index) => {
            expense.id = Date.now().toString() + '_' + index + '_' + Math.random().toString(36).substr(2, 9);
            this.expenses.push(expense);

            // ふるさと納税への自動同期
            this.syncToFurusato(expense);
        });
        this.saveExpenses(); // 最後に1回だけ保存
    }

    // 支出を更新
    updateExpense(id, updatedData) {
        const index = this.expenses.findIndex(e => e.id === id);
        if (index !== -1) {
            const oldExpense = { ...this.expenses[index] };
            this.expenses[index] = { ...this.expenses[index], ...updatedData };
            this.saveExpenses();

            // ふるさと納税への自動同期（更新後のデータ）
            this.syncToFurusato(this.expenses[index], oldExpense);

            return true;
        }
        return false;
    }

    // 支出を削除
    deleteExpense(id) {
        // 削除前のデータを取得してふるさと納税から削除
        const expenseToDelete = this.expenses.find(e => e.id === id);
        if (expenseToDelete) {
            this.removeFurusatoSync(expenseToDelete);
        }

        this.expenses = this.expenses.filter(e => e.id !== id);
        this.saveExpenses();
    }

    // 特定のデータが既存データと重複しているかチェック
    isDuplicate(expense) {
        const key = `${expense.date}|${expense.category}|${expense.subcategory || ''}|${expense.amount}|${expense.place || ''}|${expense.description || ''}`;
        return this.expenses.some(e => {
            const existingKey = `${e.date}|${e.category}|${e.subcategory || ''}|${e.amount}|${e.place || ''}|${e.description || ''}`;
            return existingKey === key;
        });
    }

    // 重複データを検出
    findDuplicates() {
        const duplicates = [];
        const seen = new Map();

        this.expenses.forEach((expense, index) => {
            // 重複判定用のキーを生成（日付、カテゴリ、小項目、金額、場所、商品名）
            const key = `${expense.date}|${expense.category}|${expense.subcategory || ''}|${expense.amount}|${expense.place || ''}|${expense.description || ''}`;

            if (seen.has(key)) {
                // 重複データを記録
                const originalIndex = seen.get(key);
                if (!duplicates.find(d => d.key === key)) {
                    duplicates.push({
                        key,
                        indices: [originalIndex, index],
                        ids: [this.expenses[originalIndex].id, expense.id],
                        data: expense
                    });
                } else {
                    const dup = duplicates.find(d => d.key === key);
                    dup.indices.push(index);
                    dup.ids.push(expense.id);
                }
            } else {
                seen.set(key, index);
            }
        });

        return duplicates;
    }

    // 重複データを削除（最初の1件を残して、残りを削除）
    removeDuplicates() {
        const duplicates = this.findDuplicates();
        let removedCount = 0;

        duplicates.forEach(dup => {
            // 最初のID以外を削除
            for (let i = 1; i < dup.ids.length; i++) {
                this.expenses = this.expenses.filter(e => e.id !== dup.ids[i]);
                removedCount++;
            }
        });

        if (removedCount > 0) {
            this.saveExpenses();
        }

        return { duplicateGroups: duplicates.length, removedCount };
    }

    // すべての支出を取得
    getAllExpenses() {
        // 元の配列を変更しないようにコピーを返す
        return [...this.expenses];
    }

    // 期間でフィルタ
    getExpensesByMonth(year, month) {
        return this.expenses.filter(e => {
            const date = new Date(e.date);
            return date.getFullYear() === year && date.getMonth() + 1 === month;
        });
    }

    // 年でフィルタ
    getExpensesByYear(year) {
        return this.expenses.filter(e => {
            const date = new Date(e.date);
            return date.getFullYear() === year;
        });
    }

    // カテゴリ別集計
    getCategoryTotals(expenses) {
        const totals = {};
        expenses.forEach(e => {
            totals[e.category] = (totals[e.category] || 0) + parseInt(e.amount);
        });
        return totals;
    }

    // 月別集計
    getMonthlyTotals(year) {
        const monthlyData = Array(12).fill(0);
        this.expenses.forEach(e => {
            const date = new Date(e.date);
            if (date.getFullYear() === year) {
                monthlyData[date.getMonth()] += parseInt(e.amount);
            }
        });
        return monthlyData;
    }

    // ふるさと納税への自動同期
    syncToFurusato(expense, oldExpense = null) {
        // 小項目に「ふるさと納税」が含まれているかチェック
        const isFurusato = expense.subcategory && expense.subcategory.includes('ふるさと納税');
        const wasOldFurusato = oldExpense && oldExpense.subcategory && oldExpense.subcategory.includes('ふるさと納税');

        if (!isFurusato && !wasOldFurusato) {
            // ふるさと納税ではない場合は何もしない
            return;
        }

        // FurusatoManagerのインスタンスを取得（グローバル変数から）
        if (typeof furusatoManager === 'undefined') {
            console.warn('FurusatoManagerが初期化されていません');
            return;
        }

        // 更新の場合：古いデータがふるさと納税で、新しいデータがふるさと納税でない場合は削除
        if (oldExpense && wasOldFurusato && !isFurusato) {
            this.removeFurusatoSync(oldExpense);
            return;
        }

        // 更新の場合：古いデータを削除してから新しいデータを追加
        if (oldExpense && wasOldFurusato) {
            this.removeFurusatoSync(oldExpense);
        }

        // ふるさと納税データに変換
        const date = new Date(expense.date);
        const furusatoData = {
            year: date.getFullYear().toString(),
            amount: parseInt(expense.amount),
            item: expense.description || '（商品名なし）',
            applicant: expense.place || '（申請先不明）',
            municipality: ''
        };

        // 重複チェック
        const existingData = furusatoManager.getAll();
        const isDuplicate = existingData.some(existing =>
            existing.year === furusatoData.year &&
            parseInt(existing.amount) === parseInt(furusatoData.amount) &&
            existing.item === furusatoData.item &&
            existing.applicant === furusatoData.applicant
        );

        if (!isDuplicate) {
            furusatoManager.add(furusatoData);
            console.log('ふるさと納税に自動追加:', furusatoData);

            // ふるさと納税UIを更新（グローバル変数から）
            if (typeof furusatoUI !== 'undefined') {
                furusatoUI.render();
            }
        }
    }

    // ふるさと納税から削除（同期）
    removeFurusatoSync(expense) {
        // 小項目に「ふるさと納税」が含まれているかチェック
        if (!expense.subcategory || !expense.subcategory.includes('ふるさと納税')) {
            return;
        }

        // FurusatoManagerのインスタンスを取得
        if (typeof furusatoManager === 'undefined') {
            console.warn('FurusatoManagerが初期化されていません');
            return;
        }

        // 該当するふるさと納税データを検索して削除
        const date = new Date(expense.date);
        const year = date.getFullYear().toString();
        const amount = parseInt(expense.amount);
        const item = expense.description || '（商品名なし）';
        const applicant = expense.place || '（申請先不明）';

        const existingData = furusatoManager.getAll();
        const toDelete = existingData.find(existing =>
            existing.year === year &&
            parseInt(existing.amount) === amount &&
            existing.item === item &&
            existing.applicant === applicant
        );

        if (toDelete) {
            furusatoManager.delete(toDelete.id);
            console.log('ふるさと納税から自動削除:', toDelete);

            // ふるさと納税UIを更新
            if (typeof furusatoUI !== 'undefined') {
                furusatoUI.render();
            }
        }
    }
}

// ========================================
// UI管理
// ========================================

class UI {
    constructor(manager) {
        this.manager = manager;
        this.stateManager = new StateManager();
        this.monthlyChart = null;
        this.categoryChart = null;

        // カテゴリ別の小項目マスター（画像から更新）
        this.subcategoryMaster = {
            '食品': ['葉物・生鮮野菜', '根菜', '野菜その他', '肉類', '魚類・貝類', '肉類加工品', '魚類加工品', 'レトルト類・冷凍食品', 'お米', 'パン', '麺類', 'たまご', '牛乳・乳製品', '豆腐・納豆類', 'ドリンク', 'お菓子類', '調味料', '粉類', '食品その他'],
            '日用品': ['キッチン用品', 'バス用品', 'ペーパー類', '掃除用品', '文具', '日用品その他', '洗剤類', '美容関連用品', '薬類', '虫対策類', '衛生用品'],
            '外食費': ['外食', 'To Go', '外食その他'],
            '衣類': ['tettaインナー', 'natsukiインナー', 'tetta衣類', 'natsuki衣類', '衣類その他'],
            '家具・家電': ['家具', '家電', '家具家電その他'],
            '美容': ['スキンケア', 'ヘアケア', 'コスメ', 'エステ', '美容院', '美容皮膚科', 'ネイルサロン', '美容その他'],
            '医療費': ['診察料', '薬代', '医療費その他'],
            '交際費': ['記念日', '冠婚葬祭費', '贈答', '友人・知人', '寄付', '宮崎家関連', '宇野家関連', '交際費その他'],
            'レジャー': ['旅行', 'おでかけ', 'キャンプ', 'BBQ', '趣味', 'レジャーその他'],
            'ガソリン・ETC': ['ガソリン', 'ETC', 'ガソリンその他'],
            '光熱費': ['下水道', '水道', '電気', 'ガス', '光熱費その他'],
            '通信費': ['通信', 'ケータイ', '通信その他'],
            '保険': ['AIG損保', 'アクサダイレクト', '哲太メットライフ', '奈月メットライフ', '保険その他'],
            '車関連【車検・税金・積立】': ['積立', '車検', '税金', '車関連その他'],
            '税金': ['固定資産税', '市民税', '国民健康保険', '年金', 'ふるさと納税', '税金その他'],
            '経費': ['natsuki', 'tetta', '経費その他'],
            'ローン': ['ローン住居', 'ローン家電', 'ローン家具', 'ローンその他'],
            'その他': ['その他']
        };

        this.init();
    }

    init() {
        // 変数を先に初期化
        this.bulkInputRows = [];
        this.bulkRowIdCounter = 0; // 一括入力行のユニークIDカウンター
        this.currentSort = { column: 'date', direction: 'desc' }; // デフォルトは日付降順
        this.displayLimit = 100; // 初期表示件数
        this.displayOffset = 0; // 表示開始位置
        this.filteredExpenses = []; // フィルター適用後のデータ（一覧タブ用）
        this.analysisData = null; // 分析データ（分析タブ用）

        // UIセットアップ
        this.setupEventListeners();
        this.setupTabs();
        this.setDefaultDate();
        this.populateYearSelect();
        this.populateTimelineYearSelects();
        this.loadSavedReportsList();
        this.setDefaultMemoMonth();

        // 保存された状態を復元（タブ復元前に実行）
        this.restoreFilterState();

        // タブを復元（フィルター復元後に実行）
        this.restoreActiveTab();

        // 初期表示
        this.renderExpenseList();
    }

    // イベントリスナー設定
    setupEventListeners() {
        // 支出フォーム送信
        document.getElementById('expenseForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleAddExpense();
        });

        // 編集フォーム送信
        document.getElementById('editForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleUpdateExpense();
        });

        // 期間タイプ変更
        document.getElementById('dateRangeType').addEventListener('change', (e) => {
            this.handleDateRangeTypeChange(e.target.value);
            this.stateManager.updateState('dateRangeType', e.target.value);
        });

        // 検索ボックス
        document.getElementById('expenseSearch').addEventListener('input', (e) => {
            this.renderExpenseList(true); // フィルター変更時は表示件数リセット
        });

        // 月フィルター
        document.getElementById('monthFilter').addEventListener('change', (e) => {
            this.renderExpenseList(true); // フィルター変更時は表示件数リセット
            this.stateManager.updateState('monthFilter', e.target.value);
        });

        // カスタム期間の適用ボタン
        document.getElementById('applyFilter').addEventListener('click', () => {
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            this.stateManager.updateState('startDate', startDate);
            this.stateManager.updateState('endDate', endDate);
            this.renderExpenseList(true);
        });

        // 開始日・終了日の変更（Enterキーで適用）
        document.getElementById('startDate').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.stateManager.updateState('startDate', e.target.value);
                this.renderExpenseList(true);
            }
        });
        document.getElementById('endDate').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.stateManager.updateState('endDate', e.target.value);
                this.renderExpenseList(true);
            }
        });

        // フィルタークリア
        document.getElementById('clearFilter').addEventListener('click', () => {
            document.getElementById('expenseSearch').value = '';
            document.getElementById('dateRangeType').value = 'all';
            document.getElementById('monthFilter').value = '';
            document.getElementById('startDate').value = '';
            document.getElementById('endDate').value = '';
            this.handleDateRangeTypeChange('all');
            this.renderExpenseList(true); // フィルター変更時は表示件数リセット
        });

        // レポート保存
        document.getElementById('saveReportBtn').addEventListener('click', () => {
            this.saveCurrentReport();
        });

        // 保存済みレポート選択
        document.getElementById('savedReports').addEventListener('change', (e) => {
            this.loadSavedReport(e.target.value);
        });

        // 推移タブ - 更新ボタン
        document.getElementById('updateTimelineBtn').addEventListener('click', () => {
            const startYear = document.getElementById('timelineStartYear').value;
            const endYear = document.getElementById('timelineEndYear').value;
            this.stateManager.updateState('timelineStartYear', startYear);
            this.stateManager.updateState('timelineEndYear', endYear);
            this.renderTimeline();
        });

        // メモ - 読込ボタン
        document.getElementById('loadMemoBtn').addEventListener('click', () => {
            this.loadMemo();
        });

        // メモ - 保存ボタン
        document.getElementById('saveMemoBtn').addEventListener('click', () => {
            this.saveMemo();
        });

        // 分析期間タイプ変更
        document.getElementById('analysisDateRangeType').addEventListener('change', (e) => {
            this.handleAnalysisDateRangeTypeChange(e.target.value);
            this.stateManager.updateState('analysisDateRangeType', e.target.value);
        });

        // 分析年月変更
        document.getElementById('analysisYear').addEventListener('change', (e) => {
            this.stateManager.updateState('analysisYear', e.target.value);
            this.updateAnalysis();
        });

        document.getElementById('analysisMonth').addEventListener('change', (e) => {
            this.stateManager.updateState('analysisMonth', e.target.value);
            this.updateAnalysis();
        });

        // 分析カテゴリ変更時に小項目を更新
        document.getElementById('analysisCategory').addEventListener('change', () => {
            this.updateAnalysisSubcategoryOptions();
            this.updateAnalysis();
        });

        // 分析小項目変更
        document.getElementById('analysisSubcategory').addEventListener('change', () => {
            this.updateAnalysis();
        });

        // 分析カスタム期間変更
        document.getElementById('analysisStartDate').addEventListener('change', (e) => {
            this.stateManager.updateState('analysisStartDate', e.target.value);
            this.updateAnalysis();
        });

        document.getElementById('analysisEndDate').addEventListener('change', (e) => {
            this.stateManager.updateState('analysisEndDate', e.target.value);
            this.updateAnalysis();
        });

        // 年度メモ保存ボタン
        document.getElementById('saveYearMemoBtn').addEventListener('click', () => {
            this.saveYearMemo();
        });

        // エクスポート・インポート
        document.getElementById('removeDuplicatesBtn').addEventListener('click', () => {
            this.removeDuplicateData();
        });

        document.getElementById('exportBtn').addEventListener('click', () => {
            this.exportData();
        });

        document.getElementById('importBtn').addEventListener('click', () => {
            this.importData();
        });

        // モーダル閉じる
        document.querySelector('.close').addEventListener('click', () => {
            this.closeModal();
        });

        window.addEventListener('click', (e) => {
            if (e.target.id === 'editModal') {
                this.closeModal();
            }
        });

        // 一括入力モード切り替え
        document.getElementById('normalModeBtn').addEventListener('click', () => {
            this.switchMode('normal');
        });

        document.getElementById('bulkModeBtn').addEventListener('click', () => {
            this.switchMode('bulk');
        });

        // 一括入力: 行追加
        document.getElementById('addRowBtn').addEventListener('click', () => {
            this.addBulkInputRow();
        });

        // 一括入力: 5行追加
        document.getElementById('add5RowsBtn').addEventListener('click', () => {
            this.addBulkInputRows(5);
        });

        // 一括入力: 10行追加
        document.getElementById('add10RowsBtn').addEventListener('click', () => {
            this.addBulkInputRows(10);
        });

        // 一括入力: まとめて記録
        document.getElementById('saveBulkBtn').addEventListener('click', () => {
            this.saveBulkExpenses();
        });

        // 一括入力: 共通の日付と場所を全行に反映
        document.getElementById('applyCommonBtn').addEventListener('click', () => {
            this.applyCommonValues();
        });

        // 通常入力: カテゴリ変更時に小項目を更新
        document.getElementById('category').addEventListener('change', () => {
            this.updateNormalSubcategoryOptions();
        });

        // 通常入力: 税込ボタン
        document.getElementById('normalTaxBtn').addEventListener('click', () => {
            this.applyNormalTax();
        });

        // 編集モーダル: カテゴリ変更時に小項目を更新
        document.getElementById('editCategory').addEventListener('change', () => {
            this.updateEditSubcategoryOptions();
        });

        // テーブルヘッダーのソート
        document.querySelectorAll('.sortable').forEach(th => {
            th.addEventListener('click', () => {
                const sortColumn = th.dataset.sort;
                this.sortExpenseList(sortColumn);
            });
        });

        // 通常入力: 毎月繰り返すチェックボックス
        document.getElementById('normalRepeatMonthly').addEventListener('change', (e) => {
            const options = document.getElementById('normalRepeatOptions');
            options.style.display = e.target.checked ? 'block' : 'none';
        });

        // 一括入力: 毎月繰り返すチェックボックス
        document.getElementById('bulkRepeatMonthly').addEventListener('change', (e) => {
            const options = document.getElementById('bulkRepeatOptions');
            options.style.display = e.target.checked ? 'block' : 'none';
        });

        // 編集モーダル: 毎月繰り返すチェックボックス
        document.getElementById('editRepeatMonthly').addEventListener('change', (e) => {
            const options = document.getElementById('editRepeatOptions');
            options.style.display = e.target.checked ? 'block' : 'none';
        });

        // Topixモーダル: 閉じる
        document.getElementById('topixModalClose').addEventListener('click', () => {
            this.closeTopixModal();
        });

        document.getElementById('cancelTopixBtn').addEventListener('click', () => {
            this.closeTopixModal();
        });

        window.addEventListener('click', (e) => {
            if (e.target.id === 'topixModal') {
                this.closeTopixModal();
            }
        });

        // Topixモーダル: 追加
        document.getElementById('addTopixBtn').addEventListener('click', () => {
            this.addNewTopix();
        });

        // Topixモーダル: 保存
        document.getElementById('saveTopixBtn').addEventListener('click', () => {
            this.saveTopixChanges();
        });

        // Topixモーダル: Enterキーで追加
        document.getElementById('newTopixInput').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.addNewTopix();
            }
        });
    }

    // タブ切り替え
    setupTabs() {
        const tabButtons = document.querySelectorAll('.tab-btn');
        tabButtons.forEach(btn => {
            btn.addEventListener('click', () => {
                const targetTab = btn.dataset.tab;

                // タブボタンのアクティブ切り替え
                tabButtons.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');

                // タブコンテンツの表示切り替え
                document.querySelectorAll('.tab-content').forEach(content => {
                    content.classList.remove('active');
                });
                document.getElementById(`${targetTab}-tab`).classList.add('active');

                // タブ状態をlocalStorageに保存
                localStorage.setItem('activeTab', targetTab);

                // タブに応じて更新
                if (targetTab === 'analysis') {
                    this.updateAnalysis();
                } else if (targetTab === 'timeline') {
                    this.renderTimeline();
                }
            });
        });
    }

    // アクティブタブを復元
    restoreActiveTab() {
        const savedTab = localStorage.getItem('activeTab');
        if (savedTab) {
            const tabBtn = document.querySelector(`.tab-btn[data-tab="${savedTab}"]`);
            if (tabBtn) {
                tabBtn.click();
            }
        }
    }

    // フィルター状態を復元
    restoreFilterState() {
        const state = this.stateManager.loadState();
        if (!state) return;

        // 一覧タブのフィルター復元
        if (state.dateRangeType) {
            document.getElementById('dateRangeType').value = state.dateRangeType;
            this.handleDateRangeTypeChange(state.dateRangeType);
        }
        if (state.monthFilter) {
            document.getElementById('monthFilter').value = state.monthFilter;
        }
        if (state.startDate) {
            document.getElementById('startDate').value = state.startDate;
        }
        if (state.endDate) {
            document.getElementById('endDate').value = state.endDate;
        }

        // 推移タブの年度復元
        if (state.timelineStartYear) {
            document.getElementById('timelineStartYear').value = state.timelineStartYear;
        }
        if (state.timelineEndYear) {
            document.getElementById('timelineEndYear').value = state.timelineEndYear;
        }

        // 分析タブの設定復元
        if (state.analysisDateRangeType) {
            document.getElementById('analysisDateRangeType').value = state.analysisDateRangeType;
            this.handleAnalysisDateRangeTypeChange(state.analysisDateRangeType);
        }
        if (state.analysisYear) {
            document.getElementById('analysisYear').value = state.analysisYear;
        }
        if (state.analysisMonth) {
            document.getElementById('analysisMonth').value = state.analysisMonth;
        }
        if (state.analysisStartDate) {
            document.getElementById('analysisStartDate').value = state.analysisStartDate;
        }
        if (state.analysisEndDate) {
            document.getElementById('analysisEndDate').value = state.analysisEndDate;
        }
    }

    // デフォルトの日付を今日に設定
    setDefaultDate() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('date').value = today;
    }

    // 年選択プルダウンを作成
    populateYearSelect() {
        const select = document.getElementById('analysisYear');
        const currentYear = new Date().getFullYear();

        // データから最小年と最大年を取得
        const expenses = this.manager.getAllExpenses();
        let minYear = 2020;
        let maxYear = currentYear;
        if (expenses.length > 0) {
            const years = expenses.map(e => new Date(e.date).getFullYear());
            minYear = Math.min(...years, 2020); // 最小でも2020年から
            maxYear = Math.max(...years, currentYear);
        }

        // 未来10年まで拡張
        const futureYear = currentYear + 10;
        maxYear = Math.max(maxYear, futureYear);

        // 最小年から未来年まで
        for (let year = minYear; year <= maxYear; year++) {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = `${year}年`;
            select.appendChild(option);
        }

        // デフォルトは現在年
        select.value = currentYear;
    }

    // 推移タブの年選択プルダウンを作成
    populateTimelineYearSelects() {
        const startSelect = document.getElementById('timelineStartYear');
        const endSelect = document.getElementById('timelineEndYear');
        const currentYear = new Date().getFullYear();

        // データから最小年と最大年を取得
        const expenses = this.manager.getAllExpenses();
        let minYear = 2020;
        let maxYear = currentYear;
        if (expenses.length > 0) {
            const years = expenses.map(e => new Date(e.date).getFullYear());
            minYear = Math.min(...years, 2020); // 最小でも2020年から
            maxYear = Math.max(...years, currentYear);
        }

        // 未来10年まで拡張
        const futureYear = currentYear + 10;
        maxYear = Math.max(maxYear, futureYear);

        // 最小年から未来年まで
        for (let year = minYear; year <= maxYear; year++) {
            const option1 = document.createElement('option');
            option1.value = year;
            option1.textContent = `${year}年`;
            startSelect.appendChild(option1);

            const option2 = document.createElement('option');
            option2.value = year;
            option2.textContent = `${year}年`;
            endSelect.appendChild(option2);
        }

        // デフォルトは全期間（最小年から現在年まで）
        startSelect.value = minYear;
        endSelect.value = currentYear;
    }

    // メモの対象月をデフォルト設定
    setDefaultMemoMonth() {
        const today = new Date();
        const yearMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
        document.getElementById('memoYearMonth').value = yearMonth;
    }

    // 支出追加
    handleAddExpense() {
        const expense = {
            date: document.getElementById('date').value,
            category: document.getElementById('category').value,
            subcategory: document.getElementById('subcategory').value,
            amount: document.getElementById('amount').value,
            place: document.getElementById('place').value,
            description: document.getElementById('description').value,
            notes: document.getElementById('notes').value
        };

        // 毎月繰り返すチェックの確認
        const repeatCheckbox = document.getElementById('normalRepeatMonthly');
        const repeatMonths = parseInt(document.getElementById('normalRepeatMonths').value) || 1;

        if (repeatCheckbox.checked && repeatMonths > 1) {
            // 複数月分の支出を作成
            const baseDate = new Date(expense.date);
            for (let i = 0; i < repeatMonths; i++) {
                const newDate = new Date(baseDate);
                newDate.setMonth(baseDate.getMonth() + i);

                const repeatedExpense = {
                    ...expense,
                    date: newDate.toISOString().split('T')[0]
                };
                this.manager.addExpense(repeatedExpense);
            }
            this.showMessage(`${repeatMonths}ヶ月分の支出を記録しました`);
        } else {
            this.manager.addExpense(expense);
            this.showMessage('支出を記録しました');
        }

        this.renderExpenseList();

        // フォームをリセット
        document.getElementById('expenseForm').reset();
        this.setDefaultDate();
    }

    // 期間タイプ変更時の表示切り替え
    handleDateRangeTypeChange(type) {
        const monthFilter = document.getElementById('monthFilter');
        const startDate = document.getElementById('startDate');
        const endDate = document.getElementById('endDate');
        const applyBtn = document.getElementById('applyFilter');

        // すべて非表示にする
        monthFilter.style.display = 'none';
        startDate.style.display = 'none';
        endDate.style.display = 'none';
        applyBtn.style.display = 'none';

        // タイプに応じて表示
        if (type === 'month') {
            monthFilter.style.display = 'inline-block';
        } else if (type === 'custom') {
            startDate.style.display = 'inline-block';
            endDate.style.display = 'inline-block';
            applyBtn.style.display = 'inline-block';
        }

        // 全期間またはタイプ変更時にリスト更新
        if (type === 'all') {
            this.renderExpenseList(true);
        }
    }

    // 支出一覧を表示
    renderExpenseList(resetLimit = false) {
        try {
            const tbody = document.getElementById('expenseTableBody');
            if (!tbody) {
                console.error('テーブルボディが見つかりません');
                return;
            }

            const dateRangeType = document.getElementById('dateRangeType').value;
            const monthFilter = document.getElementById('monthFilter').value;
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;

            // フィルター変更時は表示件数をリセット
            if (resetLimit) {
                this.displayLimit = 100;
            }

            let expenses = this.manager.getAllExpenses();
            console.log(`全データ件数: ${expenses.length}件`);

            // 期間フィルター適用
            if (dateRangeType === 'month' && monthFilter) {
                const [year, month] = monthFilter.split('-').map(Number);
                console.log(`月フィルター: ${year}年${month}月`);
                expenses = this.manager.getExpensesByMonth(year, month);
                console.log(`フィルター後: ${expenses.length}件`);
            } else if (dateRangeType === 'custom' && startDate && endDate) {
                console.log(`カスタム期間: ${startDate} 〜 ${endDate}`);
                expenses = expenses.filter(e => {
                    return e.date >= startDate && e.date <= endDate;
                });
                console.log(`フィルター後: ${expenses.length}件`);
            }

            // 検索フィルター適用
            const searchText = document.getElementById('expenseSearch').value.toLowerCase().trim();
            if (searchText) {
                expenses = expenses.filter(e => {
                    const description = (e.description || '').toLowerCase();
                    const place = (e.place || '').toLowerCase();
                    const notes = (e.notes || '').toLowerCase();
                    const category = (e.category || '').toLowerCase();
                    const subcategory = (e.subcategory || '').toLowerCase();
                    return description.includes(searchText) ||
                           place.includes(searchText) ||
                           notes.includes(searchText) ||
                           category.includes(searchText) ||
                           subcategory.includes(searchText);
                });
                console.log(`検索フィルター後: ${expenses.length}件 (検索: "${searchText}")`);
            }

            // ソート適用
            expenses = this.applySortToExpenses(expenses);

            // フィルター適用後のデータを保存（エクスポート用）
            this.filteredExpenses = expenses;

            // ヘッダーのソート表示を更新
            this.updateSortIndicators();

            // 合計計算（全体）
            const total = expenses.reduce((sum, e) => sum + parseInt(e.amount || 0), 0);
            document.getElementById('displayTotal').textContent = total.toLocaleString();

            // フィルター情報を更新
            const filterInfo = searchText ? `検索結果（"${searchText}"）` : '表示期間';
            const filterCount = searchText ? `(${expenses.length}件)` : '';
            document.getElementById('filterInfo').textContent = filterInfo;
            document.getElementById('filterCount').textContent = filterCount;

            // 表示件数を制限
            const displayExpenses = expenses.slice(0, this.displayLimit);
            const hasMore = expenses.length > this.displayLimit;

            console.log(`表示件数: ${displayExpenses.length}/${expenses.length}件`);

            // テーブルに表示
            tbody.innerHTML = '';
            displayExpenses.forEach((expense, index) => {
                try {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${expense.date || '-'}</td>
                        <td>${expense.place || '-'}</td>
                        <td><span class="badge" data-category="${expense.category || ''}">${expense.category || '-'}</span></td>
                        <td>${expense.subcategory || '-'}</td>
                        <td>${parseInt(expense.amount || 0).toLocaleString()}円</td>
                        <td>${expense.description || '-'}</td>
                        <td>${expense.notes || '-'}</td>
                        <td class="action-buttons">
                            <button class="btn btn-edit" onclick="ui.openEditModal('${expense.id}')">編集</button>
                            <button class="btn btn-danger" onclick="ui.deleteExpense('${expense.id}')">削除</button>
                        </td>
                    `;
                    tbody.appendChild(row);
                } catch (rowError) {
                    console.error(`行${index}の表示エラー:`, rowError, expense);
                }
            });

            // 「もっと見る」ボタンの表示/非表示
            this.updateLoadMoreButton(hasMore, displayExpenses.length, expenses.length);

            console.log('一覧表示完了');
        } catch (error) {
            console.error('renderExpenseListエラー:', error);
            alert('データの表示中にエラーが発生しました。ブラウザのコンソールを確認してください。');
        }
    }

    // 「もっと見る」ボタンを更新
    updateLoadMoreButton(hasMore, displayed, total) {
        let loadMoreBtn = document.getElementById('loadMoreBtn');

        // ボタンが存在しない場合は作成
        if (!loadMoreBtn) {
            const tableContainer = document.querySelector('.table-container');
            loadMoreBtn = document.createElement('div');
            loadMoreBtn.id = 'loadMoreBtn';
            loadMoreBtn.style.cssText = 'text-align: center; padding: 20px;';
            tableContainer.parentElement.insertBefore(loadMoreBtn, tableContainer.nextSibling);
        }

        if (hasMore) {
            loadMoreBtn.innerHTML = `
                <button class="btn btn-secondary" onclick="ui.loadMoreExpenses()">
                    もっと見る（${displayed}/${total}件表示中）
                </button>
            `;
            loadMoreBtn.style.display = 'block';
        } else if (total > 0) {
            loadMoreBtn.innerHTML = `<p style="color: #666;">全${total}件を表示中</p>`;
            loadMoreBtn.style.display = 'block';
        } else {
            loadMoreBtn.style.display = 'none';
        }
    }

    // もっと見るボタンの処理
    loadMoreExpenses() {
        this.displayLimit += 100;
        this.renderExpenseList(false);
    }

    // ソート処理
    sortExpenseList(column) {
        // 同じ列をクリックした場合は昇順/降順を切り替え
        if (this.currentSort.column === column) {
            this.currentSort.direction = this.currentSort.direction === 'asc' ? 'desc' : 'asc';
        } else {
            // 異なる列の場合は新しい列で降順
            this.currentSort.column = column;
            this.currentSort.direction = 'desc';
        }

        this.renderExpenseList(true); // ソート変更時は表示件数リセット
    }

    // 支出データにソートを適用
    applySortToExpenses(expenses) {
        const { column, direction } = this.currentSort;
        const sorted = [...expenses];

        sorted.sort((a, b) => {
            let aVal, bVal;

            switch (column) {
                case 'date':
                    aVal = new Date(a.date);
                    bVal = new Date(b.date);
                    break;
                case 'amount':
                    aVal = parseInt(a.amount);
                    bVal = parseInt(b.amount);
                    break;
                case 'category':
                case 'subcategory':
                case 'place':
                case 'description':
                case 'notes':
                    aVal = (a[column] || '').toLowerCase();
                    bVal = (b[column] || '').toLowerCase();
                    break;
                default:
                    return 0;
            }

            if (aVal < bVal) return direction === 'asc' ? -1 : 1;
            if (aVal > bVal) return direction === 'asc' ? 1 : -1;
            return 0;
        });

        return sorted;
    }

    // ソート表示インジケーターを更新
    updateSortIndicators() {
        // すべてのヘッダーからソートクラスを削除
        document.querySelectorAll('.sortable').forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
        });

        // 現在のソート列にクラスを追加
        const currentHeader = document.querySelector(`.sortable[data-sort="${this.currentSort.column}"]`);
        if (currentHeader) {
            currentHeader.classList.add(`sort-${this.currentSort.direction}`);
        }
    }

    // 編集モーダルを開く
    openEditModal(id) {
        const expense = this.manager.expenses.find(e => e.id === id);
        if (!expense) return;

        document.getElementById('editId').value = expense.id;
        document.getElementById('editDate').value = expense.date;
        document.getElementById('editCategory').value = expense.category;

        // 小項目の選択肢を更新してから値を設定
        this.updateEditSubcategoryOptions();
        document.getElementById('editSubcategory').value = expense.subcategory || '';

        document.getElementById('editAmount').value = expense.amount;
        document.getElementById('editPlace').value = expense.place || '';
        document.getElementById('editDescription').value = expense.description || '';
        document.getElementById('editNotes').value = expense.notes || '';

        document.getElementById('editModal').style.display = 'block';
    }

    // モーダルを閉じる
    closeModal() {
        document.getElementById('editModal').style.display = 'none';
    }

    // 支出更新
    handleUpdateExpense() {
        const id = document.getElementById('editId').value;
        const updatedData = {
            date: document.getElementById('editDate').value,
            category: document.getElementById('editCategory').value,
            subcategory: document.getElementById('editSubcategory').value,
            amount: document.getElementById('editAmount').value,
            place: document.getElementById('editPlace').value,
            description: document.getElementById('editDescription').value,
            notes: document.getElementById('editNotes').value
        };

        // 現在のデータを更新
        this.manager.updateExpense(id, updatedData);

        // 毎月繰り返すチェックの確認
        const repeatCheckbox = document.getElementById('editRepeatMonthly');
        const repeatMonths = parseInt(document.getElementById('editRepeatMonths').value) || 1;

        if (repeatCheckbox.checked && repeatMonths > 1) {
            // 追加で繰り返し分を作成（2ヶ月目から）
            const baseDate = new Date(updatedData.date);
            for (let i = 1; i < repeatMonths; i++) {
                const newDate = new Date(baseDate);
                newDate.setMonth(baseDate.getMonth() + i);

                const repeatedExpense = {
                    ...updatedData,
                    date: newDate.toISOString().split('T')[0]
                };
                this.manager.addExpense(repeatedExpense);
            }
            this.showMessage(`支出を更新し、${repeatMonths - 1}ヶ月分を追加しました`);
        } else {
            this.showMessage('支出を更新しました');
        }

        this.renderExpenseList();
        this.closeModal();
    }

    // 支出削除
    deleteExpense(id) {
        if (confirm('本当に削除しますか?')) {
            this.manager.deleteExpense(id);
            this.renderExpenseList();
            this.showMessage('支出を削除しました');
        }
    }

    // レポート保存
    saveCurrentReport() {
        const reportName = prompt('レポート名を入力してください:', `レポート ${new Date().toLocaleDateString()}`);
        if (!reportName) return;

        const report = {
            name: reportName,
            dateRangeType: document.getElementById('dateRangeType').value,
            monthFilter: document.getElementById('monthFilter').value,
            startDate: document.getElementById('startDate').value,
            endDate: document.getElementById('endDate').value,
        };

        this.manager.addReport(report);
        this.loadSavedReportsList();
        this.showMessage('レポートを保存しました');
    }

    // 保存済みレポート一覧を読み込み
    loadSavedReportsList() {
        const select = document.getElementById('savedReports');
        const reports = this.manager.getAllReports();

        // 既存のオプションをクリア（最初のオプションを除く）
        while (select.options.length > 1) {
            select.remove(1);
        }

        // レポートを追加
        reports.forEach(report => {
            const option = document.createElement('option');
            option.value = report.id;
            option.textContent = report.name;
            select.appendChild(option);
        });
    }

    // 保存済みレポートを読み込み
    loadSavedReport(reportId) {
        if (!reportId) return;

        const report = this.manager.getAllReports().find(r => r.id === reportId);
        if (!report) return;

        // フィルター条件を復元
        document.getElementById('dateRangeType').value = report.dateRangeType;
        document.getElementById('monthFilter').value = report.monthFilter || '';
        document.getElementById('startDate').value = report.startDate || '';
        document.getElementById('endDate').value = report.endDate || '';

        // UIを更新
        this.handleDateRangeTypeChange(report.dateRangeType);
        this.renderExpenseList(true);

        this.showMessage(`レポート「${report.name}」を読み込みました`);
    }

    // 推移テーブルをレンダリング
    renderTimeline() {
        const startYear = parseInt(document.getElementById('timelineStartYear').value);
        const endYear = parseInt(document.getElementById('timelineEndYear').value);

        if (startYear > endYear) {
            alert('開始年は終了年以前である必要があります');
            return;
        }

        // 年月のリストを生成
        const months = [];
        for (let year = startYear; year <= endYear; year++) {
            for (let month = 1; month <= 12; month++) {
                months.push({ year, month, key: `${year}-${String(month).padStart(2, '0')}` });
            }
        }

        // カテゴリリスト
        const categories = ['食品', '日用品', '外食費', '衣類', '家具・家電', '美容', '医療費', '交際費', 'レジャー', 'ガソリン・ETC', '光熱費', '通信費', '保険', '車関連【車検・税金・積立】', '税金', '経費', 'ローン', 'その他'];

        // カテゴリごと・月ごとの集計
        const data = {};
        categories.forEach(cat => {
            data[cat] = {};
            months.forEach(m => {
                data[cat][m.key] = 0;
            });
        });

        // データを集計
        this.manager.getAllExpenses().forEach(expense => {
            const date = expense.date;
            const yearMonth = date.substring(0, 7); // YYYY-MM
            if (data[expense.category] && data[expense.category][yearMonth] !== undefined) {
                data[expense.category][yearMonth] += parseInt(expense.amount || 0);
            }
        });

        // テーブルヘッダーを生成（ソート可能）
        const thead = document.getElementById('timelineTableHead');
        thead.innerHTML = `
            <tr>
                <th class="sticky-col sortable-header" data-reset="true" title="クリックで元の順序に戻す" style="cursor: pointer;">カテゴリ</th>
                ${months.map((m, idx) => `<th class="sortable-header" data-month-index="${idx}" title="クリックでソート">${m.year.toString().substring(2)}.${String(m.month).padStart(2, '0')}</th>`).join('')}
            </tr>
        `;

        // カテゴリヘッダーにリセットイベントを追加
        const categoryHeader = thead.querySelector('th[data-reset="true"]');
        categoryHeader.addEventListener('click', () => {
            this.renderTimeline(); // 元の順序で再描画
        });

        // 月ヘッダーにソートイベントを追加
        thead.querySelectorAll('.sortable-header[data-month-index]').forEach(th => {
            th.addEventListener('click', () => {
                const monthIndex = parseInt(th.getAttribute('data-month-index'));
                this.sortTimelineByMonth(categories, data, months, monthIndex);
            });
        });

        // テーブルボディを生成
        const tbody = document.getElementById('timelineTableBody');
        tbody.innerHTML = '';

        // テーブル自体にtable-layout: fixedを設定
        const table = tbody.closest('table');
        if (table) {
            table.style.tableLayout = 'fixed';
            table.style.width = 'max-content';
        }

        categories.forEach(category => {
            const row = document.createElement('tr');

            // その行の最大値を見つける
            const amounts = months.map(m => data[category][m.key]);
            const maxAmount = Math.max(...amounts);

            const cells = months.map(m => {
                const amount = data[category][m.key];
                const displayAmount = amount > 0 ? `¥${amount.toLocaleString()}` : '';

                if (amount > 0 && maxAmount > 0) {
                    // 金額の比率を計算（最大値に対する割合）
                    const ratio = amount / maxAmount;
                    // 比率に応じて色の濃さを調整（0.3〜1.0）
                    const alpha = 0.3 + (ratio * 0.7);
                    const bgColor = this.getCategoryColorWithAlpha(category, alpha);
                    return `<td style="background:${bgColor}">${displayAmount}</td>`;
                } else {
                    return `<td>${displayAmount}</td>`;
                }
            });

            row.innerHTML = `
                <td class="sticky-col"><span class="badge" data-category="${category}">${category}</span></td>
                ${cells.join('')}
            `;
            tbody.appendChild(row);
        });

        // 合計行
        const totalRow = document.createElement('tr');
        totalRow.style.fontWeight = '500';
        totalRow.style.borderTop = '2px solid #202124';
        const totalCells = months.map(m => {
            const total = categories.reduce((sum, cat) => sum + data[cat][m.key], 0);
            return `<td>¥${total.toLocaleString()}</td>`;
        });
        totalRow.innerHTML = `
            <td class="sticky-col">合計</td>
            ${totalCells.join('')}
        `;
        tbody.appendChild(totalRow);

        // 出来事行（複数行対応）
        // 各月の出来事を配列として取得
        const monthsEvents = months.map(m => {
            const memo = this.manager.getMemo(m.key);
            console.log(`Topix取得: ${m.key}`, memo);
            const events = memo.events ? memo.events.trim() : '';
            return events ? events.split('\n').filter(e => e.trim()).slice(0, 10) : [];
        });

        console.log('monthsEvents配列:', monthsEvents);
        console.log('months配列の長さ:', months.length);
        console.log('monthsEvents配列の長さ:', monthsEvents.length);

        // 常に10行のTopix行を作成
        const topixRowCount = 10;
        console.log('Topix行数:', topixRowCount);

        // 10行のTopix行を作成
        for (let eventIndex = 0; eventIndex < topixRowCount; eventIndex++) {
            const memoRow = document.createElement('tr');
            memoRow.className = 'timeline-memo-row';
            memoRow.style.display = 'table-row';

            // 左端のTopixラベルセルを追加
            const labelCell = document.createElement('td');
            labelCell.className = 'memo-label';
            labelCell.textContent = 'Topix';
            // 直接スタイルを適用
            labelCell.style.minWidth = '100px';
            labelCell.style.maxWidth = '100px';
            labelCell.style.width = '100px';
            labelCell.style.position = 'sticky';
            labelCell.style.left = '0';
            labelCell.style.backgroundColor = '#fff8e1';
            labelCell.style.zIndex = '10';
            labelCell.style.fontWeight = '500';
            labelCell.style.color = '#202124';
            labelCell.style.display = 'table-cell';
            labelCell.style.boxSizing = 'border-box';
            labelCell.style.padding = '8px';
            memoRow.appendChild(labelCell);

            // 各月のセルを追加
            months.forEach((m, monthIndex) => {
                const events = monthsEvents[monthIndex];
                const event = events ? events[eventIndex] : null;

                const cell = document.createElement('td');
                cell.className = 'memo-cell';
                // 各セルに幅を設定
                cell.style.minWidth = '80px';
                cell.style.width = '80px';
                cell.style.display = 'table-cell';
                cell.style.boxSizing = 'border-box';
                cell.style.padding = '8px';

                if (event && event.trim()) {
                    const preview = event.length > 15 ? event.substring(0, 15) + '...' : event;
                    cell.title = event;
                    cell.textContent = preview;

                    if (eventIndex === 0 && monthIndex >= 4 && monthIndex < 10) {
                        console.log(`月${m.key}(index=${monthIndex}): "${preview}"`);
                    }
                }

                // クリックイベントを追加してtopix編集モーダルを開く
                cell.dataset.yearMonth = m.key;
                cell.addEventListener('click', () => {
                    this.openTopixModal(m.key);
                });

                memoRow.appendChild(cell);
            });

            console.log(`Topix行${eventIndex}: セル数=${memoRow.children.length - 1}`);
            tbody.appendChild(memoRow);
        }
    }

    // HTMLエスケープ
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // 出来事詳細を表示
    showMemoDetail(cell) {
        const memo = cell.getAttribute('data-memo');
        const month = cell.getAttribute('data-month');

        if (memo) {
            // エスケープされたHTMLエンティティを元に戻す
            const textarea = document.createElement('textarea');
            textarea.innerHTML = memo;
            const decodedMemo = textarea.value;

            alert(`【${month}の出来事】\n\n${decodedMemo}`);
        }
    }

    // カテゴリごとの薄い背景色を取得
    getCategoryColor(category) {
        const colors = {
            '食品': '#e8f5e9',
            '日用品': '#e3f2fd',
            '外食費': '#fff3e0',
            '衣類': '#f3e5f5',
            '家具・家電': '#e0f2f1',
            '美容': '#fce4ec',
            '医療費': '#ffebee',
            '交際費': '#fff8e1',
            'レジャー': '#e1f5fe',
            'ガソリン・ETC': '#efebe9',
            '光熱費': '#fbe9e7',
            '通信費': '#e8eaf6',
            '保険': '#e0f7fa',
            '車関連【車検・税金・積立】': '#f1f8e9',
            '税金': '#fce4ec',
            '経費': '#e0f2f1',
            'ローン': '#ede7f6',
            'その他': '#f5f5f5'
        };
        return colors[category] || '#f5f5f5';
    }

    // カテゴリごとの背景色をアルファ値付きで取得
    getCategoryColorWithAlpha(category, alpha) {
        const colorMap = {
            '食品': '76, 175, 80',       // green
            '日用品': '33, 150, 243',     // blue
            '外食費': '255, 152, 0',      // orange
            '衣類': '156, 39, 176',       // purple
            '家具・家電': '0, 150, 136',  // teal
            '美容': '233, 30, 99',        // pink
            '医療費': '244, 67, 54',      // red
            '交際費': '255, 193, 7',      // amber
            'レジャー': '3, 169, 244',    // light blue
            'ガソリン・ETC': '121, 85, 72', // brown
            '光熱費': '255, 87, 34',      // deep orange
            '通信費': '103, 58, 183',     // deep purple
            '保険': '0, 188, 212',        // cyan
            '車関連【車検・税金・積立】': '139, 195, 74', // light green
            '税金': '233, 30, 99',        // pink
            '経費': '0, 150, 136',        // teal
            'ローン': '103, 58, 183',     // deep purple
            'その他': '158, 158, 158'     // grey
        };

        const rgb = colorMap[category] || '158, 158, 158';
        return `rgba(${rgb}, ${alpha})`;
    }

    // 推移表を特定の月でソート
    sortTimelineByMonth(categories, data, months, monthIndex) {
        const month = months[monthIndex];

        // 指定された月の金額でカテゴリをソート（降順）
        const sortedCategories = [...categories].sort((a, b) => {
            const amountA = data[a][month.key] || 0;
            const amountB = data[b][month.key] || 0;
            return amountB - amountA; // 降順
        });

        // テーブルボディを再生成
        const tbody = document.getElementById('timelineTableBody');
        tbody.innerHTML = '';

        // 選択された月の合計を計算
        let selectedMonthTotal = 0;

        sortedCategories.forEach(category => {
            const row = document.createElement('tr');

            // その行の最大値を見つける
            const amounts = months.map(m => data[category][m.key]);
            const maxAmount = Math.max(...amounts);

            const cells = months.map((m, idx) => {
                const amount = data[category][m.key];
                const displayAmount = amount > 0 ? `¥${amount.toLocaleString()}` : '';

                // 選択された月の場合、合計に加算
                if (idx === monthIndex) {
                    selectedMonthTotal += amount;
                }

                if (amount > 0 && maxAmount > 0) {
                    const ratio = amount / maxAmount;
                    const alpha = 0.3 + (ratio * 0.7);
                    const bgColor = this.getCategoryColorWithAlpha(category, alpha);

                    // 選択された列を強調表示
                    const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8;' : '';
                    return `<td style="background:${bgColor}; ${highlight}">${displayAmount}</td>`;
                } else {
                    const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8;' : '';
                    return `<td style="${highlight}">${displayAmount}</td>`;
                }
            });

            row.innerHTML = `
                <td class="sticky-col"><span class="badge" data-category="${category}">${category}</span></td>
                ${cells.join('')}
            `;
            tbody.appendChild(row);
        });

        // 合計行
        const totalRow = document.createElement('tr');
        totalRow.style.fontWeight = '500';
        totalRow.style.borderTop = '2px solid #202124';
        const totalCells = months.map((m, idx) => {
            const total = categories.reduce((sum, cat) => sum + data[cat][m.key], 0);
            const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8; background: #e8f0fe;' : '';
            return `<td style="${highlight}">¥${total.toLocaleString()}</td>`;
        });
        totalRow.innerHTML = `
            <td class="sticky-col">合計</td>
            ${totalCells.join('')}
        `;
        tbody.appendChild(totalRow);

        // 選択された月の合計を表示するメッセージ
        const infoRow = document.createElement('tr');
        infoRow.style.background = '#e8f0fe';
        infoRow.style.fontWeight = '500';
        infoRow.innerHTML = `
            <td class="sticky-col" colspan="${months.length + 1}" style="text-align: center; padding: 12px;">
                ${month.year}年${month.month}月でソート済み - 合計: ¥${selectedMonthTotal.toLocaleString()}
            </td>
        `;
        tbody.appendChild(infoRow);

        // 出来事行を再追加（複数行対応）
        // 各月の出来事を配列として取得
        const monthsEvents = months.map(m => {
            const memo = this.manager.getMemo(m.key);
            console.log(`Topix取得(ソート後): ${m.key}`, memo);
            const events = memo.events ? memo.events.trim() : '';
            return events ? events.split('\n').filter(e => e.trim()).slice(0, 10) : [];
        });

        // 常に10行のTopix行を作成
        const topixRowCount = 10;

        for (let eventIndex = 0; eventIndex < topixRowCount; eventIndex++) {
            const memoRow = document.createElement('tr');
            memoRow.className = 'timeline-memo-row';
            memoRow.style.display = 'table-row';

            // 左端のTopixラベルセルを追加
            const labelCell = document.createElement('td');
            labelCell.className = 'memo-label';
            labelCell.textContent = 'Topix';
            // 直接スタイルを適用
            labelCell.style.minWidth = '100px';
            labelCell.style.maxWidth = '100px';
            labelCell.style.width = '100px';
            labelCell.style.position = 'sticky';
            labelCell.style.left = '0';
            labelCell.style.backgroundColor = '#fff8e1';
            labelCell.style.zIndex = '10';
            labelCell.style.fontWeight = '500';
            labelCell.style.color = '#202124';
            labelCell.style.display = 'table-cell';
            labelCell.style.boxSizing = 'border-box';
            labelCell.style.padding = '8px';
            memoRow.appendChild(labelCell);

            // 各月のセルを追加
            months.forEach((m, idx) => {
                const events = monthsEvents[idx];
                const event = events ? events[eventIndex] : null;

                const cell = document.createElement('td');
                cell.className = 'memo-cell';

                // 選択された列を強調表示
                if (idx === monthIndex) {
                    cell.style.border = '2px solid #1a73e8';
                }

                if (event && event.trim()) {
                    const preview = event.length > 15 ? event.substring(0, 15) + '...' : event;
                    cell.title = event;
                    cell.textContent = preview;
                }

                // クリックイベントを追加してtopix編集モーダルを開く
                cell.dataset.yearMonth = m.key;
                cell.addEventListener('click', () => {
                    this.openTopixModal(m.key);
                });

                memoRow.appendChild(cell);
            });

            tbody.appendChild(memoRow);
        }
    }

    // メモを読み込み
    loadMemo() {
        const yearMonth = document.getElementById('memoYearMonth').value;
        if (!yearMonth) {
            alert('対象月を選択してください');
            return;
        }

        const memo = this.manager.getMemo(yearMonth);
        document.getElementById('memoEvents').value = memo.events;
        document.getElementById('memoPlans').value = memo.plans;

        this.showMessage(`${yearMonth}のメモを読み込みました`);
    }

    // メモを保存
    saveMemo() {
        const yearMonth = document.getElementById('memoYearMonth').value;
        if (!yearMonth) {
            alert('対象月を選択してください');
            return;
        }

        const events = document.getElementById('memoEvents').value;
        const plans = document.getElementById('memoPlans').value;

        this.manager.saveMemo(yearMonth, events, plans);
        this.showMessage(`${yearMonth}のメモを保存しました`);
    }

    // 年度メモを読み込み
    loadYearMemo(year) {
        const memo = this.manager.getYearMemo(year.toString());
        document.getElementById('yearMemoText').value = memo;
        // 現在表示している年度を保存
        this.currentYearForMemo = year;
    }

    // 年度メモを保存
    saveYearMemo() {
        const year = this.currentYearForMemo;
        if (!year) {
            alert('年度が指定されていません');
            return;
        }

        const memo = document.getElementById('yearMemoText').value;
        this.manager.saveYearMemo(year.toString(), memo);
        this.showMessage(`${year}年のメモを保存しました`);
    }

    // 分析期間タイプ変更時の表示切り替え
    handleAnalysisDateRangeTypeChange(type) {
        const yearLabel = document.getElementById('analysisYearLabel');
        const yearSelect = document.getElementById('analysisYear');
        const monthLabel = document.getElementById('analysisMonthLabel');
        const monthSelect = document.getElementById('analysisMonth');
        const startDateLabel = document.getElementById('analysisStartDateLabel');
        const startDate = document.getElementById('analysisStartDate');
        const endDateLabel = document.getElementById('analysisEndDateLabel');
        const endDate = document.getElementById('analysisEndDate');

        // すべて非表示
        yearLabel.style.display = 'none';
        yearSelect.style.display = 'none';
        monthLabel.style.display = 'none';
        monthSelect.style.display = 'none';
        startDateLabel.style.display = 'none';
        startDate.style.display = 'none';
        endDateLabel.style.display = 'none';
        endDate.style.display = 'none';

        // タイプに応じて表示
        if (type === 'year') {
            yearLabel.style.display = 'inline';
            yearSelect.style.display = 'inline';
        } else if (type === 'month') {
            yearLabel.style.display = 'inline';
            yearSelect.style.display = 'inline';
            monthLabel.style.display = 'inline';
            monthSelect.style.display = 'inline';
        } else if (type === 'custom') {
            startDateLabel.style.display = 'inline';
            startDate.style.display = 'inline';
            endDateLabel.style.display = 'inline';
            endDate.style.display = 'inline';
        }

        this.updateAnalysis();
    }

    // 分析タブの小項目選択肢を更新
    updateAnalysisSubcategoryOptions() {
        const categorySelect = document.getElementById('analysisCategory');
        const subcategoryLabel = document.getElementById('analysisSubcategoryLabel');
        const subcategorySelect = document.getElementById('analysisSubcategory');
        const selectedCategory = categorySelect.value;

        // 小項目の選択肢をクリア
        subcategorySelect.innerHTML = '<option value="">全小項目</option>';

        // カテゴリが選択されている場合
        if (selectedCategory && this.subcategoryMaster[selectedCategory]) {
            subcategoryLabel.style.display = 'inline';
            subcategorySelect.style.display = 'inline';

            this.subcategoryMaster[selectedCategory].forEach(subcat => {
                const option = document.createElement('option');
                option.value = subcat;
                option.textContent = subcat;
                subcategorySelect.appendChild(option);
            });
        } else {
            subcategoryLabel.style.display = 'none';
            subcategorySelect.style.display = 'none';
        }
    }

    // 分析データ更新
    updateAnalysis() {
        const dateRangeType = document.getElementById('analysisDateRangeType').value;
        const year = parseInt(document.getElementById('analysisYear').value);
        const month = document.getElementById('analysisMonth').value;
        const startDate = document.getElementById('analysisStartDate').value;
        const endDate = document.getElementById('analysisEndDate').value;
        const selectedCategory = document.getElementById('analysisCategory').value;
        const selectedSubcategory = document.getElementById('analysisSubcategory').value;

        let expenses = this.manager.getAllExpenses();

        // 期間フィルター適用
        if (dateRangeType === 'year' && year) {
            expenses = this.manager.getExpensesByYear(year);
        } else if (dateRangeType === 'month' && year && month) {
            expenses = this.manager.getExpensesByMonth(year, parseInt(month));
        } else if (dateRangeType === 'custom' && startDate && endDate) {
            expenses = expenses.filter(e => {
                return e.date >= startDate && e.date <= endDate;
            });
        }

        // カテゴリフィルター適用
        if (selectedCategory) {
            expenses = expenses.filter(e => e.category === selectedCategory);
        }

        // 小項目フィルター適用
        if (selectedSubcategory) {
            expenses = expenses.filter(e => e.subcategory === selectedSubcategory);
        }

        // 分析データを保存（エクスポート用）
        this.analysisData = {
            expenses: expenses,
            dateRangeType: dateRangeType,
            year: year,
            month: month,
            startDate: startDate,
            endDate: endDate,
            category: selectedCategory,
            subcategory: selectedSubcategory
        };

        // タイトル更新
        if (selectedSubcategory) {
            document.getElementById('categoryStatsTitle').textContent =
                `${selectedCategory} - ${selectedSubcategory}`;
            document.getElementById('monthlyChartTitle').textContent =
                `${selectedCategory} - ${selectedSubcategory} - 月別推移`;
        } else if (selectedCategory) {
            document.getElementById('categoryStatsTitle').textContent =
                `${selectedCategory} - 小項目別集計`;
            document.getElementById('monthlyChartTitle').textContent =
                `${selectedCategory} - 月別推移`;
        } else {
            document.getElementById('categoryStatsTitle').textContent = 'カテゴリ別集計';
            document.getElementById('monthlyChartTitle').textContent = '月別推移';
        }

        // カテゴリ別集計または小項目別集計
        if (selectedSubcategory) {
            // 小項目が選択されている場合は簡易統計を表示
            this.renderSimpleStats(expenses, selectedCategory, selectedSubcategory);
        } else if (selectedCategory) {
            this.renderSubcategoryStats(expenses, selectedCategory);
        } else {
            this.renderCategoryStats(expenses);
        }

        // グラフ更新
        this.updateChartsWithPeriod(expenses, dateRangeType, year, selectedCategory, selectedSubcategory);

        // 年別表の表示・非表示
        if (dateRangeType === 'year' && year && !selectedCategory) {
            // 年指定かつカテゴリ未選択の場合のみ年別表を表示
            document.getElementById('yearlyTableCard').style.display = 'block';
            this.renderYearlyTable(year);
        } else {
            document.getElementById('yearlyTableCard').style.display = 'none';
        }

        // 年度メモの表示・非表示
        if (dateRangeType === 'year' && year) {
            // 年指定の場合は年度メモを表示
            document.getElementById('yearMemoCard').style.display = 'block';
            document.getElementById('yearMemoTitle').textContent = `${year}年 メモ`;
            this.loadYearMemo(year);
        } else {
            document.getElementById('yearMemoCard').style.display = 'none';
        }
    }

    // 簡易統計表示（小項目選択時）
    renderSimpleStats(expenses, category, subcategory) {
        const container = document.getElementById('categoryStats');
        const total = expenses.reduce((sum, e) => sum + parseInt(e.amount), 0);
        const count = expenses.length;
        const average = count > 0 ? Math.round(total / count) : 0;

        container.innerHTML = `
            <div class="category-summary">
                <div class="summary-item">
                    <span>合計金額:</span>
                    <strong>${total.toLocaleString()}円</strong>
                </div>
                <div class="summary-item">
                    <span>件数:</span>
                    <strong>${count}件</strong>
                </div>
                <div class="summary-item">
                    <span>平均:</span>
                    <strong>${average.toLocaleString()}円</strong>
                </div>
            </div>
        `;
    }

    // グラフ更新（期間対応版）
    updateChartsWithPeriod(expenses, dateRangeType, year, selectedCategory = null, selectedSubcategory = null) {
        // 年指定の場合は従来通り12ヶ月、それ以外は年月別
        const useYearMonth = dateRangeType !== 'year';

        if (selectedSubcategory) {
            // 小項目選択時：月別推移と月別データ一覧、取引明細
            const monthlyData = this.getMonthlyDataFromExpenses(expenses, useYearMonth);
            this.renderMonthlyChart(monthlyData, false, useYearMonth);
            this.renderMonthlyDataTable(expenses);
            this.renderSubcategoryDetails(expenses, selectedSubcategory);
        } else if (selectedCategory) {
            // カテゴリ選択時：月別推移と月別データ一覧
            const monthlyData = this.getMonthlyDataFromExpenses(expenses, useYearMonth);
            this.renderMonthlyChart(monthlyData, false, useYearMonth);
            this.renderMonthlyDataTable(expenses);
            // 取引明細を非表示
            document.getElementById('subcategoryDetailsCard').style.display = 'none';
        } else {
            // 全カテゴリ：月別推移（カテゴリ別）と月別データ一覧
            const monthlyCategoryData = this.getMonthlyCategoryDataFromExpenses(expenses, useYearMonth);
            this.renderMonthlyChart(monthlyCategoryData, true, useYearMonth); // カテゴリ別フラグと年月フラグ
            this.renderMonthlyDataTable(expenses);
            // 取引明細を非表示
            document.getElementById('subcategoryDetailsCard').style.display = 'none';
        }
    }

    // 支出データから月別データを生成
    getMonthlyDataFromExpenses(expenses, useYearMonth = false) {
        if (!useYearMonth) {
            // 年指定時：従来通り12ヶ月
            const monthlyData = Array(12).fill(0);
            expenses.forEach(e => {
                const date = new Date(e.date);
                const month = date.getMonth();
                monthlyData[month] += parseInt(e.amount);
            });
            return monthlyData;
        } else {
            // 全期間・月指定・カスタム期間：年月別
            const monthlyData = {};
            expenses.forEach(e => {
                const yearMonth = e.date.substring(0, 7); // YYYY-MM
                monthlyData[yearMonth] = (monthlyData[yearMonth] || 0) + parseInt(e.amount);
            });

            // ソートして配列に変換
            const sortedKeys = Object.keys(monthlyData).sort();
            const sortedData = {};
            sortedKeys.forEach(key => {
                sortedData[key] = monthlyData[key];
            });

            return sortedData;
        }
    }

    // 支出データからカテゴリ別月別データを生成
    getMonthlyCategoryDataFromExpenses(expenses, useYearMonth = false) {
        const categories = ['食品', '日用品', '外食費', '衣類', '家具・家電', '美容', '医療費', '交際費', 'レジャー', 'ガソリン・ETC', '光熱費', '通信費', '保険', '車関連【車検・税金・積立】', '税金', '経費', 'ローン', 'その他'];

        if (!useYearMonth) {
            // 年指定時：従来通り12ヶ月
            const categoryData = {};
            categories.forEach(cat => {
                categoryData[cat] = Array(12).fill(0);
            });

            expenses.forEach(e => {
                const date = new Date(e.date);
                const month = date.getMonth();
                const category = e.category;
                if (categoryData[category]) {
                    categoryData[category][month] += parseInt(e.amount);
                }
            });

            return categoryData;
        } else {
            // 全期間・月指定・カスタム期間：年月別
            const categoryData = {};
            categories.forEach(cat => {
                categoryData[cat] = {};
            });

            // 全ての年月を収集
            const allYearMonths = new Set();
            expenses.forEach(e => {
                const yearMonth = e.date.substring(0, 7); // YYYY-MM
                allYearMonths.add(yearMonth);
            });

            // ソートされた年月リスト
            const sortedYearMonths = Array.from(allYearMonths).sort();

            // 各カテゴリの年月データを初期化
            categories.forEach(cat => {
                sortedYearMonths.forEach(ym => {
                    categoryData[cat][ym] = 0;
                });
            });

            // データを集計
            expenses.forEach(e => {
                const yearMonth = e.date.substring(0, 7);
                const category = e.category;
                if (categoryData[category]) {
                    categoryData[category][yearMonth] = (categoryData[category][yearMonth] || 0) + parseInt(e.amount);
                }
            });

            return categoryData;
        }
    }

    // カテゴリ別統計表示
    renderCategoryStats(expenses) {
        const categoryTotals = this.manager.getCategoryTotals(expenses);
        const container = document.getElementById('categoryStats');

        container.innerHTML = '';
        Object.entries(categoryTotals).forEach(([category, amount]) => {
            const stat = document.createElement('div');
            stat.className = 'category-stat';
            stat.setAttribute('data-category', category);
            stat.innerHTML = `
                <h4>${category}</h4>
                <div class="amount">${amount.toLocaleString()}円</div>
            `;
            container.appendChild(stat);
        });
    }

    // 小項目別統計表示
    renderSubcategoryStats(expenses, category) {
        const container = document.getElementById('categoryStats');

        // 小項目別集計
        const subcategoryTotals = {};
        expenses.forEach(e => {
            const subcat = e.subcategory || '(未分類)';
            subcategoryTotals[subcat] = (subcategoryTotals[subcat] || 0) + parseInt(e.amount);
        });

        // 統計情報
        const total = expenses.reduce((sum, e) => sum + parseInt(e.amount), 0);
        const count = expenses.length;
        const average = count > 0 ? Math.round(total / count) : 0;

        container.innerHTML = `
            <div class="category-summary">
                <div class="summary-item">
                    <span>合計金額:</span>
                    <strong>${total.toLocaleString()}円</strong>
                </div>
                <div class="summary-item">
                    <span>件数:</span>
                    <strong>${count}件</strong>
                </div>
                <div class="summary-item">
                    <span>平均:</span>
                    <strong>${average.toLocaleString()}円</strong>
                </div>
            </div>
            <hr style="margin: 20px 0; border: none; border-top: 1px solid #ddd;">
        `;

        Object.entries(subcategoryTotals)
            .sort((a, b) => b[1] - a[1])
            .forEach(([subcat, amount]) => {
                const stat = document.createElement('div');
                stat.className = 'category-stat';
                const percentage = total > 0 ? Math.round((amount / total) * 100) : 0;
                stat.innerHTML = `
                    <h4>${subcat}</h4>
                    <div class="amount">${amount.toLocaleString()}円 (${percentage}%)</div>
                `;
                container.appendChild(stat);
            });
    }


    // 月別推移グラフ
    renderMonthlyChart(data, isCategoryBased = false, useYearMonth = false) {
        const ctx = document.getElementById('monthlyChart').getContext('2d');

        if (this.monthlyChart) {
            this.monthlyChart.destroy();
        }

        let datasets;
        let options;
        let labels;

        // ラベルの生成
        if (useYearMonth) {
            // 年月形式のラベル（例：2020-01, 2020-02...）
            if (isCategoryBased) {
                // カテゴリ別の場合、最初のカテゴリからキーを取得
                const firstCategory = Object.keys(data)[0];
                labels = Object.keys(data[firstCategory] || {});
            } else {
                // 単一データの場合
                labels = Object.keys(data);
            }
            // YYYY-MM形式を YY.MM 形式に変換
            labels = labels.map(ym => {
                const [year, month] = ym.split('-');
                return `${year.substring(2)}.${month}`;
            });
        } else {
            // 従来の12ヶ月形式
            labels = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
        }

        if (isCategoryBased) {
            // カテゴリ別の積み上げ棒グラフ
            const categoryColors = {
                '食品': 'rgba(76, 175, 80, 0.8)',
                '日用品': 'rgba(33, 150, 243, 0.8)',
                '外食費': 'rgba(255, 152, 0, 0.8)',
                '衣類': 'rgba(156, 39, 176, 0.8)',
                '家具・家電': 'rgba(0, 150, 136, 0.8)',
                '美容': 'rgba(233, 30, 99, 0.8)',
                '医療費': 'rgba(244, 67, 54, 0.8)',
                '交際費': 'rgba(255, 193, 7, 0.8)',
                'レジャー': 'rgba(3, 169, 244, 0.8)',
                'ガソリン・ETC': 'rgba(121, 85, 72, 0.8)',
                '光熱費': 'rgba(255, 87, 34, 0.8)',
                '通信費': 'rgba(103, 58, 183, 0.8)',
                '保険': 'rgba(0, 188, 212, 0.8)',
                '車関連【車検・税金・積立】': 'rgba(139, 195, 74, 0.8)',
                '税金': 'rgba(233, 30, 99, 0.8)',
                '経費': 'rgba(0, 150, 136, 0.8)',
                'ローン': 'rgba(103, 58, 183, 0.8)',
                'その他': 'rgba(158, 158, 158, 0.8)'
            };

            datasets = Object.entries(data).map(([category, monthlyData]) => ({
                label: category,
                data: useYearMonth ? Object.values(monthlyData) : monthlyData,
                backgroundColor: categoryColors[category] || 'rgba(158, 158, 158, 0.8)',
                borderWidth: 0
            }));

            options = {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        labels: {
                            boxWidth: 12,
                            padding: 8,
                            font: {
                                size: 10
                            }
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            title: function(tooltipItems) {
                                // tooltipItems[0]のdataIndexを使用してラベルを取得
                                if (tooltipItems && tooltipItems.length > 0) {
                                    const index = tooltipItems[0].dataIndex;
                                    return this.chart.data.labels[index];
                                }
                                return '';
                            },
                            label: function(context) {
                                const label = context.dataset.label || '';
                                const value = context.parsed.y || 0;
                                return label + ': ' + value.toLocaleString() + '円';
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        ticks: {
                            maxRotation: 45,
                            minRotation: 45
                        }
                    },
                    y: {
                        stacked: true,
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return value.toLocaleString() + '円';
                            }
                        }
                    }
                }
            };
        } else {
            // 単一データの棒グラフ
            datasets = [{
                label: '支出額',
                data: useYearMonth ? Object.values(data) : data,
                backgroundColor: '#667eea',
                borderColor: '#667eea',
                borderWidth: 1
            }];

            options = {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            title: function(tooltipItems) {
                                // tooltipItems[0]のdataIndexを使用してラベルを取得
                                if (tooltipItems && tooltipItems.length > 0) {
                                    const index = tooltipItems[0].dataIndex;
                                    return this.chart.data.labels[index];
                                }
                                return '';
                            },
                            label: function(context) {
                                const value = context.parsed.y || 0;
                                return '支出額: ' + value.toLocaleString() + '円';
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        ticks: {
                            maxRotation: 45,
                            minRotation: 45
                        }
                    },
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return value.toLocaleString() + '円';
                            }
                        }
                    }
                }
            };
        }

        this.monthlyChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: options
        });
    }

    // 月別データ一覧テーブル
    renderMonthlyDataTable(expenses) {
        const container = document.getElementById('monthlyDataTable');

        // 月ごとにデータを集計
        const monthlyDetails = Array(12).fill(null).map(() => ({
            count: 0,
            total: 0
        }));

        expenses.forEach(e => {
            const date = new Date(e.date);
            const month = date.getMonth();
            monthlyDetails[month].count++;
            monthlyDetails[month].total += parseInt(e.amount);
        });

        // テーブルHTML生成
        let html = `
            <table class="monthly-data-table">
                <thead>
                    <tr>
                        <th>月</th>
                        <th>件数</th>
                        <th>合計金額</th>
                        <th>平均</th>
                    </tr>
                </thead>
                <tbody>
        `;

        monthlyDetails.forEach((detail, index) => {
            const monthName = `${index + 1}月`;
            const count = detail.count;
            const total = detail.total;
            const average = count > 0 ? Math.round(total / count) : 0;

            if (count > 0) {
                html += `
                    <tr>
                        <td>${monthName}</td>
                        <td>${count}件</td>
                        <td>${total.toLocaleString()}円</td>
                        <td>${average.toLocaleString()}円</td>
                    </tr>
                `;
            }
        });

        html += `
                </tbody>
            </table>
        `;

        container.innerHTML = html;
    }

    // 小項目別取引明細テーブル
    renderSubcategoryDetails(expenses, subcategory) {
        const container = document.getElementById('subcategoryDetailsTable');
        const card = document.getElementById('subcategoryDetailsCard');
        const title = document.getElementById('subcategoryDetailsTitle');

        if (!expenses || expenses.length === 0) {
            card.style.display = 'none';
            return;
        }

        // カードを表示
        card.style.display = 'block';
        title.textContent = `${subcategory} - 取引明細（全${expenses.length}件）`;

        // 日付でソート（新しい順）
        const sortedExpenses = [...expenses].sort((a, b) => {
            return new Date(b.date) - new Date(a.date);
        });

        // テーブルHTML生成
        let html = `
            <table class="subcategory-details-table">
                <thead>
                    <tr>
                        <th>日付</th>
                        <th>カテゴリ</th>
                        <th>小項目</th>
                        <th>金額</th>
                        <th>場所</th>
                        <th>商品名・メモ</th>
                    </tr>
                </thead>
                <tbody>
        `;

        sortedExpenses.forEach(expense => {
            html += `
                <tr>
                    <td>${expense.date}</td>
                    <td><span class="badge" data-category="${expense.category}">${expense.category}</span></td>
                    <td>${expense.subcategory || '-'}</td>
                    <td style="text-align: right; font-weight: 500;">¥${parseInt(expense.amount).toLocaleString()}</td>
                    <td>${expense.place || '-'}</td>
                    <td>${expense.description || '-'}</td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
        `;

        container.innerHTML = html;

        // バッジの色を適用
        setTimeout(() => {
            const badges = container.querySelectorAll('.badge');
            badges.forEach(badge => {
                const category = badge.getAttribute('data-category');
                const color = this.getCategoryColor(category);
                badge.style.backgroundColor = color;
            });
        }, 0);
    }

    // 年別カテゴリ推移表
    renderYearlyTable(year) {
        // 月のリストを生成（1月〜12月）
        const months = [];
        for (let month = 1; month <= 12; month++) {
            months.push({
                year,
                month,
                key: `${year}-${String(month).padStart(2, '0')}`,
                label: `${String(year).substring(2)}.${String(month).padStart(2, '0')}`
            });
        }

        // カテゴリリスト
        const categories = ['食品', '日用品', '外食費', '衣類', '家具・家電', '美容', '医療費', '交際費', 'レジャー', 'ガソリン・ETC', '光熱費', '通信費', '保険', '車関連【車検・税金・積立】', '税金', '経費', 'ローン', 'その他'];

        // カテゴリごと・月ごとの集計
        const data = {};
        categories.forEach(cat => {
            data[cat] = {};
            months.forEach(m => {
                data[cat][m.key] = 0;
            });
        });

        // データを集計
        this.manager.getAllExpenses().forEach(expense => {
            const date = expense.date;
            const yearMonth = date.substring(0, 7); // YYYY-MM
            if (data[expense.category] && data[expense.category][yearMonth] !== undefined) {
                data[expense.category][yearMonth] += parseInt(expense.amount || 0);
            }
        });

        // テーブルヘッダーを生成（ソート可能）
        const thead = document.getElementById('yearlyTableHead');
        thead.innerHTML = `
            <tr>
                <th class="sticky-col sortable-header" data-reset="true" title="クリックで元の順序に戻す" style="cursor: pointer;">カテゴリ</th>
                ${months.map((m, idx) => `<th class="sortable-header" data-month-index="${idx}" title="クリックでソート">${m.label}</th>`).join('')}
            </tr>
        `;

        // カテゴリヘッダーにリセットイベントを追加
        const categoryHeader = thead.querySelector('th[data-reset="true"]');
        categoryHeader.addEventListener('click', () => {
            this.renderYearlyTable(year); // 元の順序で再描画
        });

        // 月ヘッダーにソートイベントを追加
        thead.querySelectorAll('.sortable-header[data-month-index]').forEach(th => {
            th.addEventListener('click', () => {
                const monthIndex = parseInt(th.getAttribute('data-month-index'));
                this.sortYearlyTableByMonth(categories, data, months, monthIndex, year);
            });
        });

        // テーブルボディを生成
        const tbody = document.getElementById('yearlyTableBody');
        tbody.innerHTML = '';

        categories.forEach(category => {
            const row = document.createElement('tr');

            // その行の最大値を見つける
            const amounts = months.map(m => data[category][m.key]);
            const maxAmount = Math.max(...amounts);

            const cells = months.map(m => {
                const amount = data[category][m.key];
                const displayAmount = amount > 0 ? `¥${amount.toLocaleString()}` : '';

                if (amount > 0 && maxAmount > 0) {
                    // 金額の比率を計算（最大値に対する割合）
                    const ratio = amount / maxAmount;
                    // 比率に応じて色の濃さを調整（0.3〜1.0）
                    const alpha = 0.3 + (ratio * 0.7);
                    const bgColor = this.getCategoryColorWithAlpha(category, alpha);
                    return `<td style="background:${bgColor}">${displayAmount}</td>`;
                } else {
                    return `<td>${displayAmount}</td>`;
                }
            });

            row.innerHTML = `
                <td class="sticky-col"><span class="badge" data-category="${category}">${category}</span></td>
                ${cells.join('')}
            `;
            tbody.appendChild(row);
        });

        // 合計行
        const totalRow = document.createElement('tr');
        totalRow.style.fontWeight = '500';
        totalRow.style.borderTop = '2px solid #202124';
        const totalCells = months.map(m => {
            const total = categories.reduce((sum, cat) => sum + data[cat][m.key], 0);
            return `<td>¥${total.toLocaleString()}</td>`;
        });
        totalRow.innerHTML = `
            <td class="sticky-col">合計</td>
            ${totalCells.join('')}
        `;
        tbody.appendChild(totalRow);

        // Topix行（出来事）
        const topixRow = document.createElement('tr');
        topixRow.style.backgroundColor = '#fff9e6';
        topixRow.style.borderTop = '3px solid #ffc107';
        topixRow.style.height = '60px';
        const topixCells = months.map(m => {
            const memo = this.manager.getMemo(m.key);
            const events = memo.events ? memo.events.split('\n').filter(e => e.trim()).join('、') : '';
            return `<td style="font-size: 11px; padding: 8px; white-space: normal; overflow: hidden; line-height: 1.4; vertical-align: top;">${events}</td>`;
        });
        topixRow.innerHTML = `
            <td class="sticky-col" style="background: #fff9e6; vertical-align: top;"><strong>Topix</strong></td>
            ${topixCells.join('')}
        `;
        tbody.appendChild(topixRow);

        // 予定行
        const plansRow = document.createElement('tr');
        plansRow.style.backgroundColor = '#e6f4ff';
        plansRow.style.height = '60px';
        const plansCells = months.map(m => {
            const memo = this.manager.getMemo(m.key);
            const plans = memo.plans ? memo.plans.split('\n').filter(p => p.trim()).join('、') : '';
            return `<td style="font-size: 11px; padding: 8px; white-space: normal; overflow: hidden; line-height: 1.4; vertical-align: top;">${plans}</td>`;
        });
        plansRow.innerHTML = `
            <td class="sticky-col" style="background: #e6f4ff; vertical-align: top;"><strong>予定</strong></td>
            ${plansCells.join('')}
        `;
        tbody.appendChild(plansRow);

        // タイトル更新
        document.getElementById('yearlyTableTitle').textContent = `${year}年 カテゴリ別推移`;
    }

    // 年別表を特定の月でソート
    sortYearlyTableByMonth(categories, data, months, monthIndex, year) {
        const month = months[monthIndex];

        // 指定された月の金額でカテゴリをソート（降順）
        const sortedCategories = [...categories].sort((a, b) => {
            const amountA = data[a][month.key] || 0;
            const amountB = data[b][month.key] || 0;
            return amountB - amountA; // 降順
        });

        // テーブルボディを再生成
        const tbody = document.getElementById('yearlyTableBody');
        tbody.innerHTML = '';

        // 選択された月の合計を計算
        let selectedMonthTotal = 0;

        sortedCategories.forEach(category => {
            const row = document.createElement('tr');

            // その行の最大値を見つける
            const amounts = months.map(m => data[category][m.key]);
            const maxAmount = Math.max(...amounts);

            const cells = months.map((m, idx) => {
                const amount = data[category][m.key];
                const displayAmount = amount > 0 ? `¥${amount.toLocaleString()}` : '';

                // 選択された月の場合、合計に加算
                if (idx === monthIndex) {
                    selectedMonthTotal += amount;
                }

                if (amount > 0 && maxAmount > 0) {
                    const ratio = amount / maxAmount;
                    const alpha = 0.3 + (ratio * 0.7);
                    const bgColor = this.getCategoryColorWithAlpha(category, alpha);

                    // 選択された列を強調表示
                    const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8;' : '';
                    return `<td style="background:${bgColor}; ${highlight}">${displayAmount}</td>`;
                } else {
                    const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8;' : '';
                    return `<td style="${highlight}">${displayAmount}</td>`;
                }
            });

            row.innerHTML = `
                <td class="sticky-col"><span class="badge" data-category="${category}">${category}</span></td>
                ${cells.join('')}
            `;
            tbody.appendChild(row);
        });

        // 合計行
        const totalRow = document.createElement('tr');
        totalRow.style.fontWeight = '500';
        totalRow.style.borderTop = '2px solid #202124';
        const totalCells = months.map((m, idx) => {
            const total = categories.reduce((sum, cat) => sum + data[cat][m.key], 0);
            const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8; background: #e8f0fe;' : '';
            return `<td style="${highlight}">¥${total.toLocaleString()}</td>`;
        });
        totalRow.innerHTML = `
            <td class="sticky-col">合計</td>
            ${totalCells.join('')}
        `;
        tbody.appendChild(totalRow);

        // 選択された月の合計を表示するメッセージ
        const infoRow = document.createElement('tr');
        infoRow.style.background = '#e8f0fe';
        infoRow.style.fontWeight = '500';
        infoRow.innerHTML = `
            <td class="sticky-col" colspan="${months.length + 1}" style="text-align: center; padding: 12px;">
                ${month.year}年${month.month}月でソート済み - 合計: ¥${selectedMonthTotal.toLocaleString()}
            </td>
        `;
        tbody.appendChild(infoRow);

        // Topix行（出来事）
        const topixRow = document.createElement('tr');
        topixRow.style.backgroundColor = '#fff9e6';
        topixRow.style.borderTop = '3px solid #ffc107';
        topixRow.style.height = '60px';
        const topixCells = months.map((m, idx) => {
            const memo = this.manager.getMemo(m.key);
            const events = memo.events ? memo.events.split('\n').filter(e => e.trim()).join('、') : '';
            const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8;' : '';
            return `<td style="font-size: 11px; padding: 8px; white-space: normal; overflow: hidden; line-height: 1.4; vertical-align: top; ${highlight}">${events}</td>`;
        });
        topixRow.innerHTML = `
            <td class="sticky-col" style="background: #fff9e6; vertical-align: top;"><strong>Topix</strong></td>
            ${topixCells.join('')}
        `;
        tbody.appendChild(topixRow);

        // 予定行
        const plansRow = document.createElement('tr');
        plansRow.style.backgroundColor = '#e6f4ff';
        plansRow.style.height = '60px';
        const plansCells = months.map((m, idx) => {
            const memo = this.manager.getMemo(m.key);
            const plans = memo.plans ? memo.plans.split('\n').filter(p => p.trim()).join('、') : '';
            const highlight = idx === monthIndex ? 'border: 2px solid #1a73e8;' : '';
            return `<td style="font-size: 11px; padding: 8px; white-space: normal; overflow: hidden; line-height: 1.4; vertical-align: top; ${highlight}">${plans}</td>`;
        });
        plansRow.innerHTML = `
            <td class="sticky-col" style="background: #e6f4ff; vertical-align: top;"><strong>予定</strong></td>
            ${plansCells.join('')}
        `;
        tbody.appendChild(plansRow);

        // タイトル更新
        document.getElementById('yearlyTableTitle').textContent = `${year}年 カテゴリ別推移（${month.month}月でソート済み）`;
    }

    // データエクスポート
    exportData() {
        // 現在のアクティブタブを取得
        const activeTab = document.querySelector('.tab-btn.active').getAttribute('data-tab');

        let data, filename, csv;

        if (activeTab === 'list') {
            // 一覧タブ：フィルター適用後のデータをエクスポート
            data = this.filteredExpenses;
            filename = `家計簿_一覧_${new Date().toISOString().split('T')[0]}.csv`;
            csv = this.convertToCSV(data);
            this.showMessage(`フィルター適用後のデータ（${data.length}件）をエクスポートしました`);
        } else if (activeTab === 'analysis') {
            // 分析タブ：分析結果をエクスポート
            if (!this.analysisData || !this.analysisData.expenses) {
                this.showMessage('分析データがありません');
                return;
            }
            csv = this.convertAnalysisToCSV(this.analysisData);
            filename = `家計簿_分析_${new Date().toISOString().split('T')[0]}.csv`;
            this.showMessage(`分析結果（${this.analysisData.expenses.length}件）をエクスポートしました`);
        } else {
            // その他のタブ：全データをエクスポート
            data = this.manager.getAllExpenses();
            filename = `家計簿_全データ_${new Date().toISOString().split('T')[0]}.csv`;
            csv = this.convertToCSV(data);
            this.showMessage(`全データ（${data.length}件）をエクスポートしました`);
        }

        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);

        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.click();
    }

    // CSVに変換
    convertToCSV(data) {
        const headers = ['日付', 'カテゴリ', '小項目', '金額', '場所', '商品名・メモ'];
        const rows = data.map(e => [
            e.date,
            e.category,
            e.subcategory || '',
            e.amount,
            e.place || '',
            e.description || ''
        ]);

        const csv = [headers, ...rows]
            .map(row => row.map(cell => `"${cell}"`).join(','))
            .join('\n');

        return '\uFEFF' + csv; // BOM追加でExcelで文字化け防止
    }

    // 分析データをCSVに変換
    convertAnalysisToCSV(analysisData) {
        const { expenses, dateRangeType, year, month, startDate, endDate, category, subcategory } = analysisData;

        // ヘッダー情報
        let headerInfo = [];
        headerInfo.push(['=== 分析条件 ===']);

        // 期間情報
        if (dateRangeType === 'year' && year) {
            headerInfo.push(['期間', `${year}年`]);
        } else if (dateRangeType === 'month' && year && month) {
            headerInfo.push(['期間', `${year}年${month}月`]);
        } else if (dateRangeType === 'custom' && startDate && endDate) {
            headerInfo.push(['期間', `${startDate} 〜 ${endDate}`]);
        }

        // カテゴリ・小項目情報
        if (category) {
            headerInfo.push(['カテゴリ', category]);
        }
        if (subcategory) {
            headerInfo.push(['小項目', subcategory]);
        }

        // 集計情報
        const total = expenses.reduce((sum, e) => sum + parseInt(e.amount || 0), 0);
        const count = expenses.length;
        const average = count > 0 ? Math.round(total / count) : 0;

        headerInfo.push(['', '']);
        headerInfo.push(['=== 集計結果 ===']);
        headerInfo.push(['件数', `${count}件`]);
        headerInfo.push(['合計金額', `${total.toLocaleString()}円`]);
        headerInfo.push(['平均金額', `${average.toLocaleString()}円`]);

        // カテゴリ別集計（カテゴリが選択されていない場合）
        if (!category) {
            headerInfo.push(['', '']);
            headerInfo.push(['=== カテゴリ別集計 ===']);
            headerInfo.push(['カテゴリ', '金額', '割合']);

            const categoryTotals = {};
            expenses.forEach(e => {
                categoryTotals[e.category] = (categoryTotals[e.category] || 0) + parseInt(e.amount || 0);
            });

            Object.entries(categoryTotals)
                .sort((a, b) => b[1] - a[1])
                .forEach(([cat, amount]) => {
                    const percentage = ((amount / total) * 100).toFixed(1);
                    headerInfo.push([cat, `${amount.toLocaleString()}円`, `${percentage}%`]);
                });
        }

        // 小項目別集計（カテゴリが選択されているが小項目が選択されていない場合）
        if (category && !subcategory) {
            headerInfo.push(['', '']);
            headerInfo.push(['=== 小項目別集計 ===']);
            headerInfo.push(['小項目', '金額', '割合']);

            const subcategoryTotals = {};
            expenses.forEach(e => {
                const sub = e.subcategory || '(未分類)';
                subcategoryTotals[sub] = (subcategoryTotals[sub] || 0) + parseInt(e.amount || 0);
            });

            Object.entries(subcategoryTotals)
                .sort((a, b) => b[1] - a[1])
                .forEach(([sub, amount]) => {
                    const percentage = ((amount / total) * 100).toFixed(1);
                    headerInfo.push([sub, `${amount.toLocaleString()}円`, `${percentage}%`]);
                });
        }

        // 明細データ
        headerInfo.push(['', '']);
        headerInfo.push(['=== 明細データ ===']);
        const detailHeaders = ['日付', 'カテゴリ', '小項目', '金額', '場所', '商品名・メモ'];
        headerInfo.push(detailHeaders);

        const detailRows = expenses.map(e => [
            e.date,
            e.category,
            e.subcategory || '',
            e.amount,
            e.place || '',
            e.description || ''
        ]);

        const allRows = [...headerInfo, ...detailRows];
        const csv = allRows
            .map(row => row.map(cell => `"${cell}"`).join(','))
            .join('\n');

        return '\uFEFF' + csv; // BOM追加でExcelで文字化け防止
    }

    // 重複データを削除
    removeDuplicateData() {
        // まず重複を検出
        const duplicates = this.manager.findDuplicates();

        if (duplicates.length === 0) {
            alert('重複データは見つかりませんでした。');
            return;
        }

        // 重複データの詳細を表示
        let message = `${duplicates.length}グループの重複データが見つかりました。\n\n`;
        message += '以下のデータが重複しています：\n\n';

        duplicates.slice(0, 5).forEach((dup, index) => {
            const data = dup.data;
            message += `${index + 1}. ${data.date} - ${data.category}`;
            if (data.subcategory) message += ` (${data.subcategory})`;
            message += ` - ${data.amount}円`;
            if (data.place) message += ` - ${data.place}`;
            if (data.description) message += ` - ${data.description}`;
            message += ` (${dup.ids.length}件)\n`;
        });

        if (duplicates.length > 5) {
            message += `\n... 他 ${duplicates.length - 5}グループ\n`;
        }

        message += '\n各グループの最初の1件を残して、残りを削除しますか？';

        if (confirm(message)) {
            const result = this.manager.removeDuplicates();
            this.renderExpenseList();
            alert(`重複削除完了\n\n削除件数: ${result.removedCount}件\n（${result.duplicateGroups}グループ）`);
        }
    }

    // データインポート
    importData() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv';

        input.onchange = async (e) => {
            const file = e.target.files[0];
            if (!file) return;

            // 既存データがある場合は、追加するか置き換えるかを確認
            const existingData = this.manager.getAllExpenses();
            let replaceMode = false;

            if (existingData.length > 0) {
                const message = `現在${existingData.length}件のデータがあります。\n\nOK: 既存データに追加\nキャンセル: 既存データを削除して新規インポート`;
                const addToExisting = confirm(message);
                replaceMode = !addToExisting;

                if (replaceMode) {
                    // 本当に削除するか再確認
                    const confirmDelete = confirm('既存データを削除してもよろしいですか？この操作は取り消せません。');
                    if (!confirmDelete) {
                        return; // インポートをキャンセル
                    }
                }
            }

            const reader = new FileReader();

            reader.onload = async (event) => {
                try {
                    const csv = event.target.result;
                    await this.processCSVImport(csv, replaceMode);
                } catch (error) {
                    console.error('Import error:', error);
                    alert('インポートに失敗しました: ' + error.message);
                }
            };

            reader.readAsText(file, 'UTF-8');
        };

        input.click();
    }

    // CSV解析とインポート処理（バッチ処理対応）
    async processCSVImport(csv, replaceMode = false) {
        // プログレス表示用のオーバーレイを作成
        const overlay = this.createProgressOverlay();
        document.body.appendChild(overlay);

        try {
            // 既存データを削除する場合
            if (replaceMode) {
                this.manager.expenses = [];
                this.manager.saveExpenses();
                console.log('既存データを削除しました');
            }

            // BOMを除去
            csv = csv.replace(/^\uFEFF/, '');

            // CSVをパース
            const lines = csv.split('\n');
            const headers = this.parseCSVLine(lines[0]);

            // データ行を取得（ヘッダーと空行を除く）
            const dataLines = lines.slice(1).filter(line => line.trim());

            console.log(`インポート開始: ${dataLines.length}件のデータ`);

            let successCount = 0;
            let errorCount = 0;
            const allExpenses = []; // 一時的に全データを保持

            // 全行を解析してメモリに保持
            for (let i = 0; i < dataLines.length; i++) {
                const line = dataLines[i];

                // 進捗を更新（解析フェーズ）
                if (i % 100 === 0) {
                    const progress = Math.round((i / dataLines.length) * 50); // 50%まで
                    this.updateProgress(overlay, progress, `${i}/${dataLines.length}件解析中...`);
                    await new Promise(resolve => setTimeout(resolve, 0));
                }

                try {
                    const values = this.parseCSVLine(line);

                    if (values.length >= 4) {
                        const [date, category, subcategory, amount, place, description] = values;

                        if (date && category && amount) {
                            allExpenses.push({
                                date: date.trim(),
                                category: category.trim(),
                                subcategory: (subcategory || '').trim(),
                                amount: amount.trim(),
                                place: (place || '').trim(),
                                description: (description || '').trim()
                            });
                            successCount++;
                        }
                    }
                } catch (error) {
                    console.error(`行 ${i + 2} でエラー:`, error);
                    errorCount++;
                }
            }

            // 進捗を更新（保存フェーズ）
            this.updateProgress(overlay, 75, `${allExpenses.length}件のデータを保存中...`);
            await new Promise(resolve => setTimeout(resolve, 100));

            // 一括でデータベースに追加（最後に1回だけlocalStorageに保存）
            try {
                this.manager.addExpensesBatch(allExpenses);
                console.log(`保存完了: ${allExpenses.length}件`);
            } catch (storageError) {
                overlay.remove();
                if (storageError.message.includes('容量超過')) {
                    alert(`データ容量エラー: ${allExpenses.length}件のデータは大きすぎます。\n\n推奨: 年別にCSVファイルを分けてインポートしてください。`);
                }
                throw storageError;
            }

            // 完了
            this.updateProgress(overlay, 100, '完了！');
            await new Promise(resolve => setTimeout(resolve, 500));

            // オーバーレイを削除
            overlay.remove();

            // 画面を更新
            this.renderExpenseList();

            // 結果を表示
            const message = `インポート完了: ${successCount}件成功${errorCount > 0 ? `, ${errorCount}件失敗` : ''}`;
            this.showMessage(message);
            console.log(message);

        } catch (error) {
            if (overlay && overlay.parentElement) {
                overlay.remove();
            }
            throw error;
        }
    }

    // CSV行をパース（ダブルクォートで囲まれた値に対応）
    parseCSVLine(line) {
        const values = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];

            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    // エスケープされたダブルクォート
                    current += '"';
                    i++;
                } else {
                    // クォートの開始/終了
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                // カンマ区切り（クォート外）
                values.push(current);
                current = '';
            } else {
                current += char;
            }
        }

        // 最後の値を追加
        values.push(current);

        return values;
    }

    // プログレス表示用のオーバーレイを作成
    createProgressOverlay() {
        const overlay = document.createElement('div');
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.8);
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            z-index: 10000;
        `;

        overlay.innerHTML = `
            <div style="background: white; padding: 30px; border-radius: 10px; text-align: center; min-width: 300px;">
                <h3 style="margin: 0 0 20px 0;">データをインポート中...</h3>
                <div style="width: 100%; height: 20px; background: #e0e0e0; border-radius: 10px; overflow: hidden;">
                    <div id="progressBar" style="width: 0%; height: 100%; background: linear-gradient(90deg, #6366f1, #8b5cf6); transition: width 0.3s;"></div>
                </div>
                <p id="progressText" style="margin: 10px 0 0 0; color: #666;">0%</p>
            </div>
        `;

        return overlay;
    }

    // プログレス表示を更新
    updateProgress(overlay, percent, text) {
        const progressBar = overlay.querySelector('#progressBar');
        const progressText = overlay.querySelector('#progressText');

        if (progressBar) progressBar.style.width = percent + '%';
        if (progressText) progressText.textContent = text || (percent + '%');
    }

    // メッセージ表示
    showMessage(text) {
        // 簡易的なトースト通知
        const toast = document.createElement('div');
        toast.textContent = text;
        toast.style.cssText = `
            position: fixed;
            bottom: 30px;
            right: 30px;
            background: #28a745;
            color: white;
            padding: 15px 25px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            z-index: 10000;
            animation: slideIn 0.3s;
        `;

        document.body.appendChild(toast);

        setTimeout(() => {
            toast.style.animation = 'slideOut 0.3s';
            setTimeout(() => toast.remove(), 300);
        }, 2000);
    }

    // ========================================
    // 一括入力モード
    // ========================================

    // モード切り替え
    switchMode(mode) {
        const normalMode = document.getElementById('normalInputMode');
        const bulkMode = document.getElementById('bulkInputMode');
        const normalBtn = document.getElementById('normalModeBtn');
        const bulkBtn = document.getElementById('bulkModeBtn');

        if (mode === 'normal') {
            normalMode.style.display = 'block';
            bulkMode.style.display = 'none';
            normalBtn.classList.add('active');
            bulkBtn.classList.remove('active');
        } else {
            normalMode.style.display = 'none';
            bulkMode.style.display = 'block';
            normalBtn.classList.remove('active');
            bulkBtn.classList.add('active');

            // 共通日付を今日に設定
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('bulkCommonDate').value = today;

            // 一括入力モードに切り替えたら、初期行を追加
            if (this.bulkInputRows.length === 0) {
                // カウンターベースのIDで重複なく追加
                this.addBulkInputRow();
                this.addBulkInputRow();
                this.addBulkInputRow();
            }
        }
    }

    // 一括入力行を追加
    addBulkInputRow() {
        const tbody = document.getElementById('bulkInputTableBody');
        this.bulkRowIdCounter++;
        const rowId = this.bulkRowIdCounter;
        const today = new Date().toISOString().split('T')[0];

        // 既存の1行目からカテゴリを取得（2行目以降の場合）
        const firstRow = tbody.querySelector('tr:first-child');
        let defaultCategory = '';
        if (firstRow && this.bulkInputRows.length > 0) {
            defaultCategory = firstRow.querySelector('.bulk-category').value;
        }

        const row = document.createElement('tr');
        row.dataset.rowId = rowId;
        row.innerHTML = `
            <td><input type="date" class="bulk-date" value="${today}"></td>
            <td><input type="text" class="bulk-place" placeholder="場所"></td>
            <td><input type="number" class="bulk-amount" min="0" placeholder="金額"></td>
            <td><button class="tax-btn" onclick="ui.applyTax(${rowId})">税込</button></td>
            <td><input type="text" class="bulk-description" placeholder="商品名"></td>
            <td>
                <select class="bulk-category" onchange="ui.updateSubcategoryOptions(${rowId})">
                    <option value="">選択</option>
                    <option value="食品" ${defaultCategory === '食品' ? 'selected' : ''}>食品</option>
                    <option value="日用品" ${defaultCategory === '日用品' ? 'selected' : ''}>日用品</option>
                    <option value="外食費" ${defaultCategory === '外食費' ? 'selected' : ''}>外食費</option>
                    <option value="衣類" ${defaultCategory === '衣類' ? 'selected' : ''}>衣類</option>
                    <option value="家具・家電" ${defaultCategory === '家具・家電' ? 'selected' : ''}>家具・家電</option>
                    <option value="美容" ${defaultCategory === '美容' ? 'selected' : ''}>美容</option>
                    <option value="医療費" ${defaultCategory === '医療費' ? 'selected' : ''}>医療費</option>
                    <option value="交際費" ${defaultCategory === '交際費' ? 'selected' : ''}>交際費</option>
                    <option value="レジャー" ${defaultCategory === 'レジャー' ? 'selected' : ''}>レジャー</option>
                    <option value="ガソリン・ETC" ${defaultCategory === 'ガソリン・ETC' ? 'selected' : ''}>ガソリン・ETC</option>
                    <option value="光熱費" ${defaultCategory === '光熱費' ? 'selected' : ''}>光熱費</option>
                    <option value="通信費" ${defaultCategory === '通信費' ? 'selected' : ''}>通信費</option>
                    <option value="保険" ${defaultCategory === '保険' ? 'selected' : ''}>保険</option>
                    <option value="車関連【車検・税金・積立】" ${defaultCategory === '車関連【車検・税金・積立】' ? 'selected' : ''}>車関連【車検・税金・積立】</option>
                    <option value="税金" ${defaultCategory === '税金' ? 'selected' : ''}>税金</option>
                    <option value="経費" ${defaultCategory === '経費' ? 'selected' : ''}>経費</option>
                    <option value="ローン" ${defaultCategory === 'ローン' ? 'selected' : ''}>ローン</option>
                    <option value="その他" ${defaultCategory === 'その他' ? 'selected' : ''}>その他</option>
                </select>
            </td>
            <td>
                <select class="bulk-subcategory" data-row-id="${rowId}">
                    <option value="">選択</option>
                </select>
            </td>
            <td><input type="text" class="bulk-notes" placeholder="メモ"></td>
            <td><button class="delete-row-btn" onclick="ui.deleteBulkInputRow(${rowId})">削除</button></td>
        `;

        tbody.appendChild(row);
        this.bulkInputRows.push(rowId);

        // 小項目の選択肢を更新（カテゴリが選択されているかどうかに関わらず実行）
        this.updateSubcategoryOptions(rowId);

        // 1行目の場合、カテゴリ変更時に2行目以降も更新するイベントリスナーを追加
        // this.bulkInputRows.lengthはpush後なので、=== 1の時は1行目
        if (this.bulkInputRows.length === 1) {
            const categorySelect = row.querySelector('.bulk-category');
            categorySelect.addEventListener('change', () => {
                this.propagateCategoryToOtherRows();
            });
        }
    }

    // 複数行を一度に追加
    addBulkInputRows(count) {
        for (let i = 0; i < count; i++) {
            // カウンターベースのIDで重複なく追加
            this.addBulkInputRow();
        }
    }

    // 通常入力の消費税を適用（1.1倍）
    applyNormalTax() {
        const amountInput = document.getElementById('amount');
        const currentValue = parseFloat(amountInput.value);

        if (!isNaN(currentValue) && currentValue > 0) {
            const taxIncluded = Math.round(currentValue * 1.1);
            amountInput.value = taxIncluded;
        }
    }

    // 一括入力の消費税を適用（1.1倍）
    applyTax(rowId) {
        const row = document.querySelector(`tr[data-row-id="${rowId}"]`);
        const amountInput = row.querySelector('.bulk-amount');
        const currentValue = parseFloat(amountInput.value);

        if (!isNaN(currentValue) && currentValue > 0) {
            const taxIncluded = Math.round(currentValue * 1.1);
            amountInput.value = taxIncluded;
        }
    }

    // 一括入力行を削除
    deleteBulkInputRow(rowId) {
        const row = document.querySelector(`tr[data-row-id="${rowId}"]`);
        if (row) {
            row.remove();
            this.bulkInputRows = this.bulkInputRows.filter(id => id !== rowId);
        }
    }

    // カテゴリに応じて小項目の選択肢を更新
    updateSubcategoryOptions(rowId) {
        const row = document.querySelector(`tr[data-row-id="${rowId}"]`);
        const categorySelect = row.querySelector('.bulk-category');
        const subcategorySelect = row.querySelector('.bulk-subcategory');
        const selectedCategory = categorySelect.value;

        // 現在選択されている小項目を保持
        const currentSubcategory = subcategorySelect.value;

        // 小項目の選択肢をクリア
        subcategorySelect.innerHTML = '<option value="">選択</option>';

        // カテゴリが選択されている場合、対応する小項目を追加
        if (selectedCategory && this.subcategoryMaster[selectedCategory]) {
            this.subcategoryMaster[selectedCategory].forEach(subcat => {
                const option = document.createElement('option');
                option.value = subcat;
                option.textContent = subcat;
                if (subcat === currentSubcategory) {
                    option.selected = true;
                }
                subcategorySelect.appendChild(option);
            });
        }
    }

    // 1行目のカテゴリを他の行にも反映
    propagateCategoryToOtherRows() {
        const tbody = document.getElementById('bulkInputTableBody');
        const firstRow = tbody.querySelector('tr:first-child');
        if (!firstRow) return;

        const selectedCategory = firstRow.querySelector('.bulk-category').value;

        // 2行目以降の全ての行に同じカテゴリを設定
        const allRows = tbody.querySelectorAll('tr');
        allRows.forEach((row, index) => {
            if (index > 0) { // 1行目以外
                const categorySelect = row.querySelector('.bulk-category');
                const rowId = row.dataset.rowId;

                // カテゴリを設定
                categorySelect.value = selectedCategory;

                // 小項目の選択肢を更新（rowIdは文字列なのでそのまま使用）
                this.updateSubcategoryOptions(rowId);
            }
        });
    }

    // 通常入力フォームの小項目を更新
    updateNormalSubcategoryOptions() {
        const categorySelect = document.getElementById('category');
        const subcategorySelect = document.getElementById('subcategory');
        const selectedCategory = categorySelect.value;

        // 現在選択されている小項目を保持
        const currentSubcategory = subcategorySelect.value;

        // 小項目の選択肢をクリア
        subcategorySelect.innerHTML = '<option value="">選択してください</option>';

        // カテゴリが選択されている場合、対応する小項目を追加
        if (selectedCategory && this.subcategoryMaster[selectedCategory]) {
            this.subcategoryMaster[selectedCategory].forEach(subcat => {
                const option = document.createElement('option');
                option.value = subcat;
                option.textContent = subcat;
                if (subcat === currentSubcategory) {
                    option.selected = true;
                }
                subcategorySelect.appendChild(option);
            });
        }
    }

    // 編集モーダルの小項目を更新
    updateEditSubcategoryOptions() {
        const categorySelect = document.getElementById('editCategory');
        const subcategorySelect = document.getElementById('editSubcategory');
        const selectedCategory = categorySelect.value;

        // 現在選択されている小項目を保持
        const currentSubcategory = subcategorySelect.value;

        // 小項目の選択肢をクリア
        subcategorySelect.innerHTML = '<option value="">選択してください</option>';

        // カテゴリが選択されている場合、対応する小項目を追加
        if (selectedCategory && this.subcategoryMaster[selectedCategory]) {
            this.subcategoryMaster[selectedCategory].forEach(subcat => {
                const option = document.createElement('option');
                option.value = subcat;
                option.textContent = subcat;
                if (subcat === currentSubcategory) {
                    option.selected = true;
                }
                subcategorySelect.appendChild(option);
            });
        }
    }

    // 共通の日付と場所を全行に反映
    applyCommonValues() {
        const commonDate = document.getElementById('bulkCommonDate').value;
        const commonPlace = document.getElementById('bulkCommonPlace').value;

        if (!commonDate && !commonPlace) {
            alert('共通日付または共通場所を入力してください');
            return;
        }

        const tbody = document.getElementById('bulkInputTableBody');
        const rows = tbody.querySelectorAll('tr');

        rows.forEach(row => {
            if (commonDate) {
                const dateInput = row.querySelector('.bulk-date');
                if (dateInput) {
                    dateInput.value = commonDate;
                }
            }

            if (commonPlace) {
                const placeInput = row.querySelector('.bulk-place');
                if (placeInput) {
                    placeInput.value = commonPlace;
                }
            }
        });

        this.showMessage('全行に日付と場所を反映しました');
    }

    // 一括入力データを保存
    saveBulkExpenses() {
        const tbody = document.getElementById('bulkInputTableBody');
        const rows = tbody.querySelectorAll('tr');
        let savedCount = 0;
        const errors = [];
        const validExpenses = [];

        // 毎月繰り返すチェックの確認
        const repeatCheckbox = document.getElementById('bulkRepeatMonthly');
        const repeatMonths = parseInt(document.getElementById('bulkRepeatMonths').value) || 1;

        rows.forEach((row, index) => {
            const date = row.querySelector('.bulk-date').value;
            const category = row.querySelector('.bulk-category').value;
            const subcategory = row.querySelector('.bulk-subcategory').value;
            const amount = row.querySelector('.bulk-amount').value;
            const place = row.querySelector('.bulk-place').value;
            const description = row.querySelector('.bulk-description').value;
            const notes = row.querySelector('.bulk-notes').value;

            // 必須項目チェック
            if (date && category && amount && parseFloat(amount) > 0) {
                validExpenses.push({
                    date, category, subcategory, amount, place, description, notes
                });
            } else if (date || category || amount || place || description || notes) {
                // 一部だけ入力されている場合はエラー
                errors.push(`行${index + 1}: 日付、カテゴリ、金額は必須です`);
            }
        });

        // 重複チェック
        if (validExpenses.length > 0) {
            const duplicateExpenses = validExpenses.filter(expense =>
                this.manager.isDuplicate(expense)
            );

            // 重複がある場合は警告を表示
            if (duplicateExpenses.length > 0) {
                let message = `${duplicateExpenses.length}件の重複データが検出されました。\n\n`;
                message += '以下のデータが既に登録されています：\n\n';
                duplicateExpenses.slice(0, 3).forEach((expense, index) => {
                    message += `${index + 1}. ${expense.date} - ${expense.category}`;
                    if (expense.subcategory) message += ` (${expense.subcategory})`;
                    message += ` - ${expense.amount}円`;
                    if (expense.place) message += ` - ${expense.place}`;
                    if (expense.description) message += ` - ${expense.description}`;
                    message += '\n';
                });
                if (duplicateExpenses.length > 3) {
                    message += `\n... 他 ${duplicateExpenses.length - 3}件\n`;
                }
                message += '\nそれでも保存しますか？';

                if (!confirm(message)) {
                    return; // 保存をキャンセル
                }
            }
        }

        // 繰り返し処理
        if (validExpenses.length > 0) {
            if (repeatCheckbox.checked && repeatMonths > 1) {
                // 複数月分の支出を作成
                validExpenses.forEach(expense => {
                    const baseDate = new Date(expense.date);
                    for (let i = 0; i < repeatMonths; i++) {
                        const newDate = new Date(baseDate);
                        newDate.setMonth(baseDate.getMonth() + i);

                        const repeatedExpense = {
                            ...expense,
                            date: newDate.toISOString().split('T')[0]
                        };
                        this.manager.addExpense(repeatedExpense);
                        savedCount++;
                    }
                });
            } else {
                // 通常の保存
                validExpenses.forEach(expense => {
                    this.manager.addExpense(expense);
                    savedCount++;
                });
            }

            this.renderExpenseList();
            this.showMessage(`${savedCount}件の支出を記録しました`);

            // テーブルをクリア
            tbody.innerHTML = '';
            this.bulkInputRows = [];
            // bulkRowIdCounterはリセットしない（同じrowIdの再利用を防ぐため）

            // 新しい行を3つ追加
            this.addBulkInputRow();
            this.addBulkInputRow();
            this.addBulkInputRow();
        } else {
            alert('記録するデータがありません。\n日付、カテゴリ、金額を入力してください。');
        }

        if (errors.length > 0) {
            alert('以下の行に入力エラーがあります:\n' + errors.join('\n'));
        }
    }

    // ========================================
    // Topix編集モーダル
    // ========================================

    // Topixモーダルを開く
    openTopixModal(yearMonth) {
        this.currentEditingYearMonth = yearMonth;
        const memo = this.manager.getMemo(yearMonth);
        const events = memo.events ? memo.events.split('\n').filter(e => e.trim()) : [];

        // モーダルタイトルと月情報を設定
        const [year, month] = yearMonth.split('-');
        document.querySelector('.topix-month-info').textContent = `${year}年${parseInt(month)}月のTopix`;

        // Topixリストを表示
        this.renderTopixList(events);

        // 入力欄をクリア
        document.getElementById('newTopixInput').value = '';

        // モーダルを表示
        document.getElementById('topixModal').style.display = 'block';
    }

    // Topixリストを描画
    renderTopixList(events) {
        const listContainer = document.getElementById('topixList');
        listContainer.innerHTML = '';

        if (events.length === 0) {
            listContainer.innerHTML = '<div class="topix-empty-message">Topixがまだ登録されていません</div>';
            return;
        }

        events.forEach((event, index) => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'topix-item';
            itemDiv.dataset.index = index;

            const textDiv = document.createElement('div');
            textDiv.className = 'topix-item-text';
            textDiv.textContent = event;
            textDiv.contentEditable = true;

            // 編集時の処理
            textDiv.addEventListener('focus', () => {
                textDiv.classList.add('editing');
            });

            textDiv.addEventListener('blur', () => {
                textDiv.classList.remove('editing');
            });

            // Enterキーで次の項目に移動
            textDiv.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    textDiv.blur();
                }
            });

            const actionsDiv = document.createElement('div');
            actionsDiv.className = 'topix-item-actions';

            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'topix-delete-btn';
            deleteBtn.textContent = '×';
            deleteBtn.title = '削除';
            deleteBtn.addEventListener('click', () => {
                if (confirm('このTopixを削除しますか？')) {
                    itemDiv.remove();
                    // 空になった場合はメッセージを表示
                    if (listContainer.children.length === 0) {
                        listContainer.innerHTML = '<div class="topix-empty-message">Topixがまだ登録されていません</div>';
                    }
                }
            });

            actionsDiv.appendChild(deleteBtn);
            itemDiv.appendChild(textDiv);
            itemDiv.appendChild(actionsDiv);
            listContainer.appendChild(itemDiv);
        });
    }

    // 新しいTopixを追加
    addNewTopix() {
        const input = document.getElementById('newTopixInput');
        const newTopix = input.value.trim();

        if (!newTopix) {
            return;
        }

        const listContainer = document.getElementById('topixList');

        // 空メッセージを削除
        const emptyMessage = listContainer.querySelector('.topix-empty-message');
        if (emptyMessage) {
            emptyMessage.remove();
        }

        // 新しいアイテムを作成
        const itemDiv = document.createElement('div');
        itemDiv.className = 'topix-item';

        const textDiv = document.createElement('div');
        textDiv.className = 'topix-item-text';
        textDiv.textContent = newTopix;
        textDiv.contentEditable = true;

        textDiv.addEventListener('focus', () => {
            textDiv.classList.add('editing');
        });

        textDiv.addEventListener('blur', () => {
            textDiv.classList.remove('editing');
        });

        textDiv.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                textDiv.blur();
            }
        });

        const actionsDiv = document.createElement('div');
        actionsDiv.className = 'topix-item-actions';

        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'topix-delete-btn';
        deleteBtn.textContent = '×';
        deleteBtn.title = '削除';
        deleteBtn.addEventListener('click', () => {
            if (confirm('このTopixを削除しますか？')) {
                itemDiv.remove();
                if (listContainer.children.length === 0) {
                    listContainer.innerHTML = '<div class="topix-empty-message">Topixがまだ登録されていません</div>';
                }
            }
        });

        actionsDiv.appendChild(deleteBtn);
        itemDiv.appendChild(textDiv);
        itemDiv.appendChild(actionsDiv);
        listContainer.appendChild(itemDiv);

        // 入力欄をクリア
        input.value = '';
        input.focus();
    }

    // Topix変更を保存
    saveTopixChanges() {
        const listContainer = document.getElementById('topixList');
        const items = listContainer.querySelectorAll('.topix-item');

        // 全てのTopixを配列に収集
        const events = [];
        items.forEach(item => {
            const text = item.querySelector('.topix-item-text').textContent.trim();
            if (text) {
                events.push(text);
            }
        });

        // メモデータを取得
        const memo = this.manager.getMemo(this.currentEditingYearMonth);

        // Topixを保存（既存のplansは保持）
        this.manager.saveMemo(this.currentEditingYearMonth, events.join('\n'), memo.plans || '');

        // モーダルを閉じる
        this.closeTopixModal();

        // 推移タブが表示されている場合は再描画
        const activeTab = document.querySelector('.tab-content.active');
        if (activeTab && activeTab.id === 'timeline-tab') {
            this.renderTimeline();
        }

        // 成功メッセージ
        const [year, month] = this.currentEditingYearMonth.split('-');
        alert(`${year}年${parseInt(month)}月のTopixを保存しました（${events.length}件）`);
    }

    // Topixモーダルを閉じる
    closeTopixModal() {
        document.getElementById('topixModal').style.display = 'none';
        this.currentEditingYearMonth = null;
    }
}

// ========================================
// ふるさと納税管理
// ========================================

class FurusatoManager {
    constructor() {
        this.storageKey = 'furusatoTaxData';
        this.loadData();
    }

    loadData() {
        const data = localStorage.getItem(this.storageKey);
        this.data = data ? JSON.parse(data) : [];

        let needsSave = false;

        // データ移行：古いデータでapplicantが空でmunicipalityに値がある場合、入れ替える
        this.data.forEach(entry => {
            if (!entry.applicant && entry.municipality) {
                entry.applicant = entry.municipality;
                entry.municipality = '';
                needsSave = true;
            }
        });

        // 重複IDを修正
        const idMap = new Map();
        this.data.forEach((entry, index) => {
            if (idMap.has(entry.id)) {
                // 重複しているIDを発見
                const newId = Date.now().toString() + '_' + index + '_' + Math.random().toString(36).substr(2, 9);
                console.log(`重複ID検出: ${entry.id} → ${newId} (品物: ${entry.item})`);
                entry.id = newId;
                needsSave = true;
            } else {
                idMap.set(entry.id, true);
            }
        });

        if (needsSave) {
            console.log('データを修正して保存します');
            this.saveData();
        }
    }

    saveData() {
        localStorage.setItem(this.storageKey, JSON.stringify(this.data));
    }

    add(entry) {
        // ユニークなIDを生成（既存のIDと重複しないようにする）
        let id = Date.now().toString();
        let counter = 0;
        while (this.data.some(e => e.id === id)) {
            counter++;
            id = (Date.now() + counter).toString();
        }

        const newEntry = {
            id,
            year: entry.year,
            amount: parseInt(entry.amount),
            item: entry.item,
            applicant: entry.applicant,
            municipality: entry.municipality || '',
            itemReceived: false,
            documentReceived: false
        };
        this.data.push(newEntry);
        this.saveData();
        return newEntry;
    }

    update(id, entry) {
        const index = this.data.findIndex(e => e.id === id);
        if (index !== -1) {
            this.data[index] = {
                ...this.data[index],
                year: entry.year,
                amount: parseInt(entry.amount),
                item: entry.item,
                applicant: entry.applicant,
                municipality: entry.municipality || ''
            };
            this.saveData();
            return this.data[index];
        }
        return null;
    }

    delete(id) {
        this.data = this.data.filter(e => e.id !== id);
        this.saveData();
    }

    toggleItemReceived(id) {
        const entry = this.data.find(e => e.id === id);
        if (entry) {
            entry.itemReceived = !entry.itemReceived;
            this.saveData();
        }
    }

    toggleDocumentReceived(id) {
        const entry = this.data.find(e => e.id === id);
        if (entry) {
            entry.documentReceived = !entry.documentReceived;
            this.saveData();
        }
    }

    getByYear(year) {
        return this.data.filter(e => e.year === year);
    }

    getAll() {
        return this.data;
    }

    importFromExpenses(expenses) {
        let importedCount = 0;
        let skippedCount = 0;

        expenses.forEach(expense => {
            // 重複チェック：同じ年、金額、品物、申請先のものがあればスキップ
            const isDuplicate = this.data.some(existing =>
                existing.year === expense.year &&
                parseInt(existing.amount) === parseInt(expense.amount) &&
                existing.item === expense.item &&
                existing.applicant === expense.applicant
            );

            if (!isDuplicate) {
                this.add(expense);
                importedCount++;
            } else {
                skippedCount++;
            }
        });

        return { importedCount, skippedCount };
    }
}

class FurusatoUI {
    constructor(manager, expenseManager) {
        this.manager = manager;
        this.expenseManager = expenseManager;
        this.stateManager = new StateManager();
        this.currentYear = new Date().getFullYear().toString();
        this.editingId = null;
        this.init();
    }

    init() {
        // 年度選択肢を動的に生成
        this.populateFurusatoYearSelects();

        // 年度選択
        document.getElementById('furusatoYear').addEventListener('change', (e) => {
            this.currentYear = e.target.value;
            this.stateManager.updateState('furusatoYear', e.target.value);
            this.render();
        });

        // 検索ボックス
        document.getElementById('furusatoSearch').addEventListener('input', (e) => {
            this.render();
        });

        // 新規追加ボタン
        document.getElementById('addFurusatoBtn').addEventListener('click', () => {
            this.openModal();
        });

        // インポートボタン
        document.getElementById('importFurusatoBtn').addEventListener('click', () => {
            this.importFromExpenses();
        });

        // モーダル関連
        document.getElementById('furusatoModalClose').addEventListener('click', () => {
            this.closeModal();
        });

        document.getElementById('cancelFurusatoBtn').addEventListener('click', () => {
            this.closeModal();
        });

        document.getElementById('furusatoForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveEntry();
        });

        // 保存された年度を復元
        const state = this.stateManager.loadState();
        if (state && state.furusatoYear) {
            this.currentYear = state.furusatoYear;
        }

        // 初期表示
        document.getElementById('furusatoYear').value = this.currentYear;
        this.render();
    }

    // ふるさと納税の年度選択肢を動的に生成
    populateFurusatoYearSelects() {
        const yearSelect = document.getElementById('furusatoYear');
        const yearInputSelect = document.getElementById('furusatoYearInput');
        const currentYear = new Date().getFullYear();

        // データから最小年と最大年を取得
        const allData = this.manager.getAll();
        const expenses = this.expenseManager.getAllExpenses();

        let minYear = 2020;
        let maxYear = currentYear;

        // ふるさと納税データから年を取得
        if (allData.length > 0) {
            const years = allData.map(e => parseInt(e.year));
            minYear = Math.min(...years, minYear);
            maxYear = Math.max(...years, maxYear);
        }

        // 支出データからも年を取得
        if (expenses.length > 0) {
            const expenseYears = expenses.map(e => new Date(e.date).getFullYear());
            minYear = Math.min(...expenseYears, minYear);
            maxYear = Math.max(...expenseYears, maxYear);
        }

        // 未来10年まで拡張
        const futureYear = currentYear + 10;
        maxYear = Math.max(maxYear, futureYear);

        // 両方のセレクトボックスに年を追加
        for (let year = minYear; year <= maxYear; year++) {
            const option1 = document.createElement('option');
            option1.value = year;
            option1.textContent = `${year}年`;
            yearSelect.appendChild(option1);

            const option2 = document.createElement('option');
            option2.value = year;
            option2.textContent = `${year}年`;
            yearInputSelect.appendChild(option2);
        }
    }

    render() {
        const container = document.getElementById('furusatoTableContainer');
        let entries = this.manager.getByYear(this.currentYear);

        // 検索フィルター
        const searchText = document.getElementById('furusatoSearch').value.toLowerCase().trim();
        if (searchText) {
            entries = entries.filter(entry => {
                const item = (entry.item || '').toLowerCase();
                const applicant = (entry.applicant || '').toLowerCase();
                const municipality = (entry.municipality || '').toLowerCase();
                return item.includes(searchText) ||
                       applicant.includes(searchText) ||
                       municipality.includes(searchText);
            });
        }

        console.log(`Rendering ${entries.length} entries for year ${this.currentYear}:`, entries);

        if (entries.length === 0) {
            const message = searchText
                ? `検索結果が見つかりません（検索: "${searchText}"）`
                : `${this.currentYear}年のふるさと納税データがありません`;
            container.innerHTML = `
                <div style="text-align: center; padding: 40px; color: #5f6368;">
                    <p>${message}</p>
                    ${!searchText ? '<p>「新規追加」ボタンから登録してください</p>' : ''}
                </div>
            `;
            return;
        }

        // 合計計算
        const total = entries.reduce((sum, e) => sum + e.amount, 0);

        const searchInfo = searchText ? ` （検索: "${searchText}"・${entries.length}件）` : '';
        let html = `
            <div class="furusato-summary">
                <strong>${this.currentYear}年 合計：¥${total.toLocaleString()}${searchInfo}</strong>
            </div>
            <table class="furusato-table">
                <thead>
                    <tr>
                        <th style="width: 35px;">No.</th>
                        <th style="width: 90px;">金額</th>
                        <th style="min-width: 200px;">品物</th>
                        <th style="min-width: 130px;">申請先</th>
                        <th style="min-width: 130px;">自治体</th>
                        <th style="width: 65px; text-align: center;">品物受取</th>
                        <th style="width: 65px; text-align: center;">納税書</th>
                        <th style="width: 70px; text-align: center;">操作</th>
                    </tr>
                </thead>
                <tbody>
        `;

        entries.forEach((entry, index) => {
            const escapedId = String(entry.id).replace(/'/g, "\\'");
            html += `
                <tr data-entry-id="${entry.id}">
                    <td style="text-align: center;">${index + 1}</td>
                    <td style="text-align: right; font-weight: 500;">¥${entry.amount.toLocaleString()}</td>
                    <td>${entry.item}</td>
                    <td>${entry.applicant || ''}</td>
                    <td>${entry.municipality || '-'}</td>
                    <td style="text-align: center;">
                        <input type="checkbox" ${entry.itemReceived ? 'checked' : ''}
                               data-action="toggleItem" data-id="${entry.id}"
                               style="cursor: pointer; width: 18px; height: 18px;">
                    </td>
                    <td style="text-align: center;">
                        <input type="checkbox" ${entry.documentReceived ? 'checked' : ''}
                               data-action="toggleDocument" data-id="${entry.id}"
                               style="cursor: pointer; width: 18px; height: 18px;">
                    </td>
                    <td style="text-align: center;">
                        <button class="btn-icon furusato-edit-btn" data-id="${entry.id}" title="編集">✏️</button>
                        <button class="btn-icon furusato-delete-btn" data-id="${entry.id}" title="削除">🗑️</button>
                    </td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
        `;

        container.innerHTML = html;

        // イベントリスナーを設定
        this.attachEventListeners();
    }

    attachEventListeners() {
        const container = document.getElementById('furusatoTableContainer');

        // 編集ボタン
        container.querySelectorAll('.furusato-edit-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-id');
                console.log('Edit button clicked, ID:', id);
                this.editEntry(id);
            });
        });

        // 削除ボタン
        container.querySelectorAll('.furusato-delete-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-id');
                this.deleteEntry(id);
            });
        });

        // チェックボックス
        container.querySelectorAll('input[type="checkbox"][data-action]').forEach(checkbox => {
            checkbox.addEventListener('change', (e) => {
                const id = e.currentTarget.getAttribute('data-id');
                const action = e.currentTarget.getAttribute('data-action');
                if (action === 'toggleItem') {
                    this.toggleItem(id);
                } else if (action === 'toggleDocument') {
                    this.toggleDocument(id);
                }
            });
        });
    }

    openModal(entry = null) {
        const modal = document.getElementById('furusatoModal');
        const title = document.getElementById('furusatoModalTitle');
        const form = document.getElementById('furusatoForm');

        // フォームを完全にリセット
        form.reset();

        if (entry) {
            // 編集モード
            title.textContent = 'ふるさと納税を編集';
            document.getElementById('furusatoId').value = entry.id;
            document.getElementById('furusatoYearInput').value = entry.year;
            document.getElementById('furusatoAmount').value = entry.amount;
            document.getElementById('furusatoItem').value = entry.item;
            document.getElementById('furusatoApplicant').value = entry.applicant || '';
            document.getElementById('furusatoMunicipality').value = entry.municipality || '';
            this.editingId = entry.id;
        } else {
            // 新規追加モード
            title.textContent = 'ふるさと納税を追加';
            document.getElementById('furusatoId').value = '';
            document.getElementById('furusatoYearInput').value = this.currentYear;
            document.getElementById('furusatoAmount').value = '';
            document.getElementById('furusatoItem').value = '';
            document.getElementById('furusatoApplicant').value = '';
            document.getElementById('furusatoMunicipality').value = '';
            this.editingId = null;
        }

        modal.style.display = 'block';
    }

    closeModal() {
        document.getElementById('furusatoModal').style.display = 'none';
        this.editingId = null;
    }

    saveEntry() {
        const entry = {
            year: document.getElementById('furusatoYearInput').value,
            amount: document.getElementById('furusatoAmount').value,
            item: document.getElementById('furusatoItem').value,
            applicant: document.getElementById('furusatoApplicant').value,
            municipality: document.getElementById('furusatoMunicipality').value
        };

        if (this.editingId) {
            this.manager.update(this.editingId, entry);
        } else {
            this.manager.add(entry);
        }

        this.closeModal();
        this.currentYear = entry.year;
        document.getElementById('furusatoYear').value = entry.year;
        this.render();
    }

    editEntry(id) {
        console.log('Editing entry with ID:', id);
        const entry = this.manager.getAll().find(e => e.id === id);
        console.log('Found entry:', entry);
        if (entry) {
            this.openModal(entry);
        } else {
            alert('エントリが見つかりませんでした。ID: ' + id);
        }
    }

    deleteEntry(id) {
        if (confirm('このふるさと納税データを削除しますか？')) {
            this.manager.delete(id);
            this.render();
        }
    }

    toggleItem(id) {
        this.manager.toggleItemReceived(id);
    }

    toggleDocument(id) {
        this.manager.toggleDocumentReceived(id);
    }

    importFromExpenses() {
        if (!confirm('支出データから小項目「ふるさと納税」のデータをインポートしますか？\n\n既存のデータと重複するものはスキップされます。')) {
            return;
        }

        // 全支出データから小項目「ふるさと納税」を抽出
        const allExpenses = this.expenseManager.getAllExpenses();
        const furusatoExpenses = allExpenses.filter(e =>
            e.subcategory && e.subcategory.includes('ふるさと納税')
        );

        if (furusatoExpenses.length === 0) {
            alert('小項目「ふるさと納税」のデータが見つかりませんでした。');
            return;
        }

        // データ変換
        const convertedData = furusatoExpenses.map(expense => {
            const date = new Date(expense.date);
            return {
                year: date.getFullYear().toString(),
                amount: parseInt(expense.amount),
                item: expense.description || '（商品名なし）',
                applicant: expense.place || '（申請先不明）',
                municipality: ''
            };
        });

        // インポート実行
        const result = this.manager.importFromExpenses(convertedData);

        alert(`インポート完了\n\n新規追加: ${result.importedCount}件\n重複でスキップ: ${result.skippedCount}件`);

        // 表示を更新
        this.render();
    }
}

// ========================================
// アプリ初期化
// ========================================

const manager = new ExpenseManager();
const ui = new UI(manager);
const furusatoManager = new FurusatoManager();
const furusatoUI = new FurusatoUI(furusatoManager, manager);

// アニメーション用CSS
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    @keyframes slideOut {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
    }
    .badge {
        background: #667eea;
        color: white;
        padding: 4px 12px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 500;
    }
`;
document.head.appendChild(style);
