// ========================================
// データ管理
// ========================================

class ExpenseManager {
    constructor() {
        this.expenses = this.loadExpenses();
        this.reports = this.loadReports();
        this.memos = this.loadMemos();
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
        return this.memos[yearMonth] || { events: '', plans: '' };
    }

    // 月のメモを保存
    saveMemo(yearMonth, events, plans) {
        this.memos[yearMonth] = { events, plans };
        this.saveMemos();
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
        return expense;
    }

    // 複数の支出を一括追加（インポート用・保存は最後に1回だけ）
    addExpensesBatch(expenses) {
        expenses.forEach((expense, index) => {
            expense.id = Date.now().toString() + '_' + index + '_' + Math.random().toString(36).substr(2, 9);
            this.expenses.push(expense);
        });
        this.saveExpenses(); // 最後に1回だけ保存
    }

    // 支出を更新
    updateExpense(id, updatedData) {
        const index = this.expenses.findIndex(e => e.id === id);
        if (index !== -1) {
            this.expenses[index] = { ...this.expenses[index], ...updatedData };
            this.saveExpenses();
            return true;
        }
        return false;
    }

    // 支出を削除
    deleteExpense(id) {
        this.expenses = this.expenses.filter(e => e.id !== id);
        this.saveExpenses();
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
}

// ========================================
// UI管理
// ========================================

class UI {
    constructor(manager) {
        this.manager = manager;
        this.monthlyChart = null;
        this.categoryChart = null;

        // カテゴリ別の小項目マスター（実際のExcelデータから抽出）
        this.subcategoryMaster = {
            '食品': ['葉物・生鮮野菜', '根菜', '野菜その他', '肉類', '魚類・貝類', '肉類加工品', '魚類加工品', 'レトルト類・冷凍食品', 'お米', 'パン', '麺類', 'たまご', '牛乳・乳製品', '豆腐・納豆類', 'ドリンク', 'お菓子類', '調味料', '粉類', '食品その他'],
            '日用品': ['キッチン用品', 'バス用品', 'ペーパー類', '掃除用品', '文具', '日用品その他', '洗剤類', '美容関連用品', '薬類', '虫対策類', '衛生用品'],
            '外食費': ['To Go', 'お惣菜', 'お茶', 'からやま', 'とんとん亭', 'はな膳', 'びっくりドンキー', 'まるみや', 'カウベル', 'ケンタッキー', 'ココス', 'サイゼリヤ', 'スシロー', 'ドトール', '外食その他'],
            '衣類': ['natsukiインナー', 'natsuki衣類', 'tettaインナー', 'tetta衣類', '衣類その他'],
            '家具・家電': ['キッチン家電', '家具', '家具家電その他', '家電', '照明関連', '空調家電', '寝具関連'],
            '美容': ['スキンケア', 'コスメ', 'ヘアケア', 'エステ', 'ネイルサロン', 'ハイフ', '美容その他'],
            '医療費': ['医療費その他', '哲太　整骨院', '哲太マッサージ', '奈月　整骨院', '整骨院', '薬代', '診察料'],
            '交際費': ['交際費その他', '友人・知人', '宇野家関連', '宮崎家関連', '敬老の日', '父の日', '母の日', '自治会費', '記念日'],
            'レジャー': ['BBQ', 'おでかけ', 'お家BBQ', 'キャンプ', 'キャンプ場', 'キャンプ用品', 'キャンプ食材', 'レジャーその他'],
            'ガソリン・ETC': ['Costcoガソリン', 'ETC', 'ガソリン', 'ガソリンその他'],
            '光熱費': ['ガス', '下水道', '水道', '電気'],
            '通信費': ['ケータイ', '通信'],
            '保険': ['AIG損保', 'アクサダイレクト', '保険その他', '哲太メットライフ', '奈月メットライフ'],
            '車関連【車検・税金・積立】': ['ガソリン', '税金', '積立', '車検', '車検・税金積立', '車関連その他'],
            '税金': ['ふるさと納税', '固定資産', '固定資産税', '奈月市民税', '年金　奈月', '税金その他'],
            '経費': ['natsuki', 'tetta', '奈月携帯', '経費その他'],
            'ローン': ['PC', 'キッチン家電', 'ローン', 'ローン住居', 'ローン家電'],
            'その他': ['その他']
        };

        this.init();
    }

    init() {
        // 変数を先に初期化
        this.bulkInputRows = [];
        this.currentSort = { column: 'date', direction: 'desc' }; // デフォルトは日付降順
        this.displayLimit = 100; // 初期表示件数
        this.displayOffset = 0; // 表示開始位置

        // UIセットアップ
        this.setupEventListeners();
        this.setupTabs();
        this.setDefaultDate();
        this.populateYearSelect();
        this.populateTimelineYearSelects();
        this.loadSavedReportsList();
        this.setDefaultMemoMonth();

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
        });

        // 月フィルター
        document.getElementById('monthFilter').addEventListener('change', () => {
            this.renderExpenseList(true); // フィルター変更時は表示件数リセット
        });

        // カスタム期間の適用ボタン
        document.getElementById('applyFilter').addEventListener('click', () => {
            this.renderExpenseList(true);
        });

        // 開始日・終了日の変更（Enterキーで適用）
        document.getElementById('startDate').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.renderExpenseList(true);
        });
        document.getElementById('endDate').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.renderExpenseList(true);
        });

        // フィルタークリア
        document.getElementById('clearFilter').addEventListener('click', () => {
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

        // 分析年月変更
        document.getElementById('analysisYear').addEventListener('change', () => {
            this.updateAnalysis();
        });

        document.getElementById('analysisMonth').addEventListener('change', () => {
            this.updateAnalysis();
        });

        document.getElementById('analysisCategory').addEventListener('change', () => {
            this.updateAnalysis();
        });

        // エクスポート・インポート
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

                // タブに応じて更新
                if (targetTab === 'analysis') {
                    this.updateAnalysis();
                } else if (targetTab === 'timeline') {
                    this.renderTimeline();
                }
            });
        });
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

        // 2020年から現在年まで
        for (let year = 2020; year <= currentYear; year++) {
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

        // 2020年から現在年まで
        for (let year = 2020; year <= currentYear; year++) {
            const option1 = document.createElement('option');
            option1.value = year;
            option1.textContent = `${year}年`;
            startSelect.appendChild(option1);

            const option2 = document.createElement('option');
            option2.value = year;
            option2.textContent = `${year}年`;
            endSelect.appendChild(option2);
        }

        // デフォルトは現在年
        startSelect.value = currentYear;
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
            description: document.getElementById('description').value
        };

        this.manager.addExpense(expense);
        this.renderExpenseList();

        // フォームをリセット
        document.getElementById('expenseForm').reset();
        this.setDefaultDate();

        // 成功メッセージ
        this.showMessage('支出を記録しました');
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

            // ソート適用
            expenses = this.applySortToExpenses(expenses);

            // ヘッダーのソート表示を更新
            this.updateSortIndicators();

            // 合計計算（全体）
            const total = expenses.reduce((sum, e) => sum + parseInt(e.amount || 0), 0);
            document.getElementById('displayTotal').textContent = total.toLocaleString();

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
                        <td><span class="badge" data-category="${expense.category || ''}">${expense.category || '-'}</span></td>
                        <td>${expense.subcategory || '-'}</td>
                        <td>${parseInt(expense.amount || 0).toLocaleString()}円</td>
                        <td>${expense.place || '-'}</td>
                        <td>${expense.description || '-'}</td>
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
            description: document.getElementById('editDescription').value
        };

        this.manager.updateExpense(id, updatedData);
        this.renderExpenseList();
        this.closeModal();
        this.showMessage('支出を更新しました');
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

        // テーブルヘッダーを生成
        const thead = document.getElementById('timelineTableHead');
        thead.innerHTML = `
            <tr>
                <th class="sticky-col">カテゴリ</th>
                ${months.map(m => `<th>${m.year.toString().substring(2)}.${String(m.month).padStart(2, '0')}</th>`).join('')}
            </tr>
        `;

        // テーブルボディを生成
        const tbody = document.getElementById('timelineTableBody');
        tbody.innerHTML = '';

        categories.forEach(category => {
            const row = document.createElement('tr');
            const cells = months.map(m => {
                const amount = data[category][m.key];
                const displayAmount = amount > 0 ? `¥${amount.toLocaleString()}` : '';
                const bgColor = amount > 0 ? this.getCategoryColor(category) : '';
                return `<td style="background:${bgColor}">${displayAmount}</td>`;
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

        // 出来事行
        const memoRow = document.createElement('tr');
        memoRow.className = 'timeline-memo-row';
        const memoCells = months.map(m => {
            const memo = this.manager.getMemo(m.key);
            const events = memo.events ? memo.events.trim() : '';

            if (events) {
                // 改行を削除して1行にする
                const oneLine = events.replace(/\n/g, ' ');
                return `<td class="memo-cell" title="${this.escapeHtml(events)}">${this.escapeHtml(oneLine)}</td>`;
            } else {
                return `<td class="memo-cell"></td>`;
            }
        });
        memoRow.innerHTML = `
            <td class="sticky-col memo-label">出来事</td>
            ${memoCells.join('')}
        `;
        tbody.appendChild(memoRow);
    }

    // HTMLエスケープ
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
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

    // 分析データ更新
    updateAnalysis() {
        const year = parseInt(document.getElementById('analysisYear').value);
        const month = document.getElementById('analysisMonth').value;
        const selectedCategory = document.getElementById('analysisCategory').value;

        let expenses;
        if (month) {
            expenses = this.manager.getExpensesByMonth(year, parseInt(month));
        } else {
            expenses = this.manager.getExpensesByYear(year);
        }

        // カテゴリフィルター適用
        if (selectedCategory) {
            expenses = expenses.filter(e => e.category === selectedCategory);
        }

        // タイトル更新
        if (selectedCategory) {
            document.getElementById('categoryStatsTitle').textContent =
                `${selectedCategory} - 小項目別集計`;
            document.getElementById('monthlyChartTitle').textContent =
                `${selectedCategory} - 月別推移`;
            document.getElementById('categoryChartTitle').textContent =
                `${selectedCategory} - 小項目別内訳`;
        } else {
            document.getElementById('categoryStatsTitle').textContent = 'カテゴリ別集計';
            document.getElementById('monthlyChartTitle').textContent = '月別推移';
            document.getElementById('categoryChartTitle').textContent = 'カテゴリ別内訳';
        }

        // カテゴリ別集計または小項目別集計
        if (selectedCategory) {
            this.renderSubcategoryStats(expenses, selectedCategory);
        } else {
            this.renderCategoryStats(expenses);
        }

        // グラフ更新
        this.updateCharts(year, selectedCategory);
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

    // グラフ更新
    updateCharts(year, selectedCategory = null) {
        if (selectedCategory) {
            // 特定カテゴリの月別推移
            const yearExpenses = this.manager.getExpensesByYear(year);
            const categoryExpenses = yearExpenses.filter(e => e.category === selectedCategory);

            // 月別集計
            const monthlyData = Array(12).fill(0);
            categoryExpenses.forEach(e => {
                const date = new Date(e.date);
                monthlyData[date.getMonth()] += parseInt(e.amount);
            });
            this.renderMonthlyChart(monthlyData);

            // 小項目別集計
            const subcategoryTotals = {};
            categoryExpenses.forEach(e => {
                const subcat = e.subcategory || '(未分類)';
                subcategoryTotals[subcat] = (subcategoryTotals[subcat] || 0) + parseInt(e.amount);
            });
            this.renderCategoryChart(subcategoryTotals);
        } else {
            // 全カテゴリの月別推移
            const monthlyData = this.manager.getMonthlyTotals(year);
            this.renderMonthlyChart(monthlyData);

            // カテゴリ別円グラフ
            const yearExpenses = this.manager.getExpensesByYear(year);
            const categoryTotals = this.manager.getCategoryTotals(yearExpenses);
            this.renderCategoryChart(categoryTotals);
        }
    }

    // 月別推移グラフ
    renderMonthlyChart(data) {
        const ctx = document.getElementById('monthlyChart').getContext('2d');

        if (this.monthlyChart) {
            this.monthlyChart.destroy();
        }

        this.monthlyChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'],
                datasets: [{
                    label: '支出額',
                    data: data,
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    tension: 0.4,
                    fill: true
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return value.toLocaleString() + '円';
                            }
                        }
                    }
                }
            }
        });
    }

    // カテゴリ別円グラフ
    renderCategoryChart(data) {
        const ctx = document.getElementById('categoryChart').getContext('2d');

        if (this.categoryChart) {
            this.categoryChart.destroy();
        }

        const labels = Object.keys(data);
        const values = Object.values(data);
        const colors = [
            '#667eea', '#764ba2', '#f093fb', '#4facfe',
            '#43e97b', '#fa709a', '#fee140', '#30cfd0'
        ];

        this.categoryChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [{
                    data: values,
                    backgroundColor: colors.slice(0, labels.length)
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        position: 'bottom'
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = context.parsed || 0;
                                return label + ': ' + value.toLocaleString() + '円';
                            }
                        }
                    }
                }
            }
        });
    }

    // データエクスポート
    exportData() {
        const data = this.manager.getAllExpenses();
        const csv = this.convertToCSV(data);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);

        link.setAttribute('href', url);
        link.setAttribute('download', `家計簿_${new Date().toISOString().split('T')[0]}.csv`);
        link.click();

        this.showMessage('データをエクスポートしました');
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
                this.addBulkInputRow();
                this.addBulkInputRow();
                this.addBulkInputRow();
            }
        }
    }

    // 一括入力行を追加
    addBulkInputRow() {
        const tbody = document.getElementById('bulkInputTableBody');
        const rowId = Date.now();
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
            <td><input type="number" class="bulk-amount" min="0" placeholder="金額"></td>
            <td><button class="tax-btn" onclick="ui.applyTax(${rowId})">税込</button></td>
            <td><input type="text" class="bulk-place" placeholder="場所"></td>
            <td><input type="text" class="bulk-description" placeholder="商品名・メモ"></td>
            <td><button class="delete-row-btn" onclick="ui.deleteBulkInputRow(${rowId})">削除</button></td>
        `;

        tbody.appendChild(row);
        this.bulkInputRows.push(rowId);

        // 小項目の選択肢を更新
        if (defaultCategory) {
            this.updateSubcategoryOptions(rowId);
        }

        // 1行目のカテゴリ変更時に、2行目以降も更新
        if (this.bulkInputRows.length === 1) {
            const categorySelect = row.querySelector('.bulk-category');
            categorySelect.addEventListener('change', () => {
                this.propagateCategoryToOtherRows(rowId);
            });
        }
    }

    // 複数行を一度に追加
    addBulkInputRows(count) {
        for (let i = 0; i < count; i++) {
            // 少しずつIDをずらすため、小さい遅延を入れる
            setTimeout(() => {
                this.addBulkInputRow();
            }, i * 10);
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
    propagateCategoryToOtherRows(firstRowId) {
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

                // 小項目の選択肢を更新
                this.updateSubcategoryOptions(parseInt(rowId));
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

        rows.forEach((row, index) => {
            const date = row.querySelector('.bulk-date').value;
            const category = row.querySelector('.bulk-category').value;
            const subcategory = row.querySelector('.bulk-subcategory').value;
            const amount = row.querySelector('.bulk-amount').value;
            const place = row.querySelector('.bulk-place').value;
            const description = row.querySelector('.bulk-description').value;

            // 必須項目チェック
            if (date && category && amount && parseFloat(amount) > 0) {
                const expense = {
                    date, category, subcategory, amount, place, description
                };
                this.manager.addExpense(expense);
                savedCount++;
            } else if (date || category || amount || place || description) {
                // 一部だけ入力されている場合はエラー
                errors.push(`行${index + 1}: 日付、カテゴリ、金額は必須です`);
            }
        });

        if (savedCount > 0) {
            this.renderExpenseList();
            this.showMessage(`${savedCount}件の支出を記録しました`);

            // テーブルをクリア
            tbody.innerHTML = '';
            this.bulkInputRows = [];

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
}

// ========================================
// アプリ初期化
// ========================================

const manager = new ExpenseManager();
const ui = new UI(manager);

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
