<script>
/**
 * StudyQuest Optimized Application - GAS Integration Version
 * 最適化されたアーキテクチャを統合したメインアプリケーション
 */

// Core State Management
<?!= include('core-state.js'); ?>

// Data Layer
<?!= include('data-layer.js'); ?>

// UI Renderer  
<?!= include('ui-renderer.js'); ?>

// API Optimizer
<?!= include('api-optimizer.js'); ?>

// Memory Manager
<?!= include('memory-manager.js'); ?>

// GAS Error Handler
<?!= include('gas-error-handler.js'); ?>

// Main Application Integration
class StudyQuestApp {
  constructor() {
    console.log('🚀 Initializing StudyQuest Optimized Application...');
    
    // 初期化開始時刻
    this.initStartTime = Date.now();
    
    // コアコンポーネントの初期化
    this.initializeCoreComponents();
    
    // アプリケーション状態
    this.reactionTypes = [
      { key: 'LIKE', icon: 'hand-thumb-up-outline', solidIcon: 'hand-thumb-up-solid', label: 'いいね' },
      { key: 'UNDERSTAND', icon: 'lightbulb-outline', solidIcon: 'lightbulb-solid', label: '理解' },
      { key: 'CURIOUS', icon: 'magnifying-glass-plus-outline', solidIcon: 'magnifying-glass-plus-solid', label: '興味' }
    ];
    
    this.initializeApplication();
  }

  /**
   * コアコンポーネントの初期化
   */
  initializeCoreComponents() {
    try {
      // 1. Core State Manager
      this.coreState = new StudyQuestCoreState();
      console.log('✅ Core State initialized');
      
      // 2. GAS Error Handler
      this.errorHandler = new GASErrorHandler(this.coreState);
      console.log('✅ Error Handler initialized');
      
      // 3. Data Manager
      this.dataManager = new GASApiManager(this.coreState);
      console.log('✅ Data Manager initialized');
      
      // 4. Memory Manager
      this.memoryManager = new AutoMemoryManager(this.coreState);
      console.log('✅ Memory Manager initialized');
      
      // 5. UI Renderer
      this.uiRenderer = new UIRenderer(this.coreState);
      console.log('✅ UI Renderer initialized');
      
      // 6. Polling Manager
      this.pollingManager = new OptimizedPollingManager(this.coreState, this.dataManager);
      console.log('✅ Polling Manager initialized');
      
      // 7. Batch Processing Manager
      this.batchManager = new BatchProcessingManager(this.coreState, this.dataManager);
      console.log('✅ Batch Manager initialized');
      
    } catch (error) {
      console.error('❌ Core component initialization failed:', error);
      this.handleInitializationError(error);
    }
  }

  /**
   * アプリケーションの初期化
   */
  async initializeApplication() {
    try {
      // GAS実行時間チェック
      this.coreState.checkGASExecutionTime();
      
      console.log('🔧 Starting application initialization...');
      
      // 1. ユーザーIDの設定と検証
      await this.initializeUser();
      
      // 2. DOM要素の初期化
      this.initializeDOMElements();
      
      // 3. イベントリスナーの設定
      this.setupEventListeners();
      
      // 4. 初期データの読み込み
      await this.loadInitialData();
      
      // 5. ポーリングの開始
      this.startServices();
      
      // 6. 初期化完了
      this.finalizeInitialization();
      
    } catch (error) {
      console.error('❌ Application initialization failed:', error);
      await this.errorHandler.handleError(error, 'applicationInitialization');
    }
  }

  /**
   * ユーザー初期化
   */
  async initializeUser() {
    // URLパラメータまたはグローバル変数からユーザーIDを取得
    let userId = '';
    
    if (typeof USER_ID !== 'undefined') {
      userId = USER_ID;
    } else {
      const urlParams = new URLSearchParams(window.location.search);
      userId = urlParams.get('userId') || '';
    }
    
    // ユーザーIDの検証
    const validatedUserId = this.coreState.validateUserId(userId);
    
    this.coreState.updateState({ userId: validatedUserId });
    
    // 管理者権限のチェック
    try {
      const isAdmin = await this.dataManager.checkAdmin();
      this.coreState.updateState({ showAdminFeatures: isAdmin });
      
      if (isAdmin) {
        console.log('👑 Admin privileges detected');
        this.updateAdminUI();
      }
    } catch (error) {
      console.warn('⚠️ Admin check failed:', error);
      // 管理者チェック失敗は致命的でない
    }
  }

  /**
   * DOM要素の初期化
   */
  initializeDOMElements() {
    // UIRendererから要素取得
    this.elements = this.uiRenderer.elements;
    
    // 必須要素の存在確認
    const requiredElements = ['answersContainer'];
    const missing = requiredElements.filter(key => !this.elements[key]);
    
    if (missing.length > 0) {
      throw new Error(`Required DOM elements missing: ${missing.join(', ')}`);
    }
    
    console.log('✅ DOM elements initialized');
  }

  /**
   * イベントリスナーの設定
   */
  setupEventListeners() {
    // メモリ管理付きでイベントリスナーを設定
    
    // 回答カードのクリックイベント（イベント委譲）
    if (this.elements.answersContainer) {
      const handleAnswerClick = (e) => {
        this.handleAnswerCardClick(e);
      };
      
      this.memoryManager.trackEventListener(
        this.elements.answersContainer,
        'click',
        handleAnswerClick
      );
    }
    
    // フィルター変更
    if (this.elements.classFilter) {
      const handleFilterChange = async () => {
        await this.handleFilterChange();
      };
      
      this.memoryManager.trackEventListener(
        this.elements.classFilter,
        'change',
        handleFilterChange
      );
    }
    
    // ソート順変更
    if (this.elements.sortOrder) {
      const handleSortChange = async () => {
        await this.handleSortChange();
      };
      
      this.memoryManager.trackEventListener(
        this.elements.sortOrder,
        'change',
        handleSortChange
      );
    }
    
    // 管理者トグル
    if (this.elements.adminToggleBtn) {
      const handleAdminToggle = () => {
        this.toggleAdminMode();
      };
      
      this.memoryManager.trackEventListener(
        this.elements.adminToggleBtn,
        'click',
        handleAdminToggle
      );
    }
    
    // 公開終了ボタン
    if (this.elements.endPublicationBtn) {
      const handleEndPublication = () => {
        this.endPublication();
      };
      
      this.memoryManager.trackEventListener(
        this.elements.endPublicationBtn,
        'click',
        handleEndPublication
      );
    }
    
    // バッチ処理完了イベント
    const handleBatchComplete = (event) => {
      this.handleBatchProcessingComplete(event);
    };
    
    this.memoryManager.trackEventListener(
      document,
      'batchProcessingComplete',
      handleBatchComplete
    );
    
    // ウィンドウリサイズ
    const handleResize = this.debounce(() => {
      this.uiRenderer.adjustLayout();
    }, 200);
    
    this.memoryManager.trackEventListener(window, 'resize', handleResize, { passive: true });
    
    // ページ可視性変更
    const handleVisibilityChange = () => {
      if (document.hidden) {
        this.pollingManager.stopPolling();
        this.memoryManager.performOptimization();
      } else {
        this.pollingManager.startPolling();
      }
    };
    
    this.memoryManager.trackEventListener(
      document,
      'visibilitychange',
      handleVisibilityChange,
      { passive: true }
    );
    
    console.log('✅ Event listeners setup completed');
  }

  /**
   * 初期データの読み込み
   */
  async loadInitialData() {
    this.showLoadingOverlay('データを読み込んでいます...');
    
    try {
      // 現在のフィルター設定を取得
      const classFilter = this.elements.classFilter?.value || 'すべて';
      const sortOrder = this.elements.sortOrder?.value || 'newest';
      const adminMode = this.coreState.state.showAdminFeatures;
      
      console.log('📥 Loading initial data...', { classFilter, sortOrder, adminMode });
      
      // データ取得
      const data = await this.dataManager.getPublishedSheetData(
        classFilter,
        sortOrder,
        adminMode,
        true // キャッシュをバイパス
      );
      
      if (!data || !Array.isArray(data.data)) {
        throw new Error('Invalid data format received');
      }
      
      // 状態を更新
      this.coreState.updateState({
        currentAnswers: data.data,
        lastSeenCount: data.data.length,
        initialized: true
      });
      
      // UIに表示
      await this.uiRenderer.renderBoard(data.data, { isInitialLoad: true });
      
      // クラスフィルターを設定
      this.populateClassFilter(data.data);
      
      console.log(`✅ Initial data loaded: ${data.data.length} items`);
      
    } catch (error) {
      console.error('❌ Initial data loading failed:', error);
      await this.errorHandler.handleError(error, 'initialDataLoad');
      
      // エラー時は空状態を表示
      this.displayEmptyState();
    } finally {
      this.hideLoadingOverlay();
    }
  }

  /**
   * サービスの開始
   */
  startServices() {
    // ポーリング開始
    if (this.coreState.state.initialized) {
      this.pollingManager.startPolling();
      console.log('✅ Polling service started');
    }
    
    console.log('✅ Background services started');
  }

  /**
   * 初期化の最終化
   */
  finalizeInitialization() {
    const initDuration = Date.now() - this.initStartTime;
    
    this.coreState.updatePerformanceMetrics({
      initTime: initDuration
    });
    
    console.log(`🎉 StudyQuest Optimized Application initialized successfully in ${initDuration}ms`);
    
    // グローバル参照を設定（デバッグ用）
    window.studyQuestApp = this;
    
    // 初期化完了イベントを発火
    const event = new CustomEvent('studyQuestInitialized', {
      detail: { initTime: initDuration }
    });
    document.dispatchEvent(event);
  }

  /**
   * イベントハンドラー
   */
  handleAnswerCardClick(e) {
    const answerCard = e.target.closest('.answer-card');
    if (!answerCard || answerCard.classList.contains('hidden-card')) {
      return;
    }
    
    const rowIndex = answerCard.dataset.rowIndex;
    if (!rowIndex) return;
    
    // リアクションボタンのクリック
    const reactionBtn = e.target.closest('.reaction-btn');
    if (reactionBtn) {
      e.stopPropagation();
      if (!reactionBtn.disabled) {
        this.handleReaction(rowIndex, reactionBtn.dataset.reaction);
      }
      return;
    }
    
    // ハイライトボタンのクリック
    const highlightBtn = e.target.closest('.highlight-btn');
    if (highlightBtn) {
      e.stopPropagation();
      if (!highlightBtn.disabled) {
        this.handleHighlight(rowIndex);
      }
      return;
    }
    
    // カード全体のクリック（詳細表示など）
    this.showAnswerModal(rowIndex);
  }

  async handleReaction(rowIndex, reaction) {
    try {
      // バッチ処理に追加
      this.batchManager.addReactionToBatch(rowIndex, reaction);
      
      // 楽観的UI更新
      this.updateReactionUI(rowIndex, reaction, true);
      
    } catch (error) {
      console.error('❌ Reaction handling failed:', error);
      await this.errorHandler.handleError(error, 'reactionHandling');
    }
  }

  async handleHighlight(rowIndex) {
    try {
      // バッチ処理に追加
      this.batchManager.addHighlightToBatch(rowIndex);
      
      // 楽観的UI更新
      this.updateHighlightUI(rowIndex, true);
      
    } catch (error) {
      console.error('❌ Highlight handling failed:', error);
      await this.errorHandler.handleError(error, 'highlightHandling');
    }
  }

  async handleFilterChange() {
    try {
      await this.refreshData({ reason: 'filterChange' });
    } catch (error) {
      await this.errorHandler.handleError(error, 'filterChange');
    }
  }

  async handleSortChange() {
    try {
      await this.refreshData({ reason: 'sortChange' });
    } catch (error) {
      await this.errorHandler.handleError(error, 'sortChange');
    }
  }

  handleBatchProcessingComplete(event) {
    const { type, operations, result } = event.detail;
    
    console.log(`✅ Batch processing completed: ${type}`, { operations: operations.length });
    
    // UI の同期更新
    this.syncUIWithServerState(result);
  }

  /**
   * ユーティリティメソッド
   */
  async refreshData(options = {}) {
    this.showLoadingOverlay('データを更新中...');
    
    try {
      const classFilter = this.elements.classFilter?.value || 'すべて';
      const sortOrder = this.elements.sortOrder?.value || 'newest';
      const adminMode = this.coreState.state.showAdminFeatures;
      
      const data = await this.dataManager.getPublishedSheetData(
        classFilter,
        sortOrder,
        adminMode,
        true // キャッシュをバイパス
      );
      
      if (data && Array.isArray(data.data)) {
        this.coreState.updateState({
          currentAnswers: data.data,
          lastSeenCount: data.data.length
        });
        
        await this.uiRenderer.renderBoard(data.data, { forceRefresh: true });
        
        console.log(`✅ Data refreshed: ${data.data.length} items`);
      }
      
    } catch (error) {
      throw error;
    } finally {
      this.hideLoadingOverlay();
    }
  }

  updateReactionUI(rowIndex, reaction, optimistic = false) {
    const card = document.querySelector(`[data-row-index="${rowIndex}"]`);
    if (!card) return;
    
    const reactionBtn = card.querySelector(`[data-reaction="${reaction}"]`);
    if (!reactionBtn) return;
    
    // 楽観的更新の場合は一時的な状態を設定
    if (optimistic) {
      reactionBtn.classList.add('processing');
      reactionBtn.disabled = true;
    }
  }

  updateHighlightUI(rowIndex, optimistic = false) {
    const card = document.querySelector(`[data-row-index="${rowIndex}"]`);
    if (!card) return;
    
    const highlightBtn = card.querySelector('.highlight-btn');
    if (!highlightBtn) return;
    
    if (optimistic) {
      highlightBtn.classList.add('processing');
      highlightBtn.disabled = true;
    }
  }

  syncUIWithServerState(serverState) {
    // サーバー状態とUIを同期
    if (serverState && Array.isArray(serverState)) {
      serverState.forEach(item => {
        const card = document.querySelector(`[data-row-index="${item.rowIndex}"]`);
        if (card) {
          this.uiRenderer.applyReactionStyles(card, item);
        }
      });
    }
  }

  populateClassFilter(data) {
    if (!this.elements.classFilter || !Array.isArray(data)) return;
    
    const uniqueClasses = ['すべて', ...new Set(data.map(r => r.class).filter(Boolean))];
    this.elements.classFilter.innerHTML = uniqueClasses
      .map(c => `<option value="${this.escapeHtml(c)}">${this.escapeHtml(c)}</option>`)
      .join('');
  }

  showAnswerModal(rowIndex) {
    // モーダル表示の実装
    console.log('📋 Showing answer modal for row:', rowIndex);
  }

  toggleAdminMode() {
    const currentMode = this.coreState.state.showAdminFeatures;
    this.coreState.updateState({ showAdminFeatures: !currentMode });
    
    console.log(`🔧 Admin mode ${!currentMode ? 'enabled' : 'disabled'}`);
  }

  async endPublication() {
    try {
      await this.dataManager.runGas('clearActiveSheet');
      console.log('📋 Publication ended');
    } catch (error) {
      await this.errorHandler.handleError(error, 'endPublication');
    }
  }

  updateAdminUI() {
    if (this.elements.adminToggleBtn) {
      this.elements.adminToggleBtn.classList.remove('hidden');
    }
  }

  showLoadingOverlay(message = '') {
    if (this.elements.loadingOverlay) {
      this.elements.loadingOverlay.style.display = 'flex';
      const messageEl = this.elements.loadingOverlay.querySelector('.loading-message');
      if (messageEl) {
        messageEl.textContent = message;
      }
    }
  }

  hideLoadingOverlay() {
    if (this.elements.loadingOverlay) {
      this.elements.loadingOverlay.style.display = 'none';
    }
  }

  displayEmptyState() {
    if (this.elements.answersContainer) {
      this.elements.answersContainer.innerHTML = `
        <div class="empty-state text-center py-12">
          <div class="text-gray-400 text-lg mb-4">
            📝 まだ回答がありません
          </div>
          <div class="text-gray-500 text-sm">
            回答が投稿されると、ここに表示されます
          </div>
        </div>
      `;
    }
  }

  escapeHtml(text) {
    if (typeof text !== 'string') return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  handleInitializationError(error) {
    console.error('🚨 Critical initialization error:', error);
    
    // 最小限のエラー表示
    const container = document.getElementById('answers') || document.body;
    container.innerHTML = `
      <div style="padding: 20px; text-align: center; color: #dc2626;">
        <h2>アプリケーションの初期化に失敗しました</h2>
        <p>${this.escapeHtml(error.message)}</p>
        <button onclick="location.reload()" style="padding: 10px 20px; margin-top: 10px;">
          再読み込み
        </button>
      </div>
    `;
  }

  /**
   * 診断情報の取得
   */
  getDiagnosticInfo() {
    return {
      initTime: Date.now() - this.initStartTime,
      coreState: this.coreState.serializeState(),
      memoryStats: this.memoryManager.getResourceStats(),
      errorStats: this.errorHandler.getErrorStats(),
      pollingStats: this.pollingManager.getStats(),
      batchStats: this.batchManager.getStats(),
      dataStats: this.dataManager.getStats()
    };
  }

  /**
   * アプリケーションの破棄
   */
  destroy() {
    console.log('🧹 Destroying StudyQuest Application...');
    
    // 各コンポーネントの破棄
    this.pollingManager?.destroy();
    this.batchManager?.destroy();
    this.memoryManager?.destroy();
    this.errorHandler?.destroy();
    this.dataManager?.destroy();
    this.uiRenderer?.destroy();
    this.coreState?.destroy();
    
    // グローバル参照をクリア
    window.studyQuestApp = null;
    
    console.log('✅ StudyQuest Application destroyed');
  }
}

// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
  try {
    new StudyQuestApp();
  } catch (error) {
    console.error('❌ Failed to initialize StudyQuest Application:', error);
    
    // フォールバック エラー表示
    const container = document.getElementById('answers') || document.body;
    container.innerHTML = `
      <div style="padding: 20px; text-align: center; color: #dc2626; font-family: sans-serif;">
        <h2>🚨 アプリケーション起動エラー</h2>
        <p>アプリケーションの初期化に失敗しました。</p>
        <p>エラー: ${error.message}</p>
        <button onclick="location.reload()" style="padding: 10px 20px; margin-top: 10px; cursor: pointer;">
          再読み込み
        </button>
      </div>
    `;
  }
});

// ページアンロード時のクリーンアップ
window.addEventListener('beforeunload', () => {
  if (window.studyQuestApp) {
    window.studyQuestApp.destroy();
  }
});

</script>