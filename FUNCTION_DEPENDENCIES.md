# 関数依存関係・初期化順序定義書

## 1. 概要

このドキュメントは、Everyone's Answer Board システムにおける JavaScript 関数の依存関係と初期化順序を定義し、`ReferenceError: function is not defined` エラーを防ぐための指針を提供する。

## 2. モジュール読み込み順序

### 2.1 AdminPanel.html での読み込み順序
```html
<!-- 順序1: セキュリティ・基盤 -->
<?!= include('SharedSecurityHeaders'); ?>

<!-- 順序2: スタイリング・設定 -->
<?!= include('SharedTailwindConfig'); ?>
<?!= include('UnifiedStyles'); ?>

<!-- 順序3: 共通ユーティリティ -->
<?!= include('SharedUtilities'); ?>
<?!= include('adminPanel-layout.css'); ?>

<!-- 順序4: クライアント最適化・キャッシュ -->
<?!= include('ClientOptimizer'); ?>
<?!= include('UnifiedCache.js'); ?>

<!-- 順序5: 共通モーダル -->
<?!= include('SharedModals'); ?>

<!-- 順序6: 管理パネル中核 (最重要) -->
<?!= include('adminPanel-core.js'); ?>

<!-- 順序7: ステータス管理 -->
<?!= include('adminPanel-status.js'); ?>

<!-- 順序8: API 通信 -->
<?!= include('adminPanel-api.js'); ?>

<!-- 順序9: UI 管理 -->
<?!= include('adminPanel-ui.js'); ?>

<!-- 順序10: フォーム管理 -->
<?!= include('adminPanel-forms.js'); ?>

<!-- 順序11: イベント処理 -->
<?!= include('adminPanel-events.js'); ?>

<!-- 順序12: メイン制御 (最後) -->
<?!= include('adminPanel.js'); ?>
```

### 2.2 依存関係の理由
1. **adminPanel-core.js が最初**: 全ての共通変数・関数を定義
2. **adminPanel-status.js が早期**: 他のモジュールがステータス関数を使用
3. **adminPanel-api.js が中間**: UI がAPIを呼び出すため
4. **adminPanel-ui.js が後期**: 他のモジュールの関数を使用
5. **adminPanel.js が最後**: 全ての初期化後にメインロジックを実行

## 3. 重要な関数依存関係マップ

### 3.1 グローバル変数依存関係
```
adminPanel-core.js.html で定義される重要な変数:
├── currentStatus (全モジュールで使用)
├── currentConfig (設定管理で使用)
├── selectedSheet (シート選択で使用)
├── selectedSheetId (API呼び出しで使用)
├── processingCache (キャッシュシステム)
├── cacheManager (統合キャッシュ)
└── isInitializing (初期化制御)

依存モジュール:
├── adminPanel-status.js → currentStatus
├── adminPanel-api.js → selectedSheetId, processingCache
├── adminPanel-ui.js → currentStatus, currentConfig, selectedSheet
├── adminPanel-events.js → currentStatus, selectedSheetId
├── adminPanel-forms.js → currentConfig
└── adminPanel.js → 全ての変数
```

### 3.2 関数呼び出し依存関係

#### 3.2.1 初期化関数の依存関係
```
DOMContentLoaded イベント処理:
┌─────────────────────────────────────┐
│ adminPanel-core.js                  │
│ └── initializeAdminPanelMaster()    │ ← 最初に実行される
│     ├── setGlobalLoadingState()     │
│     ├── initializeCache()           │
│     └── loadStatusAndInitialize()   │
└─────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────┐
│ adminPanel-ui.js                    │
│ └── DOMContentLoaded Handler        │ ← 2番目に実行
│     ├── initializeStepNavigation()  │
│     ├── setupEventListeners()       │
│     └── updateUIWithNewStatus()     │
└─────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────┐
│ adminPanel-forms.js                 │
│ └── DOMContentLoaded Handler        │ ← 3番目に実行
│     └── setupFormHandlers()         │
└─────────────────────────────────────┘
```

#### 3.2.2 UI 関数の依存関係
```
UI関数呼び出しチェーン:
updateUIWithNewStatus()
├── navigateToStep() ← adminPanel-ui.js で定義
├── showFormConfigModal() ← adminPanel-ui.js で定義
├── hideFormConfigModal() ← adminPanel-ui.js で定義
├── toggleSection() ← adminPanel-ui.js で定義
└── updateStatusDisplay() ← adminPanel-ui.js で定義

呼び出し元:
├── adminPanel-core.js → loadStatusAndInitialize()
├── adminPanel-api.js → API レスポンス処理
├── adminPanel-events.js → イベントハンドラー
└── adminPanel.js → ステップ変更処理
```

#### 3.2.3 API 関数の依存関係
```
API関数呼び出しチェーン:
runGasWithUserId() ← adminPanel-api.js で定義
├── 呼び出し元:
│   ├── adminPanel-ui.js → loadStatus()
│   ├── adminPanel-events.js → saveConfiguration()
│   ├── adminPanel-forms.js → フォーム送信
│   └── adminPanel.js → データ保存・公開
│
├── 依存関数:
│   ├── getWebAppUrls() ← url.gs で定義
│   ├── showGlobalLoadingOverlay() ← SharedUtilities.html で定義
│   └── hideGlobalLoadingOverlay() ← SharedUtilities.html で定義
│
└── キャッシュ関数:
    ├── executeWithCache() ← adminPanel-core.js で定義
    └── cacheManager.get() ← UnifiedCache.js で定義
```

## 4. 初期化タイミング制御

### 4.1 DOMContentLoaded ハンドラーの優先順位
```javascript
// adminPanel-core.js.html (最高優先度)
if (document.readyState !== 'loading') {
  initializeAdminPanelMaster();
} else {
  document.addEventListener('DOMContentLoaded', initializeAdminPanelMaster);
}

// adminPanel-ui.js.html (中優先度)
document.addEventListener('DOMContentLoaded', function() {
  // adminPanel-core.js の初期化を待つ
  waitForCoreInitialization().then(() => {
    initializeStepNavigation();
    setupEventListeners();
  });
});

// adminPanel-forms.js.html (低優先度) 
document.addEventListener('DOMContentLoaded', function() {
  // UI 初期化を待つ
  waitForUIInitialization().then(() => {
    setupFormHandlers();
  });
});
```

### 4.2 非同期初期化の実装
```javascript
// 依存関数の存在チェック
function waitForFunction(functionName, maxRetries = 50) {
  return new Promise((resolve, reject) => {
    let retries = 0;
    
    function check() {
      if (typeof window[functionName] === 'function') {
        resolve();
      } else if (retries < maxRetries) {
        retries++;
        setTimeout(check, 100); // 100ms間隔でチェック
      } else {
        reject(new Error(`Function ${functionName} not found after ${maxRetries} retries`));
      }
    }
    
    check();
  });
}

// 使用例
async function safelyCallUIFunction() {
  try {
    await waitForFunction('updateUIWithNewStatus');
    updateUIWithNewStatus(currentStatus);
  } catch (error) {
    console.error('UI function not available:', error);
  }
}
```

## 5. エラー防止策

### 5.1 関数存在チェック
```javascript
// 安全な関数呼び出しパターン
function safeCall(functionName, ...args) {
  if (typeof window[functionName] === 'function') {
    return window[functionName](...args);
  } else {
    console.error(`Function ${functionName} is not defined`);
    return null;
  }
}

// 使用例
safeCall('navigateToStep', 2);
safeCall('updateUIWithNewStatus', status);
```

### 5.2 遅延初期化パターン
```javascript
// 遅延初期化マネージャー
class DeferredInitializer {
  constructor() {
    this.pendingCalls = [];
    this.initialized = false;
  }
  
  defer(functionName, ...args) {
    if (this.initialized && typeof window[functionName] === 'function') {
      return window[functionName](...args);
    } else {
      this.pendingCalls.push({ functionName, args });
    }
  }
  
  flush() {
    this.initialized = true;
    this.pendingCalls.forEach(call => {
      if (typeof window[call.functionName] === 'function') {
        window[call.functionName](...call.args);
      }
    });
    this.pendingCalls = [];
  }
}

// グローバルインスタンス
window.deferredInit = new DeferredInitializer();
```

### 5.3 関数定義ガード
```javascript
// 関数定義時のガード
function defineFunction(name, fn) {
  if (typeof window[name] === 'undefined') {
    window[name] = fn;
  } else {
    console.warn(`Function ${name} already defined, skipping`);
  }
}

// 使用例
defineFunction('navigateToStep', function(step) {
  // 実装
});
```

## 6. 重要な関数一覧と定義場所

### 6.1 Core Functions (adminPanel-core.js.html)
```javascript
Global Variables:
├── currentStatus          // システム全体のステータス
├── currentConfig          // 現在の設定
├── selectedSheet          // 選択されたシート
├── selectedSheetId        // 選択されたシートID
├── processingCache        // 処理キャッシュマップ
└── cacheManager           // 統合キャッシュマネージャー

Core Functions:
├── initializeAdminPanelMaster()     // メイン初期化
├── setGlobalLoadingState()          // ローディング状態設定
├── executeWithCache()               // キャッシュ付き実行
├── createGlobalLoadingOverlay()     // グローバルローディング作成
└── loadStatusAndInitialize()        // ステータス読み込み＆初期化
```

### 6.2 UI Functions (adminPanel-ui.js.html)
```javascript
UI Management:
├── updateUIWithNewStatus(status)    // ステータスによるUI更新
├── navigateToStep(step)             // ステップナビゲーション
├── showFormConfigModal()            // モーダル表示
├── hideFormConfigModal()            // モーダル非表示
├── toggleSection(sectionId)         // セクション開閉
└── initializeStepNavigation()       // ステップナビ初期化

State Management:
├── managedSectionStates()           // セクション状態管理
├── updateStepValidation()           // ステップ検証更新
└── refreshUIElements()              // UI要素リフレッシュ
```

### 6.3 API Functions (adminPanel-api.js.html)
```javascript
API Communication:
├── runGasWithUserId(action, message, ...args)  // GAS API呼び出し
├── loadStatus()                                // ステータス読み込み
├── saveConfiguration(config)                   // 設定保存
└── proceedWithSaveAndPublish(config)          // 保存・公開実行

Data Processing:
├── normalizeConfigJson(data)        // 設定正規化
├── validateSheetStructure(sheet)    // シート構造検証
└── generateFormUrl(config)          // フォームURL生成
```

### 6.4 Event Functions (adminPanel-events.js.html)
```javascript
Event Handlers:
├── setupSheetSelectionHandlers()    // シート選択イベント
├── setupConfigurationHandlers()     // 設定イベント
├── handleSheetSelection(event)      // シート選択処理
└── handleConfigurationChange(event) // 設定変更処理

Form Processing:
├── processFormSubmission(form)      // フォーム送信処理
├── validateFormData(data)           // フォームデータ検証
└── showFormConfigModalSafely()      // 安全なモーダル表示
```

## 7. 初期化エラーの診断方法

### 7.1 エラー診断チェックリスト
```javascript
// 初期化診断関数
function diagnoseInitializationErrors() {
  const diagnostics = {
    timestamp: new Date().toISOString(),
    domState: document.readyState,
    criticalFunctions: {},
    criticalVariables: {},
    errors: []
  };
  
  // 重要な関数の存在チェック
  const criticalFunctions = [
    'updateUIWithNewStatus',
    'navigateToStep', 
    'showFormConfigModal',
    'hideFormConfigModal',
    'toggleSection',
    'runGasWithUserId'
  ];
  
  criticalFunctions.forEach(func => {
    diagnostics.criticalFunctions[func] = typeof window[func] === 'function';
    if (typeof window[func] !== 'function') {
      diagnostics.errors.push(`Missing function: ${func}`);
    }
  });
  
  // 重要な変数の存在チェック
  const criticalVariables = [
    'currentStatus',
    'currentConfig', 
    'selectedSheet',
    'selectedSheetId',
    'processingCache'
  ];
  
  criticalVariables.forEach(variable => {
    diagnostics.criticalVariables[variable] = typeof window[variable] !== 'undefined';
    if (typeof window[variable] === 'undefined') {
      diagnostics.errors.push(`Missing variable: ${variable}`);
    }
  });
  
  console.log('Initialization Diagnostics:', diagnostics);
  return diagnostics;
}
```

### 7.2 実行時エラー回復
```javascript
// エラー回復メカニズム
function recoverFromInitializationError(error) {
  console.error('Initialization error detected:', error);
  
  // 部分的な再初期化を試行
  setTimeout(() => {
    if (typeof initializeAdminPanelMaster === 'function') {
      console.log('Attempting recovery initialization...');
      initializeAdminPanelMaster();
    }
  }, 1000);
  
  // ユーザーへの通知
  showMessage('システムの初期化中にエラーが発生しました。ページを再読み込みしてください。', 'error');
}

// グローバルエラーハンドラー
window.addEventListener('error', function(event) {
  if (event.error && event.error.message.includes('is not defined')) {
    recoverFromInitializationError(event.error);
  }
});
```

## 8. ベストプラクティス

### 8.1 関数定義の推奨パターン
```javascript
// ✅ 推奨: 存在チェック付き定義
if (typeof updateUIWithNewStatus === 'undefined') {
  function updateUIWithNewStatus(status) {
    // 実装
  }
}

// ✅ 推奨: 即座に利用可能にする
window.updateUIWithNewStatus = function(status) {
  // 実装
};

// ❌ 非推奨: ブロックスコープ内での定義
if (someCondition) {
  function someFunction() {
    // ホイスティングの問題あり
  }
}
```

### 8.2 モジュール間通信パターン
```javascript
// ✅ 推奨: イベントベース通信
function notifyStatusChange(status) {
  window.dispatchEvent(new CustomEvent('statusChanged', { 
    detail: { status } 
  }));
}

// ✅ 推奨: コールバック登録パターン
const callbacks = {
  onStatusUpdate: [],
  onConfigChange: []
};

function registerCallback(event, callback) {
  if (callbacks[event]) {
    callbacks[event].push(callback);
  }
}

// ❌ 非推奨: 直接的な相互参照
// moduleA.someFunction() ← 依存関係が複雑になる
```

### 8.3 初期化タイミングの制御
```javascript
// ✅ 推奨: 段階的初期化
const initializationSteps = [
  'core',
  'ui', 
  'api',
  'events'
];

let currentInitStep = 0;

function proceedToNextInitStep() {
  if (currentInitStep < initializationSteps.length) {
    const step = initializationSteps[currentInitStep];
    console.log(`Initializing step: ${step}`);
    
    switch(step) {
      case 'core':
        initializeCore();
        break;
      case 'ui':
        initializeUI();
        break;
      // ... 他のステップ
    }
    
    currentInitStep++;
  }
}
```

この依存関係定義に従うことで、関数未定義エラーを防ぎ、安定したシステム初期化を実現できる。