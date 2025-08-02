# システムアーキテクチャ設計書

## 1. アーキテクチャ概要

### 1.1 システム全体図
```
┌─────────────────────────────────────────────────────────────┐
│                        Client Side                          │
├─────────────────┬─────────────────┬─────────────────────────┤
│   User Pages    │   Admin Panel   │   Login & Setup Pages  │
│                 │                 │                         │
│ • Page.html     │ • AdminPanel    │ • Login.html           │
│ • Question/     │   .html         │ • AppSetupPage.html    │
│   Answer UI     │ • Real-time     │ • User Authentication  │
│ • Reactions     │   Monitoring    │ • Initial Setup        │
│ • Real-time     │ • User Mgmt     │                         │
│   Updates       │ • Config Mgmt   │                         │
└─────────────────┴─────────────────┴─────────────────────────┘
                              │
                    ┌─────────▼─────────┐
                    │   Unified APIs    │
                    │                   │
                    │ • Authentication  │
                    │ • Data Operations │
                    │ • Cache Layer     │
                    │ • Error Handling  │
                    └─────────┬─────────┘
                              │
┌─────────────────────────────▼─────────────────────────────────┐
│                  Google Apps Script Core                     │
├─────────────────┬─────────────────┬─────────────────────────┤
│  Data Layer     │  Business Logic │    System Services      │
│                 │                 │                         │
│ • Database.gs   │ • UserMgmt.gs   │ • url.gs               │
│ • UserData.gs   │ • QuestionMgmt  │ • CacheManager.js      │
│ • SheetOps.gs   │   .gs           │ • ErrorBoundary.html   │
│ • FormOps.gs    │ • AdminOps.gs   │ • SecurityHeaders.html │
└─────────────────┴─────────────────┴─────────────────────────┘
                              │
┌─────────────────────────────▼─────────────────────────────────┐
│                    Google Services                           │
├─────────────────┬─────────────────┬─────────────────────────┤
│  Data Storage   │   Identity      │    Platform Services    │
│                 │                 │                         │
│ • Google Sheets │ • OAuth 2.0     │ • Drive API            │
│ • Google Forms  │ • Directory API │ • Script Service       │
│ • Drive Files   │ • Domain Auth   │ • Cache Service        │
└─────────────────┴─────────────────┴─────────────────────────┘
```

### 1.2 技術スタック
- **Frontend**: HTML5, CSS3, JavaScript (ES5+)
- **Backend**: Google Apps Script (V8 Runtime)
- **Database**: Google Sheets API
- **Authentication**: Google OAuth 2.0
- **Caching**: Google CacheService + In-Memory Maps
- **Styling**: Custom CSS + TailwindCSS (fallback)

## 2. レイヤー別詳細設計

### 2.1 プレゼンテーション層

#### 2.1.1 ページ構成
```
src/
├── Page.html                 # メインユーザーページ
├── AdminPanel.html          # 管理者パネル
├── Login.html              # ログインページ
├── AppSetupPage.html       # セットアップページ
└── Error.html             # エラーページ
```

#### 2.1.2 JavaScript モジュール構成
```
Frontend Modules:
├── Core Modules
│   ├── SharedUtilities.html      # 共通ユーティリティ
│   ├── UnifiedCache.js.html      # 統合キャッシュ
│   ├── ClientOptimizer.html      # クライアント最適化
│   └── ErrorBoundary.html        # エラーハンドリング
│
├── User Interface
│   ├── page.js.html             # メインページ機能
│   ├── login.js.html            # ログイン機能
│   └── UnifiedStyles.html       # 統合スタイル
│
└── Admin Interface
    ├── adminPanel-core.js.html   # 管理パネル中核
    ├── adminPanel-ui.js.html     # UI管理
    ├── adminPanel-api.js.html    # API通信
    ├── adminPanel-events.js.html # イベント処理
    ├── adminPanel-forms.js.html  # フォーム管理
    ├── adminPanel-status.js.html # ステータス管理
    └── adminPanel.js.html        # メイン制御
```

#### 2.1.3 CSS アーキテクチャ
```
Styling Architecture:
├── CSS Variables System
│   ├── Color Palette (Dark Theme)
│   ├── Typography Scale
│   ├── Spacing System
│   └── Animation Timings
│
├── Component System
│   ├── Button Variants
│   ├── Form Controls
│   ├── Card Components
│   ├── Modal System
│   └── Loading States
│
└── Responsive Design
    ├── Mobile-First Approach
    ├── Breakpoint System
    ├── Touch-Friendly Controls
    └── Performance Optimizations
```

### 2.2 アプリケーション層

#### 2.2.1 ビジネスロジック構成
```
Business Logic:
├── User Management
│   ├── Authentication Flow
│   ├── Authorization Checks
│   ├── Profile Management
│   └── Session Management
│
├── Content Management
│   ├── Question Operations
│   ├── Answer Operations
│   ├── Reaction System
│   └── Content Validation
│
├── Admin Operations
│   ├── System Configuration
│   ├── User Administration
│   ├── Content Moderation
│   └── Analytics & Reporting
│
└── System Services
    ├── URL Management
    ├── Cache Management
    ├── Error Handling
    └── Security Services
```

#### 2.2.2 API 設計
```
API Endpoints:
├── Authentication APIs
│   ├── doGet(e) - Entry Point Router
│   ├── authenticateUser(email)
│   ├── checkUserPermission(userId)
│   └── validateDomainAccess(domain)
│
├── Data APIs
│   ├── getUserData(userId)
│   ├── getQuestions(filters)
│   ├── submitQuestion(data)
│   ├── submitAnswer(data)
│   └── updateReactions(reactionData)
│
├── Admin APIs
│   ├── getSystemStatus()
│   ├── updateConfiguration(config)
│   ├── manageUsers(operation, data)
│   └── generateReports(type, params)
│
└── System APIs
    ├── getWebAppUrls(userId)
    ├── clearSystemCache()
    ├── performHealthCheck()
    └── exportData(format, filters)
```

### 2.3 データ層

#### 2.3.1 データベース設計
```
Google Sheets Database:
├── Users Sheet
│   ├── Columns: userId, email, name, domain, role, lastAccess
│   ├── Indexes: userId (Primary), email (Unique)
│   └── Constraints: email validation, domain restrictions
│
├── Questions Sheet
│   ├── Columns: questionId, title, content, authorId, createdAt, status
│   ├── Indexes: questionId (Primary), authorId (Foreign Key)
│   └── Constraints: content length limits, status validation
│
├── Answers Sheet
│   ├── Columns: answerId, questionId, content, authorId, createdAt, reactions
│   ├── Indexes: answerId (Primary), questionId (Foreign Key)
│   └── Constraints: content validation, reaction format
│
└── Configuration Sheet
    ├── Columns: configKey, configValue, lastUpdated
    ├── Indexes: configKey (Primary)
    └── Constraints: JSON validation for complex values
```

#### 2.3.2 データアクセス層
```
Data Access Layer:
├── Database Operations (Database.gs)
│   ├── CRUD Operations
│   ├── Batch Operations
│   ├── Transaction Management
│   └── Data Validation
│
├── Sheet Operations (SheetOps.gs)
│   ├── Sheet Access Control
│   ├── Column Management
│   ├── Row Operations
│   └── Data Formatting
│
├── Cache Layer (UnifiedCache.js)
│   ├── Multi-level Caching
│   ├── TTL Management
│   ├── Cache Invalidation
│   └── Performance Optimization
│
└── Data Models
    ├── User Model
    ├── Question Model
    ├── Answer Model
    └── Configuration Model
```

## 3. システム間連携

### 3.1 Google Services 統合

#### 3.1.1 Sheets API 連携
```javascript
// Sheet Operations Architecture
class SheetOperations {
  constructor(sheetId) {
    this.sheet = SpreadsheetApp.openById(sheetId);
    this.cache = new UnifiedCache('sheet_ops');
  }
  
  // Batch Read Operations
  batchRead(ranges) {
    const cached = this.cache.get('batch_' + ranges.join('|'));
    if (cached) return cached;
    
    const values = this.sheet.getRange(ranges).getValues();
    this.cache.put('batch_' + ranges.join('|'), values, 300);
    return values;
  }
  
  // Optimized Write Operations
  batchWrite(data) {
    // Transaction-like behavior
    const backup = this.createBackup();
    try {
      this.sheet.getRange(data.range).setValues(data.values);
      this.cache.invalidate('batch_*');
      return { success: true };
    } catch (error) {
      this.restoreBackup(backup);
      throw error;
    }
  }
}
```

#### 3.1.2 OAuth 2.0 統合
```javascript
// Authentication Flow
class AuthenticationManager {
  static authenticateUser(email) {
    const user = Session.getActiveUser();
    if (!user.getEmail()) {
      throw new Error('認証が必要です');
    }
    
    // Domain validation
    const domain = user.getEmail().split('@')[1];
    if (!this.isAllowedDomain(domain)) {
      throw new Error('許可されていないドメインです');
    }
    
    return {
      userId: user.getEmail(),
      domain: domain,
      role: this.getUserRole(user.getEmail())
    };
  }
}
```

### 3.2 キャッシュアーキテクチャ

#### 3.2.1 多層キャッシュシステム
```javascript
// Unified Cache Manager
class UnifiedCacheManager {
  constructor() {
    this.levels = {
      memory: new Map(),           // L1: In-memory cache
      script: CacheService.getScriptCache(), // L2: Script cache
      user: CacheService.getUserCache()      // L3: User cache
    };
  }
  
  get(key, factory, options = {}) {
    // L1: Memory cache
    if (this.levels.memory.has(key)) {
      return this.levels.memory.get(key);
    }
    
    // L2: Script cache
    const scriptCached = this.levels.script.get(key);
    if (scriptCached) {
      const parsed = JSON.parse(scriptCached);
      this.levels.memory.set(key, parsed);
      return parsed;
    }
    
    // L3: Factory function
    if (factory) {
      const value = factory();
      this.put(key, value, options.ttl || 3600);
      return value;
    }
    
    return null;
  }
  
  put(key, value, ttl = 3600) {
    // Store in all levels
    this.levels.memory.set(key, value);
    this.levels.script.put(key, JSON.stringify(value), ttl);
  }
}
```

### 3.3 セキュリティアーキテクチャ

#### 3.3.1 認証・認可フロー
```
Authentication Flow:
1. User Access Request
   ↓
2. OAuth 2.0 Validation (Google)
   ↓
3. Domain Restriction Check
   ↓
4. Role-Based Access Control
   ↓
5. Session Management
   ↓
6. Resource Authorization
```

#### 3.3.2 セキュリティ層実装
```javascript
// Security Layer
class SecurityManager {
  static validateRequest(request) {
    // CSRF Protection
    if (!this.validateCSRFToken(request.token)) {
      throw new SecurityError('Invalid CSRF token');
    }
    
    // Input Validation
    const sanitized = this.sanitizeInput(request.data);
    
    // Rate Limiting
    if (!this.checkRateLimit(request.user)) {
      throw new SecurityError('Rate limit exceeded');
    }
    
    return sanitized;
  }
  
  static sanitizeInput(data) {
    if (typeof data === 'string') {
      return data
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#x27;');
    }
    return data;
  }
}
```

## 4. パフォーマンス最適化

### 4.1 フロントエンド最適化

#### 4.1.1 リソース最適化
```javascript
// Client-Side Optimization
class ClientOptimizer {
  static init() {
    // Lazy Loading
    this.setupLazyLoading();
    
    // Resource Preloading
    this.preloadCriticalResources();
    
    // Script Optimization
    this.optimizeScriptLoading();
  }
  
  static setupLazyLoading() {
    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          this.loadComponent(entry.target);
          observer.unobserve(entry.target);
        }
      });
    });
    
    document.querySelectorAll('[data-lazy]').forEach(el => {
      observer.observe(el);
    });
  }
}
```

#### 4.1.2 状態管理最適化
```javascript
// State Management Optimization
class StateManager {
  constructor() {
    this.state = {};
    this.subscribers = new Map();
    this.batchedUpdates = [];
    this.updateTimer = null;
  }
  
  setState(updates) {
    // Batch updates for performance
    this.batchedUpdates.push(updates);
    
    if (!this.updateTimer) {
      this.updateTimer = setTimeout(() => {
        this.flushUpdates();
      }, 16); // 60fps
    }
  }
  
  flushUpdates() {
    const mergedUpdates = Object.assign({}, ...this.batchedUpdates);
    Object.assign(this.state, mergedUpdates);
    
    this.notifySubscribers(mergedUpdates);
    
    this.batchedUpdates = [];
    this.updateTimer = null;
  }
}
```

### 4.2 バックエンド最適化

#### 4.2.1 データベース最適化
```javascript
// Database Optimization
class DatabaseOptimizer {
  static optimizeQuery(operation, params) {
    // Query caching
    const cacheKey = `query_${operation}_${JSON.stringify(params)}`;
    const cached = cache.get(cacheKey);
    if (cached) return cached;
    
    // Batch operations
    if (Array.isArray(params)) {
      return this.batchOperation(operation, params);
    }
    
    // Single operation
    const result = this.executeOperation(operation, params);
    cache.put(cacheKey, result, 300);
    return result;
  }
  
  static batchOperation(operation, paramsArray) {
    const chunks = this.chunkArray(paramsArray, 100);
    const results = [];
    
    chunks.forEach(chunk => {
      const batchResult = this.executeBatch(operation, chunk);
      results.push(...batchResult);
    });
    
    return results;
  }
}
```

## 5. エラーハンドリング & ログ

### 5.1 エラー処理アーキテクチャ
```javascript
// Error Handling Architecture
class ErrorBoundary {
  static handle(error, context) {
    // Categorize error
    const category = this.categorizeError(error);
    
    // Log error
    this.logError(error, context, category);
    
    // Handle based on category
    switch (category) {
      case 'AUTHENTICATION':
        return this.handleAuthError(error);
      case 'AUTHORIZATION':
        return this.handleAuthzError(error);
      case 'DATA':
        return this.handleDataError(error);
      case 'SYSTEM':
        return this.handleSystemError(error);
      default:
        return this.handleGenericError(error);
    }
  }
}
```

### 5.2 ログシステム
```javascript
// Logging System
class LogManager {
  static log(level, message, context = {}) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      level: level,
      message: message,
      context: context,
      user: Session.getActiveUser().getEmail(),
      sessionId: this.getSessionId()
    };
    
    // Console logging
    console[level.toLowerCase()](message, context);
    
    // Persistent logging
    this.writeToSheet(logEntry);
    
    // Alert on critical errors
    if (level === 'ERROR' || level === 'FATAL') {
      this.sendAlert(logEntry);
    }
  }
}
```

## 6. 展開・運用アーキテクチャ

### 6.1 環境分離
```
Environment Architecture:
├── Development
│   ├── Local Testing
│   ├── Feature Branches
│   └── Integration Testing
│
├── Staging
│   ├── Pre-production Testing
│   ├── User Acceptance Testing
│   └── Performance Testing
│
└── Production
    ├── Live Environment
    ├── Monitoring & Alerting
    └── Backup & Recovery
```

### 6.2 モニタリング
```javascript
// System Monitoring
class SystemMonitor {
  static healthCheck() {
    return {
      timestamp: new Date().toISOString(),
      status: 'healthy',
      metrics: {
        memory: this.getMemoryUsage(),
        cache: this.getCacheStatus(),
        database: this.getDatabaseStatus(),
        performance: this.getPerformanceMetrics()
      }
    };
  }
}
```

このアーキテクチャ設計により、スケーラブルで保守性の高い、高性能なシステムを実現する。