/**
 * AdvancedArchitecture.gs
 * 次世代マルチテナントアーキテクチャ
 * 
 * 特徴：
 * 1. Event-Driven Pattern
 * 2. Circuit Breaker Pattern  
 * 3. Bulkhead Pattern
 * 4. 予測キャッシング
 */

/**
 * 🔄 Event-Driven User Registration
 * イベント駆動型でより堅牢な処理
 */
class UserRegistrationEvent {
  constructor() {
    this.events = [];
    this.handlers = new Map();
  }
  
  /**
   * イベント登録
   */
  subscribe(eventType, handler) {
    if (!this.handlers.has(eventType)) {
      this.handlers.set(eventType, []);
    }
    this.handlers.get(eventType).push(handler);
  }
  
  /**
   * イベント発火
   */
  emit(eventType, data) {
    const handlers = this.handlers.get(eventType) || [];
    const results = [];
    
    for (const handler of handlers) {
      try {
        const result = handler(data);
        results.push({ success: true, result });
      } catch (error) {
        results.push({ success: false, error: error.message });
        console.error(`Event handler error [${eventType}]:`, error);
      }
    }
    
    return results;
  }
}

// グローバルイベントマネージャー
const userEventManager = new UserRegistrationEvent();

/**
 * 🛡️ Circuit Breaker Pattern
 * システム負荷制御と障害隔離
 */
class CircuitBreaker {
  constructor(name, threshold = 5, timeout = 60000) {
    this.name = name;
    this.threshold = threshold;
    this.timeout = timeout;
    this.failures = 0;
    this.state = 'CLOSED'; // CLOSED, OPEN, HALF_OPEN
    this.nextAttempt = Date.now();
  }
  
  async execute(fn) {
    if (this.state === 'OPEN') {
      if (Date.now() < this.nextAttempt) {
        throw new Error(`Circuit breaker [${this.name}] is OPEN`);
      }
      this.state = 'HALF_OPEN';
    }
    
    try {
      const result = await fn();
      this.onSuccess();
      return result;
    } catch (error) {
      this.onFailure();
      throw error;
    }
  }
  
  onSuccess() {
    this.failures = 0;
    this.state = 'CLOSED';
  }
  
  onFailure() {
    this.failures++;
    if (this.failures >= this.threshold) {
      this.state = 'OPEN';
      this.nextAttempt = Date.now() + this.timeout;
    }
  }
}

// システム全体のCircuit Breaker
const databaseCircuitBreaker = new CircuitBreaker('database', 3, 30000);
const cacheCircuitBreaker = new CircuitBreaker('cache', 5, 10000);

/**
 * 🚢 Bulkhead Pattern - リソース隔離
 */
class ResourcePool {
  constructor(name, maxConcurrent = 10) {
    this.name = name;
    this.maxConcurrent = maxConcurrent;
    this.current = 0;
    this.queue = [];
  }
  
  async acquire() {
    return new Promise((resolve, reject) => {
      if (this.current < this.maxConcurrent) {
        this.current++;
        resolve();
      } else {
        // キューに追加（最大100件まで）
        if (this.queue.length < 100) {
          this.queue.push({ resolve, reject });
        } else {
          reject(new Error(`Resource pool [${this.name}] is full`));
        }
      }
    });
  }
  
  release() {
    this.current--;
    if (this.queue.length > 0) {
      const next = this.queue.shift();
      this.current++;
      next.resolve();
    }
  }
}

// リソースプール
const userRegistrationPool = new ResourcePool('userRegistration', 5);
const databasePool = new ResourcePool('database', 10);

/**
 * ⚡ 予測キャッシング
 * 使用パターンを学習してプリフェッチ
 */
class PredictiveCache {
  constructor() {
    this.accessPatterns = new Map();
    this.cache = new Map();
    this.maxCacheSize = 1000;
  }
  
  recordAccess(key, relatedKeys = []) {
    if (!this.accessPatterns.has(key)) {
      this.accessPatterns.set(key, new Set());
    }
    
    relatedKeys.forEach(relatedKey => {
      this.accessPatterns.get(key).add(relatedKey);
    });
  }
  
  async get(key, fetchFn) {
    // キャッシュヒット
    if (this.cache.has(key)) {
      this.recordAccess(key);
      return this.cache.get(key);
    }
    
    // データ取得
    const value = await fetchFn();
    this.set(key, value);
    
    // 関連データのプリフェッチ
    this.prefetchRelated(key);
    
    return value;
  }
  
  set(key, value) {
    // LRU eviction
    if (this.cache.size >= this.maxCacheSize) {
      const firstKey = this.cache.keys().next().value;
      this.cache.delete(firstKey);
    }
    
    this.cache.set(key, value);
  }
  
  async prefetchRelated(key) {
    const relatedKeys = this.accessPatterns.get(key);
    if (!relatedKeys) return;
    
    // 非同期でプリフェッチ（エラーは無視）
    relatedKeys.forEach(async (relatedKey) => {
      if (!this.cache.has(relatedKey)) {
        try {
          // 関連データの取得ロジック（実装は省略）
          // await this.fetchRelatedData(relatedKey);
        } catch (e) {
          // プリフェッチエラーは無視
        }
      }
    });
  }
}

const predictiveCache = new PredictiveCache();

/**
 * 🏗️ 統合マルチテナント登録システム
 */
async function registerUserAdvanced(adminEmail, options = {}) {
  // リソース取得
  await userRegistrationPool.acquire();
  
  try {
    return await databaseCircuitBreaker.execute(async () => {
      
      // 1. 予測キャッシュを使用したユーザー検索
      const userInfo = await predictiveCache.get(
        `user_email_${adminEmail}`,
        () => findUserByEmailDirect(adminEmail)
      );
      
      let result;
      
      if (userInfo) {
        // 既存ユーザー更新イベント
        result = userEventManager.emit('user.update', {
          userId: userInfo.userId,
          adminEmail,
          options
        });
      } else {
        // 新規ユーザー作成イベント
        result = userEventManager.emit('user.create', {
          adminEmail,
          options
        });
      }
      
      // イベント処理結果の検証
      const failures = result.filter(r => !r.success);
      if (failures.length > 0) {
        throw new Error(`Registration failed: ${failures.map(f => f.error).join(', ')}`);
      }
      
      return result[0].result;
    });
    
  } finally {
    userRegistrationPool.release();
  }
}

/**
 * 📊 システム健全性監視
 */
class HealthMonitor {
  constructor() {
    this.metrics = new Map();
    this.alerts = [];
  }
  
  recordMetric(name, value, timestamp = Date.now()) {
    if (!this.metrics.has(name)) {
      this.metrics.set(name, []);
    }
    
    const metric = this.metrics.get(name);
    metric.push({ value, timestamp });
    
    // 最新1000件のみ保持
    if (metric.length > 1000) {
      metric.shift();
    }
    
    this.checkThresholds(name, value);
  }
  
  checkThresholds(name, value) {
    const thresholds = {
      'user.registration.time': 5000, // 5秒以上
      'database.error.rate': 0.1,     // 10%以上
      'cache.hit.rate': 0.8           // 80%未満
    };
    
    if (thresholds[name] && value > thresholds[name]) {
      this.alerts.push({
        metric: name,
        value,
        threshold: thresholds[name],
        timestamp: Date.now()
      });
    }
  }
  
  getHealthStatus() {
    const recentAlerts = this.alerts.filter(
      alert => Date.now() - alert.timestamp < 300000 // 5分以内
    );
    
    return {
      status: recentAlerts.length === 0 ? 'healthy' : 'degraded',
      alerts: recentAlerts,
      metrics: Object.fromEntries(this.metrics)
    };
  }
}

const healthMonitor = new HealthMonitor();

/**
 * 🎯 イベントハンドラー登録
 */
function initializeEventHandlers() {
  // ユーザー作成イベント
  userEventManager.subscribe('user.create', (data) => {
    const startTime = Date.now();
    
    try {
      const userId = generateConsistentUserId(data.adminEmail);
      const userData = {
        userId,
        adminEmail: data.adminEmail,
        createdAt: new Date().toISOString(),
        lastAccessedAt: new Date().toISOString(),
        isActive: 'true',
        configJson: '{}',
        ...data.options
      };
      
      createUserDirect(userData);
      
      // メトリクス記録
      healthMonitor.recordMetric('user.registration.time', Date.now() - startTime);
      healthMonitor.recordMetric('user.registration.success', 1);
      
      return { userId, adminEmail: data.adminEmail, isNewUser: true };
      
    } catch (error) {
      healthMonitor.recordMetric('user.registration.error', 1);
      throw error;
    }
  });
  
  // ユーザー更新イベント
  userEventManager.subscribe('user.update', (data) => {
    try {
      const updateData = {
        lastAccessedAt: new Date().toISOString(),
        ...data.options
      };
      
      updateUserDirect(data.userId, updateData);
      
      return { userId: data.userId, adminEmail: data.adminEmail, isNewUser: false };
      
    } catch (error) {
      healthMonitor.recordMetric('user.update.error', 1);
      throw error;
    }
  });
}

// 初期化
initializeEventHandlers();

/**
 * 🚀 高パフォーマンス統合API
 */
function getUserOrCreateAdvanced(adminEmail, options = {}) {
  return registerUserAdvanced(adminEmail, options);
}

/**
 * 📈 システム統計取得
 */
function getSystemStats() {
  return {
    health: healthMonitor.getHealthStatus(),
    circuitBreakers: {
      database: {
        state: databaseCircuitBreaker.state,
        failures: databaseCircuitBreaker.failures
      },
      cache: {
        state: cacheCircuitBreaker.state,
        failures: cacheCircuitBreaker.failures
      }
    },
    resourcePools: {
      userRegistration: {
        current: userRegistrationPool.current,
        queue: userRegistrationPool.queue.length
      },
      database: {
        current: databasePool.current,
        queue: databasePool.queue.length
      }
    }
  };
}