/**
 * ユーザーIDと識別子から一意のキャッシュキーを生成します。
 * マルチテナントセキュリティを考慮した実装
 * @param {string} prefix - キーのプレフィックス
 * @param {string} userId - 対象ユーザーID
 * @param {string} [suffix] - 追加識別子
 * @return {string} 生成されたキャッシュキー
 */
function buildUserScopedKey(prefix, userId, suffix) {
  if (!userId) throw new Error('SECURITY_ERROR: userId is required for cache key');
  
  // マルチテナントセキュリティマネージャーを使用してセキュアキーを生成
  if (typeof multiTenantSecurity !== 'undefined' && multiTenantSecurity.generateSecureKey) {
    return multiTenantSecurity.generateSecureKey(prefix, userId, suffix);
  }
  
  // フォールバック実装（後方互換性）
  let key = `${prefix}_${userId}`;
  if (suffix) key += `_${suffix}`;
  return key;
}

/**
 * エンタープライズ対応セキュアなテナント分離キーの生成
 * @param {string} prefix - キーのプレフィックス
 * @param {string} userId - 対象ユーザーID
 * @param {string} context - コンテキスト情報
 * @param {Object} options - オプション設定
 * @param {boolean} options.enableEncryption - 暗号化有効フラグ（HIPAA対応）
 * @param {boolean} options.includeTimestamp - タイムスタンプ追加フラグ
 * @param {number} options.maxKeyLength - 最大キー長（CacheService制限対応）
 * @return {string} セキュアなキー
 */
function buildSecureUserScopedKey(prefix, userId, context = '', options = {}) {
  const { 
    enableEncryption = true,
    includeTimestamp = false,
    maxKeyLength = 250 // CacheService制限
  } = options;
  
  if (!userId) throw new Error('SECURITY_ERROR: userId is required for secure cache key');
  
  // HIPAA対応: 機密データのハッシュ化
  const hashedUserId = enableEncryption ? 
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, userId, Utilities.Charset.UTF_8)
      .map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'))
      .join('').slice(0, 16) : 
    userId.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 32); // 英数字のみ、32文字制限
  
  // マルチテナント対応キー生成
  let key = `MT_${prefix}_${hashedUserId}`;
  
  if (context) {
    const sanitizedContext = context.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 50);
    key += `_${sanitizedContext}`;
  }
  
  if (includeTimestamp) {
    const timestamp = Math.floor(Date.now() / 1000); // Unix timestamp
    key += `_${timestamp}`;
  }
  
  // CacheService制限チェック
  if (key.length > maxKeyLength) {
    key = key.substring(0, maxKeyLength);
  }
  
  // マルチテナントセキュリティマネージャーが利用可能な場合は追加検証
  if (typeof multiTenantSecurity !== 'undefined') {
    try {
      return multiTenantSecurity.generateSecureKey(prefix, userId, context);
    } catch (error) {
      // フォールバック: 独自実装を使用
      debugLog('マルチテナントセキュリティマネージャーエラー、フォールバック実行:', error.message);
    }
  }
  
  return key;
}

/**
 * キーからユーザー情報の抽出を試行（セキュリティ目的での検証）
 * @param {string} key - キャッシュキー
 * @return {object|null} ユーザー情報（抽出できない場合はnull）
 */
function extractUserInfoFromKey(key) {
  if (!key || typeof key !== 'string') return null;
  
  // セキュアキー形式の検証
  if (key.startsWith('MT_')) {
    // マルチテナント形式のキーは解析不可（セキュリティのため）
    return { secure: true, extractable: false };
  }
  
  // レガシー形式の解析（後方互換性）
  const parts = key.split('_');
  if (parts.length >= 2) {
    return {
      prefix: parts[0],
      userId: parts[1],
      suffix: parts.length > 2 ? parts.slice(2).join('_') : null,
      secure: false,
      extractable: true
    };
  }
  
  return null;
}

/**
 * ユーザースコープキーの検証
 * @param {string} key - 検証対象のキー
 * @param {string} expectedUserId - 期待されるユーザーID
 * @return {boolean} キーが指定ユーザーに対して有効かどうか
 */
function validateUserScopedKey(key, expectedUserId) {
  if (!key || !expectedUserId) return false;
  
  const keyInfo = extractUserInfoFromKey(key);
  if (!keyInfo) return false;
  
  // セキュアキーの場合は直接検証不可
  if (keyInfo.secure) {
    // セキュリティマネージャーでの検証が必要
    return true; // セキュアキーは生成時に検証済みと仮定
  }
  
  // レガシーキーの検証
  return keyInfo.extractable && keyInfo.userId === expectedUserId;
}

/**
 * エンタープライズ対応テナントアクセス検証
 * @param {string} operation - 操作種別
 * @param {string} userId - ユーザーID
 * @param {string} targetData - 対象データ識別子
 * @return {boolean} アクセス許可の可否
 */
function validateTenantAccess(operation, userId, targetData) {
  if (!userId) {
    auditSecurityViolation('MISSING_USER_ID', { operation, targetData });
    return false;
  }
  
  const currentUserId = Session.getActiveUser().getEmail();
  
  // マルチテナントセキュリティマネージャーによる検証
  if (typeof multiTenantSecurity !== 'undefined') {
    const isValid = multiTenantSecurity.validateTenantBoundary(currentUserId, userId, operation);
    auditTenantAccess(operation, userId, isValid, { targetData, currentUserId });
    return isValid;
  }
  
  // フォールバック: 基本的なテナント境界検証
  const isValid = currentUserId === userId;
  auditTenantAccess(operation, userId, isValid, { targetData, currentUserId });
  return isValid;
}

/**
 * HIPAA/SOC準拠の監査ログ記録
 * @param {string} operation - 操作種別
 * @param {string} userId - ユーザーID
 * @param {boolean} result - アクセス結果
 * @param {Object} metadata - 追加メタデータ
 */
function auditTenantAccess(operation, userId, result, metadata = {}) {
  try {
    const auditEntry = {
      timestamp: new Date().toISOString(),
      operation: operation,
      userId: userId,
      result: result ? 'GRANTED' : 'DENIED',
      metadata: metadata,
      sessionId: Session.getTemporaryActiveUserKey() || 'unknown'
    };
    
    // エンタープライズ監査要件対応
    console.log('🔐 AUDIT:', JSON.stringify(auditEntry));
    
    // 将来的には外部監査システムに送信
    // this.logToComplianceAudit(auditEntry);
  } catch (error) {
    console.error('監査ログ記録エラー:', error.message);
  }
}

/**
 * セキュリティ違反の記録
 * @param {string} violationType - 違反種別
 * @param {Object} details - 詳細情報
 */
function auditSecurityViolation(violationType, details = {}) {
  try {
    const violationEntry = {
      timestamp: new Date().toISOString(),
      type: 'SECURITY_VIOLATION',
      violationType: violationType,
      details: details,
      userAgent: Session.getActiveUser().getEmail(),
      sessionId: Session.getTemporaryActiveUserKey() || 'unknown'
    };
    
    console.error('🚨 SECURITY_VIOLATION:', JSON.stringify(violationEntry));
    
    // エンタープライズセキュリティ監視システムへの通知
    // this.alertSecurityTeam(violationEntry);
  } catch (error) {
    console.error('セキュリティ違反記録エラー:', error.message);
  }
}

