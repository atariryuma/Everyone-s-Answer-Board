/**
 * 基盤クラス集
 * ConfigurationManagerとAccessControllerの定義
 * 2024年GASベストプラクティス準拠
 */

// ===============================
// ConfigurationManagerクラス（データベース専用版）
// ===============================

class ConfigurationManager {
  /**
   * ユーザー設定を取得（データベースから直接読み取り）
   * @param {string} userId ユーザーID
   * @return {Object|null} ユーザー設定オブジェクト
   */
  getUserConfig(userId) {
    if (!userId) return null;
    
    const user = DB.findUserById(userId);
    if (!user || !user.configJson) return null;
    
    try {
      return JSON.parse(user.configJson);
    } catch (e) {
      console.error(`設定JSON解析エラー (${userId}):`, e);
      return null;
    }
  }

  /**
   * 最適化されたユーザー情報を取得（動的フィールド付き）
   * @param {string} userId ユーザーID
   * @return {Object|null} 統合されたユーザー情報
   */
  getOptimizedUserInfo(userId) {
    if (typeof getOptimizedUserInfo === 'function') {
      return getOptimizedUserInfo(userId);
    } else {
      // フォールバック：従来の方法
      return this.getUserConfig(userId);
    }
  }

  /**
   * ユーザー設定を保存（データベースに直接書き込み）
   * @param {string} userId ユーザーID
   * @param {Object} config 設定オブジェクト
   * @return {boolean} 保存成功可否
   */
  setUserConfig(userId, config) {
    if (!userId || !config) return false;
    
    config.userId = userId;
    config.lastModified = new Date().toISOString();
    
    try {
      const updated = updateUser(userId, {
        configJson: JSON.stringify(config),
        lastAccessedAt: new Date().toISOString()
      });
      
      if (updated) {
        console.log(`設定更新完了: ${userId}`);
        return true;
      }
      return false;
    } catch (error) {
      console.error(`設定保存エラー (${userId}):`, error);
      return false;
    }
  }

  /**
   * ユーザー設定を削除
   * @param {string} userId ユーザーID
   * @return {boolean} 削除成功可否
   */
  removeUserConfig(userId) {
    if (!userId) return false;
    return this.setUserConfig(userId, {});
  }

  /**
   * 公開設定を取得（ゲストアクセス用）
   * @param {string} userId ユーザーID
   * @return {Object|null} 公開設定オブジェクト
   */
  getPublicConfig(userId) {
    const config = this.getUserConfig(userId);
    if (!config || !config.isPublic) return null;

    return {
      userId: userId,
      title: config.title || 'Answer Board',
      description: config.description || '',
      allowAnonymous: config.allowAnonymous || false,
      columns: config.columns || this.getDefaultColumns(),
      theme: config.theme || 'default'
    };
  }

  /**
   * デフォルト列設定を取得
   * @return {Array} デフォルト列配列
   */
  getDefaultColumns() {
    return [
      { name: 'timestamp', label: 'タイムスタンプ', type: 'datetime', required: false },
      { name: 'email', label: 'メールアドレス', type: 'email', required: false },
      { name: 'class', label: 'クラス', type: 'text', required: false },
      { name: 'opinion', label: '回答', type: 'textarea', required: true },
      { name: 'reason', label: '理由', type: 'textarea', required: false },
      { name: 'name', label: '名前', type: 'text', required: false }
    ];
  }

  /**
   * 新規ユーザー設定を初期化
   * @param {string} userId ユーザーID
   * @param {string} email ユーザーメール
   * @return {Object} 初期化された設定
   */
  initializeUserConfig(userId, email) {
    // ConfigOptimizerの最適化形式を使用（重複データを除去）
    const optimizedConfig = {
      title: `${email}の回答ボード`,
      setupStatus: 'pending',
      formCreated: false,
      appPublished: false,
      isPublic: false,
      allowAnonymous: false,
      sheetName: null,
      columnMapping: {},
      theme: 'default',
      lastModified: new Date().toISOString()
    };

    // columnsは必要時のみ保持（デフォルト列設定）
    optimizedConfig.columns = this.getDefaultColumns();

    console.info('Base.gs 初期化: 最適化済みconfigJSON使用', {
      userId,
      email,
      optimizedSize: JSON.stringify(optimizedConfig).length,
      removedFields: ['userId', 'userEmail', 'createdAt', 'description'] // DB列に移行済み
    });

    const success = this.setUserConfig(userId, optimizedConfig);
    return success ? optimizedConfig : null;
  }

  /**
   * 設定を更新
   * @param {string} userId ユーザーID
   * @param {Object} updates 更新内容
   * @return {boolean} 更新成功可否
   */
  updateUserConfig(userId, updates) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    const updatedConfig = {
      ...currentConfig,
      ...updates,
      lastModified: new Date().toISOString()
    };

    return this.setUserConfig(userId, updatedConfig);
  }

  /**
   * 設定の存在確認
   * @param {string} userId ユーザーID
   * @return {boolean} 存在可否
   */
  hasUserConfig(userId) {
    const config = this.getUserConfig(userId);
    return config !== null && Object.keys(config).length > 0;
  }
}

// ===============================
// AccessControllerクラス
// ===============================

class AccessController {
  constructor(configManager) {
    this.configManager = configManager;
  }

  /**
   * ユーザーアクセスを検証
   * @param {string} targetUserId 対象ユーザーID
   * @param {string} mode アクセスモード (view, edit, admin)
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {Object} アクセス結果
   */
  verifyAccess(targetUserId, mode = 'view', currentUserEmail = null) {
    try {
      if (!targetUserId) {
        return this.createAccessResult(false, 'invalid', null, '無効なユーザーIDです');
      }

      const config = this.configManager.getUserConfig(targetUserId);
      
      if (!config) {
        return this.createAccessResult(false, 'not_found', null, '指定されたユーザーが見つかりません');
      }

      switch (mode) {
        case 'admin':
          return this.verifyAdminAccess(config, currentUserEmail);
        case 'edit':
          return this.verifyEditAccess(config, currentUserEmail);
        case 'view':
        default:
          return this.verifyViewAccess(config, currentUserEmail);
      }

    } catch (error) {
      console.error('アクセス検証エラー:', error);
      return this.createAccessResult(false, 'error', null, 'アクセス検証でエラーが発生しました');
    }
  }

  /**
   * 管理者アクセスを検証
   */
  verifyAdminAccess(config, currentUserEmail) {
    const systemAdminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    
    // システム管理者
    if (currentUserEmail === systemAdminEmail) {
      return this.createAccessResult(true, 'system_admin', config);
    }
    
    // オーナー
    if (currentUserEmail === config.userEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    return this.createAccessResult(false, 'guest', null, '管理者権限がありません');
  }

  /**
   * 編集アクセスを検証
   */
  verifyEditAccess(config, currentUserEmail) {
    // オーナーか確認
    if (currentUserEmail === config.userEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    // 他のユーザーは編集不可
    return this.createAccessResult(false, 'guest', null, '編集権限がありません');
  }

  /**
   * 閲覧アクセスを検証
   */
  verifyViewAccess(config, currentUserEmail) {
    // オーナーは常に閲覧可能
    if (currentUserEmail === config.userEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    // 公開設定を確認
    if (config.isPublic) {
      return this.createAccessResult(true, 'guest', config);
    }

    // 匿名アクセス許可を確認
    if (config.allowAnonymous) {
      return this.createAccessResult(true, 'anonymous', config);
    }

    // デフォルトは閲覧不可
    return this.createAccessResult(false, 'guest', null, '閲覧権限がありません');
  }

  /**
   * アクセス結果オブジェクトを作成
   */
  createAccessResult(allowed, userType, config = null, message = null) {
    return {
      allowed,
      userType,
      config,
      message,
      timestamp: new Date().toISOString()
    };
  }

  /**
   * ユーザーレベルを取得
   */
  getUserLevel(targetUserId, currentUserEmail) {
    if (!targetUserId || !currentUserEmail) return 'none';

    const config = this.configManager.getUserConfig(targetUserId);
    if (!config) return 'none';

    const systemAdminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    
    if (currentUserEmail === systemAdminEmail) {
      return 'system_admin';
    }
    
    if (currentUserEmail === config.userEmail) {
      return 'owner';
    }

    if (Session.getActiveUser().getEmail() === currentUserEmail) {
      return 'authenticated_user';
    }

    return 'guest';
  }
}