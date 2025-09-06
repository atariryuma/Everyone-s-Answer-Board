/**
 * ConfigManager.gs - 統一configJSON管理システム
 *
 * 🎯 目的: 全configJson操作の単一責任者として論理整合性を確保
 * 🚀 CLAUDE.md完全準拠: configJSON中心型アーキテクチャの統一実装
 *
 * 責任範囲:
 * - configJsonの読み込み・保存・更新・削除
 * - 設定構築・検証・サニタイズ
 * - 既存システムとの互換性維持
 */

/**
 * ConfigManager - 統一configJSON管理システム
 * 全configJson操作の唯一の責任者
 */
const ConfigManager = Object.freeze({
  // ========================================
  // 📖 読み込み系メソッド
  // ========================================

  /**
   * ユーザー設定取得（統合版）
   * @param {string} userId - ユーザーID
   * @returns {Object|null} 統合設定オブジェクト
   */
  getUserConfig(userId) {
    if (!userId || !this.validateUserId(userId)) {
      console.warn('ConfigManager.getUserConfig: 無効なuserID:', userId);
      return null;
    }

    try {
      const user = DB.findUserById(userId);
      if (!user) {
        console.warn('ConfigManager.getUserConfig: ユーザーが見つかりません:', userId);
        // デフォルト設定を返す（CLAUDE.md準拠：心理的安全性重視）
        return {
          setupStatus: 'pending',
          appPublished: false,
          displaySettings: { showNames: false, showReactions: false },
          userId: userId,
        };
      }

      // 無限再帰回避: 直接JSONパース（user.parsedConfigではなく生configJsonから）
      let baseConfig;
      try {
        baseConfig = JSON.parse(user.configJson || '{}');
      } catch (parseError) {
        console.warn('ConfigManager.getUserConfig: JSON解析エラー:', parseError.message);
        baseConfig = {};
      }

      // 🔧 自動修復機能: configJsonが二重になっていたら修正
      if (baseConfig.configJson) {
        console.warn('⚠️ ConfigManager.getUserConfig: 二重構造を検出 - 自動修復開始');

        if (typeof baseConfig.configJson === 'string') {
          try {
            // ネストしたconfigJsonを展開
            const nestedConfig = JSON.parse(baseConfig.configJson);

            // 内側のデータを外側にマージ（内側優先）
            baseConfig = { ...baseConfig, ...nestedConfig };

            // configJsonフィールドを削除
            delete baseConfig.configJson;
            delete baseConfig.configJSON;

            // 修復したデータをDBに保存
            this.saveConfig(userId, baseConfig);

            console.log('✅ ConfigManager.getUserConfig: 二重構造を自動修復完了', {
              userId: userId,
              fixedFields: Object.keys(baseConfig),
            });
          } catch (parseError) {
            console.error(
              '❌ ConfigManager.getUserConfig: ネストしたconfigJson解析エラー',
              parseError.message
            );
            // パースできない場合はフィールドだけ削除
            delete baseConfig.configJson;
          }
        } else {
          // 文字列でない場合も削除
          delete baseConfig.configJson;
        }
      }

      const enhancedConfig = this.enhanceConfigWithDynamicUrls(baseConfig, userId);

      console.log('✅ ConfigManager.getUserConfig: 設定取得完了', {
        userId,
        configFields: Object.keys(enhancedConfig).length,
        hasSpreadsheetId: !!enhancedConfig.spreadsheetId,
        spreadsheetId: enhancedConfig.spreadsheetId
          ? `${enhancedConfig.spreadsheetId.substring(0, 8)}...`
          : 'null',
        configKeys: Object.keys(enhancedConfig),
      });

      return enhancedConfig;
    } catch (error) {
      console.error('❌ ConfigManager.getUserConfig: エラー', {
        userId,
        error: error.message,
        stack: error.stack,
      });
      return null;
    }
  },

  /**
   * 生configJson取得（デバッグ・移行用）
   * @param {string} userId - ユーザーID
   * @returns {string|null} JSON文字列
   */
  getRawConfig(userId) {
    if (!userId) return null;

    const user = DB.findUserById(userId);
    return user?.configJson || null;
  },

  // ========================================
  // 💾 保存系メソッド（統一）
  // ========================================

  /**
   * 設定保存（唯一の保存メソッド）
   * @param {string} userId - ユーザーID
   * @param {Object} config - 保存する設定オブジェクト
   * @returns {boolean} 保存成功可否
   */
  saveConfig(userId, config) {
    if (!userId || !config) {
      console.error('ConfigManager.saveConfig: 無効な引数', { userId: !!userId, config: !!config });
      return false;
    }

    try {
      // 🚫 二重構造防止（第2層防御）: configJsonフィールドを強制削除
      const cleanConfig = { ...config };
      delete cleanConfig.configJson;
      delete cleanConfig.configJSON;

      // 大文字小文字のバリエーションも削除
      Object.keys(cleanConfig).forEach((key) => {
        if (key.toLowerCase() === 'configjson') {
          console.warn(`⚠️ ConfigManager.saveConfig: 危険なフィールド "${key}" を削除`);
          delete cleanConfig[key];
        }
      });

      // 🔍 デバッグ: cleanConfig詳細ログ
      console.log('🔍 ConfigManager.saveConfig: cleanConfig詳細', {
        userId,
        cleanConfigKeys: Object.keys(cleanConfig),
        spreadsheetId: cleanConfig.spreadsheetId,
        sheetName: cleanConfig.sheetName,
        formUrl: cleanConfig.formUrl,
      });

      // 設定の検証とサニタイズ
      const validatedConfig = this.validateAndSanitizeConfig(cleanConfig);
      if (!validatedConfig) {
        console.error('ConfigManager.saveConfig: 設定検証失敗', { userId, config });
        return false;
      }

      // 🔍 デバッグ: validatedConfig詳細ログ
      console.log('🔍 ConfigManager.saveConfig: validatedConfig詳細', {
        userId,
        validatedConfigKeys: Object.keys(validatedConfig),
        spreadsheetId: validatedConfig.spreadsheetId,
        sheetName: validatedConfig.sheetName,
        formUrl: validatedConfig.formUrl,
      });

      // タイムスタンプ更新
      validatedConfig.lastModified = new Date().toISOString();

      // 🔧 修正: DB.updateUserInDatabaseを直接使用（updateUserではなく）
      // updateUserは個別フィールドマージ用、完全なconfigJson置き換えはupdateUserInDatabase
      let success = false;
      try {
        DB.updateUserInDatabase(userId, {
          configJson: JSON.stringify(validatedConfig),
          lastModified: validatedConfig.lastModified,
        });
        success = true;
      } catch (dbError) {
        console.error('❌ ConfigManager.saveConfig: DB更新エラー:', dbError.message);
        success = false;
      }

      if (success) {
        console.log('✅ ConfigManager.saveConfig: 設定保存完了', {
          userId,
          configSize: JSON.stringify(validatedConfig).length,
          configFields: Object.keys(validatedConfig),
          timestamp: validatedConfig.lastModified,
        });
      } else {
        console.error('❌ ConfigManager.saveConfig: データベース更新失敗', { userId });
      }

      return success;
    } catch (error) {
      console.error('❌ ConfigManager.saveConfig: 保存エラー', {
        userId,
        error: error.message,
        stack: error.stack,
      });
      return false;
    }
  },

  // ========================================
  // ✏️ 編集系メソッド
  // ========================================

  /**
   * データソース更新
   * @param {string} userId - ユーザーID
   * @param {Object} dataSource - {spreadsheetId, sheetName}
   * @returns {boolean} 更新成功可否
   */
  updateDataSource(userId, { spreadsheetId, sheetName }) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    const updatedConfig = {
      ...currentConfig,
      spreadsheetId: spreadsheetId || currentConfig.spreadsheetId,
      sheetName: sheetName || currentConfig.sheetName,
      spreadsheetUrl: spreadsheetId
        ? `https://docs.google.com/spreadsheets/d/${spreadsheetId}`
        : currentConfig.spreadsheetUrl,
      setupStatus: spreadsheetId && sheetName ? 'data_connected' : currentConfig.setupStatus,
      lastModified: new Date().toISOString(),
    };

    return this.saveConfig(userId, updatedConfig);
  },

  /**
   * 表示設定更新
   * @param {string} userId - ユーザーID
   * @param {Object} displaySettings - {showNames, showReactions, displayMode}
   * @returns {boolean} 更新成功可否
   */
  updateDisplaySettings(userId, { showNames, showReactions, displayMode }) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    const updatedConfig = {
      ...currentConfig,
      displaySettings: {
        ...(currentConfig.displaySettings || {}),
        ...(showNames !== undefined && { showNames }),
        ...(showReactions !== undefined && { showReactions }),
      },
      ...(displayMode && { displayMode }),
      lastModified: new Date().toISOString(),
    };

    return this.saveConfig(userId, updatedConfig);
  },

  /**
   * アプリステータス更新（データソース情報保護対応）
   * @param {string} userId - ユーザーID
   * @param {Object} status - {appPublished, setupStatus, formUrl, formTitle, preserveDataSource, spreadsheetId, sheetName, appUrl}
   * @returns {boolean} 更新成功可否
   */
  updateAppStatus(
    userId,
    {
      appPublished,
      setupStatus,
      formUrl,
      formTitle,
      preserveDataSource = true,
      spreadsheetId,
      sheetName,
      appUrl,
    }
  ) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    const updatedConfig = {
      ...currentConfig,
      ...(appPublished !== undefined && { appPublished }),
      ...(setupStatus && { setupStatus }),
      ...(formUrl !== undefined && { formUrl }),
      ...(formTitle !== undefined && { formTitle }),
      ...(appPublished && { publishedAt: new Date().toISOString() }),
      // 🔒 明示的なデータソース情報の設定（フォールバック対応）
      ...(spreadsheetId && { spreadsheetId }),
      ...(sheetName && { sheetName }),
      ...(appUrl && { appUrl }),
      lastModified: new Date().toISOString(),
    };

    // 🔒 データソース情報の明示的な保護（connectDataSource設定の保持）
    if (preserveDataSource && currentConfig) {
      const dataSourceFields = [
        'spreadsheetId',
        'sheetName',
        'spreadsheetUrl',
        'columnMapping',
        'opinionHeader',
      ];
      dataSourceFields.forEach((field) => {
        if (currentConfig[field] !== undefined && updatedConfig[field] === undefined) {
          updatedConfig[field] = currentConfig[field];
          console.log(`🔒 ConfigManager.updateAppStatus: ${field}を保護`, {
            userId,
            field,
            value: currentConfig[field],
          });
        }
      });
    }

    return this.saveConfig(userId, updatedConfig);
  },

  /**
   * 設定の部分更新（汎用）
   * @param {string} userId - ユーザーID
   * @param {Object} updates - 更新データ
   * @returns {boolean} 更新成功可否
   */
  updateConfig(userId, updates) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    const updatedConfig = {
      ...currentConfig,
      ...updates,
      lastModified: new Date().toISOString(),
    };

    return this.saveConfig(userId, updatedConfig);
  },

  // ========================================
  // 🏗️ 構築系メソッド
  // ========================================

  /**
   * 初期設定構築（新規ユーザー用）
   * @param {Object} userData - ユーザーデータ
   * @returns {Object} 初期設定オブジェクト
   */
  buildInitialConfig(userData = {}) {
    const now = new Date().toISOString();

    return {
      // 監査情報
      createdAt: now,
      lastModified: now,
      lastAccessedAt: now,

      // セットアップ情報
      setupStatus: 'pending',
      appPublished: false,

      // 表示設定（プライバシー重視：デフォルトOFF）
      displaySettings: {
        showNames: false,
        showReactions: false,
      },
      displayMode: 'anonymous',

      // データソース情報（空で開始）
      spreadsheetId: userData.spreadsheetId || null,
      sheetName: userData.sheetName || null,
      spreadsheetUrl: null,

      // フォーム情報
      formUrl: null,
      formTitle: null,

      // メタ情報
      configVersion: '2.0',
      claudeMdCompliant: true,
    };
  },

  /**
   * ドラフト設定構築
   * @param {Object} currentConfig - 現在の設定
   * @param {Object} updates - 更新データ
   * @returns {Object} ドラフト設定オブジェクト
   */
  buildDraftConfig(currentConfig, updates) {
    const baseConfig = currentConfig || this.buildInitialConfig();

    return {
      // 既存の重要データを継承
      createdAt: baseConfig.createdAt || new Date().toISOString(),
      lastAccessedAt: baseConfig.lastAccessedAt || new Date().toISOString(),

      // データソース情報（更新または継承）
      spreadsheetId: updates.spreadsheetId || baseConfig.spreadsheetId,
      sheetName: updates.sheetName || baseConfig.sheetName,
      spreadsheetUrl: updates.spreadsheetId
        ? `https://docs.google.com/spreadsheets/d/${updates.spreadsheetId}`
        : baseConfig.spreadsheetUrl,

      // 表示設定
      displaySettings: {
        showNames:
          updates.showNames !== undefined
            ? updates.showNames
            : baseConfig.displaySettings?.showNames || false,
        showReactions:
          updates.showReactions !== undefined
            ? updates.showReactions
            : baseConfig.displaySettings?.showReactions || false,
      },

      // アプリ設定
      setupStatus: baseConfig.setupStatus || 'pending',
      appPublished: baseConfig.appPublished || false,

      // フォーム情報（継承）
      ...(baseConfig.formUrl && {
        formUrl: baseConfig.formUrl,
        formTitle: baseConfig.formTitle,
      }),

      // 列マッピング情報（継承）
      ...(baseConfig.columnMapping && { columnMapping: baseConfig.columnMapping }),
      ...(baseConfig.opinionHeader && { opinionHeader: baseConfig.opinionHeader }),

      // ドラフト状態
      isDraft: !baseConfig.appPublished,

      // メタ情報
      configVersion: '2.0',
      claudeMdCompliant: true,
      lastModified: new Date().toISOString(),
    };
  },

  /**
   * 公開設定構築
   * @param {Object} currentConfig - 現在の設定
   * @returns {Object} 公開設定オブジェクト
   */
  buildPublishConfig(currentConfig) {
    if (!currentConfig || !currentConfig.spreadsheetId || !currentConfig.sheetName) {
      throw new Error('公開に必要な設定が不足しています');
    }

    return {
      ...currentConfig,
      appPublished: true,
      setupStatus: 'completed',
      publishedAt: new Date().toISOString(),
      isDraft: false,
      lastModified: new Date().toISOString(),
    };
  },

  // ========================================
  // 🔍 検証・サニタイズ系メソッド
  // ========================================

  /**
   * 設定検証とサニタイズ
   * @param {Object} config - 検証する設定
   * @returns {Object|null} サニタイズ済み設定またはnull
   */
  validateAndSanitizeConfig(config) {
    if (!config || typeof config !== 'object') {
      console.warn('ConfigManager.validateAndSanitizeConfig: 無効な設定オブジェクト');
      return null;
    }

    try {
      // 基本検証
      const sanitized = { ...config };

      // 🚨 重複ネスト防止: configJsonフィールドの除去
      if ('configJson' in sanitized) {
        console.warn(
          'ConfigManager.validateAndSanitizeConfig: configJsonフィールドを除去（重複ネスト防止）'
        );
        delete sanitized.configJson;
      }

      // 必須タイムスタンプの確保
      if (!sanitized.createdAt) {
        sanitized.createdAt = new Date().toISOString();
      }
      if (!sanitized.lastModified) {
        sanitized.lastModified = new Date().toISOString();
      }

      // displaySettingsの正規化
      if (sanitized.displaySettings && typeof sanitized.displaySettings === 'object') {
        sanitized.displaySettings = {
          showNames: Boolean(sanitized.displaySettings.showNames),
          showReactions: Boolean(sanitized.displaySettings.showReactions),
        };
      }

      // spreadsheetUrlの動的生成
      if (sanitized.spreadsheetId && !sanitized.spreadsheetUrl) {
        sanitized.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${sanitized.spreadsheetId}`;
      }

      return sanitized;
    } catch (error) {
      console.error('ConfigManager.validateAndSanitizeConfig: サニタイズエラー', error);
      return null;
    }
  },

  /**
   * ユーザーID検証
   * @param {string} userId - ユーザーID
   * @returns {boolean} 有効性
   */
  validateUserId(userId) {
    return userId && typeof userId === 'string' && userId.trim().length > 0;
  },

  // ========================================
  // 🔧 内部ユーティリティメソッド
  // ========================================

  /**
   * 動的URL生成して設定を拡張
   * @param {Object} config - 基本設定
   * @param {string} userId - ユーザーID
   * @returns {Object} URL追加済み設定
   */
  enhanceConfigWithDynamicUrls(config, userId) {
    try {
      const enhanced = { ...config };

      // WebAppURL生成（main.gsの安定したgetWebAppUrl関数を使用）
      if (!enhanced.appUrl) {
        try {
          // main.gsのgetWebAppUrl関数を使用（複数手法でURL取得）
          const baseUrl = getWebAppUrl();
          if (baseUrl) {
            enhanced.appUrl = `${baseUrl}?mode=view&userId=${userId}`;
          } else {
            console.warn(
              'ConfigManager.enhanceConfigWithDynamicUrls: getWebAppUrl()がnullを返しました'
            );
          }
        } catch (urlError) {
          console.warn(
            'ConfigManager.enhanceConfigWithDynamicUrls: getWebAppUrl()使用エラー:',
            urlError.message
          );
          // URL生成失敗時はappUrlを設定しない（undefined のまま）
        }
      }

      // SpreadsheetURL生成
      if (enhanced.spreadsheetId && !enhanced.spreadsheetUrl) {
        enhanced.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${enhanced.spreadsheetId}`;
      }

      return enhanced;
    } catch (error) {
      console.warn('ConfigManager.enhanceConfigWithDynamicUrls: URL生成エラー', error.message);
      return config; // エラー時は元のconfigをそのまま返す
    }
  },

  // ========================================
  // 🗄️ 統一DB操作メソッド（キャッシュ付き）
  // ========================================

  /**
   * 統一ユーザー情報取得（キャッシュ付き）
   * @param {string} email - ユーザーメールアドレス
   * @returns {Object|null} ユーザー情報
   */
  getUserInfo(email) {
    if (!email) return null;

    try {
      // 5分キャッシュでユーザー情報取得
      const cacheKey = `userInfo-${email}`;

      // CacheManagerが利用可能な場合
      if (typeof cacheManager !== 'undefined' && cacheManager) {
        return cacheManager.get(
          cacheKey,
          () => {
            console.log(`ConfigManager: ユーザー情報取得 - ${email}`);
            return DB.findUserByEmail(email);
          },
          { ttl: 300 }
        );
      }

      // フォールバック: 直接取得
      console.log(`ConfigManager: ユーザー情報取得（キャッシュなし） - ${email}`);
      return DB.findUserByEmail(email);
    } catch (error) {
      console.error('ConfigManager.getUserInfo エラー:', error.message);
      return null;
    }
  },

  /**
   * 統一ユーザー情報+設定取得
   * @param {string} email - ユーザーメールアドレス
   * @returns {Object|null} ユーザー情報+設定
   */
  getUserWithConfig(email) {
    try {
      const userInfo = this.getUserInfo(email);
      if (!userInfo) {
        console.log(`ConfigManager: ユーザーが見つかりません - ${email}`);
        return null;
      }

      const config = this.getUserConfig(userInfo.userId);

      return {
        ...userInfo,
        config: config || this.buildInitialConfig(),
      };
    } catch (error) {
      console.error('ConfigManager.getUserWithConfig エラー:', error.message);
      return null;
    }
  },
});

// ========================================
// 🌐 グローバル関数は削除済み（ConfigManager名前空間に統一）
// ========================================
