/**
 * @fileoverview ConfigService - 統一設定管理サービス
 * 
 * 🎯 責任範囲:
 * - configJSON の CRUD操作
 * - 設定検証・サニタイズ
 * - 動的設定生成（URL等）
 * - 設定マイグレーション
 * 
 * 🔄 置き換え対象:
 * - ConfigManager (ConfigManager.gs)
 * - UnifiedManager.config
 * - ConfigurationManager (削除済み)
 */

/* global UserService, DB, PROPS_KEYS, CONSTANTS, SecurityValidator, AppCacheService, SecurityService */

/**
 * ConfigService - 統一設定管理サービス
 * configJSON中心型アーキテクチャの核となるサービス
 */
const ConfigService = Object.freeze({

  // ===========================================
  // 📖 設定読み込み・取得
  // ===========================================

  /**
   * ユーザー設定取得（統合版）
   * @param {string} userId - ユーザーID
   * @returns {Object|null} 統合設定オブジェクト
   */
  getUserConfig(userId) {
    if (!userId || !this.validateUserId(userId)) {
      console.warn('ConfigService.getUserConfig: 無効なuserID - デフォルト設定を返却:', userId);
      return this.getDefaultConfig(userId);
    }

    const cacheKey = `config_${userId}`;

    try {
      // キャッシュから取得試行
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }

      // データベースから取得
      const user = DB.findUserById(userId);
      if (!user) {
        console.warn('ConfigService.getUserConfig: ユーザーが見つかりません:', userId);
        return this.getDefaultConfig(userId);
      }

      // configJsonパース・修復
      const baseConfig = this.parseAndRepairConfig(user.configJson, userId);

      // 動的URL生成
      const enhancedConfig = this.enhanceConfigWithDynamicUrls(baseConfig, userId);

      // キャッシュに保存（10分間）
      CacheService.getScriptCache().put(
        cacheKey, 
        JSON.stringify(enhancedConfig), 
        600
      );

      return enhancedConfig;
    } catch (error) {
      console.error('ConfigService.getUserConfig: エラー', {
        userId,
        error: error.message
      });
      return this.getDefaultConfig(userId);
    }
  },

  

  /**
   * デフォルト設定取得
   * @param {string} userId - ユーザーID
   * @returns {Object} デフォルト設定
   */
  getDefaultConfig(userId) {
    return {
      userId,
      setupStatus: 'pending',
      appPublished: false,
      displaySettings: {
        showNames: false,
        showReactions: false
      },
      userPermissions: {
        isOwner: false,
        isSystemAdmin: false,
        accessLevel: 'viewer',
        canEdit: false,
        canView: true,
        canReact: true
      },
      createdAt: new Date().toISOString(),
      lastModified: new Date().toISOString()
    };
  },

  /**
   * configJson解析・自動修復
   * @param {string} configJson - 生のconfigJson文字列
   * @param {string} userId - ユーザーID
   * @returns {Object} 解析・修復済み設定オブジェクト
   */
  parseAndRepairConfig(configJson, userId) {
    try {
      let baseConfig;
      
      // 基本JSON解析
      try {
        baseConfig = JSON.parse(configJson || '{}');
      } catch (parseError) {
        console.warn('ConfigService.parseAndRepairConfig: JSON解析エラー:', parseError.message);
        return this.getDefaultConfig(userId);
      }

      // 🔧 二重構造修復（レガシー構造の互換性処理）
      if (baseConfig.configJson || baseConfig.configJSON) {
        console.warn('⚠️ ConfigService: 二重構造を検出 - 自動修復開始');
        baseConfig = this.repairNestedConfig(baseConfig, userId);
      }

      // 🔧 必須フィールド補完
      baseConfig = this.ensureRequiredFields(baseConfig, userId);

      return baseConfig;
    } catch (error) {
      console.error('ConfigService.parseAndRepairConfig: エラー', error.message);
      return this.getDefaultConfig(userId);
    }
  },

  /**
   * ネストした設定構造を修復
   * @param {Object} config - 設定オブジェクト
   * @param {string} userId - ユーザーID
   * @returns {Object} 修復済み設定オブジェクト
   */
  repairNestedConfig(config, userId) {
    try {
      let nestedConfig = config.configJson || config.configJSON;

      if (typeof nestedConfig === 'string') {
        nestedConfig = JSON.parse(nestedConfig);
      }

      // 内側のデータを外側にマージ（内側優先）
      const mergedConfig = { ...config, ...nestedConfig };

      // configJsonフィールドを削除
      delete mergedConfig.configJson;
      delete mergedConfig.configJSON;

      // 修復したデータをDBに保存
      this.saveConfig(userId, mergedConfig);

      console.info('✅ ConfigService: configJson修復完了', {
        userId,
        fixedFields: Object.keys(mergedConfig)
      });

      return mergedConfig;
    } catch (error) {
      console.error('ConfigService.repairNestedConfig: エラー', error.message);
      // フォールバック: 外側のデータのみ使用
      const { configJson, configJSON, ...cleanConfig } = config;
      return cleanConfig;
    }
  },

  /**
   * 必須フィールド補完
   * @param {Object} config - 設定オブジェクト
   * @param {string} userId - ユーザーID
   * @returns {Object} 補完済み設定オブジェクト
   */
  ensureRequiredFields(config, userId) {
    const timestamp = new Date().toISOString();

    return {
      // デフォルト値
      setupStatus: 'pending',
      appPublished: false,
      displaySettings: {
        showNames: false,
        showReactions: false
      },
      createdAt: timestamp,
      // 既存設定を上書き
      ...config,
      // 強制更新フィールド
      userId,
      lastModified: timestamp
    };
  },

  /**
   * 動的URL生成・設定拡張
   * @param {Object} baseConfig - 基本設定
   * @param {string} userId - ユーザーID
   * @returns {Object} 拡張済み設定
   */
  enhanceConfigWithDynamicUrls(baseConfig, userId) {
    try {
      const enhanced = { ...baseConfig };

      // スプレッドシートURL生成
      if (baseConfig.spreadsheetId && !baseConfig.spreadsheetUrl) {
        enhanced.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${baseConfig.spreadsheetId}/edit`;
      }

      // アプリURL生成（公開済みの場合）
      if (baseConfig.appPublished && !baseConfig.appUrl) {
        try {
          enhanced.appUrl = ScriptApp.getService().getUrl();
        } catch (urlError) {
          console.warn('ConfigService: アプリURL取得エラー', urlError.message);
        }
      }

      // フォーム存在確認
      if (baseConfig.formUrl) {
        enhanced.hasValidForm = this.validateFormUrl(baseConfig.formUrl);
      }

      // 設定完了度計算
      enhanced.completionScore = this.calculateCompletionScore(enhanced);

      // 権限情報を追加
      enhanced.userPermissions = this.generateUserPermissions(userId);

      return enhanced;
    } catch (error) {
      console.error('ConfigService.enhanceConfigWithDynamicUrls: エラー', error.message);
      return baseConfig;
    }
  },

  /**
   * ユーザー権限情報生成
   * @param {string} userId - ユーザーID
   * @returns {Object} 権限情報
   */
  generateUserPermissions(userId) {
    try {
      // UserServiceを使って権限レベルを取得
      const accessLevel = UserService.getAccessLevel(userId);
      const isOwner = UserService.verifyOwnership(userId);
      const currentEmail = UserService.getCurrentEmail();
      const isSystemAdmin = UserService.isSystemAdmin(currentEmail);

      return {
        isOwner,
        isSystemAdmin,
        accessLevel,
        canEdit: isOwner || isSystemAdmin,
        canView: true, // 基本的に全員閲覧可能
        canReact: true // 基本的に全員リアクション可能
      };
    } catch (error) {
      console.error('ConfigService.generateUserPermissions: エラー', error.message);
      return {
        isOwner: false,
        isSystemAdmin: false,
        accessLevel: 'guest',
        canEdit: false,
        canView: true,
        canReact: true
      };
    }
  },

  // ===========================================
  // 💾 設定保存・更新
  // ===========================================

  /**
   * 設定保存（統合版）
   * @param {string} userId - ユーザーID
   * @param {Object} config - 設定オブジェクト
   * @returns {boolean} 保存成功可否
   */
  saveConfig(userId, config) {
    if (!userId || !config) {
      console.error('ConfigService.saveConfig: 無効なパラメータ');
      return false;
    }

    try {
      // 設定検証・サニタイズ
      const validatedConfig = this.validateAndSanitizeConfig(config, userId);
      if (!validatedConfig.isValid) {
        console.error('ConfigService.saveConfig: 設定検証エラー', validatedConfig.errors);
        return false;
      }

      // タイムスタンプ更新
      const configToSave = {
        ...validatedConfig.config,
        lastModified: new Date().toISOString()
      };

      // データベース保存
      const success = DB.updateUserConfig(userId, configToSave);
      if (!success) {
        console.error('ConfigService.saveConfig: DB保存エラー');
        return false;
      }

      // キャッシュクリア
      this.clearConfigCache(userId);

      console.info('ConfigService.saveConfig: 設定保存完了', { userId });
      return true;
    } catch (error) {
      console.error('ConfigService.saveConfig: エラー', {
        userId,
        error: error.message
      });
      return false;
    }
  },

  /**
   * 部分設定更新
   * @param {string} userId - ユーザーID
   * @param {Object} partialConfig - 部分設定
   * @returns {boolean} 更新成功可否
   */
  updatePartialConfig(userId, partialConfig) {
    try {
      // 現在の設定取得
      const currentConfig = this.getUserConfig(userId);
      if (!currentConfig) {
        console.error('ConfigService.updatePartialConfig: 現在設定が見つかりません');
        return false;
      }

      // 部分更新をマージ
      const updatedConfig = {
        ...currentConfig,
        ...partialConfig,
        lastModified: new Date().toISOString()
      };

      // 保存
      return this.saveConfig(userId, updatedConfig);
    } catch (error) {
      console.error('ConfigService.updatePartialConfig: エラー', {
        userId,
        error: error.message
      });
      return false;
    }
  },

  // ===========================================
  // 🔍 設定検証・サニタイズ
  // ===========================================

  /**
   * 設定検証・サニタイズ
   * @param {Object} config - 設定オブジェクト
   * @param {string} userId - ユーザーID
   * @returns {Object} 検証結果
   */
  validateAndSanitizeConfig(config, userId) {
    const result = {
      isValid: true,
      errors: [],
      config: { ...config }
    };

    try {
      // 必須フィールドチェック
      if (!config.setupStatus) {
        result.config.setupStatus = 'pending';
      }

      if (typeof config.appPublished !== 'boolean') {
        result.config.appPublished = false;
      }

      // スプレッドシートID検証
      if (config.spreadsheetId && !this.validateSpreadsheetId(config.spreadsheetId)) {
        result.errors.push('無効なスプレッドシートID');
        result.isValid = false;
      }

      // フォームURL検証
      if (config.formUrl && !this.validateFormUrl(config.formUrl)) {
        result.errors.push('無効なフォームURL');
        result.isValid = false;
      }

      // displaySettings検証
      if (config.displaySettings) {
        result.config.displaySettings = this.sanitizeDisplaySettings(config.displaySettings);
      }

      // columnMapping検証
      if (config.columnMapping) {
        result.config.columnMapping = this.sanitizeColumnMapping(config.columnMapping);
      }

      // サイズ制限チェック（32KB制限）
      const configSize = JSON.stringify(result.config).length;
      if (configSize > 32000) {
        result.errors.push('設定データが大きすぎます');
        result.isValid = false;
      }

    } catch (error) {
      result.errors.push(`検証エラー: ${error.message}`);
      result.isValid = false;
    }

    return result;
  },

  /**
   * 表示設定サニタイズ
   * @param {Object} displaySettings - 表示設定
   * @returns {Object} サニタイズ済み表示設定
   */
  sanitizeDisplaySettings(displaySettings) {
    return {
      showNames: !!displaySettings.showNames,
      showReactions: !!displaySettings.showReactions,
      displayMode: ['anonymous', 'named', 'email'].includes(displaySettings.displayMode) 
        ? displaySettings.displayMode 
        : 'anonymous'
    };
  },

  /**
   * 列マッピングサニタイズ
   * @param {Object} columnMapping - 列マッピング
   * @returns {Object} サニタイズ済み列マッピング
   */
  sanitizeColumnMapping(columnMapping) {
    const sanitized = { mapping: {} };

    if (columnMapping.mapping) {
      ['answer', 'reason', 'class', 'name'].forEach(key => {
        const value = columnMapping.mapping[key];
        if (typeof value === 'number' && value >= 0) {
          sanitized.mapping[key] = value;
        }
      });
    }

    return sanitized;
  },

  // ===========================================
  // 🔧 ユーティリティ・バリデーション
  // ===========================================

  /**
   * ユーザーID検証
   * @param {string} userId - ユーザーID
   * @returns {boolean} 有効かどうか
   */
  validateUserId(userId) {
    return !!(userId && typeof userId === 'string' && userId.length > 0);
  },

  /**
   * スプレッドシートID検証
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {boolean} 有効かどうか
   */
  validateSpreadsheetId(spreadsheetId) {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') return false;
    return /^[a-zA-Z0-9-_]+$/.test(spreadsheetId) && spreadsheetId.length > 10;
  },

  /**
   * フォームURL検証
   * @param {string} formUrl - フォームURL
   * @returns {boolean} 有効かどうか
   */
  validateFormUrl(formUrl) {
    if (!formUrl || typeof formUrl !== 'string') return false;
    return formUrl.includes('forms.gle') || formUrl.includes('docs.google.com/forms');
  },

  /**
   * Google Formを作成
   * @param {string} userId - ユーザーID
   * @param {Object} config - フォーム設定
   * @returns {Object} 作成結果
   */
  createForm(userId, config) {
    try {
      console.log('ConfigService.createForm: 開始', { userId, configTitle: config?.title });

      if (!userId) {
        return { success: false, error: 'ユーザーIDが必要です' };
      }

      if (!config || !config.title) {
        return { success: false, error: 'フォームタイトルが必要です' };
      }

      // Google Formsを作成
      const form = FormApp.create(config.title);

      if (config.description) {
        form.setDescription(config.description);
      }

      // 基本的な質問を追加（設定可能にする）
      const questionTitle = config.questionText || 'あなたの回答を入力してください';
      form.addTextItem()
        .setTitle(questionTitle)
        .setRequired(true);

      // 名前フィールド（オプション）
      if (config.collectName !== false) {
        form.addTextItem()
          .setTitle('お名前')
          .setRequired(true);
      }

      const formUrl = form.getPublishedUrl();
      const editUrl = form.getEditUrl();

      console.log('ConfigService.createForm: フォーム作成成功', { formUrl });

      // ユーザー設定にフォーム情報を保存
      const userConfig = this.getUserConfig(userId) || {};
      userConfig.formUrl = formUrl;
      userConfig.formEditUrl = editUrl;
      userConfig.formTitle = config.title;
      userConfig.formCreated = true;
      userConfig.formCreatedAt = new Date().toISOString();

      this.saveUserConfig(userId, userConfig);

      return {
        success: true,
        formUrl,
        editUrl,
        title: config.title,
        message: 'フォームが正常に作成されました'
      };

    } catch (error) {
      console.error('ConfigService.createForm エラー:', error.message);
      return {
        success: false,
        error: `フォーム作成に失敗しました: ${error.message}`
      };
    }
  },

  /**
   * セットアップステップ判定（Core.gsより移行）
   * @param {Object} userInfo - ユーザー情報
   * @param {Object} configJson - 設定JSON（オプション）
   * @returns {number} セットアップステップ (1-3)
   */
  determineSetupStep(userInfo, configJson) {
    // 統一データソース：configJSON優先
    const config = JSON.parse(userInfo.configJson || '{}') || configJson || {};
    const setupStatus = config.setupStatus || 'pending';

    // Step 1: データソース未設定 OR セットアップ初期状態
    if (
      !userInfo ||
      !config.spreadsheetId ||
      config.spreadsheetId.trim() === '' ||
      setupStatus === 'pending'
    ) {
      return 1;
    }

    // Step 2: データソース設定済み + セットアップ未完了
    if (setupStatus !== 'completed' || !config.formCreated) {
      console.info('ConfigService.determineSetupStep: Step 2 - データソース設定済み、セットアップ未完了');
      return 2;
    }

    // Step 3: 全セットアップ完了 + 公開済み
    if (setupStatus === 'completed' && config.formCreated && config.appPublished) {
      return 3;
    }

    // フォールバック: Step 2
    return 2;
  },

  /**
   * システムセットアップ状態確認
   * @returns {boolean} システムがセットアップされているか
   */
  isSystemSetup() {
    try {
      // まずコア必須プロパティが揃っていること
      if (!this.hasCoreSystemProps()) return false;

      // コア必須プロパティが揃っていればセットアップ済みとして扱う
      return true;
    } catch (error) {
      console.warn('ConfigService.isSystemSetup: システム状態確認エラー', error.message);
      return false;
    }
  },

  /**
   * 設定完了度計算
   * @param {Object} config - 設定オブジェクト
   * @returns {number} 完了度スコア（0-100）
   */
  calculateCompletionScore(config) {
    let score = 0;
    const weights = {
      spreadsheetId: 30,
      sheetName: 20,
      formUrl: 20,
      columnMapping: 20,
      displaySettings: 10
    };

    if (config.spreadsheetId) score += weights.spreadsheetId;
    if (config.sheetName) score += weights.sheetName;
    if (config.formUrl) score += weights.formUrl;
    if (config.columnMapping && Object.keys(config.columnMapping.mapping || {}).length > 0) {
      score += weights.columnMapping;
    }
    if (config.displaySettings) score += weights.displaySettings;

    return Math.min(score, 100);
  },

  // ===========================================
  // 🧹 キャッシュ・メンテナンス
  // ===========================================

  /**
   * 設定キャッシュクリア
   * @param {string} userId - ユーザーID
   */
  clearConfigCache(userId) {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(`config_${userId}`);
      
      // 関連キャッシュもクリア
      cache.remove(`user_info_${userId}`);
      
      console.info('ConfigService.clearConfigCache: キャッシュクリア完了', { userId });
    } catch (error) {
      console.error('ConfigService.clearConfigCache: エラー', error.message);
    }
  },

  /**
   * 全設定キャッシュクリア
   */
  clearAllConfigCache() {
    try {
      // スクリプトキャッシュ全体をクリア（設定関連のみ）
      const cache = CacheService.getScriptCache();
      cache.removeAll();
      
      console.info('ConfigService.clearAllConfigCache: 全キャッシュクリア完了');
    } catch (error) {
      console.error('ConfigService.clearAllConfigCache: エラー', error.message);
    }
  },

  /**
   * サービス状態診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const results = {
      service: 'ConfigService',
      timestamp: new Date().toISOString(),
      checks: []
    };

    try {
      // データベース接続確認
      results.checks.push({
        name: 'Database Connection',
        status: DB.isHealthy ? DB.isHealthy() ? '✅' : '❌' : '⚠️',
        details: 'DB接続テスト'
      });

      // キャッシュサービス確認
      const cache = CacheService.getScriptCache();
      cache.put('test_config_cache', 'test', 10);
      const testResult = cache.get('test_config_cache');
      results.checks.push({
        name: 'Cache Service',
        status: testResult === 'test' ? '✅' : '❌',
        details: 'キャッシュ読み書きテスト'
      });

      // 設定検証機能確認
      const testConfig = { setupStatus: 'test', appPublished: false };
      const validation = this.validateAndSanitizeConfig(testConfig, 'test-user');
      results.checks.push({
        name: 'Config Validation',
        status: validation.isValid ? '✅' : '❌',
        details: '設定検証機能テスト'
      });

      results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    } catch (error) {
      results.checks.push({
        name: 'Service Diagnosis',
        status: '❌',
        details: error.message
      });
      results.overall = '❌';
    }

    return results;
  },

  /**
   * コアシステム必須プロパティの存在確認
   * - ADMIN_EMAIL / DATABASE_SPREADSHEET_ID / SERVICE_ACCOUNT_CREDS
   * @returns {boolean} 3つすべて存在すれば true
   */
  hasCoreSystemProps() {
    try {
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
      const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      const creds = props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS);

      if (!adminEmail || !dbId || !creds) return false;

      // SERVICE_ACCOUNT_CREDS はJSONであることを軽く検証
      try {
        const parsed = JSON.parse(creds);
        if (!parsed || typeof parsed !== 'object') return false;
      } catch (_e) {
        return false;
      }
      return true;
    } catch (error) {
      console.warn('ConfigService.hasCoreSystemProps: エラー', error.message);
      return false;
    }
  },

  // ===========================================
  // 📊 管理パネル用API関数（main.gsから移動）
  // ===========================================

  /**
   * 設定の下書き保存
   * AdminPanel.js.html から呼び出される
   *
   * @param {Object} config - 保存する設定
   * @returns {Object} 保存結果
   */
  saveDraftConfiguration(config) {
    try {
      if (!config || typeof config !== 'object') {
        return {
          success: false,
          message: '設定データが無効です'
        };
      }

      // ユーザー情報の取得
      let userInfo = UserService.getCurrentUserInfo();
      let userId = userInfo && userInfo.userId;

      if (!userId) {
        const email = UserService.getCurrentEmail();
        if (!email) {
          return { success: false, message: 'ユーザー情報が見つかりません' };
        }

        // DB から直接ユーザー検索
        try {
          const dbUser = DB.findUserByEmail(email);
          if (dbUser && dbUser.userId) {
            userId = dbUser.userId;
            userInfo = { userId, userEmail: email };
          } else {
            // ユーザーを自動作成
            const created = UserService.createUser(email);
            if (created && created.userId) {
              userId = created.userId;
              userInfo = created;
            } else {
              console.error('ConfigService.saveDraftConfiguration: ユーザー作成に失敗 - userIdが見つかりません', { created });
              return { success: false, message: 'ユーザー作成に失敗しました' };
            }
          }
        } catch (e) {
          return { success: false, message: 'ユーザー情報の処理に失敗しました' };
        }
      }

      // 設定を保存
      const result = this.saveUserConfig(userId, config);

      return {
        success: result ? true : false,
        config: result || config,
        message: result ? '下書きを保存しました' : '保存に失敗しました'
      };

    } catch (error) {
      console.error('ConfigService.saveDraftConfiguration エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
  },

  /**
   * アプリケーションの公開
   * AdminPanel.js.html から呼び出される
   *
   * @param {Object} publishConfig - 公開設定
   * @returns {Object} 公開結果
   */
  publishApplication(publishConfig) {
    try {
      if (!publishConfig || typeof publishConfig !== 'object') {
        return {
          success: false,
          message: '公開設定が無効です'
        };
      }

      // 必須フィールドの確認
      if (!publishConfig.spreadsheetId || !publishConfig.sheetName) {
        return {
          success: false,
          message: 'スプレッドシートIDとシート名は必須です'
        };
      }

      // ユーザー情報の取得
      let userInfo = UserService.getCurrentUserInfo();
      let userId = userInfo && userInfo.userId;

      if (!userId) {
        const email = UserService.getCurrentEmail();
        if (!email) {
          return { success: false, message: 'ユーザー情報が見つかりません' };
        }

        // DB から直接ユーザー検索
        try {
          const dbUser = DB.findUserByEmail(email);
          if (dbUser && dbUser.userId) {
            userId = dbUser.userId;
            userInfo = { userId, userEmail: email };
          } else {
            return { success: false, message: 'ユーザーが見つかりません' };
          }
        } catch (e) {
          return { success: false, message: 'ユーザー情報の処理に失敗しました' };
        }
      }

      // 現在の設定を取得
      const currentConfig = this.getUserConfig(userId) || {};

      // 公開設定をマージ
      const updatedConfig = {
        ...currentConfig,
        ...publishConfig,
        appPublished: true,
        publishedAt: new Date().toISOString(),
        setupComplete: true
      };

      // 設定を保存
      const result = this.saveUserConfig(userId, updatedConfig);

      if (result) {
        return {
          success: true,
          config: result,
          message: 'アプリケーションを公開しました'
        };
      } else {
        return {
          success: false,
          message: '公開設定の保存に失敗しました'
        };
      }

    } catch (error) {
      console.error('ConfigService.publishApplication エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
  }

});
