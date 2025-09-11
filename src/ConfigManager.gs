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
          userId,
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

            console.info('✅ configJson修復完了', {
              userId,
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

      // 設定取得完了（詳細ログは削除 - 必要時のみエラー情報をログ出力）

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
      // 🚨 第1層防御: 二重構造の厳格検出・警告
      const duplicateFields = Object.keys(config).filter(key => 
        key.toLowerCase() === 'configjson' || 
        (typeof config[key] === 'string' && this.isJSONString(config[key]) && key.toLowerCase().includes('config'))
      );
      
      if (duplicateFields.length > 0) {
        console.error('🚨 ConfigManager.saveConfig: 二重構造検出 - 保存を拒否', {
          userId: userId,
          duplicateFields: duplicateFields,
          source: new Error().stack.split('\n')[2]
        });
        // 厳格モード: 二重構造を検出したら保存を拒否
        throw new Error(`二重構造検出: フィールド [${duplicateFields.join(', ')}] - 保存を拒否しました`);
      }

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

      // cleanConfig準備完了（詳細ログは削除）

      // 設定の検証とサニタイズ
      const validatedConfig = this.validateAndSanitizeConfig(cleanConfig);
      if (!validatedConfig) {
        console.error('ConfigManager.saveConfig: 設定検証失敗', { userId, config });
        return false;
      }

      // 設定検証完了（詳細ログは削除）

      // タイムスタンプ更新
      validatedConfig.lastModified = new Date().toISOString();

      // 🔥 修正: DB.updateUserを完全置換モードで使用
      // configJsonの完全置換により、古いデータの残存を防止
      let success = false;
      try {
        DB.updateUser(userId, validatedConfig, { replaceConfig: true });
        success = true;
        console.log('✅ ConfigManager.saveConfig: 完全置換モードで更新成功', {
          userId,
          configFields: Object.keys(validatedConfig)
        });
      } catch (dbError) {
        console.error('❌ ConfigManager.saveConfig: DB更新エラー:', dbError.message);
        success = false;
      }

      if (!success) {
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
      configVersion: '3.0',
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
      configVersion: '3.0',
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

      // version/etag/sourceKey補完
      if (!sanitized.configVersion) {
        sanitized.configVersion = '3.0';
      }
      if (!sanitized.sourceKey && sanitized.spreadsheetId && sanitized.sheetName) {
        sanitized.sourceKey = `${sanitized.spreadsheetId}::${sanitized.sheetName}`;
      }
      if (!sanitized.etag) {
        try {
          sanitized.etag = `${Utilities.getUuid()}-${new Date().getTime()}`;
        } catch (e) {
          sanitized.etag = new Date().toISOString();
        }
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
   * JSON文字列判定（安全版）
   * @param {*} value - 判定する値
   * @returns {boolean} JSON文字列かどうか
   */
  isJSONString(value) {
    if (typeof value !== 'string') return false;
    if (value.length < 2) return false; // 最小 "{}" or "[]"
    
    // 明らかにJSONでないパターンを除外
    const trimmed = value.trim();
    if (!((trimmed.startsWith('{') && trimmed.endsWith('}')) || 
          (trimmed.startsWith('[') && trimmed.endsWith(']')))) {
      return false;
    }
    
    try {
      JSON.parse(value);
      return true;
    } catch (e) {
      return false;
    }
  },

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

  /**
   * setupStatus/appPublished整合性修正
   * @param {string} userId - ユーザーID
   * @returns {boolean} 修正成功可否
   */
  fixSetupConsistency(userId) {
    try {
      const config = this.getUserConfig(userId);
      if (!config) {
        console.error('ConfigManager.fixSetupConsistency: 設定が見つかりません', { userId });
        return false;
      }

      console.log('🔧 ConfigManager.fixSetupConsistency: 整合性チェック開始', {
        userId,
        currentSetupStatus: config.setupStatus,
        currentAppPublished: config.appPublished,
        hasSpreadsheetId: !!config.spreadsheetId,
        hasSheetName: !!config.sheetName
      });

      let needsFix = false;
      const updates = {};

      // Rule 1: spreadsheetIdとsheetNameが揃っていればdata_connected以上
      if (config.spreadsheetId && config.sheetName) {
        if (config.setupStatus === 'pending' || !config.setupStatus) {
          updates.setupStatus = 'data_connected';
          needsFix = true;
        }
      }

      // Rule 2: appPublishedがtrueならsetupStatusはcompleted
      if (config.appPublished === true) {
        if (config.setupStatus !== 'completed') {
          updates.setupStatus = 'completed';
          needsFix = true;
        }
      }

      // Rule 3: setupStatusがcompletedならappPublishedもtrue
      if (config.setupStatus === 'completed') {
        if (config.appPublished !== true) {
          updates.appPublished = true;
          if (!config.publishedAt) {
            updates.publishedAt = new Date().toISOString();
          }
          needsFix = true;
        }
      }

      // Rule 4: publishedAtがあってappPublishedがfalseなら整合性エラー
      if (config.publishedAt && config.appPublished !== true) {
        updates.appPublished = true;
        updates.setupStatus = 'completed';
        needsFix = true;
      }

      if (needsFix) {
        console.log('🔧 ConfigManager.fixSetupConsistency: 整合性修正適用', {
          userId,
          updates,
          before: {
            setupStatus: config.setupStatus,
            appPublished: config.appPublished
          }
        });

        const success = this.updateConfig(userId, updates);
        if (success) {
          console.log('⚡ セットアップ状態修正完了', {
            userId,
            after: {
              setupStatus: updates.setupStatus || config.setupStatus,
              appPublished: updates.appPublished !== undefined ? updates.appPublished : config.appPublished
            }
          });
        }
        return success;
      } else {
        return true;
      }
    } catch (error) {
      console.error('❌ ConfigManager.fixSetupConsistency: エラー', {
        userId,
        error: error.message
      });
      return false;
    }
  },

  /**
   * フォーム情報復元（Google Sheetsから自動検出）
   * @param {string} userId - ユーザーID
   * @returns {boolean} 復元成功可否
   */
  restoreFormInfo(userId) {
    try {
      const config = this.getUserConfig(userId);
      if (!config || !config.spreadsheetId) {
        console.error('ConfigManager.restoreFormInfo: スプレッドシートIDが見つかりません', { userId });
        return false;
      }

      console.log('🔧 ConfigManager.restoreFormInfo: フォーム情報復元開始', {
        userId,
        spreadsheetId: config.spreadsheetId,
        currentFormUrl: config.formUrl,
        currentFormTitle: config.formTitle
      });

      // フォーム情報が既に存在する場合はスキップ
      if (config.formUrl && config.formTitle) {
        return true;
      }

      try {
        // スプレッドシートからフォーム情報を取得
        const spreadsheet = new ConfigurationManager().getSpreadsheet(config.spreadsheetId);
        const formUrl = spreadsheet.getFormUrl();
        
        if (formUrl) {
          const updates = {
            formUrl: formUrl
          };

          // フォームタイトルも取得を試みる
          try {
            const form = FormApp.openByUrl(formUrl);
            if (form) {
              updates.formTitle = form.getTitle();
            }
          } catch (formError) {
            console.warn('ConfigManager.restoreFormInfo: フォームタイトル取得失敗', formError.message);
            // タイトル取得失敗時はスプレッドシート名をフォールバック
            updates.formTitle = spreadsheet.getName() + ' (フォーム)';
          }

          const success = this.updateConfig(userId, updates);
          if (success) {
            console.log('✅ フォーム情報復元完了', {
              userId,
              formUrl: updates.formUrl,
              formTitle: updates.formTitle
            });
          }
          return success;
        } else {
          console.warn('ConfigManager.restoreFormInfo: スプレッドシートにフォームが関連付けられていません', {
            userId,
            spreadsheetId: config.spreadsheetId
          });
          return false;
        }
      } catch (spreadsheetError) {
        console.error('ConfigManager.restoreFormInfo: スプレッドシートアクセスエラー', {
          userId,
          spreadsheetId: config.spreadsheetId,
          error: spreadsheetError.message
        });
        return false;
      }
    } catch (error) {
      console.error('❌ ConfigManager.restoreFormInfo: エラー', {
        userId,
        error: error.message
      });
      return false;
    }
  },

  /**
   * 既存ユーザーデータの一括修正（二重構造解消）
   * @returns {Object} 修正結果
   */
  fixAllDoubleStructure() {
    try {
      const results = {
        totalUsers: 0,
        fixedUsers: 0,
        errorUsers: 0,
        skippedUsers: 0,
        details: []
      };

      // 全ユーザーを取得
      const allUsers = DB.getAllUsers();
      results.totalUsers = allUsers.length;


      allUsers.forEach(user => {
        try {
          if (!user.configJson) {
            results.skippedUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'skipped',
              reason: 'configJsonが空'
            });
            return;
          }

          // configJsonをパース
          let config;
          try {
            config = JSON.parse(user.configJson);
          } catch (parseError) {
            results.errorUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'error',
              reason: `JSON解析エラー: ${parseError.message}`
            });
            return;
          }

          // 二重構造を検出
          const hasDoubleStructure = config.configJson || config.configJSON;
          if (!hasDoubleStructure) {
            results.skippedUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'skipped',
              reason: '二重構造なし'
            });
            return;
          }

          // 修正処理：getUserConfigを呼び出すことで自動修復させる
          const fixedConfig = this.getUserConfig(user.userId);
          if (fixedConfig) {
            results.fixedUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'fixed',
              reason: '二重構造を自動修復'
            });
          } else {
            results.errorUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'error',
              reason: '自動修復失敗'
            });
          }
        } catch (userError) {
          results.errorUsers++;
          results.details.push({
            userId: user.userId || 'unknown',
            email: user.userEmail || 'unknown',
            status: 'error',
            reason: `処理エラー: ${userError.message}`
          });
        }
      });

      console.log('⚡ 全ユーザーの構造修正完了', {
        total: results.totalUsers,
        fixed: results.fixedUsers,
        error: results.errorUsers,
        skipped: results.skippedUsers
      });

      return results;
    } catch (error) {
      console.error('❌ ConfigManager.fixAllDoubleStructure: エラー', error.message);
      throw error;
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
            return DB.findUserByEmail(email);
          },
          { ttl: 300 }
        );
      }

      // フォールバック: 直接取得
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

  // ========================================
  // 🛡️ 予防システム・監視機能
  // ========================================

  /**
   * 二重構造予防システム初期化
   * @returns {boolean} 初期化成功可否
   */
  initPreventionSystem() {
    try {
      console.log('🛡️ ConfigManager.initPreventionSystem: 二重構造予防システム初期化');
      
      // 予防システムのフラグを立てる
      this._preventionSystemActive = true;
      
      console.log('✅ ConfigManager.initPreventionSystem: 予防システム初期化完了');
      return true;
    } catch (error) {
      console.error('❌ ConfigManager.initPreventionSystem: エラー', error.message);
      return false;
    }
  },

  /**
   * システム健全性チェック
   * @returns {Object} チェック結果
   */
  performHealthCheck() {
    try {
      
      const results = {
        totalUsers: 0,
        healthyUsers: 0,
        doubleStructureUsers: 0,
        errorUsers: 0,
        details: [],
        timestamp: new Date().toISOString()
      };

      // 全ユーザーをチェック
      const allUsers = DB.getAllUsers();
      results.totalUsers = allUsers.length;

      allUsers.forEach(user => {
        try {
          if (!user.configJson) {
            results.healthyUsers++;
            return;
          }

          // configJsonを安全に解析
          let config;
          try {
            config = JSON.parse(user.configJson);
          } catch (parseError) {
            results.errorUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              issue: 'json_parse_error',
              description: parseError.message
            });
            return;
          }

          // 二重構造チェック
          const hasDoubleStructure = config.configJson || config.configJSON;
          if (hasDoubleStructure) {
            results.doubleStructureUsers++;
            results.details.push({
              userId: user.userId,
              email: user.userEmail,
              issue: 'double_structure',
              description: '二重構造を検出'
            });
          } else {
            results.healthyUsers++;
          }
        } catch (userError) {
          results.errorUsers++;
          results.details.push({
            userId: user.userId || 'unknown',
            email: user.userEmail || 'unknown',
            issue: 'processing_error',
            description: userError.message
          });
        }
      });

      const healthScore = Math.round((results.healthyUsers / results.totalUsers) * 100);
      
      console.log('📊 データベース健康状態診断完了', {
        total: results.totalUsers,
        healthy: results.healthyUsers,
        doubleStructure: results.doubleStructureUsers,
        errors: results.errorUsers,
        healthScore: `${healthScore}%`
      });

      results.healthScore = healthScore;
      return results;
    } catch (error) {
      console.error('❌ ConfigManager.performHealthCheck: エラー', error.message);
      throw error;
    }
  },

  /**
   * 総合修復処理（全問題を一括解決）
   * @returns {Object} 修復結果
   */
  performCompleteRepair() {
    try {
      console.log('🔧 ConfigManager.performCompleteRepair: 総合修復処理開始');
      
      const repairResults = {
        doubleStructureRepair: null,
        consistencyRepair: null,
        formInfoRepair: null,
        timestamp: new Date().toISOString()
      };

      // Phase 1: 二重構造修復
      console.log('Phase 1: 二重構造一括修復');
      repairResults.doubleStructureRepair = this.fixAllDoubleStructure();

      // Phase 2: 整合性修復（各ユーザーごと）
      console.log('Phase 2: 整合性修復');
      const allUsers = DB.getAllUsers();
      let consistencyFixed = 0;
      let consistencyErrors = 0;

      allUsers.forEach(user => {
        try {
          const success = this.fixSetupConsistency(user.userId);
          if (success) {
            consistencyFixed++;
          } else {
            consistencyErrors++;
          }
        } catch (error) {
          consistencyErrors++;
          console.warn('ConfigManager.performCompleteRepair: 整合性修復エラー', {
            userId: user.userId,
            error: error.message
          });
        }
      });

      repairResults.consistencyRepair = {
        totalUsers: allUsers.length,
        fixed: consistencyFixed,
        errors: consistencyErrors
      };

      // Phase 3: フォーム情報復元（各ユーザーごと）
      console.log('Phase 3: フォーム情報復元');
      let formInfoFixed = 0;
      let formInfoErrors = 0;

      allUsers.forEach(user => {
        try {
          const success = this.restoreFormInfo(user.userId);
          if (success) {
            formInfoFixed++;
          } else {
            formInfoErrors++;
          }
        } catch (error) {
          formInfoErrors++;
          console.warn('ConfigManager.performCompleteRepair: フォーム情報復元エラー', {
            userId: user.userId,
            error: error.message
          });
        }
      });

      repairResults.formInfoRepair = {
        totalUsers: allUsers.length,
        fixed: formInfoFixed,
        errors: formInfoErrors
      };

      console.log('✅ ConfigManager.performCompleteRepair: 総合修復処理完了', repairResults);
      return repairResults;
    } catch (error) {
      console.error('❌ ConfigManager.performCompleteRepair: エラー', error.message);
      throw error;
    }
  }
});

// ========================================
// 🌐 グローバル関数は削除済み（ConfigManager名前空間に統一）
// ========================================
