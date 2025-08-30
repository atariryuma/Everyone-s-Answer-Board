/**
 * 統一ユーティリティモジュール
 * StudyQuestの共通機能を統合し、重複コードを削減
 * commonUtilities.gsの機能も統合
 */

// =============================================================================
// 統合：共通ユーティリティ関数（元commonUtilities.gs）
// =============================================================================


/**
 * 統一ユーティリティクラス
 * 28個の生成関数と258箇所のログパターンを統合
 */
class UUtilities {
  // 後方互換性のためのurlFactoryエイリアス（generatorFactory.urlに統合済み）
  static get urlFactory() {
    return {
      generateUserUrls: (userId, options = {}) => UUtilities.generatorFactory.url.user(userId, options),
      generateUnpublishedUrl: (userId) => UUtilities.generatorFactory.url.unpublished(userId),
      buildAdminUrl: (userId) => UUtilities.generatorFactory.url.admin(userId)
    };
  }

  // 後方互換性のためのformFactoryエイリアス（generatorFactory.formに統合済み）
  static get formFactory() {
    return {
      create: (type, options) => UUtilities.generatorFactory.form.create(type, options),
      createCustomUI: (requestUserId, config) => UUtilities.generatorFactory.form.customUI(requestUserId, config),
      createQuickStartUI: (requestUserId) => UUtilities.generatorFactory.form.quickStartUI(requestUserId)
    };
  }

  // 後方互換性のためのuserFactoryエイリアス（generatorFactory.userに統合済み）
  static get userFactory() {
    return {
      create: (userData) => UUtilities.generatorFactory.user.create(userData),
      createFolder: (userEmail) => UUtilities.generatorFactory.user.folder(userEmail)
    };
  }

  /**
   * 【Phase 6最適化】統合生成ファクトリ - 28個の生成関数を一元管理
   * すべてのcreate/generate/build関数を統合し、重複を排除
   */
  static get generatorFactory() {
    return {
    /**
     * レスポンス生成統合（既存responseFactoryを統合）
     */
    response: {
      success: (data = null, message = null) => createSuccessResponse(data, message),
      error: (error, message = null, data = null) => createErrorResponse(error, message, data),
      unified: (success, data = null, message = null, error = null) => createUnifiedResponse(success, data, message, error)
    },

    /**
     * HTML生成統合 - HtmlServiceの統一管理
     */
    html: {
      output: (content) => HtmlService.createHtmlOutput(content),
      template: (fileName) => HtmlService.createTemplateFromFile(fileName),
      templateWithData: (fileName, data) => {
        const template = HtmlService.createTemplateFromFile(fileName);
        Object.keys(data).forEach(key => template[key] = data[key]);
        return template;
      },
      secureRedirect: (targetUrl, message) => createSecureRedirect(targetUrl, message)
    },

    /**
     * キー生成統合 - 各種キー生成の統一管理
     */
    key: {
      userScoped: (prefix, userId, suffix) => buildUserScopedKey(prefix, userId, suffix),
      // secure: 未実装関数のため削除
      batchCache: (operation, id, params, options = {}) => {
        if (typeof options.cachePrefix === 'string') {
          return `${options.cachePrefix}:${operation}:${id}:${JSON.stringify(params)}`;
        }
        return `batch:${operation}:${id}:${JSON.stringify(params)}`;
      },
      sheetDataCache: (operation, spreadsheetId, params) => `sheetData:${operation}:${spreadsheetId}:${JSON.stringify(params)}`
    },

    /**
     * URL生成統合（既存urlFactoryを統合）
     */
    url: {
      user: (userId, options = {}) => {
        if (options.cacheBusting || options.forceFresh) {
          return generateUserUrlsWithCacheBusting(userId, options);
        }
        return generateUserUrls(userId);
      },
      unpublished: (userId) => generateUnpublishedStateUrl(userId),
      admin: (userId) => buildUserAdminUrl(userId)
    },

    /**
     * フォーム生成統合（既存formFactoryを統合）
     */
    form: {
      unified: (type, userEmail, userId, overrides) => createUnifiedForm(type, userEmail, userId, overrides),
      factory: (options) => createFormFactory(options),
      customUI: (requestUserId, config) => createCustomFormUI(requestUserId, config),
      quickStartUI: (requestUserId) => createQuickStartFormUI(requestUserId),
      forSpreadsheet: (spreadsheetId, sheetName) => createFormForSpreadsheet(spreadsheetId, sheetName),
      create: (type, options) => {
        switch (type) {
          case 'quickstart':
          case 'custom':
          case 'study':
            return createUnifiedForm(type, options.userEmail, options.userId, options.overrides);
          case 'factory':
            return createFormFactory(options);
          default:
            throw new Error(`未対応のフォームタイプ: ${type}`);
        }
      }
    },

    /**
     * ユーザー生成統合（既存userFactoryを統合）
     */
    user: {
      create: (userData) => DB.createUser(userData),
      folder: (userEmail) => createUserFolder(userEmail)
    },

    /**
     * サービス関連生成統合
     */
    service: {
      sheetsService: (accessToken) => createSheetsService(accessToken),
      serviceAccountToken: () => generateNewServiceAccountToken()
    },

    /**
     * コンテキスト生成統合
     */
    context: {
      execution: (requestUserId, options = {}) => createExecutionContext(requestUserId, options),
      response: (context) => buildResponseFromContext(context)
    },

    /**
     * ボード生成統合
     */
    board: {
      fromAdmin: (requestUserId) => createBoardFromAdmin(requestUserId)
    }
    };
  }

  // 後方互換性のためのresponseFactoryエイリアス（generatorFactory.responseに統合済み）
  static get responseFactory() {
    return {
      success: (data = null, message = null) => UUtilities.generatorFactory.response.success(data, message),
      error: (error, message = null, data = null) => UUtilities.generatorFactory.response.error(error, message, data),
      unified: (success, data = null, message = null, error = null) => UUtilities.generatorFactory.response.unified(success, data, message, error)
    };
  }

  /**
   * URL検証・サニタイズ
   * @param {string} url - 検証対象URL
   * @returns {string} サニタイズされたURL
   */
  static sanitizeURL(url) {
    if (!url || typeof url !== 'string') {
      return '';
    }

    try {
      let sanitized = String(url).trim();
      if (
        (sanitized.startsWith('"') && sanitized.endsWith('"')) ||
        (sanitized.startsWith("'") && sanitized.endsWith("'"))
      ) {
        sanitized = sanitized.slice(1, -1);
      }
      sanitized = sanitized.replace(/\\"/g, '"').replace(/\\'/g, "'");

      if (sanitized.includes('javascript:') || sanitized.includes('data:')) {
        UUtilities.logger.warn('URL検証', '危険なURLスキームを検出', sanitized);
        return '';
      }

      return sanitized;
    } catch (error) {
      UUtilities.logger.error('URL検証', 'URL検証エラー', error.message);
      return '';
    }
  }
  /**
   * 【Phase 7最適化】統合ログ管理
   * 冗長なログ出力パターン（258箇所）を統一
   */
  static get logger() {
    return {
    error: (context, message, details = null) => {
      const logMessage = `[ERROR] ${context}: ${message}`;
      if (details) {
        Log.error(logMessage, details);
      } else {
        Log.error(logMessage);
      }
    },
    warn: (context, message, details = null) => {
      const logMessage = `[WARN] ${context}: ${message}`;
      if (details) {
        Log.warn(logMessage, details);
      } else {
        Log.warn(logMessage);
      }
    },
    info: (context, message, details = null) => {
      const logMessage = `[INFO] ${context}: ${message}`;
      if (details) {
        Log.info(logMessage, details);
      } else {
        Log.info(logMessage);
      }
    },
    debug: (context, message, details = null) => {
      const logMessage = `[DEBUG] ${context}: ${message}`;
      if (details) {
        Log.info(logMessage, details);
      } else {
        Log.info(logMessage);
      }
    }
    };
  }

  /**
   * 【Phase 7最適化】統合エラーハンドリングヘルパー
   * try-catchパターンの冗長性を削減
   */
  static get safeExecute() {
    return {
    /**
     * 安全な関数実行（エラー時はレスポンス返却）
     */
    withResponse: (fn, context, errorMessage = '処理中にエラーが発生しました') => {
      try {
        const result = fn();
        return result;
      } catch (error) {
        UUtilities.logger.error(context, errorMessage, error.message);
        return UUtilities.generatorFactory.response.error(error.message, errorMessage);
      }
    },
    
    /**
     * 安全な非同期関数実行
     */
    async: async (fn, context, errorMessage = '非同期処理中にエラーが発生しました') => {
      try {
        const result = await fn();
        return result;
      } catch (error) {
        UUtilities.logger.error(context, errorMessage, error.message);
        throw error;
      }
    },

    /**
     * 安全なサービス初期化
     */
    service: (fn, serviceName) => {
      try {
        const service = fn();
        UUtilities.logger.debug('Service', `${serviceName}初期化成功`);
        return service;
      } catch (error) {
        UUtilities.logger.error('Service', `${serviceName}初期化失敗`, error.message);
        throw error;
      }
    }
    };
  }
}

// 重複クラス削除完了 - UnifiedAPIClient, UnifiedValidation は他のファイルで定義済み
// 検証機能は簡素化済み

// 後方互換性ラッパー関数は削除済み
// 統合されたクラスを直接使用してください
