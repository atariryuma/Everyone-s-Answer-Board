/**
 * @fileoverview Configuration Optimizer
 * データベース最適化とconfigJSON削減のためのユーティリティ
 */

/**
 * configJSONを最適化（重複データと不要フィールドを削除）
 * @param {Object} currentConfig - 現在のconfigJSON
 * @param {Object} userInfo - データベース行データ
 * @returns {Object} 最適化されたconfig
 */
function optimizeConfigJson(currentConfig, userInfo) {
  if (!currentConfig || typeof currentConfig !== 'object') {
    console.error('optimizeConfigJson: Invalid currentConfig');
    return null;
  }

  // 最小限のconfigJSONフィールドのみ保持
  const optimizedConfig = {
    title: currentConfig.title || `${userInfo.userEmail || '不明'}の回答ボード`,
    setupStatus: determineCorrectSetupStatus(currentConfig),
    formCreated: determineFormCreatedStatus(currentConfig),
    appPublished: currentConfig.appPublished || false,
    isPublic: currentConfig.isPublic || false,
    allowAnonymous: currentConfig.allowAnonymous || false,
    sheetName: currentConfig.sheetName || null,
    columnMapping: currentConfig.columnMapping || {},
    lastModified: new Date().toISOString(),
  };

  // columnsは本当に必要な場合のみ保持
  if (currentConfig.columns && Array.isArray(currentConfig.columns)) {
    optimizedConfig.columns = currentConfig.columns;
  }

  console.info('Config optimized', {
    originalSize: JSON.stringify(currentConfig).length,
    optimizedSize: JSON.stringify(optimizedConfig).length,
    reductionPercent: Math.round(
      (1 - JSON.stringify(optimizedConfig).length / JSON.stringify(currentConfig).length) * 100
    ),
  });

  return optimizedConfig;
}

/**
 * 正しいsetupStatusを判定
 */
function determineCorrectSetupStatus(config) {
  if (config.appPublished && config.spreadsheetId && config.sheetName) {
    return 'completed';
  } else if (config.spreadsheetId && config.sheetName) {
    return 'configured';
  }
  return 'pending';
}

/**
 * 正しいformCreatedStatusを判定
 */
function determineFormCreatedStatus(config) {
  return !!(config.appPublished && config.spreadsheetId);
}

/**
 * 動的フィールドを生成（configJSONに保存しない）
 * @param {Object} userInfo - データベースユーザー情報
 * @param {Object} config - 最適化されたconfig
 * @returns {Object} 動的に生成されたフィールド
 */
function generateDynamicFields(userInfo, config) {
  const userId = userInfo.userId;
  const baseUrl = getWebAppUrl();

  return {
    // 基本情報（データベース列から）
    userId: userInfo.userId,
    userEmail: userInfo.userEmail,
    createdAt: userInfo.createdAt,
    spreadsheetId: userInfo.spreadsheetId,

    // 動的生成フィールド
    appName: generateAppName(config.title),
    appUrl: `${baseUrl}?userId=${userId}`,
    formUrl: userInfo.formUrl || null,
    spreadsheetUrl: userInfo.spreadsheetId
      ? `https://docs.google.com/spreadsheets/d/${userInfo.spreadsheetId}/edit`
      : null,

    // 計算フィールド
    setupStep: determineSetupStep(userInfo, config),
    setupComplete: config.setupStatus === 'completed',
  };
}

/**
 * アプリ名を生成（重複除去）
 */
function generateAppName(title) {
  if (!title) return 'みんなの回答ボード';
  return title.replace(/回答ボード$/, '') + '回答ボード';
}

/**
 * ユーザー情報とconfigJSONを最適化統合
 * @param {string} userId - ユーザーID
 * @returns {Object} 統合されたユーザー情報
 */
function getOptimizedUserInfo(userId) {
  try {
    // データベースからユーザー情報取得
    const userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error(`ユーザーが見つかりません: ${userId}`);
    }

    // configJSONを解析
    let config = {};
    if (userInfo.configJson) {
      try {
        config = JSON.parse(userInfo.configJson);
      } catch (e) {
        console.error('configJSON解析エラー:', e);
      }
    }

    // configJSONを最適化
    const optimizedConfig = optimizeConfigJson(config, userInfo);

    // 動的フィールドを生成
    const dynamicFields = generateDynamicFields(userInfo, optimizedConfig);

    // 統合結果を返す
    return {
      ...dynamicFields,
      ...optimizedConfig,
    };
  } catch (error) {
    console.error('getOptimizedUserInfo エラー:', error);
    return null;
  }
}

/**
 * データベース構造の最適化を実行
 * @param {string} userId - ユーザーID
 * @returns {boolean} 最適化成功可否
 */
function optimizeUserDatabase(userId) {
  try {
    const userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }

    let config = {};
    if (userInfo.configJson) {
      config = JSON.parse(userInfo.configJson);
    }

    // データベース列に移行すべきデータ
    const dbUpdates = {};

    // spreadsheetIdをDBへ移行
    if (config.spreadsheetId && !userInfo.spreadsheetId) {
      dbUpdates.spreadsheetId = config.spreadsheetId;
    }

    // formUrlの生成（必要に応じて）
    if (config.formUrl && !userInfo.formUrl) {
      dbUpdates.formUrl = config.formUrl;
    }

    // configJSONを最適化
    const optimizedConfig = optimizeConfigJson(config, userInfo);
    dbUpdates.configJson = JSON.stringify(optimizedConfig);

    // lastAccessedAtを更新
    dbUpdates.lastAccessedAt = new Date().toISOString();

    // データベースを更新
    const success = updateUser(userId, dbUpdates);

    if (success) {
      // キャッシュクリア
      clearUserInfoCache(userId);
      console.info('データベース最適化完了:', userId);
    }

    return success;
  } catch (error) {
    console.error('optimizeUserDatabase エラー:', error);
    return false;
  }
}

/**
 * 特定ユーザーのデータベース最適化を実行（管理者用）
 * @param {string} targetUserId - 対象ユーザーID
 * @returns {Object} 最適化結果
 */
function optimizeSpecificUser(targetUserId) {
  const userId = targetUserId || '882d95c7-1fef-4739-a4b5-4ca02feaa69b';

  console.info('ユーザーデータベース最適化開始:', userId);

  const success = optimizeUserDatabase(userId);

  if (success) {
    const optimizedInfo = getOptimizedUserInfo(userId);
    return {
      success: true,
      userId: userId,
      optimizedData: optimizedInfo,
      message: 'データベース最適化が完了しました',
    };
  } else {
    return {
      success: false,
      userId: userId,
      message: '最適化に失敗しました',
    };
  }
}
