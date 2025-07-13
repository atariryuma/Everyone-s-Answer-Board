/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// =================================================================
// ユーザー情報キャッシュ（関数実行中の重複取得を防ぐ）
// =================================================================

var _executionUserInfoCache = null;

/**
 * 関数実行中のユーザー情報キャッシュをクリア
 */
function clearExecutionUserInfoCache() {
  _executionUserInfoCache = null;
}

/**
 * ユーザー情報を取得（実行中はキャッシュを使用）
 * @param {string} userId ユーザーID
 * @returns {Object} ユーザー情報
 */
function getCachedUserInfo(userId) {
  if (_executionUserInfoCache && _executionUserInfoCache.userId === userId) {
    return _executionUserInfoCache.userInfo;
  }
  
  const userInfo = findUserById(userId);
  _executionUserInfoCache = { userId, userInfo };
  return userInfo;
}

// =================================================================
// メインロジック
// =================================================================

// doGetLegacy function removed - consolidated into main doGet in UltraOptimizedCore.gs

/**
 * 意見ヘッダーを安全に取得する関数（テンプレート変数の問題を回避）
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {string} 意見ヘッダー
 */
function getOpinionHeaderSafely(userId, sheetName) {
  try {
    const userInfo = getCachedUserInfo(userId);
    if (!userInfo) {
      return 'お題';
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
    const sheetConfig = config[sheetConfigKey] || {};
    
    const opinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';
    
    debugLog('getOpinionHeaderSafely:', {
      userId: userId,
      sheetName: sheetName,
      opinionHeader: opinionHeader
    });
    
    return opinionHeader;
  } catch (e) {
    console.error('getOpinionHeaderSafely error:', e);
    return 'お題';
  }
}

/**
 * フォームURLを安全に取得する関数
 * configJsonにformUrlがない場合、スプレッドシートから取得を試みる
 * @param {Object} configJson - 設定JSON
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string} フォームURL
 */
function getFormUrlSafely(configJson, spreadsheetId) {
  try {
    // 既にconfigJsonにformUrlがある場合はそれを返す
    if (configJson && configJson.formUrl) {
      return configJson.formUrl;
    }
    
    // スプレッドシートIDがない場合は空文字を返す
    if (!spreadsheetId) {
      return '';
    }
    
    // スプレッドシートからフォームURLを取得を試みる
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const form = spreadsheet.getFormUrl();
      if (form) {
        debugLog('getFormUrlSafely: スプレッドシートからフォームURLを取得:', form);
        return form;
      }
    } catch (e) {
      debugLog('getFormUrlSafely: スプレッドシートからのフォームURL取得に失敗:', e.message);
    }
    
    return '';
  } catch (e) {
    console.error('getFormUrlSafely error:', e);
    return '';
  }
}

/**
 * 新規ユーザーを登録する（データベース登録のみ）
 * フォーム作成はクイックスタートで実行される
 */
function registerNewUser(adminEmail) {
  var activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // ドメイン制限チェック
  var domainInfo = getDeployUserDomainInfo();
  if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
    throw new Error(`ドメインアクセスが制限されています。許可されたドメイン: ${domainInfo.deployDomain}, 現在のドメイン: ${domainInfo.currentDomain}`);
  }

  // 既存ユーザーチェック（1ユーザー1行の原則）
  var existingUser = findUserByEmail(adminEmail);
  var userId, appUrls;
  
  if (existingUser) {
    // 既存ユーザーの場合は情報を更新
    userId = existingUser.userId;
    var existingConfig = JSON.parse(existingUser.configJson || '{}');
    
    // 設定をリセット（新規登録状態に戻す）
    var updatedConfig = {
      ...existingConfig,
      setupStatus: 'pending',
      lastRegistration: new Date().toISOString(),
      formCreated: false,
      appPublished: false
    };
    
    // 既存ユーザー情報を更新
    updateUser(userId, {
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: JSON.stringify(updatedConfig)
    });
    // キャッシュを無効化して最新状態を反映
    invalidateUserCache(userId, adminEmail, existingUser.spreadsheetId, false);
    
    debugLog('✅ 既存ユーザー情報を更新しました: ' + adminEmail);
    appUrls = generateAppUrls(userId);
    
    return {
      userId: userId,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupRequired: true,
      message: '既存ユーザーの情報を更新しました。クイックスタートで新しいフォームを作成してください。',
      isExistingUser: true
    };
  }

  // 新規ユーザーの場合
  userId = Utilities.getUuid();
  
  var initialConfig = {
    setupStatus: 'pending',
    createdAt: new Date().toISOString(),
    formCreated: false,
    appPublished: false
  };
  
  var userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: '',
    spreadsheetUrl: '',
    createdAt: new Date().toISOString(),
    configJson: JSON.stringify(initialConfig),
    lastAccessedAt: new Date().toISOString(),
    isActive: 'true'
  };

  try {
    createUser(userData);
    debugLog('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);
    // 生成されたユーザー情報のキャッシュをクリア
    invalidateUserCache(userId, adminEmail, null, false);
  } catch (e) {
    console.error('データベースへのユーザー登録に失敗: ' + e.message);
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す
  appUrls = generateAppUrls(userId);
  return {
    userId: userId,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupRequired: true,
    message: 'ユーザー登録が完了しました！次にクイックスタートでフォームを作成してください。',
    isExistingUser: false
  };
}

/**
 * リアクションを追加/削除する (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function addReaction(requestUserId, rowIndex, reactionKey, sheetName) {
  verifyUserAccess(requestUserId);
  clearExecutionUserInfoCache();

  try {
    var reactingUserEmail = Session.getActiveUser().getEmail();
    var ownerUserId = requestUserId; // requestUserId を使用

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    var boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    var result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      // 更新後のリアクション情報を取得
      var updatedReactions = getRowReactions(boardOwnerInfo.spreadsheetId, sheetName, rowIndex, reactingUserEmail);

      return {
        status: "ok",
        reactions: updatedReactions
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    console.error('addReaction エラー: ' + e.message);
    return {
      status: "error",
      message: e.message
    };
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * マルチテナント環境でのユーザーアクセス権限を検証します。
 * リクエストを投げたユーザー (Session.getActiveUser().getEmail()) が、
 * requestUserId のデータにアクセスする権限を持っているかを確認します。
 * 権限がない場合はエラーをスローします。
 * @param {string} requestUserId - アクセスを要求しているユーザーのID
 * @throws {Error} 認証エラーまたは権限エラー
 */
function verifyUserAccess(requestUserId) {
  clearExecutionUserInfoCache(); // キャッシュをクリアして最新のユーザー情報を取得
  const activeUserEmail = Session.getActiveUser().getEmail();
  const requestedUserInfo = findUserById(requestUserId);

  if (!requestedUserInfo) {
    throw new Error(`認証エラー: 指定されたユーザーID (${requestUserId}) が見つかりません。`);
  }

  // リクエストユーザーのメールアドレスと、要求されたユーザーIDのadminEmailが一致するか検証
  if (activeUserEmail !== requestedUserInfo.adminEmail) {
    throw new Error(`権限エラー: ${activeUserEmail} はユーザーID ${requestUserId} のデータにアクセスする権限がありません。`);
  }
  debugLog(`✅ ユーザーアクセス検証成功: ${activeUserEmail} は ${requestUserId} のデータにアクセスできます。`);
}

/**
 * 公開されたシートのデータを取得 (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode, bypassCache) {
  verifyUserAccess(requestUserId);
  clearExecutionUserInfoCache(); // キャッシュをクリアして最新のユーザー情報を取得

  try {
    // キャッシュキー生成（パフォーマンス向上）
    var requestKey = `publishedData_${requestUserId}_${classFilter}_${sortOrder}_${adminMode}`;

    // キャッシュバイパス時は直接実行
    if (bypassCache === true) {
      debugLog('🔄 キャッシュバイパス：最新データを直接取得');
      return executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode);
    }

    return cacheManager.get(requestKey, () => {
      return executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode);
    }, { ttl: 600 }); // 10分間キャッシュ
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

/**
 * 実際のデータ取得処理（キャッシュ制御から分離） (マルチテナント対応版)
 */
function executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode) {
    try {
      var currentUserId = requestUserId; // requestUserId を使用
      debugLog('getPublishedSheetData: userId=%s, classFilter=%s, sortOrder=%s, adminMode=%s', currentUserId, classFilter, sortOrder, adminMode);

      var userInfo = getCachedUserInfo(currentUserId);
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }
      debugLog('getPublishedSheetData: userInfo=%s', JSON.stringify(userInfo));

      var configJson = JSON.parse(userInfo.configJson || '{}');
      debugLog('getPublishedSheetData: configJson=%s', JSON.stringify(configJson));

    // 公開対象のスプレッドシートIDとシート名を取得
    var publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    var publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // シート固有の設定を取得 (sheetKey is based only on sheet name)
    var sheetKey = 'sheet_' + publishedSheetName;
    var sheetConfig = configJson[sheetKey] || {};
    debugLog('getPublishedSheetData: sheetConfig=%s', JSON.stringify(sheetConfig));

    // Check if current user is the board owner
    var isOwner = (configJson.ownerId === currentUserId);
    debugLog('getPublishedSheetData: isOwner=%s, ownerId=%s, currentUserId=%s', isOwner, configJson.ownerId, currentUserId);

    // データ取得
    var sheetData = getSheetData(currentUserId, publishedSheetName, classFilter, sortOrder, adminMode);
    debugLog('getPublishedSheetData: sheetData status=%s, totalCount=%s', sheetData.status, sheetData.totalCount);

    if (sheetData.status === 'error') {
      throw new Error(sheetData.message);
    }

    // Page.html期待形式に変換
    // 設定からヘッダー名を取得。未定義の場合のみデフォルト値を使用。
    var mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    var reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
    var classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
    var nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    debugLog('getPublishedSheetData: Configured Headers - mainHeaderName=%s, reasonHeaderName=%s, classHeaderName=%s, nameHeaderName=%s', mainHeaderName, reasonHeaderName, classHeaderName, nameHeaderName);

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    var headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);
    debugLog('getPublishedSheetData: Available headerIndices=%s', JSON.stringify(headerIndices));

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    var mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName,
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);
    debugLog('getPublishedSheetData: Mapped indices=%s', JSON.stringify(mappedIndices));

    var formattedData = formatSheetDataForFrontend(sheetData.data, mappedIndices, headerIndices, adminMode, isOwner, sheetData.displayMode);

    debugLog('getPublishedSheetData: formattedData length=%s', formattedData.length);

    // ボードのタイトルを実際のスプレッドシートのヘッダーから取得
    let headerTitle = publishedSheetName || '今日のお題';
    if (mappedIndices.opinionHeader !== undefined) {
      for (var actualHeader in headerIndices) {
        if (headerIndices[actualHeader] === mappedIndices.opinionHeader) {
          headerTitle = actualHeader;
          debugLog('getPublishedSheetData: Using actual header as title: "%s"', headerTitle);
          break;
        }
      }
    }

    var finalDisplayMode = (adminMode === true) ? DISPLAY_MODES.NAMED : (sheetData.displayMode || DISPLAY_MODES.ANONYMOUS);

    var result = {
      header: headerTitle,
      sheetName: publishedSheetName,
      showCounts: (adminMode === true) ? true : (configJson.showCounts === true),
      displayMode: finalDisplayMode,
      data: formattedData,
      rows: formattedData // 後方互換性のため
    };

    debugLog('🔍 最終結果:', {
      adminMode: adminMode,
      originalDisplayMode: sheetData.displayMode,
      finalDisplayMode: finalDisplayMode,
      dataCount: formattedData.length,
      showCounts: result.showCounts
    });
    debugLog('getPublishedSheetData: Returning result=%s', JSON.stringify(result));
    return result;

  } catch (e) {
    console.error('公開シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      rows: []
    };
  }
}

/**
 * 増分データ取得機能：指定された基準点以降の新しいデータのみを取得 (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @param {number} sinceRowCount - この行数以降のデータを取得
 * @returns {object} 新しいデータのみを含む結果
 */
function getIncrementalSheetData(requestUserId, classFilter, sortOrder, adminMode, sinceRowCount) {
  verifyUserAccess(requestUserId);
  try {
    debugLog('🔄 増分データ取得開始: sinceRowCount=%s', sinceRowCount);

    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = getCachedUserInfo(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    var publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    var publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // スプレッドシートとシートを取得
    var spreadsheet = SpreadsheetApp.openById(publishedSpreadsheetId);
    var sheet = spreadsheet.getSheetByName(publishedSheetName);

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + publishedSheetName);
    }

    var lastRow = sheet.getLastRow(); // スプレッドシートの最終行
    var headerRow = 1; // ヘッダー行は1行目と仮定

    // 実際に読み込むべき開始行を計算 (sinceRowCountはデータ行数なので、+1してヘッダーを考慮)
    // sinceRowCountが0の場合、ヘッダーの次の行から読み込む
    var startRowToRead = sinceRowCount + headerRow + 1;

    // 新しいデータがない場合
    if (lastRow < startRowToRead) {
      debugLog('🔍 増分データ分析: 新しいデータなし。lastRow=%s, startRowToRead=%s', lastRow, startRowToRead);
      return {
        header: '', // 必要に応じて設定
        sheetName: publishedSheetName,
        showCounts: false, // 必要に応じて設定
        displayMode: '', // 必要に応じて設定
        data: [],
        rows: [],
        totalCount: lastRow - headerRow, // ヘッダーを除いたデータ総数
        newCount: 0,
        isIncremental: true
      };
    }

    // 読み込む行数
    var numRowsToRead = lastRow - startRowToRead + 1;

    // 必要なデータのみをスプレッドシートから直接取得
    // getRange(row, column, numRows, numColumns)
    // ここでは全列を取得すると仮定 (A列から最終列まで)
    var lastColumn = sheet.getLastColumn();
    var rawNewData = sheet.getRange(startRowToRead, 1, numRowsToRead, lastColumn).getValues();

    debugLog('📥 スプレッドシートから直接取得した新しいデータ:', rawNewData.length, '件');

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    var headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    var sheetConfig = configJson['sheet_' + publishedSheetName] || {};
    var mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    var reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
    var classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
    var nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    var mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName,
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);

    // ユーザー情報と管理者モードの取得
    var isOwner = (configJson.ownerId === currentUserId);
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // 新しいデータを既存の処理パイプラインと同様に加工
    var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    var rosterMap = buildRosterMap([]); // roster is not used
    var processedData = rawNewData.map(function(row, idx) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, startRowToRead + idx, isOwner);
    });

    // 取得した生データをPage.htmlが期待する形式にフォーマット
    var formattedNewData = formatSheetDataForFrontend(processedData, mappedIndices, headerIndices, adminMode, isOwner, displayMode);

    debugLog('✅ 増分データ取得完了: %s件の新しいデータを返します', formattedNewData.length);

    return {
      header: '', // 必要に応じて設定
      sheetName: publishedSheetName,
      showCounts: false, // 必要に応じて設定
      displayMode: displayMode,
      data: formattedNewData,
      rows: formattedNewData, // 後方互換性のため
      totalCount: lastRow - headerRow, // ヘッダーを除いたデータ総数
      newCount: formattedNewData.length,
      isIncremental: true
    };
  } catch (e) {
    console.error('増分データ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: '増分データの取得に失敗しました: ' + e.message
    };
  }
}

/**
 * 利用可能なシート一覧を取得 (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getAvailableSheets(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用

    if (!currentUserId) {
      console.warn('getAvailableSheets: No current user ID set');
      return [];
    }

    var sheets = getSheetsList(currentUserId);

    if (!sheets || sheets.length === 0) {
      console.warn('getAvailableSheets: No sheets found for user:', currentUserId);
      return [];
    }

    // Page.html期待形式に変換: [{name: string}]
    return sheets.map(function(sheet) {
      return {
        name: typeof sheet === 'string' ? sheet : (sheet.name || sheet.title || sheet)
      };
    });
  } catch (e) {
    console.error('getAvailableSheets エラー: ' + e.message);
    console.error('Error details:', e.stack);
    return [];
  }
}

/**
 * ボードデータを再読み込み (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function refreshBoardData(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // キャッシュをクリア
    invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId, false);

    // 最新のステータスを取得
    return getAppConfig(requestUserId);
  } catch (e) {
    console.error('refreshBoardData の再読み込みに失敗: ' + e.message);
    return { status: 'error', message: 'ボードデータの再読み込みに失敗しました: ' + e.message };
  }
}

/**
 * スプレッドシートの生データをフロントエンドが期待する形式にフォーマットするヘルパー関数
 * @param {Array<Object>} rawData - getSheetDataから返された生データ（originalData, reactionCountsなどを含む）
 * @param {Object} mappedIndices - 設定されたヘッダー名と実際の列インデックスのマッピング
 * @param {Object} headerIndices - 実際のヘッダー名と列インデックスのマッピング
 * @param {boolean} adminMode - 管理者モードかどうか
 * @param {boolean} isOwner - 現在のユーザーがボードのオーナーかどうか
 * @param {string} displayMode - 表示モード（'named' or 'anonymous'）
 * @returns {Array<Object>} フォーマットされたデータ
 */
function formatSheetDataForFrontend(rawData, mappedIndices, headerIndices, adminMode, isOwner, displayMode) {
  return rawData.map(function(row, index) {
    var classIndex = mappedIndices.classHeader;
    var opinionIndex = mappedIndices.opinionHeader;
    var reasonIndex = mappedIndices.reasonHeader;
    var nameIndex = mappedIndices.nameHeader;

    debugLog('formatSheetDataForFrontend: Row %s - classIndex=%s, opinionIndex=%s, reasonIndex=%s, nameIndex=%s', index, classIndex, opinionIndex, reasonIndex, nameIndex);
    debugLog('formatSheetDataForFrontend: Row data length=%s, first few values=%s', row.originalData ? row.originalData.length : 'undefined', row.originalData ? JSON.stringify(row.originalData.slice(0, 5)) : 'undefined');
    
    var nameValue = '';
    var shouldShowName = (adminMode === true || displayMode === DISPLAY_MODES.NAMED || isOwner);
    var hasNameIndex = nameIndex !== undefined;
    var hasOriginalData = row.originalData && row.originalData.length > 0;
    
    if (shouldShowName && hasNameIndex && hasOriginalData) {
      nameValue = row.originalData[nameIndex] || '';
    }
    
    if (!nameValue && shouldShowName && hasOriginalData) {
      var emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
      if (emailIndex !== undefined && row.originalData[emailIndex]) {
        nameValue = row.originalData[emailIndex].split('@')[0];
      }
    }
    
    debugLog('🔍 サーバー側名前データ詳細:', {
      rowIndex: row.rowNumber || (index + 2),
      shouldShowName: shouldShowName,
      adminMode: adminMode,
      displayMode: displayMode,
      isOwner: isOwner,
      nameIndex: nameIndex,
      hasNameIndex: hasNameIndex,
      hasOriginalData: hasOriginalData,
      originalDataLength: row.originalData ? row.originalData.length : 'undefined',
      nameValue: nameValue,
      rawNameData: hasOriginalData && nameIndex !== undefined ? row.originalData[nameIndex] : 'N/A'
    });

    return {
      rowIndex: row.rowNumber || (index + 2),
      name: nameValue,
      email: row.originalData && row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]] ? row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]] : '',
      class: (classIndex !== undefined && row.originalData && row.originalData[classIndex]) ? row.originalData[classIndex] : '',
      opinion: (opinionIndex !== undefined && row.originalData && row.originalData[opinionIndex]) ? row.originalData[opinionIndex] : '',
      reason: (reasonIndex !== undefined && row.originalData && row.originalData[reasonIndex]) ? row.originalData[reasonIndex] : '',
      reactions: {
        UNDERSTAND: { count: row.understandCount || 0, reacted: false },
        LIKE: { count: row.likeCount || 0, reacted: false },
        CURIOUS: { count: row.curiousCount || 0, reacted: false }
      },
      highlight: row.isHighlighted || false
    };
  });
}

/**
 * アプリ設定を取得（最適化版） (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getAppConfig(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');

    // --- Auto-healing for inconsistent setup states ---
    var needsUpdate = false;
    if (configJson.formUrl && !configJson.formCreated) {
      configJson.formCreated = true;
      needsUpdate = true;
    }
    if (configJson.formCreated && configJson.setupStatus !== 'completed') {
      configJson.setupStatus = 'completed';
      needsUpdate = true;
    }
    if (configJson.publishedSheetName && !configJson.appPublished) {
      configJson.appPublished = true;
      needsUpdate = true;
    }
    if (needsUpdate) {
      try {
        updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
      } catch (updateErr) {
        console.warn('Config auto-heal failed: ' + updateErr.message);
      }
    }

    var sheets = getSheetsList(currentUserId);
    var appUrls = generateAppUrls(currentUserId);
    
    // 回答数を取得
    var answerCount = 0;
    var totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          // リアクション数の概算計算（詳細実装は後回し）
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (countError) {
      console.warn('回答数の取得に失敗: ' + countError.message);
    }
    
    return {
      status: 'success',
      userId: currentUserId,
      adminEmail: userInfo.adminEmail,
      publishedSpreadsheetId: configJson.publishedSpreadsheetId || '',
      publishedSheetName: configJson.publishedSheetName || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false,
      appPublished: configJson.appPublished || false, // AdminPanel.htmlで使用される
      availableSheets: sheets,
      allSheets: sheets, // AdminPanel.htmlで使用される
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: getFormUrlSafely(configJson, userInfo.spreadsheetId),
      editFormUrl: configJson.editFormUrl || '',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      activeSheetName: configJson.publishedSheetName || '',
      appUrls: appUrls,
      // AdminPanel.htmlが期待する表示設定プロパティ
      showNames: configJson.showNames || false,
      showCounts: configJson.showCounts === true,
      // データベース詳細情報
      userInfo: {
        userId: currentUserId,
        adminEmail: userInfo.adminEmail,
        spreadsheetId: userInfo.spreadsheetId || '',
        spreadsheetUrl: userInfo.spreadsheetUrl || '',
        createdAt: userInfo.createdAt || '',
        lastAccessedAt: userInfo.lastAccessedAt || '',
        isActive: userInfo.isActive || 'false',
        configJson: userInfo.configJson || '{}'
      },
      // 統計情報
      answerCount: answerCount,
      totalReactions: totalReactions,
      // システム状態
      systemStatus: {
        setupStatus: configJson.setupStatus || 'unknown',
        formCreated: configJson.formCreated || false,
        appPublished: configJson.appPublished || false,
        lastUpdated: new Date().toISOString()
      },
      // ユーザーのセットアップ段階を判定
      setupStep: (userInfo && userInfo.spreadsheetId) ? (configJson.publishedSheetName ? 3 : 2) : 1
    };
  } catch (e) {
    console.error('アプリ設定取得エラー: ' + e.message);
    return {
      status: 'error',
      message: '設定の取得に失敗しました: ' + e.message
    };
  }
}

/**
 * シート設定を保存する
 * AdminPanel.htmlから呼び出される
 * @param {string} spreadsheetId - 設定対象のスプレッドシートID
 * @param {string} sheetName - 設定対象のシート名
 * @param {object} config - 保存するシート固有の設定
 */
function saveSheetConfig(userId, spreadsheetId, sheetName, config) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    if (!config || typeof config !== 'object') {
      throw new Error('無効なconfigです: ' + config);
    }
    
    var currentUserId = userId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    // save using sheet-specific key expected by getConfig
    var sheetKey = 'sheet_' + sheetName;
    configJson[sheetKey] = config;

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    debugLog('✅ シート設定を保存しました: %s', sheetKey);
    return { status: 'success', message: 'シート設定を保存しました。' };
  } catch (e) {
    console.error('シート設定保存エラー: ' + e.message);
    return { status: 'error', message: 'シート設定の保存に失敗しました: ' + e.message };
  }
}

/**
 * 表示するシートを切り替える
 * AdminPanel.htmlから呼び出される
 * @param {string} spreadsheetId - 公開対象のスプレッドシートID
 * @param {string} sheetName - 公開対象のシート名
 */
function switchToSheet(userId, spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    
    var currentUserId = userId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    configJson.publishedSpreadsheetId = spreadsheetId;
    configJson.publishedSheetName = sheetName;
    configJson.appPublished = true; // シートを切り替えたら公開状態にする

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    debugLog('✅ 表示シートを切り替えました: %s - %s', spreadsheetId, sheetName);
    return { status: 'success', message: '表示シートを切り替えました。' };
  } catch (e) {
    console.error('シート切り替えエラー: ' + e.message);
    return { status: 'error', message: '表示シートの切り替えに失敗しました: ' + e.message };
  }
}

// =================================================================
// セットアップ関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行） (マルチテナント対応版)
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    if (typeof dbId !== 'string' || dbId.length !== 44) {
      throw new Error('無効なスプレッドシートIDです。IDは44文字の文字列である必要があります。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    var adminEmail = Session.getEffectiveUser().getEmail();
    if (adminEmail) {
      props.setProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL, adminEmail);
    }

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    debugLog('✅ セットアップが正常に完了しました。');
    return { status: 'success', message: 'セットアップが正常に完了しました。' };
  } catch (e) {
    console.error('セットアップエラー:', e);
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * セットアップ状態をテストする (マルチテナント対応版)
 */
function testSetup() {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var creds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    
    if (!dbId) {
      return { status: 'error', message: 'データベーススプレッドシートIDが設定されていません。' };
    }
    
    if (!creds) {
      return { status: 'error', message: 'サービスアカウント認証情報が設定されていません。' };
    }
    
    // データベースへの接続テスト
    try {
      var userInfo = findUserByEmail(Session.getActiveUser().getEmail());
      return { 
        status: 'success', 
        message: 'セットアップは正常に完了しています。システムは使用準備が整いました。',
        details: {
          databaseConnected: true,
          userCount: userInfo ? 'ユーザー登録済み' : '未登録',
          serviceAccountConfigured: true
        }
      };
    } catch (dbError) {
      return { 
        status: 'warning', 
        message: '設定は保存されていますが、データベースアクセスに問題があります。',
        details: { error: dbError.message }
      };
    }
    
  } catch (e) {
    console.error('セットアップテストエラー:', e);
    return { status: 'error', message: 'セットアップテストに失敗しました: ' + e.message };
  }
}

// =================================================================
// ヘルパー関数
// =================================================================

// include 関数は main.gs で定義されています


function getResponsesData(userId, sheetName) {
  var userInfo = getCachedUserInfo(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var service = getSheetsService();
    var spreadsheetId = userInfo.spreadsheetId;
    var range = "'" + (sheetName || 'フォームの回答 1') + "'!A:Z";
    
    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var values = response.valueRanges[0].values || [];
    
    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1),
      headers: values[0]
    };
  } catch (e) {
    console.error('回答データの取得に失敗: ' + e.message);
    return { status: 'error', message: '回答データの取得に失敗しました: ' + e.message };
  }
}

// =================================================================
// HTML依存関数（UI連携）
// =================================================================

/**
 * 現在のユーザーのステータスを取得し、UIに表示するための情報を返します。(マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} ユーザーのステータス情報
 */
function getCurrentUserStatus(requestUserId) {
  try {
    const activeUserEmail = Session.getActiveUser().getEmail();
    
    // requestUserIdが無効な場合は、メールアドレスでユーザーを検索
    let userInfo;
    if (requestUserId && requestUserId.trim() !== '') {
      verifyUserAccess(requestUserId);
      userInfo = findUserById(requestUserId);
    } else {
      userInfo = findUserByEmail(activeUserEmail);
    }

    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません。' };
    }

    return {
      status: 'success',
      userInfo: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt
      }
    };
  } catch (e) {
    console.error('getCurrentUserStatus エラー: ' + e.message);
    return { status: 'error', message: 'ステータス取得に失敗しました: ' + e.message };
  }
}

/**
 * アクティブなフォーム情報を取得 (マルチテナント対応版)
 * Page.htmlから呼び出される（パラメータなし）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getActiveFormInfo(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    // フォーム回答数を取得
    var answerCount = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheet) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheet);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
        }
      }
    } catch (countError) {
      console.warn('回答数の取得に失敗: ' + countError.message);
    }

    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      editUrl: configJson.editFormUrl || '',  // AdminPanel.htmlが期待するフィールド名
      formId: extractFormIdFromUrl(configJson.formUrl || configJson.editFormUrl || ''),
      spreadsheetUrl: userInfo.spreadsheetUrl || '',
      answerCount: answerCount,
      isFormActive: !!(configJson.formUrl && configJson.formCreated)
    };
  } catch (e) {
    console.error('getActiveFormInfo エラー: ' + e.message);
    return { status: 'error', message: 'フォーム情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ユーザーが管理者かどうかをチェックする (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {boolean} 管理者の場合はtrue、そうでない場合はfalse
 */
function checkAdmin(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    // Session.getActiveUser().getEmail() が requestUserId の adminEmail と一致するかどうかを verifyUserAccess で既にチェック済み
    // ここでは単に userInfo.adminEmail と Session.getActiveUser().getEmail() が一致するかを返す
    return Session.getActiveUser().getEmail() === userInfo.adminEmail;
  } catch (e) {
    console.error('checkAdmin エラー: ' + e.message);
    return false;
  }
}

/**
 * データ数を取得する (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {number} データ数
 */
function getDataCount(requestUserId, classFilter, sortOrder, adminMode) {
  verifyUserAccess(requestUserId);
  try {
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    const configJson = JSON.parse(userInfo.configJson || '{}');

    if (!configJson.publishedSpreadsheetId || !configJson.publishedSheetName) {
      return {
        count: 0,
        lastUpdate: new Date().toISOString(),
        status: 'error',
        message: 'スプレッドシート設定なし'
      };
    }

    // シートデータを取得して件数を返す
    const sheetData = getSheetData(requestUserId, configJson.publishedSheetName, classFilter, sortOrder, adminMode);
    if (sheetData.status === 'success') {
      return {
        count: sheetData.totalCount || 0,
        lastUpdate: new Date().toISOString(),
        status: 'success'
      };
    } else {
      return {
        count: 0,
        lastUpdate: new Date().toISOString(),
        status: 'error',
        message: sheetData.message
      };
    }
  } catch (e) {
    console.error('getDataCount エラー: ' + e.message);
    return {
      count: 0,
      lastUpdate: new Date().toISOString(),
      status: 'error',
      message: e.message
    };
  }
}

/**
 * フォーム設定を更新 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} title - 新しいタイトル
 * @param {string} description - 新しい説明
 */
function updateFormSettings(requestUserId, title, description) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    if (configJson.editFormUrl) {
      try {
        var formId = extractFormIdFromUrl(configJson.editFormUrl);
        var form = FormApp.openById(formId);
        
        if (title) {
          form.setTitle(title);
        }
        if (description) {
          form.setDescription(description);
        }
        
        return {
          status: 'success',
          message: 'フォーム設定が更新されました'
        };
      } catch (formError) {
        console.error('フォーム更新エラー: ' + formError.message);
        return { status: 'error', message: 'フォームの更新に失敗しました: ' + formError.message };
      }
    } else {
      return { status: 'error', message: 'フォームが見つかりません' };
    }
  } catch (e) {
    console.error('updateFormSettings エラー: ' + e.message);
    return { status: 'error', message: 'フォーム設定の更新に失敗しました: ' + e.message };
  }
}

/**
 * システム設定を保存 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 */
function saveSystemConfig(requestUserId, config) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // システム設定を更新
    configJson.systemConfig = {
      cacheEnabled: config.cacheEnabled,
      autosaveInterval: config.autosaveInterval,
      logLevel: config.logLevel,
      updatedAt: new Date().toISOString()
    };
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return {
      status: 'success',
      message: 'システム設定が保存されました'
    };
  } catch (e) {
    console.error('saveSystemConfig エラー: ' + e.message);
    return { status: 'error', message: 'システム設定の保存に失敗しました: ' + e.message };
  }
}

/**
 * ハイライト状態の切り替え (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function toggleHighlight(requestUserId, rowIndex, sheetName) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 管理者権限チェック - 現在のユーザーがボードの所有者かどうかを確認
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (activeUserEmail !== userInfo.adminEmail) {
      throw new Error('ハイライト機能は管理者のみ使用できます');
    }

    var result = processHighlightToggle(
      userInfo.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      return {
        status: "ok",
        highlight: result.highlighted || false
      };
    } else {
      throw new Error(result.message || 'ハイライト切り替えに失敗しました');
    }
  } catch (e) {
    console.error('toggleHighlight エラー: ' + e.message);
    return { 
      status: "error", 
      message: e.message 
    };
  }
}



/**
 * 利用可能なシート一覧を取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */

/**
 * クイックスタートセットアップ（完全版） (マルチテナント対応版)
 * フォルダ作成、フォーム作成、スプレッドシート作成、ボード公開まで一括実行
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function quickStartSetup(requestUserId) {
  // 新規ユーザー（requestUserIdがundefinedまたはnull）の場合はverifyUserAccessをスキップ
  if (requestUserId) {
    verifyUserAccess(requestUserId);
  }
  try {
    debugLog('🚀 クイックスタートセットアップ開始: ' + requestUserId);
    
    // ユーザー情報の取得
    var userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var userEmail = userInfo.adminEmail;
    
    // 既にセットアップ済みかチェック
    if (configJson.formCreated && userInfo.spreadsheetId) {
      var appUrls = generateAppUrls(requestUserId);
      return {
        status: 'already_completed',
        message: 'クイックスタートは既に完了しています。',
        webAppUrl: appUrls.webAppUrl,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl,
        setupUrl: appUrls.setupUrl
      };
    }
    
    // ステップ1: ユーザー専用フォルダを作成
    debugLog('📁 ステップ1: フォルダ作成中...');
    var folder = createUserFolder(userEmail);
    
    // ステップ2: Googleフォームとスプレッドシートを作成
    debugLog('📝 ステップ2: フォーム作成中...');
    var formAndSsInfo = createStudyQuestForm(userEmail, requestUserId);
    
    // 作成したファイルをフォルダに移動
    if (folder) {
      try {
        var formFile = DriveApp.getFileById(formAndSsInfo.formId);
        var ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);
        
        folder.addFile(formFile);
        folder.addFile(ssFile);
        
        // 元の場所から削除（Myドライブから移動）
        DriveApp.getRootFolder().removeFile(formFile);
        DriveApp.getRootFolder().removeFile(ssFile);
        
        debugLog('📁 ファイルをフォルダに移動しました: ' + folder.getName());
      } catch (moveError) {
        // フォルダ移動に失敗しても処理は継続
        console.warn('ファイル移動に失敗しましたが、処理を継続します: ' + moveError.message);
      }
    }
    
    // ステップ3: データベースを更新
    debugLog('💾 ステップ3: データベース更新中...');
    
    // クイックスタート用の適切な初期設定を作成
    var sheetConfigKey = 'sheet_' + formAndSsInfo.sheetName;
    var quickStartSheetConfig = {
      opinionHeader: '今日のテーマについて、あなたの考えや意見を聞かせてください',
      reasonHeader: 'そう考える理由や体験があれば教えてください（任意）',
      nameHeader: '名前',
      classHeader: 'クラス',
      showNames: false,
      showCounts: false,
      lastModified: new Date().toISOString()
    };
    
    var updatedConfig = {
      ...configJson,
      setupStatus: 'completed',
      formCreated: true,
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      publishedSpreadsheetId: formAndSsInfo.spreadsheetId,
      publishedSheetName: formAndSsInfo.sheetName,
      appPublished: true,
      folderId: folder ? folder.getId() : '',
      folderUrl: folder ? folder.getUrl() : '',
      completedAt: new Date().toISOString(),
      [sheetConfigKey]: quickStartSheetConfig
    };
    
    updateUser(requestUserId, {
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      configJson: JSON.stringify(updatedConfig)
    });
    // セットアップ完了後に関連キャッシュをクリア
    invalidateUserCache(requestUserId, userEmail, formAndSsInfo.spreadsheetId, true);
    
    // ステップ4: 回答ボードを公開状態に設定
    debugLog('🌐 ステップ4: 回答ボード公開中...');
    
    debugLog('✅ クイックスタートセットアップ完了: ' + requestUserId);
    
    var appUrls = generateAppUrls(requestUserId);
    return {
      status: 'success',
      message: 'クイックスタートが完了しました！回答ボードをお楽しみください。',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupUrl: appUrls.setupUrl,
      formUrl: updatedConfig.formUrl,
      editFormUrl: updatedConfig.editFormUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      folderUrl: updatedConfig.folderUrl
    };
    
  } catch (e) {
    console.error('❌ quickStartSetup エラー: ' + e.message);
    
    // エラー時はセットアップ状態をリセット
    try {
      var currentConfig = JSON.parse(userInfo.configJson || '{}');
      currentConfig.setupStatus = 'error';
      currentConfig.lastError = e.message;
      currentConfig.errorAt = new Date().toISOString();
      
      updateUser(requestUserId, {
        configJson: JSON.stringify(currentConfig)
      });
      invalidateUserCache(requestUserId, userEmail, null, false);
    } catch (updateError) {
      console.error('エラー状態の更新に失敗: ' + updateError.message);
    }
    
    return {
      status: 'error',
      message: 'クイックスタートセットアップに失敗しました: ' + e.message,
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: ''
    };
  }
}

/**
 * ユーザー専用フォルダを作成 (マルチテナント対応版)
 */
function createUserFolder(userEmail) {
  try {
    var rootFolderName = "StudyQuest - ユーザーデータ";
    var userFolderName = "StudyQuest - " + userEmail + " - ファイル";
    
    // ルートフォルダを検索または作成
    var rootFolder;
    var folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }
    
    // ユーザー専用フォルダを作成
    var userFolders = rootFolder.getFoldersByName(userFolderName);
    if (userFolders.hasNext()) {
      return userFolders.next(); // 既存フォルダを返す
    } else {
      return rootFolder.createFolder(userFolderName);
    }
    
  } catch (e) {
    console.error('createUserFolder エラー: ' + e.message);
    return null; // フォルダ作成に失敗してもnullを返して処理を継続
  }
}

/**
 * ハイライト切り替え処理
 */
function processHighlightToggle(spreadsheetId, sheetName, rowIndex) {
  try {
    var service = getSheetsService();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    var highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
    
    if (highlightColumnIndex === undefined) {
      throw new Error('ハイライト列が見つかりません');
    }
    
    // 現在の値を取得
    var range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var isHighlighted = false;
    if (response && response.valueRanges && response.valueRanges[0] && 
        response.valueRanges[0].values && response.valueRanges[0].values[0] &&
        response.valueRanges[0].values[0][0]) {
      isHighlighted = response.valueRanges[0].values[0][0] === 'true';
    }
    
    // 値を切り替え
    var newValue = isHighlighted ? 'false' : 'true';
    updateSheetsData(service, spreadsheetId, range, [[newValue]]);
    
    return {
      status: 'success',
      highlighted: !isHighlighted,
      message: isHighlighted ? 'ハイライトを解除しました' : 'ハイライトしました'
    };
  } catch (e) {
    console.error('ハイライト処理エラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// =================================================================
// 互換性関数（後方互換性のため）
// =================================================================

function getWebAppUrl() {
  return getWebAppUrlCached();
}


function getHeaderIndices(spreadsheetId, sheetName) {
  debugLog('getHeaderIndices received in core.gs: spreadsheetId=%s, sheetName=%s', spreadsheetId, sheetName);
  return getHeadersCached(spreadsheetId, sheetName);
}

function clearAllCaches() {
  cacheManager.clearExpired();
  clearServiceAccountTokenCache();
}

/**
 * GoogleフォームURLからフォームIDを抽出
 */
function extractFormIdFromUrl(url) {
  if (!url) return '';
  
  try {
    // Regular expression to extract form ID from Google Forms URLs
    var formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }
    
    // Alternative pattern for e/ URLs  
    var eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (eFormIdMatch && eFormIdMatch[1]) {
      return eFormIdMatch[1];
    }
    
    return '';
  } catch (e) {
    console.warn('フォームID抽出エラー: ' + e.message);
    return '';
  }
}

// =================================================================
// リアクション処理関数
// =================================================================

/**
 * リアクション処理
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, reactingUserEmail) {
  try {
    // LockServiceを使って競合を防ぐ
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      var service = getSheetsService();
      var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
      
      // すべてのリアクション列を取得してユーザーの重複リアクションをチェック
      var allReactionRanges = [];
      var allReactionColumns = {};
      var targetReactionColumnIndex = null;
      
      // 全リアクション列の情報を準備
      REACTION_KEYS.forEach(function(key) {
        var columnName = COLUMN_HEADERS[key];
        var columnIndex = headerIndices[columnName];
        if (columnIndex !== undefined) {
          var range = "'" + sheetName + "'!" + String.fromCharCode(65 + columnIndex) + rowIndex;
          allReactionRanges.push(range);
          allReactionColumns[key] = {
            columnIndex: columnIndex,
            range: range
          };
          if (key === reactionKey) {
            targetReactionColumnIndex = columnIndex;
          }
        }
      });
      
      if (targetReactionColumnIndex === null) {
        throw new Error('対象リアクション列が見つかりません: ' + reactionKey);
      }
      
      // 全リアクション列の現在の値を一括取得
      var response = batchGetSheetsData(service, spreadsheetId, allReactionRanges);
      var updateData = [];
      var userAction = null;
      var targetCount = 0;
      
      // 各リアクション列を処理
      var rangeIndex = 0;
      REACTION_KEYS.forEach(function(key) {
        if (!allReactionColumns[key]) return;
        
        var currentReactionString = '';
        if (response && response.valueRanges && response.valueRanges[rangeIndex] && 
            response.valueRanges[rangeIndex].values && response.valueRanges[rangeIndex].values[0] &&
            response.valueRanges[rangeIndex].values[0][0]) {
          currentReactionString = response.valueRanges[rangeIndex].values[0][0];
        }
        
        var currentReactions = parseReactionString(currentReactionString);
        var userIndex = currentReactions.indexOf(reactingUserEmail);
        
        if (key === reactionKey) {
          // 対象リアクション列の処理
          if (userIndex >= 0) {
            // 既にリアクション済み → 削除（トグル）
            currentReactions.splice(userIndex, 1);
            userAction = 'removed';
          } else {
            // 未リアクション → 追加
            currentReactions.push(reactingUserEmail);
            userAction = 'added';
          }
          targetCount = currentReactions.length;
        } else {
          // 他のリアクション列からユーザーを削除（1人1リアクション制限）
          if (userIndex >= 0) {
            currentReactions.splice(userIndex, 1);
            debugLog('他のリアクションから削除: ' + reactingUserEmail + ' from ' + key);
          }
        }
        
        // 更新データを準備
        var updatedReactionString = currentReactions.join(', ');
        updateData.push({
          range: allReactionColumns[key].range,
          values: [[updatedReactionString]]
        });
        
        rangeIndex++;
      });
      
      // すべての更新を一括実行
      if (updateData.length > 0) {
        batchUpdateSheetsData(service, spreadsheetId, updateData);
      }
      
      debugLog('リアクション切り替え完了: ' + reactingUserEmail + ' → ' + reactionKey + ' (' + userAction + ')', {
        updatedRanges: updateData.length,
        targetCount: targetCount,
        allColumns: Object.keys(allReactionColumns)
      });
      
      return { 
        status: 'success', 
        message: 'リアクションを更新しました。',
        action: userAction,
        count: targetCount
      };
      
    } finally {
      lock.releaseLock();
    }
    
  } catch (e) {
    console.error('リアクション処理エラー: ' + e.message);
    return { 
      status: 'error', 
      message: 'リアクションの処理に失敗しました: ' + e.message 
    };
  }
}

// =================================================================
// フォーム作成機能
// =================================================================

/**
 * フォーム作成ファクトリ
 * @param {Object} options - フォーム作成オプション
 * @param {string} options.userEmail - ユーザーのメールアドレス
 * @param {string} options.userId - ユーザーID
 * @param {string} options.questions - 質問タイプ（'default'など）
 * @param {string} options.formDescription - フォームの説明
 * @returns {Object} フォーム作成結果
 */
function createFormFactory(options) {
  try {
    var userEmail = options.userEmail;
    var userId = options.userId;
    var formDescription = options.formDescription || 'みんなの回答ボードへの投稿フォームです。';
    
    // タイムスタンプ生成
    var now = new Date();
    var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm:ss');
    
    // フォームタイトル生成
    var formTitle = options.formTitle || ('📝 みんなの回答ボード - ' + userEmail + ' - ' + dateTimeString);
    
    // フォーム作成
    var form = FormApp.create(formTitle);
    form.setDescription(formDescription);
    
    // 基本的な質問を追加
    addUnifiedQuestions(form, options.questions || 'default', options.customConfig || {});
    
    // スプレッドシート作成
    var spreadsheetResult = createLinkedSpreadsheet(userEmail, form, dateTimeString);
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheetResult.spreadsheetId,
      spreadsheetUrl: spreadsheetResult.spreadsheetUrl,
      sheetName: spreadsheetResult.sheetName
    };
    
  } catch (error) {
    console.error('createFormFactory エラー:', error.message);
    throw new Error('フォーム作成ファクトリでエラーが発生しました: ' + error.message);
  }
}

/**
 * 統一された質問をフォームに追加
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 */
function addUnifiedQuestions(form, questionType, customConfig) {
  try {
    var config = getQuestionConfig(questionType, customConfig);

    form.setCollectEmail(true);

    if (questionType === 'simple') {
      var classItem = form.addListItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setChoiceValues(config.classQuestion.choices);
      classItem.setRequired(true);

      var nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setRequired(true);

      var mainItem = form.addTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setRequired(true);

      var reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      var validation = FormApp.createParagraphTextValidation()
        .requireTextLengthLessThanOrEqualTo(140)
        .build();
      reasonItem.setValidation(validation);
      reasonItem.setRequired(false);
    } else if (questionType === 'custom' && customConfig) {
      console.log('addUnifiedQuestions - custom mode with config:', JSON.stringify(customConfig));
      
      // クラス選択肢（有効な場合のみ）
      if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && customConfig.classQuestion.choices.length > 0) {
        var classItem = form.addListItem();
        classItem.setTitle('クラス');
        classItem.setChoiceValues(customConfig.classQuestion.choices);
        classItem.setRequired(true);
      }

      // 名前欄（常にオン）
      var nameItem = form.addTextItem();
      nameItem.setTitle('名前');
      nameItem.setRequired(false);

      // メイン質問
      var mainQuestionTitle = customConfig.mainQuestion ? customConfig.mainQuestion.title : '今回のテーマについて、あなたの考えや意見を聞かせてください';
      var mainItem;
      var questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';
      
      switch(questionType) {
        case 'text':
          mainItem = form.addTextItem();
          break;
        case 'multiple':
          mainItem = form.addCheckboxItem();
          if (customConfig.mainQuestion && customConfig.mainQuestion.choices && customConfig.mainQuestion.choices.length > 0) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (customConfig.mainQuestion && customConfig.mainQuestion.includeOthers && typeof mainItem.showOtherOption === 'function') {
            mainItem.showOtherOption(true);
          }
          break;
        case 'choice':
          mainItem = form.addMultipleChoiceItem();
          if (customConfig.mainQuestion && customConfig.mainQuestion.choices && customConfig.mainQuestion.choices.length > 0) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (customConfig.mainQuestion && customConfig.mainQuestion.includeOthers && typeof mainItem.showOtherOption === 'function') {
            mainItem.showOtherOption(true);
          }
          break;
        default:
          mainItem = form.addParagraphTextItem();
      }
      mainItem.setTitle(mainQuestionTitle);
      mainItem.setRequired(true);

      // 理由欄（常にオン）
      var reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle('そう考える理由や体験があれば教えてください（任意）');
      reasonItem.setRequired(false);
    } else {
      var classItem = form.addTextItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setHelpText(config.classQuestion.helpText);
      classItem.setRequired(true);

      var mainItem = form.addParagraphTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setHelpText(config.mainQuestion.helpText);
      mainItem.setRequired(true);

      var reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      reasonItem.setRequired(false);

      var nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setHelpText(config.nameQuestion.helpText);
      nameItem.setRequired(false);
    }

    debugLog('フォームに統一質問を追加しました: ' + questionType);
    
  } catch (error) {
    console.error('addUnifiedQuestions エラー:', error.message);
    throw error;
  }
}

/**
 * 質問設定を取得
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 * @returns {Object} 質問設定
 */
function getQuestionConfig(questionType, customConfig) {
  // 統一されたテンプレート設定（simple のみ使用）
  var config = {
    classQuestion: {
      title: 'クラス',
      helpText: '',
      choices: ['クラス1', 'クラス2', 'クラス3', 'クラス4']
    },
    nameQuestion: {
      title: '名前',
      helpText: ''
    },
    mainQuestion: {
      title: '今日のテーマについて、あなたの考えや意見を聞かせてください',
      helpText: '',
      choices: ['気づいたことがある。', '疑問に思うことがある。', 'もっと知りたいことがある。'],
      type: 'paragraph'
    },
    reasonQuestion: {
      title: 'そう考える理由や体験があれば教えてください（任意）',
      helpText: '',
      type: 'paragraph'
    }
  };
  
  // カスタム設定をマージ
  if (customConfig && typeof customConfig === 'object') {
    for (var key in customConfig) {
      if (config[key]) {
        Object.assign(config[key], customConfig[key]);
      }
    }
  }
  
  return config;
}

/**
 * 質問テンプレート取得用 API
 * @returns {ContentService.TextOutput} JSON形式の質問設定
 */
function doGetQuestionConfig() {
  try {
    // 現在の日時を取得してタイトルに含める
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    
    const cfg = getQuestionConfig('simple');
    
    // タイトルにタイムスタンプを追加
    cfg.formTitle = `フォーム作成 - ${timestamp}`;
    
    return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('doGetQuestionConfig error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Failed to get question config',
      details: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * クラス選択肢をデータベースに保存
 */
function saveClassChoices(userId, classChoices) {
  try {
    var currentUserId = userId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.savedClassChoices = classChoices;
    configJson.lastClassChoicesUpdate = new Date().toISOString();
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    debugLog('クラス選択肢が保存されました:', classChoices);
    return { status: 'success', message: 'クラス選択肢が保存されました' };
  } catch (error) {
    console.error('クラス選択肢保存エラー:', error.message);
    return { status: 'error', message: 'クラス選択肢の保存に失敗しました: ' + error.message };
  }
}

/**
 * 保存されたクラス選択肢を取得
 */
function getSavedClassChoices(userId) {
  try {
    var currentUserId = userId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var savedClassChoices = configJson.savedClassChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4'];
    
    return { 
      status: 'success', 
      classChoices: savedClassChoices,
      lastUpdate: configJson.lastClassChoicesUpdate
    };
  } catch (error) {
    console.error('クラス選択肢取得エラー:', error.message);
    return { status: 'error', message: 'クラス選択肢の取得に失敗しました: ' + error.message };
  }
}

/**
 * クイックスタート用フォーム作成（デフォルト設定）
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} userId - ユーザーID
 * @returns {Object} フォーム作成結果
 */
function createQuickStartForm(userEmail, userId) {
  try {
    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm:ss');
    const formTitle = `みんなの回答ボード - ${dateTimeString}`;
    
    const defaultConfig = {
      opinionHeader: '今日の学習について、あなたの考えや感想を聞かせてください',  // mainQuestion を opinionHeader に統一
      questionType: 'text',
      enableClass: false,
      includeOthers: false
    };
    
    return createFormFactory({
      userEmail: userEmail,
      userId: userId,
      formTitle: formTitle,
      questions: 'custom',
      formDescription: 'このフォームに回答すると、みんなの回答ボードに反映されます。',
      customConfig: defaultConfig
    });
  } catch (error) {
    console.error('createQuickStartForm Error:', error.message);
    throw new Error('クイックスタートフォームの作成に失敗しました: ' + error.message);
  }
}

/**
 * カスタムフォーム作成（管理パネル用）
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} userId - ユーザーID
 * @param {Object} config - カスタム設定
 * @returns {Object} フォーム作成結果
 */
function createCustomForm(userEmail, userId, config) {
  try {
    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm:ss');
    const baseTitle = config.formTitle || 'カスタムフォーム';
    const formTitle = `${baseTitle} - ${dateTimeString}`;
    
    // AdminPanelのconfig構造を内部形式に変換 (opinionHeaderに統一)
    const convertedConfig = {
      mainQuestion: {
        title: config.opinionHeader || '今日の学習について、あなたの考えや感想を聞かせてください',  // mainQuestion を opinionHeader に統一
        type: config.questionType || 'text',
        choices: config.choices || [],
        includeOthers: config.includeOthers || false
      },
      enableClass: config.enableClass || false,
      classQuestion: {
        choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4']
      }
    };
    
    console.log('createCustomForm - converted config:', JSON.stringify(convertedConfig));
    
    return createFormFactory({
      userEmail: userEmail,
      userId: userId,
      formTitle: formTitle,
      questions: 'custom',
      formDescription: 'このフォームに回答すると、みんなの回答ボードに反映されます。',
      customConfig: convertedConfig
    });
  } catch (error) {
    console.error('createCustomForm Error:', error.message);
    throw new Error('カスタムフォームの作成に失敗しました: ' + error.message);
  }
}

/**
 * カスタム設定でStudyQuestフォームを作成（廃止予定）
 * @deprecated createCustomFormを使用してください
 */
function createStudyQuestFormWithConfig(userEmail, userId, formTitle, config) {
  try {
    var formResult = createFormFactory({
      userEmail: userEmail,
      userId: userId,
      formTitle: formTitle,
      questions: 'custom',
      formDescription: 'このフォームに回答すると、みんなの回答ボードに反映されます。',
      customConfig: config
    });
    
    var form = FormApp.openById(formResult.formId);
    
    // Email収集タイプの設定
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      }
    } catch (undocumentedError) {
      console.warn('Email collection type setting failed:', undocumentedError.message);
    }
    
    return formResult;
  } catch (error) {
    console.error('createStudyQuestFormWithConfig Error:', error.message);
    throw new Error('カスタムフォームの作成に失敗しました: ' + error.message);
  }
}

/**
 * リンクされたスプレッドシートを作成
 * @param {string} userEmail - ユーザーメール
 * @param {GoogleAppsScript.Forms.Form} form - フォーム
 * @param {string} dateTimeString - 日時文字列
 * @returns {Object} スプレッドシート情報
 */
function createLinkedSpreadsheet(userEmail, form, dateTimeString) {
  try {
    // スプレッドシート名を設定
    var spreadsheetName = userEmail + ' - 回答データ - ' + dateTimeString;
    
    // 新しいスプレッドシートを作成
    var spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
    var spreadsheetId = spreadsheetObj.getId();
    
    // スプレッドシートの共有設定を同一ドメイン閲覧可能に設定
    try {
      var file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }
      
      // 同一ドメインで閲覧可能に設定（教育機関対応）
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      debugLog('スプレッドシートを同一ドメイン閲覧可能に設定しました: ' + spreadsheetId);
      
      // 作成者（現在のユーザー）は所有者として保持
      debugLog('作成者は所有者として権限を保持: ' + userEmail);
      
    } catch (sharingError) {
      console.warn('共有設定の変更に失敗しましたが、処理を続行します: ' + sharingError.message);
    }
    
    // フォームの回答先としてスプレッドシートを設定
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
    
    // スプレッドシート名を再設定（フォーム設定後に名前が変わる場合があるため）
    spreadsheetObj.rename(spreadsheetName);
    
    // シート名を取得（通常は「フォームの回答 1」）
    var sheets = spreadsheetObj.getSheets();
    var sheetName = sheets[0].getName();
    
    // サービスアカウントとスプレッドシートを共有
    try {
      shareSpreadsheetWithServiceAccount(spreadsheetId);
      debugLog('サービスアカウントとの共有完了: ' + spreadsheetId);
    } catch (shareError) {
      console.error('サービスアカウント共有エラー:', shareError.message);
      console.error('スプレッドシート作成は完了しましたが、サービスアカウントとの共有に失敗しました。手動で共有してください。');
    }
    
    // リアクション列をスプレッドシートに追加
    try {
      addReactionColumnsToSpreadsheet(spreadsheetId, sheetName);
      debugLog('リアクション列の追加完了: ' + spreadsheetId + ', シート: ' + sheetName);
    } catch (reactionError) {
      console.warn('リアクション列の追加に失敗しましたが、処理を続行します:', reactionError.message);
    }
    
    return {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetObj.getUrl(),
      sheetName: sheetName
    };
    
  } catch (error) {
    console.error('createLinkedSpreadsheet エラー:', error.message);
    throw new Error('スプレッドシート作成に失敗しました: ' + error.message);
  }
}

/**
 * スプレッドシートをサービスアカウントと共有
 * @param {string} spreadsheetId - スプレッドシートID
 */
function shareSpreadsheetWithServiceAccount(spreadsheetId) {
  try {
    var serviceAccountEmail = getServiceAccountEmail();
    
    if (!serviceAccountEmail || serviceAccountEmail === 'サービスアカウント未設定' || serviceAccountEmail === 'サービスアカウント設定エラー') {
      throw new Error('サービスアカウントのメールアドレスが取得できません: ' + serviceAccountEmail);
    }
    
    debugLog('サービスアカウント共有開始:', serviceAccountEmail, 'スプレッドシート:', spreadsheetId);
    
    // DriveAppを使用してスプレッドシートをサービスアカウントと共有
    try {
      var file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }
      file.addEditor(serviceAccountEmail);
    } catch (driveError) {
      console.error('DriveApp error:', driveError.message);
      throw new Error('Drive API操作に失敗しました: ' + driveError.message);
    }
    
    debugLog('サービスアカウント共有成功:', serviceAccountEmail);
    
  } catch (error) {
    console.error('shareSpreadsheetWithServiceAccount エラー:', error.message);
    throw new Error('サービスアカウントとの共有に失敗しました: ' + error.message);
  }
}

/**
 * すべての既存スプレッドシートをサービスアカウントと共有
 * @returns {Object} 共有結果
 */
function shareAllSpreadsheetsWithServiceAccount() {
  try {
    debugLog('全スプレッドシートのサービスアカウント共有開始');
    
    var allUsers = getAllUsers();
    var results = [];
    var successCount = 0;
    var errorCount = 0;
    
    for (var i = 0; i < allUsers.length; i++) {
      var user = allUsers[i];
      if (user.spreadsheetId && user.isActive === 'true') {
        try {
          shareSpreadsheetWithServiceAccount(user.spreadsheetId);
          results.push({
            userId: user.userId,
            adminEmail: user.adminEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'success'
          });
          successCount++;
          debugLog('共有成功:', user.adminEmail, user.spreadsheetId);
        } catch (shareError) {
          results.push({
            userId: user.userId,
            adminEmail: user.adminEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'error',
            error: shareError.message
          });
          errorCount++;
          console.error('共有失敗:', user.adminEmail, shareError.message);
        }
      }
    }
    
    debugLog('全スプレッドシート共有完了:', successCount + '件成功', errorCount + '件失敗');
    
    return {
      status: 'completed',
      totalUsers: allUsers.length,
      successCount: successCount,
      errorCount: errorCount,
      results: results
    };
    
  } catch (error) {
    console.error('shareAllSpreadsheetsWithServiceAccount エラー:', error.message);
    throw new Error('全スプレッドシート共有処理でエラーが発生しました: ' + error.message);
  }
}

/**
 * フォーム作成
 */
// questionType defaults to 'simple'
function createStudyQuestForm(userEmail, userId, formTitle, questionType) {
  try {
    // パフォーマンス測定開始
    var profiler = (typeof globalProfiler !== 'undefined') ? globalProfiler : {
      start: function() {},
      end: function() {}
    };
    profiler.start('createForm');
    
    // 共通ファクトリを使用してフォーム作成
    var formResult = createFormFactory({
      userEmail: userEmail,
      userId: userId,
      formTitle: formTitle,
      questions: questionType || 'simple',
      formDescription: 'このフォームに回答すると、みんなの回答ボードに反映されます。'
    });
    
    // カスタマイズされた設定を追加
    var form = FormApp.openById(formResult.formId);
    
    // Email収集タイプの設定（可能な場合）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      }
    } catch (undocumentedError) {
      // 失敗しても続行
      console.warn('Email collection type setting failed:', undocumentedError.message);
    }
    
    // 確認メッセージの設定（回答ボードURLを含む）
    var appUrls = generateAppUrls(userId);
    var boardUrl = appUrls.viewUrl || (appUrls.webAppUrl + '?userId=' + encodeURIComponent(userId || ''));
    
    var confirmationMessage = 'ご回答ありがとうございます！ボードはこちら: ' + boardUrl;

    if (form.getPublishedUrl()) {
      confirmationMessage += '\n\n回答の編集はこちら: ' + form.getPublishedUrl();
    }
    
    form.setConfirmationMessage(confirmationMessage);
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(formResult.spreadsheetId);
    
    // リアクション列をスプレッドシートに追加
    addReactionColumnsToSpreadsheet(formResult.spreadsheetId, formResult.sheetName);
    
    profiler.end('createForm');
    return formResult;
    
  } catch (e) {
    console.error('createStudyQuestFormエラー: ' + e.message);
    throw new Error('フォームの作成に失敗しました: ' + e.message);
  }
}

/**
 * サービスアカウントをスプレッドシートに追加
 */
function addServiceAccountToSpreadsheet(spreadsheetId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;
    
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // サービスアカウントを編集者として追加
    if (serviceAccountEmail) {
      spreadsheet.addEditor(serviceAccountEmail);
      debugLog('サービスアカウント (' + serviceAccountEmail + ') をスプレッドシートの編集者として追加しました。');
      
      // セッション管理：サービスアカウントアクセス権限の記録
      try {
        const sessionData = {
          serviceAccountEmail: serviceAccountEmail,
          spreadsheetId: spreadsheetId,
          accessGranted: new Date().toISOString(),
          accessType: 'service_account_editor',
          securityLevel: 'domain_view'
        };
        debugLog('サービスアカウントアクセス権限を記録しました:', JSON.stringify(sessionData));
      } catch (sessionLogError) {
        console.warn('セッション記録でエラー:', sessionLogError.message);
      }
    }
    
    // 同一ドメインユーザーは共有設定により閲覧可能
    debugLog('同一ドメインユーザーは共有設定により閲覧可能です');
    
  } catch (e) {
    console.error('サービスアカウントの追加に失敗: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * 既存ユーザーのスプレッドシートアクセス権限を修復
 * ユーザーがスプレッドシートにアクセスできない場合に使用
 * @param {string} userEmail - ユーザーのメールアドレス
 * @param {string} spreadsheetId - スプレッドシートID
 */
function repairUserSpreadsheetAccess(userEmail, spreadsheetId) {
  try {
    debugLog('スプレッドシートアクセス権限の修復を開始: ' + userEmail + ' -> ' + spreadsheetId);
    
    // DriveApp経由で共有設定を変更
    var file;
    try {
      file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }
    } catch (driveError) {
      console.error('DriveApp.getFileById error:', driveError.message);
      throw new Error('スプレッドシートへのアクセスに失敗しました: ' + driveError.message);
    }
    
    // ドメイン全体でアクセス可能に設定
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
      debugLog('スプレッドシートをドメイン全体で編集可能に設定しました');
    } catch (domainSharingError) {
      console.warn('ドメイン共有設定に失敗: ' + domainSharingError.message);
      
      // ドメイン共有に失敗した場合は個別にユーザーを追加
      try {
        file.addEditor(userEmail);
        debugLog('ユーザーを個別に編集者として追加しました: ' + userEmail);
      } catch (individualError) {
        console.error('個別ユーザー追加も失敗: ' + individualError.message);
      }
    }
    
    // SpreadsheetApp経由でも編集者として追加
    try {
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      spreadsheet.addEditor(userEmail);
      debugLog('SpreadsheetApp経由でユーザーを編集者として追加: ' + userEmail);
    } catch (spreadsheetAddError) {
      console.warn('SpreadsheetApp経由の追加で警告: ' + spreadsheetAddError.message);
    }
    
    // サービスアカウントも確認
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;
    
    if (serviceAccountEmail) {
      try {
        file.addEditor(serviceAccountEmail);
        debugLog('サービスアカウントも編集者として追加: ' + serviceAccountEmail);
      } catch (serviceError) {
        console.warn('サービスアカウント追加で警告: ' + serviceError.message);
      }
    }
    
    return {
      success: true,
      message: 'スプレッドシートアクセス権限を修復しました。ドメイン全体でアクセス可能です。'
    };
    
  } catch (e) {
    console.error('スプレッドシートアクセス権限の修復に失敗: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 管理パネル専用：権限問題の緊急修復
 * @param {string} userEmail - ユーザーメール
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} 修復結果
 */
function emergencyAdminPanelRepair(userEmail, spreadsheetId) {
  try {
    debugLog('緊急修復開始: 管理パネルアクセス用');
    
    // 1. サービスアカウント権限の強制追加
    addServiceAccountToSpreadsheet(spreadsheetId);
    debugLog('ステップ1: サービスアカウント権限追加完了');
    
    // 2. ユーザー権限の強制追加
    const repairResult = repairUserSpreadsheetAccess(userEmail, spreadsheetId);
    debugLog('ステップ2: ユーザー権限修復結果:', repairResult);
    
    // 3. 権限確認テスト
    try {
      const testAccess = SpreadsheetApp.openById(spreadsheetId);
      testAccess.getName();
      debugLog('ステップ3: 権限確認テスト成功');
    } catch (testError) {
      console.warn('ステップ3: 権限確認テスト失敗:', testError.message);
    }
    
    // 4. サービスアカウントアクセステスト
    try {
      const service = getSheetsService();
      const testData = getSpreadsheetsData(service, spreadsheetId);
      debugLog('ステップ4: サービスアカウントアクセステスト成功');
    } catch (serviceTestError) {
      console.warn('ステップ4: サービスアカウントアクセステスト失敗:', serviceTestError.message);
    }
    
    return {
      success: true,
      message: '管理パネルアクセス権限を緊急修復しました',
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('緊急修復エラー:', error.message);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * スプレッドシートにリアクション列を追加
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
    
    var additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    // 効率的にヘッダー情報を取得
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var startCol = currentHeaders.length + 1;
    
    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
    
    // スタイリングを一括適用
    var allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
    
    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      console.warn('Auto-resize failed:', resizeError.message);
    }
    
    debugLog('リアクション列を追加しました: ' + sheetName);
  }
  catch (e) {
    console.error('リアクション列追加エラー: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートの共有設定をチェックする
 * @param {string} spreadsheetId - チェックするスプレッドシートのID
 * @returns {object} status ('success' or 'error') と message
 */
function checkSpreadsheetSharingPermission(spreadsheetId) {
  try {
    var service = getDriveService(); // Drive APIサービスを取得
    var file = service.files.get({
      fileId: spreadsheetId,
      fields: 'permissions'
    });

    var isRestricted = true;
    var publicAccessFound = false;

    if (file.permissions) {
      for (var i = 0; i < file.permissions.length; i++) {
        var permission = file.permissions[i];
        // typeが'anyone'または'domain'でroleが'reader'以上の場合は公開状態
        if (permission.type === 'anyone' || (permission.type === 'domain' && permission.role !== 'reader')) {
          publicAccessFound = true;
          isRestricted = false;
          break;
        }
      }
    }

    if (publicAccessFound) {
      return {
        status: 'error',
        message: 'スプレッドシートの共有設定が「制限付き」になっていません。セキュリティのため、「制限付き」に設定してください。'
      };
    } else {
      return {
        status: 'success',
        message: 'スプレッドシートの共有設定は「制限付き」です。'
      };
    }
  } catch (e) {
    console.error('スプレッドシート共有設定チェックエラー: ' + e.message);
    return {
      status: 'error',
      message: 'スプレッドシートの共有設定の確認中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * シートデータ取得
 */
function getSheetData(userId, sheetName, classFilter, sortMode, adminMode) {
  // キャッシュキー生成（ユーザー、シート、フィルタ条件ごとに個別キャッシュ）
  var cacheKey = `sheetData_${userId}_${sheetName}_${classFilter}_${sortMode}`;
  
  // 管理モードの場合はキャッシュをバイパス（最新データを取得）
  if (adminMode === true) {
    debugLog('🔄 管理モード：シートデータキャッシュをバイパス');
    return executeGetSheetData(userId, sheetName, classFilter, sortMode);
  }
  
  return cacheManager.get(cacheKey, () => {
    return executeGetSheetData(userId, sheetName, classFilter, sortMode);
  }, { ttl: 300 }); // 5分間キャッシュ
}

/**
 * 実際のシートデータ取得処理（キャッシュ制御から分離）
 */
function executeGetSheetData(userId, sheetName, classFilter, sortMode) {
    try {
      var userInfo = getCachedUserInfo(userId);
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }
      
      var spreadsheetId = userInfo.spreadsheetId;
      var service = getSheetsService();
      
      // フォーム回答データのみを取得（名簿機能は使用しない）
      var ranges = [sheetName + '!A:Z'];
      
      var responses = batchGetSheetsData(service, spreadsheetId, ranges);
      var sheetData = responses.valueRanges[0].values || [];
    
    // 名簿機能は使用せず、空の配列を設定
    var rosterData = [];
    
    
    if (sheetData.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0
      };
    }
    
    var headers = sheetData[0];
    var dataRows = sheetData.slice(1);
    
    // ヘッダーインデックスを取得（キャッシュ利用）
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    
    // 名簿マップを作成（キャッシュ利用）
    var rosterMap = buildRosterMap(rosterData);
    
    // 表示モードを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
    
    // Check if current user is the board owner
    var isOwner = (configJson.ownerId === userId);
    debugLog('getSheetData: isOwner=%s, ownerId=%s, userId=%s', isOwner, configJson.ownerId, userId);
    
    // データを処理
    var processedData = dataRows.map(function(row, index) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2, isOwner);
    });
    
    // フィルタリング
    var filteredData = processedData;
    if (classFilter && classFilter !== 'すべて') {
      var classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        filteredData = processedData.filter(function(row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }
    
    // ソート適用
    var sortedData = applySortMode(filteredData, sortMode || 'newest');
    
    return {
      status: 'success',
      data: sortedData,
      headers: headers,
      totalCount: sortedData.length,
      displayMode: displayMode
    };
    
  } catch (e) {
    console.error('シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      headers: []
    };
  }
}

/**
 * シート一覧取得
 */
function getSheetsList(userId) {
  try {
    debugLog('getSheetsList: Start for userId:', userId);
    
    // キャッシュキーを生成
    const cacheKey = `sheets_list_${userId}`;
    
    // キャッシュの使用を試行、失敗時は直接実行
    try {
      if (typeof cacheManager !== 'undefined' && cacheManager) {
        debugLog('getSheetsList: Using cacheManager');
        const result = cacheManager.get(cacheKey, () => {
          debugLog('getSheetsList: Cache miss, executing getSheetsListInternal');
          return getSheetsListInternal(userId);
        }, { ttl: 1800 }); // 30分間キャッシュ
        debugLog('getSheetsList: cacheManager.get returned:', result);
        
        if (result === undefined || result === null) {
          console.warn('getSheetsList: cacheManager returned undefined/null, falling back to direct execution');
          return getSheetsListInternal(userId);
        }
        return result;
      } else {
        debugLog('getSheetsList: cacheManager not available, executing directly');
        return getSheetsListInternal(userId);
      }
    } catch (cacheError) {
      console.warn('getSheetsList: Cache error, falling back to direct execution:', cacheError.message);
      console.warn('getSheetsList: Cache error details:', cacheError.stack);
      return getSheetsListInternal(userId);
    }
    
  } catch (e) {
    console.error('getSheetsList: シート一覧取得エラー:', e.message);
    return [];
  }
}

/**
 * 内部シート一覧取得関数
 */
function getSheetsListInternal(userId) {
  debugLog('getSheetsListInternal: Start for userId:', userId);
  
  var userInfo = findUserById(userId);
  if (!userInfo) {
    console.warn('getSheetsListInternal: User not found:', userId);
    return [];
  }
  
  if (!userInfo.spreadsheetId) {
    console.warn('getSheetsListInternal: No spreadsheet ID for user:', userId);
    return [];
  }
  
  var service = getSheetsService();
  var spreadsheet;
  
  try {
    spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
  } catch (accessError) {
    console.warn('getSheetsListInternal: アクセスエラーを検出。サービスアカウント権限を修復中...', accessError.message);
    
    // サービスアカウントの権限修復を試行
    try {
      addServiceAccountToSpreadsheet(userInfo.spreadsheetId);
      debugLog('getSheetsListInternal: サービスアカウント権限を追加しました。再試行中...');
      
      // 少し待ってから再試行
      Utilities.sleep(1000);
      spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
      
    } catch (repairError) {
      console.error('getSheetsListInternal: 権限修復に失敗:', repairError.message);
      return {
        error: true,
        message: 'スプレッドシートへのアクセス権限がありません。管理者にお問い合わせください。',
        details: accessError.message
      };
    }
  }
  
  if (!spreadsheet || !spreadsheet.sheets || !Array.isArray(spreadsheet.sheets)) {
    console.error('getSheetsListInternal: Invalid spreadsheet data');
    return [];
  }
  
  var sheets = spreadsheet.sheets.map(function(sheet) {
    if (!sheet.properties) {
      console.warn('getSheetsListInternal: Sheet missing properties:', sheet);
      return null;
    }
    return {
      name: sheet.properties.title,
      id: sheet.properties.sheetId
    };
  }).filter(function(sheet) { return sheet !== null; });
  
  debugLog('getSheetsListInternal: Successfully returning', sheets.length, 'sheets:', sheets);
  return sheets;
}

/**
 * 名簿マップを構築（名簿機能無効化のため空のマップを返す）
 * @param {array} rosterData - 名簿データ（使用されません）
 * @returns {object} 空の名簿マップ
 */
function buildRosterMap(rosterData) {
  // 名簿機能は使用しないため、常に空のマップを返す
  // 名前はフォーム入力から直接取得
  return {};
}

/**
 * 行データを処理（スコア計算、名前変換など）
 */
function processRowData(row, headers, headerIndices, rosterMap, displayMode, rowNumber, isOwner) {
  var processedRow = {
    rowNumber: rowNumber,
    originalData: row,
    score: 0,
    likeCount: 0,
    understandCount: 0,
    curiousCount: 0,
    isHighlighted: false
  };
  
  // リアクションカウント計算
  REACTION_KEYS.forEach(function(reactionKey) {
    var columnName = COLUMN_HEADERS[reactionKey];
    var columnIndex = headerIndices[columnName];
    
    if (columnIndex !== undefined && row[columnIndex]) {
      var reactions = parseReactionString(row[columnIndex]);
      var count = reactions.length;
      
      switch (reactionKey) {
        case 'LIKE':
          processedRow.likeCount = count;
          break;
        case 'UNDERSTAND':
          processedRow.understandCount = count;
          break;
        case 'CURIOUS':
          processedRow.curiousCount = count;
          break;
      }
    }
  });
  
  // ハイライト状態チェック
  var highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }
  
  // スコア計算
  processedRow.score = calculateRowScore(processedRow);
  
  // 名前の表示処理（フォーム入力の名前を使用）
  var nameIndex = headerIndices[COLUMN_HEADERS.NAME];
  if (nameIndex !== undefined && row[nameIndex] && (displayMode === DISPLAY_MODES.NAMED || isOwner)) {
    processedRow.displayName = row[nameIndex];
  } else if (displayMode === DISPLAY_MODES.NAMED || isOwner) {
    // 名前入力がない場合のフォールバック
    processedRow.displayName = '匿名';
  }
  
  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScore(rowData) {
  var baseScore = 1.0;
  
  // いいね！による加算
  var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
  
  // その他のリアクションも軽微な加算
  var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;
  
  // ハイライトによる大幅加算
  var highlightBonus = rowData.isHighlighted ? 0.5 : 0;
  
  // ランダム要素（同じスコアの項目をランダムに並べるため）
  var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
  
  return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
}

/**
 * データにソートを適用
 */
function applySortMode(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort(function(a, b) { return b.score - a.score; });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArray(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort(function(a, b) { return b.likeCount - a.likeCount; });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

/**
 * リアクション文字列をパース
 */
function parseReactionString(val) {
  if (!val) return [];
  return val.toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

/**
 * ヘルパー関数：ヘッダー配列から指定した名前のインデックスを取得
 * COLUMN_HEADERSと統一された方式を使用
 */
function getHeaderIndex(headers, headerName) {
  if (!headers || !headerName) return -1;
  return headers.indexOf(headerName);
}

/**
 * COLUMN_HEADERSキーから適切なヘッダー名を取得
 * @param {string} columnKey - COLUMN_HEADERSのキー（例：'OPINION', 'CLASS'）
 * @returns {string} ヘッダー名
 */
function getColumnHeaderName(columnKey) {
  return COLUMN_HEADERS[columnKey] || '';
}

/**
 * 統一されたヘッダーインデックス取得関数
 * @param {array} headers - ヘッダー配列
 * @param {string} columnKey - COLUMN_HEADERSのキー
 * @returns {number} インデックス（見つからない場合は-1）
 */
function getColumnIndex(headers, columnKey) {
  var headerName = getColumnHeaderName(columnKey);
  return getHeaderIndex(headers, headerName);
}

/**
 * 設定された列名と実際のスプレッドシートヘッダーをマッピング
 * @param {Object} configHeaders - 設定されたヘッダー名
 * @param {Object} actualHeaderIndices - 実際のヘッダーインデックスマップ
 * @returns {Object} マッピングされたインデックス
 */
function mapConfigToActualHeaders(configHeaders, actualHeaderIndices) {
  var mappedIndices = {};
  var availableHeaders = Object.keys(actualHeaderIndices);
  debugLog('mapConfigToActualHeaders: Available headers in spreadsheet: %s', JSON.stringify(availableHeaders));
  
  // 各設定ヘッダーでマッピングを試行
  for (var configKey in configHeaders) {
    var configHeaderName = configHeaders[configKey];
    var mappedIndex = undefined;
    
    debugLog('mapConfigToActualHeaders: Trying to map %s = "%s"', configKey, configHeaderName);
    
    if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {
      // 完全一致
      mappedIndex = actualHeaderIndices[configHeaderName];
      debugLog('mapConfigToActualHeaders: Exact match found for %s: index %s', configKey, mappedIndex);
    } else if (configHeaderName) {
      // 部分一致で検索（大文字小文字を区別しない）
      var normalizedConfigName = configHeaderName.toLowerCase().trim();
      
      for (var actualHeader in actualHeaderIndices) {
        var normalizedActualHeader = actualHeader.toLowerCase().trim();
        if (normalizedActualHeader === normalizedConfigName) {
          mappedIndex = actualHeaderIndices[actualHeader];
          debugLog('mapConfigToActualHeaders: Case-insensitive match found for %s: "%s" -> index %s', configKey, actualHeader, mappedIndex);
          break;
        }
      }
      
      // 部分一致検索
      if (mappedIndex === undefined) {
        for (var actualHeader in actualHeaderIndices) {
          var normalizedActualHeader = actualHeader.toLowerCase().trim();
          if (normalizedActualHeader.includes(normalizedConfigName) || normalizedConfigName.includes(normalizedActualHeader)) {
            mappedIndex = actualHeaderIndices[actualHeader];
            debugLog('mapConfigToActualHeaders: Partial match found for %s: "%s" -> index %s', configKey, actualHeader, mappedIndex);
            break;
          }
        }
      }
    }
    
    // opinionHeader（メイン質問）の特別処理：見つからない場合は最も長い質問様ヘッダーを使用
    if (mappedIndex === undefined && configKey === 'opinionHeader') {
      var standardHeaders = ['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '理由', 'なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト'];
      var questionHeaders = [];
      
      for (var header in actualHeaderIndices) {
        var isStandardHeader = false;
        for (var i = 0; i < standardHeaders.length; i++) {
          if (header.toLowerCase().includes(standardHeaders[i].toLowerCase()) || 
              standardHeaders[i].toLowerCase().includes(header.toLowerCase())) {
            isStandardHeader = true;
            break;
          }
        }
        
        if (!isStandardHeader && header.length > 10) { // 質問は通常長い
          questionHeaders.push({header: header, index: actualHeaderIndices[header]});
        }
      }
      
      if (questionHeaders.length > 0) {
        // 最も長いヘッダーを選択（通常メイン質問が最も長い）
        var longestHeader = questionHeaders.reduce(function(prev, current) {
          return (prev.header.length > current.header.length) ? prev : current;
        });
        mappedIndex = longestHeader.index;
        debugLog('mapConfigToActualHeaders: Auto-detected main question header for %s: "%s" -> index %s', configKey, longestHeader.header, mappedIndex);
      }
    }
    
    mappedIndices[configKey] = mappedIndex;
    
    if (mappedIndex === undefined) {
      debugLog('mapConfigToActualHeaders: WARNING - No match found for %s = "%s"', configKey, configHeaderName);
    }
  }
  
  debugLog('mapConfigToActualHeaders: Final mapping result: %s', JSON.stringify(mappedIndices));
  return mappedIndices;
}

/**
 * 特定の行のリアクション情報を取得
 */
function getRowReactions(spreadsheetId, sheetName, rowIndex, userEmail) {
  try {
    var service = getSheetsService();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    
    var reactionData = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
    
    // 各リアクション列からデータを取得
    ['UNDERSTAND', 'LIKE', 'CURIOUS'].forEach(function(reactionKey) {
      var columnName = COLUMN_HEADERS[reactionKey];
      var columnIndex = headerIndices[columnName];
      
      if (columnIndex !== undefined) {
        var range = sheetName + '!' + String.fromCharCode(65 + columnIndex) + rowIndex;
        try {
          var response = batchGetSheetsData(service, spreadsheetId, [range]);
          var cellValue = '';
          if (response && response.valueRanges && response.valueRanges[0] && 
              response.valueRanges[0].values && response.valueRanges[0].values[0] &&
              response.valueRanges[0].values[0][0]) {
            cellValue = response.valueRanges[0].values[0][0];
          }
          
          if (cellValue) {
            var reactions = parseReactionString(cellValue);
            reactionData[reactionKey].count = reactions.length;
            reactionData[reactionKey].reacted = reactions.indexOf(userEmail) !== -1;
          }
        } catch (cellError) {
          console.warn('リアクション取得エラー(' + reactionKey + '): ' + cellError.message);
        }
      }
    });
    
    return reactionData;
  } catch (e) {
    console.error('getRowReactions エラー: ' + e.message);
    return {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
  }
}

// =================================================================
// 追加のコアファンクション（Code.gsから移行）
// =================================================================



/**
 * 軽量な件数チェック（新着通知用）
 * 実際のデータではなく件数のみを返す
 */


/**
 * 回答ボードのデータを強制的に再読み込み
 */



// =================================================================
// ユーティリティ関数
// =================================================================

/**
 * 現在のユーザーのステータス情報を取得（AppSetupPage.htmlから呼び出される）
 * @returns {object} ユーザーのステータス情報
 */

/**
 * isActive状態を更新（AppSetupPage.htmlから呼び出される）
 * @param {boolean} isActive - 新しいisActive状態
 * @returns {object} 更新結果
 */
function updateIsActiveStatus(requestUserId, isActive) {
  // 新規ユーザー（requestUserIdがundefinedまたはnull）の場合はverifyUserAccessをスキップ
  if (requestUserId) {
    verifyUserAccess(requestUserId);
  }
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return {
        status: 'error',
        message: 'ユーザーが認証されていません'
      };
    }

    // 現在のユーザー情報を取得
    var userInfo = findUserByEmail(activeUserEmail);
    if (!userInfo) {
      return {
        status: 'error',
        message: 'ユーザーがデータベースに登録されていません'
      };
    }

    // 編集者権限があるか確認（自分自身の状態変更も含む）
    if (!isTrue(userInfo.isActive)) {
      return {
        status: 'error',
        message: 'この操作を実行する権限がありません'
      };
    }

    // isActive状態を更新
    var newIsActiveValue = isActive ? 'true' : 'false';
    var updateResult = updateUser(userInfo.userId, {
      isActive: newIsActiveValue,
      lastAccessedAt: new Date().toISOString()
    });

    if (updateResult.success) {
      var statusMessage = isActive 
        ? 'アプリが正常に有効化されました' 
        : 'アプリが正常に無効化されました';
      
      return {
        status: 'success',
        message: statusMessage,
        newStatus: newIsActiveValue
      };
    } else {
      return {
        status: 'error',
        message: 'データベースの更新に失敗しました'
      };
    }
  } catch (e) {
    console.error('updateIsActiveStatus エラー: ' + e.message);
    return {
      status: 'error',
      message: 'isActive状態の更新に失敗しました: ' + e.message
    };
  }
}

/**
 * セットアップページへのアクセス権限を確認
 * @returns {boolean} アクセス権限があればtrue
 */
function hasSetupPageAccess() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return false;
    }

    // データベースに登録され、かつisActiveがtrueのユーザーのみアクセス可能
    var userInfo = findUserByEmail(activeUserEmail);
    return userInfo && isTrue(userInfo.isActive);
  } catch (e) {
    console.error('hasSetupPageAccess エラー: ' + e.message);
    return false;
  }
}

/**
 * メールアドレスからドメインを抽出
 * @param {string} email - メールアドレス
 * @returns {string} ドメイン部分
 */
function getEmailDomain(email) {
  return email.split('@')[1] || '';
}

/**
 * Drive APIサービスを取得
 * @returns {object} Drive APIサービス
 */
function getDriveService() {
  var accessToken = getServiceAccountTokenCached();
  return {
    accessToken: accessToken,
    baseUrl: 'https://www.googleapis.com/drive/v3',
    files: {
      get: function(params) {
        var url = this.baseUrl + '/files/' + params.fileId + '?fields=' + encodeURIComponent(params.fields);
        var response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + this.accessToken }
        });
        return JSON.parse(response.getContentText());
      }
    }
  };
}

/**
 * 現在のアクセスユーザーがデバッグを有効にすべきかを判定
 * @returns {boolean} デバッグモードを有効にすべき場合はtrue
 */
function shouldEnableDebugMode() {
  return isSystemAdmin();
}

/**
 * デプロイユーザーかどうか判定
 * 1. GASスクリプトの編集権限を確認
 * 2. 教育機関ドメインの管理者権限を確認
 * @returns {boolean} 管理者権限を持つ場合 true
 */
function isSystemAdmin() {
  try {
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    var currentUserEmail = Session.getActiveUser().getEmail();
    return adminEmail && currentUserEmail && adminEmail === currentUserEmail;
  } catch (e) {
    console.error('isSystemAdmin エラー: ' + e.message);
    return false;
  }
}

function isDeployUser() {
  return isSystemAdmin();
}

// =================================================================
// マルチテナント対応関数（同時アクセス対応）
// =================================================================

/**
 * マルチテナント対応: ステータス取得
 * @param {string} requestUserId - リクエストされたユーザーID（第一引数）
 * @param {boolean} forceRefresh - キャッシュを強制的にリフレッシュするかどうか（第二引数、オプション）
 * @returns {Object} ステータス情報
 */
function getStatus(requestUserId, forceRefresh = false) {
  try {
    console.log('getStatus - requestUserId:', requestUserId);
    
    // セキュリティチェック: 認証済みユーザーと要求されたuserIdが一致するか確認
    const currentUserEmail = Session.getActiveUser().getEmail();
    const requestedUserInfo = findUserById(requestUserId);
    
    if (!requestedUserInfo) {
      throw new Error('指定されたユーザーIDが見つかりません: ' + requestUserId);
    }
    
    if (requestedUserInfo.adminEmail !== currentUserEmail) {
      throw new Error('他のユーザーのデータにはアクセスできません');
    }
    
    // 従来のgetStatus処理をuserIdベースで実行
    if (forceRefresh) {
      clearExecutionUserInfoCache(); // 実行レベルキャッシュをクリア
    }
    let userInfo = getCachedUserInfo(requestUserId);
    
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }
    
    // 自動修復: スプレッドシートIDがあるのにconfigJsonにpublishedSpreadsheetIdがない場合
    let configJson = JSON.parse(userInfo.configJson || '{}');
    if (userInfo.spreadsheetId && !configJson.publishedSpreadsheetId) {
      debugLog('getStatus: Auto-repairing missing publishedSpreadsheetId');
      const repairResult = repairUserConfigJson(requestUserId);
      if (repairResult.status === 'success') {
        // 修復後、キャッシュをクリアして最新のユーザー情報を再取得
        invalidateUserCache(requestUserId, userInfo.adminEmail, userInfo.spreadsheetId, true);
        userInfo = findUserById(requestUserId);
        // configJsonも更新
        configJson = JSON.parse(userInfo.configJson || '{}');
        debugLog('getStatus: After repair - formCreated:', configJson.formCreated, 'setupStatus:', configJson.setupStatus);
      }
    }

    // 追加の自動修復: スプレッドシートが存在し、フォームの回答シートがあるのにformCreatedがfalseの場合
    if (userInfo.spreadsheetId && configJson.publishedSpreadsheetId && !configJson.formCreated) {
      try {
        const sheets = getSheetsList(requestUserId);
        const hasFormResponseSheet = sheets && sheets.some(sheet => {
          const sheetName = typeof sheet === 'string' ? sheet : (sheet.name || sheet.title || '');
          return sheetName.toLowerCase().includes('フォームの回答') || 
                 sheetName.toLowerCase().includes('form responses');
        });
        
        if (hasFormResponseSheet) {
          debugLog('getStatus: Auto-repairing formCreated status - form response sheet detected');
          configJson.formCreated = true;
          configJson.setupStatus = configJson.appPublished ? 'published' : 'completed';
          
          updateUser(requestUserId, {
            configJson: JSON.stringify(configJson)
          });
          
          // userInfoも再取得して同期
          userInfo = findUserById(requestUserId);
          configJson = JSON.parse(userInfo.configJson || '{}');
          debugLog('getStatus: formCreated status repaired to:', configJson.formCreated);
        }
      } catch (e) {
        console.warn('getStatus: Failed to auto-repair formCreated status:', e.message);
      }
    }
    
    // mainQuestionからopinionHeaderへの移行処理
    if (configJson.mainQuestion && configJson.publishedSheetName) {
      try {
        const sheetKey = 'sheet_' + configJson.publishedSheetName;
        if (!configJson[sheetKey]) {
          configJson[sheetKey] = {};
        }
        // 既存のopinionHeaderがない場合のみ移行
        if (!configJson[sheetKey].opinionHeader) {
          configJson[sheetKey].opinionHeader = configJson.mainQuestion;
          console.log('✅ [Migration] Moved mainQuestion to opinionHeader for sheet:', configJson.publishedSheetName);
        }
        delete configJson.mainQuestion;
        
        // 設定を保存
        updateUser(requestUserId, {
          configJson: JSON.stringify(configJson)
        });
        
        // userInfoも再取得して同期
        userInfo = findUserById(requestUserId);
        configJson = JSON.parse(userInfo.configJson || '{}');
      } catch (e) {
        console.warn('getStatus: Failed to migrate mainQuestion to opinionHeader:', e.message);
      }
    }
    
    // スプレッドシート情報を取得
    let sheetNames = [];
    try {
      if (userInfo.spreadsheetId) {
        debugLog('getStatus: Attempting to get sheet list for userId:', requestUserId);
        debugLog('getStatus: userInfo.spreadsheetId:', userInfo.spreadsheetId);
        const sheets = getSheetsList(requestUserId);
        debugLog('getStatus: getSheetsList returned:', sheets);
        debugLog('getStatus: getSheetsList type:', typeof sheets);
        
        if (sheets === undefined) {
          console.warn('getStatus: getSheetsList returned undefined, trying direct call');
          const directSheets = getSheetsListInternal(requestUserId);
          debugLog('getStatus: getSheetsListInternal returned:', directSheets);
          sheetNames = directSheets || [];
        } else {
          sheetNames = sheets || [];
        }
      } else {
        console.warn('getStatus: No spreadsheetId found in userInfo');
      }
    } catch (sheetError) {
      console.error('スプレッドシート情報の取得に失敗:', sheetError.message);
      console.error('getStatus sheetError details:', sheetError.stack);
    }
    
    // 設定情報は既に上でパース済み
    
    // カスタムフォーム情報を取得
    let customFormInfo = null;
    if (configJson.formCreated && configJson.lastFormCreatedAt) {
      try {
        // カスタムフォームの設定情報を取得
        // customFormInfo は opinionHeader に統一したため削除
        customFormInfo = null;
      } catch (e) {
        console.warn('カスタムフォーム情報の取得に失敗:', e.message);
      }
    }
    
    // フォームURLの取得と修復
    let formUrl = configJson.formUrl || configJson.editFormUrl;
    
    // フォームURLがない場合、スプレッドシートから取得を試みる
    if (!formUrl && userInfo.spreadsheetId && configJson.formCreated) {
      try {
        const retrievedFormUrl = getFormUrlSafely(configJson, userInfo.spreadsheetId);
        if (retrievedFormUrl) {
          formUrl = retrievedFormUrl;
          // configJsonを更新
          configJson.formUrl = formUrl;
          updateUser(requestUserId, {
            configJson: JSON.stringify(configJson)
          });
          debugLog('getStatus: フォームURLを自動修復しました:', formUrl);
        }
      } catch (e) {
        console.warn('getStatus: フォームURL自動修復に失敗:', e.message);
      }
    }
    
    // 公開状態の判定
    const isPublished = !!(configJson.appPublished && configJson.publishedSpreadsheetId && configJson.publishedSheetName);

    let topic = '（問題文未設定）';
    // opinionHeaderのみを使用するように統一
    if (configJson.publishedSheetName) {
      const sheetKey = 'sheet_' + configJson.publishedSheetName;
      const sheetConfig = configJson[sheetKey] || {};
      topic = sheetConfig.opinionHeader || configJson.publishedSheetName;
    }

    return {
      status: 'success',
      userInfo: userInfo,
      sheetNames: sheetNames,
      setupStep: configJson.setupStatus === 'completed' ? 3 : 2,
      activeSheetName: configJson.publishedSheetName || '',
      publishedSheetName: configJson.publishedSheetName || null,
      isPublished: isPublished,
      appPublished: configJson.appPublished || false,
      formUrl: formUrl || null,
      webAppUrl: getWebAppUrlCached(),
      appUrls: generateAppUrls(requestUserId),
      customFormInfo: customFormInfo,
      currentTopic: topic
    };
    
  } catch (error) {
    console.error('getStatus error:', error.message);
    return { status: 'error', message: error.message };
  }
}

/**
 * マルチテナント対応: 設定保存・公開
 * @param {string} requestUserId - リクエストされたユーザーID（第一引数）
 * @param {Object} settingsData - 設定データ
 * @returns {Object} 実行結果
 */

/**
 * マルチテナント対応: アクティブフォーム情報取得
 * @param {string} requestUserId - リクエストされたユーザーID
 * @returns {Object} フォーム情報
 */
/**
 * 管理者によるユーザー削除（AppSetupPage.html用ラッパー）
 * @param {string} targetUserId - 削除対象ユーザーID
 * @param {string} reason - 削除理由
 */
function deleteUserAccountByAdminForUI(targetUserId, reason) {
  try {
    const result = deleteUserAccountByAdmin(targetUserId, reason);
    return {
      status: 'success',
      message: result.message,
      deletedUser: result.deletedUser
    };
  } catch (error) {
    console.error('deleteUserAccountByAdmin wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 削除ログ取得（AppSetupPage.html用ラッパー）
 */
function getDeletionLogsForUI() {
  try {
    const logs = getDeletionLogs();
    return {
      status: 'success',
      logs: logs
    };
  } catch (error) {
    console.error('getDeletionLogs wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}



/**
 * ボードを作成 (createBoardFromAdminのエイリアス)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function createBoard(requestUserId) {
  return createBoardFromAdmin(requestUserId);
}

/**
 * ユーザーのセットアップページアクセス権限をチェックする
 */
function checkSetupPageAccess() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userInfo = findUserByEmail(activeUserEmail);
    var hasAccess = hasSetupPageAccess();
    
    return {
      status: 'success',
      email: activeUserEmail,
      userInfo: userInfo,
      hasAccess: hasAccess,
      isActive: userInfo ? userInfo.isActive : null
    };
  } catch (e) {
    return {
      status: 'error',
      message: e.message
    };
  }
}

/**
 * 全ユーザー一覧を取得（AppSetupPage.html用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getAllUsersForAdminForUI(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    const result = getAllUsersForAdmin();
    return {
      status: 'success',
      users: result
    };
  } catch (error) {
    console.error('getAllUsersForAdminForUI wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * カスタムフォーム作成（AdminPanel.html用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} config - フォーム設定
 */
function createCustomFormUI(requestUserId, config) {
  try {
    verifyUserAccess(requestUserId);
    const activeUserEmail = Session.getActiveUser().getEmail();
    
    const result = createCustomForm(activeUserEmail, requestUserId, config);
    
    // 既存ユーザーの情報を更新（スプレッドシート情報を追加）
    const existingUser = findUserById(requestUserId);
    if (existingUser) {
      console.log('createCustomFormUI - updating user data for:', requestUserId);
      
      const updatedConfigJson = JSON.parse(existingUser.configJson || '{}');
      updatedConfigJson.formUrl = result.formUrl;
      updatedConfigJson.editFormUrl = result.editFormUrl || result.formUrl;
      updatedConfigJson.formCreated = true;
      updatedConfigJson.lastFormCreatedAt = new Date().toISOString();
      updatedConfigJson.setupStatus = 'completed';
      updatedConfigJson.appPublished = true;
      updatedConfigJson.publishedSpreadsheetId = result.spreadsheetId;
      updatedConfigJson.publishedSheetName = result.sheetName;
      
      // 自動公開フラグを設定（新規フォーム作成時）
      updatedConfigJson.autoPublishAfterCreation = true;
      updatedConfigJson.readyForAutoPublish = true;
      
      // カスタムフォーム設定情報を保存 (opinionHeaderに統一)
      updatedConfigJson.formTitle = config.formTitle;
      updatedConfigJson.questionType = config.questionType;
      updatedConfigJson.choices = config.choices;
      updatedConfigJson.includeOthers = config.includeOthers;
      updatedConfigJson.enableClass = config.enableClass;
      updatedConfigJson.classChoices = config.classChoices;
      
      // 既存のmainQuestionをopinionHeaderに移行して削除
      if (updatedConfigJson.mainQuestion) {
        // シート固有の設定に移行
        const sheetKey = 'sheet_' + result.sheetName;
        if (!updatedConfigJson[sheetKey]) {
          updatedConfigJson[sheetKey] = {};
        }
        // 既存のopinionHeaderがない場合のみ移行
        if (!updatedConfigJson[sheetKey].opinionHeader) {
          updatedConfigJson[sheetKey].opinionHeader = updatedConfigJson.mainQuestion;
        }
        delete updatedConfigJson.mainQuestion;
        console.log('✅ Migrated mainQuestion to opinionHeader for sheet:', result.sheetName);
      }
      
      // 新しく作成されたスプレッドシート情報をメインのユーザー情報として更新
      const updateData = {
        spreadsheetId: result.spreadsheetId,
        spreadsheetUrl: result.spreadsheetUrl,
        configJson: JSON.stringify(updatedConfigJson),
        lastAccessedAt: new Date().toISOString()
      };
      
      console.log('createCustomFormUI - update data:', JSON.stringify(updateData));
      
      updateUser(requestUserId, updateData);
      
      // キャッシュをクリアして次回取得時に最新データを確保
      invalidateUserCache(requestUserId, activeUserEmail, result.spreadsheetId, true);
    } else {
      console.warn('createCustomFormUI - user not found:', requestUserId);
    }
    
    return {
      status: 'success',
      message: 'カスタムフォームが正常に作成されました！',
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      formTitle: result.formTitle,
      spreadsheetId: result.spreadsheetId,
      sheetName: result.sheetName,
      autoPublishReady: true
    };
  } catch (error) {
    console.error('createCustomFormUI error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}



/**
 * フォーム作成後の自動設定保存と公開処理
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 * @param {object} formData - フォーム作成結果データ
 */
function autoSaveAndPublishAfterFormCreation(requestUserId, sheetName, formData) {
  try {
    verifyUserAccess(requestUserId);
    console.log('autoSaveAndPublishAfterFormCreation 開始 - userId:', requestUserId, 'sheetName:', sheetName);
    
    // 基本設定オブジェクトを作成
    const basicConfig = {
      opinionHeader: formData.opinionHeader || 'お題',  // mainQuestion を opinionHeader に統一
      displayMode: 'anonymous',
      showCounts: false
    };
    
    // Step 1: 設定を保存
    console.log('自動設定保存開始...');
    saveSheetConfig(requestUserId, formData.spreadsheetId, sheetName, basicConfig);
    
    // Step 2: シートをアクティブ化
    console.log('シート自動アクティブ化開始...');
    switchToSheet(requestUserId, formData.spreadsheetId, sheetName);
    
    // Step 3: 表示オプションを設定
    const displayOptions = {
      displayMode: 'anonymous',
      showCounts: false
    };
    setDisplayOptions(requestUserId, displayOptions);
    
    // Step 4: 公開状態を更新
    console.log('自動公開状態更新開始...');
    const userInfo = getUserInfo(requestUserId);
    if (userInfo) {
      const currentConfig = JSON.parse(userInfo.configJson || '{}');
      currentConfig.appPublished = true;
      currentConfig.setupStatus = 'published';
      currentConfig.lastPublishedAt = new Date().toISOString();
      currentConfig.autoPublishCompleted = true;
      
      updateUser(requestUserId, {
        configJson: JSON.stringify(currentConfig)
      });
      
      // キャッシュクリア
      invalidateUserCache(requestUserId, userInfo.adminEmail, userInfo.spreadsheetId, true);
    }
    
    console.log('autoSaveAndPublishAfterFormCreation 完了');
    
    return {
      status: 'success',
      message: 'フォーム作成、設定保存、公開が自動で完了しました！',
      published: true,
      autoPublishCompleted: true
    };
    
  } catch (error) {
    console.error('autoSaveAndPublishAfterFormCreation error:', error.message);
    return {
      status: 'error',
      message: '自動設定・公開処理でエラーが発生しました: ' + error.message,
      published: false
    };
  }
}

/**
 * 現在のユーザーアカウントを削除（AdminPanel.html用）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function deleteCurrentUserAccount(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    const result = deleteUserAccount(requestUserId);
    
    return {
      status: 'success',
      message: 'アカウントが正常に削除されました',
      result: result
    };
  } catch (error) {
    console.error('deleteCurrentUserAccount error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 既存ユーザーのconfigJsonを修復する関数
 * フォーム作成済みだがpublishedSpreadsheetIdが欠如している場合に使用
 */
function repairUserConfigJson(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    const configJson = JSON.parse(userInfo.configJson || '{}');
    
    // スプレッドシートIDがあるが、configJsonにpublishedSpreadsheetIdがない場合
    if (userInfo.spreadsheetId && !configJson.publishedSpreadsheetId) {
      console.log('repairUserConfigJson: Repairing missing publishedSpreadsheetId for user:', requestUserId);
      
      configJson.publishedSpreadsheetId = userInfo.spreadsheetId;
      
      // シート名を取得して設定
      try {
        const sheets = getSheetsListInternal(requestUserId);
        if (sheets && sheets.length > 0) {
          // 最初のシートを公開シートとして設定（通常は「フォームの回答 1」）
          configJson.publishedSheetName = sheets[0].name;
          console.log('repairUserConfigJson: Set publishedSheetName to:', sheets[0].name);
        }
      } catch (sheetError) {
        console.warn('repairUserConfigJson: Failed to get sheet names:', sheetError.message);
      }
      
      // フォーム作成状態を正しく設定
      // スプレッドシートIDがあるということは、フォームも作成済みである
      if (userInfo.spreadsheetId && !configJson.formCreated) {
        configJson.formCreated = true;
        configJson.setupStatus = 'completed';
        configJson.appPublished = true; // スプレッドシートがあれば公開可能
        console.log('repairUserConfigJson: Set formCreated=true, setupStatus=completed, appPublished=true');
      }
      
      // 更新
      updateUser(requestUserId, {
        configJson: JSON.stringify(configJson)
      });
      
      console.log('repairUserConfigJson: Successfully repaired configJson for user:', requestUserId);
      return { status: 'success', message: 'configJsonを修復しました' };
    } else {
      return { status: 'no_action', message: '修復の必要はありません' };
    }
    
  } catch (error) {
    console.error('repairUserConfigJson error:', error.message);
    return { status: 'error', message: error.message };
  }
}

/**
 * シートを有効化（AdminPanel.html用のシンプル版）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 */
function activateSheetSimple(requestUserId, sheetName) {
  try {
    verifyUserAccess(requestUserId);
    const userInfo = findUserById(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }
    
    return activateSheet(requestUserId, userInfo.spreadsheetId, sheetName);
  } catch (error) {
    console.error('activateSheetSimple error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}
