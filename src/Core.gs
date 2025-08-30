/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

/**
 * 自動停止時間を計算する
 * @param {string} publishedAt - 公開開始時間のISO文字列
 * @param {number} minutes - 自動停止までの分数
 * @returns {object} 停止時間情報
 */
function getAutoStopTime(publishedAt, minutes) {
  try {
    const publishTime = new Date(publishedAt);
    const stopTime = new Date(publishTime.getTime() + minutes * 60 * 1000);

    return {
      publishTime: publishTime,
      stopTime: stopTime,
      publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
      stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
      remainingMinutes: Math.max(
        0,
        Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))
      ),
    };
  } catch (error) {
    console.error('自動停止時間計算エラー:', error);
    return null;
  }
}

/**
 * セットアップステップを統一的に判定する関数
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} configJson - 設定JSON
 * @returns {number} セットアップステップ (1-3)
 */
function determineSetupStep(userInfo, configJson) {
  const setupStatus = configJson ? configJson.setupStatus : 'pending';

  // Step 1: データソース未設定 OR セットアップ初期状態
  if (
    !userInfo ||
    !userInfo.spreadsheetId ||
    userInfo.spreadsheetId.trim() === '' ||
    setupStatus === 'pending'
  ) {
    console.log('determineSetupStep: Step 1 - データソース未設定またはセットアップ初期状態');
    return 1;
  }

  // Step 2: データソース設定済み + (再設定中 OR セットアップ未完了)
  if (
    !configJson ||
    setupStatus === 'reconfiguring' ||
    setupStatus !== 'completed' ||
    !configJson.formCreated
  ) {
    console.log(
      'determineSetupStep: Step 2 - データソース設定済み、セットアップ未完了または再設定中'
    );
    return 2;
  }

  // Step 3: 全セットアップ完了 + 公開済み
  if (setupStatus === 'completed' && configJson.formCreated && configJson.appPublished) {
    console.log('determineSetupStep: Step 3 - 全セットアップ完了・公開済み');
    return 3;
  }

  // フォールバック: Step 2
  console.log('determineSetupStep: フォールバック - Step 2');
  return 2;
}

// =================================================================
// ユーザー情報キャッシュ（関数実行中の重複取得を防ぐ）
// =================================================================

let _executionSheetsServiceCache = null;

/**
 * ヘッダー整合性検証（リアルタイム検証用）
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(userId) {
  try {
    console.log('🔍 Starting header integrity validation for userId:', userId);

    const userInfo = getCurrentUserInfo();
    if (!userInfo || !userInfo.spreadsheetId) {
      return {
        success: false,
        error: 'User spreadsheet not found',
        timestamp: new Date().toISOString(),
      };
    }

    const spreadsheetId = userInfo.spreadsheetId;
    const sheetName = userInfo.sheetName || 'EABDB';

    // 理由列のヘッダー検証を重点的に実施
    const indices = getHeaderIndices(spreadsheetId, sheetName);

    const validationResults = {
      success: true,
      timestamp: new Date().toISOString(),
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      headerValidation: {
        reasonColumnIndex: indices[COLUMN_HEADERS.REASON],
        opinionColumnIndex: indices[COLUMN_HEADERS.OPINION],
        hasReasonColumn: indices[COLUMN_HEADERS.REASON] !== undefined,
        hasOpinionColumn: indices[COLUMN_HEADERS.OPINION] !== undefined,
      },
      issues: [],
    };

    // 理由列の必須チェック
    if (indices[COLUMN_HEADERS.REASON] === undefined) {
      validationResults.success = false;
      validationResults.issues.push('Reason column (理由) not found in headers');
    }

    // 回答列の必須チェック
    if (indices[COLUMN_HEADERS.OPINION] === undefined) {
      validationResults.success = false;
      validationResults.issues.push('Opinion column (回答) not found in headers');
    }

    // ログ出力
    if (validationResults.success) {
      console.log('✅ Header integrity validation passed');
    } else {
      console.warn('⚠️ Header integrity validation failed:', validationResults.issues);
    }

    return validationResults;
  } catch (error) {
    console.error('❌ Header integrity validation error:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
    };
  }
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
    // unifiedUserManager.gsの関数を使用
    const userInfo = getCurrentUserInfo();
    if (!userInfo) {
      return 'お題';
    }

    const config = JSON.parse(userInfo.configJson || '{}');
    const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
    const sheetConfig = config[sheetConfigKey] || {};

    const opinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';

    console.log('getOpinionHeaderSafely:', {
      userId: userId,
      sheetName: sheetName,
      opinionHeader: opinionHeader,
    });

    return opinionHeader;
  } catch (e) {
    console.error('getOpinionHeaderSafely error:', e);
    return 'お題';
  }
}

/**
 * 新規ユーザーを登録する（データベース登録のみ）
 * フォーム作成はクイックスタートで実行される
 */
function registerNewUser(adminEmail) {
  const activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // ドメイン制限チェック
  const domainInfo = Deploy.domain();
  if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
    throw new Error(
      `ドメインアクセスが制限されています。許可されたドメイン: ${domainInfo.deployDomain}, 現在のドメイン: ${domainInfo.currentDomain}`
    );
  }

  // 既存ユーザーチェック（1ユーザー1行の原則）
  const existingUser = DB.findUserByEmail(adminEmail);
  let userId, appUrls;

  if (existingUser) {
    // 既存ユーザーの場合は最小限の更新のみ（設定は保護）
    userId = existingUser.userId;
    const existingConfig = JSON.parse(existingUser.configJson || '{}');

    // 最終アクセス時刻とアクティブ状態のみ更新（設定は保護）
    updateUser(userId, {
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      // 注意: configJsonは更新しない（既存の設定を保護）
    });

    // キャッシュを無効化して最新状態を反映
    invalidateUserCache(userId, adminEmail, existingUser.spreadsheetId, false);

    console.log('✅ 既存ユーザーの最終アクセス時刻を更新しました（設定は保護）: ' + adminEmail);
    appUrls = generateUserUrls(userId);

    return {
      userId: userId,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupRequired: false, // 既存ユーザーはセットアップ完了済みと仮定
      message: 'おかえりなさい！管理パネルへリダイレクトします。',
      isExistingUser: true,
    };
  }

  // 新規ユーザーの場合
  userId = Utilities.getUuid();

  const initialConfig = {
    setupStatus: 'pending',
    createdAt: new Date().toISOString(),
    formCreated: false,
    appPublished: false,
  };

  const userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: '',
    spreadsheetUrl: '',
    createdAt: new Date().toISOString(),
    configJson: JSON.stringify(initialConfig),
    lastAccessedAt: new Date().toISOString(),
    isActive: 'true',
  };

  try {
    DB.createUser(userData);
    console.log('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);
    // 生成されたユーザー情報のキャッシュをクリア
    invalidateUserCache(userId, adminEmail, null, false);
  } catch (e) {
    console.error('データベースへのユーザー登録に失敗: ' + e.message);
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す
  appUrls = generateUserUrls(userId);
  return {
    userId: userId,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupRequired: true,
    message: 'ユーザー登録が完了しました！次にクイックスタートでフォームを作成してください。',
    isExistingUser: false,
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
    const reactingUserEmail = User.email();
    const ownerUserId = requestUserId; // requestUserId を使用

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    const boardOwnerInfo = DB.findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    const result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      // 更新後のリアクション情報を取得
      const updatedReactions = getRowReactions(
        boardOwnerInfo.spreadsheetId,
        sheetName,
        rowIndex,
        reactingUserEmail
      );

      return {
        status: 'ok',
        reactions: updatedReactions,
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    console.error('addReaction エラー: ' + e.message);
    return {
      status: 'error',
      message: e.message,
    };
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

/**
 * バッチリアクション処理（既存addReaction機能を保持したまま追加）
 * 複数のリアクション操作を効率的に一括処理
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {Array} batchOperations - バッチ操作配列 [{rowIndex, reaction, timestamp}, ...]
 * @returns {object} バッチ処理結果
 */
function addReactionBatch(requestUserId, batchOperations) {
  verifyUserAccess(requestUserId);
  clearExecutionUserInfoCache();

  try {
    // 入力検証
    if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
      throw new Error('バッチ操作が無効です');
    }

    // バッチサイズ制限（安全性のため）
    const MAX_BATCH_SIZE = 20;
    if (batchOperations.length > MAX_BATCH_SIZE) {
      throw new Error(`バッチサイズが制限を超えています (最大${MAX_BATCH_SIZE}件)`);
    }

    console.log('🔄 バッチリアクション処理開始:', batchOperations.length + '件');

    const reactingUserEmail = User.email();
    const ownerUserId = requestUserId;

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    const boardOwnerInfo = DB.findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    // バッチ処理結果を格納
    const batchResults = [];
    const processedRows = new Set(); // 重複行の追跡

    // 既存のsheetNameを取得（最初の操作から）
    const sheetName = getCurrentSheetName(boardOwnerInfo.spreadsheetId);

    // バッチ操作を順次処理（既存のprocessReaction関数を再利用）
    for (let i = 0; i < batchOperations.length; i++) {
      const operation = batchOperations[i];

      try {
        // 入力検証
        if (!operation.rowIndex || !operation.reaction) {
          console.warn('無効な操作をスキップ:', operation);
          continue;
        }

        // 既存のprocessReaction関数を使用（100%互換性保証）
        const result = processReaction(
          boardOwnerInfo.spreadsheetId,
          sheetName,
          operation.rowIndex,
          operation.reaction,
          reactingUserEmail
        );

        if (result && result.status === 'success') {
          // 更新後のリアクション情報を取得
          const updatedReactions = getRowReactions(
            boardOwnerInfo.spreadsheetId,
            sheetName,
            operation.rowIndex,
            reactingUserEmail
          );

          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            reactions: updatedReactions,
            status: 'success',
          });

          processedRows.add(operation.rowIndex);
        } else {
          console.warn('リアクション処理失敗:', operation, result.message);
          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            status: 'error',
            message: result.message || 'リアクション処理失敗',
          });
        }
      } catch (operationError) {
        console.error('個別操作エラー:', operation, operationError.message);
        batchResults.push({
          rowIndex: operation.rowIndex,
          reaction: operation.reaction,
          status: 'error',
          message: operationError.message,
        });
      }
    }

    // 成功した行の最新状態を収集
    const finalResults = [];
    processedRows.forEach(function (rowIndex) {
      try {
        const latestReactions = getRowReactions(
          boardOwnerInfo.spreadsheetId,
          sheetName,
          rowIndex,
          reactingUserEmail
        );
        finalResults.push({
          rowIndex: rowIndex,
          reactions: latestReactions,
        });
      } catch (error) {
        console.warn('最終状態取得エラー:', rowIndex, error.message);
      }
    });

    console.log('✅ バッチリアクション処理完了:', {
      total: batchOperations.length,
      processed: processedRows.size,
      success: batchResults.filter((r) => r.status === 'success').length,
    });

    return {
      success: true,
      data: finalResults,
      processedCount: batchOperations.length,
      successCount: batchResults.filter((r) => r.status === 'success').length,
      timestamp: new Date().toISOString(),
      details: batchResults, // デバッグ用詳細情報
    };
  } catch (error) {
    console.error('addReactionBatch エラー:', error.message);

    // バッチ処理失敗時は個別処理にフォールバック可能であることを示す
    return {
      success: false,
      error: error.message,
      fallbackToIndividual: true, // クライアント側が個別処理にフォールバック可能
      timestamp: new Date().toISOString(),
    };
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

/**
 * 現在のシート名を取得するヘルパー関数
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string} シート名
 */
function getCurrentSheetName(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    // デフォルトでは最初のシートを使用
    if (sheets.length > 0) {
      return sheets[0].getName();
    }

    throw new Error('シートが見つかりません');
  } catch (error) {
    console.warn('シート名取得エラー:', error.message);
    return 'Sheet1'; // フォールバック
  }
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * マルチテナント環境でのユーザーアクセス権限を検証します。
 * リクエストを投げたユーザー (User.email()) が、
 * requestUserId のデータにアクセスする権限を持っているかを確認します。
 * 権限がない場合はエラーをスローします。
 * @param {string} requestUserId - アクセスを要求しているユーザーのID
 * @throws {Error} 認証エラーまたは権限エラー
 */
function verifyUserAccess(requestUserId) {
  // 新しいAccessControllerシステムを使用（後方互換性のためのラッパー関数）
  const currentUserEmail = Session.getActiveUser().getEmail();
  const result = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
  
  if (!result.allowed) {
    throw new Error(`認証エラー: ${result.message}`);
  }
  
  console.log(`✅ ユーザーアクセス検証成功: ${currentUserEmail} は ${requestUserId} のデータにアクセスできます。`);
  return result;
}

/**
 * 実際のデータ取得処理（キャッシュ制御から分離） (マルチテナント対応版)
 */
function executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode) {
  try {
    const currentUserId = requestUserId; // requestUserId を使用
    console.log(
      'getPublishedSheetData: userId=%s, classFilter=%s, sortOrder=%s, adminMode=%s',
      currentUserId,
      classFilter,
      sortOrder,
      adminMode
    );

    const userInfo = getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    console.log('getPublishedSheetData: userInfo=%s', JSON.stringify(userInfo));

    const configJson = JSON.parse(userInfo.configJson || '{}');
    console.log('getPublishedSheetData: configJson=%s', JSON.stringify(configJson));

    // セットアップ状況を確認
    const setupStatus = configJson.setupStatus || 'pending';

    // 公開対象のスプレッドシートIDとシート名を取得
    const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    const publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      if (setupStatus === 'pending') {
        // セットアップ未完了の場合は適切なメッセージを返す
        return {
          status: 'setup_required',
          message:
            'セットアップを完了してください。データ準備、シート・列設定、公開設定の順番で進めてください。',
          data: [],
          setupStatus: setupStatus,
        };
      }
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // シート固有の設定を取得 (sheetKey is based only on sheet name)
    const sheetKey = 'sheet_' + publishedSheetName;
    const sheetConfig = configJson[sheetKey] || {};
    console.log('getPublishedSheetData: sheetConfig=%s', JSON.stringify(sheetConfig));

    // Check if current user is the board owner
    const isOwner = configJson.ownerId === currentUserId;
    console.log(
      'getPublishedSheetData: isOwner=%s, ownerId=%s, currentUserId=%s',
      isOwner,
      configJson.ownerId,
      currentUserId
    );

    // データ取得
    const sheetData = getSheetData(
      currentUserId,
      publishedSheetName,
      classFilter,
      sortOrder,
      adminMode
    );
    console.log(
      'getPublishedSheetData: sheetData status=%s, totalCount=%s',
      sheetData.status,
      sheetData.totalCount
    );

    if (sheetData.status === 'error') {
      throw new Error(sheetData.message);
    }

    // Page.html期待形式に変換
    // 設定からヘッダー名を取得。setupStatus未完了時は安全なデフォルト値を使用。
    let mainHeaderName;
    if (setupStatus === 'pending') {
      mainHeaderName = 'セットアップ中...';
    } else {
      mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    }

    // その他のヘッダーフィールドも安全に取得
    let reasonHeaderName, classHeaderName, nameHeaderName;
    if (setupStatus === 'pending') {
      reasonHeaderName = 'セットアップ中...';
      classHeaderName = 'セットアップ中...';
      nameHeaderName = 'セットアップ中...';
    } else {
      reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
      classHeaderName =
        sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
      nameHeaderName =
        sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    }
    console.log(
      'getPublishedSheetData: Configured Headers - mainHeaderName=%s, reasonHeaderName=%s, classHeaderName=%s, nameHeaderName=%s',
      mainHeaderName,
      reasonHeaderName,
      classHeaderName,
      nameHeaderName
    );

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    const headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);
    console.log('getPublishedSheetData: Available headerIndices=%s', JSON.stringify(headerIndices));

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    const mappedIndices = mapConfigToActualHeaders(
      {
        opinionHeader: mainHeaderName,
        reasonHeader: reasonHeaderName,
        classHeader: classHeaderName,
        nameHeader: nameHeaderName,
      },
      headerIndices
    );
    console.log('getPublishedSheetData: Mapped indices=%s', JSON.stringify(mappedIndices));

    const formattedData = formatSheetDataForFrontend(
      sheetData.data,
      mappedIndices,
      headerIndices,
      adminMode,
      isOwner,
      sheetData.displayMode
    );

    console.log('getPublishedSheetData: formattedData length=%s', formattedData.length);
    console.log('getPublishedSheetData: formattedData content=%s', JSON.stringify(formattedData));

    // ボードのタイトルを実際のスプレッドシートのヘッダーから取得
    let headerTitle = publishedSheetName || '今日のお題';
    if (mappedIndices.opinionHeader !== undefined) {
      for (var actualHeader in headerIndices) {
        if (headerIndices[actualHeader] === mappedIndices.opinionHeader) {
          headerTitle = actualHeader;
          console.log('getPublishedSheetData: Using actual header as title: "%s"', headerTitle);
          break;
        }
      }
    }

    const finalDisplayMode =
      adminMode === true ? DISPLAY_MODES.NAMED : configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    const result = {
      header: headerTitle,
      sheetName: publishedSheetName,
      showCounts: adminMode === true ? true : configJson.showCounts === true,
      displayMode: finalDisplayMode,
      data: formattedData,
      rows: formattedData, // 後方互換性のため
    };

    console.log('🔍 最終結果:', {
      adminMode: adminMode,
      originalDisplayMode: sheetData.displayMode,
      finalDisplayMode: finalDisplayMode,
      dataCount: formattedData.length,
      showCounts: result.showCounts,
    });
    console.log('getPublishedSheetData: Returning result=%s', JSON.stringify(result));
    return result;
  } catch (e) {
    console.error('公開シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      rows: [],
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
    console.log('🔄 増分データ取得開始: sinceRowCount=%s', sinceRowCount);

    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    const setupStatus = configJson.setupStatus || 'pending';
    const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    const publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      if (setupStatus === 'pending') {
        // セットアップ未完了の場合は適切なメッセージを返す
        return {
          status: 'setup_required',
          message: 'セットアップを完了してください。',
          incrementalData: [],
          setupStatus: setupStatus,
        };
      }
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // スプレッドシートとシートを取得
    const ss = SpreadsheetApp.openById(publishedSpreadsheetId);
    console.log('DEBUG: Spreadsheet object obtained: %s', ss ? ss.getName() : 'null');

    const sheet = ss.getSheetByName(publishedSheetName);
    console.log('DEBUG: Sheet object obtained: %s', sheet ? sheet.getName() : 'null');

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + publishedSheetName);
    }

    const lastRow = sheet.getLastRow(); // スプレッドシートの最終行
    const headerRow = 1; // ヘッダー行は1行目と仮定

    // 実際に読み込むべき開始行を計算 (sinceRowCountはデータ行数なので、+1してヘッダーを考慮)
    // sinceRowCountが0の場合、ヘッダーの次の行から読み込む
    const startRowToRead = sinceRowCount + headerRow + 1;

    // 新しいデータがない場合
    if (lastRow < startRowToRead) {
      console.log(
        '🔍 増分データ分析: 新しいデータなし。lastRow=%s, startRowToRead=%s',
        lastRow,
        startRowToRead
      );
      return {
        header: '', // 必要に応じて設定
        sheetName: publishedSheetName,
        showCounts: configJson.showCounts === true,
        displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
        data: [],
        rows: [],
        totalCount: lastRow - headerRow, // ヘッダーを除いたデータ総数
        newCount: 0,
        isIncremental: true,
      };
    }

    // 読み込む行数
    const numRowsToRead = lastRow - startRowToRead + 1;

    // 必要なデータのみをスプレッドシートから直接取得
    // getRange(row, column, numRows, numColumns)
    // ここでは全列を取得すると仮定 (A列から最終列まで)
    const lastColumn = sheet.getLastColumn();
    const rawNewData = sheet.getRange(startRowToRead, 1, numRowsToRead, lastColumn).getValues();

    console.log('📥 スプレッドシートから直接取得した新しいデータ:', rawNewData.length, '件');

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    const headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    const sheetConfig = configJson['sheet_' + publishedSheetName] || {};
    const mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    const reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeaderName =
      sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
    const nameHeaderName =
      sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    const mappedIndices = mapConfigToActualHeaders(
      {
        opinionHeader: mainHeaderName,
        reasonHeader: reasonHeaderName,
        classHeader: classHeaderName,
        nameHeader: nameHeaderName,
      },
      headerIndices
    );

    // ユーザー情報と管理者モードの取得
    const isOwner = configJson.ownerId === currentUserId;
    const displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // 新しいデータを既存の処理パイプラインと同様に加工
    const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    const rosterMap = buildRosterMap([]); // roster is not used
    const processedData = rawNewData.map(function (row, idx) {
      return processRowData(
        row,
        headers,
        headerIndices,
        rosterMap,
        displayMode,
        startRowToRead + idx,
        isOwner
      );
    });

    // 取得した生データをPage.htmlが期待する形式にフォーマット
    const formattedNewData = formatSheetDataForFrontend(
      processedData,
      mappedIndices,
      headerIndices,
      adminMode,
      isOwner,
      displayMode
    );

    console.log('✅ 増分データ取得完了: %s件の新しいデータを返します', formattedNewData.length);

    return {
      header: '', // 必要に応じて設定
      sheetName: publishedSheetName,
      showCounts: false, // 必要に応じて設定
      displayMode: displayMode,
      data: formattedNewData,
      rows: formattedNewData, // 後方互換性のため
      totalCount: lastRow - headerRow, // ヘッダーを除いたデータ総数
      newCount: formattedNewData.length,
      isIncremental: true,
    };
  } catch (e) {
    console.error('増分データ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: '増分データの取得に失敗しました: ' + e.message,
    };
  }
}

/**
 * 指定されたユーザーのスプレッドシートからシートのリストを取得します。
 * @param {string} userId - ユーザーID
 * @returns {Array<Object>} シートのリスト（例: [{name: 'Sheet1', id: 'sheetId1'}, ...]）
 */
function getSheetsList(userId) {
  try {
    const userInfo = DB.findUserById(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      console.warn('getSheetsList: User info or spreadsheetId not found for userId:', userId);
      return [];
    }

    const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(function (sheet) {
      return {
        name: sheet.getName(),
        id: sheet.getSheetId(), // シートIDも必要に応じて取得
      };
    });

    console.log(
      '✅ getSheetsList: Found sheets for userId %s: %s',
      userId,
      JSON.stringify(sheetList)
    );
    return sheetList;
  } catch (e) {
    console.error('getSheetsList エラー: ' + e.message);
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
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
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
function formatSheetDataForFrontend(
  rawData,
  mappedIndices,
  headerIndices,
  adminMode,
  isOwner,
  displayMode
) {
  // 現在のユーザーメールを取得（リアクション状態判定用）
  const currentUserEmail = User.email();

  return rawData.map(function (row, index) {
    const classIndex = mappedIndices.classHeader;
    const opinionIndex = mappedIndices.opinionHeader;
    const reasonIndex = mappedIndices.reasonHeader;
    const nameIndex = mappedIndices.nameHeader;

    console.log(
      'formatSheetDataForFrontend: Row %s - classIndex=%s, opinionIndex=%s, reasonIndex=%s, nameIndex=%s',
      index,
      classIndex,
      opinionIndex,
      reasonIndex,
      nameIndex
    );
    console.log(
      'formatSheetDataForFrontend: Row data length=%s, first few values=%s',
      row.originalData ? row.originalData.length : 'undefined',
      row.originalData ? JSON.stringify(row.originalData.slice(0, 5)) : 'undefined'
    );

    let nameValue = '';
    const shouldShowName = adminMode === true || displayMode === DISPLAY_MODES.NAMED || isOwner;
    const hasNameIndex = nameIndex !== undefined;
    const hasOriginalData = row.originalData && row.originalData.length > 0;

    if (shouldShowName && hasNameIndex && hasOriginalData) {
      nameValue = row.originalData[nameIndex] || '';
    }

    if (!nameValue && shouldShowName && hasOriginalData) {
      const emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
      if (emailIndex !== undefined && row.originalData[emailIndex]) {
        nameValue = row.originalData[emailIndex].split('@')[0];
      }
    }

    console.log('🔍 サーバー側名前データ詳細:', {
      rowIndex: row.rowNumber || index + 2,
      shouldShowName: shouldShowName,
      adminMode: adminMode,
      displayMode: displayMode,
      isOwner: isOwner,
      nameIndex: nameIndex,
      hasNameIndex: hasNameIndex,
      hasOriginalData: hasOriginalData,
      originalDataLength: row.originalData ? row.originalData.length : 'undefined',
      nameValue: nameValue,
      rawNameData: hasOriginalData && nameIndex !== undefined ? row.originalData[nameIndex] : 'N/A',
    });

    // リアクション状態を判定するヘルパー関数
    function checkReactionState(reactionKey) {
      const columnName = COLUMN_HEADERS[reactionKey];
      const columnIndex = headerIndices[columnName];
      let count = 0;
      let reacted = false;

      if (columnIndex !== undefined && row.originalData && row.originalData[columnIndex]) {
        const reactionString = row.originalData[columnIndex].toString();
        if (reactionString) {
          const reactions = parseReactionString(reactionString);
          count = reactions.length;
          reacted = reactions.indexOf(currentUserEmail) !== -1;
        }
      }

      // フォールバック: プリカウントされた値を使用
      if (count === 0) {
        if (reactionKey === 'UNDERSTAND') count = row.understandCount || 0;
        else if (reactionKey === 'LIKE') count = row.likeCount || 0;
        else if (reactionKey === 'CURIOUS') count = row.curiousCount || 0;
      }

      return { count: count, reacted: reacted };
    }

    // 理由列の値を取得
    let reasonValue = '';
    if (
      reasonIndex !== undefined &&
      row.originalData &&
      row.originalData[reasonIndex] !== undefined
    ) {
      reasonValue = row.originalData[reasonIndex] || '';
    }

    return {
      rowIndex: row.rowNumber || index + 2,
      name: nameValue,
      email:
        row.originalData && row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]]
          ? row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]]
          : '',
      class:
        classIndex !== undefined && row.originalData && row.originalData[classIndex]
          ? row.originalData[classIndex]
          : '',
      opinion:
        opinionIndex !== undefined && row.originalData && row.originalData[opinionIndex]
          ? row.originalData[opinionIndex]
          : '',
      reason: reasonValue,
      reactions: {
        UNDERSTAND: checkReactionState('UNDERSTAND'),
        LIKE: checkReactionState('LIKE'),
        CURIOUS: checkReactionState('CURIOUS'),
      },
      highlight: row.isHighlighted || false,
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
    const currentUserId = requestUserId;
    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');

    // --- Auto-healing for inconsistent setup states ---
    let needsUpdate = false;
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

    const sheets = getSheetsList(currentUserId);
    const appUrls = generateUserUrls(currentUserId);

    // 回答数を取得
    let answerCount = 0;
    let totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        const responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
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
      formUrl: configJson.formUrl || '',
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
        configJson: userInfo.configJson || '{}',
      },
      // 統計情報
      answerCount: answerCount,
      totalReactions: totalReactions,
      // システム状態
      systemStatus: {
        setupStatus: configJson.setupStatus || 'unknown',
        formCreated: configJson.formCreated || false,
        appPublished: configJson.appPublished || false,
        lastUpdated: new Date().toISOString(),
      },
      // ユーザーのセットアップ段階を判定（統一化されたロジック）
      setupStep: determineSetupStep(userInfo, configJson),
    };
  } catch (e) {
    console.error('アプリ設定取得エラー: ' + e.message);
    return {
      status: 'error',
      message: '設定の取得に失敗しました: ' + e.message,
    };
  }
}

/**
 * シート設定を保存する（統合版：Batch/Optimized機能統合）
 * AdminPanel.htmlから呼び出される
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - 設定対象のスプレッドシートID
 * @param {string} sheetName - 設定対象のシート名
 * @param {object} config - 保存するシート固有の設定
 * @param {object} options - オプション設定
 * @param {object} options.sheetsService - 共有SheetsService（最適化用）
function switchToSheet(userId, spreadsheetId, sheetName, options = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    
    const currentUserId = userId;
    
    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    const userInfo = options.userInfo || DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');

    configJson.publishedSpreadsheetId = spreadsheetId;
    configJson.publishedSheetName = sheetName;
    configJson.appPublished = true; // シートを切り替えたら公開状態にする
    configJson.lastModified = new Date().toISOString();

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    console.log('✅ 表示シートを切り替えました: %s - %s', spreadsheetId, sheetName);
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
function getResponsesData(userId, sheetName) {
  const userInfo = getCurrentUserInfo();
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    const service = getSheetsServiceCached();
    const spreadsheetId = userInfo.spreadsheetId;
    const range = "'" + (sheetName || 'フォームの回答 1') + "'!A:Z";

    const response = batchGetSheetsData(service, spreadsheetId, [range]);
    const values = response.valueRanges[0].values || [];

    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1),
      headers: values[0],
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
    const activeUserEmail = User.email();

    // requestUserIdが無効な場合は、メールアドレスでユーザーを検索
    let userInfo;
    if (requestUserId && requestUserId.trim() !== '') {
      verifyUserAccess(requestUserId);
      userInfo = DB.findUserById(requestUserId);
    } else {
      userInfo = DB.findUserByEmail(activeUserEmail);
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
        lastAccessedAt: userInfo.lastAccessedAt,
      },
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
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');

    // フォーム回答数を取得
    let answerCount = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheet) {
        const responseData = getResponsesData(currentUserId, configJson.publishedSheet);
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
      editUrl: configJson.editFormUrl || '', // AdminPanel.htmlが期待するフィールド名
      formId: extractFormIdFromUrl(configJson.formUrl || configJson.editFormUrl || ''),
      spreadsheetUrl: userInfo.spreadsheetUrl || '',
      answerCount: answerCount,
      isFormActive: !!(configJson.formUrl && configJson.formCreated),
    };
  } catch (e) {
    console.error('getActiveFormInfo エラー: ' + e.message);
    return { status: 'error', message: 'フォーム情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * 指定シートのデータ行数を取得します。
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} sheetName シート名
 * @param {string} classFilter クラスフィルタ
 * @returns {number} データ行数
 */
function countSheetRows(spreadsheetId, sheetName, classFilter) {
  const key = `rowCount_${spreadsheetId}_${sheetName}_${classFilter}`;
  return cacheManager.get(
    key,
    () => {
      const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
      if (!sheet) return 0;

      const lastRow = sheet.getLastRow();
      if (!classFilter || classFilter === 'すべて') {
        return Math.max(0, lastRow - 1);
      }

      const headerIndices = getHeaderIndices(spreadsheetId, sheetName);
      const classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex === undefined) {
        return Math.max(0, lastRow - 1);
      }

      const values = sheet.getRange(2, classIndex + 1, lastRow - 1, 1).getValues();
      return values.reduce((cnt, row) => (row[0] === classFilter ? cnt + 1 : cnt), 0);
    },
    { ttl: 30, enableMemoization: true }
  );
}

/**
 * データ数を取得する (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {number} データ数
 */
function getDataCount(requestUserId, classFilter, sortOrder, adminMode) {
  verifyUserAccess(requestUserId);
  try {
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    const configJson = JSON.parse(userInfo.configJson || '{}');

    if (!configJson.publishedSpreadsheetId || !configJson.publishedSheetName) {
      return {
        count: 0,
        lastUpdate: new Date().toISOString(),
        status: 'error',
        message: 'スプレッドシート設定なし',
      };
    }

    const count = countSheetRows(
      configJson.publishedSpreadsheetId,
      configJson.publishedSheetName,
      classFilter
    );
    return {
      count,
      lastUpdate: new Date().toISOString(),
      status: 'success',
    };
  } catch (e) {
    console.error('getDataCount エラー: ' + e.message);
    return {
      count: 0,
      lastUpdate: new Date().toISOString(),
      status: 'error',
      message: e.message,
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
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');

    if (configJson.editFormUrl) {
      try {
        const formId = extractFormIdFromUrl(configJson.editFormUrl);
        const form = FormApp.openById(formId);

        if (title) {
          form.setTitle(title);
        }
        if (description) {
          form.setDescription(description);
        }

        return {
          status: 'success',
          message: 'フォーム設定が更新されました',
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
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');

    // システム設定を更新
    configJson.systemConfig = {
      cacheEnabled: config.cacheEnabled,
      autosaveInterval: config.autosaveInterval,
      logLevel: config.logLevel,
      updatedAt: new Date().toISOString(),
    };

    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson),
    });

    return {
      status: 'success',
      message: 'システム設定が保存されました',
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
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 管理者権限チェック - 現在のユーザーがボードの所有者かどうかを確認
    const activeUserEmail = User.email();
    if (activeUserEmail !== userInfo.adminEmail) {
      throw new Error('ハイライト機能は管理者のみ使用できます');
    }

    const result = processHighlightToggle(
      userInfo.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      return {
        status: 'ok',
        highlight: result.highlighted || false,
      };
    } else {
      throw new Error(result.message || 'ハイライト切り替えに失敗しました');
    }
  } catch (e) {
    console.error('toggleHighlight エラー: ' + e.message);
    return {
      status: 'error',
      message: e.message,
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
// createQuickStartFiles - クイックスタート機能は不要のため削除済み

/**
 * ユーザー専用フォルダを作成 (マルチテナント対応版)
 */
function createUserFolder(userEmail) {
  try {
    const rootFolderName = 'StudyQuest - ユーザーデータ';
    const userFolderName = 'StudyQuest - ' + userEmail + ' - ファイル';

    // ルートフォルダを検索または作成
    let rootFolder;
    const folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    // ユーザー専用フォルダを作成
    const userFolders = rootFolder.getFoldersByName(userFolderName);
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
    const service = getSheetsServiceCached();
    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    const highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];

    if (highlightColumnIndex === undefined) {
      throw new Error('ハイライト列が見つかりません');
    }

    // 現在の値を取得
    const range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
    const response = batchGetSheetsData(service, spreadsheetId, [range]);
    let isHighlighted = false;
    if (
      response &&
      response.valueRanges &&
      response.valueRanges[0] &&
      response.valueRanges[0].values &&
      response.valueRanges[0].values[0] &&
      response.valueRanges[0].values[0][0]
    ) {
      isHighlighted = response.valueRanges[0].values[0][0] === 'true';
    }

    // 値を切り替え
    const newValue = isHighlighted ? 'false' : 'true';
    updateSheetsData(service, spreadsheetId, range, [[newValue]]);

    // ハイライト更新後のキャッシュ無効化
    try {
      if (
        typeof cacheManager !== 'undefined' &&
        typeof cacheManager.invalidateSheetData === 'function'
      ) {
        cacheManager.invalidateSheetData(spreadsheetId, sheetName);
        console.log('ハイライト更新後のキャッシュ無効化完了: ' + spreadsheetId);
      }
    } catch (cacheError) {
      console.warn('ハイライト後のキャッシュ無効化エラー:', cacheError.message);
    }

    return {
      status: 'success',
      highlighted: !isHighlighted,
      message: isHighlighted ? 'ハイライトを解除しました' : 'ハイライトしました',
    };
  } catch (e) {
    console.error('ハイライト処理エラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// =================================================================
// 互換性関数（後方互換性のため）
// =================================================================

// getWebAppUrl function removed - now using the unified version from url.gs

function getHeaderIndices(spreadsheetId, sheetName) {
  console.log(
    'getHeaderIndices received in core.gs: spreadsheetId=%s, sheetName=%s',
    spreadsheetId,
    sheetName
  );

  const cacheKey = 'hdr_' + spreadsheetId + '_' + sheetName;
  const indices = getHeadersCached(spreadsheetId, sheetName);

  // 理由列が取得できていない場合はキャッシュを無効化して再取得
  if (!indices || indices[COLUMN_HEADERS.REASON] === undefined) {
    console.log('getHeaderIndices: Reason header missing, refreshing cache for key=%s', cacheKey);
    cacheManager.remove(cacheKey);
    indices = getHeadersCached(spreadsheetId, sheetName);
  }

  return indices;
}

function getSheetColumns(userId, sheetId) {
  verifyUserAccess(userId);
  try {
    const userInfo = DB.findUserById(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
    const sheet = spreadsheet.getSheetById(sheetId);

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + sheetId);
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return []; // 列がない場合
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // クライアントサイドが期待する形式に変換
    const columns = headers.map(function (headerName) {
      return {
        id: headerName,
        name: headerName,
      };
    });

    console.log('✅ getSheetColumns: Found %s columns for sheetId %s', columns.length, sheetId);
    return columns;
  } catch (e) {
    console.error('getSheetColumns エラー: ' + e.message);
    console.error('Error details:', e.stack);
    throw new Error('列の取得に失敗しました: ' + e.message);
  }
}

/**
 * GoogleフォームURLからフォームIDを抽出
 */
function extractFormIdFromUrl(url) {
  if (!url) return '';

  try {
    // Regular expression to extract form ID from Google Forms URLs
    const formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }

    // Alternative pattern for e/ URLs
    const eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
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
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);

      const service = getSheetsServiceCached();
      const headerIndices = getHeaderIndices(spreadsheetId, sheetName);

      // すべてのリアクション列を取得してユーザーの重複リアクションをチェック
      const allReactionRanges = [];
      const allReactionColumns = {};
      let targetReactionColumnIndex = null;

      // 全リアクション列の情報を準備
      REACTION_KEYS.forEach(function (key) {
        const columnName = COLUMN_HEADERS[key];
        const columnIndex = headerIndices[columnName];
        if (columnIndex !== undefined) {
          const range = "'" + sheetName + "'!" + String.fromCharCode(65 + columnIndex) + rowIndex;
          allReactionRanges.push(range);
          allReactionColumns[key] = {
            columnIndex: columnIndex,
            range: range,
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
      const response = batchGetSheetsData(service, spreadsheetId, allReactionRanges);
      const updateData = [];
      let userAction = null;
      let targetCount = 0;

      // 各リアクション列を処理
      let rangeIndex = 0;
      REACTION_KEYS.forEach(function (key) {
        if (!allReactionColumns[key]) return;

        let currentReactionString = '';
        if (
          response &&
          response.valueRanges &&
          response.valueRanges[rangeIndex] &&
          response.valueRanges[rangeIndex].values &&
          response.valueRanges[rangeIndex].values[0] &&
          response.valueRanges[rangeIndex].values[0][0]
        ) {
          currentReactionString = response.valueRanges[rangeIndex].values[0][0];
        }

        const currentReactions = parseReactionString(currentReactionString);
        const userIndex = currentReactions.indexOf(reactingUserEmail);

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
            console.log('他のリアクションから削除: ' + reactingUserEmail + ' from ' + key);
          }
        }

        // 更新データを準備
        const updatedReactionString = currentReactions.join(', ');
        updateData.push({
          range: allReactionColumns[key].range,
          values: [[updatedReactionString]],
        });

        rangeIndex++;
      });

      // すべての更新を一括実行
      if (updateData.length > 0) {
        batchUpdateSheetsData(service, spreadsheetId, updateData);
      }

      console.log(
        'リアクション切り替え完了: ' +
          reactingUserEmail +
          ' → ' +
          reactionKey +
          ' (' +
          userAction +
          ')',
        {
          updatedRanges: updateData.length,
          targetCount: targetCount,
          allColumns: Object.keys(allReactionColumns),
        }
      );

      // リアクション更新後のキャッシュ無効化
      try {
        if (
          typeof cacheManager !== 'undefined' &&
          typeof cacheManager.invalidateSheetData === 'function'
        ) {
          cacheManager.invalidateSheetData(spreadsheetId, sheetName);
          console.log('リアクション更新後のキャッシュ無効化完了: ' + spreadsheetId);
        }
      } catch (cacheError) {
        console.warn('リアクション後のキャッシュ無効化エラー:', cacheError.message);
      }

      return {
        status: 'success',
        message: 'リアクションを更新しました。',
        action: userAction,
        count: targetCount,
      };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    console.error('リアクション処理エラー: ' + e.message);
    return {
      status: 'error',
      message: 'リアクションの処理に失敗しました: ' + e.message,
    };
  }
}

// =================================================================
// フォーム作成機能
// =================================================================

/**
 * 回答ボードの公開を停止する (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// NOTE: unpublishBoard関数の重複を回避するため、config.gsの実装を使用
// function unpublishBoard(requestUserId) {
//   verifyUserAccess(requestUserId);
//   try {
//     const currentUserId = requestUserId;
//     const userInfo = DB.findUserById(currentUserId);
//     if (!userInfo) {
//       throw new Error('ユーザー情報が見つかりません');
//     }

//     const configJson = JSON.parse(userInfo.configJson || '{}');

//     configJson.publishedSpreadsheetId = '';
//     configJson.publishedSheetName = '';
//     configJson.appPublished = false; // 公開状態をfalseにする
//     configJson.setupStatus = 'completed'; // 公開停止後もセットアップは完了状態とする

//     updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
//     invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId, true);

//     console.log('✅ 回答ボードの公開を停止しました: %s', currentUserId);
//     return { status: 'success', message: '回答ボードの公開を停止しました。' };
//   } catch (e) {
//     console.error('公開停止エラー: ' + e.message);
//     return { status: 'error', message: '回答ボードの公開停止に失敗しました: ' + e.message };
//   }
// }

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
    const userEmail = options.userEmail;
    const userId = options.userId;
    const formDescription = options.formDescription || 'みんなの回答ボードへの投稿フォームです。';

    // タイムスタンプ生成（日時を含む）
    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // フォームタイトル生成
    const formTitle = options.formTitle || 'みんなの回答ボード ' + dateTimeString;

    // フォーム作成
    const form = FormApp.create(formTitle);
    form.setDescription(formDescription);

    // 基本的な質問を追加
    addUnifiedQuestions(form, options.questions || 'default', options.customConfig || {});

    // メールアドレス収集を有効化（スプレッドシート連携前に設定）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      } else {
        form.setCollectEmail(true);
      }
    } catch (emailError) {
      console.warn('Email collection setting failed:', emailError.message);
    }

    // スプレッドシート作成
    const spreadsheetResult = createLinkedSpreadsheet(userEmail, form, dateTimeString);

    // リアクション関連列を追加
    addReactionColumnsToSpreadsheet(spreadsheetResult.spreadsheetId, spreadsheetResult.sheetName);

    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      viewFormUrl: form.getPublishedUrl(),
      editFormUrl: typeof form.getEditUrl === 'function' ? form.getEditUrl() : '',
      spreadsheetId: spreadsheetResult.spreadsheetId,
      spreadsheetUrl: spreadsheetResult.spreadsheetUrl,
      sheetName: spreadsheetResult.sheetName,
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
    const config = getQuestionConfig(questionType, customConfig);

    form.setCollectEmail(false);

    if (questionType === 'simple') {
      const classItem = form.addListItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setChoiceValues(config.classQuestion.choices);
      classItem.setRequired(true);

      const nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setRequired(true);

      const mainItem = form.addTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setRequired(true);

      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      const validation = FormApp.createParagraphTextValidation()
        .requireTextLengthLessThanOrEqualTo(140)
        .build();
      reasonItem.setValidation(validation);
      reasonItem.setRequired(false);
    } else if (questionType === 'custom' && customConfig) {
      console.log('addUnifiedQuestions - custom mode with config:', JSON.stringify(customConfig));

      // クラス選択肢（有効な場合のみ）
      if (
        customConfig.enableClass &&
        customConfig.classQuestion &&
        customConfig.classQuestion.choices &&
        customConfig.classQuestion.choices.length > 0
      ) {
        const classItem = form.addListItem();
        classItem.setTitle('クラス');
        classItem.setChoiceValues(customConfig.classQuestion.choices);
        classItem.setRequired(true);
      }

      // 名前欄（常にオン）
      const nameItem = form.addTextItem();
      nameItem.setTitle('名前');
      nameItem.setRequired(false);

      // メイン質問
      const mainQuestionTitle = customConfig.mainQuestion
        ? customConfig.mainQuestion.title
        : '今回のテーマについて、あなたの考えや意見を聞かせてください';
      let mainItem;
      const questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';

      switch (questionType) {
        case 'text':
          mainItem = form.addTextItem();
          break;
        case 'multiple':
          mainItem = form.addCheckboxItem();
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.choices &&
            customConfig.mainQuestion.choices.length > 0
          ) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.includeOthers &&
            typeof mainItem.showOtherOption === 'function'
          ) {
            mainItem.showOtherOption(true);
          }
          break;
        case 'choice':
          mainItem = form.addMultipleChoiceItem();
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.choices &&
            customConfig.mainQuestion.choices.length > 0
          ) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.includeOthers &&
            typeof mainItem.showOtherOption === 'function'
          ) {
            mainItem.showOtherOption(true);
          }
          break;
        default:
          mainItem = form.addParagraphTextItem();
      }
      mainItem.setTitle(mainQuestionTitle);
      mainItem.setRequired(true);

      // 理由欄（常にオン）
      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle('そう考える理由や体験があれば教えてください。');
      reasonItem.setRequired(false);
    } else {
      const classItem = form.addTextItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setHelpText(config.classQuestion.helpText);
      classItem.setRequired(true);

      const mainItem = form.addParagraphTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setHelpText(config.mainQuestion.helpText);
      mainItem.setRequired(true);

      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      reasonItem.setRequired(false);

      const nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setHelpText(config.nameQuestion.helpText);
      nameItem.setRequired(false);
    }

    console.log('フォームに統一質問を追加しました: ' + questionType);
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
  const config = {
    classQuestion: {
      title: 'クラス',
      helpText: '',
      choices: ['クラス1', 'クラス2', 'クラス3', 'クラス4'],
    },
    nameQuestion: {
      title: '名前',
      helpText: '',
    },
    mainQuestion: {
      title: '今日の学習について、あなたの考えや感想を聞かせてください',
      helpText: '',
      choices: ['気づいたことがある。', '疑問に思うことがある。', 'もっと知りたいことがある。'],
      type: 'paragraph',
    },
    reasonQuestion: {
      title: 'そう考える理由や体験があれば教えてください。',
      helpText: '',
      type: 'paragraph',
    },
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

    return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(
      ContentService.MimeType.JSON
    );
  } catch (error) {
    console.error('doGetQuestionConfig error:', error);
    return ContentService.createTextOutput(
      JSON.stringify({
        error: 'Failed to get question config',
        details: error.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * クラス選択肢をデータベースに保存
 */
function saveClassChoices(userId, classChoices) {
  try {
    const currentUserId = userId;
    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.savedClassChoices = classChoices;
    configJson.lastClassChoicesUpdate = new Date().toISOString();

    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson),
    });

    console.log('クラス選択肢が保存されました:', classChoices);
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
    const currentUserId = userId;
    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    const savedClassChoices = configJson.savedClassChoices || [
      'クラス1',
      'クラス2',
      'クラス3',
      'クラス4',
    ];

    return {
      status: 'success',
      classChoices: savedClassChoices,
      lastUpdate: configJson.lastClassChoicesUpdate,
    };
  } catch (error) {
    console.error('クラス選択肢取得エラー:', error.message);
    return { status: 'error', message: 'クラス選択肢の取得に失敗しました: ' + error.message };
  }
}

/**
 * 統合フォーム作成関数（Phase 2: 最適化版）
 * プリセット設定を使用して重複コードを排除
 */
const FORM_PRESETS = {
  quickstart: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'custom',
    description:
      'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {
      mainQuestion: '今日の学習について、あなたの考えや感想を聞かせてください',
      questionType: 'text',
      enableClass: false,
      includeOthers: false,
    },
  },
  custom: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'custom',
    description:
      'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {}, // Will be overridden by user input
  },
  study: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'simple', // Default, can be overridden
    description:
      'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {},
  },
};

/**
 * 統合フォーム作成関数
 * @param {string} presetType - 'quickstart' | 'custom' | 'study'
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} userId - ユーザーID
 * @param {Object} overrides - プリセットを上書きする設定
 * @returns {Object} フォーム作成結果
 */
function createUnifiedForm(presetType, userEmail, userId, overrides = {}) {
  try {
    const preset = FORM_PRESETS[presetType];
    if (!preset) {
      throw new Error(`未知のプリセットタイプ: ${presetType}`);
    }

    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // タイトル生成（上書き可能）
    const titlePrefix = overrides.titlePrefix || preset.titlePrefix;
    const formTitle = overrides.formTitle || `${titlePrefix} ${dateTimeString}`;

    // 設定をマージ（プリセット + ユーザー設定）
    const mergedConfig = { ...preset.config, ...overrides.customConfig };

    const factoryOptions = {
      userEmail: userEmail,
      userId: userId,
      formTitle: formTitle,
      questions: overrides.questions || preset.questions,
      formDescription: overrides.formDescription || preset.description,
      customConfig: mergedConfig,
    };

    return createFormFactory(factoryOptions);
  } catch (error) {
    console.error(`createUnifiedForm Error (${presetType}):`, error.message);
    throw new Error(`フォーム作成に失敗しました (${presetType}): ` + error.message);
  }
}

/**
 * カスタムフォーム作成（管理パネル用）
 * @deprecated createUnifiedForm('custom', ...) を使用してください
 * 互換性のため保持、内部でcreateUnifiedFormを使用
 */
function createCustomForm(userEmail, userId, config) {
  try {
    console.warn('createCustomForm() is deprecated. Use createUnifiedForm("custom", ...) instead.');

    // AdminPanelのconfig構造を内部形式に変換
    const convertedConfig = {
      mainQuestion: {
        title: config.mainQuestion || '今日の学習について、あなたの考えや感想を聞かせてください',
        type: config.responseType || config.questionType || 'text', // responseTypeを優先して使用
        choices: config.questionChoices || config.choices || [], // questionChoicesを優先して使用
        includeOthers: config.includeOthers || false,
      },
      enableClass: config.enableClass || false,
      classQuestion: {
        choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4'],
      },
    };

    console.log('createCustomForm - converted config:', JSON.stringify(convertedConfig));

    const overrides = {
      titlePrefix: config.formTitle || 'カスタムフォーム',
      customConfig: convertedConfig,
    };

    return createUnifiedForm('custom', userEmail, userId, overrides);
  } catch (error) {
    console.error('createCustomForm Error:', error.message);
    throw new Error('カスタムフォームの作成に失敗しました: ' + error.message);
  }
}

/**
 * リンクされたスプレッドシートを作成
 * @param {string} userEmail - ユーザーメール
 * @param {GoogleAppsScript.Forms.Form} form - フォーム
 * @param {string} dateTimeString - 日付時刻文字列
 * @returns {Object} スプレッドシート情報
 */
function createLinkedSpreadsheet(userEmail, form, dateTimeString) {
  try {
    // スプレッドシート名を設定
    const spreadsheetName = userEmail + ' - 回答データ - ' + dateTimeString;

    // 新しいスプレッドシートを作成
    const spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
    const spreadsheetId = spreadsheetObj.getId();

    // スプレッドシートの共有設定を同一ドメイン閲覧可能に設定
    try {
      const file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }

      // 同一ドメインで閲覧可能に設定（教育機関対応）
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      console.log('スプレッドシートを同一ドメイン閲覧可能に設定しました: ' + spreadsheetId);

      // 作成者（現在のユーザー）は所有者として保持
      console.log('作成者は所有者として権限を保持: ' + userEmail);
    } catch (sharingError) {
      console.warn('共有設定の変更に失敗しましたが、処理を続行します: ' + sharingError.message);
    }

    // フォームの回答先としてスプレッドシートを設定
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

    // スプレッドシート名を再設定（フォーム設定後に名前が変わる場合があるため）
    spreadsheetObj.rename(spreadsheetName);

    // シート名を取得（通常は「フォームの回答 1」）
    const sheets = spreadsheetObj.getSheets();
    const sheetName = String(sheets[0].getName());
    // シート名が不正な値でないことを確認
    if (!sheetName || sheetName === 'true') {
      sheetName = 'Sheet1'; // または適切なデフォルト値
      console.warn(
        '不正なシート名が検出されました。デフォルトのシート名を使用します: ' + sheetName
      );
    }

    // サービスアカウントとスプレッドシートを共有（失敗しても処理継続）
    try {
      shareSpreadsheetWithServiceAccount(spreadsheetId);
      console.log('サービスアカウントとの共有完了: ' + spreadsheetId);
    } catch (shareError) {
      console.warn('サービスアカウント共有エラー（処理継続）:', shareError.message);
      // 権限エラーの場合でも、スプレッドシート作成自体は成功とみなす
      console.log(
        'スプレッドシート作成は完了しました。サービスアカウント共有は後で設定してください。'
      );
    }

    return {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetObj.getUrl(),
      sheetName: sheetName,
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
function shareAllSpreadsheetsWithServiceAccount() {
  try {
    console.log('全スプレッドシートのサービスアカウント共有開始');

    const allUsers = getAllUsers();
    const results = [];
    let successCount = 0;
    let errorCount = 0;

    for (var i = 0; i < allUsers.length; i++) {
      const user = allUsers[i];
      if (user.spreadsheetId && user.isActive === 'true') {
        try {
          shareSpreadsheetWithServiceAccount(user.spreadsheetId);
          results.push({
            userId: user.userId,
            adminEmail: user.adminEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'success',
          });
          successCount++;
          console.log('共有成功:', user.adminEmail, user.spreadsheetId);
        } catch (shareError) {
          results.push({
            userId: user.userId,
            adminEmail: user.adminEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'error',
            error: shareError.message,
          });
          errorCount++;
          console.error('共有失敗:', user.adminEmail, shareError.message);
        }
      }
    }

    console.log('全スプレッドシート共有完了:', successCount + '件成功', errorCount + '件失敗');

    return {
      status: 'completed',
      totalUsers: allUsers.length,
      successCount: successCount,
      errorCount: errorCount,
      results: results,
    };
  } catch (error) {
    console.error('shareAllSpreadsheetsWithServiceAccount エラー:', error.message);
    throw new Error('全スプレッドシート共有処理でエラーが発生しました: ' + error.message);
  }
}

/**
 * フォーム作成
 */
/**
 * StudyQuestフォーム作成（追加機能付き）
 * @deprecated createUnifiedForm('study', ...) を使用してください
 */
function repairUserSpreadsheetAccess(userEmail, spreadsheetId) {
  try {
    console.log('スプレッドシートアクセス権限の修復を開始: ' + userEmail + ' -> ' + spreadsheetId);

    // DriveApp経由で共有設定を変更
    let file;
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
      console.log('スプレッドシートをドメイン全体で編集可能に設定しました');
    } catch (domainSharingError) {
      console.warn('ドメイン共有設定に失敗: ' + domainSharingError.message);

      // ドメイン共有に失敗した場合は個別にユーザーを追加
      try {
        file.addEditor(userEmail);
        console.log('ユーザーを個別に編集者として追加しました: ' + userEmail);
      } catch (individualError) {
        console.error('個別ユーザー追加も失敗: ' + individualError.message);
      }
    }

    // SpreadsheetApp経由でも編集者として追加
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      spreadsheet.addEditor(userEmail);
      console.log('SpreadsheetApp経由でユーザーを編集者として追加: ' + userEmail);
    } catch (spreadsheetAddError) {
      console.warn('SpreadsheetApp経由の追加で警告: ' + spreadsheetAddError.message);
    }

    // サービスアカウントも確認
    const props = PropertiesService.getScriptProperties();
    const serviceAccountCreds = JSON.parse(
      props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS)
    );
    const serviceAccountEmail = serviceAccountCreds.client_email;

    if (serviceAccountEmail) {
      try {
        file.addEditor(serviceAccountEmail);
        console.log('サービスアカウントも編集者として追加: ' + serviceAccountEmail);
      } catch (serviceError) {
        console.warn('サービスアカウント追加で警告: ' + serviceError.message);
      }
    }

    return {
      success: true,
      message: 'スプレッドシートアクセス権限を修復しました。ドメイン全体でアクセス可能です。',
    };
  } catch (e) {
    console.error('スプレッドシートアクセス権限の修復に失敗: ' + e.message);
    return {
      success: false,
      error: e.message,
    };
  }
}

/**
 * 管理パネル専用：権限問題の緊急修復
 * @param {string} userEmail - ユーザーメール
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} 修復結果
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

    const additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT,
    ];

    // 効率的にヘッダー情報を取得
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const startCol = currentHeaders.length + 1;

    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);

    // スタイリングを一括適用
    const allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      console.warn('Auto-resize failed:', resizeError.message);
    }

    console.log('リアクション列を追加しました: ' + sheetName);
  } catch (e) {
    console.error('リアクション列追加エラー: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートの共有設定をチェックする
 * @param {string} spreadsheetId - チェックするスプレッドシートのID
 * @returns {object} status ('success' or 'error') と message
 */
function getSheetData(userId, sheetName, classFilter, sortMode, adminMode) {
  // キャッシュキー生成（ユーザー、シート、フィルタ条件ごとに個別キャッシュ）
  const cacheKey = `sheetData_${userId}_${sheetName}_${classFilter}_${sortMode}`;

  // 管理モードの場合はキャッシュをバイパス（最新データを取得）
  if (adminMode === true) {
    console.log('🔄 管理モード：シートデータキャッシュをバイパス');
    return executeGetSheetData(userId, sheetName, classFilter, sortMode);
  }

  return cacheManager.get(
    cacheKey,
    () => {
      return executeGetSheetData(userId, sheetName, classFilter, sortMode);
    },
    { ttl: 180 }
  ); // 3分間キャッシュ（短縮してデータ整合性向上）
}

/**
 * 実際のシートデータ取得処理（キャッシュ制御から分離）
 */
function executeGetSheetData(userId, sheetName, classFilter, sortMode) {
  try {
    const userInfo = getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const spreadsheetId = userInfo.spreadsheetId;
    const service = getSheetsServiceCached();

    // フォーム回答データのみを取得（名簿機能は使用しない）
    const ranges = [sheetName + '!A:Z'];

    const responses = batchGetSheetsData(service, spreadsheetId, ranges);
    console.log('DEBUG: batchGetSheetsData responses: %s', JSON.stringify(responses));
    const sheetData = responses.valueRanges[0].values || [];
    console.log('DEBUG: sheetData length: %s', sheetData.length);

    // 名簿機能は使用せず、空の配列を設定
    const rosterData = [];

    if (sheetData.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0,
      };
    }

    const headers = sheetData[0];
    const dataRows = sheetData.slice(1);

    // ヘッダーインデックスを取得（キャッシュ利用）
    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);

    // 名簿マップを作成（キャッシュ利用）
    const rosterMap = buildRosterMap(rosterData);

    // 表示モードを取得
    const configJson = JSON.parse(userInfo.configJson || '{}');
    const displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // Check if current user is the board owner
    const isOwner = configJson.ownerId === userId;
    console.log(
      'getSheetData: isOwner=%s, ownerId=%s, userId=%s',
      isOwner,
      configJson.ownerId,
      userId
    );

    // データを処理
    const processedData = dataRows.map(function (row, index) {
      return processRowData(
        row,
        headers,
        headerIndices,
        rosterMap,
        displayMode,
        index + 2,
        isOwner
      );
    });

    // フィルタリング
    let filteredData = processedData;
    if (classFilter && classFilter !== 'すべて') {
      const classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        filteredData = processedData.filter(function (row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }

    // ソート適用
    const sortedData = applySortMode(filteredData, sortMode || 'newest');

    return {
      status: 'success',
      data: sortedData,
      headers: headers,
      totalCount: sortedData.length,
      displayMode: displayMode,
    };
  } catch (e) {
    console.error('シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      headers: [],
    };
  }
}

/**
 * シート一覧取得
 */
function getSheetsList(userId) {
  try {
    console.log('getSheetsList: Start for userId:', userId);
    const userInfo = DB.findUserById(userId);
    if (!userInfo) {
      console.warn('getSheetsList: User not found:', userId);
      return [];
    }

    console.log('getSheetsList: UserInfo found:', {
      userId: userInfo.userId,
      adminEmail: userInfo.adminEmail,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl,
    });

    if (!userInfo.spreadsheetId) {
      console.warn('getSheetsList: No spreadsheet ID for user:', userId);
      return [];
    }

    console.log("getSheetsList: User's spreadsheetId:", userInfo.spreadsheetId);

    const service = getSheetsServiceCached();
    if (!service) {
      console.error('❌ getSheetsList: Sheets service not initialized');
      return [];
    }

    console.log('✅ getSheetsList: Service validated successfully');

    console.log('getSheetsList: SheetsService obtained, attempting to fetch spreadsheet data...');

    let spreadsheet;
    try {
      spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
    } catch (accessError) {
      console.warn(
        'getSheetsList: アクセスエラーを検出。サービスアカウント権限を修復中...',
        accessError.message
      );

      // サービスアカウントの権限修復を試行
      try {
        addServiceAccountToSpreadsheet(userInfo.spreadsheetId);
        console.log('getSheetsList: サービスアカウント権限を追加しました。再試行中...');

        // 少し待ってから再試行
        Utilities.sleep(1000);
        spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
      } catch (repairError) {
        console.error('getSheetsList: 権限修復に失敗:', repairError.message);

        // 最終手段：ユーザー権限での修復も試行
        try {
          const currentUserEmail = User.email();
          if (currentUserEmail === userInfo.adminEmail) {
            repairUserSpreadsheetAccess(currentUserEmail, userInfo.spreadsheetId);
            console.log('getSheetsList: ユーザー権限での修復を実行しました。');
          }
        } catch (finalRepairError) {
          console.error('getSheetsList: 最終修復も失敗:', finalRepairError.message);
        }

        return [];
      }
    }

    console.log('getSheetsList: Raw spreadsheet data:', spreadsheet);
    if (!spreadsheet) {
      console.error('getSheetsList: No spreadsheet data returned');
      return [];
    }

    if (!spreadsheet.sheets) {
      console.error(
        'getSheetsList: Spreadsheet data missing sheets property. Available properties:',
        Object.keys(spreadsheet)
      );
      return [];
    }

    if (!Array.isArray(spreadsheet.sheets)) {
      console.error('getSheetsList: sheets property is not an array:', typeof spreadsheet.sheets);
      return [];
    }

    const sheets = spreadsheet.sheets
      .map(function (sheet) {
        if (!sheet.properties) {
          console.warn('getSheetsList: Sheet missing properties:', sheet);
          return null;
        }
        return {
          name: sheet.properties.title,
          id: sheet.properties.sheetId,
        };
      })
      .filter(function (sheet) {
        return sheet !== null;
      });

    console.log('getSheetsList: Successfully returning', sheets.length, 'sheets:', sheets);
    return sheets;
  } catch (e) {
    console.error('getSheetsList: シート一覧取得エラー:', e.message);
    console.error('getSheetsList: Error details:', e.stack);
    console.error('getSheetsList: Error for userId:', userId);
    return [];
  }
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
  const processedRow = {
    rowNumber: rowNumber,
    originalData: row,
    score: 0,
    likeCount: 0,
    understandCount: 0,
    curiousCount: 0,
    isHighlighted: false,
  };

  // リアクションカウント計算
  REACTION_KEYS.forEach(function (reactionKey) {
    const columnName = COLUMN_HEADERS[reactionKey];
    const columnIndex = headerIndices[columnName];

    if (columnIndex !== undefined && row[columnIndex]) {
      const reactions = parseReactionString(row[columnIndex]);
      const count = reactions.length;

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
  const highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }

  // スコア計算
  processedRow.score = calculateRowScore(processedRow);

  // 名前の表示処理（フォーム入力の名前を使用）
  const nameIndex = headerIndices[COLUMN_HEADERS.NAME];
  if (
    nameIndex !== undefined &&
    row[nameIndex] &&
    (displayMode === DISPLAY_MODES.NAMED || isOwner)
  ) {
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
  const baseScore = 1.0;

  // いいね！による加算
  const likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;

  // その他のリアクションも軽微な加算
  const reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;

  // ハイライトによる大幅加算
  const highlightBonus = rowData.isHighlighted ? 0.5 : 0;

  // ランダム要素（同じスコアの項目をランダムに並べるため）
  const randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;

  return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
}

/**
 * データにソートを適用
 */
function applySortMode(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort(function (a, b) {
        return b.score - a.score;
      });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArray(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort(function (a, b) {
        return b.likeCount - a.likeCount;
      });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
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

/**
 * 設定された列名と実際のスプレッドシートヘッダーをマッピング
 * @param {Object} configHeaders - 設定されたヘッダー名
 * @param {Object} actualHeaderIndices - 実際のヘッダーインデックスマップ
 * @returns {Object} マッピングされたインデックス
 */
function mapConfigToActualHeaders(configHeaders, actualHeaderIndices) {
  const mappedIndices = {};
  const availableHeaders = Object.keys(actualHeaderIndices);
  console.log(
    'mapConfigToActualHeaders: Available headers in spreadsheet: %s',
    JSON.stringify(availableHeaders)
  );

  // 各設定ヘッダーでマッピングを試行
  for (var configKey in configHeaders) {
    const configHeaderName = configHeaders[configKey];
    let mappedIndex = undefined;

    console.log('mapConfigToActualHeaders: Trying to map %s = "%s"', configKey, configHeaderName);

    if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {
      // 完全一致
      mappedIndex = actualHeaderIndices[configHeaderName];
      console.log(
        'mapConfigToActualHeaders: Exact match found for %s: index %s',
        configKey,
        mappedIndex
      );
    } else if (configHeaderName) {
      // 部分一致で検索（大文字小文字を区別しない）
      const normalizedConfigName = configHeaderName.toLowerCase().trim();

      for (var actualHeader in actualHeaderIndices) {
        const normalizedActualHeader = actualHeader.toLowerCase().trim();
        if (normalizedActualHeader === normalizedConfigName) {
          mappedIndex = actualHeaderIndices[actualHeader];
          console.log(
            'mapConfigToActualHeaders: Case-insensitive match found for %s: "%s" -> index %s',
            configKey,
            actualHeader,
            mappedIndex
          );
          break;
        }
      }

      // 部分一致検索
      if (mappedIndex === undefined) {
        for (var actualHeader in actualHeaderIndices) {
          const normalizedActualHeader = actualHeader.toLowerCase().trim();
          if (
            normalizedActualHeader.includes(normalizedConfigName) ||
            normalizedConfigName.includes(normalizedActualHeader)
          ) {
            mappedIndex = actualHeaderIndices[actualHeader];
            console.log(
              'mapConfigToActualHeaders: Partial match found for %s: "%s" -> index %s',
              configKey,
              actualHeader,
              mappedIndex
            );
            break;
          }
        }
      }
    }

    // opinionHeader（メイン質問）の特別処理：見つからない場合は最も長い質問様ヘッダーを使用
    if (mappedIndex === undefined && configKey === 'opinionHeader') {
      const standardHeaders = [
        'タイムスタンプ',
        'メールアドレス',
        'クラス',
        '名前',
        '理由',
        'なるほど！',
        'いいね！',
        'もっと知りたい！',
        'ハイライト',
      ];
      const questionHeaders = [];

      for (var header in actualHeaderIndices) {
        let isStandardHeader = false;
        for (var i = 0; i < standardHeaders.length; i++) {
          if (
            header.toLowerCase().includes(standardHeaders[i].toLowerCase()) ||
            standardHeaders[i].toLowerCase().includes(header.toLowerCase())
          ) {
            isStandardHeader = true;
            break;
          }
        }

        if (!isStandardHeader && header.length > 10) {
          // 質問は通常長い
          questionHeaders.push({ header: header, index: actualHeaderIndices[header] });
        }
      }

      if (questionHeaders.length > 0) {
        // 最も長いヘッダーを選択（通常メイン質問が最も長い）
        const longestHeader = questionHeaders.reduce(function (prev, current) {
          return prev.header.length > current.header.length ? prev : current;
        });
        mappedIndex = longestHeader.index;
        console.log(
          'mapConfigToActualHeaders: Auto-detected main question header for %s: "%s" -> index %s',
          configKey,
          longestHeader.header,
          mappedIndex
        );
      }
    }

    // reasonHeader（理由列）の特別処理：見つからない場合は理由らしいヘッダーを自動検出
    if (mappedIndex === undefined && configKey === 'reasonHeader') {
      const reasonKeywords = ['理由', 'reason', 'なぜ', 'why', '根拠', 'わけ'];

      console.log(
        'mapConfigToActualHeaders: Searching for reason header with keywords: %s',
        JSON.stringify(reasonKeywords)
      );

      for (var header in actualHeaderIndices) {
        const normalizedHeader = header.toLowerCase().trim();
        for (var k = 0; k < reasonKeywords.length; k++) {
          if (
            normalizedHeader.includes(reasonKeywords[k]) ||
            reasonKeywords[k].includes(normalizedHeader)
          ) {
            mappedIndex = actualHeaderIndices[header];
            console.log(
              'mapConfigToActualHeaders: Auto-detected reason header for %s: "%s" -> index %s (keyword: %s)',
              configKey,
              header,
              mappedIndex,
              reasonKeywords[k]
            );
            break;
          }
        }
        if (mappedIndex !== undefined) break;
      }

      // より広範囲の検索（部分一致）
      if (mappedIndex === undefined) {
        for (var header in actualHeaderIndices) {
          const normalizedHeader = header.toLowerCase().trim();
          if (
            normalizedHeader.indexOf('理由') !== -1 ||
            normalizedHeader.indexOf('reason') !== -1
          ) {
            mappedIndex = actualHeaderIndices[header];
            console.log(
              'mapConfigToActualHeaders: Found reason header by partial match for %s: "%s" -> index %s',
              configKey,
              header,
              mappedIndex
            );
            break;
          }
        }
      }
    }

    mappedIndices[configKey] = mappedIndex;

    if (mappedIndex === undefined) {
      console.log(
        'mapConfigToActualHeaders: WARNING - No match found for %s = "%s"',
        configKey,
        configHeaderName
      );
    }
  }

  console.log('mapConfigToActualHeaders: Final mapping result: %s', JSON.stringify(mappedIndices));
  return mappedIndices;
}

/**
 * 特定の行のリアクション情報を取得
 */
function getRowReactions(spreadsheetId, sheetName, rowIndex, userEmail) {
  try {
    const service = getSheetsServiceCached();
    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);

    const reactionData = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false },
    };

    // 各リアクション列からデータを取得
    ['UNDERSTAND', 'LIKE', 'CURIOUS'].forEach(function (reactionKey) {
      const columnName = COLUMN_HEADERS[reactionKey];
      const columnIndex = headerIndices[columnName];

      if (columnIndex !== undefined) {
        const range = sheetName + '!' + String.fromCharCode(65 + columnIndex) + rowIndex;
        try {
          const response = batchGetSheetsData(service, spreadsheetId, [range]);
          let cellValue = '';
          if (
            response &&
            response.valueRanges &&
            response.valueRanges[0] &&
            response.valueRanges[0].values &&
            response.valueRanges[0].values[0] &&
            response.valueRanges[0].values[0][0]
          ) {
            cellValue = response.valueRanges[0].values[0][0];
          }

          if (cellValue) {
            const reactions = parseReactionString(cellValue);
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
      CURIOUS: { count: 0, reacted: false },
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
    const activeUserEmail = User.email();
    if (!activeUserEmail) {
      return {
        status: 'error',
        message: 'ユーザーが認証されていません',
      };
    }

    // 現在のユーザー情報を取得
    const userInfo = DB.findUserByEmail(activeUserEmail);
    if (!userInfo) {
      return {
        status: 'error',
        message: 'ユーザーがデータベースに登録されていません',
      };
    }

    // 編集者権限があるか確認（自分自身の状態変更も含む）
    if (!isTrue(userInfo.isActive)) {
      return {
        status: 'error',
        message: 'この操作を実行する権限がありません',
      };
    }

    // isActive状態を更新
    const newIsActiveValue = isActive ? 'true' : 'false';
    const updateResult = updateUser(userInfo.userId, {
      isActive: newIsActiveValue,
      lastAccessedAt: new Date().toISOString(),
    });

    if (updateResult.success) {
      const statusMessage = isActive
        ? 'アプリが正常に有効化されました'
        : 'アプリが正常に無効化されました';

      return {
        status: 'success',
        message: statusMessage,
        newStatus: newIsActiveValue,
      };
    } else {
      return {
        status: 'error',
        message: 'データベースの更新に失敗しました',
      };
    }
  } catch (e) {
    console.error('updateIsActiveStatus エラー: ' + e.message);
    return {
      status: 'error',
      message: 'isActive状態の更新に失敗しました: ' + e.message,
    };
  }
}

/**
 * セットアップページへのアクセス権限を確認
 * @returns {boolean} アクセス権限があればtrue
 */
function hasSetupPageAccess() {
  try {
    const activeUserEmail = User.email();
    if (!activeUserEmail) {
      return false;
    }

    // データベースに登録され、かつisActiveがtrueのユーザーのみアクセス可能
    const userInfo = DB.findUserByEmail(activeUserEmail);
    return userInfo && isTrue(userInfo.isActive);
  } catch (e) {
    console.error('hasSetupPageAccess エラー: ' + e.message);
    return false;
  }
}

/**
 * Drive APIサービスを取得
 * @returns {object} Drive APIサービス
 */
function getDriveService() {
  const accessToken = getServiceAccountTokenCached();
  return {
    accessToken: accessToken,
    baseUrl: 'https://www.googleapis.com/drive/v3',
    files: {
      get: function (params) {
        const url =
          this.baseUrl + '/files/' + params.fileId + '?fields=' + encodeURIComponent(params.fields);
        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + this.accessToken },
        });
        return JSON.parse(response.getContentText());
      },
    },
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
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    const currentUserEmail = User.email();
    return adminEmail && currentUserEmail && adminEmail === currentUserEmail;
  } catch (e) {
    console.error('isSystemAdmin エラー: ' + e.message);
    return false;
  }
}

// =================================================================
// マルチテナント対応関数（同時アクセス対応）
// =================================================================

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
      deletedUser: result.deletedUser,
    };
  } catch (error) {
    console.error('deleteUserAccountByAdmin wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message,
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
      logs: logs,
    };
  } catch (error) {
    console.error('getDeletionLogs wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

function getAllUsersForAdminForUI(requestUserId) {
  try {
    const result = getAllUsersForAdmin();
    return {
      status: 'success',
      users: result,
    };
  } catch (error) {
    console.error('getAllUsersForAdminForUI wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * カスタムフォーム作成（AdminPanel.html用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} config - フォーム設定
 */
function createForm(requestUserId, config) {
  try {
    // セキュリティチェック: ユーザー認証と入力検証
    verifyUserAccess(requestUserId);
    const activeUserEmail = User.email();

    // 入力検証
    if (!config || typeof config !== 'object') {
      throw new Error('フォーム設定が無効です');
    }

    // AdminPanelのconfig構造を内部形式に変換（createCustomForm の処理を統合）
    const convertedConfig = {
      mainQuestion: {
        title: config.mainQuestion || '今日の学習について、あなたの考えや感想を聞かせてください',
        type: config.responseType || config.questionType || 'text', // responseTypeを優先して使用
        choices: config.questionChoices || config.choices || [], // questionChoicesを優先して使用
        includeOthers: config.includeOthers || false,
      },
      enableClass: config.enableClass || false,
      classQuestion: {
        choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4'],
      },
    };

    console.log('createForm - converted config:', JSON.stringify(convertedConfig));

    const overrides = {
      titlePrefix: config.formTitle || 'カスタムフォーム',
      customConfig: convertedConfig,
    };

    const result = createUnifiedForm('custom', activeUserEmail, requestUserId, overrides);

    // ユーザー専用フォルダを作成・取得してファイルを移動
    let folder = null;
    let moveResults = { form: false, spreadsheet: false };
    try {
      console.log('📁 ユーザーフォルダ作成とファイル移動開始...');
      folder = createUserFolder(activeUserEmail);

      if (folder) {
        // フォームファイルを移動
        try {
          const formFile = DriveApp.getFileById(result.formId);
          if (formFile) {
            // 既にフォルダに存在するかチェック
            const formParents = formFile.getParents();
            let isFormAlreadyInFolder = false;

            while (formParents.hasNext()) {
              if (formParents.next().getId() === folder.getId()) {
                isFormAlreadyInFolder = true;
                break;
              }
            }

            if (!isFormAlreadyInFolder) {
              folder.addFile(formFile);
              DriveApp.getRootFolder().removeFile(formFile);
              moveResults.form = true;
              console.log('✅ カスタムフォームファイル移動完了');
            }
          }
        } catch (formMoveError) {
          console.warn('カスタムフォームファイル移動エラー:', formMoveError.message);
        }

        // スプレッドシートファイルを移動
        try {
          const ssFile = DriveApp.getFileById(result.spreadsheetId);
          if (ssFile) {
            // 既にフォルダに存在するかチェック
            const ssParents = ssFile.getParents();
            let isSsAlreadyInFolder = false;

            while (ssParents.hasNext()) {
              if (ssParents.next().getId() === folder.getId()) {
                isSsAlreadyInFolder = true;
                break;
              }
            }

            if (!isSsAlreadyInFolder) {
              folder.addFile(ssFile);
              DriveApp.getRootFolder().removeFile(ssFile);
              moveResults.spreadsheet = true;
              console.log('✅ カスタムスプレッドシートファイル移動完了');
            }
          }
        } catch (ssMoveError) {
          console.warn('カスタムスプレッドシートファイル移動エラー:', ssMoveError.message);
        }

        console.log('📁 カスタムセットアップファイル移動結果:', moveResults);
      }
    } catch (folderError) {
      console.warn('カスタムセットアップフォルダ処理エラー:', folderError.message);
    }

    // 既存ユーザーの情報を更新（スプレッドシート情報を追加）
    const existingUser = DB.findUserById(requestUserId);
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
      updatedConfigJson.folderId = folder ? folder.getId() : '';
      updatedConfigJson.folderUrl = folder ? folder.getUrl() : '';

      // カスタムフォーム設定情報を保存
      // カスタムフォーム設定情報をシート固有のキーの下に保存
      const sheetKey = 'sheet_' + result.sheetName;
      updatedConfigJson[sheetKey] = {
        ...(updatedConfigJson[sheetKey] || {}), // 既存のシート設定を保持
        formTitle: config.formTitle,
        mainQuestion: config.mainQuestion,
        questionType: config.questionType,
        choices: config.choices,
        includeOthers: config.includeOthers,
        enableClass: config.enableClass,
        classChoices: config.classChoices,
        lastModified: new Date().toISOString(),
      };

      // 以前の実行で誤ってトップレベルに追加された可能性のあるプロパティを削除
      delete updatedConfigJson.formTitle;
      delete updatedConfigJson.mainQuestion;
      delete updatedConfigJson.questionType;
      delete updatedConfigJson.choices;
      delete updatedConfigJson.includeOthers;
      delete updatedConfigJson.enableClass;
      delete updatedConfigJson.classChoices;

      // 新しく作成されたスプレッドシート情報をメインのユーザー情報として更新
      const updateData = {
        spreadsheetId: result.spreadsheetId,
        spreadsheetUrl: result.spreadsheetUrl,
        configJson: JSON.stringify(updatedConfigJson),
        lastAccessedAt: new Date().toISOString(),
      };

      console.log('createCustomFormUI - update data:', JSON.stringify(updateData));

      updateUser(requestUserId, updateData);

      // カスタムフォーム作成後の包括的キャッシュ同期（Quick Startと同様）
      console.log('🗑️ カスタムフォーム作成後の包括的キャッシュ同期中...');
      synchronizeCacheAfterCriticalUpdate(
        requestUserId,
        activeUserEmail,
        existingUser.spreadsheetId,
        result.spreadsheetId
      );
    } else {
      console.warn('createCustomFormUI - user not found:', requestUserId);
    }

    return {
      status: 'success',
      message: 'カスタムフォームが正常に作成されました！',
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      formTitle: result.formTitle,
      sheetName: result.sheetName,
      folderId: folder ? folder.getId() : '',
      folderUrl: folder ? folder.getUrl() : '',
      filesMovedToFolder: moveResults,
    };
  } catch (error) {
    console.error('createCustomFormUI error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * クイックスタート用フォーム作成（UI用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// createQuickStartFormUI - 削除済み（クイックスタート機能不要）

/**
 * @deprecated createCustomFormUIを使用してください
 */
function deleteCurrentUserAccount(requestUserId) {
  try {
    if (!requestUserId) {
      throw new Error('認証エラー: ユーザーIDが指定されていません');
    }
    verifyUserAccess(requestUserId);
    const result = deleteUserAccount(requestUserId);

    return {
      status: 'success',
      message: 'アカウントが正常に削除されました',
      result: result,
    };
  } catch (error) {
    console.error('deleteCurrentUserAccount error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
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
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    return activateSheet(requestUserId, userInfo.spreadsheetId, sheetName);
  } catch (error) {
    console.error('activateSheetSimple error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * ログインフロー修正の検証テスト
 * @returns {object} テスト結果
 */

/**
 * セキュリティ強化とキャッシュ最適化のパフォーマンステスト
 * @returns {object} パフォーマンステスト結果
 */

/**
 * 統合ログインフロー処理 - セキュリティ強化とパフォーマンス最適化
 * 従来の複数API呼び出し（getCurrentUserStatus → getExistingBoard → registerNewUser）を1回に集約
 * @returns {object} ログイン結果とリダイレクトURL
 */
function getLoginStatus() {
  try {
    const activeUserEmail = User.email();
    if (!activeUserEmail) {
      return { status: 'error', message: 'ログインユーザーの情報を取得できませんでした。' };
    }

    const cacheKey = 'login_status_' + activeUserEmail;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {
      console.warn('getLoginStatus: キャッシュ読み込みエラー -', e.message);
    }

    // 簡素化された認証：データベースから直接ユーザー情報を取得
    const userInfo = DB.findUserByEmail(activeUserEmail);

    let result;
    if (
      userInfo &&
      (userInfo.isActive === true || String(userInfo.isActive).toLowerCase() === 'true')
    ) {
      const urls = generateUserUrls(userInfo.userId);
      result = {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: 'ログインが完了しました',
      };
    } else if (userInfo) {
      const urls = generateUserUrls(userInfo.userId);
      result = {
        status: 'setup_required',
        userId: userInfo.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: 'セットアップを完了してください',
      };
    } else {
      result = { status: 'unregistered', userEmail: activeUserEmail };
    }

    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 30);
    } catch (e) {
      console.warn('getLoginStatus: キャッシュ保存エラー -', e.message);
    }

    return result;
  } catch (error) {
    console.error('getLoginStatus error:', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * Confirms registration for the active user on demand.
 * @returns {Object} Registration result
 */
function confirmUserRegistration() {
  try {
    const activeUserEmail = User.email();
    if (!activeUserEmail) {
      return { status: 'error', message: 'ユーザー情報を取得できませんでした。' };
    }
    return registerNewUser(activeUserEmail);
  } catch (error) {
    console.error('confirmUserRegistration error:', error);
    return { status: 'error', message: error.message };
  }
}

// =================================================================
// 統合API（フェーズ2最適化）
// =================================================================

/**
 * 統合初期データ取得API - OPTIMIZED
 * 従来の5つのAPI呼び出し（getCurrentUserStatus, getUserId, getAppConfig, getSheetDetails）を統合
 * @param {string} requestUserId - リクエスト元のユーザーID（省略可能）
 * @param {string} targetSheetName - 詳細を取得するシート名（省略可能）
 * @returns {Object} 統合された初期データ
 */
function getInitialData(requestUserId, targetSheetName) {
  console.log('🚀 getInitialData: 統合初期化開始', { requestUserId, targetSheetName });

  try {
    const startTime = new Date().getTime();

    // === ステップ1: ユーザー認証とユーザー情報取得（キャッシュ活用） ===
    const activeUserEmail = User.email();
    const currentUserId = requestUserId;

    // UserID の解決
    if (!currentUserId) {
      currentUserId = getUserId();
    }

    // Phase3 Optimization: Use execution-level cache to avoid duplicate database queries
    clearExecutionUserInfoCache(); // Clear any stale cache

    // ユーザー認証
    verifyUserAccess(currentUserId);
    const userInfo = getCurrentUserInfo(); // Use cached version
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // === ステップ1.5: データ整合性の自動チェックと修正 ===
    try {
      console.log('🔍 データ整合性の自動チェック開始...');
      const consistencyResult = fixUserDataConsistency(currentUserId);
      if (consistencyResult.updated) {
        console.log('✅ データ整合性が自動修正されました');
        // 修正後は最新データを再取得
        clearExecutionUserInfoCache();
        userInfo = getCurrentUserInfo();
      }
    } catch (consistencyError) {
      console.warn('⚠️ データ整合性チェック中にエラー:', consistencyError.message);
      // エラーが発生しても初期化処理は続行
    }

    // === ステップ2: 設定データの取得と自動修復 ===
    const configJson = JSON.parse(userInfo.configJson || '{}');

    // Auto-healing for inconsistent setup states
    let needsUpdate = false;
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
        userInfo.configJson = JSON.stringify(configJson);
      } catch (updateErr) {
        console.warn('Config auto-heal failed: ' + updateErr.message);
      }
    }

    // === ステップ3: シート一覧とアプリURL生成 ===
    const sheets = getSheetsList(currentUserId);
    const appUrls = generateUserUrls(currentUserId);

    // === ステップ4: 回答数とリアクション数の取得 ===
    let answerCount = 0;
    let totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        const responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (err) {
      console.warn('Answer count retrieval failed:', err.message);
    }

    // === ステップ5: セットアップステップの決定 ===
    const setupStep = determineSetupStep(userInfo, configJson);

    // 公開シート設定とヘッダー情報を取得
    const publishedSheetName = configJson.publishedSheetName || '';
    const sheetConfigKey = publishedSheetName ? 'sheet_' + publishedSheetName : '';
    const activeSheetConfig =
      sheetConfigKey && configJson[sheetConfigKey] ? configJson[sheetConfigKey] : {};

    const opinionHeader = activeSheetConfig.opinionHeader || '';
    const nameHeader = activeSheetConfig.nameHeader || '';
    const classHeader = activeSheetConfig.classHeader || '';

    // === ベース応答の構築 ===
    const response = {
      // ユーザー情報
      userInfo: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt,
        spreadsheetId: userInfo.spreadsheetId,
        spreadsheetUrl: userInfo.spreadsheetUrl,
        configJson: userInfo.configJson,
      },
      // アプリ設定
      appUrls: appUrls,
      setupStep: setupStep,
      activeSheetName: configJson.publishedSheetName || null,
      webAppUrl: appUrls.webApp,
      isPublished: !!configJson.appPublished,
      answerCount: answerCount,
      totalReactions: totalReactions,
      config: {
        publishedSheetName: publishedSheetName,
        opinionHeader: opinionHeader,
        nameHeader: nameHeader,
        classHeader: classHeader,
        showNames: configJson.showNames || false,
        showCounts: configJson.showCounts !== undefined ? configJson.showCounts : false,
        displayMode: configJson.displayMode || 'anonymous',
        setupStatus: configJson.setupStatus || 'initial',
        isPublished: !!configJson.appPublished,
      },
      // シート情報
      allSheets: sheets,
      sheetNames: sheets,
      // カスタムフォーム情報
      customFormInfo: configJson.formUrl
        ? {
            title: configJson.formTitle || 'カスタムフォーム',
            mainQuestion:
              configJson.mainQuestion || opinionHeader || configJson.publishedSheetName || '質問',
            formUrl: configJson.formUrl,
          }
        : null,
      // メタ情報
      _meta: {
        apiVersion: 'integrated_v1',
        executionTime: null,
        includedApis: ['getCurrentUserStatus', 'getUserId', 'getAppConfig'],
      },
    };

    // === ステップ6: シート詳細の取得（オプション）- 最適化版 ===
    const includeSheetDetails = targetSheetName || configJson.publishedSheetName;

    // デバッグ: シート詳細取得パラメータの確認
    console.log('🔍 getInitialData: シート詳細取得パラメータ確認:', {
      targetSheetName: targetSheetName,
      publishedSheetName: configJson.publishedSheetName,
      includeSheetDetails: includeSheetDetails,
      hasSpreadsheetId: !!userInfo.spreadsheetId,
      willIncludeSheetDetails: !!(includeSheetDetails && userInfo.spreadsheetId),
    });

    // publishedSheetNameが空の場合のフォールバック処理
    if (!includeSheetDetails && userInfo.spreadsheetId && configJson) {
      console.warn('⚠️ シート名が指定されていません。デフォルトシート名を検索中...');
      try {
        // 一般的なシート名パターンを試す
        const commonSheetNames = [
          'フォームの回答 1',
          'フォーム回答 1',
          'Form Responses 1',
          'Sheet1',
          'シート1',
        ];
        const tempService = getSheetsServiceCached();
        const spreadsheetInfo = getSpreadsheetsData(tempService, userInfo.spreadsheetId);

        if (spreadsheetInfo && spreadsheetInfo.sheets && spreadsheetInfo.sheets.length > 0) {
          // 既知のシート名から最初に見つかったものを使用
          for (const commonName of commonSheetNames) {
            if (spreadsheetInfo.sheets.some((sheet) => sheet.properties.title === commonName)) {
              includeSheetDetails = commonName;
              console.log('✅ フォールバックシート名を使用:', commonName);
              break;
            }
          }

          // それでも見つからない場合は最初のシートを使用
          if (!includeSheetDetails) {
            includeSheetDetails = spreadsheetInfo.sheets[0].properties.title;
            console.log('✅ 最初のシートを使用:', includeSheetDetails);
          }
        }
      } catch (fallbackError) {
        console.warn('⚠️ フォールバックシート名検索に失敗:', fallbackError.message);
      }
    }

    if (includeSheetDetails && userInfo.spreadsheetId) {
      try {
        // 最適化: getSheetsServiceの重複呼び出しを避けるため、一度だけ作成して再利用
        const sharedSheetsService = getSheetsServiceCached();

        // ExecutionContext を最適化版で作成（sheetsService と userInfo を渡して重複作成を回避）
        const context = createExecutionContext(currentUserId, {
          reuseService: sharedSheetsService,
          reuseUserInfo: userInfo,
        });
        const sheetDetails = getSheetDetailsFromContext(
          context,
          userInfo.spreadsheetId,
          includeSheetDetails
        );
        response.sheetDetails = sheetDetails;
        response._meta.includedApis.push('getSheetDetails');
        console.log('✅ シート詳細を統合応答に追加 (最適化版):', includeSheetDetails);
        // getInitialData内で生成したcontextの変更をコミット
        commitAllChanges(context);
      } catch (sheetErr) {
        console.warn('Sheet details retrieval failed:', sheetErr.message);
        response.sheetDetailsError = sheetErr.message;
      }
    }

    // === 実行時間の記録 ===
    const endTime = new Date().getTime();
    response._meta.executionTime = endTime - startTime;

    console.log('🎯 getInitialData: 統合初期化完了', {
      executionTime: response._meta.executionTime + 'ms',
      userId: currentUserId,
      setupStep: setupStep,
      hasSheetDetails: !!response.sheetDetails,
    });

    return response;
  } catch (error) {
    console.error('❌ getInitialData error:', error);
    return {
      status: 'error',
      message: error.message,
      _meta: {
        apiVersion: 'integrated_v1',
        error: error.message,
      },
    };
  }
}

/**
 * 手動データ整合性修正（管理パネル用）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} 修正結果
 */
function fixDataConsistencyManual(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    console.log('🔧 手動データ整合性修正実行:', requestUserId);

    const result = fixUserDataConsistency(requestUserId);

    if (result.status === 'success') {
      if (result.updated) {
        return {
          status: 'success',
          message: 'データ整合性の問題を修正しました。ページを再読み込みしてください。',
          details: result.message,
        };
      } else {
        return {
          status: 'success',
          message: 'データ整合性に問題は見つかりませんでした。',
          details: result.message,
        };
      }
    } else {
      return {
        status: 'error',
        message: 'データ整合性修正中にエラーが発生しました: ' + result.message,
      };
    }
  } catch (error) {
    console.error('❌ 手動データ整合性修正エラー:', error);
    return {
      status: 'error',
      message: '修正処理中にエラーが発生しました: ' + error.message,
    };
  }
}

/**
 * getSheetDetailsの内部実装（統合API用）
 * @param {string} requestUserId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} シート詳細
 */

/**
 * アプリケーション状態取得（UI用）
 * @returns {Object} アプリケーション状態情報
 */
function getApplicationStatusForUI() {
  try {
    const accessCheck = Access.check();
    const isEnabled = getApplicationEnabled();
    const adminEmail = User.email();

    return {
      status: 'success',
      isEnabled: isEnabled,
      isSystemAdmin: accessCheck.isSystemAdmin,
      adminEmail: adminEmail,
      lastUpdated: new Date().toISOString(),
      message: accessCheck.accessReason,
    };
  } catch (error) {
    console.error('getApplicationStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * アプリケーション状態設定（UI用）
 * @param {boolean} enabled - 有効化するかどうか
 * @returns {Object} 設定結果
 */
function setApplicationStatusForUI(enabled) {
  try {
    const result = setApplicationEnabled(enabled);
    return {
      status: 'success',
      enabled: result.enabled,
      message: result.message,
      timestamp: result.timestamp,
      adminEmail: result.adminEmail,
    };
  } catch (error) {
    console.error('setApplicationStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

// =============================================================================
// 高精度AI列判定システム
// =============================================================================

/**
 * ヘッダー配列から列マッピングを高精度で推測
 * Googleフォームやカスタムシートのヘッダーを自動判定
 * @param {Array<string>} headers - ヘッダー行の配列
 * @returns {Object} 推測された列マッピング
 */
function identifyHeaders(headers) {
  const find = (keys) => {
    const header = headers.find((h) => {
      const hStr = String(h || '').toLowerCase();
      return keys.some((k) => hStr.includes(k.toLowerCase()));
    });
    return header ? String(header) : '';
  };

  console.log('guessHeadersFromArray: Available headers:', headers);

  // Googleフォーム特有のヘッダー構造に対応
  const isGoogleForm = headers.some(
    (h) => String(h || '').includes('タイムスタンプ') || String(h || '').includes('メールアドレス')
  );

  let question = '';
  let answer = '';
  let reason = '';
  let name = '';
  let classHeader = '';

  if (isGoogleForm) {
    console.log('guessHeadersFromArray: Detected Google Form response sheet');

    // Googleフォームの一般的な構造: タイムスタンプ, メールアドレス, [質問1], [質問2], ...
    const nonMetaHeaders = headers.filter((h) => {
      const hStr = String(h || '').toLowerCase();
      return (
        !hStr.includes('タイムスタンプ') &&
        !hStr.includes('timestamp') &&
        !hStr.includes('メールアドレス') &&
        !hStr.includes('email')
      );
    });

    console.log('guessHeadersFromArray: Non-meta headers:', nonMetaHeaders);

    // より柔軟な推測ロジック
    for (let i = 0; i < nonMetaHeaders.length; i++) {
      const header = nonMetaHeaders[i];
      const hStr = String(header || '').toLowerCase();

      // 質問を含む長いテキストを質問ヘッダーとして推測
      if (
        !question &&
        (hStr.includes('だろうか') ||
          hStr.includes('ですか') ||
          hStr.includes('でしょうか') ||
          hStr.length > 20)
      ) {
        question = header;
        // 同じ内容が複数列にある場合、回答用として2番目を使用
        if (i + 1 < nonMetaHeaders.length && nonMetaHeaders[i + 1] === header) {
          answer = header;
          continue;
        }
      }

      // 回答・意見に関するヘッダー
      if (
        !answer &&
        (hStr.includes('回答') ||
          hStr.includes('答え') ||
          hStr.includes('意見') ||
          hStr.includes('考え'))
      ) {
        answer = header;
      }

      // 理由に関するヘッダー
      if (!reason && (hStr.includes('理由') || hStr.includes('詳細') || hStr.includes('説明'))) {
        reason = header;
      }

      // 名前に関するヘッダー
      if (!name && (hStr.includes('名前') || hStr.includes('氏名') || hStr.includes('学生'))) {
        name = header;
      }

      // クラスに関するヘッダー
      if (
        !classHeader &&
        (hStr.includes('クラス') || hStr.includes('組') || hStr.includes('学級'))
      ) {
        classHeader = header;
      }
    }

    // フォールバック: まだ回答が見つからない場合
    if (!answer && nonMetaHeaders.length > 0) {
      // 最初の非メタヘッダーを回答として使用
      answer = nonMetaHeaders[0];
    }
  } else {
    // 通常のシート用の推測ロジック
    question = find(['質問', '問題', 'question', 'Q']);
    answer = find(['回答', '答え', 'answer', 'A', '意見', 'opinion']);
    reason = find(['理由', 'reason', '詳細', 'detail']);
    name = find(['名前', '氏名', 'name', '学生', 'student']);
    classHeader = find(['クラス', 'class', '組', '学級']);
  }

  // リアクション列も検出（システム列は完全一致のみ）
  const understand = headers.find((h) => String(h || '').trim() === 'なるほど！') || '';
  const like = headers.find((h) => String(h || '').trim() === 'いいね！') || '';
  const curious = headers.find((h) => String(h || '').trim() === 'もっと知りたい！') || '';
  const highlight = headers.find((h) => String(h || '').trim() === 'ハイライト') || '';

  const guessed = {
    questionHeader: question,
    answerHeader: answer,
    reasonHeader: reason,
    nameHeader: name,
    classHeader: classHeader,
    understandHeader: understand,
    likeHeader: like,
    curiousHeader: curious,
    highlightHeader: highlight,
  };

  console.log('guessHeadersFromArray: Guessed headers:', guessed);

  // 最終フォールバック: 何も推測できない場合
  if (!question && !answer && headers.length > 0) {
    console.log('guessHeadersFromArray: No specific headers found, using positional mapping');

    // タイムスタンプとメールを除外して最初の列を回答として使用
    const usableHeaders = headers.filter((h) => {
      const hStr = String(h || '').toLowerCase();
      return (
        !hStr.includes('タイムスタンプ') &&
        !hStr.includes('timestamp') &&
        !hStr.includes('メールアドレス') &&
        !hStr.includes('email') &&
        String(h || '').trim() !== ''
      );
    });

    if (usableHeaders.length > 0) {
      guessed.answerHeader = usableHeaders[0];
      if (usableHeaders.length > 1) guessed.reasonHeader = usableHeaders[1];
      if (usableHeaders.length > 2) guessed.nameHeader = usableHeaders[2];
      if (usableHeaders.length > 3) guessed.classHeader = usableHeaders[3];
    }
  }

  console.log('guessHeadersFromArray: Final result:', guessed);
  return guessed;
}
