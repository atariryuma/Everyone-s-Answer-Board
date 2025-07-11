/**
 * config.gs - 軽量化版
 * 新サービスアカウントアーキテクチャで必要最小限の関数のみ
 */

const CONFIG_SHEET_NAME = 'Config';

/**
 * 現在のユーザーのスプレッドシートを取得
 * AdminPanel.htmlから呼び出される
 */
function getCurrentSpreadsheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーIDが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }
    
    return SpreadsheetApp.openById(userInfo.spreadsheetId);
  } catch (e) {
    console.error('getCurrentSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートの取得に失敗しました: ' + e.message);
  }
}

/**
 * アクティブなスプレッドシートのURLを取得
 * AdminPanel.htmlから呼び出される
 * @returns {string} スプレッドシートURL
 */
function openActiveSpreadsheet() {
  try {
    var ss = getCurrentSpreadsheet();
    return ss.getUrl();
  } catch (e) {
    console.error('openActiveSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートのURL取得に失敗しました: ' + e.message);
  }
}

/**
 * 現在のスクリプト実行ユーザーのユニークIDを取得する。
 * まずプロパティストアを確認し、なければセッション情報から取得・保存する。
 * @returns {string} 現在のユーザーのユニークID
 */
function getUserId() {
  const props = PropertiesService.getUserProperties();
  let userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    // アクティブユーザーのメールアドレスからIDを生成（より安定した方法）
    const email = Session.getActiveUser().getEmail();
    if (email) {
      // メールアドレスをハッシュ化してIDとするなど、ユニークなIDを生成
      // ここでは簡易的にメールアドレスをそのまま使っていますが、
      // 必要に応じてUtilities.computeDigest等でハッシュ化してください。
      userId = email; 
      props.setProperty('CURRENT_USER_ID', userId);
    } else {
      // それでも取得できない場合はエラー
      throw new Error('ユーザーIDを取得できませんでした。');
    }
  }
  return userId;
}

// 他の関数も同様に、存在することを確認
function getUserInfo() {
  const userId = getUserId();
  // findUserById関数を呼び出すなど、ユーザー情報を取得するロジック
  return findUserById(userId); 
}

function getSheetHeaders(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    console.warn(`シート '${sheetName}' に列が存在しません`);
    return [];
  }
  
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

/**
 * 指定されたシートの設定情報を取得する、新しい推奨関数。
 * 保存された設定があればそれを最優先し、なければ自動マッピングを試みる。
 * @param {string} sheetName - 設定を取得するシート名
 * @param {boolean} forceRefresh - キャッシュを無視して強制的に再取得するかどうか
 * @returns {object} 統一された設定オブジェクト
 */
function getConfig(sheetName, forceRefresh = false) {
  const userId = getUserId(); // ★修正点: 関数の最初にIDを取得
  const userCache = CacheService.getUserCache();
  const cacheKey = 'config_v3_' + userId + '_' + sheetName; 
  
  if (!forceRefresh) {
    const cached = userCache.get(cacheKey);
    if (cached) {
      console.log('設定キャッシュヒット: %s', cacheKey);
      return JSON.parse(cached);
    }
  }

  try {
    console.log('設定を取得中: sheetName=%s, userId=%s', sheetName, userId);
    const userInfo = getUserInfo(); // 依存関係を明確化
    const headers = getSheetHeaders(userInfo.spreadsheetId, sheetName);

    // 1. 返却する設定オブジェクトの器を準備
    let finalConfig = {
      sheetName: sheetName,
      opinionHeader: '',
      reasonHeader: '',
      nameHeader: '',
      classHeader: '',
      showNames: false,
      showCounts: true, 
      availableHeaders: headers || [],
      hasExistingConfig: false
    };

    // 2. 保存済みの設定があるか確認
    const configJson = userInfo.configJson ? JSON.parse(userInfo.configJson) : {};
    const sheetConfigKey = 'sheet_' + sheetName;
    const savedSheetConfig = configJson[sheetConfigKey];

    if (savedSheetConfig && Object.keys(savedSheetConfig).length > 0) {
      // 3.【最優先】保存済みの設定を適用する
      console.log('保存済み設定を適用します:', JSON.stringify(savedSheetConfig));
      finalConfig.hasExistingConfig = true;
      finalConfig.opinionHeader = savedSheetConfig.opinionHeader || '';
      finalConfig.reasonHeader = savedSheetConfig.reasonHeader || '';
      finalConfig.nameHeader = savedSheetConfig.nameHeader || '';
      finalConfig.classHeader = savedSheetConfig.classHeader || '';
      finalConfig.showNames = savedSheetConfig.showNames || false;
      finalConfig.showCounts = savedSheetConfig.showCounts !== undefined ? savedSheetConfig.showCounts : true;

    } else if (headers && headers.length > 0) {
      // 4.【保存設定がない場合のみ】新しい自動マッピングを実行
      console.log('保存済み設定がないため、自動マッピングを実行します。');
      const guessedConfig = autoMapHeaders(headers); 
      finalConfig.opinionHeader = guessedConfig.opinionHeader || '';
      finalConfig.reasonHeader = guessedConfig.reasonHeader || '';
      finalConfig.nameHeader = guessedConfig.nameHeader || '';
      finalConfig.classHeader = guessedConfig.classHeader || '';
    }

    // 5. 最終的な設定をキャッシュに保存
    userCache.put(cacheKey, JSON.stringify(finalConfig), 3600); // 1時間キャッシュ

    console.log('getConfig: 最終設定を返します: %s', JSON.stringify(finalConfig));
    return finalConfig;

  } catch (error) {
    console.error('getConfigでエラー:', error.message, error.stack);
    throw new Error('シート設定の取得中にエラーが発生しました: ' + error.message);
  }
}

/**
 * 列ヘッダーのリストから、各項目に最もふさわしい列名を推測する（大幅強化版）
 * @param {Array<string>} headers - スプレッドシートのヘッダーリスト
 * @param {string} sheetName - シート名（データ分析用）
 * @returns {object} 推測されたマッピング結果
 */
function autoMapHeaders(headers, sheetName = null) {
  // 高精度マッピングルール（優先度順）
  const mappingRules = {
    opinionHeader: {
      exact: ['今日のテーマについて、あなたの考えや意見を聞かせてください', 'あなたの回答・意見', '回答・意見'],
      high: ['今日のテーマ', 'あなたの考え', '意見', '回答', '答え', '質問への回答'],
      medium: ['answer', 'response', 'opinion', 'comment', '投稿', 'コメント', '内容'],
      low: ['テキスト', 'text', '記述', '入力', '自由記述']
    },
    reasonHeader: {
      exact: ['そう考える理由や体験があれば教えてください', '理由・根拠', '根拠・理由'],
      high: ['そう考える理由', '理由', '根拠', '説明'],
      medium: ['詳細', 'reason', '理由説明', '補足'],
      low: ['なぜ', 'why', '背景', '経験']
    },
    nameHeader: {
      exact: ['名前', '氏名', 'name'],
      high: ['ニックネーム', '呼び名', '表示名'],
      medium: ['nickname', 'display_name', '投稿者名'],
      low: ['ユーザー', 'user', '投稿者', '回答者']
    },
    classHeader: {
      exact: ['あなたのクラス', 'クラス名', 'クラス'],
      high: ['組', '学年', 'class'],
      medium: ['グループ', 'チーム', 'group', 'team'],
      low: ['所属', '部門', 'section', '学校']
    }
  };

  const result = {};
  const usedHeaders = new Set();
  const headerScores = {}; // ヘッダー毎のスコア記録

  // 1. ヘッダーの前処理と分析
  const processedHeaders = headers.map((header, index) => {
    const cleaned = String(header || '').trim();
    const lower = cleaned.toLowerCase();
    
    return {
      original: header,
      cleaned: cleaned,
      lower: lower,
      index: index,
      length: cleaned.length,
      wordCount: cleaned.split(/\s+/).length,
      hasJapanese: /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(cleaned),
      hasEnglish: /[A-Za-z]/.test(cleaned),
      hasNumbers: /\d/.test(cleaned),
      isMetadata: isMetadataColumn(cleaned)
    };
  });

  // 2. 各項目に対して最適なヘッダーをスコア計算で決定
  Object.keys(mappingRules).forEach(mappingKey => {
    let bestHeader = null;
    let bestScore = 0;

    processedHeaders.forEach(headerInfo => {
      if (usedHeaders.has(headerInfo.index) || headerInfo.isMetadata) {
        return;
      }

      const score = calculateHeaderScore(headerInfo, mappingRules[mappingKey], mappingKey);
      
      if (score > bestScore) {
        bestScore = score;
        bestHeader = headerInfo;
      }
    });

    if (bestHeader && bestScore > 0.3) { // 閾値以上のスコアが必要
      result[mappingKey] = bestHeader.original;
      usedHeaders.add(bestHeader.index);
      headerScores[mappingKey] = bestScore;
    } else {
      result[mappingKey] = '';
    }
  });

  // 3. 主要な回答列が見つからない場合のフォールバック処理
  if (!result.opinionHeader) {
    const candidateHeaders = processedHeaders.filter(h => 
      !usedHeaders.has(h.index) && 
      !h.isMetadata &&
      h.cleaned.length > 0
    );

    if (candidateHeaders.length > 0) {
      // 最も長い説明的なヘッダーを選択
      const fallbackHeader = candidateHeaders.reduce((best, current) => {
        const currentPriority = calculateFallbackPriority(current);
        const bestPriority = calculateFallbackPriority(best);
        return currentPriority > bestPriority ? current : best;
      });
      
      result.opinionHeader = fallbackHeader.original;
      usedHeaders.add(fallbackHeader.index);
    }
  }

  // 4. データ内容を分析して判定精度を向上（シート名が提供された場合）
  if (sheetName) {
    try {
      const contentAnalysis = analyzeColumnContent(sheetName, processedHeaders);
      result = refineResultsWithContentAnalysis(result, contentAnalysis, processedHeaders);
    } catch (e) {
      console.warn('データ内容分析でエラーが発生しましたが、処理を続行します:', e.message);
    }
  }

  console.log('高精度自動判定結果:', JSON.stringify({
    result: result,
    scores: headerScores,
    totalHeaders: headers.length,
    usedHeaders: usedHeaders.size
  }));

  return result;
}

/**
 * ヘッダーのスコア計算（キーワードマッチング + コンテキスト分析）
 */
function calculateHeaderScore(headerInfo, rules, mappingType) {
  let score = 0;
  const header = headerInfo.lower;

  // 完全一致（最高スコア）
  if (rules.exact) {
    for (const pattern of rules.exact) {
      if (header === pattern.toLowerCase()) {
        return 1.0;
      }
      if (header.includes(pattern.toLowerCase())) {
        score = Math.max(score, 0.9);
      }
    }
  }

  // 高優先度マッチング
  if (rules.high) {
    for (const pattern of rules.high) {
      const patternLower = pattern.toLowerCase();
      if (header.includes(patternLower)) {
        score = Math.max(score, 0.8 * (patternLower.length / header.length));
      }
    }
  }

  // 中優先度マッチング
  if (rules.medium) {
    for (const pattern of rules.medium) {
      const patternLower = pattern.toLowerCase();
      if (header.includes(patternLower)) {
        score = Math.max(score, 0.6 * (patternLower.length / header.length));
      }
    }
  }

  // 低優先度マッチング
  if (rules.low) {
    for (const pattern of rules.low) {
      const patternLower = pattern.toLowerCase();
      if (header.includes(patternLower)) {
        score = Math.max(score, 0.4 * (patternLower.length / header.length));
      }
    }
  }

  // コンテキスト分析によるスコア調整
  score = adjustScoreByContext(score, headerInfo, mappingType);

  return score;
}

/**
 * コンテキストに基づくスコア調整
 */
function adjustScoreByContext(score, headerInfo, mappingType) {
  // 長い説明的なヘッダーは質問項目である可能性が高い
  if (mappingType === 'opinionHeader' && headerInfo.length > 20) {
    score *= 1.2;
  }

  // 短いヘッダーは名前やクラス項目の可能性が高い
  if ((mappingType === 'nameHeader' || mappingType === 'classHeader') && headerInfo.length < 10) {
    score *= 1.1;
  }

  // 日本語が含まれる場合の調整
  if (headerInfo.hasJapanese) {
    score *= 1.1; // 日本語項目を優先
  }

  // 数字が含まれるヘッダーの調整
  if (headerInfo.hasNumbers) {
    if (mappingType === 'classHeader') {
      score *= 1.1; // クラス番号の可能性
    } else {
      score *= 0.9; // その他では若干減点
    }
  }

  return Math.min(score, 1.0); // 最大1.0に制限
}

/**
 * メタデータ列かどうかの判定
 */
function isMetadataColumn(header) {
  const metadataPatterns = [
    'タイムスタンプ', 'timestamp', 'メールアドレス', 'email', 
    'id', 'uuid', '更新日時', '作成日時', 'created_at', 'updated_at'
  ];
  
  const headerLower = header.toLowerCase();
  return metadataPatterns.some(pattern => 
    headerLower.includes(pattern.toLowerCase())
  );
}

/**
 * フォールバック用の優先度計算
 */
function calculateFallbackPriority(headerInfo) {
  let priority = 0;
  
  // 長い説明的なヘッダーを優先
  priority += headerInfo.length * 0.1;
  
  // 日本語が含まれる場合は優先
  if (headerInfo.hasJapanese) {
    priority += 10;
  }
  
  // 質問らしいパターン
  if (headerInfo.lower.includes('ください') || headerInfo.lower.includes('何') || 
      headerInfo.lower.includes('どう') || headerInfo.lower.includes('?')) {
    priority += 15;
  }
  
  return priority;
}

/**
 * 列の実際のデータ内容を分析
 */
function analyzeColumnContent(sheetName, processedHeaders) {
  try {
    const spreadsheet = getCurrentSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return {};

    const analysis = {};
    const lastRow = Math.min(sheet.getLastRow(), 21); // 最大20行まで分析
    
    processedHeaders.forEach((headerInfo, index) => {
      if (lastRow <= 1) return; // データがない場合

      try {
        const columnData = sheet.getRange(2, index + 1, lastRow - 1, 1).getValues()
          .map(row => String(row[0] || '').trim())
          .filter(value => value.length > 0);

        if (columnData.length === 0) return;

        analysis[index] = {
          avgLength: columnData.reduce((sum, val) => sum + val.length, 0) / columnData.length,
          maxLength: Math.max(...columnData.map(val => val.length)),
          hasLongText: columnData.some(val => val.length > 50),
          hasShortText: columnData.every(val => val.length < 20),
          containsReasonWords: columnData.some(val => 
            val.includes('なぜなら') || val.includes('理由') || val.includes('から') || val.includes('ので')
          ),
          containsClassPattern: columnData.some(val => 
            /^\d+[年組]/.test(val) || /^[A-Z]\d*$/.test(val)
          ),
          dataCount: columnData.length
        };
      } catch (e) {
        console.warn(`列${index + 1}の分析でエラー:`, e.message);
      }
    });

    return analysis;
  } catch (e) {
    console.warn('データ内容分析でエラー:', e.message);
    return {};
  }
}

/**
 * データ内容分析結果を使って判定結果を精緻化
 */
function refineResultsWithContentAnalysis(result, contentAnalysis, processedHeaders) {
  const refinedResult = { ...result };

  // 回答列の精緻化
  if (refinedResult.opinionHeader) {
    const opinionIndex = processedHeaders.findIndex(h => h.original === refinedResult.opinionHeader);
    const analysis = contentAnalysis[opinionIndex];
    
    if (analysis && analysis.avgLength < 10 && !analysis.hasLongText) {
      // 短いテキストばかりの場合、本当に回答列か確認
      const betterCandidate = findBetterOpinionColumn(contentAnalysis, processedHeaders, refinedResult);
      if (betterCandidate) {
        refinedResult.opinionHeader = betterCandidate;
      }
    }
  }

  // 理由列の精緻化
  if (refinedResult.reasonHeader) {
    const reasonIndex = processedHeaders.findIndex(h => h.original === refinedResult.reasonHeader);
    const analysis = contentAnalysis[reasonIndex];
    
    if (analysis && analysis.containsReasonWords) {
      // 理由を示すキーワードが含まれている場合、信頼度を上げる
      console.log('理由列の信頼度が高いことを確認しました');
    }
  }

  // クラス列の精緻化
  if (refinedResult.classHeader) {
    const classIndex = processedHeaders.findIndex(h => h.original === refinedResult.classHeader);
    const analysis = contentAnalysis[classIndex];
    
    if (analysis && analysis.containsClassPattern) {
      console.log('クラス列でクラス番号パターンを確認しました');
    }
  }

  return refinedResult;
}

/**
 * より適切な回答列を探す
 */
function findBetterOpinionColumn(contentAnalysis, processedHeaders, currentResult) {
  const usedHeaders = Object.values(currentResult).filter(v => v);
  
  for (let i = 0; i < processedHeaders.length; i++) {
    const header = processedHeaders[i];
    const analysis = contentAnalysis[i];
    
    if (usedHeaders.includes(header.original) || header.isMetadata) {
      continue;
    }
    
    if (analysis && (analysis.hasLongText || analysis.avgLength > 30)) {
      return header.original;
    }
  }
  
  return null;
}

/**
 * シート設定の保存、アクティブ化、最新ステータスの取得を一つのトランザクションで実行する統合関数
 * @param {string} spreadsheetId - 対象のスプレッドシートID
 * @param {string} sheetName - 対象のシート名
 * @param {object} config - 保存する設定オブジェクト
 * @returns {object} { success: boolean, message: string, status: object } 形式のオブジェクト
 */
function saveAndActivateSheet(spreadsheetId, sheetName, config) {
  // ★★★ここから追加：引数の検証処理★★★
  if (typeof spreadsheetId !== 'string' || !spreadsheetId) {
    throw new Error('無効なspreadsheetIdです。スプレッドシートIDは必須です。');
  }
  if (typeof sheetName !== 'string' || !sheetName) {
    throw new Error('無効なsheetNameです。シート名は必須です。');
  }
  if (typeof config !== 'object' || config === null) {
    throw new Error('無効なconfigオブジェクトです。設定オブジェクトは必須です。');
  }
  // ★★★ここまで追加★★★

  try {
    console.log('saveAndActivateSheet開始: sheetName=%s', sheetName);

    // 1. 関連キャッシュを事前に削除
    const props = PropertiesService.getUserProperties();
    const currentUserId = props.getProperty('CURRENT_USER_ID');
    if (currentUserId) {
      const userInfo = findUserById(currentUserId);
      if (userInfo) {
        invalidateUserCache(currentUserId, userInfo.adminEmail, spreadsheetId);
        console.log('saveAndActivateSheet: 事前キャッシュクリア完了');
      }
    }

    // 2. 設定を保存
    saveSheetConfig(spreadsheetId, sheetName, config);
    console.log('saveAndActivateSheet: 設定保存完了');

    // 3. シートをアクティブ化
    switchToSheet(spreadsheetId, sheetName);
    console.log('saveAndActivateSheet: シート切り替え完了');

    // 4. 表示オプションを設定
    const displayOptions = {
      showNames: !!config.showNames, // booleanに変換
      showCounts: config.showCounts !== undefined ? !!config.showCounts : true // booleanに変換、デフォルトtrue
    };
    setDisplayOptions(displayOptions);
    console.log('saveAndActivateSheet: 表示オプション設定完了');

    // 5. 最新のステータスをキャッシュを無視して取得
    const finalStatus = getStatus(true);
    console.log('saveAndActivateSheet: 統合処理完了');

    return finalStatus;

  } catch (error) {
    console.error('saveAndActivateSheetで致命的なエラー:', error.message, error.stack);
    // クライアントには分かりやすいエラーメッセージを返す
    throw new Error('設定の保存・適用中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 *【新しい推奨関数】設定を保存し、ボードを公開する統合関数
 * @param {string} sheetName - 対象のシート名
 * @param {object} config - 保存・適用する設定オブジェクト
 * @returns {object} 最新のステータスオブジェクト
 */
function saveAndPublish(sheetName, config) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 30秒待機

  try {
    console.log('saveAndPublish開始: sheetName=%s', sheetName);
    
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません。');
    }
    const spreadsheetId = userInfo.spreadsheetId;

    // 1. 関連キャッシュを事前にクリア
    invalidateUserCache(userInfo.userId, userInfo.adminEmail, spreadsheetId);
    console.log('saveAndPublish: 事前キャッシュクリア完了');

    // 2. 設定を保存
    saveSheetConfig(spreadsheetId, sheetName, config);
    console.log('saveAndPublish: 設定保存完了');

    // 3. シートをアクティブ化
    switchToSheet(spreadsheetId, sheetName);
    console.log('saveAndPublish: シート切り替え完了');

    // 4. 表示オプションを更新
    const displayOptions = {
      showNames: !!config.showNames,
      showCounts: config.showCounts !== undefined ? !!config.showCounts : true
    };
    setDisplayOptions(displayOptions);
    console.log('saveAndPublish: 表示オプション設定完了');

    // 5. 最新のステータスを強制再取得して返す
    const finalStatus = getStatus(true); // forceRefresh = true
    console.log('saveAndPublish: 統合処理完了、最新ステータスを返します。');
    
    return finalStatus;

  } catch (error) {
    console.error('saveAndPublishで致命的なエラー:', error.message, error.stack);
    throw new Error('設定の保存と公開中にサーバーエラーが発生しました: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 *【新しい推奨関数】設定を下書きとして保存する（公開はしない）
 * @param {string} sheetName - シート名
 * @param {object} config - 設定オブジェクト
 * @returns {object} 保存完了メッセージ
 */
function saveDraftConfig(sheetName, config) {
  try {
    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error('無効なsheetNameです。シート名は必須です。');
    }
    if (typeof config !== 'object' || config === null) {
      throw new Error('無効なconfigオブジェクトです。設定オブジェクトは必須です。');
    }

    console.log('saveDraftConfig開始: sheetName=%s', sheetName);

    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません。');
    }

    // 設定を保存
    saveSheetConfig(userInfo.spreadsheetId, sheetName, config);
    
    // 関連キャッシュをクリア
    invalidateUserCache(userInfo.userId, userInfo.adminEmail, userInfo.spreadsheetId);

    console.log('saveDraftConfig: 設定保存完了');

    return {
      success: true,
      message: '設定が下書きとして保存されました。'
    };

  } catch (error) {
    console.error('saveDraftConfigで致命的なエラー:', error.message, error.stack);
    throw new Error('設定の保存中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * シートを公開状態にする（設定は既に保存済みであることを前提）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} 最新のステータス
 */
function activateSheet(spreadsheetId, sheetName) {
  const startTime = Date.now();
  try {
    if (typeof spreadsheetId !== 'string' || !spreadsheetId) {
      throw new Error('無効なspreadsheetIdです。スプレッドシートIDは必須です。');
    }
    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error('無効なsheetNameです。シート名は必須です。');
    }

    console.log('activateSheet開始: sheetName=%s', sheetName);

    // 最適化: 一度だけユーザー情報を取得してキャッシュクリア
    const props = PropertiesService.getUserProperties();
    const currentUserId = props.getProperty('CURRENT_USER_ID');
    if (currentUserId) {
      // ユーザー情報を一度だけ取得してキャッシュ
      const userInfo = findUserById(currentUserId);
      if (userInfo) {
        // 最小限のキャッシュクリア（スプレッドシート関連のみ）
        const keysToRemove = [
          'config_v3_' + currentUserId + '_' + sheetName,
          'sheets_' + spreadsheetId,
          'data_' + spreadsheetId
        ];
        keysToRemove.forEach(key => cacheManager.remove(key));
        console.log('activateSheet: 最小限キャッシュクリア完了');
      }
    }

    // シートをアクティブ化（効率化）
    const switchResult = switchToSheet(spreadsheetId, sheetName);
    console.log('activateSheet: シート切り替え完了');

    // 最新のステータスを取得（キャッシュ活用）
    const finalStatus = getStatus(false);
    console.log('activateSheet: 公開処理完了');

    const executionTime = Date.now() - startTime;
    console.log('activateSheet完了: 実行時間 %dms', executionTime);
    
    return finalStatus;

  } catch (error) {
    const executionTime = Date.now() - startTime;
    console.error('activateSheetで致命的なエラー (実行時間: %dms):', executionTime, error.message, error.stack);
    throw new Error('シートの公開中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * シート固有の設定を保存
 * @param {string} sheetName - シート名
 * @param {object} cfg - 設定オブジェクト
 */
/**
 * シートヘッダーを取得
 * AdminPanel.htmlから呼び出される
 */
function getSheetHeaders(sheetName) {
  try {
    var spreadsheet = getCurrentSpreadsheet();
    
    // シート名が指定されていない場合、利用可能なシートを表示
    if (!sheetName) {
      var allSheets = spreadsheet.getSheets();
      var sheetNames = allSheets.map(function(sheet) { return sheet.getName(); });
      console.log('利用可能なシート:', sheetNames);
      throw new Error('シート名が指定されていません。利用可能なシート: ' + sheetNames.join(', '));
    }
    
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      // シートが見つからない場合、利用可能なシートを表示
      var allSheets = spreadsheet.getSheets();
      var sheetNames = allSheets.map(function(sheet) { return sheet.getName(); });
      console.log('要求されたシート「' + sheetName + '」が見つかりません');
      console.log('利用可能なシート:', sheetNames);
      
      // デフォルトシートを試行
      if (sheetNames.length > 0) {
        var firstSheet = allSheets[0];
        console.log('デフォルトシート「' + firstSheet.getName() + '」を使用します');
        sheet = firstSheet;
      } else {
        throw new Error('シートが見つかりません: ' + sheetName + '. 利用可能なシート: ' + sheetNames.join(', '));
      }
    }
    
    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return [];
    }
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    return headers.filter(String); // 空のヘッダーを除外
  } catch (e) {
    console.error('getSheetHeaders エラー: ' + e.message);
    throw new Error('シートヘッダーの取得に失敗しました: ' + e.message);
  }
}

/**
 * スプレッドシートの列名から自動的にconfig設定を推測する
 */
function autoMapSheetHeaders(sheetName, overrides) {
  try {
    var headers = getSheetHeaders(sheetName);
    if (!headers || headers.length === 0) {
      return null;
    }
    
    // 新しい高精度自動判定機能を使用
    const mappingResult = autoMapHeaders(headers, sheetName);

    // モーダル入力値による上書き
    if (overrides) {
      if (overrides.mainQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.mainQuestion.trim());
        if (match) mappingResult.opinionHeader = match;
      }
      if (overrides.reasonQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.reasonQuestion.trim());
        if (match) mappingResult.reasonHeader = match;
      }
      if (overrides.nameQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.nameQuestion.trim());
        if (match) mappingResult.nameHeader = match;
      }
      if (overrides.classQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.classQuestion.trim());
        if (match) mappingResult.classHeader = match;
      }
    }

    // 従来のフォーマットに変換（後方互換性のため）
    var mapping = {
      mainHeader: mappingResult.opinionHeader || '',
      rHeader: mappingResult.reasonHeader || '',
      nameHeader: mappingResult.nameHeader || '',
      classHeader: mappingResult.classHeader || ''
    };
    
    debugLog('高精度自動マッピング結果 for ' + sheetName + ': ' + JSON.stringify(mapping));
    return mapping;
    
  } catch (e) {
    console.error('autoMapSheetHeaders エラー: ' + e.message);
    return null;
  }
}

/**
 * スプレッドシートURLを追加してシート検出を実行
 * Unpublished.htmlから呼び出される
 */
function addSpreadsheetUrl(url) {
  try {
    var spreadsheetId = url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
    if (!spreadsheetId) {
      throw new Error('無効なスプレッドシートURLです。');
    }
    
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(spreadsheetId);
    
    // ユーザー情報にスプレッドシートIDとURLを更新
    updateUser(currentUserId, {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: url
    });
    
    // シートリストを即座に取得
    var sheets = [];
    try {
      sheets = getSheetsList(currentUserId);
      console.log('シート検出完了:', {
        spreadsheetId: spreadsheetId,
        sheetCount: sheets.length,
        sheets: sheets.map(s => s.name)
      });
    } catch (sheetError) {
      console.warn('シートリスト取得でエラー:', sheetError.message);
      // シートリスト取得に失敗してもスプレッドシート追加は成功とする
    }
    
    // キャッシュをクリアして最新の状態を確保
    cacheManager.clearByPattern(spreadsheetId);
    
    return { 
      status: 'success', 
      message: 'スプレッドシートが正常に追加されました。', 
      sheets: sheets,
      spreadsheetId: spreadsheetId,
      autoSelectFirst: sheets.length > 0 ? sheets[0].name : null,
      needsRefresh: true // UI側でのリフレッシュが必要
    };
  } catch (e) {
    console.error('addSpreadsheetUrl エラー: ' + e.message);
    throw new Error('スプレッドシートの追加に失敗しました: ' + e.message);
  }
}

/**
 * シートを切り替える
 * AdminPanel.htmlから呼び出される
 */
/**
 * アクティブシートをクリア（公開停止）
 * AdminPanel.htmlから呼び出される
 */
function clearActiveSheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // 完全な公開状態のクリア（正しいプロパティ名を使用）
    configJson.publishedSheet = ''; // 後方互換性のため残す
    configJson.publishedSheetName = ''; // 正しいプロパティ名
    configJson.publishedSpreadsheetId = ''; // スプレッドシートIDもクリア
    configJson.appPublished = false; // 公開停止
    
    // データベースを更新
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    // キャッシュを無効化して即座にUIに反映
    try {
      if (typeof invalidateUserCache === 'function') {
        invalidateUserCache(currentUserId);
      }
    } catch (cacheError) {
      console.warn('キャッシュ無効化でエラーが発生しましたが、処理を続行します:', cacheError.message);
    }
    
    console.log('回答ボードの公開を正常に停止しました - ユーザーID:', currentUserId);
    return {
      success: true,
      message: '✅ 回答ボードの公開を停止しました。',
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    console.error('clearActiveSheet エラー: ' + e.message);
    throw new Error('回答ボードの公開停止に失敗しました: ' + e.message);
  }
}

/**
 * 表示オプションを設定
 * AdminPanel.htmlから呼び出される
 */
function setDisplayOptions(options) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.showNames = options.showNames;
    configJson.showCounts = options.showCounts;
    
    // showNamesからdisplayModeを設定
    configJson.displayMode = options.showNames ? 'named' : 'anonymous';
    
    console.log('setDisplayOptions: 設定更新', {
      showNames: options.showNames,
      showCounts: options.showCounts,
      displayMode: configJson.displayMode
    });
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return '表示オプションを保存しました。';
  } catch (e) {
    console.error('setDisplayOptions エラー: ' + e.message);
    throw new Error('表示オプションの保存に失敗しました: ' + e.message);
  }
}

/**
 * 管理者権限チェック
 * Page.htmlから呼び出される
 */
function checkAdmin() {
  debugLog('checkAdmin function called.');
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var response = batchGetSheetsData(service, dbId, [sheetName + '!A:H']);
    var data = (response && response.valueRanges && response.valueRanges[0] && response.valueRanges[0].values) ? response.valueRanges[0].values : [];
    
    if (data.length === 0) {
      debugLog('checkAdmin: No data found in Users sheet or sheet is empty.');
      return false;
    }
    
    var headers = data[0];
    var emailIndex = headers.indexOf('adminEmail');
    var isActiveIndex = headers.indexOf('isActive');
    
    if (emailIndex === -1 || isActiveIndex === -1) {
      console.warn('checkAdmin: Missing required headers in Users sheet.');
      return false;
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === activeUserEmail && isTrue(data[i][isActiveIndex])) {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error('checkAdmin エラー: ' + e.message);
    return false;
  }
}

/**
 * 管理者からボードを作成
 * AdminPanel.htmlから呼び出される
 */
function createBoardFromAdmin() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userId = Utilities.getUuid(); // 新しいユーザーIDを生成
    
    // フォームとスプレッドシートを作成
    var formAndSsInfo = createStudyQuestForm(activeUserEmail, userId);
    
    // 中央データベースにユーザー情報を登録
    var initialConfig = {
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      createdAt: new Date().toISOString(),
      publishedSheet: formAndSsInfo.sheetName, // 作成時に公開シートを設定
      appPublished: true // 作成時に公開状態にする
    };
    
    var userData = {
      userId: userId,
      adminEmail: activeUserEmail,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      createdAt: new Date().toISOString(),
      configJson: JSON.stringify(initialConfig),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
    };
    
    createUser(userData);
    
    // 成功レスポンスを返す
    var appUrls = generateAppUrls(userId);
    return {
      status: 'success',
      message: '新しいボードが作成され、公開されました！',
      formUrl: formAndSsInfo.formUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      formTitle: formAndSsInfo.formTitle // フォームタイトルも返す
    };
  } catch (e) {
    console.error('createBoardFromAdmin エラー: ' + e.message);
    throw new Error('ボードの作成に失敗しました: ' + e.message);
  }
}

/**
 * 既存ボード情報を取得
 * Registration.htmlから呼び出される
 */
function getExistingBoard() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userInfo = findUserByEmail(activeUserEmail);
    
    if (userInfo && isTrue(userInfo.isActive)) {
      var appUrls = generateAppUrls(userInfo.userId);
      return {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl
      };
    } else if (userInfo && String(userInfo.isActive).toLowerCase() === 'false') {
      return {
        status: 'setup_required',
        userId: userInfo.userId
      };
    } else {
      return {
        status: 'new_user'
      };
    }
  } catch (e) {
    console.error('getExistingBoard エラー: ' + e.message);
    return { status: 'error', message: '既存ボード情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ユーザー認証を検証
 * Registration.htmlから呼び出される
 */
function verifyUserAuthentication() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      return { authenticated: true, email: email };
    } else {
      return { authenticated: false, email: null };
    }
  } catch (e) {
    console.error('verifyUserAuthentication エラー: ' + e.message);
    return { authenticated: false, email: null, error: e.message };
  }
}

/**
 * 指定されたシートのヘッダーを自動でマッピングし、その結果を返す。
 * この関数は AdminPanel.html から直接呼び出されることを想定しています。
 * @param {string} sheetName - 対象のスプレッドシート名。
 * @returns {object} 推測されたヘッダーのマッピング結果。
 */
function getGuessedHeaders(sheetName) {
  try {
    const spreadsheet = getCurrentSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません。`);
    }
    
    // ヘッダー行（1行目）を取得
    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) {
      throw new Error(`シート '${sheetName}' に列が存在しません`);
    }
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // 改良された高精度判定ロジックを呼び出し（シート名も渡してデータ分析を有効化）
    const mappingResult = autoMapHeaders(headers, sheetName);
    
    // 既存の設定も取得
    let existingConfig = null;
    try {
      existingConfig = getConfig(sheetName);
    } catch (configError) {
      console.warn('既存設定の取得に失敗しましたが、処理を続行します: ' + configError.toString());
    }
    
    // 結果にallHeadersと既存設定も含めて返す
    const result = {
      ...mappingResult,
      allHeaders: headers || [],
      existingConfig: existingConfig
    };
    
    console.log('高精度自動判定を実行しました: sheet=%s, result=%s', sheetName, JSON.stringify(result));
    return result;

  } catch (e) {
    console.error('ヘッダーの自動判定に失敗しました: ' + e.toString());
    // エラーが発生した場合は、クライアント側で処理できるようエラー情報を返す
    return { error: e.message };
  }
}

/**
 * 指定されたシートのヘッダー情報と既存設定をまとめて取得します。
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} { allHeaders: Array<string>, guessedConfig: object, existingConfig: object }
 */
function getSheetDetails(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || !sheetName) {
      throw new Error('spreadsheetIdとsheetNameは必須です');
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) {
      throw new Error(`シート '${sheetName}' に列が存在しません`);
    }
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0] || [];
    const guessed = autoMapHeaders(headers);

    let existing = {};
    try {
      existing = getConfig(sheetName, true) || {};
    } catch (e) {
      console.warn('getConfig failed in getSheetDetails:', e.message);
    }

    return {
      allHeaders: headers,
      guessedConfig: guessed,
      existingConfig: existing
    };

  } catch (error) {
    console.error('getSheetDetails error:', error.message);
    throw new Error('シート情報の取得に失敗しました: ' + error.message);
  }
}

/**
 * デバッグ用：現在のスプレッドシートとシートの状態を確認
 */
function debugCurrentSpreadsheetState() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      return { error: 'ユーザーIDが設定されていません' };
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return { error: 'ユーザー情報が見つかりません' };
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
    var sheets = spreadsheet.getSheets();
    var sheetNames = sheets.map(function(sheet) { return sheet.getName(); });
    
    return {
      userId: currentUserId,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      publishedSheet: configJson.publishedSheet,
      availableSheets: sheetNames,
      sheetCount: sheets.length,
      hasPublishedSheetInSpreadsheet: sheetNames.indexOf(configJson.publishedSheet) !== -1
    };
    
  } catch (e) {
    console.error('debugCurrentSpreadsheetState エラー: ' + e.message);
    return { error: e.message };
  }
}
