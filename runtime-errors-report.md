# Runtime Errors Report

Total: 944 undefined identifiers

## debug
- Occurrences: 620
- Files: Core.gs, autoInit.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, session-utils.gs, ulog.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Core.gs:248: ULog.debug('🔍 getSetupStep: UI状態ベースのステップ判定開始', {
  - Core.gs:262: ULog.debug('🔧 ステップ1判定: 手動停止による完全リセット', {
  - Core.gs:271: ULog.debug('🔧 ステップ1判定: データソース未設定', {

## ERROR
- Occurrences: 590
- Files: Core.gs, autoInit.gs, config.gs, constants.gs, database.gs, debugConfig.gs, lockManager.gs, main.gs, secretManager.gs, setup.gs, systemIntegrationManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, unifiedValidationSystem.gs, url.gs, workflowValidation.gs
- Samples:
  - Core.gs:34: * @param {string} [severity=UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM] エラーの重要度
  - Core.gs:34: * @param {string} [severity=UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM] エラーの重要度
  - Core.gs:35: * @param {string} [category=UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM] エラーカテゴリ

## ID
- Occurrences: 380
- Files: Core.gs, autoInit.gs, config.gs, constants.gs, database.gs, main.gs, secretManager.gs, session-utils.gs, setup.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs, url.gs
- Samples:
  - Core.gs:352: * @param {string} requestUserId - リクエスト元のユーザーID（オプション）
  - Core.gs:382: configJson.publishedSpreadsheetId = ''; // スプレッドシートIDもクリア
  - Core.gs:583: * @param {string} userId - ユーザーID

## warn
- Occurrences: 274
- Files: Core.gs, autoInit.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, session-utils.gs, ulog.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Core.gs:59: ULog.warn(`⚠️ MEDIUM SEVERITY [${context}]:`, errorInfo.message, errorInfo.metadata);
  - Core.gs:224: ULog.warn('[WARN]', message, ...args);
  - Core.gs:543: ULog.warn('configJson解析エラー:', parseError.message);

## CATEGORIES
- Occurrences: 261
- Files: Core.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, resilientExecutor.gs, session-utils.gs, setup.gs, ulog.gs, workflowValidation.gs
- Samples:
  - Core.gs:35: * @param {string} [category=UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM] エラーカテゴリ
  - Core.gs:35: * @param {string} [category=UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM] エラーカテゴリ
  - Core.gs:43: category = UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,

## push
- Occurrences: 219
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, secretManager.gs, session-utils.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, validationMigration.gs
- Samples:
  - Core.gs:495: errors.push(`必須フィールド '${field}' が未定義です`);
  - Core.gs:497: errors.push(
  - Core.gs:505: errors.push("publishedSheetNameが不正な値 'true' になっています");

## URL
- Occurrences: 175
- Files: Core.gs, config.gs, database.gs, main.gs, resilientExecutor.gs, session-utils.gs, unifiedCacheManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Core.gs:431: !hasFormUrl; // フォームURL未設定
  - Core.gs:2188: // アクティブシートのフォームURLを優先（未設定時のみグローバルにフォールバック）
  - Core.gs:2200: // フォームURLの解決失敗は致命的ではない

## toISOString
- Occurrences: 154
- Files: Core.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, secretManager.gs, session-utils.gs, systemIntegrationManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Core.gs:81: timestamp: new Date().toISOString(),
  - Core.gs:97: timestamp: new Date().toISOString(),
  - Core.gs:107: timestamp: new Date().toISOString(),

## SEVERITY
- Occurrences: 149
- Files: Core.gs, config.gs, constants.gs, main.gs, setup.gs, unifiedCacheManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:34: * @param {string} [severity=UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM] エラーの重要度
  - Core.gs:34: * @param {string} [severity=UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM] エラーの重要度
  - Core.gs:42: severity = UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,

## API
- Occurrences: 122
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, resilientExecutor.gs, secretManager.gs, session-utils.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, url.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Core.gs:3: * 主要な業務ロジックとAPI エンドポイント
  - Core.gs:1185: // 通常のアクセスを試行（SERVICE_ACCOUNTでも同じSpreadsheetApp API使用）
  - Core.gs:1224: // processReactionの戻り値から直接リアクション状態を取得（API呼び出し削減）

## stringify
- Occurrences: 105
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, secretManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, url.gs, workflowValidation.gs
- Samples:
  - Core.gs:53: ULog.error(`🚨 CRITICAL ERROR [${context}]:`, JSON.stringify(errorInfo, null, 2));
  - Core.gs:56: ULog.error(`❌ HIGH SEVERITY [${context}]:`, JSON.stringify(errorInfo, null, 2));
  - Core.gs:387: configJson: JSON.stringify(configJson),

## parse
- Occurrences: 101
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, secretManager.gs, setup.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs
- Samples:
  - Core.gs:371: const configJson = JSON.parse(userInfo.configJson || '{}');
  - Core.gs:540: config = JSON.parse(configJsonString);
  - Core.gs:888: const config = JSON.parse(userInfo.configJson || '{}');

## includes
- Occurrences: 95
- Files: Core.gs, config.gs, database.gs, debugConfig.gs, main.gs, resilientExecutor.gs, secretManager.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:510: if (config.setupStatus && !validSetupStatuses.includes(config.setupStatus)) {
  - Core.gs:5011: normalizedActualHeader.includes(normalizedConfigName) ||
  - Core.gs:5012: normalizedConfigName.includes(normalizedActualHeader)

## trim
- Occurrences: 80
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Core.gs:270: if (!userInfo || !userInfo.spreadsheetId || userInfo.spreadsheetId.trim() === '') {
  - Core.gs:293: configJson.opinionHeader.trim() !== '' && // 意見列が設定済み
  - Core.gs:295: configJson.activeSheetName.trim() !== ''; // アクティブシートが選択済み

## SYSTEM
- Occurrences: 68
- Files: Core.gs, config.gs, constants.gs, debugConfig.gs, main.gs, setup.gs, ulog.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:35: * @param {string} [category=UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM] エラーカテゴリ
  - Core.gs:43: category = UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,
  - Core.gs:344: UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM

## MAIN_MAIN_UNIFIED_CONSTANTS
- Occurrences: 68
- Files: main.gs
- Samples:
  - main.gs:131: MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
  - main.gs:131: MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
  - main.gs:132: MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE

## MEDIUM
- Occurrences: 52
- Files: Core.gs, config.gs, constants.gs, main.gs, setup.gs, workflowValidation.gs
- Samples:
  - Core.gs:34: * @param {string} [severity=UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM] エラーの重要度
  - Core.gs:42: severity = UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
  - Core.gs:58: case UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM:

## forEach
- Occurrences: 49
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, session-utils.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs
- Samples:
  - Core.gs:1360: processedRows.forEach(function (rowIndex) {
  - Core.gs:3530: REACTION_KEYS.forEach(function (key) {
  - Core.gs:3558: REACTION_KEYS.forEach(function (key) {

## remove
- Occurrences: 45
- Files: Core.gs, config.gs, database.gs, session-utils.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs
- Samples:
  - Core.gs:1013: CacheService.getScriptCache().remove('user_' + userId);
  - Core.gs:1014: CacheService.getScriptCache().remove('email_' + adminEmail);
  - Core.gs:3434: cacheManager.remove(cacheKey);

## getScriptProperties
- Occurrences: 44
- Files: Core.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, secretManager.gs, setup.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:4358: var props = PropertiesService.getScriptProperties();
  - Core.gs:5342: const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
  - Core.gs:5361: var props = PropertiesService.getScriptProperties();

## AUTH
- Occurrences: 44
- Files: debugConfig.gs, main.gs, session-utils.gs, ulog.gs, unifiedValidationSystem.gs
- Samples:
  - debugConfig.gs:25: AUTH: true, // 認証関連
  - debugConfig.gs:55: * @param {string} category - カテゴリ (CACHE, AUTH, DATABASE, UI, PERFORMANCE)
  - debugConfig.gs:121: category = 'AUTH';

## CRITICAL
- Occurrences: 43
- Files: Core.gs, autoInit.gs, config.gs, constants.gs, main.gs, systemIntegrationManager.gs, ulog.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - Core.gs:52: case UNIFIED_CONSTANTS.ERROR.SEVERITY.CRITICAL:
  - Core.gs:53: ULog.error(`🚨 CRITICAL ERROR [${context}]:`, JSON.stringify(errorInfo, null, 2));
  - autoInit.gs:72: if (securityCheck.overallStatus === 'CRITICAL') {

## tests
- Occurrences: 41
- Files: unifiedValidationSystem.gs, workflowValidation.gs
- Samples:
  - unifiedValidationSystem.gs:90: results.tests.userAuth = this._checkUserAuthentication(options.userId);
  - unifiedValidationSystem.gs:91: results.tests.sessionValid = this._checkSessionValidity(options.userId);
  - unifiedValidationSystem.gs:96: results.tests.userAccess = this._checkUserAccess(options.userId);

## getProperty
- Occurrences: 40
- Files: Core.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, secretManager.gs, session-utils.gs, setup.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:4360: props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS)
  - Core.gs:5342: const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
  - Core.gs:5362: var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);

## toLowerCase
- Occurrences: 37
- Files: Core.gs, config.gs, database.gs, main.gs, resilientExecutor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:1107: messageText.toLowerCase().indexOf('duplicate') !== -1
  - Core.gs:4820: processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  - Core.gs:4990: var normalizedConfigName = configHeaderName.toLowerCase().trim();

## manager
- Occurrences: 37
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:2004: this.manager.remove(`user_${identifier}`);
  - unifiedCacheManager.gs:2005: this.manager.remove(`userinfo_${identifier}`);
  - unifiedCacheManager.gs:2006: this.manager.remove(`unified_user_info_${identifier}`);

## toString
- Occurrences: 35
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, resilientExecutor.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, url.gs
- Samples:
  - Core.gs:1104: var messageText = (e && (e.message || e.toString())) || '';
  - Core.gs:2074: var reactionString = row.originalData[columnIndex].toString();
  - Core.gs:3958: details: error.toString(),

## map
- Occurrences: 35
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, session-utils.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, validationMigration.gs
- Samples:
  - Core.gs:1868: var processedData = rawNewData.map(function (row, idx) {
  - Core.gs:1938: return sheets.map(function (sheet) {
  - Core.gs:2016: return rawData.map(function (row, index) {

## substring
- Occurrences: 34
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, secretManager.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - Core.gs:123: value: typeof value === 'string' ? value.substring(0, 100) : String(value).substring(0, 100),
  - Core.gs:123: value: typeof value === 'string' ? value.substring(0, 100) : String(value).substring(0, 100),
  - Core.gs:2121: opinionValue.substring(0, 50),

## sleep
- Occurrences: 34
- Files: Core.gs, config.gs, constants.gs, database.gs, resilientExecutor.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:1036: Utilities.sleep(stage.delay);
  - Core.gs:3005: Utilities.sleep(500);
  - Core.gs:4151: Utilities.sleep(1000); // フォーム連携完了を待つ

## DATABASE
- Occurrences: 33
- Files: Core.gs, config.gs, constants.gs, database.gs, debugConfig.gs, main.gs, unifiedCacheManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:88: UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE,
  - Core.gs:105: category: UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE,
  - Core.gs:634: UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE

## replace
- Occurrences: 33
- Files: main.gs, session-utils.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs
- Samples:
  - main.gs:40: .replace(/\\/g, '\\\\')
  - main.gs:41: .replace(/\n/g, '\\n')
  - main.gs:42: .replace(/\r/g, '\\r')

## HIGH
- Occurrences: 31
- Files: Core.gs, config.gs, constants.gs, main.gs, setup.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:55: case UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH:
  - Core.gs:56: ULog.error(`❌ HIGH SEVERITY [${context}]:`, JSON.stringify(errorInfo, null, 2));
  - Core.gs:633: UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,

## components
- Occurrences: 31
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:133: this.components.unifiedSecretManager.status = 'INITIALIZED';
  - systemIntegrationManager.gs:134: this.components.unifiedSecretManager.instance = unifiedSecretManager;
  - systemIntegrationManager.gs:142: this.components.resilientExecutor.status = 'INITIALIZED';

## validationLevels
- Occurrences: 31
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:35: validate(category, level = this.validationLevels.STANDARD, options = {}) {
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t

## getTime
- Occurrences: 30
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, ulog.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:327: const stopTime = new Date(publishTime.getTime() + minutes * 60 * 1000);
  - Core.gs:336: Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))
  - Core.gs:336: Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))

## join
- Occurrences: 30
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:648: ULog.debug('📋 Auto-healing実行:', changes.join(', '));
  - Core.gs:758: validation.errors.join(', '),
  - Core.gs:765: ULog.info('✅ 状態遷移実行:', Object.keys(newValues).join(', '));

## DB
- Occurrences: 30
- Files: Core.gs, config.gs, constants.gs, database.gs, debugConfig.gs, ulog.gs
- Samples:
  - Core.gs:651: // DB更新失敗時は元の設定を返す
  - config.gs:1220: // 変更をDBに反映
  - config.gs:2410: // 変更をpendingUpdatesに蓄積（DB書き込みはしない）

## getContentText
- Occurrences: 30
- Files: Core.gs, database.gs, secretManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:5329: return JSON.parse(response.getContentText());
  - database.gs:1322: 'Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText()
  - database.gs:1326: return JSON.parse(response.getContentText());

## GAS
- Occurrences: 30
- Files: Core.gs, config.gs, constants.gs, database.gs, secretManager.gs, session-utils.gs, ulog.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs, url.gs, workflowValidation.gs
- Samples:
  - Core.gs:5355: * 1. GASスクリプトの編集権限を確認
  - config.gs:3485: // GAS環境ではcommitAllChanges実行をtry-catch強化で最適化
  - constants.gs:1014: // JavaScript標準関数のGAS互換ポリフィル

## clearByPattern
- Occurrences: 30
- Files: config.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:2465: cacheManager.clearByPattern(`batchGet_${context.pendingUpdates.spreadsheetId}_*`);
  - config.gs:2526: cacheManager.clearByPattern(`publishedData_${userId}_`, { strict: false, maxKeys: 200 });
  - config.gs:2527: cacheManager.clearByPattern(`sheetData_${userId}_`, { strict: false, maxKeys: 200 });

## getResponseCode
- Occurrences: 30
- Files: database.gs, resilientExecutor.gs, secretManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - database.gs:1320: if (response.getResponseCode() !== 200) {
  - database.gs:1322: 'Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText()
  - database.gs:1347: if (response.getResponseCode() !== 200) {

## CACHE
- Occurrences: 29
- Files: constants.gs, debugConfig.gs, ulog.gs
- Samples:
  - constants.gs:19: * @property {string} CACHE - キャッシュエラー
  - constants.gs:48: CACHE: 'cache',
  - constants.gs:60: CACHE: {

## checks
- Occurrences: 29
- Files: database.gs
- Samples:
  - database.gs:1831: diagnosticResult.checks.databaseConfig = {
  - database.gs:1846: diagnosticResult.checks.serviceConnection = { success: true };
  - database.gs:1848: diagnosticResult.checks.serviceConnection = {

## rgba
- Occurrences: 28
- Files: main.gs
- Samples:
  - main.gs:1943: background: rgba(31, 41, 55, 0.95);
  - main.gs:1949: box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5);
  - main.gs:1950: border: 1px solid rgba(75, 85, 99, 0.3);

## filter
- Occurrences: 27
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs, unifiedValidationSystem.gs, url.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Core.gs:1380: success: batchResults.filter((r) => r.success === true).length,
  - Core.gs:1386: successCount: batchResults.filter((r) => r.success === true).length,
  - Core.gs:4549: filteredData = processedData.filter(function (row) {

## DEBUG_MODE
- Occurrences: 26
- Files: Core.gs, constants.gs, debugConfig.gs, main.gs
- Samples:
  - Core.gs:5341: // PropertiesServiceでDEBUG_MODEが有効に設定されているかをチェック
  - Core.gs:5342: const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
  - constants.gs:228: DEBUG_MODE: 'DEBUG_MODE',

## getScriptCache
- Occurrences: 25
- Files: Core.gs, config.gs, database.gs, main.gs, session-utils.gs, unifiedCacheManager.gs, unifiedUserManager.gs, url.gs, workflowValidation.gs
- Samples:
  - Core.gs:1013: CacheService.getScriptCache().remove('user_' + userId);
  - Core.gs:1014: CacheService.getScriptCache().remove('email_' + adminEmail);
  - Core.gs:1095: CacheService.getScriptCache().put('user_' + userId, JSON.stringify(createdUser), 600); // 10分キャッシュ

## HTML
- Occurrences: 25
- Files: Core.gs, constants.gs, main.gs, session-utils.gs, unifiedUtilities.gs
- Samples:
  - Core.gs:2413: // HTML依存関数（UI連携）
  - constants.gs:510: * HTMLからアクセス可能な定数のサブセットを返す
  - main.gs:7: * HTML ファイルを読み込む include ヘルパー

## freeze
- Occurrences: 25
- Files: constants.gs
- Samples:
  - constants.gs:247: Object.freeze(UNIFIED_CONSTANTS);
  - constants.gs:248: Object.freeze(UNIFIED_CONSTANTS.ERROR);
  - constants.gs:249: Object.freeze(UNIFIED_CONSTANTS.ERROR.SEVERITY);

## UI
- Occurrences: 24
- Files: Core.gs, config.gs, constants.gs, debugConfig.gs, main.gs, ulog.gs
- Samples:
  - Core.gs:248: ULog.debug('🔍 getSetupStep: UI状態ベースのステップ判定開始', {
  - Core.gs:290: // Step 2 vs Step 3: UI設定状態に基づく判定
  - Core.gs:2413: // HTML依存関数（UI連携）

## indexOf
- Occurrences: 24
- Files: Core.gs, config.gs, database.gs, main.gs
- Samples:
  - Core.gs:1106: messageText.indexOf('既に登録されています') !== -1 ||
  - Core.gs:1107: messageText.toLowerCase().indexOf('duplicate') !== -1
  - Core.gs:2078: reacted = reactions.indexOf(currentUserEmail) !== -1;

## getId
- Occurrences: 24
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:2782: if (formParents.next().getId() === folder.getId()) {
  - Core.gs:2782: if (formParents.next().getId() === folder.getId()) {
  - Core.gs:2789: ULog.debug('📝 フォームファイルを移動中: %s → %s', formFile.getId(), folder.getName());

## HEADERS
- Occurrences: 24
- Files: constants.gs, database.gs
- Samples:
  - constants.gs:116: HEADERS: [
  - constants.gs:131: HEADERS: [
  - constants.gs:144: HEADERS: [

## setTitle
- Occurrences: 23
- Files: Core.gs, config.gs, main.gs
- Samples:
  - Core.gs:2643: form.setTitle(title);
  - Core.gs:3767: classItem.setTitle(config.classQuestion.title);
  - Core.gs:3772: nameItem.setTitle(config.nameQuestion.title);

## SHEET_NAME
- Occurrences: 23
- Files: constants.gs, database.gs, unifiedUserManager.gs
- Samples:
  - constants.gs:319: SHEET_NAME: 'Users',
  - constants.gs:334: SHEET_NAME: 'DeleteLogs',
  - constants.gs:340: SHEET_NAME: 'DiagnosticLogs',

## INFO
- Occurrences: 22
- Files: Core.gs, autoInit.gs, constants.gs, debugConfig.gs, ulog.gs, unifiedUtilities.gs, workflowValidation.gs
- Samples:
  - Core.gs:232: ULog.info('[INFO]', message, ...args);
  - autoInit.gs:27: logLevel: 'INFO',
  - constants.gs:234: LOG_LEVEL: 'INFO',

## customConfig
- Occurrences: 22
- Files: Core.gs
- Samples:
  - Core.gs:3792: customConfig.enableClass &&
  - Core.gs:3793: customConfig.classQuestion &&
  - Core.gs:3794: customConfig.classQuestion.choices &&

## startsWith
- Occurrences: 21
- Files: config.gs, constants.gs, main.gs, secretManager.gs, session-utils.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs
- Samples:
  - config.gs:399: key.startsWith('CURRENT_USER_ID_') &&
  - config.gs:1788: const sheetConfigKeys = Object.keys(configJson).filter((key) => key.startsWith('sheet_'));
  - config.gs:3341: 利用可能な設定: Object.keys(configJson).filter((k) => k.startsWith('sheet_')),

## removeAll
- Occurrences: 21
- Files: config.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:422: // removeAll() はキー配列が必要なため、自動期限切れを利用
  - config.gs:1849: CacheService.getScriptCache().removeAll([
  - session-utils.gs:30: * - removeAll() がサポートされる環境では全面削除

## OK
- Occurrences: 21
- Files: constants.gs, secretManager.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs
- Samples:
  - constants.gs:1117: const performBasicHealthCheck = () => ({ status: 'ok', message: 'Basic health OK' });
  - secretManager.gs:609: results.secretManagerStatus = 'OK';
  - secretManager.gs:622: results.propertiesServiceStatus = 'OK';

## XFrameOptionsMode
- Occurrences: 21
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:1499: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
  - main.gs:1500: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  - main.gs:1519: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {

## getName
- Occurrences: 20
- Files: Core.gs, config.gs, database.gs
- Samples:
  - Core.gs:1422: return sheets[0].getName();
  - Core.gs:1793: ULog.debug('DEBUG: Spreadsheet object obtained: %s', ss ? ss.getName() : 'null');
  - Core.gs:1796: ULog.debug('DEBUG: Sheet object obtained: %s', sheet ? sheet.getName() : 'null');

## slice
- Occurrences: 20
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs
- Samples:
  - Core.gs:2033: row.originalData ? JSON.stringify(row.originalData.slice(0, 5)) : 'undefined'
  - Core.gs:2397: data: values.slice(1),
  - Core.gs:4471: var dataRows = sheetData.slice(1);

## criticalIssues
- Occurrences: 20
- Files: database.gs, unifiedSecurityManager.gs
- Samples:
  - database.gs:1837: diagnosticResult.summary.criticalIssues.push('DATABASE_SPREADSHEET_ID が設定されていません');
  - database.gs:1852: diagnosticResult.summary.criticalIssues.push(
  - database.gs:1877: diagnosticResult.summary.criticalIssues.push(

## AdminPanel
- Occurrences: 19
- Files: Core.gs, config.gs, main.gs
- Samples:
  - Core.gs:1965: * AdminPanel.htmlから呼び出される
  - Core.gs:2211: appPublished: configJson.appPublished || false, // AdminPanel.htmlで使用される
  - Core.gs:2213: allSheets: sheets, // AdminPanel.htmlで使用される

## unifiedResult
- Occurrences: 19
- Files: validationMigration.gs
- Samples:
  - validationMigration.gs:434: isValid: unifiedResult.success,
  - validationMigration.gs:435: errors: unifiedResult.success ? [] : extractErrorMessages(unifiedResult)
  - validationMigration.gs:440: hasAccess: unifiedResult.success,

## constants
- Occurrences: 18
- Files: Core.gs, config.gs, database.gs, main.gs
- Samples:
  - Core.gs:13: // ERROR_SEVERITY is defined in constants.gs
  - Core.gs:18: // ERROR_CATEGORIES is defined in constants.gs
  - config.gs:13: // EXECUTION_MAX_LIFETIME is defined in constants.gs

## test
- Occurrences: 18
- Files: Core.gs, config.gs, constants.gs, database.gs, secretManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:186: return retryablePatterns.some((pattern) => pattern.test(coreErrorMessage));
  - config.gs:307: if (!spreadsheetIdPattern.test(userInfo.spreadsheetId)) {
  - config.gs:605: hasJapanese: /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(cleaned),

## WARN
- Occurrences: 18
- Files: Core.gs, constants.gs, debugConfig.gs, ulog.gs, unifiedUtilities.gs, validationMigration.gs
- Samples:
  - Core.gs:224: ULog.warn('[WARN]', message, ...args);
  - constants.gs:599: console.warn(`[WARN] ${functionName}: ${error.message}`, errorInfo);
  - constants.gs:824: logUnified('WARN', functionName, `試行${attempt}失敗、${waitTime}ms後に再試行`, {

## cacheError
- Occurrences: 18
- Files: Core.gs, config.gs, database.gs, main.gs, session-utils.gs, url.gs
- Samples:
  - Core.gs:1099: ULog.warn('registerNewUser: キャッシュ設定でエラー:', cacheError.message);
  - Core.gs:3404: ULog.warn('ハイライト後のキャッシュ無効化エラー:', cacheError.message);
  - Core.gs:3636: ULog.warn('リアクション後のキャッシュ無効化エラー:', cacheError.message);

## SHEETS
- Occurrences: 18
- Files: constants.gs
- Samples:
  - constants.gs:112: SHEETS: {
  - constants.gs:258: Object.freeze(UNIFIED_CONSTANTS.SHEETS);
  - constants.gs:259: Object.freeze(UNIFIED_CONSTANTS.SHEETS.DATABASE);

## monitoringResult
- Occurrences: 18
- Files: database.gs
- Samples:
  - database.gs:2842: alertMessage += '発生時刻: ' + monitoringResult.timestamp + '\n';
  - database.gs:2843: alertMessage += 'システム状態: ' + monitoringResult.summary.overallHealth + '\n\n';
  - database.gs:2845: if (monitoringResult.summary.alerts.length > 0) {

## circuitBreaker
- Occurrences: 18
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:55: if (this.circuitBreaker.state === 'OPEN') {
  - resilientExecutor.gs:56: const timeSinceLastFailure = Date.now() - this.circuitBreaker.lastFailureTime;
  - resilientExecutor.gs:57: if (timeSinceLastFailure < this.circuitBreaker.openTimeout) {

## getEmail
- Occurrences: 17
- Files: Core.gs, config.gs, secretManager.gs, setup.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:1438: * リクエストを投げたユーザー (Session.getActiveUser().getEmail()) が、
  - config.gs:1991: const activeUserEmail = Session.getActiveUser().getEmail();
  - config.gs:2045: const activeUserEmail = Session.getActiveUser().getEmail();

## deleteProperty
- Occurrences: 17
- Files: config.gs, constants.gs, database.gs, main.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:390: props.deleteProperty('CURRENT_USER_ID');
  - config.gs:411: props.deleteProperty(key);
  - constants.gs:1165: props.deleteProperty(sessionKey);

## setProperty
- Occurrences: 17
- Files: config.gs, constants.gs, main.gs, secretManager.gs, session-utils.gs, setup.gs, unifiedCacheManager.gs, unifiedUserManager.gs, workflowValidation.gs
- Samples:
  - config.gs:3849: props.setProperty('APPLICATION_ENABLED', enabledValue);
  - constants.gs:1150: props.setProperty(sessionKey, JSON.stringify(newSession));
  - constants.gs:1177: props.setProperty(sessionKey, JSON.stringify(session));

## ALLOWALL
- Occurrences: 17
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:1499: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
  - main.gs:1500: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  - main.gs:2074: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {

## memoCache
- Occurrences: 17
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:82: this.memoCache.set(key, { value: newValue, createdAt: Date.now(), ttl });
  - unifiedCacheManager.gs:132: if (enableMemoization && this.memoCache.has(key)) {
  - unifiedCacheManager.gs:134: const memoEntry = this.memoCache.get(key);

## VALIDATION
- Occurrences: 16
- Files: Core.gs, config.gs, constants.gs, main.gs, ulog.gs, workflowValidation.gs
- Samples:
  - Core.gs:132: UNIFIED_CONSTANTS.ERROR.CATEGORIES.VALIDATION,
  - config.gs:3678: UNIFIED_CONSTANTS.ERROR.CATEGORIES.VALIDATION
  - config.gs:3697: UNIFIED_CONSTANTS.ERROR.CATEGORIES.VALIDATION

## some
- Occurrences: 16
- Files: Core.gs, config.gs, database.gs, resilientExecutor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:186: return retryablePatterns.some((pattern) => pattern.test(coreErrorMessage));
  - Core.gs:6646: if (spreadsheetInfo.sheets.some((sheet) => sheet.properties.title === commonName)) {
  - config.gs:786: return metadataPatterns.some((pattern) => headerLower.includes(pattern.toLowerCase()));

## openById
- Occurrences: 16
- Files: Core.gs, config.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:1186: return SpreadsheetApp.openById(spreadsheetId);
  - Core.gs:1190: return SpreadsheetApp.openById(spreadsheetId);
  - Core.gs:2640: var form = FormApp.openById(formId);

## getRange
- Occurrences: 16
- Files: Core.gs, config.gs, unifiedSheetDataManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:1833: // getRange(row, column, numRows, numColumns)
  - Core.gs:1836: var rawNewData = sheet.getRange(startRowToRead, 1, numRowsToRead, lastColumn).getValues();
  - Core.gs:1866: var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];

## AI
- Occurrences: 16
- Files: Core.gs
- Samples:
  - Core.gs:4482: // シート固有の設定を取得（最新のAI判定結果を反映）
  - Core.gs:4486: // AI判定結果またはguessedConfigがある場合、それを優先使用
  - Core.gs:4562: // AI判定結果またはシート設定からメイン質問を特定

## getUserProperties
- Occurrences: 16
- Files: config.gs, constants.gs, database.gs, main.gs, session-utils.gs
- Samples:
  - config.gs:387: const props = PropertiesService.getUserProperties();
  - constants.gs:1140: const props = PropertiesService.getUserProperties();
  - database.gs:34: return PropertiesService.getUserProperties();

## TTL
- Occurrences: 16
- Files: constants.gs, database.gs, unifiedCacheManager.gs
- Samples:
  - constants.gs:61: /** @type {Object} TTL設定（秒） */
  - constants.gs:62: TTL: {
  - constants.gs:252: Object.freeze(UNIFIED_CONSTANTS.CACHE.TTL);

## recommendations
- Occurrences: 16
- Files: database.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - database.gs:1956: diagnosticResult.summary.recommendations.push('クリティカルな問題を解決してください');
  - database.gs:1959: diagnosticResult.summary.recommendations.push('キャッシュをクリアすることを推奨します');
  - database.gs:2819: securityResult.recommendations.push('セキュリティ設定は適切です');

## COMPREHENSIVE
- Occurrences: 16
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:22: COMPREHENSIVE: 'comprehensive'
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t
  - unifiedValidationSystem.gs:95: if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {

## max
- Occurrences: 15
- Files: Core.gs, config.gs, database.gs, main.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:334: remainingMinutes: Math.max(
  - Core.gs:2557: return Math.max(0, lastRow - 1);
  - Core.gs:2563: return Math.max(0, lastRow - 1);

## isArray
- Occurrences: 15
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs
- Samples:
  - Core.gs:1269: if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
  - Core.gs:4735: if (!Array.isArray(spreadsheet.sheets)) {
  - config.gs:3664: isArray: Array.isArray(newConfig),

## gradient
- Occurrences: 15
- Files: main.gs
- Samples:
  - main.gs:1934: background: linear-gradient(135deg, #1e3a8a 0%, #7c3aed 100%);
  - main.gs:1966: background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);
  - main.gs:2477: background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);

## constructor
- Occurrences: 14
- Files: Core.gs, config.gs, constants.gs, resilientExecutor.gs, secretManager.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:24: constructor() {
  - config.gs:22: constructor(requestUserId, userInfo) {
  - config.gs:2636: contextConstructor: context && context.constructor && context.constructor.name,

## getActiveUser
- Occurrences: 14
- Files: Core.gs, config.gs, secretManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:1438: * リクエストを投げたユーザー (Session.getActiveUser().getEmail()) が、
  - config.gs:1991: const activeUserEmail = Session.getActiveUser().getEmail();
  - config.gs:2045: const activeUserEmail = Session.getActiveUser().getEmail();

## Z0
- Occurrences: 14
- Files: Core.gs, config.gs, constants.gs, database.gs, main.gs, secretManager.gs
- Samples:
  - Core.gs:3492: var formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
  - Core.gs:3498: var eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
  - config.gs:306: const spreadsheetIdPattern = /^[a-zA-Z0-9-_]{44}$/;

## mainQuestion
- Occurrences: 14
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3776: mainItem.setTitle(config.mainQuestion.title);
  - Core.gs:3810: ? customConfig.mainQuestion.title
  - Core.gs:3813: var questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';

## entries
- Occurrences: 13
- Files: Core.gs, config.gs, constants.gs, systemIntegrationManager.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - Core.gs:493: for (const [field, expectedType] of Object.entries(requiredFields)) {
  - config.gs:396: for (const [key, value] of Object.entries(allProperties)) {
  - constants.gs:717: for (const [field, validators] of Object.entries(rules.fields)) {

## put
- Occurrences: 13
- Files: Core.gs, config.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, url.gs, workflowValidation.gs
- Samples:
  - Core.gs:1095: CacheService.getScriptCache().put('user_' + userId, JSON.stringify(createdUser), 600); // 10分キャッシュ
  - Core.gs:1096: CacheService.getScriptCache().put('email_' + adminEmail, JSON.stringify(createdUser), 600);
  - Core.gs:6226: CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), cacheTtl);

## getSheetByName
- Occurrences: 13
- Files: Core.gs, config.gs, database.gs, main.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:1606: var sheet = spreadsheet.getSheetByName(publishedSheetName);
  - Core.gs:1795: var sheet = ss.getSheetByName(publishedSheetName);
  - Core.gs:2552: const sheet = openSpreadsheetOptimized(spreadsheetId).getSheetByName(sheetName);

## fromCharCode
- Occurrences: 13
- Files: Core.gs, config.gs, database.gs, secretManager.gs
- Samples:
  - Core.gs:1614: 範囲: `A1:${String.fromCharCode(64 + lastColumn)}${lastRow}`,
  - Core.gs:3376: var range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
  - Core.gs:3534: var range = "'" + sheetName + "'!" + String.fromCharCode(65 + columnIndex) + rowIndex;

## getValues
- Occurrences: 13
- Files: Core.gs, config.gs, unifiedSheetDataManager.gs, workflowValidation.gs
- Samples:
  - Core.gs:1836: var rawNewData = sheet.getRange(startRowToRead, 1, numRowsToRead, lastColumn).getValues();
  - Core.gs:1866: var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  - Core.gs:2566: const values = sheet.getRange(2, classIndex + 1, lastRow - 1, 1).getValues();

## setRequired
- Occurrences: 13
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3769: classItem.setRequired(true);
  - Core.gs:3773: nameItem.setRequired(true);
  - Core.gs:3777: mainItem.setRequired(true);

## SERVICE_ACCOUNT_CREDS
- Occurrences: 13
- Files: Core.gs, constants.gs, database.gs, main.gs, secretManager.gs, setup.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:4360: props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS)
  - constants.gs:225: SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  - constants.gs:225: SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',

## googleapis
- Occurrences: 13
- Files: Core.gs, database.gs, secretManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:5321: baseUrl: 'https://www.googleapis.com/drive/v3',
  - database.gs:1298: baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets',
  - database.gs:1303: 'https://sheets.googleapis.com/v4/spreadsheets/' +

## ADMIN_EMAIL
- Occurrences: 13
- Files: Core.gs, database.gs, main.gs, setup.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:5362: var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
  - database.gs:2670: SCRIPT_PROPS_KEYS.ADMIN_EMAIL,
  - database.gs:2790: var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);

## headerInfo
- Occurrences: 13
- Files: config.gs
- Samples:
  - config.gs:618: if (usedHeaders.has(headerInfo.index) || headerInfo.isMetadata) {
  - config.gs:618: if (usedHeaders.has(headerInfo.index) || headerInfo.isMetadata) {
  - config.gs:687: const header = headerInfo.lower;

## DATABASE_SPREADSHEET_ID
- Occurrences: 13
- Files: config.gs, constants.gs, database.gs, main.gs, secretManager.gs, setup.gs, unifiedSecurityManager.gs
- Samples:
  - config.gs:2563: const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
  - constants.gs:226: DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  - constants.gs:226: DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',

## Core
- Occurrences: 13
- Files: constants.gs, debugConfig.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, url.gs
- Samples:
  - constants.gs:313: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead (defined in Core.gs) */
  - constants.gs:315: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION instead (defined in Core.gs) */
  - constants.gs:1098: // UnifiedErrorHandler は Core.gs で定義済みのため削除

## unifiedUserManager
- Occurrences: 13
- Files: database.gs, main.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - database.gs:441: // getAllUsersForAdmin() は unifiedUserManager.gs に統合済み
  - database.gs:751: // findUserById(), findUserByIdForViewer(), findUserByIdFresh() は unifiedUserManager.gs に統合済み
  - database.gs:822: // → unifiedUserManager.gsの実装を使用してください

## addMetaTag
- Occurrences: 13
- Files: main.gs
- Samples:
  - main.gs:2335: htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  - main.gs:2337: ULog.warn('⚠️ addMetaTag(viewport) failed:', e.message);
  - main.gs:2341: htmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');

## validate
- Occurrences: 12
- Files: Core.gs, main.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - Core.gs:470: const result = UnifiedValidation.validate('configuration', 'basic', { config });
  - main.gs:1048: const result = UnifiedValidation.validate('authentication', 'basic', {
  - unifiedValidationSystem.gs:35: validate(category, level = this.validationLevels.STANDARD, options = {}) {

## REASON
- Occurrences: 12
- Files: Core.gs, constants.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:823: reasonColumnIndex: indices[COLUMN_HEADERS.REASON],
  - Core.gs:825: hasReasonColumn: indices[COLUMN_HEADERS.REASON] !== undefined,
  - Core.gs:833: if (indices[COLUMN_HEADERS.REASON] === undefined) {

## EMAIL
- Occurrences: 12
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:2046: var emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
  - Core.gs:2129: row.originalData && row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]]
  - Core.gs:2130: ? row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]]

## rules
- Occurrences: 12
- Files: config.gs, constants.gs
- Samples:
  - config.gs:690: if (rules.exact) {
  - config.gs:691: for (const pattern of rules.exact) {
  - config.gs:702: if (rules.high) {

## clearAll
- Occurrences: 12
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:702: clearAll() {
  - unifiedCacheManager.gs:732: `[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, ScriptCache: ${scr
  - unifiedCacheManager.gs:943: window.sharedUtilities.throttle.clearAll();

## STANDARD
- Occurrences: 12
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:21: STANDARD: 'standard',
  - unifiedValidationSystem.gs:35: validate(category, level = this.validationLevels.STANDARD, options = {}) {
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t

## unifiedCacheManager
- Occurrences: 11
- Files: Core.gs, config.gs, constants.gs, database.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:789: * @deprecated この関数は削除されました。unifiedCacheManager.gs の getSheetsServiceCached() を使用してください。
  - Core.gs:792: // → unifiedCacheManager.gsの実装を使用してください
  - Core.gs:868: // clearAllExecutionCache() は unifiedCacheManager.gs に統合済み

## A1
- Occurrences: 11
- Files: Core.gs, config.gs, database.gs
- Samples:
  - Core.gs:1614: 範囲: `A1:${String.fromCharCode(64 + lastColumn)}${lastRow}`,
  - config.gs:1543: const range = sheet.getRange('A1:Z10');
  - database.gs:137: appendSheetsData(service, dbId, `'${logSheetName}'!A1`, [

## split
- Occurrences: 11
- Files: Core.gs, config.gs, constants.gs, main.gs, ulog.gs
- Samples:
  - Core.gs:2048: nameValue = row.originalData[emailIndex].split('@')[0];
  - Core.gs:4927: .split(',')
  - Core.gs:5310: return email.split('@')[1] || '';

## UNDERSTAND
- Occurrences: 11
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:2084: if (reactionKey === 'UNDERSTAND') count = row.understandCount || 0;
  - Core.gs:2139: UNDERSTAND: checkReactionState('UNDERSTAND'),
  - Core.gs:2139: UNDERSTAND: checkReactionState('UNDERSTAND'),

## LIKE
- Occurrences: 11
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:2085: else if (reactionKey === 'LIKE') count = row.likeCount || 0;
  - Core.gs:2140: LIKE: checkReactionState('LIKE'),
  - Core.gs:2140: LIKE: checkReactionState('LIKE'),

## CURIOUS
- Occurrences: 11
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:2086: else if (reactionKey === 'CURIOUS') count = row.curiousCount || 0;
  - Core.gs:2141: CURIOUS: checkReactionState('CURIOUS'),
  - Core.gs:2141: CURIOUS: checkReactionState('CURIOUS'),

## getUrl
- Occurrences: 11
- Files: Core.gs, config.gs, main.gs, session-utils.gs, unifiedCacheManager.gs, url.gs
- Samples:
  - Core.gs:2963: folderUrl: folder ? folder.getUrl() : '',
  - Core.gs:4197: spreadsheetUrl = spreadsheetObj.getUrl();
  - Core.gs:4210: spreadsheetUrl = spreadsheetObj.getUrl();

## fetch
- Occurrences: 11
- Files: Core.gs, database.gs, resilientExecutor.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:5326: var response = UrlFetchApp.fetch(url, {
  - database.gs:1407: const response = UrlFetchApp.fetch(url, {
  - database.gs:1466: const response = UrlFetchApp.fetch(url, {

## memoryUsage
- Occurrences: 11
- Files: config.gs
- Samples:
  - config.gs:66: this.memoryUsage.current = resourcesSize + updatesSize;
  - config.gs:68: if (this.memoryUsage.current > this.memoryUsage.peak) {
  - config.gs:68: if (this.memoryUsage.current > this.memoryUsage.peak) {

## delete
- Occurrences: 11
- Files: config.gs, constants.gs, secretManager.gs, unifiedCacheManager.gs, unifiedValidationSystem.gs
- Samples:
  - config.gs:125: this.resources.delete(key);
  - config.gs:232: this.activeContexts.delete(userId);
  - constants.gs:1086: delete(key) {

## getProperties
- Occurrences: 11
- Files: config.gs, database.gs, main.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:395: const allProperties = props.getProperties();
  - database.gs:2777: var allProps = props.getProperties();
  - main.gs:770: const allProperties = properties.getProperties();

## BATCH_SIZE
- Occurrences: 11
- Files: constants.gs, main.gs
- Samples:
  - constants.gs:70: BATCH_SIZE: {
  - constants.gs:253: Object.freeze(UNIFIED_CONSTANTS.CACHE.BATCH_SIZE);
  - constants.gs:307: /** @deprecated Use UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.LARGE instead */

## execute
- Occurrences: 11
- Files: resilientExecutor.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs
- Samples:
  - resilientExecutor.gs:47: async execute(operation, options = {}) {
  - resilientExecutor.gs:349: return resilientExecutor.execute(operation, {
  - resilientExecutor.gs:363: return resilientExecutor.execute(operation, {

## SECURITY_ERROR
- Occurrences: 11
- Files: secretManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - secretManager.gs:60: throw new Error('SECURITY_ERROR: 無効な秘密情報名');
  - secretManager.gs:132: throw new Error(`SECURITY_ERROR: 秘密情報が見つかりません: ${secretName}`);
  - secretManager.gs:152: throw new Error('SECURITY_ERROR: 秘密情報名と値は必須です');

## HEALTHY
- Occurrences: 11
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - systemIntegrationManager.gs:249: healthResult.overallStatus = 'HEALTHY';
  - unifiedSecurityManager.gs:763: healthCheckResult.overallStatus = 'HEALTHY';
  - unifiedValidationSystem.gs:332: passed: result.overallStatus === 'HEALTHY',

## SECTION
- Occurrences: 11
- Files: unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - unifiedCacheManager.gs:8: // SECTION 1: 統合キャッシュマネージャークラス（元cache.gs）
  - unifiedCacheManager.gs:1114: // SECTION 2: 統一実行レベルキャッシュ管理システム（元unifiedExecutionCache.gs）
  - unifiedCacheManager.gs:1265: // SECTION 3: SpreadsheetApp最適化システム（元spreadsheetCache.gs）

## sharedUtilities
- Occurrences: 11
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:910: window.sharedUtilities.cache &&
  - unifiedCacheManager.gs:911: typeof window.sharedUtilities.cache.clear === 'function'
  - unifiedCacheManager.gs:913: window.sharedUtilities.cache.clear();

## pattern
- Occurrences: 10
- Files: Core.gs, config.gs, database.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:186: return retryablePatterns.some((pattern) => pattern.test(coreErrorMessage));
  - config.gs:692: if (header === pattern.toLowerCase()) {
  - config.gs:695: if (header.includes(pattern.toLowerCase())) {

## OPINION
- Occurrences: 10
- Files: Core.gs, constants.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:824: opinionColumnIndex: indices[COLUMN_HEADERS.OPINION],
  - Core.gs:826: hasOpinionColumn: indices[COLUMN_HEADERS.OPINION] !== undefined,
  - Core.gs:838: if (indices[COLUMN_HEADERS.OPINION] === undefined) {

## Page
- Occurrences: 10
- Files: Core.gs
- Samples:
  - Core.gs:1196: * Page.htmlから呼び出される - フロントエンド期待形式に対応
  - Core.gs:1508: * Page.htmlから呼び出される - フロントエンド期待形式に対応
  - Core.gs:1632: // Page.html期待形式に変換

## moveTo
- Occurrences: 10
- Files: Core.gs
- Samples:
  - Core.gs:2791: // 推奨メソッドmoveTo()を使用してファイル移動
  - Core.gs:2792: formFile.moveTo(folder);
  - Core.gs:2830: // 推奨メソッドmoveTo()を使用してファイル移動

## resources
- Occurrences: 10
- Files: config.gs
- Samples:
  - config.gs:34: this.resources.set('sheetsService', this.sheetsService);
  - config.gs:44: this.resources.set('sheetsService', this.sheetsService);
  - config.gs:64: const resourcesSize = this.resources.size * 100; // 概算値

## FORMS
- Occurrences: 10
- Files: constants.gs
- Samples:
  - constants.gs:165: FORMS: {
  - constants.gs:263: Object.freeze(UNIFIED_CONSTANTS.FORMS);
  - constants.gs:264: Object.freeze(UNIFIED_CONSTANTS.FORMS.PRESETS);

## SCORING
- Occurrences: 10
- Files: constants.gs
- Samples:
  - constants.gs:204: SCORING: {
  - constants.gs:266: Object.freeze(UNIFIED_CONSTANTS.SCORING);
  - constants.gs:267: Object.freeze(UNIFIED_CONSTANTS.SCORING.WEIGHTS);

## REGEX
- Occurrences: 10
- Files: constants.gs
- Samples:
  - constants.gs:239: REGEX: {
  - constants.gs:271: Object.freeze(UNIFIED_CONSTANTS.REGEX);
  - constants.gs:360: /** @deprecated Use UNIFIED_CONSTANTS.REGEX.EMAIL instead */

## getStats
- Occurrences: 10
- Files: database.gs, resilientExecutor.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - database.gs:2031: var stats = cacheManager.getStats();
  - database.gs:2746: var cacheStats = cacheManager.getStats();
  - resilientExecutor.gs:237: getStats() {

## setXFrameOptionsMode
- Occurrences: 10
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:1500: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  - main.gs:1520: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
  - main.gs:1566: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);

## GET
- Occurrences: 10
- Files: secretManager.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - secretManager.gs:94: this.logSecretAccess('GET', secretName, {
  - secretManager.gs:119: this.logSecretAccess('GET', secretName, {
  - secretManager.gs:227: method: 'GET',

## component
- Occurrences: 10
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:239: if (component.status === 'INITIALIZED') {
  - systemIntegrationManager.gs:314: if (component.status === 'INITIALIZED' && component.instance) {
  - systemIntegrationManager.gs:314: if (component.status === 'INITIALIZED' && component.instance) {

## relatedIds
- Occurrences: 10
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:508: if (relatedIds.length > maxRelated) {
  - unifiedCacheManager.gs:509: ULog.warn(`[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`);
  - unifiedCacheManager.gs:510: relatedIds = relatedIds.slice(0, maxRelated);

## database
- Occurrences: 9
- Files: Core.gs, config.gs, constants.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:86: `database.${operation}`,
  - Core.gs:93: ULog.error(`⚠️ LOGGING ERROR [database.${operation}]:`, {
  - Core.gs:103: context: `database.${operation}`,

## parseError
- Occurrences: 9
- Files: Core.gs, config.gs, database.gs, main.gs
- Samples:
  - Core.gs:543: ULog.warn('configJson解析エラー:', parseError.message);
  - config.gs:2661: console.error('[ERROR]', '❌ Failed to recover context from JSON:', parseError.message);
  - database.gs:1540: throw new Error(`レスポンス解析失敗: ${parseError.message}`);

## getLastColumn
- Occurrences: 9
- Files: Core.gs, config.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:1609: var lastColumn = sheet.getLastColumn();
  - Core.gs:1835: var lastColumn = sheet.getLastColumn();
  - Core.gs:3460: var lastColumn = sheet.getLastColumn();

## NAME
- Occurrences: 9
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:1652: sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
  - Core.gs:1850: sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
  - Core.gs:4827: var nameIndex = headerIndices[COLUMN_HEADERS.NAME];

## getFileById
- Occurrences: 9
- Files: Core.gs
- Samples:
  - Core.gs:2772: var formFile = DriveApp.getFileById(formAndSsInfo.formId);
  - Core.gs:2773: var ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);
  - Core.gs:3710: const formFile = DriveApp.getFileById(form.getId());

## array
- Occurrences: 9
- Files: Core.gs, constants.gs, unifiedBatchProcessor.gs
- Samples:
  - Core.gs:4911: for (var i = array.length - 1; i > 0; i--) {
  - constants.gs:862: for (let i = 0; i < array.length; i += size) {
  - constants.gs:863: chunks.push(array.slice(i, i + size));

## WARNING
- Occurrences: 9
- Files: Core.gs, secretManager.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs, validationMigration.gs
- Samples:
  - Core.gs:5151: 'mapConfigToActualHeaders: WARNING - No match found for %s = "%s"',
  - secretManager.gs:669: results.criticalSecretsStatus = 'WARNING';
  - systemIntegrationManager.gs:251: healthResult.overallStatus = 'WARNING';

## UNKNOWN
- Occurrences: 9
- Files: autoInit.gs, secretManager.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - autoInit.gs:144: diagnosticsStatus: diagnostics?.healthCheck?.overallStatus || 'UNKNOWN',
  - secretManager.gs:589: secretManagerStatus: 'UNKNOWN',
  - secretManager.gs:590: propertiesServiceStatus: 'UNKNOWN',

## has
- Occurrences: 9
- Files: config.gs, constants.gs, database.gs, secretManager.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:119: if (this.resources.has(key)) {
  - config.gs:229: if (this.activeContexts.has(userId)) {
  - config.gs:618: if (usedHeaders.has(headerInfo.index) || headerInfo.isMetadata) {

## validator
- Occurrences: 9
- Files: constants.gs
- Samples:
  - constants.gs:721: if (validator.type === 'email' && !this.validateEmail(value)) {
  - constants.gs:724: if (validator.type === 'spreadsheetId' && !this.validateSpreadsheetId(value)) {
  - constants.gs:727: if (validator.type === 'userId' && !this.validateUserId(value)) {

## createTemplateFromFile
- Occurrences: 9
- Files: main.gs, unifiedUtilities.gs
- Samples:
  - main.gs:1494: const template = HtmlService.createTemplateFromFile('LoginPage');
  - main.gs:1514: const template = HtmlService.createTemplateFromFile('SetupPage');
  - main.gs:1559: const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');

## OPEN
- Occurrences: 9
- Files: resilientExecutor.gs, unifiedSecurityManager.gs
- Samples:
  - resilientExecutor.gs:23: state: 'CLOSED', // CLOSED, OPEN, HALF_OPEN
  - resilientExecutor.gs:27: failureThreshold: 5, // OPENになる連続失敗回数
  - resilientExecutor.gs:28: openTimeout: 60000, // OPEN状態の継続時間（1分）

## LEVELS
- Occurrences: 9
- Files: ulog.gs
- Samples:
  - ulog.gs:17: const levelValue = ULog.LEVELS[level] || ULog.LEVELS.INFO;
  - ulog.gs:17: const levelValue = ULog.LEVELS[level] || ULog.LEVELS.INFO;
  - ulog.gs:17: const levelValue = ULog.LEVELS[level] || ULog.LEVELS.INFO;

## overall
- Occurrences: 9
- Files: unifiedValidationSystem.gs, workflowValidation.gs
- Samples:
  - unifiedValidationSystem.gs:821: passed: result.overall.success,
  - unifiedValidationSystem.gs:822: message: `包括ワークフロー検証: ${result.overall.passed}/${result.overall.total} 通過`,
  - unifiedValidationSystem.gs:822: message: `包括ワークフロー検証: ${result.overall.passed}/${result.overall.total} 通過`,

## LOW
- Occurrences: 8
- Files: Core.gs, constants.gs, main.gs
- Samples:
  - Core.gs:61: case UNIFIED_CONSTANTS.ERROR.SEVERITY.LOW:
  - Core.gs:62: ULog.info(`ℹ️ LOW SEVERITY [${context}]:`, errorInfo.message);
  - Core.gs:131: UNIFIED_CONSTANTS.ERROR.SEVERITY.LOW,

## floor
- Occurrences: 8
- Files: Core.gs, config.gs, database.gs, main.gs, resilientExecutor.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:336: Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))
  - Core.gs:4912: var j = Math.floor(Math.random() * (i + 1));
  - config.gs:2617: num = Math.floor(num / 26);

## classQuestion
- Occurrences: 8
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3767: classItem.setTitle(config.classQuestion.title);
  - Core.gs:3768: classItem.setChoiceValues(config.classQuestion.choices);
  - Core.gs:3794: customConfig.classQuestion.choices &&

## retryError
- Occurrences: 8
- Files: Core.gs, config.gs, database.gs
- Samples:
  - Core.gs:4219: ULog.warn(`❌ リトライ${retry + 1}回目失敗:`, retryError.message);
  - config.gs:3071: console.error('[ERROR]', '❌ 強制リフレッシュリトライも失敗:', retryError.message);
  - config.gs:3073: `Sheets APIサービスの復旧に完全に失敗しました。初期エラー: ${serviceError.message}, リトライエラー: ${retryError.message}`

## unifiedSecurityManager
- Occurrences: 8
- Files: Core.gs, constants.gs, secretManager.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:4246: // → unifiedSecurityManager.gsの実装を使用してください
  - Core.gs:4308: // → unifiedSecurityManager.gsの実装を使用してください
  - Core.gs:5371: // isDeployUser() は unifiedSecurityManager.gs に統合済み

## fallbackError
- Occurrences: 8
- Files: Core.gs, config.gs, main.gs, resilientExecutor.gs, session-utils.gs
- Samples:
  - Core.gs:4615: ULog.error('[ERROR]', 'フォールバック処理も失敗: ' + fallbackError.message);
  - Core.gs:6660: ULog.warn('⚠️ フォールバックシート名検索に失敗:', fallbackError.message);
  - config.gs:2199: console.error('[ERROR]', '❌ フォールバック処理に失敗:', fallbackError.message);

## sheetError
- Occurrences: 8
- Files: Core.gs, config.gs, database.gs
- Samples:
  - Core.gs:5728: ULog.warn('⚠️ シートアクティベーション失敗（処理継続）:', sheetError.message);
  - Core.gs:5733: error: sheetError.message,
  - Core.gs:5939: ULog.warn('⚠️ QuickStartシートアクティベーション失敗（処理継続）:', sheetError.message);

## setTimeout
- Occurrences: 8
- Files: config.gs, main.gs, resilientExecutor.gs, unifiedBatchProcessor.gs
- Samples:
  - config.gs:85: this.memoryMonitorTimer = setTimeout(() => {
  - config.gs:97: this.cleanupTimer = setTimeout(() => {
  - config.gs:3914: setTimeout(() => {

## clear
- Occurrences: 8
- Files: config.gs, secretManager.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:163: this.resources.clear();
  - secretManager.gs:542: this.secretCache.clear();
  - unifiedCacheManager.gs:694: this.memoCache.clear();

## activeContexts
- Occurrences: 8
- Files: config.gs
- Samples:
  - config.gs:219: if (this.activeContexts.size >= this.maxConcurrentContexts) {
  - config.gs:224: this.activeContexts.set(context.requestUserId, context);
  - config.gs:225: ULog.debug(`📝 コンテキスト登録: ${context.requestUserId} (total: ${this.activeContexts.size})`);

## openError
- Occurrences: 8
- Files: config.gs
- Samples:
  - config.gs:323: console.error('[ERROR]', '❌ SpreadsheetApp.openById 権限エラー:', openError.message);
  - config.gs:331: errorType: openError.name,
  - config.gs:332: errorMessage: openError.message,

## CURRENT_USER_ID
- Occurrences: 8
- Files: config.gs, database.gs, session-utils.gs
- Samples:
  - config.gs:390: props.deleteProperty('CURRENT_USER_ID');
  - database.gs:1259: var currentUserId = userProps.getProperty('CURRENT_USER_ID');
  - database.gs:1262: userProps.deleteProperty('CURRENT_USER_ID');

## getUserCache
- Occurrences: 8
- Files: config.gs, session-utils.gs, unifiedSecurityManager.gs
- Samples:
  - config.gs:420: const userCache = CacheService.getUserCache();
  - config.gs:490: const userCache = CacheService.getUserCache();
  - session-utils.gs:15: return resilientExecutor.execute(() => CacheService.getUserCache(), {

## val
- Occurrences: 8
- Files: config.gs
- Samples:
  - config.gs:841: avgLength: columnData.reduce((sum, val) => sum + val.length, 0) / columnData.length,
  - config.gs:842: maxLength: Math.max(...columnData.map((val) => val.length)),
  - config.gs:843: hasLongText: columnData.some((val) => val.length > 50),

## SHORT
- Occurrences: 8
- Files: constants.gs
- Samples:
  - constants.gs:63: SHORT: 300, // 5分
  - constants.gs:80: SHORT: 1000, // 1秒
  - constants.gs:301: /** @deprecated Use UNIFIED_CONSTANTS.CACHE.TTL.SHORT instead */

## TIMEOUTS
- Occurrences: 8
- Files: constants.gs
- Samples:
  - constants.gs:79: TIMEOUTS: {
  - constants.gs:254: Object.freeze(UNIFIED_CONSTANTS.TIMEOUTS);
  - constants.gs:304: /** @deprecated Use UNIFIED_CONSTANTS.TIMEOUTS.EXECUTION_MAX instead */

## DISPLAY
- Occurrences: 8
- Files: constants.gs
- Samples:
  - constants.gs:103: DISPLAY: {
  - constants.gs:256: Object.freeze(UNIFIED_CONSTANTS.DISPLAY);
  - constants.gs:257: Object.freeze(UNIFIED_CONSTANTS.DISPLAY.MODES);

## LIMITS
- Occurrences: 8
- Files: constants.gs
- Samples:
  - constants.gs:193: LIMITS: {
  - constants.gs:265: Object.freeze(UNIFIED_CONSTANTS.LIMITS);
  - constants.gs:310: /** @deprecated Use UNIFIED_CONSTANTS.LIMITS.HISTORY_ITEMS instead */

## searchParams
- Occurrences: 8
- Files: main.gs, url.gs
- Samples:
  - main.gs:1886: url.searchParams.set('mode', mode);
  - main.gs:1890: url.searchParams.set(key, params[key]);
  - url.gs:193: url.searchParams.set('_cb', Date.now().toString());

## createHtmlOutput
- Occurrences: 8
- Files: main.gs, session-utils.gs, unifiedUtilities.gs
- Samples:
  - main.gs:1906: return HtmlService.createHtmlOutput().setContent(
  - main.gs:2070: const htmlOutput = HtmlService.createHtmlOutput(userActionRedirectHtml);
  - main.gs:2427: const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setTitle('StudyQuest - 準備中');

## INITIALIZED
- Occurrences: 8
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:133: this.components.unifiedSecretManager.status = 'INITIALIZED';
  - systemIntegrationManager.gs:142: this.components.resilientExecutor.status = 'INITIALIZED';
  - systemIntegrationManager.gs:158: this.components.multiTenantSecurity.status = 'INITIALIZED';

## executionCache
- Occurrences: 8
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1999: this.executionCache.clearUserInfo();
  - unifiedCacheManager.gs:2000: this.executionCache.syncWithUnifiedCache('userDataChange');
  - unifiedCacheManager.gs:2036: this.executionCache.clearAll();

## AppsScript
- Occurrences: 8
- Files: url.gs
- Samples:
  - url.gs:38: if (typeof AppsScript === 'undefined') {
  - url.gs:41: } else if (!AppsScript.Script) {
  - url.gs:42: ULog.debug('AppsScript.Script not available, using fallback method');

## WRITE_OPERATION
- Occurrences: 7
- Files: Core.gs, database.gs, lockManager.gs
- Samples:
  - Core.gs:363: return executeWithStandardizedLock('WRITE_OPERATION', 'clearActiveSheet', () => {
  - Core.gs:3520: return executeWithStandardizedLock('WRITE_OPERATION', 'processReaction', () => {
  - database.gs:107: return executeWithStandardizedLock('WRITE_OPERATION', 'logAccountDeletion', () => {

## stage
- Occurrences: 7
- Files: Core.gs
- Samples:
  - Core.gs:1036: Utilities.sleep(stage.delay);
  - Core.gs:1052: stage: stage.method,
  - Core.gs:1067: stage: stage.method,

## CLASS
- Occurrences: 7
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:1650: sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
  - Core.gs:1848: sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
  - Core.gs:2561: const classIndex = headerIndices[COLUMN_HEADERS.CLASS];

## ANONYMOUS
- Occurrences: 7
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:1703: adminMode === true ? DISPLAY_MODES.NAMED : configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
  - Core.gs:1820: displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
  - Core.gs:1863: var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

## hasNext
- Occurrences: 7
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:2781: while (formParents.hasNext()) {
  - Core.gs:2816: while (ssParents.hasNext()) {
  - Core.gs:3343: if (folders.hasNext()) {

## next
- Occurrences: 7
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:2782: if (formParents.next().getId() === folder.getId()) {
  - Core.gs:2817: if (ssParents.next().getId() === folder.getId()) {
  - Core.gs:3344: rootFolder = folders.next();

## create
- Occurrences: 7
- Files: Core.gs, config.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs, unifiedUtilities.gs
- Samples:
  - Core.gs:3703: var form = FormApp.create(formTitle);
  - Core.gs:4106: var spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
  - config.gs:1650: const form = FormApp.create(formTitle);

## addTextItem
- Occurrences: 7
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3771: var nameItem = form.addTextItem();
  - Core.gs:3775: var mainItem = form.addTextItem();
  - Core.gs:3804: var nameItem = form.addTextItem();

## serviceError
- Occurrences: 7
- Files: Core.gs, config.gs, database.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:4369: ULog.warn('サービスアカウント追加で警告: ' + serviceError.message);
  - config.gs:3037: console.error('[ERROR]', '❌ SheetsService復旧エラー:', serviceError.message);
  - config.gs:3038: console.error('[ERROR]', '❌ Error stack:', serviceError.stack);

## MD5
- Occurrences: 7
- Files: config.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - config.gs:402: Utilities.DigestAlgorithm.MD5,
  - session-utils.gs:120: Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
  - unifiedBatchProcessor.gs:417: // MD5ハッシュ化でキーを短縮

## find
- Occurrences: 7
- Files: config.gs, database.gs, unifiedSheetDataManager.gs
- Samples:
  - config.gs:1274: const match = headers.find((h) => String(h).trim() === overrides.mainQuestion.trim());
  - config.gs:1278: const match = headers.find((h) => String(h).trim() === overrides.reasonQuestion.trim());
  - config.gs:1282: const match = headers.find((h) => String(h).trim() === overrides.nameQuestion.trim());

## apiError
- Occurrences: 7
- Files: config.gs, url.gs
- Samples:
  - config.gs:2244: console.error('[ERROR]', '❌ Sheets API取得エラー:', apiError.message);
  - config.gs:2268: `ヘッダー取得に失敗しました。Sheets API: ${apiError.message}, SpreadsheetApp: ${spreadsheetError.message}`
  - config.gs:3104: console.error('[ERROR]', '❌ getSpreadsheetsData failed:', apiError.message);

## dbError
- Occurrences: 7
- Files: config.gs, database.gs, main.gs, setup.gs
- Samples:
  - config.gs:3498: console.error('[ERROR]', `❌ DB書き込み失敗 (試行${dbWriteAttempts}):`, dbError.message);
  - config.gs:3502: throw new Error('DB書き込み処理に失敗しました: ' + dbError.message);
  - config.gs:3506: if (dbError.message.includes('503') || dbError.message.includes('429')) {

## NETWORK
- Occurrences: 7
- Files: constants.gs, main.gs
- Samples:
  - constants.gs:20: * @property {string} NETWORK - ネットワークエラー
  - constants.gs:49: NETWORK: 'network',
  - constants.gs:425: NETWORK: 'network',

## chunkArray
- Occurrences: 7
- Files: constants.gs, unifiedBatchProcessor.gs
- Samples:
  - constants.gs:856: chunkArray(array, size = UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.MEDIUM) {
  - constants.gs:935: const chunks = this.chunkArray(items, batchSize);
  - unifiedBatchProcessor.gs:69: const chunkedRanges = this.chunkArray(ranges, this.config.maxBatchSize);

## toFixed
- Occurrences: 7
- Files: database.gs, resilientExecutor.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs
- Samples:
  - database.gs:2868: (monitoringResult.summary.metrics.cacheHitRate * 100).toFixed(1) +
  - resilientExecutor.gs:243: ? ((this.stats.successfulExecutions / this.stats.totalExecutions) * 100).toFixed(2) + '%'
  - systemIntegrationManager.gs:435: totalOperations > 0 ? ((totalErrors / totalOperations) * 100).toFixed(2) + '%' : '0%';

## GOOGLE_CLIENT_ID
- Occurrences: 7
- Files: main.gs
- Samples:
  - main.gs:760: ULog.debug('Getting GOOGLE_CLIENT_ID from script properties...');
  - main.gs:762: const clientId = properties.getProperty('GOOGLE_CLIENT_ID');
  - main.gs:764: ULog.debug('GOOGLE_CLIENT_ID retrieved:', clientId ? 'Found' : 'Not found');

## evaluate
- Occurrences: 7
- Files: main.gs
- Samples:
  - main.gs:1495: const htmlOutput = template.evaluate().setTitle('StudyQuest - ログイン');
  - main.gs:1515: const htmlOutput = template.evaluate().setTitle('StudyQuest - 初回セットアップ');
  - main.gs:1561: const htmlOutput = appSetupTemplate.evaluate().setTitle('アプリ設定 - StudyQuest');

## resolve
- Occurrences: 7
- Files: resilientExecutor.gs, unifiedCacheManager.gs
- Samples:
  - resilientExecutor.gs:140: Promise.resolve()
  - resilientExecutor.gs:144: resolve(result);
  - unifiedCacheManager.gs:266: const newValues = Promise.resolve(valuesFn(missingKeys));

## systemMetrics
- Occurrences: 7
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:92: this.systemMetrics.startupTime = new Date().toISOString();
  - systemIntegrationManager.gs:416: systemIntegrationManager.systemMetrics.lastHealthCheck = new Date().toISOString();
  - systemIntegrationManager.gs:433: systemIntegrationManager.systemMetrics.totalRequests = totalOperations;

## prototype
- Occurrences: 7
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:844: CacheManager.prototype.clearInProgress = false;
  - unifiedCacheManager.gs:845: CacheManager.prototype.pendingClears = [];
  - unifiedCacheManager.gs:852: CacheManager.prototype.clearAllFrontendCaches = function (options = {}) {

## preWarmedItems
- Occurrences: 7
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1815: results.preWarmedItems.push('service_account_token');
  - unifiedCacheManager.gs:1826: results.preWarmedItems.push('user_by_email');
  - unifiedCacheManager.gs:1831: results.preWarmedItems.push('user_by_id');

## batchOperations
- Occurrences: 6
- Files: Core.gs
- Samples:
  - Core.gs:1269: if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
  - Core.gs:1275: if (batchOperations.length > MAX_BATCH_SIZE) {
  - Core.gs:1279: ULog.debug('🔄 バッチリアクション処理開始:', batchOperations.length + '件');

## getLastRow
- Occurrences: 6
- Files: Core.gs, config.gs, main.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:1608: var lastRow = sheet.getLastRow();
  - Core.gs:1802: var lastRow = sheet.getLastRow(); // スプレッドシートの最終行
  - Core.gs:2555: const lastRow = sheet.getLastRow();

## NAMED
- Occurrences: 6
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:1703: adminMode === true ? DISPLAY_MODES.NAMED : configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
  - Core.gs:2037: var shouldShowName = adminMode === true || displayMode === DISPLAY_MODES.NAMED || isOwner;
  - Core.gs:4831: (displayMode === DISPLAY_MODES.NAMED || isOwner)

## reduce
- Occurrences: 6
- Files: Core.gs, config.gs, unifiedBatchProcessor.gs
- Samples:
  - Core.gs:2567: return values.reduce((cnt, row) => (row[0] === classFilter ? cnt + 1 : cnt), 0);
  - Core.gs:5062: var longestHeader = questionHeaders.reduce(function (prev, current) {
  - config.gs:648: const fallbackHeader = candidateHeaders.reduce((best, current) => {

## addParagraphTextItem
- Occurrences: 6
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3779: var reasonItem = form.addParagraphTextItem();
  - Core.gs:3854: mainItem = form.addParagraphTextItem();
  - Core.gs:3860: var reasonItem = form.addParagraphTextItem();

## accessError
- Occurrences: 6
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:4667: var statusMatch = ((accessError && accessError.message) || '').match(
  - Core.gs:4675: accessError.message
  - database.gs:1875: error: accessError.message,

## pow
- Occurrences: 6
- Files: Core.gs, constants.gs, resilientExecutor.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:4705: Utilities.sleep(Math.pow(2, attempt) * 1000);
  - constants.gs:823: const waitTime = exponentialBackoff ? delay * Math.pow(2, attempt - 1) : delay;
  - resilientExecutor.gs:159: let delay = this.config.baseDelay * Math.pow(this.config.backoffFactor, attempt);

## SECURITY
- Occurrences: 6
- Files: constants.gs, ulog.gs
- Samples:
  - constants.gs:26: * @property {string} SECURITY - セキュリティエラー
  - constants.gs:55: SECURITY: 'security',
  - constants.gs:431: SECURITY: 'security',

## LONG
- Occurrences: 6
- Files: constants.gs
- Samples:
  - constants.gs:65: LONG: 3600, // 1時間
  - constants.gs:82: LONG: 5000, // 5秒
  - constants.gs:438: * @param {string} duration - 期間（SHORT, MEDIUM, LONG, EXTENDED）

## example
- Occurrences: 6
- Files: constants.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - constants.gs:1212: generateUserUrls: () => 'http://example.com',
  - constants.gs:1213: generateUnpublishedUrl: () => 'http://example.com/unpublished'
  - systemIntegrationManager.gs:151: 'test@example.com',

## propsError
- Occurrences: 6
- Files: database.gs, main.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - database.gs:1266: ULog.warn('Failed to clear user properties:', propsError.message);
  - main.gs:960: errors.push(`PropertiesService エラー: ${propsError.message}`);
  - session-utils.gs:254: authResetResult.errors.push(`プロパティクリアエラー: ${propsError.message}`);

## details
- Occurrences: 6
- Files: database.gs
- Samples:
  - database.gs:2263: result.details.duplicates = duplicateResult.duplicates;
  - database.gs:2274: result.details.missingFields = missingFieldsResult.missing;
  - database.gs:2285: result.details.invalidData = invalidDataResult.invalid;

## PERFORMANCE
- Occurrences: 6
- Files: debugConfig.gs, ulog.gs
- Samples:
  - debugConfig.gs:28: PERFORMANCE: true, // パフォーマンス計測
  - debugConfig.gs:55: * @param {string} category - カテゴリ (CACHE, AUTH, DATABASE, UI, PERFORMANCE)
  - debugConfig.gs:135: category = 'PERFORMANCE';

## DENY
- Occurrences: 6
- Files: main.gs
- Samples:
  - main.gs:1519: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
  - main.gs:1520: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
  - main.gs:1565: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {

## getService
- Occurrences: 6
- Files: main.gs, session-utils.gs, url.gs
- Samples:
  - main.gs:1623: const baseUrl = ScriptApp.getService().getUrl();
  - session-utils.gs:263: const loginPageUrl = ScriptApp.getService().getUrl();
  - session-utils.gs:426: const fallbackUrl = ScriptApp.getService().getUrl() + '?mode=login';

## secretTypes
- Occurrences: 6
- Files: secretManager.gs
- Samples:
  - secretManager.gs:38: SERVICE_ACCOUNT_CREDS: this.secretTypes.SERVICE_ACCOUNT,
  - secretManager.gs:39: DATABASE_SPREADSHEET_ID: this.secretTypes.DATABASE_CREDS,
  - secretManager.gs:40: WEBHOOK_SECRET: this.secretTypes.WEBHOOK_SECRET,

## secretCache
- Occurrences: 6
- Files: secretManager.gs
- Samples:
  - secretManager.gs:516: return this.secretCache.has(secretName);
  - secretManager.gs:520: const cached = this.secretCache.get(secretName);
  - secretManager.gs:525: const cached = this.secretCache.get(secretName);

## FAILED
- Occurrences: 6
- Files: secretManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - secretManager.gs:569: if (action.includes('ERROR') || action.includes('FAILED')) {
  - unifiedCacheManager.gs:732: `[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, ScriptCache: ${scr
  - unifiedCacheManager.gs:732: `[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, ScriptCache: ${scr

## validateTenantBoundary
- Occurrences: 6
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - systemIntegrationManager.gs:150: const testResult = multiTenantSecurity.validateTenantBoundary(
  - unifiedSecurityManager.gs:291: validateTenantBoundary(requestUserId, targetUserId, operation) {
  - unifiedSecurityManager.gs:492: if (!multiTenantSecurity.validateTenantBoundary(currentUserId, userId, `cache_${operation}`)) {

## _getCallerInfo
- Occurrences: 6
- Files: ulog.gs
- Samples:
  - ulog.gs:77: const caller = ULog._getCallerInfo();
  - ulog.gs:88: const caller = ULog._getCallerInfo();
  - ulog.gs:99: const caller = ULog._getCallerInfo();

## op
- Occurrences: 6
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:357: return secureMultiTenantCacheOperation('get', op.key, userId);
  - unifiedBatchProcessor.gs:359: return secureMultiTenantCacheOperation('set', op.key, userId, op.value, {
  - unifiedBatchProcessor.gs:359: return secureMultiTenantCacheOperation('set', op.key, userId, op.value, {

## getHealth
- Occurrences: 6
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:746: getHealth() {
  - unifiedCacheManager.gs:1093: health: this.getHealth(),
  - unifiedCacheManager.gs:1901: const health = cacheManager.getHealth();

## BASIC
- Occurrences: 6
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:20: BASIC: 'basic',
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t
  - unifiedValidationSystem.gs:119: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t

## _calculateSummary
- Occurrences: 6
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:106: return this._calculateSummary(results);
  - unifiedValidationSystem.gs:136: return this._calculateSummary(results);
  - unifiedValidationSystem.gs:168: return this._calculateSummary(results);

## add
- Occurrences: 5
- Files: Core.gs, config.gs, database.gs
- Samples:
  - Core.gs:1327: processedRows.add(operation.rowIndex);
  - config.gs:633: usedHeaders.add(bestHeader.index);
  - config.gs:655: usedHeaders.add(fallbackHeader.index);

## updateError
- Occurrences: 5
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:3320: ULog.error('[ERROR]', 'エラー状態の更新に失敗: ' + updateError.message);
  - Core.gs:5586: ULog.error('[ERROR]', 'エラー状態の更新に失敗: ' + updateError.message);
  - database.gs:1025: var errorMessage = updateError.toString();

## HIGHLIGHT
- Occurrences: 5
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:3369: var highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  - Core.gs:4395: COLUMN_HEADERS.HIGHLIGHT,
  - Core.gs:4818: var highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];

## splice
- Occurrences: 5
- Files: Core.gs, main.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:3580: currentReactions.splice(userIndex, 1);
  - Core.gs:3591: currentReactions.splice(userIndex, 1);
  - main.gs:392: configJson.historyArray.splice(SERVER_MAX_HISTORY);

## getPublishedUrl
- Occurrences: 5
- Files: Core.gs, config.gs, constants.gs
- Samples:
  - Core.gs:3740: formUrl: form.getPublishedUrl(),
  - Core.gs:3741: viewFormUrl: form.getPublishedUrl(),
  - config.gs:1605: const formUrl = form.getPublishedUrl();

## reasonQuestion
- Occurrences: 5
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3780: reasonItem.setTitle(config.reasonQuestion.title);
  - Core.gs:3781: reasonItem.setHelpText(config.reasonQuestion.helpText);
  - Core.gs:3875: reasonItem.setTitle(config.reasonQuestion.title);

## setHelpText
- Occurrences: 5
- Files: Core.gs
- Samples:
  - Core.gs:3781: reasonItem.setHelpText(config.reasonQuestion.helpText);
  - Core.gs:3866: classItem.setHelpText(config.classQuestion.helpText);
  - Core.gs:3871: mainItem.setHelpText(config.mainQuestion.helpText);

## urlError
- Occurrences: 5
- Files: Core.gs, session-utils.gs
- Samples:
  - Core.gs:4203: ULog.warn('⚠️ 初回URL取得失敗、リトライします:', urlError.message);
  - session-utils.gs:276: authResetResult.errors.push(`URL取得エラー: ${urlError.message}`);
  - session-utils.gs:277: throw new Error(`ログインページURL取得に失敗: ${urlError.message}`);

## addEditor
- Occurrences: 5
- Files: Core.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:4341: file.addEditor(userEmail);
  - Core.gs:4351: spreadsheet.addEditor(userEmail);
  - Core.gs:4366: file.addEditor(serviceAccountEmail);

## random
- Occurrences: 5
- Files: Core.gs, resilientExecutor.gs, unifiedUserManager.gs, url.gs
- Samples:
  - Core.gs:4878: var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
  - Core.gs:4912: var j = Math.floor(Math.random() * (i + 1));
  - resilientExecutor.gs:164: delay = delay * (0.5 + Math.random() * 0.5);

## hasOwnProperty
- Occurrences: 5
- Files: Core.gs, secretManager.gs, ulog.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:4980: if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {
  - Core.gs:6992: if (updateData.hasOwnProperty(field)) {
  - secretManager.gs:467: return this.criticalSecrets.hasOwnProperty(secretName);

## AppSetupPage
- Occurrences: 5
- Files: Core.gs
- Samples:
  - Core.gs:5240: * 現在のユーザーのステータス情報を取得（AppSetupPage.htmlから呼び出される）
  - Core.gs:5245: * isActive状態を更新（AppSetupPage.htmlから呼び出される）
  - Core.gs:5390: * 管理者によるユーザー削除（AppSetupPage.html用ラッパー）

## _meta
- Occurrences: 5
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:6680: response._meta.includedApis.push('getSheetDetails');
  - Core.gs:6692: response._meta.executionTime = endTime - startTime;
  - Core.gs:6695: executionTime: response._meta.executionTime + 'ms',

## _trackMemoryUsage
- Occurrences: 5
- Files: config.gs
- Samples:
  - config.gs:35: this._trackMemoryUsage('sheetsService');
  - config.gs:45: this._trackMemoryUsage('sheetsService');
  - config.gs:61: _trackMemoryUsage(resourceKey) {

## CURRENT_USER_ID_
- Occurrences: 5
- Files: config.gs, session-utils.gs, unifiedUserManager.gs
- Samples:
  - config.gs:399: key.startsWith('CURRENT_USER_ID_') &&
  - config.gs:401: `CURRENT_USER_ID_${Utilities.computeDigest(
  - session-utils.gs:119: 'CURRENT_USER_ID_' +

## computeDigest
- Occurrences: 5
- Files: config.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - config.gs:401: `CURRENT_USER_ID_${Utilities.computeDigest(
  - session-utils.gs:120: Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
  - unifiedBatchProcessor.gs:418: const paramsHash = Utilities.computeDigest(

## DigestAlgorithm
- Occurrences: 5
- Files: config.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - config.gs:402: Utilities.DigestAlgorithm.MD5,
  - session-utils.gs:120: Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
  - unifiedBatchProcessor.gs:419: Utilities.DigestAlgorithm.MD5,

## UTF_8
- Occurrences: 5
- Files: config.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - config.gs:404: Utilities.Charset.UTF_8
  - session-utils.gs:120: Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
  - unifiedBatchProcessor.gs:421: Utilities.Charset.UTF_8

## Charset
- Occurrences: 5
- Files: config.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - config.gs:404: Utilities.Charset.UTF_8
  - session-utils.gs:120: Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
  - unifiedBatchProcessor.gs:421: Utilities.Charset.UTF_8

## Unpublished
- Occurrences: 5
- Files: config.gs, main.gs
- Samples:
  - config.gs:1006: * 既存の設定で回答ボードを再公開 (Unpublished.htmlから呼び出される)
  - config.gs:1309: * Unpublished.htmlから呼び出される
  - main.gs:2276: * ErrorBoundaryを回避して確実にUnpublished.htmlを表示

## newConfig
- Occurrences: 5
- Files: config.gs
- Samples:
  - config.gs:3663: newConfigType: typeof newConfig,
  - config.gs:3688: if (newConfig.length > 0 && typeof newConfig[0] === 'object') {
  - config.gs:3688: if (newConfig.length > 0 && typeof newConfig[0] === 'object') {

## CONFIG
- Occurrences: 5
- Files: constants.gs, unifiedValidationSystem.gs
- Samples:
  - constants.gs:25: * @property {string} CONFIG - 設定エラー
  - constants.gs:54: CONFIG: 'config',
  - constants.gs:430: CONFIG: 'config',

## LOG
- Occurrences: 5
- Files: constants.gs
- Samples:
  - constants.gs:158: LOG: {
  - constants.gs:262: Object.freeze(UNIFIED_CONSTANTS.SHEETS.LOG);
  - constants.gs:354: /** @deprecated Use UNIFIED_CONSTANTS.SHEETS.LOG instead */

## ACCESS
- Occurrences: 5
- Files: constants.gs, main.gs
- Samples:
  - constants.gs:518: ACCESS: 'access',
  - constants.gs:1299: ACCESS: 'access',
  - main.gs:1670: errorType = ERROR_TYPES.ACCESS;

## USER
- Occurrences: 5
- Files: constants.gs, main.gs
- Samples:
  - constants.gs:521: USER: 'user',
  - constants.gs:1302: USER: 'user',
  - main.gs:1662: let errorType = ERROR_TYPES.USER;

## BATCH
- Occurrences: 5
- Files: constants.gs
- Samples:
  - constants.gs:1121: const batchGet = (keys) => console.log(`[BATCH] Getting ${keys.length} items`);
  - constants.gs:1122: const fallbackBatchGet = (keys) => console.log(`[BATCH] Fallback get for ${keys.length} items`);
  - constants.gs:1123: const batchUpdate = (updates) => console.log(`[BATCH] Updating ${updates.length} items`);

## tokenError
- Occurrences: 5
- Files: database.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs
- Samples:
  - database.gs:724: console.error('[ERROR]', '❌ Failed to get service account token:', tokenError.message);
  - database.gs:725: throw new Error('サービスアカウントトークンの取得に失敗しました: ' + tokenError.message);
  - unifiedBatchProcessor.gs:253: tokenError.message

## parsed
- Occurrences: 5
- Files: database.gs
- Samples:
  - database.gs:1545: success: !!parsed.updates,
  - database.gs:1546: updatedRows: parsed.updates?.updatedRows || 0,
  - database.gs:1547: updatedColumns: parsed.updates?.updatedColumns || 0,

## actions
- Occurrences: 5
- Files: database.gs
- Samples:
  - database.gs:2140: result.summary.actions.push('サービスアカウントの再共有実行');
  - database.gs:2157: result.summary.actions.push('アクセス権限修復成功');
  - database.gs:2164: result.summary.actions.push('修復後テスト失敗: ' + retestError.message);

## vulnerabilities
- Occurrences: 5
- Files: database.gs
- Samples:
  - database.gs:2782: securityResult.vulnerabilities.push({
  - database.gs:2792: securityResult.vulnerabilities.push({
  - database.gs:2803: securityResult.vulnerabilities.push({

## POST
- Occurrences: 5
- Files: database.gs, secretManager.gs, unifiedBatchProcessor.gs
- Samples:
  - database.gs:3033: method: 'POST',
  - secretManager.gs:286: method: 'POST',
  - secretManager.gs:311: method: 'POST',

## scale
- Occurrences: 5
- Files: main.gs
- Samples:
  - main.gs:2002: 0% { transform: scale(1); }
  - main.gs:2003: 50% { transform: scale(1.05); }
  - main.gs:2004: 100% { transform: scale(1); }

## CLOSED
- Occurrences: 5
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:23: state: 'CLOSED', // CLOSED, OPEN, HALF_OPEN
  - resilientExecutor.gs:23: state: 'CLOSED', // CLOSED, OPEN, HALF_OPEN
  - resilientExecutor.gs:201: this.circuitBreaker.state = 'CLOSED';

## HALF_OPEN
- Occurrences: 5
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:23: state: 'CLOSED', // CLOSED, OPEN, HALF_OPEN
  - resilientExecutor.gs:26: successThreshold: 3, // HALF_OPENでの成功回数しきい値
  - resilientExecutor.gs:62: this.circuitBreaker.state = 'HALF_OPEN';

## reject
- Occurrences: 5
- Files: resilientExecutor.gs, unifiedCacheManager.gs
- Samples:
  - resilientExecutor.gs:137: reject(new Error(`Operation timed out after ${timeoutMs}ms`));
  - resilientExecutor.gs:148: reject(error);
  - unifiedCacheManager.gs:1006: reject(error);

## logSecretAccess
- Occurrences: 5
- Files: secretManager.gs
- Samples:
  - secretManager.gs:76: this.logSecretAccess('CACHE_HIT', secretName, logMeta);
  - secretManager.gs:94: this.logSecretAccess('GET', secretName, {
  - secretManager.gs:119: this.logSecretAccess('GET', secretName, {

## NOT_INITIALIZED
- Occurrences: 5
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:14: resilientExecutor: { status: 'NOT_INITIALIZED', instance: null },
  - systemIntegrationManager.gs:15: unifiedBatchProcessor: { status: 'NOT_INITIALIZED', instance: null },
  - systemIntegrationManager.gs:16: multiTenantSecurity: { status: 'NOT_INITIALIZED', instance: null },

## updateProcessingMetrics
- Occurrences: 5
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:113: this.updateProcessingMetrics(Date.now() - startTime, true);
  - unifiedBatchProcessor.gs:197: this.updateProcessingMetrics(Date.now() - startTime, true);
  - unifiedBatchProcessor.gs:315: this.updateProcessingMetrics(Date.now() - startTime, true);

## clearOp
- Occurrences: 5
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:967: const success = clearOp.operation();
  - unifiedCacheManager.gs:968: results.push({ name: clearOp.name, success });
  - unifiedCacheManager.gs:971: ULog.debug(`✅ ${clearOp.name} cleared successfully`);

## CACHE_KEY_PREFIX
- Occurrences: 5
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1276: CACHE_KEY_PREFIX: 'ss_cache_',
  - unifiedCacheManager.gs:1290: const cacheKey = `${SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX}${spreadsheetId}`;
  - unifiedCacheManager.gs:1431: const cacheKey = `${SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX}${spreadsheetId}`;

## logSecurityViolation
- Occurrences: 5
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:293: this.logSecurityViolation('MISSING_USER_ID', { requestUserId, targetUserId, operation });
  - unifiedSecurityManager.gs:311: this.logSecurityViolation('TENANT_BOUNDARY_VIOLATION', {
  - unifiedSecurityManager.gs:336: this.logSecurityViolation('INVALID_ACCESS_PATTERN', { userId, dataType, operation });

## getCoreSheetHeaders
- Occurrences: 5
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:316: getCoreSheetHeaders(spreadsheetId, sheetName, options = {}) {
  - unifiedSheetDataManager.gs:368: return this.getCoreSheetHeaders(spreadsheetId, sheetName, {
  - unifiedSheetDataManager.gs:686: const headers = this.getCoreSheetHeaders(spreadsheetId, sheetName);

## getCoreSpreadsheetInfo
- Occurrences: 5
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:529: getCoreSpreadsheetInfo(userId, options = {}) {
  - unifiedSheetDataManager.gs:571: const info = this.getCoreSpreadsheetInfo(userId);
  - unifiedSheetDataManager.gs:583: const info = this.getCoreSpreadsheetInfo(userId, { usePublished });

## validationCategories
- Occurrences: 5
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:43: case this.validationCategories.AUTH:
  - unifiedValidationSystem.gs:46: case this.validationCategories.CONFIG:
  - unifiedValidationSystem.gs:49: case this.validationCategories.DATA:

## Script
- Occurrences: 5
- Files: url.gs
- Samples:
  - url.gs:44: } else if (!AppsScript.Script.Deployments) {
  - url.gs:45: ULog.debug('AppsScript.Script.Deployments not available, using fallback method');
  - url.gs:55: deploymentsList = AppsScript.Script.Deployments.list(scriptId, {

## getUuid
- Occurrences: 4
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:25: this.sessionId = Utilities.getUuid();
  - Core.gs:145: errorId: Utilities.getUuid(),
  - Core.gs:969: userId = Utilities.getUuid();

## debugConfig
- Occurrences: 4
- Files: Core.gs, url.gs
- Samples:
  - Core.gs:219: // debugLog関数はdebugConfig.gsで統一定義されていますが、テスト環境でのfallback定義
  - Core.gs:220: // debugLog は debugConfig.gs で統一制御されるため、重複定義を削除
  - url.gs:7: // debugLog関数はdebugConfig.gsで統一定義されていますが、テスト環境でのfallback定義

## JP
- Occurrences: 4
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:332: publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
  - Core.gs:333: stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
  - main.gs:1324: `時刻: ${new Date().toLocaleString('ja-JP')}`,

## SERVICE_ACCOUNT
- Occurrences: 4
- Files: Core.gs, secretManager.gs
- Samples:
  - Core.gs:1185: // 通常のアクセスを試行（SERVICE_ACCOUNTでも同じSpreadsheetApp API使用）
  - secretManager.gs:29: SERVICE_ACCOUNT: 'service_account',
  - secretManager.gs:38: SERVICE_ACCOUNT_CREDS: this.secretTypes.SERVICE_ACCOUNT,

## getSheets
- Occurrences: 4
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:1418: var sheets = spreadsheet.getSheets();
  - Core.gs:4156: var sheets = spreadsheetObj.getSheets();
  - Core.gs:4389: var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

## USER_NOT_FOUND
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:2381: return createErrorResponse('USER_NOT_FOUND', 'ユーザー情報が見つかりません', null);
  - Core.gs:2435: return createErrorResponse('USER_NOT_FOUND', 'ユーザー情報が見つかりません。', null);
  - Core.gs:3999: return createErrorResponse('USER_NOT_FOUND', 'ユーザー情報が見つかりません', null);

## formError
- Occurrences: 4
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:2651: ULog.error('[ERROR]', 'フォーム更新エラー: ' + formError.message);
  - Core.gs:2652: return createErrorResponse(formError.message, 'フォームの更新に失敗しました: ' + formError.message, null);
  - Core.gs:2652: return createErrorResponse(formError.message, 'フォームの更新に失敗しました: ' + formError.message, null);

## getParents
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:2778: var formParents = formFile.getParents();
  - Core.gs:2813: var ssParents = ssFile.getParents();
  - Core.gs:5753: const formParents = formFile.getParents();

## invalidateSheetData
- Occurrences: 4
- Files: Core.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:3400: cacheManager.invalidateSheetData(spreadsheetId, sheetName);
  - Core.gs:3632: cacheManager.invalidateSheetData(spreadsheetId, sheetName);
  - unifiedCacheManager.gs:348: invalidateSheetData(spreadsheetId, sheetName = null) {

## MM
- Occurrences: 4
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:3697: var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  - Core.gs:3943: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - Core.gs:4069: const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

## formatDate
- Occurrences: 4
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:3697: var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  - Core.gs:3943: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - Core.gs:4069: const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

## GoogleAppsScript
- Occurrences: 4
- Files: Core.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:3755: * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
  - Core.gs:4096: * @param {GoogleAppsScript.Forms.Form} form - フォーム
  - unifiedCacheManager.gs:1283: * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト

## setChoiceValues
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:3768: classItem.setChoiceValues(config.classQuestion.choices);
  - Core.gs:3799: classItem.setChoiceValues(customConfig.classQuestion.choices);
  - Core.gs:3826: mainItem.setChoiceValues(customConfig.mainQuestion.choices);

## nameQuestion
- Occurrences: 4
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3772: nameItem.setTitle(config.nameQuestion.title);
  - Core.gs:3880: nameItem.setTitle(config.nameQuestion.title);
  - Core.gs:3881: nameItem.setHelpText(config.nameQuestion.helpText);

## shareError
- Occurrences: 4
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:4187: ULog.warn('サービスアカウント共有エラー（処理継続）:', shareError.message);
  - Core.gs:4280: error: shareError.message,
  - Core.gs:4283: ULog.error('[ERROR]', '共有失敗:', user.adminEmail, shareError.message);

## rowData
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:4869: var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
  - Core.gs:4872: var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;
  - Core.gs:4872: var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;

## sort
- Occurrences: 4
- Files: Core.gs, unifiedCacheManager.gs, url.gs
- Samples:
  - Core.gs:4889: return data.sort(function (a, b) {
  - Core.gs:4899: return data.sort(function (a, b) {
  - unifiedCacheManager.gs:1400: .sort((a, b) => b.timestamp - a.timestamp);

## valueRange
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:5199: if (valueRange && valueRange.values && valueRange.values[0] && valueRange.values[0][0]) {
  - Core.gs:5199: if (valueRange && valueRange.values && valueRange.values[0] && valueRange.values[0][0]) {
  - Core.gs:5199: if (valueRange && valueRange.values && valueRange.values[0] && valueRange.values[0][0]) {

## www
- Occurrences: 4
- Files: Core.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:5321: baseUrl: 'https://www.googleapis.com/drive/v3',
  - unifiedSecurityManager.gs:36: const tokenUrl = 'https://www.googleapis.com/oauth2/v4/token';
  - unifiedSecurityManager.gs:45: scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',

## getSystemMetrics
- Occurrences: 4
- Files: autoInit.gs, systemIntegrationManager.gs
- Samples:
  - autoInit.gs:132: const metrics = systemIntegrationManager.getSystemMetrics();
  - systemIntegrationManager.gs:305: getSystemMetrics() {
  - systemIntegrationManager.gs:415: const metrics = systemIntegrationManager.getSystemMetrics();

## clearSecretCache
- Occurrences: 4
- Files: autoInit.gs, secretManager.gs, systemIntegrationManager.gs
- Samples:
  - autoInit.gs:166: unifiedSecretManager.clearSecretCache();
  - secretManager.gs:210: this.clearSecretCache(secretName);
  - secretManager.gs:538: clearSecretCache(secretName) {

## cleanup
- Occurrences: 4
- Files: config.gs
- Samples:
  - config.gs:123: resource.cleanup();
  - config.gs:155: resource.cleanup();
  - config.gs:253: cleanup() {

## min
- Occurrences: 4
- Files: config.gs, database.gs, resilientExecutor.gs
- Samples:
  - config.gs:765: return Math.min(score, 1.0); // 最大1.0に制限
  - config.gs:826: const lastRow = Math.min(sheet.getLastRow(), 21); // 最大20行まで分析
  - database.gs:3165: matrix[i][j] = Math.min(

## findIndex
- Occurrences: 4
- Files: config.gs, unifiedUserManager.gs
- Samples:
  - config.gs:877: const opinionIndex = processedHeaders.findIndex(
  - config.gs:897: const reasonIndex = processedHeaders.findIndex(
  - config.gs:910: const classIndex = processedHeaders.findIndex((h) => h.original === refinedResult.classHeader);

## BATCH_OPERATION
- Occurrences: 4
- Files: config.gs, lockManager.gs
- Samples:
  - config.gs:1018: return executeWithStandardizedLock('BATCH_OPERATION', 'republishBoard', () => {
  - config.gs:3449: return executeWithStandardizedLock('BATCH_OPERATION', 'saveAndPublish', () => {
  - lockManager.gs:11: BATCH_OPERATION: 30000, // バッチ処理: 30秒 (saveAndPublish等)

## clearUserInfo
- Occurrences: 4
- Files: config.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:1841: cache.clearUserInfo();
  - unifiedCacheManager.gs:1190: clearUserInfo() {
  - unifiedCacheManager.gs:1208: this.clearUserInfo();

## syncWithUnifiedCache
- Occurrences: 4
- Files: config.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:3530: execCache.syncWithUnifiedCache('dataUpdate');
  - unifiedCacheManager.gs:1239: syncWithUnifiedCache(operation) {
  - unifiedCacheManager.gs:2000: this.executionCache.syncWithUnifiedCache('userDataChange');

## AUTHORIZATION
- Occurrences: 4
- Files: constants.gs, main.gs
- Samples:
  - constants.gs:17: * @property {string} AUTHORIZATION - 認可エラー
  - constants.gs:46: AUTHORIZATION: 'authorization',
  - constants.gs:422: AUTHORIZATION: 'authorization',

## EXECUTION_MAX
- Occurrences: 4
- Files: constants.gs
- Samples:
  - constants.gs:85: EXECUTION_MAX: 300000, // 5分
  - constants.gs:304: /** @deprecated Use UNIFIED_CONSTANTS.TIMEOUTS.EXECUTION_MAX instead */
  - constants.gs:463: EXECUTION_MAX: 300000, // 5分

## MODES
- Occurrences: 4
- Files: constants.gs
- Samples:
  - constants.gs:104: MODES: {
  - constants.gs:257: Object.freeze(UNIFIED_CONSTANTS.DISPLAY.MODES);
  - constants.gs:291: /** @deprecated Use UNIFIED_CONSTANTS.DISPLAY.MODES instead */

## fn
- Occurrences: 4
- Files: constants.gs, unifiedUtilities.gs
- Samples:
  - constants.gs:1196: try { return fn(); }
  - unifiedUtilities.gs:249: const result = fn();
  - unifiedUtilities.gs:262: const result = await fn();

## CRITICAL_OPERATION
- Occurrences: 4
- Files: database.gs, lockManager.gs
- Samples:
  - database.gs:508: 'CRITICAL_OPERATION',
  - database.gs:2944: return executeWithStandardizedLock('CRITICAL_OPERATION', 'deleteUserAccount', async () => {
  - lockManager.gs:10: CRITICAL_OPERATION: 20000, // クリティカルな操作: 20秒

## alerts
- Occurrences: 4
- Files: database.gs
- Samples:
  - database.gs:2845: if (monitoringResult.summary.alerts.length > 0) {
  - database.gs:2847: for (var i = 0; i < monitoringResult.summary.alerts.length; i++) {
  - database.gs:2882: if (monitoringResult.summary.alerts.length > 0) {

## str2
- Occurrences: 4
- Files: database.gs
- Samples:
  - database.gs:3151: for (let i = 0; i <= str2.length; i++) {
  - database.gs:3160: for (let i = 1; i <= str2.length; i++) {
  - database.gs:3162: if (str2.charAt(i - 1) === str1.charAt(j - 1)) {

## str1
- Occurrences: 4
- Files: database.gs
- Samples:
  - database.gs:3155: for (let j = 0; j <= str1.length; j++) {
  - database.gs:3161: for (let j = 1; j <= str1.length; j++) {
  - database.gs:3162: if (str2.charAt(i - 1) === str1.charAt(j - 1)) {

## HTTP
- Occurrences: 4
- Files: main.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - main.gs:36: // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理
  - unifiedCacheManager.gs:1628: // レスポンスオブジェクト検証とHTTPステータスチェック
  - unifiedCacheManager.gs:1634: throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);

## globalProfiler
- Occurrences: 4
- Files: main.gs
- Samples:
  - main.gs:644: if (typeof globalProfiler !== 'undefined') {
  - main.gs:645: globalProfiler.start('logging');
  - main.gs:669: if (typeof globalProfiler !== 'undefined') {

## top
- Occurrences: 4
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:1907: `<script>window.top.location.href = '${sanitizedUrl}';</script>`
  - main.gs:2038: window.top.location.href = url;
  - session-utils.gs:461: // 最も確実な方法：window.top.location.href

## translateY
- Occurrences: 4
- Files: main.gs
- Samples:
  - main.gs:1978: transform: translateY(-2px);
  - main.gs:2607: transform: translateY(-2px);
  - main.gs:2616: transform: translateY(-2px);

## endsWith
- Occurrences: 4
- Files: main.gs, unifiedUtilities.gs
- Samples:
  - main.gs:2114: (cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
  - main.gs:2115: (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))
  - unifiedUtilities.gs:179: (sanitized.startsWith('"') && sanitized.endsWith('"')) ||

## failed
- Occurrences: 4
- Files: resilientExecutor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - resilientExecutor.gs:102: `[ResilientExecutor] ${operationName} failed (attempt ${attempt + 1}). Retrying in ${delay}ms: ${err
  - unifiedCacheManager.gs:1697: `[getHeadersWithRetry] All ${maxRetries} attempts failed. Last error:`,
  - unifiedSecurityManager.gs:77: console.error('[ERROR]', 'Token request failed. Status:', responseCode);

## resetStats
- Occurrences: 4
- Files: resilientExecutor.gs, systemIntegrationManager.gs, unifiedCacheManager.gs
- Samples:
  - resilientExecutor.gs:251: resetStats() {
  - systemIntegrationManager.gs:141: resilientExecutor.resetStats();
  - unifiedCacheManager.gs:726: this.resetStats();

## WEBHOOK_SECRET
- Occurrences: 4
- Files: secretManager.gs
- Samples:
  - secretManager.gs:32: WEBHOOK_SECRET: 'webhook_secret',
  - secretManager.gs:40: WEBHOOK_SECRET: this.secretTypes.WEBHOOK_SECRET,
  - secretManager.gs:40: WEBHOOK_SECRET: this.secretTypes.WEBHOOK_SECRET,

## charCodeAt
- Occurrences: 4
- Files: secretManager.gs
- Samples:
  - secretManager.gs:416: const charCode = value.charCodeAt(i) ^ key.charCodeAt(i % key.length);
  - secretManager.gs:416: const charCode = value.charCodeAt(i) ^ key.charCodeAt(i % key.length);
  - secretManager.gs:443: const charCode = encryptedString.charCodeAt(i) ^ key.charCodeAt(i % key.length);

## deleteAllProperties
- Occurrences: 4
- Files: session-utils.gs
- Samples:
  - session-utils.gs:201: props.deleteAllProperties();
  - session-utils.gs:203: ULog.debug('✅ PropertiesService.deleteAllProperties() でクリア完了');
  - session-utils.gs:346: props.deleteAllProperties();

## deleteError
- Occurrences: 4
- Files: session-utils.gs
- Samples:
  - session-utils.gs:227: authResetResult.errors.push(`プロパティ削除エラー(${key}): ${deleteError.message}`);
  - session-utils.gs:246: `重要プロパティ削除エラー(${key}): ${deleteError.message}`
  - session-utils.gs:362: ULog.warn(`プロパティ削除エラー(${key}): ${deleteError.message}`);

## instance
- Occurrences: 4
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:318: metrics.components.resilientExecutor = component.instance.getStats();
  - systemIntegrationManager.gs:321: metrics.components.unifiedBatchProcessor = component.instance.getMetrics();
  - systemIntegrationManager.gs:324: metrics.components.multiTenantSecurity = component.instance.getSecurityStats();

## getMetrics
- Occurrences: 4
- Files: systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:321: metrics.components.unifiedBatchProcessor = component.instance.getMetrics();
  - unifiedBatchProcessor.gs:560: getMetrics() {
  - unifiedBatchProcessor.gs:649: const metrics = unifiedBatchProcessor.getMetrics();

## getSecurityStats
- Occurrences: 4
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:324: metrics.components.multiTenantSecurity = component.instance.getSecurityStats();
  - unifiedSecurityManager.gs:465: getSecurityStats() {
  - unifiedSecurityManager.gs:578: const stats = multiTenantSecurity.getSecurityStats();

## L1
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:136: ULog.debug(`[Cache] L1(Memo) hit: ${key}`);
  - unifiedCacheManager.gs:136: ULog.debug(`[Cache] L1(Memo) hit: ${key}`);
  - unifiedCacheManager.gs:142: ULog.warn(`[Cache] L1(Memo) error: ${key}`, e.message);

## L2
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:151: ULog.debug(`[Cache] L2(Script) hit: ${key}`);
  - unifiedCacheManager.gs:151: ULog.debug(`[Cache] L2(Script) hit: ${key}`);
  - unifiedCacheManager.gs:162: ULog.warn(`[Cache] L2(Script) error: ${key}`, e.message);

## L3
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:172: ULog.debug(`[Cache] L3(Properties) hit: ${key}`);
  - unifiedCacheManager.gs:172: ULog.debug(`[Cache] L3(Properties) hit: ${key}`);
  - unifiedCacheManager.gs:181: ULog.warn(`[Cache] L3(Properties) error: ${key}`, e.message);

## pendingClears
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:861: this.pendingClears.push({ resolve, reject });
  - unifiedCacheManager.gs:1091: pendingClears: this.pendingClears.length,
  - unifiedCacheManager.gs:1101: const pending = this.pendingClears.splice(0);

## isExpired
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1136: if (this.isExpired()) {
  - unifiedCacheManager.gs:1165: if (this.isExpired()) {
  - unifiedCacheManager.gs:1217: isExpired() {

## clearSheetsService
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1199: clearSheetsService() {
  - unifiedCacheManager.gs:1209: this.clearSheetsService();
  - unifiedCacheManager.gs:2062: this.executionCache.clearSheetsService();

## MEMORY_CACHE_TTL
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1273: MEMORY_CACHE_TTL: 300000, // 5分間（メモリキャッシュ）
  - unifiedCacheManager.gs:1308: if (memoryEntry && now - memoryEntry.timestamp < SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
  - unifiedCacheManager.gs:1414: if (now - entry.timestamp > SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {

## MAX_CACHE_SIZE
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1275: MAX_CACHE_SIZE: 50, // 最大キャッシュエントリ数
  - unifiedCacheManager.gs:1394: if (memoryKeys.length > SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE) {
  - unifiedCacheManager.gs:1403: const entriesToDelete = sortedEntries.slice(SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE);

## global
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1509: if (typeof global !== 'undefined') {
  - unifiedCacheManager.gs:1510: global.cacheManager = cacheManager;
  - unifiedCacheManager.gs:2303: if (typeof global !== 'undefined') {

## clearUserInfoCache
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1996: clearUserInfoCache(identifier = null) {
  - unifiedCacheManager.gs:2146: this.clearUserInfoCache(userId || email);
  - unifiedCacheManager.gs:2181: this.clearUserInfoCache();

## JWT
- Occurrences: 4
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:27: * 新しいJWTトークンを生成
  - unifiedSecurityManager.gs:41: // JWT生成
  - unifiedSecurityManager.gs:42: const jwtHeader = { alg: 'RS256', typ: 'JWT' };

## generateSecureKey
- Occurrences: 4
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:350: generateSecureKey(baseKey, userId, context = '') {
  - unifiedSecurityManager.gs:502: const secureKey = multiTenantSecurity.generateSecureKey(baseKey, userId, options.context);
  - unifiedSecurityManager.gs:601: const secureKey1 = multiTenantSecurity.generateSecureKey('test', testUserId1);

## getCoreSheetData
- Occurrences: 4
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:51: getCoreSheetData(options) {
  - unifiedSheetDataManager.gs:813: return unifiedSheetManager.getCoreSheetData({
  - unifiedSheetDataManager.gs:844: return unifiedSheetManager.getCoreSheetData({

## _getUserInfo
- Occurrences: 4
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:116: const userInfo = this._getUserInfo(userId, options);
  - unifiedSheetDataManager.gs:532: const userInfo = this._getUserInfo(userId, { useExecutionCache: true, ttl: 300 });
  - unifiedSheetDataManager.gs:627: const userInfo = this._getUserInfo(userId, { useExecutionCache: true, ttl: 300 });

## toLocaleString
- Occurrences: 3
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:332: publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
  - Core.gs:333: stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
  - main.gs:1324: `時刻: ${new Date().toLocaleString('ja-JP')}`,

## action
- Occurrences: 3
- Files: Core.gs, secretManager.gs
- Samples:
  - Core.gs:915: * @param {string} adminEmail - Email address of the user performing the action.
  - secretManager.gs:569: if (action.includes('ERROR') || action.includes('FAILED')) {
  - secretManager.gs:569: if (action.includes('ERROR') || action.includes('FAILED')) {

## main
- Occurrences: 3
- Files: Core.gs, config.gs, main.gs
- Samples:
  - Core.gs:2374: // include 関数は main.gs で定義されています
  - config.gs:10: // メモリ管理用の実行レベル変数 (main.gsと統一)
  - main.gs:2790: ULog.warn('🔧 main.gs: publishedSheetNameを空文字にリセットしました');

## formMoveError
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:2800: moveErrors.push('フォームファイル移動エラー: ' + formMoveError.message);
  - Core.gs:2801: ULog.error('[ERROR]', '❌ フォームファイルの移動に失敗:', formMoveError.message);
  - Core.gs:5775: ULog.warn('カスタムフォームファイル移動エラー:', formMoveError.message);

## ssMoveError
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:2843: moveErrors.push('スプレッドシートファイル移動エラー: ' + ssMoveError.message);
  - Core.gs:2844: ULog.error('[ERROR]', '❌ スプレッドシートファイルの移動に失敗:', ssMoveError.message);
  - Core.gs:5805: ULog.warn('カスタムスプレッドシートファイル移動エラー:', ssMoveError.message);

## HH
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:3697: var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  - Core.gs:3943: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - Core.gs:4069: const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

## getEditUrl
- Occurrences: 3
- Files: Core.gs, config.gs, constants.gs
- Samples:
  - Core.gs:3742: editFormUrl: typeof form.getEditUrl === 'function' ? form.getEditUrl() : '',
  - config.gs:1652: ULog.debug('📝 フォーム作成成功:', form.getEditUrl());
  - constants.gs:1292: getEditUrl: (form) => form ? form.getEditUrl() : ''

## choices
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:3795: customConfig.classQuestion.choices.length > 0
  - Core.gs:3824: customConfig.mainQuestion.choices.length > 0
  - Core.gs:3841: customConfig.mainQuestion.choices.length > 0

## assign
- Occurrences: 3
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3927: Object.assign(config[key], customConfig[key]);
  - config.gs:2368: userInfo: Object.assign({}, userInfo), // 浅いコピー（userInfoは単純オブジェクト）
  - config.gs:3248: Object.assign(configJson, topLevelUpdates);

## SPREADSHEET
- Occurrences: 3
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:4145: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  - config.gs:1601: if (destinationType === FormApp.DestinationType.SPREADSHEET) {
  - config.gs:1703: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

## DestinationType
- Occurrences: 3
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:4145: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  - config.gs:1601: if (destinationType === FormApp.DestinationType.SPREADSHEET) {
  - config.gs:1703: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

## TIMESTAMP
- Occurrences: 3
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:4854: var timestampIndex = headerIndices[COLUMN_HEADERS.TIMESTAMP];
  - constants.gs:90: TIMESTAMP: 'タイムスタンプ',
  - constants.gs:279: TIMESTAMP: 'タイムスタンプ',

## componentsInitialized
- Occurrences: 3
- Files: autoInit.gs, systemIntegrationManager.gs
- Samples:
  - autoInit.gs:32: componentsInitialized: initResult.componentsInitialized.length,
  - systemIntegrationManager.gs:62: initResult.componentsInitialized.push(componentName);
  - systemIntegrationManager.gs:105: initialized: initResult.componentsInitialized.length,

## getAuditLog
- Occurrences: 3
- Files: autoInit.gs, secretManager.gs, unifiedSecurityManager.gs
- Samples:
  - autoInit.gs:183: const auditLog = unifiedSecretManager.getAuditLog();
  - secretManager.gs:578: getAuditLog() {
  - unifiedSecurityManager.gs:942: const auditLog = unifiedSecretManager.getAuditLog();

## _startMemoryMonitoring
- Occurrences: 3
- Files: config.gs
- Samples:
  - config.gs:55: this._startMemoryMonitoring();
  - config.gs:84: _startMemoryMonitoring() {
  - config.gs:88: this._startMemoryMonitoring(); // 再帰的監視

## destroy
- Occurrences: 3
- Files: config.gs
- Samples:
  - config.gs:100: this.destroy();
  - config.gs:136: destroy() {
  - config.gs:231: context.destroy();

## unregister
- Occurrences: 3
- Files: config.gs
- Samples:
  - config.gs:228: unregister(userId) {
  - config.gs:249: this.unregister(oldestContext);
  - config.gs:255: this.unregister(userId);

## getFormUrl
- Occurrences: 3
- Files: config.gs
- Samples:
  - config.gs:1503: const form = sheet.getFormUrl();
  - config.gs:1509: // getFormUrl()がエラーになる場合があるのでcatchする
  - config.gs:3315: const url = targetSheet.getFormUrl();

## Login
- Occurrences: 3
- Files: config.gs, main.gs
- Samples:
  - config.gs:2039: * Login.htmlから呼び出される
  - config.gs:2074: * Login.htmlから呼び出される
  - main.gs:679: * AdminPanel.html と Login.html から共通で呼び出される

## batchError
- Occurrences: 3
- Files: config.gs, database.gs
- Samples:
  - config.gs:3131: console.error('[ERROR]', '❌ batchGetSheetsData failed:', batchError.message);
  - config.gs:3132: throw new Error('ヘッダー行の取得に失敗: ' + batchError.message);
  - database.gs:3055: console.error('[ERROR]', '❌ Deletion operation failed:', batchError.message);

## AUTHENTICATION
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:16: * @property {string} AUTHENTICATION - 認証エラー
  - constants.gs:45: AUTHENTICATION: 'authentication',
  - constants.gs:421: AUTHENTICATION: 'authentication',

## USER_INPUT
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:23: * @property {string} USER_INPUT - ユーザー入力エラー
  - constants.gs:52: USER_INPUT: 'user_input',
  - constants.gs:428: USER_INPUT: 'user_input',

## EXTERNAL
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:24: * @property {string} EXTERNAL - 外部サービスエラー
  - constants.gs:53: EXTERNAL: 'external',
  - constants.gs:429: EXTERNAL: 'external',

## EXTENDED
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:66: EXTENDED: 21600, // 6時間
  - constants.gs:438: * @param {string} duration - 期間（SHORT, MEDIUM, LONG, EXTENDED）
  - constants.gs:446: EXTENDED: 21600, // 6時間

## LARGE
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:73: LARGE: 100,
  - constants.gs:307: /** @deprecated Use UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.LARGE instead */
  - constants.gs:489: return getConstant('CACHE.BATCH_SIZE.LARGE', 100);

## COLUMNS
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:89: COLUMNS: {
  - constants.gs:255: Object.freeze(UNIFIED_CONSTANTS.COLUMNS);
  - constants.gs:277: /** @deprecated Use UNIFIED_CONSTANTS.COLUMNS instead */

## DELETE_LOG
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:129: DELETE_LOG: {
  - constants.gs:260: Object.freeze(UNIFIED_CONSTANTS.SHEETS.DELETE_LOG);
  - constants.gs:332: /** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DELETE_LOG instead */

## DIAGNOSTIC_LOG
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:142: DIAGNOSTIC_LOG: {
  - constants.gs:261: Object.freeze(UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG);
  - constants.gs:338: /** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG instead */

## RETRY_COUNT
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:199: RETRY_COUNT: 3,
  - constants.gs:762: retries = UNIFIED_CONSTANTS.LIMITS.RETRY_COUNT,
  - constants.gs:804: maxRetries = UNIFIED_CONSTANTS.LIMITS.RETRY_COUNT,

## SPREADSHEET_ID
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:241: SPREADSHEET_ID: /^[a-zA-Z0-9-_]{44}$/,
  - constants.gs:683: return UNIFIED_CONSTANTS.REGEX.SPREADSHEET_ID.test(spreadsheetId.trim());
  - constants.gs:683: return UNIFIED_CONSTANTS.REGEX.SPREADSHEET_ID.test(spreadsheetId.trim());

## USER_ID
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:242: USER_ID: /^[a-f0-9]{32}$/,
  - constants.gs:695: return UNIFIED_CONSTANTS.REGEX.USER_ID.test(userId.trim());
  - constants.gs:695: return UNIFIED_CONSTANTS.REGEX.USER_ID.test(userId.trim());

## MAIN
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:494: * @param {string} type - 質問タイプ（MAIN, REASON）
  - constants.gs:499: MAIN: 'あなたの考えや気づいたことを教えてください',
  - constants.gs:502: return mapping[type] || mapping.MAIN;

## validateRequired
- Occurrences: 3
- Files: constants.gs, workflowValidation.gs
- Samples:
  - constants.gs:641: validateRequired(params, required) {
  - constants.gs:709: const requiredResult = this.validateRequired(data, rules.required);
  - workflowValidation.gs:13: const validation = UValidate.validateRequired({ data }, ['data']);

## executeWithTimeout
- Occurrences: 3
- Files: constants.gs, resilientExecutor.gs
- Samples:
  - constants.gs:758: executeWithTimeout(operation, options = {}) {
  - resilientExecutor.gs:76: const result = this.executeWithTimeout(operation, this.config.timeoutMs);
  - resilientExecutor.gs:134: async executeWithTimeout(operation, timeoutMs) {

## ALERT
- Occurrences: 3
- Files: constants.gs
- Samples:
  - constants.gs:1257: console.log('[ALERT]', message);
  - constants.gs:1258: Logger.log('[ALERT] ' + message);
  - constants.gs:1262: console.log('[ALERT]', message);

## steps
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:234: transactionLog.steps.push('診断ログシートヘッダーを作成');
  - database.gs:246: transactionLog.steps.push('診断ログシートを新規作成');
  - database.gs:275: transactionLog.steps.push('診断ログエントリーを追加');

## UNFORMATTED_VALUE
- Occurrences: 3
- Files: database.gs, unifiedBatchProcessor.gs
- Samples:
  - database.gs:1388: valueRenderOption: 'UNFORMATTED_VALUE',
  - unifiedBatchProcessor.gs:48: valueRenderOption = 'UNFORMATTED_VALUE',
  - unifiedBatchProcessor.gs:147: responseValueRenderOption = 'UNFORMATTED_VALUE',

## RAW
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:1461: valueInputOption: 'RAW',
  - database.gs:1506: ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';
  - database.gs:1783: '?valueInputOption=RAW';

## spreadsheetAccess
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:2116: result.checks.spreadsheetAccess.canWrite = true;
  - database.gs:2118: result.checks.spreadsheetAccess.canWrite = false;
  - database.gs:2119: result.checks.spreadsheetAccess.writeError = writeError.message;

## then
- Occurrences: 3
- Files: lockManager.gs, resilientExecutor.gs
- Samples:
  - lockManager.gs:71: return result.then(
  - resilientExecutor.gs:141: .then(() => operation())
  - resilientExecutor.gs:142: .then((result) => {

## WORKFLOW
- Occurrences: 3
- Files: main.gs, unifiedValidationSystem.gs
- Samples:
  - main.gs:148: ULog.info(`自動停止実行完了: ${userInfo.adminEmail} (期限: ${config.scheduledEndAt})`, {}, ULog.CATEGORIES.WOR
  - unifiedValidationSystem.gs:16: WORKFLOW: 'workflow'
  - unifiedValidationSystem.gs:55: case this.validationCategories.WORKFLOW:

## performanceMetrics
- Occurrences: 3
- Files: main.gs
- Samples:
  - main.gs:1264: systemDiagnostics.performanceMetrics.authDuration = authDuration + 'ms';
  - main.gs:1268: systemDiagnostics.performanceMetrics.totalRequestTime = totalRequestTime + 'ms';
  - main.gs:1340: systemDiagnostics.performanceMetrics.totalRequestTime = totalRequestTime + 'ms';

## line
- Occurrences: 3
- Files: main.gs
- Samples:
  - main.gs:1718: const trimmed = line.trim();
  - main.gs:1724: const jsonStart = line.indexOf('{');
  - main.gs:1727: const jsonStr = line.substring(jsonStart);

## DOCTYPE
- Occurrences: 3
- Files: main.gs
- Samples:
  - main.gs:1925: <!DOCTYPE html>
  - main.gs:2401: <!DOCTYPE html>
  - main.gs:2467: <!DOCTYPE html>

## googleusercontent
- Occurrences: 3
- Files: main.gs
- Samples:
  - main.gs:2137: // 開発モードURLのチェック（googleusercontent.comは有効なデプロイURLも含むため調整）
  - main.gs:2143: // 最終的な URL 妥当性チェック（googleusercontent.comも有効URLとして認識）
  - main.gs:2146: cleanUrl.includes('googleusercontent.com') ||

## DATABASE_CREDS
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:31: DATABASE_CREDS: 'database_creds',
  - secretManager.gs:39: DATABASE_SPREADSHEET_ID: this.secretTypes.DATABASE_CREDS,
  - secretManager.gs:486: case this.secretTypes.DATABASE_CREDS:

## getSecret
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:50: getSecret(secretName, options = {}) {
  - secretManager.gs:657: const value = this.getSecret(secretName, { auditLog: false });
  - secretManager.gs:720: return unifiedSecretManager.getSecret(key);

## cacheSecret
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:91: this.cacheSecret(secretName, secretValue);
  - secretManager.gs:116: this.cacheSecret(secretName, secretValue);
  - secretManager.gs:531: cacheSecret(secretName, secretValue) {

## isCriticalSecret
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:160: if (this.isCriticalSecret(secretName)) {
  - secretManager.gs:375: if (this.isCriticalSecret(secretName)) {
  - secretManager.gs:466: isCriticalSecret(secretName) {

## encryptValue
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:197: const valueToStore = encrypt ? this.encryptValue(secretValue) : secretValue;
  - secretManager.gs:409: encryptValue(value) {
  - secretManager.gs:632: const encrypted = this.encryptValue(testValue);

## secretmanager
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:225: `https://secretmanager.googleapis.com/v1/${secretPath}:access`,
  - secretManager.gs:263: const secretsListUrl = `https://secretmanager.googleapis.com/v1/projects/${this.config.projectId}/se
  - secretManager.gs:601: `https://secretmanager.googleapis.com/v1/projects/${this.config.projectId}/secrets`,

## decryptValue
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:362: value = this.decryptValue(value);
  - secretManager.gs:431: decryptValue(encryptedValue) {
  - secretManager.gs:633: const decrypted = this.decryptValue(encrypted);

## ENC
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:420: return 'ENC:' + Utilities.base64Encode(encrypted);
  - secretManager.gs:433: if (!encryptedValue.startsWith('ENC:')) {
  - secretManager.gs:459: return typeof value === 'string' && value.startsWith('ENC:');

## componentsFailedToInitialize
- Occurrences: 3
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:65: initResult.componentsFailedToInitialize.push({
  - systemIntegrationManager.gs:95: initResult.success = initResult.componentsFailedToInitialize.length === 0;
  - systemIntegrationManager.gs:106: failed: initResult.componentsFailedToInitialize.length,

## updateCacheMetrics
- Occurrences: 3
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:58: this.updateCacheMetrics(true);
  - unifiedBatchProcessor.gs:129: this.updateCacheMetrics(false);
  - unifiedBatchProcessor.gs:539: updateCacheMetrics(hit) {

## concat
- Occurrences: 3
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:100: allValueRanges = allValueRanges.concat(chunkResult.valueRanges || []);
  - unifiedBatchProcessor.gs:307: allReplies = allReplies.concat(chunkResult.replies || []);
  - unifiedBatchProcessor.gs:374: allResults = allResults.concat(chunkResults);

## spreadsheetCache
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:3: * cache.gs、unifiedExecutionCache.gs、spreadsheetCache.gs の機能を統合
  - unifiedCacheManager.gs:1265: // SECTION 3: SpreadsheetApp最適化システム（元spreadsheetCache.gs）
  - unifiedCacheManager.gs:2234: // メモリキャッシュから削除（spreadsheetCache.gsの機能）

## DRY
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:514: `🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`
  - unifiedCacheManager.gs:527: ULog.debug(`[Cache] DRY RUN: Would clear pattern: ${pattern}`);
  - unifiedCacheManager.gs:547: ULog.debug(`[Cache] DRY RUN: Would clear related pattern: ${pattern}`);

## RUN
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:514: `🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`
  - unifiedCacheManager.gs:527: ULog.debug(`[Cache] DRY RUN: Would clear pattern: ${pattern}`);
  - unifiedCacheManager.gs:547: ULog.debug(`[Cache] DRY RUN: Would clear related pattern: ${pattern}`);

## _getInvalidationPatterns
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:518: const patterns = this._getInvalidationPatterns(entityType, entityId);
  - unifiedCacheManager.gs:541: const relatedPatterns = this._getInvalidationPatterns(entityType, relatedId);
  - unifiedCacheManager.gs:591: _getInvalidationPatterns(entityType, entityId) {

## unifiedCache
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:882: typeof window.unifiedCache.clear === 'function'
  - unifiedCacheManager.gs:884: window.unifiedCache.clear();
  - unifiedCacheManager.gs:1059: typeof window !== 'undefined' && !!(window.unifiedCache && window.unifiedCache.clear),

## gasOptimizer
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:896: typeof window.gasOptimizer.clearCache === 'function'
  - unifiedCacheManager.gs:898: window.gasOptimizer.clearCache();
  - unifiedCacheManager.gs:1064: typeof window !== 'undefined' && !!(window.gasOptimizer && window.gasOptimizer.clearCache),

## SESSION_CACHE_TTL
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1274: SESSION_CACHE_TTL: 1800000, // 30分間（PropertiesServiceキャッシュ）
  - unifiedCacheManager.gs:1321: if (now - sessionEntry.timestamp < SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL) {
  - unifiedCacheManager.gs:1487: sessionTTL: SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL,

## base64EncodeWebSafe
- Occurrences: 3
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:51: const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  - unifiedSecurityManager.gs:52: const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
  - unifiedSecurityManager.gs:55: const encodedSignature = Utilities.base64EncodeWebSafe(signature);

## logDataAccess
- Occurrences: 3
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:299: this.logDataAccess('TENANT_BOUNDARY_VALID', { requestUserId, operation });
  - unifiedSecurityManager.gs:306: this.logDataAccess('ADMIN_ACCESS_GRANTED', { requestUserId, targetUserId, operation });
  - unifiedSecurityManager.gs:435: logDataAccess(accessType, details) {

## _executeDataFetch
- Occurrences: 3
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:86: return this._executeDataFetch(options);
  - unifiedSheetDataManager.gs:92: () => this._executeDataFetch(options),
  - unifiedSheetDataManager.gs:101: _executeDataFetch(options) {

## _fetchHeadersWithRetry
- Occurrences: 3
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:340: () => this._fetchHeadersWithRetry(spreadsheetId, sheetName, maxRetries),
  - unifiedSheetDataManager.gs:388: _fetchHeadersWithRetry(spreadsheetId, sheetName, maxRetries) {
  - unifiedSheetDataManager.gs:498: const recoveredHeaders = this._fetchHeadersWithRetry(spreadsheetId, sheetName, 1);

## _validateRequiredHeaders
- Occurrences: 3
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:348: () => this._validateRequiredHeaders(headers),
  - unifiedSheetDataManager.gs:452: _validateRequiredHeaders(indices) {
  - unifiedSheetDataManager.gs:500: const recoveredValidation = this._validateRequiredHeaders(recoveredHeaders);

## field
- Occurrences: 3
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:140: if (field.toLowerCase().includes('email')) {
  - unifiedUserManager.gs:162: (header) => header && header.toString().toLowerCase() === field.toLowerCase()
  - unifiedUserManager.gs:175: if (field.toLowerCase().includes('email')) {

## _sanitizeOptions
- Occurrences: 3
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:878: options: this._sanitizeOptions(options),
  - unifiedValidationSystem.gs:927: options: this._sanitizeOptions(options),
  - unifiedValidationSystem.gs:939: _sanitizeOptions(options) {

## Deployments
- Occurrences: 3
- Files: url.gs
- Samples:
  - url.gs:55: deploymentsList = AppsScript.Script.Deployments.list(scriptId, {
  - url.gs:61: deploymentsList = AppsScript.Script.Deployments.list();
  - url.gs:65: 'AppsScript.Script.Deployments.list response:',

## deploymentConfig
- Occurrences: 3
- Files: url.gs
- Samples:
  - url.gs:56: fields: 'deployments(deploymentId,deploymentConfig(webApp(url)),updateTime)',
  - url.gs:77: .filter((d) => d.deploymentConfig && d.deploymentConfig.webApp)
  - url.gs:81: finalUrl = webAppDeployments[0].deploymentConfig.webApp.url;

## _buildErrorInfo
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:48: const errorInfo = this._buildErrorInfo(error, context, severity, category, metadata);
  - Core.gs:141: _buildErrorInfo(error, context, severity, category, metadata) {

## _isRetryableError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:80: retryable: this._isRetryableError(error),
  - Core.gs:174: _isRetryableError(error) {

## ISO
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:320: * @param {string} publishedAt - 公開開始時間のISO文字列
  - Core.gs:6911: * @returns {string|null} 作成日のISO文字列、取得失敗時はnull

## configJsonString
- Occurrences: 2
- Files: Core.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:539: if (configJsonString && configJsonString.trim() !== '' && configJsonString !== '{}') {
  - unifiedSecurityManager.gs:245: if (!configJsonString || configJsonString.trim() === '' || configJsonString === '{}') {

## column
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:834: headerValidationData.issues.push('Reason column (理由) not found in headers');
  - Core.gs:839: headerValidationData.issues.push('Opinion column (回答) not found in headers');

## method
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:1048: createdUser = method();
  - Core.gs:6303: existingUser = stage.method();

## stageError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:1068: error: stageError.message,
  - Core.gs:6316: error: stageError.message,

## countError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2185: ULog.warn('回答数の取得に失敗: ' + countError.message);
  - Core.gs:2495: ULog.warn('回答数の取得に失敗: ' + countError.message);

## setDescription
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2646: form.setDescription(description);
  - Core.gs:3704: form.setDescription(formDescription);

## generalError
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:2868: ULog.error('[ERROR]', '❌ ファイル移動処理で予期しないエラー:', generalError.message);
  - main.gs:992: errors.push(`システムチェック中の一般エラー: ${generalError.message}`);

## getFoldersByName
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3342: var folders = DriveApp.getFoldersByName(rootFolderName);
  - Core.gs:3350: var userFolders = rootFolder.getFoldersByName(userFolderName);

## createFolder
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3346: rootFolder = DriveApp.createFolder(rootFolderName);
  - Core.gs:3354: return rootFolder.createFolder(userFolderName);

## rootFolder
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3350: var userFolders = rootFolder.getFoldersByName(userFolderName);
  - Core.gs:3354: return rootFolder.createFolder(userFolderName);

## folderMoveError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3715: ULog.warn('⚠️ フォーム作成直後の移動に失敗（後で再移動されます）:', folderMoveError.message);
  - Core.gs:4123: folderMoveError.message

## setCollectEmail
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3726: form.setCollectEmail(true);
  - Core.gs:3763: form.setCollectEmail(false);

## Forms
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3755: * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
  - Core.gs:4096: * @param {GoogleAppsScript.Forms.Form} form - フォーム

## addListItem
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3766: var classItem = form.addListItem();
  - Core.gs:3797: var classItem = form.addListItem();

## showOtherOption
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3833: mainItem.showOtherOption(true);
  - Core.gs:3850: mainItem.showOtherOption(true);

## getScriptTimeZone
- Occurrences: 2
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:3943: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - workflowValidation.gs:147: const formatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

## createTextOutput
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3950: return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(
  - Core.gs:3955: return ContentService.createTextOutput(

## setMimeType
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3950: return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(
  - Core.gs:3960: ).setMimeType(ContentService.MimeType.JSON);

## MimeType
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3951: ContentService.MimeType.JSON
  - Core.gs:3960: ).setMimeType(ContentService.MimeType.JSON);

## setSharing
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:4135: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  - Core.gs:4334: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## Access
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:4135: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  - Core.gs:4334: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## Permission
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:4135: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  - Core.gs:4334: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## setDestination
- Occurrences: 2
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:4145: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  - config.gs:1703: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

## docs
- Occurrences: 2
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:4225: spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
  - config.gs:1549: if (typeof cellValue === 'string' && cellValue.includes('docs.google.com/forms/')) {

## driveError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:4328: ULog.error('[ERROR]', 'DriveApp.getFileById error:', driveError.message);
  - Core.gs:4329: throw new Error('スプレッドシートへのアクセスに失敗しました: ' + driveError.message);

## setValues
- Occurrences: 2
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:4403: sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
  - workflowValidation.gs:74: testRange.setValues(testData);

## status
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:4426: * @returns {object} status ('success' or 'error') と message
  - main.gs:844: * Retrieves the administrator domain for the login page with domain match status.

## repairError
- Occurrences: 2
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:4687: ULog.error('[ERROR]', 'getSheetsList: 権限修復に失敗:', repairError.message);
  - database.gs:2168: result.summary.actions.push('権限修復失敗: ' + repairError.message);

## LIKE_MULTIPLIER_FACTOR
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:4869: var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
  - main.gs:85: LIKE_MULTIPLIER_FACTOR: 0.1,

## RANDOM_SCORE_FACTOR
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:4878: var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
  - main.gs:86: RANDOM_SCORE_FACTOR: 0.01,

## actualHeader
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:4993: var normalizedActualHeader = actualHeader.toLowerCase().trim();
  - Core.gs:5009: var normalizedActualHeader = actualHeader.toLowerCase().trim();

## DATABASE_UPDATE_FAILED
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:5276: return createErrorResponse('DATABASE_UPDATE_FAILED', result.message);
  - Core.gs:5466: return createErrorResponse(result.message, result.error || 'DATABASE_UPDATE_FAILED');

## folderError
- Occurrences: 2
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:5811: ULog.warn('カスタムセットアップフォルダ処理エラー:', folderError.message);
  - database.gs:1160: folderError.message

## stage2Error
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:6110: searchAttempts.push({ method: 'findUserByEmail', error: stage2Error.message });
  - Core.gs:6111: ULog.warn('getLoginStatus: Stage 2失敗:', stage2Error.message);

## stage3Error
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:6130: error: stage3Error.message,
  - Core.gs:6132: ULog.warn('getLoginStatus: Stage 3失敗:', stage3Error.message);

## verifyError
- Occurrences: 2
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:6385: ULog.warn('confirmUserRegistration: 登録後検証でエラー:', verifyError.message);
  - database.gs:3074: console.error('[ERROR]', '⚠️ Deletion verification failed:', verifyError.message);

## registrationError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:6422: error: registrationError.message,
  - Core.gs:6428: message: registrationError.message,

## err
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:6531: ULog.warn('Answer count retrieval failed:', err.message);
  - main.gs:2823: ULog.warn('アクセス権設定警告:', err.message);

## sheetErr
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:6685: ULog.warn('Sheet details retrieval failed:', sheetErr.message);
  - Core.gs:6686: response.sheetDetailsError = sheetErr.message;

## _scheduleAutoCleanup
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:52: this._scheduleAutoCleanup();
  - config.gs:96: _scheduleAutoCleanup() {

## isHealthy
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:188: isHealthy() {
  - config.gs:208: isHealthy: this.isHealthy(),

## _cleanupOldest
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:221: this._cleanupOldest();
  - config.gs:237: _cleanupOldest() {

## clearExpired
- Occurrences: 2
- Files: config.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:275: cacheManager.clearExpired();
  - unifiedCacheManager.gs:693: clearExpired() {

## every
- Occurrences: 2
- Files: config.gs, unifiedValidationSystem.gs
- Samples:
  - config.gs:844: hasShortText: columnData.every((val) => val.length < 20),
  - unifiedValidationSystem.gs:1064: overallSuccess: Object.values(results).every(r => r.success),

## item
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:1686: item.setTitle(headerStr);
  - config.gs:1688: item.setRequired(true); // 最初の質問は必須

## spreadsheetError
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:2265: spreadsheetError.message
  - config.gs:2268: `ヘッダー取得に失敗しました。Sheets API: ${apiError.message}, SpreadsheetApp: ${spreadsheetError.message}`

## pendingUpdates
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:2463: if (context.pendingUpdates.spreadsheetId) {
  - config.gs:2465: cacheManager.clearByPattern(`batchGet_${context.pendingUpdates.spreadsheetId}_*`);

## string
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:2892: 'Parameter order mismatch: context must be an object, not a string. Check function call parameters.'
  - config.gs:2932: 'getSheetDetailsFromContext: sheetName parameter must be a valid non-empty string. Received: ' +

## spreadsheets
- Occurrences: 2
- Files: config.gs, database.gs
- Samples:
  - config.gs:3091: typeof context.sheetsService.spreadsheets.get === 'function'
  - database.gs:1390: // includeGridData は values:batchGet では無効（spreadsheets.get 系のみ）

## setUserInfo
- Occurrences: 2
- Files: config.gs, unifiedCacheManager.gs
- Samples:
  - config.gs:3524: execCache.setUserInfo(context.requestUserId, freshUserInfo);
  - unifiedCacheManager.gs:1154: setUserInfo(userId, userInfo) {

## APPLICATION_ENABLED
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:3818: const value = props.getProperty('APPLICATION_ENABLED');
  - config.gs:3849: props.setProperty('APPLICATION_ENABLED', enabledValue);

## POLLING
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:83: POLLING: 30000, // 30秒
  - constants.gs:461: POLLING: 30000, // 30秒

## QUEUE
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:84: QUEUE: 60000, // 1分
  - constants.gs:462: QUEUE: 60000, // 1分

## PRESETS
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:170: PRESETS: {
  - constants.gs:264: Object.freeze(UNIFIED_CONSTANTS.FORMS.PRESETS);

## HISTORY_ITEMS
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:198: HISTORY_ITEMS: 50,
  - constants.gs:310: /** @deprecated Use UNIFIED_CONSTANTS.LIMITS.HISTORY_ITEMS instead */

## WEIGHTS
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:210: WEIGHTS: {
  - constants.gs:267: Object.freeze(UNIFIED_CONSTANTS.SCORING.WEIGHTS);

## BONUSES
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:217: BONUSES: {
  - constants.gs:268: Object.freeze(UNIFIED_CONSTANTS.SCORING.BONUSES);

## SCRIPT_PROPS
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:224: SCRIPT_PROPS: {
  - constants.gs:269: Object.freeze(UNIFIED_CONSTANTS.SCRIPT_PROPS);

## CLIENT_ID
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:227: CLIENT_ID: 'CLIENT_ID',
  - constants.gs:227: CLIENT_ID: 'CLIENT_ID',

## ENABLED
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:233: ENABLED: PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true',
  - constants.gs:363: /** @deprecated Use UNIFIED_CONSTANTS.DEBUG.ENABLED instead */

## instead
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:313: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead (defined in Core.gs) */
  - constants.gs:315: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION instead (defined in Core.gs) */

## DEFAULT_QUESTIONS
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:526: DEFAULT_QUESTIONS: UNIFIED_CONSTANTS.FORMS.DEFAULT_QUESTIONS,
  - constants.gs:526: DEFAULT_QUESTIONS: UNIFIED_CONSTANTS.FORMS.DEFAULT_QUESTIONS,

## validateEmail
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:667: validateEmail(email) {
  - constants.gs:721: if (validator.type === 'email' && !this.validateEmail(value)) {

## validateSpreadsheetId
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:679: validateSpreadsheetId(spreadsheetId) {
  - constants.gs:724: if (validator.type === 'spreadsheetId' && !this.validateSpreadsheetId(value)) {

## validateUserId
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:691: validateUserId(userId) {
  - constants.gs:727: if (validator.type === 'userId' && !this.validateUserId(value)) {

## items
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:989: totalItems: items.length,
  - constants.gs:999: total: items.length,

## ES2019
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:1030: * ES2019のObject.fromEntriesポリフィル
  - constants.gs:1314: console.log('   - Object.fromEntries: ES2019ポリフィル');

## unifiedValidationSystem
- Occurrences: 2
- Files: constants.gs, unifiedUtilities.gs
- Samples:
  - constants.gs:1244: // UnifiedValidationSystem は unifiedValidationSystem.gs で定義済み
  - unifiedUtilities.gs:288: // これらの機能は既に unifiedValidationSystem.gs 等で実装されています

## msgBox
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:1255: return Browser.msgBox(message);
  - constants.gs:1270: const response = Browser.msgBox(message, Browser.Buttons.YES_NO);

## Buttons
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:1270: const response = Browser.msgBox(message, Browser.Buttons.YES_NO);
  - constants.gs:1271: return response === Browser.Buttons.YES;

## CONFIRM
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:1273: console.log('[CONFIRM]', message, '(auto-confirmed)');
  - constants.gs:1277: console.log('[CONFIRM]', message, '(auto-confirmed)');

## DOM
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:1282: // DOM/Browser操作のダミー実装（GASでは不要）
  - constants.gs:1317: console.log('   - DOM操作関数: ダミー実装');

## I1
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:229: range: `'${logSheetName}'!A1:I1`,
  - database.gs:241: range: `'${logSheetName}'!A1:I1`,

## INSERT_ROWS
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:270: insertDataOption: 'INSERT_ROWS',
  - database.gs:1506: ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';

## ROWS
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:567: dimension: 'ROWS',
  - database.gs:3008: dimension: 'ROWS',

## req
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:973: '  ' + (index + 1) + '. 範囲: ' + req.range + ', 値: ' + JSON.stringify(req.values)
  - database.gs:973: '  ' + (index + 1) + '. 範囲: ' + req.range + ', 値: ' + JSON.stringify(req.values)

## SERIAL_NUMBER
- Occurrences: 2
- Files: database.gs, unifiedBatchProcessor.gs
- Samples:
  - database.gs:1389: dateTimeRenderOption: 'SERIAL_NUMBER',
  - unifiedBatchProcessor.gs:47: dateTimeRenderOption = 'SERIAL_NUMBER',

## dataError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:1924: error: dataError.message,
  - database.gs:1926: diagnosticResult.summary.criticalIssues.push('ユーザーデータ取得失敗: ' + dataError.message);

## saError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2092: error: saError.message,
  - database.gs:2094: result.summary.issues.push('サービスアカウント設定エラー: ' + saError.message);

## writeError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2119: result.checks.spreadsheetAccess.writeError = writeError.message;
  - database.gs:2120: result.summary.issues.push('スプレッドシート書き込み権限不足: ' + writeError.message);

## retestError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2161: error: retestError.message,
  - database.gs:2164: result.summary.actions.push('修復後テスト失敗: ' + retestError.message);

## fixError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2307: console.error('[ERROR]', '❌ 自動修復エラー:', fixError.message);
  - database.gs:2308: result.summary.issues.push('自動修復に失敗: ' + fixError.message);

## configError
- Occurrences: 2
- Files: database.gs, unifiedCacheManager.gs
- Samples:
  - database.gs:2693: error: configError.message,
  - unifiedCacheManager.gs:1843: results.errors.push('sheet_headers: ' + configError.message);

## benchmarks
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2738: perfResult.benchmarks.databaseAccess = Date.now() - dbTestStart;
  - database.gs:2752: perfResult.benchmarks.cacheCheck = Date.now() - cacheTestStart;

## perfError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2754: console.error('[ERROR]', 'パフォーマンステスト実行エラー:', perfError.message);
  - database.gs:2755: perfResult.error = perfError.message;

## securityError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2825: console.error('[ERROR]', 'セキュリティチェック実行エラー:', securityError.message);
  - database.gs:2826: securityResult.error = securityError.message;

## charAt
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:3162: if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
  - database.gs:3162: if (str2.charAt(i - 1) === str1.charAt(j - 1)) {

## READ_OPERATION
- Occurrences: 2
- Files: lockManager.gs
- Samples:
  - lockManager.gs:8: READ_OPERATION: 5000, // 読み取り専用操作: 5秒
  - lockManager.gs:16: * @param {string} operationType - 操作タイプ (READ_OPERATION, WRITE_OPERATION, CRITICAL_OPERATION, BATCH_

## getContent
- Occurrences: 2
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:13: return HtmlService.createHtmlOutputFromFile(path).getContent();
  - session-utils.gs:509: const outputContent = htmlOutput.getContent();

## sessionError
- Occurrences: 2
- Files: main.gs, unifiedCacheManager.gs
- Samples:
  - main.gs:1215: error: sessionError.message
  - unifiedCacheManager.gs:1367: ULog.debug('セッションキャッシュ保存エラー:', sessionError.message);

## title
- Occurrences: 2
- Files: main.gs, unifiedSheetDataManager.gs
- Samples:
  - main.gs:1658: const titleLower = title.toLowerCase();
  - unifiedSheetDataManager.gs:645: !sheet.title.startsWith('名簿') // 名簿シート除外

## technical
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:1730: structured.technical.push(line);
  - main.gs:1745: structured.technical.push(trimmed);

## event
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:2036: event.preventDefault();
  - main.gs:2721: const button = event.target;

## style
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:2059: mainButton.style.animation = 'pulse 1s infinite';
  - main.gs:2060: mainButton.style.boxShadow = '0 0 20px rgba(16, 185, 129, 0.5)';

## script
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:2145: cleanUrl.includes('script.google.com') ||
  - main.gs:2726: google.script.run

## UTF
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:2404: <meta charset="UTF-8">
  - main.gs:2470: <meta charset="UTF-8">

## DEBUG_MODE_LAST_MODIFIED
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:2996: PropertiesService.getScriptProperties().getProperty('DEBUG_MODE_LAST_MODIFIED') ||
  - main.gs:3037: DEBUG_MODE_LAST_MODIFIED: new Date().toISOString(),

## handleSuccess
- Occurrences: 2
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:79: this.handleSuccess(operationName);
  - resilientExecutor.gs:198: handleSuccess(operationName) {

## isRetryableError
- Occurrences: 2
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:91: if (!this.isRetryableError(error) || !isIdempotent) {
  - resilientExecutor.gs:175: isRetryableError(error) {

## calculateBackoffDelay
- Occurrences: 2
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:100: const delay = this.calculateBackoffDelay(attempt);
  - resilientExecutor.gs:158: calculateBackoffDelay(attempt) {

## handleFailure
- Occurrences: 2
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:110: this.handleFailure(operationName, lastError);
  - resilientExecutor.gs:211: handleFailure(operationName, error) {

## getProjectIdFromEnvironment
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:14: projectId: options.projectId || this.getProjectIdFromEnvironment(),
  - secretManager.gs:503: getProjectIdFromEnvironment() {

## isSecretCached
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:71: if (useCache && this.isSecretCached(secretName)) {
  - secretManager.gs:515: isSecretCached(secretName) {

## getSecretFromCache
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:72: const cached = this.getSecretFromCache(secretName);
  - secretManager.gs:519: getSecretFromCache(secretName) {

## isCacheExpired
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:73: if (cached && !this.isCacheExpired(secretName)) {
  - secretManager.gs:524: isCacheExpired(secretName) {

## CACHE_HIT
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:76: this.logSecretAccess('CACHE_HIT', secretName, logMeta);
  - secretManager.gs:568: // エラー時のみログ出力（通常の GET/CACHE_HIT は記録しない）

## getSecretFromManager
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:87: secretValue = this.getSecretFromManager(secretName, version);
  - secretManager.gs:219: getSecretFromManager(secretName, version = 'latest') {

## getSecretFromProperties
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:112: secretValue = this.getSecretFromProperties(secretName);
  - secretManager.gs:343: getSecretFromProperties(secretName) {

## setSecret
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:142: setSecret(secretName, secretValue, options = {}) {
  - secretManager.gs:731: return unifiedSecretManager.setSecret(key, value, options);

## validateCriticalSecret
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:161: if (!this.validateCriticalSecret(secretName, secretValue)) {
  - secretManager.gs:474: validateCriticalSecret(secretName, secretValue) {

## setSecretInManager
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:183: this.setSecretInManager(secretName, secretValue);
  - secretManager.gs:261: setSecretInManager(secretName, secretValue) {

## setSecretInProperties
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:198: this.setSecretInProperties(secretName, valueToStore, { encrypted: encrypt });
  - secretManager.gs:388: setSecretInProperties(secretName, secretValue, options = {}) {

## newBlob
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:246: ? Utilities.newBlob(Utilities.base64Decode(data.payload.data)).getDataAsString()
  - secretManager.gs:438: const encryptedString = Utilities.newBlob(encrypted).getDataAsString();

## base64Decode
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:246: ? Utilities.newBlob(Utilities.base64Decode(data.payload.data)).getDataAsString()
  - secretManager.gs:437: const encrypted = Utilities.base64Decode(encryptedValue.substring(4));

## getDataAsString
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:246: ? Utilities.newBlob(Utilities.base64Decode(data.payload.data)).getDataAsString()
  - secretManager.gs:438: const encryptedString = Utilities.newBlob(encrypted).getDataAsString();

## base64Encode
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:318: data: Utilities.base64Encode(secretValue),
  - secretManager.gs:420: return 'ENC:' + Utilities.base64Encode(encrypted);

## isEncryptedValue
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:360: if (this.isEncryptedValue(value)) {
  - secretManager.gs:458: isEncryptedValue(value) {

## ENCRYPTION_KEY_2024
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:412: const key = 'ENCRYPTION_KEY_2024';
  - secretManager.gs:439: const key = 'ENCRYPTION_KEY_2024';

## encryptedValue
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:433: if (!encryptedValue.startsWith('ENC:')) {
  - secretManager.gs:437: const encrypted = Utilities.base64Decode(encryptedValue.substring(4));

## metadata
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:554: timestamp: metadata.timestamp || new Date().toISOString(),
  - secretManager.gs:557: userEmail: metadata.userEmail,

## DISABLED
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:615: results.secretManagerStatus = 'DISABLED';
  - secretManager.gs:646: results.encryptionStatus = 'DISABLED';

## LAST_ACCESS_EMAIL
- Occurrences: 2
- Files: session-utils.gs
- Samples:
  - session-utils.gs:208: 'LAST_ACCESS_EMAIL',
  - session-utils.gs:374: 'LAST_ACCESS_EMAIL',

## USER_CACHE_KEY
- Occurrences: 2
- Files: session-utils.gs
- Samples:
  - session-utils.gs:211: 'USER_CACHE_KEY',
  - session-utils.gs:377: 'USER_CACHE_KEY',

## SESSION_ID
- Occurrences: 2
- Files: session-utils.gs
- Samples:
  - session-utils.gs:212: 'SESSION_ID',
  - session-utils.gs:378: 'SESSION_ID',

## isCriticalComponent
- Occurrences: 2
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:73: if (this.isCriticalComponent(componentName)) {
  - systemIntegrationManager.gs:194: isCriticalComponent(componentName) {

## schedulePeriodicTasks
- Occurrences: 2
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:100: this.schedulePeriodicTasks();
  - systemIntegrationManager.gs:263: schedulePeriodicTasks() {

## resetMetrics
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedBatchProcessor.gs
- Samples:
  - systemIntegrationManager.gs:179: unifiedBatchProcessor.resetMetrics();
  - unifiedBatchProcessor.gs:575: resetMetrics() {

## scheduleMetricsUpdate
- Occurrences: 2
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:271: this.scheduleMetricsUpdate();
  - systemIntegrationManager.gs:283: scheduleMetricsUpdate() {

## getProjectTriggers
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:285: const triggers = ScriptApp.getProjectTriggers();
  - unifiedSecurityManager.gs:962: const triggers = ScriptApp.getProjectTriggers();

## getHandlerFunction
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:289: if (trigger.getHandlerFunction() === 'updateSystemMetrics') {
  - unifiedSecurityManager.gs:966: if (trigger.getHandlerFunction() === 'runScheduledSecurityHealthCheck') {

## trigger
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:289: if (trigger.getHandlerFunction() === 'updateSystemMetrics') {
  - unifiedSecurityManager.gs:966: if (trigger.getHandlerFunction() === 'runScheduledSecurityHealthCheck') {

## deleteTrigger
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:290: ScriptApp.deleteTrigger(trigger);
  - unifiedSecurityManager.gs:967: ScriptApp.deleteTrigger(trigger);

## newTrigger
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:295: ScriptApp.newTrigger('updateSystemMetrics').timeBased().everyMinutes(15).create();
  - unifiedSecurityManager.gs:972: ScriptApp.newTrigger('runScheduledSecurityHealthCheck')

## timeBased
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:295: ScriptApp.newTrigger('updateSystemMetrics').timeBased().everyMinutes(15).create();
  - unifiedSecurityManager.gs:973: .timeBased()

## everyMinutes
- Occurrences: 2
- Files: systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - systemIntegrationManager.gs:295: ScriptApp.newTrigger('updateSystemMetrics').timeBased().everyMinutes(15).create();
  - unifiedSecurityManager.gs:974: .everyMinutes(intervalMinutes)

## getSystemStatus
- Occurrences: 2
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:377: getSystemStatus() {
  - systemIntegrationManager.gs:453: systemStatus: systemIntegrationManager.getSystemStatus(),

## INTEGRATION
- Occurrences: 2
- Files: ulog.gs
- Samples:
  - ulog.gs:214: INTEGRATION: 'INTEGRATION'
  - ulog.gs:214: INTEGRATION: 'INTEGRATION'

## generateBatchCacheKey
- Occurrences: 2
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:52: const cacheKey = this.generateBatchCacheKey('batchGet', spreadsheetId, ranges, options);
  - unifiedBatchProcessor.gs:415: generateBatchCacheKey(operation, spreadsheetId, params, options = {}) {

## fetchError
- Occurrences: 2
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:282: error: fetchError.message,
  - unifiedBatchProcessor.gs:286: throw new Error('API request failed: ' + fetchError.message);

## MT_
- Occurrences: 2
- Files: unifiedBatchProcessor.gs, unifiedSecurityManager.gs
- Samples:
  - unifiedBatchProcessor.gs:444: `MT_*_${spreadsheetId}_*`,
  - unifiedSecurityManager.gs:359: let secureKey = `MT_${userHash}_${baseKey}`;

## unifiedExecutionCache
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:3: * cache.gs、unifiedExecutionCache.gs、spreadsheetCache.gs の機能を統合
  - unifiedCacheManager.gs:1114: // SECTION 2: 統一実行レベルキャッシュ管理システム（元unifiedExecutionCache.gs）

## validateKey
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:46: if (!this.validateKey(key)) {
  - unifiedCacheManager.gs:110: validateKey(key) {

## valueFn
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:48: return valueFn ? valueFn() : null;
  - unifiedCacheManager.gs:70: newValue = valueFn();

## getFromCacheHierarchy
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:52: const cacheResult = this.getFromCacheHierarchy(
  - unifiedCacheManager.gs:130: getFromCacheHierarchy(key, enableMemoization, ttl, usePropertiesFallback = false) {

## parseScriptCacheValue
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:152: const parsedValue = this.parseScriptCacheValue(key, cachedValue);
  - unifiedCacheManager.gs:195: parseScriptCacheValue(key, cachedValue) {

## handleCorruptedCacheEntry
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:164: this.handleCorruptedCacheEntry(key);
  - unifiedCacheManager.gs:214: handleCorruptedCacheEntry(key) {

## parsePropertiesValue
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:173: const parsedValue = this.parsePropertiesValue(propsValue);
  - unifiedCacheManager.gs:787: parsePropertiesValue(value) {

## setToCacheHierarchy
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:176: this.setToCacheHierarchy(key, parsedValue, Math.max(ttl, 3600), enableMemoization);
  - unifiedCacheManager.gs:806: setToCacheHierarchy(key, value, ttl, enableMemoization) {

## WEB_APP_URL
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:197: const rawStringKeys = ['WEB_APP_URL'];
  - unifiedCacheManager.gs:475: 'WEB_APP_URL', // WebアプリURL

## valuesFn
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:266: const newValues = Promise.resolve(valuesFn(missingKeys));
  - unifiedCacheManager.gs:306: const newValues = valuesFn(missingKeys);

## _isProtectedKey
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:421: if (this._isProtectedKey(key)) {
  - unifiedCacheManager.gs:472: _isProtectedKey(key) {

## SA_TOKEN_CACHE
- Occurrences: 2
- Files: unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - unifiedCacheManager.gs:474: 'SA_TOKEN_CACHE', // サービスアカウントトークン
  - unifiedSecurityManager.gs:12: const AUTH_CACHE_KEY = 'SA_TOKEN_CACHE';

## invalidateRelated
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:490: invalidateRelated(entityType, entityId, relatedIds = [], options = {}) {
  - unifiedCacheManager.gs:2142: this.manager.invalidateRelated('spreadsheet', dbSpreadsheetId, [userId]);

## _invalidateCrossEntityCache
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:557: this._invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog);
  - unifiedCacheManager.gs:629: _invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog) {

## relatedId
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:634: if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {
  - unifiedCacheManager.gs:634: if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {

## dom
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:926: typeof window.sharedUtilities.dom.clearElementCache === 'function'
  - unifiedCacheManager.gs:928: window.sharedUtilities.dom.clearElementCache();

## throttle
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:941: typeof window.sharedUtilities.throttle.clearAll === 'function'
  - unifiedCacheManager.gs:943: window.sharedUtilities.throttle.clearAll();

## setSheetsService
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1182: setSheetsService(service) {
  - unifiedCacheManager.gs:2081: this.executionCache.setSheetsService(service);

## Spreadsheet
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1283: * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト
  - unifiedCacheManager.gs:1495: * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト

## optimizationOpportunities
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1952: analysis.optimizationOpportunities.push('プリウォーミング戦略の導入');
  - unifiedCacheManager.gs:1956: analysis.optimizationOpportunities.push('TTL延長によるさらなる高速化');

## auth
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:3: * auth.gs、multiTenantSecurity.gs、securityHealthCheck.gs の機能を統合
  - unifiedSecurityManager.gs:8: // SECTION 1: 認証管理システム（元auth.gs）

## securityHealthCheck
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:3: * auth.gs、multiTenantSecurity.gs、securityHealthCheck.gs の機能を統合
  - unifiedSecurityManager.gs:619: // SECTION 3: セキュリティヘルスチェック統合システム（元securityHealthCheck.gs）

## checkAdminAccess
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:304: const hasAdminAccess = this.checkAdminAccess(requestUserId, targetUserId, operation);
  - unifiedSecurityManager.gs:400: checkAdminAccess(requestUserId, targetUserId, operation) {

## TENANT_BOUNDARY_VIOLATION
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:311: this.logSecurityViolation('TENANT_BOUNDARY_VIOLATION', {
  - unifiedSecurityManager.gs:425: if (violationType === 'TENANT_BOUNDARY_VIOLATION') {

## validateDataAccessPattern
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:327: validateDataAccessPattern(userId, dataType, operation) {
  - unifiedSecurityManager.gs:497: if (!multiTenantSecurity.validateDataAccessPattern(userId, 'user_cache', operation)) {

## generateUserHash
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:356: const userHash = this.generateUserHash(userId);
  - unifiedSecurityManager.gs:375: generateUserHash(userId) {

## handleCriticalSecurityViolation
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:426: this.handleCriticalSecurityViolation(logEntry);
  - unifiedSecurityManager.gs:452: handleCriticalSecurityViolation(logEntry) {

## getEditors
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:1027: const permissions = spreadsheet.getEditors();
  - unifiedSecurityManager.gs:1066: const permissions = spreadsheet.getEditors();

## editor
- Occurrences: 2
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:1028: const isAlreadyEditor = permissions.some((editor) => editor.getEmail() === serviceAccountEmail);
  - unifiedSecurityManager.gs:1067: const isAlreadyEditor = permissions.some((editor) => editor.getEmail() === serviceAccountEmail);

## _validateCoreParams
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:66: this._validateCoreParams(userId, spreadsheetId, sheetName);
  - unifiedSheetDataManager.gs:709: _validateCoreParams(userId, spreadsheetId, sheetName) {

## _generateCacheKey
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:69: const cacheKey = this._generateCacheKey({
  - unifiedSheetDataManager.gs:731: _generateCacheKey(params) {

## _getCacheConfig
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:81: const cacheConfig = this._getCacheConfig(dataType, adminMode);
  - unifiedSheetDataManager.gs:751: _getCacheConfig(dataType, adminMode) {

## _fetchPublishedData
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:121: return this._fetchPublishedData(userInfo, options);
  - unifiedSheetDataManager.gs:138: _fetchPublishedData(userInfo, options) {

## _fetchIncrementalData
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:123: return this._fetchIncrementalData(userInfo, options);
  - unifiedSheetDataManager.gs:182: _fetchIncrementalData(userInfo, options) {

## _fetchFullSheetData
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:126: return this._fetchFullSheetData(userInfo, options);
  - unifiedSheetDataManager.gs:259: _fetchFullSheetData(userInfo, options) {

## _getRawSheetData
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:170: const rawData = this._getRawSheetData(publishedSpreadsheetId, publishedSheetName, sheetConfig);
  - unifiedSheetDataManager.gs:765: _getRawSheetData(spreadsheetId, sheetName, sheetConfig) {

## _handleHeaderValidationFailure
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:353: return this._handleHeaderValidationFailure(spreadsheetId, sheetName, headers, validation);
  - unifiedSheetDataManager.gs:487: _handleHeaderValidationFailure(spreadsheetId, sheetName, headers, validation) {

## _fetchSpreadsheetInfo
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:559: () => this._fetchSpreadsheetInfo(spreadsheetId),
  - unifiedSheetDataManager.gs:591: _fetchSpreadsheetInfo(spreadsheetId) {

## invalidate
- Occurrences: 2
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:62: invalidate(userId, email) {
  - unifiedUserManager.gs:148: unifiedUserCache.invalidate(userId, email);

## identifier
- Occurrences: 2
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:452: if (type === 'email' || (!type && typeof identifier === 'string' && identifier.includes('@'))) {
  - unifiedUserManager.gs:452: if (type === 'email' || (!type && typeof identifier === 'string' && identifier.includes('@'))) {

## commonUtilities
- Occurrences: 2
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:4: * commonUtilities.gsの機能も統合
  - unifiedUtilities.gs:8: // 統合：共通ユーティリティ関数（元commonUtilities.gs）

## DATA
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:14: DATA: 'data_integrity',
  - unifiedValidationSystem.gs:49: case this.validationCategories.DATA:

## _validateAuthentication
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:44: result = this._validateAuthentication(level, options);
  - unifiedValidationSystem.gs:82: _validateAuthentication(level, options) {

## _validateConfiguration
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:47: result = this._validateConfiguration(level, options);
  - unifiedValidationSystem.gs:112: _validateConfiguration(level, options) {

## _validateDataIntegrity
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:50: result = this._validateDataIntegrity(level, options);
  - unifiedValidationSystem.gs:142: _validateDataIntegrity(level, options) {

## _validateSystemHealth
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:53: result = this._validateSystemHealth(level, options);
  - unifiedValidationSystem.gs:174: _validateSystemHealth(level, options) {

## _validateWorkflow
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:56: result = this._validateWorkflow(level, options);
  - unifiedValidationSystem.gs:207: _validateWorkflow(level, options) {

## _formatValidationResult
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:62: const finalResult = this._formatValidationResult(
  - unifiedValidationSystem.gs:859: _formatValidationResult(validationId, category, level, result, startTime, options) {

## _handleValidationError
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:75: return this._handleValidationError(validationId, category, level, error, startTime, options);
  - unifiedValidationSystem.gs:890: _handleValidationError(validationId, category, level, error, startTime, options) {

## _checkUserAuthentication
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:90: results.tests.userAuth = this._checkUserAuthentication(options.userId);
  - unifiedValidationSystem.gs:241: _checkUserAuthentication(userId) {

## _checkSessionValidity
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:91: results.tests.sessionValid = this._checkSessionValidity(options.userId);
  - unifiedValidationSystem.gs:262: _checkSessionValidity(userId) {

## _checkUserAccess
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:96: results.tests.userAccess = this._checkUserAccess(options.userId);
  - unifiedValidationSystem.gs:278: _checkUserAccess(userId) {

## _checkAdminPermissions
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:97: results.tests.adminCheck = this._checkAdminPermissions(options.userId);
  - unifiedValidationSystem.gs:294: _checkAdminPermissions(userId) {

## _checkAuthenticationHealth
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:102: results.tests.authHealth = this._checkAuthenticationHealth();
  - unifiedValidationSystem.gs:311: _checkAuthenticationHealth() {

## _performAuthSecurityCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:103: results.tests.securityCheck = this._performAuthSecurityCheck(options.userId);
  - unifiedValidationSystem.gs:328: _performAuthSecurityCheck(userId) {

## _checkConfigJson
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:120: results.tests.configJson = this._checkConfigJson(options.config);
  - unifiedValidationSystem.gs:345: _checkConfigJson(config) {

## _checkConfigState
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:121: results.tests.configState = this._checkConfigState(options.config, options.userInfo);
  - unifiedValidationSystem.gs:362: _checkConfigState(config, userInfo) {

## _checkSystemConfiguration
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:126: results.tests.systemConfig = this._checkSystemConfiguration();
  - unifiedValidationSystem.gs:378: _checkSystemConfiguration() {

## _checkSystemDependencies
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:127: results.tests.dependencies = this._checkSystemDependencies();
  - unifiedValidationSystem.gs:395: _checkSystemDependencies() {

## _checkServiceAccountConfiguration
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:132: results.tests.serviceAccount = this._checkServiceAccountConfiguration();
  - unifiedValidationSystem.gs:412: _checkServiceAccountConfiguration() {

## _checkApplicationAccess
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:133: results.tests.applicationAccess = this._checkApplicationAccess();
  - unifiedValidationSystem.gs:429: _checkApplicationAccess() {

## _checkHeaderIntegrity
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:150: results.tests.headerIntegrity = this._checkHeaderIntegrity(options.userId);
  - unifiedValidationSystem.gs:445: _checkHeaderIntegrity(userId) {

## _checkRequiredHeaders
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:151: results.tests.requiredHeaders = this._checkRequiredHeaders(options.indices);
  - unifiedValidationSystem.gs:462: _checkRequiredHeaders(indices) {

## _checkForDuplicates
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:156: results.tests.duplicates = this._checkForDuplicates(options.headers, options.userRows);
  - unifiedValidationSystem.gs:478: _checkForDuplicates(headers, userRows) {

## _checkMissingRequiredFields
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:157: results.tests.missingFields = this._checkMissingRequiredFields(options.headers, options.userRows);
  - unifiedValidationSystem.gs:495: _checkMissingRequiredFields(headers, userRows) {

## _checkInvalidDataFormats
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:158: results.tests.invalidFormats = this._checkInvalidDataFormats(options.headers, options.userRows);
  - unifiedValidationSystem.gs:512: _checkInvalidDataFormats(headers, userRows) {

## _checkOrphanedData
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:163: results.tests.orphanedData = this._checkOrphanedData(options.headers, options.userRows);
  - unifiedValidationSystem.gs:529: _checkOrphanedData(headers, userRows) {

## _performDataIntegrityCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:164: results.tests.dataIntegrity = this._performDataIntegrityCheck(options);
  - unifiedValidationSystem.gs:546: _performDataIntegrityCheck(options) {

## _checkUserScopedKey
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:165: results.tests.userScopedKey = this._checkUserScopedKey(options.key, options.expectedUserId);
  - unifiedValidationSystem.gs:563: _checkUserScopedKey(key, expectedUserId) {

## _performBasicHealthCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:182: results.tests.basicHealth = this._performBasicHealthCheck();
  - unifiedValidationSystem.gs:579: _performBasicHealthCheck() {

## _checkCacheStatus
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:183: results.tests.cacheStatus = this._checkCacheStatus(options.userId);
  - unifiedValidationSystem.gs:596: _checkCacheStatus(userId) {

## _checkDatabaseHealth
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:188: results.tests.databaseHealth = this._checkDatabaseHealth();
  - unifiedValidationSystem.gs:613: _checkDatabaseHealth() {

## _performPerformanceCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:189: results.tests.performanceCheck = this._performPerformanceCheck();
  - unifiedValidationSystem.gs:630: _performPerformanceCheck() {

## _performSecurityCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:194: results.tests.securityHealth = this._performSecurityCheck();
  - unifiedValidationSystem.gs:647: _performSecurityCheck() {

## _performMultiTenantHealthCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:195: results.tests.multiTenantHealth = this._performMultiTenantHealthCheck();
  - unifiedValidationSystem.gs:664: _performMultiTenantHealthCheck() {

## _performServiceAccountHealthCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:196: results.tests.serviceAccountHealth = this._performServiceAccountHealthCheck();
  - unifiedValidationSystem.gs:681: _performServiceAccountHealthCheck() {

## _checkUnifiedUserManagerHealth
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:197: results.tests.userManagerHealth = this._checkUnifiedUserManagerHealth();
  - unifiedValidationSystem.gs:698: _checkUnifiedUserManagerHealth() {

## _performUnifiedBatchHealthCheck
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:198: results.tests.batchProcessorHealth = this._performUnifiedBatchHealthCheck();
  - unifiedValidationSystem.gs:715: _performUnifiedBatchHealthCheck() {

## _validateWorkflowGAS
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:215: results.tests.workflowGAS = this._validateWorkflowGAS(options.data);
  - unifiedValidationSystem.gs:732: _validateWorkflowGAS(data) {

## _checkCurrentPublicationStatus
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:216: results.tests.publicationStatus = this._checkCurrentPublicationStatus(options.userId);
  - unifiedValidationSystem.gs:749: _checkCurrentPublicationStatus(userId) {

## _validateWorkflowWithSheet
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:221: results.tests.workflowWithSheet = this._validateWorkflowWithSheet();
  - unifiedValidationSystem.gs:766: _validateWorkflowWithSheet() {

## _checkAndHandleAutoStop
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:222: results.tests.autoStop = this._checkAndHandleAutoStop(options.config, options.userInfo);
  - unifiedValidationSystem.gs:783: _checkAndHandleAutoStop(config, userInfo) {

## _checkIfNewOrUpdatedForm
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:223: results.tests.formUpdates = this._checkIfNewOrUpdatedForm(options.requestUserId, options.spreadsheet
  - unifiedValidationSystem.gs:800: _checkIfNewOrUpdatedForm(requestUserId, spreadsheetId, sheetName) {

## _comprehensiveWorkflowValidation
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:228: results.tests.comprehensiveWorkflow = this._comprehensiveWorkflowValidation();
  - unifiedValidationSystem.gs:817: _comprehensiveWorkflowValidation() {

## round
- Occurrences: 2
- Files: unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - unifiedValidationSystem.gs:875: passRate: result.summary.total > 0 ? Math.round((result.summary.passed / result.summary.total) * 100
  - validationMigration.gs:493: migrationProgress: Math.round(((stats.total - stats.deprecated) / stats.total) * 100)

## list
- Occurrences: 2
- Files: url.gs
- Samples:
  - url.gs:55: deploymentsList = AppsScript.Script.Deployments.list(scriptId, {
  - url.gs:61: deploymentsList = AppsScript.Script.Deployments.list();

## webApp
- Occurrences: 2
- Files: url.gs
- Samples:
  - url.gs:56: fields: 'deployments(deploymentId,deploymentConfig(webApp(url)),updateTime)',
  - url.gs:81: finalUrl = webAppDeployments[0].deploymentConfig.webApp.url;

## busting
- Occurrences: 2
- Files: url.gs
- Samples:
  - url.gs:217: * Generate URL for unpublished state with aggressive cache busting.
  - url.gs:253: * Generate URLs with optional cache busting.

## LOGGING
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:93: ULog.error(`⚠️ LOGGING ERROR [database.${operation}]:`, {

## loggingError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:95: loggingError: loggingError.message,

## activeSheetName
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:295: configJson.activeSheetName.trim() !== ''; // アクティブシートが選択済み

## completed
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:602: changes.push(`setupStatus: ${configJson.setupStatus} → completed (form作成済み)`);

## EABDB
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:813: const sheetName = userInfo.sheetName || 'EABDB';

## one
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:913: * Register a new user or update an existing one.

## flags
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:916: * @returns {Object} Registration result including URLs and flags.

## methodError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1060: ULog.debug('registerNewUser: 検証方法でエラー:', methodError.message);

## lookupError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1124: ULog.warn('registerNewUser: 既存ユーザー検索でエラー:', lookupError.message);

## operationError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1348: operationError.message,

## operations
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1396: { batchSize: operations.length }

## diagnosisError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1625: ULog.warn('診断処理でエラー:', diagnosisError.message);

## FORM_NOT_FOUND
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2655: return createErrorResponse('FORM_NOT_FOUND', 'フォームが見つかりません', null);

## core
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3423: 'getHeaderIndices received in core.gs: spreadsheetId=%s, sheetName=%s',

## getSheetById
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3454: var sheet = spreadsheet.getSheetById(sheetId);

## NOTE
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3678: // NOTE: unpublishBoard関数はconfig.gsに実装済み

## VERIFIED
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3724: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);

## setEmailCollectionType
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3724: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);

## EmailCollectionType
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3724: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);

## emailError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3729: ULog.warn('Email collection setting failed:', emailError.message);

## createParagraphTextValidation
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3782: var validation = FormApp.createParagraphTextValidation()

## requireTextLengthLessThanOrEqualTo
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3783: .requireTextLengthLessThanOrEqualTo(140)

## build
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3784: .build();

## setValidation
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3785: reasonItem.setValidation(validation);

## addCheckboxItem
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3820: mainItem = form.addCheckboxItem();

## addMultipleChoiceItem
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3837: mainItem = form.addMultipleChoiceItem();

## DOMAIN
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4135: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

## VIEW
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4135: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

## sharingError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4141: ULog.warn('共有設定の変更に失敗しましたが、処理を続行します: ' + sharingError.message);

## rename
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4148: spreadsheetObj.rename(spreadsheetName);

## sheetNameError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4176: ULog.error('[ERROR]', '❌ フォーム連携後のシート名取得エラー:', sheetNameError.message);

## DOMAIN_WITH_LINK
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4334: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## EDIT
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4334: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## domainSharingError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4337: ULog.warn('ドメイン共有設定に失敗: ' + domainSharingError.message);

## individualError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4344: ULog.error('[ERROR]', '個別ユーザー追加も失敗: ' + individualError.message);

## spreadsheetAddError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4354: ULog.warn('SpreadsheetApp経由の追加で警告: ' + spreadsheetAddError.message);

## E3F2FD
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4407: allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

## setFontWeight
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4407: allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

## setBackground
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4407: allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

## autoResizeColumns
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4411: sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());

## getNumColumns
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4411: sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());

## resizeError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4413: ULog.warn('Auto-resize failed:', resizeError.message);

## finalRepairError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4697: ULog.error('[ERROR]', 'getSheetsList: 最終修復も失敗:', finalRepairError.message);

## property
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4729: 'getSheetsList: Spreadsheet data missing sheets property. Available properties:',

## reverse
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4893: return data.reverse();

## actualHeaderIndices
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4980: if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {

## prev
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:5063: return prev.header.length > current.header.length ? prev : current;

## current
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:5063: return prev.header.length > current.header.length ? prev : current;

## Code
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:5224: // 追加のコアファンクション（Code.gsから移行）

## USER_NOT_AUTHENTICATED
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:5254: return createErrorResponse('USER_NOT_AUTHENTICATED', 'ユーザーが認証されていません');

## USER_ID_REQUIRED
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:5452: return createErrorResponse('ユーザーIDが指定されていません', 'USER_ID_REQUIRED');

## setName
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:5634: sheet.setName(config.sheetName);

## LOGIN_USER_NOT_FOUND
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6070: error: 'LOGIN_USER_NOT_FOUND',

## finalError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6159: searchAttempts.push({ method: `final_method_${i + 1}`, error: finalError.message });

## USER_INFO_FETCH_FAILED
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6261: error: 'USER_INFO_FETCH_FAILED',

## OPTIMIZED
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6453: * 統合初期データ取得API - OPTIMIZED

## consistencyError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6501: ULog.warn('⚠️ データ整合性チェック中にエラー:', consistencyError.message);

## stepError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6540: ULog.warn('setupStep決定でエラー、デフォルト値(1)を使用:', stepError.message);

## includedApis
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6680: response._meta.includedApis.push('getSheetDetails');

## getDateCreated
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6921: const createdDate = file.getDateCreated();

## update
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:6971: * @param {Object} updateData - Data to update (e.g., {spreadsheetId: "..."})

## ALREADY_INITIALIZED
- Occurrences: 1
- Files: autoInit.gs
- Samples:
  - autoInit.gs:17: skipReason: 'ALREADY_INITIALIZED',

## addResource
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:108: addResource(key, resource) {

## removeResource
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:118: removeResource(key) {

## cleanupError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:158: ULog.warn(`リソースクリーンアップエラー (${key}):`, cleanupError.message);

## getMemoryStats
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:202: getMemoryStats() {

## register
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:218: register(context) {

## propertyError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:416: ULog.warn('プロパティクリーンアップエラー:', propertyError.message);

## scriptCacheError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:434: ULog.debug('スクリプトキャッシュクリーンアップスキップ:', scriptCacheError.message);

## formSyncError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1226: ULog.warn('activateSheet: フォームURL同期で警告:', formSyncError.message);

## formDetectionError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1373: ULog.warn('フォーム自動検出でエラー:', formDetectionError.message);

## cfgErr
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1469: ULog.warn('addFormUrl: シート固有の更新に失敗（グローバルのみ反映）:', cfgErr.message);

## reverseError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1523: console.debug('⚠️ 逆検索エラー:', reverseError.message);

## createError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1535: ULog.warn('⚠️ フォーム自動作成エラー:', createError.message);

## Z10
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1543: const range = sheet.getRange('A1:Z10');

## cellError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1562: ULog.warn('⚠️ セル内容検索エラー:', sheet.getName(), cellError.message);

## searchFiles
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1587: const formFiles = DriveApp.searchFiles('mimeType="application/vnd.google-apps.form"');

## vnd
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1587: const formFiles = DriveApp.searchFiles('mimeType="application/vnd.google-apps.form"');

## apps
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1587: const formFiles = DriveApp.searchFiles('mimeType="application/vnd.google-apps.form"');

## getDestinationType
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1600: const destinationType = form.getDestinationType();

## getDestinationId
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1602: const destinationId = form.getDestinationId();

## formCheckError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1611: console.debug('⚠️ フォームチェックエラー:', formFile.getName(), formCheckError.message);

## destError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1706: ULog.warn('⚠️ 回答先設定でエラー:', destError.message);

## cacheServiceError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:1856: ULog.debug('CacheService clear 警告:', cacheServiceError.message);

## PHASE3
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2297: // PHASE3 OPTIMIZATION: トランザクション型実行アーキテクチャ

## OPTIMIZATION
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2297: // PHASE3 OPTIMIZATION: トランザクション型実行アーキテクチャ

## keyError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2520: ULog.warn('キャッシュキー削除エラー:', key, keyError.message);

## patternErr
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2532: patternErr.message

## patternError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2542: ULog.warn('パターンベースキャッシュ削除エラー:', patternError.message);

## recreating
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2653: ULog.warn('⚠️ SheetsService lost during serialization, recreating...');

## mismatch
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2884: '❌ This suggests parameter order mismatch. Expected: (context:object, spreadsheetId:string, sheetNam

## parameters
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:2892: 'Parameter order mismatch: context must be an object, not a string. Check function call parameters.'

## syncErr
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:3252: ULog.warn('saveSheetConfigInContext: トップレベル同期で警告（処理継続）:', syncErr.message);

## getFormErr
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:3321: ULog.debug('getFormUrl 未取得（例外）:', getFormErr.message);

## sheetOpenErr
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:3325: ULog.debug('シートアクセス失敗（フォームURL検出スキップ）:', sheetOpenErr.message);

## formSyncErr
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:3352: ULog.warn('switchToSheetInContext: フォームURL同期で警告:', formSyncErr.message);

## warmingError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:3537: warmingError.message

## conversionError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:3702: const error = '配列形式の設定データ変換に失敗しました: ' + conversionError.message;

## SMALL
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:71: SMALL: 20,

## XLARGE
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:74: XLARGE: 200,

## OPINION_SURVEY
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:171: OPINION_SURVEY: {

## REFLECTION
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:177: REFLECTION: {

## FEEDBACK
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:183: FEEDBACK: {

## STRING_SHORT
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:194: STRING_SHORT: 50,

## STRING_MEDIUM
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:195: STRING_MEDIUM: 100,

## STRING_LONG
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:196: STRING_LONG: 500,

## STRING_MAX
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:197: STRING_MAX: 1000,

## MAX_CONCURRENT_FLOWS
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:200: MAX_CONCURRENT_FLOWS: 5,

## LIKE_MULTIPLIER
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:205: LIKE_MULTIPLIER: 0.1,

## RANDOM_FACTOR
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:206: RANDOM_FACTOR: 0.01,

## LIKE_WEIGHT
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:211: LIKE_WEIGHT: 1.0,

## LENGTH_WEIGHT
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:212: LENGTH_WEIGHT: 0.1,

## RECENCY_WEIGHT
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:213: RECENCY_WEIGHT: 0.05,

## MULTIPLE_REACTIONS
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:218: MULTIPLE_REACTIONS: 0.2,

## DETAILED_REASON
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:219: DETAILED_REASON: 0.15,

## LOG_LEVEL
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:234: LOG_LEVEL: 'INFO',

## TRACE_ENABLED
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:235: TRACE_ENABLED: false,

## path
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:382: const parts = path.split('.');

## ulog
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:622: // ULogクラスはulog.gsで定義されている

## validateComplex
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:704: validateComplex(data, rules) {

## executeWithRetry
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:802: executeWithRetry(operation, options = {}) {

## safeFilter
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:875: safeFilter(array, predicate, functionName = 'safeFilter') {

## safeMap
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:901: safeMap(array, mapper, functionName = 'safeMap') {

## processBatch
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:927: processBatch(items, processor, options = {}) {

## processor
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:949: return processor(item);

## called
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1024: console.log('[GAS-Polyfill] clearTimeout called (no-op in GAS):', timeoutId);

## pair
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1057: const [key, value] = pair.split('=');

## append
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1070: append(key, value) {

## SETUP
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1101: const setupStep = (step) => console.log(`[SETUP] ${step}`);

## CLEANUP
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1103: const cleanupExpired = () => console.log('[CLEANUP] Expired items cleaned');

## DMS
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1104: const dms = () => console.log('[DMS] Document management operation');

## DELETE
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1111: const deleteAll = () => console.log('[DELETE] All items deleted');

## COMPONENT
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1115: const initializeComponent = (component) => console.log(`[COMPONENT] Initialized ${component}`);

## updates
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1123: const batchUpdate = (updates) => console.log(`[BATCH] Updating ${updates.length} items`);

## LOGIN
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1185: const updateLoginStatus = (userId, status) => console.log(`[LOGIN] Status ${status} for ${userId}`);

## UNIFIED
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1190: const unified = () => console.log('[UNIFIED] Operation');

## SAFE_EXECUTE
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1197: catch(e) { console.error('[SAFE_EXECUTE]', e.message); return null; }

## unifiedSheetDataManager
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1241: // UnifiedSheetDataManager は unifiedSheetDataManager.gs で定義済み

## secretManager
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1249: // UnifiedSecretManager は secretManager.gs で定義済み

## YES_NO
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1270: const response = Browser.msgBox(message, Browser.Buttons.YES_NO);

## YES
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1271: return response === Browser.Buttons.YES;

## RELOAD
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1283: const reload = () => console.log('[RELOAD] Page reload not applicable in GAS');

## PREVENT_DEFAULT
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1284: const preventDefault = () => console.log('[PREVENT_DEFAULT] Event handling not applicable in GAS');

## FORM
- Occurrences: 1
- Files: constants.gs
- Samples:
  - constants.gs:1290: setEmailCollectionType: () => console.log('[FORM] setEmailCollectionType'),

## appendError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:171: throw new Error(`ログエントリの追加に失敗: ${appendError.message}`);

## Z1
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:221: const logSheetData = batchGetSheetsData(service, dbId, [`'${logSheetName}'!A1:Z1`]);

## setDate
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:390: cutoffDate.setDate(cutoffDate.getDate() - 30); // 30日前

## getDate
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:390: cutoffDate.setDate(cutoffDate.getDate() - 30); // 30日前

## OR
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:635: // 本人削除 OR 管理者削除

## generation
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:729: console.error('[ERROR]', '❌ Access token is null or undefined after generation.');

## UNAUTHENTICATED
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1029: errorMessage.includes('UNAUTHENTICATED') ||

## ACCESS_TOKEN_EXPIRED
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1030: errorMessage.includes('ACCESS_TOKEN_EXPIRED')

## cacheInvalidationError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1144: ULog.warn('createUser - キャッシュ無効化で警告:', cacheInvalidationError.message);

## cacheErr
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1574: ULog.warn('appendSheetsData: キャッシュクリアでエラー:', cacheErr.message);

## invalidateErr
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1579: ULog.warn('appendSheetsData: キャッシュ無効化でエラー:', invalidateErr.message);

## ALL_USERS
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1812: ULog.debug('🔍 データベース診断開始:', targetUserId || 'ALL_USERS');

## statsError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2035: ULog.warn('キャッシュ統計取得エラー:', statsError.message);

## serviceAccount
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2138: if (result.checks.serviceAccount && result.checks.serviceAccount.email) {

## UUID
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2463: // ユーザーID形式チェック（UUIDかどうか）

## permError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2659: error: permError.message,

## B1
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2736: "'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A1:B1",

## alertError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2900: console.error('[ERROR]', '❌ アラート送信エラー:', alertError.message);

## deletionError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:3051: console.error('[ERROR]', '❌ User deletion failed:', deletionError.message);

## ENVIRONMENT
- Occurrences: 1
- Files: debugConfig.gs
- Samples:
  - debugConfig.gs:39: const envSetting = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT');

## logLevels
- Occurrences: 1
- Files: debugConfig.gs
- Samples:
  - debugConfig.gs:77: const levelValue = DEBUG_CONFIG.logLevels[level] || DEBUG_CONFIG.logLevels.DEBUG;

## args
- Occurrences: 1
- Files: debugConfig.gs
- Samples:
  - debugConfig.gs:90: const logData = args.length > 0 ? { additionalData: args } : {};

## getScriptLock
- Occurrences: 1
- Files: lockManager.gs
- Samples:
  - lockManager.gs:26: const lock = LockService.getScriptLock();

## LockService
- Occurrences: 1
- Files: lockManager.gs
- Samples:
  - lockManager.gs:26: const lock = LockService.getScriptLock();

## waitLock
- Occurrences: 1
- Files: lockManager.gs
- Samples:
  - lockManager.gs:29: lock.waitLock(timeout);

## releaseLock
- Occurrences: 1
- Files: lockManager.gs
- Samples:
  - lockManager.gs:47: lock.releaseLock();

## V8
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:3: * V8ランタイム、最新パフォーマンス技術、安定性強化を統合

## createHtmlOutputFromFile
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:13: return HtmlService.createHtmlOutputFromFile(path).getContent();

## str
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:34: const strValue = str.toString();

## HTTPS
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:36: // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理

## unshift
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:385: configJson.historyArray.unshift(serverHistoryItem);

## pop
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:405: configJson.historyArray.pop();

## TRUE
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:589: * Accepts boolean true, 'true', or 'TRUE'.

## EXECUTION_LIMITS
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:624: EXECUTION_LIMITS: {

## MAX_TIME
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:625: MAX_TIME: 330000, // 5.5分（安全マージン）

## API_RATE_LIMIT
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:627: API_RATE_LIMIT: 90, // 100秒間隔での制限

## CACHE_STRATEGY
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:630: CACHE_STRATEGY: {

## L1_TTL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:631: L1_TTL: 300, // Level 1: 5分

## L2_TTL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:632: L2_TTL: 3600, // Level 2: 1時間

## L3_TTL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:633: L3_TTL: 21600, // Level 3: 6時間（最大）

## start
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:645: globalProfiler.start('logging');

## end
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:670: globalProfiler.end('logging');

## PerformanceOptimizer
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:733: // PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、

## preAuthError
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:916: error: preAuthError.message

## executorError
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:970: errors.push(`resilientExecutor エラー: ${executorError.message}`);

## _DEPENDENCY_TEST_KEY
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:978: CacheService.getScriptCache().get('_DEPENDENCY_TEST_KEY');

## utilsError
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:989: errors.push(`Utilities エラー: ${utilsError.message}`);

## request
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1007: // Clear execution-level cache for new request

## setup
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1021: // Check system setup (highest priority)

## route
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1111: * Handle default route (no mode parameter)

## diagError
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1250: ULog.warn('handleAdminMode: システム診断でエラー', { error: diagError.message }, ULog.CATEGORIES.SYSTEM);

## propError
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1314: propertiesDiagnostics = 'error: ' + propError.message;

## HANDLING
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1645: // ERROR HANDLING & CATEGORIZATION

## CATEGORIZATION
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1645: // ERROR HANDLING & CATEGORIZATION

## XSS
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1904: // XSS攻撃を防ぐため、URLをサニタイズ

## setContent
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1906: return HtmlService.createHtmlOutput().setContent(

## NATIVE
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2244: .setSandboxMode(HtmlService.SandboxMode.NATIVE);

## setSandboxMode
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2244: .setSandboxMode(HtmlService.SandboxMode.NATIVE);

## SandboxMode
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2244: .setSandboxMode(HtmlService.SandboxMode.NATIVE);

## templateError
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2291: throw new Error('Unpublished.htmlテンプレートの読み込みに失敗: ' + templateError.message);

## ErrorBoundary
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2374: // フォールバック: ErrorBoundary.htmlを回避して確実にUnpublishedページを表示

## blur
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2505: backdrop-filter: blur(20px);

## bezier
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2584: transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);

## media
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2666: @media (max-width: 640px) {

## withSuccessHandler
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2727: .withSuccessHandler((result) => {

## withFailureHandler
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2733: .withFailureHandler((error) => {

## CONTROL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:2980: // DEBUG_MODE & USER ACCESS CONTROL API

## setProperties
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:3035: props.setProperties({

## CLIENT
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:3393: console.error('[CLIENT ERROR]', JSON.stringify(errorInfo, null, 2));

## msg
- Occurrences: 1
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:191: return retryableMessages.some((msg) => resilientErrorMessage.toLowerCase().includes(msg.toLowerCase(

## API_KEY
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:30: API_KEY: 'api_key',

## ENCRYPTION_KEY
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:33: ENCRYPTION_KEY: 'encryption_key',

## secretName
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:59: if (!secretName || typeof secretName !== 'string') {

## SET
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:168: this.logSecretAccess('SET', secretName, {

## payload
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:246: ? Utilities.newBlob(Utilities.base64Decode(data.payload.data)).getDataAsString()

## decryptError
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:364: warnLog(`暗号化値の復号化に失敗:`, { secretName, error: decryptError.message });

## _ENCRYPTED
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:395: props.setProperty(`${secretName}_ENCRYPTED`, 'true');

## XOR
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:411: // 簡易暗号化（Base64エンコード + 固定キーでのXOR）

## criticalSecrets
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:467: return this.criticalSecrets.hasOwnProperty(secretName);

## GCP_PROJECT_ID
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:506: return PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');

## shift
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:565: this.auditLog.shift();

## TEST_KEY
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:621: props.getProperty('TEST_KEY'); // 存在しないキーでのテスト

## Cache
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:71: // 2) プレフィックス + email で個別削除（Cache.remove）

## sanitizeError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:432: ULog.error('Fallback URL sanitization failed', { error: sanitizeError.message }, ULog.CATEGORIES.AUT

## preview
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:482: ULog.debug('📄 Generated HTML preview (first 200 chars):', redirectScript.substring(0, 200));

## frameError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:504: ULog.warn('XFrameOptionsMode設定失敗:', frameError.message);

## NO
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:516: outputContent ? outputContent.substring(0, 100) : 'NO CONTENT'

## CONTENT
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:516: outputContent ? outputContent.substring(0, 100) : 'NO CONTENT'

## contentError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:519: ULog.warn('⚠️ Cannot access HtmlOutput content:', contentError.message);

## getEffectiveUser
- Occurrences: 1
- Files: setup.gs
- Samples:
  - setup.gs:22: var adminEmail = Session.getEffectiveUser().getEmail();

## DATABASE_ID_NOT_SET
- Occurrences: 1
- Files: setup.gs
- Samples:
  - setup.gs:59: error: 'DATABASE_ID_NOT_SET',

## SERVICE_ACCOUNT_NOT_SET
- Occurrences: 1
- Files: setup.gs
- Samples:
  - setup.gs:68: error: 'SERVICE_ACCOUNT_NOT_SET',

## healthError
- Occurrences: 1
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:88: initResult.warnings.push(`初期ヘルスチェック失敗: ${healthError.message}`);

## SHUTDOWN
- Occurrences: 1
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:357: component.status = 'SHUTDOWN';

## fromEntries
- Occurrences: 1
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:380: components: Object.fromEntries(

## setLevel
- Occurrences: 1
- Files: ulog.gs
- Samples:
  - ulog.gs:157: static setLevel(level) {

## setPrefix
- Occurrences: 1
- Files: ulog.gs
- Samples:
  - ulog.gs:170: static setPrefix(newPrefix) {

## startTimer
- Occurrences: 1
- Files: ulog.gs
- Samples:
  - ulog.gs:179: static startTimer(label) {

## includeGridData
- Occurrences: 1
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:78: includeGridData: includeGridData.toString(),

## USER_ENTERED
- Occurrences: 1
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:145: valueInputOption = 'USER_ENTERED',

## DELETE_ROWS
- Occurrences: 1
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:284: requestType: chunk[0]?.deleteDimension ? 'DELETE_ROWS' : 'OTHER',

## OTHER
- Occurrences: 1
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:284: requestType: chunk[0]?.deleteDimension ? 'DELETE_ROWS' : 'OTHER',

## all
- Occurrences: 1
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:373: const chunkResults = Promise.all(chunkPromises);

## rate
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:94: `[Cache] Performance: ${hitRate}% hit rate (${this.stats.hits}/${this.stats.totalOps}), ${this.stats

## removeError
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:220: ULog.warn(`[Cache] Failed to clean corrupted entry: ${key}`, removeError.message);

## getAll
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:295: const cachedValues = this.scriptCache.getAll(keys);

## putAll
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:312: this.scriptCache.putAll(newCacheValues, ttl);

## override
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:408: `[Cache] Pattern too broad for safe removal: ${pattern}. Use strict=true to override.`

## limit
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:431: ULog.warn(`[Cache] Reached maxKeys limit (${maxKeys}) for pattern: ${pattern}`);

## protected
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:462: `[Cache] Pattern clear completed: ${successCount} removed, ${failedRemovals} failed, ${skippedCount}

## SYSTEM_CONFIG
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:476: 'SYSTEM_CONFIG', // システム設定

## DOMAIN_INFO
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:477: 'DOMAIN_INFO', // ドメイン情報

## IDs
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:509: ULog.warn(`[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`);

## LIVE
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:514: `🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`

## automatically
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:719: ULog.debug('[Cache] Script cache will expire automatically.');

## waiting
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:858: ULog.debug('🔄 Cache clear already in progress, waiting...');

## getUserInfo
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1135: getUserInfo(userId) {

## headerName
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1664: if (headerName && headerName.trim() !== '' && headerName !== 'タイムスタンプ') {

## retry
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1688: ULog.debug(`[getHeadersWithRetry] Waiting ${waitTime}ms before retry...`);

## clearPattern
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:2136: } else if (typeof clearPattern === 'string' && clearPattern !== 'false') {

## private_key
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:34: const privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化

## RS256
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:42: const jwtHeader = { alg: 'RS256', typ: 'JWT' };

## computeRsaSha256Signature
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:54: const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);

## MISSING_USER_ID
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:293: this.logSecurityViolation('MISSING_USER_ID', { requestUserId, targetUserId, operation });

## TENANT_BOUNDARY_VALID
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:299: this.logDataAccess('TENANT_BOUNDARY_VALID', { requestUserId, operation });

## ADMIN_ACCESS_GRANTED
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:306: this.logDataAccess('ADMIN_ACCESS_GRANTED', { requestUserId, targetUserId, operation });

## INVALID_ACCESS_PATTERN
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:336: this.logSecurityViolation('INVALID_ACCESS_PATTERN', { userId, dataType, operation });

## _T
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:365: secureKey += `_T${timestamp}`;

## _TENANT_SALT_2024
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:381: userId + '_TENANT_SALT_2024',

## SECURITY_VIOLATION
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:416: type: 'SECURITY_VIOLATION',

## DATA_ACCESS
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:440: type: 'DATA_ACCESS',

## CACHE_OPERATION_FAILED
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:531: multiTenantSecurity.logSecurityViolation('CACHE_OPERATION_FAILED', {

## L117
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:1152: // この重複関数は削除されました。L117のgetServiceAccountEmail()を使用してください。

## OBJECT
- Occurrences: 1
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:644: sheet.sheetType !== 'OBJECT' && // オブジェクトシート除外

## USER_
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:247: userId = 'USER_' + Date.now() + '_' + Math.random().toString(36).substring(2, 8);

## 10
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:443: * 10. getOrFetchUserInfo - 汎用ユーザー情報取得

## 11
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:470: * 11. getUserWithFallback - フォールバック付きユーザー取得

## 12
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:493: * 12. getConfigUserInfo - 設定用ユーザー情報取得

## 13
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:504: * 13. getCurrentUserEmail - 現在のユーザーメール取得

## 14
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:512: * 14. getUserId - ユーザーID取得・生成

## 15
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:521: * 15. getAllUsers - 全ユーザー一覧取得

## admin
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:22: buildAdminUrl: (userId) => UUtilities.generatorFactory.url.admin(userId)

## customUI
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:30: createCustomUI: (requestUserId, config) => UUtilities.generatorFactory.form.customUI(requestUserId, 

## quickStartUI
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:31: createQuickStartUI: (requestUserId) => UUtilities.generatorFactory.form.quickStartUI(requestUserId)

## sanitizeURL
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:171: static sanitizeURL(url) {

## REDACTED
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:947: sanitized[key] = '[REDACTED]';

## getValidationResult
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:957: getValidationResult(validationId) {

## getAllValidationResults
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:964: getAllValidationResults() {

## from
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:965: return Array.from(this.results.values());

## clearValidationResults
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:971: clearValidationResults(olderThanHours = 24) {

## utilities
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:2: * @fileoverview Web app URL utilities.

## retrieval
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:3: * Uses Apps Script API deployments for stable URL retrieval.

## deployment
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:17: * by returning the domain specific URL of the current deployment.

## failure
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:19: * @returns {string} Web app URL or empty string on failure.

## WEB_APP_URL_CACHE
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:23: const cacheKey = 'WEB_APP_URL_CACHE';

## getScriptId
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:50: const scriptId = ScriptApp.getScriptId();

## deploymentsList
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:69: const deployments = deploymentsList.deployments || [];

## URLs
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:136: * Generate user specific URLs.

## substr
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:202: url.searchParams.set('_t', Math.random().toString(36).substr(2, 9));

## E2E
- Occurrences: 1
- Files: workflowValidation.gs
- Samples:
  - workflowValidation.gs:3: * Claude Code開発環境のE2Eテスト用

## getActiveSheet
- Occurrences: 1
- Files: workflowValidation.gs
- Samples:
  - workflowValidation.gs:69: const sheet = SpreadsheetApp.getActiveSheet();
