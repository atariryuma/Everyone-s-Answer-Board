# Runtime Errors Report

Total: 922 undefined identifiers

## error
- Occurrences: 688
- Files: Code.gs, Core.gs, ErrorHandling.gs, auth.gs, autoInit.gs, cache.gs, config.gs, constants.gs, database.gs, debugConfig.gs, errorHandler.gs, lockManager.gs, main.gs, monitoring.gs, resilientExecutor.gs, secretManager.gs, session-utils.gs, systemIntegrationManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, unifiedValidationSystem.gs, url.gs, workflowValidation.gs
- Samples:
  - Code.gs:52: console.error(`Error in ${context}:`, error);
  - Code.gs:54: return { status: 'error', message: error.message || 'エラーが発生しました。' };
  - Code.gs:85: console.error('Failed to get user email:', e);

## ID
- Occurrences: 331
- Files: Code.gs, Core.gs, auth.gs, autoInit.gs, cache.gs, constants.gs, database.gs, keyUtils.gs, main.gs, secretManager.gs, session-utils.gs, setup.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs, url.gs
- Samples:
  - Code.gs:418: throw new Error('ユーザーIDが指定されていません。');
  - Code.gs:421: throw new Error('無効なユーザーIDの形式です。');
  - Code.gs:433: // ユーザーIDがない場合は登録画面を表示

## push
- Occurrences: 242
- Files: Code.gs, Core.gs, cache.gs, config.gs, constants.gs, database.gs, main.gs, monitoring.gs, secretManager.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, validationMigration.gs
- Samples:
  - Code.gs:1086: if (nameHeader) required.push(nameHeader);
  - Code.gs:1154: required.push(nameHeader);
  - Code.gs:1282: required.push(nameHeader);

## warn
- Occurrences: 223
- Files: Code.gs, Core.gs, auth.gs, autoInit.gs, cache.gs, constants.gs, database.gs, debugConfig.gs, errorHandler.gs, main.gs, monitoring.gs, session-utils.gs, ulog.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Code.gs:571: console.warn('Config not found for sheet:', activeSheetName, error);
  - Code.gs:794: console.warn('Failed to auto-create config for new sheet:', configError.message);
  - Code.gs:851: console.warn('Failed to get sheets:', error);

## URL
- Occurrences: 195
- Files: Code.gs, Core.gs, auth.gs, cache.gs, database.gs, main.gs, resilientExecutor.gs, session-utils.gs, unifiedCacheManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Code.gs:700: * 新しいスプレッドシートURLを追加し、アクティブシートとして設定します。
  - Code.gs:701: * @param {string} spreadsheetUrl - 追加するスプレッドシートのURL
  - Code.gs:729: // URLからスプレッドシートIDを抽出

## ERROR
- Occurrences: 170
- Files: autoInit.gs, constants.gs, debugConfig.gs, errorHandler.gs, lockManager.gs, monitoring.gs, secretManager.gs, setup.gs, systemIntegrationManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, unifiedValidationSystem.gs, workflowValidation.gs
- Samples:
  - autoInit.gs:54: console.error('[ERROR]', '❌ 自動システム初期化エラー:', error);
  - autoInit.gs:73: console.error('[ERROR]', '🚨 初期セキュリティチェックで重要な問題を検出', securityCheck);
  - autoInit.gs:147: console.error('[ERROR]', '定期メンテナンスエラー:', error);

## includes
- Occurrences: 129
- Files: Code.gs, Core.gs, auth.gs, cache.gs, database.gs, debugConfig.gs, main.gs, resilientExecutor.gs, secretManager.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, unifiedValidationSystem.gs, url.gs
- Samples:
  - Code.gs:171: isAdmin: admins.includes(userEmail)
  - Code.gs:173: return admins.includes(userEmail);
  - Code.gs:180: isAdmin: admins.includes(userEmail)

## summary
- Occurrences: 128
- Files: database.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - database.gs:892: if (opts.autoRepair && diagnosis.summary.criticalIssues.length > 0) {
  - database.gs:1790: diagnosticResult.summary.criticalIssues.push('DATABASE_SPREADSHEET_ID が設定されていません');
  - database.gs:1791: diagnosticResult.summary.overallStatus = 'critical';

## toISOString
- Occurrences: 127
- Files: Code.gs, Core.gs, auth.gs, cache.gs, constants.gs, database.gs, debugConfig.gs, errorHandler.gs, main.gs, monitoring.gs, secretManager.gs, session-utils.gs, systemIntegrationManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Code.gs:1183: const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
  - Code.gs:1311: const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
  - Code.gs:2315: timestamp: timestamp.toISOString(),

## debug
- Occurrences: 100
- Files: autoInit.gs, constants.gs, debugConfig.gs, monitoring.gs, ulog.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs
- Samples:
  - autoInit.gs:13: ULog.debug('💨 システムは既に初期化済みです');
  - autoInit.gs:95: ULog.debug('🔥 システムキャッシュウォームアップ開始');
  - autoInit.gs:100: ULog.debug('📊 データベースID キャッシュ済み');

## API
- Occurrences: 99
- Files: Code.gs, Core.gs, cache.gs, constants.gs, database.gs, monitoring.gs, resilientExecutor.gs, secretManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, url.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Code.gs:1743: throw new Error('Google Forms API is not available');
  - Code.gs:1815: if (error.message.includes('Google Forms API is not available') || error.message.includes('FormApp')
  - Core.gs:3: * 主要な業務ロジックとAPI エンドポイント

## stringify
- Occurrences: 95
- Files: Code.gs, Core.gs, auth.gs, cache.gs, constants.gs, database.gs, errorHandler.gs, main.gs, monitoring.gs, secretManager.gs, ulog.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, workflowValidation.gs
- Samples:
  - Code.gs:1548: cache.put(cacheKey, JSON.stringify(indices), 21600); // 6 hours cache
  - Code.gs:1942: JSON.stringify({
  - Code.gs:2082: cache.put(cacheKey, JSON.stringify(userInfo), 1800);

## parse
- Occurrences: 87
- Files: Code.gs, Core.gs, auth.gs, cache.gs, constants.gs, database.gs, main.gs, monitoring.gs, secretManager.gs, setup.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs
- Samples:
  - Code.gs:1527: return JSON.parse(cached);
  - Code.gs:2051: return JSON.parse(cached);
  - Code.gs:2069: userInfo[header] = JSON.parse(data[i][index] || '{}');

## getProperty
- Occurrences: 80
- Files: Code.gs, Core.gs, auth.gs, constants.gs, database.gs, debugConfig.gs, main.gs, monitoring.gs, secretManager.gs, session-utils.gs, setup.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, workflowValidation.gs
- Samples:
  - Code.gs:66: const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  - Code.gs:123: const str = props.getProperty(`ADMIN_EMAILS_${spreadsheetId}`);
  - Code.gs:128: const str = props.getProperty('ADMIN_EMAILS') || '';

## getScriptProperties
- Occurrences: 63
- Files: Code.gs, Core.gs, auth.gs, constants.gs, database.gs, debugConfig.gs, main.gs, monitoring.gs, secretManager.gs, setup.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, workflowValidation.gs
- Samples:
  - Code.gs:119: const props = PropertiesService.getScriptProperties();
  - Code.gs:193: const props = PropertiesService.getScriptProperties();
  - Code.gs:203: const props = PropertiesService.getScriptProperties();

## getEmail
- Occurrences: 60
- Files: Code.gs, Core.gs, auth.gs, database.gs, errorHandler.gs, main.gs, monitoring.gs, secretManager.gs, setup.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Code.gs:79: const email = Session.getActiveUser().getEmail();
  - Code.gs:138: const ownerEmail = owner.getEmail();
  - Code.gs:291: if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {

## forEach
- Occurrences: 59
- Files: Code.gs, Core.gs, cache.gs, config.gs, constants.gs, database.gs, main.gs, session-utils.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs
- Samples:
  - Code.gs:1387: REACTION_KEYS.forEach((k, idx) => {
  - Code.gs:1392: Object.keys(lists).forEach(key => {
  - Code.gs:1581: required.forEach(h => {

## trim
- Occurrences: 58
- Files: Code.gs, Core.gs, auth.gs, cache.gs, config.gs, constants.gs, database.gs, main.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Code.gs:124: if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:129: adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:214: .map(s => s.trim())

## getActiveUser
- Occurrences: 55
- Files: Code.gs, Core.gs, auth.gs, database.gs, errorHandler.gs, main.gs, monitoring.gs, secretManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Code.gs:79: const email = Session.getActiveUser().getEmail();
  - Code.gs:291: if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
  - Code.gs:333: if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {

## remove
- Occurrences: 52
- Files: Code.gs, Core.gs, auth.gs, cache.gs, database.gs, session-utils.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, url.gs
- Samples:
  - Code.gs:1561: cache.remove(`headers_${sheetName}`);
  - Code.gs:2177: cache.remove(cacheKey);
  - Core.gs:2268: cacheManager.remove(cacheKey);

## map
- Occurrences: 48
- Files: Code.gs, Core.gs, config.gs, constants.gs, database.gs, session-utils.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, validationMigration.gs
- Samples:
  - Code.gs:124: if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:129: adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:214: .map(s => s.trim())

## replace
- Occurrences: 46
- Files: Code.gs, auth.gs, main.gs, session-utils.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Code.gs:935: .replace(/[\uFF01-\uFF5E]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
  - Code.gs:936: .replace(/\s+/g, ' ')
  - Code.gs:1140: const sheetHeaders = allValues[0].map(h => String(h || '').replace(/\s+/g,''));

## keys
- Occurrences: 46
- Files: Code.gs, Core.gs, cache.gs, constants.gs, database.gs, main.gs, secretManager.gs, session-utils.gs, ulog.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs, unifiedValidationSystem.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Code.gs:944: const keyNorm = keys.map(k => normalize(k));
  - Code.gs:1392: Object.keys(lists).forEach(key => {
  - Code.gs:2180: auditLog('CONFIG_UPDATED', userId, { currentUser, updatedFields: Object.keys(sanitizedConfig), newCo

## tests
- Occurrences: 41
- Files: unifiedValidationSystem.gs, workflowValidation.gs
- Samples:
  - unifiedValidationSystem.gs:90: results.tests.userAuth = this._checkUserAuthentication(options.userId);
  - unifiedValidationSystem.gs:91: results.tests.sessionValid = this._checkSessionValidity(options.userId);
  - unifiedValidationSystem.gs:96: results.tests.userAccess = this._checkUserAccess(options.userId);

## clearByPattern
- Occurrences: 38
- Files: cache.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:245: clearByPattern(pattern, options = {}) {
  - cache.gs:364: const removed = this.clearByPattern(pattern, { strict: false, maxKeys: 100 });
  - cache.gs:384: const removed = this.clearByPattern(pattern, { strict: false, maxKeys: 50 });

## setTitle
- Occurrences: 37
- Files: Code.gs, Core.gs, main.gs
- Samples:
  - Code.gs:437: .setTitle('StudyQuest - 新規登録')
  - Code.gs:453: .setTitle('エラー');
  - Code.gs:472: return output.setTitle('StudyQuest - 回答ボード');

## manager
- Occurrences: 37
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:2004: this.manager.remove(`user_${identifier}`);
  - unifiedCacheManager.gs:2005: this.manager.remove(`userinfo_${identifier}`);
  - unifiedCacheManager.gs:2006: this.manager.remove(`unified_user_info_${identifier}`);

## SHEET_NAME
- Occurrences: 36
- Files: Code.gs, constants.gs, database.gs, main.gs, unifiedUserManager.gs
- Samples:
  - Code.gs:12: SHEET_NAME: 'Users',
  - Code.gs:754: const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  - Code.gs:1934: const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);

## SEVERITY
- Occurrences: 35
- Files: constants.gs, errorHandler.gs, setup.gs, unifiedCacheManager.gs, workflowValidation.gs
- Samples:
  - constants.gs:36: SEVERITY: {
  - constants.gs:249: Object.freeze(UNIFIED_CONSTANTS.ERROR.SEVERITY);
  - constants.gs:274: /** @deprecated Use UNIFIED_CONSTANTS.ERROR.SEVERITY instead */

## toString
- Occurrences: 34
- Files: Code.gs, Core.gs, cache.gs, constants.gs, database.gs, main.gs, resilientExecutor.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, url.gs
- Samples:
  - Code.gs:103: return (email || '').toString().split('@').pop().toLowerCase();
  - Code.gs:212: .toString()
  - Code.gs:2297: const randomString = uuid + timestamp + Math.random().toString(36);

## indexOf
- Occurrences: 33
- Files: Code.gs, Core.gs, database.gs
- Samples:
  - Code.gs:757: const userIdIndex = headers.indexOf('userId');
  - Code.gs:758: const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
  - Code.gs:759: const spreadsheetUrlIndex = headers.indexOf('spreadsheetUrl');

## CATEGORIES
- Occurrences: 33
- Files: constants.gs, debugConfig.gs, resilientExecutor.gs, setup.gs, ulog.gs, workflowValidation.gs
- Samples:
  - constants.gs:44: CATEGORIES: {
  - constants.gs:250: Object.freeze(UNIFIED_CONSTANTS.ERROR.CATEGORIES);
  - constants.gs:275: // Legacy constants removed - use UNIFIED_CONSTANTS.ERROR.SEVERITY and UNIFIED_CONSTANTS.ERROR.CATEG

## GAS
- Occurrences: 32
- Files: Code.gs, Core.gs, auth.gs, cache.gs, constants.gs, database.gs, errorHandler.gs, secretManager.gs, ulog.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs, url.gs, workflowValidation.gs
- Samples:
  - Code.gs:414: // GAS Webアプリケーションのエントリーポイント
  - Core.gs:4032: * 1. GASスクリプトの編集権限を確認
  - auth.gs:3: * GAS互換の関数ベースの実装

## getContentText
- Occurrences: 32
- Files: Core.gs, auth.gs, cache.gs, database.gs, secretManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:4016: return JSON.parse(response.getContentText());
  - auth.gs:81: console.error('Response:', response.getContentText());
  - auth.gs:85: var responseData = JSON.parse(response.getContentText());

## getResponseCode
- Occurrences: 32
- Files: auth.gs, cache.gs, database.gs, resilientExecutor.gs, secretManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - auth.gs:79: if (response.getResponseCode() !== 200) {
  - auth.gs:80: console.error('Token request failed. Status:', response.getResponseCode());
  - auth.gs:82: throw new Error('サービスアカウントトークンの取得に失敗しました。認証情報を確認してください。Status: ' + response.getResponseCode());

## CRITICAL
- Occurrences: 32
- Files: autoInit.gs, constants.gs, errorHandler.gs, systemIntegrationManager.gs, ulog.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - autoInit.gs:72: if (securityCheck.overallStatus === 'CRITICAL') {
  - constants.gs:11: * @property {string} CRITICAL - 致命的なエラー
  - constants.gs:40: CRITICAL: 'critical',

## toLowerCase
- Occurrences: 31
- Files: Code.gs, Core.gs, auth.gs, main.gs, resilientExecutor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Code.gs:103: return (email || '').toString().split('@').pop().toLowerCase();
  - Code.gs:938: .toLowerCase();
  - Code.gs:1207: highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',

## getRange
- Occurrences: 31
- Files: Code.gs, Core.gs, config.gs, unifiedSheetDataManager.gs, workflowValidation.gs
- Samples:
  - Code.gs:764: userDb.getRange(i + 1, spreadsheetIdIndex + 1).setValue(spreadsheetId);
  - Code.gs:765: userDb.getRange(i + 1, spreadsheetUrlIndex + 1).setValue(spreadsheetUrl);
  - Code.gs:784: const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];

## DATABASE_SPREADSHEET_ID
- Occurrences: 31
- Files: Core.gs, constants.gs, database.gs, main.gs, monitoring.gs, secretManager.gs, setup.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:1371: props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);
  - Core.gs:1395: var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  - constants.gs:226: DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',

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

## filter
- Occurrences: 30
- Files: Code.gs, Core.gs, cache.gs, config.gs, constants.gs, database.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs, unifiedValidationSystem.gs, validationMigration.gs, workflowValidation.gs
- Samples:
  - Code.gs:124: if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:129: adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:215: .filter(Boolean);

## getSheetByName
- Occurrences: 30
- Files: Code.gs, Core.gs, config.gs, database.gs, unifiedSheetDataManager.gs
- Samples:
  - Code.gs:301: const sheet = spreadsheet.getSheetByName(sheetName);
  - Code.gs:650: const sheet = spreadsheet.getSheetByName(targetSheetName);
  - Code.gs:754: const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);

## memoCache
- Occurrences: 30
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:48: if (enableMemoization && this.memoCache.has(key)) {
  - cache.gs:50: const memoEntry = this.memoCache.get(key);
  - cache.gs:60: this.memoCache.delete(key);

## CACHE
- Occurrences: 30
- Files: constants.gs, debugConfig.gs, errorHandler.gs, ulog.gs
- Samples:
  - constants.gs:19: * @property {string} CACHE - キャッシュエラー
  - constants.gs:48: CACHE: 'cache',
  - constants.gs:60: CACHE: {

## checks
- Occurrences: 30
- Files: database.gs, monitoring.gs
- Samples:
  - database.gs:1784: diagnosticResult.checks.databaseConfig = {
  - database.gs:1799: diagnosticResult.checks.serviceConnection = { status: 'success' };
  - database.gs:1801: diagnosticResult.checks.serviceConnection = {

## UNDERSTAND
- Occurrences: 28
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:32: UNDERSTAND: 'なるほど！',
  - Code.gs:45: const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];
  - Code.gs:865: return sum + (reactions.UNDERSTAND?.count || 0) +

## LIKE
- Occurrences: 28
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:33: LIKE: 'いいね！',
  - Code.gs:45: const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];
  - Code.gs:866: (reactions.LIKE?.count || 0) +

## CURIOUS
- Occurrences: 28
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:34: CURIOUS: 'もっと知りたい！',
  - Code.gs:45: const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];
  - Code.gs:867: (reactions.CURIOUS?.count || 0);

## HEADERS
- Occurrences: 27
- Files: Code.gs, constants.gs, database.gs, main.gs
- Samples:
  - Code.gs:13: HEADERS: [
  - Code.gs:2247: sheet.appendRow(USER_DB_CONFIG.HEADERS);
  - constants.gs:116: HEADERS: [

## slice
- Occurrences: 27
- Files: Code.gs, Core.gs, cache.gs, constants.gs, database.gs, main.gs, monitoring.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - Code.gs:1090: const entries = values.slice(1).map(row => ({
  - Code.gs:1164: const dataRows = allValues.slice(1);
  - Code.gs:1292: const dataRows = allValues.slice(1);

## sleep
- Occurrences: 27
- Files: Core.gs, cache.gs, constants.gs, database.gs, resilientExecutor.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:1991: Utilities.sleep(500);
  - Core.gs:3491: Utilities.sleep(1000);
  - cache.gs:735: Utilities.sleep(waitTime);

## INFO
- Occurrences: 27
- Files: autoInit.gs, constants.gs, debugConfig.gs, monitoring.gs, ulog.gs, unifiedUtilities.gs, workflowValidation.gs
- Samples:
  - autoInit.gs:27: logLevel: 'INFO',
  - constants.gs:234: LOG_LEVEL: 'INFO',
  - constants.gs:769: logUnified('INFO', functionName, '実行開始', { timeout: timeout });

## substring
- Occurrences: 27
- Files: constants.gs, database.gs, errorHandler.gs, secretManager.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - constants.gs:1054: params = params.substring(1);
  - database.gs:121: String(executorEmail).substring(0, 255), // 長さ制限
  - database.gs:122: String(targetUserId).substring(0, 255),

## warnLog
- Occurrences: 26
- Files: resilientExecutor.gs, secretManager.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedUserManager.gs
- Samples:
  - resilientExecutor.gs:119: warnLog(
  - resilientExecutor.gs:218: warnLog(
  - resilientExecutor.gs:315: warnLog(`resilientUrlFetch: クライアントエラー ${responseCode} - ${url}`);

## getValues
- Occurrences: 25
- Files: Code.gs, Core.gs, config.gs, unifiedSheetDataManager.gs, workflowValidation.gs
- Samples:
  - Code.gs:755: const data = userDb.getDataRange().getValues();
  - Code.gs:784: const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];
  - Code.gs:917: const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

## join
- Occurrences: 25
- Files: Code.gs, Core.gs, cache.gs, config.gs, constants.gs, database.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedUserManager.gs, unifiedValidationSystem.gs
- Samples:
  - Code.gs:1400: REACTION_KEYS.map(k => lists[k].arr.join(','))
  - Code.gs:1601: throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`);
  - Code.gs:2399: scriptProps.setProperty(key, list.join(','));

## warnings
- Occurrences: 25
- Files: autoInit.gs, database.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - autoInit.gs:34: warnings: initResult.warnings.length,
  - database.gs:1881: diagnosticResult.summary.warnings.push(cacheStatus.staleEntries + ' 個の古いキャッシュエントリが見つかりました');
  - database.gs:1892: diagnosticResult.summary.overallStatus = diagnosticResult.summary.warnings.length > 0 ? 'warning' : 

## freeze
- Occurrences: 25
- Files: constants.gs
- Samples:
  - constants.gs:247: Object.freeze(UNIFIED_CONSTANTS);
  - constants.gs:248: Object.freeze(UNIFIED_CONSTANTS.ERROR);
  - constants.gs:249: Object.freeze(UNIFIED_CONSTANTS.ERROR.SEVERITY);

## getUserProperties
- Occurrences: 24
- Files: Code.gs, constants.gs, database.gs, main.gs, session-utils.gs
- Samples:
  - Code.gs:65: const props = PropertiesService.getUserProperties();
  - Code.gs:153: const userProps = PropertiesService.getUserProperties();
  - Code.gs:235: const userProps = PropertiesService.getUserProperties();

## OK
- Occurrences: 24
- Files: cache.gs, constants.gs, monitoring.gs, secretManager.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs
- Samples:
  - cache.gs:560: console.log(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, Script
  - cache.gs:560: console.log(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, Script
  - constants.gs:1117: const performBasicHealthCheck = () => ({ status: 'ok', message: 'Basic health OK' });

## setProperty
- Occurrences: 23
- Files: Code.gs, Core.gs, constants.gs, main.gs, monitoring.gs, secretManager.gs, session-utils.gs, setup.gs, unifiedCacheManager.gs, unifiedUserManager.gs, workflowValidation.gs
- Samples:
  - Code.gs:194: props.setProperty('IS_PUBLISHED', 'true');
  - Code.gs:195: props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
  - Code.gs:204: props.setProperty('IS_PUBLISHED', 'false');

## HTML
- Occurrences: 23
- Files: Core.gs, auth.gs, constants.gs, main.gs, session-utils.gs, unifiedUtilities.gs
- Samples:
  - Core.gs:1472: // HTML依存関数（UI連携）
  - auth.gs:223: * @returns {HtmlOutput} 表示するHTMLコンテンツ
  - constants.gs:510: * HTMLからアクセス可能な定数のサブセットを返す

## MEDIUM
- Occurrences: 23
- Files: constants.gs, errorHandler.gs, setup.gs, workflowValidation.gs
- Samples:
  - constants.gs:9: * @property {string} MEDIUM - 中程度のエラー
  - constants.gs:38: MEDIUM: 'medium',
  - constants.gs:64: MEDIUM: 600, // 10分

## SYSTEM
- Occurrences: 23
- Files: constants.gs, debugConfig.gs, errorHandler.gs, setup.gs, ulog.gs, unifiedValidationSystem.gs
- Samples:
  - constants.gs:22: * @property {string} SYSTEM - システムエラー
  - constants.gs:51: SYSTEM: 'system',
  - constants.gs:427: SYSTEM: 'system',

## openById
- Occurrences: 22
- Files: Code.gs, Core.gs, database.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - Code.gs:74: return SpreadsheetApp.openById(spreadsheetId);
  - Code.gs:300: const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  - Code.gs:649: const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

## getScriptCache
- Occurrences: 22
- Files: Code.gs, Core.gs, cache.gs, session-utils.gs, unifiedCacheManager.gs, unifiedUserManager.gs, url.gs, workflowValidation.gs
- Samples:
  - Code.gs:1520: const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  - Code.gs:1558: const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  - Code.gs:1710: const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;

## customConfig
- Occurrences: 22
- Files: Core.gs
- Samples:
  - Core.gs:2613: if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && 
  - Core.gs:2613: if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && 
  - Core.gs:2613: if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && 

## global
- Occurrences: 21
- Files: Code.gs, cache.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:60: if (typeof global !== 'undefined' && global.getConfig) {
  - Code.gs:60: if (typeof global !== 'undefined' && global.getConfig) {
  - Code.gs:61: getConfig = global.getConfig;

## getId
- Occurrences: 21
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1773: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  - Code.gs:1805: formId: form.getId(),
  - Code.gs:1807: spreadsheetId: spreadsheet.getId(),

## WARN
- Occurrences: 21
- Files: constants.gs, debugConfig.gs, monitoring.gs, ulog.gs, unifiedUtilities.gs, validationMigration.gs
- Samples:
  - constants.gs:599: console.warn(`[WARN] ${functionName}: ${error.message}`, errorInfo);
  - constants.gs:824: logUnified('WARN', functionName, `試行${attempt}失敗、${waitTime}ms後に再試行`, {
  - constants.gs:877: logUnified('WARN', functionName, '配列以外の値が渡されました', { type: typeof array });

## EMAIL
- Occurrences: 20
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:27: EMAIL: 'メールアドレス',
  - Code.gs:1085: const required = [answerHeader, reasonHeader, classHeader, COLUMN_HEADERS.EMAIL];
  - Code.gs:1144: COLUMN_HEADERS.EMAIL,

## REASON
- Occurrences: 20
- Files: Code.gs, Core.gs, cache.gs, constants.gs, main.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:30: REASON: '理由',
  - Code.gs:1081: const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
  - Code.gs:1136: const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;

## TTL
- Occurrences: 20
- Files: auth.gs, cache.gs, constants.gs, database.gs, unifiedCacheManager.gs
- Samples:
  - auth.gs:151: ttl: 60 // 短いTTLで最新性を確保
  - cache.gs:14: this.defaultTTL = 21600; // デフォルトTTL（6時間）
  - cache.gs:1008: details: 'TTL設定の見直し、メモ化の活用、キャッシュキー設計の最適化を検討してください。'

## relatedIds
- Occurrences: 20
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:350: if (relatedIds.length > maxRelated) {
  - cache.gs:351: console.warn(`[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`);
  - cache.gs:352: relatedIds = relatedIds.slice(0, maxRelated);

## XFrameOptionsMode
- Occurrences: 20
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:150: HtmlService.XFrameOptionsMode.DENY) {
  - main.gs:151: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
  - main.gs:724: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {

## split
- Occurrences: 19
- Files: Code.gs, Core.gs, constants.gs, main.gs, ulog.gs
- Samples:
  - Code.gs:103: return (email || '').toString().split('@').pop().toLowerCase();
  - Code.gs:124: if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  - Code.gs:129: adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);

## SERVICE_ACCOUNT_CREDS
- Occurrences: 19
- Files: Core.gs, auth.gs, constants.gs, database.gs, main.gs, secretManager.gs, setup.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:1370: props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
  - Core.gs:1396: var creds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
  - Core.gs:3178: var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));

## recommendations
- Occurrences: 19
- Files: cache.gs, database.gs, systemIntegrationManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - cache.gs:1005: analysis.recommendations.push({
  - cache.gs:1013: analysis.recommendations.push({
  - cache.gs:1021: analysis.recommendations.push({

## criticalIssues
- Occurrences: 19
- Files: database.gs, unifiedSecurityManager.gs
- Samples:
  - database.gs:892: if (opts.autoRepair && diagnosis.summary.criticalIssues.length > 0) {
  - database.gs:1790: diagnosticResult.summary.criticalIssues.push('DATABASE_SPREADSHEET_ID が設定されていません');
  - database.gs:1805: diagnosticResult.summary.criticalIssues.push('Sheets サービス接続失敗: ' + serviceError.message);

## unifiedResult
- Occurrences: 19
- Files: validationMigration.gs
- Samples:
  - validationMigration.gs:434: isValid: unifiedResult.success,
  - validationMigration.gs:435: errors: unifiedResult.success ? [] : extractErrorMessages(unifiedResult)
  - validationMigration.gs:440: hasAccess: unifiedResult.success,

## OPINION
- Occurrences: 18
- Files: Code.gs, Core.gs, cache.gs, constants.gs, main.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:29: OPINION: '回答',
  - Code.gs:1080: const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
  - Code.gs:1135: const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;

## HIGHLIGHT
- Occurrences: 18
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:35: HIGHLIGHT: 'ハイライト'
  - Code.gs:1151: COLUMN_HEADERS.HIGHLIGHT
  - Code.gs:1181: const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];

## getName
- Occurrences: 18
- Files: Code.gs, Core.gs, database.gs
- Samples:
  - Code.gs:746: console.log('Spreadsheet name:', testSpreadsheet.getName());
  - Code.gs:774: const firstSheetName = sheets[0].getName();
  - Code.gs:805: console.log('Successfully added spreadsheet:', testSpreadsheet.getName());

## SHEETS
- Occurrences: 18
- Files: constants.gs
- Samples:
  - constants.gs:112: SHEETS: {
  - constants.gs:258: Object.freeze(UNIFIED_CONSTANTS.SHEETS);
  - constants.gs:259: Object.freeze(UNIFIED_CONSTANTS.SHEETS.DATABASE);

## circuitBreaker
- Occurrences: 18
- Files: resilientExecutor.gs
- Samples:
  - resilientExecutor.gs:55: if (this.circuitBreaker.state === 'OPEN') {
  - resilientExecutor.gs:56: const timeSinceLastFailure = Date.now() - this.circuitBreaker.lastFailureTime;
  - resilientExecutor.gs:57: if (timeSinceLastFailure < this.circuitBreaker.openTimeout) {

## CURRENT_USER_ID
- Occurrences: 17
- Files: Code.gs, database.gs, session-utils.gs
- Samples:
  - Code.gs:154: const userId = userProps.getProperty('CURRENT_USER_ID');
  - Code.gs:236: const userId = userProps.getProperty('CURRENT_USER_ID');
  - Code.gs:283: const userId = props.getProperty('CURRENT_USER_ID');

## createTemplateFromFile
- Occurrences: 17
- Files: Code.gs, main.gs, unifiedUtilities.gs
- Samples:
  - Code.gs:435: return HtmlService.createTemplateFromFile('Registration')
  - Code.gs:462: const template = HtmlService.createTemplateFromFile('Unpublished');
  - Code.gs:499: const template = HtmlService.createTemplateFromFile('AdminPanel');

## some
- Occurrences: 17
- Files: Code.gs, Core.gs, cache.gs, database.gs, errorHandler.gs, resilientExecutor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedValidationSystem.gs, url.gs
- Samples:
  - Code.gs:945: const found = normalizedHeaders.find(h => keyNorm.some(k => h.normalized.includes(k)));
  - Code.gs:952: const isGoogleForm = normalizedHeaders.some(h =>
  - Core.gs:4834: if (spreadsheetInfo.sheets.some(sheet => sheet.properties.title === commonName)) {

## getTime
- Occurrences: 17
- Files: Code.gs, Core.gs, constants.gs, database.gs, ulog.gs, unifiedValidationSystem.gs
- Samples:
  - Code.gs:2296: const timestamp = new Date().getTime();
  - Code.gs:2324: const cacheKey = `audit_${timestamp.getTime()}_${uuid}`;
  - Core.gs:15: const stopTime = new Date(publishTime.getTime() + (minutes * 60 * 1000));

## googleusercontent
- Occurrences: 17
- Files: main.gs, url.gs
- Samples:
  - main.gs:1184: // 開発モードURLのチェック（googleusercontent.comは有効なデプロイURLも含むため調整）
  - main.gs:1190: // 最終的な URL 妥当性チェック（googleusercontent.comも有効URLとして認識）
  - main.gs:1192: cleanUrl.includes('googleusercontent.com') ||

## test
- Occurrences: 16
- Files: Code.gs, constants.gs, database.gs, errorHandler.gs, secretManager.gs, url.gs, workflowValidation.gs
- Samples:
  - Code.gs:99: return emailRegex.test(email);
  - Code.gs:420: if (!/^[a-zA-Z0-9-_]{10,}$/.test(userId)) {
  - Code.gs:1618: if (/script\.googleusercontent\.com/.test(url) && deployId) {

## evaluate
- Occurrences: 16
- Files: Code.gs, main.gs
- Samples:
  - Code.gs:436: .evaluate()
  - Code.gs:470: const output = template.evaluate();
  - Code.gs:503: const output = template.evaluate();

## setRequired
- Occurrences: 16
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1755: .setRequired(true);
  - Code.gs:1759: .setRequired(true);
  - Code.gs:1764: .setRequired(true);

## UI
- Occurrences: 16
- Files: Core.gs, constants.gs, debugConfig.gs, main.gs, ulog.gs
- Samples:
  - Core.gs:1472: // HTML依存関数（UI連携）
  - Core.gs:1476: * 現在のユーザーのステータスを取得し、UIに表示するための情報を返します。(マルチテナント対応版)
  - Core.gs:4307: * クイックスタート用フォーム作成（UI用ラッパー）

## googleapis
- Occurrences: 16
- Files: Core.gs, auth.gs, database.gs, secretManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:4009: baseUrl: 'https://www.googleapis.com/drive/v3',
  - auth.gs:46: var tokenUrl = "https://www.googleapis.com/oauth2/v4/token";
  - auth.gs:55: scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",

## COMPREHENSIVE
- Occurrences: 16
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:22: COMPREHENSIVE: 'comprehensive'
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t
  - unifiedValidationSystem.gs:95: if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {

## CLASS
- Occurrences: 15
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:28: CLASS: 'クラス',
  - Code.gs:1082: const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
  - Code.gs:1137: const classHeaderGuess = cfg.classHeader || COLUMN_HEADERS.CLASS;

## put
- Occurrences: 15
- Files: Code.gs, Core.gs, cache.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedUserManager.gs, url.gs, workflowValidation.gs
- Samples:
  - Code.gs:1548: cache.put(cacheKey, JSON.stringify(indices), 21600); // 6 hours cache
  - Code.gs:1718: cache.put(key, String(attempts + 1), 3600);
  - Code.gs:2082: cache.put(cacheKey, JSON.stringify(userInfo), 1800);

## fetch
- Occurrences: 15
- Files: Core.gs, auth.gs, cache.gs, database.gs, resilientExecutor.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:4013: var response = UrlFetchApp.fetch(url, {
  - auth.gs:69: var response = UrlFetchApp.fetch(tokenUrl, {
  - cache.gs:679: var response = UrlFetchApp.fetch(url, {

## constructor
- Occurrences: 15
- Files: cache.gs, constants.gs, errorHandler.gs, monitoring.gs, resilientExecutor.gs, secretManager.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs, unifiedValidationSystem.gs
- Samples:
  - cache.gs:11: constructor() {
  - constants.gs:1049: constructor(params = {}) {
  - errorHandler.gs:34: constructor() {

## HIGH
- Occurrences: 15
- Files: constants.gs, errorHandler.gs, setup.gs, unifiedCacheManager.gs
- Samples:
  - constants.gs:10: * @property {string} HIGH - 重要なエラー
  - constants.gs:39: HIGH: 'high',
  - constants.gs:372: * @param {string} path - 定数パス（例: 'ERROR.SEVERITY.HIGH'）

## NAME
- Occurrences: 14
- Files: Code.gs, Core.gs, constants.gs, main.gs
- Samples:
  - Code.gs:31: NAME: '名前',
  - Code.gs:1686: COLUMN_HEADERS.NAME,
  - Code.gs:1845: COLUMN_HEADERS.NAME,

## getUrl
- Occurrences: 14
- Files: Code.gs, Core.gs, main.gs, session-utils.gs, unifiedCacheManager.gs, url.gs
- Samples:
  - Code.gs:1646: current = ScriptApp.getService().getUrl();
  - Code.gs:1808: spreadsheetUrl: spreadsheet.getUrl(),
  - Code.gs:1871: spreadsheetUrl: spreadsheet.getUrl(),

## clearAll
- Occurrences: 14
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:531: clearAll() {
  - cache.gs:560: console.log(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, Script
  - unifiedCacheManager.gs:702: clearAll() {

## preWarmedItems
- Occurrences: 14
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:897: results.preWarmedItems.push('service_account_token');
  - cache.gs:908: results.preWarmedItems.push('user_by_email');
  - cache.gs:913: results.preWarmedItems.push('user_by_id');

## DATABASE
- Occurrences: 14
- Files: constants.gs, database.gs, debugConfig.gs, errorHandler.gs, unifiedCacheManager.gs, workflowValidation.gs
- Samples:
  - constants.gs:18: * @property {string} DATABASE - データベースエラー
  - constants.gs:47: DATABASE: 'database',
  - constants.gs:114: DATABASE: {

## LockService
- Occurrences: 13
- Files: Code.gs, Core.gs, database.gs, lockManager.gs
- Samples:
  - Code.gs:1363: const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  - Code.gs:1363: const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  - Code.gs:1441: const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;

## AdminPanel
- Occurrences: 13
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:1014: * AdminPanel.htmlから呼び出される
  - Core.gs:1204: appPublished: configJson.appPublished || false, // AdminPanel.htmlで使用される
  - Core.gs:1206: allSheets: sheets, // AdminPanel.htmlで使用される

## mainQuestion
- Occurrences: 13
- Files: Core.gs
- Samples:
  - Core.gs:2598: mainItem.setTitle(config.mainQuestion.title);
  - Core.gs:2626: var mainQuestionTitle = customConfig.mainQuestion ? customConfig.mainQuestion.title : '今回のテーマについて、あな
  - Core.gs:2628: var questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';

## delete
- Occurrences: 13
- Files: cache.gs, constants.gs, secretManager.gs, unifiedCacheManager.gs, unifiedValidationSystem.gs
- Samples:
  - cache.gs:60: this.memoCache.delete(key);
  - cache.gs:185: this.memoCache.delete(key);
  - cache.gs:216: this.memoCache.delete(key);

## startsWith
- Occurrences: 13
- Files: constants.gs, main.gs, secretManager.gs, session-utils.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs, unifiedUtilities.gs, url.gs
- Samples:
  - constants.gs:1053: if (params.startsWith('?')) {
  - main.gs:1162: if ((cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
  - main.gs:1163: (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))) {

## actions
- Occurrences: 13
- Files: database.gs
- Samples:
  - database.gs:2066: result.summary.actions.push('サービスアカウントの再共有実行');
  - database.gs:2083: result.summary.actions.push('アクセス権限修復成功');
  - database.gs:2091: result.summary.actions.push('修復後テスト失敗: ' + retestError.message);

## deleteProperty
- Occurrences: 12
- Files: Code.gs, constants.gs, database.gs, main.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:205: if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');
  - constants.gs:1165: props.deleteProperty(sessionKey);
  - constants.gs:1171: props.deleteProperty(sessionKey);

## Z0
- Occurrences: 12
- Files: Code.gs, Core.gs, constants.gs, database.gs, main.gs, secretManager.gs
- Samples:
  - Code.gs:420: if (!/^[a-zA-Z0-9-_]{10,}$/.test(userId)) {
  - Code.gs:731: const urlPattern = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  - Code.gs:2277: /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,

## Page
- Occurrences: 12
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:563: // 既存のPage.htmlを使用
  - Core.gs:363: * Page.htmlから呼び出される - フロントエンド期待形式に対応
  - Core.gs:388: // Page.html期待形式に変換

## fromCharCode
- Occurrences: 12
- Files: Code.gs, Core.gs, database.gs, secretManager.gs
- Samples:
  - Code.gs:935: .replace(/[\uFF01-\uFF5E]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
  - Core.gs:2218: var range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
  - Core.gs:2368: var range = "'" + sheetName + "'!" + String.fromCharCode(65 + columnIndex) + rowIndex;

## ADMIN_EMAIL
- Occurrences: 12
- Files: Core.gs, database.gs, main.gs, setup.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:1375: props.setProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL, adminEmail);
  - Core.gs:4039: var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
  - database.gs:2757: SCRIPT_PROPS_KEYS.ADMIN_EMAIL

## DEBUG_MODE
- Occurrences: 12
- Files: constants.gs, debugConfig.gs, main.gs
- Samples:
  - constants.gs:228: DEBUG_MODE: 'DEBUG_MODE',
  - constants.gs:228: DEBUG_MODE: 'DEBUG_MODE',
  - constants.gs:233: ENABLED: PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true',

## STANDARD
- Occurrences: 12
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:21: STANDARD: 'standard',
  - unifiedValidationSystem.gs:35: validate(category, level = this.validationLevels.STANDARD, options = {}) {
  - unifiedValidationSystem.gs:89: if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === t

## TIMESTAMP
- Occurrences: 11
- Files: Code.gs, constants.gs, main.gs
- Samples:
  - Code.gs:26: TIMESTAMP: 'タイムスタンプ',
  - Code.gs:1147: COLUMN_HEADERS.TIMESTAMP,
  - Code.gs:1182: const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];

## createHtmlOutput
- Occurrences: 11
- Files: Code.gs, main.gs, session-utils.gs, unifiedUtilities.gs
- Samples:
  - Code.gs:452: return HtmlService.createHtmlOutput('無効なユーザーIDです。')
  - Code.gs:475: return HtmlService.createHtmlOutput('認証が必要です。Googleアカウントでログインしてください。')
  - Code.gs:485: return HtmlService.createHtmlOutput('システムエラーが発生しました。管理者にお問い合わせください。')

## getDataRange
- Occurrences: 11
- Files: Code.gs, config.gs
- Samples:
  - Code.gs:755: const data = userDb.getDataRange().getValues();
  - Code.gs:1072: const values = sheet.getDataRange().getValues();
  - Code.gs:1105: const allValues = sheet.getDataRange().getValues();

## hasOwnProperty
- Occurrences: 11
- Files: Code.gs, Core.gs, cache.gs, database.gs, secretManager.gs, ulog.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:2199: if (config.hasOwnProperty(field)) {
  - Code.gs:2212: if (sanitized.hasOwnProperty('showDetails')) {
  - Code.gs:2213: if (!sanitized.hasOwnProperty('showNames')) {

## getHealth
- Occurrences: 11
- Files: cache.gs, monitoring.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:573: getHealth() {
  - cache.gs:979: const health = cacheManager.getHealth();
  - monitoring.gs:266: cacheHealth: cacheManager?.getHealth() || 'unavailable',

## BATCH_SIZE
- Occurrences: 11
- Files: constants.gs, main.gs
- Samples:
  - constants.gs:70: BATCH_SIZE: {
  - constants.gs:253: Object.freeze(UNIFIED_CONSTANTS.CACHE.BATCH_SIZE);
  - constants.gs:307: /** @deprecated Use UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.LARGE instead */

## Core
- Occurrences: 11
- Files: constants.gs, debugConfig.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - constants.gs:313: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead (defined in Core.gs) */
  - constants.gs:315: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION instead (defined in Core.gs) */
  - constants.gs:1098: // UnifiedErrorHandler は Core.gs で定義済みのため削除

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

## getLastColumn
- Occurrences: 10
- Files: Code.gs, Core.gs, config.gs, unifiedSheetDataManager.gs
- Samples:
  - Code.gs:784: const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];
  - Code.gs:917: const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  - Code.gs:1537: const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

## getScriptLock
- Occurrences: 10
- Files: Code.gs, Core.gs, database.gs, lockManager.gs
- Samples:
  - Code.gs:1363: const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  - Code.gs:1441: const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  - Code.gs:2138: const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;

## releaseLock
- Occurrences: 10
- Files: Code.gs, Core.gs, database.gs, lockManager.gs
- Samples:
  - Code.gs:1415: try { lock.releaseLock(); } catch (e) {}
  - Code.gs:1479: lock.releaseLock();
  - Code.gs:2188: lock.releaseLock();

## WEB_APP_URL
- Occurrences: 10
- Files: Code.gs, cache.gs, unifiedCacheManager.gs, url.gs
- Samples:
  - Code.gs:1608: props.setProperties({ WEB_APP_URL: (url || '').trim() });
  - Code.gs:1630: let stored = (props.getProperty('WEB_APP_URL') || '').trim();
  - Code.gs:1638: props.setProperties({ WEB_APP_URL: converted.trim() });

## create
- Occurrences: 10
- Files: Code.gs, Core.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs, unifiedUtilities.gs
- Samples:
  - Code.gs:1747: const form = FormApp.create(`StudyQuest - 回答フォーム - ${userEmail.split('@')[0]}`);
  - Code.gs:1772: const spreadsheet = SpreadsheetApp.create(`StudyQuest - 回答データ - ${userEmail.split('@')[0]}`);
  - Code.gs:1833: const spreadsheet = SpreadsheetApp.create(`StudyQuest - 回答データ - ${userEmail.split('@')[0]}`);

## cacheError
- Occurrences: 10
- Files: Core.gs, database.gs, main.gs, session-utils.gs
- Samples:
  - Core.gs:2238: console.warn('ハイライト後のキャッシュ無効化エラー:', cacheError.message);
  - Core.gs:2453: console.warn('リアクション後のキャッシュ無効化エラー:', cacheError.message);
  - database.gs:768: console.warn('fetchUserFromDatabase - キャッシュクリア警告:', cacheError.message);

## toFixed
- Occurrences: 10
- Files: cache.gs, database.gs, resilientExecutor.gs, systemIntegrationManager.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:125: const hitRate = (this.stats.hits / this.stats.totalOps * 100).toFixed(1);
  - cache.gs:575: const hitRate = this.stats.totalOps > 0 ? (this.stats.hits / this.stats.totalOps * 100).toFixed(1) :
  - cache.gs:576: const errorRate = this.stats.totalOps > 0 ? (this.stats.errors / this.stats.totalOps * 100).toFixed(

## VALIDATION
- Occurrences: 10
- Files: constants.gs, errorHandler.gs, ulog.gs, workflowValidation.gs
- Samples:
  - constants.gs:21: * @property {string} VALIDATION - バリデーションエラー
  - constants.gs:50: VALIDATION: 'validation',
  - constants.gs:426: VALIDATION: 'validation',

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
  - database.gs:1954: var stats = cacheManager.getStats();
  - database.gs:2831: var cacheStats = cacheManager.getStats();
  - resilientExecutor.gs:237: getStats() {

## alerts
- Occurrences: 10
- Files: database.gs
- Samples:
  - database.gs:2617: monitoringResult.summary.alerts.push('システムヘルスが危険な状態です');
  - database.gs:2623: monitoringResult.summary.alerts.push('ヘルスチェック自体が失敗しました');
  - database.gs:2634: monitoringResult.summary.alerts.push('レスポンス時間が異常に長くなっています');

## DENY
- Occurrences: 10
- Files: main.gs
- Samples:
  - main.gs:150: HtmlService.XFrameOptionsMode.DENY) {
  - main.gs:151: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
  - main.gs:745: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {

## setXFrameOptionsMode
- Occurrences: 10
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:151: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
  - main.gs:725: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  - main.gs:746: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);

## ALLOWALL
- Occurrences: 10
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:724: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
  - main.gs:725: htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  - main.gs:1133: if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {

## infoLog
- Occurrences: 10
- Files: resilientExecutor.gs, setup.gs, systemIntegrationManager.gs, unifiedUserManager.gs
- Samples:
  - resilientExecutor.gs:308: infoLog(`resilientUrlFetch: リトライで成功 ${url}`);
  - setup.gs:30: infoLog('✅ セットアップが正常に完了しました。');
  - systemIntegrationManager.gs:56: infoLog('🚀 統合システム初期化開始');

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

## validate
- Occurrences: 10
- Files: unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - unifiedValidationSystem.gs:35: validate(category, level = this.validationLevels.STANDARD, options = {}) {
  - unifiedValidationSystem.gs:996: return UnifiedValidation.validate(category, level, options);
  - unifiedValidationSystem.gs:1003: return UnifiedValidation.validate('authentication', level, { userId });

## sort
- Occurrences: 9
- Files: Code.gs, Core.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:1215: rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  - Code.gs:1217: rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
  - Code.gs:1219: rows.sort((a, b) => b.score - a.score);

## setValues
- Occurrences: 9
- Files: Code.gs, Core.gs, config.gs, workflowValidation.gs
- Samples:
  - Code.gs:1399: reactionRange.setValues([
  - Code.gs:1696: sheet.getRange(1, 1, 1, all.length).setValues([all]);
  - Code.gs:1787: sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);

## getUuid
- Occurrences: 9
- Files: Code.gs, Core.gs, auth.gs, errorHandler.gs, monitoring.gs, workflowValidation.gs
- Samples:
  - Code.gs:1930: const userId = Utilities.getUuid();
  - Code.gs:2295: const uuid = Utilities.getUuid();
  - Code.gs:2323: const uuid = (typeof Utilities !== 'undefined' && Utilities.getUuid) ? Utilities.getUuid() : Math.ra

## isArray
- Occurrences: 9
- Files: Core.gs, constants.gs, database.gs
- Samples:
  - Core.gs:425: if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
  - Core.gs:3523: if (!Array.isArray(spreadsheet.sheets)) {
  - constants.gs:857: if (!Array.isArray(array) || size <= 0) {

## array
- Occurrences: 9
- Files: Core.gs, constants.gs, unifiedBatchProcessor.gs
- Samples:
  - Core.gs:3663: for (var i = array.length - 1; i > 0; i--) {
  - constants.gs:862: for (let i = 0; i < array.length; i += size) {
  - constants.gs:863: chunks.push(array.slice(i, i + size));

## WARNING
- Occurrences: 9
- Files: Core.gs, secretManager.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs, validationMigration.gs
- Samples:
  - Core.gs:3817: debugLog('mapConfigToActualHeaders: WARNING - No match found for %s = "%s"', configKey, configHeader
  - secretManager.gs:669: results.criticalSecretsStatus = 'WARNING';
  - systemIntegrationManager.gs:251: healthResult.overallStatus = 'WARNING';

## UNKNOWN
- Occurrences: 9
- Files: autoInit.gs, secretManager.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - autoInit.gs:144: diagnosticsStatus: diagnostics?.healthCheck?.overallStatus || 'UNKNOWN',
  - secretManager.gs:589: secretManagerStatus: 'UNKNOWN',
  - secretManager.gs:590: propertiesServiceStatus: 'UNKNOWN',

## clear
- Occurrences: 9
- Files: cache.gs, secretManager.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:523: this.memoCache.clear();
  - cache.gs:537: this.memoCache.clear();
  - secretManager.gs:542: this.secretCache.clear();

## removeAll
- Occurrences: 9
- Files: cache.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:546: this.scriptCache.removeAll();
  - session-utils.gs:38: userCache.removeAll(['config_v3_', 'user_', 'email_', 'hdr_', 'data_', 'sheets_']);
  - session-utils.gs:74: userCache.removeAll([]); // ユーザーキャッシュを全てクリア

## validator
- Occurrences: 9
- Files: constants.gs
- Samples:
  - constants.gs:721: if (validator.type === 'email' && !this.validateEmail(value)) {
  - constants.gs:724: if (validator.type === 'spreadsheetId' && !this.validateSpreadsheetId(value)) {
  - constants.gs:727: if (validator.type === 'userId' && !this.validateUserId(value)) {

## details
- Occurrences: 9
- Files: database.gs
- Samples:
  - database.gs:2273: result.details.duplicates = duplicateResult.duplicates;
  - database.gs:2282: result.details.missingFields = missingFieldsResult.missing;
  - database.gs:2291: result.details.invalidData = invalidDataResult.invalid;

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

## script
- Occurrences: 8
- Files: Code.gs, main.gs, url.gs
- Samples:
  - Code.gs:1620: const base = `https://script.google.com/macros/s/${deployId}/exec`;
  - main.gs:1191: var isValidUrl = cleanUrl.includes('script.google.com') ||
  - main.gs:1218: const fallbackUrl = 'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVq

## addTextItem
- Occurrences: 8
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1753: form.addTextItem()
  - Code.gs:1757: form.addTextItem()
  - Core.gs:2593: var nameItem = form.addTextItem();

## DB
- Occurrences: 8
- Files: Code.gs, Core.gs, constants.gs, database.gs, debugConfig.gs, ulog.gs
- Samples:
  - Code.gs:2251: // DBアクセスは必ずこの関数を経由させる
  - Core.gs:374: // ボードオーナーの情報をDBから取得（キャッシュ利用）
  - Core.gs:440: // ボードオーナーの情報をDBから取得（キャッシュ利用）

## ANONYMOUS
- Occurrences: 8
- Files: Core.gs, constants.gs, main.gs
- Samples:
  - Core.gs:774: var finalDisplayMode = (adminMode === true) ? DISPLAY_MODES.NAMED : (configJson.displayMode || DISPL
  - Core.gs:873: displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
  - Core.gs:911: var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

## JWT
- Occurrences: 8
- Files: auth.gs, unifiedSecurityManager.gs
- Samples:
  - auth.gs:2: * @fileoverview 認証管理 - JWTトークンキャッシュと最適化
  - auth.gs:22: * 新しいJWTトークンを生成
  - auth.gs:51: // JWT生成

## pattern
- Occurrences: 8
- Files: cache.gs, database.gs, errorHandler.gs, unifiedCacheManager.gs, url.gs
- Samples:
  - cache.gs:248: if (!pattern || typeof pattern !== 'string') {
  - cache.gs:254: if (!strict && (pattern.length < 3 || pattern === '*' || pattern === '.*')) {
  - database.gs:279: if (invalidReasonPatterns.some(pattern => pattern.test(reason))) {

## FAILED
- Occurrences: 8
- Files: cache.gs, secretManager.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - cache.gs:560: console.log(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, Script
  - cache.gs:560: console.log(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, Script
  - secretManager.gs:569: if (action.includes('ERROR') || action.includes('FAILED')) {

## LOW
- Occurrences: 8
- Files: constants.gs, errorHandler.gs
- Samples:
  - constants.gs:8: * @property {string} LOW - 軽微なエラー
  - constants.gs:37: LOW: 'low',
  - constants.gs:401: * @param {string} level - レベル（LOW, MEDIUM, HIGH, CRITICAL）

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

## steps
- Occurrences: 8
- Files: database.gs
- Samples:
  - database.gs:56: transactionLog.steps.push('validation_complete');
  - database.gs:70: transactionLog.steps.push('lock_acquired');
  - database.gs:101: transactionLog.steps.push('sheet_created');

## GOOGLE_CLIENT_ID
- Occurrences: 8
- Files: main.gs
- Samples:
  - main.gs:301: console.log('Getting GOOGLE_CLIENT_ID from script properties...');
  - main.gs:303: var clientId = properties.getProperty('GOOGLE_CLIENT_ID');
  - main.gs:305: console.log('GOOGLE_CLIENT_ID retrieved:', clientId ? 'Found' : 'Not found');

## searchParams
- Occurrences: 8
- Files: main.gs, url.gs
- Samples:
  - main.gs:989: url.searchParams.set('mode', mode);
  - main.gs:993: url.searchParams.set(key, params[key]);
  - url.gs:356: url.searchParams.set('_cb', Date.now().toString());

## endsWith
- Occurrences: 8
- Files: main.gs, unifiedUtilities.gs, url.gs
- Samples:
  - main.gs:1162: if ((cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
  - main.gs:1163: (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))) {
  - unifiedUtilities.gs:179: (sanitized.startsWith('"') && sanitized.endsWith('"')) ||

## resolve
- Occurrences: 8
- Files: monitoring.gs, resilientExecutor.gs, unifiedCacheManager.gs
- Samples:
  - monitoring.gs:194: resolve({
  - resilientExecutor.gs:140: Promise.resolve()
  - resilientExecutor.gs:144: resolve(result);

## execute
- Occurrences: 8
- Files: resilientExecutor.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs
- Samples:
  - resilientExecutor.gs:47: async execute(operation, options = {}) {
  - resilientExecutor.gs:349: return resilientExecutor.execute(operation, {
  - resilientExecutor.gs:363: return resilientExecutor.execute(operation, {

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

## addMetaTag
- Occurrences: 7
- Files: Code.gs, main.gs
- Samples:
  - Code.gs:438: .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  - Code.gs:600: .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  - main.gs:1397: .addMetaTag('viewport', 'width=device-width, initial-scale=1');

## splice
- Occurrences: 7
- Files: Code.gs, Core.gs, unifiedCacheManager.gs
- Samples:
  - Code.gs:1158: required.splice(1,0,classHeaderGuess);
  - Code.gs:1286: required.splice(1,0,classHeaderGuess);
  - Code.gs:1394: if (idx > -1) lists[key].arr.splice(idx, 1);

## arr
- Occurrences: 7
- Files: Code.gs
- Samples:
  - Code.gs:1391: const wasReacted = lists[reactionKey].arr.includes(userEmail);
  - Code.gs:1393: const idx = lists[key].arr.indexOf(userEmail);
  - Code.gs:1394: if (idx > -1) lists[key].arr.splice(idx, 1);

## getService
- Occurrences: 7
- Files: Code.gs, main.gs, session-utils.gs, url.gs
- Samples:
  - Code.gs:1646: current = ScriptApp.getService().getUrl();
  - main.gs:837: const baseUrl = ScriptApp.getService().getUrl();
  - session-utils.gs:90: const loginPageUrl = ScriptApp.getService().getUrl();

## addParagraphTextItem
- Occurrences: 7
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1761: form.addParagraphTextItem()
  - Code.gs:1766: form.addParagraphTextItem()
  - Core.gs:2601: var reasonItem = form.addParagraphTextItem();

## setHelpText
- Occurrences: 7
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1763: .setHelpText('質問に対するあなたの回答を入力してください')
  - Code.gs:1768: .setHelpText('その回答を選んだ理由を教えてください')
  - Core.gs:2603: reasonItem.setHelpText(config.reasonQuestion.helpText);

## random
- Occurrences: 7
- Files: Code.gs, Core.gs, resilientExecutor.gs, unifiedUserManager.gs, url.gs
- Samples:
  - Code.gs:2297: const randomString = uuid + timestamp + Math.random().toString(36);
  - Code.gs:2323: const uuid = (typeof Utilities !== 'undefined' && Utilities.getUuid) ? Utilities.getUuid() : Math.ra
  - Core.gs:3634: var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;

## base64EncodeWebSafe
- Occurrences: 7
- Files: Code.gs, auth.gs, unifiedSecurityManager.gs
- Samples:
  - Code.gs:2304: return Utilities.base64EncodeWebSafe(bytes);
  - auth.gs:61: var encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  - auth.gs:62: var encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));

## floor
- Occurrences: 7
- Files: Core.gs, auth.gs, database.gs, resilientExecutor.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:22: remainingMinutes: Math.max(0, Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60)))
  - Core.gs:3664: var j = Math.floor(Math.random() * (i + 1));
  - auth.gs:48: var now = Math.floor(Date.now() / 1000);

## NAMED
- Occurrences: 7
- Files: Core.gs, constants.gs, main.gs
- Samples:
  - Core.gs:774: var finalDisplayMode = (adminMode === true) ? DISPLAY_MODES.NAMED : (configJson.displayMode || DISPL
  - Core.gs:1062: var shouldShowName = (adminMode === true || displayMode === DISPLAY_MODES.NAMED || isOwner);
  - Core.gs:3608: if (nameIndex !== undefined && row[nameIndex] && (displayMode === DISPLAY_MODES.NAMED || isOwner)) {

## getFileById
- Occurrences: 7
- Files: Core.gs
- Samples:
  - Core.gs:1818: var formFile = DriveApp.getFileById(formAndSsInfo.formId);
  - Core.gs:1819: var ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);
  - Core.gs:2940: var file = DriveApp.getFileById(spreadsheetId);

## classQuestion
- Occurrences: 7
- Files: Core.gs
- Samples:
  - Core.gs:2589: classItem.setTitle(config.classQuestion.title);
  - Core.gs:2590: classItem.setChoiceValues(config.classQuestion.choices);
  - Core.gs:2613: if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && 

## addEditor
- Occurrences: 7
- Files: Core.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:3013: file.addEditor(serviceAccountEmail);
  - Core.gs:3185: spreadsheet.addEditor(serviceAccountEmail);
  - Core.gs:3243: file.addEditor(userEmail);

## www
- Occurrences: 7
- Files: Core.gs, auth.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:4009: baseUrl: 'https://www.googleapis.com/drive/v3',
  - auth.gs:46: var tokenUrl = "https://www.googleapis.com/oauth2/v4/token";
  - auth.gs:55: scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",

## HTTP
- Occurrences: 7
- Files: cache.gs, main.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - cache.gs:684: // HTTPステータスチェック
  - cache.gs:686: throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  - main.gs:32: // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理

## chunkArray
- Occurrences: 7
- Files: constants.gs, unifiedBatchProcessor.gs
- Samples:
  - constants.gs:856: chunkArray(array, size = UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.MEDIUM) {
  - constants.gs:935: const chunks = this.chunkArray(items, batchSize);
  - unifiedBatchProcessor.gs:69: const chunkedRanges = this.chunkArray(ranges, this.config.maxBatchSize);

## example
- Occurrences: 7
- Files: constants.gs, main.gs, systemIntegrationManager.gs, unifiedSecurityManager.gs
- Samples:
  - constants.gs:1212: generateUserUrls: () => 'http://example.com',
  - constants.gs:1213: generateUnpublishedUrl: () => 'http://example.com/unpublished'
  - main.gs:1413: template.adminEmail = userInfo.adminEmail || 'admin@example.com';

## vulnerabilities
- Occurrences: 7
- Files: database.gs
- Samples:
  - database.gs:2653: if (securityResult.vulnerabilities.length > 0) {
  - database.gs:2654: monitoringResult.summary.alerts.push(securityResult.vulnerabilities.length + '件のセキュリティ問題が検出されました');
  - database.gs:2868: securityResult.vulnerabilities.push({

## AUTH
- Occurrences: 7
- Files: debugConfig.gs, ulog.gs, unifiedValidationSystem.gs
- Samples:
  - debugConfig.gs:25: AUTH: true, // 認証関連
  - debugConfig.gs:55: * @param {string} category - カテゴリ (CACHE, AUTH, DATABASE, UI, PERFORMANCE)
  - debugConfig.gs:121: category = 'AUTH';

## identifier
- Occurrences: 7
- Files: main.gs, unifiedUserManager.gs
- Samples:
  - main.gs:637: if (typeof identifier === 'object' && identifier !== null) {
  - main.gs:639: email = identifier.email;
  - main.gs:640: userId = identifier.userId;

## rgba
- Occurrences: 7
- Files: main.gs
- Samples:
  - main.gs:1046: background: rgba(31, 41, 55, 0.95);
  - main.gs:1052: box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5);
  - main.gs:1053: border: 1px solid rgba(75, 85, 99, 0.3);

## reject
- Occurrences: 7
- Files: monitoring.gs, resilientExecutor.gs, unifiedCacheManager.gs
- Samples:
  - monitoring.gs:188: reject(new Error(`ヘルスチェックタイムアウト: ${name}`));
  - monitoring.gs:201: reject(error);
  - resilientExecutor.gs:137: reject(new Error(`Operation timed out after ${timeoutMs}ms`));

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

## charCodeAt
- Occurrences: 6
- Files: Code.gs, secretManager.gs
- Samples:
  - Code.gs:222: hash = ((hash << 5) - hash) + str.charCodeAt(i);
  - Code.gs:935: .replace(/[\uFF01-\uFF5E]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
  - secretManager.gs:416: const charCode = value.charCodeAt(i) ^ key.charCodeAt(i % key.length);

## setSandboxMode
- Occurrences: 6
- Files: Code.gs, main.gs
- Samples:
  - Code.gs:471: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  - Code.gs:504: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  - Code.gs:531: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);

## SandboxMode
- Occurrences: 6
- Files: Code.gs, main.gs
- Samples:
  - Code.gs:471: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  - Code.gs:504: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  - Code.gs:531: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);

## setValue
- Occurrences: 6
- Files: Code.gs
- Samples:
  - Code.gs:764: userDb.getRange(i + 1, spreadsheetIdIndex + 1).setValue(spreadsheetId);
  - Code.gs:765: userDb.getRange(i + 1, spreadsheetUrlIndex + 1).setValue(spreadsheetUrl);
  - Code.gs:1471: cell.setValue(newValue);

## reduce
- Occurrences: 6
- Files: Code.gs, Core.gs, unifiedBatchProcessor.gs
- Samples:
  - Code.gs:863: sheetData.rows.reduce((sum, row) => {
  - Code.gs:1403: const reactions = REACTION_KEYS.reduce((obj, k) => {
  - Core.gs:1604: return values.reduce((cnt, row) => row[0] === classFilter ? cnt + 1 : cnt, 0);

## waitLock
- Occurrences: 6
- Files: Code.gs, Core.gs, database.gs, lockManager.gs
- Samples:
  - Code.gs:1444: lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
  - Core.gs:2353: lock.waitLock(10000);
  - database.gs:317: lock.waitLock(15000);

## getPublishedUrl
- Occurrences: 6
- Files: Code.gs, Core.gs, constants.gs
- Samples:
  - Code.gs:1806: formUrl: form.getPublishedUrl(),
  - Core.gs:2561: formUrl: form.getPublishedUrl(),
  - Core.gs:2562: viewFormUrl: form.getPublishedUrl(),

## appendRow
- Occurrences: 6
- Files: Code.gs, config.gs
- Samples:
  - Code.gs:1935: userDb.appendRow([
  - Code.gs:2247: sheet.appendRow(USER_DB_CONFIG.HEADERS);
  - config.gs:10: sheet.appendRow(headers);

## module
- Occurrences: 6
- Files: Code.gs, ErrorHandling.gs, config.gs
- Samples:
  - Code.gs:2428: if (typeof module !== 'undefined') {
  - Code.gs:2431: module.exports = {
  - ErrorHandling.gs:10: if (typeof module !== 'undefined') {

## max
- Occurrences: 6
- Files: Core.gs, database.gs, unifiedCacheManager.gs, unifiedSheetDataManager.gs
- Samples:
  - Core.gs:22: remainingMinutes: Math.max(0, Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60)))
  - Core.gs:1594: return Math.max(0, lastRow - 1);
  - Core.gs:1600: return Math.max(0, lastRow - 1);

## batchOperations
- Occurrences: 6
- Files: Core.gs
- Samples:
  - Core.gs:425: if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
  - Core.gs:431: if (batchOperations.length > MAX_BATCH_SIZE) {
  - Core.gs:435: console.log('🔄 バッチリアクション処理開始:', batchOperations.length + '件');

## hasNext
- Occurrences: 6
- Files: Core.gs
- Samples:
  - Core.gs:1827: while (formParents.hasNext()) {
  - Core.gs:1856: while (ssParents.hasNext()) {
  - Core.gs:2184: if (folders.hasNext()) {

## next
- Occurrences: 6
- Files: Core.gs
- Samples:
  - Core.gs:1828: if (formParents.next().getId() === folder.getId()) {
  - Core.gs:1857: if (ssParents.next().getId() === folder.getId()) {
  - Core.gs:2185: rootFolder = folders.next();

## serviceError
- Occurrences: 6
- Files: Core.gs, database.gs, unifiedSecurityManager.gs
- Samples:
  - Core.gs:3269: console.warn('サービスアカウント追加で警告: ' + serviceError.message);
  - database.gs:752: serviceError.type = 'SERVICE_ERROR';
  - database.gs:1803: error: serviceError.message

## failed
- Occurrences: 6
- Files: auth.gs, cache.gs, resilientExecutor.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs, unifiedSheetDataManager.gs
- Samples:
  - auth.gs:80: console.error('Token request failed. Status:', response.getResponseCode());
  - cache.gs:741: console.error(`[getHeadersWithRetry] All ${maxRetries} attempts failed. Last error:`, lastError.toSt
  - resilientExecutor.gs:102: `[ResilientExecutor] ${operationName} failed (attempt ${attempt + 1}). Retrying in ${delay}ms: ${err

## has
- Occurrences: 6
- Files: cache.gs, constants.gs, database.gs, secretManager.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:48: if (enableMemoization && this.memoCache.has(key)) {
  - constants.gs:1082: has(key) {
  - database.gs:2363: if (userId && seenUserIds.has(userId)) {

## DRY
- Occurrences: 6
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:355: console.log(`🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`);
  - cache.gs:367: console.log(`[Cache] DRY RUN: Would clear pattern: ${pattern}`);
  - cache.gs:387: console.log(`[Cache] DRY RUN: Would clear related pattern: ${pattern}`);

## RUN
- Occurrences: 6
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:355: console.log(`🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`);
  - cache.gs:367: console.log(`[Cache] DRY RUN: Would clear pattern: ${pattern}`);
  - cache.gs:387: console.log(`[Cache] DRY RUN: Would clear related pattern: ${pattern}`);

## _getInvalidationPatterns
- Occurrences: 6
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:358: const patterns = this._getInvalidationPatterns(entityType, entityId);
  - cache.gs:381: const relatedPatterns = this._getInvalidationPatterns(entityType, relatedId);
  - cache.gs:423: _getInvalidationPatterns(entityType, entityId) {

## resetStats
- Occurrences: 6
- Files: cache.gs, resilientExecutor.gs, systemIntegrationManager.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:555: this.resetStats();
  - cache.gs:596: resetStats() {
  - resilientExecutor.gs:251: resetStats() {

## NETWORK
- Occurrences: 6
- Files: constants.gs, errorHandler.gs
- Samples:
  - constants.gs:20: * @property {string} NETWORK - ネットワークエラー
  - constants.gs:49: NETWORK: 'network',
  - constants.gs:425: NETWORK: 'network',

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

## unifiedCacheManager
- Occurrences: 6
- Files: constants.gs, unifiedSecurityManager.gs
- Samples:
  - constants.gs:1242: // UnifiedCacheAPI は unifiedCacheManager.gs で定義済み
  - constants.gs:1245: // CacheManager は unifiedCacheManager.gs で定義済み
  - constants.gs:1246: // UnifiedExecutionCache は unifiedCacheManager.gs で定義済み

## A1
- Occurrences: 6
- Files: database.gs
- Samples:
  - database.gs:105: appendSheetsData(service, dbId, `'${logSheetName}'!A1`, [DELETE_LOG_SHEET_CONFIG.HEADERS]);
  - database.gs:1173: appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);
  - database.gs:1252: // 2. 作成直後にヘッダーを追加（A1記法でレンジを指定）

## getProperties
- Occurrences: 6
- Files: database.gs, main.gs, session-utils.gs, unifiedCacheManager.gs
- Samples:
  - database.gs:2863: var allProps = props.getProperties();
  - main.gs:311: var allProperties = properties.getProperties();
  - main.gs:335: var allProperties = properties.getProperties();

## PERFORMANCE
- Occurrences: 6
- Files: debugConfig.gs, ulog.gs
- Samples:
  - debugConfig.gs:28: PERFORMANCE: true, // パフォーマンス計測
  - debugConfig.gs:55: * @param {string} category - カテゴリ (CACHE, AUTH, DATABASE, UI, PERFORMANCE)
  - debugConfig.gs:135: category = 'PERFORMANCE';

## registerCheck
- Occurrences: 6
- Files: monitoring.gs
- Samples:
  - monitoring.gs:123: registerCheck(name, checkFunction, options = {}) {
  - monitoring.gs:301: healthChecker.registerCheck('database', () => {
  - monitoring.gs:312: healthChecker.registerCheck('sheetsApi', () => {

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

## getUserCache
- Occurrences: 6
- Files: session-utils.gs, unifiedSecurityManager.gs
- Samples:
  - session-utils.gs:16: const userCache = CacheService.getUserCache();
  - session-utils.gs:72: const userCache = CacheService.getUserCache();
  - session-utils.gs:118: const userCache = CacheService.getUserCache();

## MD5
- Occurrences: 6
- Files: session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - session-utils.gs:19: const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, c
  - unifiedBatchProcessor.gs:417: // MD5ハッシュ化でキーを短縮
  - unifiedBatchProcessor.gs:419: Utilities.DigestAlgorithm.MD5,

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

## unifiedUserManager
- Occurrences: 6
- Files: unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - unifiedCacheManager.gs:2348: // → unifiedUserManager.gsの実装を使用してください
  - unifiedSecurityManager.gs:548: * @deprecated 統合実装：unifiedUserManager.getSecureUserInfo() を使用してください
  - unifiedSecurityManager.gs:551: // → unifiedUserManager.gsの実装を使用してください

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

## LIKE_MULTIPLIER_FACTOR
- Occurrences: 5
- Files: Code.gs, Core.gs, main.gs
- Samples:
  - Code.gs:38: LIKE_MULTIPLIER_FACTOR: 0.05 // 1いいね！ごとにスコアが5%増加
  - Code.gs:1186: const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
  - Code.gs:1314: const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);

## IFRAME
- Occurrences: 5
- Files: Code.gs
- Samples:
  - Code.gs:471: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  - Code.gs:504: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  - Code.gs:531: if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);

## setProperties
- Occurrences: 5
- Files: Code.gs
- Samples:
  - Code.gs:492: PropertiesService.getUserProperties().setProperties({
  - Code.gs:1608: props.setProperties({ WEB_APP_URL: (url || '').trim() });
  - Code.gs:1638: props.setProperties({ WEB_APP_URL: converted.trim() });

## insertSheet
- Occurrences: 5
- Files: Code.gs, config.gs
- Samples:
  - Code.gs:1679: const sheet = ss.insertSheet(sheetName);
  - Code.gs:1982: configSheet = spreadsheet.insertSheet('Config');
  - config.gs:8: sheet = getCurrentSpreadsheet().insertSheet('Config');

## computeDigest
- Occurrences: 5
- Files: Code.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - Code.gs:2299: const bytes = Utilities.computeDigest(
  - session-utils.gs:19: const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, c
  - unifiedBatchProcessor.gs:418: const paramsHash = Utilities.computeDigest(

## DigestAlgorithm
- Occurrences: 5
- Files: Code.gs, session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - Code.gs:2300: Utilities.DigestAlgorithm.SHA_256,
  - session-utils.gs:19: const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, c
  - unifiedBatchProcessor.gs:419: Utilities.DigestAlgorithm.MD5,

## updateError
- Occurrences: 5
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:2159: console.error('エラー状態の更新に失敗: ' + updateError.message);
  - Core.gs:4486: console.warn('processLoginFlow: 最終アクセス時刻更新エラー:', updateError.message);
  - database.gs:1090: var errorMessage = updateError.toString();

## invalidateSheetData
- Occurrences: 5
- Files: Core.gs, cache.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:2234: cacheManager.invalidateSheetData(spreadsheetId, sheetName);
  - Core.gs:2449: cacheManager.invalidateSheetData(spreadsheetId, sheetName);
  - cache.gs:198: invalidateSheetData(spreadsheetId, sheetName = null) {

## globalProfiler
- Occurrences: 5
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:3099: var profiler = (typeof globalProfiler !== 'undefined') ? globalProfiler : {
  - main.gs:183: if (typeof globalProfiler !== 'undefined') {
  - main.gs:184: globalProfiler.start('logging');

## accessError
- Occurrences: 5
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:3483: console.warn('getSheetsList: アクセスエラーを検出。サービスアカウント権限を修復中...', accessError.message);
  - database.gs:1824: error: accessError.message
  - database.gs:1826: diagnosticResult.summary.criticalIssues.push('スプレッドシートアクセス失敗: ' + accessError.message);

## AUTHORIZATION
- Occurrences: 5
- Files: constants.gs, errorHandler.gs
- Samples:
  - constants.gs:17: * @property {string} AUTHORIZATION - 認可エラー
  - constants.gs:46: AUTHORIZATION: 'authorization',
  - constants.gs:422: AUTHORIZATION: 'authorization',

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

## pow
- Occurrences: 5
- Files: constants.gs, resilientExecutor.gs, unifiedSheetDataManager.gs
- Samples:
  - constants.gs:823: const waitTime = exponentialBackoff ? delay * Math.pow(2, attempt - 1) : delay;
  - resilientExecutor.gs:159: let delay = this.config.baseDelay * Math.pow(this.config.backoffFactor, attempt);
  - resilientExecutor.gs:321: const delay = baseDelay * Math.pow(2, attempt);

## BATCH
- Occurrences: 5
- Files: constants.gs
- Samples:
  - constants.gs:1121: const batchGet = (keys) => console.log(`[BATCH] Getting ${keys.length} items`);
  - constants.gs:1122: const fallbackBatchGet = (keys) => console.log(`[BATCH] Fallback get for ${keys.length} items`);
  - constants.gs:1123: const batchUpdate = (updates) => console.log(`[BATCH] Updating ${updates.length} items`);

## fn
- Occurrences: 5
- Files: constants.gs, monitoring.gs, unifiedUtilities.gs
- Samples:
  - constants.gs:1196: try { return fn(); }
  - monitoring.gs:192: const result = check.fn();
  - unifiedUtilities.gs:249: const result = fn();

## database
- Occurrences: 5
- Files: constants.gs, errorHandler.gs, unifiedBatchProcessor.gs, unifiedCacheManager.gs
- Samples:
  - constants.gs:1208: // deleteUserAccount関数群は実際の実装がdatabase.gs、Core.gsに存在するため削除
  - errorHandler.gs:113: `database.${operation}`,
  - unifiedBatchProcessor.gs:613: // 同名のグローバル関数が database.gs にも存在しており、Promise を返す実装と同期実装が競合していました。

## unifiedSecurityManager
- Occurrences: 5
- Files: constants.gs, secretManager.gs, unifiedCacheManager.gs
- Samples:
  - constants.gs:1243: // MultiTenantSecurityManager は unifiedSecurityManager.gs で定義済み
  - secretManager.gs:705: // → unifiedSecurityManager.gsの実装を使用してください
  - secretManager.gs:712: // → unifiedSecurityManager.gsの実装を使用してください

## tokenError
- Occurrences: 5
- Files: database.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs
- Samples:
  - database.gs:566: console.error('❌ Failed to get service account token:', tokenError.message);
  - database.gs:567: throw new Error('サービスアカウントトークンの取得に失敗しました: ' + tokenError.message);
  - unifiedBatchProcessor.gs:253: tokenError.message

## parseError
- Occurrences: 5
- Files: database.gs
- Samples:
  - database.gs:1473: console.error('❌ JSON解析エラー:', parseError.message);
  - database.gs:1475: throw new Error('APIレスポンスのJSON解析に失敗: ' + parseError.message);
  - database.gs:1634: console.warn('エラーレスポンスのJSON解析に失敗:', parseError.message);

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

## CURRENT_SPREADSHEET_ID
- Occurrences: 4
- Files: Code.gs
- Samples:
  - Code.gs:66: const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  - Code.gs:494: CURRENT_SPREADSHEET_ID: userInfo.spreadsheetId
  - Code.gs:624: const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');

## getLastRow
- Occurrences: 4
- Files: Code.gs, Core.gs, unifiedSheetDataManager.gs
- Samples:
  - Code.gs:783: if (firstSheet && firstSheet.getLastRow() > 0) {
  - Core.gs:859: var lastRow = sheet.getLastRow(); // スプレッドシートの最終行
  - Core.gs:1592: const lastRow = sheet.getLastRow();

## tryLock
- Occurrences: 4
- Files: Code.gs, Core.gs, database.gs
- Samples:
  - Code.gs:1365: if (lock && !lock.tryLock(TIME_CONSTANTS.LOCK_WAIT_MS)) {
  - Code.gs:2140: if (lock && !lock.tryLock(10000)) {
  - Core.gs:44: if (!lock.tryLock(10000)) {

## concat
- Occurrences: 4
- Files: Code.gs, unifiedBatchProcessor.gs
- Samples:
  - Code.gs:1695: const all = headers.concat(req);
  - unifiedBatchProcessor.gs:100: allValueRanges = allValueRanges.concat(chunkResult.valueRanges || []);
  - unifiedBatchProcessor.gs:307: allReplies = allReplies.concat(chunkResult.replies || []);

## AUTO_CREATE
- Occurrences: 4
- Files: Code.gs
- Samples:
  - Code.gs:1906: if (spreadsheetUrl === 'AUTO_CREATE') {
  - Code.gs:1943: activeSheetName: spreadsheetUrl === 'AUTO_CREATE' ? '回答データ' : '',
  - Code.gs:1959: message: spreadsheetUrl === 'AUTO_CREATE' ?

## getParents
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:1824: var formParents = formFile.getParents();
  - Core.gs:1853: var ssParents = ssFile.getParents();
  - Core.gs:4173: const formParents = formFile.getParents();

## addFile
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:1836: folder.addFile(formFile);
  - Core.gs:1865: folder.addFile(ssFile);
  - Core.gs:4184: folder.addFile(formFile);

## getRootFolder
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:1838: DriveApp.getRootFolder().removeFile(formFile);
  - Core.gs:1867: DriveApp.getRootFolder().removeFile(ssFile);
  - Core.gs:4185: DriveApp.getRootFolder().removeFile(formFile);

## removeFile
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:1838: DriveApp.getRootFolder().removeFile(formFile);
  - Core.gs:1867: DriveApp.getRootFolder().removeFile(ssFile);
  - Core.gs:4185: DriveApp.getRootFolder().removeFile(formFile);

## MM
- Occurrences: 4
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:2530: var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  - Core.gs:2743: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - Core.gs:2858: const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

## formatDate
- Occurrences: 4
- Files: Core.gs, workflowValidation.gs
- Samples:
  - Core.gs:2530: var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  - Core.gs:2743: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - Core.gs:2858: const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

## GoogleAppsScript
- Occurrences: 4
- Files: Core.gs, unifiedCacheManager.gs
- Samples:
  - Core.gs:2577: * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
  - Core.gs:2925: * @param {GoogleAppsScript.Forms.Form} form - フォーム
  - unifiedCacheManager.gs:1283: * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト

## setChoiceValues
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:2590: classItem.setChoiceValues(config.classQuestion.choices);
  - Core.gs:2616: classItem.setChoiceValues(customConfig.classQuestion.choices);
  - Core.gs:2637: mainItem.setChoiceValues(customConfig.mainQuestion.choices);

## reasonQuestion
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:2602: reasonItem.setTitle(config.reasonQuestion.title);
  - Core.gs:2603: reasonItem.setHelpText(config.reasonQuestion.helpText);
  - Core.gs:2674: reasonItem.setTitle(config.reasonQuestion.title);

## instead
- Occurrences: 4
- Files: Core.gs, constants.gs
- Samples:
  - Core.gs:2891: console.warn('createCustomForm() is deprecated. Use createUnifiedForm("custom", ...) instead.');
  - Core.gs:3092: console.warn('createStudyQuestForm() is deprecated. Use createUnifiedForm("study", ...) instead.');
  - constants.gs:313: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead (defined in Core.gs) */

## shareError
- Occurrences: 4
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:2976: console.warn('サービスアカウント共有エラー（処理継続）:', shareError.message);
  - Core.gs:3059: error: shareError.message
  - Core.gs:3062: console.error('共有失敗:', user.adminEmail, shareError.message);

## driveError
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:3015: console.error('DriveApp error:', driveError.message);
  - Core.gs:3016: throw new Error('Drive API操作に失敗しました: ' + driveError.message);
  - Core.gs:3230: console.error('DriveApp.getFileById error:', driveError.message);

## end
- Occurrences: 4
- Files: Core.gs, main.gs, monitoring.gs
- Samples:
  - Core.gs:3156: profiler.end('createForm');
  - main.gs:203: globalProfiler.end('logging');
  - monitoring.gs:144: timer.end();

## rowData
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:3625: var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
  - Core.gs:3628: var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;
  - Core.gs:3628: var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;

## AppSetupPage
- Occurrences: 4
- Files: Core.gs
- Samples:
  - Core.gs:3900: * 現在のユーザーのステータス情報を取得（AppSetupPage.htmlから呼び出される）
  - Core.gs:3905: * isActive状態を更新（AppSetupPage.htmlから呼び出される）
  - Core.gs:4071: * 管理者によるユーザー削除（AppSetupPage.html用ラッパー）

## SA_TOKEN_CACHE
- Occurrences: 4
- Files: auth.gs, cache.gs, unifiedCacheManager.gs, unifiedSecurityManager.gs
- Samples:
  - auth.gs:7: var AUTH_CACHE_KEY = 'SA_TOKEN_CACHE';
  - cache.gs:316: 'SA_TOKEN_CACHE',           // サービスアカウントトークン
  - unifiedCacheManager.gs:474: 'SA_TOKEN_CACHE', // サービスアカウントトークン

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

## valueFn
- Occurrences: 4
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:44: return valueFn();
  - cache.gs:103: newValue = valueFn();
  - unifiedCacheManager.gs:48: return valueFn ? valueFn() : null;

## _isProtectedKey
- Occurrences: 4
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:267: if (this._isProtectedKey(key)) {
  - cache.gs:314: _isProtectedKey(key) {
  - unifiedCacheManager.gs:421: if (this._isProtectedKey(key)) {

## _invalidateCrossEntityCache
- Occurrences: 4
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:397: this._invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog);
  - cache.gs:465: _invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog) {
  - unifiedCacheManager.gs:557: this._invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog);

## relatedId
- Occurrences: 4
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:470: if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {
  - cache.gs:470: if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {
  - unifiedCacheManager.gs:634: if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {

## optimizationOpportunities
- Occurrences: 4
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:1030: analysis.optimizationOpportunities.push('プリウォーミング戦略の導入');
  - cache.gs:1034: analysis.optimizationOpportunities.push('TTL延長によるさらなる高速化');
  - unifiedCacheManager.gs:1952: analysis.optimizationOpportunities.push('プリウォーミング戦略の導入');

## AUTHENTICATION
- Occurrences: 4
- Files: constants.gs, errorHandler.gs
- Samples:
  - constants.gs:16: * @property {string} AUTHENTICATION - 認証エラー
  - constants.gs:45: AUTHENTICATION: 'authentication',
  - constants.gs:421: AUTHENTICATION: 'authentication',

## USER_INPUT
- Occurrences: 4
- Files: constants.gs, errorHandler.gs
- Samples:
  - constants.gs:23: * @property {string} USER_INPUT - ユーザー入力エラー
  - constants.gs:52: USER_INPUT: 'user_input',
  - constants.gs:428: USER_INPUT: 'user_input',

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

## rules
- Occurrences: 4
- Files: constants.gs
- Samples:
  - constants.gs:708: if (rules.required) {
  - constants.gs:709: const requiredResult = this.validateRequired(data, rules.required);
  - constants.gs:716: if (rules.fields) {

## valueRange
- Occurrences: 4
- Files: database.gs
- Samples:
  - database.gs:1497: const hasValues = valueRange.values && valueRange.values.length > 0;
  - database.gs:1497: const hasValues = valueRange.values && valueRange.values.length > 0;
  - database.gs:1498: console.log(`📊 範囲[${index}] ${ranges[index]}: ${hasValues ? valueRange.values.length + '行' : 'データなし

## then
- Occurrences: 4
- Files: lockManager.gs, monitoring.gs, resilientExecutor.gs
- Samples:
  - lockManager.gs:71: return result.then(
  - monitoring.gs:410: .then(result => {
  - resilientExecutor.gs:141: .then(() => operation())

## FATAL
- Occurrences: 4
- Files: monitoring.gs
- Samples:
  - monitoring.gs:16: FATAL: 4
  - monitoring.gs:52: case 'FATAL':
  - monitoring.gs:58: if (level === 'ERROR' || level === 'FATAL') {

## startTimer
- Occurrences: 4
- Files: monitoring.gs, ulog.gs
- Samples:
  - monitoring.gs:87: startTimer(operationName) {
  - monitoring.gs:142: const timer = logger.startTimer(`HealthCheck: ${name}`);
  - monitoring.gs:360: const timer = logger.startTimer('MonitoringDashboard');

## check
- Occurrences: 4
- Files: monitoring.gs
- Samples:
  - monitoring.gs:148: if (result.status === 'unhealthy' && check.critical) {
  - monitoring.gs:162: if (check.critical) {
  - monitoring.gs:189: }, check.timeout);

## setTimeout
- Occurrences: 4
- Files: monitoring.gs, resilientExecutor.gs, unifiedBatchProcessor.gs
- Samples:
  - monitoring.gs:187: const timeout = setTimeout(() => {
  - resilientExecutor.gs:136: const timeoutId = setTimeout(() => {
  - resilientExecutor.gs:230: return new Promise((resolve) => setTimeout(resolve, ms));

## globalContextManager
- Occurrences: 4
- Files: monitoring.gs
- Samples:
  - monitoring.gs:345: const contextCount = globalContextManager.activeContexts.size;
  - monitoring.gs:346: return contextCount < globalContextManager.maxConcurrentContexts;
  - monitoring.gs:372: activeContexts: globalContextManager.activeContexts.size,

## WEBHOOK_SECRET
- Occurrences: 4
- Files: secretManager.gs
- Samples:
  - secretManager.gs:32: WEBHOOK_SECRET: 'webhook_secret',
  - secretManager.gs:40: WEBHOOK_SECRET: this.secretTypes.WEBHOOK_SECRET,
  - secretManager.gs:40: WEBHOOK_SECRET: this.secretTypes.WEBHOOK_SECRET,

## POST
- Occurrences: 4
- Files: secretManager.gs, unifiedBatchProcessor.gs
- Samples:
  - secretManager.gs:286: method: 'POST',
  - secretManager.gs:311: method: 'POST',
  - unifiedBatchProcessor.gs:169: method: 'POST',

## UTF_8
- Occurrences: 4
- Files: session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - session-utils.gs:19: const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, c
  - unifiedBatchProcessor.gs:421: Utilities.Charset.UTF_8
  - unifiedSecurityManager.gs:382: Utilities.Charset.UTF_8

## Charset
- Occurrences: 4
- Files: session-utils.gs, unifiedBatchProcessor.gs, unifiedSecurityManager.gs, unifiedUserManager.gs
- Samples:
  - session-utils.gs:19: const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, c
  - unifiedBatchProcessor.gs:421: Utilities.Charset.UTF_8
  - unifiedSecurityManager.gs:382: Utilities.Charset.UTF_8

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

## clearUserInfoCache
- Occurrences: 4
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1996: clearUserInfoCache(identifier = null) {
  - unifiedCacheManager.gs:2146: this.clearUserInfoCache(userId || email);
  - unifiedCacheManager.gs:2181: this.clearUserInfoCache();

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

## LOCK_WAIT_MS
- Occurrences: 3
- Files: Code.gs
- Samples:
  - Code.gs:42: LOCK_WAIT_MS: 10000
  - Code.gs:1365: if (lock && !lock.tryLock(TIME_CONSTANTS.LOCK_WAIT_MS)) {
  - Code.gs:1444: lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);

## IS_PUBLISHED
- Occurrences: 3
- Files: Code.gs
- Samples:
  - Code.gs:194: props.setProperty('IS_PUBLISHED', 'true');
  - Code.gs:204: props.setProperty('IS_PUBLISHED', 'false');
  - Code.gs:259: const isPublished = props.getProperty('IS_PUBLISHED') === 'true';

## DEPLOY_ID
- Occurrences: 3
- Files: Code.gs
- Samples:
  - Code.gs:266: deployId: props.getProperty('DEPLOY_ID'),
  - Code.gs:1629: const deployId = props.getProperty('DEPLOY_ID');
  - Code.gs:1667: props.setProperties({ DEPLOY_ID: (id || '').trim() });

## assign
- Occurrences: 3
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:580: Object.assign(template, {
  - Code.gs:2169: const newConfig = Object.assign({}, currentConfig, sanitizedConfig);
  - Core.gs:2727: Object.assign(config[key], customConfig[key]);

## setDescription
- Occurrences: 3
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1750: form.setDescription('StudyQuestで使用する回答フォームです。質問に対する回答を入力してください。');
  - Core.gs:1680: form.setDescription(description);
  - Core.gs:2537: form.setDescription(formDescription);

## setFontWeight
- Occurrences: 3
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1791: allHeadersRange.setFontWeight('bold');
  - Code.gs:1856: headerRange.setFontWeight('bold');
  - Core.gs:3314: allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

## E3F2FD
- Occurrences: 3
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1792: allHeadersRange.setBackground('#E3F2FD');
  - Code.gs:1857: headerRange.setBackground('#E3F2FD');
  - Core.gs:3314: allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

## setBackground
- Occurrences: 3
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1792: allHeadersRange.setBackground('#E3F2FD');
  - Code.gs:1857: headerRange.setBackground('#E3F2FD');
  - Core.gs:3314: allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

## autoResizeColumns
- Occurrences: 3
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1796: sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
  - Code.gs:1861: sheet.autoResizeColumns(1, headers.length);
  - Core.gs:3318: sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());

## getEditUrl
- Occurrences: 3
- Files: Code.gs, Core.gs, constants.gs
- Samples:
  - Code.gs:1809: editFormUrl: form.getEditUrl()
  - Core.gs:2563: editFormUrl: typeof form.getEditUrl === 'function' ? form.getEditUrl() : '',
  - constants.gs:1292: getEditUrl: (form) => form ? form.getEditUrl() : ''

## DATABASE_ID
- Occurrences: 3
- Files: Code.gs
- Samples:
  - Code.gs:2239: if (props.getProperty('DATABASE_ID')) {
  - Code.gs:2244: props.setProperty('DATABASE_ID', db.getId());
  - Code.gs:2253: const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_ID');

## OR
- Occurrences: 3
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:101: // Step 1: データソース未設定 OR セットアップ初期状態
  - Core.gs:107: // Step 2: データソース設定済み + (再設定中 OR セットアップ未完了)
  - database.gs:420: // 本人削除 OR 管理者削除

## add
- Occurrences: 3
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:489: processedRows.add(operation.rowIndex);
  - database.gs:2371: seenUserIds.add(userId);
  - database.gs:2383: seenEmails.add(email);

## dbError
- Occurrences: 3
- Files: Core.gs, database.gs, setup.gs
- Samples:
  - Core.gs:1422: details: { error: dbError.message }
  - database.gs:2725: error: dbError.message
  - setup.gs:88: details: { error: dbError.message },

## main
- Occurrences: 3
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:1436: // include 関数は main.gs で定義されています
  - main.gs:1317: console.error('❌ main.gs: publishedSheetNameが不正な型です:', typeof config.publishedSheetName, config.publ
  - main.gs:1318: console.warn('🔧 main.gs: publishedSheetNameを空文字にリセットしました');

## formMoveError
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:1846: moveErrors.push('フォームファイル移動エラー: ' + formMoveError.message);
  - Core.gs:1847: console.error('❌ フォームファイルの移動に失敗:', formMoveError.message);
  - Core.gs:4191: console.warn('カスタムフォームファイル移動エラー:', formMoveError.message);

## ssMoveError
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:1875: moveErrors.push('スプレッドシートファイル移動エラー: ' + ssMoveError.message);
  - Core.gs:1876: console.error('❌ スプレッドシートファイルの移動に失敗:', ssMoveError.message);
  - Core.gs:4217: console.warn('カスタムスプレッドシートファイル移動エラー:', ssMoveError.message);

## HH
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:2530: var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  - Core.gs:2743: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - Core.gs:2858: const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

## nameQuestion
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:2594: nameItem.setTitle(config.nameQuestion.title);
  - Core.gs:2679: nameItem.setTitle(config.nameQuestion.title);
  - Core.gs:2680: nameItem.setHelpText(config.nameQuestion.helpText);

## choices
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:2613: if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && 
  - Core.gs:2636: if (customConfig.mainQuestion && customConfig.mainQuestion.choices && customConfig.mainQuestion.choi
  - Core.gs:2645: if (customConfig.mainQuestion && customConfig.mainQuestion.choices && customConfig.mainQuestion.choi

## getScriptTimeZone
- Occurrences: 3
- Files: Core.gs, monitoring.gs, workflowValidation.gs
- Samples:
  - Core.gs:2743: const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  - monitoring.gs:259: timeZone: Session.getScriptTimeZone(),
  - workflowValidation.gs:147: const formatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

## fallbackError
- Occurrences: 3
- Files: Core.gs, resilientExecutor.gs, session-utils.gs
- Samples:
  - Core.gs:4848: console.warn('⚠️ フォールバックシート名検索に失敗:', fallbackError.message);
  - resilientExecutor.gs:120: `[ResilientExecutor] Fallback failed for ${operationName}: ${fallbackError.message}`
  - session-utils.gs:287: console.error('❌ Fallback HTML creation failed:', fallbackError.message);

## _meta
- Occurrences: 3
- Files: Core.gs
- Samples:
  - Core.gs:4864: response._meta.includedApis.push('getSheetDetails');
  - Core.gs:4876: response._meta.executionTime = endTime - startTime;
  - Core.gs:4879: executionTime: response._meta.executionTime + 'ms',

## checkApplicationAccess
- Occurrences: 3
- Files: Core.gs, main.gs, unifiedValidationSystem.gs
- Samples:
  - Core.gs:4956: const accessCheck = checkApplicationAccess();
  - main.gs:422: const accessCheck = checkApplicationAccess();
  - unifiedValidationSystem.gs:431: const result = checkApplicationAccess(); // 既存関数利用

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

## valuesFn
- Occurrences: 3
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:156: const newValues = valuesFn(missingKeys);
  - unifiedCacheManager.gs:266: const newValues = Promise.resolve(valuesFn(missingKeys));
  - unifiedCacheManager.gs:306: const newValues = valuesFn(missingKeys);

## invalidateRelated
- Occurrences: 3
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:332: invalidateRelated(entityType, entityId, relatedIds = [], options = {}) {
  - unifiedCacheManager.gs:490: invalidateRelated(entityType, entityId, relatedIds = [], options = {}) {
  - unifiedCacheManager.gs:2142: this.manager.invalidateRelated('spreadsheet', dbSpreadsheetId, [userId]);

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

## spreadsheets
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:532: hasGet: !!(service.spreadsheets && typeof service.spreadsheets.get === 'function'),
  - database.gs:537: if (!service.spreadsheets || typeof service.spreadsheets.get !== 'function') {
  - database.gs:540: getType: service.spreadsheets ? typeof service.spreadsheets.get : 'no spreadsheets'

## dataError
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:777: dataError.type = 'DATA_ACCESS_ERROR';
  - database.gs:1870: error: dataError.message
  - database.gs:1872: diagnosticResult.summary.criticalIssues.push('ユーザーデータ取得失敗: ' + dataError.message);

## RAW
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:1531: valueInputOption: 'RAW'
  - database.gs:1559: ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';
  - database.gs:1728: '?valueInputOption=RAW';

## spreadsheetAccess
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:2040: result.checks.spreadsheetAccess.canWrite = true;
  - database.gs:2043: result.checks.spreadsheetAccess.canWrite = false;
  - database.gs:2044: result.checks.spreadsheetAccess.writeError = writeError.message;

## perfError
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:2642: console.error('❌ パフォーマンスチェックエラー:', perfError.message);
  - database.gs:2840: console.error('パフォーマンステスト実行エラー:', perfError.message);
  - database.gs:2841: perfResult.error = perfError.message;

## securityError
- Occurrences: 3
- Files: database.gs
- Samples:
  - database.gs:2657: console.error('❌ セキュリティチェックエラー:', securityError.message);
  - database.gs:2912: console.error('セキュリティチェック実行エラー:', securityError.message);
  - database.gs:2913: securityResult.error = securityError.message;

## top
- Occurrences: 3
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:1009: return HtmlService.createHtmlOutput().setContent(`<script>window.top.location.href = '${sanitizedUrl
  - session-utils.gs:200: // 最も確実な方法：window.top.location.href
  - session-utils.gs:201: window.top.location.href = '${safeLoginUrl}';

## ERROR_LOGS
- Occurrences: 3
- Files: monitoring.gs
- Samples:
  - monitoring.gs:70: const errorLogs = JSON.parse(props.getProperty('ERROR_LOGS') || '[]');
  - monitoring.gs:78: props.setProperty('ERROR_LOGS', JSON.stringify(errorLogs));
  - monitoring.gs:397: const errorLogs = JSON.parse(props.getProperty('ERROR_LOGS') || '[]');

## round
- Occurrences: 3
- Files: monitoring.gs, unifiedValidationSystem.gs, validationMigration.gs
- Samples:
  - monitoring.gs:234: healthPercentage: Math.round((healthy / total) * 100)
  - unifiedValidationSystem.gs:875: passRate: result.summary.total > 0 ? Math.round((result.summary.passed / result.summary.total) * 100
  - validationMigration.gs:493: migrationProgress: Math.round(((stats.total - stats.deprecated) / stats.total) * 100)

## SERVICE_ACCOUNT
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:29: SERVICE_ACCOUNT: 'service_account',
  - secretManager.gs:38: SERVICE_ACCOUNT_CREDS: this.secretTypes.SERVICE_ACCOUNT,
  - secretManager.gs:478: case this.secretTypes.SERVICE_ACCOUNT:

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

## getResilientScriptProperties
- Occurrences: 3
- Files: secretManager.gs
- Samples:
  - secretManager.gs:345: const props = getResilientScriptProperties();
  - secretManager.gs:390: const props = getResilientScriptProperties();
  - secretManager.gs:620: const props = getResilientScriptProperties();

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

## CURRENT_USER_ID_
- Occurrences: 3
- Files: session-utils.gs, unifiedUserManager.gs
- Samples:
  - session-utils.gs:19: const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, c
  - session-utils.gs:29: if (key.startsWith('CURRENT_USER_ID_') && key !== currentUserKey) {
  - unifiedUserManager.gs:235: 'CURRENT_USER_ID_' +

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

## spreadsheetCache
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:3: * cache.gs、unifiedExecutionCache.gs、spreadsheetCache.gs の機能を統合
  - unifiedCacheManager.gs:1265: // SECTION 3: SpreadsheetApp最適化システム（元spreadsheetCache.gs）
  - unifiedCacheManager.gs:2234: // メモリキャッシュから削除（spreadsheetCache.gsの機能）

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

## clearUserInfo
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1190: clearUserInfo() {
  - unifiedCacheManager.gs:1208: this.clearUserInfo();
  - unifiedCacheManager.gs:1999: this.executionCache.clearUserInfo();

## syncWithUnifiedCache
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1239: syncWithUnifiedCache(operation) {
  - unifiedCacheManager.gs:2000: this.executionCache.syncWithUnifiedCache('userDataChange');
  - unifiedCacheManager.gs:2037: this.executionCache.syncWithUnifiedCache('systemChange');

## SESSION_CACHE_TTL
- Occurrences: 3
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1274: SESSION_CACHE_TTL: 1800000, // 30分間（PropertiesServiceキャッシュ）
  - unifiedCacheManager.gs:1321: if (now - sessionEntry.timestamp < SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL) {
  - unifiedCacheManager.gs:1487: sessionTTL: SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL,

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

## ADMIN_EMAILS_
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:123: const str = props.getProperty(`ADMIN_EMAILS_${spreadsheetId}`);
  - Code.gs:2393: const key = `ADMIN_EMAILS_${spreadsheetId}`;

## PUBLISHED_SHEET_NAME
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:195: props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
  - Code.gs:205: if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');

## UNPUBLISHED_ACCESS
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:541: auditLog('UNPUBLISHED_ACCESS', validatedUserId, { viewerEmail, isOwner: false });
  - Code.gs:557: auditLog('UNPUBLISHED_ACCESS', validatedUserId, { viewerEmail, isOwner: false });

## ASCII
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:926: * trimming/collapsing whitespace and converting full-width ASCII before
  - Code.gs:933: // Normalize header strings: convert full-width ASCII, collapse whitespace and lowercase

## saveError
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:1129: console.warn('Failed to save auto-created config in getSheetData:', saveError.message);
  - Code.gs:1257: console.warn('Failed to save auto-created config:', saveError.message);

## SPREADSHEET
- Occurrences: 2
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1773: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  - Core.gs:2957: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

## setDestination
- Occurrences: 2
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1773: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  - Core.gs:2957: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

## DestinationType
- Occurrences: 2
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1773: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  - Core.gs:2957: form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

## getNumColumns
- Occurrences: 2
- Files: Code.gs, Core.gs
- Samples:
  - Code.gs:1796: sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
  - Core.gs:3318: sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());

## getActiveSheet
- Occurrences: 2
- Files: Code.gs, workflowValidation.gs
- Samples:
  - Code.gs:1834: const sheet = spreadsheet.getActiveSheet();
  - workflowValidation.gs:69: const sheet = SpreadsheetApp.getActiveSheet();

## setName
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:1835: sheet.setName('回答データ');
  - Code.gs:2246: sheet.setName(USER_DB_CONFIG.SHEET_NAME);

## available
- Occurrences: 2
- Files: Code.gs, database.gs
- Samples:
  - Code.gs:2414: * Return existing board info for the current user if available.
  - database.gs:1200: * Polls the database until a user record becomes available.

## require
- Occurrences: 2
- Files: Code.gs
- Samples:
  - Code.gs:2429: handleError = require('./ErrorHandling.gs').handleError;
  - Code.gs:2430: const { getConfig, saveSheetConfig, createConfigSheet } = require('./config.gs');

## JP
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:20: publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
  - Core.gs:21: stopTimeFormatted: stopTime.toLocaleString('ja-JP'),

## toLocaleString
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:20: publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
  - Core.gs:21: stopTimeFormatted: stopTime.toLocaleString('ja-JP'),

## column
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:193: validationResults.issues.push('Reason column (理由) not found in headers');
  - Core.gs:199: validationResults.issues.push('Opinion column (回答) not found in headers');

## operationError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:501: console.error('個別操作エラー:', operation, operationError.message);
  - Core.gs:506: message: operationError.message

## updateErr
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:1173: console.warn('Config auto-heal failed: ' + updateErr.message);
  - Core.gs:4725: console.warn('Config auto-heal failed: ' + updateErr.message);

## countError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:1193: console.warn('回答数の取得に失敗: ' + countError.message);
  - Core.gs:1539: console.warn('回答数の取得に失敗: ' + countError.message);

## displayOptions
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:1297: configJson.displayMode = options.displayOptions.showNames ? 'named' : 'anonymous';
  - Core.gs:1298: configJson.showCounts = options.displayOptions.showCounts;

## getEffectiveUser
- Occurrences: 2
- Files: Core.gs, setup.gs
- Samples:
  - Core.gs:1373: var adminEmail = Session.getEffectiveUser().getEmail();
  - setup.gs:22: var adminEmail = Session.getEffectiveUser().getEmail();

## formError
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:1688: console.error('フォーム更新エラー: ' + formError.message);
  - Core.gs:1689: return { status: 'error', message: 'フォームの更新に失敗しました: ' + formError.message };

## getFoldersByName
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2183: var folders = DriveApp.getFoldersByName(rootFolderName);
  - Core.gs:2191: var userFolders = rootFolder.getFoldersByName(userFolderName);

## createFolder
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2187: rootFolder = DriveApp.createFolder(rootFolderName);
  - Core.gs:2195: return rootFolder.createFolder(userFolderName);

## rootFolder
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2191: var userFolders = rootFolder.getFoldersByName(userFolderName);
  - Core.gs:2195: return rootFolder.createFolder(userFolderName);

## VERIFIED
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2545: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
  - Core.gs:3125: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);

## setEmailCollectionType
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2545: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
  - Core.gs:3125: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);

## EmailCollectionType
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2545: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
  - Core.gs:3125: form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);

## setCollectEmail
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2547: form.setCollectEmail(true);
  - Core.gs:2585: form.setCollectEmail(false);

## Forms
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2577: * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
  - Core.gs:2925: * @param {GoogleAppsScript.Forms.Form} form - フォーム

## addListItem
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2588: var classItem = form.addListItem();
  - Core.gs:2614: var classItem = form.addListItem();

## showOtherOption
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2640: mainItem.showOtherOption(true);
  - Core.gs:2649: mainItem.showOtherOption(true);

## createTextOutput
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2750: return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(ContentService.MimeType.JSON
  - Core.gs:2753: return ContentService.createTextOutput(JSON.stringify({

## setMimeType
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2750: return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(ContentService.MimeType.JSON
  - Core.gs:2756: })).setMimeType(ContentService.MimeType.JSON);

## MimeType
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2750: return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(ContentService.MimeType.JSON
  - Core.gs:2756: })).setMimeType(ContentService.MimeType.JSON);

## deprecated
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2891: console.warn('createCustomForm() is deprecated. Use createUnifiedForm("custom", ...) instead.');
  - Core.gs:3092: console.warn('createStudyQuestForm() is deprecated. Use createUnifiedForm("study", ...) instead.');

## setSharing
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2946: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  - Core.gs:3236: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## Access
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2946: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  - Core.gs:3236: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## Permission
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:2946: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  - Core.gs:3236: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## status
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:3334: * @returns {object} status ('success' or 'error') と message
  - main.gs:377: * Retrieves the administrator domain for the login page with domain match status.

## repairError
- Occurrences: 2
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:3495: console.error('getSheetsList: 権限修復に失敗:', repairError.message);
  - database.gs:2096: result.summary.actions.push('権限修復失敗: ' + repairError.message);

## RANDOM_SCORE_FACTOR
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:3634: var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
  - main.gs:107: RANDOM_SCORE_FACTOR: 0.01

## actualHeader
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:3732: var normalizedActualHeader = actualHeader.toLowerCase().trim();
  - Core.gs:3743: var normalizedActualHeader = actualHeader.toLowerCase().trim();

## Code
- Occurrences: 2
- Files: Core.gs, config.gs
- Samples:
  - Core.gs:3878: // 追加のコアファンクション（Code.gsから移行）
  - config.gs:73: // ヘッダーを推測（Code.gsの関数を使用）

## folderError
- Occurrences: 2
- Files: Core.gs, database.gs
- Samples:
  - Core.gs:4223: console.warn('カスタムセットアップフォルダ処理エラー:', folderError.message);
  - database.gs:1187: console.warn('createUser - フォルダ作成でエラーが発生しましたが、処理を続行します: ' + folderError.message);

## err
- Occurrences: 2
- Files: Core.gs, main.gs
- Samples:
  - Core.gs:4745: console.warn('Answer count retrieval failed:', err.message);
  - main.gs:1354: try { addServiceAccountToSpreadsheet(userInfo.spreadsheetId); } catch (err) { console.warn('アクセス権設定警

## createExecutionContext
- Occurrences: 2
- Files: Core.gs, unifiedUtilities.gs
- Samples:
  - Core.gs:4858: const context = createExecutionContext(currentUserId, {
  - unifiedUtilities.gs:144: execution: (requestUserId, options = {}) => createExecutionContext(requestUserId, options),

## sheetErr
- Occurrences: 2
- Files: Core.gs
- Samples:
  - Core.gs:4869: console.warn('Sheet details retrieval failed:', sheetErr.message);
  - Core.gs:4870: response.sheetDetailsError = sheetErr.message;

## private_key
- Occurrences: 2
- Files: auth.gs, unifiedSecurityManager.gs
- Samples:
  - auth.gs:44: var privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化
  - unifiedSecurityManager.gs:34: const privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化

## RS256
- Occurrences: 2
- Files: auth.gs, unifiedSecurityManager.gs
- Samples:
  - auth.gs:52: var jwtHeader = { alg: "RS256", typ: "JWT" };
  - unifiedSecurityManager.gs:42: const jwtHeader = { alg: 'RS256', typ: 'JWT' };

## computeRsaSha256Signature
- Occurrences: 2
- Files: auth.gs, unifiedSecurityManager.gs
- Samples:
  - auth.gs:64: var signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  - unifiedSecurityManager.gs:54: const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);

## configJsonString
- Occurrences: 2
- Files: auth.gs, unifiedSecurityManager.gs
- Samples:
  - auth.gs:350: if (!configJsonString || configJsonString.trim() === '' || configJsonString === '{}') {
  - unifiedSecurityManager.gs:245: if (!configJsonString || configJsonString.trim() === '' || configJsonString === '{}') {

## removeError
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:92: console.warn(`[Cache] Failed to remove corrupted cache entry: ${key}`, removeError.message);
  - unifiedCacheManager.gs:220: ULog.warn(`[Cache] Failed to clean corrupted entry: ${key}`, removeError.message);

## rate
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:126: debugLog(`[Cache] Performance: ${hitRate}% hit rate (${this.stats.hits}/${this.stats.totalOps}), ${t
  - unifiedCacheManager.gs:94: `[Cache] Performance: ${hitRate}% hit rate (${this.stats.hits}/${this.stats.totalOps}), ${this.stats

## getAll
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:145: const cachedValues = this.scriptCache.getAll(keys);
  - unifiedCacheManager.gs:295: const cachedValues = this.scriptCache.getAll(keys);

## putAll
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:162: this.scriptCache.putAll(newCacheValues, ttl);
  - unifiedCacheManager.gs:312: this.scriptCache.putAll(newCacheValues, ttl);

## override
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:255: console.warn(`[Cache] Pattern too broad for safe removal: ${pattern}. Use strict=true to override.`)
  - unifiedCacheManager.gs:408: `[Cache] Pattern too broad for safe removal: ${pattern}. Use strict=true to override.`

## limit
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:277: console.warn(`[Cache] Reached maxKeys limit (${maxKeys}) for pattern: ${pattern}`);
  - unifiedCacheManager.gs:431: ULog.warn(`[Cache] Reached maxKeys limit (${maxKeys}) for pattern: ${pattern}`);

## protected
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:305: debugLog(`[Cache] Pattern clear completed: ${successCount} removed, ${failedRemovals} failed, ${skip
  - unifiedCacheManager.gs:462: `[Cache] Pattern clear completed: ${successCount} removed, ${failedRemovals} failed, ${skippedCount}

## SYSTEM_CONFIG
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:318: 'SYSTEM_CONFIG',           // システム設定
  - unifiedCacheManager.gs:476: 'SYSTEM_CONFIG', // システム設定

## DOMAIN_INFO
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:319: 'DOMAIN_INFO'              // ドメイン情報
  - unifiedCacheManager.gs:477: 'DOMAIN_INFO', // ドメイン情報

## IDs
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:351: console.warn(`[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`);
  - unifiedCacheManager.gs:509: ULog.warn(`[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`);

## LIVE
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:355: console.log(`🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`);
  - unifiedCacheManager.gs:514: `🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`

## clearExpired
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:522: clearExpired() {
  - unifiedCacheManager.gs:693: clearExpired() {

## headerName
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:713: if (headerName && headerName.trim() !== '' && headerName !== 'タイムスタンプ') {
  - unifiedCacheManager.gs:1664: if (headerName && headerName.trim() !== '' && headerName !== 'タイムスタンプ') {

## retry
- Occurrences: 2
- Files: cache.gs, unifiedCacheManager.gs
- Samples:
  - cache.gs:734: console.log(`[getHeadersWithRetry] Waiting ${waitTime}ms before retry...`);
  - unifiedCacheManager.gs:1688: ULog.debug(`[getHeadersWithRetry] Waiting ${waitTime}ms before retry...`);

## resetMemoizationCache
- Occurrences: 2
- Files: cache.gs
- Samples:
  - cache.gs:836: if (typeof resetMemoizationCache === 'function') {
  - cache.gs:837: resetMemoizationCache();

## getUi
- Occurrences: 2
- Files: config.gs
- Samples:
  - config.gs:167: SpreadsheetApp.getUi().alert(`${created.join('と')}シートを作成しました。設定を入力してください。`);
  - config.gs:169: SpreadsheetApp.getUi().alert('Configシートは既に存在します。');

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

## DEFAULT_MAIN_QUESTION
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:166: DEFAULT_MAIN_QUESTION: 'あなたの考えや気づいたことを教えてください',
  - constants.gs:313: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead (defined in Core.gs) */

## DEFAULT_REASON_QUESTION
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:167: DEFAULT_REASON_QUESTION: 'そう考える理由や体験があれば教えてください（任意）',
  - constants.gs:315: /** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION instead (defined in Core.gs) */

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

## ACCESS
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:518: ACCESS: 'access',
  - constants.gs:1299: ACCESS: 'access',

## USER
- Occurrences: 2
- Files: constants.gs
- Samples:
  - constants.gs:521: USER: 'user',
  - constants.gs:1302: USER: 'user',

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

## sheetError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:115: throw new Error(`ログシートの準備に失敗: ${sheetError.message}`);
  - database.gs:487: console.warn('ログシートの読み取りに失敗:', sheetError.message);

## ROWS
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:358: dimension: 'ROWS',
  - database.gs:3081: dimension: 'ROWS',

## CONFIG_ERROR
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:743: configError.type = 'CONFIG_ERROR';
  - database.gs:874: if (error.type === 'CONFIG_ERROR' || error.type === 'FIELD_ERROR') {

## FIELD_ERROR
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:800: fieldError.type = 'FIELD_ERROR';
  - database.gs:874: if (error.type === 'CONFIG_ERROR' || error.type === 'FIELD_ERROR') {

## TRUE
- Occurrences: 2
- Files: database.gs, main.gs
- Samples:
  - database.gs:836: if (user.isActive === true || user.isActive === 'true' || user.isActive === 'TRUE') {
  - main.gs:116: * Accepts boolean true, 'true', or 'TRUE'.

## min
- Occurrences: 2
- Files: database.gs, resilientExecutor.gs
- Samples:
  - database.gs:918: var waitTime = Math.min(1000 * retryAttempt, 3000); // 最大3秒
  - resilientExecutor.gs:160: delay = Math.min(delay, this.config.maxDelay);

## milliseconds
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:1202: * @param {number} maxWaitMs - Maximum wait time in milliseconds.
  - database.gs:1203: * @param {number} intervalMs - Poll interval in milliseconds.

## propsError
- Occurrences: 2
- Files: database.gs, unifiedCacheManager.gs
- Samples:
  - database.gs:1293: console.warn('Failed to clear user properties:', propsError.message);
  - unifiedCacheManager.gs:2244: ULog.debug('PropertiesService削除エラー:', propsError.message);

## saError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2016: error: saError.message
  - database.gs:2018: result.summary.issues.push('サービスアカウント設定エラー: ' + saError.message);

## writeError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2044: result.checks.spreadsheetAccess.writeError = writeError.message;
  - database.gs:2045: result.summary.issues.push('スプレッドシート書き込み権限不足: ' + writeError.message);

## retestError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2088: error: retestError.message,
  - database.gs:2091: result.summary.actions.push('修復後テスト失敗: ' + retestError.message);

## permError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2176: repairResult.actions.push('サービスアカウント権限確認失敗: ' + permError.message);
  - database.gs:2746: error: permError.message

## fixError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2311: console.error('❌ 自動修復エラー:', fixError.message);
  - database.gs:2312: result.summary.issues.push('自動修復に失敗: ' + fixError.message);

## healthError
- Occurrences: 2
- Files: database.gs, systemIntegrationManager.gs
- Samples:
  - database.gs:2622: console.error('❌ ヘルスチェックエラー:', healthError.message);
  - systemIntegrationManager.gs:88: initResult.warnings.push(`初期ヘルスチェック失敗: ${healthError.message}`);

## alertError
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2674: console.error('❌ アラート送信エラー:', alertError.message);
  - database.gs:2981: console.error('❌ アラート送信エラー:', alertError.message);

## benchmarks
- Occurrences: 2
- Files: database.gs
- Samples:
  - database.gs:2823: perfResult.benchmarks.databaseAccess = Date.now() - dbTestStart;
  - database.gs:2837: perfResult.benchmarks.cacheCheck = Date.now() - cacheTestStart;

## logLevels
- Occurrences: 2
- Files: debugConfig.gs, monitoring.gs
- Samples:
  - debugConfig.gs:77: const levelValue = DEBUG_CONFIG.logLevels[level] || DEBUG_CONFIG.logLevels.DEBUG;
  - monitoring.gs:18: this.currentLevel = this.logLevels.INFO;

## _buildErrorInfo
- Occurrences: 2
- Files: errorHandler.gs
- Samples:
  - errorHandler.gs:52: const errorInfo = this._buildErrorInfo(error, context, severity, category, metadata);
  - errorHandler.gs:192: _buildErrorInfo(error, context, severity, category, metadata) {

## _isRetryableError
- Occurrences: 2
- Files: errorHandler.gs
- Samples:
  - errorHandler.gs:107: retryable: this._isRetryableError(error),
  - errorHandler.gs:227: _isRetryableError(error) {

## READ_OPERATION
- Occurrences: 2
- Files: lockManager.gs
- Samples:
  - lockManager.gs:8: READ_OPERATION: 5000, // 読み取り専用操作: 5秒
  - lockManager.gs:16: * @param {string} operationType - 操作タイプ (READ_OPERATION, WRITE_OPERATION, CRITICAL_OPERATION, BATCH_

## WRITE_OPERATION
- Occurrences: 2
- Files: lockManager.gs
- Samples:
  - lockManager.gs:9: WRITE_OPERATION: 10000, // 通常の書き込み操作: 10秒（リアクション高速化）
  - lockManager.gs:16: * @param {string} operationType - 操作タイプ (READ_OPERATION, WRITE_OPERATION, CRITICAL_OPERATION, BATCH_

## CRITICAL_OPERATION
- Occurrences: 2
- Files: lockManager.gs
- Samples:
  - lockManager.gs:10: CRITICAL_OPERATION: 20000, // クリティカルな操作: 20秒
  - lockManager.gs:16: * @param {string} operationType - 操作タイプ (READ_OPERATION, WRITE_OPERATION, CRITICAL_OPERATION, BATCH_

## BATCH_OPERATION
- Occurrences: 2
- Files: lockManager.gs
- Samples:
  - lockManager.gs:11: BATCH_OPERATION: 30000, // バッチ処理: 30秒 (saveAndPublish等)
  - lockManager.gs:16: * @param {string} operationType - 操作タイプ (READ_OPERATION, WRITE_OPERATION, CRITICAL_OPERATION, BATCH_

## V8
- Occurrences: 2
- Files: main.gs, monitoring.gs
- Samples:
  - main.gs:3: * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
  - monitoring.gs:258: gasVersion: 'V8 Runtime',

## getContent
- Occurrences: 2
- Files: main.gs, session-utils.gs
- Samples:
  - main.gs:15: return tmpl.evaluate().getContent();
  - session-utils.gs:248: const outputContent = htmlOutput.getContent();

## gradient
- Occurrences: 2
- Files: main.gs
- Samples:
  - main.gs:1037: background: linear-gradient(135deg, #1e3a8a 0%, #7c3aed 100%);
  - main.gs:1069: background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);

## _persistLog
- Occurrences: 2
- Files: monitoring.gs
- Samples:
  - monitoring.gs:59: this._persistLog(logEntry);
  - monitoring.gs:66: _persistLog(logEntry) {

## shift
- Occurrences: 2
- Files: monitoring.gs, secretManager.gs
- Samples:
  - monitoring.gs:75: errorLogs.shift();
  - secretManager.gs:565: this.auditLog.shift();

## runAllChecks
- Occurrences: 2
- Files: monitoring.gs
- Samples:
  - monitoring.gs:135: async runAllChecks() {
  - monitoring.gs:409: return healthChecker.runAllChecks()

## _runSingleCheck
- Occurrences: 2
- Files: monitoring.gs
- Samples:
  - monitoring.gs:143: const result = await this._runSingleCheck(name, check);
  - monitoring.gs:185: async _runSingleCheck(name, check) {

## fromEntries
- Occurrences: 2
- Files: monitoring.gs, systemIntegrationManager.gs
- Samples:
  - monitoring.gs:174: checks: Object.fromEntries(results),
  - systemIntegrationManager.gs:380: components: Object.fromEntries(

## _generateSummary
- Occurrences: 2
- Files: monitoring.gs
- Samples:
  - monitoring.gs:175: summary: this._generateSummary(results)
  - monitoring.gs:209: _generateSummary(results) {

## activeContexts
- Occurrences: 2
- Files: monitoring.gs
- Samples:
  - monitoring.gs:345: const contextCount = globalContextManager.activeContexts.size;
  - monitoring.gs:372: activeContexts: globalContextManager.activeContexts.size,

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

## action
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:569: if (action.includes('ERROR') || action.includes('FAILED')) {
  - secretManager.gs:569: if (action.includes('ERROR') || action.includes('FAILED')) {

## DISABLED
- Occurrences: 2
- Files: secretManager.gs
- Samples:
  - secretManager.gs:615: results.secretManagerStatus = 'DISABLED';
  - secretManager.gs:646: results.encryptionStatus = 'DISABLED';

## deleteAllProperties
- Occurrences: 2
- Files: session-utils.gs
- Samples:
  - session-utils.gs:86: props.deleteAllProperties();
  - session-utils.gs:131: props.deleteAllProperties();

## urlError
- Occurrences: 2
- Files: session-utils.gs
- Samples:
  - session-utils.gs:162: console.warn('⚠️ WebAppURL取得失敗、フォールバック使用:', urlError.message);
  - session-utils.gs:163: console.warn('⚠️ URLエラースタック:', urlError.stack);

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

## UNFORMATTED_VALUE
- Occurrences: 2
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:48: valueRenderOption = 'UNFORMATTED_VALUE',
  - unifiedBatchProcessor.gs:147: responseValueRenderOption = 'UNFORMATTED_VALUE',

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

## clearCache
- Occurrences: 2
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:898: window.gasOptimizer.clearCache();
  - unifiedCacheManager.gs:1022: gasOptimizer: () => typeof window !== 'undefined' && window.gasOptimizer?.clearCache(),

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

## isPublishedBoard
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:142: if (!isPublishedBoard(userInfo)) {
  - unifiedSheetDataManager.gs:186: if (!isPublishedBoard(userInfo)) {

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

## getEffectiveSpreadsheetId
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:582: getEffectiveSpreadsheetId(userId, usePublished = false) {
  - unifiedSheetDataManager.gs:938: return unifiedSheetManager.getEffectiveSpreadsheetId(requestUserId, true);

## getSheetDetails
- Occurrences: 2
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:673: getSheetDetails(userId, spreadsheetId, sheetName) {
  - unifiedSheetDataManager.gs:960: return unifiedSheetManager.getSheetDetails(requestUserId, spreadsheetId, sheetName);

## invalidate
- Occurrences: 2
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:62: invalidate(userId, email) {
  - unifiedUserManager.gs:148: unifiedUserCache.invalidate(userId, email);

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

## WORKFLOW
- Occurrences: 2
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:16: WORKFLOW: 'workflow'
  - unifiedValidationSystem.gs:55: case this.validationCategories.WORKFLOW:

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

## getActiveSpreadsheet
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:70: return SpreadsheetApp.getActiveSpreadsheet();

## pop
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:103: return (email || '').toString().split('@').pop().toLowerCase();

## emailA
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:107: if (!emailA || !emailB || typeof emailA !== 'string' || typeof emailB !== 'string') {

## emailB
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:107: if (!emailA || !emailB || typeof emailA !== 'string' || typeof emailB !== 'string') {

## ADMIN_EMAILS
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:128: const str = props.getProperty('ADMIN_EMAILS') || '';

## getOwner
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:136: const owner = ss.getOwner();

## ACTIVE_SHEET_NAME
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:260: const activeSheetName = appSettings.activeSheetName || props.getProperty('ACTIVE_SHEET_NAME') || '';

## ACTIVE_SHEET_CHANGED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:317: auditLog('ACTIVE_SHEET_CHANGED', userId, { sheetName, userEmail: userInfo.adminEmail });

## ACTIVE_SHEET_CLEARED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:347: auditLog('ACTIVE_SHEET_CLEARED', userId, { userEmail: userInfo.adminEmail });

## SHOW_DETAILS_UPDATED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:378: auditLog('SHOW_DETAILS_UPDATED', userId, { showNames: Boolean(flag), showCounts: Boolean(flag), user

## DISPLAY_OPTIONS_UPDATED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:406: auditLog('DISPLAY_OPTIONS_UPDATED', userId, { options, userEmail: userInfo.adminEmail });

## UNPUBLISHED_ACCESS_NO_LOGIN
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:469: auditLog('UNPUBLISHED_ACCESS_NO_LOGIN', validatedUserId, { error: e });

## ACCESS_DENIED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:483: auditLog('ACCESS_DENIED', validatedUserId, { viewerEmail, adminEmail: userInfo.adminEmail });

## ADMIN_ACCESS
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:502: auditLog('ADMIN_ACCESS', validatedUserId, { viewerEmail });

## ADMIN_ACCESS_NO_SHEET
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:529: auditLog('ADMIN_ACCESS_NO_SHEET', validatedUserId, { viewerEmail });

## VIEW_ACCESS
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:596: auditLog('VIEW_ACCESS', validatedUserId, { viewerEmail, sheetName: activeSheetName });

## spreadsheetUrl
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:732: const match = spreadsheetUrl.match(urlPattern);

## SPREADSHEET_ADDED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:798: auditLog('SPREADSHEET_ADDED', userId, {

## isSheetHidden
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:903: const visibleSheets = allSheets.filter(sheet => !sheet.isSheetHidden());

## variants
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:924: * Keywords cover common variants (e.g. コメント vs 回答) and include

## characters
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:925: * both full/half-width characters. Header strings are normalized by

## searching
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:927: * searching.

## getValue
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:1469: const current = !!cell.getValue();

## insertColumnAfter
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:1583: sheet.insertColumnAfter(lastCol);

## REGISTER
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:1888: checkRateLimit('REGISTER', userEmail);

## AUTO_FORM_CREATED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:1912: auditLog('AUTO_FORM_CREATED', '', { userEmail, spreadsheetId, formId: formResult.formId });

## UNAUTHORIZED_USER_ACCESS
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2106: auditLog('UNAUTHORIZED_USER_ACCESS', userId, { currentUser });

## UNAUTHORIZED_CONFIG_UPDATE
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2148: auditLog('UNAUTHORIZED_CONFIG_UPDATE', userId, { currentUser, config });

## CONFIG_UPDATED
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2180: auditLog('CONFIG_UPDATED', userId, { currentUser, updatedFields: Object.keys(sanitizedConfig), newCo

## done
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2240: console.log('Setup already done.');

## initialized
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2255: throw new Error('Database not initialized. Please run setup().');

## SHA_256
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2300: Utilities.DigestAlgorithm.SHA_256,

## logging
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2312: // Use cache for temporary audit logging (since we can't create sheets dynamically in all environmen

## AUDIT
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2330: console.log(`AUDIT: ${action} by ${currentUser} for ${userId}`, details);

## only
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2359: * Find existing user by email only (used for single board per user).

## ErrorHandling
- Occurrences: 1
- Files: Code.gs
- Samples:
  - Code.gs:2429: handleError = require('./ErrorHandling.gs').handleError;

## ISO
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:8: * @param {string} publishedAt - 公開開始時間のISO文字列

## EABDB
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:171: const sheetName = userInfo.sheetName || 'EABDB';

## UltraOptimizedCore
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:233: // doGetLegacy function removed - consolidated into main doGet in UltraOptimizedCore.gs

## getSheetId
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1000: id: sheet.getSheetId() // シートIDも必要に応じて取得

## generalError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:1891: console.error('❌ ファイル移動処理で予期しないエラー:', generalError.message);

## core
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2260: debugLog('getHeaderIndices received in core.gs: spreadsheetId=%s, sheetName=%s', spreadsheetId, shee

## getSheetById
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2284: var sheet = spreadsheet.getSheetById(sheetId);

## NOTE
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2485: // NOTE: unpublishBoard関数の重複を回避するため、config.gsの実装を使用

## unpublishBoard
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2486: // function unpublishBoard(requestUserId) {

## emailError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2550: console.warn('Email collection setting failed:', emailError.message);

## createParagraphTextValidation
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2604: var validation = FormApp.createParagraphTextValidation()

## requireTextLengthLessThanOrEqualTo
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2605: .requireTextLengthLessThanOrEqualTo(140)

## build
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2606: .build();

## setValidation
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2607: reasonItem.setValidation(validation);

## addCheckboxItem
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2635: mainItem = form.addCheckboxItem();

## addMultipleChoiceItem
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2644: mainItem = form.addMultipleChoiceItem();

## DOMAIN
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2946: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

## VIEW
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2946: file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

## sharingError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2953: console.warn('共有設定の変更に失敗しましたが、処理を続行します: ' + sharingError.message);

## rename
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:2960: spreadsheetObj.rename(spreadsheetName);

## undocumentedError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3128: console.warn('Email collection type setting failed:', undocumentedError.message);

## setConfirmationMessage
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3146: form.setConfirmationMessage(confirmationMessage);

## serviceAccountError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3152: console.warn('サービスアカウント追加に失敗しましたが、処理を継続します:', serviceAccountError.message);

## sessionLogError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3199: console.warn('セッション記録でエラー:', sessionLogError.message);

## DOMAIN_WITH_LINK
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3236: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## EDIT
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3236: file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

## domainSharingError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3239: console.warn('ドメイン共有設定に失敗: ' + domainSharingError.message);

## individualError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3246: console.error('個別ユーザー追加も失敗: ' + individualError.message);

## spreadsheetAddError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3256: console.warn('SpreadsheetApp経由の追加で警告: ' + spreadsheetAddError.message);

## resizeError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3320: console.warn('Auto-resize failed:', resizeError.message);

## finalRepairError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3505: console.error('getSheetsList: 最終修復も失敗:', finalRepairError.message);

## property
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3519: console.error('getSheetsList: Spreadsheet data missing sheets property. Available properties:', Obje

## reverse
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3647: return data.reverse();

## val
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3677: return val.toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);

## actualHeaderIndices
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3723: if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {

## prev
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3776: return (prev.header.length > current.header.length) ? prev : current;

## cellError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:3861: console.warn('リアクション取得エラー(' + reactionKey + '): ' + cellError.message);

## activateSheet
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4392: return activateSheet(requestUserId, userInfo.spreadsheetId, sheetName);

## demand
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4630: * Confirms registration for the active user on demand.

## OPTIMIZED
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4651: * 統合初期データ取得API - OPTIMIZED

## consistencyError
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4699: console.warn('⚠️ データ整合性チェック中にエラー:', consistencyError.message);

## getSheetDetailsFromContext
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4862: var sheetDetails = getSheetDetailsFromContext(context, userInfo.spreadsheetId, includeSheetDetails);

## includedApis
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4864: response._meta.includedApis.push('getSheetDetails');

## commitAllChanges
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4867: commitAllChanges(context);

## getApplicationEnabled
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4957: const isEnabled = getApplicationEnabled();

## setApplicationEnabled
- Occurrences: 1
- Files: Core.gs
- Samples:
  - Core.gs:4984: const result = setApplicationEnabled(enabled);

## 2a
- Occurrences: 1
- Files: auth.gs
- Samples:
  - auth.gs:236: // 2a. アクティブユーザーの場合

## 2b
- Occurrences: 1
- Files: auth.gs
- Samples:
  - auth.gs:256: // 2b. 非アクティブユーザーの場合

## 3a
- Occurrences: 1
- Files: auth.gs
- Samples:
  - auth.gs:269: // 3a. 新規ユーザーデータを準備（初期設定でpending状態）

## 3b
- Occurrences: 1
- Files: auth.gs
- Samples:
  - auth.gs:288: // 3b. データベースに作成

## 3c
- Occurrences: 1
- Files: auth.gs
- Samples:
  - auth.gs:295: // 3c. 新規ユーザーの管理パネルへリダイレクト

## updateUserField
- Occurrences: 1
- Files: auth.gs
- Samples:
  - auth.gs:336: updateUserField(userId, 'lastAccessedAt', now);

## ALREADY_INITIALIZED
- Occurrences: 1
- Files: autoInit.gs
- Samples:
  - autoInit.gs:17: skipReason: 'ALREADY_INITIALIZED',

## memoError
- Occurrences: 1
- Files: cache.gs
- Samples:
  - cache.gs:840: console.warn('メモ化キャッシュリセットをスキップ:', memoError.message);

## one
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:7: console.log('Config sheet not found, creating new one');

## autoError
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:46: console.error('Auto config creation failed:', autoError.message);

## fill
- Occurrences: 1
- Files: config.gs
- Samples:
  - config.gs:132: const row = new Array(headers.length).fill('');

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

## rollbackActions
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:113: transactionLog.rollbackActions.push('remove_created_sheet');

## appendError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:158: throw new Error(`ログエントリの追加に失敗: ${appendError.message}`);

## generation
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:571: console.error('❌ Access token is null or undefined after generation.');

## SERVICE_ERROR
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:752: serviceError.type = 'SERVICE_ERROR';

## DATA_ACCESS_ERROR
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:777: dataError.type = 'DATA_ACCESS_ERROR';

## NO_DATA_ERROR
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:787: noDataError.type = 'NO_DATA_ERROR';

## FALSE
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:838: } else if (user.isActive === false || user.isActive === 'false' || user.isActive === 'FALSE') {

## USER_NOT_FOUND
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:862: notFoundError.type = 'USER_NOT_FOUND';

## diagError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:909: console.error('❌ 診断・修復処理でエラー:', diagError.message);

## UNAUTHENTICATED
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1092: if (errorMessage.includes('401') || errorMessage.includes('UNAUTHENTICATED') || errorMessage.include

## ACCESS_TOKEN_EXPIRED
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1092: if (errorMessage.includes('401') || errorMessage.includes('UNAUTHENTICATED') || errorMessage.include

## INSERT_ROWS
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1559: ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';

## ALL_USERS
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1765: console.log('🔍 データベース診断開始:', targetUserId || 'ALL_USERS');

## statsError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:1958: console.warn('キャッシュ統計取得エラー:', statsError.message);

## serviceAccount
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2064: if (result.checks.serviceAccount && result.checks.serviceAccount.email) {

## GENERAL
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2132: console.log('🔧 自動修復開始:', targetUserId || 'GENERAL');

## verifyError
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2195: repairResult.summary = '修復実行したが検証に失敗: ' + verifyError.message;

## UUID
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2448: // ユーザーID形式チェック（UUIDかどうか）

## B1
- Occurrences: 1
- Files: database.gs
- Samples:
  - database.gs:2822: var testData = batchGetSheetsData(service, dbId, ["'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A1:B1"]);

## ENVIRONMENT
- Occurrences: 1
- Files: debugConfig.gs
- Samples:
  - debugConfig.gs:39: const envSetting = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT');

## args
- Occurrences: 1
- Files: debugConfig.gs
- Samples:
  - debugConfig.gs:90: const logData = args.length > 0 ? { additionalData: args } : {};

## ms
- Occurrences: 1
- Files: errorHandler.gs
- Samples:
  - errorHandler.gs:166: `Performance threshold exceeded: ${operation} took ${duration}ms (threshold: ${threshold}ms)`,

## performance
- Occurrences: 1
- Files: errorHandler.gs
- Samples:
  - errorHandler.gs:167: `performance.${operation}`,

## HTTPS
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:32: // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理

## EXECUTION_LIMITS
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:163: EXECUTION_LIMITS: {

## MAX_TIME
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:164: MAX_TIME: 330000, // 5.5分（安全マージン）

## API_RATE_LIMIT
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:166: API_RATE_LIMIT: 90 // 100秒間隔での制限

## CACHE_STRATEGY
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:169: CACHE_STRATEGY: {

## L1_TTL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:170: L1_TTL: 300,     // Level 1: 5分

## L2_TTL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:171: L2_TTL: 3600,    // Level 2: 1時間

## L3_TTL
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:172: L3_TTL: 21600    // Level 3: 6時間（最大）

## apply
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:213: console.log.apply(console, arguments);

## Login
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:221: * AdminPanel.html と Login.html から共通で呼び出される

## PerformanceOptimizer
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:272: // PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、

## XSS
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1007: // XSS攻撃を防ぐため、URLをサニタイズ

## setContent
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1009: return HtmlService.createHtmlOutput().setContent(`<script>window.top.location.href = '${sanitizedUrl

## DOCTYPE
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1028: <!DOCTYPE html>

## translateY
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1081: transform: translateY(-2px);

## okinawa
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1218: const fallbackUrl = 'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVq

## ed
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1218: const fallbackUrl = 'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVq

## NATIVE
- Occurrences: 1
- Files: main.gs
- Samples:
  - main.gs:1284: .setSandboxMode(HtmlService.SandboxMode.NATIVE);

## fatal
- Occurrences: 1
- Files: monitoring.gs
- Samples:
  - monitoring.gs:104: fatal(message, metadata = {}) { this.log('FATAL', message, metadata); }

## getActiveLocale
- Occurrences: 1
- Files: monitoring.gs
- Samples:
  - monitoring.gs:260: locale: Session.getActiveLocale(),

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

## TEST_KEY
- Occurrences: 1
- Files: secretManager.gs
- Samples:
  - secretManager.gs:621: props.getProperty('TEST_KEY'); // 存在しないキーでのテスト

## scriptCacheError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:53: console.warn('スクリプトキャッシュクリア中のエラー: ' + scriptCacheError.message);

## LAST_ACCESS_EMAIL
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:84: // PropertiesServiceもクリアする（LAST_ACCESS_EMAILなど）

## sanitizeError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:171: console.error('❌ Fallback URL sanitization failed:', sanitizeError.message);

## preview
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:221: console.log('📄 Generated HTML preview (first 200 chars):', redirectScript.substring(0, 200));

## frameError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:243: console.warn('XFrameOptionsMode設定失敗:', frameError.message);

## NO
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:250: console.log('📋 HtmlOutput content preview:', outputContent ? outputContent.substring(0, 100) : 'NO 

## CONTENT
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:250: console.log('📋 HtmlOutput content preview:', outputContent ? outputContent.substring(0, 100) : 'NO 

## contentError
- Occurrences: 1
- Files: session-utils.gs
- Samples:
  - session-utils.gs:252: console.warn('⚠️ Cannot access HtmlOutput content:', contentError.message);

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

## SHUTDOWN
- Occurrences: 1
- Files: systemIntegrationManager.gs
- Samples:
  - systemIntegrationManager.gs:357: component.status = 'SHUTDOWN';

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

## SERIAL_NUMBER
- Occurrences: 1
- Files: unifiedBatchProcessor.gs
- Samples:
  - unifiedBatchProcessor.gs:47: dateTimeRenderOption = 'SERIAL_NUMBER',

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

## setUserInfo
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1154: setUserInfo(userId, userInfo) {

## sessionError
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:1367: ULog.debug('セッションキャッシュ保存エラー:', sessionError.message);

## clearPattern
- Occurrences: 1
- Files: unifiedCacheManager.gs
- Samples:
  - unifiedCacheManager.gs:2136: } else if (typeof clearPattern === 'string' && clearPattern !== 'false') {

## formUrl
- Occurrences: 1
- Files: unifiedSecurityManager.gs
- Samples:
  - unifiedSecurityManager.gs:258: if (config.formCreated === true && config.formUrl && config.formUrl.trim()) {

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

## title
- Occurrences: 1
- Files: unifiedSheetDataManager.gs
- Samples:
  - unifiedSheetDataManager.gs:645: !sheet.title.startsWith('名簿') // 名簿シート除外

## findIndex
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:161: const fieldIndex = headers.findIndex(

## getResilientUserProperties
- Occurrences: 1
- Files: unifiedUserManager.gs
- Samples:
  - unifiedUserManager.gs:243: const props = getResilientUserProperties();

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

## createFormForSpreadsheet
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:109: forSpreadsheet: (spreadsheetId, sheetName) => createFormForSpreadsheet(spreadsheetId, sheetName),

## buildResponseFromContext
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:145: response: (context) => buildResponseFromContext(context)

## createBoardFromAdmin
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:152: fromAdmin: (requestUserId) => createBoardFromAdmin(requestUserId)

## sanitizeURL
- Occurrences: 1
- Files: unifiedUtilities.gs
- Samples:
  - unifiedUtilities.gs:171: static sanitizeURL(url) {

## validateUserAuthentication
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:243: // validateUserAuthentication() の統合実装

## validateConfigJson
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:347: const result = validateConfigJson(config); // 既存関数利用

## validateConfigJsonState
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:364: const result = validateConfigJsonState(config, userInfo); // 既存関数利用

## validateSystemDependencies
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:397: const result = validateSystemDependencies(); // 既存関数利用

## checkAndHandleAutoStop
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:785: const result = checkAndHandleAutoStop(config, userInfo); // 既存関数利用

## checkIfNewOrUpdatedForm
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:802: const result = checkIfNewOrUpdatedForm(requestUserId, spreadsheetId, sheetName); // 既存関数利用

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

## every
- Occurrences: 1
- Files: unifiedValidationSystem.gs
- Samples:
  - unifiedValidationSystem.gs:1064: overallSuccess: Object.values(results).every(r => r.success),

## com
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:87: /^https:\/\/[a-z0-9-]+\.googleusercontent\.com\/$/ // googleusercontent.com (デプロイ形式)

## getScriptId
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:171: var scriptId = ScriptApp.getScriptId();

## reset
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:218: console.log('Forcing URL system reset...');

## substr
- Occurrences: 1
- Files: url.gs
- Samples:
  - url.gs:368: url.searchParams.set('_t', Math.random().toString(36).substr(2, 9));

## E2E
- Occurrences: 1
- Files: workflowValidation.gs
- Samples:
  - workflowValidation.gs:3: * Claude Code開発環境のE2Eテスト用
