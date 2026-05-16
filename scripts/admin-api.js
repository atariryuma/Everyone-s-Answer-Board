/**
 * Admin API クライアント
 *
 * Usage:
 *   npm run api -- <operation> [--key value ...] [--json '{"...":"..."}'] [--file path.json]
 *
 * Examples:
 *   npm run api -- getAppStatus
 *   npm run api -- systemDiagnosis
 *   npm run api -- getUsers
 *   npm run api -- getLogs --limit 20
 *   npm run api -- perfMetrics --category api
 *   npm run api -- listProperties
 *
 *   npm run api -- findUser --email t260781p@naha-okinawa.ed.jp
 *   npm run api -- getUserConfig --userId adb24b94-...
 *   npm run api -- exportConfigs --file backup.json    # ファイルへ保存
 *   npm run api -- setUserConfig --userId <uuid> --patch '{"displaySettings":{"boardMode":"numberline"}}'
 *   npm run api -- bulkSetUserConfig --filter '{"isPublished":false}' --patch '{...}' --dryRun
 *   npm run api -- runColumnAnalysis --userId <uuid>
 *   npm run api -- previewBoard --userId <uuid>
 *
 *   # Lesson workspace (Phase 1+2)
 *   npm run api -- lesson.list --userId <uuid>
 *   npm run api -- lesson.create --userId <uuid> --name "5/15 テスト" --template doutoku-3phase
 *   npm run api -- lesson.updateDraft --userId <uuid> --lessonId <id> --fieldPath classes --value '["5-1","5-2"]'
 *   npm run api -- lesson.start --userId <uuid> --lessonId <id>
 *   npm run api -- lesson.advance --userId <uuid> --lessonId <id> --direction next
 *   npm run api -- lesson.end --userId <uuid> --lessonId <id>
 *   npm run api -- lesson.review --userId <uuid> --lessonId <id>
 *
 * Flag types:
 *   --key value           : 文字列。数値文字列は数値化される
 *   --key=value           : 同上、= 区切り
 *   --flag                : boolean true（直後に値が無いとき）
 *   --json '<json>'       : 任意 JSON。`--patch '{...}'` で部分更新に使う
 *   --file <path>         : ファイルから JSON を読み込み、すべての params に展開
 *
 * 結果の出力:
 *   stdout: JSON。`--output <path>` で書き出し。
 */
const fs = require('fs');
const path = require('path');
const { getAccessToken, postJSONSync, getConfig, parseEnvFromArgs } = require('./lib/gas-auth');

const OPERATIONS = [
  // user / app
  'getUsers', 'toggleUserActive', 'toggleUserBoard',
  // Board publish lifecycle (unified)
  'unpublishBoard', 'republishMyBoard',
  'getLogs', 'disableApp', 'enableApp', 'getAppStatus',
  // system
  'systemDiagnosis', 'autoRepair', 'cacheReset',
  'perfMetrics', 'perfDiagnosis',
  // properties
  'getProperty', 'setProperty', 'listProperties',
  // user config (v2)
  'findUser', 'getUserConfig', 'exportConfigs',
  'setUserConfig', 'bulkSetUserConfig',
  'runColumnAnalysis', 'previewBoard', 'exportBoardData',
  // Form operations (v3)
  'listMyForms', 'validateFormUrl', 'connectForm', 'createForm', 'customizeForm',
  'setFormAllowResubmit', 'uploadLessonImage',
  // Multi-board profiles (v4)
  'listProfiles', 'saveProfile', 'loadProfile', 'deleteProfile',
  // Data ops (v5) — テストデータ投入専用
  'appendRows', 'clearDataRows',
  // Drive file rename (v5.1) — customizeForm 後の Drive file name 同期 etc.
  'renameDriveFile',
  // Drive sharing (v6)
  'shareWithDomain',
  // SS sharing repair (v7) — SA editor 共有が抜けた既存 SS の遡及修復
  'repairSpreadsheetSharing',
  // Lesson workspace (Phase 1+2) — wizard / runner / replay / archive
  'lesson.create', 'lesson.updateDraft', 'lesson.start',
  'lesson.advance', 'lesson.end',
  'lesson.list', 'lesson.review', 'lesson.delete',
  'lesson.duplicate', 'lesson.templates', 'lesson.knownClasses',
  'lesson.importFromProfiles',
];

/**
 * Parse args supporting:
 *   --key value         → string (numeric coerced)
 *   --key=value         → string (numeric coerced)
 *   --flag (no value)   → boolean true
 *   --json '<obj/arr>'  → parsed JSON, stored under 'json' or merged top-level
 *   --patch '<obj>'     → parsed JSON, stored under 'patch'
 *   --filter '<obj>'    → parsed JSON, stored under 'filter'
 *   --options '<obj>'   → parsed JSON, stored under 'options'
 *   --file <path>       → load JSON file and merge into params
 *
 * Note: JSON-aware keys (json/patch/filter/options) are auto-parsed.
 */
function parseArgs(args) {
  const operation = args[0];
  const params = {};
  let outputPath = null;

  const JSON_KEYS = new Set(['json', 'patch', 'filter', 'options', 'templateOptions', 'schema', 'snapshot', 'rows']);
  // value は lesson.updateDraft の汎用引数。JSON parse 試行 → 失敗時は raw string fallback
  //   (`--value '["5-1"]'` で array、`--value foo` で string が両方扱える)。
  const JSON_OR_STRING_KEYS = new Set(['value']);

  for (let i = 1; i < args.length; i++) {
    const raw = args[i];
    if (!raw || !raw.startsWith('--')) continue;

    let key, value;
    if (raw.includes('=')) {
      const eq = raw.indexOf('=');
      key = raw.slice(2, eq);
      value = raw.slice(eq + 1);
    } else {
      key = raw.slice(2);
      const next = args[i + 1];
      // boolean flag detection: next is missing OR next is another --flag
      if (next === undefined || next.startsWith('--')) {
        value = true;
      } else {
        value = next;
        i++;
      }
    }

    if (key === 'output') { outputPath = value; continue; }
    if (key === 'file' && typeof value === 'string') {
      // load whole file as params object
      try {
        const text = fs.readFileSync(path.resolve(value), 'utf8');
        const obj = JSON.parse(text);
        if (obj && typeof obj === 'object') Object.assign(params, obj);
      } catch (e) {
        console.error(`Failed to read --file ${value}: ${e.message}`);
        process.exit(2);
      }
      continue;
    }

    if (JSON_KEYS.has(key) && typeof value === 'string') {
      try {
        params[key] = JSON.parse(value);
      } catch (e) {
        console.error(`--${key} value is not valid JSON: ${e.message}`);
        process.exit(2);
      }
      continue;
    }

    if (JSON_OR_STRING_KEYS.has(key) && typeof value === 'string') {
      // JSON 試行 → 失敗 (素の文字列) なら raw string でセット。
      try { params[key] = JSON.parse(value); }
      catch (_) { params[key] = value; }
      continue;
    }

    // boolean flags: dryRun / publish / verbose
    if (typeof value === 'boolean') {
      params[key] = value;
      continue;
    }

    // numeric coercion only for non-JSON keys
    if (typeof value === 'string' && value !== '' && !isNaN(value) && !/^[0+]\d/.test(value)) {
      // Avoid coercing IDs that happen to be all-numeric (rare for our UUIDs).
      // Skip coercion if value has leading 0 or + (likely intentional string).
      params[key] = Number(value);
    } else {
      params[key] = value;
    }
  }

  return { operation, params, outputPath };
}

const { env, remainingArgs } = parseEnvFromArgs(process.argv.slice(2));
if (!remainingArgs[0] || remainingArgs[0] === '--help' || remainingArgs[0] === '-h') {
  console.log('Usage: npm run api -- <operation> [--env <env>] [--key value ...]');
  console.log('\nOperations:');
  OPERATIONS.forEach(op => console.log(`  ${op}`));
  console.log('\nJSON-aware flags: --json --patch --filter --options');
  console.log('I/O flags:        --file <path>  --output <path>');
  console.log('\nExamples:');
  console.log("  npm run api -- findUser --email teacher@example.com");
  console.log("  npm run api -- setUserConfig --userId <uuid> --patch '{\"displaySettings\":{\"boardMode\":\"numberline\"}}'");
  console.log("  npm run api -- exportConfigs --output configs.json");
  console.log("  npm run api -- bulkSetUserConfig --filter '{\"isPublished\":false}' --patch '{\"allowResubmit\":true}' --dryRun");
  console.log('\nLesson workspace (Phase 1+2):');
  console.log("  npm run api -- lesson.list --userId <uuid>");
  console.log("  npm run api -- lesson.create --userId <uuid> --name '5/15 道徳テスト' --template doutoku-3phase");
  console.log("  npm run api -- lesson.updateDraft --userId <uuid> --lessonId <id> --fieldPath classes --value '[\"5-1\",\"5-2\"]'");
  console.log("  npm run api -- lesson.updateDraft --userId <uuid> --lessonId <id> --fieldPath name --value 新タイトル");
  console.log("  npm run api -- lesson.start --userId <uuid> --lessonId <id>");
  console.log("  npm run api -- lesson.advance --userId <uuid> --lessonId <id> --direction next");
  console.log("  npm run api -- lesson.end --userId <uuid> --lessonId <id>");
  console.log("  npm run api -- lesson.review --userId <uuid> --lessonId <id>     # full payload");
  console.log("  npm run api -- lesson.delete --userId <uuid> --lessonId <id>");
  process.exit(0);
}

const { operation, params, outputPath } = parseArgs(remainingArgs);

try {
  const config = getConfig(env);
  const token = getAccessToken(config.tokenName);
  const payload = { action: 'adminApi', apiKey: config.apiKey, operation, params };
  const json = postJSONSync(config.prodUrl, payload, token);
  const formatted = JSON.stringify(json, null, 2);
  if (outputPath) {
    fs.writeFileSync(path.resolve(outputPath), formatted);
    console.log(`Wrote ${formatted.length} bytes to ${outputPath}  (success=${json.success})`);
  } else {
    console.log(formatted);
  }
  process.exit(json.success ? 0 : 1);
} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}
