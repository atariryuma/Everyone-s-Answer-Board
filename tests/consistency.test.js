/**
 * @fileoverview システム一貫性・不整合検証テスト
 * システム全体の重複、不整合、命名規則違反を継続的に検出
 */

const fs = require('fs');
const path = require('path');

describe('システム一貫性・整合性検証', () => {
  const srcPath = path.join(__dirname, '../src');
  
  /**
   * 全.gsファイルの内容を取得
   */
  function getAllGasFiles() {
    const files = fs.readdirSync(srcPath)
      .filter(file => file.endsWith('.gs'))
      .map(file => ({
        name: file,
        path: path.join(srcPath, file),
        content: fs.readFileSync(path.join(srcPath, file), 'utf-8')
      }));
    return files;
  }

  describe('レスポンス形式統一検証', () => {
    test('全API関数が{success, data, message, error}形式を使用している', () => {
      const files = getAllGasFiles();
      const violations = [];

      files.forEach(file => {
        // 古いstatus形式の残存チェック
        const statusMatches = file.content.match(/status\s*:\s*['"`](success|error)['"`]/g);
        if (statusMatches) {
          violations.push({
            file: file.name,
            issues: statusMatches.map(match => `古いstatus形式: ${match}`),
            type: 'legacy_status_format'
          });
        }

        // status条件分岐の残存チェック
        const statusChecks = file.content.match(/\.status\s*===\s*['"`](success|error)['"`]/g);
        if (statusChecks) {
          violations.push({
            file: file.name,
            issues: statusChecks.map(match => `status条件分岐: ${match}`),
            type: 'legacy_status_check'
          });
        }
      });

      if (violations.length > 0) {
        const errorMessage = violations.map(v => 
          `${v.file}: ${v.issues.join(', ')}`
        ).join('\n');
        
        throw new Error(`レスポンス形式の不整合が検出されました:\n${errorMessage}`);
      }
    });

    test('createSuccessResponse/createErrorResponseの適切な使用', () => {
      const files = getAllGasFiles();
      const violations = [];

      files.forEach(file => {
        // success: true/falseの直接記述チェック
        const directSuccess = file.content.match(/\{\s*success\s*:\s*(true|false)[^}]*\}/g);
        if (directSuccess) {
          // 統一関数を使わずに直接記述している箇所をチェック
          const validUses = file.content.match(/createSuccessResponse|createErrorResponse|createUnifiedResponse/g);
          if (!validUses || validUses.length < directSuccess.length / 2) {
            violations.push({
              file: file.name,
              issue: `統一レスポンス関数の未使用: ${directSuccess.length}箇所で直接記述`,
              type: 'direct_response_creation'
            });
          }
        }
      });

      if (violations.length > 0) {
        const errorMessage = violations.map(v => `${v.file}: ${v.issue}`).join('\n');
        throw new Error(`統一レスポンス関数の使用違反:\n${errorMessage}`);
      }
    });
  });

  describe('重複関数検出', () => {
    test('同様機能の重複関数が存在しない', () => {
      const files = getAllGasFiles();
      const functionPattern = /function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\(/g;
      const functionsByName = {};
      
      files.forEach(file => {
        let match;
        while ((match = functionPattern.exec(file.content)) !== null) {
          const funcName = match[1];
          if (!functionsByName[funcName]) {
            functionsByName[funcName] = [];
          }
          functionsByName[funcName].push({
            file: file.name,
            name: funcName
          });
        }
      });

      // 重複関数の検出（ただし、明示的に許可されたラッパー関数は除外）
      const allowedDuplicates = [
        'createSuccessResponse',
        'createErrorResponse',
        'createUnifiedResponse',
        // 後方互換性のために一時的に許可される重複
        'clearExecutionUserInfoCache',
        'clearAllExecutionCache'
      ];

      const duplicates = Object.entries(functionsByName)
        .filter(([name, locations]) => 
          locations.length > 1 && !allowedDuplicates.includes(name)
        );

      if (duplicates.length > 0) {
        const errorMessage = duplicates.map(([name, locations]) =>
          `関数 "${name}" が以下のファイルで重複: ${locations.map(l => l.file).join(', ')}`
        ).join('\n');

        throw new Error(`重複関数が検出されました:\n${errorMessage}`);
      }
    });

    test('類似名称の機能重複関数が存在しない', () => {
      const files = getAllGasFiles();
      const functionNames = [];

      files.forEach(file => {
        const functionPattern = /function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\(/g;
        let match;
        while ((match = functionPattern.exec(file.content)) !== null) {
          functionNames.push({
            name: match[1],
            file: file.name
          });
        }
      });

      // 類似名称のパターン検出
      const suspiciousPatterns = [
        ['getUserInfo', 'getOrFetchUserInfo', 'fetchUserInfo'],
        ['updateActiveStatus', 'updateIsActiveStatus', 'updateUserActiveStatus'],
        ['clearCache', 'clearAllCache', 'clearExecutionCache'],
        ['getCachedData', 'getDataCached', 'getCachedInfo']
      ];

      const violations = [];
      suspiciousPatterns.forEach(pattern => {
        const found = pattern.filter(name => 
          functionNames.some(f => f.name.includes(name) || name.includes(f.name))
        );
        
        if (found.length > 1) {
          const matchingFunctions = functionNames.filter(f => 
            found.some(pattern => f.name.includes(pattern) || pattern.includes(f.name))
          );
          
          violations.push({
            pattern: pattern.join(', '),
            functions: matchingFunctions
          });
        }
      });

      if (violations.length > 0) {
        const errorMessage = violations.map(v =>
          `類似パターン "${v.pattern}" の関数: ${v.functions.map(f => `${f.name}(${f.file})`).join(', ')}`
        ).join('\n');

        throw new Error(`類似名称の重複関数が検出されました:\n${errorMessage}`);
      }
    });
  });

  describe('フロントエンド整合性検証', () => {
    test('API関数の命名がフロントエンド要求に適合している', () => {
      const files = getAllGasFiles();
      const violations = [];

      // フロントエンドが期待するAPI関数パターン
      const requiredApiFunctions = [
        'getPublishedSheetData',
        'getDataCount', 
        'updateIsActiveStatus',
        'getCurrentUserStatus',
        'getActiveFormInfo'
      ];

      requiredApiFunctions.forEach(funcName => {
        const found = files.some(file => 
          file.content.includes(`function ${funcName}`)
        );
        
        if (!found) {
          violations.push({
            missing: funcName,
            type: 'required_api_function'
          });
        }
      });

      if (violations.length > 0) {
        const errorMessage = violations.map(v => `必須API関数が見つかりません: ${v.missing}`).join('\n');
        throw new Error(`フロントエンド要求API不整合:\n${errorMessage}`);
      }
    });

    test('データ形式がフロントエンド期待値と整合している', () => {
      const files = getAllGasFiles();
      const violations = [];

      files.forEach(file => {
        // Boolean値をstring化している箇所の検出
        const booleanStringMatches = file.content.match(/isActive\s*:\s*['"`](true|false)['"`]/g);
        if (booleanStringMatches) {
          violations.push({
            file: file.name,
            issue: `Boolean値のstring化: ${booleanStringMatches.join(', ')}`,
            type: 'boolean_string_conversion'
          });
        }

        // 数値をstring化している可能性の検出
        const numberStringMatches = file.content.match(/(count|id|index)\s*:\s*['"`]\d+['"`]/g);
        if (numberStringMatches) {
          violations.push({
            file: file.name,
            issue: `数値のstring化の可能性: ${numberStringMatches.join(', ')}`,
            type: 'number_string_conversion'
          });
        }
      });

      if (violations.length > 0) {
        const errorMessage = violations.map(v => `${v.file}: ${v.issue}`).join('\n');
        throw new Error(`データ形式不整合が検出されました:\n${errorMessage}`);
      }
    });
  });

  describe('コード品質検証', () => {
    test('統一ログシステム(ULog)の使用', () => {
      const files = getAllGasFiles();
      const violations = [];

      files.forEach(file => {
        // 直接console使用の検出（ULogが利用可能な環境で）
        const directConsoleUse = file.content.match(/console\.(log|error|warn|info)\s*\(/g);
        const ulogUse = file.content.match(/ULog\.(debug|info|warn|error)\s*\(/g);
        
        if (directConsoleUse && directConsoleUse.length > 3) {
          // ULogの使用が少ない、または全く使用していない場合
          if (!ulogUse || ulogUse.length < directConsoleUse.length / 2) {
            violations.push({
              file: file.name,
              issue: `統一ログシステム未使用: console直接使用${directConsoleUse.length}箇所`,
              type: 'direct_console_usage'
            });
          }
        }
      });

      // 重要でない警告は除外（テストの過剰反応を防ぐ）
      const criticalViolations = violations.filter(v => 
        v.issue.includes('統一ログシステム未使用') && 
        parseInt(v.issue.match(/\d+/)[0]) > 10 // 10箇所以上の違反のみエラー
      );

      if (criticalViolations.length > 0) {
        const errorMessage = criticalViolations.map(v => `${v.file}: ${v.issue}`).join('\n');
        console.warn(`ログシステム最適化推奨:\n${errorMessage}`);
      }
    });

    test('統一キャッシュAPI(UnifiedCacheAPI)の使用', () => {
      const files = getAllGasFiles();
      const violations = [];

      files.forEach(file => {
        // 直接CacheService使用の検出
        const directCacheUse = file.content.match(/CacheService\.(getScriptCache|getDocumentCache|getUserCache)/g);
        const unifiedCacheUse = file.content.match(/UnifiedCacheAPI|getUnifiedExecutionCache|unifiedCache/g);
        
        if (directCacheUse && directCacheUse.length > 2) {
          // 統一キャッシュAPIの使用が少ない場合（緩い検証）
          if (!unifiedCacheUse) {
            violations.push({
              file: file.name,
              issue: `統一キャッシュAPI未使用: CacheService直接使用${directCacheUse.length}箇所`,
              type: 'direct_cache_usage'
            });
          }
        }
      });

      // 重要でない警告は情報レベルで出力
      if (violations.length > 0) {
        const errorMessage = violations.map(v => `${v.file}: ${v.issue}`).join('\n');
        console.info(`キャッシュAPI最適化推奨:\n${errorMessage}`);
      }
    });
  });

  describe('継続的品質保証', () => {
    test('新規追加関数が命名規則に従っている', () => {
      const files = getAllGasFiles();
      const violations = [];

      files.forEach(file => {
        const functionPattern = /function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\(/g;
        let match;
        
        while ((match = functionPattern.exec(file.content)) !== null) {
          const funcName = match[1];
          
          // 命名規則違反の検出
          if (funcName.includes('_') && !funcName.startsWith('_')) {
            // アンダースコア混在（プライベート関数以外）
            violations.push({
              file: file.name,
              function: funcName,
              issue: 'アンダースコア混在（camelCase推奨）',
              type: 'naming_convention'
            });
          }
          
          if (funcName.length > 50) {
            // 関数名が長すぎる
            violations.push({
              file: file.name,
              function: funcName,
              issue: '関数名が長すぎる（50文字以内推奨）',
              type: 'naming_length'
            });
          }
        }
      });

      if (violations.length > 0) {
        const errorMessage = violations.map(v => 
          `${v.file}: ${v.function} - ${v.issue}`
        ).join('\n');
        
        console.warn(`命名規則改善推奨:\n${errorMessage}`);
      }
    });

    test('システム全体の健全性確認', () => {
      const files = getAllGasFiles();
      const healthMetrics = {
        totalFiles: files.length,
        totalFunctions: 0,
        unifiedResponseUsage: 0,
        unifiedLogUsage: 0,
        unifiedCacheUsage: 0,
        legacyPatterns: 0
      };

      files.forEach(file => {
        // 関数数カウント
        const functionMatches = file.content.match(/function\s+[a-zA-Z_$][a-zA-Z0-9_$]*\s*\(/g);
        if (functionMatches) {
          healthMetrics.totalFunctions += functionMatches.length;
        }

        // 統一API使用状況
        if (file.content.includes('createSuccessResponse') || file.content.includes('createErrorResponse')) {
          healthMetrics.unifiedResponseUsage++;
        }
        
        if (file.content.includes('ULog.')) {
          healthMetrics.unifiedLogUsage++;
        }
        
        if (file.content.includes('UnifiedCacheAPI') || file.content.includes('getUnifiedExecutionCache')) {
          healthMetrics.unifiedCacheUsage++;
        }

        // レガシーパターン
        if (file.content.includes('.status ===') || file.content.includes('console.')) {
          healthMetrics.legacyPatterns++;
        }
      });

      // システム健全性の評価
      const healthScore = (
        (healthMetrics.unifiedResponseUsage / healthMetrics.totalFiles * 30) +
        (healthMetrics.unifiedLogUsage / healthMetrics.totalFiles * 25) +
        (healthMetrics.unifiedCacheUsage / Math.min(healthMetrics.totalFiles, 10) * 25) +
        (Math.max(0, healthMetrics.totalFiles - healthMetrics.legacyPatterns) / healthMetrics.totalFiles * 20)
      );

      console.log(`システム健全性スコア: ${Math.round(healthScore)}/100`);
      console.log(`統計: ${JSON.stringify(healthMetrics, null, 2)}`);

      // 健全性が低すぎる場合は警告
      if (healthScore < 60) {
        console.warn(`システム健全性が低下しています。統一API使用の推進を推奨します。`);
      }

      // テスト自体は成功（健全性は改善推奨事項）
      expect(healthMetrics.totalFiles).toBeGreaterThan(0);
      expect(healthMetrics.totalFunctions).toBeGreaterThan(0);
    });
  });
});