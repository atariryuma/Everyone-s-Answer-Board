/**
 * スクリプトプロパティの大量キャッシュ問題を調査・清掃するスクリプト
 */

function investigateAndCleanupCacheProperties() {
  console.log('🔍 スクリプトプロパティキャッシュ調査開始...');
  
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    const allKeys = Object.keys(allProps);
    
    console.log(`📊 総プロパティ数: ${allKeys.length}`);
    
    // ss_cache_ で始まるキーを特定
    const cacheKeys = allKeys.filter(key => key.startsWith('ss_cache_'));
    console.log(`🗂️ ss_cache_ キーの数: ${cacheKeys.length}`);
    
    // 各キャッシュエントリの詳細分析
    const cacheAnalysis = {
      totalCount: cacheKeys.length,
      totalSize: 0,
      oldEntries: [],
      validEntries: [],
      invalidEntries: [],
      spreadsheetIds: new Set()
    };
    
    const now = Date.now();
    const thirtyDaysAgo = now - (30 * 24 * 60 * 60 * 1000); // 30日前
    
    cacheKeys.forEach((key, index) => {
      try {
        const value = allProps[key];
        const size = value ? value.length : 0;
        cacheAnalysis.totalSize += size;
        
        // JSON解析を試行
        const parsedValue = JSON.parse(value);
        const spreadsheetId = parsedValue.spreadsheetId;
        const timestamp = parsedValue.timestamp;
        
        if (spreadsheetId) {
          cacheAnalysis.spreadsheetIds.add(spreadsheetId);
        }
        
        const entry = {
          key: key,
          spreadsheetId: spreadsheetId,
          timestamp: timestamp,
          age: timestamp ? (now - timestamp) : null,
          size: size
        };
        
        if (timestamp && timestamp < thirtyDaysAgo) {
          cacheAnalysis.oldEntries.push(entry);
        } else if (timestamp && spreadsheetId) {
          cacheAnalysis.validEntries.push(entry);
        } else {
          cacheAnalysis.invalidEntries.push(entry);
        }
        
        // 進捗表示（大量データの場合）
        if (index % 100 === 0) {
          console.log(`処理中... ${index + 1}/${cacheKeys.length}`);
        }
        
      } catch (parseError) {
        cacheAnalysis.invalidEntries.push({
          key: key,
          error: parseError.message,
          size: value ? value.length : 0
        });
      }
    });
    
    // 結果レポート
    console.log('\n📋 === キャッシュ分析レポート ===');
    console.log(`総キャッシュエントリ数: ${cacheAnalysis.totalCount}`);
    console.log(`総データサイズ: ${(cacheAnalysis.totalSize / 1024).toFixed(2)} KB`);
    console.log(`ユニークなスプレッドシートID数: ${cacheAnalysis.spreadsheetIds.size}`);
    console.log(`有効なエントリ: ${cacheAnalysis.validEntries.length}`);
    console.log(`30日以上古いエントリ: ${cacheAnalysis.oldEntries.length}`);
    console.log(`無効なエントリ: ${cacheAnalysis.invalidEntries.length}`);
    
    // 古いエントリの詳細
    if (cacheAnalysis.oldEntries.length > 0) {
      console.log('\n🗑️ === 30日以上古いエントリ（削除対象） ===');
      cacheAnalysis.oldEntries.slice(0, 10).forEach(entry => {
        const ageInDays = Math.floor(entry.age / (24 * 60 * 60 * 1000));
        console.log(`${entry.key} - ${ageInDays}日前 - ${entry.size}B`);
      });
      if (cacheAnalysis.oldEntries.length > 10) {
        console.log(`...他 ${cacheAnalysis.oldEntries.length - 10} 件`);
      }
    }
    
    // 無効なエントリの詳細
    if (cacheAnalysis.invalidEntries.length > 0) {
      console.log('\n❌ === 無効なエントリ（削除対象） ===');
      cacheAnalysis.invalidEntries.slice(0, 10).forEach(entry => {
        console.log(`${entry.key} - ${entry.error || 'パース失敗'} - ${entry.size}B`);
      });
      if (cacheAnalysis.invalidEntries.length > 10) {
        console.log(`...他 ${cacheAnalysis.invalidEntries.length - 10} 件`);
      }
    }
    
    // スプレッドシートID別の統計
    console.log('\n📊 === スプレッドシートID別統計（上位10件） ===');
    const spreadsheetStats = {};
    cacheAnalysis.validEntries.forEach(entry => {
      const id = entry.spreadsheetId;
      if (!spreadsheetStats[id]) {
        spreadsheetStats[id] = { count: 0, totalSize: 0 };
      }
      spreadsheetStats[id].count++;
      spreadsheetStats[id].totalSize += entry.size;
    });
    
    const sortedStats = Object.entries(spreadsheetStats)
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 10);
    
    sortedStats.forEach(([id, stats]) => {
      console.log(`${id}: ${stats.count}件, ${(stats.totalSize / 1024).toFixed(2)}KB`);
    });
    
    return {
      totalProperties: allKeys.length,
      cacheProperties: cacheKeys.length,
      analysis: cacheAnalysis,
      recommendation: generateCleanupRecommendation(cacheAnalysis)
    };
    
  } catch (error) {
    console.error('❌ 調査中にエラーが発生しました:', error);
    return { error: error.message };
  }
}

function generateCleanupRecommendation(analysis) {
  const recommendations = [];
  
  if (analysis.oldEntries.length > 0) {
    recommendations.push({
      priority: 'high',
      action: '古いキャッシュエントリの削除',
      count: analysis.oldEntries.length,
      description: '30日以上古いキャッシュエントリを削除してプロパティ数を削減'
    });
  }
  
  if (analysis.invalidEntries.length > 0) {
    recommendations.push({
      priority: 'high',
      action: '無効なキャッシュエントリの削除',
      count: analysis.invalidEntries.length,
      description: '破損または無効な形式のキャッシュエントリを削除'
    });
  }
  
  if (analysis.totalCount > 100) {
    recommendations.push({
      priority: 'medium',
      action: 'キャッシュサイズ制限の実装',
      description: 'MAX_CACHE_SIZE制限の強化により将来的な蓄積を防止'
    });
  }
  
  if (analysis.totalSize > 100 * 1024) { // 100KB以上
    recommendations.push({
      priority: 'medium',
      action: 'キャッシュデータの軽量化',
      description: 'キャッシュするデータの軽量化により容量を削減'
    });
  }
  
  return recommendations;
}

/**
 * 古いキャッシュエントリを安全に削除する関数
 */
function cleanupOldCacheEntries(dryRun = true) {
  console.log(`🧹 キャッシュクリーンアップ開始 (${dryRun ? 'DRY RUN' : 'LIVE'})`);
  
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    const cacheKeys = Object.keys(allProps).filter(key => key.startsWith('ss_cache_'));
    
    const now = Date.now();
    const thirtyDaysAgo = now - (30 * 24 * 60 * 60 * 1000);
    
    const keysToDelete = [];
    const invalidKeys = [];
    
    cacheKeys.forEach(key => {
      try {
        const value = allProps[key];
        const parsedValue = JSON.parse(value);
        const timestamp = parsedValue.timestamp;
        
        if (!timestamp || timestamp < thirtyDaysAgo) {
          keysToDelete.push(key);
        }
      } catch (parseError) {
        invalidKeys.push(key);
      }
    });
    
    console.log(`削除対象: 古いエントリ ${keysToDelete.length}件, 無効エントリ ${invalidKeys.length}件`);
    
    if (!dryRun) {
      // 実際の削除を実行
      const allKeysToDelete = [...keysToDelete, ...invalidKeys];
      
      // バッチ削除（GASの制限を考慮して小分けに実行）
      const batchSize = 50;
      for (let i = 0; i < allKeysToDelete.length; i += batchSize) {
        const batch = allKeysToDelete.slice(i, i + batchSize);
        batch.forEach(key => {
          props.deleteProperty(key);
        });
        console.log(`削除済み: ${Math.min(i + batchSize, allKeysToDelete.length)}/${allKeysToDelete.length}`);
        
        // レート制限対策で少し待機
        if (i + batchSize < allKeysToDelete.length) {
          Utilities.sleep(100);
        }
      }
      
      console.log('✅ クリーンアップ完了');
    } else {
      console.log('ℹ️ DRY RUN完了 - 実際の削除は行われませんでした');
    }
    
    return {
      success: true,
      deletedCount: dryRun ? 0 : (keysToDelete.length + invalidKeys.length),
      oldEntriesFound: keysToDelete.length,
      invalidEntriesFound: invalidKeys.length
    };
    
  } catch (error) {
    console.error('❌ クリーンアップ中にエラーが発生しました:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 特定のスプレッドシートIDのキャッシュをすべて削除
 */
function deleteSpreadsheetCache(spreadsheetId, dryRun = true) {
  console.log(`🗑️ スプレッドシートキャッシュ削除: ${spreadsheetId} (${dryRun ? 'DRY RUN' : 'LIVE'})`);
  
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    
    const targetKeys = Object.keys(allProps).filter(key => 
      key.startsWith('ss_cache_') && allProps[key].includes(spreadsheetId)
    );
    
    console.log(`削除対象キー数: ${targetKeys.length}`);
    
    if (!dryRun && targetKeys.length > 0) {
      targetKeys.forEach(key => {
        props.deleteProperty(key);
      });
      console.log('✅ 削除完了');
    }
    
    return {
      success: true,
      foundKeys: targetKeys.length,
      deletedCount: dryRun ? 0 : targetKeys.length
    };
    
  } catch (error) {
    console.error('❌ スプレッドシートキャッシュ削除エラー:', error);
    return { success: false, error: error.message };
  }
}

// newUser_関連機能は削除されました
// マルチテナント対応により、レガシーnewUser_プロパティは不要

// generateNewUserCleanupRecommendation関数も削除（不要）

// cleanupNewUserProperties関数も削除（不要）
// マルチテナント対応により、レガシー機能は全て削除