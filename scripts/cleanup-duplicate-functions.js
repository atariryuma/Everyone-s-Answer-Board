#!/usr/bin/env node

/**
 * @fileoverview 重複関数クリーンアップ・推奨アクション
 * 検出された重複関数の修正方法を提案
 */

const fs = require('fs');
const path = require('path');

/**
 * 重複関数の修正推奨アクション
 */
const CLEANUP_ACTIONS = [
  {
    functionName: 'generateUserUrls',
    issue: '同名だが内容が異なる関数',
    locations: [
      'AdminPanelBackend.gs:571-578',
      'main.gs:721-730'
    ],
    analysis: {
      problem: '同じ機能を実装した関数が2箇所に存在',
      impact: 'ハードコードされたURLと動的URL生成の不整合',
      risk: 'HIGH - URL変更時の修正漏れリスク'
    },
    recommendations: [
      {
        priority: 1,
        action: 'main.gsの動的版を採用（より安全・保守性が高い）',
        reason: 'encodeURIComponent使用、getWebAppUrl()による動的URL生成',
        steps: [
          '1. AdminPanelBackend.gsのgenerateUserUrls関数を削除',
          '2. AdminPanelBackend.gsでmain.gsの関数を参照するように変更',
          '3. ハードコードされたURLを完全削除'
        ]
      }
    ]
  },
  {
    functionName: 'verifyAdminAccess',
    issue: '同名だが内容が異なる関数',
    locations: [
      'Base.gs:406-436',
      'security.gs:126-203'
    ],
    analysis: {
      problem: '認証ロジックが2箇所に分散',
      impact: '認証の不整合、保守の複雑化',
      risk: 'CRITICAL - セキュリティ脆弱性の原因'
    },
    recommendations: [
      {
        priority: 1,
        action: 'security.gsの詳細版を採用（セキュリティ強化版）',
        reason: '3重チェック、構造化ログ、詳細なエラーハンドリング',
        steps: [
          '1. Base.gsのverifyAdminAccess関数を削除',
          '2. Base.gsでsecurity.gsの関数を参照',
          '3. 全ての呼び出し箇所でsecurity.gsの関数を使用'
        ]
      }
    ]
  },
  {
    functionName: 'verifyUserAccess',
    issue: '同名だが内容が異なる関数',
    locations: [
      'Core.gs:451-464', 
      'security.gs:241-253'
    ],
    analysis: {
      problem: 'ユーザー認証ロジックが重複',
      impact: '認証ルールの不整合',
      risk: 'MEDIUM - アクセス制御の混乱'
    },
    recommendations: [
      {
        priority: 1,
        action: 'security.gsの統一版を採用',
        reason: 'App.getAccess().verifyAccess()使用でより統一された設計',
        steps: [
          '1. Core.gsのverifyUserAccess関数を削除',
          '2. Core.gsでsecurity.gsの関数を参照',
          '3. 統一されたアクセス制御パターンに統合'
        ]
      }
    ]
  }
];

/**
 * クリーンアップレポートを生成
 */
function generateCleanupReport() {
  console.log('\n🔧 ==============================================');
  console.log('     重複関数クリーンアップ推奨アクション');
  console.log('==============================================\n');

  CLEANUP_ACTIONS.forEach((action, index) => {
    console.log(`${index + 1}. 🚨 ${action.functionName}`);
    console.log(`   問題: ${action.issue}`);
    console.log(`   場所: ${action.locations.join(', ')}`);
    console.log(`   影響度: ${action.analysis.risk}`);
    console.log(`   問題詳細: ${action.analysis.problem}`);
    
    action.recommendations.forEach((rec, i) => {
      console.log(`\n   💡 推奨アクション ${i + 1}:`);
      console.log(`      ${rec.action}`);
      console.log(`      理由: ${rec.reason}`);
      console.log(`      手順:`);
      rec.steps.forEach((step, j) => {
        console.log(`         ${step}`);
      });
    });
    console.log('');
  });

  console.log('📋 まとめ:');
  console.log(`   - 重複関数: ${CLEANUP_ACTIONS.length}組`);
  console.log(`   - CRITICAL: ${CLEANUP_ACTIONS.filter(a => a.analysis.risk.includes('CRITICAL')).length}個`);
  console.log(`   - HIGH: ${CLEANUP_ACTIONS.filter(a => a.analysis.risk.includes('HIGH')).length}個`);
  console.log(`   - MEDIUM: ${CLEANUP_ACTIONS.filter(a => a.analysis.risk.includes('MEDIUM')).length}個`);
  
  console.log('\n🎯 実行優先順位:');
  console.log('   1. verifyAdminAccess (CRITICAL - セキュリティ)');
  console.log('   2. generateUserUrls (HIGH - URL管理)');
  console.log('   3. verifyUserAccess (MEDIUM - アクセス制御)');
  
  console.log('\n⚠️  注意事項:');
  console.log('   - 修正前に必ずバックアップを取得');
  console.log('   - 一つずつ段階的に修正');
  console.log('   - 修正後はテスト実行で動作確認');
  console.log('   - 特にverifyAdminAccessはセキュリティに直結');
}

/**
 * メイン実行
 */
function main() {
  generateCleanupReport();
  
  console.log('\n✅ 重複関数クリーンアップ推奨アクションの生成完了！');
  
  return CLEANUP_ACTIONS;
}

// 直接実行時のみmain実行
if (require.main === module) {
  main();
}

module.exports = {
  CLEANUP_ACTIONS,
  generateCleanupReport,
  main
};