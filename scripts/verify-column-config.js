#!/usr/bin/env node

/**
 * @fileoverview 列設定検証スクリプト
 * データベースに保存されている列設定とスプレッドシート構造を検証
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// GAS環境のモック
function createMockGASEnvironment() {
  const mockContext = {
    console,
    
    // Properties Service Mock
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => {
          const mockProperties = {
            'SERVICE_ACCOUNT_CREDS': JSON.stringify({
              type: 'service_account',
              project_id: 'mock-project'
            }),
            'DB_SPREADSHEET_ID': '1CHELsGNUowfzcfK0s6fi04WBCA9MKyonvTIQAL0HMG0'
          };
          return mockProperties[key] || null;
        }
      }),
      getUserProperties: () => ({
        getProperty: (key) => {
          const mockUserProps = {
            'CURRENT_USER_ID': '6d222374-e377-44fa-ac72-4acb8bd80e08',
            'CURRENT_SPREADSHEET_ID': '1vgrxDsH-aigRbj4_5vbK0_s87S4o5R3Sk5BAuiUBa5M'
          };
          return mockUserProps[key] || null;
        }
      })
    },
    
    // Session Mock
    Session: {
      getActiveUser: () => ({
        getEmail: () => '35t22@naha-okinawa.ed.jp'
      })
    },
    
    // SpreadsheetApp Mock
    SpreadsheetApp: {
      openById: (id) => {
        console.log(`[MOCK] Opening spreadsheet: ${id}`);
        return {
          getName: () => 'Mock Spreadsheet',
          getSheetByName: (name) => ({
            getName: () => name,
            getLastRow: () => 10,
            getLastColumn: () => 5,
            getRange: (row, col, numRows, numCols) => ({
              getValues: () => {
                if (row === 1) {
                  // ヘッダー行をモック
                  return [['タイムスタンプ', '質問', '回答', '理由', 'クラス']];
                }
                return [['2025/08/30 10:00:00', 'サンプル質問', 'サンプル回答', 'サンプル理由', '1年A組']];
              }
            })
          }),
          getSheets: () => [
            { getName: () => 'フォームの回答 1' }
          ]
        };
      }
    },

    // Cache Mock
    CacheService: {
      getScriptCache: () => ({
        get: () => null,
        put: () => {},
        remove: () => {}
      })
    },

    // Utilities Mock
    Utilities: {
      sleep: () => {}
    }
  };

  return mockContext;
}

async function verifyColumnConfiguration() {
  console.log('🔍 列設定検証を開始...');
  
  try {
    // GASファイルを読み込み
    const mainCode = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf-8');
    const databaseCode = fs.readFileSync(path.join(__dirname, '../src/database.gs'), 'utf-8');
    const coreCode = fs.readFileSync(path.join(__dirname, '../src/Core.gs'), 'utf-8');
    
    // モック環境を作成
    const context = createMockGASEnvironment();
    vm.createContext(context);
    
    // コードを実行
    vm.runInContext(mainCode, context);
    vm.runInContext(databaseCode, context);
    vm.runInContext(coreCode, context);
    
    console.log('✅ GAS環境の初期化完了');
    
    // 1. ユーザー情報の取得テスト
    console.log('\n📋 Phase 1: ユーザー情報確認');
    try {
      const userId = '6d222374-e377-44fa-ac72-4acb8bd80e08';
      console.log(`ユーザーID: ${userId}`);
      
      // ユーザー情報取得をシミュレート
      if (context.getUserInfo) {
        const userInfo = context.getUserInfo(userId);
        console.log('ユーザー情報:', JSON.stringify(userInfo, null, 2));
      }
      
      if (context.findUserById) {
        const userInfo = context.findUserById(userId);
        console.log('findUserById結果:', JSON.stringify(userInfo, null, 2));
      }
      
    } catch (error) {
      console.warn('ユーザー情報取得エラー:', error.message);
    }
    
    // 2. スプレッドシート構造確認
    console.log('\n📊 Phase 2: スプレッドシート構造確認');
    try {
      const spreadsheetId = '1vgrxDsH-aigRbj4_5vbK0_s87S4o5R3Sk5BAuiUBa5M';
      const sheetName = 'フォームの回答 1';
      
      console.log(`対象スプレッドシート: ${spreadsheetId}`);
      console.log(`対象シート: ${sheetName}`);
      
      // ヘッダー情報取得テスト
      if (context.getHeaderIndices) {
        const headers = context.getHeaderIndices(spreadsheetId, sheetName);
        console.log('ヘッダーインデックス:', JSON.stringify(headers, null, 2));
      }
      
    } catch (error) {
      console.warn('スプレッドシート構造確認エラー:', error.message);
    }
    
    // 3. 列設定確認
    console.log('\n🔧 Phase 3: 列設定確認');
    try {
      if (context.getSheetConfigForUser) {
        const config = context.getSheetConfigForUser();
        console.log('現在の列設定:', JSON.stringify(config, null, 2));
      }
      
      if (context.getAppSettingsForUser) {
        const settings = context.getAppSettingsForUser();
        console.log('アプリ設定:', JSON.stringify(settings, null, 2));
      }
      
    } catch (error) {
      console.warn('列設定確認エラー:', error.message);
    }
    
    // 4. データ取得テスト
    console.log('\n📥 Phase 4: データ取得テスト');
    try {
      if (context.getPublishedSheetData) {
        const data = context.getPublishedSheetData('フォームの回答 1');
        console.log('公開データ:', JSON.stringify(data, null, 2));
      }
      
    } catch (error) {
      console.warn('データ取得テスト エラー:', error.message);
    }
    
    console.log('\n✅ 列設定検証完了');
    
  } catch (error) {
    console.error('❌ 検証エラー:', error);
    throw error;
  }
}

// 実行
if (require.main === module) {
  verifyColumnConfiguration()
    .then(() => {
      console.log('\n🎉 検証処理が完了しました');
    })
    .catch(error => {
      console.error('\n💥 検証処理でエラーが発生:', error.message);
      process.exit(1);
    });
}

module.exports = { verifyColumnConfiguration };