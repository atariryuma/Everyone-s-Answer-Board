const fs = require('fs');
const vm = require('vm');

describe('ゲストアクセステスト', () => {
  // ConfigurationManagerとAccessControllerは同じファイルに統合
  const accessControllerCode = fs.readFileSync('src/AccessController.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: jest.fn((key) => {
            if (key === 'user_config_public_user') {
              return JSON.stringify({
                tenantId: 'public_user',
                ownerEmail: 'owner@example.com',
                title: 'Public Answer Board',
                isPublic: true,
                allowAnonymous: true,
                columns: [
                  { name: 'question', label: '質問', type: 'text', required: true },
                  { name: 'answer', label: '回答', type: 'textarea', required: false }
                ]
              });
            }
            if (key === 'user_config_private_user') {
              return JSON.stringify({
                tenantId: 'private_user',
                ownerEmail: 'owner@example.com',
                title: 'Private Answer Board',
                isPublic: false,
                allowAnonymous: false
              });
            }
            return null;
          })
        })
      },
      CacheService: {
        getScriptCache: () => ({
          get: jest.fn(() => null),
          put: jest.fn(),
          remove: jest.fn()
        })
      }
    };
    vm.createContext(context);
    vm.runInContext(accessControllerCode, context);
  });

  describe('公開ボードのゲストアクセス', () => {
    test('ゲストが公開ボードに閲覧アクセスできる', () => {
      const result = context.accessController.verifyAccess('public_user', 'view', null);
      
      expect(result.allowed).toBe(true);
      expect(result.userType).toBe('guest');
      expect(result.config.title).toBe('Public Answer Board');
      expect(result.config.isPublic).toBeUndefined(); // publicConfigでは内部設定は隠される
    });

    test('ゲストが公開ボードの編集はできない', () => {
      const result = context.accessController.verifyAccess('public_user', 'edit', null);
      
      expect(result.allowed).toBe(false);
      expect(result.userType).toBe('unauthorized');
    });

    test('所有者は自分のボードに全アクセス権限がある', () => {
      const viewResult = context.accessController.verifyAccess('public_user', 'view', 'owner@example.com');
      const editResult = context.accessController.verifyAccess('public_user', 'edit', 'owner@example.com');
      const adminResult = context.accessController.verifyAccess('public_user', 'admin', 'owner@example.com');
      
      expect(viewResult.allowed).toBe(true);
      expect(viewResult.userType).toBe('owner');
      expect(editResult.allowed).toBe(true);
      expect(editResult.userType).toBe('owner');
      expect(adminResult.allowed).toBe(true);
      expect(adminResult.userType).toBe('owner');
    });
  });

  describe('プライベートボードのアクセス制御', () => {
    test('ゲストがプライベートボードにアクセスできない', () => {
      const result = context.accessController.verifyAccess('private_user', 'view', null);
      
      expect(result.allowed).toBe(false);
      expect(result.userType).toBe('private');
    });

    test('認証ユーザーもプライベートボードにアクセスできない', () => {
      const result = context.accessController.verifyAccess('private_user', 'view', 'other@example.com');
      
      expect(result.allowed).toBe(false);
      expect(result.userType).toBe('private');
    });

    test('所有者のみプライベートボードにアクセスできる', () => {
      const result = context.accessController.verifyAccess('private_user', 'view', 'owner@example.com');
      
      expect(result.allowed).toBe(true);
      expect(result.userType).toBe('owner');
    });
  });

  describe('存在しないユーザーのアクセス制御', () => {
    test('存在しないユーザーIDでアクセス拒否', () => {
      const result = context.accessController.verifyAccess('non_existent_user', 'view', null);
      
      expect(result.allowed).toBe(false);
      expect(result.userType).toBe('not_found');
    });
  });

  describe('ConfigurationManager統合テスト', () => {
    test('公開設定の取得が正しく動作する', () => {
      const publicConfig = context.configManager.getPublicConfig('public_user');
      
      expect(publicConfig).not.toBeNull();
      expect(publicConfig.tenantId).toBe('public_user');
      expect(publicConfig.title).toBe('Public Answer Board');
      expect(publicConfig.allowAnonymous).toBe(true);
      expect(publicConfig.columns).toHaveLength(2);
    });

    test('プライベートボードの公開設定がnullになる', () => {
      const publicConfig = context.configManager.getPublicConfig('private_user');
      
      expect(publicConfig).toBeNull();
    });
  });
});