const fs = require('fs');
const vm = require('vm');

describe('AuthorizationService', () => {
  let context;

  beforeEach(() => {
    // 必要なモックをセットアップ
    context = {
      Session: { getActiveUser: jest.fn() },
      isDeployUser: jest.fn(),
      findUserById: jest.fn(),
      console: { log: jest.fn(), warn: jest.fn() },
    };
    const code = fs.readFileSync('src/AuthorizationService.gs', 'utf8');
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('システム管理者はどのユーザーのデータにもアクセスできる', () => {
    context.isDeployUser.mockReturnValue(true);
    context.Session.getActiveUser.mockReturnValue({ getEmail: () => 'admin@example.com' });
    expect(() => context.AuthorizationService.verifyUserAccess('any_user_id')).not.toThrow();
  });

  test('一般ユーザーは自分のデータにのみアクセスできる', () => {
    context.isDeployUser.mockReturnValue(false);
    context.Session.getActiveUser.mockReturnValue({ getEmail: () => 'user@example.com' });
    context.findUserById.mockReturnValue({ adminEmail: 'user@example.com' });

    expect(() => context.AuthorizationService.verifyUserAccess('my_user_id')).not.toThrow();
  });

  test('一般ユーザーが他人のデータにアクセスしようとするとエラーをスローする', () => {
    context.isDeployUser.mockReturnValue(false);
    context.Session.getActiveUser.mockReturnValue({ getEmail: () => 'user@example.com' });
    context.findUserById.mockReturnValue({ adminEmail: 'another_user@example.com' });

    expect(() => context.AuthorizationService.verifyUserAccess('another_user_id')).toThrow('権限エラー');
  });

  test('同一ドメインならボードにアクセスできる', () => {
    context.Session.getActiveUser.mockReturnValue({ getEmail: () => 'member@example.com' });
    expect(() => context.AuthorizationService.verifyBoardAccess('owner@example.com')).not.toThrow();
  });

  test('異なるドメインならボードアクセスエラー', () => {
    context.Session.getActiveUser.mockReturnValue({ getEmail: () => 'user@other.com' });
    expect(() => context.AuthorizationService.verifyBoardAccess('owner@example.com')).toThrow('権限エラー');
  });
});
