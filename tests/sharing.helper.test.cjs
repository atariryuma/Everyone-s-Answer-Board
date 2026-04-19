const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadSharingHelperContext(overrides = {}) {
  const fileActions = [];
  const file = overrides.file || {
    getSharingAccess: () => 'ANYONE',
    getSharingPermission: () => 'VIEW',
    setSharing: (access, permission) => { fileActions.push({ access, permission }); }
  };

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    DriveApp: {
      Access: {
        DOMAIN: 'DOMAIN',
        DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK',
        ANYONE: 'ANYONE'
      },
      Permission: {
        EDIT: 'EDIT',
        VIEW: 'VIEW'
      },
      getFileById: overrides.getFileById || (() => file)
    },
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/SharingHelper.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'SharingHelper.js' });
  return { context, fileActions, file };
}

test('setupDomainWideSharing: configures DOMAIN_WITH_LINK + EDIT on a non-shared file', () => {
  const { context, fileActions } = loadSharingHelperContext();

  const result = context.setupDomainWideSharing('ss-1', 'user@example.com');

  assert.equal(result.success, true);
  assert.equal(fileActions.length, 1);
  assert.equal(fileActions[0].access, 'DOMAIN_WITH_LINK');
  assert.equal(fileActions[0].permission, 'EDIT');
});

test('setupDomainWideSharing: no-op when already DOMAIN + EDIT', () => {
  const { context, fileActions } = loadSharingHelperContext({
    file: {
      getSharingAccess: () => 'DOMAIN',
      getSharingPermission: () => 'EDIT',
      setSharing: () => {}
    }
  });

  const result = context.setupDomainWideSharing('ss-1', 'user@example.com');

  assert.equal(result.success, true);
  assert.match(result.message, /Already/);
  assert.equal(fileActions.length, 0);
});

test('setupDomainWideSharing: no-op when already DOMAIN_WITH_LINK + EDIT', () => {
  let setCalls = 0;
  const { context } = loadSharingHelperContext({
    file: {
      getSharingAccess: () => 'DOMAIN_WITH_LINK',
      getSharingPermission: () => 'EDIT',
      setSharing: () => { setCalls += 1; }
    }
  });

  const result = context.setupDomainWideSharing('ss-1', 'user@example.com');
  assert.equal(result.success, true);
  assert.equal(setCalls, 0);
});

test('setupDomainWideSharing: upgrades permission when access is DOMAIN but permission is VIEW', () => {
  const { context, fileActions } = loadSharingHelperContext({
    file: {
      getSharingAccess: () => 'DOMAIN',
      getSharingPermission: () => 'VIEW',
      setSharing: (a, p) => { fileActions.push({ access: a, permission: p }); }
    }
  });

  const result = context.setupDomainWideSharing('ss-1', 'user@example.com');
  // Even though access is DOMAIN, permission is VIEW → re-apply
  assert.equal(result.success, true);
  assert.equal(fileActions.length, 1);
});

test('setupDomainWideSharing: rejects email without @domain', () => {
  const { context } = loadSharingHelperContext();
  assert.throws(
    () => context.setupDomainWideSharing('ss-1', 'no-at-sign'),
    /Invalid email format/
  );
});

test('setupDomainWideSharing: propagates DriveApp errors', () => {
  const { context } = loadSharingHelperContext({
    getFileById: () => { throw new Error('file not found'); }
  });
  assert.throws(
    () => context.setupDomainWideSharing('ss-ghost', 'user@example.com'),
    /file not found/
  );
});
