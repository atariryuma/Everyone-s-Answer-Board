const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createContentServiceStub() {
  return {
    MimeType: { JSON: 'application/json' },
    createTextOutput(body) {
      return {
        body,
        mimeType: null,
        setMimeType(mimeType) {
          this.mimeType = mimeType;
          return this;
        }
      };
    }
  };
}

function loadMainContext(overrides = {}) {
  const context = {
    console: {
      log: () => {},
      warn: () => {},
      error: () => {}
    },
    ContentService: createContentServiceStub(),
    createSuccessResponse: (message, data, extra) => ({ success: true, message, ...(data && { data }), ...extra }),
    createAuthError: () => ({ success: false, message: 'auth required' }),
    createUserNotFoundError: () => ({ success: false, message: 'user not found' }),
    createErrorResponse: (message, data, extra) => ({ success: false, message, error: message, ...extra }),
    createExceptionResponse: (error) => ({ success: false, message: error.message, error: error.message }),
    createAdminRequiredError: () => ({ success: false, message: 'admin required' }),
    Session: {
      getActiveUser: () => ({
        getEmail: () => 'teacher@example.com'
      })
    },
    evaluateDomainRestriction: () => ({ allowed: true }),
    findUserByEmail: () => ({ userId: 'user-1' }),
    getUserSheetData: () => ({ rows: [{ id: 1 }] }),
    addReaction: (userId, rowId, reactionType) => ({ success: true, userId, rowId, reactionType }),
    toggleHighlight: (userId, rowId) => ({ success: true, userId, rowId }),
    SystemController: {
      publishApp: (config) => ({ success: true, config })
    },
    getCachedProperty: (key) => key === 'ADMIN_API_KEY' ? 'test-secret-key' : key === 'ADMIN_EMAIL' ? 'admin@example.com' : null,
    dispatchAdminOperation: (operation, params) => {
      if (operation === 'getUsers') return { success: true, users: [] };
      if (operation === 'getAppStatus') return { success: true, status: 'enabled' };
      return { success: false, message: `Unknown operation: ${operation}`, error: 'UNKNOWN_OPERATION' };
    },
    timingSafeEqual: (a, b) => typeof a === 'string' && typeof b === 'string' && a === b
  };

  Object.assign(context, overrides);

  // Stub setCachedProperty AFTER overrides so it reads the final PropertiesService.
  if (!context.setCachedProperty) {
    context.setCachedProperty = (key, value) => {
      try {
        context.PropertiesService.getScriptProperties().setProperty(key, value);
      } catch (_) { /* stub best-effort */ }
    };
  }

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/main.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'main.js' });
  return context;
}

function parseResponse(output) {
  assert.ok(output);
  assert.equal(output.mimeType, 'application/json');
  return JSON.parse(output.body);
}

function createPostEvent(payload, contentType = 'application/json') {
  const contents = typeof payload === 'string' ? payload : JSON.stringify(payload);
  return {
    postData: {
      contents,
      type: contentType
    }
  };
}

test('doPost: body missing should return MISSING_POST_BODY', () => {
  const context = loadMainContext();
  const response = parseResponse(context.doPost(undefined));
  assert.equal(response.error, 'MISSING_POST_BODY');
});

test('doPost: non-JSON content type should be rejected', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'getData' }, 'text/plain');
  const response = parseResponse(context.doPost(event));
  assert.equal(response.error, 'UNSUPPORTED_CONTENT_TYPE');
});

test('doPost: invalid JSON should return JSON_PARSE_ERROR', () => {
  const context = loadMainContext();
  const event = createPostEvent('{"action":', 'application/json');
  const response = parseResponse(context.doPost(event));
  assert.equal(response.error, 'JSON_PARSE_ERROR');
});

test('doPost: payload must be object', () => {
  const context = loadMainContext();
  const event = createPostEvent(['not-object'], 'application/json');
  const response = parseResponse(context.doPost(event));
  assert.equal(response.error, 'INVALID_REQUEST_PAYLOAD');
});

test('doPost: unknown action should be rejected', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'unknownAction' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.error, 'UNKNOWN_ACTION');
});

test('doPost: addReaction requires rowId and reactionType', () => {
  const context = loadMainContext();
  const missingRowEvent = createPostEvent({ action: 'addReaction', userId: 'u1' });
  const missingRowResponse = parseResponse(context.doPost(missingRowEvent));
  assert.equal(missingRowResponse.success, false);
  assert.match(missingRowResponse.message, /row/i);

  const missingReactionEvent = createPostEvent({ action: 'addReaction', userId: 'u1', rowId: 'row_2' });
  const missingReactionResponse = parseResponse(context.doPost(missingReactionEvent));
  assert.equal(missingReactionResponse.success, false);
  assert.match(missingReactionResponse.message, /Reaction type required/);
});

test('doPost: publishApp requires config object', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'publishApp' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.message, 'Publish config is required');
});

test('doPost: getData success path', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'getData' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, true);
  assert.deepEqual(response.data, { rows: [{ id: 1 }] });
});

// --- adminApi tests ---

// --- setupApiKey tests ---

test('doPost: setupApiKey sets key when admin and not configured', () => {
  const props = {};
  const context = loadMainContext({
    Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => props[k] || null,
        setProperty: (k, v) => { props[k] = v; }
      })
    }
  });
  const event = createPostEvent({ action: 'setupApiKey', apiKey: 'my-secret-key-1234' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, true);
  assert.equal(props['ADMIN_API_KEY'], 'my-secret-key-1234');
});

test('doPost: setupApiKey rejects when already configured', () => {
  const context = loadMainContext({
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => 'existing-key'
      })
    }
  });
  const event = createPostEvent({ action: 'setupApiKey', apiKey: 'new-key-12345678' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'ALREADY_CONFIGURED');
});

test('doPost: setupApiKey rejects non-admin', () => {
  const context = loadMainContext({
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => null
      })
    }
  });
  const event = createPostEvent({ action: 'setupApiKey', apiKey: 'long-enough-key-1234' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'ADMIN_REQUIRED');
});

test('doPost: setupApiKey rejects short key for admin', () => {
  const context = loadMainContext({
    Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => null
      })
    }
  });
  const event = createPostEvent({ action: 'setupApiKey', apiKey: 'short' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'INVALID_KEY_FORMAT');
});

// --- adminApi tests ---

test('doPost: adminApi rejects invalid API key', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'adminApi', apiKey: 'wrong-key', operation: 'getUsers' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'INVALID_API_KEY');
});

test('doPost: adminApi rejects missing API key', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'adminApi', operation: 'getUsers' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'INVALID_API_KEY');
});

test('doPost: adminApi returns error when ADMIN_API_KEY not configured', () => {
  const context = loadMainContext({
    getCachedProperty: () => null
  });
  const event = createPostEvent({ action: 'adminApi', apiKey: 'any-key', operation: 'getUsers' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'API_KEY_NOT_CONFIGURED');
});

test('doPost: adminApi dispatches with valid key', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'adminApi', apiKey: 'test-secret-key', operation: 'getUsers' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, true);
  assert.deepEqual(response.users, []);
});

test('doPost: adminApi forwards unknown operation error', () => {
  const context = loadMainContext();
  const event = createPostEvent({ action: 'adminApi', apiKey: 'test-secret-key', operation: 'badOp' });
  const response = parseResponse(context.doPost(event));
  assert.equal(response.success, false);
  assert.equal(response.error, 'UNKNOWN_OPERATION');
});
