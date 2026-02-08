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
    createAuthError: () => ({ success: false, message: 'auth required' }),
    createUserNotFoundError: () => ({ success: false, message: 'user not found' }),
    createErrorResponse: (message) => ({ success: false, message, error: message }),
    createExceptionResponse: (error) => ({ success: false, message: error.message, error: error.message }),
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
    }
  };

  Object.assign(context, overrides);

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
