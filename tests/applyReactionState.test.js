const { JSDOM } = require('jsdom');

function applyReactionState(answers) {
  const savedReactions = this.loadReactionState();
  let modified = false;

  answers.forEach(answer => {
    if (!savedReactions[answer.rowIndex]) {
      savedReactions[answer.rowIndex] = {};
    }
    const savedReaction = savedReactions[answer.rowIndex];
    if (answer.reactions) {
      Object.keys(answer.reactions).forEach(type => {
        const info = answer.reactions[type];
        if (!info) return;
        const serverReacted = !!info.reacted;
        const localReacted = !!savedReaction[type];

        info.reacted = serverReacted;

        if (serverReacted) {
          if (!localReacted) {
            if (!savedReactions[answer.rowIndex]) {
              savedReactions[answer.rowIndex] = {};
            }
            savedReactions[answer.rowIndex][type] = {
              reacted: true,
              timestamp: new Date().toISOString()
            };
            modified = true;
          }
        } else if (localReacted) {
          delete savedReaction[type];
          modified = true;
        }

        if (Object.keys(savedReaction).length === 0) {
          delete savedReactions[answer.rowIndex];
        }
      });
    }
  });

  if (modified) {
    localStorage.setItem(this.reactionStorageKey, JSON.stringify(savedReactions));
  }
}

describe('applyReactionState', () => {
  let window, app;
  beforeEach(() => {
    const dom = new JSDOM('', { url: 'http://localhost' });
    window = dom.window;
    global.window = window;
    global.document = window.document;
    global.localStorage = window.localStorage;
    app = {
      reactionStorageKey: 'reactions_test_Sheet1',
      loadReactionState() {
        return JSON.parse(localStorage.getItem(this.reactionStorageKey) || '{}');
      },
      applyReactionState
    };
    localStorage.clear();
  });

  test('clears stale local reactions when server shows none', () => {
    localStorage.setItem(app.reactionStorageKey, JSON.stringify({
      1: { LIKE: { reacted: true, timestamp: 't' } }
    }));

    const answers = [{ rowIndex: 1, reactions: { LIKE: { reacted: false, count: 0 } } }];
    app.applyReactionState(answers);

    expect(answers[0].reactions.LIKE.reacted).toBe(false);
    expect(localStorage.getItem(app.reactionStorageKey)).toBe('{}');
  });

  test('stores reaction when server reacted true', () => {
    const answers = [{ rowIndex: 2, reactions: { LIKE: { reacted: true, count: 1 } } }];
    app.applyReactionState(answers);

    const stored = JSON.parse(localStorage.getItem(app.reactionStorageKey) || '{}');
    expect(stored['2']?.LIKE?.reacted).toBe(true);
    expect(answers[0].reactions.LIKE.reacted).toBe(true);
  });
});
