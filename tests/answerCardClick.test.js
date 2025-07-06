const { JSDOM } = require('jsdom');

function setupEventDelegation(app) {
  if (app.eventDelegationSetup) return;
  app.handlers.onAnswersContainerClick = (e) => {
    const answerCard = e.target.closest('.answer-card');
    if (!answerCard || answerCard.classList.contains('hidden-card')) {
      return;
    }
    const rowIndex = answerCard.dataset.rowIndex;
    if (!rowIndex) return;

    const reactionBtn = e.target.closest('.reaction-btn');
    if (reactionBtn) {
      e.stopPropagation();
      if (!reactionBtn.disabled) {
        app.handleReaction(rowIndex, reactionBtn.dataset.reaction);
      }
      return;
    }

    const highlightBtn = e.target.closest('.highlight-btn');
    if (highlightBtn) {
      e.stopPropagation();
      if (!highlightBtn.disabled) {
        app.handleHighlight(rowIndex);
      }
      return;
    }

    app.showAnswerModal(rowIndex);
  };
  if (app.elements.answersContainer) {
    app.elements.answersContainer.addEventListener(
      'click',
      app.handlers.onAnswersContainerClick
    );
    app.eventDelegationSetup = true;
  }
}

describe('answer card click handling', () => {
  let dom;
  let app;

  beforeEach(() => {
    dom = new JSDOM(`<!DOCTYPE html><body>
      <div id="loading-overlay" class="loading-overlay hidden"></div>
      <div id="answers">
        <div class="answer-card" data-row-index="1">
          <button class="reaction-btn" data-row-index="1" data-reaction="LIKE"></button>
          <button class="highlight-btn" data-row-index="1"></button>
        </div>
      </div>
    </body>`);
    global.window = dom.window;
    global.document = dom.window.document;

    app = {
      elements: { answersContainer: document.getElementById('answers') },
      showAnswerModal: jest.fn(),
      handleReaction: jest.fn(),
      handleHighlight: jest.fn(),
      handlers: {},
      eventDelegationSetup: false
    };
    setupEventDelegation(app);
  });

  afterEach(() => {
    delete global.window;
    delete global.document;
  });

  test('card body click triggers showAnswerModal', () => {
    const card = document.querySelector('.answer-card');
    card.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
    expect(app.showAnswerModal).toHaveBeenCalledWith('1');
  });

  test('reaction button triggers handleReaction', () => {
    const btn = document.querySelector('.reaction-btn');
    btn.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
    expect(app.handleReaction).toHaveBeenCalledWith('1', 'LIKE');
    expect(app.showAnswerModal).not.toHaveBeenCalled();
  });

  test('highlight button triggers handleHighlight', () => {
    const btn = document.querySelector('.highlight-btn');
    btn.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
    expect(app.handleHighlight).toHaveBeenCalledWith('1');
    expect(app.showAnswerModal).not.toHaveBeenCalled();
  });

  test('hidden card does not trigger actions', () => {
    const card = document.querySelector('.answer-card');
    card.classList.add('hidden-card');
    card.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
    expect(app.showAnswerModal).not.toHaveBeenCalled();
    expect(app.handleReaction).not.toHaveBeenCalled();
    expect(app.handleHighlight).not.toHaveBeenCalled();
  });

});
