const { JSDOM } = require('jsdom');

function createAnswerCard(data, reactionTypes) {
  const card = document.createElement('div');
  card.className = 'answer-card';
  const total = reactionTypes.reduce((sum, rt) => sum + (data.reactions?.[rt.key]?.count || 0), 0);
  if (total >= 10) {
    card.classList.add('reaction-border-3');
  } else if (total >= 5) {
    card.classList.add('reaction-border-2');
  } else if (total > 0) {
    card.classList.add('reaction-border-1');
  }
  return card;
}

function applyUpdates(items, reactionTypes) {
  items.forEach(item => {
    const card = document.querySelector(`.answer-card[data-row-index="${item.rowIndex}"]`);
    if (!card) return;
    const total = reactionTypes.reduce((sum, rt) => sum + (item.reactions?.[rt.key]?.count || 0), 0);
    ['reaction-border-1','reaction-border-2','reaction-border-3'].forEach(c => card.classList.remove(c));
    if (total >= 10) {
      card.classList.add('reaction-border-3');
    } else if (total >= 5) {
      card.classList.add('reaction-border-2');
    } else if (total > 0) {
      card.classList.add('reaction-border-1');
    }
  });
}

const reactionTypes = [
  { key: 'LIKE' },
  { key: 'UNDERSTAND' },
  { key: 'CURIOUS' }
];

afterEach(() => {
  delete global.document;
});

test('createAnswerCard adds border class based on total reactions', () => {
  const dom = new JSDOM('<body></body>');
  global.document = dom.window.document;

  const card1 = createAnswerCard({ reactions: { LIKE: { count: 1 } } }, reactionTypes);
  expect(card1.classList.contains('reaction-border-1')).toBe(true);

  const card2 = createAnswerCard({ reactions: { LIKE: { count: 5 } } }, reactionTypes);
  expect(card2.classList.contains('reaction-border-2')).toBe(true);

  const card3 = createAnswerCard({ reactions: { LIKE: { count: 10 } } }, reactionTypes);
  expect(card3.classList.contains('reaction-border-3')).toBe(true);
});

test('applyUpdates toggles border classes', () => {
  const dom = new JSDOM('<div class="answer-card" data-row-index="1"></div>');
  global.document = dom.window.document;

  applyUpdates([{ rowIndex: 1, reactions: { LIKE: { count: 6 } } }], reactionTypes);
  const card = dom.window.document.querySelector('.answer-card');
  expect(card.classList.contains('reaction-border-2')).toBe(true);

  applyUpdates([{ rowIndex: 1, reactions: { LIKE: { count: 11 } } }], reactionTypes);
  expect(card.classList.contains('reaction-border-3')).toBe(true);

  applyUpdates([{ rowIndex: 1, reactions: { LIKE: { count: 0 } } }], reactionTypes);
  expect(card.classList.contains('reaction-border-1')).toBe(false);
  expect(card.classList.contains('reaction-border-2')).toBe(false);
  expect(card.classList.contains('reaction-border-3')).toBe(false);
});
