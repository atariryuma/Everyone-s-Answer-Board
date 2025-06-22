const { JSDOM } = require('jsdom');

function createAnswerCard(data, reactionTypes) {
  const card = document.createElement('div');
  card.className = 'answer-card';
  const active = reactionTypes
    .filter(rt => data.reactions && data.reactions[rt.key] && data.reactions[rt.key].count > 0)
    .map(rt => rt.key);
  if (active.length === 1) {
    if (active[0] === 'LIKE') card.classList.add('reaction-bg-like');
    if (active[0] === 'UNDERSTAND') card.classList.add('reaction-bg-understand');
    if (active[0] === 'CURIOUS') card.classList.add('reaction-bg-curious');
  }
  return card;
}

function applyUpdates(items, reactionTypes) {
  items.forEach(item => {
    const card = document.querySelector(`.answer-card[data-row-index="${item.rowIndex}"]`);
    if (!card) return;
    [
      'reaction-bg-like',
      'reaction-bg-understand',
      'reaction-bg-curious',
      'reaction-bg-like-understand',
      'reaction-bg-like-curious',
      'reaction-bg-understand-curious',
      'reaction-bg-all'
    ].forEach(c => card.classList.remove(c));
    const active = reactionTypes
      .filter(rt => item.reactions && item.reactions[rt.key] && item.reactions[rt.key].count > 0)
      .map(rt => rt.key);
    if (active.length === 1) {
      if (active[0] === 'LIKE') card.classList.add('reaction-bg-like');
      if (active[0] === 'UNDERSTAND') card.classList.add('reaction-bg-understand');
      if (active[0] === 'CURIOUS') card.classList.add('reaction-bg-curious');
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

test('createAnswerCard adds reaction background for single reactions', () => {
  const dom = new JSDOM('<body></body>');
  global.document = dom.window.document;

  const likeCard = createAnswerCard({ reactions: { LIKE: { count: 1 } } }, reactionTypes);
  expect(likeCard.classList.contains('reaction-bg-like')).toBe(true);

  const understandCard = createAnswerCard({ reactions: { UNDERSTAND: { count: 2 } } }, reactionTypes);
  expect(understandCard.classList.contains('reaction-bg-understand')).toBe(true);

  const curiousCard = createAnswerCard({ reactions: { CURIOUS: { count: 3 } } }, reactionTypes);
  expect(curiousCard.classList.contains('reaction-bg-curious')).toBe(true);
});

test('applyUpdates toggles reaction background classes', () => {
  const dom = new JSDOM('<div class="answer-card" data-row-index="1"></div>');
  global.document = dom.window.document;
  const card = dom.window.document.querySelector('.answer-card');

  applyUpdates([{ rowIndex: 1, reactions: { LIKE: { count: 1 } } }], reactionTypes);
  expect(card.classList.contains('reaction-bg-like')).toBe(true);

  applyUpdates([{ rowIndex: 1, reactions: { UNDERSTAND: { count: 2 } } }], reactionTypes);
  expect(card.classList.contains('reaction-bg-like')).toBe(false);
  expect(card.classList.contains('reaction-bg-understand')).toBe(true);

  applyUpdates([{ rowIndex: 1, reactions: {} }], reactionTypes);
  expect(card.classList.contains('reaction-bg-understand')).toBe(false);
  expect(card.classList.contains('reaction-bg-like')).toBe(false);
  expect(card.classList.contains('reaction-bg-curious')).toBe(false);
});
