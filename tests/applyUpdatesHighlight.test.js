const { JSDOM } = require('jsdom');

function getIcon() {
  return '<svg></svg>';
}


function applyUpdates(items) {
  items.forEach(item => {
    const card = document.querySelector(`.answer-card[data-row-index="${item.rowIndex}"]`);
    if (!card) return;
    const preview = card.querySelector('.answer-preview');
    card.classList.toggle('highlighted', item.highlight);
    const btn = card.querySelector('.highlight-btn');
    if (btn) {
      btn.classList.toggle('liked', item.highlight);
      btn.setAttribute('aria-pressed', String(item.highlight));
      const label = item.highlight ? 'ハイライトを解除する' : 'ハイライトする';
      btn.setAttribute('aria-label', label);
    }
    let badge = card.querySelector('.highlight-badge');
    if (item.highlight && !badge) {
      badge = document.createElement('span');
      badge.className = 'highlight-badge';
      badge.innerHTML = getIcon('star', 'w-6 h-6', true);
      if (preview) preview.insertBefore(badge, preview.firstChild);
    } else if (!item.highlight && badge) {
      badge.remove();
    }
  });
}

afterEach(() => {
  delete global.document;
});

test('highlight badge toggles based on item state', () => {
  const dom = new JSDOM(`
    <div class="answer-card" data-row-index="1">
      <div class="answer-preview"></div>
      <button class="highlight-btn like-btn" aria-pressed="false"></button>
    </div>
  `);
  global.document = dom.window.document;

  const card = dom.window.document.querySelector('.answer-card');
  const preview = card.querySelector('.answer-preview');
  const btn = card.querySelector('.highlight-btn');

  applyUpdates([{ rowIndex: 1, highlight: true }]);
  expect(card.classList.contains('highlighted')).toBe(true);
  expect(preview.querySelector('.highlight-badge')).not.toBeNull();
  expect(btn.classList.contains('liked')).toBe(true);
  expect(btn.getAttribute('aria-pressed')).toBe('true');

  applyUpdates([{ rowIndex: 1, highlight: false }]);
  expect(card.classList.contains('highlighted')).toBe(false);
  expect(preview.querySelector('.highlight-badge')).toBeNull();
  expect(btn.classList.contains('liked')).toBe(false);
  expect(btn.getAttribute('aria-pressed')).toBe('false');
});
