const { JSDOM } = require('jsdom');

function updateReactionButtonUI(rowIndex, reaction, count, reacted) {
  document.querySelectorAll(`[data-row-index="${rowIndex}"][data-reaction="${reaction}"]`).forEach(btn => {
    const countEl = btn.querySelector('.reaction-count');
    if (countEl) {
      countEl.textContent = count;
    }
    btn.classList.toggle('liked', reacted);
  });
}

test('updateReactionButtonUI applies solid icon when reacted', () => {
  const dom = new JSDOM(`
    <button class="reaction-btn like-btn" data-row-index="1" data-reaction="LIKE">
      <svg class="w-5 h-5"></svg>
      <span class="reaction-count">0</span>
    </button>
  `);
  global.document = dom.window.document;

  const btn = dom.window.document.querySelector('button');

  updateReactionButtonUI(1, 'LIKE', 4, true);
  expect(btn.classList.contains('liked')).toBe(true);
  expect(btn.querySelector('.reaction-count').textContent).toBe('4');

  updateReactionButtonUI(1, 'LIKE', 1, false);
  expect(btn.classList.contains('liked')).toBe(false);
  expect(btn.querySelector('.reaction-count').textContent).toBe('1');

  delete global.document;
});
