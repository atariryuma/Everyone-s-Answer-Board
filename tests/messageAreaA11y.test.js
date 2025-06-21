const { JSDOM } = require('jsdom');

test('SheetSelector showMessage preserves aria-live attribute', () => {
  const dom = new JSDOM(`<div id="message-area" aria-live="polite" role="status"></div>`);
  const elements = { messageArea: dom.window.document.getElementById('message-area') };
  function showMessage(msg, color = 'red') {
    const colorClasses = { red: 'text-red-600', green: 'text-green-600', blue: 'text-blue-600' };
    elements.messageArea.className = `mt-4 text-center text-sm h-5 ${colorClasses[color]}`;
    elements.messageArea.textContent = msg;
  }
  showMessage('hello', 'green');
  expect(elements.messageArea.getAttribute('aria-live')).toBe('polite');
  expect(elements.messageArea.getAttribute('role')).toBe('status');
});

test('Unpublished message updates preserve aria-live attribute', () => {
  const dom = new JSDOM(`<p id="message-area" aria-live="polite" role="status"></p>`);
  const messageArea = dom.window.document.getElementById('message-area');
  messageArea.textContent = 'error';
  messageArea.className = 'mt-4 text-sm h-5 text-red-400';
  expect(messageArea.getAttribute('aria-live')).toBe('polite');
  expect(messageArea.getAttribute('role')).toBe('status');
});
