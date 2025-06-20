const { JSDOM } = require('jsdom');

test('polling stops and restarts on visibility change', () => {
  const dom = new JSDOM(`<!DOCTYPE html><body></body>`, { pretendToBeVisual: true });
  global.document = dom.window.document;
  global.window = dom.window;

  let hidden = false;
  Object.defineProperty(document, 'hidden', {
    get: () => hidden,
    set: (val) => { hidden = val; },
    configurable: true
  });

  let id = 1;
  window.setInterval = jest.fn(() => id++);
  window.clearInterval = jest.fn();

  let pollingInterval = null;
  function startPolling() {
    if (pollingInterval) {
      window.clearInterval(pollingInterval);
    }
    pollingInterval = window.setInterval(() => {}, 15000);
  }
  function stopPolling() {
    if (pollingInterval) {
      window.clearInterval(pollingInterval);
      pollingInterval = null;
    }
  }

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      stopPolling();
    } else {
      startPolling();
    }
  });

  startPolling();
  expect(window.setInterval).toHaveBeenCalledTimes(1);

  document.hidden = true;
  document.dispatchEvent(new dom.window.Event('visibilitychange'));
  expect(window.clearInterval).toHaveBeenCalledTimes(1);
  expect(pollingInterval).toBeNull();

  document.hidden = false;
  document.dispatchEvent(new dom.window.Event('visibilitychange'));
  expect(window.setInterval).toHaveBeenCalledTimes(2);

  delete global.document;
  delete global.window;
});
