const { JSDOM } = require('jsdom');

function createDeferred() {
  let resolve, reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

async function handleReaction(rowIndex, reaction, gas) {
  const btns = document.querySelectorAll(`[data-row-index="${rowIndex}"][data-reaction="${reaction}"]`);
  btns.forEach(btn => {
    btn.disabled = true;
    btn.setAttribute('aria-disabled', 'true');
  });
  try {
    await gas.addReaction(rowIndex, reaction);
  } catch (e) {
    // ignore
  } finally {
    btns.forEach(btn => {
      btn.disabled = false;
      btn.setAttribute('aria-disabled', 'false');
    });
  }
}

test('reaction buttons disable during async call', async () => {
  const dom = new JSDOM(`<button data-row-index="1" data-reaction="LIKE"></button>`);
  global.document = dom.window.document;
  const def = createDeferred();
  const gas = { addReaction: jest.fn(() => def.promise) };

  const promise = handleReaction(1, 'LIKE', gas);
  const btn = dom.window.document.querySelector('button');
  expect(btn.disabled).toBe(true);
  expect(btn.getAttribute('aria-disabled')).toBe('true');

  def.resolve({ status: 'ok' });
  await promise;
  expect(btn.disabled).toBe(false);
  expect(btn.getAttribute('aria-disabled')).toBe('false');

  delete global.document;
});
