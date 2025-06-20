const fs = require('fs');
const { JSDOM } = require('jsdom');

test('size slider value persists with localStorage', () => {
  const html = fs.readFileSync('./src/Page.html', 'utf8');
  const start = html.lastIndexOf('<script>');
  const end = html.lastIndexOf('</script>');
  let script = html.slice(start + 8, end);
  script = script.replace(/<\?=.+?\?>/g, 'false');
  script = script.replace(/new StudyQuestApp\(\);/, '');
  script += '\nwindow.StudyQuestApp = StudyQuestApp;';

  const dom = new JSDOM(`<!DOCTYPE html><body>
    <div id="main-container"></div>
    <header id="main-header"></header>
    <main id="answers"></main>
    <input id="sizeSlider" value="4">
    <span id="sliderValue"></span>
    <div id="headingLabel"></div>
    <div id="sheetNameText"></div>
    <button id="adminToggleBtn"></button>
    <div id="answerCount"></div>
    <div id="answerModalContainer"></div>
    <button id="answerModalCloseBtn"></button>
    <div id="answerModalCard"></div>
    <div id="modalAnswer"></div>
    <div id="modalStudentName"></div>
    <div id="modalReactions"></div>
    <select id="classFilter"></select>
    <select id="sortOrder"></select>
    <div id="controlsFooter"></div>
  </body>`, { runScripts: 'outside-only', url: 'https://example.com' });

  const { window } = dom;
  global.window = window;
  global.document = window.document;
  global.localStorage = window.localStorage;
  window.gsap = { fromTo: () => {}, to: () => {} };

  const vm = require('vm');
  vm.runInContext(script, dom.getInternalVMContext());
  window.StudyQuestApp.prototype.renderIcons = () => {};
  window.StudyQuestApp.prototype.verifyAdmin = () => {};
  window.StudyQuestApp.prototype.adjustLayout = () => {};
  window.StudyQuestApp.prototype.loadInitialData = () => {};

  window.localStorage.setItem('boardColumns', '5');
  const app = new window.StudyQuestApp();
  expect(window.document.getElementById('sizeSlider').value).toBe('5');
  expect(window.document.getElementById('sliderValue').textContent).toBe('5');

  window.document.getElementById('sizeSlider').value = '3';
  window.document.getElementById('sizeSlider').dispatchEvent(new window.Event('input'));
  expect(window.localStorage.getItem('boardColumns')).toBe('3');

  delete global.window;
  delete global.document;
  delete global.localStorage;
});
