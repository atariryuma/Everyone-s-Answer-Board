/**
 * scripts/deploy-single.js: エラーの報告先 (scripts/ops-report.local.json) が無いときは配置しない。
 *
 * Why: このファイルは git に入らない。別の PC に clone しただけの状態で配置すると、送り先が空で
 *   焼き込まれ、本番のエラーの報告が黙って止まる (2026-11-11 に気づいた)。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

const { loadOpsReport, stage } = require('../scripts/deploy-single.js');

function withFile(content, fn) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-report-'));
  const file = path.join(dir, 'ops-report.local.json');
  try {
    if (content !== null) fs.writeFileSync(file, content);
    return fn(file);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('ファイルが無ければ、配置を止める (空の送り先を返さない)', () => {
  withFile(null, (file) => {
    assert.throws(() => loadOpsReport({ file }), /ops-report\.local\.json がありません.*配置しません/);
  });
});

test('url と token がそろっていれば、その値を返す', () => {
  withFile(JSON.stringify({ url: 'https://example.invalid/report', token: 'ops_dummy' }), (file) => {
    assert.deepStrictEqual(loadOpsReport({ file }), { url: 'https://example.invalid/report', token: 'ops_dummy' });
  });
});

test('url か token が欠けていれば、配置を止める', () => {
  withFile(JSON.stringify({ url: 'https://example.invalid/report', token: '' }), (file) => {
    assert.throws(() => loadOpsReport({ file }), /url か token がありません/);
  });
  withFile(JSON.stringify({ token: 'ops_dummy' }), (file) => {
    assert.throws(() => loadOpsReport({ file }), /url か token がありません/);
  });
  withFile('null', (file) => {
    assert.throws(() => loadOpsReport({ file }), /url か token がありません/);
  });
});

test('文字列でない値・空白だけの値・https でない url は、配置を止める', () => {
  withFile(JSON.stringify({ url: {}, token: 'ops_dummy' }), (file) => {
    assert.throws(() => loadOpsReport({ file }), /url か token がありません/);
  });
  withFile(JSON.stringify({ url: 'https://example.invalid/report', token: '   ' }), (file) => {
    assert.throws(() => loadOpsReport({ file }), /url か token がありません/);
  });
  for (const url of ['http://example.invalid/report', 'example.invalid/report', 'https://', 'https://exa mple.invalid/report',
    'https://?/report', 'https://example.invalid:abc/report', 'ftp://example.invalid/report']) {
    withFile(JSON.stringify({ url, token: 'ops_dummy' }), (file) => {
      assert.throws(() => loadOpsReport({ file }), /https:\/\/ の URL ではありません/, url);
    });
  }
});

test('前後の空白は落として返す', () => {
  withFile(JSON.stringify({ url: ' https://example.invalid/report\n', token: ' ops_dummy ' }), (file) => {
    assert.deepStrictEqual(loadOpsReport({ file }), { url: 'https://example.invalid/report', token: 'ops_dummy' });
  });
});

test('配置の名前がフォルダの外を指すときは、消す前に止める', () => {
  const ops = { url: '', token: '' };
  for (const name of ['../src', '..', 'a/b', 'a\\b', 'a b', '', undefined]) {
    assert.throws(() => stage({ name }, ops), /name が英数字/, String(name));
  }
  assert.ok(fs.existsSync(path.join(__dirname, '..', 'src', 'main.js')), 'src/ が消えている');
});

test('JSON が壊れていれば止める。エラー文に中身 (鍵) を出さない', () => {
  withFile('{"url":"https://example.invalid/report","token":"ops_SECRETVALUE"', (file) => {
    let message = '';
    try { loadOpsReport({ file }); } catch (e) { message = e.message; }
    assert.match(message, /JSON として読めません/);
    assert.ok(!message.includes('ops_SECRETVALUE'), 'エラー文に鍵が出ている');
    assert.ok(!message.includes('example.invalid'), 'エラー文に送り先が出ている');
  });
});

test('--no-ops-report (skip) のときだけ、空の送り先で進める', () => {
  withFile(null, (file) => {
    assert.deepStrictEqual(loadOpsReport({ file, skip: true }), { url: '', token: '' });
  });
});

test('読み込むだけでは配置が走らない (require で main を呼ばない)', () => {
  assert.strictEqual(typeof loadOpsReport, 'function');
});
