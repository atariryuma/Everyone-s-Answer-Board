/**
 * @fileoverview ひな形をコピーした先生への案内 (スプレッドシートの最初のタブ「はじめに」と、メニュー)。
 *
 * Why 最初のタブか: コピーして開いた瞬間に、許可の前でも必ず目に入るのはシートだけ。開いたときに
 *   サイドバーやダイアログを自動で出すことは、許可の前 (onOpen) にはできず、Google の指針でも勧められていない。
 *   メニューは「はじめに」へ戻るための入口にとどめる (ダイアログを出すと appsscript.json に権限を足すことになり、
 *   動いている配置の先生みんなに許可のし直しが要る)。
 * 焼き込みのある配置 (deploy-single.js の create / deploy) では何もしない。先生の DB にタブを足さない。
 * 手順の文を変えたら GUIDE_VERSION を上げる (開いたときに書き直す)。docs/DISTRIBUTION.md の手順とそろえる。
 */

/* global DEPLOY_ADMIN_EMAIL, DEPLOY_DATABASE_SPREADSHEET_ID, getWebAppUrl, teacherOnlyError_, createSuccessResponse */

const GUIDE_SHEET_NAME = 'はじめに';
const GUIDE_VERSION = '2';
const GUIDE_VERSION_LABEL = '案内の版';
const GUIDE_STATUS_ROW = 3;
const GUIDE_URL_ROW = 4;
const GUIDE_STATUS_PENDING = '⏳ まだ準備の途中です。下の「準備」の 1〜6 をしてください。';
const GUIDE_STATUS_READY = '✅ 準備ができました。下の URL が、あなたの管理パネルです。ブックマークしてください。';

// [種類, A 列, B 列]。種類: title / status / section / step / note
const GUIDE_LINES = [
  ['title', 'みんなの回答ボード', ''],
  ['note', '', '児童の考えを 1 つの画面に集めて、みんなで見合うための道具です。このスプレッドシートが、あなた専用の回答ボードになります。'],
  ['status', '今の状態', GUIDE_STATUS_PENDING],
  ['status', '管理パネルの URL', '（準備が済むと、ここに出ます）'],
  ['note', '', ''],
  ['section', '準備', '最初の 1 回だけ。5 分ほどです。'],
  ['step', '1', 'まず、このスプレッドシートを自分のものにします。上のメニューの「ファイル」→「コピーを作成」を押し、できたコピーの方で 2 から先を進めます（すでにコピーした方を開いているなら、そのまま 2 へ）。コピーは自分の「マイドライブ」に置きます（共有ドライブでは動きません）。'],
  ['step', '2', '上のメニューの「拡張機能」→「Apps Script」を押します。新しいタブが開きます。'],
  ['step', '3', '開いたタブの右上の青いボタン「デプロイ」→「新しいデプロイ」を押します。「種類の選択」と出たら、歯車から「ウェブアプリ」を選びます。'],
  ['step', '4', '「次のユーザーとして実行」が「自分」、「アクセスできるユーザー」が「（あなたの学校や市の組織の名前）内の全員」になっていることを確かめて、「デプロイ」を押します。'],
  ['step', '5', '「アクセスを承認」を押し、自分のアカウントを選んで「許可」を押します。'],
  ['step', '6', '「ウェブアプリ」の下に出た URL を開きます。管理パネルが出たら、準備は終わりです。このタブに戻ると、上の「今の状態」が ✅ に変わっています。'],
  ['note', '', ''],
  ['section', '授業で使う', ''],
  ['step', '1', '管理パネルを開きます（上の「管理パネルの URL」）。'],
  ['step', '2', '管理パネルで、ボードにする Google フォームを選ぶか、授業モードの授業を作ります。'],
  ['step', '3', '「公開する」を押し、「ボードを開く」の URL を児童に配ります（クラスルームなど）。'],
  ['step', '4', 'はじめて授業で使う前に、児童のアカウントでその URL が開けるかを 1 度確かめます。'],
  ['note', '', ''],
  ['section', '困ったとき', ''],
  ['step', '・', '自分で開くと「まだ準備できていません」と出る → このスプレッドシートが共有ドライブにあるか、持ち主が自分ではありません。マイドライブにコピーし直して、準備をやり直します。'],
  ['step', '・', '児童が開くと Google のエラーの画面になる → 準備の 4 を確かめます。合っていれば、組織の設定で止められています（学校や市の管理者に相談してください。アプリでは直せません）。'],
  ['step', '・', '児童が開くと「同一ドメインのアカウントで」と出る → 児童と先生のアカウントの @ より後ろが違います。今は、先生と同じ組織の児童だけが開けます。'],
  ['note', '', ''],
  ['section', '気をつけること', ''],
  ['step', '・', 'このスプレッドシートには、児童の回答が入ります。ほかの人と共有したり、コピーして渡したりしないでください。ほかの先生には、もらったときの URL を伝えます。'],
  ['step', '・', '「users」「lessons」「lesson_responses」と、授業の名前の付いたタブは、アプリが使います。消したり、名前を変えたりしないでください。'],
  ['step', '・', 'このタブ「はじめに」は、消しても動きます。メニューの「回答ボード」→「はじめに（使い方）を開く」で、もう一度出せます。'],
  ['note', '', ''],
  ['note', GUIDE_VERSION_LABEL, GUIDE_VERSION]
];

// 焼き込みのある配置 (自分で配置した先生) かどうか。Deployment.js が無い (テスト) ときは「ある」と同じに扱い、何もしない。
function isBakedDeployment_() {
  if (typeof DEPLOY_ADMIN_EMAIL === 'undefined') return true;
  return !!(DEPLOY_ADMIN_EMAIL || DEPLOY_DATABASE_SPREADSHEET_ID);
}

// スプレッドシートの画面から呼ばれたときだけ UI を返す。ウェブアプリから (児童の画面の google.script.run で)
//   呼ばれたときは null。onOpen と openBoardGuide は名前の都合で公開関数になるので、UI が無ければ何も書かない。
function guideUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (_) { /* スプレッドシートの画面の外から呼ばれた。案内は要らない */
    return null;
  }
}

/**
 * スプレッドシートを開いたとき (simple trigger。許可の前でも動く)。
 * コピーした配置でだけ、「はじめに」が無い・古いときに書き、メニューを足す。
 */
function onOpen() {
  if (isBakedDeployment_()) return;
  const ui = guideUi_();
  if (!ui) return;
  try {
    ui.createMenu('回答ボード').addItem('はじめに（使い方）を開く', 'openBoardGuide').addToUi();
  } catch (e) {
    console.warn('onOpen: メニューを足せませんでした:', e && e.message);
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const created = ensureGuideSheet_(ss);
    if (created) ss.setActiveSheet(created);   // 作った・書き直したときだけ前に出す (毎回は出さない)
  } catch (e) {
    console.warn('onOpen: 「はじめに」を書けませんでした:', e && e.message);
  }
}

/** メニュー「回答ボード → はじめに（使い方）を開く」。 */
function openBoardGuide() {
  if (isBakedDeployment_()) return;
  if (!guideUi_()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureGuideSheet_(ss);
  ss.setActiveSheet(ss.getSheetByName(GUIDE_SHEET_NAME));
}

/**
 * 「はじめに」が無い・版が古いときに書く。今の版ならシートに触らない。
 * @returns {Sheet|null} 書いたときはそのシート、何もしなかったら null
 */
function ensureGuideSheet_(ss) {
  let sheet = ss.getSheetByName(GUIDE_SHEET_NAME);
  let status = null;
  if (sheet) {
    const last = sheet.getLastRow();
    const values = last > 0 ? sheet.getRange(1, 1, last, 2).getValues() : [];
    const versionRow = values.find((r) => r[0] === GUIDE_VERSION_LABEL);
    if (versionRow && String(versionRow[1]) === GUIDE_VERSION) return null;
    // 書き直す前に、準備が済んだ印と URL を控える (初回の準備は 1 度しか走らない)
    const s = values[GUIDE_STATUS_ROW - 1];
    const u = values[GUIDE_URL_ROW - 1];
    if (s && String(s[1]).indexOf('✅') === 0 && u) status = { text: String(s[1]), url: String(u[1] || '') };
    sheet.clear();
  } else {
    sheet = ss.insertSheet(GUIDE_SHEET_NAME, 0);
  }

  const rows = GUIDE_LINES.map((l) => [l[1], l[2]]);
  if (status) {
    rows[GUIDE_STATUS_ROW - 1][1] = status.text;
    rows[GUIDE_URL_ROW - 1][1] = status.url;
  }
  // 「1」などを数や日付に変えさせない・「=」で始まっても式にしない
  sheet.getRange(1, 1, rows.length, 2).setNumberFormat('@').setValues(rows)
    .setWrap(true).setVerticalAlignment('top').setFontSize(11);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 780);

  const a1 = (kind) => GUIDE_LINES.map((l, i) => (l[0] === kind ? `A${i + 1}:B${i + 1}` : null)).filter(Boolean);
  sheet.getRangeList(a1('title')).setFontSize(18).setFontWeight('bold');
  sheet.getRangeList(a1('section')).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRangeList(a1('status')).setFontWeight('bold').setBackground('#fef7e0');
  sheet.getRangeList(GUIDE_LINES.map((l, i) => (l[0] === 'step' ? `A${i + 1}` : null)).filter(Boolean))
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 初回の準備が済んだことを「はじめに」に書く (helpers.js の ensureCopiedDeploymentBootstrap_ から)。
 * 先生がウェブアプリを開いた実行の中なので、URL が分かる。
 */
function markGuideReady_(ss) {
  ensureGuideSheet_(ss);
  const sheet = ss.getSheetByName(GUIDE_SHEET_NAME);
  const url = typeof getWebAppUrl === 'function' ? String(getWebAppUrl() || '') : '';
  sheet.getRange(GUIDE_STATUS_ROW, 2, 2, 1).setValues([
    [GUIDE_STATUS_READY],
    [url || '（URL を読めませんでした。準備の 6 で開いた URL をブックマークしてください）']
  ]);
}

/**
 * 前の回答ボードを引き継いだことを「はじめに」に書く (Relocation.js の adoptPreviousBoard から)。
 * データは前のスプレッドシートに入ったままなので、消さないように名前を書いておく。
 */
function markGuideAdopted_(ss, previousName) {
  ensureGuideSheet_(ss);
  const sheet = ss.getSheetByName(GUIDE_SHEET_NAME);
  const url = typeof getWebAppUrl === 'function' ? String(getWebAppUrl() || '') : '';
  sheet.getRange(GUIDE_STATUS_ROW, 2, 2, 1).setValues([
    ['✅ 前の回答ボードを引き継ぎました。授業や回答は、スプレッドシート「' + previousName + '」に入っています。そちらは消さないでください。'],
    [url]
  ]);
}

// ─── 管理パネルの「はじめの 4 歩」 ─────────────────────────────────────
// コピーした配置の先生に、準備のあと授業で使えるようになるまでを管理パネルで案内する (AdminPanel.html の #first-steps)。
//   1 準備 (管理パネルが開けた時点で済み) / 2 授業をつくる (画面が授業の数で判定) /
//   3 URL を配る (URL コピーを押した) / 4 児童が開いた (先生・共同編集者でない人にボードを返した)。
//   3 と 4 はスクリプトのプロパティに残す。焼き込みのある配置では出さず、何も書かない。
const FIRST_STEPS_SHARED_KEY = 'FIRST_STEPS_URL_SHARED';
const FIRST_STEPS_CHILD_KEY = 'FIRST_STEPS_CHILD_OPENED';
const FIRST_STEPS_CHILD_CACHE_KEY = 'first_steps_child_opened';

/** 管理パネルが開くときに読む。先生だけ。 */
function getFirstStepsState() {
  const denied = teacherOnlyError_();
  if (denied) return denied;
  if (isBakedDeployment_()) return createSuccessResponse('', { show: false });
  const props = PropertiesService.getScriptProperties();
  return createSuccessResponse('', {
    show: true,
    urlShared: props.getProperty(FIRST_STEPS_SHARED_KEY) === '1',
    childOpened: props.getProperty(FIRST_STEPS_CHILD_KEY) === '1'
  });
}

/** 管理パネルで「URL コピー」を押したとき。先生だけ。 */
function markFirstStepUrlShared() {
  const denied = teacherOnlyError_();
  if (denied) return denied;
  if (!isBakedDeployment_()) PropertiesService.getScriptProperties().setProperty(FIRST_STEPS_SHARED_KEY, '1');
  return createSuccessResponse('', { urlShared: true });
}

/**
 * 児童にボードを返したことを 1 度だけ残す (main.js の handleViewMode_ から)。
 * 児童が開くたびに通るので、済んだあとは CacheService を 1 回読むだけにする (プロパティを読まない)。
 * 失敗しても児童の画面は止めない。
 */
function noteChildOpenedBoard_() {
  if (isBakedDeployment_()) return;
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get(FIRST_STEPS_CHILD_CACHE_KEY)) return;
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(FIRST_STEPS_CHILD_KEY) !== '1') props.setProperty(FIRST_STEPS_CHILD_KEY, '1');
    cache.put(FIRST_STEPS_CHILD_CACHE_KEY, '1', 21600);
  } catch (e) {
    // 4 歩目の印が付かないだけ (次に児童が開いたときにもう一度試す)
    console.warn('noteChildOpenedBoard_: 印を残せませんでした:', e && e.message);
  }
}
