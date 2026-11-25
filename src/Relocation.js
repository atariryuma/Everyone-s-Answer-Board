/**
 * @fileoverview コピーした配置を新しい版にするときの、データの引き継ぎと、前の URL からの転送。
 *
 * 新しい版はコピーし直しになる (スクリプトは自分自身を更新できない。docs/DISTRIBUTION.md)。そのとき:
 *   - 引き継ぎ: 新しいコピーが、前のスプレッドシートをそのまま DB として使う (DATABASE_SPREADSHEET_ID を付け替える)。
 *     データは写さない。userId・授業・回答・設定の中のシートの参照が 1 つも変わらないので、書き換えが要らず、
 *     写し損ねも起きない。前のスプレッドシートはデータの置き場として残る。
 *   - 転送: 引き継いだとき、DB のタブ「引っ越し先」に新しいウェブアプリの URL を書く。どの版の配置も doGet で
 *     これを読み、自分の URL でなければ、開いた人を新しい方へ送る (児童に URL を配り直さなくてよい)。
 *     「引っ越し先」の読み方は、もう配ったコピーの中にあるので変えない (変えると前の版が転送できなくなる)。
 * 焼き込みのある配置 (deploy-single.js の create / deploy) では何もしない。
 */

/* global isBakedDeployment_, openDatabase, getWebAppUrl, teacherOnlyError_, createSuccessResponse, createErrorResponse,
   getCachedProperty_, setCachedProperty_, clearDatabaseUserCache_, invalidateAccessDecisionCache_, validateSpreadsheetId,
   markGuideAdopted_, FIRST_STEPS_SHARED_KEY, FIRST_STEPS_CHILD_KEY */

const MOVED_SHEET_NAME = '引っ越し先';
const MOVED_CACHE_KEY = 'moved_to_v1';
const MOVED_CACHE_TTL = 600;
const MOVED_NONE = 'none';
// 送り先にしてよいのは Apps Script のウェブアプリの URL だけ (シートに別のサイトを書かれても送らない)
const WEB_APP_EXEC_URL_RE = /^(https:\/\/script\.google\.com\/(?:a\/macros\/[A-Za-z0-9.-]+\/|macros\/))s\/([A-Za-z0-9_-]+)\/exec$/;
const MOVED_SHEET_LABEL = '新しい回答ボードの URL';
// このスクリプトが紐づいたスプレッドシートの ID (初回の準備が残す。helpers.js)
const CONTAINER_KEY = 'CONTAINER_SPREADSHEET_ID';
// 引き継いだ相手 (前のスプレッドシート) の ID。あれば、もう引き継がない
const ADOPTED_KEY = 'ADOPTED_FROM_SPREADSHEET_ID';
// 引き継いだあと、紐づいたスプレッドシートに作るタブ。「ここはアプリの本体で、データは別にある」の印
const POINTER_SHEET_NAME = 'データの置き場';
const POINTER_SHEET_LABEL = '授業と回答が入っているスプレッドシート';
const PREVIOUS_BOARD_SCAN_LIMIT = 15;
const PREVIOUS_BOARD_RESULT_LIMIT = 5;

function webAppDeploymentId_(url) {
  const m = WEB_APP_EXEC_URL_RE.exec(String(url || '').trim());
  return m ? m[2] : '';
}

// URL の、配置の ID より前の部分 (組織のドメインが入る)。同じ組織のウェブアプリかを比べる。
function webAppUrlPrefix_(url) {
  const m = WEB_APP_EXEC_URL_RE.exec(String(url || '').trim());
  return m ? m[1] : '';
}

/**
 * この配置が引っ越し済みなら、送り先の URL を返す (doGet から)。引っ越していない・自分が送り先なら ''。
 * 開くたびに通るので、結果を 10 分 cache に置く (引っ越していないことも置く)。失敗したら送らない。
 * @param {string} [arrivedAs] - URL の moved の値 (送られてきたときの、送り先の配置の ID)。
 *   いまの引っ越し先と同じなら送らない (自分の URL を読み違えても回り続けない)。次の引っ越しで送り先が変われば、
 *   前の moved が付いたブックマークからも新しい方へ送る。
 */
function movedDestination_(arrivedAs) {
  if (isBakedDeployment_()) return '';
  try {
    const cache = CacheService.getScriptCache();
    let url = cache.get(MOVED_CACHE_KEY);
    if (!url) {
      const db = openDatabase();
      const sheet = db ? db.getSheetByName(MOVED_SHEET_NAME) : null;
      const written = sheet ? String(sheet.getRange(1, 2).getValue() || '').trim() : '';
      url = webAppDeploymentId_(written) ? written : MOVED_NONE;
      cache.put(MOVED_CACHE_KEY, url, MOVED_CACHE_TTL);
    }
    if (url === MOVED_NONE) return '';
    const target = webAppDeploymentId_(url);
    if (arrivedAs && arrivedAs === target) return '';
    const ownUrl = getWebAppUrl();
    const own = webAppDeploymentId_(ownUrl);
    // 自分の URL が分からないときは送らない (自分へ送って回り続けるのを避ける)
    if (!own || own === target) return '';
    // 同じ組織のウェブアプリにだけ送る (タブを書き換えられても、よその組織のアプリへ userId を渡さない)
    if (webAppUrlPrefix_(url) !== webAppUrlPrefix_(ownUrl)) return '';
    return url;
  } catch (e) {
    console.warn('movedDestination_: 引っ越し先を読めませんでした (送らずに続ける):', e && e.message);
    return '';
  }
}

/** 送り先の URL に、開こうとしていたページ (mode / userId / lessonId) を付ける。moved は送り先の配置の ID (movedDestination_ の arrivedAs)。 */
function movedTargetUrl_(baseUrl, params) {
  const p = params || {};
  const query = ['moved=' + webAppDeploymentId_(baseUrl)];
  ['mode', 'userId', 'lessonId'].forEach((key) => {
    const value = p[key];
    if (typeof value === 'string' && value && value.length <= 200) query.push(key + '=' + encodeURIComponent(value));
  });
  return baseUrl + '?' + query.join('&');
}

// 前の回答ボードとして使えるかを確かめて、分かったことを返す。使えなければ { ok:false, why }。
function previousBoardInfo_(spreadsheetId, adminEmail) {
  const file = DriveApp.getFileById(spreadsheetId);
  const owner = file.getOwner();
  const ownerEmail = owner ? String(owner.getEmail() || '').trim().toLowerCase() : '';
  if (!ownerEmail || ownerEmail !== adminEmail) return { ok: false, why: '自分が持ち主のスプレッドシートではありません' };
  const ss = SpreadsheetApp.openById(spreadsheetId);
  // 引き継ぎを済ませたコピー (アプリの本体) は、データを持っていない。相手にすると空の DB を引き継いでしまう
  const pointer = ss.getSheetByName(POINTER_SHEET_NAME);
  if (pointer && String(pointer.getRange(1, 1).getValue() || '').trim() === POINTER_SHEET_LABEL) {
    return { ok: false, why: 'そのスプレッドシートには授業や回答が入っていません（タブ「' + POINTER_SHEET_NAME + '」に書いてあるスプレッドシートを選んでください）' };
  }
  const users = ss.getSheetByName('users');
  const lessons = ss.getSheetByName('lessons');
  if (!users || !lessons) return { ok: false, why: '回答ボードのスプレッドシートではありません（users / lessons のタブがありません）' };
  const rows = users.getLastRow() > 1 ? users.getRange(2, 1, users.getLastRow() - 1, 2).getValues() : [];
  const mine = rows.some((r) => String(r[1] || '').trim().toLowerCase() === adminEmail);
  if (!mine) return { ok: false, why: 'このスプレッドシートには、あなたの登録がありません' };
  return {
    ok: true,
    spreadsheet: ss,
    id: spreadsheetId,
    name: file.getName(),
    updatedAt: file.getLastUpdated().toISOString(),
    lessonCount: Math.max(0, lessons.getLastRow() - 1)
  };
}

// このスクリプトが紐づいたスプレッドシートの ID。初回の準備 (doGet の中) が残した値を先に使う。
//   Why: 画面からの呼び出し (google.script.run) の中で getActiveSpreadsheet が紐づけ元を返すかを、実物で確かめていない。
function containerSpreadsheetId_() {
  const saved = String(PropertiesService.getScriptProperties().getProperty(CONTAINER_KEY) || '');
  if (saved) return saved;
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    return active ? active.getId() : '';
  } catch (_) { /* 紐づけ元が分からないだけ (呼び出し側が、無い場合を扱う) */
    return '';
  }
}

// いまの DB と、このスクリプトが紐づいたスプレッドシート (引き継ぎの相手にしないもの)
function relocationSelfIds_() {
  return [String(getCachedProperty_('DATABASE_SPREADSHEET_ID') || ''), containerSpreadsheetId_()].filter(Boolean);
}

/** 管理パネル「前の回答ボードを探す」。先生だけ。名前に「回答ボード」が入った、自分が持ち主のスプレッドシートを見る。 */
function findPreviousBoards() {
  const denied = teacherOnlyError_();
  if (denied) return denied;
  if (isBakedDeployment_()) return createErrorResponse('この配置では使えません');
  try {
    const adminEmail = String(getCachedProperty_('ADMIN_EMAIL') || '').toLowerCase();
    const skip = relocationSelfIds_();
    const files = DriveApp.searchFiles(
      "mimeType = 'application/vnd.google-apps.spreadsheet' and title contains '回答ボード' and 'me' in owners and trashed = false");
    const found = [];
    for (let scanned = 0; files.hasNext() && scanned < PREVIOUS_BOARD_SCAN_LIMIT; scanned += 1) {
      const id = files.next().getId();
      if (skip.indexOf(id) !== -1) continue;
      let info = null;
      try {
        info = previousBoardInfo_(id, adminEmail);
      } catch (_) { /* 開けないファイルは候補にしないだけ */ }
      if (info && info.ok) found.push({ id: info.id, name: info.name, updatedAt: info.updatedAt, lessonCount: info.lessonCount });
    }
    found.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
    return createSuccessResponse('', { boards: found.slice(0, PREVIOUS_BOARD_RESULT_LIMIT) });
  } catch (e) {
    console.warn('findPreviousBoards failed:', e && e.message);
    return createErrorResponse('前の回答ボードを探せませんでした: ' + (e && e.message ? e.message : 'unknown'));
  }
}

/**
 * 管理パネル「引き継ぐ」。先生だけ。前のスプレッドシートを、この配置の DB にする。
 * @param {string} idOrUrl - 前のスプレッドシートの ID か URL
 */
function adoptPreviousBoard(idOrUrl) {
  const denied = teacherOnlyError_();
  if (denied) return denied;
  if (isBakedDeployment_()) return createErrorResponse('この配置では使えません');
  const text = typeof idOrUrl === 'string' ? idOrUrl.trim() : '';
  const fromUrl = /\/d\/([A-Za-z0-9_-]{40,50})/.exec(text);
  const targetId = fromUrl ? fromUrl[1] : text;
  if (!validateSpreadsheetId(targetId).isValid) {
    return createErrorResponse('スプレッドシートの URL か ID を読み取れませんでした');
  }
  if (relocationSelfIds_().indexOf(targetId) !== -1) {
    return createErrorResponse('それは、いま使っているスプレッドシートです');
  }
  // 送り先として書くのは、デプロイした URL (…/exec) だけ。テスト用の URL から開いているときは止める
  const ownUrl = String(getWebAppUrl() || '').trim();
  if (!webAppDeploymentId_(ownUrl)) {
    return createErrorResponse('このページの URL を読めませんでした。デプロイしたウェブアプリの URL から開き直して、もう一度お試しください');
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse('ほかの処理が動いています。少し待ってから、もう一度お試しください');
  const previousDbId = String(getCachedProperty_('DATABASE_SPREADSHEET_ID') || '');
  let switched = false;
  const props = PropertiesService.getScriptProperties();
  try {
    // 引き継ぎは 1 度だけ。2 つ目を引き継ぐと、1 つ目の前の URL を開いた児童が 2 つ目のデータへ送られる。
    //   ロックの中で確かめる (同時に 2 つ押されても、後の方を断る)
    if (props.getProperty(ADOPTED_KEY)) {
      return createErrorResponse('この回答ボードは、すでに前の回答ボードを引き継いでいます');
    }
    const adminEmail = String(getCachedProperty_('ADMIN_EMAIL') || '').toLowerCase();
    const info = previousBoardInfo_(targetId, adminEmail);
    if (!info.ok) return createErrorResponse(info.why);

    // 同じ名前のタブを人が作っていたら、上書きしない (アプリが書いたタブは A1 が決まった見出し。空のタブは使ってよい)
    const existing = info.spreadsheet.getSheetByName(MOVED_SHEET_NAME);
    if (existing && existing.getLastRow() > 0) {
      const head = String(existing.getRange(1, 1).getValue() || '').trim();
      if (head !== MOVED_SHEET_LABEL) {
        return createErrorResponse('前のスプレッドシートに「' + MOVED_SHEET_NAME + '」という名前のタブがあります。名前を変えてから、もう一度お試しください');
      }
    }

    props.setProperty(ADOPTED_KEY, targetId);
    switched = true;
    setCachedProperty_('DATABASE_SPREADSHEET_ID', targetId);
    // 前の URL を開いた人を、この配置へ送る印。ここで失敗したら DB を戻す (送り先の無い引き継ぎを残さない)
    const sheet = existing || info.spreadsheet.insertSheet(MOVED_SHEET_NAME);
    sheet.getRange(1, 1, 3, 2).setNumberFormat('@').setValues([
      [MOVED_SHEET_LABEL, ownUrl],
      ['書いた日時', new Date().toISOString()],
      ['このタブについて', 'アプリが使います。前の URL を開いた人を、上の URL へ送ります。消すと送らなくなります。']
    ]);
    // 印を書いたら、もう DB を戻さない (戻すと、前の URL を開いた児童が空の DB の配置へ送られる)。
    //   ここから下は、失敗しても引き継ぎそのものは済んでいる。
    switched = false;

    try {
      clearDatabaseUserCache_();
      invalidateAccessDecisionCache_(targetId);
      if (previousDbId) invalidateAccessDecisionCache_(previousDbId);
      CacheService.getScriptCache().remove(MOVED_CACHE_KEY);
    } catch (e) {
      // cache は長くても 6 時間で切れる。それまで古い利用者の情報を読むことがある
      console.warn('adoptPreviousBoard: cache を消せませんでした:', e && e.message);
    }
    try {
      // 前から使っている先生なので、「はじめの 4 歩」は出さない
      props.setProperty(FIRST_STEPS_SHARED_KEY, '1');
      props.setProperty(FIRST_STEPS_CHILD_KEY, '1');
    } catch (e) {
      // 「はじめの 4 歩」のカードが出るだけ
      console.warn('adoptPreviousBoard: 「はじめの 4 歩」の印を残せませんでした:', e && e.message);
    }
    try {
      // 紐づいたスプレッドシート (アプリの本体) に、データの置き場を書く。次に新しい版にするとき、
      //   このスプレッドシートを引き継ぎの相手に選べないようにする印にもなる (previousBoardInfo_)
      const containerId = containerSpreadsheetId_();
      const own = containerId ? SpreadsheetApp.openById(containerId) : null;
      if (own) {
        const pointerSheet = own.getSheetByName(POINTER_SHEET_NAME) || own.insertSheet(POINTER_SHEET_NAME);
        pointerSheet.getRange(1, 1, 3, 2).setNumberFormat('@').setValues([
          [POINTER_SHEET_LABEL, info.name],
          ['URL', 'https://docs.google.com/spreadsheets/d/' + targetId + '/edit'],
          ['このタブについて', 'このスプレッドシートはアプリの本体です。授業や回答は、上のスプレッドシートに入っています。どちらも消さないでください。']
        ]);
        if (typeof markGuideAdopted_ === 'function') markGuideAdopted_(own, info.name);
      }
    } catch (e) {
      // 案内のタブに書けないだけ (引き継ぎは済んでいる)
      console.warn('adoptPreviousBoard: 紐づいたスプレッドシートに書けませんでした:', e && e.message);
    }
    return createSuccessResponse('引き継ぎました', { name: info.name, lessonCount: info.lessonCount });
  } catch (e) {
    if (switched && previousDbId) {
      try {
        props.deleteProperty(ADOPTED_KEY);
        setCachedProperty_('DATABASE_SPREADSHEET_ID', previousDbId);
      } catch (revertError) {
        console.error('adoptPreviousBoard: DB を戻せませんでした:', revertError && revertError.message);
      }
    }
    console.warn('adoptPreviousBoard failed:', e && e.message);
    return createErrorResponse('引き継げませんでした: ' + (e && e.message ? e.message : 'unknown'));
  } finally {
    lock.releaseLock();
  }
}
