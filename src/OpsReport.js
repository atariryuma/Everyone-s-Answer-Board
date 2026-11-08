/**
 * @fileoverview ops-monitor へエラーの「場所」だけを送る（ops-monitor の clients/gas/OpsReport.js をこのアプリ向けにしたもの）。
 *
 * 送るもの: 処理名（where）・Error の name・ファイル名と行。
 * 送らないもの: エラーの文面・logError_ の context（閲覧者のメールが入ることがある）・引数・シートの値。
 *
 * 送り先と鍵は、先生ごとの配置で scripts/deploy-single.js が Deployment.js に焼き込む
 *   （DEPLOY_OPS_REPORT_URL / DEPLOY_OPS_REPORT_TOKEN。元は scripts/ops-report.local.json）。
 *   Why プロパティでないか: Admin API は名前に TOKEN を含むプロパティを書けない（保護キー）ので、
 *   配置のたびに全員へ確実に届く Deployment.js に載せる。値が空なら何も送らない（リポジトリの既定）。
 *
 * どの先生の配置から来たかは送らない（全配置で 1 つのアプリ名・1 つの鍵）。
 */

// 受け口（ops-monitor の src/receiver.ts）が受け付ける形にそろえる
const OPS_PLACE_RE_ = /^[\w.$:-]{1,64}$/;
const OPS_FILE_RE_ = /^[\w./\-぀-ヿ一-鿿]{1,100}$/;
const OPS_KIND_RE_ = /^[A-Za-z_$][\w$]{0,63}$/;
// 同じ場所の報告は 10 分に 1 回まで（一覧表示の中などで同じ失敗が続いても、外への通信を増やさない）
const OPS_THROTTLE_SEC_ = 600;

/**
 * エラーの場所を ops-monitor へ送る。何が起きても投げない（本来の処理を止めない）。
 * @param {string} where - 処理名（logError_ の funcName など）
 * @param {*} err - Error オブジェクト（name と stack だけを見る）
 * @param {{file?: string, line?: number}} [place] - 場所が分かっているときに渡す（画面側のエラー）
 */
function opsReport_(where, err, place) {
  try {
    const url = typeof DEPLOY_OPS_REPORT_URL === 'string' ? DEPLOY_OPS_REPORT_URL : '';
    const token = typeof DEPLOY_OPS_REPORT_TOKEN === 'string' ? DEPLOY_OPS_REPORT_TOKEN : '';
    if (!url || !token) return;

    const w = OPS_PLACE_RE_.test(String(where)) ? String(where) : 'unknown';
    const kind = err && typeof err.name === 'string' && OPS_KIND_RE_.test(err.name) ? err.name : 'Error';
    const p = place || opsPlace_(err);
    const file = typeof p.file === 'string' && OPS_FILE_RE_.test(p.file) ? p.file : '?';
    const line = Number.isInteger(p.line) && p.line >= 0 ? p.line : 0;

    try {
      const cache = CacheService.getScriptCache();
      const key = 'ops_' + Utilities.base64EncodeWebSafe([w, kind, file, line].join('|')).slice(0, 200);
      if (cache.get(key)) return;
      cache.put(key, '1', OPS_THROTTLE_SEC_);
    } catch (cacheErr) {
      // キャッシュが使えないときは間引かずに送る
    }

    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ where: w, kind, file, line }),
      muteHttpExceptions: true
    });
  } catch (ignored) {
    // 報告の失敗で本来の処理を止めない
  }
}

/** スタックの最初の「自分のコード」の行から ファイル名:行 を取り出す。文面は見ない */
function opsPlace_(err) {
  const stack = err && typeof err.stack === 'string' ? err.stack.split('\n') : [];
  for (let i = 0; i < stack.length; i++) {
    // 例: "    at doGet_ (main:12:5)" / "    at main:12:5"
    const m = stack[i].match(/^\s+at (?:.*?\()?([^\s():]+):(\d+):\d+\)?$/);
    if (m && m[1] !== 'OpsReport' && OPS_FILE_RE_.test(m[1])) {
      return { file: m[1], line: Number(m[2]) };
    }
  }
  return { file: '?', line: 0 };
}
