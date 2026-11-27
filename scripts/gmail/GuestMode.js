/**
 * 個人の Gmail で配置したときだけ足すファイル (single-tenants.json の extraFiles)。学校の配置には入れない。
 *
 * Why: 「自分として実行」のウェブアプリでは、開いた人のメール (Session.getActiveUser) は配置した人と
 *   同じ Workspace のドメインの人にしか返らない。@gmail.com の参加者には空が返り、送信が NO_EMAIL で断られる。
 *   そこで、メールが空の人だけを「端末ごとの番号」で見分ける。番号は ブラウザの localStorage に置き、
 *   ページを開くときは URL の g=、画面からの呼び出しは guestRpc の引数で届ける。
 *   見分けの名前は guest-<番号>@gmail.com。Gmail のアドレスに「-」は使えないので、本物の人 (管理者を含む) とは重ならない。
 *
 * つなぎ目 (src にある。このファイルが無い配置では何もしない):
 *   - getCurrentEmail (main.js): 本物のメールが空のときだけ guestEmail_() を使う
 *   - doGet (main.js): 最初に guestBeginGet_(e)、返す前に guestDecorateOutput_(out)
 */
const GUEST_ID_RE_ = /^[a-z0-9]{24,40}$/;
let __guestIdForThisRun = null;

function guestEmail_() {
  return __guestIdForThisRun ? 'guest-' + __guestIdForThisRun + '@gmail.com' : null;
}

function guestRealEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function guestNewId_() {
  return Utilities.getUuid().replace(/-/g, '').toLowerCase();
}

/**
 * doGet の最初に呼ぶ。本物のメールがある人 (管理者) には何もしない。
 * メールが空で g= が無い人には、端末の番号を決めてから開き直すページを返す。
 * @returns {HtmlOutput|null} 返すページ (無ければ null で、そのまま続ける)
 */
function guestBeginGet_(e) {
  __guestIdForThisRun = null;
  if (guestRealEmail_()) return null;
  const params = (e && e.parameter) || {};
  const g = typeof params.g === 'string' ? params.g : '';
  if (GUEST_ID_RE_.test(g)) {
    __guestIdForThisRun = g;
    return null;
  }
  const base = ScriptApp.getService().getUrl();
  const keep = {};
  Object.keys(params).forEach(function (k) { if (k !== 'g') keep[k] = String(params[k]); });
  const html =
    '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:sans-serif;display:flex;min-height:90vh;align-items:center;justify-content:center;margin:0;background:#f7f7f5}' +
    'a{display:inline-block;padding:18px 36px;font-size:22px;border-radius:12px;background:#1f6feb;color:#fff;text-decoration:none}</style></head>' +
    '<body><a id="go" target="_top" href="#">回答ボードを開く</a><script>' +
    '(function(){var base=' + JSON.stringify(base) + ',keep=' + JSON.stringify(keep) + ',fresh=' + JSON.stringify(guestNewId_()) + ';' +
    'var id=null;try{id=localStorage.getItem("eabGuestId");}catch(e){}' +
    'if(!/^[a-z0-9]{24,40}$/.test(id||"")){id=fresh;try{localStorage.setItem("eabGuestId",id);}catch(e){}}' +
    'var q=[];for(var k in keep){q.push(encodeURIComponent(k)+"="+encodeURIComponent(keep[k]));}q.push("g="+id);' +
    'var url=base+"?"+q.join("&");document.getElementById("go").href=url;' +
    'try{window.top.location.href=url;}catch(e){}})();' +
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('回答ボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * doGet が返すページの頭に、画面からの呼び出しを guestRpc 経由にする小さなスクリプトを入れる。
 * 番号で入った人のページだけ。管理者 (本物のメールがある人) のページはそのまま。
 */
function guestDecorateOutput_(out) {
  if (!__guestIdForThisRun || !out || typeof out.getContent !== 'function') return out;
  const script = '<script>(function(){var id=' + JSON.stringify(__guestIdForThisRun) + ';' +
    'try{localStorage.setItem("eabGuestId",id);}catch(e){}' +
    'function wrap(r){return new Proxy({},{get:function(_,p){' +
    'if(p==="withSuccessHandler"||p==="withFailureHandler"||p==="withUserObject"){return function(h){return wrap(r[p](h));};}' +
    'if(typeof p!=="string")return undefined;' +
    'return function(){return r.guestRpc(id,p,Array.prototype.slice.call(arguments));};}});}' +
    'function install(){if(typeof google==="undefined"||!google.script||!google.script.run||google.script.run.__guest)return false;' +
    'var real=google.script.run;var w=wrap(real);' +
    'try{Object.defineProperty(google.script,"run",{configurable:true,get:function(){return w;}});}catch(e){google.script.run=w;}' +
    'return true;}' +
    'if(!install()){var n=0;var t=setInterval(function(){if(install()||++n>200)clearInterval(t);},10);}' +
    '})();</script>';
  const content = out.getContent();
  const at = content.search(/<head[^>]*>/i);
  const next = at >= 0
    ? content.replace(/<head[^>]*>/i, function (m) { return m + script; })
    : script + content;
  out.setContent(next);
  return out;
}

// 画面からの呼び出しの入口 (番号で入った人だけが使う)。呼べるのは、google.script.run から元々呼べる関数だけ。
const GUEST_RPC_DENY_ = { guestRpc: true, doGet: true, doPost: true };
function guestRpc(id, name, args) {
  if (!GUEST_ID_RE_.test(String(id || ''))) throw new Error('ページを開き直してください');
  const fnName = String(name || '');
  if (!/^[A-Za-z][A-Za-z0-9]*$/.test(fnName) || GUEST_RPC_DENY_[fnName]) throw new Error('Unknown server function: ' + fnName);
  const fn = globalThis[fnName];
  if (typeof fn !== 'function') throw new Error('Unknown server function: ' + fnName);
  // 本物のメールがある人は番号を使わない (管理者がこの入口を通っても、管理者のまま)
  __guestIdForThisRun = guestRealEmail_() ? null : String(id);
  try {
    return fn.apply(null, Array.isArray(args) ? args : []);
  } finally {
    __guestIdForThisRun = null;
  }
}
