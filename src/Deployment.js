/**
 * @fileoverview この配置の先生 (管理者) と、データを入れるスプレッドシート。
 *
 * リポジトリでは空。scripts/deploy-single.js が先生ごとにこのファイルだけを書き換えて配置する。
 *   値はプロパティより先に効く (helpers.js の ensureDeploymentBootstrap_ が初回にプロパティへ写す)。
 *   Why ファイルで渡すか: 管理者を「最初に開いた人」で決めると、児童が先に開いたときに
 *   児童が管理者になる。配置する人が決めた値をコードに焼き込めば、その隙間がない。
 *
 * 空のままスプレッドシートに紐づけて配置したとき (ひな形をコピーして配った配置。docs/DISTRIBUTION.md) は、
 *   そのスプレッドシートの持ち主が管理者に、そのスプレッドシートが DB になる
 *   (helpers.js の ensureCopiedDeploymentBootstrap_)。
 */
const DEPLOY_ADMIN_EMAIL = '';
const DEPLOY_DATABASE_SPREADSHEET_ID = '';
// ops-monitor への報告の送り先と鍵（OpsReport.js）。空なら送らない。
const DEPLOY_OPS_REPORT_URL = '';
const DEPLOY_OPS_REPORT_TOKEN = '';
