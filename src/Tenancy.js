/**
 * @fileoverview 配置の種類。本体 (みんなで 1 つ) では空のまま。
 *
 * 先生 1 人専用の配置 (scripts/deploy-single.js) では、このファイルだけが差し替えられ、
 *   DEPLOY_TENANCY = 'single' と、その先生のメール・DB にするスプレッドシートが入る。
 *   値はプロパティより先に効く (isSingleTenant_ / ensureSingleTenantBootstrap_)。
 *   Why ファイルで渡すか: 管理者を「最初に開いた人」で決めると、児童が先に開いたときに
 *   児童が管理者になる。配置する人が決めた値をコードに焼き込めば、その隙間がない。
 */
const DEPLOY_TENANCY = '';
const DEPLOY_ADMIN_EMAIL = '';
const DEPLOY_DATABASE_SPREADSHEET_ID = '';
