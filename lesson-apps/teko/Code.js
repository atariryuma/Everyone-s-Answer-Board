/**
 * @fileoverview 授業専用アプリ (スプレッドシート紐づけ版)。1 授業 = 1 スプレッドシート。
 *
 * なぜ本体 (みんなの回答ボード) と別に作るか:
 *   本体は「開いた人の権限で動く (USER_ACCESSING)」ため、児童アカウント側の設定が変わると
 *   入口で「このウェブアプリへのアクセス権がありません」になり、こちらでは直せない (2026-10-29)。
 *   このアプリは「自分 (教師) として実行」し、児童は自分の権限を何も使わない。
 *   教師 1 人の専用アプリなので、他の先生の Drive やシートを考えなくてよい。
 *
 * 授業ごとに変えるのは LESSON だけ。コピーして書き換えれば次の授業になる。
 *
 * 守っていること (本体の授業モードと同じ):
 *   - 何ができるか (書く / 読む / 止まる) はフェーズが決め、サーバで検証する (画面の表示は権限ではない)
 *   - 「考える」では他の児童の回答を返さない
 *   - 児童同士に名前を見せない (名前・メールは教師の画面だけ)
 *   - 1 フェーズ 1 人 1 行。同じフェーズで送り直したら行を更新する
 */

/* ── 授業ごとに書き換えるところ ───────────────────────────── */
const LESSON = {
  title: 'てこのはたらきとしくみ（6年理科）',
  classes: ['6年1組', '6年2組', '6年3組', '6年4組'],
  // 座標系は授業を通して 1 つ (● と ★ を同じ地図で比べるため)。
  //   横 5 列 (まん中 = 変わらない)、縦 5 段 (下 = 迷いあり / 上 = 迷いなし)。
  axis: {
    xLow: '手ごたえが小さくなる', xMid: '変わらない', xHigh: '手ごたえが大きくなる',
    yLow: '迷いあり', yHigh: '迷いなし'
  },
  // role: input (書く) / browse (読む) / discuss (画面をとじる) / reinput (書き直す) / reflect (ふりかえる)
  phases: [
    { name: '考える', role: 'input',
      question: 'おもり（作用点）を支点に近づけると、手ごたえはどうなるだろう？ 変わらないと思ったら、まん中に。' },
    { name: '出会う', role: 'browse',
      question: '友達はどこに、なぜ置いたのだろう。自分とちがう予想の理由を読んでみよう' },
    { name: '議論する', role: 'discuss',
      question: '意見が分かれたところは、どうやって確かめる？' },
    { name: 'もう一度考える', role: 'reinput',
      question: '実験をしたいま、おもりを支点に近づけると、手ごたえはどうなる？' },
    { name: 'ふりかえる', role: 'reflect',
      question: '予想と結果をくらべて、自分の考えはどう変わった？ 変わらなかった？' }
  ]
};
/* ─────────────────────────────────────────────────────── */

const GRID = 5;                 // 横 5 × 縦 5
const REASON_MAX = 400;
const SHEET_ANSWERS = '回答';
const ANSWER_HEADERS = ['時刻', 'メールアドレス', 'クラス', 'フェーズ', '横', '縦', '理由'];
const COL = { ts: 0, email: 1, cls: 2, phase: 3, x: 4, y: 5, reason: 6 };
const CACHE_TTL = 10;           // 秒。書き込み時は version を上げるので古い値は返らない

function doGet(e) {
  const view = (e && e.parameter && e.parameter.view) === 'teacher' ? 'Teacher' : 'Student';
  const t = HtmlService.createTemplateFromFile(view);
  t.lessonTitle = LESSON.title;
  return t.evaluate()
    .setTitle(LESSON.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ── 本人と教師の判定 ───────────────────────────────────── */

function currentEmail_() {
  // 「自分として実行」でも、同じ Google Workspace ドメインなら開いた人のメールが取れる。
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}

function isTeacher_(email) {
  if (!email) return false;
  // 「自分として実行」では effective user = このアプリを出した教師。比較にだけ使う (権限は増やさない)。
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (email === owner) return true;
  const extra = String(PropertiesService.getScriptProperties().getProperty('TEACHERS') || '')
    .toLowerCase().split(/[,\s]+/).filter(Boolean);
  return extra.indexOf(email) !== -1;
}

function requireTeacher_() {
  const email = currentEmail_();
  if (!isTeacher_(email)) throw new Error('この操作は先生だけができます');
  return email;
}

/* ── 授業の状態 (いまの学級・フェーズ) ──────────────────── */

function readState_() {
  const p = PropertiesService.getScriptProperties();
  const phase = Number(p.getProperty('PHASE'));
  const cls = p.getProperty('CLASS');
  return {
    phase: Number.isInteger(phase) && phase >= 0 && phase < LESSON.phases.length ? phase : 0,
    cls: LESSON.classes.indexOf(cls) !== -1 ? cls : LESSON.classes[0],
    version: Number(p.getProperty('VERSION')) || 0
  };
}

function bumpVersion_() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('VERSION', String((Number(p.getProperty('VERSION')) || 0) + 1));
}

// 投稿を受け付けないフェーズは、直前の入力フェーズの回答を見る (自分のフェーズには回答が無いため)。
function sourcePhase_(phaseIndex) {
  for (let i = phaseIndex; i >= 0; i--) {
    const role = LESSON.phases[i].role;
    if (role === 'input' || role === 'reinput') return i;
  }
  return 0;
}

function firstInputPhase_() {
  return LESSON.phases.findIndex(p => p.role === 'input');
}

/* ── シート I/O ─────────────────────────────────────────── */

function answersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_ANSWERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ANSWERS);
    sh.getRange(1, 1, 1, ANSWER_HEADERS.length).setValues([ANSWER_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 学級の全回答 (全フェーズ)。児童 30 人が数秒おきに読むので、version 付き key で cache する。
function classRows_(state) {
  const key = 'rows:' + state.version + ':' + state.cls;
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) { /* 壊れた cache は読み直せばよい */ }
  }
  const values = answersSheet_().getDataRange().getValues().slice(1);
  const rows = values
    .filter(r => r[COL.cls] === state.cls)
    .map(r => ({
      email: String(r[COL.email]).toLowerCase(),
      phase: Number(r[COL.phase]),
      x: Number(r[COL.x]),
      y: Number(r[COL.y]),
      reason: String(r[COL.reason] || '')
    }));
  try {
    cache.put(key, JSON.stringify(rows), CACHE_TTL);
  } catch (_) { /* 100KB 超などで入らなくても、次回また読むだけ */ }
  return rows;
}

/* ── 児童の API ─────────────────────────────────────────── */

/**
 * 児童の画面が数秒おきに呼ぶ。いまのフェーズ (未来のフェーズは見せない) と、
 * そのフェーズで見てよいものだけを返す。
 */
function getStudentView() {
  const email = currentEmail_();
  if (!email) throw new Error('学校のアカウントでログインしてください');
  const state = readState_();
  const phase = LESSON.phases[state.phase];
  const rows = classRows_(state);
  const mine = rows.filter(r => r.email === email);
  const myAnswer = i => {
    const r = mine.find(m => m.phase === i);
    return r ? { x: r.x, y: r.y, reason: r.reason } : null;
  };

  const firstInput = firstInputPhase_();
  const view = {
    title: LESSON.title,
    cls: state.cls,
    axis: LESSON.axis,
    grid: GRID,
    phaseIndex: state.phase,
    phase: { name: phase.name, role: phase.role, question: phase.question },
    mine: myAnswer(state.phase),
    first: myAnswer(firstInput),       // さいしょの考え (● として、読むだけで出す)
    others: []
  };

  // 「考える」の間は他の児童の回答を返さない (隠すのではなく、送らない)。
  if (phase.role !== 'input') {
    const src = sourcePhase_(state.phase);
    view.sourcePhaseName = LESSON.phases[src].name;
    view.others = rows
      .filter(r => r.phase === src)
      .map(r => ({ x: r.x, y: r.y, reason: r.reason, isMine: r.email === email }));
  }
  if (phase.role === 'reflect') {
    const reinput = LESSON.phases.findIndex(p => p.role === 'reinput');
    view.second = reinput >= 0 ? myAnswer(reinput) : null;
  }
  return view;
}

/** 児童の送信。書けるフェーズかどうかはここで決める (画面の状態は信用しない)。 */
function submitAnswer(payload) {
  const email = currentEmail_();
  if (!email) throw new Error('学校のアカウントでログインしてください');
  const x = Number(payload && payload.x);
  const y = Number(payload && payload.y);
  const reason = String((payload && payload.reason) || '').trim().slice(0, REASON_MAX);
  if (!Number.isInteger(x) || !Number.isInteger(y) || x < 1 || x > GRID || y < 1 || y > GRID) {
    return { ok: false, code: 'BAD_POSITION', message: '位置をえらんでください' };
  }
  if (!reason) return { ok: false, code: 'NO_REASON', message: '理由を書いてください' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, code: 'BUSY', message: '混み合っています。もう一度「送る」を押してください' };
  }
  try {
    const state = readState_();
    const role = LESSON.phases[state.phase].role;
    // 送信を押した後にフェーズが進んでいたら受け付けない (前のフェーズには送れない)。
    if (Number(payload.phaseIndex) !== state.phase) {
      return { ok: false, code: 'PHASE_CHANGED', message: 'いまは「' + LESSON.phases[state.phase].name + '」です' };
    }
    if (role !== 'input' && role !== 'reinput') {
      return { ok: false, code: 'NOT_INPUT_PHASE', message: 'いまは送る時間ではありません' };
    }
    const sh = answersSheet_();
    const values = sh.getDataRange().getValues();
    const row = [new Date(), email, state.cls, state.phase, x, y, reason];
    // 同じ学級・フェーズ・本人の行があれば更新 (置き直しを許す)。
    let target = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][COL.email]).toLowerCase() === email &&
          values[i][COL.cls] === state.cls && Number(values[i][COL.phase]) === state.phase) {
        target = i + 1;
        break;
      }
    }
    if (target > 0) {
      sh.getRange(target, 1, 1, row.length).setValues([row]);
    } else {
      // appendRow は表の範囲を推定して列がずれることがあるので、行番号を決めて書く。
      sh.getRange(values.length + 1, 1, 1, row.length).setValues([row]);
    }
    bumpVersion_();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ── 教師の API ─────────────────────────────────────────── */

function getTeacherView() {
  requireTeacher_();
  const state = readState_();
  const rows = classRows_(state);
  const src = sourcePhase_(state.phase);
  const counts = LESSON.phases.map((p, i) => rows.filter(r => r.phase === i).length);
  return {
    title: LESSON.title,
    classes: LESSON.classes,
    cls: state.cls,
    axis: LESSON.axis,
    grid: GRID,
    phases: LESSON.phases.map(p => ({ name: p.name, role: p.role, question: p.question })),
    phaseIndex: state.phase,
    sourcePhaseIndex: src,
    counts: counts,
    // 名前 (メールの @ より前) は教師の画面だけに出す。
    answers: rows.map(r => ({
      who: r.email.split('@')[0], phase: r.phase, x: r.x, y: r.y, reason: r.reason
    }))
  };
}

function setPhase(index) {
  requireTeacher_();
  const i = Number(index);
  if (!Number.isInteger(i) || i < 0 || i >= LESSON.phases.length) throw new Error('フェーズの番号が正しくありません');
  PropertiesService.getScriptProperties().setProperty('PHASE', String(i));
  bumpVersion_();
  return getTeacherView();
}

/** 学級を切り替える。前の学級の回答が次の学級に混ざらないよう、フェーズは最初に戻す。 */
function setClass(cls) {
  requireTeacher_();
  if (LESSON.classes.indexOf(cls) === -1) throw new Error('学級名が正しくありません');
  const p = PropertiesService.getScriptProperties();
  p.setProperties({ CLASS: cls, PHASE: '0' });
  bumpVersion_();
  return getTeacherView();
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
