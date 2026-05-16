/**
 * @fileoverview DataApis - データ操作 API（公開ボードのデータ取得、列分析、
 *   profile 切替、フォルダ作成、リアクションのバッチ取得など）。クロスファイルの
 *   依存関係は下の global 宣言を参照。
 */

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, findPublishedBoardOwner, getUserConfig, getConfigOrDefault, DEFAULT_DISPLAY_SETTINGS, saveUserConfig, openSpreadsheet, getSheetInfo, getUserSheetData, getBatchedAdminAuth, getFormInfo, invalidateSheetHeadersCache, performIntegratedColumnDiagnostics, setupDomainWideSharing, applySpreadsheetSharingDefaults, validateAccess, createAuthError, createUserNotFoundError, createErrorResponse, createExceptionResponse, emailToShortHash */
// GAS built-ins (DriveApp, SpreadsheetApp, ScriptApp, URL, FormApp, UrlFetchApp, Utilities, Session)
// は eslint.config.js の globals に登録済み — ここで再宣言しない。

/**
 * ユーザー専用フォルダを作成
 * 各ユーザーのマイドライブに「みんなの回答ボード」フォルダを作成
 *
 * Why owner-filter: getFoldersByName() は共有されてきた同名フォルダも返すため、
 *   他人が共有した「みんなの回答ボード」が先にヒットすると新規ファイルがそこへ
 *   流れ込みプライバシー事故になりうる。getOwner().getEmail() == 自分 を必須にする。
 *
 * @returns {GoogleAppsScript.Drive.Folder|null} 作成/取得したフォルダ
 */
// 教材画像 (Form 添付用) のサイズ・形式ガード。
//   Drive にアップロード後に form.addImageItem で埋め込むので、
//   過大ファイルは reject (児童端末の通信負荷 + Forms 表示崩れ防止)。
const LESSON_IMAGE_MAX_BYTES = 5 * 1024 * 1024;  // 5 MB
const LESSON_IMAGE_ALLOWED_MIMES = Object.freeze({
  'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg',
  'image/gif': 'gif', 'image/webp': 'webp'
});

/**
 * 教材画像をユーザー Drive フォルダに保存し、後で Form 添付に使う fileId を返す。
 *   フロント (wizard) から drag&drop / file picker / paste で取得した画像を base64 で受け取り、
 *   Blob にデコードしてユーザーの Drive folder に保存する。教師は URL を一切意識しない。
 *
 * @param {string} base64Data - "data:image/png;base64,xxxx" 形式の data URL or 純粋な base64
 * @param {string} [filename] - 任意のファイル名 (省略時 image-<timestamp>.<ext>)
 * @returns {Object} { success, fileId, fileName, mimeType, size }
 */
function uploadLessonImage(base64Data, filename) {
  try {
    const email = getCurrentEmail();
    if (!email) return createAuthError();
    if (typeof base64Data !== 'string' || base64Data.length === 0) {
      return createErrorResponse('画像データが指定されていません');
    }
    // data URL prefix を剥がして MIME / payload を分離
    let mimeType = '';
    let payload = base64Data;
    const dataUrlMatch = base64Data.match(/^data:([\w\/\-+.]+);base64,(.+)$/);
    if (dataUrlMatch) {
      mimeType = dataUrlMatch[1].toLowerCase();
      payload = dataUrlMatch[2];
    }
    if (!mimeType || !LESSON_IMAGE_ALLOWED_MIMES[mimeType]) {
      return createErrorResponse('対応していない画像形式です (png/jpeg/gif/webp のみ)');
    }
    // base64 のおおよそのサイズ (4 文字 → 3 bytes)
    const approxBytes = Math.floor(payload.length * 3 / 4);
    if (approxBytes > LESSON_IMAGE_MAX_BYTES) {
      return createErrorResponse(`画像サイズが大きすぎます (${Math.round(approxBytes / 1024 / 1024)} MB) — 5 MB 以下にしてください`);
    }
    const bytes = Utilities.base64Decode(payload);
    const ext = LESSON_IMAGE_ALLOWED_MIMES[mimeType];
    // 許可文字: ASCII 英数記号 + Hiragana/Katakana/CJK Ideographs。
    //   範囲は ぁ (ぁ) から 鿿 (CJK 末端)。ASCII 制御文字や全角空白は除外。
    const safeName = (typeof filename === 'string' && filename.trim())
      ? filename.trim().slice(0, 80).replace(/[^a-zA-Z0-9._\-ぁ-鿿]/g, '_')
      : ('image-' + Date.now() + '.' + ext);
    const blob = Utilities.newBlob(bytes, mimeType, safeName);
    const folder = createUserFolder();
    if (!folder) return createErrorResponse('Drive フォルダ作成に失敗しました');
    const file = folder.createFile(blob);
    return createSuccessResponse('uploaded', {
      fileId: file.getId(),
      fileName: file.getName(),
      mimeType,
      size: approxBytes
    });
  } catch (error) {
    logError_('uploadLessonImage', error);
    return createExceptionResponse(error);
  }
}

function createUserFolder() {
  try {
    const folderName = 'みんなの回答ボード';
    const myEmail = Session.getActiveUser().getEmail();

    const folders = DriveApp.getFoldersByName(folderName);
    while (folders.hasNext()) {
      const f = folders.next();
      try {
        const owner = f.getOwner();
        if (owner && owner.getEmail && owner.getEmail() === myEmail) {
          return f;
        }
      } catch (ownerErr) {
        // getOwner() は権限不足で失敗しうる（共有ドライブ等）。その場合は無視して次へ。
      }
    }
    return DriveApp.createFolder(folderName);

  } catch (e) {
    console.error('createUserFolder エラー:', e.message);
    return null;
  }
}

/**
 * ファイルを指定フォルダへ移動（Drive 単一親モデル準拠）
 * Why: addFile/removeFile は multi-parent 時代の deprecated API。
 *      moveTo() は2020年以降の正式 API で、ルートフォルダからの除外も自動で行う。
 * @param {GoogleAppsScript.Drive.File} file
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @private
 */
function moveFileToFolder_(file, folder) {
  if (!file || !folder) return;
  if (typeof file.moveTo === 'function') {
    file.moveTo(folder);
    return;
  }
  // 古い GAS 環境向けフォールバック（CI スタブ含む）
  folder.addFile(file);
  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    // 既に root に無いなら無視
  }
}

/**
 * 自分がオーナーのGoogleフォームを30件取得
 * @returns {Object} フォーム一覧
 */
function getForms() {
  try {
    const files = DriveApp.getFilesByType('application/vnd.google-apps.form');
    const forms = [];

    while (files.hasNext() && forms.length < 30) {
      const file = files.next();
      forms.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl()
      });
    }

    return { success: true, forms };
  } catch (error) {
    logError_('getForms', error);
    return { success: false, error: error.message, forms: [] };
  }
}

// Why: 線形尺度の範囲は detectNumericScaleColumns の判定（0-10 の整数, レンジ 1-9）と
//      整合する必要がある。テンプレートと検出ロジックが乖離すると、テンプレートで作った
//      M1/M2 フォームの数値列が auto モードで検出されない事故になる。
const TEMPLATE_SCALE_MIN = 1;
const TEMPLATE_SCALE_MAX = 5;
// 児童の学年に合わせて選べる粒度 (Forms LinearScale の max は 1〜10、min は 0〜1)。
//   3 = 低学年向け (賛成 / どちらでもない / 反対)
//   5 = 既定 (中学年〜)
//   7 = 高学年向け (細かい揺らぎを表現)
const TEMPLATE_SCALE_POINTS_ALLOWED = Object.freeze([3, 5, 7]);
const TEMPLATE_SCALE_POINTS_DEFAULT = 5;

const TEMPLATE_LABELS = Object.freeze({
  board: '掲示板',
  numberline: '数直線',
  matrix: 'マトリクス',
  pie: '円グラフ'
});

function addScaleItemTo_(form, { title, helpText, lowLabel, highLabel, scalePoints }) {
  const points = TEMPLATE_SCALE_POINTS_ALLOWED.includes(scalePoints)
    ? scalePoints : TEMPLATE_SCALE_POINTS_DEFAULT;
  form.addScaleItem()
    .setTitle(title)
    .setRequired(true)
    .setHelpText(helpText)
    .setBounds(TEMPLATE_SCALE_MIN, points)
    .setLabels(lowLabel, highLabel);
}

/**
 * テンプレートフォームを作成
 *
 * templateType で生成されるフォームの形が変わる：
 *   - 'board'      (既定): クラス / 名前 / 回答(選択肢) / 理由            — 既存掲示板モード
 *   - 'numberline' (M1)  : クラス / 名前 / 立場(線形尺度1-5) / 理由       — 数直線可視化
 *   - 'matrix'     (M2)  : クラス / 名前 / X軸(線形尺度) / Y軸(線形尺度) / 理由 — 散布図
 *
 * Why M1/M2 テンプレートを用意するか: detectNumericScaleColumns は実データから
 *   線形尺度列を検出するが、空のフォームでは検出できないので教師に手動で
 *   ScaleItem を追加させる手間が残る。テンプレート時点で M1/M2 用の構造を作れば
 *   そのまま「✨ 自動判定」で boardMode=numberline/matrix に乗る。
 *
 * ベストプラクティス: https://developers.google.com/apps-script/reference/forms/form
 *
 * @param {string} [templateType='board'] - 'board' | 'numberline' | 'matrix'
 * @returns {Object} 作成結果（フォームURL、スプレッドシートID等）
 */
function createTemplateForm(templateType, templateOptions) {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) return createAuthError();

    const type = TEMPLATE_LABELS[templateType] ? templateType : 'board';
    const opts = (templateOptions && typeof templateOptions === 'object') ? templateOptions : {};
    const safeStr = (v, fallback, max) => {
      const s = String(v == null ? '' : v).trim().slice(0, max || 40);
      return s || fallback;
    };
    // choices: 多肢選択用 (2〜8)。class 用には maxLength を上げた専用版を別途用意。
    const safeChoices = (arr, fallback) => {
      if (!Array.isArray(arr)) return fallback;
      const cleaned = arr.map((c) => String(c == null ? '' : c).trim().slice(0, 40)).filter(Boolean);
      if (cleaned.length < 2 || cleaned.length > 8) return fallback;
      return cleaned;
    };
    // クラス選択肢: 1 件以上 (1 クラスでも有効) / 最大 30 (学年×組で現実的に上限)。
    const safeClassChoices = (arr, fallback) => {
      if (!Array.isArray(arr)) return fallback;
      const cleaned = arr.map((c) => String(c == null ? '' : c).trim().slice(0, 30)).filter(Boolean);
      if (cleaned.length < 1 || cleaned.length > 30) return fallback;
      return cleaned;
    };

    // ----- タイトル: UI で指定された lesson 名 / phase 名を最優先 -----
    //   児童が Form 一覧で「何の Form か」を即判別できるようにするため、
    //   従来の "回答ボード [type] M/d HH:mm" は lessonName 未指定時のみ。
    const lessonName = safeStr(opts.lessonName, '', 60);
    const phaseName = safeStr(opts.phaseName, '', 30);
    let formTitle;
    if (lessonName && phaseName && lessonName !== phaseName) {
      formTitle = `${lessonName} / ${phaseName}`;
    } else if (lessonName) {
      formTitle = lessonName;
    } else if (phaseName) {
      formTitle = phaseName;
    } else {
      const now = new Date();
      const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d HH:mm');
      formTitle = `回答ボード [${TEMPLATE_LABELS[type]}] ${dateStr}`;
    }
    const form = FormApp.create(formTitle);

    // ----- フォーム全体の description: 教師が wizard で入れた question を出す -----
    //   児童は Form の冒頭にここを読む。テーマや問いを最初に提示できる。
    const question = safeStr(opts.question, '', 200);
    if (question) {
      try { form.setDescription(question); } catch (descErr) { logError_('createTemplateForm:setDescription', descErr); }
    }

    // Why VERIFIED via REST: Apps Script の setCollectEmail(true) では「回答者入力」モード
    // になりなりすまし可能。Forms REST API で emailCollectionType=VERIFIED を強制する。
    setFormEmailCollectionVerified_(form.getId());

    // 1 回答制限: opts.allowResubmit=true なら OFF (揺らぎ追跡が有効)、それ以外は ON (既定)。
    //   UI の「再投稿を許可（揺らぎを追跡）」チェックボックスを反映。
    //   注: UI で後から allowResubmit を toggle した場合は Form 側を別経路で更新する必要がある。
    form.setLimitOneResponsePerUser(!opts.allowResubmit);

    // ----- クラス選択肢: UI で教師が選んだ classes を反映 -----
    //   児童は自分の所属クラスを Form 内で選ぶので、教師指定の一覧と完全一致させる。
    const classChoices = safeClassChoices(opts.classChoices, ['クラス1', 'クラス2', 'クラス3', 'クラス4']);
    form.addListItem()
      .setTitle('クラス')
      .setRequired(true)
      .setHelpText('自分のクラスを選んでください')
      .setChoiceValues(classChoices);
    form.addTextItem()
      .setTitle('名前')
      .setRequired(true)
      .setHelpText('名前を入力してください');

    // ----- 視覚区切り: ここから「今日のテーマ」セクション -----
    //   教師の問い (question) があれば SectionHeader として強調する。
    //   児童は description より SectionHeader title の方をスキャンしやすい (Forms の視覚階層)。
    if (question) {
      try {
        form.addSectionHeaderItem()
          .setTitle('今日のテーマ')
          .setHelpText(question);
      } catch (sectionErr) {
        logError_('createTemplateForm:addSectionHeaderItem', sectionErr);
      }
    }

    // ----- 教材画像: Drive fileId 経由 (推奨) または URL fetch (legacy) -----
    //   優先: opts.imageFileId (uploadLessonImage が返す Drive file ID)。
    //   フォールバック: opts.imageUrl (旧パス)。失敗は warning で fail-soft。
    //   児童は教材タブを行き来せずに Form 内で問題と画像を一緒に見られる。
    const imageFileId = safeStr(opts.imageFileId, '', 80);
    const imageUrl = safeStr(opts.imageUrl, '', 1024);
    if (imageFileId || imageUrl) {
      try {
        let blob = null;
        if (imageFileId) {
          blob = DriveApp.getFileById(imageFileId).getBlob();
        } else {
          blob = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true }).getBlob();
        }
        const imgItem = form.addImageItem().setImage(blob);
        if (typeof imgItem.setTitle === 'function') imgItem.setTitle('教材');
      } catch (imgErr) {
        logError_('createTemplateForm:addImageItem', imgErr);
        console.warn('Form 画像埋め込みに失敗 (処理は続行):', imgErr && imgErr.message);
      }
    }

    const scalePoints = TEMPLATE_SCALE_POINTS_ALLOWED.includes(opts.scalePoints)
      ? opts.scalePoints : TEMPLATE_SCALE_POINTS_DEFAULT;

    if (type === 'numberline') {
      // scaleTitle は質問文 (UI の phase.question) を優先。短ければそのまま、長ければ phase 名にフォールバック。
      const numberlineTitle = safeStr(opts.scaleTitle, question || phaseName || '立場', 60);
      addScaleItemTo_(form, {
        title: numberlineTitle,
        helpText: 'あなたの考えにいちばん近いところを選んでください',
        lowLabel: safeStr(opts.lowLabel, 'そう思わない'),
        highLabel: safeStr(opts.highLabel, 'とてもそう思う'),
        scalePoints
      });
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('そこを選んだ理由を書いてください');
    } else if (type === 'matrix') {
      // matrix は X / Y 2 軸が両方必要。各軸 title は教師指定 (xTitle/yTitle) を尊重し、
      //   空なら lowLabel↔highLabel から自動生成 (例: "効率↓↔効率↑")。
      const xAuto = (opts.xLow || opts.xHigh) ? `${safeStr(opts.xLow, '低い', 20)} ↔ ${safeStr(opts.xHigh, '高い', 20)}` : 'X軸';
      const yAuto = (opts.yLow || opts.yHigh) ? `${safeStr(opts.yLow, '低い', 20)} ↔ ${safeStr(opts.yHigh, '高い', 20)}` : 'Y軸';
      addScaleItemTo_(form, {
        title: safeStr(opts.xTitle, xAuto, 60),
        helpText: '横の軸であなたの考えにいちばん近いところを選んでください',
        lowLabel: safeStr(opts.xLow, '低い'),
        highLabel: safeStr(opts.xHigh, '高い'),
        scalePoints
      });
      addScaleItemTo_(form, {
        title: safeStr(opts.yTitle, yAuto, 60),
        helpText: '縦の軸であなたの考えにいちばん近いところを選んでください',
        lowLabel: safeStr(opts.yLow, '低い'),
        highLabel: safeStr(opts.yHigh, '高い'),
        scalePoints
      });
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('そこを選んだ理由を書いてください');
    } else if (type === 'pie') {
      form.addMultipleChoiceItem()
        .setTitle(safeStr(opts.choiceTitle, question || phaseName || 'あなたの選択', 60))
        .setRequired(true)
        .setHelpText('選択肢から 1 つ選んでください')
        .setChoiceValues(safeChoices(opts.choices, ['A', 'B', 'どちらとも言えない']));
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('そう選んだ理由を書いてください');
    } else {
      // board (掲示板)
      form.addMultipleChoiceItem()
        .setTitle(safeStr(opts.choiceTitle, question || phaseName || '回答', 60))
        .setRequired(true)
        .setHelpText('選択肢から 1 つ選んでください')
        .setChoiceValues(safeChoices(opts.choices, ['賛成', '反対', 'どちらでもない']));
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('選んだ理由を書いてください');
    }

    // ----- 送信後の確認メッセージ: 児童に達成感を与え、再投稿可否も伝える -----
    try {
      const confirmation = opts.allowResubmit
        ? 'ありがとう！意見が届いたよ。\nもう一度考えが変わったら、また回答できるよ。'
        : 'ありがとう！意見が届いたよ。';
      form.setConfirmationMessage(confirmation);
    } catch (confErr) {
      logError_('createTemplateForm:setConfirmationMessage', confErr);
    }

    const ss = SpreadsheetApp.create(formTitle + ' (回答)');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    let folderUrl = '';
    try {
      const folder = createUserFolder();
      if (folder) {
        folderUrl = folder.getUrl();
        moveFileToFolder_(DriveApp.getFileById(form.getId()), folder);
        moveFileToFolder_(DriveApp.getFileById(ss.getId()), folder);
      }
    } catch (folderError) {
      console.warn('フォルダ作成/移動エラー（処理は続行）:', folderError.message);
    }

    // SS 共有デフォルト適用：サービスアカウント editor + ドメイン共有。
    // Why: サーバ側 DatabaseCore は Sheets REST API + Service Account JWT で SS を読む。
    //   未共有のままだと view モードで `getPublishedSheetData` が 403 で落ちる。
    //   作成直後に必ず適用し、教師の手動共有作業を不要にする。fail-soft。
    let sharingResult = null;
    try {
      sharingResult = applySpreadsheetSharingDefaults(ss.getId(), currentEmail);
      if (sharingResult && sharingResult.errors && sharingResult.errors.length) {
        console.warn('createTemplateForm sharing warnings:', sharingResult.errors.join(' / '));
      }
    } catch (sharingError) {
      console.warn('SS 共有設定エラー（処理は続行）:', sharingError.message);
    }

    return {
      success: true,
      formUrl: form.getPublishedUrl(),
      editUrl: form.getEditUrl(),
      formId: form.getId(),
      formTitle: formTitle,
      templateType: type,
      spreadsheetId: ss.getId(),
      sharing: sharingResult,
      spreadsheetUrl: ss.getUrl(),
      spreadsheetName: ss.getName(),
      sheetName: 'フォームの回答 1',
      folderUrl: folderUrl,
      wasCreated: true,
      message: `テンプレート（${TEMPLATE_LABELS[type]}）を作成しました`
    };

  } catch (error) {
    logError_('createTemplateForm', error);
    return { success: false, error: 'フォーム作成に失敗しました: ' + error.message };
  }
}

/**
 * 既存 Form の質問項目をスキーマで完全置換する。
 *
 * Why: CLI から授業用に Forms をカスタマイズしたい (4 クラス選択肢、線形尺度ラベル、
 *   設問文の書換等)。createTemplateForm は固定テンプレなので、その後の編集を
 *   automate する必要がある。
 *
 * 動作:
 *   1. 既存アイテム全削除（spreadsheet との紐付けは保たれる）
 *   2. schema.questions[] を順に addXxxItem() で再構築
 *   3. 任意で form の title/description も更新
 *
 *   注意: 既に回答が入っている Forms に対して実行すると、Spreadsheet の列構造が変わるため
 *   過去回答との対応が崩れる。新規作成直後（回答 0 件）の Forms に対してのみ使うべき。
 *
 * @param {string} formIdOrUrl - FormApp.openById|openByUrl 対応の id/URL
 * @param {Object} schema - { title?, description?, questions: [...] }
 *   各 question の shape:
 *     { type: 'list'|'text'|'multipleChoice'|'scale'|'paragraph',
 *       title: string, helpText?: string, required?: bool,
 *       choices?: string[], min?: number, max?: number, leftLabel?, rightLabel? }
 * @returns {Object} { success, formId, title, itemCount }
 */
/**
 * 既存 Form の「1 回答制限」を toggle する。
 *   再投稿許可 (= allowResubmit=true) なら setLimitOneResponsePerUser(false)。
 *   AdminPanel の「再投稿を許可（揺らぎを追跡）」チェックボックスから呼ばれる。
 *   Why 別関数: customizeForm は項目全削除を伴う破壊的 API。設定の 1 bit だけ
 *     触りたいケースでは使わない。回答が既に入っている Form でも安全に呼べる。
 * @param {string} formIdOrUrl
 * @param {boolean} allowResubmit
 * @returns {Object} { success, formId, allowResubmit }
 */
// formId / formUrl どちらでも開く共通ヘルパー。
//   失敗時は throw (呼び出し側で createErrorResponse に変換)。
function __openFormByIdOrUrl_(formIdOrUrl) {
  return formIdOrUrl.startsWith('http')
    ? FormApp.openByUrl(formIdOrUrl)
    : FormApp.openById(formIdOrUrl);
}

function setFormAllowResubmit(formIdOrUrl, allowResubmit) {
  try {
    if (!formIdOrUrl || typeof formIdOrUrl !== 'string') {
      return createErrorResponse('formId or formUrl is required');
    }
    let form;
    try {
      form = __openFormByIdOrUrl_(formIdOrUrl);
    } catch (e) {
      return createErrorResponse('Form を開けません: ' + e.message);
    }
    form.setLimitOneResponsePerUser(!allowResubmit);
    return createSuccessResponse('updated', {
      formId: form.getId(),
      allowResubmit: Boolean(allowResubmit)
    });
  } catch (error) {
    logError_('setFormAllowResubmit', error);
    return createExceptionResponse(error);
  }
}

function customizeForm(formIdOrUrl, schema) {
  try {
    if (!formIdOrUrl || typeof formIdOrUrl !== 'string') {
      return { success: false, error: 'formId or formUrl is required' };
    }
    if (!schema || typeof schema !== 'object' || !Array.isArray(schema.questions)) {
      return { success: false, error: 'schema.questions array is required' };
    }
    if (schema.questions.length > 40) {
      // Why: 過大スキーマで Forms API limit に当たる前にガード（実用上 40 で十分）
      return { success: false, error: 'too many questions (max 40)' };
    }

    let form;
    try {
      form = __openFormByIdOrUrl_(formIdOrUrl);
    } catch (e) {
      return { success: false, error: 'Form を開けません: ' + e.message };
    }

    // 既存項目を削除（後ろから消すと index ずれない）
    const existing = form.getItems();
    for (let i = existing.length - 1; i >= 0; i--) {
      try { form.deleteItem(existing[i]); } catch (_) { /* skip non-deletable */ }
    }

    // タイトル・説明
    //
    // Why: form.setTitle() は Forms 内部のタイトル（編集画面 / 回答画面に表示）を更新するが、
    //   Drive 上のファイル名は別管理で setName を呼ばないと変わらない。AdminPanel の
    //   「📋 一覧から選択」dropdown は DriveApp.getFiles() の name を表示するので、
    //   ここで file の name も同期しないと「タイトルは【導入】... なのに dropdown には
    //   "回答ボード [掲示板] 5/14 06:22" と出る」食い違いが起きる (2026-05-14 教師フィードバック)。
    if (typeof schema.title === 'string' && schema.title.trim()) {
      const cleanTitle = schema.title.trim().substring(0, 200);
      form.setTitle(cleanTitle);
      try {
        DriveApp.getFileById(form.getId()).setName(cleanTitle);
      } catch (renameErr) {
        // setName 失敗は致命的でない（権限不足等）。fall through して title だけ更新済の状態。
        console.warn('customizeForm: Drive file rename failed:', renameErr.message);
      }
    }
    if (typeof schema.description === 'string') {
      form.setDescription(schema.description.substring(0, 2000));
    }

    // 質問を順に追加
    let added = 0;
    for (const q of schema.questions) {
      if (!q || typeof q !== 'object' || !q.type || !q.title) continue;
      const title = String(q.title).substring(0, 200);
      const required = q.required !== false; // default true
      const helpText = q.helpText ? String(q.helpText).substring(0, 500) : '';
      const choices = Array.isArray(q.choices) ? q.choices.map(c => String(c).substring(0, 200)).filter(Boolean) : [];

      let item = null;
      switch (q.type) {
        case 'list':
          item = form.addListItem();
          if (choices.length) item.setChoiceValues(choices);
          break;
        case 'multipleChoice':
          item = form.addMultipleChoiceItem();
          if (choices.length) item.setChoiceValues(choices);
          break;
        case 'text':
          item = form.addTextItem();
          break;
        case 'paragraph':
          item = form.addParagraphTextItem();
          break;
        case 'scale': {
          item = form.addScaleItem();
          const min = Number.isFinite(q.min) ? q.min : 1;
          const max = Number.isFinite(q.max) ? q.max : 5;
          try { item.setBounds(min, max); } catch (e) {
            // setBounds は max-min >= 1, max <= 10 等の制約あり。失敗時はデフォルト 1-5。
            try { item.setBounds(1, 5); } catch (_) { /* give up */ }
          }
          const left = q.leftLabel ? String(q.leftLabel).substring(0, 100) : '';
          const right = q.rightLabel ? String(q.rightLabel).substring(0, 100) : '';
          if (left || right) item.setLabels(left, right);
          break;
        }
        default:
          continue; // unknown type → skip
      }
      if (!item) continue;

      item.setTitle(title);
      if (helpText) item.setHelpText(helpText);
      try { item.setRequired(required); } catch (_) { /* not all items support setRequired */ }
      added++;
    }

    // Why: Form の items を deleteItem → 再 addItem しても Spreadsheet 上の旧列は
    //      残り続け、新 items は別の新列として追加される。再 customize を繰り返すと
    //      列が累積する pathological bug。
    //
    //      根本解消策: destination Spreadsheet を新規作成し直す。Form は新しい SS に
    //      リンクされ、列は items の順序通りクリーンに構成される。旧 SS はトラッシュへ
    //      移動（即削除でなく猶予を与える）。
    //
    //      ただし「既に回答が入っている SS を捨てる」のはデータロスにつながるので、
    //      schema.resetDestination !== false（デフォルト true）のときのみ実行。
    //      回答済みのフォームを customize するなら明示的に { resetDestination: false } を渡す。
    let newSpreadsheetId = null;
    let trashedOldSpreadsheetId = null;
    const shouldReset = schema.resetDestination !== false;
    if (shouldReset) {
      try {
        const oldDestId = form.getDestinationId();
        // 旧 SS が空でない場合（既に回答あり）はリセットを諦めて警告
        let hasResponses = false;
        if (oldDestId) {
          try {
            const oldSs = SpreadsheetApp.openById(oldDestId);
            const sheets = oldSs.getSheets();
            for (const s of sheets) {
              if (s.getLastRow() > 1) { hasResponses = true; break; }
            }
          } catch (_) { /* skip */ }
        }
        if (hasResponses) {
          console.warn('customizeForm: skipped destination reset (responses exist in old spreadsheet)');
        } else {
          // 新 SS を作って再リンク。リンク前後でファイル所有者は CLI ユーザー（Form と同じ）。
          const ssTitle = form.getTitle() + ' (回答)';
          const newSs = SpreadsheetApp.create(ssTitle);
          form.setDestination(FormApp.DestinationType.SPREADSHEET, newSs.getId());
          newSpreadsheetId = newSs.getId();

          // 新 SS にも共有デフォルト適用（createTemplateForm と同じ理由）。
          // ここを忘れると customizeForm 経由で再生成された SS だけ view が 403 で死ぬ。
          try {
            const ownerEmail = (typeof getCurrentEmail === 'function') ? getCurrentEmail() : '';
            const sr = applySpreadsheetSharingDefaults(newSpreadsheetId, ownerEmail);
            if (sr && sr.errors && sr.errors.length) {
              console.warn('customizeForm sharing warnings:', sr.errors.join(' / '));
            }
          } catch (sharingError) {
            console.warn('customizeForm SS 共有設定エラー（処理は続行）:', sharingError.message);
          }

          // 旧 SS をトラッシュへ移動（即削除でなく Drive ゴミ箱経由で 30 日猶予）
          if (oldDestId && oldDestId !== newSpreadsheetId) {
            try {
              DriveApp.getFileById(oldDestId).setTrashed(true);
              trashedOldSpreadsheetId = oldDestId;
            } catch (_) { /* trash 失敗は致命的でない */ }
          }
        }
      } catch (e) {
        console.warn('customizeForm: destination reset failed:', e.message);
      }
    }

    // ヘッダーキャッシュ無効化（新 SS のキャッシュは未作成だが、安全のため旧キャッシュも消す）
    // Why (fail-soft): cache invalidation 失敗は customizeForm 全体を止める理由にならない。
    //   destination spreadsheet が trash 済 / アクセス権無し / cache 未初期化のいずれかが原因。
    //   次回 fetch 時にどちらにせよ headers は実際の sheet から取り直されるので無害。
    try {
      const destId = newSpreadsheetId || form.getDestinationId();
      if (destId && typeof invalidateSheetHeadersCache === 'function') {
        ['フォームの回答 1', 'Form Responses 1', 'Sheet1'].forEach(sn => {
          try { invalidateSheetHeadersCache(destId, sn); } catch (cacheErr) {
            // sheet 名違いは normal: 配列の他要素で当たる
          }
        });
      }
    } catch (destErr) {
      // form.getDestinationId() の失敗は珍しいが致命的でない
    }

    return {
      success: true,
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      editUrl: form.getEditUrl(),
      title: form.getTitle(),
      itemCount: added,
      newSpreadsheetId,
      trashedOldSpreadsheetId,
      message: `${added} 問の質問を再構築${newSpreadsheetId ? '・回答先 Spreadsheet を新規化' : ''}しました`
    };
  } catch (error) {
    logError_('customizeForm', error);
    return { success: false, error: error.message };
  }
}

/**
 * フォームURLから回答ボード接続情報を取得
 * スプレッドシートがなければ新規作成してリンク
 * @param {string} formUrl - GoogleフォームURL
 * @returns {Object} 接続情報
 */
function processFormUrlInput(formUrl) {
  try {
    // 1. URL検証
    if (!formUrl || typeof formUrl !== 'string') {
      return { success: false, error: 'GoogleフォームのURLを入力してください' };
    }

    // フォームURL形式チェック
    const isFormUrl = formUrl.includes('/forms/d/') || formUrl.includes('forms.gle/');
    if (!isFormUrl) {
      return { success: false, error: 'GoogleフォームのURLを入力してください' };
    }

    // 2. フォームを開く
    let form;
    try {
      form = FormApp.openByUrl(formUrl);
    } catch (e) {
      console.error('connectDataSource: FormApp.openByUrl failed:', e.message);
      return { success: false, error: `フォームにアクセスできません: ${e.message}` };
    }

    const formTitle = form.getTitle();
    const formId = form.getId();
    const publishedUrl = form.getPublishedUrl();

    // 3. 回答先スプレッドシートを検出
    //
    // Why try-catch: form.getDestinationId() は destination 未設定時に null を返さず、
    //   "フォームに応答先がありません。" という例外を投げる仕様 (Apps Script Forms API)。
    //   教師が手動で作った Form は典型的に destination が無いので、例外も null と同等扱いにして
    //   下の !spreadsheetId 分岐で auto-create に流す。
    let spreadsheetId = null;
    try {
      spreadsheetId = form.getDestinationId();
    } catch (noDestErr) {
      // expected: "フォームに応答先がありません。" — 未設定なので auto-create に進む
      spreadsheetId = null;
    }
    let wasCreated = false;
    let spreadsheetUrl = '';
    let ss;

    // 4. スプレッドシートがなければ新規作成
    if (!spreadsheetId) {
      try {
        ss = SpreadsheetApp.create(formTitle + ' (回答)');
        spreadsheetId = ss.getId();
        spreadsheetUrl = ss.getUrl();
        form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
        wasCreated = true;

        // 新規 SS 共有設定 (SA editor + ドメイン共有) — 生徒 view モード 403 防止。
        //   createTemplateForm / customizeForm 経路と同じ。Sharing が無いと viewer 403。
        try {
          const ownerEmail = (typeof getCurrentEmail === 'function') ? getCurrentEmail() : '';
          const sr = applySpreadsheetSharingDefaults(spreadsheetId, ownerEmail);
          if (sr && sr.errors && sr.errors.length) {
            console.warn('processFormUrlInput sharing warnings:', sr.errors.join(' / '));
          }
        } catch (sharingErr) {
          console.warn('processFormUrlInput SS 共有設定エラー（処理は続行）:', sharingErr.message);
        }
      } catch (e) {
        console.error('Spreadsheet creation error:', e.message);
        return { success: false, error: 'スプレッドシートの作成に失敗しました: ' + e.message };
      }
    } else {
      try {
        ss = SpreadsheetApp.openById(spreadsheetId);
        spreadsheetUrl = ss.getUrl();
      } catch (e) {
        console.error('connectDataSource: SpreadsheetApp.openById failed:', e.message);
        return { success: false, error: `回答先スプレッドシートにアクセスできません: ${e.message}` };
      }
    }

    // 5. 回答シート名を特定（新規作成時は「シート1」、既存は回答シート）
    const sheets = ss.getSheets();
    const responseSheet = sheets.find(s =>
      s.getName().includes('フォームの回答') ||
      s.getName().toLowerCase().includes('form response')
    ) || sheets[0];

    const sheetName = responseSheet.getName();

    return {
      success: true,
      formId,
      formTitle,
      formUrl: publishedUrl,
      spreadsheetId,
      spreadsheetUrl,
      spreadsheetName: ss.getName(),
      sheetName,
      wasCreated,
      message: wasCreated ? '回答先スプレッドシートを作成しました' : null
    };

  } catch (error) {
    logError_('processFormUrlInput', error);
    return { success: false, error: '予期しないエラーが発生しました' };
  }
}

/**
 * Validate header integrity for user's active sheet
 * @param {string} targetUserId - 対象ユーザーID（省略可能）
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(targetUserId) {
  try {
    const currentEmail = getCurrentEmail();
    let targetUser = targetUserId ? findUserById(targetUserId, { requestingUser: currentEmail }) : null;
    if (!targetUser && currentEmail) {
      targetUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    const config = getConfigOrDefault(targetUser.userId);

    if (!config.spreadsheetId || !config.sheetName) {
      return {
        success: false,
        error: 'Spreadsheet configuration incomplete'
      };
    }

    const dataAccess = openSpreadsheet(config.spreadsheetId, {
      useServiceAccount: false
    });
    const sheet = dataAccess.getSheet(config.sheetName);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found'
      };
    }

    const { headers } = getSheetInfo(sheet);
    const normalizedHeaders = headers.map(header => String(header || '').trim());
    const emptyHeaderCount = normalizedHeaders.filter(header => header.length === 0).length;
    const valid = normalizedHeaders.length > 0 && emptyHeaderCount < normalizedHeaders.length;

    return {
      success: valid,
      valid,
      headerCount: normalizedHeaders.length,
      emptyHeaderCount,
      headers: normalizedHeaders,
      error: valid ? null : 'ヘッダー情報が不完全です'
    };
  } catch (error) {
    logError_('validateHeaderIntegrity', error);
    return createExceptionResponse(error);
  }
}

/**
 * View 画面から profile を切替えるためのエンドポイント。
 *
 * Why: 教師がビュー画面の profile セレクタを操作 → google.script.run.loadProfileForView →
 *      該当ユーザーの config の profile を active に適用 → 200 OK で再ロードを促す。
 *      adminApi 経路の loadProfile と異なり、こちらは Session ベースの認証で動く。
 *      profile を切替えられるのは「自分のボード」または「管理者」のみ。
 *
 * @param {string} profileName
 * @param {string} [targetUserId] - 省略時は requester 自身
 * @returns {Object} { success, message, activeProfile }
 */
function loadProfileForView(profileName, targetUserId) {
  try {
    if (!profileName || typeof profileName !== 'string') {
      return createErrorResponse('profileName is required');
    }
    const email = getCurrentEmail();
    if (!email) return createAuthError();

    // 対象 user 判定
    let user;
    if (targetUserId && typeof targetUserId === 'string') {
      user = findUserById(targetUserId, { requestingUser: email });
    } else {
      user = findUserByEmail(email, { requestingUser: email });
    }
    if (!user) return createUserNotFoundError();

    // 認可: 自分のボードか、管理者か
    const isOwn = user.userEmail === email;
    const isAdmin = isAdministrator(email);
    if (!isOwn && !isAdmin) {
      return createErrorResponse('権限がありません: 自分以外のボードのプロファイル切替は管理者のみ可能です');
    }

    const cfgRes = getUserConfig(user.userId, user);
    if (!cfgRes.success) return createErrorResponse(cfgRes.message || 'config load failed');
    const cur = cfgRes.config || {};
    const profiles = Array.isArray(cur.profiles) ? cur.profiles : [];
    const p = profiles.find(x => x && x.name === profileName);
    if (!p) return createErrorResponse(`Profile "${profileName}" not found`);

    // Why (完全置換): matrix → wordcloud に切替えた時に旧 columnMapping.numericX/Y が
    //   居残るのを防ぐため、profile に無いキーは __applyProfileToConfig_ で明示上書き。
    const merged = {
      ...__applyProfileToConfig_(cur, p),
      activeProfile: p.name,
      userId: user.userId,
      profileHistory: __appendProfileHistory_(cur.profileHistory, p.name)
    };

    const saveRes = saveUserConfig(user.userId, merged, { isMainConfig: true });
    if (!saveRes.success) return saveRes;
    return createSuccessResponse('Profile loaded', { activeProfile: p.name });
  } catch (error) {
    console.error('loadProfileForView error:', error && error.message);
    return createExceptionResponse(error);
  }
}

// =====================================================================
// Profile management for AdminPanel (owner-facing).
// Why: 既存の listProfiles/saveProfile/deleteProfile は admin API 経由のみで、
//   AdminPanel の owner が直接 google.script.run できなかった。Session 認証で
//   「自分のボード」に限り profile CRUD できる薄いラッパーを用意する。
//   実装は AdminApis.js の __listProfilesCore / __saveProfileCore /
//   __deleteProfileCore に集約済み。ここでは認可（自分のボード限定）のみ担当。
// =====================================================================

function __resolveOwnUserId_() {
  const email = getCurrentEmail();
  if (!email) return { error: createAuthError() };
  const user = findUserByEmail(email, { requestingUser: email });
  if (!user) return { error: createUserNotFoundError() };
  return { userId: user.userId };
}

function listMyProfiles() {
  try {
    const resolved = __resolveOwnUserId_();
    if (resolved.error) return resolved.error;
    return __listProfilesCore(resolved.userId);
  } catch (error) {
    console.error('listMyProfiles error:', error && error.message);
    return createExceptionResponse(error);
  }
}

function saveCurrentAsProfile(name) {
  try {
    const resolved = __resolveOwnUserId_();
    if (resolved.error) return resolved.error;
    return __saveProfileCore(resolved.userId, name, { autoActivate: true });
  } catch (error) {
    console.error('saveCurrentAsProfile error:', error && error.message);
    return createExceptionResponse(error);
  }
}

function deleteMyProfile(name) {
  try {
    const resolved = __resolveOwnUserId_();
    if (resolved.error) return resolved.error;
    return __deleteProfileCore(resolved.userId, name);
  } catch (error) {
    console.error('deleteMyProfile error:', error && error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Shape the sheet-data result into the JSON-safe envelope the frontend expects.
 * Date values (raw timestamps from Sheets) are converted to ISO strings;
 * everything else is already JSON-serializable by construction.
 *
 * Why strip identity fields: showNames:false is meant to be anonymous, but the
 * frontend only hides name/email in the UI — the payload still carried them,
 * so any viewer could read peers' emails via DevTools. We now filter on the
 * server so the wire never sees the identity when identity is not permitted.
 * Admins and board owners always see everything (they can see raw sheet anyway).
 */
function buildSafePublishedDataResult(result, config, viewerContext = {}) {
  const displaySettings = (config && config.displaySettings) || DEFAULT_DISPLAY_SETTINGS;
  const includeIdentity = Boolean(
    viewerContext.isAdmin || viewerContext.isOwnBoard || displaySettings.showNames
  );
  const rows = Array.isArray(result.data) ? result.data : [];

  // Why: 可視化モード（M1/M2）で同一児童の再投稿を「揺らぎ」として追跡するために
  //      仮名化された emailHash を全モードで wire に乗せる。生メアドは showNames=false なら除去。
  //      includeIdentity=true のときも emailHash は重複して載せるが、ペイロード増は <10B/row なので無視できる。
  const safeData = rows.map(item => {
    if (typeof item !== 'object' || item === null) return item;
    const cleaned = {};
    const rawEmail = item.email;
    for (const key in item) {
      if (!includeIdentity && (key === 'email' || key === 'name')) continue;
      const v = item[key];
      cleaned[key] = v instanceof Date ? v.toISOString() : v;
    }
    // emailHash は両モードで載せる（揺らぎ追跡のため）
    if (typeof rawEmail === 'string' && rawEmail) {
      cleaned.emailHash = emailToShortHash(rawEmail);
    }
    return cleaned;
  });

  // Why: boardMode='auto' のとき、columnMapping から最適モードを推論。
  //      これでフロントは 'board'/'numberline'/'matrix' のいずれかを必ず受け取れる。
  //      列構成が変わっても教師が再設定する必要がないようにする UX 配慮。
  const columnMapping = (config && config.columnMapping) || {};
  const hasNumericX = typeof columnMapping.numericX === 'number' && columnMapping.numericX >= 0;
  const hasNumericY = typeof columnMapping.numericY === 'number' && columnMapping.numericY >= 0;
  const rawMode = displaySettings.boardMode || 'auto';
  let effectiveMode;
  if (rawMode === 'auto') {
    if (hasNumericX && hasNumericY) effectiveMode = 'matrix';
    else if (hasNumericX) effectiveMode = 'numberline';
    else effectiveMode = 'board';
  } else {
    effectiveMode = rawMode;
  }

  // 軸ラベル・象限ラベル・揺らぎ追跡フラグを wire に同梱
  const axisConfig = {
    xAxisLabels: (config && config.xAxisLabels) || null,
    yAxisLabels: (config && config.yAxisLabels) || null,
    matrixQuadrantLabels: (config && config.matrixQuadrantLabels) || null,
    allowResubmit: Boolean(config && config.allowResubmit),
    // 数値レンジは列マッピングからは取れないので、表示側でデータから推定。
    // ただし min=1,max=5 の慣例値を default として送る（Forms 線形尺度の既定）。
    defaultMin: 1,
    defaultMax: 5
  };

  // Why: multi-board の profile セレクタは「所有者または管理者のみ」が操作可能。
  //      students の wire には乗せない（混乱回避 + 内部設定の漏洩防止）。
  //      また、これを client の isTeacher 判定の根拠にも使う（profiles 配列が wire に
  //      含まれていれば、その閲覧者は teacher 権限）。
  let profileSummary = null;
  const isPrivilegedViewer = Boolean(viewerContext.isAdmin || viewerContext.isOwnBoard);
  if (isPrivilegedViewer && config && Array.isArray(config.profiles) && config.profiles.length > 0) {
    // history も teacher に同梱して、pill 上の ✓ (= 過去 active 経験あり) 表示の根拠にする。
    const teacherHistory = Array.isArray(config.profileHistory)
      ? config.profileHistory.map(h => ({ name: h.name, activatedAt: h.activatedAt || '' }))
      : [];
    profileSummary = {
      active: config.activeProfile || null,
      list: config.profiles.map(p => ({
        name: p.name,
        formTitle: p.formTitle || '',
        boardMode: (p.displaySettings && p.displaySettings.boardMode) || 'auto'
      })),
      history: teacherHistory
    };
  }

  // Option B: 生徒は過去フェーズだけは閲覧できるよう、profileHistory と current のみ
  //   サブセット情報を wire に乗せる。未来 profile (history 未登録) は漏らさない。
  //   teacher 用の full profileSummary は上記 isPrivilegedViewer 経路で送る。
  //
  // 構造: studentProfileNav = {
  //   active: '現在の profile 名',
  //   history: [{ name, activatedAt, formTitle }, ...]  // 末尾が最新
  // }
  // student client はこれだけで pill UI を組み立てる。formTitle は表示名として使う。
  let studentProfileNav = null;
  const rawHistory = (config && Array.isArray(config.profileHistory)) ? config.profileHistory : [];
  if (!isPrivilegedViewer && rawHistory.length > 0 && Array.isArray(config && config.profiles)) {
    // formTitle 補完: history は {name, activatedAt} だけ持つので profiles から表示名を引く。
    //   削除済 profile (history に名前あるが profiles から消えた) は formTitle 空のまま渡す。
    const profileByName = {};
    for (const p of config.profiles) {
      if (p && p.name) profileByName[p.name] = p;
    }
    studentProfileNav = {
      active: config.activeProfile || null,
      history: rawHistory.map(h => {
        const meta = profileByName[h.name] || {};
        return {
          name: h.name,
          activatedAt: h.activatedAt || '',
          formTitle: meta.formTitle || '',
          // 削除済 profile かどうか: students にロード不可と判定させるためのフラグ
          deleted: !profileByName[h.name]
        };
      })
    };
  }

  // viewer がティーチャー権限を持つかの最終フラグ。client 側で UI 表示判定に使う。
  // server 側で判定する方が改ざんに強い（client の window.isEditor は信用しない）。
  const viewerIsTeacher = isPrivilegedViewer;

  // Why: 生徒の「📝 回答フォーム」ボタンを polling のたびに最新値に更新するため、
  //      formUrl と formTitle を wire に載せる。profile 切替で active config の
  //      formUrl が変わったとき、5 秒以内に生徒のボタンも切替わる（教師との同期）。
  //      formUrl は config に元から保存されている情報なので追加のセキュリティリスクなし。
  const formMeta = {
    formUrl: (config && typeof config.formUrl === 'string') ? config.formUrl : '',
    formTitle: (config && typeof config.formTitle === 'string') ? config.formTitle : ''
  };

  return {
    success: true,
    data: safeData,
    header: String(result.header || result.sheetName || '回答一覧'),
    sheetName: String(result.sheetName || 'Sheet1'),
    displaySettings: { ...displaySettings, boardMode: effectiveMode },
    axisConfig,
    profiles: profileSummary,
    studentProfileNav,
    formMeta,
    viewerIsTeacher,
    // viewingPastProfile: 過去フェーズ閲覧時のみ profile 名を入れる（caller が override）。
    //   主経路 (getPublishedSheetData) では undefined のまま。
    //   getPublishedSheetDataForProfile が結果を post-process して値を入れる。
    viewingPastProfile: viewerContext.viewingPastProfile || null
  };
}

/**
 * 統合API: フロントエンド用データ取得（最適化版・クロスユーザー対応）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @param {string} targetUserId - 対象ユーザーID
 * @returns {Object} フロントエンド期待形式のデータ
 */
function getPublishedSheetData(classFilter, sortOrder, adminMode, targetUserId) {
  classFilter = classFilter || null;
  sortOrder = sortOrder || 'newest';
  adminMode = adminMode || false;
  targetUserId = targetUserId || null;

  try {
    const adminAuth = getBatchedAdminAuth({ allowNonAdmin: true });
    if (!adminAuth.success || !adminAuth.authenticated) {
      return {
        success: false,
        error: 'Authentication required',
        data: [],
        sheetName: '',
        header: '認証エラー'
      };
    }

    const { email: viewerEmail, isAdmin: isSystemAdmin } = adminAuth;

    if (targetUserId) {
      const targetUser = findPublishedBoardOwner(targetUserId, viewerEmail, {
        preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
      });
      if (!targetUser) {
        console.error('getPublishedSheetData: Target user not found', { targetUserId, viewerEmail });
        return {
          success: false,
          error: 'Target user not found',
          data: [],
          sheetName: '',
          header: 'ユーザーエラー'
        };
      }

      const targetUserConfig = getConfigOrDefault(targetUserId, targetUser);
      const isOwnBoard = targetUser.userEmail === viewerEmail;
      const isPublished = Boolean(targetUserConfig.isPublished);

      if (!isSystemAdmin && !isOwnBoard && !isPublished) {
        return {
          success: false,
          error: 'このボードは未公開です',
          data: [],
          sheetName: '',
          header: '未公開'
        };
      }

      const options = {
        classFilter: classFilter !== 'すべて' ? classFilter : undefined,
        sortBy: sortOrder || 'newest',
        includeTimestamp: true,
        adminMode: isSystemAdmin || isOwnBoard,
        requestingUser: viewerEmail,
        preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
      };

      const result = getUserSheetData(targetUser.userId, options, targetUser, targetUserConfig);

      if (!result || !result.success) {
        return {
          success: false,
          error: result?.message || 'データ取得エラー',
          data: [],
          sheetName: result?.sheetName || '',
          header: result?.header || '問題'
        };
      }

      return buildSafePublishedDataResult(result, targetUserConfig, {
        isAdmin: isSystemAdmin,
        isOwnBoard
      });
    }

    const user = findUserByEmail(viewerEmail, {
      requestingUser: viewerEmail,
      adminMode: isSystemAdmin,
      ignorePermissions: isSystemAdmin,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    });

    if (!user) {
      console.error('getPublishedSheetData: User not found (self-access)', { viewerEmail });
      return {
        success: false,
        error: 'User not found',
        data: [],
        sheetName: '',
        header: 'ユーザーエラー'
      };
    }

    const userConfig = getConfigOrDefault(user.userId, user);

    const options = {
      classFilter: classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true,
      adminMode: isSystemAdmin,
      requestingUser: viewerEmail,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    };

    const result = getUserSheetData(user.userId, options, user, userConfig);

    if (!result || !result.success) {
      return {
        success: false,
        error: result?.message || 'データ取得エラー',
        data: [],
        sheetName: result?.sheetName || '',
        header: result?.header || '問題',
      };
    }

    return buildSafePublishedDataResult(result, userConfig, {
      isAdmin: isSystemAdmin,
      isOwnBoard: true
    });
  } catch (error) {
    console.error('getPublishedSheetData: Exception caught', {
      error: error?.message || 'Unknown error',
      stack: error?.stack,
      classFilter,
      sortOrder,
      adminMode,
      targetUserId,
      timestamp: new Date().toISOString()
    });
    return {
      success: false,
      error: error?.message || 'データ取得エラー',
      data: [],
      sheetName: '',
      header: '問題'
    };
  }
}

/**
 * Option B: 生徒が過去フェーズ (profileHistory に登録済) を read-only で閲覧するための API。
 *
 * Why: getPublishedSheetData は active config だけを見る。教師が active を「振り返り」に
 *   切替えてしまうと、生徒は「導入」での自分の答えを後から見直せない。
 *   このエンドポイントは config を書き換えずに指定 profile snapshot を合成して読むだけ。
 *
 *   security model:
 *     - 公開済ボード（または own/admin）でなければアクセス不可（getPublishedSheetData と同じ）
 *     - profileName は profileHistory に含まれている必要あり（未来 profile / 削除済の名前漏れを防ぐ）
 *     - 結果には viewingPastProfile=name を載せて client が read-only UI に切替えられるようにする
 *
 * @param {string} targetUserId
 * @param {string} profileName
 * @param {string} [classFilter]
 * @param {string} [sortOrder]
 * @returns {Object} buildSafePublishedDataResult 形式 + viewingPastProfile
 */
function getPublishedSheetDataForProfile(targetUserId, profileName, classFilter, sortOrder) {
  try {
    if (!targetUserId || typeof targetUserId !== 'string') {
      return { success: false, error: 'targetUserId is required', data: [] };
    }
    if (!profileName || typeof profileName !== 'string') {
      return { success: false, error: 'profileName is required', data: [] };
    }

    const adminAuth = getBatchedAdminAuth({ allowNonAdmin: true });
    if (!adminAuth.success || !adminAuth.authenticated) {
      return { success: false, error: 'Authentication required', data: [] };
    }
    const { email: viewerEmail, isAdmin: isSystemAdmin } = adminAuth;

    const targetUser = findPublishedBoardOwner(targetUserId, viewerEmail, {
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    });
    if (!targetUser) {
      return { success: false, error: 'Target user not found', data: [] };
    }

    const targetConfig = getConfigOrDefault(targetUserId, targetUser);
    const isOwnBoard = targetUser.userEmail === viewerEmail;
    const isPublished = Boolean(targetConfig.isPublished);

    if (!isSystemAdmin && !isOwnBoard && !isPublished) {
      return { success: false, error: 'このボードは未公開です', data: [] };
    }

    // history gate: students (非 owner / 非 admin) は profileHistory にある名前のみ閲覧可。
    //   owner/admin は profiles[] にあれば全部読めるようにする（preview 用途）。
    const history = Array.isArray(targetConfig.profileHistory) ? targetConfig.profileHistory : [];
    const inHistory = history.some(h => h && h.name === profileName);
    const isPrivileged = isSystemAdmin || isOwnBoard;
    if (!isPrivileged && !inHistory) {
      return { success: false, error: 'このフェーズはまだ公開されていません', data: [] };
    }

    const profiles = Array.isArray(targetConfig.profiles) ? targetConfig.profiles : [];
    const p = profiles.find(x => x && x.name === profileName);
    if (!p) {
      // history にはあるが profiles[] から消されているケース（削除済 profile）
      return { success: false, error: '指定されたフェーズの設定が見つかりません（削除済の可能性）', data: [] };
    }

    if (!p.spreadsheetId) {
      return { success: false, error: 'このフェーズはスプレッドシート未設定です', data: [] };
    }

    // 合成 config: profile snapshot のフィールドを active 形に詰め替える。loadProfileForView と同じ
    //   semantics だが、こちらは「保存せず」一時的に getUserSheetData に渡すだけ。
    const synthesized = {
      ...targetConfig,  // isPublished / userId 等を継承
      formUrl: p.formUrl || '',
      formTitle: p.formTitle || '',
      spreadsheetId: p.spreadsheetId,
      sheetName: p.sheetName || '',
      columnMapping: p.columnMapping || {},
      displaySettings: p.displaySettings || (typeof DEFAULT_DISPLAY_SETTINGS !== 'undefined' ? DEFAULT_DISPLAY_SETTINGS : {}),
      xAxisLabels: p.xAxisLabels || null,
      yAxisLabels: p.yAxisLabels || null,
      matrixQuadrantLabels: p.matrixQuadrantLabels || null,
      allowResubmit: !!p.allowResubmit,
      // 検索用の anchor。active のままにしておくと「viewingPastProfile=null かつ active も同じ」と
      //   混乱するので、明示的に override する。
      activeProfile: targetConfig.activeProfile || null
    };

    const options = {
      classFilter: classFilter && classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true,
      adminMode: isSystemAdmin || isOwnBoard,
      requestingUser: viewerEmail,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    };

    const result = getUserSheetData(targetUserId, options, targetUser, synthesized);
    if (!result || !result.success) {
      return {
        success: false,
        error: result?.message || 'データ取得エラー',
        data: [],
        sheetName: result?.sheetName || '',
        header: result?.header || '問題'
      };
    }

    return buildSafePublishedDataResult(result, synthesized, {
      isAdmin: isSystemAdmin,
      isOwnBoard,
      viewingPastProfile: profileName
    });
  } catch (error) {
    console.error('getPublishedSheetDataForProfile error:', error && error.message);
    return {
      success: false,
      error: (error && error.message) || 'データ取得エラー',
      data: []
    };
  }
}

/**
 * GAS-Native統一設定保存API
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - 保存オプション { isDraft: boolean }
 * @returns {Object} 保存結果
 */
function saveConfig(config, options = {}) {
  const startTime = Date.now();

  try {
    const userEmail = getCurrentEmail();

    if (!userEmail) {
      return createAuthError();
    }

    const user = findUserByEmail(userEmail, { requestingUser: userEmail });

    if (!user) {
      return createUserNotFoundError();
    }

    const saveOptions = options.isDraft ?
      { isDraft: true } :
      { isMainConfig: true };

    const result = saveUserConfig(user.userId, config, saveOptions);

    return result;

  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`saveConfig: ERROR after ${duration}ms - ${error.message || 'Operation error'}`);
    return { success: false, message: error.message || 'エラーが発生しました' };
  }
}

/**
 * 時刻ベース統一新着通知システム - 全ユーザーロール対応
 * @param {string} targetUserId - 閲覧対象ユーザーID
 * @param {Object} options - オプション設定
 * @returns {Object} 時刻ベース統一通知更新結果
 */
function getNotificationUpdate(targetUserId, options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !targetUserId) {
      return { success: false, message: 'Invalid request' };
    }

    const targetUser = findPublishedBoardOwner(targetUserId, email);
    if (!targetUser) {
      return { success: false, message: 'User not found' };
    }

    const targetConfig = getConfigOrDefault(targetUser.userId, targetUser);
    const isOwnBoard = targetUser.userEmail === email;
    const isAdmin = isAdministrator(email);

    if (!isAdmin && !isOwnBoard && !targetConfig.isPublished) {
      return { success: false, message: 'Access denied' };
    }

    // Why: 'すべて' はサーバーにとっては null と同義。生のまま渡すとクラス名として
      //      フィルタされて全件 0 にマッチしなくなる（getPublishedSheetData と同じ正規化）。
    const classFilter = options.classFilter && options.classFilter !== 'すべて'
      ? options.classFilter
      : undefined;

    const userData = getUserSheetData(targetUserId, {
      includeTimestamp: true,
      classFilter,
      sortBy: options.sortOrder || 'newest',
      requestingUser: email,
      targetUserEmail: targetUser.userEmail
    }, targetUser, targetConfig);

    if (!userData.success) {
      return { success: false, message: 'Data access failed' };
    }

    const lastUpdate = new Date(options.lastUpdateTime || 0);
    const newItems = userData.data.filter(item => {
      const itemTime = new Date(item.timestamp || 0);
      return itemTime > lastUpdate;
    });

    // Why: 教師の profile 切替を生徒側 5 秒 polling で即時検知できるよう、
    //   formMeta (formUrl + formTitle) と activeProfile を載せる。生徒は前回値と
    //   比較して URL 変更を検知 → トースト + 自動データ再読込が走る。
    //   これがないと polling は「新着投稿」しか拾えず、profile 切替は気付かれない。
    return {
      success: true,
      hasNewContent: newItems.length > 0,
      newItemsCount: newItems.length,
      formMeta: {
        formUrl: (targetConfig && typeof targetConfig.formUrl === 'string') ? targetConfig.formUrl : '',
        formTitle: (targetConfig && typeof targetConfig.formTitle === 'string') ? targetConfig.formTitle : ''
      },
      activeProfile: (targetConfig && targetConfig.activeProfile) || null
    };

  } catch (error) {
    logError_('getNotificationUpdate', error);
    return { success: false, message: error.message };
  }
}

/**
 * Connect to data source - API Gateway function for DataService
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} batchOperations - バッチ操作
 * @returns {Object} 接続結果
 */
function connectDataSource(spreadsheetId, sheetName, batchOperations = null) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.warn('connectDataSource: Unauthenticated access attempt');
      return createAuthError();
    }

    try {
      setupDomainWideSharing(spreadsheetId, email);
    } catch (sharingError) {
      console.warn('connectDataSource: Domain-wide sharing setup failed (non-critical):', sharingError.message);
    }

    if (batchOperations && Array.isArray(batchOperations)) {
      return processDataSourceOperations(spreadsheetId, sheetName, batchOperations);
    }

    return getColumnAnalysis(spreadsheetId, sheetName);

  } catch (error) {
    logError_('connectDataSource', error);
    return createExceptionResponse(error);
  }
}

/**
 * バッチ処理でデータソース操作を実行
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} operations - 操作配列
 * @returns {Object} バッチ処理結果
 */
function processDataSourceOperations(spreadsheetId, sheetName, operations) {
  try {
    const results = {
      success: true,
      batchResults: {},
      message: '統合処理完了'
    };

    let columnAnalysisResult = null;

    for (const operation of operations) {
      switch (operation.type) {
        case 'validateAccess':
          if (!columnAnalysisResult) {
            columnAnalysisResult = getColumnAnalysis(spreadsheetId, sheetName);
          }
          results.batchResults.validation = {
            success: columnAnalysisResult.success,
            details: {
              connectionVerified: columnAnalysisResult.success,
              connectionError: columnAnalysisResult.success ? null : columnAnalysisResult.message
            }
          };
          break;
        case 'getFormInfo':
          results.batchResults.formInfo = getFormInfo(spreadsheetId, sheetName);
          break;
        case 'connectDataSource': {
          if (!columnAnalysisResult) {
            columnAnalysisResult = getColumnAnalysis(spreadsheetId, sheetName);
          }

          if (columnAnalysisResult.success) {
            results.mapping = columnAnalysisResult.mapping;
            results.confidence = columnAnalysisResult.confidence;
            results.headers = columnAnalysisResult.headers;
            results.data = columnAnalysisResult.data;
            results.sheet = columnAnalysisResult.sheet;
          } else {
            results.success = false;
            results.error = columnAnalysisResult.errorResponse?.message || columnAnalysisResult.message;
          }
          break;
        }
      }
    }

    return results;
  } catch (error) {
    logError_('processDataSourceOperations', error);
    return createExceptionResponse(error);
  }
}

/**
 * 列分析 - API Gateway実装（既存サービス活用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function getColumnAnalysis(spreadsheetId, sheetName) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.warn('getColumnAnalysis: Unauthenticated access attempt');
      return createAuthError();
    }

    let dataAccess;
    try {
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
      if (!dataAccess) {
        return {
          success: false,
          error: 'スプレッドシートにアクセスできません。同一ドメイン内で編集可能に設定されているか確認してください。'
        };
      }
    } catch (accessError) {
      console.warn('getColumnAnalysis: Spreadsheet access failed for user:', `${email.split('@')[0]}@***`, accessError.message);
      return {
        success: false,
        error: 'スプレッドシートにアクセス権がありません。同一ドメイン内で編集可能に設定されているか確認してください。'
      };
    }

    const sheet = dataAccess.getSheet(sheetName);

    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0
      ? (sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [])
      : [];

    let sampleData = [];
    if (lastRow > 1 && lastCol > 0) {
      const sampleSize = Math.min(10, lastRow - 1);
      try {
        sampleData = sheet.getRange(2, 1, sampleSize, lastCol).getValues();
      } catch (sampleError) {
        console.warn('getColumnAnalysis: サンプルデータ取得失敗:', sampleError.message);
        sampleData = [];
      }
    }

    const diagnostics = performIntegratedColumnDiagnostics(headers, { sampleData });

    // Why: 線形尺度（Forms「リニアスケール」）列が見つかれば、numericX/numericY として
    //      自動的に mapping に含める。教師は明示設定なしで M1/M2 モードが使えるようになる。
    //      候補が複数あるとき、信頼度上位 2 件を採用（matrix 用）。
    //      欠落・低信頼の場合は何も設定せず、auto モードは 'board' にフォールバックする。
    const mergedMapping = { ...(diagnostics.recommendedMapping || {}) };
    const numericCandidates = Array.isArray(diagnostics.numericScaleCandidates)
      ? diagnostics.numericScaleCandidates
      : [];
    const acceptedScales = numericCandidates.filter(c => c && typeof c.index === 'number' && c.confidence >= 80);
    if (acceptedScales.length >= 1 && typeof mergedMapping.numericX !== 'number') {
      mergedMapping.numericX = acceptedScales[0].index;
    }
    if (acceptedScales.length >= 2 && typeof mergedMapping.numericY !== 'number') {
      mergedMapping.numericY = acceptedScales[1].index;
    }

    try {
      const columnSetupResult = setupReactionAndHighlightColumns(spreadsheetId, sheetName, headers);
      if (columnSetupResult.columnsAdded && columnSetupResult.columnsAdded.length > 0) {
        return {
          success: true,
          sheet,
          headers: [...headers, ...columnSetupResult.columnsAdded],
          data: [],
          mapping: mergedMapping,
          confidence: diagnostics.confidence || {},
          numericScaleCandidates: numericCandidates,
          columnsAdded: columnSetupResult.columnsAdded
        };
      }
    } catch (columnError) {
      console.warn('getColumnAnalysis: Column setup failed:', columnError.message);
    }

    return {
      success: true,
      sheet,
      headers,
      data: [],
      mapping: mergedMapping,
      confidence: diagnostics.confidence || {},
      numericScaleCandidates: numericCandidates
    };
  } catch (error) {
    logError_('getColumnAnalysis', error);
    return createExceptionResponse(error);
  }
}

/**
 * リアクション列とハイライト列を事前セットアップ
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} currentHeaders - 現在のヘッダー配列
 * @returns {Object} 追加結果
 */
function setupReactionAndHighlightColumns(spreadsheetId, sheetName, currentHeaders = []) {
  try {
    // Why openSpreadsheet (not bare SpreadsheetApp.openById): DatabaseCore.openSpreadsheet
    //   は circuit breaker + service account fallback + retry を一括で提供する。直接呼びで
    //   この envelope をバイパスすると、429 storm 時に簡単にロックアウトを引き起こす。
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
    if (!dataAccess) {
      throw new Error(`Cannot open spreadsheet ${spreadsheetId}`);
    }
    const { spreadsheet } = dataAccess;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }

    const requiredColumns = [
      'UNDERSTAND',
      'LIKE',
      'CURIOUS',
      'HIGHLIGHT'
    ];

    const columnsToAdd = [];
    // Why strict equality: processReactionDirect looks up reaction columns with
    //   `header.trim() === 'UNDERSTAND'` etc. If this existence check is looser
    //   (e.g. substring), the two sides disagree and reactions break silently.
    const currentHeadersNormalized = currentHeaders.map(h => String(h || '').toUpperCase().trim());

    requiredColumns.forEach(columnName => {
      const exists = currentHeadersNormalized.includes(columnName);

      if (!exists) {
        columnsToAdd.push(columnName);
      }
    });

    const columnsAdded = [];

    if (columnsToAdd.length > 0) {
      const { lastCol } = getSheetInfo(sheet);

      try {
        // バッチ処理: 全カラムを一度に追加（パフォーマンス最適化）
        const startCol = lastCol + 1;
        const values = [columnsToAdd]; // 2D配列（1行 × n列）
        sheet.getRange(1, startCol, 1, columnsToAdd.length).setValues(values);
        columnsAdded.push(...columnsToAdd);
      } catch (colError) {
        console.error(`setupReactionAndHighlightColumns: Failed to add columns:`, colError.message);
        // フォールバック: 個別追加を試行
        columnsToAdd.forEach((columnName, index) => {
          try {
            sheet.getRange(1, lastCol + index + 1).setValue(columnName);
            columnsAdded.push(columnName);
          } catch (individualError) {
            console.error(`setupReactionAndHighlightColumns: Failed to add column '${columnName}':`, individualError.message);
          }
        });
      }
    }

    if (columnsAdded.length > 0) {
      invalidateSheetHeadersCache(spreadsheetId, sheetName);
    }

    return {
      success: true,
      columnsAdded,
      totalColumns: requiredColumns.length,
      alreadyExists: requiredColumns.length - columnsToAdd.length
    };

  } catch (error) {
    logError_('setupReactionAndHighlightColumns', error);
    return {
      success: false,
      error: error.message,
      columnsAdded: []
    };
  }
}

/**
 * Get active form info - フロントエンドエラー修正用
 * @param {string} userId - ユーザーID
 * @returns {Object} フォーム情報
 */
function getActiveFormInfo(userId) {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.error('Authentication failed');
      return {
        success: false,
        message: 'Authentication required',
        formUrl: null,
        formTitle: null
      };
    }

    let targetUserId = userId;
    if (!targetUserId) {
      const currentUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
      if (!currentUser) {
        console.error('Current user not found:', currentEmail);
        return {
          success: false,
          message: 'User not found',
          formUrl: null,
          formTitle: null
        };
      }
      targetUserId = currentUser.userId;
    }

    // 生徒（non-editor, non-admin）でも公開ボードの formUrl を取得できるよう
    // viewer-path helper を使う。findUserById を直接使うと allowPublishedRead:true を
    // 渡し忘れて silent failure を起こしやすいので、この経路では helper 経由で統一。
    const targetUser = findPublishedBoardOwner(targetUserId, currentEmail);
    const config = getConfigOrDefault(targetUserId, targetUser);

    const hasFormUrl = !!(config.formUrl && config.formUrl.trim());
    const isValidUrl = hasFormUrl && isValidFormUrl(config.formUrl);

    return {
      success: hasFormUrl,
      shouldShow: hasFormUrl,
      formUrl: hasFormUrl ? config.formUrl : null,
      formTitle: config.formTitle || 'フォーム',
      isValidUrl,
      message: hasFormUrl ?
        (isValidUrl ? 'Valid form found' : 'Form URL found but validation failed') :
        'No form URL configured'
    };
  } catch (error) {
    console.error('getActiveFormInfo ERROR:', error.message);
    return {
      success: false,
      shouldShow: false,
      message: error.message,
      formUrl: null,
      formTitle: null
    };
  }
}

/**
 * Check if URL is valid Google Forms URL
 * @param {string} url - URL to validate
 * @returns {boolean} True if valid Google Forms URL
 */
function isValidFormUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  try {
    const urlObj = new URL(url.trim());

    if (urlObj.protocol !== 'https:') {
      return false;
    }

    const validHosts = ['docs.google.com', 'forms.gle'];
    const isValidHost = validHosts.includes(urlObj.hostname);

    if (urlObj.hostname === 'docs.google.com') {
      return urlObj.pathname.includes('/forms/') || urlObj.pathname.includes('/viewform');
    }

    return isValidHost;
  } catch {
    return false;
  }
}

/**
 * Forms REST API を使用してメール収集設定を「確認済み（VERIFIED）」に設定
 * Apps Script の setCollectEmail(true) では「回答者入力」モードになるため、REST APIを使用
 * 参考: https://developers.google.com/forms/api/reference/rest/v1/forms/batchUpdate
 * @param {string} formId - フォームID
 * @private
 */
function setFormEmailCollectionVerified_(formId) {
  try {
    const url = `https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`;
    const payload = {
      requests: [{
        updateSettings: {
          settings: {
            emailCollectionType: 'VERIFIED'
          },
          updateMask: 'emailCollectionType'
        }
      }]
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.warn('Forms API emailCollectionType設定に失敗:', responseCode, response.getContentText());
      // フォールバック: Apps Script の setCollectEmail を使用
      const form = FormApp.openById(formId);
      form.setCollectEmail(true);
    }
  } catch (error) {
    console.warn('Forms REST API呼び出しエラー（フォールバック実行）:', error.message);
    // フォールバック: Apps Script の setCollectEmail を使用
    try {
      const form = FormApp.openById(formId);
      form.setCollectEmail(true);
    } catch (fallbackError) {
      console.error('フォールバックも失敗:', fallbackError.message);
    }
  }
}

