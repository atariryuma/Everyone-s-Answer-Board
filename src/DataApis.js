/**
 * @fileoverview DataApis - データ操作専用API
 *
 * ✅ V8ランタイム対応（2022年6月アップデート準拠）
 * - 関数定義は順序に関係なく呼び出し可能
 * - グローバルスコープでのコード実行を完全排除
 *
 * 依存関係（呼び出す関数）:
 * - getCurrentEmail() - main.jsで定義
 * - findUserById() - DatabaseCore.jsで定義
 * - findUserByEmail() - DatabaseCore.jsで定義
 * - getUserConfig() - ConfigService.jsで定義
 * - saveUserConfig() - ConfigService.jsで定義
 * - openSpreadsheet() - DatabaseCore.jsで定義
 * - getSheetInfo() - SystemController.jsで定義
 * - getUserSheetData() - DataService.jsで定義
 * - getBatchedAdminAuth() - main.jsで定義
 * - getQuestionText() - DataService.jsで定義
 * - getFormInfo() - DataService.jsで定義
 * - performIntegratedColumnDiagnostics() - ColumnMappingService.jsで定義
 * - setupDomainWideSharing() - helpers.jsで定義
 * - validateAccess() - DataService.jsで定義
 * - executeWithRetry() - main.jsで定義
 * - createAuthError() - helpers.jsで定義
 * - createUserNotFoundError() - helpers.jsで定義
 * - createErrorResponse() - helpers.jsで定義
 * - createExceptionResponse() - helpers.jsで定義
 *
 * 移動元: main.js
 * 移動日: 2025-12-13
 */

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, findPublishedBoardOwner, getUserConfig, getConfigOrDefault, DEFAULT_DISPLAY_SETTINGS, saveUserConfig, openSpreadsheet, getSheetInfo, getUserSheetData, getBatchedAdminAuth, getQuestionText, getFormInfo, invalidateSheetHeadersCache, performIntegratedColumnDiagnostics, setupDomainWideSharing, validateAccess, executeWithRetry, createAuthError, createUserNotFoundError, createErrorResponse, createExceptionResponse, emailToShortHash */
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
    console.error('getForms error:', error.message);
    return { success: false, error: error.message, forms: [] };
  }
}

// Why: 線形尺度の範囲は detectNumericScaleColumns の判定（0-10 の整数, レンジ 1-9）と
//      整合する必要がある。テンプレートと検出ロジックが乖離すると、テンプレートで作った
//      M1/M2 フォームの数値列が auto モードで検出されない事故になる。
const TEMPLATE_SCALE_MIN = 1;
const TEMPLATE_SCALE_MAX = 5;

const TEMPLATE_LABELS = Object.freeze({
  board: '掲示板',
  numberline: '数直線',
  matrix: 'マトリクス'
});

function addScaleItemTo_(form, { title, helpText, lowLabel, highLabel }) {
  form.addScaleItem()
    .setTitle(title)
    .setRequired(true)
    .setHelpText(helpText)
    .setBounds(TEMPLATE_SCALE_MIN, TEMPLATE_SCALE_MAX)
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
function createTemplateForm(templateType) {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      return { success: false, error: '認証が必要です' };
    }

    const type = TEMPLATE_LABELS[templateType] ? templateType : 'board';

    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d HH:mm');
    const formTitle = `回答ボード [${TEMPLATE_LABELS[type]}] ${dateStr}`;
    const form = FormApp.create(formTitle);

    // Why VERIFIED via REST: Apps Script の setCollectEmail(true) では「回答者入力」モード
    // になりなりすまし可能。Forms REST API で emailCollectionType=VERIFIED を強制する。
    // https://developers.google.com/forms/api/reference/rest/v1/forms/batchUpdate
    setFormEmailCollectionVerified_(form.getId());

    // Why 1回答制限: M1/M2 で allowResubmit と組み合わせて「揺らぎ」追跡するため、
    //   教師が後から無効化することを想定。テンプレート既定は安全側の 1 回答。
    form.setLimitOneResponsePerUser(true);

    form.addListItem()
      .setTitle('クラス')
      .setRequired(true)
      .setHelpText('所属クラスを選択してください')
      .setChoiceValues(['クラス1', 'クラス2', 'クラス3', 'クラス4']);
    form.addTextItem()
      .setTitle('名前')
      .setRequired(true)
      .setHelpText('表示名を入力してください');

    if (type === 'numberline') {
      addScaleItemTo_(form, {
        title: '立場',
        helpText: 'あなたの考えに最も近い位置を選んでください',
        lowLabel: 'そう思わない',
        highLabel: 'とてもそう思う'
      });
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('その位置を選んだ理由を記入してください');
    } else if (type === 'matrix') {
      addScaleItemTo_(form, {
        title: 'X軸（例: 効率）',
        helpText: '横軸の立場を選んでください',
        lowLabel: '低い',
        highLabel: '高い'
      });
      addScaleItemTo_(form, {
        title: 'Y軸（例: 倫理）',
        helpText: '縦軸の立場を選んでください',
        lowLabel: '低い',
        highLabel: '高い'
      });
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('その位置を選んだ理由を記入してください');
    } else {
      form.addMultipleChoiceItem()
        .setTitle('回答')
        .setRequired(true)
        .setHelpText('回答を選択してください')
        .setChoiceValues(['賛成', '反対', 'どちらでもない']);
      form.addParagraphTextItem()
        .setTitle('理由')
        .setRequired(true)
        .setHelpText('選択した理由を詳しく記入してください');
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

    return {
      success: true,
      formUrl: form.getPublishedUrl(),
      editUrl: form.getEditUrl(),
      formId: form.getId(),
      formTitle: formTitle,
      templateType: type,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      spreadsheetName: ss.getName(),
      sheetName: 'フォームの回答 1',
      folderUrl: folderUrl,
      wasCreated: true,
      message: `テンプレート（${TEMPLATE_LABELS[type]}）を作成しました`
    };

  } catch (error) {
    console.error('createTemplateForm error:', error.message);
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
      form = formIdOrUrl.startsWith('http')
        ? FormApp.openByUrl(formIdOrUrl)
        : FormApp.openById(formIdOrUrl);
    } catch (e) {
      return { success: false, error: 'Form を開けません: ' + e.message };
    }

    // 既存項目を削除（後ろから消すと index ずれない）
    const existing = form.getItems();
    for (let i = existing.length - 1; i >= 0; i--) {
      try { form.deleteItem(existing[i]); } catch (_) { /* skip non-deletable */ }
    }

    // タイトル・説明
    if (typeof schema.title === 'string' && schema.title.trim()) {
      form.setTitle(schema.title.trim().substring(0, 200));
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
    try {
      const destId = newSpreadsheetId || form.getDestinationId();
      if (destId && typeof invalidateSheetHeadersCache === 'function') {
        ['フォームの回答 1', 'Form Responses 1', 'Sheet1'].forEach(sn => {
          try { invalidateSheetHeadersCache(destId, sn); } catch (_) {}
        });
      }
    } catch (_) {}

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
    console.error('customizeForm error:', error.message);
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
    let spreadsheetId = form.getDestinationId();
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
    console.error('processFormUrlInput error:', error.message);
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
    console.error('validateHeaderIntegrity error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get board info - simplified name
 * @returns {Object} ボード情報
 */
function getBoardInfo() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.error('Authentication failed');
      return createAuthError();
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      console.error('User not found:', email);
      return { success: false, message: 'User not found' };
    }

    const config = getConfigOrDefault(user.userId);

    const isPublished = Boolean(config.isPublished);
    const baseUrl = getWebAppUrl();

    return {
      success: true,
      isPublished,
      questionText: getQuestionText(config, { targetUserEmail: user.userEmail }),
      urls: {
        view: `${baseUrl}?mode=view&userId=${user.userId}`,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || user.lastModified
    };
  } catch (error) {
    console.error('getBoardInfo ERROR:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get sheet list for spreadsheet - simplified name
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getSheetList(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      return createErrorResponse('Spreadsheet ID required');
    }

    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.warn('getSheetList: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    let dataAccess = null;

    try {
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
    } catch (accessError) {
      console.warn('getSheetList: Spreadsheet access failed:', accessError.message);
      return {
        success: false,
        error: 'スプレッドシートにアクセスできません。同一ドメイン内で編集可能に設定されているか確認してください。'
      };
    }

    if (!dataAccess || !dataAccess.spreadsheet) {
      console.warn('getSheetList: Failed to access spreadsheet:', `${spreadsheetId.substring(0, 8)}***`);
      return {
        success: false,
        error: 'スプレッドシートを開くことができませんでした。同一ドメイン内で編集可能に設定されているか確認してください。'
      };
    }

    const { spreadsheet } = dataAccess;
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => {
      const { lastRow, lastCol } = getSheetInfo(sheet);
      return {
        name: sheet.getName(),
        id: sheet.getSheetId(),
        rowCount: lastRow,
        columnCount: lastCol
      };
    });

    return {
      success: true,
      sheets: sheetList,
      accessMethod: 'normal_permissions'
    };
  } catch (error) {
    console.error('getSheetList error:', error.message);
    return {
      success: false,
      error: `シート一覧取得エラー: ${error.message}`,
      details: error.stack
    };
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

    // profile を active config に適用
    const merged = { ...cur };
    merged.formUrl = p.formUrl || '';
    merged.formTitle = p.formTitle || '';
    merged.spreadsheetId = p.spreadsheetId || '';
    merged.sheetName = p.sheetName || '';
    if (p.columnMapping) merged.columnMapping = p.columnMapping;
    if (p.displaySettings) merged.displaySettings = p.displaySettings;
    merged.xAxisLabels = p.xAxisLabels || null;
    merged.yAxisLabels = p.yAxisLabels || null;
    merged.matrixQuadrantLabels = p.matrixQuadrantLabels || null;
    merged.allowResubmit = !!p.allowResubmit;
    merged.activeProfile = p.name;
    merged.userId = user.userId;

    const saveRes = saveUserConfig(user.userId, merged, { isMainConfig: true });
    if (!saveRes.success) return saveRes;
    return createSuccessResponse('Profile loaded', { activeProfile: p.name });
  } catch (error) {
    console.error('loadProfileForView error:', error && error.message);
    return createExceptionResponse(error);
  }
}

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

  // Why: multi-board の profile セレクタを view 画面で描画するために、profile name 一覧と
  //      activeProfile を wire に乗せる。所有者/管理者のみ profile 一覧が見えるべきだが、
  //      ここでは contextual に「name と表示用メタ」だけ送る（spreadsheetId 等の機密情報は除外）。
  let profileSummary = null;
  if (config && Array.isArray(config.profiles) && config.profiles.length > 0) {
    profileSummary = {
      active: config.activeProfile || null,
      list: config.profiles.map(p => ({
        name: p.name,
        formTitle: p.formTitle || '',
        boardMode: (p.displaySettings && p.displaySettings.boardMode) || 'auto'
      }))
    };
  }

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
    formMeta
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
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    const user = findUserByEmail(userEmail, { requestingUser: userEmail });

    if (!user) {
      return { success: false, message: 'ユーザーが見つかりません' };
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

    return {
      success: true,
      hasNewContent: newItems.length > 0,
      newItemsCount: newItems.length
    };

  } catch (error) {
    console.error('getNotificationUpdate error:', error.message);
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
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
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
    console.error('connectDataSource error:', error.message);
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
          results.batchResults.formInfo = getFormInfoInternal(spreadsheetId, sheetName);
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
    console.error('processDataSourceOperations error:', error.message);
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
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
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
    console.error('getColumnAnalysis error:', error.message);
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
    console.error('setupReactionAndHighlightColumns error:', error.message);
    return {
      success: false,
      error: error.message,
      columnsAdded: []
    };
  }
}

/**
 * フォーム情報取得（バッチ処理用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfoInternal(spreadsheetId, sheetName) {
  try {
    if (typeof getFormInfo === 'function') {
      return getFormInfo(spreadsheetId, sheetName);
    } else {
      return {
        success: false,
        message: 'getFormInfo関数が初期化されていません'
      };
    }
  } catch (error) {
    const errorMessage = error && error.message ? error.message : '予期しないエラーが発生しました';
    console.error('getFormInfoInternal error:', errorMessage);
    return {
      success: false,
      message: errorMessage
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
 * スプレッドシートURL解析 - GAS-Native Implementation
 * @param {string} fullUrl - 完全なスプレッドシートURL（gid含む）
 * @returns {Object} 解析結果 {spreadsheetId, gid}
 */
function extractSpreadsheetInfo(fullUrl) {
  try {
    if (!fullUrl || typeof fullUrl !== 'string') {
      return {
        success: false,
        message: 'Invalid URL provided'
      };
    }

    const spreadsheetIdMatch = fullUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const gidMatch = fullUrl.match(/[#&]gid=(\d+)/);

    if (!spreadsheetIdMatch) {
      return {
        success: false,
        message: 'Invalid Google Sheets URL format'
      };
    }

    return {
      success: true,
      spreadsheetId: spreadsheetIdMatch[1],
      gid: gidMatch ? gidMatch[1] : '0'
    };
  } catch (error) {
    console.error('extractSpreadsheetInfo error:', error.message);
    return {
      success: false,
      message: `URL parsing error: ${error.message}`
    };
  }
}

/**
 * GIDからシート名取得 - Zero-Dependency + Batch Operations
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} gid - シートGID
 * @returns {string} シート名
 */
function getSheetNameFromGid(spreadsheetId, gid) {
  try {
    // Why openSpreadsheet (not bare SpreadsheetApp.openById + executeWithRetry):
    //   openSpreadsheet がすでに retry / SA fallback / circuit-breaker を包含するので、
    //   ここで二重に executeWithRetry を走らせるのは無駄かつ backoff が重複する。
    // null fallback は 'Sheet1' (旧挙動互換): GAS の新規スプレッドシート既定名。
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
    if (!dataAccess) {
      return 'Sheet1';
    }
    const { spreadsheet } = dataAccess;
    const sheets = spreadsheet.getSheets();

    const sheetInfos = sheets.map(sheet => ({
      name: sheet.getName(),
      gid: sheet.getSheetId().toString()
    }));

    const targetSheet = sheetInfos.find(info => info.gid === gid);
    const resultName = targetSheet ? targetSheet.name : sheetInfos[0]?.name || 'Sheet1';

    return resultName;

  } catch (error) {
    console.error('getSheetNameFromGid error:', error.message);
    return 'Sheet1';
  }
}

/**
 * 完全URL統合検証 - 既存API活用 + Performance最適化
 * @param {string} fullUrl - 完全なスプレッドシートURL
 * @returns {Object} 統合検証結果
 */
function validateCompleteSpreadsheetUrl(fullUrl) {
  const started = Date.now();
  try {
    const parseResult = extractSpreadsheetInfo(fullUrl);
    if (!parseResult.success) {
      return parseResult;
    }

    const { spreadsheetId, gid } = parseResult;
    const sheetName = getSheetNameFromGid(spreadsheetId, gid);
    const accessResult = validateAccess(spreadsheetId, true);

    const formResult = typeof getFormInfo === 'function' ?
      getFormInfo(spreadsheetId, sheetName) :
      { success: false, message: 'getFormInfo関数が初期化されていません' };

    const result = {
      success: true,
      spreadsheetId,
      gid,
      sheetName,
      hasAccess: accessResult.success,
      accessInfo: {
        spreadsheetName: accessResult.spreadsheetName,
        sheets: accessResult.sheets || []
      },
      formInfo: formResult,
      readyToConnect: accessResult.success && sheetName,
      executionTime: `${Date.now() - started}ms`
    };

    return result;

  } catch (error) {
    const errorResult = {
      success: false,
      message: `Complete validation error: ${error.message}`,
      error: error.message,
      executionTime: `${Date.now() - started}ms`
    };

    console.error('validateCompleteSpreadsheetUrl ERROR:', errorResult);
    return errorResult;
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
