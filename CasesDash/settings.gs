/**
 * settings.gs - 設定管理モジュール
 * スプレッドシートIDの保存・取得、接続テスト、スプレッドシートアクセスを担当します。
 */

// 定数 (Code.gs から移動)
const SHEET_NAME = "3PO Chat"; // 対象シート名 - 他のモジュールからも参照される可能性あり
const HEADER_ROW = 2; // ヘッダー行番号 - 他のモジュールからも参照される可能性あり

// 管理者メールアドレスリストをコード内に固定化（設定タブUIからは削除）
const ADMIN_EMAILS = [
  "daito@google.com",
  "mhiratsuka@google.com",
  "rtakahata@google.com",
];

/**
 * スプレッドシートへのアクセスを取得するヘルパー関数。
 * エラーハンドリングとシート存在確認を含む。
 * @return {{ss: Spreadsheet, sheet: Sheet, success: boolean, message: string}} スプレッドシートオブジェクト、シートオブジェクト、成功フラグ、メッセージ
 */
function getSpreadsheetAccess() {
  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetId = userProperties.getProperty("SPREADSHEET_ID");

  if (!spreadsheetId) {
    return {
      ss: null,
      sheet: null,
      success: false,
      message:
        "スプレッドシートIDが設定されていません。「設定」タブで設定してください。",
    };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return {
        ss: ss,
        sheet: null,
        success: false,
        message: `シート "${SHEET_NAME}" が見つかりません。スプレッドシートを確認してください。`,
      };
    }
    return { ss: ss, sheet: sheet, success: true, message: "接続成功" };
  } catch (error) {
    Logger.log(`スプレッドシートアクセスエラー: ${error}`);
    let errorMessage = "スプレッドシートへのアクセスに失敗しました。";
    if (error.message.includes("You do not have permission")) {
      errorMessage += "アクセス権限がない可能性があります。";
    } else if (error.message.includes("Not Found")) {
      errorMessage += "指定されたIDのスプレッドシートが見つかりません。";
    } else {
      errorMessage += `詳細: ${error.message}`;
    }
    return { ss: null, sheet: null, success: false, message: errorMessage };
  }
}

/**
 * 管理者メールアドレスリストをコード内定数から返します。
 * @return {Array<string>} 管理者メールアドレスの配列。
 */
function getAdminEmails_() {
  return ADMIN_EMAILS;
}

/**
 * システム設定 (スプレッドシートID) をユーザープロパティに保存する関数。
 * フロントエンドから呼び出される想定。
 * @param {Object} settings - 保存する設定オブジェクト (例: { spreadsheetId: '...' })
 * @return {Object} 結果オブジェクト { success: boolean, message: string }
 */
function saveSettings(settings) {
  try {
    if (!settings || !settings.spreadsheetId) {
      return {
        success: false,
        message: "スプレッドシートIDが提供されていません。",
      };
    }
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty("SPREADSHEET_ID", settings.spreadsheetId);

    // 管理者メールアドレスの設定は無視（コード内定数を使用）

    // キャッシュをクリア
    try {
      const cache = CacheService.getUserCache();
      // cache.remove("formDefinitions");
      // sheetHeaders のキャッシュクリアは dataAccess モジュールが担当する
    } catch (cacheError) {
      Logger.log(`キャッシュクリア中のエラー (無視します): ${cacheError}`);
    }

    Logger.log(`設定保存: スプレッドシートID = ${settings.spreadsheetId}`);
    return { success: true, message: "設定が正常に保存されました。" };
  } catch (error) {
    Logger.log(`設定保存エラー: ${error}`);
    return {
      success: false,
      message: `設定の保存中にエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * 現在保存されているシステム設定 (スプレッドシートID) を取得する関数。
 * フロントエンドから呼び出される想定。
 * @return {Object} 結果オブジェクト { success: boolean, settings?: { spreadsheetId: string }, message?: string }
 */
function getSettings() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const spreadsheetId = userProperties.getProperty("SPREADSHEET_ID") || "";
    Logger.log(
      `設定取得: スプレッドシートID = ${spreadsheetId}, 管理者メールアドレスはコード内定数を使用`
    );
    return {
      success: true,
      settings: { spreadsheetId: spreadsheetId, adminEmails: ADMIN_EMAILS },
    };
  } catch (error) {
    Logger.log(`設定取得エラー: ${error}`);
    return {
      success: false,
      message: `設定の取得中にエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * スプレッドシートIDの接続テストを行います。
 * @param {string} spreadsheetId - テストするスプレッドシートID。
 * @returns {{success: boolean, message: string}} テスト結果。
 */
function testConnection(spreadsheetId) {
  if (!spreadsheetId) {
    return {
      success: false,
      message: "スプレッドシートIDが指定されていません。",
    };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        message: `シート "${SHEET_NAME}" が見つかりません。`,
      };
    }
    return { success: true, message: "接続に成功しました。" };
  } catch (e) {
    let msg = "接続に失敗しました。";
    if (e.message.includes("You do not have permission")) {
      msg += " アクセス権限がない可能性があります。";
    } else if (e.message.includes("Not Found")) {
      msg += " 指定されたIDのスプレッドシートが見つかりません。";
    } else {
      msg += ` 詳細: ${e.message}`;
    }
    return { success: false, message: msg };
  }
}
