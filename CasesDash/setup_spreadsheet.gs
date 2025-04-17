/**
 * setup_spreadsheet.gs - スプレッドシート初期セットアップモジュール
 * このモジュールは、CasesDash用のスプレッドシートの初期構造を作成・設定します。
 * シートの作成、列ヘッダーの設定、データ検証ルールの適用などを行います。
 *
 * @requires settings.gs
 */

/** 対象スプレッドシートIDを取得 */
function getSpreadsheetId() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty("SPREADSHEET_ID");
}

/**
 * スプレッドシートの初期セットアップを実行します。
 * 必要なシートを作成し、列ヘッダーを設定し、データ検証ルールを適用します。
 *
 * @returns {{success: boolean, message: string}} 実行結果
 */
function setupSpreadsheet() {
  const spreadsheetId = getSpreadsheetId();
  if (!spreadsheetId) {
    return {
      success: false,
      message: "スプレッドシートIDが設定されていません。",
    };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);

    // 必要なシート名リスト
    const requiredSheets = [
      "3PO Chat",
      "OT Email",
      "3PO Email",
      "OT Chat",
      "3PO Chat",
      "OT Phone",
      "3PO Phone",
      "index",
    ];

    // シート作成または取得
    requiredSheets.forEach((sheetName) => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
      // 既存のシートはクリア（必要に応じて）
      sheet.clear();
    });

    // 3PO Chat シートのヘッダー設定例
    const chatSheet = ss.getSheetByName("3PO Chat");
    if (chatSheet) {
      const headers = [
        "Date",
        "Case",
        "Case ID",
        "Date",
        "Time",
        "Segment",
        "Category",
        "Triage",
        "Either",
        "Initiated",
        "3.0",
        "1st Assignee",
        "TRT Timer",
        "Aging Timer",
        "Destination",
        "Reason",
        "MCC",
        "to Child",
        "Final Assignee",
        "Segment",
        "Case Status",
        "AM Transfer",
        "non NCC",
        "Bug",
        "Info",
        "Date",
        "Time",
        "Reason",
        "Date",
        "Time",
        "Product",
        "Assign",
        "TRT",
        "Aging",
        "Close",
        "Other",
      ];
      chatSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // データ検証ルールやフォーミュラの設定はここに追加可能
    }

    return {
      success: true,
      message: "スプレッドシートのセットアップが完了しました。",
    };
  } catch (error) {
    return {
      success: false,
      message: `セットアップ中にエラーが発生しました: ${error.message}`,
    };
  }
}
