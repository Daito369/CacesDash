/**
 * Code.gs - メインスクリプトファイル
 * 3POチャットケース管理システム (改訂版)
 * 
 * @fileoverview Google Apps Script バックエンドロジック。
 * スプレッドシートとの連携、データ処理、ユーザー認証などを担当します。
 */

// 定数
const HEADER_ROW = 2; // ヘッダー行番号
const CACHE_EXPIRATION = 300; // キャッシュ有効期間（秒）
const DEFAULT_SHEET_NAME = '3PO Chat'; // デフォルトのシート名

// フォーム定義 (フロントエンドとバックエンドで共有するデータ定義)
const FORM_DEFINITIONS = {
  segment: ['Gold', 'Platinum', 'Titanium', 'Silver', 'Bronze - Low', 'Bronze - High'],
  productCategory: ['Search', 'Display', 'Video', 'Commerce', 'Apps', 'M&A', 'Policy', 'Billing', 'Other'],
  issueCategory: ['CBT invo-invo', 'CBT invo-auto', 'CBT (self to self)', 'LC creation', 'PP link', 'PP update', 'IDT/ Bmod', 
                'LCS billing policy', 'self serve issue', 'Unidentified Charge', 'CBT Flow', 'GQ', 'OOS', 'Bulk CBT', 
                'CBT ext request', 'MMS billing policy', 'Promotion code', 'Refund', 'Review', 'TM form', 'Trademarks issue', 
                'Under Review', 'Certificate', 'Suspend', 'AIV', 'Complaint'],
  caseStatus: ['Assigned', 'Solution Offered', 'Finished'],
  amTransfer: ['Request to AM contact', 'Optimize request', 'β product inquiry', 'Trouble shooting scope but we don\'t have access to the resource', 
              'Tag team request (LCS customer)', 'Data analysis', 'Allowlist request', 'Other'],
  nonNCC: ['Duplicate', 'Discard', 'Transfer to Platinum', 'Transfer to S/B', 'Transfer to TDCX', 'Transfer to 3PO', 
          'Transfer to OT', 'Transfer to EN Team', 'Transfer to GMB Team', 'Transfer to Other Team (not AM)'],
  reopenReason: ['Partial Answer', 'Recognition issue', 'Related question', 'Prefer Phone', 'NI reopen', 'Recurrent question']
};

/**
 * WebアプリとしてGETリクエストを受け取ったときに実行される関数。
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} HTMLサービスからの出力
 */
function doGet(e) {
  // Live Mode のパラメータをチェック
  if (e && e.parameter && e.parameter.view === 'live') {
    // Live Mode 用のテンプレートを返す
    return HtmlService.createTemplateFromFile('live')
      .evaluate()
      .setTitle('新規ケース登録 (Live Mode)')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  } else {
    // 通常のメイン画面を返す
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('3PO Chat ケース管理システム')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }
}

/**
 * HTMLテンプレート内で他のHTMLファイルをインクルードするための関数。
 * @param {string} filename - インクルードするファイル名 (拡張子なし)
 * @return {string} ファイルの内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * フォーム定義を取得する関数
 * フロントエンドから参照されるため必要
 * @return {Object} フォーム定義
 */
function getFormDefinitions() {
  return FORM_DEFINITIONS;
}

/**
 * Live Mode用にWebアプリのURLを返す関数
 * @return {Object} サービスURL情報
 */
function getService() {
  try {
    const url = ScriptApp.getService().getUrl();
    return { success: true, url: url };
  } catch (error) {
    return { success: false, message: `サービスURL取得エラー: ${error.message}` };
  }
}

/**
 * スプレッドシートへのアクセスを取得するヘルパー関数。
:start_line:85
:end_line:119
-------
 * エラーハンドリングとシート存在確認を含む。
 * @param {string} sheetName - アクセスするシート名。省略時はデフォルトシート名を使用。
 * @return {{ss: Spreadsheet, sheet: Sheet, success: boolean, message: string}} スプレッドシートオブジェクト、シートオブジェクト、成功フラグ、メッセージ
 */
function getSpreadsheetAccess(sheetName = DEFAULT_SHEET_NAME) {
  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetId = userProperties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    return { ss: null, sheet: null, success: false, message: 'スプレッドシートIDが設定されていません。「設定」タブで設定してください。' };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { ss: ss, sheet: null, success: false, message: `シート "${sheetName}" が見つかりません。スプレッドシートを確認してください。` };
    }
    return { ss: ss, sheet: sheet, success: true, message: '接続成功' };

  } catch (error) {
    Logger.log(`スプレッドシートアクセスエラー: ${error}`);
    let errorMessage = 'スプレッドシートへのアクセスに失敗しました。';
    if (error.message.includes("You do not have permission")) {
      errorMessage += 'アクセス権限がない可能性があります。';
    } else if (error.message.includes("Not Found")) {
      errorMessage += '指定されたIDのスプレッドシートが見つかりません。';
    } else {
      errorMessage += `詳細: ${error.message}`;
    }
    return { ss: null, sheet: null, success: false, message: errorMessage };
  }
}

/**
 * スプレッドシート内のすべてのシート名を取得する関数。
 * @return {Object} 結果オブジェクト { success: boolean, sheetNames?: Array<string>, message?: string }
 */
function getAllSheetNames() {
  const access = getSpreadsheetAccess(); // デフォルトシートでスプレッドシートオブジェクトを取得
  if (!access.success) {
    // スプレッドシートIDが未設定の場合でも、接続テストを試みる
    const settings = getSettings();
    if (settings.success && settings.settings.spreadsheetId) {
      try {
        const ss = SpreadsheetApp.openById(settings.settings.spreadsheetId);
        const sheets = ss.getSheets();
        const sheetNames = sheets.map(sheet => sheet.getName());
        return { success: true, sheetNames: sheetNames };
      } catch (error) {
        return { success: false, message: `スプレッドシートへのアクセスに失敗しました: ${error.message}` };
      }
    } else {
      return { success: false, message: access.message }; // ID未設定エラー
    }
  }
  const ss = access.ss;

  try {
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    Logger.log(`全シート名取得: ${sheetNames}`);
    return { success: true, sheetNames: sheetNames };
  } catch (error) {
    Logger.log(`全シート名取得エラー: ${error}`);
    return { success: false, message: `シート名の取得中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * システム設定 (スプレッドシートID) をユーザープロパティに保存する関数。
 * @param {Object} settings - 保存する設定オブジェクト (例: { spreadsheetId: '...' })
 * @return {Object} 結果オブジェクト { success: boolean, message: string }
 */
function saveSettings(settings) {
  try {
    if (!settings || !settings.spreadsheetId) {
      return { success: false, message: 'スプレッドシートIDが提供されていません。' };
    }
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('SPREADSHEET_ID', settings.spreadsheetId);
    
    // キャッシュをクリア
    try {
      const cache = CacheService.getUserCache();
      cache.remove('sheetHeaders');
      cache.remove('formDefinitions');
    } catch (cacheError) {
      Logger.log(`キャッシュクリア中のエラー (無視します): ${cacheError}`);
    }
    
    Logger.log(`設定保存: スプレッドシートID = ${settings.spreadsheetId}`);
    return { success: true, message: '設定が正常に保存されました。' };
  } catch (error) {
    Logger.log(`設定保存エラー: ${error}`);
    return { success: false, message: `設定の保存中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * 現在保存されているシステム設定 (スプレッドシートID) を取得する関数。
 * @return {Object} 結果オブジェクト { success: boolean, settings?: { spreadsheetId: string }, message?: string }
 */
function getSettings() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const spreadsheetId = userProperties.getProperty('SPREADSHEET_ID') || '';
    Logger.log(`設定取得: スプレッドシートID = ${spreadsheetId}`);
    return { success: true, settings: { spreadsheetId: spreadsheetId } };
  } catch (error) {
    Logger.log(`設定取得エラー: ${error}`);
    return { success: false, message: `設定の取得中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * 指定されたスプレッドシートIDへの接続をテストする関数。
 * シートの存在とヘッダー行の読み取りを試みる。
 * @param {string} spreadsheetId - テストするスプレッドシートID
 * @return {Object} 結果オブジェクト { success: boolean, message: string, headers?: Array<string> }
 */
function testConnection(spreadsheetId) {
  if (!spreadsheetId) {
    return { success: false, message: 'スプレッドシートIDが入力されていません。' };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return { success: false, message: `接続成功。しかし、シート "${SHEET_NAME}" が見つかりません。シート名を確認してください。` };
    }

    // ヘッダー行を取得して確認
    const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log(`接続テスト成功: ${spreadsheetId}, ヘッダー: ${headers.slice(0, 5)}...`);
    return { success: true, message: `スプレッドシート "${ss.getName()}" のシート "${SHEET_NAME}" に正常に接続できました。`, headers: headers };

  } catch (error) {
    Logger.log(`接続テストエラー: ${error}`);
    let errorMessage = 'スプレッドシートへの接続に失敗しました。';
    if (error.message.includes("You do not have permission")) {
      errorMessage += 'アクセス権限がない可能性があります。';
    } else if (error.message.includes("Not Found")) {
      errorMessage += '指定されたIDのスプレッドシートが見つかりません。';
    } else {
      errorMessage += `詳細: ${error.message}`;
    }
    return { success: false, message: errorMessage };
  }
}

/**
 * 現在のスクリプト実行ユーザーの情報を取得する関数。
 * @return {Object} 結果オブジェクト { success: boolean, email?: string, ldap?: string, message?: string }
 */
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    // Corp環境を想定し、@google.com より前をLDAPとする
    const ldap = email.split('@')[0]; 
    if (!email || !ldap) {
      throw new Error("有効なユーザー情報を取得できませんでした。");
    }
    Logger.log(`ユーザー情報取得: email=${email}, ldap=${ldap}`);
    return { success: true, email: email, ldap: ldap };
  } catch (error) {
    Logger.log(`ユーザー情報取得エラー: ${error}`);
    // より具体的なエラーメッセージを試みる
    let userMessage = 'ユーザー情報の取得に失敗しました。';
    if (error.message.includes("Authorization is required")) {
      userMessage += 'スクリプトの承認が必要です。ページを再読み込みして承認してください。';
    } else {
      userMessage += `詳細: ${error.message}`;
    }
    return { success: false, message: userMessage };
  }
}

/**
:start_line:271
:end_line:314
-------
 * スプレッドシートの全シートから現在のユーザーに関連するケースデータを取得する関数。
 * キャッシュを利用して高速化。
 * @return {Object} 結果オブジェクト { success: boolean, data?: Array<Object>, message?: string }
 */
function getAllSheetData() {
  const allSheetNamesResult = getAllSheetNames();
  if (!allSheetNamesResult.success) {
    return { success: false, message: allSheetNamesResult.message };
  }
  const sheetNames = allSheetNamesResult.sheetNames;

  // 現在のユーザー情報を取得
  const userInfo = getCurrentUser();
  if (!userInfo.success) {
    return { success: false, message: userInfo.message };
  }
  const currentUserLdap = userInfo.ldap;

  let allData = [];
  let firstHeaders = null; // 最初のシートのヘッダーを基準とする

  for (const sheetName of sheetNames) {
    // 'index' シートはスキップ
    if (sheetName.toLowerCase() === 'index') {
      continue;
    }

    const access = getSpreadsheetAccess(sheetName);
    if (!access.success) {
      Logger.log(`シート "${sheetName}" へのアクセスに失敗: ${access.message}`);
      continue; // エラーが発生したシートはスキップ
    }
    const sheet = access.sheet;

    try {
      // ヘッダー行を取得
      const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
      const headers = headerRange.getValues()[0];
      if (!firstHeaders) {
        firstHeaders = headers; // 最初の有効なシートのヘッダーを保存
      }

      const headerMap = {};
      headers.forEach((h, i) => { if (h) headerMap[h] = i; });

      // 必要なカラムのインデックスを取得
      const caseIdIndex = headerMap['Case ID'];
      const firstAssigneeIndex = headerMap['1st Assignee'];
      const caseStatusIndex = headerMap['Case Status']; // Case Status のインデックス

      // 必要なカラムが存在するかチェック
      if (caseIdIndex === undefined || firstAssigneeIndex === undefined || caseStatusIndex === undefined) {
        Logger.log(`シート "${sheetName}" に必要なヘッダー（Case ID, 1st Assignee, Case Status）が見つかりません。スキップします。`);
        continue;
      }

      // データ行を取得 (3行目以降)
      const lastRow = sheet.getLastRow();
      if (lastRow < HEADER_ROW + 1) {
        continue; // データがない場合はスキップ
      }
      const dataRange = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn());
      const values = dataRange.getDisplayValues();

      // データをオブジェクトの配列に変換し、フィルタリング
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const caseId = row[caseIdIndex];
        const firstAssignee = row[firstAssigneeIndex] || '';
        const caseStatus = row[caseStatusIndex];

        // Case IDが空、または Case Status が "Assigned" でない行はスキップ
        if (!caseId || caseStatus !== 'Assigned') {
          continue;
        }

        // 現在のユーザーが担当者(LDAP)に含まれるケースのみを抽出
        if (firstAssignee.includes(currentUserLdap)) {
          const rowData = { _sheetName: sheetName }; // シート名を付与
          headers.forEach((header, index) => {
            if (header) {
              rowData[header] = row[index];
            }
          });
          allData.push(rowData);
        }
      }
    } catch (error) {
      Logger.log(`シート "${sheetName}" のデータ取得中にエラー: ${error} - ${error.stack}`);
      // エラーが発生しても他のシートの処理を続行
    }
  }

  Logger.log(`全シートデータ取得成功: ${allData.length}件の担当ケースを取得 (ユーザー: ${currentUserLdap})`);
  // headers は返さず、データのみを返す（ヘッダー構造がシートごとに異なる可能性があるため）
  return { success: true, data: allData };
}

/**
 * キャッシュからヘッダー情報を取得する
 * @return {Array<string>|null} キャッシュされたヘッダー配列、またはnull
 */
function getCachedHeaders() {
  try {
    const cache = CacheService.getUserCache();
    const cachedData = cache.get('sheetHeaders');
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    return null;
  } catch (error) {
    Logger.log(`キャッシュ取得エラー: ${error}`);
    return null; // キャッシュエラーは無視して通常処理を続行
  }
}

/**
 * ヘッダー情報をキャッシュに保存する
 * @param {Array<string>} headers - キャッシュするヘッダー配列
 */
function setCachedHeaders(headers) {
  try {
    const cache = CacheService.getUserCache();
    cache.put('sheetHeaders', JSON.stringify(headers), CACHE_EXPIRATION);
  } catch (error) {
    Logger.log(`キャッシュ保存エラー: ${error}`);
    // キャッシュエラーは無視して通常処理を続行
  }
}

/**
 * 特定のケースIDでスプレッドシートからデータを検索する関数。
 * Case Statusに関わらず検索する。
 * @param {string} caseId - 検索するケースID
 * @return {Object} 結果オブジェクト { success: boolean, rowIndex?: number, data?: Object, message?: string }
 */
function searchCaseById(caseId) {
  if (!caseId) {
    return { success: false, message: '検索するCase IDを指定してください。' };
  }

  const access = getSpreadsheetAccess();
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // ヘッダー行を取得
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => { if (h) headerMap[h] = i; });

    const caseIdIndex = headerMap['Case ID']; // B列
    if (caseIdIndex === undefined) {
      return { success: false, message: 'ヘッダー「Case ID」が見つかりません。スプレッドシートを確認してください。' };
    }

    // データ範囲を取得して検索
    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      return { success: false, message: `ケースID "${caseId}" は見つかりませんでした。(データなし)` };
    }
    const dataRange = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn());
    const values = dataRange.getDisplayValues(); // 表示値で検索

    let foundRowIndex = -1;
    let foundData = null;

    for (let i = 0; i < values.length; i++) {
      if (values[i][caseIdIndex] === caseId) {
        foundRowIndex = i + HEADER_ROW + 1; // 実際の行番号
        foundData = {};
        headers.forEach((header, index) => {
          if (header) {
            foundData[header] = values[i][index];
          }
        });
        break; // 最初に見つかったものを返す
      }
    }

    if (foundRowIndex === -1) {
      Logger.log(`ケース検索: ID "${caseId}" は見つかりませんでした。`);
      return { success: false, message: `ケースID "${caseId}" が見つかりません。` };
    }

    Logger.log(`ケース検索成功: ID "${caseId}" を行 ${foundRowIndex} で発見。`);
    return { success: true, rowIndex: foundRowIndex, data: foundData };

  } catch (error) {
    Logger.log(`ケース検索エラー: ${error} - ${error.stack}`);
    return { success: false, message: `ケースの検索中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * Case ID(B列)と1st Assignee(L列)の両方が空欄の最初の行を見つける関数。
 * 3行目から検索を開始する。
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @return {number} 空の行の行番号。見つからない場合は最終行+1を返す。
 */
function findEmptyRow(sheet) {
  const lastRow = sheet.getLastRow();
  // データがない、またはヘッダーしかない場合
  if (lastRow < HEADER_ROW + 1) {
    return HEADER_ROW + 1; // 最初のデータ行 (3行目)
  }

  // B列(Case ID, index 1) と L列(1st Assignee, index 11) の値を取得
  // 検索範囲は3行目から最終行まで
  const range = sheet.getRange(HEADER_ROW + 1, 2, lastRow - HEADER_ROW, 11); // B列からL列まで取得
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    // B列の値 (values[i][0]) と L列の値 (values[i][10]) が両方空かチェック
    if (values[i][0] === '' && values[i][10] === '') {
      const emptyRow = i + HEADER_ROW + 1; // 実際の行番号
      Logger.log(`空行発見: ${emptyRow}行目`);
      return emptyRow;
    }
  }

  // 空の行が見つからなかった場合は、最終行の次の行を返す
  const nextRow = lastRow + 1;
  Logger.log(`空行が見つからず。次の書き込み行: ${nextRow}`);
  return nextRow;
}

/**
 * 新しいケースデータをスプレッドシートに追加する関数。
 * B列とL列が空の最初の行に書き込む。
 * @param {Object} caseData - フロントエンドから送信されたケースデータ
 * @return {Object} 結果オブジェクト { success: boolean, message: string, rowIndex?: number }
 */
function addNewCase(caseData) {
  const access = getSpreadsheetAccess();
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // 現在のユーザー情報を取得
    const userInfo = getCurrentUser();
    if (!userInfo.success) {
      return { success: false, message: userInfo.message };
    }
    const currentUserLdap = userInfo.ldap;

    // 書き込む行を探す
    const targetRow = findEmptyRow(sheet);

    // フロントエンドからのデータとデフォルト値をマッピング
    // スプレッドシートのカラム順序に合わせて配列を作成
    const rowData = [];
    // A: Case (ハイパーリンク) - スプレッドシートのARRAYFORMULAに任せる
    rowData[0] = ''; 
    // B: Case ID
    rowData[1] = caseData.caseId || '';
    // C: Date
    rowData[2] = caseData.caseOpenDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
    // D: Time
    rowData[3] = caseData.caseOpenTime || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
    // E: Segment
    rowData[4] = caseData.segment || '';
    // F: Product Category
    rowData[5] = caseData.productCategory || '';
    // G: Triage (チェックボックスを1/0に変換)
    rowData[6] = caseData.triage ? 1 : 0;
    // H: Either (チェックボックスを1/0に変換)
    rowData[7] = caseData.either ? 1 : 0;
    // I: 3.0 (チェックボックスを1/0に変換)
    rowData[8] = caseData["3.0"] ? 1 : 0;
    // J: Issue Category
    rowData[9] = caseData.issueCategory || '';
    // K: (空列)
    rowData[10] = '';
    // L: 1st Assignee (LDAPのみ)
    rowData[11] = currentUserLdap;
    // M to R: (空列) - 6列分
    for (let i = 0; i < 6; i++) { rowData.push(''); }
    // S: Final Assignee (デフォルトは1st Assigneeと同じLDAP)
    rowData[18] = caseData.finalAssignee || currentUserLdap;
    // T: Segment (再度 - 仕様通り)
    rowData[19] = caseData.segment || ''; // E列と同じ値
    // U: Case Status
    rowData[20] = caseData.caseStatus || 'Assigned';
    // V: AM Transfer
    rowData[21] = caseData.amTransfer || '';
    // W: non NCC
    rowData[22] = caseData.nonNCC || '';
    // X: Bug (チェックボックスを1/0に変換)
    rowData[23] = caseData.bug ? 1 : 0;
    // Y: Need Info (チェックボックスを1/0に変換)
    rowData[24] = caseData.needInfo ? 1 : 0;
    // Z: 1st Close Date
    rowData[25] = ''; // 新規作成時は空
    // AA: 1st Close Time
    rowData[26] = ''; // 新規作成時は空
    // AB: Reopen Reason
    rowData[27] = ''; // 新規作成時は空
    // AC: Reopen Close Date
    rowData[28] = ''; // 新規作成時は空
    // AD: Reopen Close Time
    rowData[29] = ''; // 新規作成時は空
    // AE列以降がある場合も考慮 (rowData.lengthまで書き込む)

    // データを書き込む (1行分、rowDataの要素数分の列)
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    Logger.log(`新規ケース追加成功: 行 ${targetRow}, Case ID ${caseData.caseId}, Assignee ${currentUserLdap}`);

    // キャッシュ更新のため、ヘッダー行のA2セルを再設定 (スプレッドシート関数再計算トリガーの代替)
    try {
      sheet.getRange(HEADER_ROW, 1).setValue(sheet.getRange(HEADER_ROW, 1).getValue());
    } catch (e) {
      Logger.log(`キャッシュ更新用セル操作中の軽微なエラー: ${e}`);
    }

    return { success: true, message: `ケース ${caseData.caseId} が正常に追加されました。`, rowIndex: targetRow };

  } catch (error) {
    Logger.log(`ケース追加エラー: ${error} - ${error.stack}`);
    return { success: false, message: `ケースの追加中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * 既存のケースデータを更新する関数。主に任意項目やステータス変更に使われる。
 * @param {number} rowIndex - 更新する行番号
 * @param {Object} caseData - 更新するデータ (例: { caseStatus: 'Finished', nonNCC: 'Duplicate', ... })
 * @return {Object} 結果オブジェクト { success: boolean, message: string }
 */
function updateCaseData(rowIndex, caseData) {
  if (!rowIndex || typeof rowIndex !== 'number' || rowIndex <= HEADER_ROW) {
    return { success: false, message: '無効な行番号が指定されました。' };
  }
  if (!caseData || typeof caseData !== 'object') {
    return { success: false, message: '更新データが無効です。' };
  }

  const access = getSpreadsheetAccess();
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // ヘッダーを取得して列インデックスをマッピング
    const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => { if (h) headerMap[h] = i + 1; }); // 1-based index

    // 更新対象の列と値を特定
    const updates = {}; // { columnIndex: value }
    
    // 更新可能なフィールドを定義 (キー: フロントエンドのキー, 値: スプレッドシートのヘッダー名)
    const editableFields = {
      'caseStatus': 'Case Status', // U列
      'amTransfer': 'AM Transfer', // V列
      'nonNCC': 'non NCC',        // W列
      'bug': 'Bug',               // X列
      'needInfo': 'Need Info',    // Y列
      'firstCloseDate': '1st Close Date', // Z列
      'firstCloseTime': '1st Close Time', // AA列
      'reopenReason': 'Reopen Reason', // AB列
      'reopenCloseDate': 'Reopen Close Date', // AC列
      'reopenCloseTime': 'Reopen Close Time', // AD列
      'tsConsulted': 'T&S Consulted' // スプレッドシートの該当列名に修正してください
    };

    for (const key in editableFields) {
      if (caseData.hasOwnProperty(key) && caseData[key] !== undefined) { // データが存在し、undefinedでない場合
        const headerName = editableFields[key];
        const colIndex = headerMap[headerName];
        if (colIndex) {
          // チェックボックスやスイッチの場合は TRUE/FALSE または Yes/No に変換
          if (key === 'bug' || key === 'needInfo') {
            updates[colIndex] = caseData[key] ? 1 : 0; // 1/0 形式の場合
          } else if (key === 'tsConsulted') {
            updates[colIndex] = caseData[key] ? 'Yes' : 'No'; // Yes/No 形式の場合
            updates[colIndex] = caseData[key] ? 1 : 0;
          } else {
            updates[colIndex] = caseData[key];
          }
        } else {
          Logger.log(`警告: 更新対象のヘッダー "${headerName}" がシートに見つかりません。`);
        }
      }
    }

    // Case Status が 'Solution Offered' または 'Finished' に変更された場合の処理
    const statusColIndex = headerMap['Case Status'];
    const firstCloseDateCol = headerMap['1st Close Date']; // Z列
    const firstCloseTimeCol = headerMap['1st Close Time']; // AA列
    const reopenCloseDateCol = headerMap['Reopen Close Date']; // AC列
    const reopenCloseTimeCol = headerMap['Reopen Close Time']; // AD列

    if (statusColIndex && updates[statusColIndex] && 
       (updates[statusColIndex] === 'Solution Offered' || updates[statusColIndex] === 'Finished')) {
       
      // 現在の1st Close Dateの値を取得して、初回クローズか再クローズかを判断
      let firstCloseDateValue = '';
      if (firstCloseDateCol) {
        firstCloseDateValue = sheet.getRange(rowIndex, firstCloseDateCol).getDisplayValue();
      }

      // ユーザーが明示的に日時を指定していない場合のみ自動設定
      const now = new Date();
      const dateString = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");
      const timeString = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

      if (!firstCloseDateValue && firstCloseDateCol && firstCloseTimeCol) {
        // 初回クローズ日時を更新（ユーザーが明示的に指定していない場合のみ）
        if (!updates[firstCloseDateCol]) updates[firstCloseDateCol] = dateString;
        if (!updates[firstCloseTimeCol]) updates[firstCloseTimeCol] = timeString;
        Logger.log(`初回クローズ日時を更新: 行 ${rowIndex}`);
      } else if (reopenCloseDateCol && reopenCloseTimeCol) {
        // 再オープンクローズ日時を更新（ユーザーが明示的に指定していない場合のみ）
        if (!updates[reopenCloseDateCol]) updates[reopenCloseDateCol] = dateString;
        if (!updates[reopenCloseTimeCol]) updates[reopenCloseTimeCol] = timeString;
        Logger.log(`再オープンクローズ日時を更新: 行 ${rowIndex}`);
      }
    }

    // 実際のセル更新
    let updatedCount = 0;
    for (const colIndexStr in updates) {
      const colIndex = parseInt(colIndexStr, 10);
      const value = updates[colIndex];
      try {
        sheet.getRange(rowIndex, colIndex).setValue(value);
        updatedCount++;
      } catch (cellError) {
        Logger.log(`警告: 行 ${rowIndex}, 列 ${colIndex} の更新に失敗しました: ${cellError}`);
      }
    }
    
    if (updatedCount > 0) {
      Logger.log(`ケース更新成功: 行 ${rowIndex}, ${updatedCount}項目を更新。`);
      return { success: true, message: 'ケースが正常に更新されました。' };
    } else {
      Logger.log(`ケース更新: 行 ${rowIndex}, 更新対象の有効なデータがありませんでした。`);
      return { success: false, message: '更新するデータが指定されていないか、対象列が見つかりませんでした。' };
    }
  } catch (error) {
    Logger.log(`ケース更新エラー: ${error} - ${error.stack}`);
    return { success: false, message: `ケースの更新中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * NCC (完了ケース数) の統計データを取得する関数。
 * 指定された期間 (daily, weekly, monthly, quarterly) ごとに集計する。
 * @param {string} period - 集計期間 ('daily', 'weekly', 'monthly', 'quarterly')
 * @return {Object} 結果オブジェクト { success: boolean, stats?: Object, message?: string }
 */
function getNCCStats(period) {
  const access = getSpreadsheetAccess();
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // ヘッダー行を取得
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => { if (h) headerMap[h] = i; }); // 0-based index

    // 必要な列のインデックスを取得
    const dateColIndex = headerMap['Date']; // C列
    const caseStatusColIndex = headerMap['Case Status']; // U列
    const firstAssigneeColIndex = headerMap['1st Assignee']; // L列
    const caseIdIndex = headerMap['Case ID']; // B列 (空行スキップ用)
    const nonNCCIndex = headerMap['non NCC']; // W列

    if (dateColIndex === undefined || caseStatusColIndex === undefined || 
        firstAssigneeColIndex === undefined || caseIdIndex === undefined || 
        nonNCCIndex === undefined) {
      return { 
        success: false, 
        message: '統計に必要なヘッダー（Date, Case Status, 1st Assignee, Case ID, non NCC）が見つかりません。' 
      };
    }

    // データ範囲を取得
    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      return { success: true, stats: { daily: [], weekly: [], monthly: [], quarterly: [] } }; // データなし
    }
    const dataRange = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn());
    const values = dataRange.getValues(); // 日付処理のため Object[][] で取得

    // 現在のユーザー情報を取得
    const userInfo = getCurrentUser();
    if (!userInfo.success) {
      return { success: false, message: userInfo.message };
    }
    const currentUserLdap = userInfo.ldap;

    // 統計データを格納するオブジェクト
    const stats = {
      daily: {}, weekly: {}, monthly: {}, quarterly: {}
    };

    // ステータス別のカウントを保持するオブジェクト
    const statusCounts = {
      daily: {}, weekly: {}, monthly: {}, quarterly: {}
    };

    // 各行を処理
    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      // Case IDが空の行はスキップ
      if (!row[caseIdIndex]) continue;

      // 現在のユーザーのケースのみ処理
      const firstAssignee = row[firstAssigneeColIndex] || '';
      if (!firstAssignee.includes(currentUserLdap)) continue;

      // 日付とケースステータスを取得
      const dateValue = row[dateColIndex];
      const caseStatus = row[caseStatusColIndex];
      const nonNCC = row[nonNCCIndex];

      // NCCの定義: non NCC が空欄で、Case Status が Assigned 以外
      const isNCC = !nonNCC && caseStatus && caseStatus !== 'Assigned';

      // 日付が無効な場合はスキップ
      if (!dateValue || !(dateValue instanceof Date)) {
        Logger.log(`警告: 行 ${i + HEADER_ROW + 1} の日付が無効です (${dateValue})。統計から除外します。`);
        continue;
      }
      const date = dateValue; // new Date(dateValue) は不要、既にDateオブジェクト

      // 期間キーを生成
      const dayStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const weekStr = getWeekNumber(date);
      const monthStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
      const quarterStr = getQuarter(date);

      // 各期間の統計を初期化・更新
      [ 
        {key: dayStr, obj: stats.daily, statusObj: statusCounts.daily},
        {key: weekStr, obj: stats.weekly, statusObj: statusCounts.weekly},
        {key: monthStr, obj: stats.monthly, statusObj: statusCounts.monthly},
        {key: quarterStr, obj: stats.quarterly, statusObj: statusCounts.quarterly}
      ].forEach(p => {
        if (!p.obj[p.key]) {
          p.obj[p.key] = { total: 0, finished: 0 };
        }
        if (!p.statusObj[p.key]) {
          p.statusObj[p.key] = {};
        }
        
        p.obj[p.key].total++;
        if (isNCC) {
          p.obj[p.key].finished++;
        }
        
        // ステータス別カウント
        if (caseStatus) {
          if (!p.statusObj[p.key][caseStatus]) {
            p.statusObj[p.key][caseStatus] = 0;
          }
          p.statusObj[p.key][caseStatus]++;
        }
      });
    }

    // 結果を整形して返す
    const result = {};
    for (const p in stats) {
      result[p] = Object.keys(stats[p])
        .sort() // 時系列順にソート
        .map(key => ({
          period: key,
          total: stats[p][key].total,
          finished: stats[p][key].finished,
          rate: (stats[p][key].total > 0 ? (stats[p][key].finished / stats[p][key].total * 100).toFixed(2) : '0.00'),
          statusBreakdown: statusCounts[p][key] || {}
        }));
    }

    Logger.log(`統計データ取得成功: 期間 ${period || '全期間'}, ユーザー ${currentUserLdap}`);
    // 要求された期間、または全期間のデータを返す
    return { success: true, stats: period ? result[period] : result };

  } catch (error) {
    Logger.log(`統計データ取得エラー: ${error} - ${error.stack}`);
    return { success: false, message: `統計データの取得中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * 日付からISO 8601形式の週番号を取得するヘルパー関数。
 * @param {Date} date - 対象の日付オブジェクト
 * @return {string} 年と週番号の文字列 (例: '2023-W43')
 */
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  // Set to nearest Thursday: current date + 4 - current day number
  // Make Sunday's day number 7
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  // Get first day of year
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  // Calculate full weeks to nearest Thursday
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  // Return array of year and week number
  return d.getUTCFullYear() + '-W' + String(weekNo).padStart(2, '0');
}

/**
 * 日付から四半期を取得するヘルパー関数。
 * @param {Date} date - 対象の日付オブジェクト
 * @return {string} 年と四半期の文字列 (例: '2023-Q4')
 */
function getQuarter(date) {
  const month = date.getMonth(); // 0-11
  const quarter = Math.floor(month / 3) + 1;
  return date.getFullYear() + '-Q' + quarter;
}