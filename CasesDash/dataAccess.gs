/**
 * @fileoverview dataAccess.gs - データアクセスモジュール
 * スプレッドシートからのデータ取得、検索、追加、更新を担当します。
 * ヘッダー情報のキャッシュ管理も行います。
 * P95デッドラインの計算もここで行います。
 * @requires settings.gs
 * @requires Code.gs
 */

// 定数 (Code.gs から移動)
const CACHE_EXPIRATION = 300; // キャッシュ有効期間（秒）
const P95_DURATION_HOURS = 72; // P95タイマーの期間（時間）
const MAX_API_RETRIES = 3; // API呼び出し最大リトライ回数 (一時的なエラー用)

/**
 * API呼び出しの共通ラッパー関数。
 * ネットワークエラーなど一時的な問題に対して指定回数までリトライを行う。
 * 内部関数が { success: false } を返した場合はリトライしない。
 * @param {Function} apiFunc - 実行するAPI関数（引数なしで呼び出せるもの）。
 * @returns {Object} API呼び出し結果。
 */
function callApiWithRetry(apiFunc) {
  let attempt = 0;
  let lastError = null;
  while (attempt < MAX_API_RETRIES) {
    try {
      const result = apiFunc();
      // 成功した場合、または回復不能なエラー ({ success: false }) の場合は即座に返す
      if (result && (result.success === true || result.success === false)) {
        return result;
      }
      // result が予期しない形式の場合 (success プロパティがないなど) はエラーとして扱う
      lastError = new Error("API関数が予期しない結果を返しました。");
      Logger.log(`API呼び出し試行 ${attempt + 1} 失敗: 予期しない結果 ${JSON.stringify(result)}`);

    } catch (e) {
      // 実行時エラー (ネットワークエラーなど) の場合
      lastError = e;
      Logger.log(`API呼び出し試行 ${attempt + 1} でエラー発生: ${e}`);
    }
    // リトライ前に待機
    attempt++;
    if (attempt < MAX_API_RETRIES) {
        Utilities.sleep(1000 * attempt); // エクスポネンシャルバックオフ (ミリ秒)
        Logger.log(`${1000 * attempt}ms 待機してリトライします... (${attempt}/${MAX_API_RETRIES})`);
    }
  }

  // 最大リトライ回数に達した場合
  Logger.log(`API呼び出しが最大リトライ回数 (${MAX_API_RETRIES}) に達しました。最後のエラー: ${lastError}`);
  return {
    success: false,
    message: `API呼び出しが最大リトライ回数に達しました。しばらくしてから再試行してください。 (エラー: ${lastError?.message || '不明'})`,
  };
}

/**
 * キャッシュからヘッダー情報を取得します。
 * @private
 * @function getCachedHeaders
 * @returns {Array<string>|null} キャッシュされたヘッダー配列。キャッシュが存在しないかエラーの場合はnull。
 */
function getCachedHeaders() {
  try {
    const cache = CacheService.getUserCache();
    const cachedData = cache.get("sheetHeaders");
    if (cachedData) {
      Logger.log("ヘッダーキャッシュをヒットしました。");
      return JSON.parse(cachedData);
    }
    Logger.log("ヘッダーキャッシュが見つかりませんでした。");
    return null;
  } catch (error) {
    Logger.log(`ヘッダーキャッシュ取得エラー: ${error}`);
    return null; // キャッシュエラーは無視して通常処理を続行
  }
}

/**
 * ヘッダー情報をキャッシュに保存します。
 * @private
 * @function setCachedHeaders
 * @param {Array<string>} headers - キャッシュするヘッダー配列。
 */
function setCachedHeaders(headers) {
  try {
    const cache = CacheService.getUserCache();
    cache.put("sheetHeaders", JSON.stringify(headers), CACHE_EXPIRATION);
    Logger.log(
      `ヘッダー情報をキャッシュに保存しました (有効期間: ${CACHE_EXPIRATION}秒)`
    );
  } catch (error) {
    Logger.log(`ヘッダーキャッシュ保存エラー: ${error}`);
    // キャッシュエラーは無視して通常処理を続行
  }
}

/**
 * スプレッドシートから現在のユーザーに関連するケースデータを取得します。
 * ヘッダー情報はキャッシュを利用して高速化を図ります。
 * 各ケースについて P95 デッドラインも計算して付加します。
 * フロントエンド (例: script.html) から呼び出されることを想定しています。
 *
 * @function getSheetData
 * @returns {{success: boolean, headers?: Array<string>, data?: Array<Object<string, string>>, message?: string}}
 *           成功時は { success: true, headers: Array<string>, data: Array<Object<string, string>> }、
 *           失敗時は { success: false, message: string } を返します。
 *           各 data オブジェクトには計算された `p95Deadline` (ISO 8601 文字列) が追加されます。
 */
function getSheetData() {
  // callApiWithRetry は一時的なエラーのリトライ用
  // getSheetDataInternal 内で回復不能なエラー (ヘッダー不足など) が発生した場合は、
  // { success: false } が返され、callApiWithRetry はリトライせずにそれを返す。
  return callApiWithRetry(getSheetDataInternal_);
}

/**
 * getSheetData の内部処理。callApiWithRetry から呼び出される。
 * @private
 * @returns {{success: boolean, headers?: Array<string>, data?: Array<Object<string, string>>, message?: string}}
 */
function getSheetDataInternal_() {
  const access = getSpreadsheetAccess(); // 関数名を修正
  if (!access.success) {
    // getSpreadsheetAccess_ で既にエラーメッセージ設定済み
    return access;
  }
  const sheet = access.sheet;

  try {
    // 1. ヘッダー取得と検証
    let headers = getCachedHeaders();
    if (!headers) {
      const headerRange = sheet.getRange(
        HEADER_ROW,
        1,
        1,
        sheet.getLastColumn()
      );
      headers = headerRange.getValues()[0].map((h) => String(h).trim());
      setCachedHeaders(headers);
    }

    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i; // 0-based index
    });

    // データ取得とP95計算に必要なヘッダーを定義
    const requiredHeadersForData = ["Case ID", "1st Assignee", "Case Status"]; // Case Status も追加
    const requiredHeadersForP95 = ["Date", "Time"];
    const allRequiredHeaders = [
      ...new Set([...requiredHeadersForData, ...requiredHeadersForP95]),
    ];
    const missingHeaders = allRequiredHeaders.filter(
      (h) => !(h in headerMap)
    );

    if (missingHeaders.length > 0) {
      // 回復不能なエラーなので { success: false } を返す
      return {
        success: false,
        message: `データ取得に必要なヘッダーが見つかりません: ${missingHeaders.join(
          ", "
        )}。スプレッドシートの ${HEADER_ROW}行目を確認してください。`,
      };
    }
    // 必要な列のインデックスを取得
    const caseIdIndex = headerMap["Case ID"];
    const firstAssigneeIndex = headerMap["1st Assignee"];
    const caseStatusIndex = headerMap["Case Status"];
    const dateIndex = headerMap["Date"];
    const timeIndex = headerMap["Time"];

    // 2. データ一括取得
    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      Logger.log("シートにデータ行がありません。");
      return { success: true, headers: headers, data: [] }; // データなしは成功
    }
    const dataRange = sheet.getRange(
      HEADER_ROW + 1,
      1,
      lastRow - HEADER_ROW,
      headers.length
    );
    const values = dataRange.getValues(); // シート全体のデータを一度に取得

    // 3. ユーザー情報取得
    const userInfo = getCurrentUser();
    if (!userInfo.success) {
      // ユーザー情報取得失敗も回復不能エラー
      return { success: false, message: userInfo.message };
    }
    const currentUserLdap = userInfo.ldap;

    // 4. データ処理とフィルタリング
    const data = [];
    const scriptTz = Session.getScriptTimeZone();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const caseId = row[caseIdIndex];
      // Case ID が空の行はスキップ
      if (!caseId) continue;

      // 担当者チェック (currentUserLdap が 1st Assignee に含まれるか)
      const firstAssignee = String(row[firstAssigneeIndex] || "").trim();
      if (firstAssignee.includes(currentUserLdap)) {
        const rowData = {};
        // ヘッダーに基づいてデータをオブジェクトに格納
        for (const headerName in headerMap) {
          const colIndex = headerMap[headerName];
          const cellValue = row[colIndex];
          // 日付/時刻のフォーマット処理
          if (cellValue instanceof Date) {
            if (headerName.toLowerCase().includes("date")) {
              rowData[headerName] = Utilities.formatDate(
                cellValue,
                scriptTz,
                "yyyy/MM/dd"
              );
            } else if (headerName.toLowerCase().includes("time")) {
              rowData[headerName] = Utilities.formatDate(
                cellValue,
                scriptTz,
                "HH:mm:ss"
              );
            } else {
              // その他の Date オブジェクトは ISO 形式に
              rowData[headerName] = cellValue.toISOString();
            }
          } else {
            // Date 以外は文字列に変換
            rowData[headerName] = String(cellValue);
          }
        }

        // P95 デッドライン計算
        const caseOpenDateValue = row[dateIndex]; // values 配列から直接取得
        const caseOpenTimeValue = row[timeIndex]; // values 配列から直接取得

        let p95DeadlineISO = null;
        // Date 列が有効な Date オブジェクトか確認
        if (caseOpenDateValue instanceof Date && !isNaN(caseOpenDateValue.getTime())) {
          let startTimeMillis = caseOpenDateValue.getTime();

          // Time 列も有効な Date オブジェクトなら日時を組み合わせる
          if (caseOpenTimeValue instanceof Date && !isNaN(caseOpenTimeValue.getTime())) {
            const combinedDate = new Date(caseOpenDateValue);
            combinedDate.setHours(caseOpenTimeValue.getHours());
            combinedDate.setMinutes(caseOpenTimeValue.getMinutes());
            combinedDate.setSeconds(caseOpenTimeValue.getSeconds());
            combinedDate.setMilliseconds(caseOpenTimeValue.getMilliseconds());
            startTimeMillis = combinedDate.getTime();
          }
          // Time 列が無効でも Date 列だけで計算は可能

          const durationHours = P95_DURATION_HOURS;
          const deadlineMillis = startTimeMillis + durationHours * 60 * 60 * 1000;
          p95DeadlineISO = new Date(deadlineMillis).toISOString();
        } else {
          Logger.log(`警告: Case ID ${caseId} の Date 列が無効なため P95 デッドラインを計算できません。`);
        }
        rowData["p95Deadline"] = p95DeadlineISO; // 計算結果を格納

        data.push(rowData);
      }
    }
    Logger.log(`getSheetData 成功: ${data.length} 件のデータを処理しました。`);
    return { success: true, headers: headers, data: data };

  } catch (error) {
    // 予期せぬエラー
    Logger.log(`getSheetDataInternal_ で予期せぬエラー: ${error} - ${error.stack}`);
    // 詳細なエラー情報を返す
    return {
      success: false,
      message: `データの取得中に予期せぬエラーが発生しました: ${error.message}`,
    };
  }
}


/**
 * 特定のケースIDでスプレッドシートからデータを検索します。
 * Case Statusに関わらず、シート全体を検索します。
 * フロントエンド (例: script.html) から呼び出されることを想定しています。
 *
 * @function searchCaseById
 * @param {string} caseId - 検索するケースID。
 * @returns {{success: boolean, rowIndex?: number, data?: Object<string, string>, message?: string}}
 *           成功時は { success: true, rowIndex: number, data: Object<string, string> }、
 *           見つからない場合は { success: false, message: string }、
 *           エラー時は { success: false, message: string } を返します。
 *           data はヘッダー名をキー、セルの値を文字列として持つオブジェクトです。
 */
function searchCaseById(caseId) {
  if (!caseId) {
    return { success: false, message: "検索するCase IDを指定してください。" };
  }

  // getSpreadsheetAccess は内部でエラーハンドリング済み
  const access = getSpreadsheetAccess(); // 関数名を修正
  if (!access.success) {
    return access;
  }
  const sheet = access.sheet;

  try {
    // 1. ヘッダー取得と検証
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0].map((h) => String(h).trim());
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i; // 0-based index
    });

    const caseIdHeader = "Case ID";
    const dateHeader = "Date"; // P95計算用
    const timeHeader = "Time"; // P95計算用

    if (!(caseIdHeader in headerMap)) {
      return {
        success: false,
        message: `検索に必要なヘッダー「${caseIdHeader}」が見つかりません。`,
      };
    }
    if (!(dateHeader in headerMap) || !(timeHeader in headerMap)) {
      Logger.log(`警告: P95計算に必要なヘッダー (${dateHeader}, ${timeHeader}) が見つかりません。`);
      // P95計算はスキップされるが、検索自体は続行可能
    }
    const caseIdIndex = headerMap[caseIdHeader];
    const dateIndex = headerMap[dateHeader]; // なければ undefined
    const timeIndex = headerMap[timeHeader]; // なければ undefined

    // 2. データ一括取得と検索
    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      return { success: false, message: `ケースID "${caseId}" は見つかりませんでした。(データなし)` };
    }
    const dataRange = sheet.getRange(
      HEADER_ROW + 1,
      1,
      lastRow - HEADER_ROW,
      headers.length
    );
    // getDisplayValues で文字列として取得 (検索と表示用)
    const displayValues = dataRange.getDisplayValues();
    // getValue で Date オブジェクトを取得 (P95計算用)
    const dateValues = (dateIndex !== undefined && timeIndex !== undefined)
        ? sheet.getRange(HEADER_ROW + 1, Math.min(dateIndex, timeIndex) + 1, lastRow - HEADER_ROW, Math.abs(dateIndex - timeIndex) + 1).getValues()
        : null; // Date/Time列がない場合は null

    let foundRowIndex = -1; // シート上の実際の行番号
    let foundData = null;

    for (let i = 0; i < displayValues.length; i++) {
      if (displayValues[i][caseIdIndex] === caseId) {
        foundRowIndex = i + HEADER_ROW + 1; // 実際の行番号
        foundData = {};
        // ヘッダーに基づいて表示用データを格納
        for (const headerName in headerMap) {
          foundData[headerName] = displayValues[i][headerMap[headerName]];
        }

        // P95 デッドライン計算
        let p95DeadlineISO = null;
        if (dateValues && dateIndex !== undefined && timeIndex !== undefined) {
            // dateValues から対応する行の Date/Time を取得
            const dateColOffset = dateIndex - Math.min(dateIndex, timeIndex);
            const timeColOffset = timeIndex - Math.min(dateIndex, timeIndex);
            const caseOpenDateValue = dateValues[i][dateColOffset];
            const caseOpenTimeValue = dateValues[i][timeColOffset];

            if (caseOpenDateValue instanceof Date && !isNaN(caseOpenDateValue.getTime())) {
                let startTimeMillis = caseOpenDateValue.getTime();
                if (caseOpenTimeValue instanceof Date && !isNaN(caseOpenTimeValue.getTime())) {
                    const combinedDate = new Date(caseOpenDateValue);
                    combinedDate.setHours(caseOpenTimeValue.getHours());
                    combinedDate.setMinutes(caseOpenTimeValue.getMinutes());
                    combinedDate.setSeconds(caseOpenTimeValue.getSeconds());
                    combinedDate.setMilliseconds(caseOpenTimeValue.getMilliseconds());
                    startTimeMillis = combinedDate.getTime();
                }
                const durationHours = P95_DURATION_HOURS;
                const deadlineMillis = startTimeMillis + durationHours * 60 * 60 * 1000;
                p95DeadlineISO = new Date(deadlineMillis).toISOString();
            }
        }
        foundData["p95Deadline"] = p95DeadlineISO; // 計算結果を追加

        break; // 最初に見つかった行で終了
      }
    }

    if (foundRowIndex !== -1 && foundData) {
      Logger.log(`ケース検索成功: ID "${caseId}" を行 ${foundRowIndex} で発見。`);
      return { success: true, rowIndex: foundRowIndex, data: foundData };
    } else {
      Logger.log(`ケース検索: ID "${caseId}" は見つかりませんでした。`);
      return { success: false, message: `ケースID "${caseId}" が見つかりません。` };
    }

  } catch (error) {
    Logger.log(`ケース検索エラー: ${error} - ${error.stack}`);
    return {
      success: false,
      message: `ケースの検索中にエラーが発生しました: ${error.message}`,
    };
  }
}


/**
 * 指定されたシート内で、Case ID(B列想定)と1st Assignee(L列想定)の両方が
 * 空欄である最初の行を見つけます。
 * HEADER_ROW + 1 行目から検索を開始します。
 *
 * @private
 * @function findEmptyRow
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 検索対象のシートオブジェクト。
 * @returns {number} 空の行の行番号。見つからない場合は最終行+1を返します。
 */
function findEmptyRow(sheet) {
  const lastRow = sheet.getLastRow();
  // データがない、またはヘッダーしかない場合
  if (lastRow < HEADER_ROW + 1) {
    Logger.log(`空行検索: データなし。最初の書き込み行: ${HEADER_ROW + 1}`);
    return HEADER_ROW + 1; // 最初のデータ行
  }

  // ヘッダーを取得して列インデックスを確認 (毎回取得するのは非効率だが、列挿入に対応するため)
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0].map((h) => String(h).trim());
  const headerMap = {};
  headers.forEach((h, i) => {
    if (h) headerMap[h] = i;
  });

  const caseIdHeader = "Case ID";
  const firstAssigneeHeader = "1st Assignee";

  if (!(caseIdHeader in headerMap) || !(firstAssigneeHeader in headerMap)) {
    Logger.log(
      `空行検索エラー: 必須ヘッダーが見つかりません (${caseIdHeader}, ${firstAssigneeHeader})`
    );
    // エラーを投げるか、デフォルトの列インデックスを使うか、最終行+1を返すか要検討
    // ここでは安全策として最終行+1を返す
    return lastRow + 1;
  }

  const caseIdColIndex = headerMap[caseIdHeader]; // B列想定 -> 1
  const firstAssigneeColIndex = headerMap[firstAssigneeHeader]; // L列想定 -> 11

  // B列とL列の値を取得 (効率化のため relevantColumns のみ取得)
  const startRow = HEADER_ROW + 1;
  const numRows = lastRow - HEADER_ROW;
  if (numRows <= 0) {
    // データ行がない場合
    return startRow;
  }
  // getRangeList を使うとより効率的かもしれない
  const caseIdValues = sheet
    .getRange(startRow, caseIdColIndex + 1, numRows, 1)
    .getValues();
  const assigneeValues = sheet
    .getRange(startRow, firstAssigneeColIndex + 1, numRows, 1)
    .getValues();

  for (let i = 0; i < numRows; i++) {
    // getValues() は空セルを '' で返すため、厳密比較でOK
    if (caseIdValues[i][0] === "" && assigneeValues[i][0] === "") {
      const emptyRow = i + startRow; // 実際の行番号
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
 * 新しいケースデータをスプレッドシートに追加します。
 * `findEmptyRow` で見つけた、Case ID と 1st Assignee が空の最初の行に書き込みます。
 * フロントエンド (例: script.html, live.html) から呼び出されることを想定しています。
 *
 * @function addNewCase
 * @param {Object} caseData - フロントエンドから送信されたケースデータオブジェクト。
 *                            キーはフォームのフィールド名に対応します (例: caseId, segment, productCategory)。
 * @returns {{success: boolean, message: string, rowIndex?: number}}
 *           成功時は { success: true, message: string, rowIndex: number }、
 *           失敗時は { success: false, message: string } を返します。
 */
function addNewCase(caseData) {
  const access = getSpreadsheetAccess(); // 関数名を修正
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // 現在のユーザー情報を取得 (Code.gs の関数を呼び出す)
    const userInfo = getCurrentUser();
    if (!userInfo.success) {
      return { success: false, message: userInfo.message };
    }
    const currentUserLdap = userInfo.ldap;

    // 書き込む行を探す
    const targetRow = findEmptyRow(sheet); // このモジュール内の関数を呼び出す

    // ヘッダーを取得してマッピング (列挿入に対応するため毎回取得)
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0].map((h) => String(h).trim());
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i; // 0-based index
    });

    // スプレッドシートに書き込むデータ配列を作成
    // headers 配列の順序に合わせて値を設定する
    const rowData = new Array(headers.length).fill(""); // デフォルトは空文字列

    // フロントエンドからのデータを対応する列にマッピング
    // 注意: このマッピングはスプレッドシートのヘッダー名と caseData のキー名が
    //       一致または特定のルールに従っていることを前提としています。
    //       より堅牢にするには、明確なマッピング定義が必要です。
    const now = new Date();
    const scriptTz = Session.getScriptTimeZone();
    // スプレッドシートは通常、日付/時刻を数値として保存するため、Dateオブジェクトで設定
    const todayDate = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate()
    );
    // 時刻も Date オブジェクトで設定 (日付部分は無視される想定だが、同じ日にしておく)
    const currentTimeDate = new Date(
      1899,
      11,
      30,
      now.getHours(),
      now.getMinutes(),
      now.getSeconds()
    ); // スプレッドシートの時刻シリアル値の基準日に合わせる (より安全)
    // const currentTimeDate = now; // これでも動くことが多い

    // --- 列とデータのマッピング ---
    // ヘッダー名を使って動的に列インデックスを取得し、値を設定
    function setData(headerName, value) {
      if (headerName in headerMap) {
        rowData[headerMap[headerName]] = value;
      } else {
        Logger.log(
          `警告: ヘッダー "${headerName}" が見つかりません。このフィールドの書き込みはスキップされます。`
        );
      }
    }

    setData("Case ID", caseData.caseId || "");
    // 日付と時刻は Date オブジェクトで設定
    // スプレッドシートのヘッダー名に合わせて修正
    setData(
      "Case Open Date", // D列ヘッダー
      caseData.caseOpenDate // フロントエンドのキー名
        ? new Date(caseData.caseOpenDate.replace(/-/g, "/"))
        : todayDate
    );
    setData(
      "Time", // E列ヘッダー (Case Open Time)
      caseData.caseOpenTime // フロントエンドのキー名
        ? new Date(`1899/12/30 ${caseData.caseOpenTime}`)
        : currentTimeDate
    );
    setData("Incoming Segment", caseData.segment || ""); // F列ヘッダー
    setData("Product Category", caseData.productCategory || ""); // G列ヘッダー (変更なし)
    setData("Triage", caseData.triage ? 1 : 0); // チェックボックス
    setData("Either", caseData.either ? 1 : 0); // チェックボックス
    setData("3.0", caseData["3.0"] ? 1 : 0); // チェックボックス (キー名注意)
    setData("Issue Category", caseData.issueCategory || "");
    setData("1st Assignee", currentUserLdap);
    // Final Assignee はデフォルトで 1st Assignee と同じにする
    setData("Final Assignee", caseData.finalAssignee || currentUserLdap);
    // Segment の二重設定ロジックを削除 (Incoming Segment のみ設定)
    setData("Case Status", caseData.caseStatus || "Assigned");
    setData("AM Transfer", caseData.amTransfer || "");
    setData("non NCC", caseData.nonNCC || "");
    setData("Bug", caseData.bug ? 1 : 0); // チェックボックス
    setData("Need Info", caseData.needInfo ? 1 : 0); // チェックボックス

    // その他の列 (1st Close Date など) はデフォルトで空文字列のまま

    // --- マッピングここまで ---

    // データを書き込む (1行分)
    // setValues を使うことで、日付/時刻オブジェクトが適切に処理される
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    Logger.log(
      `新規ケース追加成功: 行 ${targetRow}, Case ID ${caseData.caseId}, Assignee ${currentUserLdap}`
    );

    // キャッシュ更新のため、ヘッダー行のA列セルを再設定 (スプレッドシート関数再計算トリガーの代替)
    // HEADER_ROW は settings.gs で定義
    try {
      // A列 (インデックス 0) のヘッダーセルを操作
      if (headers.length > 0) {
        // getValue/setValue で再計算をトリガー
        const headerCell = sheet.getRange(HEADER_ROW, 1);
        headerCell.setValue(headerCell.getValue());
      }
    } catch (e) {
      Logger.log(`キャッシュ更新用セル操作中の軽微なエラー: ${e}`);
    }

    return {
      success: true,
      message: `ケース ${caseData.caseId} が正常に追加されました。`,
      rowIndex: targetRow,
    };
  } catch (error) {
    Logger.log(`ケース追加エラー: ${error} - ${error.stack}`);
    return {
      success: false,
      message: `ケースの追加中にエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * 既存のケースデータを更新します。主に任意項目やステータス変更に使用されます。
 * フロントエンド (例: script.html) から呼び出されることを想定しています。
 *
 * @function updateCaseData
 * @param {number} rowIndex - 更新する行番号 (ヘッダー行を除く、実際のシート上の行番号)。
 * @param {Object} caseData - 更新するデータを含むオブジェクト。
 *                            キーは更新可能なフィールド名 (例: caseStatus, nonNCC)、値は新しい値。
 *                            指定されなかったフィールドは更新されません。
 * @returns {{success: boolean, message: string}}
 *           成功時は { success: true, message: string }、
 *           失敗時は { success: false, message: string } を返します。
 */
function updateCaseData(rowIndex, caseData) {
  // rowIndex のバリデーション (HEADER_ROW は settings.gs で定義)
  if (!rowIndex || typeof rowIndex !== "number" || rowIndex <= HEADER_ROW) {
    return {
      success: false,
      message: `無効な行番号が指定されました: ${rowIndex}`,
    };
  }
  if (
    !caseData ||
    typeof caseData !== "object" ||
    Object.keys(caseData).length === 0
  ) {
    return { success: false, message: "更新データが無効か空です。" };
  }

  const access = getSpreadsheetAccess(); // 関数名を修正
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // ヘッダーを取得して列インデックスをマッピング (HEADER_ROW は settings.gs で定義)
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0].map((h) => String(h).trim());
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i + 1; // 1-based index for getRange()
    });

    // 更新対象の列と値を特定
    const updates = {}; // { columnIndex: value }

    // 更新可能なフィールドとスプレッドシートのヘッダー名のマッピング
    // 注意: このマッピングはスプレッドシートの構造に依存します。
    const editableFieldsMap = {
      caseStatus: "Case Status", // U列想定
      amTransfer: "AM Transfer", // V列想定
      nonNCC: "non NCC", // W列想定
      bug: "Bug", // X列想定
      needInfo: "Need Info", // Y列想定
      firstCloseDate: "1st Close Date", // Z列想定
      firstCloseTime: "1st Close Time", // AA列想定
      reopenReason: "Reopen Reason", // AB列想定
      reopenCloseDate: "Reopen Close Date", // AC列想定
      reopenCloseTime: "Reopen Close Time", // AD列想定
      sentimentScore: "Sentiment Score", // Sentiment Score 列を追加 (列名は要確認)
      // T&S Consulted は CasesDash.md に記載があるが、元のコードにはなかった。
      // 必要であればスプレッドシートに追加し、ここにもマッピングを追加する。
      // tsConsulted: "T&S Consulted",
    };

    let hasValidUpdate = false;
    for (const key in editableFieldsMap) {
      // caseData にキーが存在する場合のみ処理 (undefined も含む)
      if (caseData.hasOwnProperty(key)) {
        const headerName = editableFieldsMap[key];
        const colIndex = headerMap[headerName];
        if (colIndex) {
          let valueToSet = caseData[key]; // デフォルトはそのままの値

          // 値の型変換
          if (key === "bug" || key === "needInfo") {
            valueToSet = caseData[key] ? 1 : 0; // Boolean to 1/0
          } else if (key === "sentimentScore") {
            // 数値に変換。無効な場合は null or 空文字列？ 空文字列にする
            const score = parseFloat(valueToSet);
            valueToSet = !isNaN(score) ? score : "";
          } else if (
            key.includes("Date") &&
            typeof valueToSet === "string" &&
            /^\d{4}\/\d{2}\/\d{2}$/.test(valueToSet)
          ) {
            valueToSet = new Date(valueToSet); // String to Date
          } else if (
            key.includes("Time") &&
            typeof valueToSet === "string" &&
            /^\d{2}:\d{2}:\d{2}$/.test(valueToSet)
          ) {
            valueToSet = new Date(`1899/12/30 ${valueToSet}`); // String to Date (with base date)
          } else if (valueToSet === null || valueToSet === undefined) {
            valueToSet = ""; // null/undefined は空文字列に
          }
          // 文字列はそのまま valueToSet を使う

          updates[colIndex] = valueToSet;
          hasValidUpdate = true;
        } else {
          Logger.log(
            `警告: 更新対象のヘッダー "${headerName}" (キー: ${key}) がシートに見つかりません。このフィールドの更新はスキップされます。`
          );
        }
      }
    }

    if (!hasValidUpdate) {
      Logger.log(
        `ケース更新: 行 ${rowIndex}, 更新対象の有効なデータがありませんでした。`
      );
      // 更新データがない場合は成功として扱うこともできるが、ここではエラーとする
      return {
        success: false,
        message: "更新する有効なデータが指定されていません。",
      };
    }

    // Case Status が 'Solution Offered' または 'Finished' に変更された場合の自動日時設定
    const statusHeader = "Case Status";
    const firstCloseDateHeader = "1st Close Date";
    const firstCloseTimeHeader = "1st Close Time";
    const reopenCloseDateHeader = "Reopen Close Date";
    const reopenCloseTimeHeader = "Reopen Close Time";

    const statusColIndex = headerMap[statusHeader];
    const firstCloseDateCol = headerMap[firstCloseDateHeader];
    const firstCloseTimeCol = headerMap[firstCloseTimeHeader];
    const reopenCloseDateCol = headerMap[reopenCloseDateHeader];
    const reopenCloseTimeCol = headerMap[reopenCloseTimeHeader];

    // 更新データに Case Status が含まれ、かつそれがクローズ系ステータスの場合
    const newStatus = updates[statusColIndex]; // 更新後のステータス
    if (
      statusColIndex &&
      newStatus &&
      (newStatus === "Solution Offered" || newStatus === "Finished")
    ) {
      // 現在の1st Close Dateの値を取得して、初回クローズか再クローズかを判断
      let firstCloseDateValue = null;
      if (firstCloseDateCol) {
        try {
          // getValue() で Date オブジェクトを取得
          firstCloseDateValue = sheet
            .getRange(rowIndex, firstCloseDateCol)
            .getValue();
          // 有効な日付かチェック
          if (
            !(firstCloseDateValue instanceof Date) ||
            isNaN(firstCloseDateValue.getTime())
          ) {
            firstCloseDateValue = null;
          }
        } catch (e) {
          Logger.log(
            `警告: 行 ${rowIndex} の ${firstCloseDateHeader} 読み取り失敗: ${e}`
          );
        }
      }

      const now = new Date();
      const scriptTz = Session.getScriptTimeZone();
      const nowDate = new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate()
      );
      const nowTime = new Date(
        1899,
        11,
        30,
        now.getHours(),
        now.getMinutes(),
        now.getSeconds()
      );

      // 初回クローズの場合 (1st Close Date が空欄 または 無効な日付)
      if (!firstCloseDateValue && firstCloseDateCol && firstCloseTimeCol) {
        // ユーザーが明示的に日時を指定していない場合のみ自動設定
        if (!updates.hasOwnProperty(firstCloseDateCol)) {
          // caseDataではなくupdatesでチェック
          updates[firstCloseDateCol] = nowDate;
          Logger.log(`初回クローズ日時 (Date) を自動設定: 行 ${rowIndex}`);
        }
        if (!updates.hasOwnProperty(firstCloseTimeCol)) {
          updates[firstCloseTimeCol] = nowTime;
          Logger.log(`初回クローズ日時 (Time) を自動設定: 行 ${rowIndex}`);
        }
      }
      // 再クローズの場合 (1st Close Date が入力済み)
      else if (reopenCloseDateCol && reopenCloseTimeCol) {
        // ユーザーが明示的に日時を指定していない場合のみ自動設定
        if (!updates.hasOwnProperty(reopenCloseDateCol)) {
          updates[reopenCloseDateCol] = nowDate;
          Logger.log(
            `再オープンクローズ日時 (Date) を自動設定: 行 ${rowIndex}`
          );
        }
        if (!updates.hasOwnProperty(reopenCloseTimeCol)) {
          updates[reopenCloseTimeCol] = nowTime;
          Logger.log(
            `再オープンクローズ日時 (Time) を自動設定: 行 ${rowIndex}`
          );
        }
      }
    }

    // 実際のセル更新 (setValue はコストが高いので、更新が必要なセルのみ行う)
    let updatedCount = 0;
    for (const colIndexStr in updates) {
      const colIndex = parseInt(colIndexStr, 10);
      const value = updates[colIndex];
      try {
        // パフォーマンスのため、setValueをループ内で呼ぶ代わりにRangeListを使うことも検討
        // しかし、個々のセル更新のエラーハンドリングが難しくなるため、ここでは個別にsetValue
        // 注意: 現在の値と比較して変更がない場合はスキップする、という最適化は
        //       getValue() のコストや比較の複雑さを考えると、必ずしも効果的ではない場合がある。
        //       ここではシンプルに常に setValue を呼ぶ。
        sheet.getRange(rowIndex, colIndex).setValue(value);
        updatedCount++;
      } catch (cellError) {
        // 特定のセルの更新に失敗しても、他のセルの更新は試みる
        Logger.log(
          `警告: 行 ${rowIndex}, 列 ${colIndex} (ヘッダー: ${
            headers[colIndex - 1]
          }) の更新に失敗しました: ${cellError}`
        );
        // エラーが発生した場合、部分的な成功として処理を続けるか、全体を失敗とするか検討が必要
        // ここでは処理を続け、最終的にエラーメッセージを返す可能性がある
      }
    }

    if (updatedCount > 0) {
      Logger.log(`ケース更新成功: 行 ${rowIndex}, ${updatedCount}項目を更新。`);
      // 必要であればヘッダーキャッシュをクリア (更新内容がヘッダーに影響する場合)
      // ここではクリアしない
      return { success: true, message: "ケースが正常に更新されました。" };
    } else {
      // このパスには通常到達しないはず (hasValidUpdate でチェック済みのため)
      Logger.log(
        `ケース更新: 行 ${rowIndex}, 実際に更新された項目はありませんでした。`
      );
      return {
        success: false,
        message: "更新するデータが指定されていないか、更新に失敗しました。",
      };
    }
  } catch (error) {
    Logger.log(
      `ケース更新エラー: 行 ${rowIndex}, データ: ${JSON.stringify(
        caseData
      )}, エラー: ${error} - ${error.stack}`
    );
    return {
      success: false,
      message: `ケースの更新中に予期せぬエラーが発生しました: ${error.message}`,
    };
  }
}
