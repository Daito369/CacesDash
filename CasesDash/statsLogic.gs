/**
 * @fileoverview statsLogic.gs - 統計処理モジュール
 * NCC (Non-Contact Complete) および Sentiment Score 統計データの集計などを担当します。
 * @requires settings.gs
 * @requires Code.gs
 */

// --- 定数 ---
const SENTIMENT_TARGET_SCORE = 80; // Sentiment Score の目標値

// --- NCC 統計 ---

/**
 * NCC (Non-Contact Complete) の統計データを、指定された期間ごとに集計して取得します。
 * NCCの定義は CasesDash.md に基づきます:
 *   - Case ID が空欄ではない
 *   - Case Status が 'Assigned' 以外
 *   - non NCC が空欄
 *   - Bug がチェックされていない (値が 0)
 * 上記条件を満たす場合に、Final Assignee の NCC としてカウントされます。
 *
 * 現在の実装では、スクリプト実行ユーザー (Final Assignee) の統計のみを返します。
 *
 * フロントエンド (例: statsUI.html) から呼び出されることを想定しています。
 *
 * @function getNCCStats
 * @param {string} [period] - 集計期間を指定します。'daily', 'weekly', 'monthly', 'quarterly' のいずれか。
 *                            省略された場合は、全ての期間の統計を返します。
 * @returns {{success: boolean, stats?: (Array<Object>|Object<string, Array<Object>>), message?: string}}
 *           成功時:
 *             - period 指定あり: { success: true, stats: Array<Object> }
 *               - stats 配列の各要素: { period: string, totalCases: number, nccCount: number, statusBreakdown: Object<string, number> }
 *             - period 指定なし: { success: true, stats: Object<string, Array<Object>> }
 *               - stats オブジェクトのキー: 'daily', 'weekly', 'monthly', 'quarterly'
 *               - 各キーの値は period 指定ありの場合と同じ配列構造
 *           失敗時は { success: false, message: string } を返します。
 */
function getNCCStats(period) {
  const access = getSpreadsheetAccess(); // settings.gs の関数を呼び出す
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // ヘッダー行を取得 (HEADER_ROW は settings.gs で定義)
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0].map((h) => String(h).trim());
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i; // 0-based index
    });

    // --- NCC計算に必要なヘッダー名と、対応する列インデックスを取得 ---
    const requiredHeaders = {
      date: "Date", // C列想定
      caseStatus: "Case Status", // U列想定
      finalAssignee: "Final Assignee", // S列想定
      caseId: "Case ID", // B列想定
      nonNCC: "non NCC", // W列想定
      bug: "Bug", // X列想定
    };
    const colIndices = {};
    const missingHeaders = [];
    for (const key in requiredHeaders) {
      const headerName = requiredHeaders[key];
      if (headerName in headerMap) {
        colIndices[key] = headerMap[headerName];
      } else {
        missingHeaders.push(headerName);
      }
    }

    if (missingHeaders.length > 0) {
      return {
        success: false,
        message: `統計に必要なヘッダーが見つかりません: ${missingHeaders.join(
          ", "
        )}。スプレッドシートの ${HEADER_ROW}行目を確認してください。`,
      };
    }
    // --- ヘッダー取得ここまで ---

    // データ範囲を取得 (HEADER_ROW + 1 行目以降)
    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      Logger.log("統計計算: データ行が存在しません。");
      // データがない場合も成功として空の統計を返す
      const emptyStats = { daily: [], weekly: [], monthly: [], quarterly: [] };
      return { success: true, stats: period ? emptyStats[period] : emptyStats };
    }
    // 全列取得 (日付処理のため getValues() を使用)
    const dataRange = sheet.getRange(
      HEADER_ROW + 1,
      1,
      lastRow - HEADER_ROW,
      headers.length // ヘッダーの数に合わせる
    );
    const values = dataRange.getValues();

    // 現在のユーザー情報を取得 (Code.gs の関数を呼び出す)
    const userInfo = getCurrentUser();
    if (!userInfo.success) {
      return { success: false, message: userInfo.message };
    }
    const currentUserLdap = userInfo.ldap;

    // 統計データを格納するオブジェクト (ユーザーLDAPをキーとする)
    const userStats = {}; // { ldap: { daily: {}, weekly: {}, monthly: {}, quarterly: {} } }
    // ステータス別カウントを格納するオブジェクト (ユーザーLDAPをキーとする)
    const userStatusCounts = {}; // { ldap: { daily: {}, weekly: {}, monthly: {}, quarterly: {} } }

    // 各データ行を処理して統計を集計
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const currentRowNum = i + HEADER_ROW + 1; // ログ表示用

      // Case IDが空の行はスキップ
      if (!row[colIndices.caseId]) continue;

      // NCC計算の対象となるユーザー (Final Assignee) を取得
      const finalAssignee = String(row[colIndices.finalAssignee] || "").trim();
      if (!finalAssignee) continue; // Final Assignee がいない場合はNCCカウント対象外

      // 日付、ケースステータス、nonNCC、Bugフラグを取得
      const dateValue = row[colIndices.date];
      const caseStatus = String(row[colIndices.caseStatus] || "").trim();
      const nonNCCValue = String(row[colIndices.nonNCC] || "").trim();
      // Bug列の値が '1' または true (チェックボックス) の場合に true とする
      const isBug = row[colIndices.bug] == 1 || row[colIndices.bug] === true;

      // --- NCC 条件判定 (CasesDash.md に基づく) ---
      const isNCC =
        row[colIndices.caseId] && // Case ID が空欄ではない
        caseStatus &&
        caseStatus !== "Assigned" && // Case Status が Assigned 以外
        !nonNCCValue && // non NCC が空欄
        !isBug; // Bug がチェックされていない
      // --- NCC 条件判定ここまで ---

      // 日付が無効な場合は警告ログを出力してスキップ
      if (
        !dateValue ||
        !(dateValue instanceof Date) ||
        isNaN(dateValue.getTime())
      ) {
        Logger.log(
          `警告: 行 ${currentRowNum} の日付が無効です (${dateValue})。統計から除外します。`
        );
        continue;
      }
      const date = dateValue;

      // 期間キーを生成
      const dayStr = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );
      const weekStr = getWeekNumber_(date); // このファイル内の関数を呼び出す
      const monthStr = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "yyyy-MM"
      );
      const quarterStr = getQuarter_(date); // このファイル内の関数を呼び出す

      // ユーザーごとの統計オブジェクトを初期化 (初回出現時)
      if (!userStats[finalAssignee]) {
        userStats[finalAssignee] = {
          daily: {},
          weekly: {},
          monthly: {},
          quarterly: {},
        };
        userStatusCounts[finalAssignee] = {
          daily: {},
          weekly: {},
          monthly: {},
          quarterly: {},
        };
      }

      // 各期間の統計を更新
      const periods = [
        {
          key: dayStr,
          statsObj: userStats[finalAssignee].daily,
          statusObj: userStatusCounts[finalAssignee].daily,
        },
        {
          key: weekStr,
          statsObj: userStats[finalAssignee].weekly,
          statusObj: userStatusCounts[finalAssignee].weekly,
        },
        {
          key: monthStr,
          statsObj: userStats[finalAssignee].monthly,
          statusObj: userStatusCounts[finalAssignee].monthly,
        },
        {
          key: quarterStr,
          statsObj: userStats[finalAssignee].quarterly,
          statusObj: userStatusCounts[finalAssignee].quarterly,
        },
      ];

      periods.forEach((p) => {
        // 期間オブジェクトの初期化 (初回出現時)
        if (!p.statsObj[p.key]) {
          p.statsObj[p.key] = { totalCases: 0, nccCount: 0 };
        }
        if (!p.statusObj[p.key]) {
          p.statusObj[p.key] = {};
        }

        // カウントを更新
        p.statsObj[p.key].totalCases++; // その期間に該当ユーザーがFinal Assigneeだったケース総数
        if (isNCC) {
          p.statsObj[p.key].nccCount++; // NCC条件を満たすケース数
        }

        // ステータス別カウント (Final Assignee基準)
        if (caseStatus) {
          p.statusObj[p.key][caseStatus] =
            (p.statusObj[p.key][caseStatus] || 0) + 1;
        }
      });
    } // End of row processing loop

    // 結果を整形 (実行ユーザーのデータのみ抽出)
    const currentUserResult = {
      daily: [],
      weekly: [],
      monthly: [],
      quarterly: [],
    };
    if (userStats[currentUserLdap]) {
      for (const p in userStats[currentUserLdap]) {
        // p は 'daily', 'weekly', ...
        currentUserResult[p] = Object.keys(userStats[currentUserLdap][p])
          .sort() // 期間キー (日付、週番号など) でソート
          .map((key) => ({
            // key は '2023-01-01', '2023-W01' など
            period: key,
            totalCases: userStats[currentUserLdap][p][key].totalCases,
            nccCount: userStats[currentUserLdap][p][key].nccCount,
            // NCC Rate はここでは計算しない (Daily Average 7 が目標のため)
            statusBreakdown: userStatusCounts[currentUserLdap][p][key] || {},
          }));
      }
    } else {
      Logger.log(
        `統計計算: ユーザー ${currentUserLdap} のNCCデータが見つかりませんでした。`
      );
    }

    Logger.log(`NCC統計データ取得成功: ユーザー ${currentUserLdap}`);
    // 要求された期間、または全期間のデータを返す
    return {
      success: true,
      stats: period ? currentUserResult[period] || [] : currentUserResult,
    };
  } catch (error) {
    Logger.log(`NCC統計データ取得エラー: ${error} - ${error.stack}`);
    return {
      success: false,
      message: `NCC統計データの取得中に予期せぬエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * 日付からISO 8601形式の週番号を取得するヘルパー関数。
 * @param {Date} date - 対象の日付オブジェクト
 * @return {string} 年と週番号の文字列 (例: '2023-W43')
 */
function getWeekNumber(date) {
  const d = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  );
  // Set to nearest Thursday: current date + 4 - current day number
  // Make Sunday's day number 7
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  // Get first day of year
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  // Calculate full weeks to nearest Thursday
  const weekNo = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  // Return array of year and week number
  return d.getUTCFullYear() + "-W" + String(weekNo).padStart(2, "0");
}

/**
 * 日付から四半期を取得するヘルパー関数。
 * @param {Date} date - 対象の日付オブジェクト
 * @return {string} 年と四半期の文字列 (例: '2023-Q4')
 */
function getQuarter(date) {
  const month = date.getMonth(); // 0-11
  const quarter = Math.floor(month / 3) + 1;
  return date.getFullYear() + "-Q" + quarter;
}
// --- Sentiment Score 統計 ---

/**
 * Sentiment Score の統計データを、指定された期間ごとに集計して取得します。
 * スクリプト実行ユーザー (Final Assignee) の統計のみを返します。
 *
 * フロントエンド (例: statsUI.html) から呼び出されることを想定しています。
 *
 * @function getSentimentStats
 * @param {string} [period] - 集計期間を指定します。'daily', 'weekly', 'monthly', 'quarterly' のいずれか。
 *                            省略された場合は、全ての期間の統計を返します。
 * @returns {{success: boolean, stats?: (Array<Object>|Object<string, Array<Object>>), message?: string}}
 *           成功時:
 *             - period 指定あり: { success: true, stats: Array<Object> }
 *               - stats 配列の各要素: { period: string, averageScore: number, targetAchievedRate: number, scoreCount: number }
 *             - period 指定なし: { success: true, stats: Object<string, Array<Object>> }
 *               - stats オブジェクトのキー: 'daily', 'weekly', 'monthly', 'quarterly'
 *               - 各キーの値は period 指定ありの場合と同じ配列構造
 *           失敗時は { success: false, message: string } を返します。
 */
function getSentimentStats(period) {
  const access = getSpreadsheetAccess(); // settings.gs の関数を呼び出す
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    // ヘッダー行を取得 (HEADER_ROW は settings.gs で定義)
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0].map((h) => String(h).trim());
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i; // 0-based index
    });

    // --- Sentiment Score 計算に必要なヘッダー名と、対応する列インデックスを取得 ---
    const requiredHeaders = {
      date: "Date", // C列想定
      finalAssignee: "Final Assignee", // S列想定
      sentimentScore: "Sentiment Score", // 列名は要確認
    };
    const colIndices = {};
    const missingHeaders = [];
    for (const key in requiredHeaders) {
      const headerName = requiredHeaders[key];
      if (headerName in headerMap) {
        colIndices[key] = headerMap[headerName];
      } else {
        missingHeaders.push(headerName);
      }
    }

    if (missingHeaders.length > 0) {
      return {
        success: false,
        message: `Sentiment Score 統計に必要なヘッダーが見つかりません: ${missingHeaders.join(
          ", "
        )}。スプレッドシートの ${HEADER_ROW}行目を確認してください。`,
      };
    }
    // --- ヘッダー取得ここまで ---

    // データ範囲を取得 (HEADER_ROW + 1 行目以降)
    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      Logger.log("Sentiment Score 統計計算: データ行が存在しません。");
      const emptyStats = { daily: [], weekly: [], monthly: [], quarterly: [] };
      return { success: true, stats: period ? emptyStats[period] : emptyStats };
    }
    // 全列取得 (日付処理のため getValues() を使用)
    const dataRange = sheet.getRange(
      HEADER_ROW + 1,
      1,
      lastRow - HEADER_ROW,
      headers.length
    );
    const values = dataRange.getValues();

    // 現在のユーザー情報を取得
    const userInfo = getCurrentUser();
    if (!userInfo.success) {
      return { success: false, message: userInfo.message };
    }
    const currentUserLdap = userInfo.ldap;

    // 統計データを格納するオブジェクト (ユーザーLDAPをキーとする)
    // { ldap: { daily: { '2023-01-01': { scoreSum: 170, count: 2, targetAchievedCount: 1 } }, weekly: {...} } }
    const userSentimentStats = {};

    // 各データ行を処理して統計を集計
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const currentRowNum = i + HEADER_ROW + 1;

      // Final Assignee が現在のユーザーと一致するか確認
      const finalAssignee = String(row[colIndices.finalAssignee] || "").trim();
      if (finalAssignee !== currentUserLdap) continue;

      // Sentiment Score を取得し、数値に変換
      const scoreValue = row[colIndices.sentimentScore];
      const score = parseFloat(scoreValue);

      // 有効なスコア (数値) がある場合のみ処理
      if (!isNaN(score)) {
        // 日付を取得
        const dateValue = row[colIndices.date];
        if (
          !dateValue ||
          !(dateValue instanceof Date) ||
          isNaN(dateValue.getTime())
        ) {
          Logger.log(
            `警告: 行 ${currentRowNum} の日付が無効です (${dateValue})。Sentiment Score 統計から除外します。`
          );
          continue;
        }
        const date = dateValue;

        // 期間キーを生成
        const dayStr = Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
        const weekStr = getWeekNumber_(date);
        const monthStr = Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          "yyyy-MM"
        );
        const quarterStr = getQuarter_(date);

        // ユーザー統計オブジェクト初期化
        if (!userSentimentStats[currentUserLdap]) {
          userSentimentStats[currentUserLdap] = {
            daily: {},
            weekly: {},
            monthly: {},
            quarterly: {},
          };
        }

        // 各期間の統計を更新
        const periods = [
          { key: dayStr, statsObj: userSentimentStats[currentUserLdap].daily },
          {
            key: weekStr,
            statsObj: userSentimentStats[currentUserLdap].weekly,
          },
          {
            key: monthStr,
            statsObj: userSentimentStats[currentUserLdap].monthly,
          },
          {
            key: quarterStr,
            statsObj: userSentimentStats[currentUserLdap].quarterly,
          },
        ];

        periods.forEach((p) => {
          // 期間オブジェクト初期化
          if (!p.statsObj[p.key]) {
            p.statsObj[p.key] = {
              scoreSum: 0,
              count: 0,
              targetAchievedCount: 0,
            };
          }

          // 集計値を更新
          p.statsObj[p.key].scoreSum += score;
          p.statsObj[p.key].count++;
          if (score >= SENTIMENT_TARGET_SCORE) {
            p.statsObj[p.key].targetAchievedCount++;
          }
        });
      }
    } // End of row processing loop

    // 結果を整形
    const currentUserResult = {
      daily: [],
      weekly: [],
      monthly: [],
      quarterly: [],
    };
    if (userSentimentStats[currentUserLdap]) {
      for (const p in userSentimentStats[currentUserLdap]) {
        // p は 'daily', 'weekly', ...
        currentUserResult[p] = Object.keys(
          userSentimentStats[currentUserLdap][p]
        )
          .sort() // 期間キーでソート
          .map((key) => {
            const stats = userSentimentStats[currentUserLdap][p][key];
            const count = stats.count;
            const averageScore = count > 0 ? stats.scoreSum / count : 0;
            const targetAchievedRate =
              count > 0 ? (stats.targetAchievedCount / count) * 100 : 0;
            return {
              period: key,
              averageScore: parseFloat(averageScore.toFixed(1)), // 小数点以下1桁
              targetAchievedRate: parseFloat(targetAchievedRate.toFixed(1)), // 小数点以下1桁
              scoreCount: count, // スコアが入力された件数
            };
          });
      }
    } else {
      Logger.log(
        `Sentiment Score 統計計算: ユーザー ${currentUserLdap} のデータが見つかりませんでした。`
      );
    }

    Logger.log(
      `Sentiment Score 統計データ取得成功: ユーザー ${currentUserLdap}`
    );
    // 要求された期間、または全期間のデータを返す
    return {
      success: true,
      stats: period ? currentUserResult[period] || [] : currentUserResult,
    };
  } catch (error) {
    Logger.log(
      `Sentiment Score 統計データ取得エラー: ${error} - ${error.stack}`
    );
    return {
      success: false,
      message: `Sentiment Score 統計データの取得中に予期せぬエラーが発生しました: ${error.message}`,
    };
  }
}

// --- ヘルパー関数 ---

/**
 * 日付オブジェクトからISO 8601形式の週番号文字列を取得します。
 * 例: 2023年1月1日 -> '2023-W52' (ISO 8601では週は月曜始まり)
 *
 * @private
 * @function getWeekNumber_
 * @param {Date} date - 対象の日付オブジェクト。
 * @returns {string} 年と週番号の文字列 (例: '2023-W43')。
 */
function getWeekNumber_(date) {
  // Dateオブジェクトのコピーを作成して元のオブジェクトを変更しないようにする
  const d = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  );
  // ISO 8601 週番号計算:
  // 週の最初の日 (月曜日) を基準にするため、曜日番号を調整 (日曜日=7)
  const dayNum = d.getUTCDay() || 7;
  // その年の1月4日を含む週を第1週とするため、日付を調整
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  // 年の最初の日 (1月1日) を取得
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  // 年の開始日から調整後の日付までの日数を計算し、7で割って週番号を算出
  const weekNo = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  // 年と週番号を 'YYYY-Www' 形式で返す (週番号は2桁ゼロ埋め)
  return d.getUTCFullYear() + "-W" + String(weekNo).padStart(2, "0");
}

/**
 * 日付オブジェクトから四半期文字列を取得します。
 * 例: 2023年10月1日 -> '2023-Q4'
 *
 * @private
 * @function getQuarter_
 * @param {Date} date - 対象の日付オブジェクト。
 * @returns {string} 年と四半期の文字列 (例: '2023-Q4')。
 */
function getQuarter_(date) {
  const month = date.getMonth(); // 0 (1月) から 11 (12月)
  const quarter = Math.floor(month / 3) + 1; // 0-2 -> Q1, 3-5 -> Q2, 6-8 -> Q3, 9-11 -> Q4
  return date.getFullYear() + "-Q" + quarter;
}

/**
 * カスタムレポートを生成します。
 * ユーザーが指定した期間と条件に基づき、ケースデータを集計して返します。
 *
 * @function generateCustomReport
 * @param {string} startDate - レポート開始日 (ISO 8601形式の文字列)
 * @param {string} endDate - レポート終了日 (ISO 8601形式の文字列)
 * @param {Object} filters - フィルター条件のオブジェクト (例: { segment: 'Gold', status: 'Assigned' })
 * @returns {{success: boolean, report?: Array<Object>, message?: string}}
 *           成功時は集計結果の配列を返し、失敗時はエラーメッセージを返します。
 */
function generateCustomReport(startDate, endDate, filters) {
  const access = getSpreadsheetAccess();
  if (!access.success) {
    return { success: false, message: access.message };
  }
  const sheet = access.sheet;

  try {
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0].map((h) => String(h).trim());
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h) headerMap[h] = i;
    });

    const lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW + 1) {
      return { success: true, report: [] };
    }
    const dataRange = sheet.getRange(
      HEADER_ROW + 1,
      1,
      lastRow - HEADER_ROW,
      headers.length
    );
    const values = dataRange.getValues();

    const start = new Date(startDate);
    const end = new Date(endDate);
    if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) {
      return { success: false, message: "無効な期間指定です。" };
    }

    const report = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const dateVal = row[headerMap["Date"]];
      if (!(dateVal instanceof Date) || isNaN(dateVal.getTime())) continue;

      if (dateVal < start || dateVal > end) continue;

      let match = true;
      for (const key in filters) {
        if (!filters[key]) continue;
        const colIndex = headerMap[key];
        if (colIndex === undefined) {
          match = false;
          break;
        }
        const cellValue = String(row[colIndex]).trim();
        if (cellValue !== filters[key]) {
          match = false;
          break;
        }
      }
      if (!match) continue;

      const record = {};
      headers.forEach((h, idx) => {
        record[h] = row[idx];
      });
      report.push(record);
    }

    return { success: true, report: report };
  } catch (error) {
    Logger.log(`カスタムレポート生成エラー: ${error}`);
    return {
      success: false,
      message: `レポート生成中にエラーが発生しました: ${error.message}`,
    };
  }
}
