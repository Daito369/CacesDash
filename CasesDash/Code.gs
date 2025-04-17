/**
 * @fileoverview Code.gs - メインスクリプトファイル (改訂版)
 * CasesDash Web アプリケーションのエントリーポイント、HTMLテンプレート処理、
 * 共通の定数定義、およびユーザー情報取得を担当します。
 * 他のモジュール (settings.gs, dataAccess.gs, statsLogic.gs) と連携して動作します。
 */

/**
 * フォームで使用される選択肢の定義。
 * フロントエンドとバックエンドで共有されます。
 * @const {Object<string, Array<string>>} FORM_DEFINITIONS
 */
const FORM_DEFINITIONS = {
  segment: [
    "Gold",
    "Platinum",
    "Titanium",
    "Silver",
    "Bronze - Low",
    "Bronze - High",
  ],
  productCategory: [
    "Search",
    "Display",
    "Video",
    "Commerce",
    "Apps",
    "M&A",
    "Policy",
    "Billing",
    "Other",
  ],
  issueCategory: [
    // このリストは非常に長いため、将来的にはスプレッドシート等から動的に取得することも検討
    "CBT invo-invo",
    "CBT invo-auto",
    "CBT (self to self)",
    "LC creation",
    "PP link",
    "PP update",
    "IDT/ Bmod",
    "LCS billing policy",
    "self serve issue",
    "Unidentified Charge",
    "CBT Flow",
    "GQ",
    "OOS",
    "Bulk CBT",
    "CBT ext request",
    "MMS billing policy",
    "Promotion code",
    "Refund",
    "Review",
    "TM form",
    "Trademarks issue",
    "Under Review",
    "Certificate",
    "Suspend",
    "AIV",
    "Complaint",
  ],
  caseStatus: ["Assigned", "Solution Offered", "Finished"],
  amTransfer: [
    "Request to AM contact",
    "Optimize request",
    "β product inquiry",
    "Trouble shooting scope but we don't have access to the resource",
    "Tag team request (LCS customer)",
    "Data analysis",
    "Allowlist request",
    "Other",
  ],
  nonNCC: [
    "Duplicate",
    "Discard",
    "Transfer to Platinum",
    "Transfer to S/B",
    "Transfer to TDCX",
    "Transfer to 3PO",
    "Transfer to OT",
    "Transfer to EN Team",
    "Transfer to GMB Team",
    "Transfer to Other Team (not AM)",
  ],
  reopenReason: [
    "Partial Answer",
    "Recognition issue",
    "Related question",
    "Prefer Phone",
    "NI reopen",
    "Recurrent question",
  ],
};

/**
 * WebアプリケーションへのGETリクエストを処理するメイン関数。
 * URLパラメータ `view=live` が存在する場合は Live Mode 用のUI (`live.html`) を、
 * それ以外の場合はメインのダッシュボードUI (`index.html`) を返します。
 *
 * @param {GoogleAppsScript.Events.DoGet} e - Apps Script から渡されるイベントオブジェクト。URLパラメータを含む。
 * @returns {GoogleAppsScript.HTML.HtmlOutput} 評価済みのHTMLテンプレート。
 */
function doGet(e) {
  let template;
  let title;

  // Live Mode のパラメータをチェック
  if (e && e.parameter && e.parameter.view === "live") {
    template = HtmlService.createTemplateFromFile("live");
    title = "新規ケース登録 (Live Mode)";
    Logger.log("Live Mode UI を表示します。");
  } else {
    template = HtmlService.createTemplateFromFile("index");
    title = "CasesDash - 3PO Chat ケース管理"; // アプリケーション名を反映
    Logger.log("メインダッシュボード UI を表示します。");
  }

  // テンプレートを評価し、タイトルとメタタグを設定して返す
  return template
    .evaluate()
    .setTitle(title)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME) // セキュリティのため IFRAME モードを推奨
    .addMetaTag(
      // レスポンシブデザインのための viewport 設定
      "viewport",
      "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"
    );
}

/**
 * HTMLテンプレート内で `<!= include('filename'); >` 構文を使用して
 * 他のHTMLファイルの内容をインクルードするためのヘルパー関数。
 *
 * @param {string} filename - インクルードするHTMLファイル名 (拡張子 `.html` は不要)。
 * @returns {string} 指定されたHTMLファイルの内容。
 */
function include(filename) {
  // createHtmlOutputFromFile は HtmlOutput を返すため、getContent() で文字列を取得
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * フォームで使用する選択肢の定義 (`FORM_DEFINITIONS`) を取得します。
 * フロントエンド (例: script.html) から `google.script.run` 経由で呼び出され、
 * ドロップダウンリストなどを動的に生成するために使用されます。
 *
 * @returns {Object<string, Array<string>>} フォーム定義オブジェクト。
 */
function getFormDefinitions() {
  // 定数をそのまま返す
  return FORM_DEFINITIONS;
}

/**
 * Webアプリケーションの現在のデプロイメントURLを取得します。
 * 主に Live Mode のポップアップウィンドウを開くためにフロントエンドから使用されます。
 *
 * @returns {{success: boolean, url?: string, message?: string}}
 *           成功時は { success: true, url: string }、
 *           失敗時は { success: false, message: string } を返します。
 */
function getService() {
  try {
    // getService() は非推奨ではないが、getUrl() の方が一般的
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      throw new Error(
        "デプロイメントURLを取得できませんでした。スクリプトがWebアプリとしてデプロイされているか確認してください。"
      );
    }
    Logger.log(`サービスURL取得成功: ${url}`);
    return { success: true, url: url };
  } catch (error) {
    Logger.log(`サービスURL取得エラー: ${error}`);
    return {
      success: false,
      message: `サービスURLの取得中にエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * 現在スクリプトを実行しているユーザーの情報を取得します。
 * Session.getActiveUser().getEmail() を使用してメールアドレスを取得し、
 * '@' より前の部分をLDAPユーザー名として抽出します (Google Corp 環境を想定)。
 *
 * @returns {{success: boolean, email?: string, ldap?: string, message?: string}}
 *           成功時は { success: true, email: string, ldap: string }、
 *           失敗時は { success: false, message: string } を返します。
 *           承認が必要な場合やユーザー情報が取得できない場合に失敗します。
 */
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    // メールアドレスが取得できない場合 (例: 承認前) はエラー
    if (!email) {
      throw new Error(
        "ユーザーのメールアドレスを取得できませんでした。スクリプトの承認が完了しているか確認してください。"
      );
    }

    // '@' を含むかチェックし、含まない場合はそのまま LDAP とする (外部ユーザーなどの可能性)
    let ldap;
    const emailParts = email.split("@");
    if (emailParts.length > 1) {
      ldap = emailParts[0]; // '@' より前を LDAP とする
    } else {
      ldap = email; // '@' がない場合はメールアドレス全体を識別子とする
      Logger.log(
        `ユーザー ${email} は '@' を含まないため、全体をLDAPとして扱います。`
      );
    }

    if (!ldap) {
      // 通常発生しないはずだが念のため
      throw new Error("LDAPユーザー名の抽出に失敗しました。");
    }

    Logger.log(`ユーザー情報取得: email=${email}, ldap=${ldap}`);
    return { success: true, email: email, ldap: ldap };
  } catch (error) {
    Logger.log(`ユーザー情報取得エラー: ${error}`);
    let userMessage = "ユーザー情報の取得に失敗しました。";
    // エラーメッセージに応じてユーザーへのフィードバックを改善
    if (error.message.includes("Authorization is required")) {
      userMessage +=
        "スクリプトの実行承認が必要です。ページを再読み込みして承認プロセスを完了してください。";
    } else if (error.message.includes("メールアドレスを取得できませんでした")) {
      userMessage = error.message; // 独自エラーメッセージをそのまま表示
    } else {
      userMessage += `詳細: ${error.message}`;
    }
    return { success: false, message: userMessage };
  }
}

/**
 * 管理者メールアドレスリストを取得
 * @returns {Array<string>}
 */
function getAdminEmails_() {
  const userProperties = PropertiesService.getUserProperties();
  const adminEmailsStr = userProperties.getProperty("ADMIN_EMAILS") || "";
  return adminEmailsStr
    .split(",")
    .map((email) => email.trim())
    .filter((email) => email.length > 0);
}
