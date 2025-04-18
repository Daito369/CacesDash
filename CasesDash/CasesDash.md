

# CasesDash.md

# CasesDash - マルチシート対応 ケース管理システム

## 1. プロジェクト概要

### 1.1 概要と目的

CasesDashは、Google 広告サポートチームのケース管理を効率化するためのウェブベースのツールです。Google スプレッドシートと連携し、複数シートにわたるケース割り当て、追跡、期限管理、統計分析を一元化します。UI/UX優先設計により、複雑なケース管理作業を直感的に行えるようにしています。

本システムは、単一シート(3PO Chat)対応から、6つの異なるシート(OT Email, 3PO Email, OT Chat, 3PO Chat, OT Phone, 3PO Phone)に対応するようにアップグレードされています。

### 1.2 背景

大規模サポートチームでは、複数のチャネル(Email/Chat/Phone)と種別(OT/3PO)を通じて多数のケースを同時に処理し、SLA（Service Level Agreement）を確実に守る必要があります。既存の管理方法はスプレッドシートベースで、以下の問題がありました：

- 「3PO Chat」シートのみの対応で、他の5つのシートをカバーできていなかった
- シート間でのデータ構造の違いに対応できていなかった
- チャネル別・種別別の統合分析が困難だった
- データの視覚化が不十分で重要情報が見えにくい
- 期限管理が手動で行われており、ミスのリスクが高い
- ケースステータスの更新や検索が非効率
- チーム全体のパフォーマンス分析が困難

CasesDashはこれらの問題を解決し、サポートチームの効率と顧客満足度を向上させます。

### 1.3 主要機能

#### マルチシート連携機能
- 6つの異なるシートとの完全連携（OT Email, 3PO Email, OT Chat, 3PO Chat, OT Phone, 3PO Phone）
- シート固有のデータ構造に対応した動的フォーム生成
- チャネル別・種別別の視覚的区別とフィルタリング

#### ダッシュボード機能
- アクティブケースの一覧表示と視覚的なステータス表示
- シート種別を視認できるバッジとカラーコーディング
- リアルタイムのSLAタイマー（残り応答時間のカウントダウン）
- P95タイマー（Case Open から72時間/7日間のカウントダウン）
- T&S Consulted 状態の切り替え機能
- ケースの詳細表示と編集機能
- ダッシュボードからのケース削除機能

#### ケース管理機能
- シート選択機能付き新規ケース登録フォーム（必須項目検証付き）
- シートタイプに応じた動的フィールド表示
- セグメント、製品カテゴリ、問題カテゴリなどの属性管理
- Case Status、AM Transfer、non NCCなどの状態管理
- 初回クローズおよび再オープン情報の管理
- Live Mode（別ウィンドウでの登録機能）

#### 検索機能
- すべてのシートを対象とした統合検索
- シート指定検索オプション
- Case ID による即時検索
- 検索結果からのケース詳細表示と編集
- 非アクティブケースの再アクティブ化機能

#### 統計分析機能
- シート別およびシート横断の統計表示
- 日別/週別/月別/四半期別の統計表示
- ケース解決率とNCCの視覚化
- グラフによるデータの視覚化（折れ線/円/棒グラフ）
- ステータス別の詳細分析
- チャネル比較分析（Email vs Chat vs Phone）

#### その他の機能
- ダークモード/ライトモード切り替え
- レスポンシブデザイン対応
- スプレッドシート設定と接続テスト機能
- エラーハンドリングとユーザーフィードバック

### 1.4 ターゲットユーザー

- カスタマーサポートの専門家（Email/Chat/Phone対応）
- ケース管理担当者
- チームリーダーとマネージャー
- サポート品質分析官
- カスタマーサクセスチーム

### 1.5 解決する問題

- **効率性**: 複数シートにまたがるケース情報の入力、更新、検索、分析を効率化
- **一元管理**: 異なるチャネル・種別のケースを単一のインターフェースで管理
- **正確性**: 自動化されたタイマーとリマインダーでSLA違反を防止
- **可視性**: チャネル・種別を区別した明確なダッシュボードと統計
- **一貫性**: 標準化されたプロセスでシート間のデータ入力と管理を統一
- **分析力**: チャネル横断・種別横断のリアルタイム統計によるパフォーマンス指標追跡

## 2. 使用技術

### 2.1 フロントエンド

#### コア技術

- **HTML5**: 最新のセマンティックマークアップ
- **CSS3**: カスタム変数、グリッドレイアウト、フレックスボックス
- **JavaScript (ES6+)**: モダンなJavaScript機能とブラウザAPI

#### UI フレームワーク

- **Material Design Components for Web**: Google の公式 Material Design3 実装
  - MDC Tab Bar, Dialog, Select, TextField, Button, Checkbox等
  - Material Icons & Symbols (アイコンフォント)

#### その他のライブラリ

- **Chart.js**: 統計データの視覚化
- **Flatpickr**: 日付と時間の選択コンポーネント
- **Google Fonts**: Google SansとRobotoフォントファミリー

### 2.2 バックエンド

- **Google Apps Script (GAS)**: サーバーサイドロジック
  - HTML Service: UIレンダリング
  - Properties Service: 設定保存
  - Spreadsheet Service: データ操作
  - Session および User Service: ユーザー認証

### 2.3 データストレージ

- **Google Spreadsheets**: プライマリデータストア
  - シート名: "OT Email", "3PO Email", "OT Chat", "3PO Chat", "OT Phone", "3PO Phone"
  - 構造化データ (シート別に列定義された表形式)
  - フォーミュラとArrayFormula
  - データ検証ルール

### 2.4 開発・デプロイツール

- **Google Apps Script Editor**: コード編集とデバッグ
- **GitHub**: コードバージョン管理と共有
- **GAS Deployment**: Webアプリケーションとしてのデプロイ
- **Version Management**: デプロイバージョン管理

## 3. システムアーキテクチャ

### 3.1 全体アーキテクチャ

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│                 │     │                 │     │                 │
│  Web クライアント │ ←→ │  Google Apps     │ ←→ │     Google       │
│  (ブラウザ)      │     │  Script (GAS)   │     │  Spreadsheet    │
│                 │     │                 │     │  (複数シート)    │
└─────────────────┘     └─────────────────┘     └─────────────────┘
     ↑                        ↑                       ↑
     │                        │                       │
     │                        │                       │
     ↓                        ↓                       ↓
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  HTML/CSS/JS    │     │  Server-side JS │     │  Structured     │
│  UI Components  │     │  API Handlers   │     │  Data Storage   │
│  Client Logic   │     │  Data Logic     │     │  & Formulas     │
└─────────────────┘     └─────────────────┘     └─────────────────┘
```

### 3.2 モジュール構成

- **フロントエンドモジュール**
  - CasesDash フォルダー内のファイルを確認して

- **バックエンドモジュール**
  - CasesDash フォルダー内のファイルを確認してここに記載してください。

- **HTMLテンプレート**
  - CasesDash フォルダー内のファイルを確認してここに記載してください。

### 3.3 データモデル

CasesDashは、6つの異なるシートに対応するための柔軟なデータモデルを採用しています。各シートは共通のデータ要素を持ちつつも、シート特有のフィールドや列構造を持っています。

#### シートの構造と特徴

各シートは以下のような特徴を持っています：

1. **OT Email**: メール経由のOT（Outbound Team）ケース
   - 主要列: Case ID, Date, Time, Segment, Product Category, Assignee など
   - 特徴: AM Initiated 列あり

2. **3PO Email**: メール経由の3POケース
   - 主要列: Case ID, Date, Time, Segment, Product Category, Issue Category, Details など
   - 特徴: Issue Category と Details 列あり

3. **OT Chat**: チャット経由のOTケース
   - 主要列: Case ID, Date, Time, Segment, Product Category, Assignee など
   - 特徴: Chat固有の対応時間指標

4. **3PO Chat**: チャット経由の3POケース
   - 主要列: Case ID, Date, Time, Segment, Product Category, Issue Category, Details など
   - 特徴: Issue Category と Details 列あり、Chat固有の対応時間指標

5. **OT Phone**: 電話経由のOTケース
   - 主要列: Case ID, Date, Time, Segment, Product Category, Assignee など
   - 特徴: Phone固有の対応時間指標

6. **3PO Phone**: 電話経由の3POケース
   - 主要列: Case ID, Date, Time, Segment, Product Category, Issue Category, Details など
   - 特徴: Issue Category と Details 列あり、Phone固有の対応時間指標

  #### シート特有のデータ構造と列マッピング

  各シートは共有する情報もありますが、データ構造が異なります。以下は各シートの詳細な列構造とその意味を示しています。

  ##### OT Email シート構造（列マッピング詳細）
  ```
  列A: Date - お問い合わせ日
  列B: Case - ケースへのハイパーリンク（=HYPERLINK("https://cases.connect.corp.google.com/#/case/"&C:C,"Cases")）
  列C: Case ID - 一意のケースID（例: 3-4505000031234）
  列D: Case Open Date - ケースが開始された日付（YYYY/MM/DD形式）
  列E: Time - ケース開始時間（HH:MM:SS形式）
  列F: Incoming Segment - Segment（ Gold, Platinum, Titanium, Silver, Bronze - Low, Bronze - High）
  列G: Product Category - Category（Search, Display, Video, Commerce, Apps, M&A, Policy, Billing, Other）
  列H: Triage - Triage（チェックボックス[0/1]）
  列I: Prefer Either - Either（チェックボックス[0/1]）
  列J: AM Initiated - AM（チェックボックス[0/1]）
  列K: Is 3.0 - 3.0（チェックボックス[0/1]）
  列L: 1st Assignee - 1st Assignee（user@google.com）
  列M: TRT Timer - リアルタイムで残りのSLA時間（=ARRAYFORMULA(IF((U:U="Assigned")*(C:C<>""),IF((AJ:AJ-NOW())>0, TEXT(AJ:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))）
  列N: Aging Timer - エージング時間（P95タイマー）(=ARRAYFORMULA(IF((U3:U="Assigned")*(C3:C<>""),IF((AL3:AL-NOW())>0, TEXT(AL3:AL-NOW(),"[h]:mm:ss"), "Missed"), "")))
  列O: Pool Transfer Destination - プール転送先
  列P: Pool Transfer Reason - 転送理由
  列Q: MCC - マーチャントCenterアカウント情報
  列R: Change to Child - 子アカウントへの変更
  列S: Final Assignee - 最終担当者（LDAP ID）
  列T: Final Segment - 最終セグメント
  列U: Case Status - ケースステータス（Assigned, Waiting, Pending, Closed）
  列V: AM Transfer - AM移管フラグ(Request to AM contact, Optimize request, β product inquiry, Trouble shooting scope but we don't have access to the resource, Tag team request (LCS customer), Data analysis, Allowlist request, Other)
  列W: non NCC - NCC外理由（NCC計算から除外する理由）
  列X: Bug - バグステータス（チェックボックス[0/1]）
  列Y: Need Info - 情報要求フラグ
  列Z: 1st Close Date - 初回クローズ日
  列AA: 1st Close Time - 初回クローズ時間
  列AB: Reopen Reason - 再オープン理由
  列AC: Reopen Close Date - 再クローズ日
  列AD: Reopen Close Time - 再クローズ時間
  列AE: Product Commerce - 商品コマース区分
  列AF: Assign Week - アサインされた週（自動計算）
  列AG: Channel - チャネル（常に"Email"）
  列AH: TRT Target - TRTターゲット時間（セグメントと製品に基づいて自動計算）
  列AI: TRT Date+Time - TRT期限日時（絶対値）
  列AJ: Aging Target - エージングターゲット時間
  列AK: Aging Date+Time - エージング期限日時（絶対値）
  列AL: Close NCC - クローズ時のNCCステータス
  列AM: Close Date - クローズ日
  列AN: Close Time - クローズ時間
  列AO: Close Week - クローズした週
  列AP: TRT - TRTステータス（SLA違反の場合1、そうでない場合0）
  列AQ: Aging - エージングステータス（P95違反の場合1、そうでない場合0）
  列AR: Reopen Close - 再オープン情報（再クローズされた場合1、そうでない場合0）
  列AS: Reassign - 再アサイン情報（再アサインされた場合1、そうでない場合0）
  ```

  ##### 3PO Email シート構造（列マッピング詳細）
  ```
  列A: Date - お問い合わせ日
  列B: Cases - ケースへのハイパーリンク（=HYPERLINK("https://cases.connect.corp.google.com/#/cases/"&C:C,"Cases")）
  列C: Case ID - 一意のケースID（例: 3-4505000031234）
  列D: Case Open Date - ケースが開始された日付（YYYY/MM/DD形式）
  列E: Time - ケース開始時間（HH:MM:SS形式）
  列F: Incoming Segment -Segment（ Gold, Platinum, Titanium, Silver, Bronze - Low, Bronze - High）
  列G: Product Category - Category（Search, Display, Video, Commerce, Apps, M&A, Policy, Billing, Other）
  列H: Triage - トリアージ情報（チェックボックス[0/1]）
  列I: Prefer Either - Either（チェックボックス[0/1]）
  列J: AM Initiated - AM（チェックボックス[0/1]）
  列K: Is 3.0 - 3.0（チェックボックス[0/1]）
  列L: Issue Category - Issue Category（CBT invo-invo, CBT invo-auto, CBT (self to self), LC creation, PP link, PP update, IDT/ Bmod, LCS billing policy, self serve issue, Unidentified Charge, CBT Flow, GQ, OOS, Bulk CBT, CBT ext request, MMS billing policy, Promotion code, Refund, Review, TM form, Trademarks issue, Under Review, Certificate, Suspend, AIV, Complaint）
  列M: Details - 詳細情報（3PO特有）
  列N: 1st Assignee - 1st Assignee（user@google.com）
  列O: TRT Timer - リアルタイムで残りのSLA時間(=ARRAYFORMULA(IF((W3:W="Assigned")*(C3:C<>""),IF((AL3:AL-NOW())>0, TEXT(AL3:AL-NOW(),"[h]:mm:ss"), "Missed"), "")))
  列P: Aging Timer - エージング時間（P95タイマー）(=ARRAYFORMULA(IF((W3:W="Assigned")*(C3:C<>""),IF((AN3:AN-NOW())>0, TEXT(AN3:AN-NOW(),"[h]:mm:ss"), "Missed"), "")))
  列Q: Pool Transfer Destination - プール転送先
  列R: Pool Transfer Reason - 転送理由
  列S: MCC - マーチャントCenterアカウント情報
  列T: Change to Child - 子アカウントへの変更
  列U: Final Assignee - 最終担当者
  列V: Final Segment - 最終セグメント
  列W: Case Status - ケースステータス
  列X: AM Transfer - AM移管フラグ
  列Y: non NCC - NCC外理由
  列Z: Bug - バグステータス
  列AA: Need Info - 情報要求フラグ
  列AB: 1st Close Date - 初回クローズ日
  列AC: 1st Close Time - 初回クローズ時間
  列AD: Reopen Reason - 再オープン理由
  列AE: Reopen Close Date - 再クローズ日
  列AF: Reopen Close Time - 再クローズ時間
  列AG: Product Commerce - 商品コマース区分（例: 2=Billing, 3=Policy）
  列AH: Assign Week - アサインされた週
  列AI: Channel - チャネル（常に"Email"）
  列AJ: TRT Target - TRTターゲット時間
  列AK: TRT Date+Time - TRT期限日時
  列AL: Aging Target - エージングターゲット時間
  列AM: Aging Date+Time - エージング期限日時
  列AN: Close NCC - クローズ時のNCCステータス
  列AO: Close Date - クローズ日
  列AP: Close Time - クローズ時間
  列AQ: Close Week - クローズした週
  列AR: TRT - TRTステータス
  列AS: Aging - エージングステータス
  列AT: Reopen Close - 再オープン情報
  列AU: Reassign - 再アサイン情報
  ```

  ##### OT Chat シート構造（列マッピング詳細）
  ```
  列A: Cases - ケースへのハイパーリンク（=HYPERLINK("https://cases.connect.corp.google.com/#/cases/"&C:C,"Cases")）
  列B: Case ID - 一意のケースID（例: 3-4505000031234）
  列C: Case Open Date - ケースが開始された日付（YYYY/MM/DD形式）
  列D: Time - ケース開始時間（HH:MM:SS形式）
  列E: Incoming Segment - Segment（ Gold, Platinum, Titanium, Silver, Bronze - Low, Bronze - High）
  列F: Product Category - Category（Search, Display, Video, Commerce, Apps, M&A, Policy, Billing, Other）
  列G: Triage - Triage（チェックボックス[0/1]）
  列H: Prefer Either - Either（チェックボックス[0/1]）
  列I: Is 3.0 - 3.0（チェックボックス[0/1]）
  列J: 1st Assignee - 1st Assignee（user@google.com）
  列K: TRT Timer - リアルタイムで残りのSLA時間（=ARRAYFORMULA(IF((S3:S="Assigned")*(B3:B<>""),IF((AH3:AH-NOW())>0, TEXT(AH3:AH-NOW(),"[h]:mm:ss"), "Missed"), ""))）
  列L: Aging Timer - エージング時間
  列M: Pool Transfer Destination - プール転送先
  列N: Pool Transfer Reason - 転送理由
  列O: MCC - マーチャントCenterアカウント情報
  列P: Change to Child - 子アカウントへの変更
  列Q: Final Assignee - 最終担当者
  列R: Final Segment - 最終セグメント
  列S: Case Status - ケースステータス
  列T: AM Transfer - AM移管フラグ
  列U: non NCC - NCC外理由
  列V: Bug - バグステータス
  列W: Need Info - 情報要求フラグ
  列X: 1st Close Date - 初回クローズ日
  列Y: 1st Close Time - 初回クローズ時間
  列Z: Reopen Reason - 再オープン理由
  列AA: Reopen Close Date - 再クローズ日
  列AB: Reopen Close Time - 再クローズ時間
  列AC: Product Commerce - 商品コマース区分
  列AD: Assign Week - アサインされた週
  列AE: Channel - チャネル（常に"Chat"）
  列AF: TRT Target - TRTターゲット時間（チャット用の短縮時間）
  列AG: TRT Date+Time - TRT期限日時
  列AH: Aging Target - エージングターゲット時間
  列AI: Aging Date+Time - エージング期限日時
  列AJ: Close NCC - クローズ時のNCCステータス
  列AK: Close Date - クローズ日
  列AL: Close Time - クローズ時間
  列AM: Close Week - クローズした週
  列AN: TRT - TRTステータス
  列AO: Aging - エージングステータス
  列AP: Reopen Close - 再オープン情報
  列AQ: Reassign - 再アサイン情報
  ```

  ##### 3PO Chat シート構造（列マッピング詳細）
  ```
  列A: Cases - ケースへのハイパーリンク（=HYPERLINK("https://cases.connect.corp.google.com/#/cases/"&C:C,"Cases")）
  列B: Case ID - 一意のケースID（例: 3-4505000031234）
  列C: Case Open Date - ケースが開始された日付（YYYY/MM/DD形式）
  列D: Time - ケース開始時間（HH:MM:SS形式）
  列E: Incoming Segment - Segment（ Gold, Platinum, Titanium, Silver, Bronze - Low, Bronze - High）
  列F: Product Category - Category（Search, Display, Video, Commerce, Apps, M&A, Policy, Billing, Other）
  列G: Triage - Triage（チェックボックス[0/1]）
  列H: Prefer Either - Either（チェックボックス[0/1]）
  列I: Is 3.0 - 3.0（チェックボックス[0/1]）
  列J: Issue Category - 問題カテゴリ（3PO特有）
  列K: Details - 詳細情報（3PO特有）
  列L: 1st Assignee - 1st Assignee（user@google.com）
  列M: TRT Timer - リアルタイムで残りのSLA時間(=ARRAYFORMULA(IF((U3:U="Assigned")*(B3:B<>""),IF((AJ3:AJ-NOW())>0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), "")))
  列N: Aging Timer - エージング時間
  列O: Pool Transfer Destination - プール転送先
  列P: Pool Transfer Reason - 転送理由
  列Q: MCC - マーチャントCenterアカウント情報
  列R: Change to Child - 子アカウントへの変更
  列S: Final Assignee - 最終担当者
  列T: Final Segment - 最終セグメント
  列U: Case Status - ケースステータス
  列V: AM Transfer - AM移管フラグ
  列W: non NCC - NCC外理由
  列X: Bug - バグステータス
  列Y: Need Info - 情報要求フラグ
  列Z: 1st Close Date - 初回クローズ日
  列AA: 1st Close Time - 初回クローズ時間
  列AB: Reopen Reason - 再オープン理由
  列AC: Reopen Close Date - 再クローズ日
  列AD: Reopen Close Time - 再クローズ時間
  列AE: Product Commerce - 商品コマース区分
  列AF: Assign Week - アサインされた週
  列AG: Channel - チャネル（常に"Chat"）
  列AH: TRT Target - TRTターゲット時間
  列AI: TRT Date+Time - TRT期限日時
  列AJ: Aging Target - エージングターゲット時間
  列AK: Aging Date+Time - エージング期限日時
  列AL: Close NCC - クローズ時のNCCステータス
  列AM: Close Date - クローズ日
  列AN: Close Time - クローズ時間
  列AO: Close Week - クローズした週
  列AP: TRT - TRTステータス
  列AQ: Aging - エージングステータス
  列AR: Reopen Close - 再オープン情報
  列AS: Reassign - 再アサイン情報
  ```

  ##### OT Phone シート構造（列マッピング詳細）
  ```
  列A: Cases - ケースへのハイパーリンク（=HYPERLINK("https://cases.connect.corp.google.com/#/cases/"&C:C,"Cases")）
  列B: Case ID - 一意のケースID（例: 3-4505000031234）
  列C: Case Open Date - ケースが開始された日付（YYYY/MM/DD形式）
  列D: Time - ケース開始時間（HH:MM:SS形式）
  列E: Incoming Segment - Segment（ Gold, Platinum, Titanium, Silver, Bronze - Low, Bronze - High）
  列F: Product Category - Category（Search, Display, Video, Commerce, Apps, M&A, Policy, Billing, Other）
  列G: Triage - Triage（チェックボックス[0/1]）
  列H: Prefer Either - Either（チェックボックス[0/1]）
  列I: Is 3.0 - 3.0（チェックボックス[0/1]）
  列J: 1st Assignee - 1st Assignee（user@google.com）
  列K: TRT Timer - リアルタイムで残りのSLA時間（=ARRAYFORMULA(IF((S3:S="Assigned")*(B3:B<>""),IF((AH3:AH-NOW())>0, TEXT(AH3:AH-NOW(),"[h]:mm:ss"), "Missed"), ""))）
  列L: Aging Timer - エージング時間
  列M: Pool Transfer Destination - プール転送先
  列N: Pool Transfer Reason - 転送理由
  列O: MCC - マーチャントCenterアカウント情報
  列P: Change to Child - 子アカウントへの変更
  列Q: Final Assignee - 最終担当者
  列R: Final Segment - 最終セグメント
  列S: Case Status - ケースステータス
  列T: AM Transfer - AM移管フラグ
  列U: non NCC - NCC外理由
  列V: Bug - バグステータス
  列W: Need Info - 情報要求フラグ
  列X: 1st Close Date - 初回クローズ日
  列Y: 1st Close Time - 初回クローズ時間
  列Z: Reopen Reason - 再オープン理由
  列AA: Reopen Close Date - 再クローズ日
  列AB: Reopen Close Time - 再クローズ時間
  列AC: Product Commerce - 商品コマース区分
  列AD: Assign Week - アサインされた週
  列AE: Channel - チャネル（常に"Phone"）
  列AF: TRT Target - TRTターゲット時間
  列AG: TRT Date+Time - TRT期限日時
  列AH: Aging Target - エージングターゲット時間
  列AI: Aging Date+Time - エージング期限日時
  列AJ: Close NCC - クローズ時のNCCステータス
  列AK: Close Date - クローズ日
  列AL: Close Time - クローズ時間
  列AM: Close Week - クローズした週
  列AN: TRT - TRTステータス
  列AO: Aging - エージングステータス
  列AP: Reopen Close - 再オープン情報
  列AQ: Reassign - 再アサイン情報
  ```

  ##### 3PO Phone シート構造（列マッピング詳細）
  ```
  列A: Cases - ケースへのハイパーリンク（=HYPERLINK("https://cases.connect.corp.google.com/#/cases/"&C:C,"Cases")）
  列B: Case ID - 一意のケースID（例: 3-4505000031234）
  列C: Case Open Date - ケースが開始された日付（YYYY/MM/DD形式）
  列D: Time - ケース開始時間（HH:MM:SS形式）
  列E: Incoming Segment - Segment（ Gold, Platinum, Titanium, Silver, Bronze - Low, Bronze - High）
  列F: Product Category - 製品カテゴリ
  列G: Triage - Triage（チェックボックス[0/1]）
  列H: Prefer Either - Either（チェックボックス[0/1]）
  列I: Is 3.0 - 3.0（チェックボックス[0/1]）
  列J: Issue Category - 問題カテゴリ（3PO特有）
  列K: Details - 詳細情報（3PO特有）
  列L: 1st Assignee - 1st Assignee（user@google.com）
  列M: TRT Timer - リアルタイムで残りのSLA時間(=ARRAYFORMULA(IF((U3:U="Assigned")*(B3:B<>""),IF((AJ3:AJ-NOW())>0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), "")))
  列N: Aging Timer - エージング時間
  列O: Pool Transfer Destination - プール転送先
  列P: Pool Transfer Reason - 転送理由
  列Q: MCC - マーチャントCenterアカウント情報
  列R: Change to Child - 子アカウントへの変更
  列S: Final Assignee - 最終担当者
  列T: Final Segment - 最終セグメント
  列U: Case Status - ケースステータス
  列V: AM Transfer - AM移管フラグ
  列W: non NCC - NCC外理由
  列X: Bug - バグステータス
  列Y: Need Info - 情報要求フラグ
  列Z: 1st Close Date - 初回クローズ日
  列AA: 1st Close Time - 初回クローズ時間
  列AB: Reopen Reason - 再オープン理由
  列AC: Reopen Close Date - 再クローズ日
  列AD: Reopen Close Time - 再クローズ時間
  列AE: Product Commerce - 商品コマース区分
  列AF: Assign Week - アサインされた週
  列AG: Channel - チャネル（常に"Phone"）
  列AH: TRT Target - TRTターゲット時間
  列AI: TRT Date+Time - TRT期限日時
  列AJ: Aging Target - エージングターゲット時間
  列AK: Aging Date+Time - エージング期限日時
  列AL: Close NCC - クローズ時のNCCステータス
  列AM: Close Date - クローズ日
  列AN: Close Time - クローズ時間
  列AO: Close Week - クローズした週
  列AP: TRT - TRTステータス
  列AQ: Aging - エージングステータス
  列AR: Reopen Close - 再オープン情報
  列AS: Reassign - 再アサイン情報
  ```

  #### 詳細なマッピングシステムの実装

  シート間の列の差異を効果的に処理するため、システムは詳細なマッピング設定を使用します。以下は完全なマッピングコードです：

  ```javascript
  /**
   * シート別の列マッピング定義
   * 各シートの特定のデータ要素がどの列に対応するかを定義
   */
  const SheetColumnMappings = {
    "OT Email": {
      caseId: "C",
      date: "D",
      time: "E",
      segment: "F",
      productCategory: "G",
      triage: "H",
      preferEither: "I",
      amInitiated: "J",
      is30: "K",
      firstAssignee: "L",
      trtTimer: "M",
      agingTimer: "N",
      poolDestination: "O",
      poolReason: "P",
      mcc: "Q",
      changeToChild: "R",
      finalAssignee: "S",
      finalSegment: "T",
      caseStatus: "U",
      amTransfer: "V",
      nonNCC: "W",
      bug: "X",
      needInfo: "Y",
      firstCloseDate: "Z",
      firstCloseTime: "AA",
      reopenReason: "AB",
      reopenCloseDate: "AC",
      reopenCloseTime: "AD",
      productCommerce: "AE",
      assignWeek: "AF",
      channel: "AG",
      trtTarget: "AH",
      trtDateTime: "AI",
      agingTarget: "AJ",
      agingDateTime: "AK",
      closeNCC: "AL",
      closeDate: "AM",
      closeTime: "AN",
      closeWeek: "AO",
      trt: "AP",
      aging: "AQ",
      reopenClose: "AR",
      reassign: "AS"
    },
    "3PO Email": {
      caseId: "C",
      date: "D",
      time: "E",
      segment: "F",
      productCategory: "G",
      triage: "H",
      preferEither: "I",
      amInitiated: "J",
      is30: "K",
      issueCategory: "L", // 3PO特有
      details: "M", // 3PO特有
      firstAssignee: "N",
      trtTimer: "O",
      agingTimer: "P",
      poolDestination: "Q",
      poolReason: "R",
      mcc: "S",
      changeToChild: "T",
      finalAssignee: "U",
      finalSegment: "V",
      caseStatus: "W",
      amTransfer: "X",
      nonNCC: "Y",
      bug: "Z",
      needInfo: "AA",
      firstCloseDate: "AB",
      firstCloseTime: "AC",
      reopenReason: "AD",
      reopenCloseDate: "AE",
      reopenCloseTime: "AF",
      productCommerce: "AG",
      assignWeek: "AH",
      channel: "AI",
      trtTarget: "AJ",
      trtDateTime: "AK",
      agingTarget: "AL",
      agingDateTime: "AM",
      closeNCC: "AN",
      closeDate: "AO",
      closeTime: "AP",
      closeWeek: "AQ",
      trt: "AR",
      aging: "AS",
      reopenClose: "AT",
      reassign: "AU"
    },
    "OT Chat": {
      caseId: "B",
      date: "C",
      time: "D",
      segment: "E",
      productCategory: "F",
      triage: "G",
      preferEither: "H",
      is30: "I",
      firstAssignee: "J",
      trtTimer: "K",
      agingTimer: "L",
      poolDestination: "M",
      poolReason: "N",
      mcc: "O",
      changeToChild: "P",
      finalAssignee: "Q",
      finalSegment: "R",
      caseStatus: "S",
      amTransfer: "T",
      nonNCC: "U",
      bug: "V",
      needInfo: "W",
      firstCloseDate: "X",
      firstCloseTime: "Y",
      reopenReason: "Z",
      reopenCloseDate: "AA",
      reopenCloseTime: "AB",
      productCommerce: "AC",
      assignWeek: "AD",
      channel: "AE",
      trtTarget: "AF",
      trtDateTime: "AG",
      agingTarget: "AH",
      agingDateTime: "AI",
      closeNCC: "AJ",
      closeDate: "AK",
      closeTime: "AL",
      closeWeek: "AM",
      trt: "AN",
      aging: "AO",
      reopenClose: "AP",
      reassign: "AQ"
    },
    "3PO Chat": {
      caseId: "B",
      date: "C",
      time: "D",
      segment: "E",
      productCategory: "F",
      triage: "G",
      preferEither: "H",
      is30: "I",
      issueCategory: "J", // 3PO特有
      details: "K", // 3PO特有
      firstAssignee: "L",
      trtTimer: "M",
      agingTimer: "N",
      poolDestination: "O",
      poolReason: "P",
      mcc: "Q",
      changeToChild: "R",
      finalAssignee: "S",
      finalSegment: "T",
      caseStatus: "U",
      amTransfer: "V",
      nonNCC: "W",
      bug: "X",
      needInfo: "Y",
      firstCloseDate: "Z",
      firstCloseTime: "AA",
      reopenReason: "AB",
      reopenCloseDate: "AC",
      reopenCloseTime: "AD",
      productCommerce: "AE",
      assignWeek: "AF",
      channel: "AG",
      trtTarget: "AH",
      trtDateTime: "AI",
      agingTarget: "AJ",
      agingDateTime: "AK",
      closeNCC: "AL",
      closeDate: "AM",
      closeTime: "AN",
      closeWeek: "AO",
      trt: "AP",
      aging: "AQ",
      reopenClose: "AR",
      reassign: "AS"
    },
    "OT Phone": {
      caseId: "B",
      date: "C",
      time: "D",
      segment: "E",
      productCategory: "F",
      triage: "G",
      preferEither: "H",
      is30: "I",
      firstAssignee: "J",
      trtTimer: "K",
      agingTimer: "L",
      poolDestination: "M",
      poolReason: "N",
      mcc: "O",
      changeToChild: "P",
      finalAssignee: "Q",
      finalSegment: "R",
      caseStatus: "S",
      amTransfer: "T",
      nonNCC: "U",
      bug: "V",
      needInfo: "W",
      firstCloseDate: "X",
      firstCloseTime: "Y",
      reopenReason: "Z",
      reopenCloseDate: "AA",
      reopenCloseTime: "AB",
      productCommerce: "AC",
      assignWeek: "AD",
      channel: "AE",
      trtTarget: "AF",
      trtDateTime: "AG",
      agingTarget: "AH",
      agingDateTime: "AI",
      closeNCC: "AJ",
      closeDate: "AK",
      closeTime: "AL",
      closeWeek: "AM",
      trt: "AN",
      aging: "AO",
      reopenClose: "AP",
      reassign: "AQ"
    },
    "3PO Phone": {
      caseId: "B",
      date: "C",
      time: "D",
      segment: "E",
      productCategory: "F",
      triage: "G",
      preferEither: "H",
      is30: "I",
      issueCategory: "J", // 3PO特有
      details: "K", // 3PO特有
      firstAssignee: "L",
      trtTimer: "M",
      agingTimer: "N",
      poolDestination: "O",
      poolReason: "P",
      mcc: "Q",
      changeToChild: "R",
      finalAssignee: "S",
      finalSegment: "T",
      caseStatus: "U",
      amTransfer: "V",
      nonNCC: "W",
      bug: "X",
      needInfo: "Y",
      firstCloseDate: "Z",
      firstCloseTime: "AA",
      reopenReason: "AB",
      reopenCloseDate: "AC",
      reopenCloseTime: "AD",
      productCommerce: "AE",
      assignWeek: "AF",
      channel: "AG",
      trtTarget: "AH",
      trtDateTime: "AI",
      agingTarget: "AJ",
      agingDateTime: "AK",
      closeNCC: "AL",
      closeDate: "AM",
      closeTime: "AN",
      closeWeek: "AO",
      trt: "AP",
      aging: "AQ",
      reopenClose: "AR",
      reassign: "AS"
    }
  };

  /**
   * シート別のデフォルト値設定
   * 新規ケース作成時のチャネル別デフォルト値を定義
   */
  const SheetDefaultValues = {
    "OT Email": {
      channel: "Email",
      trtTarget: 36, // 時間単位
      agingTarget: 72 // P95標準時間（時間単位）
    },
    "3PO Email": {
      channel: "Email",
      trtTarget: 36,
      agingTarget: 72
    },
    "OT Chat": {
      channel: "Chat",
      trtTarget: 8, // チャットは短いSLA
      agingTarget: 72
    },
    "3PO Chat": {
      channel: "Chat",
      trtTarget: 8,
      agingTarget: 72
    },
    "OT Phone": {
      channel: "Phone",
      trtTarget: 8, // 電話も短いSLA
      agingTarget: 72
    },
    "3PO Phone": {
      channel: "Phone",
      trtTarget: 8,
      agingTarget: 72
    }
  };

  /**
   * シートプロパティ定義
   * 各シートの特性を定義
   */
  const SheetProperties = {
    "OT Email": {
      type: "OT",
      channel: "Email",
      has3POFields: false,
      startRow: 3, // データ開始行
      caseIdFormat: /^3-\d{13}$/,
      caseStatusValues: ["Assigned", "Waiting", "Pending", "Closed"],
      requiredFields: ["caseId", "date", "time", "segment", "productCategory", "firstAssignee"]
    },
    "3PO Email": {
      type: "3PO",
      channel: "Email",
      has3POFields: true,
      startRow: 3,
      caseIdFormat: /^3-\d{13}$/,
      caseStatusValues: ["Assigned", "Waiting", "Pending", "Closed"],
      requiredFields: ["caseId", "date", "time", "segment", "productCategory", "issueCategory", "firstAssignee"]
    },
    "OT Chat": {
      type: "OT",
      channel: "Chat",
      has3POFields: false,
      startRow: 3,
      caseIdFormat: /^3-\d{13}$/,
      caseStatusValues: ["Assigned", "Waiting", "Pending", "Closed"],
      requiredFields: ["caseId", "date", "time", "segment", "productCategory", "firstAssignee"]
    },
    "3PO Chat": {
      type: "3PO",
      channel: "Chat",
      has3POFields: true,
      startRow: 3,
      caseIdFormat: /^3-\d{13}$/,
      caseStatusValues: ["Assigned", "Waiting", "Pending", "Closed"],
      requiredFields: ["caseId", "date", "time", "segment", "productCategory", "issueCategory", "firstAssignee"]
    },
    "OT Phone": {
      type: "OT",
      channel: "Phone",
      has3POFields: false,
      startRow: 3,
      caseIdFormat: /^3-\d{13}$/,
      caseStatusValues: ["Assigned", "Waiting", "Pending", "Closed"],
      requiredFields: ["caseId", "date", "time", "segment", "productCategory", "firstAssignee"]
    },
    "3PO Phone": {
      type: "3PO",
      channel: "Phone",
      has3POFields: true,
      startRow: 3,
      caseIdFormat: /^3-\d{13}$/,
      caseStatusValues: ["Assigned", "Waiting", "Pending", "Closed"],
      requiredFields: ["caseId", "date", "time", "segment", "productCategory", "issueCategory", "firstAssignee"]
    }
  };

  /**
   * 統合ケースモデルクラス
   * すべてのシートのケースデータを一貫した形式で扱うためのモデル
   */
  class CaseModel {
    constructor(data = {}) {
      // メタデータ（どのシートのケースか）
      this.sheetType = data.sheetType || "";
      this.channelType = data.channelType || this._deriveChannelType(data.sheetType);
      this.caseType = data.caseType || this._deriveCaseType(data.sheetType);
      
      // 共通基本情報
      this.caseId = data.caseId || "";
      this.date = data.date || new Date().toISOString().split('T')[0];
      this.time = data.time || new Date().toTimeString().split(' ')[0];
      this.segment = data.segment || "";
      this.productCategory = data.productCategory || "";
      
      // 3PO特有の情報
      this.issueCategory = data.issueCategory || "";
      this.details = data.details || "";
      
      // 担当者情報
      this.firstAssignee = data.firstAssignee || Session.getActiveUser().getEmail().split('@')[0];
      this.finalAssignee = data.finalAssignee || this.firstAssignee;
      
      // ステータス情報
      this.caseStatus = data.caseStatus || "Assigned";
      this.amTransfer = data.amTransfer || "";
      this.nonNCC = data.nonNCC || "";
      this.bug = data.bug || false;
      this.needInfo = data.needInfo || false;
      this.tsConsulted = data.tsConsulted || false;
      
      // クローズ情報
      this.firstCloseDate = data.firstCloseDate || "";
      this.firstCloseTime = data.firstCloseTime || "";
      this.reopenReason = data.reopenReason || "";
      this.reopenCloseDate = data.reopenCloseDate || "";
      this.reopenCloseTime = data.reopenCloseTime || "";
      
      // タイマー情報
      this.trtTarget = data.trtTarget || this._getDefaultTRTTarget();
      this.trtDeadline = data.trtDeadline || "";
      this.agingTarget = data.agingTarget || 72; // デフォルト72時間（P95）
      this.agingDeadline = data.agingDeadline || "";
      
      // NCC関連
      this.nccValue = data.nccValue || 0;
      
      // マッピング情報（どのシートのどの列にデータがあるか）
      this.columnMappings = data.columnMappings || {};
    }
    
    /**
     * シートタイプからチャネルタイプを導出
     */
    _deriveChannelType(sheetType) {
      if (!sheetType) return "";
      if (sheetType.includes('Email')) return "Email";
      if (sheetType.includes('Chat')) return "Chat";
      if (sheetType.includes('Phone')) return "Phone";
      return "";
    }
    
    /**
     * シートタイプからケースタイプを導出
     */
    _deriveCaseType(sheetType) {
      if (!sheetType) return "";
      if (sheetType.includes('OT')) return "OT";
      if (sheetType.includes('3PO')) return "3PO";
      return "";
    }
    
    /**
     * チャネルに応じたデフォルトTRTターゲットを取得
     */
    _getDefaultTRTTarget() {
      switch (this.channelType) {
        case "Email":
          return 36; // 36時間
        case "Chat":
        case "Phone":
          return 8;  // 8時間
        default:
          return 24; // デフォルト
      }
    }
    
    /**
     * ケースデータをスプレッドシート行に変換
     */
    toRowData() {
      const mapping = SheetColumnMappings[this.sheetType] || {};
      const rowData = {};
      
      // マッピングに基づいて値を設定
      Object.keys(mapping).forEach(field => {
        if (this[field] !== undefined) {
          const column = mapping[field];
          rowData[column] = this[field];
        }
      });
      
      return rowData;
    }
    
    /**
     * SLAタイマーのカウントダウンテキストを計算
     * @returns {string} HH:MM:SS形式のカウントダウン、または "Missed"
     */
    calculateSLACountdown() {
      if (this.caseStatus !== "Assigned" || !this.trtDeadline) {
        return "";
      }
      
      const now = new Date();
      const deadline = new Date(this.trtDeadline);
      const diffMs = deadline - now;
      
      if (diffMs <= 0) {
        return "Missed";
      }
      
      // ミリ秒を時:分:秒に変換
      const hours = Math.floor(diffMs / (1000 * 60 * 60));
      const minutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
      const seconds = Math.floor((diffMs % (1000 * 60)) / 1000);
      
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }
    
    /**
     * P95エージングタイマーのカウントダウンを計算
     * @returns {string} DD:HH:MM:SS形式のカウントダウン、または "Missed"
     */
    calculateAgingCountdown() {
      if (this.caseStatus !== "Assigned" || !this.agingDeadline) {
        return "";
      }
      
      const now = new Date();
      const deadline = new Date(this.agingDeadline);
      const diffMs = deadline - now;
      
      if (diffMs <= 0) {
        return "Missed";
      }
      
      // ミリ秒を日:時:分:秒に変換
      const days = Math.floor(diffMs / (1000 * 60 * 60 * 24));
      const hours = Math.floor((diffMs % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
      const minutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
      const seconds = Math.floor((diffMs % (1000 * 60)) / 1000);
      
      return `${days}日 ${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }
    
    /**
     * T&Sコンサルトフラグを設定し、必要に応じてエージングターゲットを更新
     * @param {boolean} consulted - T&Sコンサルト状態
     */
    setTSConsulted(consulted) {
      this.tsConsulted = consulted;
      
      // T&Sコンサルトの場合、エージングターゲットを7日間に延長
      if (consulted) {
        this.agingTarget = 168; // 7日間 = 168時間
        
        // エージングデッドラインを再計算
        const caseStartDate = new Date(`${this.date}T${this.time}`);
        const deadlineDate = new Date(caseStartDate.getTime() + this.agingTarget * 60 * 60 * 1000);
        this.agingDeadline = deadlineDate.toISOString();
      } else {
        // 通常のP95に戻す
        this.agingTarget = 72; // 3日間 = 72時間
        
        // エージングデッドラインを再計算
        const caseStartDate = new Date(`${this.date}T${this.time}`);
        const deadlineDate = new Date(caseStartDate.getTime() + this.agingTarget * 60 * 60 * 1000);
        this.agingDeadline = deadlineDate.toISOString();
      }
    }
    
    /**
     * ケース情報が有効かどうかを検証
     * @returns {Object} 検証結果 { valid: boolean, errors: string[] }
     */
    validate() {
      const errors = [];
      const sheetProps = SheetProperties[this.sheetType] || { requiredFields: [] };
      
      // 必須フィールドの検証
      sheetProps.requiredFields.forEach(field => {
        if (!this[field]) {
          errors.push(`${field} は必須項目です`);
        }
      });
      
      // ケースIDのフォーマット検証
      if (this.caseId && sheetProps.caseIdFormat && !sheetProps.caseIdFormat.test(this.caseId)) {
        errors.push(`Case ID のフォーマットが不正です（例: 3-4505000031234）`);
      }
      
      // 3POシートの追加検証
      if (sheetProps.has3POFields) {
        if (!this.issueCategory) {
          errors.push(`Issue Category は3POケースでは必須です`);
        }
      }
      
      return {
        valid: errors.length === 0,
        errors: errors
      };
    }
  }

  /**
   * シートマッピングユーティリティクラス
   * 異なるシート構造間のデータ変換を担当
   */
  class SheetMapper {
    /**
     * 指定したシートの列マッピング情報を取得
     * @param {string} sheetName - シート名
     * @returns {Object} 列マッピングオブジェクト
     */
    static getColumnMapping(sheetName) {
      return SheetColumnMappings[sheetName] || {};
    }
    
    /**
     * スプレッドシートの行データからケースオブジェクトを生成
     * @param {Object} rowData - スプレッドシートから取得した行データ
     * @param {string} sheetName - シート名
     * @returns {CaseModel} 統合ケースモデルのインスタンス
     */
    static rowToCase(rowData, sheetName) {
      const mapping = this.getColumnMapping(sheetName);
      const caseData = {
        sheetType: sheetName,
        channelType: sheetName.includes('Email') ? 'Email' : 
               sheetName.includes('Chat') ? 'Chat' : 'Phone',
        caseType: sheetName.includes('OT') ? 'OT' : '3PO'
      };
      
      // マッピングを逆変換してフィールド名を取得
      const reverseMapping = {};
      Object.keys(mapping).forEach(field => {
        reverseMapping[mapping[field]] = field;
      });
      
      // 行データをケースデータに変換
      Object.keys(rowData).forEach(column => {
        const field = reverseMapping[column];
        if (field) {
          // 型変換を適用
          switch (field) {
            case 'bug':
            case 'needInfo':
            case 'tsConsulted':
              // チェックボックス値をブール値に変換
              caseData[field] = rowData[column] === true || rowData[column] === 'TRUE' || rowData[column] === 1;
              break;
            case 'date':
            case 'firstCloseDate':
            case 'reopenCloseDate':
            case 'closeDate':
              // 日付文字列を標準化
              if (rowData[column]) {
                const date = new Date(rowData[column]);
                caseData[field] = date.toISOString().split('T')[0]; // YYYY-MM-DD
              } else {
                caseData[field] = '';
              }
              break;
            case 'time':
            case 'firstCloseTime':
            case 'reopenCloseTime':
            case 'closeTime':
              // 時間文字列を標準化
              if (rowData[column]) {
                const timeParts = rowData[column].toString().split(':');
                if (timeParts.length >= 2) {
                  caseData[field] = timeParts.slice(0, 3).join(':'); // HH:MM:SS
                } else {
                  caseData[field] = rowData[column];
                }
              } else {
                caseData[field] = '';
              }
              break;
            default:
              // その他の値はそのまま代入
              caseData[field] = rowData[column];
          }
        }
      });
      
      // デッドラインを計算
      if (caseData.date && caseData.time && caseData.trtTarget) {
        const caseStartDate = new Date(`${caseData.date}T${caseData.time}`);
        const trtDeadlineDate = new Date(caseStartDate.getTime() + parseInt(caseData.trtTarget) * 60 * 60 * 1000);
        caseData.trtDeadline = trtDeadlineDate.toISOString();
        
        const agingTarget = caseData.tsConsulted ? 168 : 72; // T&S consulted は 7日間、それ以外は3日間
        const agingDeadlineDate = new Date(caseStartDate.getTime() + agingTarget * 60 * 60 * 1000);
        caseData.agingDeadline = agingDeadlineDate.toISOString();
      }
      
      return new CaseModel(caseData);
    }
    
    /**
     * ケースオブジェクトを指定したシートの行データに変換
     * @param {CaseModel} caseData - ケースモデルインスタンス
     * @param {string} sheetName - シート名
     * @returns {Object} スプレッドシートに書き込む行データ
     */
    static caseToRow(caseData, sheetName) {
      const mapping = this.getColumnMapping(sheetName);
      const rowData = {};
      
      // マッピングに従ってデータを変換
      Object.keys(mapping).forEach(field => {
        if (caseData[field] !== undefined) {
          const column = mapping[field];
          
          // 型変換を適用
          switch (field) {
            case 'bug':
            case 'needInfo':
            case 'tsConsulted':
              // ブール値をスプレッドシートのチェックボックスに適した形式に変換
              rowData[column] = caseData[field] ? 'TRUE' : 'FALSE';
              break;
            case 'channel':
              // チャネルはシートタイプに応じて固定値を使用
              rowData[column] = SheetDefaultValues[sheetName]?.channel || caseData[field];
              break;
            default:
              rowData[column] = caseData[field];
          }
        }
      });
      
      return rowData;
    }
    
    /**
     * シートのデータ範囲をすべて取得し、ケースオブジェクトの配列に変換
     * @param {string} sheetName - シート名
     * @returns {Array<CaseModel>} ケースモデルの配列
     */
    static getAllCasesFromSheet(sheetName) {
      // この関数は実際にはGASでスプレッドシートから取得したデータを使用します
      // ここでは簡略化して疑似コードを示しています
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) return [];
      
      const startRow = SheetProperties[sheetName]?.startRow || 3;
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      const cases = [];
      
      for (let i = startRow - 1; i < values.length; i++) {
        const rowData = {};
        for (let j = 0; j < values[i].length; j++) {
          // A列は0番目、B列は1番目...
          const columnLetter = String.fromCharCode(65 + j);
          rowData[columnLetter] = values[i][j];
        }
        
        // ケースIDが存在する行のみ処理
        const mapping = this.getColumnMapping(sheetName);
        const caseIdColumn = mapping.caseId;
        if (rowData[caseIdColumn]) {
          const caseModel = this.rowToCase(rowData, sheetName);
          cases.push(caseModel);
        }
      }
      
      return cases;
    }
    
    /**
     * 全シートからのケースを取得してマージ
     * @returns {Array<CaseModel>} 全シートからのケースモデルの配列
     */
    static getAllCasesFromAllSheets() {
      const allCases = [];
      const sheets = Object.keys(SheetColumnMappings);
      
      sheets.forEach(sheetName => {
        const sheetCases = this.getAllCasesFromSheet(sheetName);
        allCases.push(...sheetCases);
      });
      
      return allCases;
    }
    
    /**
     * ケースIDでケースを検索（全シート対象）
     * @param {string} caseId - 検索するケースID
     * @returns {CaseModel|null} 見つかったケースまたはnull
     */
    static findCaseByCaseId(caseId) {
      const sheets = Object.keys(SheetColumnMappings);
      
      for (const sheetName of sheets) {
        const cases = this.getAllCasesFromSheet(sheetName);
        const foundCase = cases.find(caseModel => caseModel.caseId === caseId);
        if (foundCase) {
          return foundCase;
        }
      }
      
      return null;
    }
    
    /**
     * 新規ケースをシートに追加
     * @param {CaseModel} caseData - 追加するケースデータ
     * @returns {boolean} 成功した場合true
     */
    static addNewCase(caseData) {
      const sheetName = caseData.sheetType;
      if (!sheetName || !SheetProperties[sheetName]) {
        return false;
      }
      
      // バリデーション
      const validation = caseData.validate();
      if (!validation.valid) {
        throw new Error(validation.errors.join(', '));
      }
      
      // データ変換
      const rowData = this.caseToRow(caseData, sheetName);
      
      // スプレッドシートにデータを追加（実装例）
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) return false;
      
      const startRow = SheetProperties[sheetName].startRow;
      const lastRow = Math.max(startRow, sheet.getLastRow() + 1);
      
      // 各列にデータを設定
      Object.keys(rowData).forEach(column => {
        const columnIndex = column.charCodeAt(0) - 65; // A=0, B=1, ...
        sheet.getRange(lastRow, columnIndex + 1).setValue(rowData[column]);
      });
      
      return true;
    }
    
    /**
     * 既存のケースを更新
     * @param {CaseModel} caseData - 更新するケースデータ
     * @returns {boolean} 成功した場合true
     */
    static updateCase(caseData) {
      const sheetName = caseData.sheetType;
      if (!sheetName || !SheetProperties[sheetName]) {
        return false;
      }
      
      // ケースIDに基づいて行を検索
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) return false;
      
      const mapping = this.getColumnMapping(sheetName);
      const caseIdColumn = mapping.caseId;
      const caseIdColumnIndex = caseIdColumn.charCodeAt(0) - 65 + 1; // 1-based
      
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      let rowIndex = -1;
      
      for (let i = SheetProperties[sheetName].startRow - 1; i < values.length; i++) {
        if (values[i][caseIdColumnIndex - 1] === caseData.caseId) {
          rowIndex = i + 1; // 1-based
          break;
        }
      }
      
      if (rowIndex === -1) {
        return false; // ケースが見つからない
      }
      
      // データ変換
      const rowData = this.caseToRow(caseData, sheetName);
      
      // 各列にデータを設定
      Object.keys(rowData).forEach(column => {
        const columnIndex = column.charCodeAt(0) - 65 + 1; // 1-based
        sheet.getRange(rowIndex, columnIndex).setValue(rowData[column]);
      });
      
      return true;
    }
  }
  ```

  #### シート固有のデータフローと処理

  各シートのデータはそれぞれ異なる構造を持っているため、データフローも異なります。以下にシート別の主要なデータフローを示します：

  1. **OT Email / 3PO Email**:
     - メールケースは通常、36時間のSLAを持つ
     - 応答は非リアルタイムだが、詳細な追跡が必要
     - OTとは異なり、3POには問題カテゴリと詳細フィールドが追加される
     - 日付・時間・SLA・エージングの計算ロジックは共通

  2. **OT Chat / 3PO Chat**:
     - チャットケースは8時間の短いSLAを持つ
     - リアルタイム対応が基本のため、タイマーの視覚的表示が重要
     - OTとは異なり、3POには問題カテゴリと詳細フィールドが追加される
     - チャット特有の列構造（B列からのマッピング開始）

  3. **OT Phone / 3PO Phone**:
     - 電話ケースも8時間の短いSLAを持つ
     - 音声対応の特性に合わせたデータ管理
     - OTとは異なり、3POには問題カテゴリと詳細フィールドが追加される
     - Chat同様にB列からのマッピング開始

  すべてのシートで共通しているのは以下の要素です：

  - ケースの基本情報（ID、日付、時間など）
  - SLAとエージングの計算ロジック
  - ケースステータス管理
  - クローズと再オープンのフロー

  シートタイプ別の主な違いは以下の通りです：

  1. **チャネル別の違い**:
     - Email: 36時間SLA、列構造がC列開始
     - Chat/Phone: 8時間SLA、列構造がB列開始

  2. **OT/3PO種別の違い**:
     - OT: 標準フィールドセット
     - 3PO: 問題カテゴリと詳細フィールド追加、追加列数による列マッピングのずれ

  これらの違いを吸収するために、上記のマッピングシステムを使用して統一的なデータモデルでケース情報を扱い、シート間の違いを抽象化します。


#### 統合データモデル

```javascript
// 統合ケースオブジェクトモデル
const CaseModel = {
  // メタデータ
  sheetType: "",           // シート種別 (OT Email, 3PO Chat など)
  channelType: "",         // チャネル (Email, Chat, Phone)
  caseType: "",           // ケース種別 (OT, 3PO)
  
  // 共通基本情報
  caseId: "",              // ケースID
  date: "",                // 開始日付
  time: "",                // 開始時間
  segment: "",             // セグメント
  productCategory: "",     // 製品カテゴリ
  
  // 3POシートのみの情報
  issueCategory: "",       // 問題カテゴリ (3POのみ)
  details: "",             // 詳細情報 (3POのみ)
  
  // 担当者情報
  firstAssignee: "",       // 初回担当者
  finalAssignee: "",       // 最終担当者
  
  // ステータス情報
  caseStatus: "",          // ケースステータス
  amTransfer: "",          // AM移管
  nonNCC: "",              // NCC除外理由
  bug: false,              // バグフラグ
  needInfo: false,         // 情報待ちフラグ
  tsConsulted: false,      // T&S相談状態
  
  // クローズ情報
  firstCloseDate: "",      // 初回クローズ日
  firstCloseTime: "",      // 初回クローズ時間
  reopenReason: "",        // 再オープン理由
  reopenCloseDate: "",     // 再クローズ日
  reopenCloseTime: "",     // 再クローズ時間
  
  // タイマー情報
  trtTarget: "",           // TRT目標時間
  trtDeadline: "",         // TRT期限
  agingTarget: "",         // エージング目標
  agingDeadline: "",       // エージング期限
  
  // NCC関連
  nccValue: 0,             // NCC値
  
  // マッピング情報
  columnMappings: {}       // シート固有の列マッピング
};
```

#### シート間のマッピングシステム

異なる列構造を持つシート間でデータを統一的に扱うため、動的マッピングシステムを実装しています：

```javascript
// シート別の列マッピング
const SheetColumnMappings = {
  "OT Email": {
  caseId: "C",
  date: "D",
  time: "E",
  segment: "F",
  productCategory: "G",
  firstAssignee: "L",
  caseStatus: "U",
  // 他の列マッピング...
  },
  "3PO Email": {
  caseId: "C",
  date: "D",
  time: "E",
  segment: "F",
  productCategory: "G",
  issueCategory: "L",
  details: "M",
  firstAssignee: "N",
  caseStatus: "W",
  // 他の列マッピング...
  },
  "OT Chat": {
  caseId: "B",
  date: "C",
  time: "D",
  segment: "E",
  productCategory: "F",
  firstAssignee: "J",
  caseStatus: "S",
  // 他の列マッピング...
  },
  "3PO Chat": {
  caseId: "B",
  date: "C",
  time: "D",
  segment: "E",
  productCategory: "F",
  issueCategory: "J",
  details: "K",
  firstAssignee: "L",
  caseStatus: "U",
  // 他の列マッピング...
  },
  "OT Phone": {
  caseId: "B",
  date: "C",
  time: "D",
  segment: "E",
  productCategory: "F",
  firstAssignee: "J",
  caseStatus: "S",
  // 他の列マッピング...
  },
  "3PO Phone": {
  caseId: "B",
  date: "C",
  time: "D",
  segment: "E",
  productCategory: "F",
  issueCategory: "J",
  details: "K",
  firstAssignee: "L",
  caseStatus: "U",
  // 他の列マッピング...
  }
};
```

この動的マッピングにより、アプリケーションはシート間の構造的な差異を抽象化し、一貫したインターフェースでケースデータにアクセスできます。

## 4. 機能詳細設計

### 4.1 マルチシート対応ダッシュボード

#### UI コンポーネント

- **シートフィルターセレクター**: 表示するシートの選択
- **チャネル別バッジ**: Email/Chat/Phone別の視覚的識別子
- **種別識別子**: OT/3PO種別の視覚的区別
- **ケースカードグリッド**: シートごとに色分けされたケースカード
- **統合ステータスインジケーター**: シート横断のステータス表示
- **タイマー表示**: SLAタイマーとP95タイマー
- **アクションボタン**: 編集、削除、T&S切替

#### 動作フロー

1. ページロード時に全シートから担当ケースを取得
2. シート種別とチャネル情報を付加
3. "Assigned" ステータスのケースのみフィルタリング表示
4. シート・チャネル・種別ごとに視覚的に区別
5. 1秒ごとにタイマーを更新
6. ユーザーアクションに応じた操作を実行

### 4.2 マルチシート対応ケース管理

#### シート選択式新規ケース登録フォーム

- **シート選択ドロップダウン**: 登録先シートの選択
- **動的フォーム生成**: 選択シートに応じたフィールド表示
  - 3POシート選択時のみ「Issue Category」と「Details」フィールド表示
  - チャネル固有フィールドの動的表示
- **必須項目検証**: シート別のバリデーションルール適用
- **入力補助**: 日付選択、ドロップダウン、オートコンプリート
- **自動フィールド**: 担当者情報、タイムスタンプの自動入力

#### シートコンテキスト対応ケース編集モーダル

- **シート情報の表示**: 編集中のケースのシート種別表示
- **シート固有フィールド**: シートに応じたフィールド表示/非表示
- **基本情報表示**: 共通基本情報の読み取り専用表示
- **編集可能フィールド**: シート特性に応じた編集可能項目
- **ステータス更新**: ステータスと関連項目の更新
- **クローズ情報管理**: クローズ日時と再オープン情報の管理

### 4.3 チャネル・種別対応タイマー機能

#### チャネル別Response SLAタイマー

- **チャネル固有SLAルール**:
  - Email: 36時間
  - Chat: 8時間
  - Phone: 8時間
- **シート別の計算ロジック**: 各シートの列構造に対応した計算式
- **視覚的表示**: チャネルごとの色分けと警告表示

#### P95タイマー（ケース解決タイマー）

- **基本タイマー**: Case Openから72時間
- **T&S延長タイマー**: T&S Consulted=Yesの場合は7日間
- **種別特性対応**: OT/3POの特性に応じたカウントダウン表示
- **視覚的警告**: 残り時間に応じた警告色と点滅効果

### 4.4 マルチシート統計分析

#### 統合データ集計

- **シート横断集計**: すべてのシートを統合した集計
- **シート個別集計**: シートごとの分離統計
- **チャネル分析**: Email/Chat/Phone別の統計比較
- **種別分析**: OT/3PO種別の比較分析
- **期間別集計**: 日次、週次、月次、四半期別の統計

#### 多次元視覚化

- **統合グラフ**: すべてのシートを含む統合グラフ
- **シート別グラフタブ**: シートごとに切り替え可能なグラフ表示
- **チャネル比較グラフ**: チャネル間のパフォーマンス比較
- **種別比較グラフ**: OT/3PO間のパフォーマンス比較
- **データテーブル**: 詳細数値の表形式表示

### 4.5 統合検索システム

- **全シート検索**: すべてのシートを対象とした統合検索
- **シート指定検索**: 特定シートのみを対象とする検索
- **チャネルフィルター**: チャネル別フィルタリング
- **種別フィルター**: OT/3PO別フィルタリング
- **検索結果の視覚的区別**: シート/チャネル/種別を視覚的に表示

## 5. UI/UX 設計

### 5.1 シート対応UI要素

#### シート識別システム

- **カラーコーディング**:
  - OT Email: 青色系
  - 3PO Email: 水色系
  - OT Chat: 緑色系
  - 3PO Chat: 黄緑色系
  - OT Phone: オレンジ色系
  - 3PO Phone: 黄色系

#### チャネル識別システム

- **アイコンバッジ**:
  - Email: 封筒アイコン
  - Chat: チャットバブルアイコン
  - Phone: 電話アイコン

#### 種別識別システム

- **スタイル差別化**:
  - OT: 実線ボーダー
  - 3PO: 破線ボーダー

### 5.2 レスポンシブデザイン

- すべてのデバイスサイズに対応するフレキシブルレイアウト
- モバイル向け簡略表示モード
- タブレット最適化表示
- デスクトップ向け拡張機能

### 5.3 アクセシビリティ対応

- スクリーンリーダー対応（ARIA属性）
- キーボードナビゲーション
- 高コントラストモード
- 文字サイズ調整機能

## 6. シート間データ連携

### 6.1 シートマッピングエンジン

```javascript
/**
 * シートマッピングエンジンのコアロジック
 */
class SheetMapper {
  /**
   * 指定したシートの列マッピングを取得
   */
  static getColumnMapping(sheetName) {
  return SheetColumnMappings[sheetName] || {};
  }
  
  /**
   * ケースデータをシート固有の行データに変換
   */
  static caseToRow(caseData, sheetName) {
  const mapping = this.getColumnMapping(sheetName);
  const rowData = {};
  
  // 共通フィールドのマッピング
  Object.keys(mapping).forEach(field => {
    if (caseData[field] !== undefined) {
    const column = mapping[field];
    rowData[column] = caseData[field];
    }
  });
  
  // シート固有の処理
  if (sheetName.includes('3PO')) {
    // 3PO固有の処理
  }
  
  return rowData;
  }
  
  /**
   * シート行データからケースオブジェクトに変換
   */
  static rowToCase(rowData, sheetName) {
  const mapping = this.getColumnMapping(sheetName);
  const caseData = {
    sheetType: sheetName,
    channelType: sheetName.includes('Email') ? 'Email' : 
           sheetName.includes('Chat') ? 'Chat' : 'Phone',
    caseType: sheetName.includes('OT') ? 'OT' : '3PO'
  };
  
  // 逆マッピング
  const reverseMapping = {};
  Object.keys(mapping).forEach(field => {
    reverseMapping[mapping[field]] = field;
  });
  
  // データ変換
  Object.keys(rowData).forEach(column => {
    const field = reverseMapping[column];
    if (field) {
    caseData[field] = rowData[column];
    }
  });
  
  return caseData;
  }
}
```

### 6.2 データ検証と整合性

- シート間のデータ整合性検証
- シート固有のバリデーションルール
- 必須フィールド検証のシート別設定

### 6.3 バッチ処理と同期

- 全シートからのデータ一括取得
- 効率的なメモリ使用のための増分更新
- キャッシング戦略によるパフォーマンス最適化

## 7. 統計と分析フレームワーク

### 7.1 マルチシート集計エンジン

```javascript
/**
 * 複数シート対応の集計エンジン
 */
class MultiSheetAnalytics {
  /**
   * すべてのシートからデータを収集して集計
   */
  static aggregateAllSheets(period = 'daily') {
  const sheets = ["OT Email", "3PO Email", "OT Chat", "3PO Chat", "OT Phone", "3PO Phone"];
  const results = {
    totalCases: 0,
    bySheet: {},
    byChannel: {
    Email: 0,
    Chat: 0,
    Phone: 0
    },
    byType: {
    OT: 0,
    '3PO': 0
    }
  };
  
  // 各シートを処理
  sheets.forEach(sheetName => {
    const sheetData = this.aggregateSheet(sheetName, period);
    results.totalCases += sheetData.totalCases;
    results.bySheet[sheetName] = sheetData;
    
    // チャネル集計
    if (sheetName.includes('Email')) {
    results.byChannel.Email += sheetData.totalCases;
    } else if (sheetName.includes('Chat')) {
    results.byChannel.Chat += sheetData.totalCases;
    } else {
    results.byChannel.Phone += sheetData.totalCases;
    }
    
    // 種別集計
    if (sheetName.includes('OT')) {
    results.byType.OT += sheetData.totalCases;
    } else {
    results.byType['3PO'] += sheetData.totalCases;
    }
  });
  
  return results;
  }
  
  /**
   * 単一シートの集計を実行
   */
  static aggregateSheet(sheetName, period) {
  // シート固有の集計ロジック
  // ...
  
  return {
    totalCases: 0,
    resolvedCases: 0,
    nccCount: 0,
    byStatus: {},
    byAssignee: {}
  };
  }
}
```

### 7.2 NCC計算エンジン

```javascript
/**
 * NCC計算特化エンジン
 */
class NCCCalculator {
  /**
   * すべてのシートからNCCを計算
   */
  static calculateNCCForAllSheets(period = 'daily') {
  const sheets = ["OT Email", "3PO Email", "OT Chat", "3PO Chat", "OT Phone", "3PO Phone"];
  const results = {
    totalNCC: 0,
    bySheet: {},
    byUser: {}
  };
  
  sheets.forEach(sheetName => {
    const sheetNCC = this.calculateNCCForSheet(sheetName, period);
    results.totalNCC += sheetNCC.total;
    results.bySheet[sheetName] = sheetNCC;
    
    // ユーザー別集計を統合
    Object.keys(sheetNCC.byUser).forEach(user => {
    if (!results.byUser[user]) {
      results.byUser[user] = 0;
    }
    results.byUser[user] += sheetNCC.byUser[user];
    });
  });
  
  return results;
  }
  
  /**
   * 単一シートのNCC計算
   */
  static calculateNCCForSheet(sheetName, period) {
  // シート固有のNCC計算ロジック
  const mapping = SheetMapper.getColumnMapping(sheetName);
  const caseStatusCol = mapping.caseStatus;
  const nonNCCCol = mapping.nonNCC;
  const bugCol = mapping.bug;
  const finalAssigneeCol = mapping.finalAssignee;
  
  // シート固有の列位置でNCCを計算
  // ...
  
  return {
    total: 0,
    byUser: {}
  };
  }
}
```

### 7.3 パフォーマンス指標ダッシュボード

- シート別パフォーマンス
- チャネル別パフォーマンス
- 種別別パフォーマンス
- ユーザー別パフォーマンス
- 期間比較グラフ

## 8. 実装と導入ロードマップ

### 8.1 フェーズ1: 基盤拡張（1-2週間）

- シートマッピングエンジンの開発
- データモデルの拡張と抽象化
- 基本的なマルチシート読み取り機能

### 8.2 フェーズ2: UI拡張（2-3週間）

- シート選択機能の実装
- 動的フォーム生成の開発
- シート識別システムの視覚的実装

### 8.3 フェーズ3: 分析機能拡張（2-3週間）

- マルチシート集計エンジンの実装
- シート別/チャネル別/種別別の統計表示
- 詳細分析ダッシュボード

### 8.4 フェーズ4: 検証と最適化（1-2週間）

- 全シートとの互換性テスト
- パフォーマンス最適化
- エッジケース処理の改善

### 8.5 フェーズ5: 展開とフィードバック（継続的）

- ユーザートレーニング
- フィードバック収集と改善
- 追加機能のリリース

## 9. メンテナンスと拡張性

### 9.1 コード保守性

- モジュール化されたアーキテクチャ
- シート構造変更への対応策
- 設定ファイルによる柔軟な調整

### 9.2 将来の拡張性

- 新規シート追加の容易さ
- 新規チャネル対応のフレームワーク
- カスタムフィールドサポート

### 9.3 バックアップと復元

- 設定のバックアップ
- データ復元機能
- バージョン管理

## 10. 付録: シート構造リファレンス

### 10.1 OT Email シート構造

主要列:
- A: Date
- B: Case (ハイパーリンク)
- C: Case ID
- D: Case Open Date
- E: Time
- F: Incoming Segment
- G: Product Category
- ...

### 10.2 3PO Email シート構造

主要列:
- A: Date
- B: Cases (ハイパーリンク)
- C: Case ID
- D: Case Open Date
- E: Time
- F: Incoming Segment
- G: Product Category
- ...
- L: Issue Category (3PO固有)
- M: Details (3PO固有)
- ...

### 10.3 OT Chat シート構造

主要列:
- A: Cases (ハイパーリンク)
- B: Case ID
- C: Case Open Date
- D: Time
- E: Incoming Segment
- F: Product Category
- ...

### 10.4 3PO Chat シート構造

主要列:
- A: Cases (ハイパーリンク)
- B: Case ID
- C: Case Open Date
- D: Time
- E: Incoming Segment
- F: Product Category
- ...
- J: Issue Category (3PO固有)
- K: Details (3PO固有)
- ...

### 10.5 OT Phone シート構造

主要列:
- A: Cases (ハイパーリンク)
- B: Case ID
- C: Case Open Date
- D: Time
- E: Incoming Segment
- F: Product Category
- ...

### 10.6 3PO Phone シート構造

主要列:
- A: Cases (ハイパーリンク)
- B: Case ID
- C: Case Open Date
- D: Time
- E: Incoming Segment
- F: Product Category
- ...
- J: Issue Category (3PO固有)
- K: Details (3PO固有)
- ...

## 11. 用語集

- **SLA (Service Level Agreement)**: サービス品質保証のための応答時間枠
- **TRT (Turn Around Time)**: ケース初期応答時間目標（セグメントとチャネル別に異なる）
  - **Platinum Email**: 24時間（80%）
  - **Platinum Chat/Phone**: 6時間（80%）
  - **Titanium/Gold Email**: 36時間（80%）
  - **Titanium/Gold Chat/Phone**: 8時間（80%）
- **P95**: 一般的なケース解決時間目標、標準は3日間（72時間）
- **Bug対応**: バグ優先度に基づく特別なSLA
  - **P0**: 3時間
  - **P1**: 13時間
- **T&S Consulted**: Technical and Specialist consultation flag、解決時間目標は7日間
- **NCC**: Non-Contact Complete（通常のクローズフロー以外）
- **AM Transfer**: Account Manager移管フラグ
- **OT**: Outbound Team - 通常のサポートチーム
- **3PO**: 特殊なポリシー関連などを扱うチーム
- **Channel**: コミュニケーションチャネル（Email、Chat、Phone）
- **シートタイプ**: 6種類のシート（OT Email, 3PO Email, OT Chat, 3PO Chat, OT Phone, 3PO Phone）
- **セグメント**: 顧客重要度区分（Platinum、Titanium、Gold、Silver、Bronze）

## 12. SLAメトリック詳細

| **メトリック** | **チャネル** | **セグメント** | **標準対応(80%)** | **全体目標(95%)** | **バグ対応** | **T&Sコンサルト** |
|------------|------------|------------|---------------|--------------|-----------|--------------|
| TRT        | Email      | Platinum   | 24時間        | 3日間         | P0: 3時間<br>P1: 13時間 | 7日間        |
| TRT        | Chat/Phone | Platinum   | 6時間         | 3日間         | P0: 3時間<br>P1: 13時間 | 7日間        |
| TRT        | Email      | Titanium/Gold | 36時間     | 3日間         | P0: 3時間<br>P1: 13時間 | 7日間        |
| TRT        | Chat/Phone | Titanium/Gold | 8時間      | 3日間         | P0: 3時間<br>P1: 13時間 | 7日間        |

これらのSLAメトリックは、アプリケーション内のTRTタイマーとエージングタイマーの設定に直接反映され、セグメントとチャネルに基づいて動的に計算されます。
