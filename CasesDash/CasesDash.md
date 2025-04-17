

# CasesDash.md

# CaseDash - 3PO チャットケース管理システム

## 1. プロジェクト概要

### 1.1 概要と目的

こちらのツールは、Google 広告サポートチームがケース管理を効率化するためのウェブベースのツールです。Google スプレッドシートと連携し、チーム内のケース割り当て、追跡、期限管理、統計分析を一元化しますUI/UX優先設計により、複雑なケース管理作業を直感的に行えるようにしています。

### 1.2 背景

大規模サポートチームでは、複数のケースを同時に処理し、SLA（Service Level Agreement）を確実に守る必要があります。既存の管理方法はスプレッドシートベースで、以下の問題がありました：

- 「3PO Chat」シートのみでしか対応しておらず、指定したスプレッドシートIDの6つのシート（「### 3.3 データモデル」を参照）に対応しきれておらず、利便性が低いツールだった
- データの視覚化が不十分で重要情報が見えにくい
- 期限管理が手動で行われており、ミスのリスクが高い
- ケースステータスの更新や検索が非効率
- チーム全体のパフォーマンス分析が困難

CaseDashはこれらの問題を解決し、サポートチームの効率と顧客満足度を向上させます。

### 1.3 主要機能

### ダッシュボード機能

- アクティブケースの一覧表示と視覚的なステータス表示
- リアルタイムのSLAタイマー（残り応答時間のカウントダウン）
- P95タイマー（Case Open から72時間/7日間のカウントダウン）
- T&S Consulted 状態の切り替え機能
- ケースの詳細表示と編集機能
- ダッシュボードからのケース削除機能

### ケース管理機能

- 新規ケース登録フォーム（必須項目検証付き）
- セグメント、製品カテゴリ、問題カテゴリなどの属性管理
- Case Status、AM Transfer、non NCCなどの状態管理
- 初回クローズおよび再オープン情報の管理
- Live Mode（別ウィンドウでの登録機能）

### 検索機能

- Case ID による即時検索
- 検索結果からのケース詳細表示と編集
- 非アクティブケースの再アクティブ化機能

### 統計分析機能

- 日別/週別/月別/四半期別の統計表示
- ケース解決率とNCCの視覚化
- グラフによるデータの視覚化（折れ線/円/棒グラフ）
- ステータス別の詳細分析

### その他の機能

- ダークモード/ライトモード切り替え
- レスポンシブデザイン対応
- スプレッドシート設定と接続テスト機能
- エラーハンドリングとユーザーフィードバック

### 1.4 ターゲットユーザー

- カスタマーサポートの専門家
- ケース管理担当者
- チームリーダーとマネージャー
- サポート品質分析官
- カスタマーサクセスチーム

### 1.5 解決する問題

- **効率性**: ケース情報の入力、更新、検索、分析を効率化
- **正確性**: 自動化されたタイマーとリマインダーでSLA違反を防止
- **可視性**: ダッシュボードとグラフでケース状況を明確に把握
- **一貫性**: 標準化されたプロセスでデータ入力と管理を統一
- **分析力**: リアルタイムの統計でパフォーマンス指標を追跡

## 2. 使用技術

### 2.1 フロントエンド

### コア技術

- **HTML5**: 最新のセマンティックマークアップ
- **CSS3**: カスタム変数、グリッドレイアウト、フレックスボックス
- **JavaScript (ES6+)**: モダンなJavaScript機能とブラウザAPI

### UI フレームワーク

- **Material Design Components for Web**: Googleの公式Material Design実装
    - MDC Tab Bar, Dialog, Select, TextField, Button, Checkbox等
    - Material Icons & Symbols (アイコンフォント)

### その他のライブラリ

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
    - シート名: "3PO Chat"
    - 構造化データ (列定義された表形式)
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
│                 │     │                 │     │                 │
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
    - `script.html`: コアロジックとユーティリティ
    - `dashboard.html`: ダッシュボード機能
    - `caseManagement.html`: ケース管理機能
    - `statsUI.html`: 統計分析機能
    - `styles.html`: スタイル定義
- **バックエンドモジュール**
    - `Code.gs`: メインスクリプト（サーバーサイドロジック）
    - `setup_spreadsheet.gs`: スプレッドシート準備ツール
- **HTMLテンプレート**
    - `index.html`: メインアプリケーション
    - `live.html`: Live Mode用UI

### 3.3 データモデル

### スプレッドシートの構造（HTML形式）

```
    <h2>スプレッドシート名: Titanium/Gold 2025Q2 のコピー</h2>
    <p>分析したシート数: 6</p>
  
      <div class="sheet-container">
        <h3>シート 1: OT Email</h3>
    <div class="table-container"><table border="1"><tbody><tr><th>行/列</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>T</th><th>U</th><th>V</th><th>W</th><th>X</th><th>Y</th><th>Z</th><th>[</th><th>\</th><th>]</th><th>^</th><th>_</th><th>`</th><th>a</th><th>b</th><th>c</th><th>d</th><th>e</th><th>f</th><th>g</th><th>h</th><th>i</th><th>j</th><th>k</th><th>l</th><th>m</th><th>n</th></tr><tr><th>1</th><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Case Open</td><td>&nbsp;</td><td>Incoming</td><td>Product</td><td>&nbsp;</td><td>Prefer</td><td>AM</td><td>Is</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Pool Transfer</td><td>&nbsp;</td><td>&nbsp;</td><td>Change</td><td>&nbsp;</td><td>Final</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Need</td><td>1st Close</td><td>&nbsp;</td><td>Reopen</td><td>Reopen Close</td><td>&nbsp;</td><td>&nbsp;</td><td>Product</td><td>Assign</td><td>&nbsp;</td><td>TRT</td><td>&nbsp;</td><td>Aging</td><td>&nbsp;</td><td>Close</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><th>2</th><td>Date</td><td>Case</td><td>Case ID</td><td>Date</td><td>Time</td><td>Segment</td><td>Category</td><td>Triage</td><td>Either</td><td>Initiated</td><td>3.0</td><td>1st Assignee</td><td>TRT Timer</td><td>Aging Timer</td><td>Destination</td><td>Reason</td><td>MCC</td><td>to Child</td><td>Final Assignee</td><td>Segment</td><td>Case Status</td><td>AM Transfer</td><td>non NCC</td><td>Bug</td><td>Info</td><td>Date</td><td>Time</td><td>Reason</td><td>Date</td><td>Time</td><td>&nbsp;</td><td>Commerc</td><td>Week</td><td>Channel</td><td>Target</td><td>Date+Time</td><td>Target</td><td>Date+Time</td><td>NCC</td><td>Date</td><td>Time</td><td>Week</td><td>TRT</td><td>Aging</td><td>Reopen Close</td><td>Reassign</td></tr><tr><th>3</th><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",hyperlink("https://cases.connect.corp.google.com/#/case/"&amp;C3:C,"Cases")))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IF((U3:U="Assigned")*(C3:C&lt;&gt;""),IF((AJ3:AJ-NOW())&gt;0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>#REF!<br><span class="formula">==ARRAYFORMULA(IF((U3:U="Assigned")*(C3:C&lt;&gt;""),IF((AL3:AL-NOW())&gt;0, TEXT(AL3:AL-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>0</td><td>&nbsp;<br><span class="formula">==arrayformula(if(A3:A="","",A3:A-weekday(A3:A,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","","Email"))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(C3:C="","",xlookup(AH3:AH&amp;T3:T&amp;AF3:AF,index!M:M,index!N:N))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",D3:D+E3:E+AI3:AI))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(C3:C="","",xlookup(AH3:AH&amp;T3:T&amp;AF3:AF,index!M:M,index!O:O))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",D3:D+E3:E+AK3:AK))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",if(W3:W&lt;&gt;"",0,if(U3:U="Assigned","",IF(X3:X=1,2,1)))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM="","",if(AC3:AC="",Z3:Z,AC3:AC)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM="","",if(AD3:AD="",AA3:AA,AD3:AD)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AN3:AN="","",AN3:AN-WEEKDAY(AN3:AN,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AJ3:AJ&gt;(AN3:AN+AO3:AO),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AL3:AL&gt;(AN3:AN+AO3:AO),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AC3:AC="",0,1)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(S3:S="","",if(L3:L=S3:S,"",1)))</span></td></tr></tbody></table></div></div>
      <div class="sheet-container">
        <h3>シート 2: 3PO Email</h3>
    <div class="table-container"><table border="1"><tbody><tr><th>行/列</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>T</th><th>U</th><th>V</th><th>W</th><th>X</th><th>Y</th><th>Z</th><th>[</th><th>\</th><th>]</th><th>^</th><th>_</th><th>`</th><th>a</th><th>b</th><th>c</th><th>d</th><th>e</th><th>f</th><th>g</th><th>h</th><th>i</th><th>j</th><th>k</th><th>l</th><th>m</th><th>n</th><th>o</th><th>p</th></tr><tr><th>1</th><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Case Open</td><td>&nbsp;</td><td>Incoming</td><td>Product</td><td>&nbsp;</td><td>Prefer</td><td>AM</td><td>Is</td><td>Issue</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Pool Transfer</td><td>&nbsp;</td><td>&nbsp;</td><td>Change</td><td>&nbsp;</td><td>Final</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Need</td><td>1st Close</td><td>&nbsp;</td><td>Reopen</td><td>Reopen Close</td><td>&nbsp;</td><td>&nbsp;</td><td>Product</td><td>Assign</td><td>&nbsp;</td><td>TRT</td><td>&nbsp;</td><td>Aging</td><td>&nbsp;</td><td>Close</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><th>2</th><td>Date</td><td>Cases</td><td>Case ID</td><td>Date</td><td>Time</td><td>Segment</td><td>Category</td><td>Triage</td><td>Either</td><td>Initiated</td><td>3.0</td><td>Category</td><td>Details</td><td>1st Assignee</td><td>TRT Timer</td><td>Aging Timer</td><td>Destination</td><td>Reason</td><td>MCC</td><td>to Child</td><td>Final Assignee</td><td>Segment</td><td>Case Status</td><td>AM Transfer</td><td>non NCC</td><td>Bug</td><td>Info</td><td>Date</td><td>Time</td><td>Reason</td><td>Date</td><td>Time</td><td>&nbsp;</td><td>&nbsp;</td><td>Week</td><td>Channel</td><td>Target</td><td>Date+Time</td><td>Target</td><td>Date+Time</td><td>NCC</td><td>Date</td><td>Time</td><td>Week</td><td>TRT</td><td>Aging</td><td>Reopen Close</td><td>Reassign</td></tr><tr><th>3</th><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",hyperlink("https://cases.connect.corp.google.com/#/case/"&amp;C3:C,"Cases")))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IF((W3:W="Assigned")*(C3:C&lt;&gt;""),IF((AL3:AL-NOW())&gt;0, TEXT(AL3:AL-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>#REF!<br><span class="formula">==ARRAYFORMULA(IF((W3:W="Assigned")*(C3:C&lt;&gt;""),IF((AN3:AN-NOW())&gt;0, TEXT(AN3:AN-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IFERROR(IFS(G3:G="Billing",2,G3:G="Policy",3),""))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(A3:A="","",A3:A-weekday(A3:A,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","","Email"))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(C3:C="","",xlookup(AJ3:AJ&amp;V3:V&amp;AH3:AH,index!M:M,index!N:N))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",D3:D+E3:E+AK3:AK))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(C3:C="","",xlookup(AJ3:AJ&amp;V3:V&amp;AH3:AH,index!M:M,index!O:O))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",D3:D+E3:E+AM3:AM))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",if(Y3:Y&lt;&gt;"",0,if(W3:W="Assigned","",IF(Z3:Z=1,2,1)))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AO3:AO="","",if(AE3:AE="",AB3:AB,AE3:AE)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AO3:AO="","",if(AF3:AF="",AC3:AC,AF3:AF)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AP3:AP="","",AP3:AP-WEEKDAY(AP3:AP,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AO3:AO&lt;&gt;1,"",if(AL3:AL&gt;(AP3:AP+AQ3:AQ),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AO3:AO&lt;&gt;1,"",if(AN3:AN&gt;(AP3:AP+AQ3:AQ),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AO3:AO&lt;&gt;1,"",if(AE3:AE="",0,1)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(U3:U="","",if(N3:N=U3:U,"",1)))</span></td></tr></tbody></table></div></div>
      <div class="sheet-container">
        <h3>シート 3: OT Chat</h3>
    <div class="table-container"><table border="1"><tbody><tr><th>行/列</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>T</th><th>U</th><th>V</th><th>W</th><th>X</th><th>Y</th><th>Z</th><th>[</th><th>\</th><th>]</th><th>^</th><th>_</th><th>`</th><th>a</th><th>b</th><th>c</th><th>d</th><th>e</th><th>f</th><th>g</th><th>h</th><th>i</th><th>j</th><th>k</th><th>l</th></tr><tr><th>1</th><td>&nbsp;</td><td>&nbsp;</td><td>Case Open</td><td>&nbsp;</td><td>Incoming</td><td>Product</td><td>&nbsp;</td><td>Prefer</td><td>Is</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Pool Transfer</td><td>&nbsp;</td><td>&nbsp;</td><td>Change</td><td>&nbsp;</td><td>Final</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Need</td><td>1st Close</td><td>&nbsp;</td><td>Reopen</td><td>Reopen Close</td><td>&nbsp;</td><td>&nbsp;</td><td>Product</td><td>Assign</td><td>&nbsp;</td><td>TRT</td><td>&nbsp;</td><td>Aging</td><td>&nbsp;</td><td>Close</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><th>2</th><td>Cases</td><td>Case ID</td><td>Date</td><td>Time</td><td>Segment</td><td>Category</td><td>Triage</td><td>Either</td><td>3.0</td><td>1st Assignee</td><td>TRT Timer</td><td>Aging Timer</td><td>Destination</td><td>Reason</td><td>MCC</td><td>to Child</td><td>Final Assignee</td><td>Segment</td><td>Case Status</td><td>AM Transfer</td><td>non NCC</td><td>Bug</td><td>Info</td><td>Date</td><td>Time</td><td>Reason</td><td>Date</td><td>Time</td><td>&nbsp;</td><td>Commerce</td><td>Week</td><td>Channel</td><td>Target</td><td>Date+Time</td><td>Target</td><td>Date+Time</td><td>NCC</td><td>Date</td><td>Time</td><td>Week</td><td>TRT</td><td>Aging</td><td>Reopen Close</td><td>Reassign</td></tr><tr><th>3</th><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",hyperlink("https://cases.connect.corp.google.com/#/case/"&amp;B3:B,"Cases")))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IF((S3:S="Assigned")*(B3:B&lt;&gt;""),IF((AH3:AH-NOW())&gt;0, TEXT(AH3:AH-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>#REF!<br><span class="formula">==ARRAYFORMULA(IF((S3:S="Assigned")*(#REF!&lt;&gt;""),IF((AJ3:AJ-NOW())&gt;0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",C3:C-weekday(C3:C,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","","Chat"))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(B3:B="","",xlookup(AF3:AF&amp;R3:R&amp;AD3:AD,index!M:M,index!N:N))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AG3:AG))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(B3:B="","",xlookup(AF3:AF&amp;R3:R&amp;AD3:AD,index!M:M,index!O:O))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AI3:AI))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",if(U3:U&lt;&gt;"",0,if(S3:S="Assigned","",IF(V3:V=1,2,1)))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK="","",if(AA3:AA="",X3:X,AA3:AA)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK="","",if(AB3:AB="",Y3:Y,AB3:AB)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AL3:AL="","",AL3:AL-weekday(AL3:AL,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK&lt;&gt;1,"",if(AH3:AH&gt;(AL3:AL+AM3:AM),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK&lt;&gt;1,"",if(AJ3:AJ&gt;(AL3:AL+AM3:AM),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK&lt;&gt;1,"",if(AA3:AA="",0,1)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(Q3:Q="","",if(J3:J=Q3:Q,"",1)))</span></td></tr></tbody></table></div></div>
      <div class="sheet-container">
        <h3>シート 4: 3PO Chat</h3>
    <div class="table-container"><table border="1"><tbody><tr><th>行/列</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>T</th><th>U</th><th>V</th><th>W</th><th>X</th><th>Y</th><th>Z</th><th>[</th><th>\</th><th>]</th><th>^</th><th>_</th><th>`</th><th>a</th><th>b</th><th>c</th><th>d</th><th>e</th><th>f</th><th>g</th><th>h</th><th>i</th><th>j</th><th>k</th><th>l</th><th>m</th><th>n</th></tr><tr><th>1</th><td>&nbsp;</td><td>&nbsp;</td><td>Case Open</td><td>&nbsp;</td><td>Incoming</td><td>Product</td><td>&nbsp;</td><td>Prefer</td><td>Is</td><td>Issue</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Pool Transfer</td><td>&nbsp;</td><td>&nbsp;</td><td>Change</td><td>&nbsp;</td><td>Final</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Need</td><td>1st Close</td><td>&nbsp;</td><td>Reopen</td><td>Reopen Close</td><td>&nbsp;</td><td>&nbsp;</td><td>Product</td><td>Assign</td><td>&nbsp;</td><td>TRT</td><td>&nbsp;</td><td>Aging</td><td>&nbsp;</td><td>Close</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><th>2</th><td>Cases</td><td>Case ID</td><td>Date</td><td>Time</td><td>Segment</td><td>Category</td><td>Triage</td><td>Either</td><td>3.0</td><td>Category</td><td>Details</td><td>1st Assignee</td><td>TRT Timer</td><td>Aging Timer</td><td>Destination</td><td>Reason</td><td>MCC</td><td>to Child</td><td>Final Assignee</td><td>Segment</td><td>Case Status</td><td>AM Transfer</td><td>non NCC</td><td>Bug</td><td>Info</td><td>Date</td><td>Time</td><td>Reason</td><td>Date</td><td>Time</td><td>&nbsp;</td><td>&nbsp;</td><td>Week</td><td>Channel</td><td>Target</td><td>Date+Time</td><td>Target</td><td>Date+Time</td><td>NCC</td><td>Date</td><td>Time</td><td>Week</td><td>TRT</td><td>Aging</td><td>Reopen Close</td><td>Reassign</td></tr><tr><th>3</th><td>Cases<br><span class="formula">==arrayformula(if(B3:B="","",hyperlink("https://cases.connect.corp.google.com/#/case/"&amp;B3:B,"Cases")))</span></td><td>3-4505000038335</td><td>2025/04/02 00:00:00</td><td>1899/12/30 18:59:00</td><td>Gold</td><td>Search</td><td>1899/12/30 00:00:00</td><td>1899/12/30 00:00:00</td><td>1899/12/30 00:00:00</td><td>&nbsp;</td><td>&nbsp;</td><td>daito</td><td>Missed<br><span class="formula">==ARRAYFORMULA(IF((U3:U="Assigned")*(B3:B&lt;&gt;""),IF((AJ3:AJ-NOW())&gt;0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>kotaf</td><td>Gold</td><td>Assigned</td><td>&nbsp;</td><td>&nbsp;</td><td>1899/12/30 00:00:00</td><td>1899/12/30 00:00:00</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IFERROR(IFS(F3:F="Billing",2,F3:F="Policy",3),""))</span></td><td>2025/03/30 00:00:00<br><span class="formula">==arrayformula(if(C3:C="","",C3:C-weekday(C3:C,1)+1))</span></td><td>Chat<br><span class="formula">==arrayformula(if(B3:B="","","Chat"))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(B3:B="","",xlookup(AH3:AH&amp;T3:T&amp;AF3:AF,index!M:M,index!N:N))))</span></td><td>2025/04/02 18:59:00<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AI3:AI))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(B3:B="","",xlookup(AH3:AH&amp;T3:T&amp;AF3:AF,index!M:M,index!O:O))))</span></td><td>2025/04/02 18:59:00<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AK3:AK))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",if(W3:W&lt;&gt;"",0,if(U3:U="Assigned","",IF(X3:X=1,2,1)))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM="","",if(AC3:AC="",Z3:Z,AC3:AC)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM="","",if(AD3:AD="",AA3:AA,AD3:AD)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AN3:AN="","",AN3:AN-weekday(AN3:AN,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AJ3:AJ&gt;(AN3:AN+AO3:AO),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AL3:AL&gt;(AN3:AN+AO3:AO),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AC3:AC="",0,1)))</span></td><td>1<br><span class="formula">==arrayformula(if(S3:S="","",if(L3:L=S3:S,"",1)))</span></td></tr></tbody></table></div></div>
      <div class="sheet-container">
        <h3>シート 5: OT Phone</h3>
    <div class="table-container"><table border="1"><tbody><tr><th>行/列</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>T</th><th>U</th><th>V</th><th>W</th><th>X</th><th>Y</th><th>Z</th><th>[</th><th>\</th><th>]</th><th>^</th><th>_</th><th>`</th><th>a</th><th>b</th><th>c</th><th>d</th><th>e</th><th>f</th><th>g</th><th>h</th><th>i</th><th>j</th><th>k</th><th>l</th></tr><tr><th>1</th><td>&nbsp;</td><td>&nbsp;</td><td>Case Open</td><td>&nbsp;</td><td>Incoming</td><td>Product</td><td>&nbsp;</td><td>Prefer</td><td>Is</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Pool Transfer</td><td>&nbsp;</td><td>&nbsp;</td><td>Change</td><td>&nbsp;</td><td>Final</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Need</td><td>1st Close</td><td>&nbsp;</td><td>Reopen</td><td>Reopen Close</td><td>&nbsp;</td><td>&nbsp;</td><td>Product</td><td>Assign</td><td>&nbsp;</td><td>TRT</td><td>&nbsp;</td><td>Aging</td><td>&nbsp;</td><td>Close</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><th>2</th><td>Cases</td><td>Case ID</td><td>Date</td><td>Time</td><td>Segment</td><td>Category</td><td>Triage</td><td>Either</td><td>3.0</td><td>1st Assignee</td><td>TRT Timer</td><td>Aging Timer</td><td>Destination</td><td>Reason</td><td>MCC</td><td>to Child</td><td>Final Assignee</td><td>Segment</td><td>Case Status</td><td>AM Transfer</td><td>non NCC</td><td>Bug</td><td>Info</td><td>Date</td><td>Time</td><td>Reason</td><td>Date</td><td>Time</td><td>&nbsp;</td><td>Commerce</td><td>Week</td><td>Channel</td><td>Target</td><td>Date+Time</td><td>Target</td><td>Date+Time</td><td>NCC</td><td>Date</td><td>Time</td><td>Week</td><td>TRT</td><td>Aging</td><td>Reopen Close</td><td>Reassign</td></tr><tr><th>3</th><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",hyperlink("https://cases.connect.corp.google.com/#/case/"&amp;B3:B,"Cases")))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IF((S3:S="Assigned")*(B3:B&lt;&gt;""),IF((AH3:AH-NOW())&gt;0, TEXT(AH3:AH-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>#REF!<br><span class="formula">==ARRAYFORMULA(IF((S3:S="Assigned")*(A3:A&lt;&gt;""),IF((AJ3:AJ-NOW())&gt;0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",C3:C-weekday(C3:C,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","","Phone"))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(F3:F="","",xlookup(AF3:AF&amp;R3:R&amp;AD3:AD,index!M:M,index!N:N))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AG3:AG))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(F3:F="","",xlookup(AF3:AF&amp;R3:R&amp;AD3:AD,index!M:M,index!O:O))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AI3:AI))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",if(U3:U&lt;&gt;"",0,if(S3:S="Assigned","",IF(V3:V=1,2,1)))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK="","",if(AA3:AA="",X3:X,AA3:AA)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK="","",if(AB3:AB="",Y3:Y,AB3:AB)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AL3:AL="","",AL3:AL-weekday(AL3:AL,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK&lt;&gt;1,"",if(AH3:AH&gt;(AL3:AL+AM3:AM),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK&lt;&gt;1,"",if(AJ3:AJ&gt;(AL3:AL+AM3:AM),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AK3:AK&lt;&gt;1,"",if(AA3:AA="",0,1)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(Q3:Q="","",if(J3:J=Q3:Q,"",1)))</span></td></tr></tbody></table></div></div>
      <div class="sheet-container">
        <h3>シート 6: 3PO Phone</h3>
    <div class="table-container"><table border="1"><tbody><tr><th>行/列</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>T</th><th>U</th><th>V</th><th>W</th><th>X</th><th>Y</th><th>Z</th><th>[</th><th>\</th><th>]</th><th>^</th><th>_</th><th>`</th><th>a</th><th>b</th><th>c</th><th>d</th><th>e</th><th>f</th><th>g</th><th>h</th><th>i</th><th>j</th><th>k</th><th>l</th><th>m</th><th>n</th></tr><tr><th>1</th><td>&nbsp;</td><td>&nbsp;</td><td>Case Open</td><td>&nbsp;</td><td>Incoming</td><td>Product</td><td>&nbsp;</td><td>Prefer</td><td>Is</td><td>Issue</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Pool Transfer</td><td>&nbsp;</td><td>&nbsp;</td><td>Change</td><td>&nbsp;</td><td>Final</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>Need</td><td>1st Close</td><td>&nbsp;</td><td>Reopen</td><td>Reopen Close</td><td>&nbsp;</td><td>&nbsp;</td><td>Product</td><td>Assign</td><td>&nbsp;</td><td>TRT</td><td>&nbsp;</td><td>Aging</td><td>&nbsp;</td><td>Close</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><th>2</th><td>Cases</td><td>Case ID</td><td>Date</td><td>Time</td><td>Segment</td><td>Category</td><td>Triage</td><td>Either</td><td>3.0</td><td>Category</td><td>Details</td><td>1st Assignee</td><td>TRT Timer</td><td>Aging Timer</td><td>Destination</td><td>Reason</td><td>MCC</td><td>to Child</td><td>Final Assignee</td><td>Segment</td><td>Case Status</td><td>AM Transfer</td><td>non NCC</td><td>Bug</td><td>Info</td><td>Date</td><td>Time</td><td>Reason</td><td>Date</td><td>Time</td><td>&nbsp;</td><td>&nbsp;</td><td>Week</td><td>Channel</td><td>Target</td><td>Date+Time</td><td>Target</td><td>Date+Time</td><td>NCC</td><td>Date</td><td>Time</td><td>Week</td><td>TRT</td><td>Aging</td><td>Reopen Close</td><td>Reassign</td></tr><tr><th>3</th><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",hyperlink("https://cases.connect.corp.google.com/#/case/"&amp;B3:B,"Cases")))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IF((U3:U="Assigned")*(B3:B&lt;&gt;""),IF((AJ3:AJ-NOW())&gt;0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>#REF!<br><span class="formula">==ARRAYFORMULA(IF((U3:U="Assigned")*(A3:A&lt;&gt;""),IF((AL3:AL-NOW())&gt;0, TEXT(AL3:AL-NOW(),"[h]:mm:ss"), "Missed"), ""))</span></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;<br><span class="formula">==ARRAYFORMULA(IFERROR(IFS(F3:F="Billing",2,F3:F="Policy",3),""))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(C3:C="","",C3:C-weekday(C3:C,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","","Phone"))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(F3:F="","",xlookup(AH3:AH&amp;T3:T&amp;AF3:AF,index!M:M,index!N:N))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AI3:AI))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(iferror(if(F3:F="","",xlookup(AH3:AH&amp;T3:T&amp;AF3:AF,index!M:M,index!O:O))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",C3:C+D3:D+AK3:AK))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(B3:B="","",if(W3:W&lt;&gt;"",0,if(U3:U="Assigned","",IF(X3:X=1,2,1)))))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM="","",if(AC3:AC="",Z3:Z,AC3:AC)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM="","",if(AD3:AD="",AA3:AA,AD3:AD)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AN3:AN="","",AN3:AN-weekday(AN3:AN,1)+1))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AJ3:AJ&gt;(AN3:AN+AO3:AO),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AL3:AL&gt;(AN3:AN+AO3:AO),1,0)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(AM3:AM&lt;&gt;1,"",if(AC3:AC="",0,1)))</span></td><td>&nbsp;<br><span class="formula">==arrayformula(if(S3:S="","",if(L3:L=S3:S,"",1)))</span></td></tr></tbody></table></div></div>

```

### フローデータモデル

```
Case オブジェクト {
  caseId: string,          // ユニークID
  date: string,            // 開始日付
  time: string,            // 開始時間
  segment: string,         // セグメント
  productCategory: string, // 製品カテゴリ
  issueCategory: string,   // 問題カテゴリ
  firstAssignee: string,   // 初回担当者
  finalAssignee: string,   // 最終担当者
  caseStatus: string,      // ステータス
  amTransfer: string,      // AM移管
  nonNCC: string,          // NCC除外理由
  bug: boolean,            // バグフラグ
  needInfo: boolean,       // 情報待ちフラグ
  tsConsulted: string,     // T&S相談状態
  firstCloseDate: string,  // 初回クローズ日
  firstCloseTime: string,  // 初回クローズ時間
  reopenReason: string,    // 再オープン理由
  reopenCloseDate: string, // 再クローズ日
  reopenCloseTime: string, // 再クローズ時間
}
```

# 分析要素：
```
Google から要求されたメトリックとして、各エージェントは以下のターゲットにMeetさせる必要があるため、これをダッシュボードで把握できる仕様にしたい。

  - Target：（TRT, Sentiment, NCC）Daily, Weekly, Monthly, Daily ごとにタブを切り替えて確認できる必要がある。
    - TRT に関する SLA Target：
        - P80 (Realtime（Live channel）：8 時間, Email：36 時間)
　　　 ⇒まずはこの時間以内にCaseをClose（SO, Finished）できることが理想
        - P95（Realtime/Email（ともに）：3 日以内にClose）/ P99（Realtime/EmailのT&S consult case：7 日以内にClose）
　　     ⇒もしP80のTargetをミスしたとしても、P95/P99（T&Sケース以外 3日以内）を必ずMeetできるようにする必要がある
    - Sentiment：Target 80%（これはGoogleのシステムがカスタマーとのやり取りを分析してSentiment Score として算出する仕様です。そのため、スコアが発表されたあと、各User（Ldap@google.com）が、 各自 Dashboard に手動入力できるようにする必要がある。）
    - NCC：7 / Daily Average （「Case ID」列が「空欄ではない」 AND 「Case Status」列が 「Assigned 以外」 AND 「non NCC」列が「空欄」 AND「Bug」列に「チェックが入っていない（値が0のとき）」に「Final Assignee」列のUser（Ldap）の NCC が 1 としてとして加算される）
  - 評価（表形式+色分け）：User管理画面では P95/P99 の TRT Timer（ 24時間以上(テキスト緑で背景濃いグレー)/24時間以下(テキスト黄色で背景濃いグレー)/8時間以下(テキスト赤で背景濃いグレー, 3時間以下からテキストを点滅させる)/Missed(テキスト白で背景濃いグレー) の自分のCase（「Case Status」列が「Assigned」）を一覧で確認できる。Admin管理画面でも各User の Case の P95/P99 の TRT Timer を確認できて、それとは別のエリアで TRT Timer が 8時間以下になったすべてのUserのCase ID と Ldap を一覧で表示できる機能
```

## 4. 機能詳細設計

### 4.1 ダッシュボード

### UI コンポーネント

- ケースカードグリッド
- ステータスインジケーター
- タイマー表示（2種類）
- アクションボタン（編集、削除、T&S切替）

### 動作フロー

1. ページロード時に担当ケースを取得
2. "Assigned" ステータスのケースのみフィルタリング表示
3. 1秒ごとにタイマーを更新
4. ユーザーアクションに応じた操作を実行

### 4.2 ケース管理

### 新規ケース登録フォーム

- 必須項目検証
- 入力補助（日付選択、ドロップダウン）
- 自動フィールド（担当者情報）

### ケース編集モーダル

- 基本情報表示（読み取り専用）
- 編集可能フィールド
- ステータスと任意項目の更新
- クローズ情報の管理

### 4.3 タイマー機能

### Response SLAタイマー

- スプレッドシートの関数と同等のロジック
    - `=ARRAYFORMULA(IF((U3:U="Assigned")*(B3:B<>""),IF((AJ3:AJ-NOW())>0, TEXT(AJ3:AJ-NOW(),"[h]:mm:ss"), "Missed"), ""))`
- デッドライン計算: `AJ3:AJ = C3:C + D3:D + AI3:AI`
- 残り時間の視覚的表示と警告色

### P95タイマー

- Case Openから72時間（デフォルト）
- T&S Consulted=Yesの場合は7日間に延長
- 日+時間形式での残り時間表示
- 残り時間による警告表示

### 4.4 統計分析

### データ集計

- 期間別集計（日次、週次、月次、四半期）
- ステータス別の内訳
- 解決率計算

### 視覚化

- 折れ線グラフ（時系列トレンド）
- 円グラフ（ステータス分布）
- 棒グラフ（期間比較）
- データテーブル（詳細数値）

## 5. 今後の設計・実装方針

### 5.1 短期実装計画（1-3ヶ月）

1. **コードリファクタリングとモジュール化**
    - 各機能モジュールの明確な分離
    - 共通ユーティリティの整理
    - コード品質向上とドキュメント化
2. **UI/UXの改善**
    - 動的なダッシュボードレイアウト
    - ドラッグ＆ドロップでのケース並べ替え
    - アニメーションとトランジションの洗練
    - モバイル対応の強化
3. **バグ修正と安定性向上**
    - エッジケースの処理改善
    - エラー報告機能
    - パフォーマンス最適化

### 5.2 中期開発計画（3-6ヶ月）

1. **高度な分析機能**
    - カスタムレポート生成
    - 予測分析（トレンド予測）
    - パフォーマンス指標のダッシュボード
2. **チーム機能の強化**
    - 担当者間のケース移管
    - チーム別ビュー
    - パフォーマンス比較
3. **インテグレーション拡張**
    - カレンダー連携
    - メール通知
    - Slackなどの通知連携

### 5.3 長期ビジョン（6ヶ月以上）

1. **AIサポート**
    - ケース分類の自動化
    - 解決時間予測
    - 類似ケース検索と解決策推奨
2. **拡張可能なプラットフォーム化**
    - プラグイン構造
    - カスタマイズ可能なダッシュボード
    - 複数チーム対応
3. **エンタープライズ機能**
    - ロールベースのアクセス制御
    - 詳細な監査ログ
    - バックアップと復元機能

### 5.4 技術的課題と対策

1. **GASの制限事項**
    - 実行時間制限: バッチ処理と非同期操作の実装
    - スクリプトプロパティのサイズ制限: 効率的なデータ構造設計
    - API呼び出し制限: キャッシュ戦略と呼び出し最適化
2. **スプレッドシートのスケーラビリティ**
    - 大量データ時のパフォーマンス: インデックス設計とクエリ最適化
    - 同時編集競合: ロック機構と衝突解決
    - アーカイブ戦略: 履歴データの効率的管理
3. **UI/UXの一貫性維持**
    - コンポーネントライブラリ: 再利用可能なUIパターン
    - スタイルガイド: デザイン規則の文書化
    - ユーザビリティテスト: 継続的なフィードバック収集

## 6. 開発およびデプロイプロセス

### 6.1 開発環境

- **ローカル開発**: clasp CLI（推奨）または直接GASエディタ
- **バージョン管理**: GitHub（`.claspignore`と`.gitignore`の適切な設定）
- **ブランチ戦略**: GitHub Flow（feature branching）

### 6.2 テストアプローチ

- **単体テスト**: GASユニットテストフレームワーク
- **統合テスト**: テスト用スプレッドシートでの機能検証
- **手動テスト**: チェックリストベースのUI/UXテスト
- **クロスブラウザテスト**: 最新のブラウザ互換性確認

### 6.3 デプロイプロセス

1. コード変更とテスト完了
2. プルリクエストとコードレビュー
3. マスターブランチへのマージ
4. 新バージョンとしてデプロイ
5. デプロイバージョンの更新とユーザー通知

### 6.4 保守とサポート

- 四半期ごとのスプレッドシート更新
- バグ追跡とイシュー管理（GitHub Issues）
- ユーザーフィードバックのトリアージと対応
- 機能リクエストの評価とロードマップ更新

## 7. セキュリティとコンプライアンス

### 7.1 データセキュリティ

- Google認証によるアクセス制御
- 組織内アクセス制限
- 機密データの取り扱い方針

### 7.2 エラーハンドリングとロギング

- 構造化されたエラーメッセージ
- ユーザーフレンドリーなエラー表示
- サーバーサイドログ記録

### 7.3 コンプライアンス考慮事項

- データ保持ポリシー
- アクセス監査
- バックアップと災害復旧

## 8. 付録

### 8.1 用語集

- **SLA (Service Level Agreement)**: サービス品質保証のための応答時間枠
- **P95**: ケース解決時間、標準は72時間
- **T&S Consulted**: Technical and Specialist consultation flag
- **NCC**: Non-Contact Complete（通常のクローズフロー以外）
- **AM Transfer**: Account Manager移管フラグ
- **Live Mode**: ポップアップウィンドウでの登録モード

### 8.2 参考リソース

- [Google Apps Script ドキュメント](https://developers.google.com/apps-script)
- [Material Design Components](https://material.io/components)
- [Chart.js ドキュメント](https://www.chartjs.org/docs/latest/)
- [Flatpickr ドキュメント](https://flatpickr.js.org/)

---

**文書バージョン**: 1.0.0

**最終更新日**: 2025年4月11日

**作成者**: CaseDash開発チーム