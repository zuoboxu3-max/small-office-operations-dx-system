# GAS 業務自動化（台帳→Drive/Calendar→ログ集約）ポートフォリオ

Google Apps Script を用いて、Googleスプレッドシート台帳を起点に
- 帳票テンプレートへの転記（Drive保存）
- Googleカレンダー登録
- 実行ログ/トリガ状態の可視化（管理者スプレッドシートに集約）
を自動化した業務支援システムの公開用ポートフォリオです。

> 公開版のため、実運用環境に含まれる個人情報・URL・トークン・ID・業務固有名称はダミー化/削除しています。

---

## 1. 全体像（アーキテクチャ）

本リポジトリは **2つのGASプロジェクト** を想定しています。

- frontend（台帳側 / スプレッドシートに紐づく）
  - メニューから「チェックした行」を一括処理
  - Driveへのテンプレ転記保存
  - Calendarへの予定登録
  - 実行ログ/トリガ状態を管理者Webアプリへ送信

- backend（管理者側 / Webアプリとしてデプロイ）
  - doPost でログを受信し、管理者スプレッドシートに記録
  - TriggerStatus や FolderSizeMonthly などの運用可視化

データフロー（概略）：

[台帳Spreadsheet]
   | (メニュー実行)
   v
[frontend GAS] --(UrlFetch: log/trigger_status)--> [backend WebApp(doPost)] --> [管理者Spreadsheet(Log等)]
   | (Drive/Calendar API)
   v
[Drive / Calendar]

---

## 2. 主な機能

### 台帳側（frontend）
- チェック行の一括処理（Drive転記 + Calendar登録）
- テンプレスプレッドシートの複製・転記（マッピング定義で拡張可能）
- 実行状況のログ送信（管理者側に集約）

### 管理者側（backend）
- ログ受信（doPost）→ Logシートへ追記
- トリガ状態の受信/反映（TriggerStatus）
- 月末のみフォルダ容量を記録（FolderSizeMonthly）

---

## 3. セットアップ（公開版向け）

### 3.1 backend（管理者Webアプリ）
1) backend のコードをGASプロジェクトに配置
2) 管理者スプレッドシートを用意し、スクリプトを紐づけ
3) 初期導入ウィザード（メニュー）で
   - WebアプリURL
   - 保存先フォルダURL/ID
   - （任意）移行元フォルダURL/ID
   を設定
4) Webアプリとしてデプロイ（doPostを有効化）

### 3.2 frontend（台帳側）
1) 台帳スプレッドシートに frontend のコードを配置
2) 設定（フォルダID、カレンダーID、テンプレID）を差し替え
3) 台帳を開き、メニューから一括処理を実行

---

## 4. 設定値（ダミー化について）
本リポジトリには以下のプレースホルダが含まれます：

- YOUR_PARENT_FOLDER_ID
- your-calendar@example.com
- YOUR_TEMPLATE_SPREADSHEET_ID_*
- YOUR_ADMIN_WEBAPP_URL / YOUR_ADMIN_WEBAPP_TOKEN

実運用では、URL/Token等は PropertiesService に保持し、
コード直書きしない運用を想定しています。

---

## 5. 制約・注意事項
- Google Apps Script の実行時間制限を考慮し、一部処理はタイムアウト回避の設計を含みます。
- 公開版は構成理解を優先し、業務固有の分類/命名ルールは抽象化しています。

---

## 6. ディレクトリ構成

```text
small-office-operations-dx-system/
├─ README.md
├─ frontend/
│  ├─ AddonMain.gs.js        # 台帳側メイン（メニュー/一括処理/Drive/Calendar）
│  ├─ LoggerClient.gs.js     # 管理者Webアプリへ送信（UrlFetch）
│  └─ appsscript.json
└─ backend/
   ├─ AdminSetup...js        # 管理者メニュー/ログ共通/プロパティ管理
   ├─ AdminWebApp...js       # doPost受信口（ログ/トリガ状態）
   ├─ AdminWizard...js       # 初期導入ウィザード
   ├─ FolderSizeMonthly...js # 月末容量チェック + トリガ
   ├─ Organize50On.js        # フォルダ階層整理（50音相当の分類）
   └─ appsscript.json
