# 機械設計技術部 タスク管理アプリ 設計手順書 v1.0

> 作成日: 2026-05-08（最終レビュー反映: 2026-05-08 / Opus 修正適用）
> 対象バージョン: v1.0
> **一次情報源: `docs/REQUIREMENTS_v1.0.md`（確定設計仕様書）と `docs/wireframe.html`（UI 正本）**
> v0.1〜v0.5 は履歴用。実装中に仕様の不明点が出た場合は v1.0 を最優先で参照すること。

---

## 1. このドキュメントの位置づけ

本書は「機械設計技術部 タスク管理アプリ」を**最初から運用開始まで**実装するための、ステップバイステップの手順書です。

### 本書が対象とする読者
- GAS（Google Apps Script）の基本操作ができる実装者
- Python の基本的な開発経験がある担当者
- 本アプリを初めて実装する方

### 関連ドキュメント
| ドキュメント | 役割 |
|---|---|
| `docs/REQUIREMENTS_v1.0.md` | **機能仕様書（一次情報源・本書の姉妹ドキュメント）** |
| `docs/REQUIREMENTS_v0.1〜v0.5.md` | 過去履歴（参照不要・改変禁止） |
| `docs/wireframe.html` | UI デザイン正本（実装時に常時参照） |
| `src/Config.gs` | GAS 定数・スキーマの中央定義 |

### 実装に当たっての絶対原則
1. **デザインは `docs/wireframe.html` に厳密準拠**する。配色・タイポ・コンポーネント名を変更しない。
2. **仕様変更は `docs/REQUIREMENTS_v1.0.md` を確認**してから実装する。
3. **破壊的操作の前に必ずバックアップ**を取る（スプレッドシートのコピー）。
4. **EXE は Windows 専用**（Outlook COM API 使用）。Mac/Linux 環境での動作は対象外。

---

## 2. 全体ロードマップ

### 2.1 実装フェーズ分け

```
フェーズ 1: 基盤整備（STEP 1〜2）
  └─ GAS スカフォールド + シート生成 + マスタ初期投入

フェーズ 2: バックエンド（STEP 3〜5）
  └─ DataService（CRUD）+ MailQueueService + 外部マスタ連携

フェーズ 3: フロントエンド（STEP 6）
  └─ UI 実装（wireframe 分解 → ガント・日報・モーダル）

フェーズ 4: EXE 開発（STEP 7〜8）
  └─ Python メールエージェント + PyInstaller ビルド + 配布

フェーズ 5: 検証・リリース（STEP 9）
  └─ 結合テスト → デプロイ → 受入テスト
```

### 2.2 想定スケジュール（目安）

| フェーズ | STEP | 所要日数（目安） |
|---|---|---|
| フェーズ 1 | STEP 1〜2 | 1〜2 日 |
| フェーズ 2 | STEP 3〜5 | 3〜5 日 |
| フェーズ 3 | STEP 6 | 5〜7 日 |
| フェーズ 4 | STEP 7〜8 | 3〜5 日 |
| フェーズ 5 | STEP 9 + デプロイ | 2〜3 日 |
| **合計** | | **約 2〜3 週間** |

> 各スタッフの PC へのタスクスケジューラ登録は運用担当者が実施するため、EXE 配布後の作業は含まない。

---

## 3. 環境準備

### 3.1 開発端末: Node.js + clasp + Git

**必要バージョン**
- Node.js: **20.x LTS 以上**
- npm: 10.x 以上（Node.js に同梱）
- clasp: 2.4.x 以上
- Git: 2.40 以上

**インストール手順**

```bash
# 1. Node.js インストール確認
node --version   # v20.x.x 以上を確認

# 2. clasp をグローバルインストール
npm install -g @google/clasp

# 3. clasp のバージョン確認
clasp --version  # 2.4.x 以上を確認

# 4. Google アカウントでログイン
clasp login
# ブラウザが開く → Google アカウントでログイン → 認証完了
```

**プロジェクトのクローン（初回のみ）**

```bash
# 作業ディレクトリへ移動
cd "C:\Users\<ユーザー名>\Documents"

# リポジトリクローン（GitHub リポジトリ作成後）
git clone https://github.com/<組織名>/team-task-management.git
cd team-task-management
```

> 既に `C:\Users\imaizumi.LINEWORKS-NET\Documents\Team Task Management` に作業ディレクトリが存在する場合はクローン不要。

**`.clasp.json` 確認**

```json
{
  "scriptId": "1v1P1s5T1L9E7snpsRQT4Wpm2qJTzvt0-kkmgEZ1ec-81lUW0qdBVMlDX",
  "rootDir": "./src"
}
```

> `scriptId` と `rootDir` がこの値であることを確認する。

### 3.2 開発端末: Python 3.11 + 必要パッケージ

**必要バージョン**
- Python: **3.11.x 以上**（3.12 も可）
- pip: 23.x 以上

**インストール確認**

```powershell
python --version   # Python 3.11.x 以上を確認
pip --version
```

**仮想環境と依存パッケージ**

```powershell
# exe ディレクトリへ移動
cd "C:\Users\<ユーザー名>\Documents\Team Task Management\exe\mail_agent"

# 仮想環境作成
python -m venv .venv

# 仮想環境を有効化（PowerShell）
.\.venv\Scripts\Activate.ps1

# 依存パッケージのインストール
pip install -r requirements.txt
```

**`requirements.txt` の内容**

```
requests>=2.31.0
pywin32>=306
pyinstaller>=6.3.0
```

### 3.3 GitHub リポジトリ

```bash
# GitHub にリポジトリを作成後、リモートを設定
git remote add origin https://github.com/<組織名>/team-task-management.git
git branch -M main
git push -u origin main
```

**推奨ブランチ運用**
```
main         本番リリースブランチ
develop      開発統合ブランチ
feature/xxx  機能別作業ブランチ
```

### 3.4 Google Apps Script プロジェクト紐付け

**前提**: GAS プロジェクトが対象のスプレッドシートにバインドされていること。

```bash
# clasp でプロジェクトをプル（既存 GAS ファイルをローカルに取得）
clasp pull

# ローカルファイルを GAS へプッシュ
clasp push

# GAS エディタをブラウザで開く
clasp open
```

**GAS プロジェクト設定の確認**
1. GAS エディタ → 「プロジェクトの設定」
2. 「Chrome V8 ランタイムを使用する」が **ON** であることを確認
3. タイムゾーンが `Asia/Tokyo` であることを確認

### 3.5 共有フォルダ（EXE 配布用）

**フォルダ構成の作成**

社内の共有ファイルサーバーまたは Google Drive の共有フォルダに以下を作成する。

```
\\<サーバー名>\共有\タスク管理アプリ\
  └─ releases\
       └─ v1.0.0\
            └─ TeamTaskMail.exe
       └─ v1.1.0\  ← 将来のバージョン用
  └─ README.txt  ← インストール手順（各 PC 担当者向け）
```

**`README.txt` の記載内容**（作成すること）

```
機械設計技術部 タスク管理アプリ — EXE セットアップ手順

1. releases\<最新バージョン>\TeamTaskMail.exe を
   C:\ProgramData\TeamTaskMail\ にコピーする
2. タスクスケジューラに登録する（5章 5.4 参照）
3. 動作確認: スタートメニューから「TeamTaskMail」を検索して実行
```

### 3.6 スプレッドシート（バインド + 外部マスタ）の確認

**バインドスプレッドシート**
- GAS プロジェクトがバインドされているスプレッドシートを開く
- シート一覧に `Schedules`, `DailyReports`, `MailQueue`, `MailLog`, `Staff`, `WorkTypes`, `Settings` があることを確認（初回は `setupAll()` 実行後）

**外部マスタスプレッドシート**
- スプレッドシート ID: `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`
- シート構成:
  - `工番マスタ` シート（約 844 行）
  - `社内カレンダーマスタ` シート
- GAS のサービスアカウントまたは実行ユーザーに **閲覧権限** が付与されていることを確認

---

## 4. 実装ステップ詳細

---

### STEP 1: GAS スカフォールド整備

**所要時間目安**: 1〜2 時間

**目的**

`src/` ディレクトリ配下の GAS ファイル群を整備し、clasp で正常にプッシュ・デプロイできる状態にする。

**前提条件**
- 3.1〜3.4 の環境準備が完了していること
- GAS エディタにアクセスできること

**作業内容**

1. **現状の `src/` ファイル構成を確認する**

```
src/
  appsscript.json   ← ランタイム設定（V8 / タイムゾーン / DOMAIN アクセス）
  Code.gs           ← doGet() エントリーポイント
  Config.gs         ← 全定数・スキーマ定義
  SetupService.gs   ← setupAll() / setupInitialSheets() 等
```

2. **`Config.gs` の WorkTypes 初期データを v1.0 仕様（13項目・カテゴリなし）へ更新する**

`src/Config.gs` の `DEFAULT_WORK_TYPES` を以下に置き換える:

```javascript
const DEFAULT_WORK_TYPES = [
  ['wt01', '構想図作成',                1, true],
  ['wt02', '承認図作成',                2, true],
  ['wt03', 'バラシ図作成',              3, true],
  ['wt04', '第三者チェック',            4, true],
  ['wt05', '材料取り・仕様検討',        5, true],
  ['wt06', '設計検討',                  6, true],
  ['wt07', '出図準備',                  7, true],
  ['wt08', '出荷後図面修正',            8, true],
  ['wt09', '購入部品手配・在庫部品確認', 9, true],
  ['wt10', '試運転調整・動作確認',     10, true],
  ['wt11', '加工指示',                 11, true],
  ['wt12', '現場対応・現場工事対応',   12, true],
  ['wt13', '出張・打ち合わせ',         13, true],
];
```

> `category` 列は v1.0 で廃止。`SHEET_SCHEMA.WorkTypes` の列定義も `['id', 'name', 'displayOrder', 'active']` に変更する（既存の `'category'` を必ず削除）。

3. **`Config.gs` の `SHEET_NAMES` から `HOLIDAYS` を削除し、`SHEET_SCHEMA` の対応エントリも削除する**

社内カレンダーは外部マスタから参照するため、バインド SS 内に `Holidays` シートは不要。

4. **`Config.gs` の `SHEET_SCHEMA[MAIL_QUEUE]` に `previousRequestId` 列を追加する**

仕様 v1.0 §6.3 / §7.4.6（再送）に準拠し、再送チェーン追跡用の列を追加する。`createdAt` の**直前**に挿入する。

```javascript
[SHEET_NAMES.MAIL_QUEUE]: [
  'id',
  'requestedBy',
  'targetStaffId',
  'targetStaffName',
  'targetStaffEmail',
  'reportDate',
  'mode',              // draft / send
  'toAddresses',       // ★カンマ区切り string（"a@x,b@y"）。JSON 配列は使わない
  'ccAddresses',       // v1.0 では常に空文字列。将来拡張用
  'subjectVars',       // JSON string
  'bodyVars',          // JSON string
  'status',            // pending / picked / drafted / sent / failed
  'pickedBy',
  'pickedAt',
  'processedAt',
  'errorMessage',
  'previousRequestId', // ★再送時に元レコード id を設定
  'createdAt'
],
```

5. **`Config.gs` の `SHEET_SCHEMA[DAILY_REPORTS]` に `createdAt` 列を追加する**

仕様 v1.0 §6.3 / §9（非機能要件・監査）に準拠し、`updatedAt` の**直前**に挿入する。

```javascript
[SHEET_NAMES.DAILY_REPORTS]: [
  'id',
  'staffId',
  'reportDate',
  'section',
  'seq',
  'periodStart',
  'periodEnd',
  'kobanCode',
  'workTypeId',
  'detail',
  'linkedScheduleId',
  'createdAt',  // ★追加
  'updatedAt'
],
```

6. **`Config.gs` の `DEFAULT_SETTINGS` に `LAST_DEAD_NOTIFICATION_AT` 行を追加する**

死活監視の重複通知抑制（仕様 §7.5）のために必要。

```javascript
const DEFAULT_SETTINGS = [
  ['KOBAN_MASTER_SHEET_ID',     '',        '工番マスタ別ブックのスプレッドシートID'],
  ['KOBAN_MASTER_SHEET_NAME',   '工番マスタ', '工番マスタブック内のシート名'],
  ['POLLING_INTERVAL_SECONDS',  '30',      'ローカル EXE が MailQueue をポーリングする間隔（秒）'],
  ['MAIL_DEFAULT_MODE',         'draft',   'メール送信モードの既定値'],
  ['WEBAPP_VERSION',            '1.0.0',   'WebApp バージョン'],
  ['LAST_HEARTBEAT_TIMESTAMP',  '',        'ローカル EXE からの最終ハートビート（自動更新）'],
  ['LAST_HEARTBEAT_HOSTNAME',   '',        '最終ハートビートを送ってきた PC 名（自動更新）'],
  ['LAST_DEAD_NOTIFICATION_AT', '',        'EXE 死活通知を最後に送信した日時（6時間以内は再通知抑制）'],
  ['ADMIN_EMAIL',               '',        '管理者メールアドレス（障害通知先）'],
];
```

7. **`appsscript.json` を確認する**

`GmailApp.sendEmail()`（管理者通知メールで使用）と外部スプレッドシート参照のため、必要 OAuth スコープを明示する。

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.send_mail"
  ],
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "DOMAIN"
  }
}
```

> `script.send_mail` を含めないと `GmailApp.sendEmail` が権限エラーとなる。`spreadsheets`（広域）は外部マスタ `openById` のために必須。

8. **clasp push でエラーがないことを確認する**

```bash
clasp push
# → 「Pushed N files.」と表示されれば OK
```

9. **GAS エディタで `setupAll()` を実行する**

GAS エディタ → 「実行」ボタン → 関数に `setupAll` を指定 → 実行  
「実行ログ」に `=== Setup complete ===` が表示されれば完了。

**完了確認**
- [ ] `clasp push` がエラーなく完了する
- [ ] GAS エディタで `setupAll()` が正常終了する
- [ ] スプレッドシートに 7 シート（`Schedules` / `DailyReports` / `MailQueue` / `MailLog` / `Staff` / `WorkTypes` / `Settings`）が生成されている
- [ ] `MailQueue` のヘッダ行に **`previousRequestId`** 列がある（再送機能で必須）
- [ ] `MailQueue` の `toAddresses` がカンマ区切り文字列を保存できる単純な文字列列である（JSON 配列ではない）
- [ ] `DailyReports` のヘッダ行に **`createdAt`** 列がある
- [ ] `WorkTypes` シートに確定13項目が投入されている（`category` 列が無いこと）
- [ ] `Settings` シートに初期値 9 行が投入されている（`LAST_DEAD_NOTIFICATION_AT` を含む）
- [ ] `MailQueue_Archive` シートは初期生成不要（`archiveOldMailQueue()` 初回実行時に自動生成。仕様 §11 参照）

**よくあるエラーと対処**

- エラー: `clasp push` で「認証エラー」  
  対処: `clasp login` を再実行し、Google アカウントで再認証する。

- エラー: `setupAll` 実行時「SpreadsheetApp.getActiveSpreadsheet() が null」  
  対処: GAS プロジェクトがスプレッドシートにバインドされていない。スプレッドシートの「拡張機能 → Apps Script」からプロジェクトを開き直す。

- エラー: `WorkTypes` シートが既存で category 列がある  
  対処: `resetAllSheets()` を使って再生成するか、手動で WorkTypes シートを削除後に `setupInitialMasters()` を再実行する。

---

### STEP 2: SetupService（シート作成 + 初期マスタ投入）

**所要時間目安**: 1〜2 時間

**目的**

バインドスプレッドシートのシート構成とマスタデータが完全・正確であることを確認し、以後の STEP で安全に参照できる状態にする。

**前提条件**
- STEP 1 が完了し、7 シートが生成されていること

**作業内容**

1. **`Staff` シートへスタッフ情報を入力する**

シートを開き、`id`, `name`, `email`, `role`, `displayOrder`, `active` の各列に 6 名分を入力する:

| id | name | email | role | displayOrder | active |
|---|---|---|---|---|---|
| staff01 | 氏名1 | email1@example.com | staff | 1 | TRUE |
| ... | ... | ... | ... | ... | ... |
| staff06 | 管理者名 | admin@example.com | admin | 6 | TRUE |

> `active=FALSE` にするとメール送信対象から除外される。

2. **`Settings` シートの `KOBAN_MASTER_SHEET_ID` を入力する**

`KOBAN_MASTER_SHEET_ID` 行の `value` 列に外部マスタの ID を入力:  
`1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`

3. **外部マスタへのアクセス権を確認する**

GAS エディタで以下のテスト関数を作成・実行して、外部マスタが読めることを確認する:

```javascript
function testExternalMaster() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('KOBAN_MASTER_SHEET_ID');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('工番マスタ');
  const lastRow = sheet.getLastRow();
  Logger.log('工番マスタ行数: ' + lastRow); // 約 845 以上（ヘッダ含む）
}
```

4. **`EXE_API_TOKEN` をスクリプトプロパティから取得し、控えておく**

GAS エディタ → 「プロジェクトの設定」→「スクリプトプロパティ」  
`EXE_API_TOKEN` の値を安全な場所（パスワードマネージャー等）に控える。  
この値は STEP 7 の Python EXE 設定で使用する。

**完了確認**
- [ ] `Staff` シートに 6 名分のデータが入力されている
- [ ] `Settings` の `KOBAN_MASTER_SHEET_ID` に値が入っている
- [ ] テスト関数で外部マスタから行数が取得できる
- [ ] `EXE_API_TOKEN` を安全に控えた

**よくあるエラーと対処**

- エラー: `openById` で「権限が必要です」  
  対処: 外部マスタのスプレッドシートの「共有」設定で、GAS を実行するユーザーへ「閲覧者」権限を付与する。

- エラー: `getSheetByName` が null を返す  
  対処: 外部マスタの「工番マスタ」シート名が完全一致しているか確認する（全角スペース等に注意）。

---

### STEP 3: DataService（CRUD）

**所要時間目安**: 2〜3 時間

**目的**

スプレッドシートへの CRUD 操作を担う `DataService.gs` を実装する。UI および MailQueueService から呼び出されるすべてのデータ操作をこのサービスに集約する。

**前提条件**
- STEP 1〜2 が完了していること

**作業内容**

1. **`src/DataService.gs` を新規作成する**

以下の関数群を実装する:

```javascript
/**
 * DataService.gs
 * スプレッドシートへの CRUD 操作を提供する。
 * すべての関数は GAS の実行コンテキスト（バインド SS）で動作する。
 */

// ── 汎用ヘルパ ──────────────────────────────────────────────

/**
 * 指定シートの全データを [{列名: 値, ...}, ...] 形式で返す。
 */
function getSheetData_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row[i]]))
  );
}

/**
 * 新規行を追加し、付与した id を返す。
 */
function appendRow_(sheetName, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => record[h] !== undefined ? record[h] : '');
  sheet.appendRow(row);
  return record.id;
}

/**
 * id が一致する行を更新する。
 */
function updateRow_(sheetName, id, updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      headers.forEach((h, j) => {
        if (updates[h] !== undefined) sheet.getRange(i + 1, j + 1).setValue(updates[h]);
      });
      return true;
    }
  }
  return false;
}

// ── Schedules ────────────────────────────────────────────────

function getSchedules(startDate, endDate) {
  return getSheetData_(SHEET_NAMES.SCHEDULES).filter(r =>
    r.startDate && r.endDate &&
    r.endDate >= startDate && r.startDate <= endDate
  );
}

function createSchedule(record) {
  record.id = 'sch_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  record.createdAt = new Date().toISOString();
  record.updatedAt = record.createdAt;
  appendRow_(SHEET_NAMES.SCHEDULES, record);
  return record.id;
}

function updateSchedule(id, updates) {
  updates.updatedAt = new Date().toISOString();
  return updateRow_(SHEET_NAMES.SCHEDULES, id, updates);
}

function deleteSchedule(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SCHEDULES);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === id) { sheet.deleteRow(i + 1); return true; }
  }
  return false;
}

// ── DailyReports ─────────────────────────────────────────────

function getDailyReports(staffId, reportDate) {
  return getSheetData_(SHEET_NAMES.DAILY_REPORTS).filter(r =>
    r.staffId === staffId && r.reportDate === reportDate
  );
}

function upsertDailyReport(record) {
  const existing = getDailyReports(record.staffId, record.reportDate)
    .find(r => r.seq === record.seq && r.section === record.section);
  if (existing) {
    record.updatedAt = new Date().toISOString();
    return updateRow_(SHEET_NAMES.DAILY_REPORTS, existing.id, record);
  } else {
    record.id = 'dr_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
    record.updatedAt = new Date().toISOString();
    appendRow_(SHEET_NAMES.DAILY_REPORTS, record);
    return record.id;
  }
}

// ── Staff ─────────────────────────────────────────────────────

function getActiveStaff() {
  return getSheetData_(SHEET_NAMES.STAFF)
    .filter(r => r.active === true || r.active === 'TRUE')
    .sort((a, b) => Number(a.displayOrder) - Number(b.displayOrder));
}

// ── WorkTypes ─────────────────────────────────────────────────

function getWorkTypes() {
  return getSheetData_(SHEET_NAMES.WORK_TYPES)
    .filter(r => r.active === true || r.active === 'TRUE')
    .sort((a, b) => Number(a.displayOrder) - Number(b.displayOrder));
}

// ── Settings ──────────────────────────────────────────────────

function getSetting(key) {
  const row = getSheetData_(SHEET_NAMES.SETTINGS).find(r => r.key === key);
  return row ? row.value : null;
}

function setSetting(key, value) {
  return updateRow_(SHEET_NAMES.SETTINGS, key, { value });
}
```

2. **外部マスタ読込関数を実装する**（`src/MasterService.gs` を新規作成）

```javascript
/**
 * MasterService.gs
 * 外部マスタ（工番・社内カレンダー）の読込を担当する。
 * 起動時に全件ロードし、フロントへ JSON で配信する。
 */

function loadKobanMaster_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_ID);
  const sheetName = props.getProperty(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_NAME) || '工番マスタ';
  if (!id) return [];
  const sheet = SpreadsheetApp.openById(id).getSheetByName(sheetName);
  if (!sheet) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  // 表示しない列（数量・取込日時）は除外しない（フロント側でフィルタ）
  return rows.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}

function loadCalendarMaster_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_ID);
  if (!id) return [];
  const sheet = SpreadsheetApp.openById(id).getSheetByName('社内カレンダーマスタ');
  if (!sheet) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}
```

**完了確認**
- [ ] `DataService.gs` が GAS エディタに表示されている
- [ ] `MasterService.gs` が GAS エディタに表示されている
- [ ] `getActiveStaff()` を手動実行して Staff シートのデータが返る
- [ ] `getWorkTypes()` を手動実行して 13 項目が返る
- [ ] `loadKobanMaster_()` を手動実行して工番データが返る（約 844 件）

**よくあるエラーと対処**

- エラー: 実行時に「RangeError: Maximum call stack size exceeded」  
  対処: 無限再帰になっていないか確認する。`getSheetData_` 内で sheet が null の場合のガードを確認する。

- エラー: `getDataRange()` でエラー  
  対処: シートが完全に空の場合（ヘッダもない）に発生する。STEP 1 の `setupAll()` を再実行してヘッダ行を生成する。

---

### STEP 4: MailQueueService（API + 重複防止 + 再送）

**所要時間目安**: 3〜4 時間

**目的**

MailQueue シートを介した EXE エージェントとの通信 API を実装する。**重複送信防止**のための atomic CAS（Compare-And-Swap）と、失敗したキューの**再送機能**を含む。

**前提条件**
- STEP 3 が完了していること

**作業内容**

1. **`src/MailQueueService.gs` を新規作成する**

```javascript
/**
 * MailQueueService.gs
 *
 * EXE エージェント向け API 群:
 *   - enqueueMailRequest()  : UI → MailQueue に pending 追加
 *   - pickMailItem()        : EXE → CAS で pending 1 件を picked に遷移（重複防止の核心）
 *   - completeMailItem()    : EXE → drafted / sent / failed に更新
 *   - retryMailItem()       : 元レコードを複製して新規 pending を生成（仕様 §7.4.6 準拠 / 元レコードは不変）
 *   - getMailQueueStatus()  : UI → MailQueue の現状一覧
 *   - recoverStalePicked()  : タイムアウト復旧（GAS トリガーで定期実行）
 */

// ── キュー追加 ──────────────────────────────────────────────

/**
 * メール送信リクエストを MailQueue に追加する。
 * 同日・同スタッフの pending/picked が既にある場合は二重登録を防ぐ。
 * @param {Object} params - requestedBy, targetStaffId, reportDate, mode, toAddresses, subjectVars, bodyVars
 */
function enqueueMailRequest(params) {
  // 仕様 §7.4: 「複数回送信は制限なし」のため、ここで未処理重複チェックは行わない。
  // ただし pending/picked が既にある状態で連打されたら UI 側で軽い注意は出すこと（重複登録自体は許容）。
  // ※ 過去ガードコードが必要ならここに残してもよいが、レスポンスは success=true で返す。

  // toAddresses は **カンマ区切り string** で保存する（仕様 v1.0 §6.3 確定）。
  // 入力が配列の場合は join、文字列の場合はそのまま使う。
  const toAddrCsv = Array.isArray(params.toAddresses)
    ? params.toAddresses.filter(Boolean).join(',')
    : (params.toAddresses || '');

  const record = {
    id:                'mq_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16),
    requestedBy:       params.requestedBy,
    targetStaffId:     params.targetStaffId,
    targetStaffName:   params.targetStaffName  || '',
    targetStaffEmail:  params.targetStaffEmail || '',
    reportDate:        params.reportDate,
    mode:              params.mode || 'draft',
    toAddresses:       toAddrCsv,                               // ★カンマ区切り string
    ccAddresses:       '',                                       // ★v1.0 では常に空文字列
    subjectVars:       JSON.stringify(params.subjectVars || {}),
    bodyVars:          JSON.stringify(params.bodyVars  || {}),
    status:            'pending',
    pickedBy:          '',
    pickedAt:          '',
    processedAt:       '',
    errorMessage:      '',
    previousRequestId: '',                                       // ★初回送信時は空。再送時のみ retryMailItem が設定
    createdAt:         new Date().toISOString()
  };
  appendRow_(SHEET_NAMES.MAIL_QUEUE, record);
  return { success: true, id: record.id };
}

// ── atomic CAS: pending → picked ────────────────────────────

/**
 * EXE が pending 状態の 1 件を取得し picked に遷移させる。
 * LockService によりスプレッドシートレベルで排他制御する（重複送信防止の核心）。
 *
 * @param {string} exeHostname - 取得する EXE の識別子（PC 名）
 * @returns {Object|null} picked したレコード、または null（取得できなかった場合）
 */
function pickMailItem(exeHostname) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10 秒待機、取れなければ例外

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MAIL_QUEUE);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusIdx  = headers.indexOf('status');
    const pickedByIdx = headers.indexOf('pickedBy');
    const pickedAtIdx = headers.indexOf('pickedAt');
    const idIdx      = headers.indexOf('id');

    for (let i = 1; i < data.length; i++) {
      if (data[i][statusIdx] === 'pending') {
        // CAS: この行を即座に picked に更新
        sheet.getRange(i + 1, statusIdx  + 1).setValue('picked');
        sheet.getRange(i + 1, pickedByIdx + 1).setValue(exeHostname);
        sheet.getRange(i + 1, pickedAtIdx + 1).setValue(new Date().toISOString());
        SpreadsheetApp.flush(); // 即座に書き込み

        // 更新後のレコードを返す
        const row = sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0];
        return Object.fromEntries(headers.map((h, j) => [h, row[j]]));
      }
    }
    return null; // pending がなかった

  } finally {
    lock.releaseLock();
  }
}

// ── 完了・失敗更新 ───────────────────────────────────────────

/**
 * EXE がメール処理後に状態を更新する。
 * @param {string} id - MailQueue レコード ID
 * @param {string} status - 'drafted' | 'sent' | 'failed'
 * @param {string} [errorMessage] - 失敗時のエラーメッセージ
 */
function completeMailItem(id, status, errorMessage) {
  const updates = {
    status:       status,
    processedAt:  new Date().toISOString(),
    errorMessage: errorMessage || ''
  };
  const result = updateRow_(SHEET_NAMES.MAIL_QUEUE, id, updates);

  // MailLog に記録
  logMailEvent_(id, status === 'failed' ? 'error' : 'info', status, errorMessage || '');

  // 失敗時は管理者へ通知メール
  if (status === 'failed') {
    notifyAdminOnFailure_(id, errorMessage);
  }
  return result;
}

// ── 再送 ─────────────────────────────────────────────────────

/**
 * 既存の MailQueue レコードを **複製** し、新レコードを pending として追加する（仕様 §7.4.6 準拠）。
 *
 * 重要:
 *   - 元レコードは一切変更しない（履歴を破壊しない）
 *   - status='sent' のときの「本当に再送しますか？」確認は UI 側で先に表示し、ユーザが OK した上で本関数を呼ぶ
 *   - 新レコードは元と同じ targetStaff/mode/toAddresses/subjectVars/bodyVars を持ち、
 *     previousRequestId に元レコード id をセットして連鎖を作る
 *
 * @param {string} originalId - 元 MailQueue レコードの id
 * @returns {{success:true, newId:string} | {success:false, reason:string}}
 */
function retryMailItem(originalId) {
  const original = getSheetData_(SHEET_NAMES.MAIL_QUEUE).find(r => r.id === originalId);
  if (!original) {
    return { success: false, reason: 'NOT_FOUND' };
  }

  const newRecord = {
    id:                'mq_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16),
    requestedBy:       original.requestedBy,
    targetStaffId:     original.targetStaffId,
    targetStaffName:   original.targetStaffName,
    targetStaffEmail:  original.targetStaffEmail,
    reportDate:        original.reportDate,
    mode:              original.mode,
    toAddresses:       original.toAddresses,   // 元のままコピー（カンマ区切り string）
    ccAddresses:       original.ccAddresses,
    subjectVars:       original.subjectVars,
    bodyVars:          original.bodyVars,
    status:            'pending',
    pickedBy:          '',
    pickedAt:          '',
    processedAt:       '',
    errorMessage:      '',
    previousRequestId: originalId,             // ★連鎖
    createdAt:         new Date().toISOString()
  };
  appendRow_(SHEET_NAMES.MAIL_QUEUE, newRecord);
  logMailEvent_(originalId, 'info', 'retry',
                '再送: 新レコード ' + newRecord.id + ' を生成');
  logMailEvent_(newRecord.id, 'info', 'retry',
                'previousRequestId=' + originalId + ' から複製');

  return { success: true, newId: newRecord.id };
}

// ── ステータス取得 ────────────────────────────────────────────

function getMailQueueStatus(reportDate) {
  const rows = getSheetData_(SHEET_NAMES.MAIL_QUEUE);
  return reportDate ? rows.filter(r => r.reportDate === reportDate) : rows;
}

// ── タイムアウト復旧 ──────────────────────────────────────────

/**
 * picked のまま 10 分以上経過したレコードを pending に戻す。
 * GAS のタイムベーストリガーで 5 分ごとに実行する。
 */
function recoverStalePicked() {
  const TIMEOUT_MS = 10 * 60 * 1000; // 10 分
  const now = new Date();
  const rows = getSheetData_(SHEET_NAMES.MAIL_QUEUE);
  let recovered = 0;

  rows.forEach(r => {
    if (r.status !== 'picked') return;
    const pickedAt = r.pickedAt ? new Date(r.pickedAt) : null;
    if (!pickedAt || (now - pickedAt) < TIMEOUT_MS) return;

    updateRow_(SHEET_NAMES.MAIL_QUEUE, r.id, {
      status:   'pending',
      pickedBy: '',
      pickedAt: ''
    });
    logMailEvent_(r.id, 'warn', 'mismatch', 'タイムアウト復旧: ' + r.pickedBy);
    recovered++;
  });
  Logger.log('recoverStalePicked: ' + recovered + ' 件を復旧');
}

// ── 内部ヘルパ ────────────────────────────────────────────────

function logMailEvent_(mailQueueId, level, event, message) {
  const record = {
    id:          'ml_' + Utilities.getUuid().replace(/-/g, '').substring(0, 12),
    mailQueueId: mailQueueId,
    timestamp:   new Date().toISOString(),
    level:       level,
    event:       event,
    message:     message
  };
  appendRow_(SHEET_NAMES.MAIL_LOG, record);
}

function notifyAdminOnFailure_(mailQueueId, errorMessage) {
  const adminEmail = getActiveStaff().find(s => s.role === 'admin');
  if (!adminEmail) return;
  try {
    GmailApp.sendEmail(
      adminEmail.email,
      '[タスク管理] メール送信失敗: ' + mailQueueId,
      'MailQueue ID: ' + mailQueueId + '\nエラー: ' + errorMessage
    );
  } catch (e) {
    Logger.log('管理者通知メール送信失敗: ' + e.message);
  }
}

// ── EXE 死活アラート（仕様 §7.5：6 時間以内の重複通知を抑制） ──────

/**
 * EXE が 5 分以上応答しない場合に管理者へ 1 回だけ通知する。
 * 1 分間隔の TimeTrigger（または UI からの呼び出し）で実行する想定。
 */
function checkExeAlive() {
  const last = getSetting('LAST_HEARTBEAT_TIMESTAMP');
  if (!last) return;
  const elapsedMin = (Date.now() - new Date(last).getTime()) / 60000;
  if (elapsedMin < 5) return; // 正常範囲

  // 重複通知抑制: 6 時間以内に通知済みならスキップ
  const lastNotified = getSetting('LAST_DEAD_NOTIFICATION_AT');
  if (lastNotified) {
    const sinceLastMin = (Date.now() - new Date(lastNotified).getTime()) / 60000;
    if (sinceLastMin < 360) return;
  }

  const adminEmail = getSetting('ADMIN_EMAIL') ||
                     (getActiveStaff().find(s => s.role === 'admin') || {}).email;
  if (!adminEmail) return;

  try {
    GmailApp.sendEmail(
      adminEmail,
      '[タスク管理] EXE エージェント応答なし',
      'EXE が ' + Math.floor(elapsedMin) + ' 分応答していません。\n' +
      '最終ハートビート: ' + last + '\n' +
      'PC: ' + (getSetting('LAST_HEARTBEAT_HOSTNAME') || '不明')
    );
    setSetting('LAST_DEAD_NOTIFICATION_AT', new Date().toISOString());
  } catch (e) {
    Logger.log('死活通知メール送信失敗: ' + e.message);
  }
}
```

2. **EXE 向け API エンドポイントを `Code.gs` の `doPost` に追加する**

```javascript
// Code.gs に追記

function doPost(e) {
  // Bearer トークン認証
  const props = PropertiesService.getScriptProperties();
  const validToken = props.getProperty(SCRIPT_PROP_KEYS.EXE_API_TOKEN);
  const authHeader = e.parameter.token || '';
  if (authHeader !== validToken) {
    return jsonResponse_({ error: 'Unauthorized' }, 401);
  }

  const body = JSON.parse(e.postData.contents || '{}');
  const action = body.action;

  if (action === 'heartbeat') {
    // ハートビート受信 → Settings に最終受信時刻を記録
    setSetting('LAST_HEARTBEAT_TIMESTAMP', new Date().toISOString());
    setSetting('LAST_HEARTBEAT_HOSTNAME', body.hostname || '');
    return jsonResponse_({ ok: true });
  }
  if (action === 'pickMailItem') {
    return jsonResponse_(pickMailItem(body.hostname));
  }
  if (action === 'completeMailItem') {
    return jsonResponse_(completeMailItem(body.id, body.status, body.errorMessage));
  }
  return jsonResponse_({ error: 'Unknown action' }, 400);
}

function jsonResponse_(data, status) {
  const payload = JSON.stringify(data);
  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}
```

3. **タイムベーストリガーを設定する**

GAS エディタ → 「トリガー」→ 「トリガーを追加」を 2 件登録する:

| 関数 | イベントソース | タイマー種別 | 間隔 | 目的 |
|---|---|---|---|---|
| `recoverStalePicked` | 時間主導型 | 分ベース | 5 分おき | CAS タイムアウト復旧（picked が 10 分以上経過したら pending に戻す） |
| `checkExeAlive`      | 時間主導型 | 分ベース | 1 分おき | EXE 死活監視・管理者通知（6 時間以内は重複抑制） |

加えて、月次トリガーで `archiveOldMailQueue` を毎月 1 日に実行するよう設定する（仕様 §11、本書 7.4 参照）。

**atomic CAS の実装ポイント**

`pickMailItem()` の重複送信防止ロジック:

```
EXE-A               GAS                 EXE-B
  |                   |                   |
  |--pickMailItem()--> |                   |
  |              [ScriptLock.waitLock]     |
  |              pending → picked          |
  |              SpreadsheetApp.flush()    |
  |  <--record----|                   |
  |               [Lock解放]              |
  |                   |<--pickMailItem()--|
  |                   | (同じ行は picked  |
  |                   |  → スキップ)      |
  |                   |--null----------->|
```

LockService によりスクリプトレベルで排他制御する。`SpreadsheetApp.flush()` により書き込みを即座にコミットする。

**`pickedBy` + タイムアウト復旧の設計**

- EXE が crash した場合、picked のまま stuck になる可能性がある
- `recoverStalePicked()` が 5 分ごとに稼働し、`pickedAt` から 10 分以上経過したレコードを `pending` に戻す
- これにより、EXE 障害時でも最大 15 分（10 分タイムアウト + 5 分トリガー間隔）で自動復旧する

**完了確認**
- [ ] `MailQueueService.gs` が GAS エディタに表示されている
- [ ] `doPost` に `heartbeat` / `pickMailItem` / `completeMailItem` が実装されている
- [ ] `enqueueMailRequest()` を手動実行して MailQueue シートに行が追加される
- [ ] 追加された行の `toAddresses` がカンマ区切り文字列（例: `a@x.com,b@x.com`）で保存されている
- [ ] `pickMailItem('test-host')` を実行して picked に遷移し、レスポンスに `previousRequestId` 列が含まれる
- [ ] `recoverStalePicked()` のトリガーが設定されている（5 分間隔・タイムアウト 10 分）
- [ ] `retryMailItem(originalId)` 実行後、MailQueue に**新規行**が追加され `previousRequestId` に元 id が入っている（**元レコードは status を含め一切変更されない**）

**よくあるエラーと対処**

- エラー: `LockService.waitLock` でタイムアウト（10秒）  
  対処: 多数の EXE が同時アクセスしている可能性がある。`waitLock` の時間を増やすか、EXE のポーリング間隔を分散させる。

- エラー: `doPost` で 401 Unauthorized  
  対処: EXE が送るトークンと `EXE_API_TOKEN` スクリプトプロパティの値が一致しているか確認する。

---

### STEP 5: 工番マスタ・社内カレンダー連携

**所要時間目安**: 1〜2 時間

**目的**

`doGet()` の初期ペイロードに外部マスタ（工番・社内カレンダー）を含め、フロントが起動時に全件取得できる状態にする。

**前提条件**
- STEP 2〜3 が完了していること

**作業内容**

1. **`doGet()` を改修して初期ペイロードを含める**

```javascript
// Code.gs の doGet() を以下に差し替える

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');

  // 初期データを JSON として埋め込む（起動時全件ロード）
  const payload = {
    staff:      getActiveStaff(),
    workTypes:  getWorkTypes(),
    koban:      loadKobanMaster_(),
    calendar:   loadCalendarMaster_(),
    settings: {
      pollingIntervalSeconds: parseInt(getSetting('POLLING_INTERVAL_SECONDS') || '30'),
      mailDefaultMode:        getSetting('MAIL_DEFAULT_MODE') || 'draft',
      webappVersion:          getSetting('WEBAPP_VERSION') || '1.0.0',
    }
  };
  template.initialPayload = JSON.stringify(payload);

  return template
    .evaluate()
    .setTitle('機械設計技術部 タスク管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

2. **ペイロードサイズの確認**

GAS エディタで以下のテスト関数を実行してサイズを確認する:

```javascript
function testPayloadSize() {
  const payload = {
    staff:     getActiveStaff(),
    workTypes: getWorkTypes(),
    koban:     loadKobanMaster_(),
    calendar:  loadCalendarMaster_()
  };
  const json = JSON.stringify(payload);
  Logger.log('Payload size: ' + (json.length / 1024).toFixed(1) + ' KB');
  // 推定 100〜200 KB であることを確認
}
```

3. **社内カレンダーの背景色ロジックをフロント向けに整理する**

フロントへ渡すカレンダーデータの形式:
```json
[
  { "日付": "2026-04-04", "区分": "休日",     "曜日": "土", "備考": "" },
  { "日付": "2026-04-11", "区分": "出勤土曜", "曜日": "土", "備考": "" },
  { "日付": "2026-04-29", "区分": "祝日",     "曜日": "水", "備考": "昭和の日" }
]
```

フロント JS でのカレンダー判定ロジック（`scripts.html` に実装する）:

```javascript
// 日付文字列 "YYYY-MM-DD" → { type, note }
function getCalendarInfo(dateStr, calendarMap) {
  const entry = calendarMap.get(dateStr);
  if (entry) return { type: entry['区分'], note: entry['備考'] || '' };
  // マスタにない土日は通常休日
  const d = new Date(dateStr);
  const dow = d.getDay();
  if (dow === 0) return { type: '休日', note: '' };
  if (dow === 6) return { type: '休日', note: '' };
  return { type: '平日', note: '' };
}
```

**完了確認**
- [ ] `testPayloadSize()` で 200 KB 以下であることを確認
- [ ] `doGet()` が HTML を返し、スクリプト内で `window.INITIAL_PAYLOAD` が参照できる
- [ ] カレンダーデータが「出勤土曜」区分を正しく含んでいる

---

### STEP 6: UI 実装

**所要時間目安**: 5〜7 日

**目的**

`docs/wireframe.html` を GAS HtmlService の形式（index.html / styles.html / scripts.html）に分解し、ガント・日報・モーダルを段階的に実装する。

**前提条件**
- STEP 5 が完了し、初期ペイロードが利用可能であること
- `docs/wireframe.html` を常時参照できる状態にしてあること

**wireframe.html を 3 ファイルへ分解する手順**

GAS HtmlService は複数 HTML ファイルをテンプレートで include できる。

**分解方針**

| 元ファイルの部位 | 移動先 | 内容 |
|---|---|---|
| `<style>` タグ全体 | `styles.html` | CSS 変数・全コンポーネントスタイル |
| `<body>` の HTML 構造 | `index.html` | ウィンドウシェル・サイドバー・メインコンテンツ |
| JavaScript ロジック | `scripts.html` | 全 JS（初期化・イベントハンドラ・API 通信） |

1. **`src/styles.html` を作成する**

```html
<style>
  /* wireframe.html の <style> タグの中身をそのままコピーする */
  /* ⚠️ CSS 変数・クラス名・コンポーネント命名を一切変更しない ⚠️ */
  :root {
    --bg-base:   #f1ebe0;
    --bg-panel:  #fbf8f1;
    /* ... wireframe.html の内容をそのまま ... */
  }
</style>
```

2. **`src/index.html` を作成する**

```html
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1280">
<title>機械設計技術部 タスク管理</title>
<script src="https://cdn.tailwindcss.com"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+JP:wght@400;500;600;700&family=Noto+Serif+JP:wght@500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<?!= HtmlService.createHtmlOutputFromFile('styles').getContent(); ?>
</head>
<body class="min-h-screen">

<!-- wireframe.html の <body> 内 HTML 構造をここに配置 -->
<!-- ただし静的サンプルデータは JS で動的生成に置き換える -->

<script>
// 初期ペイロードを JS グローバルに展開する
window.INITIAL_PAYLOAD = <?!= initialPayload; ?>;
</script>
<?!= HtmlService.createHtmlOutputFromFile('scripts').getContent(); ?>
</body>
</html>
```

3. **`src/scripts.html` を作成する**

```html
<script>
/**
 * scripts.html
 * アプリケーションの全 JavaScript ロジックを担う。
 */

// ── 初期化 ───────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  const P = window.INITIAL_PAYLOAD;
  App.init(P);
});

const App = {
  staff:    [],
  workTypes: [],
  kobanMap:  new Map(),
  calendarMap: new Map(),

  init(payload) {
    this.staff     = payload.staff;
    this.workTypes = payload.workTypes;
    this.kobanMap  = new Map(payload.koban.map(k => [k['工番'], k]));
    this.calendarMap = new Map(payload.calendar.map(c => [c['日付'], c]));

    this.renderSidebar();
    this.renderGantt();
    this.updateHeartbeatStatus();
    this.startHeartbeatChecker();
  },

  // ── サイドバー ───────────────────────────────────────────

  renderSidebar() {
    const container = document.getElementById('sidebar-staff');
    if (!container) return;
    container.innerHTML = this.staff.map((s, i) => `
      <div class="nav-item" data-staff-id="${s.id}">
        <span class="w-5 h-5 rounded-full grid place-items-center text-[9px] font-semibold text-white av-${(i % 6) + 1}">
          ${s.name.charAt(0)}
        </span>
        <span>${s.name}</span>
      </div>
    `).join('');
  },

  // ── ガント ───────────────────────────────────────────────

  renderGantt() {
    // 描画範囲: 過去30日〜未来90日
    const today    = new Date();
    const startDay = new Date(today); startDay.setDate(startDay.getDate() - 30);
    const endDay   = new Date(today); endDay.setDate(endDay.getDate() + 90);

    this.renderGanttHeader(startDay, endDay, today);
    this.scrollGanttToToday();
  },

  renderGanttHeader(startDay, endDay, today) {
    // 日付ヘッダ行を動的生成する（wireframe.html のクラス名を使用）
    // ... 実装省略（詳細は下記「ガント段階的実装」参照）
  },

  scrollGanttToToday() {
    // 今日を中央に scroll
    const ganttScroll = document.querySelector('.gantt-scroll');
    if (!ganttScroll) return;
    const DAY_WIDTH = 64; // px/日
    const daysBefore = 30;
    const containerWidth = ganttScroll.clientWidth;
    ganttScroll.scrollLeft = DAY_WIDTH * daysBefore - containerWidth / 2;
  },

  // ── EXE 死活監視 ────────────────────────────────────────

  updateHeartbeatStatus() {
    google.script.run
      .withSuccessHandler(result => this.renderHeartbeat(result))
      .getSetting('LAST_HEARTBEAT_TIMESTAMP');
  },

  renderHeartbeat(timestamp) {
    const el = document.getElementById('heartbeat-status');
    if (!el) return;
    if (!timestamp) {
      el.innerHTML = '<span class="heartbeat" style="background:var(--err)"></span> EXE 未接続';
      return;
    }
    const diff = Math.floor((Date.now() - new Date(timestamp).getTime()) / 60000);
    const isStale = diff >= 5;
    el.innerHTML = `
      <span class="heartbeat" style="background: ${isStale ? 'var(--err)' : 'var(--ok)'}"></span>
      最終応答 ${diff} 分前
      ${isStale ? '<span class="badge badge-failed ml-1">要確認</span>' : ''}
    `;
  },

  startHeartbeatChecker() {
    // 1 分ごとに EXE 死活状態を更新
    setInterval(() => this.updateHeartbeatStatus(), 60 * 1000);
  },
};
</script>
```

**デザイントークンの実装ルール**

実装時に守るべき制約:

| 項目 | ルール |
|---|---|
| CSS 変数 | `wireframe.html` の `:root` 変数名を変更しない |
| フォント | `wireframe.html` と同じ Google Fonts URL を使用する |
| ガント1日幅 | `64px`（デフォルト）、`48px`（狭い）、`96px`（広い）の3択 |
| ガント行高 | `56px` 固定 |
| ガントヘッダ | 月32px + 日32px = 合計64px |
| バーカラー | `bar-clay / bar-plum / bar-ochre / bar-burgundy / bar-moss / bar-indigo / bar-stone` のみ使用 |
| アバター | `av-1` 〜 `av-6` の 6 色を staff 順に割り当て |

**ガント段階的実装の順序**

1. **Day 1**: ヘッダ行（月・日）の動的生成。今日が中央に来るよう scroll。土日・祝日の背景色。
2. **Day 2**: スタッフ行（固定列）の動的生成。アバター色の割り当て。
3. **Day 3**: セル描画（背景色 weekend / today）。ガントバーの読み取り専用表示。
4. **Day 4**: ガントバーのドラッグ操作（mousedown / mousemove / mouseup）。日付変更時の API 保存。
5. **Day 5**: 日報カード実装。工番プルダウン（インクリメンタル検索）。自動補完。
6. **Day 6**: メール送信モーダル。`enqueueMailRequest` 呼び出し。再送ボタン（`retryMailItem`）。
7. **Day 7**: 祝日ツールチップ。納入先住所ツールチップ。EXE 死活バッジ。ガント幅切替の LocalStorage 永続化（キー `ttm.ganttDayWidth`）。

> **重要**: 本書はあくまで骨組み・段階的アプローチを示すのみ。実装中の HTML 構造・配色・CSS クラス名・余白は **`docs/wireframe.html` を逐行コピー**して使うこと（仕様 v1.0 §5.1：wireframe からの逸脱は禁止）。

**ガントバードラッグ操作のスケルトン**

```javascript
// ガントバーをドラッグして期間を変更する最小骨格
function attachGanttDrag(barEl, scheduleId) {
  barEl.addEventListener('mousedown', (e) => {
    const startX = e.clientX;
    const startLeft = parseInt(barEl.style.left) || 0;
    const dayWidth = getCurrentDayWidth(); // 64 / 48 / 96
    const onMove = (ev) => {
      const dx = ev.clientX - startX;
      const dayDelta = Math.round(dx / dayWidth);
      barEl.style.left = (startLeft + dayDelta * dayWidth) + 'px';
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup',   onUp);
      // 数秒デバウンスで GAS 側に保存
      debouncedSaveSchedule(scheduleId, computeNewDates(barEl));
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
  });
}
```

**メール送信モーダル / 再送ボタンのスケルトン**

```javascript
// 「メール下書き作成」ボタン → モーダル → enqueueMailRequest
function openMailModal(staffId, reportDate) {
  // wireframe.html §5.3 のモーダルマークアップを参照すること
  const modal = document.getElementById('mail-modal');
  // 1) targetStaff の email を Staff マスタから取得して表示
  // 2) Staff active 全員の email をカンマ区切りで toAddresses プレビューに表示
  // 3) 送信モード ラジオ（既定: draft）
  // 4) 「依頼を登録する」 → google.script.run.enqueueMailRequest({...})

  document.getElementById('mail-submit').onclick = () => {
    const params = collectMailFormParams(staffId, reportDate);
    google.script.run
      .withSuccessHandler(res => {
        if (res.success) toast('依頼を登録しました');
        modal.classList.add('hidden');
      })
      .enqueueMailRequest(params);
  };
}

// 「再送する」ボタン（status=sent 時は警告→確認）
function retryMail(originalId, currentStatus) {
  if (currentStatus === 'sent') {
    if (!confirm('既に送信済みです。本当に再送しますか？')) return;
  }
  google.script.run
    .withSuccessHandler(res => {
      if (res.success) toast('再送依頼を登録しました（新ID: ' + res.newId + '）');
      else            toast('再送に失敗: ' + res.reason);
    })
    .retryMailItem(originalId);
}
```

**工番プルダウン（インクリメンタル検索）の実装ポイント**

```javascript
// 工番検索: 入力文字が 工番・受注先・品名 のどれかに部分一致するものを返す
function filterKoban(query) {
  if (!query) return [];
  const q = query.toLowerCase();
  const results = [];
  for (const [code, rec] of App.kobanMap) {
    if (
      code.toLowerCase().includes(q) ||
      (rec['受注先'] || '').toLowerCase().includes(q) ||
      (rec['品名']   || '').toLowerCase().includes(q)
    ) {
      results.push(rec);
      if (results.length >= 50) break; // 表示上限
    }
  }
  return results;
}

// プルダウン表示形式: "LW23012  住友建機㈱  4ton応用機アタッチメントポジショナー"
function formatKobanLabel(rec) {
  return [rec['工番'], rec['受注先'], rec['品名']].filter(Boolean).join('  ');
}
```

**完了確認**
- [ ] `clasp push` 後にブラウザで WebApp が表示される
- [ ] ガントチャートが今日を中央に表示する
- [ ] 土日・祝日の背景色が `wireframe.html` に一致する
- [ ] 工番プルダウンで文字入力すると候補が絞り込まれる
- [ ] 工番選択時に受注先・納入先・品名が自動補完される
- [ ] 納入先住所がマウスホバーでツールチップ表示される
- [ ] 祝日名（備考）がマウスホバーでツールチップ表示される
- [ ] EXE 死活ステータスがヘッダに表示される

---

### STEP 7: Python EXE 開発

**所要時間目安**: 3〜4 日

**目的**

MailQueue をポーリングして Outlook で下書き作成または送信を行うローカルエージェント `TeamTaskMail.exe` を Python で実装する。

**前提条件**
- STEP 4 の `doPost` が実装・デプロイされていること
- Windows PC + Outlook がインストールされていること

**ディレクトリ構成**

```
Team Task Management/
  exe/
    mail_agent/
      __init__.py
      main.py            # エントリポイント（起動・ループ制御）
      poller.py          # GAS WebApp ポーリング
      outlook_client.py  # Outlook COM 経由でメール操作
      template.py        # 件名・本文テンプレート（ハードコード）
      config.py          # 設定（WebApp URL・トークン等）
      logger.py          # ファイルログ
    requirements.txt
    build.bat            # PyInstaller ビルドスクリプト
```

**作業内容**

1. **`exe/mail_agent/config.py` を実装する**

```python
# config.py
import os

# GAS WebApp の URL（デプロイ後に設定する）
WEBAPP_URL = os.environ.get(
    "TEAM_TASK_WEBAPP_URL",
    "https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec"
)

# EXE 認証トークン（スクリプトプロパティ EXE_API_TOKEN の値を設定する）
API_TOKEN = os.environ.get("TEAM_TASK_API_TOKEN", "<ここにトークンを入力>")

# ポーリング間隔（秒）。GAS の Settings から取得する想定だが、初期値として 30 秒。
POLLING_INTERVAL = int(os.environ.get("TEAM_TASK_POLLING_INTERVAL", "30"))

# ハートビート送信間隔（秒）
HEARTBEAT_INTERVAL = 60

# このEXEを実行している PC のホスト名
import socket
HOSTNAME = socket.gethostname()
```

> **セキュリティ上の注意**: `API_TOKEN` はハードコードせず、環境変数で渡すことが望ましい。ただし本アプリでは DOMAIN 内限定運用のため、ビルド時に設定しても許容範囲とする。

2. **`exe/mail_agent/logger.py` を実装する**

```python
# logger.py
import logging
import os
from datetime import datetime

LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA", "."), "TeamTaskMail", "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(
            os.path.join(LOG_DIR, f"agent_{datetime.now().strftime('%Y%m%d')}.log"),
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("TeamTaskMail")
```

3. **`exe/mail_agent/poller.py` を実装する（HTTP クライアント）**

```python
# poller.py
import requests
import json
from .config import WEBAPP_URL, API_TOKEN, HOSTNAME
from .logger import logger


def _post(action: str, extra: dict = None) -> dict | None:
    """GAS WebApp の doPost に JSON でリクエストする。"""
    payload = {"action": action, "hostname": HOSTNAME, **(extra or {})}
    try:
        resp = requests.post(
            WEBAPP_URL,
            params={"token": API_TOKEN},
            json=payload,
            timeout=30
        )
        resp.raise_for_status()
        return resp.json()
    except requests.Timeout:
        logger.warning(f"[{action}] タイムアウト (30s)")
    except requests.HTTPError as e:
        logger.error(f"[{action}] HTTP エラー: {e.response.status_code}")
    except Exception as e:
        logger.error(f"[{action}] 予期せぬエラー: {e}")
    return None


def send_heartbeat() -> bool:
    """ハートビートを送信する。成功なら True。"""
    result = _post("heartbeat")
    return result is not None and result.get("ok") is True


def pick_mail_item() -> dict | None:
    """pending の MailQueue アイテムを 1 件取得（picked に遷移）。"""
    return _post("pickMailItem")


def complete_mail_item(item_id: str, status: str, error_message: str = "") -> bool:
    """処理完了を報告する。"""
    result = _post("completeMailItem", {
        "id": item_id,
        "status": status,
        "errorMessage": error_message
    })
    return result is not None
```

4. **`exe/mail_agent/template.py` を実装する（ハードコードテンプレート）**

```python
# template.py
# メール件名・本文のテンプレート（ハードコード）
# テンプレート変更時は EXE を再ビルドして配布する
# 仕様 v1.0 §7.4.7 を一字一句準拠する。文言を変えると業務メール体裁が壊れる。


def build_subject(vars: dict) -> str:
    """
    件名テンプレート（仕様 v1.0 §7.4.7 準拠）。
    フォーマット: 【機械設計技術部】{staffName} 業務報告 {reportDate}
    例:           【機械設計技術部】山田 太郎 業務報告 2026/05/08
    vars: { "staffName": str, "reportDate": str }
        reportDate は GAS 側で "YYYY/MM/DD" 形式で渡す（仕様 §7.4.7）。
    """
    staff_name  = vars.get("staffName", "")
    report_date = vars.get("reportDate", "")
    return f"【機械設計技術部】{staff_name} 業務報告 {report_date}"


def build_body(vars: dict) -> str:
    """
    本文テンプレート（仕様 v1.0 §7.4.7 準拠・プレーンテキスト）。

    vars: {
        "staffName":  str,
        "reportDate": str,
        "todayItems": [
            { "periodStart": str, "periodEnd": str,
              "kobanCode": str, "customer": str, "productName": str,
              "workType": str, "detail": str }
        ],
        "yesterdayItems": [...]
    }

    「（一般作業）」の場合は kobanCode/customer/productName が空文字列になるので、
    出力時に空項目を skip する（仕様 §7.4.7 注記）。
    """
    staff_name = vars.get("staffName", "")

    def format_items(items):
        if not items:
            return "　（記録なし）"
        lines = []
        for n, r in enumerate(items, start=1):
            parts = [
                f"{r.get('periodStart','')}〜{r.get('periodEnd','')}",
            ]
            for k in ("kobanCode", "customer", "productName"):
                v = r.get(k, "")
                if v:
                    parts.append(v)
            parts.append("/")
            parts.append(r.get("workType", ""))
            if r.get("detail"):
                parts.append(r["detail"])
            lines.append(f"{n}. " + "  ".join(parts))
        return "\n".join(lines)

    today_section     = format_items(vars.get("todayItems", []))
    yesterday_section = format_items(vars.get("yesterdayItems", []))

    body = (
        f"お疲れ様です。{staff_name} です。\n"
        f"本日の業務をご報告いたします。\n"
        f"\n"
        f"▼ 本日の作業内容\n"
        f"{today_section}\n"
        f"\n"
        f"▼ 前日までの作業報告\n"
        f"{yesterday_section}\n"
        f"\n"
        f"以上、よろしくお願いいたします。\n"
    )
    return body
```

5. **`exe/mail_agent/outlook_client.py` を実装する（COM クライアント）**

```python
# outlook_client.py
import win32com.client  # pywin32
from .logger import logger


def create_draft(to_addresses: list[str], subject: str, body: str) -> bool:
    """
    Outlook の下書きフォルダにメールを作成する。
    送信はしない（モード: draft）。

    :param to_addresses: 宛先メールアドレスのリスト
    :param subject: 件名
    :param body: 本文（プレーンテキスト）
    :return: 成功なら True
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = "; ".join(to_addresses)
        mail.Subject = subject
        mail.Body = body
        mail.Save()  # 下書き保存（SendKeys は使わない）
        logger.info(f"下書き作成成功: {subject}")
        return True
    except Exception as e:
        logger.error(f"下書き作成失敗: {e}")
        raise


def send_mail(to_addresses: list[str], subject: str, body: str) -> bool:
    """
    Outlook でメールを即時送信する（モード: send）。
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(to_addresses)
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        logger.info(f"送信成功: {subject}")
        return True
    except Exception as e:
        logger.error(f"送信失敗: {e}")
        raise
```

6. **`exe/mail_agent/main.py` を実装する（エントリポイント）**

```python
# main.py
import time
import json
from .config import POLLING_INTERVAL, HEARTBEAT_INTERVAL
from .poller import send_heartbeat, pick_mail_item, complete_mail_item
from .outlook_client import create_draft, send_mail
from .template import build_subject, build_body
from .logger import logger


def process_item(item: dict) -> None:
    """1 件の MailQueue アイテムを処理する。"""
    item_id = item["id"]
    mode    = item.get("mode", "draft")
    logger.info(f"処理開始: {item_id} (mode={mode})")

    try:
        # 宛先の展開（仕様 v1.0 §6.3: toAddresses はカンマ区切り string）
        raw_to = item.get("toAddresses", "") or ""
        to_addresses = [a.strip() for a in raw_to.split(",") if a.strip()]

        # subjectVars / bodyVars は JSON string（GAS 側で JSON.stringify したもの）
        subject_vars = json.loads(item.get("subjectVars") or "{}")
        body_vars    = json.loads(item.get("bodyVars")    or "{}")

        subject = build_subject(subject_vars)
        body    = build_body(body_vars)

        if mode == "send":
            send_mail(to_addresses, subject, body)
            complete_mail_item(item_id, "sent")
        else:  # draft（既定）
            create_draft(to_addresses, subject, body)
            complete_mail_item(item_id, "drafted")

    except Exception as e:
        error_msg = str(e)
        logger.error(f"処理失敗: {item_id}: {error_msg}")
        complete_mail_item(item_id, "failed", error_msg)


def run():
    """メインループ。ポーリングとハートビートを管理する。"""
    logger.info("TeamTaskMail EXE エージェント 起動")
    last_heartbeat = 0.0

    while True:
        now = time.time()

        # ハートビート送信（HEARTBEAT_INTERVAL 秒ごと）
        if now - last_heartbeat >= HEARTBEAT_INTERVAL:
            if send_heartbeat():
                logger.debug("ハートビート送信 OK")
            else:
                logger.warning("ハートビート送信失敗")
            last_heartbeat = now

        # MailQueue のポーリング
        item = pick_mail_item()
        if item:
            process_item(item)
        else:
            logger.debug("pending なし")

        time.sleep(POLLING_INTERVAL)


if __name__ == "__main__":
    run()
```

> **実行方法の注意（重要）**: `main.py` は相対インポート（`from .config import ...`）を使うため、**直接 `python main.py` で起動すると `ImportError` になる**。次のいずれかで実行すること:
>
> ```powershell
> # 方式 A（推奨・開発時）: パッケージとして実行
> cd "C:\Users\<ユーザー名>\Documents\Team Task Management\exe"
> python -m mail_agent.main
> ```
>
> ```powershell
> # 方式 B: PyInstaller でビルドした EXE を実行（本番）
> .\dist\v1.0.0\TeamTaskMail.exe
> ```
>
> PyInstaller は `--name TeamTaskMail mail_agent\main.py` のエントリ指定により、内部で同等の解決をする。

7. **`exe/mail_agent/__init__.py` を作成する（空ファイル）**

```python
# __init__.py
```

8. **`exe/requirements.txt` を作成する**

```
requests>=2.31.0
pywin32>=306
pyinstaller>=6.3.0
```

**エラー処理・リトライの設計**

```
pick_mail_item() 失敗
  → ログ出力 → 次のポーリングサイクルへ（GAS 側は picked のまま）
  → recoverStalePicked() が 10 分後に pending に戻す

process_item() 内で例外
  → complete_mail_item(id, "failed", error_message) を必ず呼ぶ
  → GAS 側で管理者通知メールを送信
  → UI に badge-failed が表示される
  → 管理者が retryMailItem(originalId) を呼ぶと、元レコードを複製した新 pending を生成して再送（仕様 §7.4.6）
```

**完了確認**
- [ ] 各モジュールが Python 構文エラーなく import できる
- [ ] `cd exe && python -m mail_agent.main` で起動してログが出力される（相対インポート対応）
- [ ] ハートビートが送信され、GAS の Settings に `LAST_HEARTBEAT_TIMESTAMP` が記録される
- [ ] テスト用の pending アイテムを MailQueue に追加すると、EXE がピックアップして処理する
- [ ] Outlook の下書きフォルダにメールが作成される
- [ ] 処理後に MailQueue の status が `drafted` に更新される

**よくあるエラーと対処**

- エラー: `pywintypes.com_error: (-2147221005, ...)` （Outlook が起動していない）  
  対処: Outlook を起動してから EXE を実行する。または Outlook のプロファイルが設定されていることを確認する。

- エラー: `requests.exceptions.ConnectionError`  
  対処: WebApp URL が正しいか確認する。ネットワーク接続を確認する。GAS WebApp がデプロイされているか確認する。

- エラー: `json.JSONDecodeError` （GAS からの応答が JSON でない）  
  対処: GAS WebApp の認証エラーや実行エラーの可能性がある。GAS エディタの実行ログを確認する。

---

### STEP 8: PyInstaller でビルド

**所要時間目安**: 1〜2 時間

**目的**

Python スクリプトをスタンドアロンの `.exe` ファイルにビルドし、共有フォルダへ配置する。

**前提条件**
- STEP 7 の Python コードが完成し、動作確認できていること
- PyInstaller がインストールされていること (`pip install pyinstaller`)

**作業内容**

1. **`exe/build.bat` を作成する**

```bat
@echo off
REM TeamTaskMail EXE ビルドスクリプト
REM 実行前に仮想環境を有効化してください:
REM   .\.venv\Scripts\activate

setlocal

set VERSION=1.0.0
set EXE_NAME=TeamTaskMail
set ENTRY=mail_agent\main.py
set DIST_DIR=dist\v%VERSION%

echo === TeamTaskMail %VERSION% ビルド開始 ===

REM バージョン情報ファイルの生成
(
echo VSVersionInfo^(
echo   ffi=FixedFileInfo^(
echo     filevers=^(1, 0, 0, 0^),
echo     prodvers=^(1, 0, 0, 0^),
echo     mask=0x3f,
echo     flags=0x0,
echo     OS=0x40004,
echo     fileType=0x1,
echo     subtype=0x0,
echo     date=^(0, 0^)
echo   ^),
echo   kids=[
echo     StringFileInfo^([
echo       StringTable^(u'040904B0', [
echo         StringStruct^(u'CompanyName',      u'機械設計技術部'^),
echo         StringStruct^(u'ProductName',      u'TeamTaskMail'^),
echo         StringStruct^(u'FileVersion',      u'%VERSION%'^),
echo         StringStruct^(u'ProductVersion',   u'%VERSION%'^),
echo         StringStruct^(u'FileDescription',  u'タスク管理 メールエージェント'^),
echo       ]^)
echo     ]^),
echo     VarFileInfo^([VarStruct^(u'Translation', [1033, 1200]^)^]^)
echo   ]
echo ^)
) > version_info.txt

REM PyInstaller でビルド
pyinstaller ^
  --onefile ^
  --noconsole ^
  --name %EXE_NAME% ^
  --version-file version_info.txt ^
  %ENTRY%

REM 出力先へコピー
if not exist %DIST_DIR% mkdir %DIST_DIR%
copy dist\%EXE_NAME%.exe %DIST_DIR%\%EXE_NAME%.exe

echo.
echo === ビルド完了 ===
echo 出力先: %DIST_DIR%\%EXE_NAME%.exe

REM 一時ファイルの削除
del version_info.txt
rmdir /s /q build
del %EXE_NAME%.spec

endlocal
```

2. **ビルドを実行する**

```powershell
# exe ディレクトリへ移動
cd "C:\Users\<ユーザー名>\Documents\Team Task Management\exe"

# 仮想環境を有効化
.\.venv\Scripts\Activate.ps1

# ビルド実行
.\build.bat
```

成功すると `exe\dist\v1.0.0\TeamTaskMail.exe` が生成される。

3. **ビルド成果物の動作確認**

```powershell
# EXE を直接実行してログが出力されることを確認
.\dist\v1.0.0\TeamTaskMail.exe
# ログウィンドウは表示されない（--noconsole）
# ログは %LOCALAPPDATA%\TeamTaskMail\logs\ に出力される
```

4. **共有フォルダへコピーする**

```powershell
# 共有フォルダのリリースディレクトリへコピー
copy ".\dist\v1.0.0\TeamTaskMail.exe" "\\<サーバー名>\共有\タスク管理アプリ\releases\v1.0.0\TeamTaskMail.exe"
```

5. **各 PC へのインストール手順（運用担当者向け）**

```
1. \\<サーバー名>\共有\タスク管理アプリ\releases\v1.0.0\TeamTaskMail.exe を開く
2. C:\ProgramData\TeamTaskMail\ フォルダを作成する
3. TeamTaskMail.exe を C:\ProgramData\TeamTaskMail\ にコピーする
4. 環境変数を設定する（PowerShell 管理者権限で実行）:
   [System.Environment]::SetEnvironmentVariable("TEAM_TASK_WEBAPP_URL", "https://...", "Machine")
   [System.Environment]::SetEnvironmentVariable("TEAM_TASK_API_TOKEN", "<トークン>", "Machine")
5. タスクスケジューラに登録する（5.4 参照）
```

**完了確認**
- [ ] `build.bat` が正常終了し `TeamTaskMail.exe` が生成される
- [ ] EXE を実行するとログが `%LOCALAPPDATA%\TeamTaskMail\logs\` に出力される
- [ ] EXE が GAS WebApp にハートビートを送信し、Settings に記録される
- [ ] 共有フォルダに EXE がコピーされている

**よくあるエラーと対処**

- エラー: PyInstaller 実行時「ModuleNotFoundError: No module named 'win32com'」  
  対処: `pip install pywin32` を再実行する。仮想環境が有効化されているか確認する。

- エラー: `--noconsole` でビルドした EXE が起動直後に終了する  
  対処: 一度 `--noconsole` を外してビルドし、エラー出力を確認する。ログファイルも確認する。

- エラー: EXE のサイズが異常に大きい（200 MB 以上）  
  対処: `--onefile` で正常。pywin32 を含むため大きめになる（通常 40〜80 MB）。`--exclude-module` で不要なモジュールを除外できる。

---

### STEP 9: 結合テスト

**所要時間目安**: 2〜3 日

**目的**

GAS WebApp と Python EXE、Outlook、スプレッドシートの全連携を検証する。

**前提条件**
- STEP 1〜8 がすべて完了していること
- GAS WebApp が DOMAIN アクセスでデプロイされていること（5.1 参照）
- EXE が少なくとも 2 台の PC に配置されていること（重複防止テストのため）

**テストシナリオ一覧**

#### T-01: 単体動作確認

| # | テスト内容 | 手順 | 期待結果 |
|---|---|---|---|
| T-01-1 | GAS setupAll | GAS エディタで `setupAll()` を実行 | 7 シートが生成される |
| T-01-2 | 工番マスタ取得 | `testPayloadSize()` を実行 | 200 KB 以下 |
| T-01-3 | DataService CRUD | `createSchedule`, `updateSchedule`, `deleteSchedule` を手動実行 | Schedules シートが正しく更新される |
| T-01-4 | EXE 起動 | `TeamTaskMail.exe` を実行 | ログが出力され、ハートビートが送信される |

#### T-02: GAS ↔ EXE 通信確認

| # | テスト内容 | 手順 | 期待結果 |
|---|---|---|---|
| T-02-1 | ハートビート | EXE 起動後 1 分待つ | `Settings` の `LAST_HEARTBEAT_TIMESTAMP` が更新される |
| T-02-2 | ピックアップ | `enqueueMailRequest()` で pending を追加 → EXE のポーリング待機 | EXE が pending を拾い、status が `picked` → `drafted` になる |
| T-02-3 | 認証失敗 | 不正なトークンで `/exec?token=invalid` にリクエスト | 401 相当のエラーレスポンスが返る |

#### T-03: Outlook 下書き作成確認

| # | テスト内容 | 手順 | 期待結果 |
|---|---|---|---|
| T-03-1 | 下書き作成 | mode=draft の MailQueue アイテムを処理 | Outlook の「下書き」フォルダにメールが作成される |
| T-03-2 | 件名・本文 | 作成された下書きを確認 | 件名が「【機械設計技術部】{氏名} 業務報告 {YYYY/MM/DD}」になっている（仕様 v1.0 §7.4.7） |
| T-03-3 | 宛先 | 作成された下書きを確認 | to には active な Staff 全員のアドレスが入っている |

#### T-04: 重複送信防止の検証手順（最重要）

**テスト環境**: EXE を 2 プロセス並走させる。可能なら 2 台の PC（PC-A / PC-B）で行うのが望ましいが、PC が 1 台しかない場合でも、以下の手順で**同一 PC で 2 つのコンソール（プロセス）を立ち上げて検証可能**。

**A. 同一 PC で 2 プロセス起動する手順（コンソール版ビルドを使う）**

PyInstaller の `--noconsole` を**外した**コンソール表示版をテスト用にビルドし、2 つのターミナルで同時起動する:

```powershell
# 1) コンソール版ビルド（テスト専用）
cd "C:\Users\<ユーザー名>\Documents\Team Task Management\exe"
.\.venv\Scripts\Activate.ps1
pyinstaller --onefile --name TeamTaskMail_console mail_agent\main.py
# → exe\dist\TeamTaskMail_console.exe が生成される

# 2) ターミナル A（PowerShell #1）
$env:TEAM_TASK_HOSTNAME = "TEST-A"
.\dist\TeamTaskMail_console.exe

# 3) ターミナル B（PowerShell #2 を別ウィンドウで開く）
$env:TEAM_TASK_HOSTNAME = "TEST-B"
.\dist\TeamTaskMail_console.exe
```

> 各コンソールに別 hostname（`TEST-A` / `TEST-B`）を設定することで、MailQueue の `pickedBy` 列でどちらが取得したかを識別できる。`config.py` の `HOSTNAME = socket.gethostname()` を `os.environ.get("TEAM_TASK_HOSTNAME", socket.gethostname())` に変更しておくと環境変数優先になり、テストが容易。

**B. 2 台の PC で実施する場合**

```
手順:
1. PC-A, PC-B の両方で TeamTaskMail.exe を起動する
2. MailQueue に pending アイテムを 1 件追加する（GAS エディタで enqueueMailRequest を実行）
3. 両 EXE のログを監視する（%LOCALAPPDATA%\TeamTaskMail\logs\ をテールする）
4. 30 秒以内（ポーリング間隔）に確認する

期待結果:
- MailQueue シートの当該行の status が `drafted` または `sent` に 1 回だけ遷移する
- どちらか 1 台の EXE のみ「処理開始: mq_xxx」ログが出る
- もう 1 台は「pending なし」ログが出る
- Outlook の下書きが 1 通のみ作成される（2 通にならない）

失敗パターン（NG）:
- MailQueue に 2 行のログ（drafted + drafted）が記録されている
- Outlook に 2 通の下書きが作成されている
```

**PowerShell でのログ監視コマンド（各 PC で実行）**

```powershell
Get-Content "$env:LOCALAPPDATA\TeamTaskMail\logs\agent_$(Get-Date -Format yyyyMMdd).log" -Wait
```

#### T-05: 再送機能の検証手順

```
手順:
1. MailQueue に pending アイテムを 1 件追加する
2. EXE を意図的に停止（タスクマネージャーで終了）
3. GAS エディタで recoverStalePicked() のタイムアウトを 1 分に短縮して実行
   （本番は 10 分だが、テスト時は変数を書き換える）
4. MailQueue の status が picked → pending に戻ることを確認
5. EXE を再起動して、再び処理されることを確認

または（failed からの再送）:
1. EXE の outlook_client.py の send/draft 処理で故意に例外を raise させる
2. MailQueue に pending を追加して EXE に処理させる
3. **元レコード**の status が failed になることを確認
4. GAS エディタで retryMailItem(originalId) を呼ぶ → 戻り値 `{success:true, newId:'mq_xxx'}`
5. **新レコード**が MailQueue に追加され、status='pending'、`previousRequestId=originalId` であることを確認
6. **元レコードは status=failed のまま**（変更されないこと）を確認
7. EXE が次のポーリングで新レコードを処理することを確認

期待結果:
- 再送後に Outlook に下書きが 1 通作成される（**新規レコード分のみ**）
- MailLog に "retry" イベントが 2 行記録されている（元レコード id と新レコード id それぞれに 1 行ずつ）
- 元レコードの履歴が破壊されていない
```

#### T-06: EXE 停止時の挙動確認

```
手順:
1. EXE を起動してハートビートが正常に送信されることを確認
2. EXE を停止する
3. UI のヘッダを監視する（updateHeartbeatStatus が 1 分ごとに更新）
4. 5 分後にヘッダの表示を確認する

期待結果:
- 「最終応答 N 分前」の N が増加する
- 5 分以上で赤バッジ「要確認」が表示される
- EXE を再起動するとバッジが消える
```

**完了確認**
- [ ] T-01〜T-06 の全テストが期待結果を満たす
- [ ] MailQueue シートに想定外のレコード（重複など）がない
- [ ] MailLog に全処理イベントが記録されている
- [ ] Outlook の下書きが意図した件名・本文・宛先になっている

---

## 5. デプロイ手順

### 5.1 GAS WebApp デプロイ

```bash
# 1. 最新コードをプッシュ
clasp push

# 2. GAS エディタを開く
clasp open
```

GAS エディタでの操作:

1. 「デプロイ」→「新しいデプロイ」
2. 種類: 「ウェブアプリ」を選択
3. 説明: `v1.0.0`（バージョン番号を必ず記入）
4. 次のユーザーとして実行: 「自分（メールアドレス）」
5. アクセスできるユーザー: **「自分のドメイン内の全員」**
6. 「デプロイ」ボタンをクリック
7. デプロイ URL（`https://script.google.com/macros/s/<ID>/exec`）をコピーして控える

> **重要**: デプロイ URL は変わらない。コードを更新するときは「デプロイを管理」→「バージョンを選択」→「新しいバージョン」を作成してデプロイを更新する。

### 5.2 DOMAIN アクセス制限

`appsscript.json` の設定が DOMAIN になっていることを確認:

```json
"webapp": {
  "executeAs": "USER_DEPLOYING",
  "access": "DOMAIN"
}
```

> `access: "DOMAIN"` により、組織のドメイン外からのアクセスは 403 になる。

### 5.3 EXE 配布

```
1. STEP 8 でビルドした TeamTaskMail.exe を確認する
   場所: exe\dist\v1.0.0\TeamTaskMail.exe

2. 共有フォルダへコピーする
   \\<サーバー名>\共有\タスク管理アプリ\releases\v1.0.0\

3. 各 PC の担当者へ展開手順を案内する（README.txt を参照）

4. 各 PC での設置:
   - C:\ProgramData\TeamTaskMail\TeamTaskMail.exe
   - 環境変数 TEAM_TASK_WEBAPP_URL / TEAM_TASK_API_TOKEN を設定
```

### 5.4 タスクスケジューラ登録

各 PC で以下の設定でタスクスケジューラを登録する:

```powershell
# 管理者権限の PowerShell で実行

$action = New-ScheduledTaskAction `
  -Execute "C:\ProgramData\TeamTaskMail\TeamTaskMail.exe"

$trigger = New-ScheduledTaskTrigger `
  -AtLogOn `
  -User "$env:USERDOMAIN\$env:USERNAME"

$settings = New-ScheduledTaskSettingsSet `
  -RestartCount 3 `
  -RestartInterval (New-TimeSpan -Minutes 1) `
  -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
  -MultipleInstances IgnoreNew

Register-ScheduledTask `
  -TaskName "TeamTaskMail" `
  -Action $action `
  -Trigger $trigger `
  -Settings $settings `
  -Description "タスク管理 メールエージェント"

Write-Host "タスクスケジューラ登録完了"
```

**設定ポイント**
- トリガー: ログオン時（ユーザーがサインインするたびに起動）
- 再試行: 失敗時に 1 分間隔で 3 回再試行
- 実行時間制限: なし（`PT0S` = 無制限）
- 多重起動: 無視（`IgnoreNew`）

---

## 6. 受入テスト項目チェックリスト

### 機能テスト

**ガントチャート**
- [ ] 起動時に今日が中央に表示される
- [ ] 過去30日〜未来90日の範囲が表示される
- [ ] 土日はグレー背景、本日はアクセント色下線＋背景
- [ ] 出勤土曜の日付ヘッダに「出勤」バッジが表示される
- [ ] 祝日の日番号がエラー色で表示される
- [ ] 祝日名が日付セルのホバーでツールチップ表示される
- [ ] ガントバーが工番別に色分けされる
- [ ] スタッフ列が固定（横スクロール時も見える）
- [ ] バーをドラッグすると期間が変更できる

**日報入力**
- [ ] 工番プルダウンにインクリメンタル検索が動作する
- [ ] 工番選択時に受注先・納入先・品名が自動補完される
- [ ] 納入先住所がホバーでツールチップ表示される
- [ ] 「（一般作業）」選択時は補完欄が空欄になる
- [ ] 作業内容が単一選択（1 行 = 1 作業）になっている
- [ ] 自由記述（detail）が 500 文字以内に制限される
- [ ] ガントの予定が日報の初期行に自動表示される（ガント連携 ON）
- [ ] 保存ボタンで DailyReports シートに記録される

**メール送信**
- [ ] 送信モードが「下書きを作成」（既定）と「送信」を選択できる
- [ ] 送信先が `active` な Staff 全員になっている
- [ ] CC/BCC は空（設定なし）になっている
- [ ] 1 日に同じスタッフから複数回送信できる
- [ ] 送信ボタン押下後に MailQueue に pending が追加される
- [ ] EXE が処理後に status が `drafted` / `sent` になる
- [ ] 失敗時に UI バッジ（badge-failed）が表示される
- [ ] 失敗時に管理者へ通知メールが送信される

### 二重送信防止（同時押下時）

- [ ] 2 台の EXE で同じ pending を同時にピックアップしようとすると、1 台のみ成功する
- [ ] MailQueue シートに同じ `targetStaffId` + `reportDate` の `drafted` が 2 行作成されない
- [ ] Outlook の下書きに同一の件名メールが 2 通作成されない

### 再送機能

- [ ] `failed` 状態のアイテムに対して再送ボタンが表示される
- [ ] 再送ボタン押下で status が `pending` に戻る
- [ ] EXE が次のポーリングで再処理する

### EXE 死活監視

- [ ] EXE 起動中はヘッダに「最終応答 N 分前」と緑ハートビートが表示される
- [ ] EXE を停止してから 5 分後に赤バッジ「要確認」が表示される
- [ ] EXE を再起動するとバッジが消える

### デザイン準拠（wireframe との一致）

wireframe.html をブラウザで開き、以下を目視確認する:

- [ ] 全体背景色が `#f1ebe0`（クリーム調）
- [ ] 見出しフォントが Noto Serif JP（セリフ体）
- [ ] 数字・工番・メールアドレスが JetBrains Mono（等幅）
- [ ] アクセント色が `#b85c3a`（クレイ）
- [ ] タイトルバーのトラフィックライト（赤黄緑）が表示される
- [ ] サイドバー幅が 232px
- [ ] KPI カード 4 列が表示される
- [ ] ペーパーグレイン（薄いノイズテクスチャ）が適用されている

---

## 7. 運用手順

### 7.1 起動・停止

**GAS WebApp**
- 起動: 常時稼働（GAS は停止不要）
- 停止: 「デプロイを管理」→「アーカイブ」（緊急時のみ）

**EXE エージェント**
- 起動: タスクスケジューラによりログオン時に自動起動
- 手動起動: `C:\ProgramData\TeamTaskMail\TeamTaskMail.exe` を実行
- 停止: タスクマネージャー → `TeamTaskMail.exe` → 「タスクの終了」

**EXE のログ確認**

```powershell
# 本日のログをリアルタイムで監視
Get-Content "$env:LOCALAPPDATA\TeamTaskMail\logs\agent_$(Get-Date -Format yyyyMMdd).log" -Wait
```

### 7.2 障害対応

| 症状 | 確認箇所 | 対処 |
|---|---|---|
| UI が開かない | GAS デプロイ状態 | clasp push → デプロイ更新 |
| ガントが空 | Schedules シート | データの有無を確認 |
| EXE 死活バッジが赤 | タスクスケジューラ | EXE を手動再起動 |
| メール送信失敗バッジ | MailQueue シート | エラーメッセージを確認 → Outlook 状態確認 → 再送 |
| 工番が表示されない | 外部マスタ権限 | GAS 実行ユーザーの閲覧権限を確認 |

**EXE が応答しない場合の復旧手順**

1. タスクマネージャーで `TeamTaskMail.exe` が動いているか確認
2. ログ (`%LOCALAPPDATA%\TeamTaskMail\logs\`) を確認し、最終エラーを特定
3. Outlook が開いているか確認（COM 接続に必要）
4. EXE を手動で再起動: `C:\ProgramData\TeamTaskMail\TeamTaskMail.exe` を実行
5. 停止中に生成された `picked` 状態のキューは、`recoverStalePicked()` が自動復旧する（最大 15 分待機）
6. 復旧後、管理者通知メールが届いているか確認

### 7.3 テンプレート変更時のリリースフロー

**ビルド担当者（今泉氏）の作業**

```
1. exe/mail_agent/template.py を編集する
2. バージョン番号を上げる（config.py の VERSION、build.bat の VERSION）
3. build.bat を実行して新しい EXE をビルドする
4. 動作確認（テスト端末で下書き作成を確認）
5. 共有フォルダの releases\<新バージョン>\ に EXE を配置する
6. 各 PC の担当者に更新通知を送る
```

**各 PC 担当者の作業**

```
1. 現在の TeamTaskMail.exe を停止する（タスクマネージャーで終了）
2. 共有フォルダの新バージョン EXE を C:\ProgramData\TeamTaskMail\ に上書きコピーする
3. EXE を手動起動して動作確認する（ハートビートが届くか）
4. タスクスケジューラのタスクは設定変更不要（パス変更がなければ）
```

### 7.4 データバックアップ

**スプレッドシートのバックアップ**

Google スプレッドシートは Google Drive の版管理機能（「変更履歴」）が自動的に保存されるが、定期的に手動コピーを取ることを推奨する。

```
毎月 1 日: スプレッドシートを Google Drive 内でコピー
  「ファイル → コピーを作成 → バックアップ_YYYYMM」
```

**データ保持ポリシー**

| シート | 保持期間 |
|---|---|
| Schedules | 無期限 |
| DailyReports | 無期限 |
| MailQueue | 90 日経過した完了レコード（drafted/sent/failed）を `MailQueue_Archive` シートへ自動移動（仕様 §11） |
| MailQueue_Archive | 無期限（手動アーカイブ・削除は管理者判断） |
| MailLog | 無期限（送信履歴の可監査性確保のため） |
| Staff / WorkTypes / Settings | 無期限 |

**MailQueue 90 日アーカイブ**（仕様 §11 準拠・GAS トリガーで月次実行）

完了レコード（`drafted` / `sent` / `failed`）のうち、`processedAt` が 90 日以上前のものを `MailQueue_Archive` シートへ移動する（**削除ではなく移動**）。アーカイブシートが無い場合は自動生成する。

```javascript
// MailQueueService.gs に追加
function archiveOldMailQueue() {
  const KEEP_DAYS = 90;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - KEEP_DAYS);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queue = ss.getSheetByName(SHEET_NAMES.MAIL_QUEUE);
  let archive = ss.getSheetByName('MailQueue_Archive');
  if (!archive) {
    archive = ss.insertSheet('MailQueue_Archive');
    // ヘッダ行をコピー
    const headers = queue.getRange(1, 1, 1, queue.getLastColumn()).getValues();
    archive.getRange(1, 1, 1, headers[0].length).setValues(headers);
  }

  const data = queue.getDataRange().getValues();
  const headers = data[0];
  const statusIdx      = headers.indexOf('status');
  const processedAtIdx = headers.indexOf('processedAt');

  // 逆順で走査して deleteRow による行ずれを回避
  let archived = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const status = data[i][statusIdx];
    const processedAt = data[i][processedAtIdx];
    if (!['drafted', 'sent', 'failed'].includes(status)) continue;
    if (!processedAt) continue;
    if (new Date(processedAt) >= cutoff) continue;

    archive.appendRow(data[i]);   // アーカイブへ移動
    queue.deleteRow(i + 1);       // queue から削除
    archived++;
  }
  Logger.log('MailQueue アーカイブ完了: ' + archived + ' 件');
}
```

### 7.5 ログの確認方法

**GAS 実行ログ**

GAS エディタ → 「実行数」 または Stackdriver Logging（Cloud Logging）でエラーを確認:

```
GAS エディタ → 「実行数」 → 失敗した実行を選択 → ログを確認
```

**EXE ログ**

```powershell
# 本日のログ
cat "$env:LOCALAPPDATA\TeamTaskMail\logs\agent_$(Get-Date -Format yyyyMMdd).log"

# エラーのみ抽出
Select-String -Path "$env:LOCALAPPDATA\TeamTaskMail\logs\*.log" -Pattern "\[ERROR\]"
```

**MailLog シート**

MailQueue ID を検索して処理経緯を確認する:

```
MailLog シートを開く → Ctrl+F で対象 mailQueueId を検索
level=error の行が失敗ログ
event=retry の行が再送ログ
```

---

## 8. トラブルシューティング

### EXE が応答しない

**症状**: ヘッダの EXE 死活バッジが赤（5 分以上未応答）

**原因と対処**

| 原因 | 確認方法 | 対処 |
|---|---|---|
| EXE が起動していない | タスクマネージャーで確認 | 手動で EXE を起動 |
| Outlook が起動していない | Outlook を確認 | Outlook を起動 |
| ネットワーク切断 | `ping 8.8.8.8` で確認 | ネットワーク接続を回復 |
| EXE がクラッシュ | ログを確認 | エラーを修正して再ビルド |
| WebApp URL が間違い | 環境変数を確認 | 正しい URL を設定 |

```powershell
# EXE プロセスの確認
Get-Process -Name "TeamTaskMail" -ErrorAction SilentlyContinue

# 環境変数の確認
echo $env:TEAM_TASK_WEBAPP_URL
echo $env:TEAM_TASK_API_TOKEN
```

### メール送信失敗

**症状**: MailQueue の status が `failed`、UI に badge-failed が表示される

**対処手順**

1. MailQueue シートの `errorMessage` 列を確認する
2. よくある原因と対処:

| エラーメッセージ | 原因 | 対処 |
|---|---|---|
| `pywintypes.com_error` | Outlook COM エラー | Outlook を再起動する |
| `ConnectionError` | ネットワーク接続 | PC のネットワークを確認 |
| `timeout` | GAS 応答遅延 | しばらく待って再送 |
| `No such user` | 宛先メールアドレス不正 | Staff シートのメールアドレスを確認 |

3. エラーを解決後、GAS エディタで `retryMailItem(id)` を実行して再送する
4. または UI の再送ボタンを押下する

### スプレッドシート編集競合

**症状**: 複数ユーザーが同時に Schedules や DailyReports を編集して競合が発生する

**設計上の対応**

本アプリでは GAS の `LockService` を DataService の書き込み時に適用することを推奨する:

```javascript
// updateRow_ に LockService を追加（DataService.gs 改修例）
function updateRow_(sheetName, id, updates) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(5000);
    // ... 既存の実装 ...
  } finally {
    lock.releaseLock();
  }
}
```

> ただし `LockService.getDocumentLock()` は同一スプレッドシートへの同時書き込みをシリアライズするため、同時アクセスが多い場合はパフォーマンスに影響する。

### 工番マスタ件数増による性能劣化

**症状**: 起動時のローディングが遅い（目安: 1000 件を超えると体感が変わる）

**対処**

1. **初期対応（無改修）**: 844 件は問題なし。2000 件程度まで JSON 配信は 1 秒以内の見込み。

2. **件数が増加した場合の対応**:
   - フロントのフィルタ処理を Web Worker に移行する（UI スレッドのブロック防止）
   - 工番マスタのキャッシュをブラウザの `sessionStorage` に保存し、ページリロード時は再取得しない

3. **抜本的対応（必要な場合）**:
   - GAS 側で工番をページネーション API に変更する
   - フロントは検索文字入力時に API を呼ぶ（都度検索方式）

---

## 9. 付録

### 9.1 コマンドチートシート

**clasp**

```bash
clasp login              # Google アカウントでログイン
clasp pull               # GAS → ローカルへ取得
clasp push               # ローカル → GAS へプッシュ
clasp open               # GAS エディタをブラウザで開く
clasp deploy             # WebApp 新バージョンをデプロイ（GAS エディタ推奨）
clasp logs               # Stackdriver ログを表示
```

**Python / EXE**

```powershell
# 仮想環境
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# EXE ビルド
cd exe
.\build.bat

# 直接実行（デバッグ用、コンソールあり版）
cd exe
python -m mail_agent.main

# EXE ログ確認
Get-Content "$env:LOCALAPPDATA\TeamTaskMail\logs\agent_$(Get-Date -Format yyyyMMdd).log" -Wait
```

**タスクスケジューラ**

```powershell
# タスク一覧
Get-ScheduledTask -TaskName "TeamTaskMail"

# タスク手動実行
Start-ScheduledTask -TaskName "TeamTaskMail"

# タスク削除（再登録時）
Unregister-ScheduledTask -TaskName "TeamTaskMail" -Confirm:$false
```

**Git**

```bash
git status
git add src/
git commit -m "feat: STEP 3 DataService 実装"
# 開発用ブランチを使う場合:
git push origin develop
# 直接 main にコミットする場合（CLAUDE.md 運用ルール準拠の最小構成）:
# git push origin main
```

> CLAUDE.md の運用ルールでは原則 `main` ブランチで運用する。`develop` ブランチを使うかどうかはチームで合意の上で決定する。

### 9.2 主要ファイル一覧

| パス | 役割 |
|---|---|
| `src/appsscript.json` | GAS ランタイム設定（V8・タイムゾーン・アクセス制限） |
| `src/Code.gs` | `doGet` / `doPost` エントリーポイント |
| `src/Config.gs` | 全定数・シートスキーマ・初期データ |
| `src/SetupService.gs` | `setupAll()` シート生成・マスタ投入 |
| `src/DataService.gs` | CRUD 操作（STEP 3 で作成） |
| `src/MasterService.gs` | 外部マスタ読込（STEP 3 で作成） |
| `src/MailQueueService.gs` | メールキュー API・CAS・再送（STEP 4 で作成） |
| `src/index.html` | UI メイン（STEP 6 で作成） |
| `src/styles.html` | CSS スタイル（wireframe.html から分解） |
| `src/scripts.html` | JavaScript ロジック（STEP 6 で作成） |
| `exe/mail_agent/main.py` | EXE エントリ・メインループ |
| `exe/mail_agent/poller.py` | HTTP クライアント |
| `exe/mail_agent/outlook_client.py` | Outlook COM クライアント |
| `exe/mail_agent/template.py` | 件名・本文テンプレート（ハードコード） |
| `exe/mail_agent/config.py` | EXE 設定（URL・トークン） |
| `exe/mail_agent/logger.py` | ファイルロガー |
| `exe/build.bat` | PyInstaller ビルドスクリプト |
| `exe/requirements.txt` | Python 依存パッケージ |
| `docs/wireframe.html` | **UI デザイン正本（常時参照）** |
| `docs/REQUIREMENTS_v1.0.md` | **設計仕様書 確定版（一次情報源）** |
| `.clasp.json` | clasp 設定（scriptId / rootDir） |

### 9.3 環境変数・スクリプトプロパティ一覧

**GAS スクリプトプロパティ（GAS エディタ → プロジェクトの設定で確認・編集）**

| キー | 役割 | 設定タイミング |
|---|---|---|
| `KOBAN_MASTER_SHEET_ID` | 外部マスタ SS ID | STEP 2 |
| `KOBAN_MASTER_SHEET_NAME` | 工番マスタシート名（既定: `工番マスタ`） | STEP 2 |
| `POLLING_INTERVAL_SECONDS` | EXE ポーリング間隔（既定: `30`） | 自動（setupAll） |
| `WEBAPP_VERSION` | アプリバージョン | 自動（setupAll） |
| `EXE_API_TOKEN` | EXE 認証トークン（ランダム生成） | 自動（setupAll） |

**Settings シート（スプレッドシート上で確認・編集）**

| key | 役割 | 初期値 |
|---|---|---|
| `KOBAN_MASTER_SHEET_ID` | 外部マスタ SS ID | （空欄→要記入） |
| `KOBAN_MASTER_SHEET_NAME` | 工番マスタシート名 | `工番マスタ` |
| `POLLING_INTERVAL_SECONDS` | ポーリング間隔（秒） | `30` |
| `MAIL_DEFAULT_MODE` | メール既定モード | `draft` |
| `WEBAPP_VERSION` | WebApp バージョン | `0.1.0` |
| `LAST_HEARTBEAT_TIMESTAMP` | 最終ハートビート時刻 | （自動更新） |
| `LAST_HEARTBEAT_HOSTNAME` | 最終ハートビート PC 名 | （自動更新） |
| `LAST_DEAD_NOTIFICATION_AT` | EXE 死活アラート最終送信時刻（6 時間以内は再通知抑制） | （自動更新） |
| `ADMIN_EMAIL` | 管理者メールアドレス（障害通知先） | （要記入） |

**EXE 環境変数（各 PC の Windows 環境変数）**

| 環境変数名 | 役割 | 設定値 |
|---|---|---|
| `TEAM_TASK_WEBAPP_URL` | GAS WebApp デプロイ URL | `https://script.google.com/macros/s/<ID>/exec` |
| `TEAM_TASK_API_TOKEN` | EXE 認証トークン | `EXE_API_TOKEN` スクリプトプロパティの値 |
| `TEAM_TASK_POLLING_INTERVAL` | ポーリング間隔（秒） | `30`（省略可能） |

**固定値（ハードコード・変更不要）**

| 項目 | 値 |
|---|---|
| GAS スクリプト ID | `1v1P1s5T1L9E7snpsRQT4Wpm2qJTzvt0-kkmgEZ1ec-81lUW0qdBVMlDX` |
| 外部マスタ SS ID | `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ` |
| clasp rootDir | `./src` |
| ガント1日幅（既定） | `64px` |
| ガント行高 | `56px` |
| EXE 死活アラート閾値 | 5 分以上未応答で赤バッジ |
| MailQueue CAS タイムアウト | **10 分**（仕様 v1.0 §7.4.5 と統一） |
| `recoverStalePicked` 実行間隔 | 5 分（GAS トリガー） |
| `checkExeAlive` 実行間隔 | 1 分（GAS トリガー）／死活通知の重複抑制窓 6 時間 |
| MailQueue 保持期間 | 90 日 |
| 業務データ保持期間 | 無期限 |

---

*本書 v1.0 — 2026-05-08 作成*  
*次版改訂予定: 機能追加・仕様変更時に `IMPLEMENTATION_GUIDE_v1.1.md` として発行*
