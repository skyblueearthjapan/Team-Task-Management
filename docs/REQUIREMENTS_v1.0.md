# 機械設計技術部 タスク管理アプリ 設計仕様書 v1.0

> **確定版。v0.1〜v0.5 を統合した単独の設計正本。**
> v0.1〜v0.5 は履歴として保存するが、以降の実装・運用はすべて本書を参照すること。

---

## 1. 概要

機械設計技術部のスタッフ約6名のタスクを一元管理する Web アプリケーション。Google スプレッドシートをバインド DB として活用し、Google Apps Script (GAS) で WebApp として提供する。日々の予定（ガントチャート）・作業内容・前日までの作業報告を記録し、必要に応じてローカル PC 常駐の Python EXE が Outlook 経由でチーム全員へメール下書き作成または直接送信を行う。

**設計の絶対原則**: UI デザインは `docs/wireframe.html` を正本とする。実装時に色・タイポグラフィ・コンポーネント命名・余白などを wireframe から逸脱してはならない。

---

## 2. 目的とスコープ

### 2.1 目的

- 機械設計技術部スタッフ全員のスケジュール・作業状況をガントチャートで可視化する
- 日次の作業報告を Outlook 経由でチーム共有しやすくする（下書き経由の手動送信を既定とし、必要時のみ直接送信）
- 工番マスタを再入力させず、選択ベースで素早く記録できるようにする

### 2.2 スコープ

**含む**
- ガントチャートによる日次スケジュール管理
- 日次の「本日の作業内容」「前日までの作業報告」入力
- 工番マスタ・スタッフマスタ・作業内容マスタとの連携
- スタッフ単位での「メール下書き作成 / 直接送信」依頼機能（Outlook 経由）
- ローカル常駐 Python EXE による Outlook 連携
- EXE 死活監視（ヘッダ表示 + 赤バッジ + 管理者メール通知）
- 重複送信防止（atomic CAS）と再送機能

**含まない**
- 工数集計・原価管理
- 顧客側への直接配信
- モバイル専用 UI（PC ブラウザ前提）

---

## 3. 利用者と権限

| 区分 | 想定人数 | 操作範囲 |
|---|---|---|
| 機械設計技術部スタッフ | 約6名 | 自分の予定・作業内容・報告を入力 / 全員分を閲覧 / メール下書き・直接送信を選択（全員可） |
| 部内管理者 | 1〜2名 | 上記に加えてマスタ編集・送信履歴閲覧・EXE 死活監視確認 |

- 全員 LINE WORKS ドメイン（`@lineworks-local.info`）のユーザー
- WebApp のアクセス制御は LINE WORKS ドメイン制限
- Staff マスタのメールアドレスは管理者がシート上で直接編集可能

---

## 4. 全体アーキテクチャ

### 4.1 構成図

```
┌──────────────────────────────────────────────────────────────┐
│                       ブラウザ (PC)                            │
│   GAS WebApp (HTML + Tailwind CDN + Vanilla JS)               │
│   wireframe.html を正本としたデザインで実装                      │
└────────────┬─────────────────────────────┬────────────────────┘
             │ google.script.run           │ HTTP (EXE→WebApp)
             ▼                             ▼
┌──────────────────────────────────────────────────────────────┐
│              Google Apps Script (V8)                          │
│  Code.gs / Config.gs / DataService.gs / MasterService.gs      │
│  SetupService.gs / MailQueueService.gs                        │
│  index.html / styles.html / scripts.html  (HtmlService)       │
│  ─── EXE 専用エンドポイント (/api/mailqueue) ───               │
│       Bearerトークン認証（Settings.EXE_API_TOKEN）               │
└────────────┬─────────────────────────────────────────────────┘
             │ SpreadsheetApp
             ▼
┌────────────────────────────┐  openById  ┌──────────────────────────────────────────────┐
│  ② バインドスプレッドシート   │◀──────────▶│  ① 外部マスタスプレッドシート（読み取り専用）   │
│  （ローカル DB）              │            │  ID: 1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ │
│  Schedules / DailyReports   │            │  シート: 工番マスタ / 社内カレンダーマスタ      │
│  MailQueue / Staff           │            └──────────────────────────────────────────────┘
│  WorkTypes / Settings        │
└────────────┬───────────────┘
             │ 30秒ごとに POST /exec ?token=...   (action=pickMailItem 他 — §7.4.4 参照)
             ▼
┌──────────────────────────────────────────────────────────────┐
│   ローカル PC 常駐 EXE（Python 3.11 + PyInstaller --onefile）   │
│   - MailQueue をポーリング → pending を取得                      │
│   - mode=draft → Outlook 下書き作成 (mail.Save())              │
│   - mode=send  → Outlook 直接送信   (mail.Send())              │
│   - メール本文テンプレートは EXE 内ハードコード                   │
│   - GAS から差し込み変数（JSON）のみ受け取り本文を生成             │
│   - 1分ごとに Settings.lastHeartbeat を更新（死活監視）           │
└────────────┬─────────────────────────────────────────────────┘
             ▼
     Outlook（下書きフォルダ または 送信済みフォルダ）
             ▼
     技術部スタッフ全員へ配信
```

### 4.2 採用技術一覧

| レイヤ | 採用技術 | 備考 |
|---|---|---|
| Frontend | Vanilla JS + Tailwind CDN | wireframe.html の CSS トークンを共通スタイルへ展開 |
| ガントチャート | 自作（CSS Grid + JS ドラッグ） | wireframe.html の実装スタイルを踏襲 |
| Backend | Google Apps Script (V8) | clasp で管理（技術部で運用、GitHub push 含む） |
| DB (ローカル) | バインドスプレッドシート | シート＝テーブル |
| DB (外部マスタ) | 既存スプレッドシート（読み取り専用） | openById で参照 |
| Deploy | clasp → GAS WebApp | LINE WORKS ドメイン制限。ステージングなし（本番1本運用） |
| ローカル送信エージェント | Python 3.11 + pywin32 | `win32com.client.Dispatch("Outlook.Application")` |
| EXE 化 | PyInstaller `--onefile --noconsole` | 配布は単一 .exe ファイル |
| フォント | Inter / Noto Sans JP / Noto Serif JP / JetBrains Mono | Google Fonts CDN |

---

## 5. 画面構成

### 5.1 wireframe.html を正本として参照

`docs/wireframe.html` に承認済みデザインが格納されている。実装は wireframe.html を分解・実装することで進める。**wireframe からの逸脱は禁止。色・コンポーネント命名・余白・タイポグラフィはすべて wireframe を正本とする。**

### 5.2 主要画面要素のレイアウト

```
┌─────────────────────────────────────────────────────────────────────┐
│ タイトルバー（デスクトップアプリ風）                                     │
│  ●●● [機械設計技術部 — タスク管理]  [EXEエージェント稼働中 · N秒前] ●   │
├───────────────┬─────────────────────────────────────────────────────┤
│ サイドバー      │ メインエリア                                          │
│ 232px 幅      │                                                       │
│               │ ── ページヘッダ ──                                     │
│ ▦ ガント/日報  │  [本日の予定 と 業務報告]  [‹前月] [本日] [次月›] [幅▼] │
│ ◉ 本日タスク  │                                                       │
│ ✉ メール下書き│ ── 統計カード（4列） ──                                  │
│ ◫ 案件一覧    │  本日のタスク件数 / 送信待ち / 送信済 / 要対応           │
│ ⊟ レポート    │                                                       │
│               │ ── ガントチャート ──                                    │
│ [スタッフ一覧] │  ヘッダ: 工番別カラー凡例                               │
│ 山田 太郎     │  行: スタッフ6名 × 列: 日付（過去30日〜未来90日）         │
│ 鈴木 花子     │  本日を中央に初期スクロール                              │
│ 田中 一郎     │                                                       │
│ 佐藤 二郎     │ ── 日報カードグリッド（3列） ──                          │
│ 高橋 三郎     │  スタッフ1人1カード / 本日の作業内容 + 前日までの報告      │
│ 渡辺 四郎     │  フッタ: [メール下書き作成] or [再送する] ボタン          │
│               │                                                       │
│ [案件一覧]    │                                                       │
│ K-12345 ...  │                                                       │
└───────────────┴─────────────────────────────────────────────────────┘
```

### 5.3 デザインシステム（色・タイポ・コンポーネント）

デザインは wireframe.html の CSS Variables を正本とする。以下は実装時に共通 CSS へ展開するトークン一覧。

#### カラートークン

| トークン | 値 | 用途 |
|---|---|---|
| `--bg-base` | `#f1ebe0` | ページ背景（クリーム調ペーパー） |
| `--bg-panel` | `#fbf8f1` | カード背景 |
| `--bg-row` | `#f6f1e6` | 偶数行背景 |
| `--bg-elev` | `#ede5d4` | ヘッダストリップ |
| `--bg-soft` | `#f7f2e6` | フッタ・入力背景 |
| `--line` | `#e1d8c5` | 通常ボーダー |
| `--line-2` | `#cfc5ae` | 中ボーダー |
| `--line-3` | `#b8ad94` | 強ボーダー |
| `--txt-1` | `#2a2520` | 本文主色 |
| `--txt-2` | `#6b6157` | 本文副色 |
| `--txt-3` | `#9a9080` | 補助テキスト |
| `--txt-4` | `#c2b9a8` | プレースホルダ |
| `--accent` | `#b85c3a` | クレイアクセント |
| `--accent-2` | `#a04d2e` | アクセントホバー |
| `--accent-soft` | `rgba(184,92,58,0.10)` | アクセント背景 |
| `--ok` | `#5d7355` | 成功 |
| `--warn` | `#b8862a` | 警告 |
| `--err` | `#a13b2a` | エラー |

#### バーパレット（ガントチャート・工番別色分け）

| クラス名 | 色値 | 用途例 |
|---|---|---|
| `.bar-clay` | `#b85c3a` | 工番1（K-12345 等） |
| `.bar-plum` | `#6b4060` | 工番2 |
| `.bar-ochre` | `#c08a2a` | 工番3（文字色 `#2a2520`） |
| `.bar-burgundy` | `#8e2f30` | 工番4 / 一般作業 |
| `.bar-moss` | `#5d7355` | 工番5 |
| `.bar-indigo` | `#3d4f6e` | 工番6 |
| `.bar-stone` | `#6e6357` | 不在・休暇 |

#### タイポグラフィ

| クラス | フォント | 用途 |
|---|---|---|
| `.serif` | Noto Serif JP | 見出し（h1, h2, section-title） |
| `.mono` | JetBrains Mono | 工番・日付・数値 |
| 本文 | Inter + Noto Sans JP | 通常テキスト |

#### 主要コンポーネント

| コンポーネント | クラス | 説明 |
|---|---|---|
| ウィンドウシェル | `.window` | 角丸14px・シャドウ付きデスクトップアプリ風枠 |
| タイトルバー | `.titlebar` | トラフィックライト（`.traffic .r/.y/.g`）＋セリフタイトル |
| パネル | `.panel` | カード背景（`--bg-panel`）＋`--line` ボーダー |
| サイドバー項目 | `.nav-item` | ホバー時クレイ薄色。アクティブ時パネル背景 |
| バッジ | `.badge-pending/.drafted/.sent/.failed/.idle` | ステータス表示 |
| ピル | `.pill` / `.pill-accent` | 日付チップ（JetBrains Mono） |
| ボタン | `.btn-primary` / `.btn-secondary` / `.btn-accent` | 各種アクション |
| ハートビート | `.heartbeat` | EXE 生存確認ドット（pulse アニメーション） |
| ルール装飾線 | `.rule` | セクション区切り（`本日の作業内容` / `前日までの報告`） |

---

## 6. データモデル

### 6.1 ローカル DB（バインドスプレッドシート）

バインドスプレッドシート ID（実装時に確認）: `1v1P1s5T1L9E7snpsRQT4Wpm2qJTzvt0-kkmgEZ1ec-81lUW0qdBVMlDX`（仮。Script Properties に格納）

| シート名 | 区分 | 用途 |
|---|---|---|
| `Schedules` | トランザクション | ガントの期間バー（予定） |
| `DailyReports` | トランザクション | 本日の作業内容・前日までの作業報告 |
| `MailQueue` | トランザクション | メール送信依頼キュー（EXE が監視） |
| `MailQueue_Archive` | アーカイブ | MailQueue の完了レコード（90日経過後に自動移動）。§11 参照 |
| `MailLog` | ログ | MailQueue のステータス遷移ログ（picked/drafted/sent/failed/retry/mismatch）。可監査性確保のため別シートで保持 |
| `Staff` | マスタ | スタッフ名・メールアドレス・表示順 |
| `WorkTypes` | マスタ | 作業内容種別（13項目） |
| `Settings` | 設定 | 外部マスタ参照ID・EXE API トークン・ハートビートタイムスタンプ |

### 6.2 外部マスタ（工番 / 社内カレンダー）

外部マスタスプレッドシート ID: `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`（読み取り専用）

- GAS 起動時に `SpreadsheetApp.openById(MASTER_ID)` で全件読み取り
- フロントへ JSON で一括配信（推定ペイロード 100〜200KB）
- 本アプリから外部マスタへの書き込みは禁止
- キャッシュ更新: リロード（F5）または明示更新ボタンで再取得

### 6.3 各シートの完全な列定義

#### Schedules（ガント予定）

| 列 | 型 | 必須 | 説明 |
|---|---|---|---|
| `id` | string | ○ | UUID |
| `staffId` | string | ○ | Staff.id への外部キー |
| `startDate` | date (YYYY-MM-DD) | ○ | 開始日 |
| `endDate` | date (YYYY-MM-DD) | ○ | 終了日（startDate 以上） |
| `kobanCode` | string \| null | – | 工番コード（一般作業時 null） |
| `workTypeId` | string \| null | – | WorkTypes.id |
| `note` | string | – | 備考（自由記述） |
| `lane` | int | ○ | 同一スタッフ内の表示レーン番号（1〜、重複時に自動採番） |
| `createdAt` | datetime | ○ | 作成日時 |
| `updatedAt` | datetime | ○ | 更新日時 |

#### DailyReports（日報）

| 列 | 型 | 必須 | 説明 |
|---|---|---|---|
| `id` | string | ○ | UUID |
| `staffId` | string | ○ | Staff.id |
| `reportDate` | date (YYYY-MM-DD) | ○ | 報告対象日 |
| `section` | enum | ○ | `today`（本日の作業内容）/ `yesterday`（前日までの報告） |
| `seq` | int | ○ | 表示順（1〜6。1スタッフ1日あたり最大6行） |
| `periodStart` | date | ○ | 期間開始 |
| `periodEnd` | date | ○ | 期間終了（periodStart 以上） |
| `kobanCode` | string \| null | – | 工番コード（一般作業時 null） |
| `workTypeId` | string \| null | – | WorkTypes.id |
| `detail` | string | – | 自由記述（上限 500 文字） |
| `linkedScheduleId` | string \| null | – | 連携元の Schedules.id |
| `createdAt` | datetime | ○ | 作成日時 |
| `updatedAt` | datetime | ○ | 更新日時 |

#### MailQueue（メール送信依頼キュー）

| 列 | 型 | 必須 | 説明 |
|---|---|---|---|
| `id` | string | ○ | UUID（レコードの一意識別子） |
| `requestedBy` | string | ○ | 依頼者の Staff.id |
| `targetStaffId` | string | ○ | 報告対象の Staff.id |
| `targetStaffName` | string | ○ | 一致確認用スナップショット（Staff.name をコピー） |
| `targetStaffEmail` | string | ○ | 一致確認用スナップショット（Staff.email をコピー） |
| `reportDate` | date | ○ | 対象日 |
| `mode` | enum | ○ | `draft`（下書き作成）/ `send`（直接送信） |
| `toAddresses` | string | ○ | **カンマ区切りメールアドレス文字列**（例: `a@x.com,b@x.com`）。Staff active 全員から自動生成。Sheets 上で目視確認できる形式を採用（JSON 配列ではなく単純な CSV）。EXE 側は `split(',')` でパース |
| `ccAddresses` | string | – | カンマ区切り。**v1.0 では常に空文字列**。将来拡張用カラム |
| `subjectVars` | JSON string | ○ | 件名差し込み変数（後述 §7.4.7） |
| `bodyVars` | JSON string | ○ | 本文差し込み変数（後述 §7.4.7） |
| `status` | enum | ○ | `pending` / `picked` / `drafted` / `sent` / `failed` |
| `pickedBy` | string \| null | – | EXE 識別子（マシン名 + プロセス ID 等） |
| `pickedAt` | datetime \| null | – | EXE が取得した日時（CAS タイムアウト判定に使用） |
| `processedAt` | datetime \| null | – | 処理完了日時 |
| `errorMessage` | string \| null | – | 失敗時の理由 |
| `previousRequestId` | string \| null | – | 再送時に元の MailQueue.id を設定（再送チェーン追跡用） |
| `createdAt` | datetime | ○ | レコード作成日時 |

#### MailLog（メールキュー ステータス遷移ログ）

MailQueue のステータス遷移・エラー・再送イベントを時系列で記録するログシート。MailQueue のレコード自体は 90 日後にアーカイブされるが、本ログは送信履歴の可監査性確保のため**長期保持**する。

| 列 | 型 | 必須 | 説明 |
|---|---|---|---|
| `id` | string | ○ | UUID |
| `mailQueueId` | string | ○ | MailQueue.id への外部キー（再送チェーンも辿れる） |
| `timestamp` | datetime | ○ | イベント発生日時（ISO 8601） |
| `level` | enum | ○ | `info` / `warn` / `error` |
| `event` | enum | ○ | `picked` / `drafted` / `sent` / `failed` / `retry` / `mismatch` / `timeout_recover` |
| `message` | string | – | 補足メッセージ（エラー詳細・再送理由等） |

#### Staff（スタッフマスタ）

| 列 | 型 | 必須 | 説明 |
|---|---|---|---|
| `id` | string | ○ | スタッフ ID（UUID または連番） |
| `name` | string | ○ | 表示名 |
| `email` | string | ○ | Outlook メールアドレス（管理者がシート上で編集可） |
| `role` | enum | ○ | `staff` / `admin` |
| `displayOrder` | int | ○ | ガント・日報カードの表示順 |
| `active` | boolean | ○ | 退職等で `false` に設定するとメール送付先から除外 |

#### WorkTypes（作業内容マスタ）

| 列 | 型 | 必須 | 説明 |
|---|---|---|---|
| `id` | string | ○ | `wt01`〜`wt13` |
| `name` | string | ○ | 作業内容名 |
| `displayOrder` | int | ○ | プルダウン表示順 |
| `active` | boolean | ○ | 非アクティブ時はプルダウンに表示しない |

**WorkTypes 初期投入データ（カテゴリ分類なし・フラット・単一選択）**

| id | name | displayOrder |
|---|---|---|
| wt01 | 構想図作成 | 1 |
| wt02 | 承認図作成 | 2 |
| wt03 | バラシ図作成 | 3 |
| wt04 | 第三者チェック | 4 |
| wt05 | 材料取り・仕様検討 | 5 |
| wt06 | 設計検討 | 6 |
| wt07 | 出図準備 | 7 |
| wt08 | 出荷後図面修正 | 8 |
| wt09 | 購入部品手配・在庫部品確認 | 9 |
| wt10 | 試運転調整・動作確認 | 10 |
| wt11 | 加工指示 | 11 |
| wt12 | 現場対応・現場工事対応 | 12 |
| wt13 | 出張・打ち合わせ | 13 |

#### Settings（共通設定）

| キー | 説明 |
|---|---|
| `MASTER_SPREADSHEET_ID` | 外部マスタスプレッドシート ID（実装側のスクリプトプロパティでは `KOBAN_MASTER_SHEET_ID` のキー名を使用。同一値） |
| `KOBAN_MASTER_SHEET_NAME` | 工番マスタシート名（既定: `工番マスタ`） |
| `EXE_API_TOKEN` | EXE 認証トークン（Bearer） |
| `LAST_HEARTBEAT_TIMESTAMP` | EXE が最後に更新したタイムスタンプ（ISO 8601） |
| `LAST_HEARTBEAT_HOSTNAME` | 最終ハートビートを送ってきた PC ホスト名 |
| `LAST_DEAD_NOTIFICATION_AT` | EXE 死活アラートを最後に送信した日時（重複通知防止用。同一 EXE に対し 6 時間以内は再通知抑制） |
| `POLLING_INTERVAL_SECONDS` | EXE ポーリング間隔（秒・既定 30） |
| `MAIL_DEFAULT_MODE` | メール送信モード既定値（`draft` / `send`・既定 `draft`） |
| `WEBAPP_VERSION` | WebApp バージョン |
| `ADMIN_EMAIL` | 管理者メールアドレス（障害通知先） |

> **キー名整合の注記**: 仕様書では論理名 `MASTER_SPREADSHEET_ID` を用いるが、実装側のスクリプトプロパティ／Settings シート上のキーは `KOBAN_MASTER_SHEET_ID` を使う。両者は同一値を指す。実装ガイドおよびソースコードの呼称を正本とし、本仕様書での記載は読みやすさのための論理名と扱う。

#### 外部マスタ: 工番マスタ（シート名: `工番マスタ`）

| 列 | 型 | 例 | 本アプリでの扱い |
|---|---|---|---|
| `工番` | string | `LW23012` | プルダウン主キー |
| `受注先` | string | `住友建機㈱` | 自動補完（表示） |
| `納入先` | string | `住友建機㈱` | 自動補完（表示） |
| `納入先住所` | string | `千葉県千葉市…` | ツールチップ表示のみ |
| `品名` | string | `4ton応用機アタッチメント…` | 自動補完（表示） |
| `数量` | int | `1` | UI 非表示（管理画面のみ） |
| `取込日時` | date | `9/1/2025` | UI 非表示（管理画面のみ） |

- 件数: 約 844 行
- 参照戦略: 起動時全件ロード → フロントで Map<工番, レコード> キャッシュ
- プルダウン表示形式: `LW23012  住友建機㈱  4ton応用機アタッチメントポジショナー`（工番＋受注先＋品名）
- 検索: インクリメンタル部分一致フィルタ（ブラウザ内処理）

#### 外部マスタ: 社内カレンダー（シート名: `社内カレンダーマスタ`）

| 列 | 型 | 例 | 説明 |
|---|---|---|---|
| `日付` | date | `2026-04-04` | キー |
| `区分` | enum | `休日` / `出勤土曜` / `祝日` | 3値。**曜日ではなくこの区分が一次情報** |
| `曜日` | string | `土` / `日` | 表示・整合性確認用 |
| `備考` | string | `昭和の日` | 祝日名（ツールチップ表示） |

---

## 7. 機能仕様

### 7.1 ガントチャート

#### 操作仕様

| 操作 | 挙動 |
|---|---|
| セル横ドラッグ | 期間バーを新規作成（開始日〜終了日を設定） |
| バー両端ドラッグ | 期間延長／短縮 |
| バー本体ドラッグ | 日程移動 |
| バーダブルクリック | 詳細編集ダイアログ（工番・作業内容・備考） |
| 保存 | 楽観的 UI 更新 → 数秒デバウンス → GAS サーバ保存 |
| 保存失敗 | バーを赤枠化 + トースト通知 |

#### 表示仕様

| 項目 | 仕様 |
|---|---|
| 期間粒度 | **日単位**（時間・半日単位なし） |
| 初期表示位置 | **「今日」を中央**に自動スクロール |
| 描画範囲 | **過去30日 / 未来90日**（合計121日分） |
| セル幅 | 64px/日（標準）/ 48px/日（狭い）/ 96px/日（広い）で切替。**選択値はブラウザ LocalStorage（キー: `ttm.ganttDayWidth`）に永続化**し、次回起動時に復元する |
| 行 | Staff マスタから取得（displayOrder 順）。同一スタッフ内の重複予定は自動レーン分割で縦積み |
| 色分け基準 | **工番別**（wireframe の `.bar-clay/.bar-plum/.bar-ochre/.bar-burgundy/.bar-moss/.bar-indigo/.bar-stone`） |
| 凡例 | ガントチャート右上に工番別カラースウォッチを表示（wireframe と同一） |

#### 社内カレンダーに基づくガント背景色ロジック

| 区分 | 営業日扱い | ガントセル背景 | 日付ヘッダ |
|---|---|---|---|
| 通常平日（カレンダー未登録の平日） | ○ | `--bg-panel`（白） | 通常表示 |
| **出勤土曜**（区分=出勤土曜） | **○** | `--bg-panel`（白） | 曜日「土」+ 小さく「出勤」バッジ |
| 休日（区分=休日 / 未登録の土日） | × | `#f0e9d8`（weekend色） | 土: 曜日色 `#4a5b6e` / 日: `--err` 色 |
| 祝日（区分=祝日） | × | `#f0e9d8` + 微赤帯 | 日番号 `--err` 色 + 備考をツールチップ |
| 本日 | – | `rgba(184,92,58,0.04)` | 下線2px（`--accent`）+ アクセント文字色 |

> **設計原則**: 土曜が必ず休みとは決め打ちしない。社内カレンダーの `区分` を一次情報とする。未登録の土日は通常休日扱い。

### 7.2 日報入力

#### 概要

下部エリアにスタッフ1人1カードのグリッド表示（3列）。wireframe の `article.panel` コンポーネントを使用。

#### 入力フィールド（1行あたり）

| フィールド | 入力方法 | 備考 |
|---|---|---|
| 期間 | 日付ピッカー or ガントから自動連携 | periodStart〜periodEnd |
| 工番 | プルダウン（工番マスタ・インクリメンタル検索） | 先頭に「（一般作業）」固定エントリ |
| 受注先 | 自動補完（工番選択時に入力） | 読み取り専用 |
| 納入先 | 自動補完（工番選択時に入力） | 読み取り専用。住所はツールチップ |
| 品名 | 自動補完（工番選択時に入力） | 読み取り専用 |
| 作業内容 | プルダウン（WorkTypes 13項目） | **単一選択**（1行=1作業内容） |
| 詳細 | テキスト入力 | 任意。**上限 500 文字** |

- 「（一般作業）」選択時は自動補完欄が空欄になり、詳細のみ入力可
- 1スタッフ1日あたり最大 6 行（seq=1〜6）
- 保存単位: **日次レコード**（毎朝リセットではなく履歴蓄積）

#### ガント↔日報の自動連携

- 既定 **ON**: ガント上で設定したバーが日報の対応行へ自動初期表示される
- 手動での追加・編集・削除が可能
- 連携元バーの ID は `linkedScheduleId` に保持

#### セクション区分

| section 値 | 表示名 | 意味 |
|---|---|---|
| `today` | 本日の作業内容 | 当日分 |
| `yesterday` | 前日までの報告 | 前日以前の作業（完了報告等） |

### 7.3 マスタ連携

- **起動時ロード対象**: Staff / WorkTypes / Settings（バインド DB）＋ 工番マスタ全件 / 社内カレンダー全件（外部マスタ）
- **工番マスタキャッシュ**: フロントで `Map<工番コード, レコード>` を保持。プルダウン選択時に O(1) で自動補完
- **祝日名ツールチップ**: ガントの日付セルにマウスホバーで表示（備考カラムの値）
- **納入先住所ツールチップ**: 日報フォームの行にマウスホバーで表示
- **マスタ更新**: 外部マスタは読み取り専用。Staff / WorkTypes はシート上で管理者が直接編集

### 7.4 メール送信

#### 7.4.1 UI フロー（モーダル・送信モード選択）

1. 日報カードフッタの **「メール下書き作成」ボタン**を押下（wireframe `.btn-accent`）
2. 送信確認モーダルが表示される（wireframe §5.3 準拠）:

```
┌─────────────────────────────────────────────────────────────┐
│  メール送信内容の確認                                          │
├─────────────────────────────────────────────────────────────┤
│  作業スタッフ: 山田 太郎                                       │
│  登録メールアドレス: yamada@lineworks-local.info               │
│       └─ ✅ 一致確認OK / ⚠ 不一致（ボタン非活性）               │
│                                                             │
│  送付先（Staff マスタ active 全員）:                          │
│    suzuki@example.com                                       │
│    tanaka@example.com  ...                                  │
│                                                             │
│  送信モード:                                                  │
│    ⚪ 下書きを作成（推奨・既定）                               │
│    ⚪ 直接送信する                                            │
│                                                             │
│  本文差し込みデータプレビュー:                                  │
│  ┌─────────────────────────────────────────────────────┐   │
│  │ ▼ 本日の作業内容                                     │   │
│  │ 1. 5/8〜5/9 LW23012 住友建機㈱ 製品A / 図面作成 …   │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                             │
│          [キャンセル]          [依頼を登録する]               │
└─────────────────────────────────────────────────────────────┘
```

3. **「依頼を登録する」** で `MailQueue` に `status='pending'` でレコード追加
4. 送信モード既定値: **「下書きを作成」**。直接送信はラジオを明示的に切り替えた場合のみ
5. 送付先: Staff マスタの `active=true` 全員。CC/BCC なし（必要なら後から追加）
6. 複数回送信: **制限なし**（制限しない）

> 注意: モーダルで表示されるのは差し込みデータのプレビューであり、最終的な本文整形は EXE 側で行う。

#### 7.4.2 スタッフとメールアドレスの一致確認（4ステップ）

| タイミング | チェック内容 | NG 時の挙動 |
|---|---|---|
| ① モーダル表示時（GAS 側） | `targetStaffId` の Staff マスタ `email` を取得し UI に明示 | 未登録の場合は「依頼を登録する」ボタン非活性 |
| ② 依頼登録時（GAS 側） | `targetStaffName` と `targetStaffEmail` を MailQueue レコードにスナップショット保存 | 保存失敗時は登録キャンセル |
| ③ EXE 取得直後 | `MailQueue.targetStaffEmail` と Staff マスタ最新 `email` を再照合 | 不一致 → `status='failed'` / `errorMessage='email_mismatch'` / 管理者通知 |
| ④ 直接送信モード時 | EXE が Outlook の現在のプロファイル送信者アドレスと `targetStaffEmail` を照合 | 不一致 → `status='failed'` / `errorMessage='outlook_profile_mismatch'` |

#### 7.4.3 送信モード（draft / send）

| モード | EXE の動作 | MailQueue.status 遷移 |
|---|---|---|
| `draft` | `mail.Save()` → Outlook の下書きフォルダに保存 | `picked` → `drafted` |
| `send` | `mail.Send()` → Outlook の送信トレイ経由で送信 | `picked` → `sent` |

#### 7.4.4 EXE ↔ GAS 間 REST エンドポイント仕様（正本・1箇所集約）

EXE は GAS WebApp の `/exec` 単一エンドポイントに対し、`POST` ボディの `action` キーで操作種別を識別する RPC スタイル API を使用する（`doPost` のディスパッチ）。本節を**唯一の正本**とし、§4.1 構成図および §8.1 EXE フローはここを参照する。

**認証**

- 全リクエストは `?token={EXE_API_TOKEN}` クエリパラメータで Bearer 相当のトークンを付与
- 不一致時は `{ "error": "Unauthorized" }` を返す（HTTP ステータスは GAS 仕様上 200 だが応答ボディで判別）

**ポーリング**

- EXE は **30 秒間隔** で `pickMailItem` を呼び続ける（ロングポーリングではない短周期ポーリング）
- ハートビートは **60 秒間隔** で別途送信

**アクション一覧（POST /exec）**

| action | リクエストボディ | レスポンス | 備考 |
|---|---|---|---|
| `heartbeat` | `{ "action":"heartbeat", "hostname":"PC-A" }` | `{ "ok": true }` | Settings.LAST_HEARTBEAT_TIMESTAMP / LAST_HEARTBEAT_HOSTNAME を更新 |
| `pickMailItem` | `{ "action":"pickMailItem", "hostname":"PC-A" }` | 取得時: 全レコードフィールド（id/mode/toAddresses/subjectVars/bodyVars 等）。pending なし: `null` | **CAS 一体実行**。pending 検索＋ picked 遷移＋ pickedBy/pickedAt 書込みを LockService 配下で 1 リクエストに集約。GET と POST の 2 段階方式は採用しない |
| `completeMailItem` | `{ "action":"completeMailItem", "id":"mq_xxx", "status":"drafted\|sent\|failed", "errorMessage":"..." }` | `{ ... }`（更新結果） | EXE の処理完了後に呼ぶ。MailLog にも自動記録 |

> **設計上の確定事項**: GET と POST の 2 段階（取得 → ピック）はレースコンディションを生じるため**採用しない**。常に POST `pickMailItem` で取得＋遷移を一体実行する。

#### 7.4.5 重複送信防止メカニズム（atomic CAS / pickedBy / タイムアウト復旧）

複数 EXE が同じレコードを同時に取得しないよう、GAS WebApp 側が Atomic Compare-and-Swap (CAS) を実行する。

**ステータス遷移図**

```
                     ┌─────────────────────────────────┐
                     │ 再送（新レコード生成）              │
                     │ previousRequestId でリンク         │
                     └────────────────┬────────────────┘
                                      │
  [UI: 依頼登録]  [EXE: CAS 成功]     │         [EXE: 完了]
  pending ──────▶ picked ──────────▶ drafted   ←─────────┐
                     │                         sent      │
              [タイムアウト]           ┌──────▶ failed ──┘
              (>10分)                 │
              │                      │  [EXE: 処理失敗]
              ▼                      └─── (retry / 管理者通知)
           pending（復旧）
           ※ GAS TimeTrigger が定期実行
```

**CAS フロー（GAS WebApp 側・action=pickMailItem の擬似コード）**

```javascript
// doPost 内: action === 'pickMailItem' のディスパッチ先
// LockService 配下で「pending 検索 → picked 書込み」を 1 リクエストで一体実行する
function pickMailItem(exeHostname) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 10 秒待機
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MailQueue');
    const data = sheet.getDataRange().getValues(); // data[0] はヘッダ行
    const headers = data[0];
    const idIdx       = headers.indexOf('id');
    const statusIdx   = headers.indexOf('status');
    const pickedByIdx = headers.indexOf('pickedBy');
    const pickedAtIdx = headers.indexOf('pickedAt');

    // ヘッダ行をスキップして探索（i=1 から開始 / シートの行番号は i+1）
    for (let i = 1; i < data.length; i++) {
      if (data[i][statusIdx] === 'pending') {
        // CAS 相当: Lock 配下で読み取った直後に同じ行へ書き込むため、
        // pending 二重チェックは省略可能（Lock により他プロセスは入れない）
        sheet.getRange(i + 1, statusIdx   + 1).setValue('picked');
        sheet.getRange(i + 1, pickedByIdx + 1).setValue(exeHostname);
        sheet.getRange(i + 1, pickedAtIdx + 1).setValue(new Date().toISOString());
        SpreadsheetApp.flush(); // 即時コミット

        // 更新後のレコードを {ヘッダ:値} で返す
        const row = sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0];
        return Object.fromEntries(headers.map((h, j) => [h, row[j]]));
      }
    }
    return null; // pending なし
  } finally {
    lock.releaseLock();
  }
}
```

**タイムアウト復旧（GAS TimeTrigger）**

CAS タイムアウト閾値は **10 分**（ポーリング間隔 30 秒 × 安全マージン。ネットワーク不調・Outlook COM 待機等を考慮した余裕値）。トリガー実行間隔は **5 分**。これにより EXE 障害時でも最大 15 分（10 分タイムアウト + 5 分トリガー間隔）で自動復旧する。

```javascript
// GAS TimeTrigger: 5 分ごとに実行
function recoverStalePicked() {
  const TIMEOUT_MS = 10 * 60 * 1000; // 10 分（実装ガイドと統一）
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MailQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusIdx   = headers.indexOf('status');
  const pickedAtIdx = headers.indexOf('pickedAt');
  const pickedByIdx = headers.indexOf('pickedBy');
  const now = Date.now();

  // i=1 から（ヘッダ行スキップ）
  for (let i = 1; i < data.length; i++) {
    if (data[i][statusIdx] !== 'picked') continue;
    const pickedAt = data[i][pickedAtIdx] ? new Date(data[i][pickedAtIdx]).getTime() : 0;
    if (now - pickedAt > TIMEOUT_MS) {
      sheet.getRange(i + 1, statusIdx   + 1).setValue('pending');
      sheet.getRange(i + 1, pickedByIdx + 1).setValue('');
      sheet.getRange(i + 1, pickedAtIdx + 1).setValue('');
      // MailLog にも timeout_recover イベントを記録
    }
  }
}
```

#### 7.4.6 再送メカニズム（UI 操作・previousRequestId・警告表示）

**再送 UI**

- 送信済み（`drafted` / `sent`）カードのフッタに **「再送する」ボタン**（wireframe `.btn-secondary`）を表示
- 送信失敗（`failed`）カードのフッタに **「↻ 再試行」ボタン**を表示

**再送フロー（仕様確定 — 実装ガイド STEP 4 もこのフローに準拠する）**

```
[UI: 再送ボタン押下]
  └─ 既存レコードの status が 'sent' の場合:
       → 確認ダイアログ「既に送信済みです。本当に再送しますか？」（警告）
       → ユーザがキャンセルした場合は何もしない
  └─ status が 'failed' / 'drafted' の場合は即時実行:
       → GAS が既存 MailQueue レコードを **複製**（既存レコードは一切変更しない）
       → 新レコードを生成（フィールドは下表のとおり）
       → 新レコードを MailQueue に appendRow で追加
       → MailLog に event='retry' イベントを記録
       → EXE が次のポーリングで新レコードを取得・処理
```

**新レコードのフィールド仕様**

| 列 | 値の取得方法 |
|---|---|
| `id` | 新規 UUID（`mq_` + 16桁） |
| `requestedBy` | 元レコードからコピー |
| `targetStaffId` / `targetStaffName` / `targetStaffEmail` | 元レコードからコピー |
| `reportDate` | 元レコードからコピー |
| `mode` | 元レコードからコピー（`draft` / `send`） |
| `toAddresses` / `ccAddresses` | 元レコードからコピー |
| `subjectVars` / `bodyVars` | 元レコードからコピー |
| `status` | **`pending`**（必ず） |
| `pickedBy` / `pickedAt` / `processedAt` / `errorMessage` | **空文字列**（リセット） |
| `previousRequestId` | **元レコードの `id`** をセット |
| `createdAt` | 現在時刻 |

> **重要**: 仕様 v0.x の旧版で「既存レコードの status を pending に戻す」と読める記述があったが、v1.0 ではこれを**明示的に禁止**する。元レコードの履歴を破壊せず、必ず新レコードを生成して `previousRequestId` で連鎖させる。

**previousRequestId チェーン**

再送を繰り返した場合、`previousRequestId` で元レコードへのリンクが連鎖する。送信履歴画面でリンクを辿ることで再送経緯が追跡可能。

#### 7.4.7 件名・本文フォーマット（EXE 内ハードコードテンプレート）

**メール本文テンプレートは EXE のソースコード内に定数として定義する。** GAS 側は差し込み変数（JSON）のみを `MailQueue.subjectVars` / `MailQueue.bodyVars` に格納する。テンプレート変更時は EXE の再ビルド・再配布が必要。

**差し込み変数 JSON（GAS → EXE）**

```json
{
  "subjectVars": {
    "staffName": "山田 太郎",
    "reportDate": "2026/05/08"
  },
  "bodyVars": {
    "staffName": "山田 太郎",
    "reportDate": "2026/05/08",
    "todayItems": [
      {
        "periodStart": "2026/05/08",
        "periodEnd": "2026/05/09",
        "kobanCode": "LW23012",
        "customer": "住友建機㈱",
        "productName": "4ton応用機アタッチメントポジショナー",
        "workType": "承認図作成",
        "detail": "外形図の修正対応"
      }
    ],
    "yesterdayItems": [
      {
        "periodStart": "2026/05/07",
        "periodEnd": "2026/05/07",
        "kobanCode": "LW24001",
        "customer": "㈱NICHIJO",
        "productName": "SK2000R",
        "workType": "第三者チェック",
        "detail": "チェック完了"
      }
    ]
  }
}
```

**件名テンプレート（EXE ハードコード）**

```
【機械設計技術部】{staffName} 業務報告 {reportDate}
```

**本文テンプレート（EXE ハードコード・プレーンテキスト）**

```
お疲れ様です。{staffName} です。
本日の業務をご報告いたします。

▼ 本日の作業内容
{todayItems を以下フォーマットで連番出力}
{N}. {periodStart}〜{periodEnd}  {kobanCode}  {customer}  {productName}  /  {workType}  {detail}

▼ 前日までの作業報告
{yesterdayItems を同様に出力}

以上、よろしくお願いいたします。
```

> 「（一般作業）」の場合は `kobanCode` / `customer` / `productName` が空文字列になる。EXE 側でこれらが空の場合は出力を省略する。

### 7.5 EXE 死活監視

#### 表示ルール

| 条件 | ヘッダ表示 | バッジ色 | 追加アクション |
|---|---|---|---|
| 1分以内に heartbeat あり | `EXEエージェント稼働中 · N秒前` | `.heartbeat`（緑パルス） | なし |
| 1〜5分間 heartbeat なし | `最終応答 N分前` | 黄色バッジ | なし |
| **5分以上 heartbeat なし** | `EXE 応答なし · N分前` | **赤バッジ** | 管理者（Settings.ADMIN_EMAIL）へメール通知。**重複防止: Settings.LAST_DEAD_NOTIFICATION_AT を参照し、前回通知から 6 時間以内であれば再通知をスキップ** |

ヘッダ内の表示は wireframe の `.heartbeat` ドットと `EXEエージェント稼働中 · N秒前` のテキストを使用。

**死活監視の重複通知防止ロジック**

```javascript
// 1 分ごとの GAS TimeTrigger（または UI からの呼び出し）
function checkExeAlive() {
  const last = getSetting('LAST_HEARTBEAT_TIMESTAMP');
  if (!last) return;
  const elapsedMin = (Date.now() - new Date(last).getTime()) / 60000;
  if (elapsedMin < 5) return; // 正常

  // 重複通知抑制
  const lastNotified = getSetting('LAST_DEAD_NOTIFICATION_AT');
  if (lastNotified) {
    const sinceLastNotifyMin = (Date.now() - new Date(lastNotified).getTime()) / 60000;
    if (sinceLastNotifyMin < 360) return; // 6 時間以内は再通知しない
  }

  GmailApp.sendEmail(getSetting('ADMIN_EMAIL'),
    '[タスク管理] EXE 応答なし', 'EXE が ' + Math.floor(elapsedMin) + ' 分応答していません');
  setSetting('LAST_DEAD_NOTIFICATION_AT', new Date().toISOString());
}
```

---

## 8. 外部連携

### 8.1 ローカル送信エージェント（Python EXE）

#### 基本情報

| 項目 | 仕様 |
|---|---|
| 言語・バージョン | Python 3.11+ |
| EXE 化 | PyInstaller `--onefile --noconsole` |
| 配布形式 | 単一 `.exe` ファイル |
| 配置場所 | **各スタッフの業務 PC**（分散配置） |
| 配布方法 | 共有フォルダにビルド済み `.exe` を配置 → 各自コピーで導入 |
| 常駐方法 | Windows タスクスケジューラ（ログオン時起動） |

#### 主要依存ライブラリ

| ライブラリ | 用途 |
|---|---|
| `pywin32` (`win32com.client`) | Outlook COM 操作 |
| `requests` | GAS WebApp エンドポイントへの HTTP 通信 |

#### 動作フロー

```
[起動]
  │
  ├─ Settings から EXE_API_TOKEN を読み込み
  │   （EXE 起動引数 or 設定ファイルで指定）
  │
  └─ メインループ（30秒ごと）— §7.4.4 REST 仕様 参照
       │
       ├─ [60秒ごと] POST /exec?token=...  body={action:'heartbeat',hostname:...}
       │       → Settings.LAST_HEARTBEAT_TIMESTAMP を更新（死活監視用）
       │
       ├─ POST /exec?token=...  body={action:'pickMailItem', hostname:'PC-A'}
       │       ↓ GAS 側で LockService 配下にて pending → picked を一体実行（CAS）
       │
       ├─ レスポンスが null（pending なし）→ 何もせず次の周回へ
       │
       └─ レスポンスにレコードあり:
            ├─ toAddresses（カンマ区切り string）を split(',') でパース
            ├─ subjectVars / bodyVars（JSON string）を json.loads でパース
            ├─ EXE 内ハードコードテンプレートで件名・本文を生成
            │
            ├─ mode='draft':
            │   outlook = win32com.client.Dispatch("Outlook.Application")
            │   mail = outlook.CreateItem(0)
            │   mail.Subject = subject
            │   mail.Body    = body
            │   mail.To      = "; ".join(toAddresses)
            │   mail.Save()  # 下書きフォルダへ
            │   → POST /exec ?token=...  body={action:'completeMailItem', id, status:'drafted'}
            │
            ├─ mode='send':
            │   [④ Outlook プロファイル照合]
            │     不一致 → completeMailItem(id, 'failed', 'outlook_profile_mismatch')
            │   mail.Send()  # 送信トレイ経由
            │   → POST /exec  body={action:'completeMailItem', id, status:'sent'}
            │
            └─ エラー時:
                → POST /exec  body={action:'completeMailItem', id, status:'failed', errorMessage:...}
                → ローカルログ（%LOCALAPPDATA%\TeamTaskMail\logs\agent_YYYYMMDD.log）
                → GAS 側で MailLog 記録 + 管理者メール通知
```

#### バージョン管理

テンプレート変更時の再配布フロー:
1. 技術部担当者がソース変更 → EXE 再ビルド（`pyinstaller --onefile --noconsole main.py`）
2. ビルド済み `.exe` を共有フォルダへ配置（ファイル名にバージョン番号を付与: `TeamTaskMailAgent_v1.2.exe`）
3. 各スタッフが共有フォルダから新 `.exe` をコピーして上書き導入

### 8.2 工番マスタ別ブック参照

- 別ブック ID: `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`（Script Properties `MASTER_SPREADSHEET_ID` に格納）
- 参照方式: `SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getSheetByName('工番マスタ')`
- 本アプリから書き込みは行わない（読み取り専用）
- 同期タイミング: 起動時全件ロード。明示更新ボタンで再取得

### 8.3 社内カレンダー参照（休日 / 出勤土曜 / 祝日 の3区分）

- 参照元: 同上スプレッドシートの `社内カレンダーマスタ` シート
- 区分値: `休日` / `出勤土曜` / `祝日` の3値
- **「出勤土曜」は営業日扱い**（ガントでグレーアウトしない）
- 未登録の土日: 通常休日扱い（区分が一次情報のため、曜日だけで休日判定しない）
- 祝日名（備考列）: ガントの日付セルにマウスホバー時にツールチップで表示

---

## 9. 非機能要件

| 区分 | 要件 |
|---|---|
| パフォーマンス | 楽観的 UI で操作レスポンス 200ms 以内 / 保存反映 5 秒以内 |
| 可用性 | GAS の SLA に準拠（Google のインフラ依存） |
| スケーラビリティ | スタッフ約6名・工番約844件・ポーリング30秒間隔の規模を前提 |
| ブラウザ対応 | PC ブラウザ（Chrome 最新版を基準）。モバイル非対応 |
| ログ | EXE: ローカルファイル `%LOCALAPPDATA%\TeamTaskMail\logs\YYYYMMDD.log`。MailQueue.errorMessage にも記録 |
| 監査 | 主要シートに `createdAt` / `updatedAt` を持たせる |
| バックアップ | Google スプレッドシートの自動バージョン管理に依存 |

---

## 10. セキュリティ・運用

| 項目 | 方針 |
|---|---|
| アクセス制御 | GAS WebApp を LINE WORKS ドメイン制限。`appsscript.json` の `webapp.access` を **`"DOMAIN"`** に設定し、`webapp.executeAs` を **`"USER_DEPLOYING"`** に設定する。デプロイ時に GAS UI で「アクセスできるユーザー: 自分のドメイン内の全員」を選択し、Workspace ドメイン `lineworks-local.info` のユーザのみ実行できる状態とする |
| EXE 認証 | Bearer トークン方式（Settings.EXE_API_TOKEN）。トークンは EXE の設定ファイルまたは起動引数で渡す。Script Properties に保管 |
| Staff マスタ更新 | 管理者のみシート上で直接編集。EXE は照合のみ（書き換え不可） |
| 直接送信モードの誤爆防止 | 送信モードの既定値は「下書きを作成」。「直接送信する」はラジオを明示的に切り替えた場合のみ選択可 |
| テンプレート管理 | EXE ハードコードのため、変更時は再ビルド・再配布。バージョン番号で管理 |
| 複数スタッフ混在環境 | ④ の Outlook プロファイル照合で送信者のなりすましを防止 |
| 機密値管理 | スプレッドシート ID・API トークン・管理者メールは Script Properties に保管。ソースコードにハードコードしない |

---

## 11. データ保持・アーカイブ運用

| シート / データ | 保持方針 |
|---|---|
| `Schedules` | **無期限保持** |
| `DailyReports` | **無期限保持** |
| `MailQueue` 完了レコード（`drafted` / `sent` / `failed`） | **90日後にアーカイブシートへ移動**（GAS TimeTrigger で定期実行） |
| EXE ローカルログ | EXE 側の独自ローテーション（例: 30日以上経過したログファイルを削除） |

**アーカイブ処理（GAS TimeTrigger 擬似コード）**

```javascript
// 毎日深夜に実行
function archiveOldMailQueue() {
  const ARCHIVE_DAYS = 90;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - ARCHIVE_DAYS);

  const ss = SpreadsheetApp.openById(DB_ID);
  const queue = ss.getSheetByName('MailQueue');
  const archive = ss.getSheetByName('MailQueue_Archive') || ss.insertSheet('MailQueue_Archive');

  const rows = queue.getDataRange().getValues();
  const toArchive = rows.filter(r => {
    const isDone = ['drafted', 'sent', 'failed'].includes(r[STATUS_COL]);
    const processedAt = new Date(r[PROCESSED_AT_COL]);
    return isDone && processedAt < cutoff;
  });

  if (toArchive.length > 0) {
    archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, toArchive[0].length)
           .setValues(toArchive);
    // queue から削除（逆順に削除して行ずれを防ぐ）
  }
}
```

---

## 12. デプロイ・配布

### 12.1 GAS WebApp（clasp 運用）

| 項目 | 内容 |
|---|---|
| ソース管理 | clasp で管理。技術部で運用（GitHub への push 含む） |
| デプロイ先 | LINE WORKS ドメイン制限の GAS WebApp（本番1本） |
| ステージング | **当面なし**（本番1本運用） |
| デプロイ手順 | `clasp push` → GAS Web エディタで「デプロイを管理」→ 新バージョンとして公開 |

### 12.2 Python EXE（配布）

| 項目 | 内容 |
|---|---|
| ビルドコマンド | `pyinstaller --onefile --noconsole main.py` |
| 配布方法 | 共有フォルダにビルド済み `.exe` を配置 → 各 PC でコピーして導入 |
| バージョン管理 | ファイル名にバージョン番号付与（例: `TeamTaskMailAgent_v1.0.exe`）。EXE 内に `VERSION = "1.0.0"` を定数として保持 |
| テンプレート変更時 | 中央でビルド → 共有フォルダへ配置 → 各 PC で上書き導入（各自判断 or 管理者連絡） |
| 動作前提 | Windows PC + Outlook がインストール済み・ログイン済み |
| 常駐設定 | Windows タスクスケジューラ（ログオン時起動）で自動起動 |

---

## 13. 用語集

| 用語 | 定義 |
|---|---|
| 工番 | 受注ごとに割り振られる製造番号（例: LW23012）。工番マスタで管理 |
| 社内カレンダー | 外部マスタスプレッドシートの「社内カレンダーマスタ」シート。休日 / 出勤土曜 / 祝日 の3区分を管理 |
| 出勤土曜 | 社内カレンダーで区分=出勤土曜 の土曜日。**営業日扱い**（ガントでグレーアウトしない） |
| 作業内容 | WorkTypes マスタの13項目のいずれか1つ（単一選択） |
| 日報 | DailyReports の1レコード。「本日の作業内容」または「前日までの作業報告」の1行 |
| メール下書き | Outlook の下書きフォルダに保存されたメール（mode=draft 時） |
| EXE | ローカル PC 常駐の Python 送信エージェント（PyInstaller でビルドした .exe ファイル） |
| MailQueue | GAS スプレッドシート上のメール送信依頼キューシート |
| CAS | Compare-and-Swap。MailQueue の重複取得を防ぐ楽観的ロック機構 |
| pickedBy | EXE が MailQueue レコードを取得した際に記録する EXE 識別子（マシン名+PID） |
| 再送 | 既存の MailQueue レコードを複製し、新しい pending レコードとして挿入する操作 |
| previousRequestId | 再送レコードが元レコードの id を保持するフィールド |
| バインド SS | GAS スクリプトにバインドされた（紐付けられた）スプレッドシート。本アプリのローカル DB |
| clasp | Google Apps Script を CLI で管理するツール |

---

## 14. 改版履歴

| バージョン | 日付 | 主要変更点 |
|---|---|---|
| v0.1 | 2026/04/中旬 | 初期ドラフト。全体アーキテクチャ・画面構成・データモデル・機能仕様の初版。確認事項40項目を列挙 |
| v0.2 | 2026/04/下旬 | EXE を Python/PyInstaller に確定。下書き/送信モードの二択を新設。メール本文テンプレートを EXE 内ハードコードに確定。スタッフ↔メール一致確認の4ステップを追加。MailQueue に mode / pickedBy / targetStaffName / targetStaffEmail を追加 |
| v0.3 | 2026/05/上旬 | 外部マスタスプレッドシート ID 確定。WorkTypes 初期13項目確定。社内カレンダーを外部マスタ参照に変更（ローカル Holidays シート廃止）。EXE を各 PC に分散配置と確定。ポーリング間隔30秒確定。デザイン（wireframe.html v0.4 ベースライン）承認 |
| v0.4 | 2026/05/05 | 工番マスタの実列構成確定（7列・844行）。社内カレンダーの実列構成と区分値（休日/出勤土曜/祝日）確定。出勤土曜=営業日扱いを明記。工番選択時の自動補完3列確定 |
| v0.5 | 2026/05/07 | 納入先住所・祝日名をツールチップ表示に確定。WorkTypes カテゴリ分類廃止→フラット13項目。作業内容=単一選択確定。数量・取込日時を UI 非表示に確定 |
| **v1.0** | **2026/05/08** | **確定版。v0.1〜v0.5 を統合。以下の事項を新規確定: ガント日単位/今日中央/過去30日〜未来90日/工番別色分け。日報: 日次レコード保持/ガント連携既定ON/自由記述500文字上限。メール: Staff active全員/CC・BCC無/複数送信許可/失敗時UIバッジ+管理者メール。EXE死活監視: 5分以上で赤バッジ+管理者通知。データ保持: Schedules・DailyReports無期限/MailQueue完了90日後アーカイブ。重複送信防止（atomic CAS + pickedBy + タイムアウト復旧）と再送メカニズム（previousRequestId・警告表示）を具体的に仕様化。clasp 技術部運用/ステージングなし確定** |
| **v1.0 (rev1)** | **2026/05/08** | **最終レビュー反映（Opus 修正適用）。設計判断を確定: ① toAddresses はカンマ区切り string、② CAS タイムアウト 10 分、③ MailLog シートは可監査性のため復活・正式定義、④ EXE↔GAS は POST `/exec` 単一エンドポイント＋action ベースで pickMailItem を CAS 一体実行（GET+POST 2段階は不採用）。再送機能を「新レコード生成＋ previousRequestId 連鎖」に確定（status を pending に戻す方式は禁止）。MailQueue_Archive を §6.1 に明記。DailyReports に createdAt 追加。Settings に LAST_DEAD_NOTIFICATION_AT 等を追加し、死活監視の重複通知（6時間抑制）を明文化。WebApp DOMAIN アクセス制御の具体・LocalStorage によるガント幅永続化・index/styles/scripts.html を追記** |
