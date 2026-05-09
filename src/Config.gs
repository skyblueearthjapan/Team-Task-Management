/**
 * グローバル設定定数。
 * シート名・列スキーマ・スクリプトプロパティのキー・初期データを一元管理する。
 * SetupService をはじめ各サービスはここを単一の正本として参照する。
 */

// ─── シート名 ────────────────────────────────────────────────
const SHEET_NAMES = {
  SCHEDULES:          'Schedules',
  DAILY_REPORTS:      'DailyReports',
  MAIL_QUEUE:         'MailQueue',
  MAIL_QUEUE_ARCHIVE: 'MailQueue_Archive',
  MAIL_LOG:           'MailLog',
  STAFF:              'Staff',
  WORK_TYPES:         'WorkTypes',
  SETTINGS:           'Settings'
};

// ─── 各シートのヘッダ列定義 ─────────────────────────────────
const SHEET_SCHEMA = {
  [SHEET_NAMES.SCHEDULES]: [
    'id',
    'staffId',
    'startDate',
    'endDate',
    'kobanCode',
    'workTypeId',
    'note',
    'lane',
    'createdAt',
    'updatedAt'
  ],
  [SHEET_NAMES.DAILY_REPORTS]: [
    'id',
    'staffId',
    'reportDate',
    'section',           // today / yesterday
    'seq',               // 1〜6
    'periodStart',
    'periodEnd',
    'kobanCode',
    'workTypeId',
    'detail',
    'duration',          // 完了見込み（"2h" / "午前中" / "2日" 等）
    'continued',         // 翌日継続フラグ（"true" / "false" 文字列）
    'linkedScheduleId',
    'createdAt',
    'updatedAt'
  ],
  [SHEET_NAMES.MAIL_QUEUE]: [
    'id',
    'requestedBy',
    'targetStaffId',
    'targetStaffName',
    'targetStaffEmail',
    'reportDate',
    'mode',              // draft / send
    'toAddresses',       // カンマ区切り string（"a@x,b@y"）。JSON 配列は使わない
    'ccAddresses',       // v1.0 では常に空文字列。将来拡張用
    'subjectVars',       // JSON string
    'bodyVars',          // JSON string
    'status',            // pending / picked / drafted / sent / failed
    'pickedBy',
    'pickedAt',
    'processedAt',
    'errorMessage',
    'previousRequestId', // 再送時に元レコード id を設定
    'createdAt'
  ],
  [SHEET_NAMES.MAIL_QUEUE_ARCHIVE]: [
    'id',
    'requestedBy',
    'targetStaffId',
    'targetStaffName',
    'targetStaffEmail',
    'reportDate',
    'mode',
    'toAddresses',
    'ccAddresses',
    'subjectVars',
    'bodyVars',
    'status',
    'pickedBy',
    'pickedAt',
    'processedAt',
    'errorMessage',
    'previousRequestId',
    'createdAt'
  ],
  [SHEET_NAMES.MAIL_LOG]: [
    'id',
    'mailQueueId',
    'timestamp',
    'level',             // info / warn / error
    'event',             // picked / drafted / sent / failed / retry / mismatch / timeout_recover
    'message'
  ],
  [SHEET_NAMES.STAFF]: [
    'id',
    'name',
    'email',
    'role',              // staff / admin / external
                         //   staff    : 通常スタッフ（UI 表示・メール宛先・送信元になりうる）
                         //   admin    : 管理者（同上 + 死活通知受信等の管理権限）
                         //   external : 送信専用メンバー（UI 非表示・メール宛先のみに含まれる）
    'displayOrder',
    'active',            // TRUE/FALSE
    'signatureName',     // メール署名用フルネーム（空なら name にフォールバック）
    'signatureEmail'     // メール署名用メール（空なら email にフォールバック）
  ],
  [SHEET_NAMES.WORK_TYPES]: [
    'id',
    'name',
    'displayOrder',
    'active'
  ],
  [SHEET_NAMES.SETTINGS]: [
    'key',
    'value',
    'description'
  ]
};

// ─── スクリプトプロパティのキー ─────────────────────────────
const SCRIPT_PROP_KEYS = {
  KOBAN_MASTER_SHEET_ID: 'KOBAN_MASTER_SHEET_ID',
  KOBAN_MASTER_SHEET_NAME: 'KOBAN_MASTER_SHEET_NAME',
  POLLING_INTERVAL_SECONDS: 'POLLING_INTERVAL_SECONDS',
  WEBAPP_VERSION: 'WEBAPP_VERSION',
  EXE_API_TOKEN: 'EXE_API_TOKEN',
  GEMINI_API_KEY: 'GEMINI_API_KEY'  // ユーザーが GAS Script Properties に手動登録
};

// ─── 初期データ：Settings ─────────────────────────────────
const DEFAULT_SETTINGS = [
  ['KOBAN_MASTER_SHEET_ID',        '',       '工番マスタ別ブックのスプレッドシートID'],
  ['KOBAN_MASTER_SHEET_NAME',      '工番マスタ', '工番マスタブック内のシート名'],
  ['POLLING_INTERVAL_SECONDS',     '30',     'ローカル EXE が MailQueue をポーリングする間隔（秒）'],
  ['MAIL_DEFAULT_MODE',            'draft',  'メール送信モードの既定値（draft / send）'],
  ['WEBAPP_VERSION',               '1.0.0',  'WebApp バージョン'],
  ['LAST_HEARTBEAT_TIMESTAMP',     '',       'ローカル EXE からの最終ハートビート（自動更新）'],
  ['LAST_HEARTBEAT_HOSTNAME',      '',       '最終ハートビートを送ってきた PC 名（自動更新）'],
  ['LAST_DEAD_NOTIFICATION_AT',    '',       '管理者への前回 EXE 死活通知時刻'],
  ['ADMIN_EMAIL',                  '',       'EXE 死活通知の送信先管理者メール'],
  ['CAS_TIMEOUT_MINUTES',          '10',     'MailQueue picked → pending 復旧の閾値（分）'],
  ['EXE_DEAD_THRESHOLD_MINUTES',   '5',      'EXE 応答なしと判定する閾値（分）'],
  ['GEMINI_MODEL',                 'gemini-2.5-flash', 'Gemini モデル名（gemini-3-flash 等に変更可）'],
  ['GEMINI_RATE_LIMIT_PER_MIN',    '5',      '同一スタッフあたり 1 分あたりの最大 AI 呼び出し回数']
];

// ─── 初期データ：WorkTypes（作業内容マスタ） ──────────────
// 仕様 v1.0 §3 / §6.3 確定13項目（カテゴリ分類なし・フラット・単一選択）
const DEFAULT_WORK_TYPES = [
  ['wt01', '構想図作成',                 1, true],
  ['wt02', '承認図作成',                 2, true],
  ['wt03', 'バラシ図作成',               3, true],
  ['wt04', '第三者チェック',             4, true],
  ['wt05', '材料取り・仕様検討',         5, true],
  ['wt06', '設計検討',                   6, true],
  ['wt07', '出図準備',                   7, true],
  ['wt08', '出荷後図面修正',             8, true],
  ['wt09', '購入部品手配・在庫部品確認', 9, true],
  ['wt10', '試運転調整・動作確認',      10, true],
  ['wt11', '加工指示',                  11, true],
  ['wt12', '現場対応・現場工事対応',    12, true],
  ['wt13', '出張・打ち合わせ',          13, true]
];

// ─── 初期データ：Staff（スタッフマスタ） ──────────────────
// 技術部スタッフ約6名分の入力枠。
// 各行をシート上で編集（name / email を記入）して使用する。
const STAFF_TEMPLATE_ROWS = [
  ['staff01', '', '', 'staff', 1, true, '', ''],
  ['staff02', '', '', 'staff', 2, true, '', ''],
  ['staff03', '', '', 'staff', 3, true, '', ''],
  ['staff04', '', '', 'staff', 4, true, '', ''],
  ['staff05', '', '', 'staff', 5, true, '', ''],
  ['staff06', '', '', 'admin', 6, true, '', '']
];

// ─── ヘッダ装飾用の色 ────────────────────────────────────
const HEADER_STYLE = {
  background: '#1f2937',
  fontColor:  '#ffffff'
};
