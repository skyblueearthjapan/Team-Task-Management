/**
 * グローバル設定定数。
 * シート名・列スキーマ・スクリプトプロパティのキー・初期データを一元管理する。
 * SetupService をはじめ各サービスはここを単一の正本として参照する。
 */

// ─── シート名 ────────────────────────────────────────────────
const SHEET_NAMES = {
  SCHEDULES: 'Schedules',
  DAILY_REPORTS: 'DailyReports',
  MAIL_QUEUE: 'MailQueue',
  MAIL_LOG: 'MailLog',
  STAFF: 'Staff',
  WORK_TYPES: 'WorkTypes',
  HOLIDAYS: 'Holidays',
  SETTINGS: 'Settings'
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
    'linkedScheduleId',
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
    'toAddresses',
    'ccAddresses',
    'subjectVars',       // JSON
    'bodyVars',          // JSON
    'status',            // pending / picked / drafted / sent / failed
    'pickedBy',
    'pickedAt',
    'processedAt',
    'errorMessage',
    'createdAt'
  ],
  [SHEET_NAMES.MAIL_LOG]: [
    'id',
    'mailQueueId',
    'timestamp',
    'level',             // info / warn / error
    'event',             // picked / drafted / sent / failed / retry / mismatch
    'message'
  ],
  [SHEET_NAMES.STAFF]: [
    'id',
    'name',
    'email',
    'role',              // staff / admin
    'displayOrder',
    'active'             // TRUE/FALSE
  ],
  [SHEET_NAMES.WORK_TYPES]: [
    'id',
    'name',
    'category',
    'displayOrder',
    'active'
  ],
  [SHEET_NAMES.HOLIDAYS]: [
    'date',
    'name'
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
  EXE_API_TOKEN: 'EXE_API_TOKEN'
};

// ─── 初期データ：Settings ─────────────────────────────────
const DEFAULT_SETTINGS = [
  ['KOBAN_MASTER_SHEET_ID', '', '工番マスタが格納されている別ブックのスプレッドシートID（後で記入）'],
  ['KOBAN_MASTER_SHEET_NAME', '工番マスタ', '工番マスタブック内のシート名'],
  ['POLLING_INTERVAL_SECONDS', '30', 'ローカル EXE が MailQueue をポーリングする間隔（秒）'],
  ['MAIL_DEFAULT_MODE', 'draft', 'メール送信モードの既定値（draft / send）'],
  ['WEBAPP_VERSION', '0.1.0', 'WebApp バージョン'],
  ['LAST_HEARTBEAT_TIMESTAMP', '', 'ローカル EXE からの最終ハートビート（自動更新）'],
  ['LAST_HEARTBEAT_HOSTNAME', '', '最終ハートビートを送ってきた PC 名（自動更新）']
];

// ─── 初期データ：WorkTypes（作業内容マスタ） ──────────────
// 部内で運用する想定の代表的な作業区分。後でシート上で編集可。
const DEFAULT_WORK_TYPES = [
  ['wt01', '図面作成',    '設計',    1,  true],
  ['wt02', '検図',        '設計',    2,  true],
  ['wt03', '出図',        '設計',    3,  true],
  ['wt04', '修正対応',    '設計',    4,  true],
  ['wt05', '打合せ',      '会議',    5,  true],
  ['wt06', '立会',        '現場',    6,  true],
  ['wt07', '報告書作成',  '事務',    7,  true],
  ['wt08', '見積対応',    '事務',    8,  true],
  ['wt09', '一般作業',    'その他',  9,  true],
  ['wt10', '不在/休暇',   'その他', 10,  true]
];

// ─── 初期データ：Staff（スタッフマスタ） ──────────────────
// 機械設計技術部スタッフ約6名分の入力枠。
// 各行をシート上で編集（name / email を記入）して使用する。
const STAFF_TEMPLATE_ROWS = [
  ['staff01', '', '', 'staff', 1, true],
  ['staff02', '', '', 'staff', 2, true],
  ['staff03', '', '', 'staff', 3, true],
  ['staff04', '', '', 'staff', 4, true],
  ['staff05', '', '', 'staff', 5, true],
  ['staff06', '', '', 'admin', 6, true]
];

// ─── ヘッダ装飾用の色 ────────────────────────────────────
const HEADER_STYLE = {
  background: '#1f2937',
  fontColor:  '#ffffff'
};
