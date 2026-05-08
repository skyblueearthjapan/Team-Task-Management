/**
 * 初期セットアップサービス。
 *
 * GAS エディタで `setupAll()` を実行すると、バインドされたスプレッドシートに
 * 必要なシート・ヘッダ・初期マスタデータ・スクリプトプロパティが一括投入される。
 *
 * 個別関数も用意しているので、再実行時は段階的に呼び出すこともできる:
 *   - setupScriptProperties() : スクリプトプロパティの初期投入
 *   - setupInitialSheets()    : シートとヘッダ行の作成
 *   - setupInitialMasters()   : Settings / WorkTypes / Staff の初期行投入
 *
 * 既存シート・既存データは保護される（上書きしない）。完全リセットしたい場合は
 * resetAllSheets() を使用する（コード内で明示的にガードを外す必要あり）。
 */

function setupAll() {
  setupScriptProperties();
  setupInitialSheets();
  setupInitialMasters();
  seedAdminSettingsIfMissing();
  Logger.log('=== Setup complete ===');
}

/**
 * スクリプトプロパティの初期値を投入する。既存値は保護する。
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperties();

  const seeds = {
    [SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_ID]:    '',
    [SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_NAME]:  '工番マスタ',
    [SCRIPT_PROP_KEYS.POLLING_INTERVAL_SECONDS]: '30',
    [SCRIPT_PROP_KEYS.WEBAPP_VERSION]:           '1.0.0',
    [SCRIPT_PROP_KEYS.EXE_API_TOKEN]:            generateRandomToken_(),
    [SCRIPT_PROP_KEYS.GEMINI_API_KEY]:           ''  // ユーザーが後で手動設定
  };

  let added = 0;
  Object.entries(seeds).forEach(([k, v]) => {
    if (current[k] === undefined || current[k] === '') {
      props.setProperty(k, v);
      added++;
    }
  });
  Logger.log('Script Properties seeded: ' + added + ' key(s)');
}

/**
 * 必要なシートを作成し、ヘッダ行を整備する。
 * 既存シートはそのまま、ヘッダ行が空の場合のみヘッダを書き込む。
 */
function setupInitialSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // MailQueue_Archive は archiveOldMailQueue() 初回実行時に自動生成（仕様 §11）。
  // ここでは作成しない。スキーマは Config.gs に保持済み。
  const SKIP_ON_SETUP = [SHEET_NAMES.MAIL_QUEUE_ARCHIVE];

  Object.entries(SHEET_SCHEMA).forEach(([name, headers]) => {
    if (SKIP_ON_SETUP.includes(name)) return;
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      Logger.log('Created sheet: ' + name);
    } else {
      Logger.log('Sheet already exists: ' + name);
    }

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    const firstRow = headerRange.getValues()[0];
    const isHeaderEmpty = firstRow.every(c => c === '' || c === null);

    if (isHeaderEmpty) {
      headerRange
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground(HEADER_STYLE.background)
        .setFontColor(HEADER_STYLE.fontColor);
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
      Logger.log('  → headers written: ' + headers.length + ' col(s)');
    }

    // 日付列にテキスト書式を適用する（Google Sheets の自動 Date 変換を防ぐ）
    applyDateColumnFormat_(sheet, name);
  });

  // 初回作成時の "シート1" / "Sheet1" を片付ける
  const defaultNames = ['シート1', 'Sheet1'];
  defaultNames.forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) {
      ss.deleteSheet(s);
      Logger.log('Deleted default sheet: ' + n);
    }
  });
}

/**
 * 各マスタ／設定シートに初期行を投入する。
 * 既にデータ（2行目以降）がある場合はスキップする。
 */
function setupInitialMasters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  seedSheetIfEmpty_(ss, SHEET_NAMES.WORK_TYPES, DEFAULT_WORK_TYPES);
  seedSheetIfEmpty_(ss, SHEET_NAMES.STAFF,      STAFF_TEMPLATE_ROWS);
  seedSheetIfEmpty_(ss, SHEET_NAMES.SETTINGS,   DEFAULT_SETTINGS);

  Logger.log('Initial masters seeded.');
}

/**
 * 内部ヘルパ: 指定シートが空（ヘッダのみ）の場合に行を投入する。
 */
function seedSheetIfEmpty_(ss, sheetName, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('seedSheetIfEmpty_: sheet not found: ' + sheetName);
    return;
  }
  if (sheet.getLastRow() > 1) {
    Logger.log('seedSheetIfEmpty_: skipped (data already present): ' + sheetName);
    return;
  }
  if (!rows || rows.length === 0) return;

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('seedSheetIfEmpty_: seeded ' + rows.length + ' row(s) into ' + sheetName);
}

/**
 * 内部ヘルパ: ランダムトークン（EXE 認証用）を生成する。
 */
function generateRandomToken_() {
  const bytes = Utilities.getUuid().replace(/-/g, '') +
                Utilities.getUuid().replace(/-/g, '');
  return bytes.substring(0, 48);
}

/**
 * ⚠️ 破壊的操作: 管理対象シートを全て削除して再構築する。
 * 暴発防止のため、コード内で confirmed=true に書き換えてから実行すること。
 */
function resetAllSheets() {
  const confirmed = false;
  if (!confirmed) {
    throw new Error('resetAllSheets はガードされています。コード内で confirmed=true にしてから実行してください。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEET_NAMES).forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) {
      ss.deleteSheet(s);
      Logger.log('Deleted: ' + name);
    }
  });
  setupAll();
}

// ─── スキーマ整合性 ───────────────────────────────────────────

/**
 * 各シートのヘッダ行が SHEET_SCHEMA の定義と一致しているかを検証する。
 * 修復は行わない（人手判断）。不一致シートの配列を返す。
 *
 * @returns {string[]} 不一致シート名の配列（空なら全シート正常）
 */
function verifySchemaIntegrity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mismatched = [];

  Object.entries(SHEET_SCHEMA).forEach(([name, expectedHeaders]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log('[verifySchemaIntegrity] WARN: sheet not found: ' + name);
      mismatched.push(name);
      return;
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      Logger.log('[verifySchemaIntegrity] WARN: sheet is empty (no columns): ' + name);
      mismatched.push(name);
      return;
    }

    const actualHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (actualHeaders.length < expectedHeaders.length) {
      Logger.log('[verifySchemaIntegrity] WARN: column count mismatch on sheet: ' + name
        + ' | expected: ' + expectedHeaders.length + ' | actual: ' + actualHeaders.length);
      mismatched.push(name);
      return;
    }

    // 期待するヘッダ列と実際の先頭 N 列を比較（N = expectedHeaders.length）
    const isMismatch = expectedHeaders.some((h, i) => actualHeaders[i] !== h);
    if (isMismatch) {
      Logger.log('[verifySchemaIntegrity] WARN: header mismatch on sheet: ' + name
        + ' | expected: ' + JSON.stringify(expectedHeaders)
        + ' | actual: ' + JSON.stringify(actualHeaders.slice(0, expectedHeaders.length)));
      mismatched.push(name);
    }
  });

  if (mismatched.length === 0) {
    Logger.log('[verifySchemaIntegrity] All sheets OK.');
  } else {
    Logger.log('[verifySchemaIntegrity] Mismatched sheets: ' + mismatched.join(', '));
  }

  return mismatched;
}

/**
 * 既存シートのヘッダ列を SHEET_SCHEMA の最新定義に合わせて不足列を末尾に追加する。
 * 列の削除・順序変更は行わない（データ破壊回避）。
 * 既存データはそのまま保持される。
 *
 * 暴発防止のため confirmed=true に書き換えてから実行すること。
 *
 * @note 列追加は末尾に行うため、SHEET_SCHEMA の論理順と物理順が
 *       異なる場合がある（既存データを保持するため）。
 *       理想的な列順への再構成が必要な場合は、setupAll を空シートで
 *       再実行する手順を運用ドキュメントで明示する。
 *       移行後は verifySchemaIntegrity() で WARN が出る場合がある。
 * @returns {Object} {sheetName: addedColumns[]} 形式の結果サマリ
 */
function migrateSchema() {
  const confirmed = false;
  if (!confirmed) {
    throw new Error('migrateSchema はガードされています。コード内で confirmed=true にしてから実行してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  Object.entries(SHEET_SCHEMA).forEach(([name, expectedHeaders]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log('[migrateSchema] sheet not found (skipped): ' + name);
      return;
    }

    const lastCol = sheet.getLastColumn();
    const actualHeaders = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      : [];

    const added = [];
    expectedHeaders.forEach(h => {
      if (!actualHeaders.includes(h) && !added.includes(h)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(h)
          .setFontWeight('bold')
          .setBackground(HEADER_STYLE.background)
          .setFontColor(HEADER_STYLE.fontColor);
        added.push(h);
        Logger.log('[migrateSchema] Added column "' + h + '" to sheet: ' + name);
      }
    });

    if (added.length > 0) {
      result[name] = added;
    }
  });

  Logger.log('[migrateSchema] Migration complete. Changed sheets: '
    + (Object.keys(result).length > 0 ? JSON.stringify(result) : 'none'));
  return result;
}

// ─── MailQueue_Archive 遅延生成 ───────────────────────────────

/**
 * MailQueue_Archive シートが存在しなければ作成し、ヘッダ行を整備する。
 * 既に存在する場合は何もしない（idempotent）。
 * MailQueueService の archiveOldMailQueue() から呼び出される想定。
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Archive シートオブジェクト
 */
function createMailQueueArchiveIfMissing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEET_NAMES.MAIL_QUEUE_ARCHIVE;
  let sheet = ss.getSheetByName(name);

  if (sheet) {
    Logger.log('[createMailQueueArchiveIfMissing] Already exists: ' + name);
    return sheet;
  }

  sheet = ss.insertSheet(name);
  Logger.log('[createMailQueueArchiveIfMissing] Created sheet: ' + name);

  const headers = SHEET_SCHEMA[name];
  if (!headers) {
    throw new Error('[createMailQueueArchiveIfMissing] SHEET_SCHEMA missing for MailQueue_Archive');
  }
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(HEADER_STYLE.background)
    .setFontColor(HEADER_STYLE.fontColor);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log('[createMailQueueArchiveIfMissing] Headers written: ' + headers.length + ' col(s)');

  return sheet;
}

// ─── Settings ヘルパ ─────────────────────────────────────────
// getSetting / setSetting は DataService.gs に集約。
// SetupService 内では DataService の公開関数を直接呼び出す。

/**
 * DEFAULT_SETTINGS に定義されたキーのうち、まだ Settings シートに存在しない行、
 * または value が空のキーを補完する。既存値は保護する（上書きしない）。
 * setupInitialMasters の補強として単体でも呼び出し可能。
 */
function seedAdminSettingsIfMissing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    Logger.log('[seedAdminSettingsIfMissing] Settings sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  // 現在の key→rowIndex マップを構築（ヘッダ行 = 1 なので row 2 から）
  const existingKeys = {};
  if (lastRow >= 2) {
    const keyData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    keyData.forEach((r, i) => {
      if (r[0] !== '') existingKeys[r[0]] = i + 2; // 1-indexed row
    });
  }

  let added = 0;
  DEFAULT_SETTINGS.forEach(([k, v, desc]) => {
    if (existingKeys[k] !== undefined) {
      // 既存行: 一切触らない（ユーザーが意図的にクリアした値を保護）
      Logger.log('[seedAdminSettingsIfMissing] Key already exists (skipped): ' + k);
    } else {
      // 存在しない行: 末尾に追加
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1, 1, 3).setValues([[k, v, desc]]);
      Logger.log('[seedAdminSettingsIfMissing] Added missing key: ' + k);
      added++;
    }
  });

  Logger.log('[seedAdminSettingsIfMissing] Done. ' + added + ' key(s) added/filled.');
}

// ─── 日付列テキスト書式 ──────────────────────────────────────

/**
 * 指定シートの日付列に '@STRING@'（テキスト）書式を適用する。
 * Google Sheets が日付文字列を Date 型として自動変換する問題を防ぐ。
 * setupInitialSheets() のループ内から呼び出される。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     - 対象シートオブジェクト
 * @param {string}                             sheetName - SHEET_NAMES に対応したシート名
 */
function applyDateColumnFormat_(sheet, sheetName) {
  var dateColumnsMap = {
    'Schedules':    ['startDate', 'endDate'],
    'DailyReports': ['reportDate', 'periodStart', 'periodEnd'],
    'MailQueue':    ['reportDate']
  };
  var cols = dateColumnsMap[sheetName];
  if (!cols) return;
  var headers = SHEET_SCHEMA[sheetName];
  if (!headers) return;
  cols.forEach(function(colName) {
    var idx = headers.indexOf(colName);
    if (idx === -1) return;
    // 1-indexed の列番号で最大 10000 行にテキスト書式を適用（ヘッダ行を除く 2 行目から）
    sheet.getRange(2, idx + 1, 10000, 1).setNumberFormat('@STRING@');
    Logger.log('  → setNumberFormat(@STRING@) on col "' + colName + '" in ' + sheetName);
  });
}

/**
 * 既存の Schedules / DailyReports / MailQueue シートで、
 * Date 型として保存されている日付セルを YYYY-MM-DD のテキストに変換し、
 * 列書式をテキストに設定する。
 *
 * 既存データがある状態で日付フィルタが効かないバグの後方移行用。
 * GAS エディタで confirmed = true に書き換えてから実行すること。
 */
function fixDateColumnsToText() {
  var confirmed = false;
  if (!confirmed) {
    throw new Error('fixDateColumnsToText はガード中です。confirmed=true にしてから実行してください。');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateColumnsMap = {
    'Schedules':    ['startDate', 'endDate'],
    'DailyReports': ['reportDate', 'periodStart', 'periodEnd'],
    'MailQueue':    ['reportDate']
  };

  Object.keys(dateColumnsMap).forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    var headers = SHEET_SCHEMA[sheetName];
    if (!headers) return;
    var cols = dateColumnsMap[sheetName];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    cols.forEach(function(colName) {
      var idx = headers.indexOf(colName);
      if (idx === -1) return;
      var range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
      var values = range.getValues();
      var converted = values.map(function(row) {
        var v = row[0];
        if (v instanceof Date) {
          return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
        }
        return [v];
      });
      // 書式をテキストに設定 → 値を再書き込み
      range.setNumberFormat('@STRING@');
      range.setValues(converted);
      Logger.log('[fixDateColumnsToText] Converted col "' + colName + '" in ' + sheetName);
    });
  });
  Logger.log('fixDateColumnsToText complete');
}
