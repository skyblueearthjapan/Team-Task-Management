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
    [SCRIPT_PROP_KEYS.WEBAPP_VERSION]:           '0.1.0',
    [SCRIPT_PROP_KEYS.EXE_API_TOKEN]:            generateRandomToken_()
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

  Object.entries(SHEET_SCHEMA).forEach(([name, headers]) => {
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
