/**
 * DataService.gs
 *
 * スプレッドシートへの汎用 CRUD 操作を提供する。
 * UI・MailQueueService・MasterService から呼び出される唯一のデータアクセス層。
 *
 * 設計方針:
 *   - SHEET_SCHEMA（Config.gs）を参照して列順を決定し、列名のハードコードを排除する。
 *   - ヘッダ行は各シートの 1 行目固定とする。
 *   - appendRow / updateRow 内で createdAt / updatedAt 列が存在する場合は自動セットする（仕様 §9 監査要件準拠）。
 *   - 不正シート名は例外、行が見つからない更新/削除は false を返す。
 *   - LockService 等の排他制御は MailQueueService（STEP 4）で扱うため、本サービスでは使用しない。
 */

// ════════════════════════════════════════════════════════════════
// A. 汎用 CRUD（公開 API）
// ════════════════════════════════════════════════════════════════

/**
 * 指定シートの全行を Object 配列で返す。
 * ヘッダ行（1 行目）をキーに変換し、データ行を [{列名: 値, ...}, ...] の形式で返す。
 *
 * @param {string} sheetName - SHEET_NAMES に定義されたシート名
 * @returns {Object[]} 全行の Object 配列（空シートは空配列）
 * @throws {Error} シートが存在しない場合
 */
function listAll(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('DataService.listAll: Sheet not found: ' + sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => rowToObject_(headers, row));
}

/**
 * id（または指定列）でレコードを1件取得する。
 * 内部で listAll を経由しているため小〜中規模シート（〜数千行）向け。
 * 大規模シート対象の場合は findRowIndex_ + 単行取得への置き換えを検討。
 *
 * @param {string} sheetName  - SHEET_NAMES に定義されたシート名
 * @param {*}      id         - 検索する値
 * @param {string} [idColumn] - 検索対象の列名（既定: 'id'）
 * @returns {Object|null} 一致した行の Object、見つからなければ null
 * @throws {Error} シートが存在しない場合
 */
function getById(sheetName, id, idColumn) {
  const col = idColumn || 'id';
  const rows = listAll(sheetName);
  return rows.find(r => r[col] === id) || null;
}

/**
 * 新規行をシートに追加する。
 * - SHEET_SCHEMA の列順に従って値を展開する。
 * - createdAt 列が存在しかつ record に未設定の場合は nowIso_() を自動セットする。
 * - updatedAt 列が存在しかつ record に未設定の場合は createdAt と同値をセットする。
 *
 * @param {string} sheetName - SHEET_NAMES に定義されたシート名
 * @param {Object} record    - 書き込む内容のオブジェクト（id は呼び出し元で設定済みであること）
 * @returns {string} record.id
 * @throws {Error} シートが存在しない場合
 */
function appendRow(sheetName, record) {
  return appendRow_(sheetName, record);
}

/**
 * id が一致する行を部分更新する。
 * updates オブジェクトに含まれるキーのみを書き換え、他の列は変更しない。
 * - updatedAt 列が存在する場合は nowIso_() で自動上書きする。
 *
 * @param {string} sheetName  - SHEET_NAMES に定義されたシート名
 * @param {*}      id         - 更新対象行の識別値
 * @param {Object} updates    - 更新内容（キー: 列名、値: 新しい値）
 * @param {string} [idColumn] - 識別列名（既定: 'id'）
 * @returns {boolean} 更新成功: true、行が見つからない: false
 * @throws {Error} シートが存在しない場合
 */
function updateRow(sheetName, id, updates, idColumn) {
  return updateRow_(sheetName, id, updates, idColumn);
}

/**
 * id が一致する行を削除する。
 *
 * @param {string} sheetName  - SHEET_NAMES に定義されたシート名
 * @param {*}      id         - 削除対象行の識別値
 * @param {string} [idColumn] - 識別列名（既定: 'id'）
 * @returns {boolean} 削除成功: true、行が見つからない: false
 * @throws {Error} シートが存在しない場合
 */
function deleteRow(sheetName, id, idColumn) {
  const col = idColumn || 'id';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('DataService.deleteRow: Sheet not found: ' + sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = headers.indexOf(col);
  if (colIdx === -1) throw new Error('DataService.deleteRow: Column not found: ' + col);
  // 下から順に削除して行番号ズレを防ぐ
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][colIdx] === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/**
 * predicate 関数に一致する行を配列で返す。
 *
 * @param {string}   sheetName - SHEET_NAMES に定義されたシート名
 * @param {Function} predicate - (row: Object) => boolean
 * @returns {Object[]} 条件に一致した行の Object 配列
 * @throws {Error} シートが存在しない場合
 */
function listWhere(sheetName, predicate) {
  return listAll(sheetName).filter(predicate);
}

// ════════════════════════════════════════════════════════════════
// B. 内部ヘルパ（末尾アンダースコア）
// ════════════════════════════════════════════════════════════════

/**
 * 指定シートの全データを 2D 配列で返す（ヘッダ行を含む）。
 * シートが空（データなし）の場合は空配列を返す。
 *
 * @param {string} sheetName
 * @returns {Array[]} 2D 配列（data[0] がヘッダ行）
 * @throws {Error} シートが存在しない場合
 */
function getSheetData_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('DataService.getSheetData_: Sheet not found: ' + sheetName);
  if (sheet.getLastRow() === 0) return [];
  return sheet.getDataRange().getValues();
}

/**
 * ヘッダ配列と行配列から Object を生成する。
 *
 * @param {string[]} headers - ヘッダ（列名）配列
 * @param {Array}    row     - 値の配列（headers と同じ長さ）
 * @returns {Object} {列名: 値, ...}
 */
function rowToObject_(headers, row) {
  return Object.fromEntries(headers.map((h, i) => [h, row[i]]));
}

/**
 * SHEET_SCHEMA の列順に従って Object を値配列に変換する。
 * スキーマに存在しないキーは無視する。スキーマにあるがオブジェクトに未設定のキーは '' とする。
 *
 * @param {string[]} headers - SHEET_SCHEMA から取得した列名配列（objectToRow_ は headers を直接受け取る設計）
 * @param {Object}   obj     - 書き込む内容のオブジェクト
 * @returns {Array} headers 順に並んだ値配列
 */
function objectToRow_(headers, obj) {
  // 注: GAS スプレッドシート上では null を直接書き込めないため空文字に変換する。
  // 仕様 §6.3 で「string | null」と定義された列（kobanCode 等）は読み取り時に空文字となるが、
  // ビジネスロジックでは「空文字 = null 相当」として扱うこと。
  return headers.map(h => (obj[h] !== undefined && obj[h] !== null) ? obj[h] : '');
}

/**
 * idColumn の値が idValue に一致する行番号（1-indexed）を返す。
 * 見つからない場合は -1 を返す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet   - 検索対象のシートオブジェクト
 * @param {string}                             idColumn - 検索対象の列名
 * @param {*}                                  idValue  - 検索する値
 * @returns {number} 1-indexed 行番号（ヘッダ込み）、見つからない場合は -1
 */
function findRowIndex_(sheet, idColumn, idValue) {
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return -1;
  const headers = data[0];
  const colIdx = headers.indexOf(idColumn);
  if (colIdx === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIdx] === idValue) return i + 1; // 1-indexed
  }
  return -1;
}

/**
 * 新規行を追加する内部ヘルパ。
 * SHEET_SCHEMA の列順に従い objectToRow_ で展開して appendRow する。
 * createdAt / updatedAt の自動セットもここで行う。
 *
 * @param {string} sheetName
 * @param {Object} record
 * @returns {string} record.id
 */
function appendRow_(sheetName, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('DataService.appendRow_: Sheet not found: ' + sheetName);

  // SHEET_SCHEMA から列順を取得する
  const schema = SHEET_SCHEMA[sheetName];
  if (!schema) throw new Error('DataService.appendRow_: No schema defined for: ' + sheetName);

  // 監査列の自動セット（createdAt / updatedAt）
  const now = nowIso_();
  if (schema.indexOf('createdAt') !== -1 && !record.createdAt) {
    record.createdAt = now;
  }
  if (schema.indexOf('updatedAt') !== -1 && !record.updatedAt) {
    record.updatedAt = record.createdAt || now;
  }

  const row = objectToRow_(schema, record);
  sheet.appendRow(row);
  return record.id;
}

/**
 * 行を部分更新する内部ヘルパ。
 * updatedAt 列が存在する場合は自動で nowIso_() をセットする。
 *
 * @param {string} sheetName
 * @param {*}      id
 * @param {Object} updates
 * @param {string} [idColumn]
 * @returns {boolean}
 */
function updateRow_(sheetName, id, updates, idColumn) {
  const col = idColumn || 'id';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('DataService.updateRow_: Sheet not found: ' + sheetName);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  const headers = data[0];
  const colIdx = headers.indexOf(col);
  if (colIdx === -1) throw new Error('DataService.updateRow_: Column not found: ' + col);

  for (let i = 1; i < data.length; i++) {
    if (data[i][colIdx] === id) {
      const rowData = data[i].slice();
      headers.forEach((h, j) => {
        if (updates[h] !== undefined) rowData[j] = updates[h];
      });
      // updatedAt 自動セット（updates に明示指定がない場合のみ）
      if (headers.indexOf('updatedAt') !== -1 && updates.updatedAt === undefined) {
        const updatedAtIdx = headers.indexOf('updatedAt');
        rowData[updatedAtIdx] = nowIso_();
      }
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);
      return true;
    }
  }
  return false;
}

// ════════════════════════════════════════════════════════════════
// C. ID 生成 / タイムスタンプ
// ════════════════════════════════════════════════════════════════

/**
 * プレフィックス付きの短縮 UUID を生成する。
 * 例: generateId_('sched') → 'sched_a1b2c3d4'
 *
 * @param {string} prefix - ID の先頭に付与するプレフィックス
 * @returns {string} prefix + '_' + UUID 短縮形（16文字）
 */
function generateId_(prefix) {
  const uuid = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  return prefix + '_' + uuid;
}

/**
 * 現在日時を ISO 8601 形式の文字列で返す。
 *
 * @returns {string} 例: "2026-05-09T10:30:00.000Z"
 */
function nowIso_() {
  return new Date().toISOString();
}

// ════════════════════════════════════════════════════════════════
// D. ドメイン別ラッパー（後続 STEP から利用）
// ════════════════════════════════════════════════════════════════

// ── Schedules ────────────────────────────────────────────────

/**
 * 指定期間内（inclusive）の Schedule レコードを返す。
 *
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate   - YYYY-MM-DD
 * @returns {Object[]}
 */
function getSchedules(startDate, endDate) {
  return listAll(SHEET_NAMES.SCHEDULES).filter(r => {
    const rEnd = String(r.endDate || '');
    const rStart = String(r.startDate || '');
    return rEnd >= startDate && rStart <= endDate;
  });
}

/**
 * 新規 Schedule レコードを追加する。
 * id / createdAt / updatedAt は自動生成。
 *
 * @param {Object} record - staffId, startDate, endDate, kobanCode, workTypeId, note, lane
 * @returns {string} 生成された id
 */
function createSchedule(record) {
  record.id = generateId_('sch');
  return appendRow_(SHEET_NAMES.SCHEDULES, record);
}

/**
 * Schedule レコードを部分更新する。updatedAt は自動セット。
 *
 * @param {string} id      - 更新対象の id
 * @param {Object} updates - 更新内容
 * @returns {boolean}
 */
function updateSchedule(id, updates) {
  return updateRow_(SHEET_NAMES.SCHEDULES, id, updates);
}

/**
 * Schedule レコードを削除する。
 *
 * @param {string} id
 * @returns {boolean}
 */
function deleteSchedule(id) {
  return deleteRow(SHEET_NAMES.SCHEDULES, id);
}

// ── DailyReports ─────────────────────────────────────────────

/**
 * 指定スタッフ・日付の DailyReport 行を返す。
 *
 * @param {string} staffId    - Staff.id
 * @param {string} reportDate - YYYY-MM-DD
 * @returns {Object[]}
 */
function getDailyReports(staffId, reportDate) {
  return listAll(SHEET_NAMES.DAILY_REPORTS).filter(r =>
    String(r.staffId) === String(staffId) &&
    String(r.reportDate) === String(reportDate)
  );
}

/**
 * DailyReport を追加または更新する（staffId + reportDate + section + seq で一意識別）。
 * 既存行があれば更新、なければ追加。
 *
 * @param {Object} record - staffId, reportDate, section, seq, ... を含む Object
 * @returns {string|boolean} 新規追加: 生成 id / 更新: true
 */
function upsertDailyReport(record) {
  // seq は数値・文字列どちらで渡されても一致するよう文字列比較する
  const seqStr = String(record.seq);
  const existing = getDailyReports(record.staffId, record.reportDate)
    .find(r => String(r.seq) === seqStr && r.section === record.section);
  if (existing) {
    // 検索キー（staffId/reportDate/section/seq）も updates に含まれるが値不変のため問題なし
    // id だけは既存行の id を保護するため除外
    const updates = {};
    Object.keys(record).forEach(function(k) {
      if (k !== 'id') updates[k] = record[k];
    });
    return updateRow_(SHEET_NAMES.DAILY_REPORTS, existing.id, updates);
  } else {
    record.id = generateId_('dr');
    return appendRow_(SHEET_NAMES.DAILY_REPORTS, record);
  }
}

/**
 * 指定日付の全スタッフ分 DailyReport 行を返す。
 * UI が1日の全スタッフ日報を一括取得する際に使用する（スタッフ数分の個別呼び出しを1回に削減）。
 *
 * @param {string} reportDate - YYYY-MM-DD
 * @returns {Object[]}
 * @note 内部で listAll を使用。DailyReports が数千行規模になった場合は最適化要検討。
 */
function getDailyReportsByDate(reportDate) {
  return listAll(SHEET_NAMES.DAILY_REPORTS).filter(r =>
    String(r.reportDate) === String(reportDate)
  );
}

// ── Staff ─────────────────────────────────────────────────────

/**
 * active なスタッフ一覧を displayOrder 昇順で返す。
 *
 * @returns {Object[]}
 */
function getActiveStaff() {
  return listAll(SHEET_NAMES.STAFF)
    .filter(r => r.active === true || r.active === 'TRUE')
    .sort((a, b) => Number(a.displayOrder) - Number(b.displayOrder));
}

// ── WorkTypes ─────────────────────────────────────────────────

/**
 * active な WorkType 一覧を displayOrder 昇順で返す。
 *
 * @returns {Object[]}
 */
function getWorkTypes() {
  return listAll(SHEET_NAMES.WORK_TYPES)
    .filter(r => r.active === true || r.active === 'TRUE')
    .sort((a, b) => Number(a.displayOrder) - Number(b.displayOrder));
}

// ── Settings ──────────────────────────────────────────────────

/**
 * Settings シートから指定キーの value を返す。
 *
 * @param {string} key - Settings.key
 * @returns {string|null} 値、またはキーが存在しなければ null
 */
function getSetting(key) {
  const row = listAll(SHEET_NAMES.SETTINGS).find(r => r.key === key);
  return row ? row.value : null;
}

/**
 * Settings シートの指定キーの value を更新する。
 * Settings シートは id 列を持たないため idColumn に 'key' を指定する。
 *
 * @param {string} key   - Settings.key
 * @param {string} value - 新しい値
 * @returns {boolean}
 */
function setSetting(key, value) {
  return updateRow_(SHEET_NAMES.SETTINGS, key, { value: value }, 'key');
}
