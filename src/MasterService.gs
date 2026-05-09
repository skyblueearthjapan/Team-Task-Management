/**
 * MasterService.gs
 *
 * 外部マスタ（工番マスタ・社内カレンダーマスタ）からデータを取得・キャッシュ・配信する。
 *
 * 設計方針:
 *   - 外部マスタは読み取り専用。本サービスから書き込み禁止。
 *   - マスタブックの ID は Settings シートから取得し、Script Properties をフォールバックとする。
 *   - CacheService（Script キャッシュ）を使い、TTL 6 時間でデータを保持する。
 *   - CacheService の 1 エントリ上限（約 100KB）を超える場合は分割保存する。
 *   - キャッシュキーにバージョン suffix (:v1) を付与し、スキーマ変更時の互換切れに備える。
 */

// ─── キャッシュ設定 ─────────────────────────────────────────────
/** @const {string} */ var CACHE_KEY_KOBAN    = 'master:koban:v1';
/** @const {string} */ var CACHE_KEY_CALENDAR = 'master:calendar:v1';
/** @const {number} */ var CACHE_TTL_SECONDS  = 6 * 60 * 60; // 6 時間
/** @const {number} */ var CACHE_CHUNK_BYTES  = 90 * 1024;   // 90 KB（100 KB 上限に余裕を持たせる）

// 社内カレンダーマスタのシート名は実データ通り固定
/** @const {string} */ var CALENDAR_SHEET_NAME = '社内カレンダーマスタ';

// ════════════════════════════════════════════════════════════════
// A. 公開関数
// ════════════════════════════════════════════════════════════════

/**
 * 工番マスタ全件を取得する。
 * キャッシュが有効な場合はキャッシュから返す。MISS 時は外部ブックから取得してキャッシュに書き戻す。
 *
 * 戻り値の各オブジェクトキーは実データの日本語列名をそのまま使用:
 *   { 工番, 受注先, 納入先, 納入先住所, 品名, 数量, 取込日時 }
 *
 * @returns {Object[]} 工番マスタ全件の Object 配列
 * @throws {Error} マスタブックを開けない場合、またはシートが見つからない場合
 */
function getKobanMaster() {
  var cached = cacheGet_(CACHE_KEY_KOBAN);
  if (cached !== null) {
    return cached;
  }

  var ss = openMasterSpreadsheet_();
  var kobanSheetName = getSetting(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_NAME)
    || PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_NAME)
    || '工番マスタ';
  var sheet = ss.getSheetByName(kobanSheetName);
  if (!sheet) {
    throw new Error('MasterService.getKobanMaster: Sheet not found: ' + kobanSheetName);
  }

  var rows = parseRows_(sheet);
  cachePut_(CACHE_KEY_KOBAN, rows, CACHE_TTL_SECONDS);
  return rows;
}

/**
 * 社内カレンダーマスタから期間内のレコードを取得する。
 * startDate・endDate を省略した場合は全件を返す。
 * キャッシュが有効な場合はキャッシュから返す。
 *
 * 戻り値の各オブジェクトキーは実データの日本語列名をそのまま使用:
 *   { 日付, 区分, 曜日, 備考 }
 *
 * 区分の値: '休日' / '出勤土曜' / '祝日'（日本語のまま）
 *
 * @param {string} [startDate] - YYYY-MM-DD 形式の開始日（省略可）
 * @param {string} [endDate]   - YYYY-MM-DD 形式の終了日（省略可）
 * @returns {Object[]} 社内カレンダーレコードの Object 配列
 * @throws {Error} マスタブックを開けない場合、またはシートが見つからない場合
 */
function getCompanyCalendar(startDate, endDate) {
  var allRecords = getCalendarAll_();

  if (!startDate && !endDate) {
    return allRecords;
  }

  return allRecords.filter(function(r) {
    var rawDate = r['日付'];
    if (!rawDate) return false;

    // Date オブジェクトの場合は YYYY-MM-DD 文字列に変換
    var d;
    if (rawDate instanceof Date) {
      var y = rawDate.getFullYear();
      var mo = ('0' + (rawDate.getMonth() + 1)).slice(-2);
      var day = ('0' + rawDate.getDate()).slice(-2);
      d = y + '-' + mo + '-' + day;
    } else {
      d = String(rawDate);
    }

    if (!d) return false;
    if (startDate && d < startDate) return false;
    if (endDate   && d > endDate)   return false;
    return true;
  });
}

/**
 * 指定日が営業日か判定する。
 *
 * 判定優先順位:
 *   1. 社内カレンダーマスタに登録済みの日付
 *      - 区分が '出勤土曜' → true（営業日）
 *      - 区分が '休日' または '祝日' → false（非営業日）
 *   2. マスタ未登録の土曜（getDay() === 6）→ false
 *   3. マスタ未登録の日曜（getDay() === 0）→ false
 *   4. マスタ未登録の平日 → true
 *
 * @param {string} dateStr - YYYY-MM-DD 形式の日付文字列
 * @returns {boolean} 営業日なら true、それ以外は false
 */
function isWorkingDay(dateStr) {
  var calendarMap = buildCalendarMap_();

  if (calendarMap[dateStr] !== undefined) {
    var category = calendarMap[dateStr];
    if (category === '出勤土曜') return true;
    // '休日' / '祝日' はすべて非営業日
    return false;
  }

  // マスタ未登録の場合は曜日で判定
  // タイムゾーン非依存で日付パース（YYYY-MM-DD 形式前提）
  var parts = String(dateStr).split('-');
  if (parts.length < 3) return false;
  var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  var dow = d.getDay(); // 0=日, 6=土
  if (dow === 0 || dow === 6) return false;
  return true;
}

/**
 * referenceDate の前日からさかのぼって最初の稼働日を返す。
 *
 * 稼働日の判定は isWorkingDay() に委譲する:
 *   - 社内カレンダーで区分='出勤土曜' → 稼働日（土曜でも）
 *   - 社内カレンダーで区分='休日' または '祝日' → 非稼働日
 *   - 未登録の土日 → 非稼働日
 *   - 未登録の平日 → 稼働日
 *
 * 例:
 *   2026-04-30(木) → 2026-04-28(火)  ※ 4/29 が祝日の場合
 *   2026-05-11(月) → 2026-05-08(金)  ※ 5/9,5/10 が非稼働日の場合
 *   2026-05-11(月) → 2026-05-09(土)  ※ 5/9 が出勤土曜の場合
 *
 * @param {string} referenceDate - YYYY-MM-DD 形式の基準日
 * @returns {string} 最初の稼働日（YYYY-MM-DD）
 * @throws {Error} 過去に稼働日が見つからない場合（365 日さかのぼっても該当なし）
 */
function getPreviousWorkingDay(referenceDate) {
  var parts = String(referenceDate).split('-');
  if (parts.length < 3) {
    throw new Error('getPreviousWorkingDay: Invalid date format: ' + referenceDate);
  }

  // 365 日分の検索でも buildCalendarMap_ を毎回呼ばないよう、冒頭で 1 回取得
  var calMap = buildCalendarMap_();

  var cur = new Date(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10) - 1,
    parseInt(parts[2], 10)
  );

  for (var i = 0; i < 365; i++) {
    cur.setDate(cur.getDate() - 1);
    var ymd = Utilities.formatDate(cur, 'JST', 'yyyy-MM-dd');
    var dow = cur.getDay(); // 0=日, 6=土
    var category = calMap[ymd];

    // 出勤土曜は稼働日
    if (category === '出勤土曜') return ymd;
    // 休日 / 祝日 は非稼働
    if (category === '休日' || category === '祝日') continue;
    // 未登録の土日は非稼働
    if (dow === 0 || dow === 6) continue;
    // それ以外（平日、未登録）は稼働日
    return ymd;
  }

  throw new Error('getPreviousWorkingDay: 365 日以内に稼働日が見つかりません: ' + referenceDate);
}

/**
 * キャッシュを強制的に削除する（管理者用）。
 * 次回 getKobanMaster / getCompanyCalendar 呼び出し時に外部ブックから再取得する。
 */
function refreshMasterCache() {
  var cache = CacheService.getScriptCache();

  // 工番マスタキャッシュ削除（分割キーも含む）
  removeChunkedCache_(cache, CACHE_KEY_KOBAN);

  // 社内カレンダーキャッシュ削除（分割キーも含む）
  removeChunkedCache_(cache, CACHE_KEY_CALENDAR);

  Logger.log('MasterService.refreshMasterCache: Cache cleared.');
}

/**
 * 起動時にフロントエンドへ配信する単一エンドポイント。
 *
 * @returns {{ kobanList: Object[], calendar: Object[], calendarMap: Object }}
 *   kobanList  : 工番マスタ全件（フロントのプルダウン用）
 *   calendar   : 社内カレンダー全件
 *   calendarMap: { "YYYY-MM-DD": "区分文字列", ... }（フロントで O(1) ルックアップ用）
 */
function getMasterPayload() {
  var kobanList   = getKobanMaster();
  var calendar    = getCalendarAll_();
  var calendarMap = buildCalendarMapFromRecords_(calendar);

  return {
    kobanList   : kobanList,
    calendar    : calendar,
    calendarMap : calendarMap
  };
}

// ════════════════════════════════════════════════════════════════
// B. 内部ヘルパ（末尾アンダースコア）
// ════════════════════════════════════════════════════════════════

/**
 * 外部マスタのスプレッドシートを openById で開いて返す。
 * ID は以下の優先順で取得する:
 *   1. Script Properties の KOBAN_MASTER_SHEET_ID（高速経路: スプレッドシート読み取り不要）
 *   2. Settings シートの KOBAN_MASTER_SHEET_ID（フォールバック: 管理者が Settings で運用変更している場合）
 *
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 * @throws {Error} ID が未設定、またはブックを開けない場合
 */
function openMasterSpreadsheet_() {
  // 高速経路: Script Properties から直接取得（Settings シート読み取りコストを回避）
  var sheetId = PropertiesService.getScriptProperties()
    .getProperty(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_ID);

  // フォールバック: Settings シート（管理者がそこで運用変更している場合）
  if (!sheetId) {
    sheetId = getSetting(SCRIPT_PROP_KEYS.KOBAN_MASTER_SHEET_ID);
  }

  if (!sheetId) {
    throw new Error(
      'MasterService.openMasterSpreadsheet_: ' +
      'KOBAN_MASTER_SHEET_ID が Script Properties および Settings シートに設定されていません。'
    );
  }

  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (e) {
    throw new Error(
      'MasterService.openMasterSpreadsheet_: ' +
      'スプレッドシートを開けませんでした (ID: ' + sheetId + '). ' + e.message
    );
  }
}

/**
 * シートのヘッダ行（1 行目）をキーとして全データ行を Object 配列で返す。
 * ヘッダ行がない、またはデータ行が 0 件の場合は空配列を返す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {Object[]} [{列名: 値, ...}, ...] の形式
 */
function parseRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];

  var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // 全列が空の行はスキップ
    var isEmpty = row.every(function(cell) { return cell === '' || cell === null || cell === undefined; });
    if (isEmpty) continue;

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      if (key === '' || key === null || key === undefined) continue;
      obj[String(key)] = row[j];
    }
    result.push(obj);
  }
  return result;
}

/**
 * CacheService からキャッシュ値を取得する（分割保存に対応）。
 * 存在しない場合は null を返す。
 *
 * @param {string} key - キャッシュキー
 * @returns {Object[]|null} デシリアライズ済みオブジェクト配列、またはキャッシュ MISS 時は null
 */
function cacheGet_(key) {
  var cache = CacheService.getScriptCache();

  // まず分割数を確認する
  var countStr = cache.get(key + ':count');
  if (countStr === null) {
    // 分割なしの単一エントリを試みる
    var single = cache.get(key);
    if (single === null) return null;
    try {
      return JSON.parse(single);
    } catch (e) {
      Logger.log('MasterService.cacheGet_: JSON parse error for key=' + key + ': ' + e.message);
      return null;
    }
  }

  // 分割保存されている場合は結合して返す
  var count = parseInt(countStr, 10);
  var chunks = [];
  for (var i = 0; i < count; i++) {
    var chunk = cache.get(key + ':' + i);
    if (chunk === null) {
      // 一部が欠損しているためキャッシュ MISS 扱い
      Logger.log('MasterService.cacheGet_: Missing chunk ' + i + ' for key=' + key);
      return null;
    }
    chunks.push(chunk);
  }

  try {
    return JSON.parse(chunks.join(''));
  } catch (e) {
    Logger.log('MasterService.cacheGet_: JSON parse error (chunked) for key=' + key + ': ' + e.message);
    return null;
  }
}

/**
 * CacheService へ値を保存する（100 KB 超は分割保存）。
 * 保存できない場合は警告ログを出力して続行する（キャッシュ失敗はフォールバック可能）。
 *
 * @param {string}   key   - キャッシュキー
 * @param {Object[]} value - 保存するオブジェクト配列
 * @param {number}   ttl   - TTL（秒）
 */
function cachePut_(key, value, ttl) {
  var cache  = CacheService.getScriptCache();
  var json   = JSON.stringify(value);
  var byteLen = Utilities.newBlob(json).getBytes().length;

  if (byteLen <= CACHE_CHUNK_BYTES) {
    // 分割不要：単一エントリとして保存
    try {
      cache.put(key, json, ttl);
    } catch (e) {
      Logger.log('MasterService.cachePut_: Failed to put cache for key=' + key + ': ' + e.message);
    }
    return;
  }

  // 100 KB 超：文字列を CACHE_CHUNK_BYTES 相当の文字数で分割して保存
  // ※ UTF-8 マルチバイト文字を考慮し、文字数ではなくバイト換算で分割する
  Logger.log(
    'MasterService.cachePut_: Data size (' + (byteLen / 1024).toFixed(1) +
    ' KB) exceeds threshold. Storing in chunks for key=' + key
  );

  // 文字列を目安の文字数ごとに分割
  // 日本語+ASCII 混在で 1 文字最大 3 バイトだが、ASCII 比率が高い場合
  // チャンクが上限を超えるリスクがあるため /4 で安全マージンを確保
  var chunkCharSize = Math.floor(CACHE_CHUNK_BYTES / 4);
  var chunks = [];
  for (var start = 0; start < json.length; start += chunkCharSize) {
    chunks.push(json.substring(start, start + chunkCharSize));
  }

  // バイト検証と保存
  var entries = {};
  entries[key + ':count'] = String(chunks.length);
  for (var i = 0; i < chunks.length; i++) {
    entries[key + ':' + i] = chunks[i];
  }

  try {
    // putAll で一括保存（TTL は共通）
    cache.putAll(entries, ttl);
  } catch (e) {
    Logger.log('MasterService.cachePut_: Failed to put chunked cache for key=' + key + ': ' + e.message);
  }
}

/**
 * 分割保存されているキャッシュエントリ（チャンク）をすべて削除する。
 *
 * @param {GoogleAppsScript.Cache.Cache} cache - CacheService.getScriptCache() の戻り値
 * @param {string}                       key   - 基底キャッシュキー
 */
function removeChunkedCache_(cache, key) {
  var countStr = cache.get(key + ':count');
  if (countStr !== null) {
    var count = parseInt(countStr, 10);
    var keysToRemove = [key + ':count'];
    for (var i = 0; i < count; i++) {
      keysToRemove.push(key + ':' + i);
    }
    cache.removeAll(keysToRemove);
  } else {
    // 分割なし単一エントリの削除
    cache.remove(key);
  }
}

/**
 * 社内カレンダーマスタ全件を取得する（内部用）。
 * キャッシュが有効な場合はキャッシュから返す。
 *
 * @returns {Object[]} 社内カレンダーレコードの Object 配列
 * @throws {Error} マスタブックを開けない場合、またはシートが見つからない場合
 */
function getCalendarAll_() {
  var cached = cacheGet_(CACHE_KEY_CALENDAR);
  if (cached !== null) {
    return cached;
  }

  var ss    = openMasterSpreadsheet_();
  var sheet = ss.getSheetByName(CALENDAR_SHEET_NAME);
  if (!sheet) {
    throw new Error('MasterService.getCalendarAll_: Sheet not found: ' + CALENDAR_SHEET_NAME);
  }

  var rows = parseRows_(sheet);
  cachePut_(CACHE_KEY_CALENDAR, rows, CACHE_TTL_SECONDS);
  return rows;
}

/**
 * 社内カレンダーレコード配列から日付→区分のマップを構築する。
 * @private
 * @param {Object[]} records - getCalendarAll_ が返すレコード配列
 * @returns {Object} { "2026-04-04": "休日", "2026-04-11": "出勤土曜", ... }
 */
function buildCalendarMapFromRecords_(records) {
  var map = {};
  records.forEach(function(r) {
    var dateKey = r['日付'];
    if (dateKey instanceof Date) {
      dateKey = Utilities.formatDate(dateKey, 'JST', 'yyyy-MM-dd');
    } else if (typeof dateKey === 'string') {
      // ISO 形式が含まれる場合に備えて日付部分のみ抽出
      dateKey = String(dateKey).slice(0, 10);
    }
    if (dateKey) map[dateKey] = r['区分'];
  });
  return map;
}

/**
 * 社内カレンダーを { "YYYY-MM-DD": "区分" } の形式で返す（O(1) ルックアップ用）。
 * 既存呼び出し互換のためのラッパー（getCalendarAll_ を呼んで buildCalendarMapFromRecords_ に委譲）。
 * @private
 * @returns {Object} { "2026-04-04": "休日", "2026-04-11": "出勤土曜", ... }
 */
function buildCalendarMap_() {
  return buildCalendarMapFromRecords_(getCalendarAll_());
}
