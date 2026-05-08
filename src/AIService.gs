/**
 * AIService.gs
 *
 * Gemini 2.5 Flash を使った自然言語解析・フィールド補完サービス。
 *
 * プライバシー保護方針:
 *   AI には「自由文 + WorkTypes リスト（id/name/displayOrder）+ 基準日」のみ送信する。
 *   工番マスタの「受注先・納入先・住所・品名」は絶対に AI に送らない。
 *   kobanCode を抽出した場合は GAS 側で工番マスタを検索して詳細を補完する。
 *
 * レート制御:
 *   CacheService（Script キャッシュ）を使い、同一 staffId が直近 60 秒で
 *   GEMINI_RATE_LIMIT_PER_MIN（既定 5）回を超えたら HTTP 429 相当を返す。
 */

// ─── Gemini API エンドポイント ────────────────────────────────
var GEMINI_API_BASE_ = 'https://generativelanguage.googleapis.com/v1beta/models/';

// ─── レート制御キープレフィックス ────────────────────────────
var RATE_LIMIT_CACHE_PREFIX_ = 'ai:rate:';

// ════════════════════════════════════════════════════════════════
// A. 公開関数
// ════════════════════════════════════════════════════════════════

/**
 * 自由文から構造化されたエントリを抽出する。
 * 工番コードのみ抽出し、詳細は GAS 側で工番マスタから補完する（プライバシー保護）。
 *
 * @param {string} text    - 自由文（複数件まとめて記述可）
 * @param {Object} context - { staffId, referenceDate }
 *   staffId       {string} レート制御の識別子
 *   referenceDate {string} YYYY-MM-DD 形式の基準日（省略時は本日）
 * @returns {Object}
 *   entries       {Array}  抽出されたエントリ配列
 *     kobanCode   {string|null}  抽出された工番コード
 *     workTypeId  {string|null}  WorkType.id
 *     duration    {string|null}  工数文字列（"3時間" "1日" "半日" 等）
 *     detail      {string|null}  作業内容
 *     periodStart {string|null}  YYYY-MM-DD
 *     periodEnd   {string|null}  YYYY-MM-DD
 *     confidence  {number}       0〜1
 *     reasoning   {string}       AI の判断根拠
 *     customer    {string|null}  GAS 側で補完（AI は設定しない）
 *     productName {string|null}  GAS 側で補完（AI は設定しない）
 *   unrecognized  {string}  AI が解釈できなかった部分
 *   error         {string}  エラーコード（エラー時のみ）
 *   message       {string}  エラー詳細（エラー時のみ）
 */
function aiParseNaturalLanguage(text, context) {
  context = context || {};
  var staffId       = context.staffId       || 'anonymous';
  var referenceDate = context.referenceDate || Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');

  // レート制御チェック
  var rateError = checkRateLimit_(staffId);
  if (rateError) return rateError;

  // API キー確認
  var apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEYS.GEMINI_API_KEY);
  if (!apiKey) {
    return { error: 'API_KEY_NOT_SET', message: 'GEMINI_API_KEY が Script Properties に設定されていません。GAS スクリプトプロパティに GEMINI_API_KEY を追加してください。' };
  }

  // WorkTypes を取得（AI に送る: id, name, displayOrder のみ。詳細は送らない）
  var workTypes = getWorkTypes().map(function(wt) {
    return { id: wt.id, name: wt.name, displayOrder: wt.displayOrder };
  });

  // システムプロンプト構築
  var systemPrompt = buildParseSystemPrompt_(referenceDate, workTypes);

  // レスポンススキーマ定義（OpenAPI 形式）
  var responseSchema = {
    type: 'OBJECT',
    properties: {
      entries: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            kobanCode:   { type: 'STRING', nullable: true },
            workTypeId:  { type: 'STRING', nullable: true },
            duration:    { type: 'STRING', nullable: true },
            detail:      { type: 'STRING', nullable: true },
            periodStart: { type: 'STRING', nullable: true },
            periodEnd:   { type: 'STRING', nullable: true },
            confidence:  { type: 'NUMBER' },
            reasoning:   { type: 'STRING' }
          },
          required: ['kobanCode', 'workTypeId', 'duration', 'detail', 'periodStart', 'periodEnd', 'confidence', 'reasoning']
        }
      },
      unrecognized: { type: 'STRING' }
    },
    required: ['entries', 'unrecognized']
  };

  // レート記録（API 呼び出し直前: 429 連打防止）
  recordRateLimit_(staffId);

  // Gemini 呼び出し
  var result = _callGemini_(systemPrompt, text, responseSchema);
  if (result.error) return result;

  // 工番マスタ補完（プライバシー保護: AI は工番詳細を知らない）
  var enriched = enrichWithKobanMaster_(result.entries || []);

  return {
    entries:      enriched,
    unrecognized: result.unrecognized || ''
  };
}

/**
 * 部分入力から不足フィールドを AI が補完する。
 * 入力済みのフィールドはそのまま、空フィールドのみ AI が提案する。
 *
 * @param {Object} hint    - { kobanCode?, workTypeId?, duration?, detail?, ... }
 * @param {Object} context - { staffId, referenceDate, section }
 *   staffId       {string} レート制御の識別子
 *   referenceDate {string} YYYY-MM-DD 形式の基準日
 *   section       {string} 'today' | 'prev'
 * @returns {Object}
 *   suggestion    {Object}  補完されたフィールド
 *     kobanCode   {string|null}
 *     workTypeId  {string|null}
 *     duration    {string|null}
 *     detail      {string|null}
 *     periodStart {string|null}
 *     periodEnd   {string|null}
 *   confidence    {number}   0〜1
 *   reasoning     {string}   AI の判断根拠
 *   error         {string}   エラーコード（エラー時のみ）
 *   message       {string}   エラー詳細（エラー時のみ）
 */
function aiSuggestForReport(hint, context) {
  context = context || {};
  var staffId       = context.staffId       || 'anonymous';
  var referenceDate = context.referenceDate || Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  var section       = context.section       || 'today';

  // レート制御チェック
  var rateError = checkRateLimit_(staffId);
  if (rateError) return rateError;

  // API キー確認
  var apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEYS.GEMINI_API_KEY);
  if (!apiKey) {
    return { error: 'API_KEY_NOT_SET', message: 'GEMINI_API_KEY が Script Properties に設定されていません。GAS スクリプトプロパティに GEMINI_API_KEY を追加してください。' };
  }

  // WorkTypes を取得（AI に送る: id, name のみ）
  var workTypes = getWorkTypes().map(function(wt) {
    return { id: wt.id, name: wt.name };
  });

  // システムプロンプト
  var systemPrompt = buildSuggestSystemPrompt_(referenceDate, workTypes, section);

  // ユーザーテキスト: hint を JSON 文字列として送信（工番詳細は含まない）
  var hintForAI = {
    kobanCode:  hint.kobanCode  || null,
    workTypeId: hint.workTypeId || null,
    duration:   hint.duration   || null,
    detail:     hint.detail     || null,
    periodStart: hint.periodStart || null,
    periodEnd:   hint.periodEnd   || null
  };
  var userText = '現在の入力状態:\n' + JSON.stringify(hintForAI, null, 2);

  // レスポンススキーマ
  var responseSchema = {
    type: 'OBJECT',
    properties: {
      suggestion: {
        type: 'OBJECT',
        properties: {
          kobanCode:   { type: 'STRING', nullable: true },
          workTypeId:  { type: 'STRING', nullable: true },
          duration:    { type: 'STRING', nullable: true },
          detail:      { type: 'STRING', nullable: true },
          periodStart: { type: 'STRING', nullable: true },
          periodEnd:   { type: 'STRING', nullable: true }
        },
        required: ['kobanCode', 'workTypeId', 'duration', 'detail', 'periodStart', 'periodEnd']
      },
      confidence: { type: 'NUMBER' },
      reasoning:  { type: 'STRING' }
    },
    required: ['suggestion', 'confidence', 'reasoning']
  };

  // レート記録（API 呼び出し直前: 429 連打防止）
  recordRateLimit_(staffId);

  // Gemini 呼び出し
  var result = _callGemini_(systemPrompt, userText, responseSchema);
  if (result.error) return result;

  // 工番マスタ補完（suggestion.kobanCode が取れた場合のみ）
  if (result.suggestion && result.suggestion.kobanCode) {
    var kobanDetail = lookupKobanDetail_(result.suggestion.kobanCode);
    result.suggestion.customer    = kobanDetail.customer    || null;
    result.suggestion.productName = kobanDetail.productName || null;
  }

  return result;
}

/**
 * 内部: Gemini API 呼び出し。
 * - エンドポイント: generativelanguage.googleapis.com/v1beta/models/{model}:generateContent
 * - 認証: x-goog-api-key ヘッダ
 * - responseSchema で構造化 JSON を強制
 * - リトライ 4 回・指数バックオフ・429/5xx 対応
 *
 * @param {string} systemPrompt    - システムプロンプト文字列
 * @param {string} userText        - ユーザー入力テキスト
 * @param {Object} responseSchema  - OpenAPI 形式のレスポンススキーマ
 * @returns {Object} パース済み JSON オブジェクト、またはエラーオブジェクト
 */
function _callGemini_(systemPrompt, userText, responseSchema) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEYS.GEMINI_API_KEY);
  if (!apiKey) {
    return { error: 'API_KEY_NOT_SET', message: 'GEMINI_API_KEY が未設定です。' };
  }

  var model   = getSetting('GEMINI_MODEL') || 'gemini-2.5-flash';
  var url     = GEMINI_API_BASE_ + model + ':generateContent';

  var payload = {
    system_instruction: {
      parts: [{ text: systemPrompt }]
    },
    contents: [
      { role: 'user', parts: [{ text: userText }] }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema:   responseSchema
    }
  };

  var options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { 'x-goog-api-key': apiKey },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var delays = [1000, 2000, 4000, 8000]; // 指数バックオフ（ms）
  var lastError = null;

  for (var attempt = 0; attempt <= delays.length; attempt++) {
    if (attempt > 0) {
      Utilities.sleep(delays[attempt - 1]);
    }

    try {
      var response    = UrlFetchApp.fetch(url, options);
      var statusCode  = response.getResponseCode();
      var rawBody     = response.getContentText();

      // 成功
      if (statusCode === 200) {
        try {
          var parsed = JSON.parse(rawBody);
          // finishReason チェック（HIGH-1）
          var finishReason = (parsed.candidates && parsed.candidates[0])
                              ? parsed.candidates[0].finishReason : 'UNKNOWN';
          if (finishReason !== 'STOP') {
            Logger.log('[AIService._callGemini_] finishReason=' + finishReason + ' attempt=' + attempt);
            lastError = { error: 'AI_FAILED', message: 'Gemini の応答が不完全です (finishReason: ' + finishReason + ')' };
            continue; // リトライ対象
          }
          // Gemini レスポンス構造から text を取り出す
          var text = parsed.candidates &&
                     parsed.candidates[0] &&
                     parsed.candidates[0].content &&
                     parsed.candidates[0].content.parts &&
                     parsed.candidates[0].content.parts[0] &&
                     parsed.candidates[0].content.parts[0].text;
          if (!text) {
            lastError = { error: 'AI_FAILED', message: 'Gemini からの応答テキストが空です。' };
            continue;
          }
          return JSON.parse(text);
        } catch (parseErr) {
          return { error: 'AI_FAILED', message: 'Gemini 応答の JSON パースに失敗しました: ' + parseErr.message };
        }
      }

      // 429 / 5xx はリトライ対象
      if (statusCode === 429 || statusCode >= 500) {
        lastError = { error: 'AI_FAILED', message: 'Gemini API エラー (HTTP ' + statusCode + '): ' + rawBody.substring(0, 200) };
        Logger.log('[AIService._callGemini_] HTTP ' + statusCode + ', attempt=' + attempt + '. Retrying...');
        continue;
      }

      // 4xx（429以外）はリトライしない
      return { error: 'AI_FAILED', message: 'Gemini API エラー (HTTP ' + statusCode + '): ' + rawBody.substring(0, 200) };

    } catch (fetchErr) {
      lastError = { error: 'AI_FAILED', message: 'UrlFetch 例外: ' + fetchErr.message };
      Logger.log('[AIService._callGemini_] Fetch exception, attempt=' + attempt + ': ' + fetchErr.message);
    }
  }

  return lastError || { error: 'AI_FAILED', message: 'Gemini API への接続に失敗しました。' };
}

// ════════════════════════════════════════════════════════════════
// B. 内部ヘルパ（末尾アンダースコア）
// ════════════════════════════════════════════════════════════════

/**
 * aiParseNaturalLanguage 用システムプロンプトを生成する。
 *
 * @param {string}   referenceDate - YYYY-MM-DD 形式の基準日
 * @param {Object[]} workTypes     - [{ id, name, displayOrder }, ...]
 * @returns {string}
 */
function buildParseSystemPrompt_(referenceDate, workTypes) {
  // ユーザー要件: タイトル名を「技術部」に統一（旧「機械設計技術部」表記を削除）
  return 'あなたは技術部のスタッフを補助するアシスタントです。\n' +
    '基準日: ' + referenceDate + '（YYYY-MM-DD）\n\n' +
    'WorkTypes 一覧（id, name, displayOrder）:\n' +
    JSON.stringify(workTypes, null, 2) + '\n\n' +
    'ルール:\n' +
    '- 「明日」→ 基準日 +1 日（営業日でなくても OK）\n' +
    '- 「来週月曜」→ 基準日に最も近い来週月曜\n' +
    '- 「半日」「3時間」「2日」等の表現はそのまま duration に保存\n' +
    '- 工番らしき文字列（例: LW23012, K-12345）が含まれていれば kobanCode に抽出\n' +
    '- 工番が含まれない / 不明確な場合は kobanCode を null に\n' +
    '- WorkTypes は name で完全一致 → なければ最も近い意味のものを選ぶ → 不明なら workTypeId を null\n' +
    '- 確信が低い項目は confidence を低く設定（0〜1）\n' +
    '- 複数の作業が記述されている場合は entries を複数件返す\n' +
    '- 出力は必ず JSON、説明文は reasoning に分離\n' +
    '- periodStart / periodEnd は基準日を起点に解釈し YYYY-MM-DD 形式で返す\n' +
    '- 日付が明示されていない場合は periodStart と periodEnd に基準日を設定する';
}

/**
 * aiSuggestForReport 用システムプロンプトを生成する。
 *
 * @param {string}   referenceDate - YYYY-MM-DD 形式の基準日
 * @param {Object[]} workTypes     - [{ id, name }, ...]
 * @param {string}   section       - 'today' | 'prev'
 * @returns {string}
 */
function buildSuggestSystemPrompt_(referenceDate, workTypes, section) {
  var sectionNote = section === 'today'
    ? '本日（' + referenceDate + '）の作業報告を補完します。'
    : '前日までの作業報告を補完します。';

  // ユーザー要件: タイトル名を「技術部」に統一（旧「機械設計技術部」表記を削除）
  return 'あなたは技術部のスタッフを補助するアシスタントです。\n' +
    '基準日: ' + referenceDate + '（YYYY-MM-DD）\n' +
    sectionNote + '\n\n' +
    'WorkTypes 一覧（id, name）:\n' +
    JSON.stringify(workTypes, null, 2) + '\n\n' +
    'ルール:\n' +
    '- 入力済みのフィールド（null でない値）はそのまま suggestion に返す\n' +
    '- null のフィールドのみ推測・補完する\n' +
    '- workTypeId は WorkTypes の id 値から選ぶ\n' +
    '- 推測が難しい場合は null を返し、confidence を低く設定\n' +
    '- periodStart / periodEnd が null の場合は基準日を設定する\n' +
    '- 出力は必ず JSON、判断根拠は reasoning に記述する';
}

/**
 * entries 配列の各エントリに対して工番マスタを検索し、
 * customer / productName を補完する（プライバシー保護: AI は工番詳細を知らない）。
 *
 * @param {Object[]} entries - AI が返した entries 配列
 * @returns {Object[]}       - customer / productName が補完された entries 配列
 */
function enrichWithKobanMaster_(entries) {
  return entries.map(function(entry) {
    var enriched = {};
    // エントリの全フィールドをコピー
    Object.keys(entry).forEach(function(k) { enriched[k] = entry[k]; });

    // kobanCode を正規化・検証
    var kobanCode = entry.kobanCode ? String(entry.kobanCode).trim() : null;
    if (kobanCode && !isValidKobanCode_(kobanCode)) {
      // ハルシネーション検出: 工番コード形式に合わない場合は null に正規化
      Logger.log('[AIService.enrichWithKobanMaster_] Invalid kobanCode normalized to null: ' + kobanCode);
      kobanCode = null;
    }
    enriched.kobanCode = kobanCode;

    // customer / productName の補完
    enriched.customer    = null;
    enriched.productName = null;

    if (kobanCode) {
      var detail = lookupKobanDetail_(kobanCode);
      enriched.customer    = detail.customer    || null;
      enriched.productName = detail.productName || null;
    }

    return enriched;
  });
}

/**
 * 工番コードで工番マスタを検索し、受注先・品名を返す。
 * 見つからない場合は空オブジェクトを返す（UI 側で「マスタ外」表示）。
 *
 * @param {string} kobanCode - 工番コード
 * @returns {{ customer: string, productName: string }}
 */
function lookupKobanDetail_(kobanCode) {
  if (!kobanCode) return {};

  try {
    var master = getKobanMaster();
    var found  = master.find(function(row) {
      return String(row['工番'] || '').trim() === String(kobanCode).trim();
    });
    if (!found) return {};
    return {
      customer:    String(found['受注先'] || '').trim() || null,
      productName: String(found['品名']   || '').trim() || null
    };
  } catch (e) {
    Logger.log('[AIService.lookupKobanDetail_] 工番マスタ検索エラー: ' + e.message);
    return {};
  }
}

/**
 * 工番コードの形式を検証する（ハルシネーション検出）。
 * 英字1〜3文字 + 数字4〜6桁、またはハイフン区切り等の一般的なパターンを許容する。
 *
 * 許容例: LW23012, K-12345, ABC001, lw-2026-001
 * 拒否例: 作業, 明日, 構想図, 12（数字のみ短すぎ）
 *
 * @param {string} code
 * @returns {boolean}
 */
function isValidKobanCode_(code) {
  if (!code || typeof code !== 'string') return false;
  var s = code.trim();
  if (s.length < 3 || s.length > 20) return false;
  // 英字を含み、かつ数字を含む（工番コードの最低条件）
  // 完全に数字のみ、または完全に日本語のみは拒否
  var hasAlpha  = /[A-Za-z]/.test(s);
  var hasDigit  = /[0-9]/.test(s);
  var hasKanji  = /[一-鿿]/.test(s);
  if (hasKanji) return false;
  return hasAlpha && hasDigit;
}

/**
 * 同一 staffId のレート制限をチェックする。
 * 直近 60 秒の呼び出し回数が上限を超えていればエラーオブジェクトを返す。
 *
 * @param {string} staffId
 * @returns {Object|null} エラーオブジェクト（超過時）または null（正常時）
 */
function checkRateLimit_(staffId) {
  var limit     = parseInt(getSetting('GEMINI_RATE_LIMIT_PER_MIN') || '5', 10);
  var cacheKey  = RATE_LIMIT_CACHE_PREFIX_ + staffId;
  var cache     = CacheService.getScriptCache();
  var storedRaw = cache.get(cacheKey);

  if (storedRaw) {
    try {
      var timestamps = JSON.parse(storedRaw);
      var now        = Date.now();
      // 直近 60 秒以内のタイムスタンプのみ残す
      var recent = timestamps.filter(function(ts) { return (now - ts) < 60000; });
      if (recent.length >= limit) {
        return {
          error:   'RATE_LIMIT_EXCEEDED',
          message: '1分間のAI呼び出し上限（' + limit + '回）を超えました。しばらくお待ちください。'
        };
      }
    } catch (e) {
      // パース失敗は無視して続行
    }
  }

  return null;
}

/**
 * レート制御用タイムスタンプを記録する。
 * CacheService に直近 60 秒のタイムスタンプ配列を保存する（TTL 70 秒）。
 *
 * @param {string} staffId
 */
function recordRateLimit_(staffId) {
  var cacheKey  = RATE_LIMIT_CACHE_PREFIX_ + staffId;
  var cache     = CacheService.getScriptCache();
  var storedRaw = cache.get(cacheKey);
  var now       = Date.now();

  var timestamps = [];
  if (storedRaw) {
    try {
      var parsed = JSON.parse(storedRaw);
      // 直近 60 秒以内のみ保持
      timestamps = parsed.filter(function(ts) { return (now - ts) < 60000; });
    } catch (e) {
      // パース失敗は空配列から開始
    }
  }

  timestamps.push(now);

  try {
    cache.put(cacheKey, JSON.stringify(timestamps), 70); // TTL 70 秒
  } catch (e) {
    Logger.log('[AIService.recordRateLimit_] Cache put failed: ' + e.message);
  }
}
