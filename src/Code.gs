/**
 * GAS WebApp エントリーポイント。
 *
 * 初回利用フロー:
 *   1. このプロジェクトをスプレッドシートにバインドして開く
 *   2. GAS エディタで `setupAll()` を実行（SetupService.gs）
 *   3. Staff シートに6名分の氏名・メールアドレスを入力
 *   4. Settings シートで KOBAN_MASTER_SHEET_ID / KOBAN_MASTER_SHEET_NAME を設定
 *
 * doPost: EXE エージェントからの RPC リクエストを受信するエンドポイント（§7.4.4）。
 *   認証: クエリパラメータまたは POST ボディの token を EXE_API_TOKEN と照合。
 *   アクション: pickMailItem / completeMailItem / heartbeat
 */

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('index');
  return tpl.evaluate()
    .setTitle('機械設計技術部 タスク管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * テンプレートファイルのコンテンツを返すヘルパ。
 * index.html 内の <?!= include('styles') ?> / <?!= include('scripts') ?> で使用。
 *
 * @param {string} filename - ファイル名（拡張子なし）
 * @returns {string} ファイルのコンテンツ文字列
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── EXE 向け doPost ルーター（§7.4.4） ─────────────────────────

/**
 * EXE エージェントからの HTTP POST リクエストを受信し、
 * action フィールドに基づいて対応する関数へディスパッチする。
 *
 * 認証: クエリパラメータ `?token=...` または POST ボディ内の `token` を
 *       スクリプトプロパティ EXE_API_TOKEN と照合する（仕様 §7.4.4）。
 * 失敗時: HTTP ステータスは GAS 仕様上常に 200 だが、レスポンスボディの
 *         status / error フィールドで呼び出し元が判別する。
 *
 * リクエスト形式:
 *   推奨（フラット構造）:
 *     { action: 'pickMailItem',    token: '...', hostname: 'PC01' }
 *     { action: 'completeMailItem',token: '...', id: 'mq_xxx', status: 'sent', errorMessage: '' }
 *     { action: 'heartbeat',       token: '...', hostname: 'PC01' }
 *   互換（args ラッパー）:
 *     { action: 'pickMailItem',    token: '...', args: { hostname: 'PC01' } }
 *     { action: 'completeMailItem',token: '...', args: { id: 'mq_xxx', result: { status: 'sent' } } }
 *     { action: 'heartbeat',       token: '...', args: { hostname: 'PC01' } }
 *
 * @param {Object} e - GAS の doPost イベントオブジェクト
 * @returns {GoogleAppsScript.Content.TextOutput} JSON レスポンス
 */
function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');

    // EXE_API_TOKEN 認証: クエリパラメータ優先、なければボディ内 token を参照
    const token      = (e.parameter && e.parameter.token) || payload.token || '';
    const validToken = PropertiesService.getScriptProperties()
                         .getProperty(SCRIPT_PROP_KEYS.EXE_API_TOKEN) || '';
    if (!validToken || token !== validToken) {
      return jsonResponse_({ status: 401, message: 'unauthorized' });
    }

    const action = payload.action || '';
    // args ラッパー互換: フラット構造優先、なければ args 内を参照
    const args   = payload.args   || {};

    switch (action) {
      case 'pickMailItem':
        return jsonResponse_(pickMailItem(payload.hostname || args.hostname || args.exeId));

      case 'completeMailItem': {
        // フラット payload を優先しつつ、args ラッパー互換のため args.result も参照する。
        // payload.status / errorMessage が "存在する" 場合は空文字列でも採用する
        // （errorMessage:"" は仕様上正常な「エラーなし」シグナル）。
        const argsResult    = (args && args.result) || {};
        const hasFlatStatus = Object.prototype.hasOwnProperty.call(payload, 'status');
        const hasFlatError  = Object.prototype.hasOwnProperty.call(payload, 'errorMessage');
        return jsonResponse_(completeMailItem(
          payload.id || args.id,
          {
            status:       hasFlatStatus ? payload.status       : argsResult.status,
            errorMessage: hasFlatError  ? payload.errorMessage : argsResult.errorMessage
          }
        ));
      }

      case 'heartbeat':
        return jsonResponse_(heartbeat(payload.hostname || args.hostname));

      default:
        return jsonResponse_({ status: 400, message: 'unknown action: ' + action });
    }

  } catch (err) {
    return jsonResponse_({ status: 500, message: String(err) });
  }
}

/**
 * Object を JSON 文字列に変換した ContentService.TextOutput を返す。
 *
 * @param {Object} obj - レスポンスとして返すオブジェクト
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
