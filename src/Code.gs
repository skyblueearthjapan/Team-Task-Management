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

// 許可するメールドメイン（doGet UI 表示用）。
// WebApp のアクセス権限を「全員」に設定する必要があるため、ここで明示的にドメイン制限を行う。
const ALLOWED_EMAIL_DOMAIN = '@lineworks-local.info';

function doGet(e) {
  // ドメイン認可チェック（ベストエフォート）:
  //   WebApp は ANYONE_ANONYMOUS でデプロイしないと EXE の doPost が通らない。
  //   ANYONE_ANONYMOUS の場合 Session.getActiveUser().getEmail() は常に空文字を返す
  //   （Google のプライバシー仕様）。
  //
  //   よってここでは「メールが取得できた場合だけドメインを判定」する。
  //   取得できない場合（匿名アクセスまたは取得不能な状態）は通す。
  //   URL 自体は社内（LINE WORKS テナント）限定で配布する運用とすることで実質的に保護する。
  const userEmail = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '';
  if (userEmail && !endsWithDomain_(userEmail, ALLOWED_EMAIL_DOMAIN)) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">' +
      '<title>アクセス権限がありません</title>' +
      '<style>body{font-family:"Yu Gothic UI","Noto Sans JP",sans-serif;' +
      'background:#f1ebe0;color:#2a2520;padding:64px;line-height:1.7;max-width:680px;margin:0 auto}' +
      'h1{font-size:24px;font-weight:600;margin:0 0 16px}' +
      'code{background:#fbf8f1;padding:2px 6px;border-radius:4px;border:1px solid #e1d8c5}' +
      '</style></head><body>' +
      '<h1>アクセス権限がありません</h1>' +
      '<p>本アプリは <code>' + ALLOWED_EMAIL_DOMAIN + '</code> ドメインのユーザーのみ利用できます。</p>' +
      '<p>現在のログインアカウント: <code>' + userEmail + '</code></p>' +
      '<p>正しいアカウントへログインしなおしてください。</p>' +
      '</body></html>'
    ).setTitle('アクセス権限がありません');
  }

  const tpl = HtmlService.createTemplateFromFile('index');
  return tpl.evaluate()
    .setTitle('技術部タスク管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * メールアドレスが指定ドメインで終わるか判定する。
 * @param {string} email
 * @param {string} domain - '@example.com' 形式
 * @returns {boolean}
 */
function endsWithDomain_(email, domain) {
  if (!email || !domain) return false;
  return String(email).toLowerCase().endsWith(String(domain).toLowerCase());
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
