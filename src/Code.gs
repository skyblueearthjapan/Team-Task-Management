/**
 * GAS WebApp エントリーポイント。
 *
 * 現フェーズはセットアップ用途中心のため、UI（index.html / scripts.html / styles.html）と
 * MailQueueService / DataService 等は次フェーズで追加する。
 *
 * 初回利用フロー:
 *   1. このプロジェクトをスプレッドシートにバインドして開く
 *   2. GAS エディタで `setupAll()` を実行（SetupService.gs）
 *   3. Staff シートに6名分の氏名・メールアドレスを入力
 *   4. Settings シートで KOBAN_MASTER_SHEET_ID / KOBAN_MASTER_SHEET_NAME を設定
 */

function doGet(e) {
  const html =
    '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">' +
    '<title>機械設計技術部 タスク管理</title>' +
    '<style>' +
    '  body{font-family:system-ui,-apple-system,"Segoe UI","Yu Gothic UI",sans-serif;' +
    '       background:#0f172a;color:#e2e8f0;padding:48px;line-height:1.7}' +
    '  h1{margin-top:0}' +
    '  code{background:#1e293b;padding:2px 6px;border-radius:4px}' +
    '  ul{padding-left:1.4em}' +
    '</style></head><body>' +
    '<h1>機械設計技術部 タスク管理</h1>' +
    '<p>本 WebApp は構築中です（v0.1 セットアップ段階）。</p>' +
    '<h3>次の手順</h3>' +
    '<ol>' +
    '  <li>GAS エディタで <code>setupAll()</code> を実行（シート生成）</li>' +
    '  <li><code>Staff</code> シートに6名分の氏名・メールアドレスを入力</li>' +
    '  <li><code>Settings</code> シートで工番マスタの ID とシート名を設定</li>' +
    '</ol>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setTitle('機械設計技術部 タスク管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
