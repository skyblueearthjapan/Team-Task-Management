/**
 * MailQueueService.gs
 *
 * EXE エージェント向け API 群:
 *   - enqueueMailRequest()   : UI → MailQueue に pending 追加
 *   - pickMailItem()         : EXE → CAS で pending 1 件を picked に遷移（重複防止の核心）
 *   - completeMailItem()     : EXE → drafted / sent / failed に更新
 *   - retryMailItem()        : 元レコードを複製して新規 pending を生成（仕様 §7.4.6 準拠）
 *   - archiveOldMailQueue()  : 90 日以上前の完了レコードを MailQueue_Archive へ移動
 *   - checkExeAlive()        : EXE 死活チェック・管理者通知（6 時間重複抑制）
 *   - notifyAdminOnFailure() : 管理者への Gmail 通知
 *   - heartbeat()            : EXE からのハートビート受信
 *   - recoverStalePicked()   : picked タイムアウト復旧（GAS TimeTrigger で定期実行）
 */

// ════════════════════════════════════════════════════════════════
// A. 公開関数
// ════════════════════════════════════════════════════════════════

// ── キュー追加 ──────────────────────────────────────────────────

/**
 * メール送信リクエストを MailQueue に追加する。
 * targetStaffId と Staff マスタの email が一致しているか確認したうえで、
 * スナップショット（targetStaffName / targetStaffEmail）をレコードに保存する。
 *
 * toAddresses は **カンマ区切り string** で保存する（仕様 v1.0 §6.3 確定）。
 * 入力が配列の場合は join、文字列の場合はそのまま使う。
 *
 * extraRecipients を指定した場合、重複排除して toAddresses に追加する。
 *
 * bodyVars に staffName / todayItems / prevItems（または yesterdayItems）が含まれる場合、
 * buildMailBody() を呼び出して fullBody を生成し bodyVars に格納する。
 * 挨拶は buildMailBody 側で「関係各位」固定（params.greeting は無視される）。
 *
 * @param {Object} params
 *   requestedBy      {string}
 *   targetStaffId    {string}
 *   reportDate       {string}  YYYY-MM-DD
 *   mode             {string}  'draft' | 'send'（省略時 'draft'）
 *   toAddresses      {string|string[]} カンマ区切り文字列、または文字列配列
 *   extraRecipients  {string|string[]} 追加宛先（省略可）
 *   ccAddresses      {string}  省略時 ''
 *   subjectVars      {Object|string}
 *   bodyVars         {Object|string}  staffName / todayItems / prevItems（または yesterdayItems）を含む場合 fullBody を自動生成（挨拶は GAS 側で「関係各位」固定）
 * @returns {{ success: true, id: string } | { success: false, reason: string }}
 */
function enqueueMailRequest(params) {
  // mode バリデーション（仕様 §6.3）
  const VALID_MODES = ['draft', 'send'];
  const mode = params.mode || 'draft';
  if (!VALID_MODES.includes(mode)) {
    return { success: false, reason: 'INVALID_MODE' };
  }

  // Staff マスタからスナップショット取得 & 整合性確認
  const staff = getById(SHEET_NAMES.STAFF, params.targetStaffId);
  if (!staff) {
    return { success: false, reason: 'STAFF_NOT_FOUND' };
  }
  if (!staff.email) {
    return { success: false, reason: 'EMAIL_NOT_SET' };
  }

  // toAddresses をカンマ区切り string に正規化
  const baseAddrs = Array.isArray(params.toAddresses)
    ? params.toAddresses.filter(Boolean)
    : (params.toAddresses ? String(params.toAddresses).split(',').map(s => s.trim()).filter(Boolean) : []);

  // extraRecipients を正規化してマージ（重複排除）
  const extraAddrs = Array.isArray(params.extraRecipients)
    ? params.extraRecipients.filter(Boolean)
    : (params.extraRecipients ? String(params.extraRecipients).split(',').map(s => s.trim()).filter(Boolean) : []);

  const toAddrSet = {};
  baseAddrs.forEach(function(a) { toAddrSet[a] = true; });
  extraAddrs.forEach(function(a) { toAddrSet[a] = true; });
  const toAddrCsv = Object.keys(toAddrSet).join(',');

  // bodyVars を Object として取得
  let bodyVarsObj = params.bodyVars;
  if (typeof bodyVarsObj === 'string') {
    try { bodyVarsObj = JSON.parse(bodyVarsObj); } catch (e) { bodyVarsObj = {}; }
  }
  bodyVarsObj = bodyVarsObj || {};

  // 署名情報を Staff マスタからフォールバック付きで設定
  bodyVarsObj.signatureCompany = '株式会社ラインワークス';
  bodyVarsObj.signatureName    = (staff.signatureName  || staff.name  || '');
  bodyVarsObj.signatureEmail   = (staff.signatureEmail || staff.email || '');
  bodyVarsObj.staffName        = bodyVarsObj.staffName || staff.name || '';

  // メール本文を GAS 側で構築（todayItems / prevItems / yesterdayItems がある場合）
  if (bodyVarsObj.todayItems || bodyVarsObj.prevItems || bodyVarsObj.yesterdayItems) {
    bodyVarsObj.fullBody = buildMailBody(bodyVarsObj);
  }

  const record = {
    id:                generateId_('mq'),
    requestedBy:       params.requestedBy  || '',
    targetStaffId:     params.targetStaffId,
    targetStaffName:   staff.name          || '',   // ★スナップショット
    targetStaffEmail:  staff.email         || '',   // ★スナップショット
    reportDate:        params.reportDate   || '',
    mode:              mode,
    toAddresses:       toAddrCsv,                   // ★カンマ区切り string
    ccAddresses:       params.ccAddresses  || '',   // v1.0 では常に空文字列
    subjectVars:       typeof params.subjectVars === 'string'
                         ? params.subjectVars
                         : JSON.stringify(params.subjectVars || {}),
    bodyVars:          JSON.stringify(bodyVarsObj),
    status:            'pending',
    pickedBy:          '',
    pickedAt:          '',
    processedAt:       '',
    errorMessage:      '',
    previousRequestId: '',   // ★初回送信時は空。再送時のみ retryMailItem が設定
    createdAt:         nowIso_()
  };

  appendRow(SHEET_NAMES.MAIL_QUEUE, record);
  return { success: true, id: record.id };
}

// ── メール本文構築 ─────────────────────────────────────────────────

/**
 * メール本文を構築して返す。
 * enqueueMailRequest から内部で呼び出されるほか、フロントエンドのプレビュー用に
 * previewMailBody() 経由でも直接呼び出し可能。
 *
 * フォーマット仕様:
 *   本日・前日ともに同じ構造:
 *     1行目: 工番コード 受注先 品名（kobanCode が空の場合は省略）
 *     2行目:    作業内容（予定工数: ...）状態語
 *   状態語:
 *     本日: 常に「継続中」
 *     前日: continued=true → 「継続中」、そうでなければ「完了」
 *   冒頭の [完了][継続] 表記は廃止。
 *
 * @param {Object} params  bodyVars オブジェクト（または JSON 文字列）
 *   staffName         {string}  スタッフ名
 *   （注: params.greeting は無視。挨拶は「関係各位」固定）
 *   todayItems        {Array}   本日の作業内容の配列
 *     [{ detail, duration, kobanCode, customer, productName }]
 *   prevItems         {Array}   前日までの作業報告の配列（yesterdayItems も受け付ける）
 *     [{ detail, kobanCode, customer, productName, continued }]
 *   signatureCompany  {string}  署名：会社名
 *   signatureName     {string}  署名：氏名
 *   signatureEmail    {string}  署名：メールアドレス
 * @returns {string} 構築済みメール本文
 */
function buildMailBody(params) {
  if (typeof params === 'string') {
    try { params = JSON.parse(params); } catch (e) { params = {}; }
  }
  params = params || {};

  // 挨拶は常に「関係各位」に固定（G: ユーザー追加要件 #2）
  var greeting    = '関係各位';
  var staffName   = params.staffName       || '';
  var todayItems  = params.todayItems      || [];
  // prevItems を正式キーとし、yesterdayItems を後方互換フォールバックとして受け付ける
  var prevItems   = params.prevItems       || params.yesterdayItems || [];
  var sigCompany  = params.signatureCompany || '株式会社ラインワークス';
  var sigName     = params.signatureName   || '';
  var sigEmail    = params.signatureEmail  || '';

  var lines = [];

  // 挨拶
  if (greeting) {
    lines.push(greeting);
    lines.push('');
  }

  lines.push('お疲れ様です。' + staffName + ' です。');
  lines.push('本日の業務をご報告いたします。');
  lines.push('');

  // ▼ 本日の作業内容
  lines.push('▼ 本日の作業内容');
  // detail が空のアイテムはスキップ
  var validTodayItems = todayItems.filter(function(item) { return item.detail && String(item.detail).trim() !== ''; });
  if (validTodayItems.length > 0) {
    validTodayItems.forEach(function(item, idx) {
      // 1行目: 工番情報（kobanCode がある場合のみ出力）
      var kobanParts = [item.kobanCode, item.customer, item.productName].filter(Boolean);
      if (kobanParts.length > 0) {
        lines.push((idx + 1) + '. ' + kobanParts.join(' '));
      }
      // 2行目: 作業内容 + （予定工数: ...） + 継続中
      var detailStr  = String(item.detail).trim();
      var durationStr = (item.duration && String(item.duration).trim() !== '')
                        ? '（予定工数: ' + String(item.duration).trim() + '）'
                        : '';
      // kobanCode がない場合は番号を2行目に付ける
      var prefix = kobanParts.length > 0 ? '   ' : (idx + 1) + '. ';
      lines.push(prefix + detailStr + durationStr + ' 継続中');
      lines.push('');
    });
  } else {
    lines.push('（なし）');
    lines.push('');
  }

  // ▼ 前日までの作業報告
  lines.push('▼ 前日までの作業報告');
  // detail が空のアイテムはスキップ
  var validPrevItems = prevItems.filter(function(item) { return item.detail && String(item.detail).trim() !== ''; });
  if (validPrevItems.length > 0) {
    validPrevItems.forEach(function(item, idx) {
      // continued の真偽値正規化（boolean / 'true' / 'TRUE' すべて対応）
      var isContinued = (item.continued === true ||
        String(item.continued).trim().toLowerCase() === 'true');
      var statusLabel = isContinued ? '継続中' : '完了';
      // 1行目: 工番情報（kobanCode がある場合のみ出力）
      var kobanParts = [item.kobanCode, item.customer, item.productName].filter(Boolean);
      if (kobanParts.length > 0) {
        lines.push((idx + 1) + '. ' + kobanParts.join(' '));
      }
      // 2行目: 作業内容 + 状態語
      var detailStr = String(item.detail).trim();
      var prefix = kobanParts.length > 0 ? '   ' : (idx + 1) + '. ';
      lines.push(prefix + detailStr + ' ' + statusLabel);
      lines.push('');
    });
  } else {
    lines.push('（なし）');
  }
  lines.push('');

  lines.push('以上、よろしくお願いいたします。');
  lines.push('');
  lines.push('-----------------------------------------------------------');
  lines.push('               ' + sigCompany);
  lines.push('　　　　　　　　　' + sigName);
  lines.push('Mail:' + sigEmail);
  lines.push('-----------------------------------------------------------');

  return lines.join('\n');
}

/**
 * メール本文プレビューを返す（フロントエンドのモーダルプレビュー用）。
 * 実体は buildMailBody と同じだが、enqueueMailRequest を経由せず直接呼び出し可能。
 *
 * @param {Object|string} params - buildMailBody と同じパラメータ
 * @returns {{ success: true, body: string } | { success: false, reason: string }}
 */
function previewMailBody(params) {
  try {
    var body = buildMailBody(params);
    return { success: true, body: body };
  } catch (e) {
    return { success: false, reason: e.message };
  }
}

// ── atomic CAS: pending → picked ────────────────────────────────

/**
 * EXE が pending 状態の最古レコード 1 件を取得し picked に遷移させる。
 * LockService.getScriptLock() によりスクリプトレベルで排他制御する（重複防止の核心）。
 *
 * picked 確定後、targetStaffEmail と Staff マスタ最新 email を再照合する（仕様 §7.4.2 ③）。
 * 不一致の場合は status='failed', errorMessage='email_mismatch' に更新し管理者通知する。
 *
 * @param {string} exeId - EXE の識別子（PC 名 等）。doPost からは payload.hostname / args.hostname / args.exeId の順で解決される
 * @returns {Object|null} picked レコード、または pending なし時は null（仕様 §7.4.4）
 */
function pickMailItem(exeId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // 15 秒待機。取得できなければ例外をスロー

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MAIL_QUEUE);
    if (!sheet) throw new Error('pickMailItem: MailQueue sheet not found');

    const data      = sheet.getDataRange().getValues();
    const headers   = data[0];
    const statusIdx   = headers.indexOf('status');
    const pickedByIdx = headers.indexOf('pickedBy');
    const pickedAtIdx = headers.indexOf('pickedAt');

    // ヘッダ行（i=0）をスキップして pending 最古レコードを探す
    for (let i = 1; i < data.length; i++) {
      if (data[i][statusIdx] !== 'pending') continue;

      // CAS: LockService 配下で pending を確認した直後に picked へ書き込む
      sheet.getRange(i + 1, statusIdx   + 1).setValue('picked');
      sheet.getRange(i + 1, pickedByIdx + 1).setValue(exeId);
      sheet.getRange(i + 1, pickedAtIdx + 1).setValue(nowIso_());
      SpreadsheetApp.flush(); // 即時コミット（他プロセスへの可視化）

      // picked 確定後のレコードを取得して返す
      const row    = sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0];
      const record = Object.fromEntries(headers.map((h, j) => [h, row[j]]));

      // MailLog: picked イベントを記録
      logMailEvent_(record.id, 'info', 'picked', 'pickedBy=' + exeId);

      // ── §7.4.2 ③: メールアドレス再照合 ──────────────────────
      // picked 確定後に Staff マスタの最新 email と突合する
      if (!verifyEmailMatch_(record.targetStaffId, record.targetStaffEmail)) {
        updateRow(SHEET_NAMES.MAIL_QUEUE, record.id, {
          status:       'failed',
          processedAt:  nowIso_(),
          errorMessage: 'email_mismatch'
        });
        logMailEvent_(record.id, 'error', 'mismatch',
          'email_mismatch: targetStaffEmail=' + record.targetStaffEmail);
        notifyAdminOnFailure(
          '[タスク管理] メールアドレス不一致: ' + record.id,
          'MailQueue ID: ' + record.id + '\n' +
          'targetStaffId: ' + record.targetStaffId + '\n' +
          'MailQueue に保存されたメール: ' + record.targetStaffEmail + '\n' +
          'Staff マスタの現在のメール: ' +
            ((getById(SHEET_NAMES.STAFF, record.targetStaffId) || {}).email || '（取得不可）')
        );
        // picked → failed に遷移したレコードを更新して返す
        record.status       = 'failed';
        record.processedAt  = nowIso_();
        record.errorMessage = 'email_mismatch';
      }

      return record;
    }

    // pending なし: 仕様 §7.4.4 準拠で null を返す
    return null;

  } finally {
    lock.releaseLock();
  }
}

// ── 完了・失敗更新 ────────────────────────────────────────────────

/**
 * EXE がメール処理後に状態を更新する。
 * result = { status: 'drafted' | 'sent' | 'failed', errorMessage?: string }
 * MailLog にもイベント行を追加する。
 *
 * @param {string} id     - MailQueue レコード ID
 * @param {Object} result - { status, errorMessage? }
 * @returns {boolean} 更新成功: true、行が見つからない: false
 */
function completeMailItem(id, result) {
  const newStatus    = result.status       || 'failed';
  const errorMessage = result.errorMessage || '';

  const updated = updateRow(SHEET_NAMES.MAIL_QUEUE, id, {
    status:       newStatus,
    processedAt:  nowIso_(),
    errorMessage: errorMessage
  });

  // MailLog に記録
  logMailEvent_(
    id,
    newStatus === 'failed' ? 'error' : 'info',
    newStatus,
    errorMessage
  );

  return updated;
}

// ── 再送 ─────────────────────────────────────────────────────────

/**
 * 既存の MailQueue レコードを **複製** し、新レコードを pending として追加する。
 * 仕様 §7.4.6 準拠: 元レコードは一切変更しない。
 * pickedBy / pickedAt / processedAt / errorMessage は空文字列（リセット）。
 *
 * @param {string} originalId - 元 MailQueue レコードの id
 * @returns {{ success: true, newId: string } | { success: false, reason: string }}
 */
function retryMailItem(originalId) {
  const original = getById(SHEET_NAMES.MAIL_QUEUE, originalId);
  if (!original) {
    return { success: false, reason: 'NOT_FOUND' };
  }

  const newRecord = {
    id:                generateId_('mq'),
    requestedBy:       original.requestedBy,
    targetStaffId:     original.targetStaffId,
    targetStaffName:   original.targetStaffName,
    targetStaffEmail:  original.targetStaffEmail,
    reportDate:        original.reportDate,
    mode:              original.mode,
    toAddresses:       original.toAddresses,    // カンマ区切り string のままコピー
    ccAddresses:       original.ccAddresses,
    subjectVars:       original.subjectVars,
    bodyVars:          original.bodyVars,
    status:            'pending',               // ★必ず pending でスタート
    pickedBy:          '',                      // ★リセット
    pickedAt:          '',                      // ★リセット
    processedAt:       '',                      // ★リセット
    errorMessage:      '',                      // ★リセット
    previousRequestId: originalId,             // ★元レコード id で連鎖
    createdAt:         nowIso_()
  };

  appendRow(SHEET_NAMES.MAIL_QUEUE, newRecord);

  // MailLog: 両レコードに retry イベントを記録
  logMailEvent_(originalId, 'info', 'retry', '再送: 新レコード ' + newRecord.id + ' を生成');
  logMailEvent_(newRecord.id, 'info', 'retry', 'previousRequestId=' + originalId + ' から複製');

  return { success: true, newId: newRecord.id };
}

// ── アーカイブ ────────────────────────────────────────────────────

/**
 * 90 日以上前の完了レコード（status: sent / drafted / failed）を
 * MailQueue_Archive へ移動する。
 * SetupService.createMailQueueArchiveIfMissing() でアーカイブシートを確保してから実行。
 * 仕様 §11 の移動方式（行コピー → MailQueue から削除）に準拠。
 * GAS TimeTrigger で日次実行する想定（トリガー登録は別途）。
 */
function archiveOldMailQueue() {
  const ARCHIVE_DAYS = 90;
  const thresholdMs  = ARCHIVE_DAYS * 24 * 60 * 60 * 1000;
  const now          = Date.now();

  // アーカイブ先シートを確保
  createMailQueueArchiveIfMissing();

  const rows = listAll(SHEET_NAMES.MAIL_QUEUE);
  const archiveTargets = rows.filter(r => {
    if (!['sent', 'drafted', 'failed'].includes(r.status)) return false;
    if (!r.processedAt) return false;
    const processed = new Date(r.processedAt).getTime();
    return (now - processed) >= thresholdMs;
  });

  if (archiveTargets.length === 0) {
    Logger.log('[archiveOldMailQueue] No records to archive.');
    return;
  }

  archiveTargets.forEach(r => {
    // アーカイブシートへコピー
    appendRow(SHEET_NAMES.MAIL_QUEUE_ARCHIVE, r);
    // MailQueue から削除
    deleteRow(SHEET_NAMES.MAIL_QUEUE, r.id);
  });

  Logger.log('[archiveOldMailQueue] Archived ' + archiveTargets.length + ' record(s).');
}

// ── EXE 死活チェック ──────────────────────────────────────────────

/**
 * EXE の死活を確認し、EXE_DEAD_THRESHOLD_MINUTES（既定 5 分）超過で管理者通知する。
 * LAST_DEAD_NOTIFICATION_AT を参照し、6 時間以内の重複通知を抑制する。
 * 通知後は LAST_DEAD_NOTIFICATION_AT を nowIso_() で更新する。
 * GAS TimeTrigger で 1 分ごとに実行する想定。
 */
function checkExeAlive() {
  const lastHb = getSetting('LAST_HEARTBEAT_TIMESTAMP');
  if (!lastHb) return; // ハートビート未受信（EXE 未起動）

  const thresholdMin = parseFloat(getSetting('EXE_DEAD_THRESHOLD_MINUTES') || '5');
  const elapsedMin   = (Date.now() - new Date(lastHb).getTime()) / 60000;

  if (elapsedMin < thresholdMin) return; // 正常範囲

  // 重複通知抑制: 前回通知から 6 時間以内はスキップ
  const lastNotified = getSetting('LAST_DEAD_NOTIFICATION_AT');
  if (lastNotified) {
    const sinceLastMin = (Date.now() - new Date(lastNotified).getTime()) / 60000;
    if (sinceLastMin < 360) return; // 6 時間 = 360 分
  }

  // 管理者通知
  const hostname = getSetting('LAST_HEARTBEAT_HOSTNAME') || '不明';
  notifyAdminOnFailure(
    '[タスク管理] EXE エージェント応答なし',
    'EXE が ' + Math.floor(elapsedMin) + ' 分応答していません。\n' +
    '最終ハートビート: ' + lastHb + '\n' +
    'PC: ' + hostname
  );

  // 通知時刻を更新（通知失敗時でも更新しない → notifyAdminOnFailure が握りつぶすため
  //  ここで setSetting を呼ぶことで「通知を試みた」記録を残す）
  setSetting('LAST_DEAD_NOTIFICATION_AT', nowIso_());
}

// ── 管理者通知 ────────────────────────────────────────────────────

/**
 * Settings.ADMIN_EMAIL 宛てに管理者通知メールを送信する。
 * ADMIN_EMAIL が空の場合は何もしない。
 * エラーはキャッチしてログのみ（通知失敗で本処理を止めない）。
 *
 * @param {string} subject - メール件名
 * @param {string} body    - メール本文
 */
function notifyAdminOnFailure(subject, body) {
  const adminEmail = getSetting('ADMIN_EMAIL');
  if (!adminEmail) return;

  try {
    GmailApp.sendEmail(adminEmail, subject, body);
  } catch (e) {
    Logger.log('[notifyAdminOnFailure] 管理者通知メール送信失敗: ' + e.message);
  }
}

// ── UI 向けクエリ ──────────────────────────────────────────────────

/**
 * 指定 staffId / reportDate に紐づく MailQueue 最新レコードを返す。
 * UI 側のカード表示（ステータスバッジ・再送ボタン）で使用する。
 * createdAt 降順で最新 1 件を返す。該当無しの場合は null を返す。
 *
 * @param {string} staffId    - Staff.id
 * @param {string} reportDate - YYYY-MM-DD
 * @returns {Object|null}
 */
function getLatestMailRequestFor(staffId, reportDate) {
  var rows = listWhere(SHEET_NAMES.MAIL_QUEUE, function (r) {
    return r.targetStaffId === staffId && r.reportDate === reportDate;
  });
  if (rows.length === 0) return null;
  rows.sort(function (a, b) {
    return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
  });
  return rows[0];
}

/**
 * MailQueue の状態別カウントを返す（ダッシュボード KPI 用）。
 * 大量データ環境を想定し、status と createdAt のみで集計する。
 *
 * @returns {{ pending:number, drafted:number, sent:number, failed:number, picked:number }}
 */
function getMailQueueStats() {
  var stats = { pending: 0, picked: 0, drafted: 0, sent: 0, failed: 0 };
  listAll(SHEET_NAMES.MAIL_QUEUE).forEach(function (r) {
    if (Object.prototype.hasOwnProperty.call(stats, r.status)) {
      stats[r.status]++;
    }
  });
  return stats;
}

// ── ハートビート受信 ──────────────────────────────────────────────

/**
 * EXE からのハートビートを受信し、Settings を更新する。
 *
 * @param {string} hostname - EXE が動作する PC 名
 * @returns {{ ok: true }}
 */
function heartbeat(hostname) {
  setSetting('LAST_HEARTBEAT_TIMESTAMP', nowIso_());
  setSetting('LAST_HEARTBEAT_HOSTNAME',  hostname || '');
  return { ok: true };
}

// ════════════════════════════════════════════════════════════════
// B. 内部ヘルパ（末尾アンダースコア）
// ════════════════════════════════════════════════════════════════

/**
 * MailLog シートにイベント行を追加する。
 *
 * @param {string} mailQueueId - 対象 MailQueue レコードの id
 * @param {string} level       - 'info' | 'warn' | 'error'
 * @param {string} event       - 'picked' | 'drafted' | 'sent' | 'failed' |
 *                               'retry' | 'mismatch' | 'timeout_recover'
 * @param {string} message     - 詳細メッセージ
 */
function logMailEvent_(mailQueueId, level, event, message) {
  const record = {
    id:          generateId_('ml'),
    mailQueueId: mailQueueId,
    timestamp:   nowIso_(),
    level:       level,
    event:       event,
    message:     message || ''
  };
  appendRow(SHEET_NAMES.MAIL_LOG, record);
}

/**
 * picked のまま CAS_TIMEOUT_MINUTES（既定 10 分）以上放置されたレコードを
 * pending に戻す。LockService で排他制御して二重復旧を防ぐ。
 * GAS TimeTrigger で 5 分ごとに実行する想定。
 */
function recoverStalePicked() {
  const timeoutMin = parseFloat(getSetting('CAS_TIMEOUT_MINUTES') || '10');
  const timeoutMs  = timeoutMin * 60 * 1000;
  const now        = Date.now();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const rows = listAll(SHEET_NAMES.MAIL_QUEUE);
    let recovered = 0;

    rows.forEach(r => {
      if (r.status !== 'picked') return;
      const pickedAt = r.pickedAt ? new Date(r.pickedAt).getTime() : 0;
      if ((now - pickedAt) < timeoutMs) return;

      updateRow(SHEET_NAMES.MAIL_QUEUE, r.id, {
        status:   'pending',
        pickedBy: '',
        pickedAt: ''
      });
      logMailEvent_(r.id, 'warn', 'timeout_recover',
        'タイムアウト復旧: pickedBy=' + r.pickedBy + ', pickedAt=' + r.pickedAt);
      recovered++;
    });

    Logger.log('[recoverStalePicked] ' + recovered + ' 件を復旧');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Staff マスタの最新 email と targetEmail を照合する。
 *
 * @param {string} targetStaffId - Staff シートの id
 * @param {string} targetEmail   - MailQueue に保存された targetStaffEmail
 * @returns {boolean} 一致: true、不一致または Staff 未発見: false
 */
function verifyEmailMatch_(targetStaffId, targetEmail) {
  const staff = getById(SHEET_NAMES.STAFF, targetStaffId);
  if (!staff) return false;
  return String(staff.email || '').trim() === String(targetEmail || '').trim();
}
