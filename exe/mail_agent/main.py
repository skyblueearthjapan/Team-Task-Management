"""
main.py - TeamTaskMail EXE エージェント エントリポイント

起動方法（開発時）:
    cd "C:\\Users\\<user>\\Documents\\Team Task Management\\exe"
    python -m mail_agent.main

    オプション:
        --debug   DEBUG レベルのログを出力する

本番（PyInstaller ビルド後）:
    dist\\v1.0.0\\TeamTaskMail.exe

メインループの挙動（仕様 §8.1）:
    - 30 秒ごとに pickMailItem をポーリング
    - 60 秒ごとに heartbeat を送信
    - KeyboardInterrupt / SIGTERM でグレースフルシャットダウン
    - 例外発生時はログ出力後にループ継続（クラッシュしない）
"""

import argparse
import json
import signal
import sys
import time
from typing import Optional

# ─── 相対インポート（パッケージ実行時）、直接実行時の互換処理 ──────
try:
    from .config import (
        EXE_API_TOKEN,
        HEARTBEAT_INTERVAL_SECONDS,
        HOSTNAME,
        POLL_INTERVAL_SECONDS,
        VERSION,
        validate,
    )
    from .logger import logger, setup_logger
    from .outlook_client import create_draft, get_default_email_address, send_directly
    from .poller import complete_mail_item, heartbeat, pick_mail_item
    from .template import build_body, build_subject
except ImportError:
    # PyInstaller でバンドルされた場合や sys.path が通っている場合のフォールバック
    from config import (  # type: ignore[no-redef]
        EXE_API_TOKEN,
        HEARTBEAT_INTERVAL_SECONDS,
        HOSTNAME,
        POLL_INTERVAL_SECONDS,
        VERSION,
        validate,
    )
    from logger import logger, setup_logger  # type: ignore[no-redef]
    from outlook_client import create_draft, get_default_email_address, send_directly  # type: ignore[no-redef]
    from poller import complete_mail_item, heartbeat, pick_mail_item  # type: ignore[no-redef]
    from template import build_body, build_subject  # type: ignore[no-redef]


# ─── グレースフルシャットダウン用フラグ ───────────────────────────
_shutdown_requested = False


def _handle_signal(signum, frame) -> None:  # noqa: ANN001
    """SIGTERM / SIGINT ハンドラ。ループを止めるフラグを立てる。"""
    global _shutdown_requested
    logger.info(f"シャットダウン要求を受信しました (signal={signum})。次のループ後に停止します。")
    _shutdown_requested = True


# ─── アイテム処理 ─────────────────────────────────────────────────

def process_item(item: dict) -> None:
    """
    MailQueue の 1 レコードを処理する。

    フロー（仕様 §8.1）:
        1. toAddresses（カンマ区切り文字列）を split でリスト化
        2. subjectVars / bodyVars（JSON 文字列）を json.loads でパース
        3. EXE 内ハードコードテンプレートで件名・本文を生成
        4. mode='draft' → create_draft → completeMailItem('drafted')
           mode='send'  → プロファイル一致確認 → send_directly → completeMailItem('sent')
           不一致       → completeMailItem('failed', 'outlook_profile_mismatch')
        5. 例外時       → completeMailItem('failed', エラーメッセージ)

    :param item: MailQueue レコード dict
    """
    item_id: str = item.get("id", "")
    mode: str = item.get("mode", "draft")
    logger.info(f"処理開始: id={item_id} mode={mode}")

    # GAS pickMailItem が §7.4.2③ email_mismatch で failed にしたレコードはスキップ
    if item.get("status") == "failed":
        logger.warning(
            f"GAS 側で既に failed: id={item_id} "
            f"errorMessage={item.get('errorMessage', '')}"
        )
        return  # GAS 側で既に状態確定しているので何もしない

    try:
        # 宛先の展開（仕様 §6.3: toAddresses はカンマ区切り string）
        raw_to: str = item.get("toAddresses", "") or ""
        to_addresses: list[str] = [a.strip() for a in raw_to.split(",") if a.strip()]

        if not to_addresses:
            raise ValueError(f"toAddresses が空です: '{raw_to}'")

        # subjectVars / bodyVars は JSON 文字列
        subject_vars: dict = json.loads(item.get("subjectVars") or "{}")
        body_vars: dict = json.loads(item.get("bodyVars") or "{}")

        subject: str = build_subject(subject_vars)
        body: str = build_body(body_vars)

        if mode == "draft":
            create_draft(to_addresses, subject, body)
            complete_mail_item(item_id, "drafted")
            logger.info(f"下書き作成完了: id={item_id}")

        elif mode == "send":
            # ④ 直接送信前のプロファイル一致確認（仕様 §7.4.2 ④）
            target_email: str = (item.get("targetStaffEmail") or "").lower().strip()
            current_email: str = get_default_email_address()

            if current_email != target_email:
                logger.warning(
                    f"Outlook プロファイル不一致: "
                    f"現在のプロファイル={current_email!r} "
                    f"対象スタッフ={target_email!r} "
                    f"id={item_id}"
                )
                complete_mail_item(item_id, "failed", "outlook_profile_mismatch")
                return

            send_directly(to_addresses, subject, body)
            complete_mail_item(item_id, "sent")
            logger.info(f"送信完了: id={item_id}")

        else:
            error_msg = f"unknown mode: {mode}"
            logger.error(f"不明なモード: {error_msg} id={item_id}")
            complete_mail_item(item_id, "failed", error_msg)

    except Exception as exc:  # noqa: BLE001
        error_msg = str(exc)[:500]
        logger.exception(f"処理失敗: id={item_id} error={error_msg}")
        # completeMailItem で GAS 側に通知（failed に遷移）
        try:
            complete_mail_item(item_id, "failed", error_msg)
        except Exception as notify_exc:  # noqa: BLE001
            logger.error(f"completeMailItem 通知にも失敗しました: {notify_exc}")


# ─── メインループ ─────────────────────────────────────────────────

def main_loop() -> None:
    """
    30 秒ごとに MailQueue をポーリングし、60 秒ごとにハートビートを送信する。
    例外が発生してもループを継続する（クラッシュしない設計）。
    """
    logger.info(f"TeamTaskMail v{VERSION} 起動 hostname={HOSTNAME}")
    last_heartbeat_at: float = 0.0

    while not _shutdown_requested:
        loop_start = time.time()

        try:
            # ─ ハートビート（60 秒ごと） ───────────────────────────
            if loop_start - last_heartbeat_at >= HEARTBEAT_INTERVAL_SECONDS:
                if heartbeat(HOSTNAME):
                    logger.debug("ハートビート送信 OK")
                else:
                    logger.warning("ハートビート送信失敗（次回ループで再試行）")
                last_heartbeat_at = loop_start

            # ─ MailQueue ポーリング ────────────────────────────────
            item: Optional[dict] = pick_mail_item(HOSTNAME)

            if item is not None:
                process_item(item)
            else:
                logger.debug("pending なし。次のポーリングまで待機します。")

        except Exception as exc:  # noqa: BLE001
            # メインループ自体の例外は記録してループ継続（クラッシュしない）
            logger.exception(f"メインループで予期せぬ例外が発生しました: {exc}")

        # ─ 次のポーリングまで待機 ─────────────────────────────────
        elapsed = time.time() - loop_start
        remaining = max(0.0, POLL_INTERVAL_SECONDS - elapsed)
        while remaining > 0 and not _shutdown_requested:
            time.sleep(min(1.0, remaining))
            remaining -= 1.0

    logger.info("TeamTaskMail を正常終了しました。")


# ─── エントリポイント ─────────────────────────────────────────────

def main() -> None:
    """
    コマンドライン引数をパースし、ロガーをセットアップしてメインループを起動する。
    """
    parser = argparse.ArgumentParser(
        prog="TeamTaskMail",
        description="機械設計技術部 タスク管理 メールエージェント",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="DEBUG レベルのログを出力する",
    )
    args = parser.parse_args()

    # ロガー初期化（config.DEBUG または --debug フラグ）
    try:
        from .config import DEBUG as CONFIG_DEBUG  # type: ignore[assignment]
    except ImportError:
        from config import DEBUG as CONFIG_DEBUG  # type: ignore[no-redef,assignment]

    setup_logger(debug=args.debug or CONFIG_DEBUG)

    # 必須環境変数の検証（未設定なら sys.exit）
    validate()

    # シグナルハンドラ登録（SIGTERM / SIGINT）
    signal.signal(signal.SIGTERM, _handle_signal)
    signal.signal(signal.SIGINT, _handle_signal)

    try:
        main_loop()
    except KeyboardInterrupt:
        logger.info("KeyboardInterrupt を受信しました。終了します。")
        sys.exit(0)


if __name__ == "__main__":
    main()
