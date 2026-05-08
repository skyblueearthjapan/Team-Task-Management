"""
config.py - TeamTaskMail 設定モジュール

環境変数から設定値を読み込む。
必須環境変数が未設定の場合は起動時に sys.exit() する。

環境変数:
    TEAMTASK_WEBAPP_URL  : GAS WebApp デプロイ URL（必須）
    TEAMTASK_API_TOKEN   : EXE 認証トークン（必須）

オプション環境変数:
    TEAMTASK_POLLING_INTERVAL : ポーリング間隔（秒）。既定 30
    TEAMTASK_DEBUG            : "1" で DEBUG ログを有効化
"""

import os
import socket
import sys

# ─── GAS WebApp URL ────────────────────────────────────────────────
WEBAPP_URL: str = os.environ.get(
    "TEAMTASK_WEBAPP_URL",
    ""
)

# ─── EXE 認証トークン ───────────────────────────────────────────────
EXE_API_TOKEN: str = os.environ.get(
    "TEAMTASK_API_TOKEN",
    ""
)

# ─── ポーリング間隔（秒） ──────────────────────────────────────────
POLL_INTERVAL_SECONDS: int = int(
    os.environ.get("TEAMTASK_POLLING_INTERVAL", "30")
)

# ─── ハートビート送信間隔（秒）。仕様 §7.4.4 : 60 秒ごと ────────
HEARTBEAT_INTERVAL_SECONDS: int = 60

# ─── このマシンのホスト名 ──────────────────────────────────────────
HOSTNAME: str = socket.gethostname()

# ─── DEBUG フラグ（CLI 引数でも切り替え可） ───────────────────────
DEBUG: bool = os.environ.get("TEAMTASK_DEBUG", "0") == "1"

# ─── EXE バージョン ────────────────────────────────────────────────
VERSION: str = "1.0.0"


def validate() -> None:
    """
    必須環境変数の存在を検証する。
    いずれか未設定の場合はエラーメッセージを表示して sys.exit(1) する。
    """
    missing = []
    if not WEBAPP_URL:
        missing.append("TEAMTASK_WEBAPP_URL")
    if not EXE_API_TOKEN:
        missing.append("TEAMTASK_API_TOKEN")

    if missing:
        print(
            f"[TeamTaskMail] 起動エラー: 以下の環境変数が設定されていません。\n"
            f"  {', '.join(missing)}\n"
            f"環境変数を設定してから再起動してください。",
            file=sys.stderr,
        )
        sys.exit(1)
