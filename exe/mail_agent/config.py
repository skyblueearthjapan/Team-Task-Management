"""
config.py - TeamTaskMail 設定モジュール

設定値の読み込み優先順位:
  1. ビルド時埋め込み (_baked_config.py) — build.bat で生成
  2. 環境変数 (TEAMTASK_WEBAPP_URL / TEAMTASK_API_TOKEN)

ビルド時に値が埋め込まれた EXE は、配布先 PC で環境変数を設定せずに
そのまま実行できる。開発時に Python で直接動かす場合は環境変数を使う。

環境変数（埋め込みがない場合のみ参照）:
    TEAMTASK_WEBAPP_URL  : GAS WebApp デプロイ URL（必須）
    TEAMTASK_API_TOKEN   : EXE 認証トークン（必須）

オプション環境変数（埋め込み有無に関わらず常に参照）:
    TEAMTASK_POLLING_INTERVAL : ポーリング間隔（秒）。既定 30
    TEAMTASK_DEBUG            : "1" で DEBUG ログを有効化
"""

import os
import socket
import sys

# ─── ビルド時埋め込み設定の読み込み（存在すれば優先） ──────────────
# _baked_config.py は build.bat が生成し、PyInstaller でバイナリへ取り込まれる。
# Git には含めない（機密のため .gitignore で除外）。
try:
    from . import _baked_config as _baked  # type: ignore
    _BAKED_WEBAPP_URL: str = getattr(_baked, "WEBAPP_URL", "") or ""
    _BAKED_API_TOKEN: str = getattr(_baked, "API_TOKEN", "") or ""
except ImportError:
    _BAKED_WEBAPP_URL = ""
    _BAKED_API_TOKEN = ""

# ─── GAS WebApp URL（埋め込み優先、無ければ環境変数） ─────────────
WEBAPP_URL: str = _BAKED_WEBAPP_URL or os.environ.get(
    "TEAMTASK_WEBAPP_URL",
    ""
)

# ─── EXE 認証トークン（埋め込み優先、無ければ環境変数） ──────────
EXE_API_TOKEN: str = _BAKED_API_TOKEN or os.environ.get(
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
