"""
logger.py - TeamTaskMail ロガー設定

ログ出力先:
    %LOCALAPPDATA%\\TeamTaskMail\\logs\\agent.log        （当日分）
    %LOCALAPPDATA%\\TeamTaskMail\\logs\\agent.log.YYYYMMDD （過去分・自動ローテーション）
    コンソール（stdout）

ログファイル名:
    agent.log        : 当日書き込み中のファイル
    agent.log.YYYYMMDD : 日跨ぎ後にローテーションされた過去ログ（最大 30 日保持）

ログフォーマット:
    [%(asctime)s] %(levelname)s [%(module)s] %(message)s

ログレベル:
    既定 INFO。config.DEBUG == True または CLI 引数 --debug で DEBUG に切り替え。
"""

import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

# ─── ログディレクトリ ──────────────────────────────────────────────
_LOCAL_APP_DATA = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
LOG_DIR = Path(_LOCAL_APP_DATA) / "TeamTaskMail" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

# ─── ログファイルベース名（handler が .YYYYMMDD サフィックスを付与） ─
LOG_FILE_BASE = LOG_DIR / "agent.log"

# ─── フォーマット ─────────────────────────────────────────────────
_LOG_FORMAT = "[%(asctime)s] %(levelname)s [%(module)s] %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# ─── ロガー生成 ───────────────────────────────────────────────────
logger: logging.Logger = logging.getLogger("TeamTaskMail")


def setup_logger(debug: bool = False) -> None:
    """
    ロガーにハンドラを設定する。
    main.py の起動時に 1 度だけ呼ぶ。

    :param debug: True にすると DEBUG レベルまで出力する
    """
    level = logging.DEBUG if debug else logging.INFO
    logger.setLevel(level)

    formatter = logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT)

    # ファイルハンドラ（日跨ぎ自動ローテーション）
    file_handler = TimedRotatingFileHandler(
        str(LOG_FILE_BASE),
        when="midnight",
        interval=1,
        backupCount=30,   # 30 日分保持
        encoding="utf-8",
        utc=False,
    )
    file_handler.suffix = "%Y%m%d"  # agent.log.20260508 形式
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)

    # コンソールハンドラ（stdout）
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)

    # 重複追加を防ぐ
    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
