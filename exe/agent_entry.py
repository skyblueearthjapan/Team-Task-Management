"""
agent_entry.py - PyInstaller のエントリポイント。

このスクリプトは、PyInstaller でビルドされた EXE が最初に実行する
ブートストラップである。`mail_agent.main:main()` を呼び出すだけの
薄いラッパー。

このラッパー方式を採用する理由:
    PyInstaller が `mail_agent\\main.py` を直接エントリにすると、
    フリーズ後に main.py が `__main__` として実行され、
    `from .config import ...` 等の相対インポートが
    "no known parent package" で失敗する。

    本ラッパーを通すと main.py は `mail_agent.main` モジュールとして
    ロードされ、相対インポートが正常に解決される。

開発時の起動（このラッパー経由でも、直接 module 実行でも可）:
    python agent_entry.py
    python -m mail_agent.main
"""
from mail_agent.main import main

if __name__ == "__main__":
    main()
