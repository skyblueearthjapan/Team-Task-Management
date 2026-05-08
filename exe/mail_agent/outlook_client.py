"""
outlook_client.py - Outlook COM クライアント

win32com.client 経由で Outlook を操作する。
Windows + Outlook インストール済み環境でのみ動作する。

公開関数:
    create_draft(to_addresses, subject, body)  : 下書き保存（mode=draft）
    send_directly(to_addresses, subject, body) : 直接送信（mode=send）
    get_default_email_address()                : 既定プロファイルのメールアドレス取得
"""

import win32com.client  # pywin32

from .logger import logger


def _get_outlook():
    """
    Outlook.Application COM オブジェクトを取得する。
    Outlook が起動していない場合も COM が自動起動を試みる。

    :raises RuntimeError: COM 接続に失敗した場合
    """
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        raise RuntimeError(
            f"Outlook への COM 接続に失敗しました。"
            f" Outlook がインストールされ、プロファイルが設定済みであることを確認してください。"
            f" 詳細: {exc}"
        ) from exc


def create_draft(to_addresses: list[str], subject: str, body: str) -> bool:
    """
    Outlook の下書きフォルダにメールを保存する（送信しない）。

    仕様 §7.4.3 mode='draft' 対応。

    :param to_addresses: 宛先メールアドレスのリスト
    :param subject:      件名
    :param body:         本文（プレーンテキスト）
    :returns: 成功なら True
    :raises:  COM エラー時は例外を再送出する
    """
    try:
        outlook = _get_outlook()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = "; ".join(to_addresses)
        mail.Subject = subject
        mail.Body = body
        mail.Save()  # 下書きフォルダへ保存
        logger.info(f"下書き作成成功: subject='{subject}' to={to_addresses}")
        return True
    except Exception as exc:
        logger.error(f"下書き作成失敗: subject='{subject}' error={exc}")
        raise


def send_directly(to_addresses: list[str], subject: str, body: str) -> bool:
    """
    Outlook でメールを直接送信する。

    仕様 §7.4.3 mode='send' 対応。
    呼び出し前に呼び出し元（main.py の process_item）が
    get_default_email_address() でプロファイル一致確認を行うこと（仕様 §7.4.2 ④）。

    :param to_addresses: 宛先メールアドレスのリスト
    :param subject:      件名
    :param body:         本文（プレーンテキスト）
    :returns: 成功なら True
    :raises:  COM エラー時は例外を再送出する
    """
    try:
        outlook = _get_outlook()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = "; ".join(to_addresses)
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        logger.info(f"送信成功: subject='{subject}' to={to_addresses}")
        return True
    except Exception as exc:
        logger.error(f"送信失敗: subject='{subject}' error={exc}")
        raise


def get_default_email_address() -> str:
    """
    現在 Outlook にサインインしている既定プロファイルのメールアドレスを返す。

    仕様 §7.4.2 ④ のプロファイル一致確認で使用する。
    複数アカウントがある場合はデフォルトアカウントを返す。

    :returns: メールアドレス文字列（小文字正規化済み）
    :raises RuntimeError: アカウント情報が取得できない場合
    """
    try:
        outlook = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        # Accounts コレクションの最初のアカウントが既定プロファイル
        accounts = namespace.Accounts
        if accounts.Count == 0:
            raise RuntimeError("Outlook にアカウントが登録されていません。")
        # COM コレクションは 1-indexed
        default_account = accounts.Item(1)
        address = default_account.SmtpAddress or ""
        logger.debug(f"既定メールアドレス取得: {address}")
        return address.lower()
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(
            f"Outlook の既定メールアドレス取得に失敗しました。詳細: {exc}"
        ) from exc
