"""
poller.py - GAS WebApp ポーリングクライアント

仕様 §7.4.4 準拠:
    - 全リクエストに ?token=EXE_API_TOKEN クエリパラメータを付与
    - action ベースの RPC スタイル（POST /exec 単一エンドポイント）
    - タイムアウト 30 秒、指数バックオフ付きリトライ 3 回（1s / 2s / 4s）

公開関数:
    pick_mail_item(host)        : pending を 1 件取得（CAS）
    complete_mail_item(...)     : 処理結果を通知
    heartbeat(host)             : 死活通知
"""

import time
from typing import Optional

import requests

from .config import EXE_API_TOKEN, HOSTNAME, WEBAPP_URL
from .logger import logger

# ─── 定数 ─────────────────────────────────────────────────────────
_TIMEOUT_SECONDS = 30
_MAX_RETRIES = 3
_BACKOFF_BASE = 1  # 初回待機秒数。次は x2、x4


# ─── 内部ユーティリティ ────────────────────────────────────────────

def _post(payload: dict) -> Optional[dict]:
    """
    GAS WebApp の doPost に JSON でリクエストする。
    失敗時は指数バックオフで最大 _MAX_RETRIES 回リトライする。

    :param payload: リクエストボディ辞書（action を含むこと）
    :returns: レスポンス JSON の dict、または None（全リトライ失敗時）
    """
    action = payload.get("action", "unknown")

    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            resp = requests.post(
                WEBAPP_URL,
                params={"token": EXE_API_TOKEN},
                json=payload,
                timeout=_TIMEOUT_SECONDS,
            )
            resp.raise_for_status()
            data = resp.json()

            # GAS は HTTP 200 でもエラーを返す場合がある（§7.4.4 注記）
            if isinstance(data, dict) and data.get("status") == 401:
                logger.error(
                    f"[{action}] 認証エラー (status=401): トークンが不正です。"
                    " 環境変数 TEAMTASK_API_TOKEN を確認してください。"
                )
                return None

            return data

        except requests.Timeout:
            logger.warning(
                f"[{action}] タイムアウト ({_TIMEOUT_SECONDS}s)"
                f" attempt={attempt}/{_MAX_RETRIES}"
            )
        except requests.HTTPError as exc:
            if exc.response is not None and exc.response.status_code == 401:
                logger.error(f"[{action}] HTTP 401 Unauthorized: リトライせず終了")
                return None
            status_code = exc.response.status_code if exc.response is not None else "?"
            logger.error(
                f"[{action}] HTTP エラー: {status_code}"
                f" attempt={attempt}/{_MAX_RETRIES}"
            )
        except requests.ConnectionError as exc:
            logger.error(
                f"[{action}] 接続エラー: {exc}"
                f" attempt={attempt}/{_MAX_RETRIES}"
            )
        except Exception as exc:  # noqa: BLE001
            logger.error(
                f"[{action}] 予期せぬエラー: {exc}"
                f" attempt={attempt}/{_MAX_RETRIES}"
            )

        # 最終試行でなければバックオフ待機
        if attempt < _MAX_RETRIES:
            wait = _BACKOFF_BASE * (2 ** (attempt - 1))  # 1s, 2s, 4s
            logger.debug(f"[{action}] {wait}s 後にリトライします")
            time.sleep(wait)

    logger.warning(f"[{action}] {_MAX_RETRIES} 回のリトライがすべて失敗しました。次のループで再試行します。")
    return None


# ─── 公開 API ─────────────────────────────────────────────────────

def pick_mail_item(host: str = HOSTNAME) -> Optional[dict]:
    """
    GAS doPost に action='pickMailItem' を POST し、
    pending な MailQueue レコードを 1 件取得する（atomic CAS）。

    :param host: EXE 識別子として使用するホスト名
    :returns: MailQueue レコード dict、または None（pending なし or 失敗）
    """
    result = _post({"action": "pickMailItem", "hostname": host})

    # GAS は pending なしのとき null（JSON null）を返す → Python では None
    if result is None:
        return None

    # 正常なレコードかどうか確認
    if isinstance(result, dict) and result.get("id"):
        return result

    # null 以外の非レコード（例: {"status": 200, "message": "..."}）
    logger.debug(f"[pickMailItem] pending なし または非レコード応答: {result}")
    return None


def complete_mail_item(
    item_id: str,
    status: str,
    error_message: Optional[str] = None,
) -> bool:
    """
    GAS doPost に action='completeMailItem' を POST し、
    MailQueue レコードの処理結果を通知する。

    :param item_id:       MailQueue.id
    :param status:        'drafted' | 'sent' | 'failed'
    :param error_message: 失敗時の理由（500 文字以内推奨）
    :returns: 通知成功なら True
    """
    payload: dict = {
        "action": "completeMailItem",
        "hostname": HOSTNAME,
        "id": item_id,
        "status": status,
        "errorMessage": (error_message or "")[:500],
    }
    result = _post(payload)
    return result is not None


def heartbeat(host: str = HOSTNAME) -> bool:
    """
    GAS doPost に action='heartbeat' を POST し、
    Settings.LAST_HEARTBEAT_TIMESTAMP を更新する（死活監視）。

    :param host: EXE 識別子として使用するホスト名
    :returns: 送信成功なら True
    """
    result = _post({"action": "heartbeat", "hostname": host})
    return result is not None and result.get("ok") is True
