"""
template.py - メール件名・本文テンプレート（EXE 内ハードコード）

仕様 v1.1 §7.4.6 準拠。
- 件名: EXE 側でハードコードテンプレ + 差し込み（変更なし）
- 本文: GAS 側 MailQueueService.buildMailBody() で構築済みの文字列を
        bodyVars.fullBody からそのまま受け取る（v1.1 変更点）

後方互換:
    fullBody が存在しない旧 v1.0 構造が来た場合は _build_body_legacy() で構築する。

テンプレート文言を変更した場合は EXE を再ビルドして配布すること。

件名フォーマット:
    【機械設計技術部】{staffName} 業務報告 {reportDate}
"""

from typing import Any


# ─── 件名ビルダー ──────────────────────────────────────────────────

def build_subject(subject_vars: dict[str, Any]) -> str:
    """
    件名テンプレートに差し込み変数を展開して返す（仕様 §7.4.6）。

    :param subject_vars: { "staffName": str, "reportDate": str }
                         reportDate は "YYYY/MM/DD" 形式（GAS 側で整形済み）
    :returns: 展開済み件名文字列
    """
    staff_name = subject_vars.get("staffName", "")
    report_date = subject_vars.get("reportDate", "")
    return f"【機械設計技術部】{staff_name} 業務報告 {report_date}"


# ─── 本文ビルダー ──────────────────────────────────────────────────

def build_body(body_vars: dict[str, Any]) -> str:
    """
    本文は GAS 側で完成形を構築する設計（v1.1 で変更）。
    body_vars["fullBody"] にそのまま整形済み文字列が入る。

    後方互換: fullBody が無い古い構造の場合は旧ロジックで構築する。

    :param body_vars: {
        "fullBody": str  # v1.1 以降: GAS 側で構築済みの完成本文
        # --- 以下は v1.0 後方互換フォールバック用 ---
        "staffName":      str,
        "todayItems":     list[dict],
        "yesterdayItems": list[dict]
    }
    :returns: 展開済み本文文字列
    """
    full_body = body_vars.get("fullBody")
    if full_body and isinstance(full_body, str) and full_body.strip():
        return full_body

    # 後方互換フォールバック（v1.0 構造）
    return _build_body_legacy(body_vars)


def _build_body_legacy(body_vars: dict[str, Any]) -> str:
    """
    v1.0 互換の本文構築ロジック。
    旧 GAS（fullBody を返さない版）と新 EXE が通信する場合に使用される。
    ※ GAS 側 buildMailBody() とは書式が異なる（署名・greeting・duration 非対応）。

    :param body_vars: {
        "staffName":      str,
        "todayItems":     list[dict],
        "yesterdayItems": list[dict]
    }
    :returns: 展開済み本文文字列
    """
    staff_name = body_vars.get("staffName", "")
    today_items = body_vars.get("todayItems", [])
    yesterday_items = body_vars.get("yesterdayItems", [])

    today_section = _format_items(today_items)
    yesterday_section = _format_items(yesterday_items)

    body = (
        f"お疲れ様です。{staff_name} です。\n"
        f"本日の業務をご報告いたします。\n"
        f"\n"
        f"▼ 本日の作業内容\n"
        f"{today_section}\n"
        f"\n"
        f"▼ 前日までの作業報告\n"
        f"{yesterday_section}\n"
        f"\n"
        f"以上、よろしくお願いいたします。\n"
    )
    return body


def _format_items(items: list[dict[str, Any]]) -> str:
    """
    作業アイテムのリストを仕様書フォーマット（連番付き）に変換する。
    _build_body_legacy() の内部ヘルパー。

    フォーマット:
        N. {periodStart}〜{periodEnd}  {kobanCode}  {customer}  {productName}  /  {workType}  {detail}

    kobanCode / customer / productName が空文字列の場合は省略する（§7.4.7 注記）。

    :param items: bodyVars.todayItems または bodyVars.yesterdayItems
    :returns: フォーマット済み文字列（アイテムなし時は「　（記録なし）」）
    """
    if not items:
        return "　（記録なし）"

    lines: list[str] = []
    for n, item in enumerate(items, start=1):
        period = f"{item.get('periodStart', '')}〜{item.get('periodEnd', '')}"

        # 工番・受注先・品名は空の場合に省略
        optional_parts: list[str] = []
        for key in ("kobanCode", "customer", "productName"):
            val = (item.get(key) or "").strip()
            if val:
                optional_parts.append(val)

        work_type = item.get("workType", "")
        detail = (item.get("detail") or "").strip()

        # 組み立て: 期間 + [工番/受注先/品名...] + "/" + 作業内容 + [詳細]
        parts = [period] + optional_parts + ["/", work_type]
        if detail:
            parts.append(detail)

        lines.append(f"{n}. " + "  ".join(parts))

    return "\n".join(lines)
