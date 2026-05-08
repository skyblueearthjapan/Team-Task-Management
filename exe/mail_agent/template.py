"""
template.py - メール件名・本文テンプレート（EXE 内ハードコード）

仕様 v1.0 §7.4.7 を完全準拠。
テンプレート文言を変更した場合は EXE を再ビルドして配布すること。

件名フォーマット:
    【機械設計技術部】{staffName} 業務報告 {reportDate}

本文フォーマット（プレーンテキスト）:
    お疲れ様です。{staffName} です。
    本日の業務をご報告いたします。

    ▼ 本日の作業内容
    {N}. {periodStart}〜{periodEnd}  {kobanCode}  {customer}  {productName}  /  {workType}  {detail}
    ...（複数行）

    ▼ 前日までの作業報告
    {同上フォーマット}

    以上、よろしくお願いいたします。

注記（仕様 §7.4.7）:
    「（一般作業）」の場合は kobanCode / customer / productName が空文字列になる。
    空の場合はそれらのフィールドを出力から省略する。
"""

from typing import Any


# ─── 件名ビルダー ──────────────────────────────────────────────────

def build_subject(subject_vars: dict[str, Any]) -> str:
    """
    件名テンプレートに差し込み変数を展開して返す。

    :param subject_vars: { "staffName": str, "reportDate": str }
                         reportDate は "YYYY/MM/DD" 形式（GAS 側で整形済み）
    :returns: 展開済み件名文字列
    """
    staff_name = subject_vars.get("staffName", "")
    report_date = subject_vars.get("reportDate", "")
    return f"【機械設計技術部】{staff_name} 業務報告 {report_date}"


# ─── 本文ビルダー ──────────────────────────────────────────────────

def _format_items(items: list[dict[str, Any]]) -> str:
    """
    作業アイテムのリストを仕様書フォーマット（連番付き）に変換する。

    フォーマット:
        N. {periodStart}〜{periodEnd}  {kobanCode}  {customer}  {productName}  /  {workType}  {detail}

    kobanCode / customer / productName が空文字列の場合は省略する（§7.4.7 注記）。

    :param items: bodyVars.todayItems または bodyVars.yesterdayItems
    :returns: フォーマット済み文字列（アイテムなし時は「　（記録なし）」）
    """
    if not items:
        return "　（記録なし）"  # 「　（記録なし）」

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


def build_body(body_vars: dict[str, Any]) -> str:
    """
    本文テンプレートに差し込み変数を展開して返す（プレーンテキスト）。

    :param body_vars: {
        "staffName":     str,
        "reportDate":    str,
        "todayItems":    list[dict],
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
