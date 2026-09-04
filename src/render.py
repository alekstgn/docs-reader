"""Человекочитаемый отчёт о расхождениях."""

from __future__ import annotations

from typing import Any

STATUS_RU = {
    "ok": "совпадает",
    "mismatch": "расхождение",
    "missing_in_report": "нет в отчёте",
    "needs_photo": "нужно фото",
    "needs_photo_skipped": "фото отложены",
    "photo_inconclusive": "фото неубедительны",
    "not_verifiable_from_docs": "не проверить по документам",
    "event_mismatch": "другое мероприятие",
    "qualitative_pending": "ожидает качественной сверки",
}


def render_markdown(pair_id: str, result: dict[str, Any]) -> str:
    usage = result.get("usage") or {}
    event_row = result["event_check"]
    rows: list[dict[str, Any]] = result["rows"]
    lines = [
        f"# Сверка: {pair_id}",
        "",
        f"- Договор: `{result['contract_file']}`",
        f"- Отчёт: `{result['report_file']}`",
        f"- Вложений в отчёте: {result.get('report_images', 0)} шт., "
        f"{result.get('report_media_mb', 0):.1f} МБ",
        f"- Этап C (фото): {result.get('photo_stage', 'skipped')}",
        f"- Токены: in={usage.get('prompt_tokens', 0)}, out={usage.get('output_tokens', 0)}, "
        f"стоимость ≈ {usage.get('cost_rub', 0)} ₽",
        "",
        "## Идентичность мероприятия",
        "",
        f"**Статус:** {STATUS_RU.get(event_row['status'], event_row['status'])}",
        "",
        f"- Договор: {event_row.get('required')}",
        f"- Отчёт: {event_row.get('claimed')}",
        f"- Комментарий: {event_row.get('comment')}",
        "",
        "## Таблица расхождений",
        "",
        "| ID | Пункт | Статус | Договор | Отчёт | Комментарий |",
        "|---|---|---|---|---|---|",
    ]
    ordered = [event_row] + rows
    for row in ordered:
        lines.append(
            "| {id} | {clause} | {status} | {req} | {cl} | {cmt} |".format(
                id=_cell(row.get("id")),
                clause=_cell(row.get("clause")),
                status=STATUS_RU.get(row.get("status"), row.get("status")),
                req=_cell(row.get("required"), 120),
                cl=_cell(row.get("claimed"), 120),
                cmt=_cell(row.get("comment"), 160),
            )
        )
    lines += ["", "## Цитаты", ""]
    for row in ordered:
        if row.get("status") in {"ok"}:
            continue
        lines.append(f"### {row.get('id')} — {row.get('clause')}")
        lines.append("")
        lines.append(f"- Договор: «{_cell(row.get('quote_contract'), 400)}»")
        lines.append(f"- Отчёт: «{_cell(row.get('quote_report'), 400)}»")
        lines.append("")
    return "\n".join(lines) + "\n"


def _cell(value: Any, limit: int = 80) -> str:
    text = "" if value is None else str(value).replace("|", "/").replace("\n", " ")
    text = " ".join(text.split())
    if len(text) > limit:
        return text[: limit - 1] + "…"
    return text
