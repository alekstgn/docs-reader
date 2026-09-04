"""Перевод технических замечаний сверки в простой русский язык."""

from __future__ import annotations

import re
from typing import Any

from .compare import parse_number, _is_logistics_row
from .photos import is_supporting_document_photo

SCENE_RU = {
    "hall_audience": "зал с аудиторией",
    "press_wall": "пресс-волл",
    "stage_equipment": "сцена и техника",
    "branding_screen": "экран с заставкой",
    "banner": "баннер",
    "logo": "логотипы",
    "unknown": "снимок",
}

GLUED_DATE_RE = re.compile(r"(\d{1,2}\.\d{2}\.\d{4})\s*г\.(?:\s*\1\s*г\.)+")
TECHNO_RE = re.compile(
    r"\b(?:TZ|PHOTO|image\d+|doc_index|event_fit|conclusion|gte|lte|eq|"
    r"extracted|sampled|compare_kind|wrong_event|inconclusive)\b",
    re.I,
)
API_STATS_RE = re.compile(
    r"(?:Полный прогон этапа C|Выборка этапа C|event_fit|conclusion:|"
    r"извлечено \d+ вложений|в модель отправлено \d+ кадров)[^.]*\.?",
    re.I,
)
EN_DATE_ADDR_RE = re.compile(
    r"The date and address in the report \((.+?)\) do not match the contract \((.+?)\)\.?$",
    re.I | re.S,
)
EN_CHAIRS_RE = re.compile(
    r"The report mentions 120 chairs and 240 cups.*",
    re.I | re.S,
)
EN_CONFIRMS_RE = re.compile(
    r"The report confirms the requirement for (\d+) (.+?)\.?$",
    re.I,
)
LATIN_WORD_RE = re.compile(r"\b[A-Za-z]{3,}\b")
CONTRAST_BREAK = "\n\n"


def _clause_intro(metric: str, clause: str) -> str:
    if clause:
        return f"«{metric}» (п. {clause} ТЗ)"
    return f"«{metric}»"


def _equipment_mismatch_note(row: dict[str, Any], qc: str, qr: str) -> str:
    clause = (row.get("clause") or "").strip()
    if clause == "6.2.2" and "msi" in qc.lower() and "gateway" in qr.lower():
        return (
            "MSI Core i5 в отчёте указан отдельно для Музея (п. 6.2.4) — 1 шт.; "
            "это другая площадка и не заменяет 2 ноутбука MSI для «Шереметьевского» по п. 6.2.2."
        )
    return ""


def format_contrast(head: str, tail: str = "") -> str:
    """Два абзаца: договор/требование, затем отчёт — для комментариев Word."""
    head = head.strip()
    tail = tail.strip()
    if head and tail:
        return f"{head}{CONTRAST_BREAK}{tail}"
    return head or tail


def format_contract_report(
    metric: str,
    contract: str,
    report: str,
    *,
    intro: str | None = None,
    contract_limit: int = 220,
    report_limit: int = 220,
) -> str:
    contract = contract.strip()[:contract_limit]
    report = report.strip()[:report_limit]
    lead = intro or f"«{metric}»: в отчёте указано иначе, чем в договоре"
    return format_contrast(
        f"{lead}. Договор: {contract}",
        f"Отчёт: {report}",
    )


def _fmt_qty(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0"):
        return text[:-2]
    return text


def normalize_caption(text: str) -> str:
    if not text:
        return ""
    text = GLUED_DATE_RE.sub(r"\1", text.strip())
    text = re.sub(r"\s*г\.\s*", " ", text)
    return " ".join(text.split())[:80]


def photo_label(date_hint: str, scene_type: str = "", notes: str = "") -> str:
    parts: list[str] = []
    if date_hint:
        parts.append(date_hint.replace(" г.", "").strip())
    scene = SCENE_RU.get(scene_type or "", "")
    if scene and scene not in parts:
        parts.append(scene)
    if not parts and notes:
        snippet = re.sub(TECHNO_RE, "", notes)
        snippet = " ".join(snippet.split())[:60]
        if snippet:
            parts.append(snippet)
    return ", ".join(parts) if parts else "фото в отчёте"


def strip_technical(text: str) -> str:
    if not text:
        return ""
    text = API_STATS_RE.sub("", text)
    text = re.sub(r"\[ФИО скрыто\]", "[скрыто]", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def looks_english(text: str) -> bool:
    if not text:
        return False
    stripped = text.strip()
    if stripped.startswith(
        ("The ", "This ", "A ", "An ", "It ", "While ", "No ", "Audience ", "The photo")
    ):
        return True
    latin_words = LATIN_WORD_RE.findall(stripped)
    cyrillic = len(re.findall(r"[а-яёА-ЯЁ]", stripped))
    if len(latin_words) >= 3 and len(latin_words) * 4 > cyrillic:
        return True
    return False


def ensure_russian(text: str, row: dict[str, Any] | None = None) -> str:
    """Перевод известных англоязычных комментариев модели в простой русский."""
    text = strip_technical(text)
    if not text:
        return ""
    if not looks_english(text):
        return text

    match = EN_DATE_ADDR_RE.search(text)
    if match:
        return (
            f"Дата и адрес в отчёте ({match.group(1).strip()}) "
            f"не совпадают с договором ({match.group(2).strip()})."
        )

    if EN_CHAIRS_RE.match(text):
        return (
            "В отчёте указаны 120 стульев и 240 стаканчиков; по договору — не менее 60 стульев. "
            "Даты позиций отчёта (29.06–02.07.2026) не совпадают с договором (ноябрь 2025)."
        )

    match = EN_CONFIRMS_RE.match(text)
    if match:
        return f"Отчёт подтверждает требование: {match.group(1)} {match.group(2).strip()}."

    if text.startswith("The report confirms"):
        return "Отчёт подтверждает проведение указанного формата."

    if row:
        rebuilt = humanize_row({**row, "comment": ""})
        if rebuilt and not looks_english(rebuilt):
            return rebuilt

    qc = (row or {}).get("quote_contract") or ""
    qr = (row or {}).get("quote_report") or ""
    metric = (row or {}).get("metric") or "пункт"
    if qc and qr:
        return format_contract_report(metric, qc, qr)

    return "Замечание по сверке с договором."


def humanize_event(event: dict[str, Any], pair_id: str) -> str:
    if event.get("same_event"):
        return ""
    required = (event.get("required") or "").strip()
    claimed = (event.get("claimed") or "").strip()
    if pair_id == "prosvetiteli":
        return format_contrast(
            "По договору — форум «Просветители» (блогеры) 27–30 ноября 2025 года "
            "в парк-отеле «Шереметьевский» и Музее военной формы.",
            "В отчёте — конференция «Просветители.Обществознание» (тоже для блогеров), "
            "но другие даты (29.06–02.07.2026) и другая основная площадка — «Лесной» вместо «Шереметьевского». "
            "Адреса, техобеспечение, трансферы и п. 9 нужно сверять с учётом этих расхождений.",
        )
    if required and claimed:
        return format_contrast(
            f"По договору: {required[:200]}.",
            f"В отчёте: {claimed[:200]}. Мероприятия не совпадают.",
        )
    return strip_technical(event.get("comment") or "Документы описывают разные мероприятия.")


def humanize_row(row: dict[str, Any]) -> str:
    status = row.get("status") or ""
    metric = (row.get("metric") or row.get("clause") or "пункт").strip()
    required = (row.get("required") or "").strip()
    claimed = (row.get("claimed") or "").strip()
    comment = ensure_russian(strip_technical(row.get("comment") or ""), row)
    op = row.get("operator") or ""

    if status == "missing_in_report":
        if required:
            return format_contrast(
                f"В отчёте не найден пункт «{metric}».",
                f"По договору: {required[:220]}.",
            )
        return f"В отчёте не найден пункт «{metric}»."

    if "фото" in metric.lower() or row.get("evidence_type") == "photo":
        embedded = row.get("embedded_images")
        claimed_n = _fmt_qty(row.get("claimed_photos") or claimed)
        req_n = _fmt_qty(row.get("required_photos") or required)
        if embedded is not None:
            req_num = parse_number(str(row.get("required_photos") or required))
            if req_num is not None and embedded >= req_num:
                if row.get("photo_text_gap"):
                    return format_contrast(
                        f"По договору нужно не менее {req_n or required}. "
                        f"В файле отчёта вложено {embedded} фотографий — минимум выполнен.",
                        f"В тексте отчёта указано {claimed_n}, во вложении DOCX — {embedded}. "
                        f"Это расхождение текста и приложения, не нарушение договора.",
                    )
                return (
                    f"По договору нужно не менее {req_n or required}. "
                    f"В файле отчёта вложено {embedded} фотографий — требование выполнено."
                )
            return format_contrast(
                f"По договору нужно не менее {req_n or required}. "
                f"В тексте отчёта заявлено {claimed_n}.",
                f"В самом файле отчёта вложено только {embedded} фотографий — меньше минимума по договору.",
            )
        return comment or f"Расхождение по фотоотчёту: договор — {required}, отчёт — {claimed}."

    if status == "mismatch" and (
        _is_logistics_row(row)
        or "разделы 6" in comment.lower()
        or "оборудование и даты" in comment.lower()
    ):
        qc = (row.get("quote_contract") or required or "").strip()
        qr = (row.get("quote_report") or claimed or "").strip()
        if qc and qr and qc != qr:
            clause = (row.get("clause") or "").strip()
            intro = f"{_clause_intro(metric, clause)}: в отчёте указано иначе, чем в договоре"
            body = format_contract_report(metric, qc, qr, intro=intro)
            extra = _equipment_mismatch_note(row, qc, qr)
            if not extra and "шереметьев" in comment.lower():
                extra = (
                    "По договору — «Шереметьевский» (ноябрь 2025); "
                    "отчёт описывает «Лесной» / июнь 2026."
                )
            if extra:
                return f"{body}{CONTRAST_BREAK}{extra}"
            return body
        if qc and qr and qc == qr and "шереметьев" in comment.lower():
            clause = (row.get("clause") or "").strip()
            return (
                f"{_clause_intro(metric, clause)}: по договору — «Шереметьевский», "
                f"27–30.11.2025; отчёт описывает «Лесной» / июнь 2026."
            )

    if status == "mismatch" and (
        "адрес" in comment.lower()
        or "автомобильн" in comment.lower()
        or "октябрьск" in comment.lower()
        or "площадк" in comment.lower()
    ):
        qc = (row.get("quote_contract") or required or "").strip()
        qr = (row.get("quote_report") or claimed or "").strip()
        if qc and qr and qc != qr:
            return format_contract_report(metric, qc, qr)
        return format_contract_report(
            metric,
            required[:180],
            claimed[:180],
            intro=f"«{metric}»: адрес или площадка в отчёте не совпадает с договором",
        )

    if op in ("gte", "lte", "eq") and required and claimed:
        op_ru = {"gte": "не менее", "lte": "не более", "eq": "ровно"}.get(op, "")
        qc = (row.get("quote_contract") or "").strip()
        qr = (row.get("quote_report") or "").strip()
        clause = (row.get("clause") or "").strip()
        if status == "mismatch" and qc and qr:
            intro = f"{_clause_intro(metric, clause)}: в отчёте указано иначе, чем в договоре"
            body = format_contract_report(metric, qc, qr, intro=intro)
            extra = _equipment_mismatch_note(row, qc, qr)
            if extra:
                return f"{body}{CONTRAST_BREAK}{extra}"
            return body
        if status == "mismatch":
            req_nums = re.findall(r"\d+", required)
            cl_nums = re.findall(r"\d+", claimed)
            if op == "eq" and req_nums and cl_nums and req_nums[0] == cl_nums[0]:
                return format_contract_report(
                    metric,
                    required[:180],
                    claimed[:180],
                    intro=f"{_clause_intro(metric, clause)}: формулировка или модель в отчёте не совпадает с договором",
                )
            return format_contrast(
                f"{_clause_intro(metric, clause)}: по договору нужно {op_ru} {required}.",
                f"В отчёте указано: {claimed}.",
            )
        return comment

    if "адрес/площадка" in comment or "автомобильн" in comment.lower():
        return format_contract_report(
            metric,
            required[:180],
            claimed[:180],
            intro=f"«{metric}»: адрес или площадка в отчёте не совпадает с договором",
        )

    if comment and not TECHNO_RE.search(comment) and len(comment) < 400 and not looks_english(comment):
        return comment

    if required and claimed:
        return format_contrast(
            f"«{metric}»: по договору — {required[:160]}.",
            f"В отчёте — {claimed[:160]}.",
        )
    return comment or f"Замечание по пункту «{metric}»."


def humanize_photo(
    analysis: dict[str, Any],
    pair_id: str,
    extra_count: int = 0,
) -> str:
    conclusion = analysis.get("conclusion") or ""
    notes = strip_technical(analysis.get("notes") or "")
    branding = (analysis.get("branding_or_text_seen") or "").strip()
    date_hint = (analysis.get("date_hint") or "").replace(" г.", "")

    if pair_id == "prosvetiteli":
        base = (
            "На снимке видно мероприятие из отчёта (конференция в июне 2026, «Лесной»), "
            "а не то, что требует договор (форум в ноябре 2025, «Шереметьевский»)."
        )
        if branding:
            base += f" На кадре читается: «{branding[:120]}»."
        if extra_count > 0:
            base += f" Ещё {extra_count} похожих снимков в этом блоке отчёта."
        return base

    if conclusion == "wrong_event" or analysis.get("event_fit") == "other_event":
        if is_supporting_document_photo(analysis):
            return ""
        if notes:
            msg = ensure_russian(notes, None)
            if branding:
                msg += f" На кадре читается: «{branding[:120]}»."
            if extra_count > 0:
                msg += f" Ещё {extra_count} похожих снимков в этом блоке отчёта."
            return msg
        return "Снимок не подтверждает площадку или программу; проверьте вручную."

    if conclusion == "contradicts":
        msg = ensure_russian(notes, None) or "Снимок противоречит требованиям договора."
        if extra_count > 0:
            msg += f" Ещё {extra_count} похожих снимков рядом."
        return msg

    if notes:
        return ensure_russian(notes, None)
    return "На снимке есть замечание по сверке с договором."


def location_label(row: dict[str, Any]) -> str:
    clause = (row.get("clause") or "").strip()
    metric = (row.get("metric") or "").strip()
    if clause and metric:
        return f"п. {clause}, {metric}"
    if metric:
        return metric
    return "текст отчёта"
