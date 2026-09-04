"""Страховка: ключевые количественные пункты ТЗ, если модель их пропустила."""

from __future__ import annotations

import re
from typing import Any

_LECTURES = re.compile(
    r"не менее\s+(\d+)\s*\([^)]*\)\s*историко-просветительских лекций",
    re.IGNORECASE,
)
_BACKDROPS = re.compile(
    r"не менее\s+(\d+)\s*\([^)]*\)\s*статичных заставок",
    re.IGNORECASE,
)
_PHOTOS = re.compile(
    r"не менее\s+(\d+)\s*\([^)]*\)\s*фотографий",
    re.IGNORECASE,
)
_PARTICIPANTS = re.compile(
    r"Общее количество участников:\s*не менее\s+(\d+)",
    re.IGNORECASE,
)
_LECTURES_DAY = re.compile(
    r"(\d{1,2}\s+июня\s+\d{4}\s+года)\s+в период[^()\n]*\((\d+)\s+лекци[^\)]*\),\s*([^\n]+)",
    re.IGNORECASE,
)
_PROGRAM_HACKATHON = re.compile(r"не менее\s+\d+[^\n]{0,40}мастер-класс", re.IGNORECASE)
_CHAIRS = re.compile(
    r"стул[^\n]{0,120}?не\s+менее\s+(\d+)",
    re.IGNORECASE,
)
_VOLUNTEERS = re.compile(
    r"волонт[^\n]{0,80}?не\s+менее\s+(\d+)|не\s+менее\s+(\d+)\s+волонт",
    re.IGNORECASE,
)
_ACCOMMODATION = re.compile(
    r"проживание\s+(\d+)\s*\([^)]*\)\s*иногородн[^\n]{0,220}Шереметьевск",
    re.IGNORECASE,
)
_TRANSPORT = re.compile(
    r"перевозк[^\n]{0,120}(\d+)\s*\([^)]*\)\s*иногородн",
    re.IGNORECASE,
)
_BUS = re.compile(
    r"автобус[^\n]{0,80}не\s+менее\s+(\d+)",
    re.IGNORECASE,
)
_COFFEE = re.compile(
    r"(\d{1,2}\s+ноябр[^\n]{0,40})кофе-брейк[^\n]{0,60}(\d+)",
    re.IGNORECASE,
)
_EQUIPMENT = re.compile(
    r"6\.2\.1[^\n]*Шереметьевск[^\n]{0,1200}",
    re.IGNORECASE | re.DOTALL,
)
_MEALS = re.compile(
    r"8\.2\.1[^\n]{0,900}",
    re.IGNORECASE | re.DOTALL,
)


def _blob(obligations: list[dict[str, Any]]) -> str:
    return " ".join(
        f"{o.get('id','')} {o.get('metric','')} {o.get('clause','')} {o.get('required','')}"
        for o in obligations
    ).lower()


def _item(
    oid: str,
    clause: str,
    metric: str,
    required: str,
    unit: str,
    operator: str,
    evidence: str,
    quote: str,
) -> dict[str, Any]:
    return {
        "id": oid,
        "clause": clause,
        "metric": metric,
        "required": required,
        "unit": unit,
        "operator": operator,
        "evidence_type": evidence,
        "quote": quote[:400],
        "source": "regex_fallback",
    }


def ensure_key_obligations(tz: str, obligations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    blob = _blob(obligations)
    extra: list[dict[str, Any]] = []

    m = _LECTURES.search(tz)
    if m and "лекц" not in blob:
        extra.append(
            _item(
                "TZ-lectures",
                "1.2 / 4.2",
                "Историко-просветительские лекции",
                f"не менее {m.group(1)}",
                "шт.",
                "gte",
                "text",
                m.group(0),
            )
        )

    m = _BACKDROPS.search(tz)
    if m and "застав" not in blob:
        extra.append(
            _item(
                "TZ-backdrops-count",
                "3.4",
                "Статичные заставки",
                f"не менее {m.group(1)}",
                "шт.",
                "gte",
                "text",
                m.group(0),
            )
        )
    elif m and not any(parse_has_number(o) and "застав" in _blob([o]) for o in obligations):
        extra.append(
            _item(
                "TZ-backdrops-count",
                "3.4",
                "Количество статичных заставок",
                f"не менее {m.group(1)}",
                "шт.",
                "gte",
                "text",
                m.group(0),
            )
        )

    m = _PHOTOS.search(tz)
    if m and "фото" not in blob:
        extra.append(
            _item(
                "TZ-photos",
                "3.2 / 4.1.6",
                "Фотоотчёт",
                f"не менее {m.group(1)}",
                "фотографий",
                "gte",
                "photo",
                m.group(0),
            )
        )

    m = _PARTICIPANTS.search(tz)
    if m and "участник" not in blob:
        extra.append(
            _item(
                "TZ-participants",
                "ТЗ шапка",
                "Количество участников",
                f"не менее {m.group(1)}",
                "человек",
                "gte",
                "text",
                m.group(0),
            )
        )

    m = _CHAIRS.search(tz)
    if m and "стул" not in blob:
        extra.append(
            _item(
                "TZ-chairs",
                "6.2 / 4.1",
                "Количество стульев",
                f"не менее {m.group(1)}",
                "шт.",
                "gte",
                "text",
                m.group(0),
            )
        )

    m = _VOLUNTEERS.search(tz)
    if m and "волонт" not in blob:
        count = m.group(1) or m.group(2)
        extra.append(
            _item(
                "TZ-volunteers",
                "7.2",
                "Количество волонтёров в день",
                f"не менее {count}",
                "человек",
                "gte",
                "text",
                m.group(0),
            )
        )

    m = _ACCOMMODATION.search(tz)
    if m and "проживан" not in blob:
        extra.append(
            _item(
                "TZ-accommodation",
                "8.3",
                "Проживание иногородних участников",
                f"{m.group(1)} человек, «Шереметьевский», 27–30.11.2025",
                "проживание",
                "text",
                "text",
                m.group(0),
            )
        )

    m = _TRANSPORT.search(tz)
    if m and "перевозк" not in blob:
        extra.append(
            _item(
                "TZ-transport-count",
                "8.1",
                "Перевозка иногородних участников",
                f"{m.group(1)}",
                "человек",
                "eq",
                "text",
                m.group(0),
            )
        )

    m = _BUS.search(tz)
    if m and "автобус" not in blob:
        extra.append(
            _item(
                "TZ-bus",
                "9.1",
                "Автобус для участников",
                f"не менее {m.group(1)}",
                "мест",
                "gte",
                "text",
                m.group(0),
            )
        )

    m = _COFFEE.search(tz)
    if m and "кофе" not in blob:
        extra.append(
            _item(
                "TZ-coffee-break",
                "8.2.2",
                "Кофе-брейк",
                f"{m.group(2)} человек, {m.group(1)}",
                "человек",
                "eq",
                "text",
                m.group(0),
            )
        )

    m = _EQUIPMENT.search(tz)
    if m and "electrovoice" not in blob and "оборудован" not in blob:
        extra.append(
            _item(
                "TZ-equipment-sheremetyevsky",
                "6.2.1",
                "Техническое оборудование (Шереметьевский)",
                "Electrovoice, Allen&Heath, Shure — по спецификации ТЗ",
                "комплект",
                "text",
                "text",
                m.group(0)[:400],
            )
        )

    m = _MEALS.search(tz)
    if m and "8.2.1" not in blob and "питани" not in blob:
        extra.append(
            _item(
                "TZ-meals",
                "8.2.1",
                "Организация питания участников",
                "обеды 28–30.11.2025 по расписанию ТЗ",
                "питание",
                "text",
                "text",
                m.group(0)[:400],
            )
        )

    for match in _LECTURES_DAY.finditer(tz):
        day, count, venue = match.group(1), match.group(2), match.group(3).strip()
        oid = "TZ-lectures-" + re.sub(r"\s+", "", day.lower())[:12]
        if oid in {o.get("id") for o in obligations}:
            continue
        extra.append(
            _item(
                oid,
                "4.2",
                f"Лекции {day}: число и адрес",
                f"{count} лекций, {venue}",
                "лекции / адрес",
                "text",
                "text",
                match.group(0),
            )
        )

    if extra:
        obligations = list(obligations) + extra
    return obligations


def audit_obligations(tz: str, obligations: list[dict[str, Any]]) -> list[str]:
    """Предупреждения о пропущенных ключевых пунктах ТЗ (после Stage A и fallback)."""
    warnings: list[str] = []
    blob = _blob(obligations)
    ids = {o.get("id", "") for o in obligations}

    expected_from_tz: list[tuple[str, re.Pattern[str], str, str]] = [
        ("TZ-lectures", _LECTURES, "лекц", "лекции"),
        ("TZ-backdrops-count", _BACKDROPS, "застав", "заставки"),
        ("TZ-photos", _PHOTOS, "фото", "фотоотчёт"),
        ("TZ-participants", _PARTICIPANTS, "участник", "участники"),
        ("TZ-chairs", _CHAIRS, "стул", "стулья (6.2)"),
        ("TZ-volunteers", _VOLUNTEERS, "волонт", "волонтёры (7.2)"),
        ("TZ-accommodation", _ACCOMMODATION, "проживан", "проживание (8.3)"),
        ("TZ-transport-count", _TRANSPORT, "перевозк", "перевозка (8.1)"),
        ("TZ-bus", _BUS, "автобус", "автобус (9.1)"),
        ("TZ-coffee-break", _COFFEE, "кофе", "кофе-брейк (8.2.2)"),
        ("TZ-equipment-sheremetyevsky", _EQUIPMENT, "оборудован", "оборудование (6.2.1)"),
        ("TZ-meals", _MEALS, "питани", "питание (8.2.1)"),
    ]
    for oid, pattern, keyword, label in expected_from_tz:
        if not pattern.search(tz):
            continue
        if oid not in ids and keyword not in blob:
            warnings.append(f"В чеклисте нет пункта «{label}» ({oid}), хотя он есть в ТЗ")

    section_markers = ("6.1", "6.2", "7.2", "8.1", "8.2", "8.3", "9.1")
    for marker in section_markers:
        if marker not in tz:
            continue
        if marker not in blob and not any(marker in (o.get("clause") or "") for o in obligations):
            warnings.append(f"Нет обязательства с привязкой к п. {marker} ТЗ")

    if any(o.get("id") == "TZ-other-items" for o in obligations):
        warnings.append(
            "Часть обязательств свёрнута в TZ-other-items — проверьте, не потеряны ли п. 6–9"
        )

    return warnings


def parse_has_number(obligation: dict[str, Any]) -> bool:
    return bool(re.search(r"\d", obligation.get("required") or ""))
