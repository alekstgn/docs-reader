"""JSON-схемы structured output для Gemini."""

from __future__ import annotations

from typing import Any

STRING = {"type": "STRING"}
BOOL = {"type": "BOOLEAN"}


def _obj(properties: dict[str, Any], required: list[str]) -> dict[str, Any]:
    return {
        "type": "OBJECT",
        "properties": properties,
        "required": required,
    }


def _arr(item: dict[str, Any]) -> dict[str, Any]:
    return {"type": "ARRAY", "items": item}


EVENT_SCHEMA = _obj(
    {
        "title": {**STRING, "description": "Официальное название мероприятия"},
        "dates": {**STRING, "description": "Даты проведения как в документе"},
        "venue": {**STRING, "description": "Площадки / адреса"},
        "city": STRING,
        "participants_required_or_claimed": STRING,
    },
    ["title", "dates", "venue", "city", "participants_required_or_claimed"],
)

OBLIGATION_ITEM = _obj(
    {
        "id": {**STRING, "description": "Стабильный id, напр. TZ-1.2-lectures"},
        "clause": {**STRING, "description": "Краткое имя пункта ТЗ"},
        "metric": {**STRING, "description": "Что измеряется"},
        "required": {**STRING, "description": "Требование как в договоре"},
        "unit": STRING,
        "operator": {
            "type": "STRING",
            "enum": ["gte", "lte", "eq", "range", "text"],
            "description": "gte=не менее, lte=не более, eq=точно, text=качественное",
        },
        "evidence_type": {
            "type": "STRING",
            "enum": ["text", "photo", "external"],
        },
        "quote": {**STRING, "description": "Точная цитата ≤400 символов"},
    },
    [
        "id",
        "clause",
        "metric",
        "required",
        "unit",
        "operator",
        "evidence_type",
        "quote",
    ],
)

STAGE_A_SCHEMA = _obj(
    {
        "event": EVENT_SCHEMA,
        "obligations": _arr(OBLIGATION_ITEM),
    },
    ["event", "obligations"],
)

CLAIM_ITEM = _obj(
    {
        "id": {**STRING, "description": "Тот же id, что в чеклисте ТЗ"},
        "found": BOOL,
        "claimed": {**STRING, "description": "Заявленный факт из отчёта или пусто"},
        "quote": {**STRING, "description": "Цитата из отчёта ≤400 символов"},
    },
    ["id", "found", "claimed", "quote"],
)

STAGE_B_SCHEMA = _obj(
    {
        "event": EVENT_SCHEMA,
        "claims": _arr(CLAIM_ITEM),
    },
    ["event", "claims"],
)

QUAL_ITEM = _obj(
    {
        "id": STRING,
        "status": {
            "type": "STRING",
            "enum": [
                "ok",
                "mismatch",
                "missing_in_report",
                "not_verifiable_from_docs",
            ],
        },
        "comment": STRING,
    },
    ["id", "status", "comment"],
)

STAGE_QUAL_SCHEMA = _obj({"items": _arr(QUAL_ITEM)}, ["items"])

PHOTO_ITEM = _obj(
    {
        "photo_id": {**STRING, "description": "Имя файла из манифеста, без пути"},
        "visible_objects": _arr(STRING),
        "matched_obligation_ids": _arr(STRING),
        "confidence": {
            "type": "STRING",
            "enum": ["high", "medium", "low"],
        },
        "notes": {
            **STRING,
            "description": "Что видно для сверки ТЗ; без имён людей и описания лиц; по-русски",
        },
        "conclusion": {
            "type": "STRING",
            "enum": ["confirms", "contradicts", "inconclusive", "wrong_event"],
            "description": "confirms=подтверждает ТЗ договора; contradicts=противоречит; inconclusive=неясно; wrong_event=другое мероприятие",
        },
        "event_fit": {
            "type": "STRING",
            "enum": ["contracted_event", "other_event", "unknown"],
            "description": "Кажется ли кадр договором, другим событием или неясно",
        },
        "branding_or_text_seen": {
            **STRING,
            "description": "Текст/логотипы на кадре без персональных данных",
        },
        "scene_type": {
            "type": "STRING",
            "enum": [
                "press_wall",
                "hall_audience",
                "stage_equipment",
                "backdrop_screen",
                "other",
            ],
        },
    },
    [
        "photo_id",
        "visible_objects",
        "matched_obligation_ids",
        "confidence",
        "notes",
        "conclusion",
        "event_fit",
        "branding_or_text_seen",
        "scene_type",
    ],
)

STAGE_C_SCHEMA = _obj({"photos": _arr(PHOTO_ITEM)}, ["photos"])
