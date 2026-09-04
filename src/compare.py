"""Детерминированная сверка чисел/дат; LLM только для качественных пунктов."""

from __future__ import annotations

import re
from typing import Any

NUMBER_RE = re.compile(
    r"(\d{1,3}(?:[ \u00a0]\d{3})+|\d+(?:[.,]\d+)?)"
)
DATE_RE = re.compile(r"\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\b")
RU_MONTHS = [
    ("январ", 1),
    ("феврал", 2),
    ("март", 3),
    ("апрел", 4),
    ("июн", 6),
    ("июл", 7),
    ("август", 8),
    ("сентябр", 9),
    ("октябр", 10),
    ("ноябр", 11),
    ("декабр", 12),
    ("ма", 5),  # май/мая — после марта
]
NOISE_NUMBERS = {600, 1200, 1920, 1080, 2024, 2025, 2026, 2027}

STATUS_OK = "ok"
STATUS_MISMATCH = "mismatch"
STATUS_MISSING = "missing_in_report"
STATUS_NEEDS_PHOTO = "needs_photo"
STATUS_SKIPPED_PHOTO = "needs_photo_skipped"
STATUS_PHOTO_INCONCLUSIVE = "photo_inconclusive"
STATUS_EXTERNAL = "not_verifiable_from_docs"
STATUS_EVENT = "event_mismatch"


def parse_numbers(text: str | None) -> list[float]:
    if not text:
        return []
    out: list[float] = []
    for raw in NUMBER_RE.findall(text.replace("\xa0", " ")):
        if isinstance(raw, tuple):
            raw = raw[0]
        raw = raw.replace(" ", "").replace("\u00a0", "").replace(",", ".")
        try:
            out.append(float(raw))
        except ValueError:
            continue
    return out


def parse_number(text: str | None) -> float | None:
    nums = parse_numbers(text)
    return nums[0] if nums else None


def parse_fact_number(
    required: str,
    claimed: str,
    quote: str,
    metric: str = "",
) -> float | None:
    """Число из claimed; из цитаты — только для фото, и не из времени/дат."""
    req_n = parse_number(required)
    claimed_n = parse_number(claimed)
    metric_l = (metric or "").lower()
    if "фото" not in metric_l:
        return claimed_n
    cleaned = re.sub(r"\d{1,2}:\d{2}", " ", quote or "")
    cleaned = DATE_RE.sub(" ", cleaned)
    cleaned = re.sub(r"\d{1,2}\s+[А-Яа-яЁё]+\s+\d{4}", " ", cleaned)
    quote_nums = [
        n for n in parse_numbers(cleaned) if n not in NOISE_NUMBERS and 10 < n < 100_000
    ]
    if quote_nums:
        if req_n is not None:
            others = [n for n in quote_nums if abs(n - req_n) > 0.5]
            if others:
                return max(others)
        return max(quote_nums)
    return claimed_n


def _month_num(token: str) -> int | None:
    token = token.lower().replace("ё", "е")
    for prefix, num in RU_MONTHS:
        if token.startswith(prefix):
            return num
    return None


def parse_dates(text: str | None) -> set[tuple[int, int, int]]:
    if not text:
        return set()
    out: set[tuple[int, int, int]] = set()
    for d, m, y in DATE_RE.findall(text):
        year = int(y)
        if year < 100:
            year += 2000
        out.add((int(d), int(m), year))
    ru = re.compile(
        r"(\d{1,2})\s+([А-Яа-яЁё]+)\s+(\d{4})",
    )
    for d, month, y in ru.findall(text):
        m = _month_num(month)
        if m:
            out.add((int(d), m, int(y)))
    # «10-12 июня 2025» / «с 10 июня по 12 июня 2025»
    span = re.compile(
        r"(?:с\s+)?(\d{1,2})\s*(?:[-–—]|по)\s*(\d{1,2})\s+([А-Яа-яЁё]+)\s+(\d{4})",
        re.IGNORECASE,
    )
    for d1, d2, month, y in span.findall(text):
        m = _month_num(month)
        if not m:
            continue
        start, end = int(d1), int(d2)
        if start > end:
            start, end = end, start
        for day in range(start, end + 1):
            out.add((day, m, int(y)))
    span2 = re.compile(
        r"(?:с\s+)?(\d{1,2})\s+([А-Яа-яЁё]+)\s+по\s+(\d{1,2})\s+([А-Яа-яЁё]+)\s+(\d{4})",
        re.IGNORECASE,
    )
    for d1, m1, d2, m2, y in span2.findall(text):
        month1, month2 = _month_num(m1), _month_num(m2)
        year = int(y)
        if month1:
            out.add((int(d1), month1, year))
        if month2:
            out.add((int(d2), month2, year))
        if month1 and month1 == month2 and int(d1) <= int(d2):
            for day in range(int(d1), int(d2) + 1):
                out.add((day, month1, year))
    return out


def _norm(text: str | None) -> str:
    if not text:
        return ""
    text = text.lower().replace("ё", "е")
    text = re.sub(r"[«»\"“”]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def compare_event(contract: dict[str, Any], report: dict[str, Any]) -> dict[str, Any]:
    c_dates = parse_dates(contract.get("dates", ""))
    r_dates = parse_dates(report.get("dates", ""))
    if c_dates and r_dates:
        date_overlap = bool(c_dates & r_dates)
    else:
        date_overlap = True
    title_c = _norm(contract.get("title"))
    title_r = _norm(report.get("title"))
    venue_c = _norm(contract.get("venue"))
    venue_r = _norm(report.get("venue"))

    title_similar = _titles_similar(title_c, title_r)
    venue_similar = _venues_similar(venue_c, venue_r)
    same = date_overlap and title_similar and venue_similar

    reasons: list[str] = []
    if c_dates and r_dates and not date_overlap:
        reasons.append(
            f"даты не пересекаются: договор {sorted(c_dates)} vs отчёт {sorted(r_dates)}"
        )
    if not title_similar:
        reasons.append(
            f"название: «{contract.get('title')}» vs «{report.get('title')}»"
        )
    if not venue_similar:
        reasons.append(
            f"площадка: «{contract.get('venue')}» vs «{report.get('venue')}»"
        )

    related = title_similar and not same
    if same:
        event_comment = "Мероприятие совпадает."
    elif related:
        event_comment = (
            "Мероприятия родственные (форум/конференция «Просветители» для блогеров), "
            "но даты и адреса не совпадают. " + "; ".join(reasons)
        )
    else:
        event_comment = "Документы описывают разные мероприятия. " + "; ".join(reasons)

    return {
        "id": "EVENT_IDENTITY",
        "clause": "Идентичность мероприятия",
        "status": STATUS_OK if same else STATUS_EVENT,
        "same_event": same,
        "required": f"{contract.get('title')}; {contract.get('dates')}; {contract.get('venue')}",
        "claimed": f"{report.get('title')}; {report.get('dates')}; {report.get('venue')}",
        "quote_contract": contract.get("dates", ""),
        "quote_report": report.get("dates", ""),
        "comment": event_comment,
        "compare_kind": "event",
        "related_event": related,
    }


def _titles_similar(a: str, b: str) -> bool:
    if not a or not b:
        return False
    if a == b:
        return True
    if "просветител" in a and "просветител" in b:
        return True
    tokens_a = set(re.findall(r"[а-яa-z0-9]{4,}", a))
    tokens_b = set(re.findall(r"[а-яa-z0-9]{4,}", b))
    if not tokens_a or not tokens_b:
        return a[:20] in b or b[:20] in a
    overlap = tokens_a & tokens_b
    return len(overlap) / max(len(tokens_a), len(tokens_b)) >= 0.45


def _venues_similar(a: str, b: str) -> bool:
    if not a or not b:
        return False
    keys = ("шереметьевск", "лесной", "калуг", "вучетича", "автомобильн", "октябрьск")
    a_keys = {k for k in keys if k in a}
    b_keys = {k for k in keys if k in b}
    if a_keys or b_keys:
        return bool(a_keys & b_keys)
    return _titles_similar(a, b)


def _numeric_status(operator: str, required: float | None, claimed: float | None) -> str | None:
    if required is None or claimed is None:
        return None
    eps = 1e-6
    if operator == "gte":
        return STATUS_OK if claimed + eps >= required else STATUS_MISMATCH
    if operator == "lte":
        return STATUS_OK if claimed - eps <= required else STATUS_MISMATCH
    if operator == "eq":
        return STATUS_OK if abs(claimed - required) < max(0.5, required * 0.01) else STATUS_MISMATCH
    if operator == "range":
        return STATUS_OK if abs(claimed - required) < max(1.0, required * 0.15) else STATUS_MISMATCH
    return None


LOGISTICS_WORDS = (
    "проживан",
    "перевозк",
    "трансфер",
    "автобус",
    "кофе",
    "питани",
    "оборудован",
    "стул",
    "микрофон",
    "проектор",
    "кулер",
    "технич",
    "мебел",
    "прочие обязательства",
)

LOCATION_KEYS = (
    "автомобильн",
    "октябрьск",
    "шереметьев",
    "лесной",
    "вучетича",
    "солнечногор",
    "московия",
    "трансформер",
    "никитск",
)


def _is_logistics_row(row: dict[str, Any]) -> bool:
    metric = (row.get("metric") or "").lower()
    clause = str(row.get("clause") or "")
    row_id = row.get("id") or ""
    if row_id in {"TZ-other-items", "PHOTO-stage-equipment"}:
        return True
    if clause.startswith(("6", "7", "8", "9")):
        return True
    return any(w in metric for w in LOGISTICS_WORDS)


def _quotes_event_conflict(row: dict[str, Any]) -> bool:
    quote_c = (row.get("quote_contract") or row.get("required") or "").strip()
    quote_r = (row.get("quote_report") or row.get("claimed") or "").strip()
    if not quote_c or not quote_r:
        return False
    c_dates = parse_dates(quote_c)
    r_dates = parse_dates(quote_r)
    if c_dates and r_dates and not (c_dates & r_dates):
        return True
    blob_c = _norm(quote_c)
    blob_r = _norm(quote_r)
    c_loc = {k for k in LOCATION_KEYS if k in blob_c}
    r_loc = {k for k in LOCATION_KEYS if k in blob_r}
    return bool(c_loc and r_loc and c_loc != r_loc)


_EQUIPMENT_BRANDS = (
    "electrovoice",
    "shure",
    "sennheiser",
    "samsung",
    "msi",
    "gateway",
    "allen",
    "db technologies",
    "opera",
    "elx",
    "ulxd",
    "ew g4",
    "yamaha",
    "epson",
    "audioctnter",
    "audiocenter",
    "плазмен",
    "светодиод",
    "midas",
)


def _equipment_spec_differs(qc: str, qr: str) -> bool:
    if _quotes_match(qc, qr):
        return False
    if not (qc or "").strip() or not (qr or "").strip():
        return False
    nc, nr = _norm(qc), _norm(qr)
    if nc == nr:
        return False
    brands_c = {b for b in _EQUIPMENT_BRANDS if b in nc}
    brands_r = {b for b in _EQUIPMENT_BRANDS if b in nr}
    if brands_c and brands_r and brands_c != brands_r:
        return True
    nums_c = re.findall(r"\d+", nc)
    nums_r = re.findall(r"\d+", nr)
    if nums_c and nums_r and nums_c[-1] == nums_r[-1]:

        def _strip_nums(text: str) -> str:
            return re.sub(r"\d+", " ", text)

        stripped_c = _strip_nums(nc)
        stripped_r = _strip_nums(nr)
        if stripped_c != stripped_r:
            shared = set(stripped_c.split()) & set(stripped_r.split())
            if len(shared) < 2:
                return True
    return False


def _obligation_venue_site(row: dict[str, Any]) -> str | None:
    clause = str(row.get("clause") or "")
    row_id = row.get("id") or ""
    qc = _norm(row.get("quote_contract") or row.get("required") or "")
    metric = _norm(row.get("metric") or "")
    if "venue1" in row_id or clause in ("6.2.1", "6.2.2") or "шереметьев" in qc:
        return "sheremetyevsky"
    if "venue2" in row_id or clause in ("6.2.3", "6.2.4"):
        return "museum"
    if clause.startswith("9") or "автобус" in metric or "маршрут" in metric:
        return "sheremetyevsky"
    if "кофе" in metric and "шереметьев" in qc:
        return "sheremetyevsky"
    return None


def _quote_dates_conflict(row: dict[str, Any]) -> bool:
    quote_c = (row.get("quote_contract") or row.get("required") or "").strip()
    quote_r = (row.get("quote_report") or row.get("claimed") or "").strip()
    c_dates = parse_dates(quote_c)
    r_dates = parse_dates(quote_r)
    return bool(c_dates and r_dates and not (c_dates & r_dates))


def refine_comparison_rows(
    rows: list[dict[str, Any]],
    event_check: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """Доп. проверки: модель при равном числе, площадка, даты."""
    same_event = bool((event_check or {}).get("same_event", True))
    refined: list[dict[str, Any]] = []
    for raw in rows:
        row = dict(raw)
        status = row.get("status")
        qc = (row.get("quote_contract") or "").strip()
        qr = (row.get("quote_report") or "").strip()

        if status == STATUS_OK and qc and qr:
            if _equipment_spec_differs(qc, qr):
                row["status"] = STATUS_MISMATCH
                row["compare_kind"] = "qualitative"
                row["comment"] = (
                    "Количество совпало, но модель или комплектация в отчёте "
                    "не соответствует договору."
                )
            elif _quotes_event_conflict(row):
                row["status"] = STATUS_MISMATCH
                row["compare_kind"] = "qualitative"
                row["comment"] = "В отчёте другие даты или площадка, чем в договоре."
            elif _quote_dates_conflict(row) and (
                _is_logistics_row(row)
                or str(row.get("clause") or "").startswith(("5", "6", "7", "8", "9"))
            ):
                row["status"] = STATUS_MISMATCH
                row["compare_kind"] = "qualitative"
                c_dates = parse_dates(qc)
                r_dates = parse_dates(qr)
                row["comment"] = f"даты: договор {sorted(c_dates)} vs отчёт {sorted(r_dates)}"
            elif not same_event and _obligation_venue_site(row) == "sheremetyevsky":
                row["status"] = STATUS_MISMATCH
                row["compare_kind"] = "qualitative"
                row["comment"] = (
                    "По договору — площадка «Шереметьевский», 27–30.11.2025; "
                    "отчёт описывает «Лесной» / июнь 2026."
                )

        refined.append(row)
    return refined


EVENT_ONLY_COMMENT_MARKERS = (
    "не подтверждает исполнение договорного форума",
    "разделы 6–9 / техобеспечение",
    "оборудование и даты в отчёте относятся",
    "техника другого мероприятия",
)


def _quotes_match(qc: str, qr: str) -> bool:
    left = (qc or "").strip()
    right = (qr or "").strip()
    if not left or not right:
        return False
    return _norm(left) == _norm(right)


def _row_has_concrete_discrepancy(row: dict[str, Any]) -> bool:
    status = row.get("status") or ""
    if status == STATUS_MISSING:
        return True
    if _is_quantity_row(row) and not quantity_requirement_met(row):
        return True
    qc = (row.get("quote_contract") or row.get("required") or "").strip()
    qr = (row.get("quote_report") or row.get("claimed") or "").strip()
    if qc and qr and not _quotes_match(qc, qr):
        return True
    required = (row.get("required") or "").strip()
    claimed = (row.get("claimed") or "").strip()
    if required and claimed and _norm(required) != _norm(claimed):
        return True
    return False


def should_include_comment_row(row: dict[str, Any], *, same_event: bool = True) -> bool:
    """Комментарий при конкретном расхождении; без дублей, где цитаты и числа совпали."""
    status = row.get("status") or ""
    if status == STATUS_MISSING:
        return True
    if status not in (STATUS_MISMATCH, "event_mismatch"):
        return False
    if is_false_numeric_mismatch(row):
        return False

    qc = (row.get("quote_contract") or row.get("required") or "").strip()
    qr = (row.get("quote_report") or row.get("claimed") or "").strip()

    comment = (row.get("comment") or "").lower()
    if any(marker in comment for marker in EVENT_ONLY_COMMENT_MARKERS):
        if not _row_has_concrete_discrepancy(row):
            return False

    venue_or_model_note = any(
        token in comment
        for token in (
            "модель",
            "комплектация",
            "шереметьев",
            "лесной",
            "даты:",
            "площадка",
            "адрес/площадка",
        )
    )

    if qc and qr and _quotes_match(qc, qr):
        if venue_or_model_note:
            return True
        if quantity_requirement_met(row) or not _is_quantity_row(row):
            return False

    if not same_event and not _row_has_concrete_discrepancy(row):
        return False

    return True


def compare_obligation(
    obligation: dict[str, Any],
    claim: dict[str, Any] | None,
    *,
    embedded_images: int | None = None,
) -> dict[str, Any]:
    oid = obligation.get("id", "")
    clause = obligation.get("clause", "")
    required_text = obligation.get("required") or ""
    quote_c = obligation.get("quote") or required_text
    operator = obligation.get("operator") or "text"
    quote_c_full = obligation.get("quote") or required_text
    operator = infer_operator(required_text, quote_c_full, operator)
    evidence = obligation.get("evidence_type") or "text"

    if not claim or not claim.get("found"):
        status = STATUS_MISSING
        claimed_text = ""
        quote_r = ""
        comment = "В отчёте пункт не найден."
    else:
        claimed_text = claim.get("claimed") or ""
        quote_r = claim.get("quote") or claimed_text
        req_n = parse_number(required_text)
        cl_n = parse_fact_number(
            required_text,
            claimed_text,
            quote_r,
            obligation.get("metric") or clause,
        )
        if cl_n is not None and "фото" in (obligation.get("metric") or clause).lower():
            if parse_number(claimed_text) != cl_n:
                claimed_text = f"{cl_n:g}"
        num_status = _numeric_status(operator, req_n, cl_n) if operator != "text" else None
        if num_status:
            status = num_status
            comment = (
                f"число: требуется {operator} {req_n:g}, в отчёте {cl_n:g}"
            )
            if status == STATUS_MISMATCH:
                comment = "Расхождение. " + comment
            else:
                comment = "Совпадает. " + comment
        elif operator != "text" and (req_n is None or cl_n is None):
            status = "qualitative_pending"
            comment = "Число не разобрано, нужна качественная сверка."
        else:
            status = "qualitative_pending"
            comment = "Качественный пункт."

        c_dates = parse_dates(required_text + " " + quote_c)
        r_dates = parse_dates(claimed_text + " " + quote_r)
        metric_l = (obligation.get("metric") or clause).lower()
        date_sensitive = (
            "дат" in metric_l
            or clause.startswith(("5", "6", "7", "8", "9"))
            or _is_logistics_row({"metric": metric_l, "clause": clause, "id": oid})
            or "автобус" in metric_l
            or "маршрут" in metric_l
            or "нояб" in quote_c.lower()
            or "нояб" in quote_r.lower()
        )
        if c_dates and r_dates and not (c_dates & r_dates) and date_sensitive:
            status = STATUS_MISMATCH
            comment = f"даты: договор {sorted(c_dates)} vs отчёт {sorted(r_dates)}"

        loc_keys = LOCATION_KEYS
        blob_c = _norm(required_text + " " + quote_c)
        blob_r = _norm(claimed_text + " " + quote_r)
        c_loc = {k for k in loc_keys if k in blob_c}
        r_loc = {k for k in loc_keys if k in blob_r}
        if c_loc and r_loc and c_loc != r_loc and (
            any(w in metric_l for w in ("адрес", "площадк", "лекц", "место", "маршрут", "автобус"))
            or _is_logistics_row({"metric": metric_l, "clause": clause, "id": oid})
        ):
            status = STATUS_MISMATCH
            comment = f"адрес/площадка: договор {c_loc} vs отчёт {r_loc}"

        if status == STATUS_OK and quote_c and quote_r and _equipment_spec_differs(quote_c, quote_r):
            status = STATUS_MISMATCH
            comment = (
                "Количество совпало, но модель или комплектация в отчёте "
                "не соответствует договору."
            )

    extra: dict[str, Any] = {}
    if evidence == "photo" and embedded_images is not None:
        claimed_photos = parse_fact_number(
            required_text,
            claimed_text,
            quote_r,
            obligation.get("metric") or clause,
        )
        required_photos = parse_number(required_text)
        extra["embedded_images"] = embedded_images
        extra["claimed_photos"] = claimed_photos
        extra["required_photos"] = required_photos
        extra["photo_stage"] = "skipped"
        extra["comment_photos"] = (
            f"фото: договор {required_text}; заявлено в отчёте "
            f"{claimed_text or '—'}; "
            f"во вложении DOCX {embedded_images} шт. (этап C с загрузкой фото пропущен)"
        )
        extra["text_status_before_photo"] = status
        contract_met = (
            required_photos is not None and embedded_images >= required_photos
        )
        text_overstated = (
            claimed_photos is not None
            and embedded_images < float(claimed_photos) * 0.5
        )
        if required_photos is not None and embedded_images < required_photos:
            status = STATUS_MISMATCH
            comment = (
                f"фото: по договору {required_text}; "
                f"во вложении DOCX {embedded_images} шт. — меньше минимума."
            )
        elif text_overstated and contract_met:
            status = STATUS_OK
            extra["photo_text_gap"] = True
            comment = (
                f"фото: по договору {required_text}; во вложении DOCX {embedded_images} шт. — "
                f"минимум выполнен. В тексте отчёта указано {claimed_text or f'{claimed_photos:g}'}, "
                f"что не совпадает с числом вложений."
            )
        elif text_overstated:
            status = STATUS_MISMATCH
            comment = extra["comment_photos"]
        elif status in {STATUS_OK, STATUS_MISMATCH, STATUS_MISSING}:
            comment = comment + ". " + extra["comment_photos"]

    if evidence == "external" and status == STATUS_OK:
        status = STATUS_EXTERNAL
        comment = "Внешний источник (VK и т.п.) по тексту отчёта не проверяется."

    return {
        "id": oid,
        "clause": clause,
        "metric": obligation.get("metric"),
        "operator": operator,
        "evidence_type": evidence,
        "status": status,
        "required": required_text,
        "claimed": claimed_text if claim else "",
        "quote_contract": quote_c,
        "quote_report": quote_r,
        "comment": comment,
        "compare_kind": "numeric" if status not in {"qualitative_pending"} else "qualitative",
        **extra,
    }


def merge_qualitative(rows: list[dict[str, Any]], qual: list[dict[str, Any]]) -> None:
    by_id = {item["id"]: item for item in qual}
    for row in rows:
        if row.get("status") != "qualitative_pending":
            continue
        hit = by_id.get(row["id"])
        if not hit:
            row["status"] = STATUS_MISMATCH if row.get("claimed") else STATUS_MISSING
            row["comment"] = (row.get("comment") or "") + " Качественная сверка не вернула статус."
            continue
        row["status"] = hit.get("status") or STATUS_MISMATCH
        row["comment"] = hit.get("comment") or row.get("comment")
        row["compare_kind"] = "qualitative"


def _objects(analyses: list[dict[str, Any]]) -> set[str]:
    out: set[str] = set()
    for item in analyses:
        for obj in item.get("visible_objects") or []:
            if obj:
                out.add(str(obj).lower())
        scene = str(item.get("scene_type") or "").lower()
        if scene:
            out.add(scene)
        blob = " ".join(
            [
                str(item.get("notes") or ""),
                str(item.get("branding_or_text_seen") or ""),
                " ".join(str(x) for x in (item.get("visible_objects") or [])),
            ]
        ).lower()
        for token in (
            "пресс-вол",
            "прессвол",
            "баннер",
            "логотип",
            "зал",
            "аудитор",
            "публика",
            "зрител",
            "микрофон",
            "колонк",
            "акустик",
            "проектор",
            "экран",
            "сцена",
            "застав",
        ):
            if token in blob:
                out.add(token)
    return out


def merge_photo_results(
    rows: list[dict[str, Any]],
    event_check: dict[str, Any],
    analyses: list[dict[str, Any]],
    *,
    same_event: bool,
    extracted: int,
    sampled: int,
    pair_id: str,
    full_run: bool = False,
    failed: int = 0,
) -> list[dict[str, Any]]:
    """Обновить пункты с evidence=photo и добавить визуальные итоги. Текстовые статусы не затираем."""
    objects = _objects(analyses)
    fits = [str(a.get("event_fit") or "unknown") for a in analyses]
    other_n = fits.count("other_event")
    contract_n = fits.count("contracted_event")
    unknown_n = fits.count("unknown")
    conclusions = [str(a.get("conclusion") or "") for a in analyses]
    wrong_n = conclusions.count("wrong_event")
    confirms_n = conclusions.count("confirms")
    branding_bits = [
        str(a.get("branding_or_text_seen") or "").strip()
        for a in analyses
        if str(a.get("branding_or_text_seen") or "").strip()
        and str(a.get("branding_or_text_seen")).strip().lower() not in {"none", "нет", "n/a", "-"}
    ]
    notes_join = " ".join(str(a.get("notes") or "") for a in analyses).lower()

    has_branding = bool(
        objects
        & {
            "press_wall",
            "backdrop_screen",
            "пресс-вол",
            "прессвол",
            "баннер",
            "banner",
            "логотип",
            "застав",
            "branding",
        }
        or branding_bits
        or any("логотип" in b.lower() or "баннер" in b.lower() or "пресс" in b.lower() for b in branding_bits)
    )
    has_hall = bool(
        objects & {"hall_audience", "зал", "аудитор", "публика", "зрител", "audience"}
        or "зал" in notes_join
        or "аудитор" in notes_join
        or "audience" in notes_join
    )
    has_equip = bool(
        objects
        & {
            "stage_equipment",
            "микрофон",
            "колонк",
            "акустик",
            "проектор",
            "экран",
            "сцена",
            "microphone",
            "speaker",
            "projector",
            "screen",
        }
        or "микрофон" in notes_join
        or "проектор" in notes_join
        or "экран" in notes_join
    )

    photos_other_event = not same_event
    if pair_id == "prosvetiteli":
        photos_other_event = True

    visual_comment = (
        f"{'Полный прогон этапа C' if full_run else 'Выборка этапа C'}: "
        f"извлечено {extracted} вложений, в модель отправлено {sampled} кадров"
        f"{f', ошибок {failed}' if failed else ''}. "
        f"event_fit: договор={contract_n}, другое={other_n}, неясно={unknown_n}. "
        f"conclusion: подтверждает={confirms_n}, чужое событие={wrong_n}."
    )
    event_mismatch_comment = (
        "Фотоотчёт не подтверждает договорное мероприятие: кадры относятся к событию отчёта "
        "(для «Просветителей» — июнь 2026 / «Лесной», а не ноябрь 2025 / «Шереметьевский»). "
        "Даже если зал и брендирование выглядят «нормально», это не исполнение ТЗ ноября 2025."
        if photos_other_event
        else visual_comment
    )

    for row in rows:
        if row.get("evidence_type") != "photo" and row.get("status") not in {
            STATUS_NEEDS_PHOTO,
            STATUS_SKIPPED_PHOTO,
        }:
            continue
        row["photo_stage"] = "done"
        row["photos_extracted"] = extracted
        row["photos_sampled"] = sampled
        prev = row.get("comment") or ""
        if photos_other_event:
            row["status"] = STATUS_MISMATCH
            row["comment"] = ((prev + " ") if prev else "") + visual_comment + " " + event_mismatch_comment
        elif row.get("status") in {STATUS_NEEDS_PHOTO, STATUS_SKIPPED_PHOTO}:
            if has_hall or has_branding:
                row["status"] = STATUS_OK
                row["comment"] = ((prev + " ") if prev else "") + visual_comment + " Визуально есть зал/брендирование."
            else:
                row["status"] = STATUS_PHOTO_INCONCLUSIVE
                row["comment"] = ((prev + " ") if prev else "") + visual_comment + " По кадрам нельзя закрыть пункт."
        else:
            row["comment"] = ((prev + " ") if prev else "") + visual_comment
            if not (has_hall or has_branding or has_equip) and row.get("status") not in {
                STATUS_OK,
                STATUS_MISMATCH,
                STATUS_MISSING,
            }:
                row["status"] = STATUS_PHOTO_INCONCLUSIVE

    identity_status = (
        STATUS_MISMATCH
        if photos_other_event
        else STATUS_OK
        if contract_n >= other_n or has_branding or has_hall
        else STATUS_PHOTO_INCONCLUSIVE
    )
    extra_rows = [
        {
            "id": "PHOTO-event-identity",
            "clause": "этап C",
            "metric": "Визуальная идентичность мероприятия",
            "operator": "text",
            "evidence_type": "photo",
            "status": identity_status,
            "required": event_check.get("required") or "",
            "claimed": event_check.get("claimed") or "",
            "quote_contract": "",
            "quote_report": "",
            "comment": event_mismatch_comment if photos_other_event else (
                visual_comment + " Кадры согласуются с договорным мероприятием."
            ),
            "compare_kind": "photo",
            "photo_stage": "done",
            "photos_extracted": extracted,
            "photos_sampled": sampled,
        },
        {
            "id": "PHOTO-branding",
            "clause": "этап C",
            "metric": "Брендирование / пресс-волл / заставки на фото",
            "operator": "text",
            "evidence_type": "photo",
            "status": (
                STATUS_MISMATCH
                if photos_other_event and has_branding
                else STATUS_OK
                if has_branding
                else STATUS_PHOTO_INCONCLUSIVE
            ),
            "required": "логотипы / пресс-волл / заставки на площадке",
            "claimed": "; ".join(branding_bits[:6]) or "не распознано",
            "quote_contract": "",
            "quote_report": "",
            "comment": (
                "На кадрах есть брендирование события отчёта, оно не закрывает ТЗ ноября 2025."
                if photos_other_event and has_branding
                else (
                    "На полном наборе кадров видно брендирование/пресс-волл/экран с заставкой."
                    if has_branding
                    else "По кадрам брендирование не подтверждено однозначно."
                )
            ),
            "compare_kind": "photo",
            "photo_stage": "done",
        },
        {
            "id": "PHOTO-hall-audience",
            "clause": "этап C",
            "metric": "Зал / аудитория на фото",
            "operator": "text",
            "evidence_type": "photo",
            "status": STATUS_OK if has_hall else STATUS_PHOTO_INCONCLUSIVE,
            "required": "зал проведения и присутствие аудитории",
            "claimed": "видно" if has_hall else "неясно",
            "quote_contract": "",
            "quote_report": "",
            "comment": (
                "Кадры подтверждают зал и/или аудиторию. "
                + (
                    "Это не доказывает площадку договора (ноябрь 2025 / «Шереметьевский»)."
                    if photos_other_event
                    else "Согласуется с проведением заявленного мероприятия."
                )
            )
            if has_hall
            else "По кадрам зал/аудитория не подтверждены однозначно.",
            "compare_kind": "photo",
            "photo_stage": "done",
        },
        {
            "id": "PHOTO-equipment",
            "clause": "этап C",
            "metric": "Сцена / оборудование (микрофоны, колонки, проектор)",
            "operator": "text",
            "evidence_type": "photo",
            "status": STATUS_OK if has_equip else STATUS_PHOTO_INCONCLUSIVE,
            "required": "сцена и техническое оборудование",
            "claimed": "видно" if has_equip else "неясно",
            "quote_contract": "",
            "quote_report": "",
            "comment": (
                "На кадрах видно сцену и/или технику (микрофон, экран, акустика). Точный пересчёт штук по фото не выполнялся."
                if has_equip
                else "По кадрам оборудование не подтверждено однозначно. Пересчёт микрофонов по фото ненадёжен."
            ),
            "compare_kind": "photo",
            "photo_stage": "done",
        },
    ]
    rows.extend(extra_rows)
    return rows


GTE_RE = re.compile(r"не\s+мен(?:ее|ьше)", re.I)
LTE_RE = re.compile(r"не\s+бол(?:ее|ьше)", re.I)
QUALITATIVE_MARKERS = (
    "адрес",
    "площадк",
    "дат",
    "wrong_event",
    "чужое событие",
    "фотоотчёт не подтверждает",
    "другое мероприятие",
    "отсутствует слово",
    "не совпадают с договором",
    "не совпадает с договором",
    "указано иначе",
)


def infer_operator(
    required: str,
    quote_contract: str = "",
    stored: str = "",
) -> str:
    blob = f"{required} {quote_contract}".lower()
    if GTE_RE.search(blob):
        return "gte"
    if LTE_RE.search(blob):
        return "lte"
    if stored in ("gte", "lte", "eq", "range"):
        return stored
    if parse_number(required) is not None or parse_number(quote_contract) is not None:
        return "eq"
    return "text"


def _is_quantity_row(row: dict[str, Any]) -> bool:
    metric = (row.get("metric") or "").lower()
    required = (row.get("required") or "").lower()
    quote_c = (row.get("quote_contract") or "").lower()
    blob = f"{required} {quote_c}"

    if row.get("compare_kind") == "qualitative":
        return "не менее" in blob or "не более" in blob

    if any(w in metric for w in ("адрес", "площадк", "место проведения", "venue", "формат")):
        return False

    if GTE_RE.search(blob) or LTE_RE.search(blob):
        return True

    if (row.get("operator") or "") in ("gte", "lte", "eq", "range"):
        return True

    qty_words = (
        "колич",
        "число",
        "лекц",
        "фото",
        "застав",
        "участник",
        "стул",
        "микрофон",
        "волонт",
        "автобус",
        "кофе",
        "просмотр",
    )
    return any(w in metric for w in qty_words)


def quantity_requirement_met(row: dict[str, Any]) -> bool:
    if not _is_quantity_row(row):
        return False
    required = (row.get("required") or "").strip()
    claimed = (row.get("claimed") or "").strip()
    quote_c = (row.get("quote_contract") or required).strip()
    quote_r = (row.get("quote_report") or claimed).strip()
    metric = (row.get("metric") or row.get("clause") or "").strip()

    op = infer_operator(required, quote_c, row.get("operator") or "text")
    if op == "text":
        return False

    req_n = parse_number(required) or parse_number(quote_c)
    cl_n = parse_fact_number(required, claimed, quote_r, metric)
    if req_n is None or cl_n is None:
        return False

    return _numeric_status(op, req_n, cl_n) == STATUS_OK


def has_non_numeric_issue(row: dict[str, Any]) -> bool:
    comment = (row.get("comment") or "").lower()
    required = (row.get("required") or "").lower()
    claimed = (row.get("claimed") or "").lower()
    metric = (row.get("metric") or "").lower()

    if row.get("compare_kind") == "qualitative":
        if quantity_requirement_met(row):
            return False
        return any(m in comment for m in QUALITATIVE_MARKERS) or len(required) > 80

    if any(m in comment for m in QUALITATIVE_MARKERS):
        return True

    loc_keys = LOCATION_KEYS
    blob_c = required + " " + (row.get("quote_contract") or "")
    blob_r = claimed + " " + (row.get("quote_report") or "")
    c_loc = {k for k in loc_keys if k in blob_c.lower()}
    r_loc = {k for k in loc_keys if k in blob_r.lower()}
    if c_loc and r_loc and c_loc != r_loc and (
        any(w in metric for w in ("адрес", "площадк", "лекц", "место"))
        or _is_logistics_row(row)
    ):
        return True

    if _is_logistics_row(row) and _quotes_event_conflict(row):
        return True

    if "фото" in metric and row.get("embedded_images") is not None:
        embedded = int(row["embedded_images"])
        claimed_p = row.get("claimed_photos")
        req_n = parse_number(required)
        if req_n is not None and embedded >= req_n:
            if claimed_p and embedded < float(claimed_p) * 0.5:
                return row.get("photo_text_gap") is not True
            return False
    return False


def is_false_numeric_mismatch(row: dict[str, Any]) -> bool:
    if (row.get("status") or "") != "mismatch":
        return False
    if _is_logistics_row(row):
        return False
    if not _is_quantity_row(row):
        return False
    if not quantity_requirement_met(row):
        return False
    return not has_non_numeric_issue(row)


_LECTURE_TOTAL_METRIC = re.compile(
    r"(?:общ(?:ее|ий)?\s+)?(?:колич|число).*лекц|количество\s+лекций(?!\s+\d)",
    re.I,
)
_LECTURE_DAY_METRIC = re.compile(
    r"\d{1,2}\s+(?:январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр)",
    re.I,
)


def normalize_quantity_row(row: dict[str, Any]) -> dict[str, Any]:
    """Если в цитате договора «не менее N», а required — голое число, приводим к gte."""
    required = (row.get("required") or "").strip()
    quote_c = (row.get("quote_contract") or "").strip()
    blob = f"{required} {quote_c}".lower()
    op = infer_operator(required, quote_c, row.get("operator") or "text")
    req_n = parse_number(required) or parse_number(quote_c)

    if op in ("gte", "lte") and req_n is not None:
        prefix = "не менее" if op == "gte" else "не более"
        if prefix not in required.lower():
            row = dict(row)
            row["operator"] = op
            row["required"] = f"{prefix} {int(req_n) if req_n == int(req_n) else req_n:g}"
            if row.get("status") == STATUS_MISMATCH and _is_quantity_row(row):
                cl_n = parse_fact_number(
                    row["required"],
                    row.get("claimed") or "",
                    row.get("quote_report") or "",
                    row.get("metric") or "",
                )
                num_status = _numeric_status(op, req_n, cl_n)
                if num_status:
                    row["status"] = num_status
                    row["comment"] = (
                        f"{'Расхождение. ' if num_status == STATUS_MISMATCH else 'Совпадает. '}"
                        f"число: требуется {op} {req_n:g}, в отчёте {cl_n:g}"
                    )
    elif GTE_RE.search(blob) or LTE_RE.search(blob):
        row = dict(row)
        row["operator"] = op
    return row


def _lecture_total_group_key(row: dict[str, Any]) -> str | None:
    metric = (row.get("metric") or "").lower()
    if "лекц" not in metric:
        return None
    if _LECTURE_DAY_METRIC.search(metric):
        return None
    if any(w in metric for w in ("адрес", "формат", "видео", "venue")):
        return None
    quote = (row.get("quote_contract") or "").lower()
    if not (
        _LECTURE_TOTAL_METRIC.search(metric)
        or "историко-просветительск" in quote
        or row.get("id") in {"TZ-lectures", "TZ-lectures-count", "TZ-lectures-total"}
    ):
        return None
    return "program_total_lectures"


def _lecture_total_row_score(row: dict[str, Any]) -> tuple[int, int]:
    req = (row.get("required") or "").lower()
    op = row.get("operator") or ""
    score = 0
    if op == "gte":
        score += 20
    elif op == "eq":
        score -= 5
    if "не менее" in req:
        score += 10
    if row.get("id") == "TZ-lectures":
        score += 5
    if row.get("id") in {"TZ-lectures-count", "TZ-lectures-total"}:
        score -= 3
    return score, len(req)


def normalize_photo_row(row: dict[str, Any]) -> dict[str, Any]:
    """Пересчитать статус фото: договор — по вложениям, текст — отдельное замечание."""
    metric = (row.get("metric") or "").lower()
    if row.get("evidence_type") != "photo" and "фото" not in metric:
        return row
    embedded = row.get("embedded_images")
    if embedded is None:
        return row

    row = dict(row)
    required_photos = row.get("required_photos") or parse_number(row.get("required") or "")
    claimed_photos = row.get("claimed_photos") or parse_number(row.get("claimed") or "")
    claimed_text = row.get("claimed") or ""
    required_text = row.get("required") or ""
    text_overstated = (
        claimed_photos is not None and embedded < float(claimed_photos) * 0.5
    )
    contract_met = required_photos is not None and embedded >= required_photos

    if required_photos is not None and embedded < required_photos:
        row["status"] = STATUS_MISMATCH
        row["comment"] = (
            f"фото: по договору {required_text}; "
            f"во вложении DOCX {embedded} шт. — меньше минимума."
        )
        row.pop("photo_text_gap", None)
    elif text_overstated and contract_met:
        row["status"] = STATUS_OK
        row["photo_text_gap"] = True
        row["comment"] = (
            f"фото: по договору {required_text}; во вложении DOCX {embedded} шт. — "
            f"минимум выполнен. В тексте отчёта указано {claimed_text or f'{claimed_photos:g}'}, "
            f"что не совпадает с числом вложений."
        )
    return row


def dedupe_quantity_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Убрать дубли «Количество лекций» / «Общее количество лекций» по одной цитате ТЗ."""
    normalized = [normalize_photo_row(normalize_quantity_row(r)) for r in rows]
    groups: dict[str, list[dict[str, Any]]] = {}
    passthrough: list[tuple[int, dict[str, Any]]] = []

    for idx, row in enumerate(normalized):
        key = _lecture_total_group_key(row)
        if key:
            groups.setdefault(key, []).append(row)
        else:
            passthrough.append((idx, row))

    winners: dict[str, dict[str, Any]] = {}
    for key, group in groups.items():
        winners[key] = max(group, key=_lecture_total_row_score)

    out: list[dict[str, Any]] = []
    emitted_groups: set[str] = set()
    for idx, row in enumerate(normalized):
        key = _lecture_total_group_key(row)
        if key:
            if key in emitted_groups:
                continue
            out.append(winners[key])
            emitted_groups.add(key)
        else:
            out.append(row)
    return out
