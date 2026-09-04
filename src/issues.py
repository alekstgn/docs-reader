"""Сбор списка замечаний из сохранённых JSON (без API)."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .humanize import humanize_event, humanize_photo, humanize_row, location_label, photo_label
from .numeric_rules import (
    dedupe_quantity_rows,
    is_false_numeric_mismatch,
    refine_comparison_rows,
    should_include_comment_row,
)
from .photos import is_supporting_document_photo

TEXT_STATUSES = {"mismatch", "missing_in_report", "event_mismatch"}
PHOTO_CONCLUSIONS = {"wrong_event", "contradicts"}
SKIP_ROW_PREFIXES = ("PHOTO-",)


def _should_skip_photo_issue(
    item: dict[str, Any],
    pair_id: str,
    *,
    same_event: bool,
) -> bool:
    if pair_id == "prosvetiteli":
        return False
    if not same_event:
        return False
    if is_supporting_document_photo(item):
        return True
    return False


@dataclass
class Issue:
    kind: str  # event | text | photo
    description: str
    location: str
    quote_report: str = ""
    metric: str = ""
    required: str = ""
    claimed: str = ""
    photo_id: str = ""
    photo_label: str = ""
    zip_name: str = ""
    doc_index: int | None = None
    anchor_first: bool = False
    dedupe_key: str = ""


def _load(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def _should_include_row(row: dict[str, Any], *, same_event: bool = True) -> bool:
    status = row.get("status") or ""
    if row.get("photo_text_gap") and status == "ok":
        return True
    if status not in TEXT_STATUSES:
        return False
    if is_false_numeric_mismatch(row):
        return False
    row_id = row.get("id") or ""
    if row_id.startswith(SKIP_ROW_PREFIXES) and status != "mismatch":
        return False
    return should_include_comment_row(row, same_event=same_event)


def _semantic_dedupe_key(issue: Issue) -> str:
    if issue.kind != "text":
        return issue.dedupe_key or issue.description[:120]
    metric = issue.metric.lower()
    if "лекц" in metric and not re.search(
        r"\d{1,2}\s+(?:январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр)",
        metric,
    ):
        if any(w in metric for w in ("общ", "колич", "число")) or metric.startswith("количество лекций"):
            nums = re.findall(r"\d+", f"{issue.required} {issue.claimed}")
            if nums:
                return f"lectures_total:{nums[-1]}"
    return issue.dedupe_key or issue.description[:120]


def _dedupe_issues(issues: list[Issue]) -> list[Issue]:
    seen: set[str] = set()
    out: list[Issue] = []
    for issue in issues:
        key = _semantic_dedupe_key(issue)
        if key in seen and issue.kind == "text":
            continue
        seen.add(key)
        out.append(issue)
    return out


def build_photo_issues(
    photos_data: dict[str, Any],
    pair_id: str,
    manifest: dict[str, Any] | None = None,
    *,
    same_event: bool = True,
) -> list[Issue]:
    analyses = photos_data.get("analyses") or []
    by_date: dict[str, list[dict[str, Any]]] = {}
    for item in analyses:
        conclusion = item.get("conclusion") or ""
        event_fit = item.get("event_fit") or ""
        if _should_skip_photo_issue(item, pair_id, same_event=same_event):
            continue
        if conclusion not in PHOTO_CONCLUSIONS and event_fit != "other_event":
            continue
        if pair_id != "prosvetiteli" and conclusion not in PHOTO_CONCLUSIONS:
            continue
        date = (item.get("date_hint") or "без даты").replace(" г.", "")
        by_date.setdefault(date, []).append(item)

    manifest_by_id: dict[str, dict[str, Any]] = {}
    if manifest:
        for p in manifest.get("photos") or []:
            manifest_by_id[p.get("photo_id") or ""] = p

    issues: list[Issue] = []
    for date, group in sorted(by_date.items()):
        first = group[0]
        photo_id = first.get("photo_id") or ""
        meta = manifest_by_id.get(photo_id, {})
        zip_name = meta.get("zip_name") or f"word/media/{photo_id}"
        doc_index = meta.get("doc_index")
        if doc_index is None:
            doc_index = first.get("doc_index")
        label = photo_label(
            date,
            first.get("scene_type") or "",
            first.get("notes") or "",
        )
        extra = len(group) - 1
        desc = humanize_photo(first, pair_id, extra_count=extra)
        issues.append(
            Issue(
                kind="photo",
                description=desc,
                location=f"фотоотчёт, {date}",
                photo_id=photo_id,
                photo_label=label,
                zip_name=zip_name,
                doc_index=doc_index,
                dedupe_key=f"photo:{date}:{desc[:80]}",
            )
        )
    return issues


def build_issues(
    pair_id: str,
    out_dir: Path,
    comparison: dict[str, Any] | None = None,
    photos: dict[str, Any] | None = None,
    manifest: dict[str, Any] | None = None,
) -> list[Issue]:
    comparison = comparison or _load(out_dir / "comparison.json")
    photos = photos or _load(out_dir / "photos.json")
    if manifest is None:
        manifest_path = out_dir.parent / "_photos" / pair_id / "manifest.json"
        manifest = _load(manifest_path) if manifest_path.is_file() else {}

    issues: list[Issue] = []

    event = comparison.get("event_check") or {}
    if not event.get("same_event", True):
        desc = humanize_event(event, pair_id)
        if desc:
            issues.append(
                Issue(
                    kind="event",
                    description=desc,
                    location="начало отчёта",
                    quote_report=(event.get("quote_report") or "")[:200],
                    anchor_first=True,
                    dedupe_key="event",
                )
            )

    same_event = bool(event.get("same_event", True))
    for row in dedupe_quantity_rows(
        refine_comparison_rows(
            comparison.get("rows") or [],
            comparison.get("event_check") or {},
        )
    ):
        if not _should_include_row(row, same_event=same_event):
            continue
        row_id = row.get("id") or ""
        if row_id.startswith(SKIP_ROW_PREFIXES):
            continue
        issues.append(
            Issue(
                kind="text",
                description=humanize_row(row),
                location=location_label(row),
                quote_report=(row.get("quote_report") or "").strip(),
                metric=(row.get("metric") or row.get("clause") or "").strip(),
                required=str(row.get("required") or "").strip(),
                claimed=str(row.get("claimed") or "").strip(),
                dedupe_key=f"text:{row_id}",
            )
        )

    issues.extend(
        build_photo_issues(
            photos,
            pair_id,
            manifest,
            same_event=bool(event.get("same_event", True)),
        )
    )
    return _dedupe_issues(issues)
