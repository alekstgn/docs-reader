"""Привязка замечаний к абзацам и оценка номера страницы в DOCX."""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

from .issues import Issue
from .photos import W_NS, _blip_ids, _load_rels, _paragraph_text

PARAS_PER_PAGE = 28

LOCATE_STOPWORDS = frozenset(
    {
        "общее",
        "количество",
        "проведение",
        "программа",
        "предусматривала",
        "предусматривать",
        "должна",
        "договор",
        "отчёте",
        "отчета",
        "отчёт",
        "отчет",
        "пункт",
        "текст",
        "мероприятия",
        "мероприятие",
        "исполнитель",
        "обеспечил",
        "обеспечить",
        "указано",
        "найден",
        "найдено",
        "начало",
    }
)

METRIC_STEMS = (
    "лекц",
    "участник",
    "фото",
    "волонт",
    "стул",
    "микрофон",
    "автобус",
    "кофе",
    "заставк",
    "пленар",
    "мастер",
    "хакатон",
    "ток-шоу",
    "адрес",
    "площадк",
    "дискусс",
    "стратег",
    "просмотр",
    "видеол",
    "ночей",
    "прожив",
    "перевоз",
    "иногород",
    "ноутбук",
    "телевиз",
    "экран",
    "обед",
    "ужин",
    "питан",
)

_SECTION_HINTS: dict[str, list[str]] = {
    "мастер": ["мастер-класс", "мастер класс", "программ"],
    "лекц": ["лекц", "историко-просветительск", "программ"],
    "кофе": ["кофе-брейк", "кофе брейк", "8.2.2"],
    "обед": ["обед", "ужин", "питан", "8.2.1"],
    "прожив": ["проживан", "7.2", "8.3", "отел"],
    "перевоз": ["перевоз", "трансфер", "8.1"],
    "технич": ["6.2", "оборудован", "технич"],
    "ноутбук": ["6.2", "ноутбук", "оборудован"],
    "адрес": ["6.2", "программ", "место проведения", "площадк"],
}


@dataclass
class DocMap:
    docx_path: Path
    paragraph_indices: list[int] = field(default_factory=list)
    paragraph_texts: list[str] = field(default_factory=list)
    image_zip_to_para: dict[str, int] = field(default_factory=dict)
    page_for_para: dict[int, int] = field(default_factory=dict)
    approximate_pages: bool = True

    def page_label(self, para_index: int | None) -> str:
        if para_index is None:
            return "—"
        page = self.page_for_para.get(para_index, 1)
        prefix = "~" if self.approximate_pages else ""
        return f"{prefix}{page}"


def _normalize(text: str) -> str:
    text = text.replace("\xa0", " ").lower()
    text = text.replace("«", '"').replace("»", '"').replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _metric_stem(metric: str, location: str = "") -> str | None:
    blob = _normalize(f"{metric} {location}")
    for stem in METRIC_STEMS:
        if stem in blob:
            return stem
    return None


def _claimed_numbers(claimed: str) -> list[str]:
    return re.findall(r"\d+", claimed or "")


def _paragraph_has_number(norm_text: str, number: str) -> bool:
    padded = f" {norm_text} "
    return (
        f" {number} " in padded
        or f"({number})" in norm_text
        or f"{number} (" in norm_text
    )


def build_doc_map(docx_path: Path) -> DocMap:
    doc_map = DocMap(docx_path=docx_path)
    with zipfile.ZipFile(docx_path) as zf:
        rels = _load_rels(zf)
        root = ET.fromstring(zf.read("word/document.xml"))
        body = root.find(f"{W_NS}body")
        if body is None:
            return doc_map

        page = 1
        para_idx = 0
        has_rendered_breaks = False

        for child in body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                text = _paragraph_text(child)
                doc_map.paragraph_indices.append(para_idx)
                doc_map.paragraph_texts.append(text)
                for blip_rid in _blip_ids(child):
                    zip_name = rels.get(blip_rid, "")
                    if zip_name:
                        doc_map.image_zip_to_para[zip_name] = para_idx
                        base = zip_name.split("/")[-1]
                        doc_map.image_zip_to_para[base] = para_idx
                page_break = False
                if child.find(f".//{W_NS}lastRenderedPageBreak") is not None:
                    has_rendered_breaks = True
                    page_break = True
                else:
                    for br in child.iter(f"{W_NS}br"):
                        if br.get(f"{W_NS}type") == "page":
                            page_break = True
                            break
                if page_break:
                    doc_map.page_for_para[para_idx] = page
                    page += 1
                else:
                    doc_map.page_for_para[para_idx] = page
                para_idx += 1
            elif tag == "tbl":
                for p in child.iter(f"{W_NS}p"):
                    text = _paragraph_text(p)
                    doc_map.paragraph_indices.append(para_idx)
                    doc_map.paragraph_texts.append(text)
                    for blip_rid in _blip_ids(p):
                        zip_name = rels.get(blip_rid, "")
                        if zip_name:
                            doc_map.image_zip_to_para[zip_name] = para_idx
                            doc_map.image_zip_to_para[zip_name.split("/")[-1]] = para_idx
                    doc_map.page_for_para[para_idx] = page
                    para_idx += 1

        doc_map.approximate_pages = not has_rendered_breaks
        if doc_map.approximate_pages:
            for i in range(len(doc_map.paragraph_texts)):
                doc_map.page_for_para[i] = 1 + i // PARAS_PER_PAGE

    return doc_map


def _quote_needles(quote: str) -> list[str]:
    if not quote:
        return []
    norm = _normalize(quote)
    needles: list[str] = []
    seen: set[str] = set()

    def add(text: str) -> None:
        text = text.strip(" -;:")
        if len(text) >= 12 and text not in seen:
            seen.add(text)
            needles.append(text)

    add(norm)
    if len(norm) > 80:
        add(norm[:80])

    for part in re.split(r"[-–—;\n]", quote):
        add(_normalize(part))

    for match in re.finditer(
        r"\d+\s*\([^)]{3,40}\)[^.;]{8,120}",
        norm,
    ):
        add(match.group(0).strip())

    for match in re.finditer(r"[а-яёa-z]{5,}[^.;]{0,80}", norm):
        fragment = match.group(0).strip()
        if any(stem in fragment for stem in METRIC_STEMS):
            add(fragment)

    return needles


def _find_by_needle(doc_map: DocMap, needle: str) -> int | None:
    if len(needle) < 8:
        return None
    best: int | None = None
    best_len = 10**9
    for idx, text in enumerate(doc_map.paragraph_texts):
        norm = _normalize(text)
        if needle in norm and len(norm) < best_len:
            best = idx
            best_len = len(norm)
    return best


def _find_by_quote(doc_map: DocMap, quote: str, claimed: str = "") -> int | None:
    needles = _quote_needles(quote)
    needles.sort(
        key=lambda n: (
            0 if re.search(r"\d", n) else 1,
            0 if any(stem in n for stem in METRIC_STEMS) else 1,
            -len(n),
        )
    )

    candidates: list[int] = []
    seen: set[int] = set()
    for needle in needles:
        para = _find_by_needle(doc_map, needle)
        if para is not None and para not in seen:
            seen.add(para)
            candidates.append(para)

    if not candidates:
        return None

    numbers = _claimed_numbers(claimed)
    if numbers:
        for para in candidates:
            norm = _normalize(doc_map.paragraph_texts[para])
            if any(_paragraph_has_number(norm, n) for n in numbers):
                return para

    return min(candidates, key=lambda p: len(doc_map.paragraph_texts[p]))


def _find_by_metric_value(
    doc_map: DocMap,
    metric: str,
    claimed: str,
    location: str,
) -> int | None:
    stem = _metric_stem(metric, location)
    numbers = _claimed_numbers(claimed)
    if not stem and not numbers:
        return None

    best: int | None = None
    best_len = 10**9
    for idx, text in enumerate(doc_map.paragraph_texts):
        norm = _normalize(text)
        if not norm:
            continue
        if stem and stem not in norm:
            continue
        if numbers and not any(_paragraph_has_number(norm, n) for n in numbers):
            continue
        if len(norm) < best_len:
            best = idx
            best_len = len(norm)
    return best


def _find_by_keywords(doc_map: DocMap, location: str, metric: str = "") -> int | None:
    blob = f"{location} {metric}"
    tokens = [
        t
        for t in re.findall(r"[а-яёa-z0-9]{4,}", blob.lower())
        if t not in LOCATE_STOPWORDS
    ]
    stem = _metric_stem(metric, location)
    if stem and stem not in tokens:
        tokens.append(stem)
    if not tokens:
        return None

    best: int | None = None
    best_score = 0
    best_len = 10**9
    for idx, text in enumerate(doc_map.paragraph_texts):
        norm = _normalize(text)
        if not norm:
            continue
        score = sum(1 for t in tokens if t in norm)
        if score > best_score or (score == best_score and score > 0 and len(norm) < best_len):
            best_score = score
            best = idx
            best_len = len(norm)

    required = 2 if len(tokens) >= 2 else 1
    if stem:
        required = max(required, 1)
        if best is not None:
            norm = _normalize(doc_map.paragraph_texts[best])
            if stem not in norm:
                return None
    return best if best_score >= required else None


def _find_by_metric_date(doc_map: DocMap, metric: str) -> int | None:
    match = re.search(
        r"(\d{1,2})\s+(январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр)",
        metric,
        re.I,
    )
    if not match:
        return None
    day, month = match.group(1), match.group(2).lower()[:4]
    for idx, text in enumerate(doc_map.paragraph_texts):
        norm = _normalize(text)
        if not re.search(rf"\b{day}\b", norm):
            continue
        if month in norm or any(m in norm for m in ("июн", "июл", "нояб", "ноя", "январ")):
            return idx
    return None


def _find_section_for_missing(
    doc_map: DocMap,
    metric: str,
    location: str,
    required: str,
) -> int | None:
    stem = _metric_stem(metric, location)
    hints: list[str] = list(_SECTION_HINTS.get(stem or "", []))
    if stem and stem not in hints:
        hints.append(stem)
    blob = _normalize(f"{metric} {location} {required}")
    if "адрес" in blob or "место" in blob:
        hints.extend(_SECTION_HINTS["адрес"])
    if "стоимост" in blob or "руб" in blob:
        hints.extend(["руб", "стоимость", "8.2", "7.2"])
    for key in ("шереметьев", "лесной", "никитск", "музей", "вучетич"):
        if key in blob:
            hints.append(key)

    best: int | None = None
    best_score = 0
    best_len = 10**9
    for idx, text in enumerate(doc_map.paragraph_texts):
        norm = _normalize(text)
        if not norm:
            continue
        score = sum(1 for hint in hints if hint in norm)
        if score > best_score or (score == best_score and score > 0 and len(norm) < best_len):
            best_score = score
            best = idx
            best_len = len(norm)
    return best if best_score >= 1 else None


def _resolve_fallback_para(
    doc_map: DocMap,
    issue: Issue,
    assigned: set[int],
) -> int:
    metric = issue.metric or issue.location or ""
    required = issue.required or ""

    for finder in (
        lambda: _find_by_metric_date(doc_map, metric),
        lambda: _find_by_quote(doc_map, required, ""),
        lambda: _find_section_for_missing(doc_map, metric, issue.location, required),
        lambda: _find_by_keywords(doc_map, issue.location, metric),
    ):
        para = finder()
        if para is not None and para not in assigned:
            return para

    for idx, text in enumerate(doc_map.paragraph_texts):
        norm = _normalize(text)
        if not norm:
            continue
        if any(token in norm for token in ("программ", "6.2.", "8.2.", "7.2")):
            if idx not in assigned:
                return idx

    for idx in range(len(doc_map.paragraph_texts)):
        if idx not in assigned and doc_map.paragraph_texts[idx].strip():
            return idx
    return 0


def locate_issue(doc_map: DocMap, issue: Issue) -> int | None:
    if issue.anchor_first:
        for idx, text in enumerate(doc_map.paragraph_texts):
            if text.strip():
                return idx
        return 0

    if issue.kind == "photo" and issue.zip_name:
        para = doc_map.image_zip_to_para.get(issue.zip_name)
        if para is not None:
            return para
        if issue.photo_id:
            para = doc_map.image_zip_to_para.get(f"word/media/{issue.photo_id}")
            if para is not None:
                return para

    para = _find_by_quote(doc_map, issue.quote_report, issue.claimed)
    if para is not None:
        return para

    if not (issue.quote_report or issue.claimed).strip():
        para = _find_by_metric_date(doc_map, issue.metric or issue.location)
        if para is not None:
            return para
        para = _find_by_quote(doc_map, issue.required, "")
        if para is not None:
            return para
        para = _find_section_for_missing(
            doc_map,
            issue.metric or "",
            issue.location,
            issue.required,
        )
        if para is not None:
            return para

    para = _find_by_metric_value(
        doc_map,
        issue.metric or issue.location,
        issue.claimed,
        issue.location,
    )
    if para is not None:
        return para

    return _find_by_keywords(doc_map, issue.location, issue.metric)


def locate_issues(doc_map: DocMap, issues: list[Issue]) -> list[tuple[Issue, int | None, str]]:
    located: list[tuple[Issue, int | None, str]] = []
    for issue in issues:
        para = locate_issue(doc_map, issue)
        page = doc_map.page_label(para)
        located.append((issue, para, page))
    return located
