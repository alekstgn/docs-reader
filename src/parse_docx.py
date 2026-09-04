"""Локальный разбор DOCX: текст и инвентарь вложений. Исходный файл в API не отправляется."""

from __future__ import annotations

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
IMAGE_EXTS = {".jpeg", ".jpg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".emf", ".wmf"}

TZ_HEADING_RE = re.compile(
    r"(?m)^Приложение\s*№\s*1\s*$",
    re.IGNORECASE,
)
TZ_TECH_RE = re.compile(
    r"(?m)^Техническое задание\s*$",
    re.IGNORECASE,
)
TZ_END_RE = re.compile(
    r"(?m)^Приложение\s*№\s*2\b",
    re.IGNORECASE,
)


@dataclass
class MediaItem:
    name: str
    size: int
    ext: str


@dataclass
class ParsedDoc:
    path: str
    file_size: int
    text: str
    paragraphs: list[str] = field(default_factory=list)
    media: list[MediaItem] = field(default_factory=list)

    @property
    def image_count(self) -> int:
        return sum(1 for m in self.media if m.ext in IMAGE_EXTS)

    @property
    def media_bytes(self) -> int:
        return sum(m.size for m in self.media)


def _paragraph_text(p: ET.Element) -> str:
    parts: list[str] = []
    for node in p.iter(f"{W_NS}t"):
        if node.text:
            parts.append(node.text)
        if node.tail:
            parts.append(node.tail)
    return "".join(parts)


def parse_docx(path: str) -> ParsedDoc:
    file_size = os.path.getsize(path)
    with zipfile.ZipFile(path) as zf:
        media: list[MediaItem] = []
        for name in zf.namelist():
            if not name.startswith("word/media/"):
                continue
            info = zf.getinfo(name)
            basename = os.path.basename(name)
            media.append(
                MediaItem(
                    name=basename,
                    size=info.file_size,
                    ext=os.path.splitext(basename)[1].lower(),
                )
            )
        root = ET.fromstring(zf.read("word/document.xml"))

    paragraphs: list[str] = []
    for p in root.iter(f"{W_NS}p"):
        text = _paragraph_text(p).strip()
        if text:
            paragraphs.append(text)

    return ParsedDoc(
        path=path,
        file_size=file_size,
        text="\n".join(paragraphs),
        paragraphs=paragraphs,
        media=media,
    )


def extract_tz(text: str) -> str:
    """Текст приложения № 1 (ТЗ), а не первое упоминание «приложение № 1» в договоре."""
    match = TZ_HEADING_RE.search(text)
    if not match:
        match = TZ_TECH_RE.search(text)
    if not match:
        return text
    body = text[match.start() :]
    end = TZ_END_RE.search(body)
    if end:
        body = body[: end.start()]
    return body.strip()


def extract_contract_header(text: str) -> str:
    """Предмет договора: название, даты, площадки (п. 1.1–1.3) и норма фотоотчёта."""
    lines = text.splitlines()
    kept: list[str] = []
    for line in lines:
        if re.match(r"^\s*2[\.\s]", line):
            break
        kept.append(line)
        if len("\n".join(kept)) > 4500:
            break
    header = "\n".join(kept)
    photo_clause = re.search(
        r"4\.1\.6\..{0,900}?фотоотчет.{0,250}",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    extra = "\n\n" + photo_clause.group(0) if photo_clause else ""
    return (header + extra).strip()


PHOTO_CAPTION_RE = re.compile(
    r"^(?:(?:\d{1,2}\.\d{2}\.\d{2,4}\s*(?:г\.?)?\s*)+|"
    r"(?:Анимированная заставка|Статичная заставка|\s)+)+$",
    re.IGNORECASE,
)
GLUED_DATE_RE = re.compile(r"(г\.)(\d{1,2}\.)")
GLUED_DATE2_RE = re.compile(r"(\d{2}\.\d{2}\.\d{2,4})(\d{2}\.\d{2}\.\d{2,4})")
HIDDEN_PII_RE = re.compile(r"Личная информация скрыта на время тестирования")


def clean_report_noise(text: str) -> str:
    """Убрать подписи дат с фото и склейки, чтобы не раздувать промпт."""
    text = GLUED_DATE_RE.sub(r"\1\n\2", text)
    text = GLUED_DATE2_RE.sub(r"\1\n\2", text)
    kept: list[str] = []
    hidden_run = 0
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if PHOTO_CAPTION_RE.match(stripped):
            continue
        if HIDDEN_PII_RE.search(stripped) and len(stripped) < 80:
            hidden_run += 1
            if hidden_run > 2:
                continue
        else:
            hidden_run = 0
        kept.append(stripped)
    return "\n".join(kept)
