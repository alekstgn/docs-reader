"""Этап C: извлечение вложений DOCX, сжатие JPEG и умная выборка кадров."""

from __future__ import annotations

import io
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps, UnidentifiedImageError

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

IMAGE_EXTS = {".jpeg", ".jpg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
DATE_RE = re.compile(
    r"\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})\b|"
    r"\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2})(?!\d)",
    re.IGNORECASE,
)
CAPTION_HINTS = (
    "застав",
    "баннер",
    "пресс",
    "логотип",
    "микрофон",
    "сцен",
    "зал",
    "экран",
    "проектор",
    "аудитор",
    "открыт",
)

MIN_BYTES = 40_000
MIN_EDGE = 220
MIN_PIXELS = 70_000
MAX_EDGE = 1280
JPEG_QUALITY = 70


@dataclass
class PhotoCandidate:
    photo_id: str
    zip_name: str
    filename: str
    nbytes: int
    width: int
    height: int
    doc_index: int
    caption: str = ""
    date_hint: str = ""
    sample_reason: str = ""
    jpeg_name: str = ""
    jpeg_bytes: int = 0

    @property
    def aspect(self) -> float:
        return self.width / self.height if self.height else 1.0


def _paragraph_text(p: ET.Element) -> str:
    parts: list[str] = []
    for node in p.iter(f"{W_NS}t"):
        if node.text:
            parts.append(node.text)
        if node.tail:
            parts.append(node.tail)
    return " ".join("".join(parts).split())


def _load_rels(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
    except KeyError:
        return {}
    out: dict[str, str] = {}
    for rel in root:
        rid = rel.attrib.get("Id") or rel.attrib.get("id")
        target = rel.attrib.get("Target") or ""
        if not rid or not target:
            continue
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = "word/" + target.lstrip("/")
        out[rid] = path.replace("\\", "/")
    return out


def _blip_ids(elem: ET.Element) -> list[str]:
    found: list[str] = []
    for blip in elem.iter(f"{A_NS}blip"):
        rid = blip.attrib.get(f"{R_NS}embed") or blip.attrib.get("embed")
        if rid:
            found.append(rid)
    return found


def _date_hint(text: str) -> str:
    match = DATE_RE.search(text or "")
    return match.group(0) if match else ""


def iter_docx_images(docx_path: Path) -> list[dict[str, Any]]:
    """Порядок появления картинок в document.xml + ближайшая подпись."""
    with zipfile.ZipFile(docx_path) as zf:
        rels = _load_rels(zf)
        root = ET.fromstring(zf.read("word/document.xml"))
        body = root.find(f"{W_NS}body")
        if body is None:
            return []
        last_text = ""
        pending: list[dict[str, Any]] = []
        ordered: list[dict[str, Any]] = []
        idx = 0
        for p in body.iter(f"{W_NS}p"):
            text = _paragraph_text(p)
            rids = _blip_ids(p)
            if rids:
                caption = text or last_text
                for rid in rids:
                    zip_name = rels.get(rid, "")
                    item = {
                        "rid": rid,
                        "zip_name": zip_name,
                        "doc_index": idx,
                        "caption": caption[:240],
                    }
                    ordered.append(item)
                    pending.append(item)
                    idx += 1
            elif text:
                if pending and (DATE_RE.search(text) or len(text) < 80):
                    for item in pending:
                        if not item["caption"] or item["caption"] == last_text:
                            item["caption"] = text[:240]
                        elif DATE_RE.search(text) and not _date_hint(item["caption"]):
                            item["caption"] = (item["caption"] + " | " + text)[:240]
                    pending = []
                last_text = text
        return ordered


def _open_image(data: bytes) -> Image.Image | None:
    try:
        im = Image.open(io.BytesIO(data))
        im = ImageOps.exif_transpose(im)
        im.load()
        return im
    except (UnidentifiedImageError, OSError, ValueError):
        return None


def _exif_date(im: Image.Image) -> str:
    try:
        exif = im.getexif()
        if not exif:
            return ""
        for tag in (36867, 306):  # DateTimeOriginal, DateTime
            value = exif.get(tag)
            if value:
                return str(value)[:19]
    except Exception:
        return ""
    return ""


def inventory_images(docx_path: Path) -> list[PhotoCandidate]:
    """Читает вложения из zip без записи всех сырых файлов на диск."""
    order = iter_docx_images(docx_path)
    by_zip: dict[str, list[dict[str, Any]]] = {}
    for item in order:
        by_zip.setdefault(item["zip_name"], []).append(item)

    candidates: list[PhotoCandidate] = []
    seen: set[str] = set()
    with zipfile.ZipFile(docx_path) as zf:
        media_names = [
            n
            for n in zf.namelist()
            if n.startswith("word/media/")
            and Path(n).suffix.lower() in IMAGE_EXTS
        ]
        fallback_index = 0
        for name in media_names:
            info = zf.getinfo(name)
            data = zf.read(name)
            im = _open_image(data)
            if im is None:
                continue
            width, height = im.size
            meta_list = by_zip.get(name) or []
            if meta_list:
                meta = meta_list.pop(0)
                doc_index = int(meta["doc_index"])
                caption = meta.get("caption") or ""
            else:
                doc_index = 10_000 + fallback_index
                caption = ""
                fallback_index += 1
            date_hint = _date_hint(caption) or _exif_date(im)
            photo_id = Path(name).name
            if photo_id in seen:
                photo_id = f"{Path(name).stem}_{doc_index}{Path(name).suffix}"
            seen.add(photo_id)
            candidates.append(
                PhotoCandidate(
                    photo_id=photo_id,
                    zip_name=name,
                    filename=Path(name).name,
                    nbytes=info.file_size,
                    width=width,
                    height=height,
                    doc_index=doc_index,
                    caption=caption,
                    date_hint=date_hint,
                )
            )
            im.close()
    candidates.sort(key=lambda p: p.doc_index)
    return candidates


def is_real_photo(p: PhotoCandidate) -> bool:
    if p.nbytes < MIN_BYTES:
        return False
    if min(p.width, p.height) < MIN_EDGE:
        return False
    if p.width * p.height < MIN_PIXELS and p.nbytes < 120_000:
        return False
    return True


def sample_photos(candidates: list[PhotoCandidate], n: int = 20) -> list[PhotoCandidate]:
    n = max(15, min(25, n))
    pool = [p for p in candidates if is_real_photo(p)]
    if len(pool) < n:
        have = {p.photo_id for p in pool}
        extra = [
            p
            for p in sorted(candidates, key=lambda x: x.nbytes, reverse=True)
            if p.photo_id not in have and p.nbytes >= 25_000 and min(p.width, p.height) >= 180
        ]
        pool = pool + extra[: n - len(pool)]
    if len(pool) <= n:
        for p in pool:
            p.sample_reason = p.sample_reason or "all_real"
        return pool
    selected: list[PhotoCandidate] = []
    seen: set[str] = set()

    def take(item: PhotoCandidate, reason: str) -> None:
        if item.photo_id in seen or len(selected) >= n:
            return
        seen.add(item.photo_id)
        copy = item
        copy.sample_reason = reason
        selected.append(copy)

    by_date: dict[str, list[PhotoCandidate]] = {}
    for p in pool:
        by_date.setdefault(p.date_hint or "_", []).append(p)
    for date, items in by_date.items():
        if date == "_":
            continue
        take(max(items, key=lambda x: x.nbytes), f"date:{date}")

    ordered = sorted(pool, key=lambda p: p.doc_index)
    steps = min(8, len(ordered))
    if ordered:
        for i in range(steps):
            idx = int(i * (len(ordered) - 1) / max(steps - 1, 1))
            take(ordered[idx], "spread")

    banners = [
        p
        for p in pool
        if p.aspect >= 1.7 or p.aspect <= 0.6 or any(h in (p.caption or "").lower() for h in CAPTION_HINTS)
    ]
    for p in sorted(banners, key=lambda x: x.nbytes, reverse=True):
        take(p, "branding_or_caption")

    for p in sorted(pool, key=lambda x: x.nbytes, reverse=True):
        take(p, "large")
        if len(selected) >= n:
            break
    return sorted(selected, key=lambda p: p.doc_index)[:n]


def compress_jpeg(data: bytes, max_edge: int = MAX_EDGE, quality: int = JPEG_QUALITY) -> bytes:
    im = _open_image(data)
    if im is None:
        raise RuntimeError("Не удалось открыть изображение для сжатия")
    rgb = im.convert("RGB")
    im.close()
    w, h = rgb.size
    scale = max_edge / max(w, h)
    if scale < 1:
        resized = rgb.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)
        rgb.close()
        rgb = resized
    buf = io.BytesIO()
    rgb.save(buf, format="JPEG", quality=quality, optimize=True)
    out = buf.getvalue()
    if len(out) > 3_500_000:
        buf = io.BytesIO()
        rgb.save(buf, format="JPEG", quality=55, optimize=True)
        out = buf.getvalue()
    rgb.close()
    return out


def select_photos(
    candidates: list[PhotoCandidate],
    n: int = 20,
    *,
    all_real: bool = False,
) -> list[PhotoCandidate]:
    if all_real:
        pool = [p for p in candidates if is_real_photo(p)]
        for p in pool:
            p.sample_reason = "all_real"
        return pool
    return sample_photos(candidates, n)


def extract_sample(
    docx_path: Path,
    dest_dir: Path,
    n: int = 20,
    *,
    all_real: bool = False,
) -> dict[str, Any]:
    dest_dir.mkdir(parents=True, exist_ok=True)
    jpeg_dir = dest_dir / "jpeg"
    jpeg_dir.mkdir(parents=True, exist_ok=True)
    all_photos = inventory_images(docx_path)
    sampled = select_photos(all_photos, n, all_real=all_real)
    total = len(sampled)
    with zipfile.ZipFile(docx_path) as zf:
        for i, photo in enumerate(sampled, 1):
            jpeg_name = Path(photo.photo_id).stem + ".jpg"
            path = jpeg_dir / jpeg_name
            if path.is_file() and path.stat().st_size > 2_000:
                photo.jpeg_name = jpeg_name
                photo.jpeg_bytes = path.stat().st_size
                continue
            raw = zf.read(photo.zip_name)
            jpeg = compress_jpeg(raw, quality=JPEG_QUALITY)
            if len(jpeg) > 2_500_000:
                jpeg = compress_jpeg(raw, max_edge=1024, quality=55)
            path.write_bytes(jpeg)
            photo.jpeg_name = jpeg_name
            photo.jpeg_bytes = len(jpeg)
            if i % 25 == 0 or i == total:
                print(f"    сжатие JPEG {i}/{total}", flush=True)
    real_n = sum(1 for p in all_photos if is_real_photo(p))
    manifest = {
        "source": str(docx_path.name),
        "extracted_total": len(all_photos),
        "real_photo_candidates": real_n,
        "sampled": len(sampled),
        "skipped_tiny": len(all_photos) - real_n,
        "coverage": "all_real" if all_real else "sample",
        "photos": [asdict(p) for p in sampled],
        "all_filenames": [p.filename for p in all_photos],
    }
    (dest_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return {"manifest": manifest, "sampled": sampled, "all": all_photos}


DOCUMENT_PHOTO_MARKERS = (
    "билет",
    "ticket",
    "квитан",
    "чек",
    "наклад",
    "проезд",
    "самолет",
    "самолёт",
    "авиа",
    "taxi",
    "такси",
    "маршрут",
    "аэрофлот",
)


def photo_analysis_blob(item: dict[str, Any]) -> str:
    return " ".join(
        [
            str(item.get("notes") or ""),
            str(item.get("branding_or_text_seen") or ""),
            str(item.get("scene_type") or ""),
            " ".join(str(x) for x in (item.get("visible_objects") or [])),
        ]
    ).lower()


def is_supporting_document_photo(item: dict[str, Any]) -> bool:
    """Билеты, чеки, проезд — не фото зала; для отчёта это нормальные приложения."""
    return any(marker in photo_analysis_blob(item) for marker in DOCUMENT_PHOTO_MARKERS)
