"""Вставка комментариев Word (OOXML) в копию отчёта DOCX."""

from __future__ import annotations

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .issues import Issue
from .locate import build_doc_map, locate_issues, _resolve_fallback_para

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

AUTHOR = "Сверка Б.Т."
COMMENT_SHADE_FILL = "FFF2CC"


def _is_comment_marker(elem: ET.Element) -> bool:
    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
    return tag in ("commentRangeStart", "commentRangeEnd", "commentReference")


def _annotate_paragraph(
    paragraph: ET.Element,
    comment_texts: list[str],
    start_id: int,
) -> tuple[list[tuple[str, str]], int]:
    """Вставить комментарии: pPr → starts → текст → ends → refs."""
    p_pr = paragraph.find(f"{W}pPr")
    content = [
        child
        for child in list(paragraph)
        if not _is_comment_marker(child) and child is not p_pr
    ]
    for child in list(paragraph):
        paragraph.remove(child)

    entries: list[tuple[str, str]] = []
    next_id = start_id

    if p_pr is not None:
        paragraph.append(p_pr)
    else:
        p_pr = ET.Element(f"{W}pPr")
        paragraph.insert(0, p_pr)

    for text in comment_texts:
        start = ET.Element(f"{W}commentRangeStart")
        start.set(f"{W}id", str(next_id))
        paragraph.append(start)
        entries.append((str(next_id), text))
        next_id += 1

    for child in content:
        paragraph.append(child)

    for cid, _ in entries:
        end = ET.Element(f"{W}commentRangeEnd")
        end.set(f"{W}id", cid)
        paragraph.append(end)
        ref_run = ET.Element(f"{W}r")
        ref = ET.SubElement(ref_run, f"{W}commentReference")
        ref.set(f"{W}id", cid)
        paragraph.append(ref_run)

    return entries, next_id


def _register_namespaces() -> None:
    ET.register_namespace("w", W_NS)


def _append_empty_comment_paragraph(comment: ET.Element) -> None:
    ET.SubElement(comment, f"{W}p")


def _append_comment_paragraph(comment: ET.Element, text: str) -> None:
    p = ET.SubElement(comment, f"{W}p")
    p_pr = ET.SubElement(p, f"{W}pPr")
    shd = ET.SubElement(p_pr, f"{W}shd")
    shd.set(f"{W}val", "clear")
    shd.set(f"{W}color", "auto")
    shd.set(f"{W}fill", COMMENT_SHADE_FILL)
    r = ET.SubElement(p, f"{W}r")
    t = ET.SubElement(r, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def _build_comments_xml(entries: list[tuple[str, str]]) -> bytes:
    comments_root = ET.Element(f"{W}comments")
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    for cid, text in entries:
        comment = ET.SubElement(comments_root, f"{W}comment")
        comment.set(f"{W}id", cid)
        comment.set(f"{W}author", AUTHOR)
        comment.set(f"{W}date", now)
        comment.set(f"{W}initials", "СР")
        paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text or "") if part.strip()]
        if not paragraphs:
            paragraphs = [""]
        for i, para_text in enumerate(paragraphs):
            if i > 0:
                _append_empty_comment_paragraph(comment)
            _append_comment_paragraph(comment, para_text)
    return ET.tostring(comments_root, encoding="utf-8", xml_declaration=True)


def _patch_document_xml(xml_bytes: bytes, para_comments: dict[int, list[str]]) -> tuple[bytes, list[tuple[str, str]]]:
    root = ET.fromstring(xml_bytes)
    body = root.find(f"{W}body")
    if body is None:
        return xml_bytes, []

    entries: list[tuple[str, str]] = []
    next_id = 0
    para_idx = 0

    def process_paragraph(p: ET.Element) -> None:
        nonlocal para_idx, next_id
        if para_idx in para_comments:
            new_entries, next_id = _annotate_paragraph(
                p, para_comments[para_idx], next_id
            )
            entries.extend(new_entries)
        para_idx += 1

    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            process_paragraph(child)
        elif tag == "tbl":
            for p in child.iter(f"{W}p"):
                process_paragraph(p)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), entries


def _patch_content_types(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    override_tag = "{%s}Override" % CT_NS
    for ov in root.findall(override_tag):
        if ov.get("PartName") == "/word/comments.xml":
            return xml_bytes
    ov = ET.SubElement(root, override_tag)
    ov.set("PartName", "/word/comments.xml")
    ov.set(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_document_rels(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    rel_tag = "{%s}Relationship" % REL_NS
    for rel in root.findall(rel_tag):
        if rel.get("Type") == COMMENTS_REL:
            return xml_bytes
    ids = []
    for rel in root.findall(rel_tag):
        rid = rel.get("Id") or ""
        m = re.match(r"rId(\d+)", rid)
        if m:
            ids.append(int(m.group(1)))
    next_id = max(ids, default=0) + 1
    rel = ET.SubElement(root, rel_tag)
    rel.set("Id", f"rId{next_id}")
    rel.set("Type", COMMENTS_REL)
    rel.set("Target", "comments.xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def annotate_report(
    report_path: Path,
    issues: list[Issue],
    output_path: Path,
) -> dict[str, Any]:
    _register_namespaces()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(report_path, output_path)

    doc_map = build_doc_map(report_path)
    located = locate_issues(doc_map, issues)

    para_comments: dict[int, list[str]] = {}
    assigned: set[int] = set()
    for issue, para_idx, _page in located:
        if para_idx is None:
            para_idx = _resolve_fallback_para(doc_map, issue, assigned)
        assigned.add(para_idx)
        para_comments.setdefault(para_idx, []).append(issue.description)

    with zipfile.ZipFile(report_path, "r") as zin:
        names = zin.namelist()
        doc_xml = zin.read("word/document.xml")
        ct_xml = zin.read("[Content_Types].xml")
        rels_xml = zin.read("word/_rels/document.xml.rels")
        skip = {
            "word/document.xml",
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/comments.xml",
        }
        other = {n: zin.read(n) for n in names if n not in skip}

    new_doc, entries = _patch_document_xml(doc_xml, para_comments)
    comments_xml = _build_comments_xml(entries)
    new_ct = _patch_content_types(ct_xml)
    new_rels = _patch_document_rels(rels_xml)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in other.items():
            zout.writestr(name, data)
        zout.writestr("word/document.xml", new_doc)
        zout.writestr("[Content_Types].xml", new_ct)
        zout.writestr("word/_rels/document.xml.rels", new_rels)
        zout.writestr("word/comments.xml", comments_xml)

    return {
        "output": str(output_path),
        "comments": len(entries),
        "issues": len(issues),
    }
