"""CLI: сверка двух пар договор/отчёт через gemini-2.5-flash-lite (ProxyAPI)."""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path
from typing import Any

from .compare import (
    compare_event,
    compare_obligation,
    dedupe_quantity_rows,
    merge_photo_results,
    merge_qualitative,
    refine_comparison_rows,
)
from .fallback import audit_obligations, ensure_key_obligations
from .gemini_client import GeminiClient, Usage
from .parse_docx import clean_report_noise, extract_contract_header, extract_tz, parse_docx
from .photos import extract_sample
from .redact import redact_names
from .render import render_markdown
from .schemas import STAGE_A_SCHEMA, STAGE_B_SCHEMA, STAGE_C_SCHEMA, STAGE_QUAL_SCHEMA

ROOT = Path(__file__).resolve().parent.parent

PAIRS = [
    {
        "id": "prosvetiteli",
        "contract": ROOT
        / "Dogovor_OOO_Milk_Edzhensi_Forum_Prosvetiteli_blogery_red_09-12-2025_3_FINAL.docx",
        "report": ROOT / "Otchet_Prosvetiteli.docx",
    },
    {
        "id": "lektoriy_kaluga",
        "contract": ROOT
        / "Dogovor_OOO_Milk_Ezhdensi_Lektoriy_Istoria_rf_Kaluga_red_03-07-2025_2_FINAL.docx",
        "report": ROOT / "Otchet2_Lektoriy_Istoria_rf_Kaluga_red_17_07.docx",
    },
]

STAGE_A_INSTRUCTIONS = """Ты извлекаешь проверяемые обязательства из ТЕХНИЧЕСКОГО ЗАДАНИЯ (приложение № 1).
Пункты 1.1–1.3 договора — только для идентичности мероприятия. НЕ ограничивайся п. 4.1.6 договора.

Верни JSON по схеме. Обязательно включи отдельные пункты:
1) название, даты, все площадки, число участников;
2) все «не менее N» / «не более N»: лекции, заставки, фото, волонтёры, оборудование, питание;
3) программу: число лекций и форматы (мастер-класс / хакатон / дискуссия / пленарка);
4) лекции ПО ДНЯМ с адресом каждой группы;
5) статичные заставки — именно КОЛИЧЕСТВО из п. 3.4 ТЗ;
6) фотоотчёт — минимум фотографий из п. 3.2 / 4.1.6;
7) разделы 6–9: техобеспечение (6.1, 6.2, 6.2.1), питание/кофе (7.2, 8.2), проживание и трансферы (8.1, 8.3), автобус (9, 9.1) — отдельным пунктом на каждое количество, адрес и дату.

id уникален для КАЖДОГО пункта (TZ-lectures, TZ-workshops, TZ-hackathon, TZ-backdrops, TZ-photos, TZ-participants, TZ-lectures-11jun-venue, TZ-mic-count, …). Не повторяй один id.
Для ОБЩЕГО числа лекций в программе — один пункт TZ-lectures с operator=gte и required «не менее N». НЕ создавай второй пункт «Общее количество лекций» с голым числом.
operator: gte=«не менее», lte=«не более», eq=точное число (только если в ТЗ нет «не менее/не более»), text=качество/адрес.
evidence_type: text | photo | external.
quote — дословная цитата ≤400 символов.
Не выдумывай. Не пропускай количества из ТЗ.
"""

STAGE_B_INSTRUCTIONS = """По чеклисту обязательств извлеки ФАКТИЧЕСКИЕ значения из содержательного отчёта.

Для КАЖДОГО id: found/claimed/quote.
claimed — число, адрес или формат ИЗ ОТЧЁТА, а не копия поля required.
Если отчёт про другое мероприятие — заполни event как в отчёте и сопоставь по смыслу (лекции к лекциям, фото к фото).
Примеры: договор «не менее 16 заставок» → claimed «22»; договор «не менее 100 фото» → claimed «1748».
quote — дословная цитата ≤180 символов. Если пункта нет — found=false, claimed="", quote="".
id в ответе копируй ТОЧНО из чеклиста, по одному claim на id.
"""

QUAL_INSTRUCTIONS = """Сравни качественные пункты (без надёжных чисел). Для каждого id верни status:
- ok — смысл требования выполнен;
- mismatch — есть противоречие (другой адрес, другой формат, другое содержание);
- missing_in_report — в отчёте нет;
- not_verifiable_from_docs — по этим текстам нельзя проверить.
Не смягчай расхождения адресов и форматов программы.
"""

STAGE_C_INSTRUCTIONS = """Ты проверяешь СЖАТЫЕ фото из содержательного отчёта на соответствие пунктам ТЗ.

Для КАЖДОГО кадра верни JSON по схеме. Смотри только на то, что видно:
- брендирование / пресс-волл / логотипы / текст на баннерах и экранах;
- зал и аудитория (без оценки лиц);
- сцена и техника (микрофоны, колонки, проектор, экран, заставка).

Правила:
1) Не называй людей и не описывай лица. Никаких ФИО. Пиши notes по-русски.
2) matched_obligation_ids — только id из чеклиста, если кадр реально относится к пункту.
3) event_fit: contracted_event если кадр похож на договорное мероприятие;
   other_event если текст/брендинг/обстановка другого события; unknown если нельзя сказать.
4) branding_or_text_seen — дословный видимый текст логотипов/баннеров без имён людей.
5) Не выдумывай оборудование, которого нет в кадре. Не считай точное число микрофонов, если не видно.
6) conclusion: confirms — кадр подтверждает пункт ТЗ именно договорного мероприятия;
   contradicts — видно противоречие ТЗ; inconclusive — нельзя сказать;
   wrong_event — кадр другого события (не того, что в договоре).
"""


def _uniquify_ids(obligations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: dict[str, int] = {}
    out: list[dict[str, Any]] = []
    for item in obligations:
        base = (item.get("id") or "TZ-item").strip() or "TZ-item"
        n = seen.get(base, 0)
        seen[base] = n + 1
        copy = dict(item)
        copy["id"] = base if n == 0 else f"{base}-{n + 1}"
        out.append(copy)
    return out


def _compact_obligations(obligations: list[dict[str, Any]], limit: int = 80) -> list[dict[str, Any]]:
    if len(obligations) <= limit:
        return obligations
    priority = (
        "лекц",
        "застав",
        "фото",
        "участник",
        "площадк",
        "хакатон",
        "мастер",
        "програм",
        "адрес",
        "дат",
        "venue",
        "мероприят",
        "оборудован",
        "стул",
        "микрофон",
        "volon",
        "волонт",
        "проживан",
        "перевозк",
        "трансфер",
        "автобус",
        "кофе",
        "питани",
        "технич",
        "6.2",
        "7.2",
        "8.1",
        "8.2",
        "8.3",
        "9.1",
        "вокзал",
        "формат",
    )

    def _protected(item: dict[str, Any]) -> bool:
        if item.get("source") == "regex_fallback":
            return True
        clause = item.get("clause") or ""
        if re.search(r"\b[6-9]\.", clause):
            return True
        blob = f"{item.get('id','')} {item.get('metric','')} {clause}".lower()
        return any(token in blob for token in priority)

    protected = [o for o in obligations if _protected(o)]
    collapsible = [o for o in obligations if not _protected(o)]
    if len(protected) + len(collapsible) <= limit:
        return _uniquify_ids(protected + collapsible)

    important: list[dict[str, Any]] = []
    rest: list[dict[str, Any]] = []
    for item in collapsible:
        blob = f"{item.get('id','')} {item.get('metric','')} {item.get('clause','')}".lower()
        if any(token in blob for token in priority):
            important.append(item)
        else:
            rest.append(item)

    room = max(0, limit - len(protected))
    kept_misc = important[:room]
    leftover = important[room:] + rest
    kept = list(protected) + kept_misc
    if leftover:
        preview = "; ".join(
            f"{o.get('metric')}: {o.get('required')}" for o in leftover[:25]
        )
        kept.append(
            {
                "id": "TZ-other-items",
                "clause": "прочие",
                "metric": f"Прочие обязательства ({len(leftover)} шт.)",
                "required": preview[:500],
                "unit": "сводка",
                "operator": "text",
                "evidence_type": "text",
                "quote": leftover[0].get("quote", "")[:400],
            }
        )
    return _uniquify_ids(kept)


def _stage_b(
    client: GeminiClient,
    event: dict[str, Any],
    obligations: list[dict[str, Any]],
    report: str,
    pair_id: str,
    batch_size: int = 12,
) -> dict[str, Any]:
    claims: list[dict[str, Any]] = []
    event_out = event
    for i in range(0, len(obligations), batch_size):
        chunk = obligations[i : i + batch_size]
        part = client.generate_json(
            _prompt_stage_b(event, chunk, report),
            STAGE_B_SCHEMA,
            stage=f"{pair_id}:B{i // batch_size + 1}",
            max_output_tokens=8192,
        )
        event_out = part.get("event") or event_out
        claims.extend(part.get("claims") or [])
    return {"event": event_out, "claims": claims}


def _prompt_stage_a(header: str, tz: str) -> str:
    return (
        f"{STAGE_A_INSTRUCTIONS}\n\n"
        f"=== ТЕХНИЧЕСКОЕ ЗАДАНИЕ (приложение № 1) — ГЛАВНЫЙ ИСТОЧНИК ===\n{tz}\n\n"
        f"=== ПРЕДМЕТ ДОГОВОРА (идентичность мероприятия) ===\n{header}\n"
    )


def _prompt_stage_b(event_hint: dict[str, Any], obligations: list[dict[str, Any]], report: str) -> str:
    checklist = [
        {
            "id": o["id"],
            "clause": o.get("clause"),
            "metric": o.get("metric"),
            "required": o.get("required"),
            "unit": o.get("unit"),
        }
        for o in obligations
    ]
    return (
        f"{STAGE_B_INSTRUCTIONS}\n\n"
        f"Событие по договору (для справки, не копируй если отчёт другой):\n"
        f"{json.dumps(event_hint, ensure_ascii=False)}\n\n"
        f"Чеклист обязательств:\n{json.dumps(checklist, ensure_ascii=False)}\n\n"
        f"=== ТЕКСТ ОТЧЁТА ===\n{report}\n"
    )


def _prompt_qual(pending: list[dict[str, Any]]) -> str:
    payload = [
        {
            "id": r["id"],
            "clause": r.get("clause"),
            "required": r.get("required"),
            "claimed": r.get("claimed"),
            "quote_contract": r.get("quote_contract"),
            "quote_report": r.get("quote_report"),
        }
        for r in pending
    ]
    return f"{QUAL_INSTRUCTIONS}\n\nПункты:\n{json.dumps(payload, ensure_ascii=False)}\n"


def _photo_checklist(obligations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    keys = (
        "фото",
        "застав",
        "пресс",
        "баннер",
        "логотип",
        "микрофон",
        "зал",
        "оборуд",
        "площад",
        "аудитор",
        "сцен",
        "проектор",
    )
    out: list[dict[str, Any]] = []
    for o in obligations:
        blob = f"{o.get('id','')} {o.get('metric','')} {o.get('clause','')} {o.get('required','')}".lower()
        if o.get("evidence_type") == "photo" or any(k in blob for k in keys):
            out.append(
                {
                    "id": o.get("id"),
                    "metric": o.get("metric"),
                    "required": (o.get("required") or "")[:180],
                }
            )
    return out[:18]


def _prompt_stage_c(
    pair: dict[str, Any],
    event_contract: dict[str, Any],
    event_report: dict[str, Any],
    same_event: bool,
    obligations: list[dict[str, Any]],
    batch: list[dict[str, Any]],
) -> str:
    extra = ""
    if pair["id"] == "prosvetiteli" or not same_event:
        extra = (
            "\nВАЖНО: договор и отчёт описывают РАЗНЫЕ мероприятия. "
            "Договор: форум «Просветители», 27–30 ноября 2025, Парк-отель «Шереметьевский» "
            "(и Музей военной формы). "
            "Отчёт: конференция «Просветители.Обществознание», 29 июня – 2 июля 2026, "
            "Парк-отель «Лесной». "
            "Если на кадре бренд/даты/площадка отчёта — event_fit=other_event. "
            "Эти фото НЕЛЬЗЯ считать подтверждением ТЗ ноября 2025.\n"
        )
    ids = [{"photo_id": p["photo_id"], "caption": p.get("caption", ""), "date_hint": p.get("date_hint", "")} for p in batch]
    return (
        f"{STAGE_C_INSTRUCTIONS}\n{extra}\n"
        f"Событие по договору:\n{json.dumps(event_contract, ensure_ascii=False)}\n\n"
        f"Событие по отчёту:\n{json.dumps(event_report, ensure_ascii=False)}\n\n"
        f"Чеклист пунктов ТЗ (для matched_obligation_ids):\n"
        f"{json.dumps(_photo_checklist(obligations), ensure_ascii=False)}\n\n"
        f"Кадры в этом запросе (photo_id копируй ТОЧНО):\n{json.dumps(ids, ensure_ascii=False)}\n"
    )


def _derive_conclusion(
    hit: dict[str, Any],
    pair: dict[str, Any],
    *,
    same_event: bool,
) -> str:
    allowed = {"confirms", "contradicts", "inconclusive", "wrong_event"}
    raw = str(hit.get("conclusion") or "").strip()
    fit = str(hit.get("event_fit") or "unknown")
    force_other = pair["id"] == "prosvetiteli" or not same_event
    if raw in allowed:
        if force_other and raw == "confirms":
            return "wrong_event"
        return raw
    if force_other or fit == "other_event":
        return "wrong_event"
    if fit == "contracted_event":
        return "confirms"
    return "inconclusive"


def _photo_record(photo: Any, hit: dict[str, Any], pair: dict[str, Any], same_event: bool, error: str = "") -> dict[str, Any]:
    event_fit = hit.get("event_fit") or ("unknown" if error else "unknown")
    rec = {
        "photo_id": photo.photo_id,
        "jpeg_name": photo.jpeg_name,
        "caption": photo.caption,
        "date_hint": photo.date_hint,
        "sample_reason": photo.sample_reason,
        "width": photo.width,
        "height": photo.height,
        "jpeg_bytes": photo.jpeg_bytes,
        "visible_objects": hit.get("visible_objects") or [],
        "matched_obligation_ids": hit.get("matched_obligation_ids") or [],
        "confidence": hit.get("confidence") or ("low" if error else "low"),
        "notes": hit.get("notes") or (f"Ошибка API: {error}" if error else ""),
        "event_fit": event_fit,
        "branding_or_text_seen": hit.get("branding_or_text_seen") or "",
        "scene_type": hit.get("scene_type") or "other",
        "conclusion": "inconclusive" if error else _derive_conclusion(hit, pair, same_event=same_event),
    }
    if error:
        rec["error"] = error[:800]
    return rec


def run_stage_c(
    client: GeminiClient,
    pair: dict[str, Any],
    out_dir: Path,
    *,
    sample_n: int = 20,
    batch_size: int = 2,
    all_photos: bool = False,
) -> dict[str, Any]:
    comparison_path = out_dir / "comparison.json"
    obligations_path = out_dir / "obligations.json"
    claims_path = out_dir / "claims.json"
    if not comparison_path.is_file() or not obligations_path.is_file():
        raise RuntimeError(f"Нет результатов A/B в {out_dir}. Сначала запустите текстовую сверку.")
    comparison = json.loads(comparison_path.read_text(encoding="utf-8"))
    stage_a = json.loads(obligations_path.read_text(encoding="utf-8"))
    stage_b = json.loads(claims_path.read_text(encoding="utf-8")) if claims_path.is_file() else {}
    photos_dir = ROOT / "output" / "_photos" / pair["id"]
    coverage = "all_real" if all_photos else "sample"
    print(
        f"  извлекаю {'все реальные' if all_photos else 'выборку'} фото из {pair['report'].name}…",
        flush=True,
    )
    pack = extract_sample(pair["report"], photos_dir, n=sample_n, all_real=all_photos)
    sampled = pack["sampled"]
    manifest = pack["manifest"]
    jpeg_dir = photos_dir / "jpeg"
    same_event = bool(comparison.get("same_event"))

    analyses: list[dict[str, Any]] = []
    done_ids: set[str] = set()
    photos_path = out_dir / "photos.json"
    prev_usage: dict[str, Any] = {}
    if all_photos and photos_path.is_file():
        prev = json.loads(photos_path.read_text(encoding="utf-8"))
        if prev.get("coverage") == "all_real" or (prev.get("sampled") or 0) > 25:
            for item in prev.get("analyses") or []:
                pid = str(item.get("photo_id") or "")
                if pid and not item.get("error"):
                    analyses.append(item)
                    done_ids.add(pid)
            prev_usage = prev.get("usage") or {}
            if done_ids:
                print(f"  продолжаю: уже готово {len(done_ids)} кадров без ошибок", flush=True)

    pending = [p for p in sampled if p.photo_id not in done_ids]
    current_batch = max(1, batch_size)
    i = 0
    batch_no = 0
    failed = 0
    extra_retries = 0
    usage_before = (client.usage.prompt_tokens, client.usage.output_tokens, client.usage.calls)

    def pair_usage_now() -> dict[str, Any]:
        new_details = client.usage.details[usage_before[2] :]
        inn = int(prev_usage.get("prompt_tokens") or 0) + (client.usage.prompt_tokens - usage_before[0])
        outt = int(prev_usage.get("output_tokens") or 0) + (client.usage.output_tokens - usage_before[1])
        return {
            "model": client.usage.as_dict().get("model"),
            "calls": int(prev_usage.get("calls") or 0) + (client.usage.calls - usage_before[2]),
            "prompt_tokens": inn,
            "output_tokens": outt,
            "cost_rub": round((inn * 26.0 + outt * 129.0) / 1_000_000, 4),
            "details": list(prev_usage.get("details") or []) + new_details,
        }

    def persist_partial() -> None:
        failed_n = sum(1 for a in analyses if a.get("error"))
        payload = {
            "pair_id": pair["id"],
            "extracted_total": manifest["extracted_total"],
            "real_photo_candidates": manifest["real_photo_candidates"],
            "sampled": len(sampled),
            "sent": len(analyses),
            "failed": failed_n,
            "skipped_tiny": manifest["skipped_tiny"],
            "coverage": coverage,
            "analyses": analyses,
            "usage": pair_usage_now(),
            "partial": True,
        }
        photos_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    while i < len(pending):
        chunk = pending[i : i + current_batch]
        orig_len = len(chunk)
        batch_no += 1
        images: list[tuple[str, bytes]] = []
        meta = []
        ready: list[Any] = []
        missing: list[Any] = []
        for photo in chunk:
            jpeg_path = jpeg_dir / photo.jpeg_name
            if not jpeg_path.is_file():
                missing.append(photo)
                continue
            data = jpeg_path.read_bytes()
            images.append(("image/jpeg", data))
            ready.append(photo)
            meta.append(
                {
                    "photo_id": photo.photo_id,
                    "caption": photo.caption,
                    "date_hint": photo.date_hint,
                    "jpeg_name": photo.jpeg_name,
                    "jpeg_bytes": photo.jpeg_bytes,
                }
            )

        def _record_missing() -> None:
            nonlocal failed
            for photo in missing:
                analyses.append(
                    _photo_record(photo, {}, pair, same_event, error=f"нет JPEG {photo.jpeg_name}")
                )
                failed += 1

        if not images:
            _record_missing()
            i += orig_len
            persist_partial()
            continue
        chunk = ready
        prompt = _prompt_stage_c(
            pair,
            stage_a.get("event") or {},
            stage_b.get("event") or comparison.get("event_check") or {},
            same_event,
            stage_a.get("obligations") or [],
            meta,
        )
        stage_name = f"{pair['id']}:C{batch_no}"
        done_now = len(analyses)
        print(
            f"  этап C {pair['id']} кадры {done_now + 1}–{done_now + len(chunk)} / {len(sampled)} "
            f"(ошибок {failed})…",
            flush=True,
        )
        try:
            part = client.generate_json_with_images(
                prompt,
                images,
                STAGE_C_SCHEMA,
                stage=stage_name,
                max_output_tokens=4096,
            )
            extra_retries = 0
        except RuntimeError as exc:
            msg = str(exc).lower()
            if current_batch > 1 and any(
                x in msg for x in ("400", "413", "size", "payload", "quota", "429", "too large")
            ):
                print(f"  уменьшаю пакет до 1 кадра", flush=True)
                current_batch = 1
                batch_no -= 1
                continue
            if extra_retries < 2 and any(x in msg for x in ("429", "unavailable", "502", "503", "504", "500")):
                extra_retries += 1
                wait = 4 * extra_retries
                print(f"  повтор через {wait} с ({exc.__class__.__name__})", flush=True)
                time.sleep(wait)
                batch_no -= 1
                continue
            err = str(exc)[:500]
            print(f"  ошибка пакета, продолжаю: {err[:180]}", flush=True)
            _record_missing()
            for photo in chunk:
                analyses.append(_photo_record(photo, {}, pair, same_event, error=err))
                failed += 1
            extra_retries = 0
            i += orig_len
            persist_partial()
            time.sleep(0.8)
            continue
        photos_out = part.get("photos") or []
        by_id = {str(p.get("photo_id") or ""): p for p in photos_out}
        _record_missing()
        for idx, photo in enumerate(chunk):
            hit = by_id.get(photo.photo_id) or by_id.get(photo.jpeg_name) or {}
            if not hit and len(photos_out) == len(chunk):
                hit = photos_out[idx]
            elif not hit and len(photos_out) == 1 and len(chunk) == 1:
                hit = photos_out[0]
            analyses.append(_photo_record(photo, hit, pair, same_event))
        i += orig_len
        extra_retries = 0
        persist_partial()
        time.sleep(0.35)

    pair_usage = pair_usage_now()
    failed_n = sum(1 for a in analyses if a.get("error"))
    photo_json = {
        "pair_id": pair["id"],
        "extracted_total": manifest["extracted_total"],
        "real_photo_candidates": manifest["real_photo_candidates"],
        "sampled": len(sampled),
        "sent": len(analyses),
        "failed": failed_n,
        "skipped_tiny": manifest["skipped_tiny"],
        "coverage": coverage,
        "analyses": analyses,
        "usage": pair_usage,
    }
    photos_path.write_text(json.dumps(photo_json, ensure_ascii=False, indent=2), encoding="utf-8")

    rows = list(comparison.get("rows") or [])
    rows = [r for r in rows if not str(r.get("id") or "").startswith("PHOTO-")]
    markers = ("Выборка этапа C", "Полный прогон этапа C", "Этап C:")
    for row in rows:
        comment = row.get("comment") or ""
        for marker in markers:
            if marker in comment:
                row["comment"] = comment.split(marker)[0].rstrip(" .")
                break
        if row.get("evidence_type") == "photo":
            row["photo_stage"] = "done"
    merge_photo_results(
        rows,
        comparison.get("event_check") or {},
        analyses,
        same_event=same_event,
        extracted=manifest["extracted_total"],
        sampled=len(sampled),
        pair_id=pair["id"],
        full_run=all_photos,
        failed=failed_n,
    )
    comparison["rows"] = rows
    comparison["photo_stage"] = "done"
    comparison["photos_extracted"] = manifest["extracted_total"]
    comparison["photos_sampled"] = len(sampled)
    comparison["photos_failed"] = failed_n
    comparison["photo_coverage"] = coverage
    comparison["photo_usage"] = {
        "model": pair_usage.get("model"),
        "calls": pair_usage.get("calls"),
        "prompt_tokens": pair_usage.get("prompt_tokens"),
        "output_tokens": pair_usage.get("output_tokens"),
        "cost_rub": pair_usage.get("cost_rub"),
    }
    comparison["mismatch_count"] = sum(
        1
        for r in [comparison.get("event_check") or {}, *rows]
        if r.get("status") in {"mismatch", "event_mismatch", "missing_in_report"}
    )
    comparison_path.write_text(json.dumps(comparison, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "discrepancies.md").write_text(
        render_markdown(pair["id"], comparison), encoding="utf-8"
    )
    print(
        f"  готово {pair['id']}: отправлено {len(analyses)}/{len(sampled)}, ошибок {failed_n}, "
        f"~{pair_usage['cost_rub']} ₽",
        flush=True,
    )
    return {**photo_json, "comparison": comparison}


def verify_pair(
    client: GeminiClient | None,
    pair: dict[str, Any],
    out_dir: Path,
    *,
    skip_llm: bool = False,
) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    contract = parse_docx(str(pair["contract"]))
    report = parse_docx(str(pair["report"]))

    header = redact_names(extract_contract_header(contract.text))
    tz = redact_names(extract_tz(contract.text))
    report_text = redact_names(clean_report_noise(report.text))

    (out_dir / "contract_redacted.txt").write_text(header + "\n\n" + tz, encoding="utf-8")
    (out_dir / "report_redacted.txt").write_text(report_text, encoding="utf-8")
    (out_dir / "media_inventory.json").write_text(
        json.dumps(
            {
                "contract_images": contract.image_count,
                "report_images": report.image_count,
                "report_media_bytes": report.media_bytes,
                "report_media": [
                    {"name": m.name, "size": m.size, "ext": m.ext} for m in report.media
                ],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    if skip_llm:
        stage_a = json.loads((out_dir / "obligations.json").read_text(encoding="utf-8"))
        stage_b = json.loads((out_dir / "claims.json").read_text(encoding="utf-8"))
        obligations = stage_a.get("obligations") or []
        prev_usage = {}
        if (out_dir / "comparison.json").is_file():
            prev_usage = json.loads((out_dir / "comparison.json").read_text(encoding="utf-8")).get("usage") or {}
    else:
        assert client is not None
        stage_a = client.generate_json(
            _prompt_stage_a(header, tz), STAGE_A_SCHEMA, stage=f"{pair['id']}:A"
        )
        raw_obligations = _uniquify_ids(
            ensure_key_obligations(tz, stage_a.get("obligations") or [])
        )
        (out_dir / "obligations_raw.json").write_text(
            json.dumps({**stage_a, "obligations": raw_obligations}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        obligations = _compact_obligations(raw_obligations)
        stage_a["obligations"] = obligations
        (out_dir / "obligations.json").write_text(
            json.dumps(stage_a, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        for warning in audit_obligations(tz, obligations):
            print(f"  ⚠ {warning}", flush=True)
        stage_b = _stage_b(
            client,
            stage_a.get("event") or {},
            obligations,
            report_text,
            pair["id"],
        )
        (out_dir / "claims.json").write_text(
            json.dumps(stage_b, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        prev_usage = None

    claims_by_id = {c.get("id"): c for c in (stage_b.get("claims") or [])}
    event_check = compare_event(stage_a.get("event") or {}, stage_b.get("event") or {})
    rows: list[dict[str, Any]] = []
    for obl in obligations:
        rows.append(
            compare_obligation(
                obl,
                claims_by_id.get(obl.get("id")),
                embedded_images=report.image_count,
            )
        )

    pending = [r for r in rows if r.get("status") == "qualitative_pending"]
    if pending and not skip_llm:
        qual = client.generate_json(
            _prompt_qual(pending), STAGE_QUAL_SCHEMA, stage=f"{pair['id']}:qual"
        )
        merge_qualitative(rows, qual.get("items") or [])
        (out_dir / "qualitative.json").write_text(
            json.dumps(qual, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    elif pending and (out_dir / "qualitative.json").is_file():
        qual = json.loads((out_dir / "qualitative.json").read_text(encoding="utf-8"))
        merge_qualitative(rows, qual.get("items") or [])

    rows = dedupe_quantity_rows(rows)

    rows = refine_comparison_rows(rows, event_check)

    usage_snapshot = prev_usage or ({} if client is None else client.usage.as_dict())
    # per-pair usage is the delta; store full snapshot, CLI will also write combined
    result = {
        "pair_id": pair["id"],
        "contract_file": pair["contract"].name,
        "report_file": pair["report"].name,
        "report_images": report.image_count,
        "report_media_mb": round(report.media_bytes / 1e6, 2),
        "photo_stage": "skipped",
        "event_check": event_check,
        "rows": rows,
        "usage": usage_snapshot,
        "same_event": event_check.get("same_event"),
        "mismatch_count": sum(
            1
            for r in [event_check, *rows]
            if r.get("status") in {"mismatch", "event_mismatch", "missing_in_report"}
        ),
    }
    (out_dir / "comparison.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "discrepancies.md").write_text(
        render_markdown(pair["id"], result), encoding="utf-8"
    )
    return result


SHEET_TITLES = {
    "prosvetiteli": "Просветители",
    "lektoriy_kaluga": "Лекторий",
}


def run_annotate(out_root: Path, selected: list[dict[str, Any]]) -> int:
    from .annotate_docx import annotate_report
    from .issues import build_issues
    from .simple_table import export_pair_problems, format_event_header, locate_report_issues

    summaries: list[str] = []

    for pair in selected:
        pair_out = out_root / pair["id"]
        cmp_path = pair_out / "comparison.json"
        if not cmp_path.is_file():
            print(f"Нет {cmp_path} — сначала запустите сверку.", file=sys.stderr)
            return 2

        comparison = json.loads(cmp_path.read_text(encoding="utf-8"))
        issues = build_issues(pair["id"], pair_out, comparison=comparison)
        report_path = Path(pair["report"])
        annotated_name = report_path.stem + "_с_комментариями.docx"
        annotated_path = pair_out / annotated_name

        result = annotate_report(report_path, issues, annotated_path)
        sheet = SHEET_TITLES.get(pair["id"], pair["id"])
        located = locate_report_issues(report_path, issues)
        event = comparison.get("event_check") or {}
        event_summary = format_event_header(event)
        xlsx_path = pair_out / "problemy_otcheta.xlsx"
        export_pair_problems(
            xlsx_path,
            event_summary=event_summary,
            located=located,
        )
        summaries.append(
            f"{sheet}: {result['comments']} комментариев → {annotated_path.name}; "
            f"таблица → {xlsx_path.name}"
        )
        print(f"  {summaries[-1]}", flush=True)

    print("Gotovo. " + "; ".join(summaries), flush=True)
    return 0


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="Сверка отчёта с ТЗ договора")
    parser.add_argument(
        "--pair",
        choices=["prosvetiteli", "lektoriy_kaluga", "all"],
        default="all",
    )
    parser.add_argument("--out", default=str(ROOT / "output"))
    parser.add_argument(
        "--recompare",
        action="store_true",
        help="Пересчитать сверку по уже сохранённым JSON без вызова API",
    )
    parser.add_argument(
        "--photos",
        action="store_true",
        help="Этап C: фото из отчёта и vision API (A/B не перезапускается)",
    )
    parser.add_argument(
        "--all-photos",
        action="store_true",
        help="Этап C: все реальные фото, не выборка 15–25 (иконки по-прежнему пропускаются)",
    )
    parser.add_argument(
        "--sample",
        type=int,
        default=20,
        help="Сколько кадров отправить на пару, если нет --all-photos (15–25)",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=2,
        help="Сколько кадров в одном запросе vision API",
    )
    parser.add_argument(
        "--export",
        action="store_true",
        help="Пересобрать Excel и PowerPoint из JSON без API",
    )
    parser.add_argument(
        "--annotate",
        action="store_true",
        help="Комментарии Word + простая таблица из сохранённых JSON (без API)",
    )
    args = parser.parse_args(argv)

    selected = PAIRS if args.pair == "all" else [p for p in PAIRS if p["id"] == args.pair]
    for pair in selected:
        if not pair["contract"].is_file() or not pair["report"].is_file():
            print(f"Нет файлов для пары {pair['id']}", file=sys.stderr)
            return 2

    out_root = Path(args.out)
    out_root.mkdir(parents=True, exist_ok=True)

    if args.photos or args.all_photos:
        client = GeminiClient(timeout=240)
        photo_summaries: list[dict[str, Any]] = []
        for pair in selected:
            mode = "все реальные фото" if args.all_photos else f"выборка {args.sample}"
            print(f"Этап C (фото, {mode}): {pair['id']}…", flush=True)
            photo = run_stage_c(
                client,
                pair,
                out_root / pair["id"],
                sample_n=args.sample,
                batch_size=max(1, args.batch_size),
                all_photos=args.all_photos,
            )
            photo_summaries.append(
                {
                    "pair_id": pair["id"],
                    "extracted": photo["extracted_total"],
                    "real_photo_candidates": photo.get("real_photo_candidates"),
                    "sampled": photo["sampled"],
                    "sent": photo.get("sent", photo["sampled"]),
                    "failed": photo.get("failed", 0),
                    "skipped_tiny": photo["skipped_tiny"],
                    "coverage": photo.get("coverage", "sample"),
                }
            )
        prev = {}
        for cand in (
            out_root / "summary.json",
            out_root / "Итоги текст" / "summary.json",
        ):
            if cand.is_file():
                loaded = json.loads(cand.read_text(encoding="utf-8"))
                details = (loaded.get("usage_text") or loaded.get("usage") or {}).get("details") or []
                if any(":C" not in str(d.get("stage") or "") for d in details) or loaded.get("usage_text"):
                    prev = loaded
                    break
        text_usage = Usage.from_dict((prev.get("usage_text") or prev.get("usage") or {}))
        # если прошлый usage уже содержал этап C, оставляем только A/B/qual
        text_details = [
            d
            for d in (text_usage.details or [])
            if ":C" not in str(d.get("stage") or "")
        ]
        text_usage = Usage(
            prompt_tokens=sum(int(d.get("prompt_tokens") or 0) for d in text_details),
            output_tokens=sum(int(d.get("output_tokens") or 0) for d in text_details),
            calls=len(text_details),
            details=text_details,
        )
        photo_usage = Usage()
        photo_summaries_all: list[dict[str, Any]] = []
        for pair in PAIRS:
            ppath = out_root / pair["id"] / "photos.json"
            if not ppath.is_file():
                continue
            pjson = json.loads(ppath.read_text(encoding="utf-8"))
            photo_usage = photo_usage.merge(Usage.from_dict(pjson.get("usage") or {}))
            photo_summaries_all.append(
                {
                    "pair_id": pair["id"],
                    "extracted": pjson.get("extracted_total"),
                    "real_photo_candidates": pjson.get("real_photo_candidates"),
                    "sampled": pjson.get("sampled"),
                    "sent": pjson.get("sent", pjson.get("sampled")),
                    "failed": pjson.get("failed", 0),
                    "skipped_tiny": pjson.get("skipped_tiny"),
                    "coverage": pjson.get("coverage", "sample"),
                }
            )
        if photo_summaries_all:
            photo_summaries = photo_summaries_all
        combined_usage = text_usage.merge(photo_usage)
        combined = {
            "pairs": [],
            "photo_pairs": photo_summaries,
            "usage": combined_usage.as_dict(),
            "usage_text": text_usage.as_dict(),
            "usage_photo": photo_usage.as_dict(),
            "photo_stage": "done",
            "photo_coverage": "all_real" if args.all_photos else "sample",
        }
        for pair in PAIRS:
            cmp = {}
            cmp_path = out_root / pair["id"] / "comparison.json"
            if not cmp_path.is_file():
                continue
            cmp = json.loads(cmp_path.read_text(encoding="utf-8"))
            prev_row = next((p for p in (prev.get("pairs") or []) if p.get("pair_id") == pair["id"]), {})
            combined["pairs"].append(
                {
                    "pair_id": pair["id"],
                    "same_event": cmp.get("same_event", prev_row.get("same_event")),
                    "event_status": (cmp.get("event_check") or {}).get("status") or prev_row.get("event_status"),
                    "event_comment": (cmp.get("event_check") or {}).get("comment") or prev_row.get("event_comment"),
                    "mismatch_count": cmp.get("mismatch_count", prev_row.get("mismatch_count")),
                    "report_images": cmp.get("report_images", prev_row.get("report_images")),
                    "photos_extracted": cmp.get("photos_extracted"),
                    "photos_sampled": cmp.get("photos_sampled"),
                    "photos_failed": cmp.get("photos_failed"),
                    "photo_coverage": cmp.get("photo_coverage"),
                    "out": str(out_root / pair["id"]),
                }
            )
        (out_root / "summary.json").write_text(
            json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        from .export import export_deliverables

        export_deliverables(out_root)
        u = combined["usage"]
        pu = combined["usage_photo"]
        print(
            f"Gotovo. photo in={pu.get('prompt_tokens', 0)} out={pu.get('output_tokens', 0)} "
            f"~ {pu.get('cost_rub', 0)} RUB extra. "
            f"total ~ {u.get('cost_rub', 0)} RUB. reports: {out_root}",
            flush=True,
        )
        return 0

    if args.annotate:
        return run_annotate(out_root, selected)

    if args.export:
        from .export import export_deliverables

        export_deliverables(out_root)
        print(f"Excel и презентация обновлены: {out_root}", flush=True)
        return 0

    client = None if args.recompare else GeminiClient()
    summaries: list[dict[str, Any]] = []
    for pair in selected:
        print(f"Обрабатываю {pair['id']}…", flush=True)
        result = verify_pair(
            client, pair, out_root / pair["id"], skip_llm=args.recompare
        )
        summaries.append(
            {
                "pair_id": result["pair_id"],
                "same_event": result["same_event"],
                "event_status": result["event_check"]["status"],
                "event_comment": result["event_check"]["comment"],
                "mismatch_count": result["mismatch_count"],
                "report_images": result["report_images"],
                "out": str(out_root / pair["id"]),
            }
        )
        print(
            f"  событие: {result['event_check']['status']}; "
            f"расхождений: {result['mismatch_count']}; "
            f"фото в DOCX: {result['report_images']}",
            flush=True,
        )

    combined = {
        "pairs": summaries,
        "usage": (json.loads((out_root / "summary.json").read_text(encoding="utf-8")).get("usage")
                  if args.recompare and (out_root / "summary.json").is_file()
                  else client.usage.as_dict()),
        "photo_stage": "skipped",
    }
    (out_root / "summary.json").write_text(
        json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    usage = combined["usage"]
    print(
        f"Gotovo. tokens in={usage.get('prompt_tokens', 0)} out={usage.get('output_tokens', 0)} "
        f"~ {usage.get('cost_rub', 0)} RUB. reports: {out_root}",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
