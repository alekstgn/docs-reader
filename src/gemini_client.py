"""Клиент нативного Google generateContent через ProxyAPI."""

from __future__ import annotations

import base64
import json
import os
import time
from dataclasses import dataclass, field
from typing import Any

import requests
from dotenv import load_dotenv

MODEL = "gemini-2.5-flash-lite"
ENDPOINT = (
    "https://api.proxyapi.ru/google/v1beta/models/"
    f"{MODEL}:generateContent"
)
INPUT_RUB_PER_MILLION = 26.0
OUTPUT_RUB_PER_MILLION = 129.0


def load_api_key() -> str:
    load_dotenv(override=False)
    for name in ("PROXYAPI_API_KEY", "PROXYAPI_KEY", "API_KEY"):
        value = os.getenv(name, "").strip()
        if value:
            return value
    env_txt = os.path.join(os.path.dirname(__file__), "..", ".env.txt")
    env_txt = os.path.abspath(env_txt)
    if os.path.isfile(env_txt):
        raw = open(env_txt, encoding="utf-8").read().strip()
        if raw and "=" not in raw.splitlines()[0]:
            return raw
        for line in raw.splitlines():
            if line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            if k.strip() in {"PROXYAPI_API_KEY", "PROXYAPI_KEY", "API_KEY"}:
                return v.strip().strip('"').strip("'")
    raise RuntimeError("Не найден ключ ProxyAPI в .env / .env.txt")


@dataclass
class Usage:
    prompt_tokens: int = 0
    output_tokens: int = 0
    calls: int = 0
    details: list[dict[str, Any]] = field(default_factory=list)

    @property
    def cost_rub(self) -> float:
        return (
            self.prompt_tokens * INPUT_RUB_PER_MILLION
            + self.output_tokens * OUTPUT_RUB_PER_MILLION
        ) / 1_000_000

    def add(self, meta: dict[str, Any], stage: str) -> None:
        prompt = int(meta.get("promptTokenCount") or 0)
        output = int(
            meta.get("candidatesTokenCount")
            or meta.get("outputTokenCount")
            or 0
        )
        self.prompt_tokens += prompt
        self.output_tokens += output
        self.calls += 1
        self.details.append(
            {
                "stage": stage,
                "prompt_tokens": prompt,
                "output_tokens": output,
                "total_tokens": int(meta.get("totalTokenCount") or prompt + output),
            }
        )

    def as_dict(self) -> dict[str, Any]:
        return {
            "model": MODEL,
            "calls": self.calls,
            "prompt_tokens": self.prompt_tokens,
            "output_tokens": self.output_tokens,
            "cost_rub": round(self.cost_rub, 4),
            "tariff": {
                "input_rub_per_million": INPUT_RUB_PER_MILLION,
                "output_rub_per_million": OUTPUT_RUB_PER_MILLION,
            },
            "details": self.details,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "Usage":
        data = data or {}
        u = cls(
            prompt_tokens=int(data.get("prompt_tokens") or 0),
            output_tokens=int(data.get("output_tokens") or 0),
            calls=int(data.get("calls") or 0),
            details=list(data.get("details") or []),
        )
        return u

    def merge(self, other: "Usage") -> "Usage":
        return Usage(
            prompt_tokens=self.prompt_tokens + other.prompt_tokens,
            output_tokens=self.output_tokens + other.output_tokens,
            calls=self.calls + other.calls,
            details=list(self.details) + list(other.details),
        )


class GeminiClient:
    def __init__(self, api_key: str | None = None, timeout: int = 180) -> None:
        self.api_key = api_key or load_api_key()
        self.timeout = timeout
        self.usage = Usage()

    def generate_json(
        self,
        prompt: str,
        schema: dict[str, Any],
        *,
        stage: str,
        temperature: float = 0.1,
        max_output_tokens: int = 16384,
    ) -> dict[str, Any]:
        return self.generate_json_parts(
            [{"text": prompt}],
            schema,
            stage=stage,
            temperature=temperature,
            max_output_tokens=max_output_tokens,
        )

    def generate_json_with_images(
        self,
        prompt: str,
        images: list[tuple[str, bytes]],
        schema: dict[str, Any],
        *,
        stage: str,
        temperature: float = 0.1,
        max_output_tokens: int = 8192,
    ) -> dict[str, Any]:
        parts: list[dict[str, Any]] = [{"text": prompt}]
        for mime, data in images:
            parts.append(
                {
                    "inlineData": {
                        "mimeType": mime,
                        "data": base64.b64encode(data).decode("ascii"),
                    }
                }
            )
        return self.generate_json_parts(
            parts,
            schema,
            stage=stage,
            temperature=temperature,
            max_output_tokens=max_output_tokens,
        )

    def generate_json_parts(
        self,
        parts: list[dict[str, Any]],
        schema: dict[str, Any],
        *,
        stage: str,
        temperature: float = 0.1,
        max_output_tokens: int = 16384,
    ) -> dict[str, Any]:
        last_schema_error: Exception | None = None
        for schema_variant in (schema, _lowercase_schema(schema), None):
            try:
                return self._post_json(
                    parts,
                    schema_variant,
                    stage=stage,
                    temperature=temperature,
                    max_output_tokens=max_output_tokens,
                )
            except RuntimeError as exc:
                message = str(exc).lower()
                if any(x in message for x in ("too large", "payload", "413", "resource exhausted")):
                    raise
                if schema_variant is not None and ("schema" in message or "400" in message):
                    last_schema_error = exc
                    continue
                raise
        raise RuntimeError(f"Не удалось применить responseSchema: {last_schema_error}")

    def _post_json(
        self,
        parts: list[dict[str, Any]],
        schema: dict[str, Any] | None,
        *,
        stage: str,
        temperature: float,
        max_output_tokens: int,
    ) -> dict[str, Any]:
        gen: dict[str, Any] = {
            "temperature": temperature,
            "maxOutputTokens": max_output_tokens,
            "responseMimeType": "application/json",
        }
        if schema is not None:
            gen["responseSchema"] = schema
        payload = {
            "contents": [{"role": "user", "parts": parts}],
            "generationConfig": gen,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        last_error: Exception | None = None
        for attempt in range(4):
            try:
                resp = requests.post(
                    ENDPOINT,
                    headers=headers,
                    json=payload,
                    timeout=self.timeout,
                )
            except requests.RequestException as exc:
                last_error = exc
                time.sleep(2 ** attempt)
                continue
            if resp.status_code in {429, 500, 502, 503, 504}:
                last_error = RuntimeError(f"HTTP {resp.status_code}: {resp.text[:500]}")
                time.sleep(2 ** attempt + 1)
                continue
            if resp.status_code >= 400:
                raise RuntimeError(
                    f"ProxyAPI HTTP {resp.status_code}: {resp.text[:1500]}"
                )
            data = resp.json()
            self.usage.add(data.get("usageMetadata") or {}, stage)
            text = _candidate_text(data)
            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"Модель вернула не-JSON на этапе {stage}: {text[:800]}"
                ) from exc
        raise RuntimeError(f"API недоступен после 4 попыток: {last_error}")


def _lowercase_schema(obj: Any) -> Any:
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        for key, value in obj.items():
            if key == "type" and isinstance(value, str):
                out[key] = value.lower()
            else:
                out[key] = _lowercase_schema(value)
        return out
    if isinstance(obj, list):
        return [_lowercase_schema(item) for item in obj]
    return obj


def _candidate_text(data: dict[str, Any]) -> str:
    candidates = data.get("candidates") or []
    if not candidates:
        raise RuntimeError(f"Пустой ответ модели: {json.dumps(data, ensure_ascii=False)[:800]}")
    parts = (candidates[0].get("content") or {}).get("parts") or []
    chunks = [p.get("text", "") for p in parts if "text" in p]
    if not chunks:
        raise RuntimeError(f"Нет text в candidates: {json.dumps(data, ensure_ascii=False)[:800]}")
    return "".join(chunks)
