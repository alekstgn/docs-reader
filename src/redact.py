"""Редакция ФИО до любого вызова внешнего API."""

from __future__ import annotations

import re

REPLACEMENT = "[ФИО скрыто]"

# Явные формы из образцов (включая реальное имя лектора в ТЗ Лектория).
KNOWN_NAMES = [
    r"Татаренкова\s+Трофима\s+Викторовича",
    r"Татаренков[аеу]?\s+Трофим[ауе]?\s+Викторович[аеу]?",
    r"Иванова\s+Ивана\s+Ивановича",
    r"Иванов\s+Иван\s+Иванович",
    r"Викторова\s+Виктора\s+Викторовича",
    r"Викторов\s+Виктор\s+Викторович",
    r"И\.\s*И\.\s*Иванов",
    r"В\.\s*В\.\s*Викторов",
]

# Фамилия Имя Отчество / родительный падеж.
FIO_RE = re.compile(
    r"\b[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}"
    r"(?:овича|евича|овны|евны|ича|ичны|вича)?\b"
)

# Не трогать устойчивые названия, которые похожи на ФИО.
KEEP = (
    "Российск",
    "Великой Отечественн",
    "Министерств",
    "Военно-историческ",
)


def redact_names(text: str) -> str:
    for pattern in KNOWN_NAMES:
        text = re.sub(pattern, REPLACEMENT, text, flags=re.IGNORECASE)
    def _sub(match: re.Match[str]) -> str:
        value = match.group(0)
        if any(token in value for token in KEEP):
            return value
        return REPLACEMENT

    return FIO_RE.sub(_sub, text)
