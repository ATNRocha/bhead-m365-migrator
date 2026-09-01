from __future__ import annotations

import json
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any


def normalize_text(value: str) -> str:
    value = (value or "").strip().upper()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return " ".join(value.split())


@dataclass(frozen=True)
class MigrationProfile:
    path: Path
    data: dict[str, Any]

    @property
    def name(self) -> str:
        return self.data.get("name", self.path.stem)

    @property
    def target(self) -> str:
        return self.data.get("target", "Excel Online / Microsoft 365")

    @property
    def sheet_names(self) -> set[str]:
        return set(self.data.get("sheet_names", []))

    def find_header_rule(self, header: str) -> dict[str, Any] | None:
        key = normalize_text(header)
        for item in self.data.get("header_rules", []):
            aliases = [normalize_text(v) for v in item.get("headers", [])]
            if key in aliases:
                return item
        return None

    def style(self, name: str) -> dict[str, str]:
        return self.data["styles"][name]


def load_profile(path: str | Path) -> MigrationProfile:
    p = Path(path)
    with p.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return MigrationProfile(path=p, data=data)
