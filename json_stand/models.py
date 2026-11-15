from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class ConvertError:
    code: str
    message: str

    def to_dict(self) -> dict[str, str]:
        return asdict(self)


@dataclass
class ConvertResult:
    ok: bool
    sheets: list[str] = field(default_factory=list)
    errors: list[ConvertError] = field(default_factory=list)
    workbook_path: str | None = None
    workbook_base64: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "sheets": self.sheets,
            "errors": [e.to_dict() for e in self.errors],
            "workbook_path": self.workbook_path,
            "workbook_base64": self.workbook_base64,
        }
