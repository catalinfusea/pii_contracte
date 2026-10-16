"""Structuri de date comune."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from enum import Enum


class Confidence(str, Enum):
    MARE = "mare"      # validat algoritmic (cifră de control, mod-97)
    MEDIE = "medie"    # tipar + ancoră de context
    MICA = "mică"      # ancoră prezentă, dar validarea a eșuat (posibil typo/OCR)


# Tipuri afișate în tabel, dar NEBIFATE implicit (date indirecte: utilizatoarea decide).
DEFAULT_UNSELECTED_KINDS = frozenset({"Loc naștere", "Emitent act", "Dată act"})


@dataclass
class Match:
    """O potrivire într-un șir de text (offset-uri în textul segmentului)."""
    kind: str
    value: str
    start: int
    end: int
    confidence: Confidence
    note: str = ""


@dataclass
class Finding:
    """O constatare într-un fișier, cu localizare."""
    id: int
    file: str
    kind: str
    value: str
    confidence: Confidence
    location: str          # ex. "pagina 2" / "corp, paragraful 14" / "antet"
    context: str           # text din jur (conține PII; doar pentru afișare locală)
    note: str = ""
    selected: bool = True  # debifat în interfață = nu se redactează
    # date interne pentru redactare
    _locator: dict = field(default_factory=dict, repr=False)


@dataclass
class FileResult:
    file: str                                   # fișierul analizat (.docx/.pdf; pentru .doc: cel convertit)
    findings: list[Finding] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    error: str | None = None
    source: str | None = None                   # originalul, dacă `file` e o conversie (ex. din .doc)

    @property
    def display_name(self) -> str:
        return os.path.basename(self.source or self.file)
