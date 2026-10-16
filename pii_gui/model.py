"""Starea tabelului de verificare, fără Tkinter (testabilă separat)."""
from __future__ import annotations

from dataclasses import dataclass

from pii_core.engine import norm_value
from pii_core.models import FileResult, Finding

LINK_NOTE = ("Link: dacă valoarea e bifată oriunde în document, linkul e eliminat "
             "chiar dacă rândul acesta e debifat.")


@dataclass(frozen=True)
class Row:
    iid: str
    file_index: int
    finding: Finding


def is_link(f: Finding) -> bool:
    return "rels" in f._locator


def _key(res: FileResult, f: Finding) -> tuple:
    loc = f._locator
    return (res.source or res.file, f.location, loc.get("start"), loc.get("end"), f.kind, norm_value(f.value))


class ReviewModel:
    def __init__(self):
        self.results: list[FileResult] = []

    # ------------------------------------------------------------ încărcare
    def clear(self) -> None:
        self.results = []

    def add_result(self, res: FileResult, previous: "ReviewModel | None" = None) -> None:
        """Adaugă un rezultat; dacă `previous` e dat (rescanare), păstrează debifările de atunci."""
        if previous is not None:
            unselected = {_key(r, f) for r in previous.results for f in r.findings if not f.selected}
            for f in res.findings:
                if _key(res, f) in unselected:
                    f.selected = False
        self.results.append(res)

    # ------------------------------------------------------------ rânduri
    @staticmethod
    def iid(file_index: int, f: Finding) -> str:
        return f"{file_index}:{f.id}"

    def rows(self, file_index: int | None = None) -> list[Row]:
        out = []
        for i, r in enumerate(self.results):
            if file_index is not None and i != file_index:
                continue
            out.extend(Row(self.iid(i, f), i, f) for f in r.findings)
        return out

    def get(self, iid: str) -> Row:
        i, fid = iid.split(":")
        i, fid = int(i), int(fid)
        for f in self.results[i].findings:
            if f.id == fid:
                return Row(iid, i, f)
        raise KeyError(iid)

    # ------------------------------------------------------------ bife
    def toggle(self, iid: str) -> list[str]:
        row = self.get(iid)
        row.finding.selected = not row.finding.selected
        return [iid]

    def set_rows(self, iids: list[str], state: bool) -> list[str]:
        for iid in iids:
            self.get(iid).finding.selected = state
        return list(iids)

    def set_same_value(self, iid: str, state: bool) -> list[str]:
        """Toate aparițiile valorii rândului, în același fișier. Întoarce rândurile modificate."""
        row = self.get(iid)
        key = norm_value(row.finding.value)
        changed = []
        for f in self.results[row.file_index].findings:
            if norm_value(f.value) == key:
                f.selected = state
                changed.append(self.iid(row.file_index, f))
        return changed

    def set_all(self, state: bool, file_index: int | None = None) -> list[str]:
        rows = self.rows(file_index)
        for r in rows:
            r.finding.selected = state
        return [r.iid for r in rows]

    # ------------------------------------------------------------ rezumat
    def counts(self) -> tuple[int, int]:
        total = sum(len(r.findings) for r in self.results)
        sel = sum(f.selected for r in self.results for f in r.findings)
        return sel, total

    def processable(self) -> list[FileResult]:
        return [r for r in self.results if not r.error]
