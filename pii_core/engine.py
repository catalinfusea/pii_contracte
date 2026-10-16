"""API-ul folosit de interfață: scan_file(), redact_file(), write_report().

Logurile conțin doar tipuri și numere de constatări, niciodată valorile.
"""
from __future__ import annotations

import csv
import logging
import os
import re

from . import docx_handler as dx
from . import pdf_handler as pd
from .detectors import detect, find_literal
from .models import DEFAULT_UNSELECTED_KINDS, Confidence, FileResult, Finding, Match

log = logging.getLogger("pii_core")

SUPPORTED = {".docx", ".pdf"}
MAX_FILE_BYTES = 100 * 1024 * 1024
# tipuri care NU se propagă prin căutare literală (prea scurte / ambigue)
_NO_PROPAGATE = {"Serie act identitate", "Loc naștere", "Emitent act", "Dată act"}
_CTX = 40


def norm_value(value: str) -> str:
    """Cheia de comparație a unei valori: spații comasate, litere mici."""
    return re.sub(r"\s+", " ", value).strip().lower()


def _context(text: str, s: int, e: int) -> str:
    a, b = max(0, s - _CTX), min(len(text), e + _CTX)
    return ("…" if a else "") + text[a:b].replace("\n", " ") + ("…" if b < len(text) else "")


_TEXT_KINDS = {"Nume", "Adresă", "Manual", "Adresă imobil", "Loc naștere", "Emitent act"}
_DATE_KINDS = {"Data nașterii", "Dată act"}


def mask(value: str, kind: str = "") -> str:
    """Valoare mascată pentru raport (raportul nu trebuie să devină el însuși un fișier cu PII).

    - texte (nume, adrese, emitenți): prima literă a fiecărui cuvânt;
    - date calendaristice: complet mascate (ziua/anul ar fi date de identificare);
    - email: prima literă + domeniul;
    - numere (CNP, IBAN, acte, telefon, permis, CF): doar ultimele 2 caractere, pentru corelarea cu documentul.
    """
    v = value.strip()
    if kind in _TEXT_KINDS:
        return " ".join(w[0] + "*" * (len(w) - 1) if len(w) > 1 else w for w in v.split())
    if kind in _DATE_KINDS:
        return re.sub(r"\d", "*", v)
    if kind == "Email" and "@" in v:
        user, dom = v.split("@", 1)
        return user[:1] + "***@" + dom
    compact = re.sub(r"\s", "", v)
    if len(compact) <= 4:
        return "*" * len(compact)
    return "*" * (len(compact) - 2) + compact[-2:]


# ------------------------------------------------------------------ scanare
def _units(path: str, ext: str):
    """Unități de text: (text, eticheta locației, locator) + avertismente."""
    if ext == ".docx":
        xml_parts, raw, warnings = dx.load(path)
        if dx.has_tracked_changes(xml_parts):
            warnings.append("Conține modificări urmărite (track changes). În copia redactată sunt "
                            "acceptate automat, iar textul șters dispare.")
        if "word/comments.xml" in xml_parts:
            warnings.append("Conține comentarii. În copia redactată sunt eliminate.")
        units = [(p.text, p.label, {"part": p.part, "par": p.index})
                 for p in dx.iter_paragraphs(xml_parts)]
        # linkuri externe: mailto:/tel: pot conține date de contact
        for rels, target in dx.hyperlink_targets(raw):
            units.append((target, "hyperlink", {"rels": rels}))
        return units, warnings
    doc = pd.open_pdf(path)
    try:
        pages, warnings = pd.page_texts(doc)
    finally:
        doc.close()
    return [(p.text, f"pagina {p.page_no + 1}", {"page": p.page_no}) for p in pages], warnings


def scan_file(path: str, manual_terms: list[str] | None = None) -> FileResult:
    res = FileResult(file=path)
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED:
        res.error = f"Format nesuportat: {ext or '(fără extensie)'}"
        return res
    try:
        if os.path.getsize(path) > MAX_FILE_BYTES:
            res.error = "Fișier prea mare (peste 100 MB)."
            return res
        units, res.warnings = _units(path, ext)
    except (dx.DocxError, pd.PdfError) as e:
        res.error = str(e)
        return res
    except Exception as e:  # noqa: BLE001
        log.exception("Eroare neașteptată la citire (%s)", ext)
        res.error = f"Eroare neașteptată la citire: {type(e).__name__}"
        return res

    per_unit: list[list[Match]] = [detect(t) for t, _, _ in units]

    # propagare: o valoare găsită o dată (ex. numele din preambul) e căutată peste tot
    known: dict[str, str] = {}
    for ms in per_unit:
        for m in ms:
            if m.kind not in _NO_PROPAGATE and m.confidence != Confidence.MICA:
                known.setdefault(norm_value(m.value), m.kind)
    extra = [(t.strip(), "Manual") for t in (manual_terms or []) if t.strip()]
    for i, (text, _, _) in enumerate(units):
        occupied = [(m.start, m.end) for m in per_unit[i]]
        candidates = [(v, k, Confidence.MEDIE, "aceeași valoare apare și în altă parte") for v, k in known.items()]
        candidates += [(v, k, Confidence.MARE, "adăugat manual") for v, k in extra]
        for value, kind, conf, note in candidates:
            for m in find_literal(text, value, kind, conf, note):
                if all(m.end <= s or m.start >= e for s, e in occupied):
                    per_unit[i].append(m)
                    occupied.append((m.start, m.end))

    fid = 0
    for (text, label, loc), ms in zip(units, per_unit):
        for m in sorted(ms, key=lambda x: x.start):
            fid += 1
            res.findings.append(Finding(
                id=fid, file=path, kind=m.kind, value=m.value, confidence=m.confidence,
                location=label, context=_context(text, m.start, m.end), note=m.note,
                selected=m.kind not in DEFAULT_UNSELECTED_KINDS,
                _locator={**loc, "start": m.start, "end": m.end}))
    counts: dict[str, int] = {}
    for f in res.findings:
        counts[f.kind] = counts.get(f.kind, 0) + 1
    log.info("Scanat %s: %s", ext, counts)
    return res


# ------------------------------------------------------------------ redactare
def output_path(src: str, out_dir: str, taken: set[str] | None = None) -> str:
    """`taken`: căi deja folosite în aceeași rulare (fișiere cu același nume din foldere diferite)."""
    base, ext = os.path.splitext(os.path.basename(src))
    out = os.path.join(out_dir, f"{base}_REDACTAT{ext}")
    n = 2
    while taken is not None and os.path.normcase(os.path.abspath(out)) in taken:
        out = os.path.join(out_dir, f"{base}_REDACTAT_{n}{ext}")
        n += 1
    if os.path.normcase(os.path.abspath(out)) == os.path.normcase(os.path.abspath(src)):
        raise ValueError("Ieșirea ar suprascrie originalul.")
    if taken is not None:
        taken.add(os.path.normcase(os.path.abspath(out)))
    return out


def mixed_values(res: FileResult) -> set[str]:
    """Valori (normalizate) cu apariții bifate ȘI debifate: păstrate intenționat în unele locuri."""
    sel = {norm_value(f.value) for f in res.findings if f.selected}
    unsel = {norm_value(f.value) for f in res.findings if not f.selected}
    return sel & unsel


def redact_file(res: FileResult, out_dir: str, taken: set[str] | None = None) -> tuple[str, list[str]]:
    """Scrie copia redactată și o verifică. Întoarce (cale, probleme_de_verificare)."""
    if res.error:
        raise ValueError(res.error)
    os.makedirs(out_dir, exist_ok=True)
    out = output_path(res.file, out_dir, taken)
    selected = [f for f in res.findings if f.selected]
    values = {f.value for f in selected}
    ext = os.path.splitext(res.file)[1].lower()

    if ext == ".docx":
        xml_parts, raw, _ = dx.load(res.file)
        pars = {(p.part, p.index): p for p in dx.iter_paragraphs(xml_parts)}
        by_par: dict[tuple, list[Finding]] = {}
        hyperlink_values: set[str] = set()
        for f in selected:
            loc = f._locator
            if "rels" in loc:
                hyperlink_values.add(f.value)
                continue
            by_par.setdefault((loc["part"], loc["par"]), []).append(f)
        for key, fs in by_par.items():
            for f in sorted(fs, key=lambda x: x._locator["start"], reverse=True):
                dx.replace_span(pars[key], f._locator["start"], f._locator["end"],
                                f"[REDACTAT:{f.kind.upper()}]")
        # linkurile care conțin valori redactate în text sunt și ele neutralizate
        dx.save(out, xml_parts, raw, res.file, values)
        check_text = dx.all_text(out)
    else:
        spans = [(f._locator["page"], f._locator["start"], f._locator["end"], f.kind) for f in selected]
        pd.redact(res.file, out, spans, values)
        check_text = pd.all_text(out)

    problems = verify(check_text, selected, mixed_values(res))
    log.info("Redactat %s: %d valori, %d probleme la verificare", ext, len(selected), len(problems))
    return out, problems


def verify_file(path: str, res: FileResult) -> list[str]:
    """Verificarea post-redactare pe orice copie derivată (ex. exportul DOCX -> PDF făcut de Word)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        text = dx.all_text(path)
    elif ext == ".pdf":
        text = pd.all_text(path)
    else:
        return [f"Format necunoscut pentru verificare: {ext}"]
    return verify(text, [f for f in res.findings if f.selected], mixed_values(res))


def verify(text: str, selected: list[Finding], kept: set[str] | None = None) -> list[str]:
    """Caută în copia redactată valorile care trebuiau să dispară.

    `kept`: valori normalizate cu cel puțin o apariție debifată - rămân legitim în text,
    deci nu sunt raportate ca probleme.
    """
    problems, seen = [], set()
    kept = kept or set()
    for f in selected:
        key = (f.kind, norm_value(f.value))
        if key in seen or key[1] in kept:
            continue
        seen.add(key)
        if find_literal(text, f.value, f.kind, f.confidence, ""):
            problems.append(f"{f.kind} {mask(f.value, f.kind)} apare încă în copia redactată")
    return problems


# ------------------------------------------------------------------ raport
REPORT_COLUMNS = ["Fișier", "Tip", "Valoare (mascată)", "Încredere", "Localizare", "Notă", "Redactat"]


def write_report(results: list[FileResult], path: str) -> None:
    """CSV cu ';' și BOM UTF-8, ca să se deschidă corect în Excel cu setări românești."""
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(REPORT_COLUMNS)
        for r in results:
            name = r.display_name
            if r.error:
                w.writerow([name, "EROARE", "", "", "", r.error, ""])
                continue
            for msg in r.warnings:
                w.writerow([name, "AVERTISMENT", "", "", "", msg, ""])
            for f in r.findings:
                w.writerow([name, f.kind, mask(f.value, f.kind), f.confidence.value, f.location,
                            f.note, "da" if f.selected else "nu"])
