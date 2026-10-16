# pii_contracte - detectarea și redactarea datelor de identificare din contracte.
# Copyright (C) 2026  CatalinFusea
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
# SPDX-License-Identifier: AGPL-3.0-or-later

"""Citire și redactare PDF cu PyMuPDF.

Redactarea folosește adnotări de redactare aplicate (apply_redactions), care
ȘTERG textul din conținutul paginii, nu doar îl acoperă. Copia e salvată
complet (garbage=4), deci obiectele vechi nu rămân în fișier.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pymupdf

from .i18n import t

MAX_PAGES = 2000
MIN_CHARS_TEXT_PAGE = 50  # sub acest prag, o pagină cu imagini e probabil scanată


class PdfError(Exception):
    pass


@dataclass
class PageText:
    page_no: int                       # 0-based
    text: str
    boxes: list[tuple | None] = field(default_factory=list)  # bbox per caracter; None = separator
    line_ids: list[int] = field(default_factory=list)       # linia fiecărui caracter


def open_pdf(path: str) -> pymupdf.Document:
    try:
        doc = pymupdf.open(path)
    except Exception as e:  # noqa: BLE001 - PyMuPDF ridică mai multe tipuri
        raise PdfError(t("err.pdf_open")) from e
    if doc.needs_pass:
        doc.close()
        raise PdfError(t("err.pdf_password"))
    if doc.page_count > MAX_PAGES:
        doc.close()
        raise PdfError(t("err.pdf_too_many_pages", n=doc.page_count))
    return doc


# Normalizare 1:1 (un caracter -> un caracter, offset-urile spre coordonate rămân valide).
# Fonturi ca Arial mapează glifa cratimei și pe U+00AD: textul extras conține soft hyphen în loc de "-".
CHAR_NORMALIZE = str.maketrans({
    "\u00ad": "-",   # soft hyphen
    "\u2010": "-",   # hyphen
    "\u2011": "-",   # non-breaking hyphen
    "\u00a0": " ",   # spațiu neseparabil
    "\u202f": " ",   # spațiu îngust neseparabil
})


def normalize_text(text: str) -> str:
    return text.translate(CHAR_NORMALIZE)


def page_texts(doc: pymupdf.Document) -> tuple[list[PageText], list[str]]:
    pages, warnings = [], []
    scanned = []
    for pno in range(doc.page_count):
        page = doc[pno]
        pt = PageText(pno, "")
        chars: list[str] = []
        line_no = 0
        raw = page.get_text("rawdict", flags=pymupdf.TEXT_PRESERVE_WHITESPACE)
        for block in raw["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    for ch in span["chars"]:
                        chars.append(normalize_text(ch["c"]))
                        pt.boxes.append(tuple(ch["bbox"]))
                        pt.line_ids.append(line_no)
                # liniile din același bloc: separate prin spațiu (adresele trec pe rândul următor)
                chars.append(" ")
                pt.boxes.append(None)
                pt.line_ids.append(line_no)
                line_no += 1
            chars.append("\n")
            pt.boxes.append(None)
            pt.line_ids.append(line_no)
        pt.text = "".join(chars)
        pages.append(pt)
        if len(pt.text.strip()) < MIN_CHARS_TEXT_PAGE and page.get_images():
            scanned.append(pno + 1)
    if scanned:
        warnings.append(t("warn.pdf_scanned", pages=", ".join(map(str, scanned))))
    if doc.is_form_pdf:
        warnings.append(t("warn.pdf_forms"))
    if any(doc[p].get_images() for p in range(doc.page_count)) and not scanned:
        warnings.append(t("warn.pdf_images"))
    return pages, warnings


def rects_for_span(pt: PageText, start: int, end: int) -> list[pymupdf.Rect]:
    """Un dreptunghi per linie pentru caracterele [start, end)."""
    by_line: dict[int, pymupdf.Rect] = {}
    for i in range(start, min(end, len(pt.boxes))):
        b = pt.boxes[i]
        if b is None:
            continue
        r = pymupdf.Rect(b)
        lid = pt.line_ids[i]
        by_line[lid] = by_line[lid] | r if lid in by_line else r
    return list(by_line.values())


def redact(src_path: str, out_path: str, spans: list[tuple[int, int, int, str]],
           redacted_values: set[str]) -> None:
    """spans: (pagina 0-based, start, end, eticheta)."""
    doc = open_pdf(src_path)
    try:
        pages, _ = page_texts(doc)
        by_page: dict[int, list] = {}
        for pno, s, e, label in spans:
            by_page.setdefault(pno, []).append((s, e, label))
        for pno, items in by_page.items():
            page = doc[pno]
            for s, e, label in items:
                for r in rects_for_span(pages[pno], s, e):
                    page.add_redact_annot(r, fill=(0, 0, 0))
            # text: eliminat; imagini: pixelii acoperiți se înnegresc; grafică: neatinsă
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_PIXELS)
        # linkuri care conțin valori redactate (mailto:, tel:)
        for page in doc:
            for link in page.get_links():
                uri = (link.get("uri") or "").lower()
                if uri and any(v.lower().replace(" ", "") in uri.replace("%20", "") for v in redacted_values):
                    page.delete_link(link)
        # metadate
        doc.set_metadata({})
        doc.del_xml_metadata()
        doc.save(out_path, garbage=4, deflate=True, clean=True)
    finally:
        doc.close()


def all_text(path: str) -> str:
    """Tot textul + metadatele + linkurile - pentru verificarea post-redactare."""
    doc = open_pdf(path)
    try:
        chunks = [str(doc.metadata), doc.get_xml_metadata() or ""]
        for page in doc:
            chunks.append(normalize_text(page.get_text()))
            chunks.extend(link.get("uri") or "" for link in page.get_links())
        return "\n".join(chunks)
    finally:
        doc.close()
