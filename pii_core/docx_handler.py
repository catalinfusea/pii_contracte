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

"""Citire și redactare DOCX direct pe XML (fără python-docx), ca să acopere
corp, tabele, casete de text, antete, subsoluri, note și track changes.

Redactarea produce o copie în care:
- valorile selectate sunt înlocuite cu [REDACTAT:<tip>];
- modificările urmărite sunt acceptate (textul șters din w:del dispare);
- comentariile sunt eliminate (inclusiv autorii);
- proprietățile documentului (autor, companie etc.) sunt golite.
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
_w = lambda tag: f"{{{W}}}{tag}"  # noqa: E731

MAX_UNCOMPRESSED = 200 * 1024 * 1024  # protecție zip bomb
MAX_ENTRIES = 5000

# părțile cu text vizibil
_TEXT_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")
_PART_LABEL = {"document": "corp", "header": "antet", "footer": "subsol",
               "footnotes": "note de subsol", "endnotes": "note de final", "comments": "comentarii"}


def _parser() -> etree.XMLParser:
    # fără entități externe / rețea (XXE), fără arbori uriași
    return etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False,
                           remove_blank_text=False)


class DocxError(Exception):
    pass


def _open_zip(path: str) -> zipfile.ZipFile:
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        raise DocxError("Fișierul nu e un .docx valid (poate e protejat cu parolă sau corupt).") from e
    infos = zf.infolist()
    if len(infos) > MAX_ENTRIES or sum(i.file_size for i in infos) > MAX_UNCOMPRESSED:
        zf.close()
        raise DocxError("Fișier .docx suspect (prea mare la dezarhivare).")
    if "word/document.xml" not in zf.namelist():
        zf.close()
        raise DocxError("Fișierul nu conține word/document.xml.")
    return zf


@dataclass
class Segment:
    elem: etree._Element | None  # w:t, w:noBreakHyphen, w:softHyphen; None = separator (tab, break)
    text: str


# Caractere pe care Word le stochează ca elemente, nu ca text în w:t:
# cratima neseparabilă (Ctrl+Shift+-) e vizibilă ("-"), cratima opțională (Ctrl+-) nu ("").
_CHAR_ELEMS = {"noBreakHyphen": "-", "softHyphen": ""}


@dataclass
class Paragraph:
    part: str
    index: int
    segments: list[Segment]

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.segments)

    @property
    def label(self) -> str:
        base = re.sub(r"\d+$", "", self.part.split("/")[-1].removesuffix(".xml"))
        return f"{_PART_LABEL.get(base, base)}, paragraful {self.index + 1}"


def _paragraphs(part: str, root: etree._Element) -> list[Paragraph]:
    out = []
    for i, p in enumerate(root.iter(_w("p"))):
        segs: list[Segment] = []
        for el in p.iter():
            if el.tag == _w("t"):
                # nu intra în paragrafe imbricate (casete de text): sunt iterate separat;
                # textul din modificări urmărite șterse nu contează (dispare la redactare)
                if _nearest_p(el) is p and not _in_deleted(el, p):
                    segs.append(Segment(el, el.text or ""))
            elif el.tag in (_w("tab"), _w("br"), _w("cr")) and _nearest_p(el) is p:
                segs.append(Segment(None, " "))
            elif (isinstance(el.tag, str) and el.tag.startswith(f"{{{W}}}")
                  and el.tag[len(W) + 2:] in _CHAR_ELEMS and _nearest_p(el) is p and not _in_deleted(el, p)):
                segs.append(Segment(el, _CHAR_ELEMS[el.tag[len(W) + 2:]]))
        out.append(Paragraph(part, i, segs))
    return out


def _in_deleted(el: etree._Element, stop: etree._Element) -> bool:
    cur = el.getparent()
    while cur is not None and cur is not stop:
        if cur.tag in (_w("del"), _w("moveFrom")):
            return True
        cur = cur.getparent()
    return False


def _nearest_p(el: etree._Element) -> etree._Element | None:
    cur = el.getparent()
    while cur is not None and cur.tag != _w("p"):
        cur = cur.getparent()
    return cur


def load(path: str) -> tuple[dict[str, etree._Element], dict[str, bytes], list[str]]:
    """Întoarce (părți XML cu text parsate, toate celelalte intrări brute, avertismente)."""
    warnings: list[str] = []
    xml_parts: dict[str, etree._Element] = {}
    raw: dict[str, bytes] = {}
    with _open_zip(path) as zf:
        for name in zf.namelist():
            data = zf.read(name)
            if _TEXT_PARTS.match(name):
                xml_parts[name] = etree.fromstring(data, _parser())
            else:
                raw[name] = data
    n_img = sum(1 for n in raw if n.startswith("word/media/"))
    if n_img:
        warnings.append(f"Conține {n_img} imagine(i). Datele din imagini (ex. buletin scanat lipit în "
                        f"document) NU sunt detectate - verificați manual.")
    if any(n.startswith("word/embeddings/") for n in raw):
        warnings.append("Conține obiecte încorporate (ex. Excel). Nu sunt analizate - verificați manual.")
    return xml_parts, raw, warnings


def iter_paragraphs(xml_parts: dict[str, etree._Element]):
    # documentul principal primul, apoi restul în ordine stabilă
    for part in sorted(xml_parts, key=lambda n: (n != "word/document.xml", n)):
        yield from _paragraphs(part, xml_parts[part])


def has_tracked_changes(xml_parts: dict[str, etree._Element]) -> bool:
    return any(next(r.iter(_w("del"), _w("ins"), _w("moveFrom"), _w("moveTo")), None) is not None
               for r in xml_parts.values())


# ------------------------------------------------------------------ redactare
def replace_span(par: Paragraph, start: int, end: int, replacement: str) -> None:
    """Înlocuiește textul [start, end) din paragraf, chiar dacă e împărțit pe mai multe w:t."""
    pos = 0
    first = True
    for seg in par.segments:
        s_len = len(seg.text)
        s_start, s_end = pos, pos + s_len
        pos = s_end
        if seg.elem is None or s_end < start or s_start > end:
            continue
        if s_len and (s_end <= start or s_start >= end):
            continue
        if s_len == 0 and not (start < s_start < end):
            continue                       # cratimă opțională doar strict în interior
        if seg.elem.tag != _w("t"):
            # element-caracter (cratimă): devine w:t, ca să poată fi golit / să primească înlocuirea
            t = etree.Element(_w("t"))
            t.text = seg.text
            seg.elem.getparent().replace(seg.elem, t)
            seg.elem = t
        a = max(start, s_start) - s_start
        b = min(end, s_end) - s_start
        new = seg.text[:a] + (replacement if first else "") + seg.text[b:]
        first = False
        seg.elem.text = new
        seg.elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    # actualizează textul segmentelor w:t pentru eventuale înlocuiri ulterioare
    # (elementele-cratimă neatinse își păstrează textul virtual)
    for seg in par.segments:
        if seg.elem is not None and seg.elem.tag == _w("t"):
            seg.text = seg.elem.text or ""


def accept_tracked_changes(root: etree._Element) -> None:
    # textul șters / mutat de acolo: eliminat
    for tag in ("del", "moveFrom"):
        for el in list(root.iter(_w(tag))):
            el.getparent().remove(el)
    # inserările / mutările: păstrate, fără învelișul de revizie (care are autor și dată)
    for tag in ("ins", "moveTo"):
        for el in list(root.iter(_w(tag))):
            parent = el.getparent()
            idx = parent.index(el)
            for child in list(el):
                parent.insert(idx, child)
                idx += 1
            parent.remove(el)
    # modificări de formatare și marcaje de interval
    for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
                "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"):
        for el in list(root.iter(_w(tag))):
            el.getparent().remove(el)


def strip_comments(root: etree._Element) -> None:
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for el in list(root.iter(_w(tag))):
            el.getparent().remove(el)
    for ref in list(root.iter(_w("commentReference"))):
        run = ref.getparent()
        if run is not None and run.tag == _w("r") and len([c for c in run if c.tag != _w("rPr")]) == 1:
            run.getparent().remove(run)
        else:
            ref.getparent().remove(ref)


_CORE_FIELDS = re.compile(
    rb"(<(dc:creator|cp:lastModifiedBy|dc:title|dc:subject|cp:keywords|dc:description|cp:category)"
    rb"(?:\s[^>]*)?>)(.*?)(</\2>)", re.S)
_APP_FIELDS = re.compile(rb"(<(Company|Manager)>)(.*?)(</\2>)", re.S)


def save(path_out: str, xml_parts: dict[str, etree._Element], raw: dict[str, bytes],
         src_path: str, redacted_hyperlinks: set[str]) -> None:
    for name, root in xml_parts.items():
        accept_tracked_changes(root)
        strip_comments(root)
    if "word/comments.xml" in xml_parts:
        root = xml_parts["word/comments.xml"]
        for c in list(root):
            root.remove(c)
    with zipfile.ZipFile(src_path) as src, zipfile.ZipFile(path_out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            name = info.filename
            if name in xml_parts:
                data = etree.tostring(xml_parts[name], xml_declaration=True, encoding="UTF-8",
                                      standalone=True)
            else:
                data = raw[name]
                if name == "docProps/core.xml":
                    data = _CORE_FIELDS.sub(lambda m: m.group(1) + m.group(4), data)
                elif name == "docProps/app.xml":
                    data = _APP_FIELDS.sub(lambda m: m.group(1) + m.group(4), data)
                elif name.endswith(".rels") and redacted_hyperlinks:
                    data = _redact_rels(data, redacted_hyperlinks)
                elif name in ("word/commentsExtended.xml", "word/commentsIds.xml",
                              "word/commentsExtensible.xml", "word/people.xml"):
                    data = _empty_root(data)
            dst.writestr(info, data)


def _redact_rels(data: bytes, values: set[str]) -> bytes:
    root = etree.fromstring(data, _parser())
    for rel in root:
        target = rel.get("Target") or ""
        low = target.lower()
        if any(v.lower() in low or v.replace(" ", "").lower() in low.replace("%20", "") for v in values):
            rel.set("Target", "#redactat")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def hyperlink_targets(raw: dict[str, bytes]) -> list[tuple[str, str]]:
    """(fișier .rels, țintă) pentru linkuri externe - pot conține mailto:/tel:."""
    out = []
    for name, data in raw.items():
        if name.startswith("word/_rels/") and name.endswith(".rels"):
            root = etree.fromstring(data, _parser())
            for rel in root:
                if rel.get("TargetMode") == "External":
                    out.append((name, rel.get("Target") or ""))
    return out


def _empty_root(data: bytes) -> bytes:
    root = etree.fromstring(data, _parser())
    for c in list(root):
        root.remove(c)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def all_text(path: str) -> str:
    """Tot textul (inclusiv metadate și linkuri) - pentru verificarea post-redactare."""
    chunks = []
    with _open_zip(path) as zf:
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                root = etree.fromstring(zf.read(name), _parser())
                chunks.append(" ".join(root.itertext()))
                for el in root.iter():
                    chunks.extend(el.attrib.values())
                # textul unui paragraf concatenat (valori sparte în mai multe w:t)
                if _TEXT_PARTS.match(name):
                    chunks.extend(p.text for p in _paragraphs(name, root))
    return "\n".join(chunks)
