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

"""Operații lungi rulate în thread separat; comunică cu interfața doar prin coadă (queue.Queue).

Mesaje trimise în coadă (tuple):
  ("progress", index, total, nume_fișier, etapă)
  ("file", FileResult)                      - la scanare
  ("redacted", nume, cale_ieșire, probleme) - la redactare
  ("file_error", nume, mesaj)               - la redactare
  ("done", rezumat: dict)
  ("fatal", mesaj)                          - eroare care oprește jobul
Niciun mesaj nu conține valori detectate, doar nume de fișiere și numere.
"""
from __future__ import annotations

import datetime as _dt
import logging
import os
import queue
import threading

from pii_core import doc_convert, engine
from pii_core.models import FileResult

log = logging.getLogger("pii_core")

INPUT_EXTS = {".docx", ".pdf", ".doc"}


def scan_job(paths: list[str], terms: list[str], doc_format: str, q: queue.Queue,
             cancel: threading.Event, converter_factory=doc_convert.WordConverter) -> None:
    results: list[FileResult] = []
    docs = [p for p in paths if os.path.splitext(p)[1].lower() == ".doc"]
    converter = None
    try:
        if docs:
            try:
                converter = converter_factory().__enter__()
            except doc_convert.ConversionError as e:
                converter = None
                conv_error = str(e)
            else:
                conv_error = None
        n = len(paths)
        for i, path in enumerate(paths, 1):
            if cancel.is_set():
                break
            name = os.path.basename(path)
            ext = os.path.splitext(path)[1].lower()
            q.put(("progress", i, n, name, "scanare"))
            if ext == ".doc":
                if converter is None:
                    res = FileResult(file=path, source=path, error=conv_error)
                else:
                    q.put(("progress", i, n, name, "conversie Word"))
                    try:
                        conv = converter.convert(path, doc_format)
                    except doc_convert.ConversionError as e:
                        res = FileResult(file=path, source=path, error=str(e))
                    else:
                        res = engine.scan_file(conv, terms)
                        res.source = path
                        res.warnings[0:0] = [f"Convertit din .doc cu Word. Copia convertită (NEREDACTATĂ) "
                                             f"e în folderul „{doc_convert.CONVERSION_DIR}” de lângă original."]
                        res.warnings[1:1] = list(getattr(converter, "last_notes", []))
            else:
                res = engine.scan_file(path, terms)
            results.append(res)
            q.put(("file", res))
    except Exception as e:  # noqa: BLE001 - nimic nu trebuie să blocheze interfața
        log.exception("Eroare la scanare")
        q.put(("fatal", f"Eroare neașteptată la scanare: {type(e).__name__}"))
        return
    finally:
        if converter is not None:
            converter.__exit__(None, None, None)
    q.put(("done", {"kind": "scan", "files": len(results), "cancelled": cancel.is_set()}))


def default_out_dir(first_file: str, now: _dt.datetime | None = None) -> str:
    stamp = (now or _dt.datetime.now()).strftime("%Y-%m-%d_%H%M")
    return os.path.join(os.path.dirname(os.path.abspath(first_file)), f"redactat_{stamp}")


OUTPUT_FORMATS = ("original", "pdf", "docx")


def _unique(path: str, taken: set[str]) -> str:
    base, ext = os.path.splitext(path)
    out, n = path, 2
    while os.path.normcase(os.path.abspath(out)) in taken or os.path.exists(out):
        out = f"{base}_{n}{ext}"
        n += 1
    taken.add(os.path.normcase(os.path.abspath(out)))
    return out


def redact_job(results: list[FileResult], out_dir: str, q: queue.Queue, cancel: threading.Event,
               out_format: str = "original", converter_factory=doc_convert.WordConverter) -> None:
    """Redactează în formatul sursei; dacă se cere alt format, Word exportă copia REDACTATĂ
    (se păstrează ambele), iar exportul e verificat și el."""
    if out_format not in OUTPUT_FORMATS:
        raise ValueError(out_format)
    done, failed, with_problems = 0, 0, 0
    taken: set[str] = set()
    converter, conv_error = None, None
    try:
        os.makedirs(out_dir, exist_ok=True)
        todo = [r for r in results if not r.error]
        n = len(todo)
        for i, res in enumerate(todo, 1):
            if cancel.is_set():
                break
            name = res.display_name
            q.put(("progress", i, n, name, "redactare"))
            try:
                out, problems = engine.redact_file(res, out_dir, taken)
            except Exception as e:  # noqa: BLE001
                log.exception("Eroare la redactare")
                failed += 1
                q.put(("file_error", name, f"Copia nu a putut fi generată ({type(e).__name__})."))
                continue
            done += 1
            file_problems = bool(problems)
            q.put(("redacted", name, out, problems))

            ext = os.path.splitext(out)[1].lower()
            if out_format != "original" and ext != doc_convert.FORMATS[out_format][1]:
                q.put(("progress", i, n, name, f"export {out_format.upper()} cu Word"))
                if converter is None and conv_error is None:
                    try:
                        converter = converter_factory().__enter__()
                    except doc_convert.ConversionError as e:
                        conv_error = str(e)
                if converter is None:
                    q.put(("file_error", name, f"Export {out_format.upper()} imposibil: {conv_error} "
                                               f"Copia redactată a rămas în formatul inițial."))
                else:
                    dst = _unique(os.path.splitext(out)[0] + doc_convert.FORMATS[out_format][1], taken)
                    try:
                        converter.export(out, out_format, dst)
                        exp_problems = engine.verify_file(dst, res)
                    except doc_convert.ConversionError as e:
                        q.put(("file_error", name, f"Export {out_format.upper()} eșuat: {e} "
                                                   f"Copia redactată a rămas în formatul inițial."))
                    else:
                        file_problems = file_problems or bool(exp_problems)
                        q.put(("redacted", name, dst, exp_problems))
            with_problems += file_problems
        stamp = _dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        report = os.path.join(out_dir, f"raport_{stamp}.csv")
        engine.write_report(results, report)
    except Exception as e:  # noqa: BLE001
        log.exception("Eroare la generare")
        q.put(("fatal", f"Eroare la generare: {type(e).__name__}"))
        return
    finally:
        if converter is not None:
            converter.__exit__(None, None, None)
    q.put(("done", {"kind": "redact", "done": done, "failed": failed, "problems": with_problems,
                    "report": report, "out_dir": out_dir, "cancelled": cancel.is_set()}))


def start(target, *args) -> threading.Thread:
    t = threading.Thread(target=target, args=args, daemon=True)
    t.start()
    return t
