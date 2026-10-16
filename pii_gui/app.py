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

"""Fereastra principală: fișiere -> scanare -> verificare cu bife -> copii redactate + raport.

Valorile complete apar DOAR în fereastră (local). Raportul CSV e mascat (pii_core.engine.mask).
"""
from __future__ import annotations

import functools
import logging
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pii_core import doc_convert
from pii_core.i18n import conf_label, get_lang, kind_label, set_lang, t
from pii_core.models import Confidence

from . import jobs, runtime
from .model import ReviewModel, is_link

log = logging.getLogger("pii_core")

ON, OFF = "\u2611", "\u2610"   # ☑ ☐
LANG_NAMES = {"ro": "Română", "en": "English"}
COLUMNS = [  # id, cheie titlu, lățime, stretch
    ("sel", None, 34, False), ("file", "col.file", 150, False), ("kind", "col.kind", 160, False),
    ("value", "col.value", 200, True), ("conf", "col.conf", 75, False), ("loc", "col.loc", 130, False),
    ("ctx", "col.ctx", 320, True), ("note", "col.note", 220, True),
]


def app_title() -> str:
    return t("app.title")


def filetypes() -> list[tuple[str, str]]:
    return [(t("ft.contracts"), "*.docx *.doc *.pdf"), ("Word", "*.docx *.doc"), ("PDF", "*.pdf"),
            (t("ft.all"), "*.*")]


def _dpi_aware() -> None:
    """Text clar pe ecrane scalate (Windows 8.1+). Eșecul nu e grav: textul doar e mai neclar."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass


class App(ttk.Frame):
    POLL_MS = 100

    def __init__(self, master: tk.Tk, state: dict | None = None):
        super().__init__(master, padding=8)
        self.master = master
        self.files: list[str] = []
        self.terms: list[str] = []
        self.model = ReviewModel()
        self.q: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.dirty = False          # lista/termenii/formatul s-au schimbat după ultima scanare
        self.closing = False
        self.last_out_dir: str | None = None
        self._previous: ReviewModel | None = None
        self.doc_format = tk.StringVar(value="docx")
        self.out_format = tk.StringVar(value="original")
        self.out_dir = tk.StringVar()
        self.filter_var = tk.StringVar()
        self.status = tk.StringVar(value=t("ui.status_start"))
        self.lang_var = tk.StringVar(value=LANG_NAMES[get_lang()])
        self._build()
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        master.protocol("WM_DELETE_WINDOW", self._on_close)
        master.title(app_title())
        if state:
            self._restore(state)
        self._refresh_buttons()

    # ================================================================== construcție UI
    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=3)
        top.columnconfigure(1, weight=2)

        # --- 1. fișiere
        ff = ttk.LabelFrame(top, text=t("ui.files"), padding=6)
        ff.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ff.columnconfigure(0, weight=1)
        self.file_list = tk.Listbox(ff, height=6, selectmode="extended", activestyle="none")
        self.file_list.grid(row=0, column=0, rowspan=3, sticky="nsew")
        sb = ttk.Scrollbar(ff, orient="vertical", command=self.file_list.yview)
        sb.grid(row=0, column=1, rowspan=3, sticky="ns")
        self.file_list.configure(yscrollcommand=sb.set)
        self.btn_add = ttk.Button(ff, text=t("ui.add"), command=self._add_files)
        self.btn_add.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        self.btn_remove = ttk.Button(ff, text=t("ui.remove_sel"), command=self._remove_files)
        self.btn_remove.grid(row=1, column=2, sticky="ew", padx=(6, 0), pady=3)
        self.btn_clear = ttk.Button(ff, text=t("ui.clear"), command=self._clear_files)
        self.btn_clear.grid(row=2, column=2, sticky="new", padx=(6, 0))

        # --- opțiuni
        of = ttk.LabelFrame(top, text=t("ui.options"), padding=6)
        of.grid(row=0, column=1, sticky="nsew")
        of.columnconfigure(0, weight=1)
        lf = ttk.Frame(of)
        lf.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
        ttk.Label(lf, text=t("ui.language")).pack(side="left")
        self.lang_cb = ttk.Combobox(lf, textvariable=self.lang_var, state="readonly", width=10,
                                    values=[LANG_NAMES[x] for x in ("ro", "en")])
        self.lang_cb.pack(side="left", padx=4)
        self.lang_cb.bind("<<ComboboxSelected>>", lambda e: self._change_lang())
        dof = ttk.Frame(of)
        dof.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(dof, text=t("ui.doc_to")).pack(side="left")
        self.rb_docx = ttk.Radiobutton(dof, text="DOCX", value="docx", variable=self.doc_format,
                                       command=self._mark_dirty)
        self.rb_pdf = ttk.Radiobutton(dof, text="PDF", value="pdf", variable=self.doc_format,
                                      command=self._mark_dirty)
        self.rb_docx.pack(side="left", padx=4)
        self.rb_pdf.pack(side="left")
        self.doc_note = ttk.Label(of, foreground="gray40", wraplength=330, justify="left",
                                  text=t("ui.doc_note", dir=doc_convert.conversion_dir()))
        self.doc_note.grid(row=6, column=0, columnspan=3, sticky="w")
        self.doc_available = doc_convert.available()
        if not self.doc_available:
            self.doc_note.configure(text=t("ui.doc_unavailable"))
        ttk.Label(of, text=t("ui.manual_terms")).grid(row=2, column=0, columnspan=3, sticky="w")
        self.term_entry = ttk.Entry(of)
        self.term_entry.grid(row=3, column=0, sticky="ew")
        self.term_entry.bind("<Return>", lambda e: self._add_term())
        self.btn_term_add = ttk.Button(of, text=t("ui.term_add"), command=self._add_term)
        self.btn_term_add.grid(row=3, column=1, padx=4)
        self.btn_term_del = ttk.Button(of, text=t("ui.term_del"), command=self._remove_term)
        self.btn_term_del.grid(row=3, column=2)
        self.term_list = tk.Listbox(of, height=3, activestyle="none")
        self.term_list.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(3, 0))

        # --- scanare
        sf = ttk.Frame(self)
        sf.grid(row=1, column=0, sticky="ew", pady=6)
        sf.columnconfigure(2, weight=1)
        self.btn_scan = ttk.Button(sf, text=t("ui.scan"), command=self._scan)
        self.btn_scan.grid(row=0, column=0)
        self.btn_cancel = ttk.Button(sf, text=t("ui.stop"), command=self._cancel)
        self.btn_cancel.grid(row=0, column=1, padx=6)
        self.progress = ttk.Progressbar(sf, mode="determinate")
        self.progress.grid(row=0, column=2, sticky="ew")
        ttk.Label(sf, textvariable=self.status, wraplength=900, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # --- 2. verificare
        vf = ttk.LabelFrame(self, text=t("ui.review"), padding=6)
        vf.grid(row=2, column=0, sticky="nsew")
        vf.columnconfigure(0, weight=1)
        vf.rowconfigure(1, weight=1)
        bar = ttk.Frame(vf)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Label(bar, text=t("ui.filter_file")).pack(side="left")
        self.filter_cb = ttk.Combobox(bar, textvariable=self.filter_var, state="readonly", width=40)
        self.filter_cb.pack(side="left", padx=(4, 12))
        self.filter_cb.bind("<<ComboboxSelected>>", lambda e: self._fill_tree())
        self.btn_val_on = ttk.Button(bar, text=t("ui.val_on"),
                                     command=lambda: self._same_value(True))
        self.btn_val_off = ttk.Button(bar, text=t("ui.val_off"),
                                      command=lambda: self._same_value(False))
        self.btn_all_on = ttk.Button(bar, text=t("ui.all_on"), command=lambda: self._all(True))
        self.btn_all_off = ttk.Button(bar, text=t("ui.all_off"), command=lambda: self._all(False))
        for b in (self.btn_val_on, self.btn_val_off, self.btn_all_on, self.btn_all_off):
            b.pack(side="left", padx=2)
        self.count_lbl = ttk.Label(bar, text="")
        self.count_lbl.pack(side="right")

        self.tree = ttk.Treeview(vf, columns=[c[0] for c in COLUMNS], show="headings", selectmode="extended")
        for cid, title_key, width, stretch in COLUMNS:
            self.tree.heading(cid, text=t(title_key) if title_key else "")
            self.tree.column(cid, width=width, minwidth=30, stretch=stretch,
                             anchor="center" if cid == "sel" else "w")
        self.tree.grid(row=1, column=0, sticky="nsew")
        ys = ttk.Scrollbar(vf, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(vf, orient="horizontal", command=self.tree.xview)
        ys.grid(row=1, column=1, sticky="ns")
        xs.grid(row=2, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.tag_configure("off", foreground="gray55")
        self.tree.tag_configure("low", background="#fff3cd")
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<space>", self._on_space)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._refresh_buttons())

        ttk.Label(vf, text=t("ui.warnings")).grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.warn = tk.Text(vf, height=5, wrap="word", state="disabled")
        self.warn.grid(row=4, column=0, columnspan=2, sticky="ew")

        # --- 3. generare
        gf = ttk.LabelFrame(self, text=t("ui.generate_frame"), padding=6)
        gf.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        gf.columnconfigure(1, weight=1)
        ttk.Label(gf, text=t("ui.out_dir")).grid(row=0, column=0, sticky="w")
        self.out_entry = ttk.Entry(gf, textvariable=self.out_dir)
        self.out_entry.grid(row=0, column=1, sticky="ew", padx=4)
        self.btn_out = ttk.Button(gf, text=t("ui.choose"), command=self._choose_out)
        self.btn_out.grid(row=0, column=2)
        fmt = ttk.Frame(gf)
        fmt.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(fmt, text=t("ui.out_as")).pack(side="left")
        self.rb_out = []
        for val, txt in (("original", t("ui.fmt_original")), ("pdf", "PDF"), ("docx", "DOCX")):
            rb = ttk.Radiobutton(fmt, text=txt, value=val, variable=self.out_format,
                                 command=self._out_format_note)
            rb.pack(side="left", padx=4)
            self.rb_out.append(rb)
        self.out_note = ttk.Label(gf, foreground="gray40", wraplength=900, justify="left", text="")
        self.out_note.grid(row=2, column=0, columnspan=3, sticky="w")
        self.btn_gen = ttk.Button(gf, text=t("ui.generate"), command=self._generate)
        self.btn_gen.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.btn_open = ttk.Button(gf, text=t("ui.open_folder"), command=self._open_out)
        self.btn_open.grid(row=3, column=2, sticky="e", pady=(6, 0))
        ttk.Label(gf, foreground="gray40", text=t("ui.footer")).grid(row=4, column=0, columnspan=3, sticky="w",
                                                                    pady=(4, 0))

    def _out_format_note(self) -> None:
        f = self.out_format.get()
        if f == "original":
            text = ""
        elif not self.doc_available:
            text = t("ui.fmt_note_unavailable")
        elif f == "pdf":
            text = t("ui.fmt_note_pdf")
        else:
            text = t("ui.fmt_note_docx")
        self.out_note.configure(text=text)

    # ================================================================== limbă
    def _state(self) -> dict:
        return {"files": list(self.files), "terms": list(self.terms), "out_dir": self.out_dir.get(),
                "doc_format": self.doc_format.get(), "out_format": self.out_format.get()}

    def _restore(self, state: dict) -> None:
        for term in state.get("terms", []):
            self.terms.append(term)
            self.term_list.insert("end", term)
        self.doc_format.set(state.get("doc_format", "docx"))
        self.out_format.set(state.get("out_format", "original"))
        self.add_paths(state.get("files", []))
        if state.get("out_dir"):
            self.out_dir.set(state["out_dir"])
        self._out_format_note()

    def _change_lang(self) -> None:
        new = {v: k for k, v in LANG_NAMES.items()}[self.lang_var.get()]
        if new == get_lang():
            return
        if self.busy or (self.model.results and not messagebox.askyesno(app_title(), t("ui.lang_rescan"),
                                                                        parent=self.master)):
            self.lang_var.set(LANG_NAMES[get_lang()])
            return
        set_lang(new)
        runtime.save_settings({**runtime.load_settings(), "lang": new})
        state, master = self._state(), self.master
        # reconstruirea se face după ce Tk termină evenimentul combobox-ului (care e distrus odată cu fereastra)
        master.after(0, lambda: (self.destroy(), App(master, state)))


    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(parent=self.master, title=t("ui.pick_files"), filetypes=filetypes())
        self.add_paths(paths)

    def add_paths(self, paths) -> None:
        skipped = []
        known = {os.path.normcase(os.path.abspath(p)) for p in self.files}
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            key = os.path.normcase(os.path.abspath(p))
            if ext not in jobs.INPUT_EXTS:
                skipped.append(os.path.basename(p))
                continue
            if key in known:
                continue
            known.add(key)
            self.files.append(os.path.abspath(p))
            self.file_list.insert("end", os.path.basename(p) + "   —   " + os.path.dirname(os.path.abspath(p)))
        if skipped:
            messagebox.showwarning(app_title(), t("ui.unsupported_skipped") + "\n" + "\n".join(skipped),
                                   parent=self.master)
        if self.files and not self.out_dir.get():
            self.out_dir.set(jobs.default_out_dir(self.files[0]))
        self._mark_dirty()

    def _remove_files(self) -> None:
        for i in sorted(self.file_list.curselection(), reverse=True):
            self.file_list.delete(i)
            del self.files[i]
        self._mark_dirty()

    def _clear_files(self) -> None:
        self.file_list.delete(0, "end")
        self.files.clear()
        self._mark_dirty()

    def _add_term(self) -> None:
        t = self.term_entry.get().strip()
        if t and t not in self.terms:
            self.terms.append(t)
            self.term_list.insert("end", t)
            self._mark_dirty()
        self.term_entry.delete(0, "end")

    def _remove_term(self) -> None:
        for i in sorted(self.term_list.curselection(), reverse=True):
            self.term_list.delete(i)
            del self.terms[i]
        self._mark_dirty()

    def _mark_dirty(self) -> None:
        if self.model.results:
            self.dirty = True
            self.status.set(t("ui.dirty"))
        self._refresh_buttons()

    # ================================================================== scanare
    def _scan(self) -> None:
        if not self.files or self.busy:
            return
        self._previous = self.model if self.model.results else None
        self.model = ReviewModel()
        self.tree.delete(*self.tree.get_children())
        self._set_warn("")
        self._start(jobs.scan_job, list(self.files), list(self.terms), self.doc_format.get())
        self.status.set(t("ui.scanning"))

    def _start(self, target, *args) -> None:
        self.busy = True
        self.cancel = threading.Event()
        self.progress.configure(value=0, maximum=1)
        self._refresh_buttons()
        jobs.start(target, *args, self.q, self.cancel)
        self.after(self.POLL_MS, self._poll)

    def _cancel(self) -> None:
        self.cancel.set()
        self.status.set(t("ui.stopping"))

    def _poll(self) -> None:
        try:
            while True:
                self._handle(self.q.get_nowait())
        except queue.Empty:
            pass
        if self.busy:
            self.after(self.POLL_MS, self._poll)

    def _handle(self, msg) -> None:
        kind = msg[0]
        if kind == "progress":
            _, i, n, name, stage = msg
            self.progress.configure(maximum=max(n, 1), value=i - 1)
            self.status.set(t("ui.progress", stage=stage[:1].upper() + stage[1:], i=i, n=n, name=name))
        elif kind == "file":
            self.model.add_result(msg[1], self._previous)
        elif kind == "redacted":
            _, name, out, problems = msg
            if problems:
                self._append_warn(t("ui.verify_problems", name=name, problems="; ".join(problems)))
        elif kind == "file_error":
            self._append_warn(f"{msg[1]}: {msg[2]}")
        elif kind == "fatal":
            self._finish()
            messagebox.showerror(app_title(), msg[1], parent=self.master)
        elif kind == "done":
            self._finish()
            info = msg[1]
            if info["kind"] == "scan":
                self._after_scan(info)
            else:
                self._after_redact(info)

    def _finish(self) -> None:
        self.busy = False
        self.progress.configure(value=self.progress.cget("maximum"))
        self._refresh_buttons()
        if self.closing:
            self.master.destroy()

    def _after_scan(self, info: dict) -> None:
        self._previous = None
        self.dirty = False
        names = [t("ui.all_files")] + [r.display_name for r in self.model.results]
        self.filter_cb.configure(values=names)
        self.filter_var.set(names[0])
        self._fill_tree()
        lines = []
        for r in self.model.results:
            if r.error:
                lines.append(t("ui.file_error", name=r.display_name, err=r.error))
            lines.extend(f"{r.display_name}: {w}" for w in r.warnings)
        if any(is_link(row.finding) for row in self.model.rows()):
            lines.append(t("ui.link_rows"))
        self._set_warn("\n".join(lines))
        sel, total = self.model.counts()
        errs = sum(1 for r in self.model.results if r.error)
        msg = t("ui.scan_done", files=info["files"], total=total)
        if errs:
            msg += t("ui.scan_errors", n=errs)
        if info.get("cancelled"):
            msg += t("ui.stopped_f")
        self.status.set(msg + t("ui.scan_next"))
        self._refresh_buttons()

    # ================================================================== tabel
    def _current_file_index(self) -> int | None:
        names = list(self.filter_cb.cget("values"))
        v = self.filter_var.get()
        return names.index(v) - 1 if v in names and names.index(v) > 0 else None

    def _fill_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for row in self.model.rows(self._current_file_index()):
            self.tree.insert("", "end", iid=row.iid, values=self._values(row), tags=self._tags(row.finding))
        self._update_count()

    def _values(self, row) -> tuple:
        f = row.finding
        note = f.note
        if is_link(f):
            note = (note + " | " if note else "") + t("ui.link_note")
        return (ON if f.selected else OFF, self.model.results[row.file_index].display_name, kind_label(f.kind),
                f.value, conf_label(f.confidence.value), f.location, f.context, note)

    @staticmethod
    def _tags(f) -> tuple:
        tags = []
        if not f.selected:
            tags.append("off")
        if f.confidence == Confidence.MICA:
            tags.append("low")
        return tuple(tags)

    def _refresh_rows(self, iids) -> None:
        for iid in iids:
            if self.tree.exists(iid):
                row = self.model.get(iid)
                self.tree.item(iid, values=self._values(row), tags=self._tags(row.finding))
        self._update_count()

    def _update_count(self) -> None:
        sel, total = self.model.counts()
        self.count_lbl.configure(text=t("ui.ticked", sel=sel, total=total))

    def _on_click(self, event):
        if self.busy:
            return None
        if self.tree.identify_region(event.x, event.y) == "cell" and self.tree.identify_column(event.x) == "#1":
            iid = self.tree.identify_row(event.y)
            if iid:
                self._refresh_rows(self.model.toggle(iid))
                return "break"
        return None

    def _on_space(self, event):
        if self.busy:
            return "break"
        for iid in self.tree.selection():
            self._refresh_rows(self.model.toggle(iid))
        return "break"

    def _same_value(self, state: bool) -> None:
        for iid in self.tree.selection():
            self._refresh_rows(self.model.set_same_value(iid, state))

    def _all(self, state: bool) -> None:
        self._refresh_rows(self.model.set_all(state, self._current_file_index()))

    # ================================================================== generare
    def _choose_out(self) -> None:
        d = filedialog.askdirectory(parent=self.master, title=t("ui.choose_out"),
                                    initialdir=os.path.dirname(self.out_dir.get()) or None)
        if d:
            self.out_dir.set(d)

    def _generate(self) -> None:
        if self.busy or not self.model.results or self.dirty:
            return
        out = self.out_dir.get().strip()
        if not out:
            messagebox.showwarning(app_title(), t("ui.need_out"), parent=self.master)
            return
        out = os.path.abspath(out)
        sources = {os.path.normcase(os.path.dirname(os.path.abspath(r.source or r.file))) for r in self.model.results}
        if os.path.normcase(out) in sources:
            if not messagebox.askyesno(app_title(), t("ui.same_folder"), parent=self.master):
                return
        ok = self.model.processable()
        sel, total = self.model.counts()
        skipped = len(self.model.results) - len(ok)
        fmt_txt = t("ui.fmt_txt_" + self.out_format.get())
        text = t("ui.confirm_generate", n=len(ok), fmt=fmt_txt, sel=sel, total=total, out=out)
        if skipped:
            text += t("ui.confirm_skipped", n=skipped)
        if not messagebox.askyesno(app_title(), text + t("ui.continue_q"), parent=self.master):
            return
        self.last_out_dir = out
        self._start(functools.partial(jobs.redact_job, out_format=self.out_format.get()), self.model.results, out)
        self.status.set(t("ui.generating"))

    def _after_redact(self, info: dict) -> None:
        msg = t("ui.gen_done", n=info["done"])
        if info["failed"]:
            msg += t("ui.gen_failed", n=info["failed"])
        if info["problems"]:
            msg += t("ui.gen_problems", n=info["problems"])
        if info.get("cancelled"):
            msg += t("ui.stopped_m")
        msg += t("ui.gen_report", name=os.path.basename(info["report"]))
        self.status.set(msg.replace("\n", " "))
        show = messagebox.showwarning if (info["failed"] or info["problems"]) else messagebox.showinfo
        show(app_title(), msg, parent=self.master)
        self._refresh_buttons()

    def _open_out(self) -> None:
        d = self.last_out_dir or self.out_dir.get()
        if d and os.path.isdir(d) and hasattr(os, "startfile"):
            os.startfile(d)  # noqa: S606 - deschide Explorer pe un folder ales de utilizator

    # ================================================================== stări
    def _refresh_buttons(self) -> None:
        idle = not self.busy
        has_results = bool(self.model.results)
        has_sel = bool(self.tree.selection())
        st = {True: "normal", False: "disabled"}
        for b in (self.btn_add, self.btn_remove, self.btn_clear, self.btn_term_add, self.btn_term_del,
                  self.btn_out):
            b.configure(state=st[idle])
        for rb in (self.rb_docx, self.rb_pdf):
            rb.configure(state=st[idle and self.doc_available])
        for rb in self.rb_out:
            rb.configure(state=st[idle and (self.doc_available or rb.cget("value") == "original")])
        self.btn_scan.configure(state=st[idle and bool(self.files)])
        self.btn_cancel.configure(state=st[self.busy])
        self.btn_gen.configure(state=st[idle and has_results and not self.dirty])
        for b in (self.btn_val_on, self.btn_val_off):
            b.configure(state=st[idle and has_sel])
        for b in (self.btn_all_on, self.btn_all_off):
            b.configure(state=st[idle and has_results])
        out_ok = bool(self.last_out_dir and os.path.isdir(self.last_out_dir))
        self.btn_open.configure(state=st[idle and out_ok and hasattr(os, "startfile")])

    def _set_warn(self, text: str) -> None:
        self.warn.configure(state="normal")
        self.warn.delete("1.0", "end")
        self.warn.insert("1.0", text)
        self.warn.configure(state="disabled")

    def _append_warn(self, line: str) -> None:
        self.warn.configure(state="normal")
        if self.warn.get("1.0", "end").strip():
            self.warn.insert("end", "\n")
        self.warn.insert("end", line)
        self.warn.configure(state="disabled")
        self.warn.see("end")

    def _on_close(self) -> None:
        if not self.busy:
            self.master.destroy()
            return
        if messagebox.askyesno(app_title(), t("ui.busy_close"), parent=self.master):
            self.closing = True
            self.cancel.set()


def main() -> None:
    from . import runtime
    log_path = runtime.setup_logging()
    if "--selftest" in sys.argv[1:]:
        sys.exit(runtime.run_selftest_cli())
    set_lang(runtime.resolve_lang(sys.argv[1:]))
    _dpi_aware()
    root = tk.Tk()

    def on_callback_error(exc, val, tb):
        # în .exe fără consolă, Tk ar scrie eroarea în stderr inexistent: o logăm și o arătăm
        import traceback
        log.error("Eroare în interfață:\n%s", "".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror(app_title(), t("ui.unexpected", name=exc.__name__, path=log_path), parent=root)
    root.report_callback_exception = on_callback_error
    root.minsize(980, 640)
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    # rânduri cu înălțimea fontului (la scalare 125-150% rândurile implicite se suprapun)
    import tkinter.font as tkfont
    lh = tkfont.nametofont("TkDefaultFont").metrics("linespace")
    style.configure("Treeview", rowheight=int(lh * 1.4))
    app = App(root)
    paths = runtime.file_args(sys.argv[1:])
    if paths:
        app.add_paths(paths)
    root.mainloop()


if __name__ == "__main__":
    main()
