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
from pii_core.models import Confidence

from . import jobs
from .model import LINK_NOTE, ReviewModel, is_link

log = logging.getLogger("pii_core")

APP_TITLE = "Redactare date de identificare - contracte"
ON, OFF = "\u2611", "\u2610"   # ☑ ☐
FILETYPES = [("Contracte (Word, PDF)", "*.docx *.doc *.pdf"), ("Word", "*.docx *.doc"),
             ("PDF", "*.pdf"), ("Toate fișierele", "*.*")]
COLUMNS = [  # id, titlu, lățime, stretch
    ("sel", "", 34, False), ("file", "Fișier", 150, False), ("kind", "Tip", 160, False),
    ("value", "Valoare", 200, True), ("conf", "Încredere", 75, False), ("loc", "Localizare", 130, False),
    ("ctx", "Context", 320, True), ("note", "Notă", 220, True),
]


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

    def __init__(self, master: tk.Tk):
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
        self.status = tk.StringVar(value="Adăugați fișiere (.docx, .doc, .pdf), apoi apăsați „Scanează”.")
        self._build()
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        master.protocol("WM_DELETE_WINDOW", self._on_close)
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
        ff = ttk.LabelFrame(top, text="1. Fișiere", padding=6)
        ff.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ff.columnconfigure(0, weight=1)
        self.file_list = tk.Listbox(ff, height=6, selectmode="extended", activestyle="none")
        self.file_list.grid(row=0, column=0, rowspan=3, sticky="nsew")
        sb = ttk.Scrollbar(ff, orient="vertical", command=self.file_list.yview)
        sb.grid(row=0, column=1, rowspan=3, sticky="ns")
        self.file_list.configure(yscrollcommand=sb.set)
        self.btn_add = ttk.Button(ff, text="Adaugă…", command=self._add_files)
        self.btn_add.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        self.btn_remove = ttk.Button(ff, text="Elimină selectate", command=self._remove_files)
        self.btn_remove.grid(row=1, column=2, sticky="ew", padx=(6, 0), pady=3)
        self.btn_clear = ttk.Button(ff, text="Golește lista", command=self._clear_files)
        self.btn_clear.grid(row=2, column=2, sticky="new", padx=(6, 0))

        # --- opțiuni
        of = ttk.LabelFrame(top, text="Opțiuni", padding=6)
        of.grid(row=0, column=1, sticky="nsew")
        of.columnconfigure(0, weight=1)
        dof = ttk.Frame(of)
        dof.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(dof, text="Fișierele .doc se convertesc în:").pack(side="left")
        self.rb_docx = ttk.Radiobutton(dof, text="DOCX", value="docx", variable=self.doc_format,
                                       command=self._mark_dirty)
        self.rb_pdf = ttk.Radiobutton(dof, text="PDF", value="pdf", variable=self.doc_format,
                                      command=self._mark_dirty)
        self.rb_docx.pack(side="left", padx=4)
        self.rb_pdf.pack(side="left")
        self.doc_note = ttk.Label(of, foreground="gray40", wraplength=330, justify="left",
                                  text=f"Conversia (neredactată) se salvează în „{doc_convert.CONVERSION_DIR}” "
                                       "lângă original și rămâne acolo.")
        self.doc_note.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))
        self.doc_available = doc_convert.available()
        if not self.doc_available:
            self.doc_note.configure(text="Conversia .doc nu e disponibilă (lipsește pywin32 sau nu e Windows). "
                                         "Fișierele .doc vor apărea cu eroare.")
        ttk.Label(of, text="Text de redactat manual (ex. un nume ratat):").grid(row=2, column=0, columnspan=3,
                                                                                sticky="w")
        self.term_entry = ttk.Entry(of)
        self.term_entry.grid(row=3, column=0, sticky="ew")
        self.term_entry.bind("<Return>", lambda e: self._add_term())
        self.btn_term_add = ttk.Button(of, text="Adaugă", command=self._add_term)
        self.btn_term_add.grid(row=3, column=1, padx=4)
        self.btn_term_del = ttk.Button(of, text="Șterge", command=self._remove_term)
        self.btn_term_del.grid(row=3, column=2)
        self.term_list = tk.Listbox(of, height=3, activestyle="none")
        self.term_list.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(3, 0))

        # --- scanare
        sf = ttk.Frame(self)
        sf.grid(row=1, column=0, sticky="ew", pady=6)
        sf.columnconfigure(2, weight=1)
        self.btn_scan = ttk.Button(sf, text="Scanează", command=self._scan)
        self.btn_scan.grid(row=0, column=0)
        self.btn_cancel = ttk.Button(sf, text="Oprește", command=self._cancel)
        self.btn_cancel.grid(row=0, column=1, padx=6)
        self.progress = ttk.Progressbar(sf, mode="determinate")
        self.progress.grid(row=0, column=2, sticky="ew")
        ttk.Label(sf, textvariable=self.status, wraplength=900, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # --- 2. verificare
        vf = ttk.LabelFrame(self, text="2. Verificare - debifați ce NU trebuie redactat "
                                       "(click pe căsuță sau tasta Spațiu)", padding=6)
        vf.grid(row=2, column=0, sticky="nsew")
        vf.columnconfigure(0, weight=1)
        vf.rowconfigure(1, weight=1)
        bar = ttk.Frame(vf)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Label(bar, text="Fișier:").pack(side="left")
        self.filter_cb = ttk.Combobox(bar, textvariable=self.filter_var, state="readonly", width=40)
        self.filter_cb.pack(side="left", padx=(4, 12))
        self.filter_cb.bind("<<ComboboxSelected>>", lambda e: self._fill_tree())
        self.btn_val_on = ttk.Button(bar, text="Bifează valoarea peste tot",
                                     command=lambda: self._same_value(True))
        self.btn_val_off = ttk.Button(bar, text="Debifează valoarea peste tot",
                                      command=lambda: self._same_value(False))
        self.btn_all_on = ttk.Button(bar, text="Bifează tot", command=lambda: self._all(True))
        self.btn_all_off = ttk.Button(bar, text="Debifează tot", command=lambda: self._all(False))
        for b in (self.btn_val_on, self.btn_val_off, self.btn_all_on, self.btn_all_off):
            b.pack(side="left", padx=2)
        self.count_lbl = ttk.Label(bar, text="")
        self.count_lbl.pack(side="right")

        self.tree = ttk.Treeview(vf, columns=[c[0] for c in COLUMNS], show="headings", selectmode="extended")
        for cid, title, width, stretch in COLUMNS:
            self.tree.heading(cid, text=title)
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

        ttk.Label(vf, text="Avertismente și erori (de verificat manual):").grid(row=3, column=0, sticky="w",
                                                                              pady=(6, 0))
        self.warn = tk.Text(vf, height=5, wrap="word", state="disabled")
        self.warn.grid(row=4, column=0, columnspan=2, sticky="ew")

        # --- 3. generare
        gf = ttk.LabelFrame(self, text="3. Generare", padding=6)
        gf.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        gf.columnconfigure(1, weight=1)
        ttk.Label(gf, text="Folder ieșire:").grid(row=0, column=0, sticky="w")
        self.out_entry = ttk.Entry(gf, textvariable=self.out_dir)
        self.out_entry.grid(row=0, column=1, sticky="ew", padx=4)
        self.btn_out = ttk.Button(gf, text="Alege…", command=self._choose_out)
        self.btn_out.grid(row=0, column=2)
        fmt = ttk.Frame(gf)
        fmt.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(fmt, text="Copii redactate ca:").pack(side="left")
        self.rb_out = []
        for val, txt in (("original", "formatul original"), ("pdf", "PDF"), ("docx", "DOCX")):
            rb = ttk.Radiobutton(fmt, text=txt, value=val, variable=self.out_format,
                                 command=self._out_format_note)
            rb.pack(side="left", padx=4)
            self.rb_out.append(rb)
        self.out_note = ttk.Label(gf, foreground="gray40", wraplength=900, justify="left", text="")
        self.out_note.grid(row=2, column=0, columnspan=3, sticky="w")
        self.btn_gen = ttk.Button(gf, text="Generează copii redactate și raport", command=self._generate)
        self.btn_gen.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.btn_open = ttk.Button(gf, text="Deschide folderul", command=self._open_out)
        self.btn_open.grid(row=3, column=2, sticky="e", pady=(6, 0))
        ttk.Label(gf, foreground="gray40",
                  text="Originalele nu se modifică. Valorile complete apar doar în această fereastră; "
                       "raportul CSV conține valori mascate.").grid(row=4, column=0, columnspan=3, sticky="w",
                                                                    pady=(4, 0))

    def _out_format_note(self) -> None:
        f = self.out_format.get()
        if f == "original":
            text = ""
        elif not self.doc_available:
            text = "Exportul în alt format folosește Word și nu e disponibil (lipsește pywin32 sau nu e Windows)."
        elif f == "pdf":
            text = ("Documentele Word se redactează în DOCX, apoi Word le exportă în PDF. Se păstrează ambele copii; "
                    "ambele sunt verificate.")
        else:
            text = ("PDF-urile se redactează ca PDF, apoi Word le convertește în DOCX (PDF Reflow) - aspectul se poate "
                    "strica. Se păstrează ambele copii; ambele sunt verificate.")
        self.out_note.configure(text=text)

    # ================================================================== fișiere / termeni
    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(parent=self.master, title="Alegeți contractele", filetypes=FILETYPES)
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
            messagebox.showwarning(APP_TITLE, "Format nesuportat, ignorate:\n" + "\n".join(skipped),
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
            self.status.set("Lista de fișiere, termenii sau formatul .doc s-au schimbat - apăsați din nou "
                            "„Scanează” (debifările existente se păstrează).")
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
        self.status.set("Scanare…")

    def _start(self, target, *args) -> None:
        self.busy = True
        self.cancel = threading.Event()
        self.progress.configure(value=0, maximum=1)
        self._refresh_buttons()
        jobs.start(target, *args, self.q, self.cancel)
        self.after(self.POLL_MS, self._poll)

    def _cancel(self) -> None:
        self.cancel.set()
        self.status.set("Se oprește după fișierul curent…")

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
            self.status.set(f"{stage.capitalize()} {i}/{n}: {name}")
        elif kind == "file":
            self.model.add_result(msg[1], self._previous)
        elif kind == "redacted":
            _, name, out, problems = msg
            if problems:
                self._append_warn(f"{name}: VERIFICAREA A GĂSIT PROBLEME în copia redactată - "
                                  + "; ".join(problems))
        elif kind == "file_error":
            self._append_warn(f"{msg[1]}: {msg[2]}")
        elif kind == "fatal":
            self._finish()
            messagebox.showerror(APP_TITLE, msg[1], parent=self.master)
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
        names = ["Toate fișierele"] + [r.display_name for r in self.model.results]
        self.filter_cb.configure(values=names)
        self.filter_var.set(names[0])
        self._fill_tree()
        lines = []
        for r in self.model.results:
            if r.error:
                lines.append(f"{r.display_name}: EROARE - {r.error}")
            lines.extend(f"{r.display_name}: {w}" for w in r.warnings)
        if any(is_link(row.finding) for row in self.model.rows()):
            lines.append("Rândurile cu localizarea „hyperlink”: " + LINK_NOTE)
        self._set_warn("\n".join(lines))
        sel, total = self.model.counts()
        errs = sum(1 for r in self.model.results if r.error)
        msg = f"Scanare terminată: {info['files']} fișiere, {total} constatări"
        if errs:
            msg += f", {errs} fișiere cu eroare (vezi mai jos)"
        if info.get("cancelled"):
            msg += " - OPRITĂ înainte de final"
        self.status.set(msg + ". Verificați tabelul, apoi generați copiile.")
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
            note = (note + " | " if note else "") + "link (vezi nota de jos)"
        return (ON if f.selected else OFF, self.model.results[row.file_index].display_name, f.kind,
                f.value, f.confidence.value, f.location, f.context, note)

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
        self.count_lbl.configure(text=f"Bifate: {sel}/{total}")

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
        d = filedialog.askdirectory(parent=self.master, title="Folder pentru copiile redactate",
                                    initialdir=os.path.dirname(self.out_dir.get()) or None)
        if d:
            self.out_dir.set(d)

    def _generate(self) -> None:
        if self.busy or not self.model.results or self.dirty:
            return
        out = self.out_dir.get().strip()
        if not out:
            messagebox.showwarning(APP_TITLE, "Alegeți folderul de ieșire.", parent=self.master)
            return
        out = os.path.abspath(out)
        sources = {os.path.normcase(os.path.dirname(os.path.abspath(r.source or r.file))) for r in self.model.results}
        if os.path.normcase(out) in sources:
            if not messagebox.askyesno(APP_TITLE, "Folderul de ieșire e același cu al originalelor. Copiile "
                                       "redactate (…_REDACTAT) vor sta lângă originale și pot fi confundate.\n\n"
                                       "Continuați?", parent=self.master):
                return
        ok = self.model.processable()
        sel, total = self.model.counts()
        skipped = len(self.model.results) - len(ok)
        fmt_txt = {"original": "în formatul original", "pdf": "în PDF (plus copia redactată originală)",
                   "docx": "în DOCX (plus copia redactată originală)"}[self.out_format.get()]
        text = (f"Se generează {len(ok)} copii redactate {fmt_txt} ({sel} din {total} constatări bifate) "
                f"și raportul CSV în:\n{out}")
        if skipped:
            text += f"\n\n{skipped} fișiere cu eroare NU vor avea copie redactată."
        if not messagebox.askyesno(APP_TITLE, text + "\n\nContinuați?", parent=self.master):
            return
        self.last_out_dir = out
        self._start(functools.partial(jobs.redact_job, out_format=self.out_format.get()), self.model.results, out)
        self.status.set("Generare…")

    def _after_redact(self, info: dict) -> None:
        msg = f"Generate: {info['done']} copii redactate"
        if info["failed"]:
            msg += f", {info['failed']} eșuate"
        if info["problems"]:
            msg += f", {info['problems']} cu probleme la verificare (vezi lista de avertismente)"
        if info.get("cancelled"):
            msg += " - OPRIT înainte de final"
        msg += f".\nRaport: {os.path.basename(info['report'])}"
        self.status.set(msg.replace("\n", " "))
        show = messagebox.showwarning if (info["failed"] or info["problems"]) else messagebox.showinfo
        show(APP_TITLE, msg, parent=self.master)
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
        if messagebox.askyesno(APP_TITLE, "O operație e în curs. Se oprește după fișierul curent și "
                               "fereastra se închide. Continuați?", parent=self.master):
            self.closing = True
            self.cancel.set()


def main() -> None:
    from . import runtime
    log_path = runtime.setup_logging()
    if "--selftest" in sys.argv[1:]:
        sys.exit(runtime.run_selftest_cli())
    _dpi_aware()
    root = tk.Tk()

    def on_callback_error(exc, val, tb):
        # în .exe fără consolă, Tk ar scrie eroarea în stderr inexistent: o logăm și o arătăm
        import traceback
        log.error("Eroare în interfață:\n%s", "".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror(APP_TITLE, f"Eroare neașteptată ({exc.__name__}).\nDetalii în:\n{log_path}",
                             parent=root)
    root.report_callback_exception = on_callback_error
    root.title(APP_TITLE)
    root.minsize(980, 640)
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    # rânduri cu înălțimea fontului (la scalare 125-150% rândurile implicite se suprapun)
    import tkinter.font as tkfont
    lh = tkfont.nametofont("TkDefaultFont").metrics("linespace")
    style.configure("Treeview", rowheight=int(lh * 1.4))
    app = App(root)
    paths = [a for a in sys.argv[1:] if not a.startswith("--")]
    if paths:
        app.add_paths(paths)
    root.mainloop()


if __name__ == "__main__":
    main()
