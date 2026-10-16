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

"""Smoke test pentru fereastră (necesită display; pe Linux: xvfb-run). Se sare dacă nu există Tk/display."""
import os
import sys
import time

import pytest

import fixtures as fx

tk = pytest.importorskip("tkinter")


@pytest.fixture
def app(monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError as e:
        if sys.platform == "win32":          # pe Windows trebuie să meargă: arată eroarea reală
            pytest.fail(f"Tk nu pornește: {e}")
        pytest.skip(f"fără display: {e}")
    from pii_gui import app as appmod
    monkeypatch.setattr(appmod.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(appmod.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(appmod.messagebox, "showwarning", lambda *a, **k: None)
    a = appmod.App(root)
    root.update()
    yield a
    root.destroy()


def pump(app, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        app.master.update()
        if not app.busy:
            return
        time.sleep(0.02)
    raise TimeoutError


def test_full_flow(app, tmp_path):
    fx.make_docx(str(tmp_path / "a.docx"))
    fx.make_pdf(str(tmp_path / "b.pdf"))
    (tmp_path / "note.txt").write_text("x")
    app.add_paths([str(tmp_path / "a.docx"), str(tmp_path / "b.pdf"), str(tmp_path / "a.docx")])
    assert len(app.files) == 2                                   # duplicat ignorat
    assert str(app.btn_gen.cget("state")) == "disabled"

    app._scan()
    pump(app)
    rows = app.tree.get_children()
    assert len(rows) == app.model.counts()[1] > 0
    assert str(app.btn_gen.cget("state")) == "normal"
    assert "modificări urmărite" in app.warn.get("1.0", "end")

    # click pe căsuța primului rând -> debifat
    first = rows[0]
    app.tree.see(first)
    app.master.update()
    x, y, w, h = app.tree.bbox(first, "#1")
    app.tree.event_generate("<Button-1>", x=x + w // 2, y=y + h // 2)
    app.master.update()
    assert app.model.get(first).finding.selected is False
    assert app.tree.set(first, "sel") == "\u2610"

    # filtru pe fișier
    app.filter_var.set(app.filter_cb.cget("values")[2])
    app._fill_tree()
    assert all(i.startswith("1:") for i in app.tree.get_children())

    # termen nou -> trebuie rescanat; debifarea se păstrează
    app.term_entry.insert(0, "1.000 lei")
    app._add_term()
    assert app.dirty and str(app.btn_gen.cget("state")) == "disabled"
    app._scan()
    pump(app)
    assert not app.dirty
    assert app.model.get(first).finding.selected is False
    assert any(r.finding.kind == "Manual" for r in app.model.rows())

    out = tmp_path / "iesire"
    app.out_dir.set(str(out))
    app._generate()
    pump(app)
    names = os.listdir(out)
    assert "a_REDACTAT.docx" in names and "b_REDACTAT.pdf" in names
    assert any(n.startswith("raport_") for n in names)
    assert "VERIFICAREA A GĂSIT PROBLEME" not in app.warn.get("1.0", "end")
