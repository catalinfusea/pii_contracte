"""Diagnostic pentru IBAN-uri ratate - afișează DOAR structura, nu valorile.
Uz:  python diag_iban.py <fișier.docx|.pdf>
Litere -> A/a, cifre -> 9, spațiu -> ·, spațiu neseparabil -> ⍽, alte spații -> <U+XXXX>.
Cuvintele uzuale (IBAN, cont, nr...) și prefixul de țară rămân vizibile."""
import os
import re
import sys

from pii_core import engine
from pii_core.detectors import find_iban, iban_ok

KEEP = {"iban", "cont", "contul", "nr", "bancar", "deschis", "la", "banca", "în", "in", "cod", "swift", "bic"}
TRIGGER = re.compile(r"(?i)\biban\b|(?<![A-Za-z])ro[\s\u00a0.\-]*\d{2}")


def shape(s: str) -> str:
    out = []
    for tok in re.split(r"(\s+)", s):
        if not tok:
            continue
        if tok.isspace():
            out.append("".join("·" if c == " " else "⍽" if c == "\u00a0" else f"<U+{ord(c):04X}>" for c in tok))
            continue
        core = tok.strip(".,:;()")
        if core.lower() in KEEP:
            out.append(tok)
            continue
        m = re.match(r"(?i)^(ro)(?=\d)", tok)
        prefix = m.group(1) if m else ""
        rest = tok[len(prefix):]
        out.append(prefix + "".join("9" if c.isdigit() else ("A" if c.isupper() else "a") if c.isalpha() else c
                                    for c in rest))
    return "".join(out)


def main(path: str) -> None:
    ext = os.path.splitext(path)[1].lower()
    units, _ = engine._units(path, ext)
    n = 0
    for text, label, _ in units:
        for m in TRIGGER.finditer(text):
            n += 1
            win = text[m.start():m.start() + 60]
            ro = re.search(r"(?i)ro", win)
            compact = re.sub(r"[^A-Za-z0-9]", "", win[ro.start():])[:24] if ro else ""
            print(f"[{n}] {label}")
            print(f"    structură : {shape(win)!s}")
            print(f"    detectat  : {'DA' if find_iban(text[m.start():m.start() + 60]) else 'NU'}")
            if len(compact) == 24:
                print(f"    mod-97 pe primele 24 caractere alfanumerice după RO: "
                      f"{'valid' if iban_ok(compact.upper()) else 'INVALID'}"
                      f"{' (litere mici)' if compact != compact.upper() else ''}")
    if not n:
        print("Niciun 'IBAN' sau 'RO99...' găsit în textul extras (poate e în imagine/scan sau câmp de formular).")


if __name__ == "__main__":
    for p in sys.argv[1:]:
        print(f"== {os.path.basename(p)}")
        main(p)
