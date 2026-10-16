"""Demo pentru validarea etapei 1 pe documente reale, înainte de interfață.
Uz:  python demo_cli.py <fișier.docx|.pdf> [...] --out <folder> [--redact] [--term "text"]
Nu modifică originalele. Afișează valorile complete DOAR în consolă (local)."""
import argparse
import os
import sys

from pii_core import engine

ap = argparse.ArgumentParser()
ap.add_argument("files", nargs="+")
ap.add_argument("--out", required=True)
ap.add_argument("--redact", action="store_true")
ap.add_argument("--term", action="append", default=[], help="text de redactat manual (repetabil)")
a = ap.parse_args()
os.makedirs(a.out, exist_ok=True)
results = []
for f in a.files:
    r = engine.scan_file(f, a.term)
    results.append(r)
    print(f"\n== {f}")
    if r.error:
        print("  EROARE:", r.error)
        continue
    for w in r.warnings:
        print("  ! ", w)
    for x in r.findings:
        print(f"  [{x.confidence.value:5}] {x.kind:22} {x.value!r:40} {x.location} {x.note}")
    if a.redact:
        out, problems = engine.redact_file(r, a.out)
        print("  -> ", out, "| verificare:", "OK" if not problems else problems)
engine.write_report(results, os.path.join(a.out, "raport.csv"))
print("\nRaport:", os.path.join(a.out, "raport.csv"))
sys.exit(0)
