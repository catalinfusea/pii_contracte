# Raportarea problemelor de securitate și de redactare

pii_contracte prelucrează documente cu date personale. Două tipuri de probleme contează în mod special:

1. **Scăpări de redactare** — o dată de identificare care rămâne în copia redactată (nedetectată sau
   detectată dar neeliminată), inclusiv în metadate, linkuri, antete/subsoluri sau în exportul PDF/DOCX.
2. **Vulnerabilități** — de exemplu un document construit special care face programul să citească fișiere
   locale, să acceseze rețeaua, să ruleze macro-uri în Word sau să suprascrie un original.

## Cum raportați
- **Nu atașați niciodată documente reale** și nu copiați date reale în descriere.
- Reproduceți problema cu **date inventate** (un CNP de test cu cifră de control validă se poate genera ca în
  `tests/fixtures.py`, funcția `gen_cnp`) și descrieți formularea exactă din jurul valorii ratate.
- Scăpările de redactare pot fi raportate ca issue în acest repository, cu textul sintetic care reproduce problema.
- Pentru o vulnerabilitate, folosiți raportarea privată a GitHub (fila **Security** → *Report a vulnerability*),
  dacă este activată pentru repository; altfel deschideți un issue fără detalii tehnice și cereți un canal privat.

## Ce nu este o vulnerabilitate
- Valorile complete afișate în fereastra aplicației (sunt necesare pentru verificare).
- Copiile neredactate din folderul `Conversie\` (comportament documentat).
- Textul din imagini și din pagini scanate (limitare documentată: fără OCR).

Programul este oferit fără garanție (vezi LICENSE). Nu există un termen garantat de răspuns.
