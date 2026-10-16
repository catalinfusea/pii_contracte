pii_contracte - redacting personal identification data in contracts

Start: double-click pii_contracte.exe. Interface language: "Language / Limba" (top right).

1. Add the files (.docx, .doc, .pdf). Originals are NOT modified.
2. Press "Scan". Review the table: untick what must NOT be redacted.
   Grey rows (indirect data: place of birth, issuer, document dates) are unticked by default.
3. Choose the output folder and format, then "Generate redacted copies and report".
4. Read the "Warnings and errors" list: everything listed there must be checked manually.

The program does NOT see text inside images or scanned pages. Always check the redacted copy before
sharing it.

.doc files are converted with Microsoft Word; the converted (UNREDACTED) copy stays in the "Converted"
folder next to the original.

Detection is designed for Romanian documents (CNP, Romanian ID cards, RO IBANs, Romanian wording).

Self-test: pii_contracte.exe --selftest (the result appears in a window).
Log for troubleshooting: %LOCALAPPDATA%\pii_contracte\pii_contracte.log (no detected values;
may contain file names).
