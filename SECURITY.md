# Reporting security and redaction problems

[Română](SECURITY.ro.md) · **English**

pii_contracte processes documents that contain personal data. Two kinds of problems matter most:

1. **Redaction leaks** — identification data left in the redacted copy (not detected, or detected but not
   removed), including in metadata, links, headers/footers or in the PDF/DOCX export.
2. **Vulnerabilities** — for example a crafted document that makes the program read local files, access the
   network, run macros in Word or overwrite an original.

## How to report
- **Never attach real documents** and never paste real data into the description.
- Reproduce the problem with **made-up data** (a test CNP with a valid check digit can be generated like
  `gen_cnp` in `tests/fixtures.py`) and describe the exact wording around the missed value.
- Redaction leaks can be reported as an issue in this repository, with the synthetic text that reproduces them.
- For a vulnerability, use GitHub private vulnerability reporting (the **Security** tab → *Report a vulnerability*)
  if it is enabled for this repository; otherwise open an issue without technical details and ask for a private
  channel.

## Not vulnerabilities
- Full values shown in the application window (needed for review).
- Unredacted copies in the `Converted\` / `Conversie\` folder (documented behaviour).
- Text in images and scanned pages (documented limitation: no OCR).

The program is provided without warranty (see LICENSE). There is no guaranteed response time.
