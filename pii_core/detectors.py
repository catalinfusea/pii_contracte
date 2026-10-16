"""Detectoare de date de identificare pentru texte în limba română.

Principiu: validare algoritmică acolo unde există (CNP, IBAN), ancore de context
în rest. Detectoarele lucrează pe un singur șir de text și întorc offset-uri.
"""
from __future__ import annotations

import re
from datetime import date

from .models import Confidence, Match

# --- litere românești (ambele variante: virgulă și sedilă) ---
UP = "A-ZĂÂÎȘȚŞŢ"
LO = "a-zăâîșțşţ"

# ---------------------------------------------------------------- CNP
_CNP_KEY = "279146358279"
# Coduri de județ cunoscute: 01-46 (41-46 = sectoarele 1-6), 47-48 (fostele
# sectoare 7-8), 51 Călărași, 52 Giurgiu. Alt cod nu invalidează CNP-ul,
# doar scade încrederea.
_CNP_COUNTIES = {f"{i:02d}" for i in range(1, 49)} | {"51", "52"}
_CNP_RE = re.compile(r"(?<!\d)([1-9]\d{12})(?!\d)")
_CNP_ANCHOR_RE = re.compile(r"(?i)\b(?:CNP|C\.N\.P\.?|cod\s+numeric\s+personal)\b")


def cnp_checksum_ok(cnp: str) -> bool:
    if len(cnp) != 13 or not cnp.isdigit():
        return False
    s = sum(int(a) * int(b) for a, b in zip(cnp[:12], _CNP_KEY)) % 11
    return (1 if s == 10 else s) == int(cnp[12])


def _cnp_date_ok(cnp: str) -> bool:
    s, yy, mm, dd = int(cnp[0]), int(cnp[1:3]), int(cnp[3:5]), int(cnp[5:7])
    centuries = {1: [1900], 2: [1900], 3: [1800], 4: [1800], 5: [2000], 6: [2000]}
    for c in centuries.get(s, [1900, 2000]):  # 7, 8, 9: secol neprecizat
        try:
            date(c + yy, mm, dd)
            return True
        except ValueError:
            continue
    return False


def find_cnp(text: str) -> list[Match]:
    out = []
    for m in _CNP_RE.finditer(text):
        cnp = m.group(1)
        before = text[max(0, m.start() - 40):m.start()]
        anchored = bool(_CNP_ANCHOR_RE.search(before))
        if cnp_checksum_ok(cnp) and _cnp_date_ok(cnp):
            if cnp[7:9] in _CNP_COUNTIES:
                out.append(Match("CNP", cnp, m.start(), m.end(), Confidence.MARE))
            else:
                out.append(Match("CNP", cnp, m.start(), m.end(), Confidence.MEDIE,
                                 f"cod județ {cnp[7:9]} necunoscut"))
        elif anchored:
            out.append(Match("CNP", cnp, m.start(), m.end(), Confidence.MICA,
                             "cifra de control sau data nu se validează"))
    return out


# ---------------------------------------------------------------- IBAN
# separatori acceptați între grupe: spațiu, spațiu neseparabil (Word), cratimă, punct
_SEP = r"[ \u00a0.\-]"
_IBAN_RE = re.compile(
    rf"(?i)(?<![A-Z0-9])([A-Z]{{2}}{_SEP}?\d{{2}}(?:{_SEP}?[A-Z0-9]{{4}}){{3,7}}(?:{_SEP}?[A-Z0-9]{{1,4}})?)(?![A-Z0-9])"
)
_IBAN_ANCHOR_RE = re.compile(r"(?i)\b(?:IBAN|cont(?:ul)?(?:\s+bancar)?|cod\s+IBAN)\b")


def _compact(iban: str) -> str:
    return re.sub(r"[ \u00a0.\-]", "", iban).upper()


def iban_ok(iban: str) -> bool:
    s = _compact(iban)
    if not 15 <= len(s) <= 34 or not s[:2].isalpha() or not s[2:4].isdigit():
        return False
    if s.startswith("RO") and len(s) != 24:
        return False
    rearranged = s[4:] + s[:4]
    try:
        num = int("".join(str(int(ch, 36)) for ch in rearranged))
    except ValueError:
        return False
    return num % 97 == 1


def find_iban(text: str) -> list[Match]:
    out = []
    for m in _IBAN_RE.finditer(text):
        val = m.group(1).rstrip(" \u00a0.-")
        # potrivirea poate înghiți un fragment în plus (ex. „RO..9012 E” din „Email”):
        # încearcă prefixele care se termină la un separator, de la cel mai lung
        ends = [len(val)] + sorted((k for k, ch in enumerate(val) if ch in " \u00a0.-"), reverse=True)
        found = False
        for e in ends:
            c = val[:e]
            if iban_ok(c):
                out.append(Match("IBAN", c, m.start(1), m.start(1) + len(c), Confidence.MARE))
                found = True
                break
        if not found:
            # IBAN RO cu structură corectă dar cifră de control greșită (typo), lângă „IBAN/cont”
            before = text[max(0, m.start() - 30):m.start()]
            for e in ends:
                c = val[:e]
                if _compact(c).startswith("RO") and len(_compact(c)) == 24 and _IBAN_ANCHOR_RE.search(before):
                    out.append(Match("IBAN", c, m.start(1), m.start(1) + len(c), Confidence.MICA,
                                     "cifra de control nu se validează"))
                    break
    return out


# ---------------------------------------------------------------- act de identitate
_CI_RE = re.compile(
    r"(?i)\bseria\s*:?\s*([A-Z]{2})\s*[,;]?\s*(?:nr\.?|num[aă]r(?:ul)?)\s*:?\s*(\d{6,7})\b"
)
_PASS_RE = re.compile(
    r"(?i)\bpa[sșş]aport(?:ul|ului)?\b[^\d\n]{0,20}?(?:nr\.?|num[aă]r(?:ul)?)\s*:?\s*([A-Z]{0,2}\d{7,9})\b"
)


def find_id_doc(text: str) -> list[Match]:
    out = []
    for m in _CI_RE.finditer(text):
        out.append(Match("Serie act identitate", m.group(1), m.start(1), m.end(1), Confidence.MEDIE))
        out.append(Match("Număr act identitate", m.group(2), m.start(2), m.end(2), Confidence.MEDIE))
    for m in _PASS_RE.finditer(text):
        out.append(Match("Număr pașaport", m.group(1), m.start(1), m.end(1), Confidence.MEDIE))
    return out


# ---------------------------------------------------------------- contact
_EMAIL_RE = re.compile(r"(?i)(?<![\w.+-])[\w.+-]+@[\w-]+(?:\.[\w-]+)*\.[a-z]{2,24}(?![\w-])")
_PHONE_RE = re.compile(
    r"(?<![\d+])((?:\+40|0040|0)[\s.-]?(?:7\d{2}|[23]\d{1,2})(?:[\s.-]?\d){6,7})(?!\d)"
)


def find_contact(text: str) -> list[Match]:
    out = [Match("Email", m.group(0), m.start(), m.end(), Confidence.MARE)
           for m in _EMAIL_RE.finditer(text)]
    for m in _PHONE_RE.finditer(text):
        digits = re.sub(r"\D", "", m.group(1))
        national = digits[2:] if digits.startswith("40") else digits[3:] if digits.startswith("0040") else digits[1:]
        if len(national) == 9:
            out.append(Match("Telefon", m.group(1), m.start(1), m.end(1), Confidence.MEDIE))
    return out


# ---------------------------------------------------------------- data nașterii
_DOB_RE = re.compile(
    r"(?i)\bn[aă]scut(?:[aă])?\s+(?:la\s+)?(?:data\s+(?:de\s+)?)?"
    r"(\d{1,2}[./-]\d{1,2}[./-]\d{4})"
)


def find_dob(text: str) -> list[Match]:
    return [Match("Data nașterii", m.group(1), m.start(1), m.end(1), Confidence.MEDIE)
            for m in _DOB_RE.finditer(text)]


# ---------------------------------------------------------------- nume
_NAME_WORD = rf"(?:[{UP}][{LO}]+|[{UP}]{{2,}})(?:-(?:[{UP}][{LO}]+|[{UP}]{{2,}}))?"
_NAME = rf"{_NAME_WORD}(?:\s+{_NAME_WORD}){{1,3}}"
# nume după o ancoră
_NAME_AFTER_RE = re.compile(
    rf"(?:\b(?:[Ss]ubsemnat(?:ul|a)|[Dd]omnul|[Dd]oamna|[Dd]l\.|[Dd]na\.|[Dd]-na|[Dd]-l"
    rf"|reprezentat[aă]?\s+(?:legal\s+)?prin(?:\s+(?:dl\.|dna\.|domnul|doamna))?"
    rf"|[Nn]ume(?:le)?\s*(?:și|şi|si)\s*prenume(?:le)?\s*:?)\s+)({_NAME})"
)
# nume urmat de o ancoră de identificare
_NAME_BEFORE_RE = re.compile(
    rf"({_NAME})\s*,\s*(?=(?:domiciliat|cu\s+domiciliul|n[aă]scut|CNP|C\.N\.P|identificat"
    rf"|posesor|posesoare|legitimat|av[aâ]nd\s+CNP|cet[aă][tțţ]ean))"
)
# cuvinte care nu sunt nume deși sunt scrise cu majusculă
_NOT_NAME = {"SC", "SRL", "SA", "PFA", "II", "IF", "SNC", "ONG", "CNP", "CUI", "ROMÂNIA", "ROMANIA",
             "CONTRACT", "PRESTATOR", "BENEFICIAR", "VÂNZĂTOR", "CUMPĂRĂTOR", "LOCATOR", "LOCATAR",
             "ÎMPRUMUTĂTOR", "ÎMPRUMUTAT", "MANDANT", "MANDATAR", "ANGAJATOR", "SALARIAT", "PĂRȚILE",
             "ÎNTRE", "INTRE", "ARTICOLUL", "DOMNUL", "DOAMNA", "DL", "DNA", "D-NA", "D-L",
             "SUBSEMNATUL", "SUBSEMNATA", "NUME", "PRENUME", "REPREZENTANT", "ADMINISTRATOR", "ART", "Str", "Strada", "Municipiul", "Județul", "Sectorul"}


def _clean_name(raw: str) -> str | None:
    words = raw.split()
    while words and words[-1].strip(",.").upper() in {w.upper() for w in _NOT_NAME}:
        words.pop()
    while words and words[0].strip(",.").upper() in {w.upper() for w in _NOT_NAME}:
        words.pop(0)
    return " ".join(words) if len(words) >= 2 else None


def find_names(text: str) -> list[Match]:
    out = []
    for rx in (_NAME_AFTER_RE, _NAME_BEFORE_RE):
        for m in rx.finditer(text):
            name = _clean_name(m.group(1))
            if name:
                s = text.find(name, m.start(1))
                out.append(Match("Nume", name, s, s + len(name), Confidence.MEDIE))
    return out


# ---------------------------------------------------------------- adresă
_ADDR_ANCHOR_RE = re.compile(
    r"(?i)\b(?:domiciliat[aă]?|cu\s+domiciliul|cu\s+re[sșş]edin[tțţ]a|re[sșş]edent[aă]?)\s+"
    r"(?:[iî]n\s+)?(?:\(?adres[aă]\)?\s*:?\s*)?"
)
# ce urmează după adresă și o încheie
_ADDR_END_RE = re.compile(
    r"\s*,?\s*(?:CNP|C\.N\.P|identificat|posesor|posesoare|legitimat|av[aâ]nd|n[aă]scut"
    r"|[iî]n\s+calitate|tel(?:efon|\.)|e-?mail|cet[aă][tțţ]ean|denumit|domiciliat|cu\s+domiciliul"
    r"|[iî]nregistrat|[iî]nscris|compus|av[aâ]nd|care\s+are|cu\s+destina[tțţ]ia|[iî]n\s+suprafa[tțţ][aă])"
    r"|\s*[;\n]|\s*\)"
)
# abrevieri uzuale în adrese: punctul de după ele NU încheie propoziția
_ADDR_ABBR = {"str", "mun", "jud", "sect", "sector", "nr", "bl", "sc", "et", "ap", "bd", "bdul", "b-dul",
              "cal", "sos", "șos", "şos", "com", "sat", "or", "loc", "aleea", "al", "int", "cart", "p-ța",
              "pța", "piața", "splai", "sp", "cod", "c.p", "jud", "mun", "ors", "nr.cad"}
_ADDR_MAX = 180


def _cut_sentence(val: str) -> str:
    for m in re.finditer(r"\.(?=\s+[A-ZĂÂÎȘȚ])|\.$", val):
        prev = re.search(r"([\w.-]+)$", val[:m.start()])
        if prev and prev.group(1).lower().rstrip(".") in _ADDR_ABBR:
            continue
        return val[:m.start()]
    return val


def _address_after(text: str, anchor_re: re.Pattern, kind: str, min_len: int = 8) -> list[Match]:
    out = []
    for m in anchor_re.finditer(text):
        s = m.end()
        window = text[s:s + _ADDR_MAX]
        e = _ADDR_END_RE.search(window)
        val = window[:e.start()] if e else window
        cut = _cut_sentence(val)
        certain = bool(e) or len(cut) < len(val)
        val = cut.rstrip(" ,.")
        if len(val) >= min_len:
            out.append(Match(kind, val, s, s + len(val), Confidence.MEDIE,
                             "" if certain else "sfârșitul e incert - verificați"))
    return out


def find_address(text: str) -> list[Match]:
    return _address_after(text, _ADDR_ANCHOR_RE, "Adresă")


# ---------------------------------------------------------------- imobil (adresă, CF, cadastral)
_PROP_ANCHOR_RE = re.compile(
    r"(?i)\b(?:situat[aă]?|amplasat[aă]?|aflat[aă]?)\s+(?:[iî]n|la)\s+(?:adresa\s*:?\s*)?"
    r"|\b(?:imobil(?:ul|ului)?|apartament(?:ul|ului)?|locuin[tțţ][aă]|spa[tțţ]iul)\s+din\s+"
)
_CF_RE = re.compile(
    r"(?i)\b(?:c[aă]r[tț]i(?:i|a)?\s+funciar[eă]|cartea\s+funciar[aă]|carte\s+funciar[aă]|C\.\s?F\.|CF\b)"
    r"[^\n.;]{0,40}?\b(?:nr\.?|num[aă]r(?:ul)?)\s*:?\s*(\d{2,9}(?:[-/][\w]+)*)"
)
_CAD_RE = re.compile(
    r"(?i)\b(?:num[aă]r(?:ul)?\s+cadastral|nr\.?\s*cad(?:astral)?\.?|cadastral\s+nr\.?)\s*:?\s*"
    r"(\d{2,9}(?:[-/][\w]+)*)"
)


def find_property(text: str) -> list[Match]:
    out = _address_after(text, _PROP_ANCHOR_RE, "Adresă imobil", min_len=10)
    for rx in (_CF_RE, _CAD_RE):
        for m in rx.finditer(text):
            out.append(Match("Carte funciară / cadastral", m.group(1), m.start(1), m.end(1), Confidence.MEDIE))
    return out


# ---------------------------------------------------------------- permis de conducere
_DL_RE = re.compile(
    r"(?i)\bpermis(?:ul|ului)?\s+de\s+conducere\b[^\n;]{0,40}?\b(?:nr\.?|num[aă]r(?:ul)?|seria)\s*:?\s*"
    r"(?=[A-Z0-9]*\d)([A-Z0-9]{5,12})\b"
)


def find_driving_licence(text: str) -> list[Match]:
    return [Match("Număr permis conducere", m.group(1), m.start(1), m.end(1), Confidence.MEDIE)
            for m in _DL_RE.finditer(text)]


# ---------------------------------------------------------------- loc naștere, emitent și date ale actelor
_POB_RE = re.compile(r"(?i)\bn[aă]scut[aă]?\b[^;\n]{0,45}?\s[iî]n\s+(?!data\b)")
_ISSUER_RE = re.compile(
    r"(?i)\b(?:eliberat[aă]?|emis[aă]?)\s+de\s+(?:c[aă]tre\s+)?"
    r"([^,;\n]{2,60}?)(?=\s+la\s+data\b|\s+la\s+\d|\s*[,;\n]|\.\s|\.$|$)"
)
_DATE = r"(\d{1,2}[./-]\d{1,2}[./-]\d{4})"
_ACT_DATE_RE = re.compile(
    rf"(?i)\b(?:(?:eliberat[aă]?|emis[aă]?)\b[^;\n]{{0,70}}?\bla\s+(?:data\s+(?:de\s+)?)?{_DATE}"
    rf"|valabil[aă]?\s+p[aâ]n[aă]\s+(?:la\s+)?(?:data\s+(?:de\s+)?)?{_DATE})"
)


def find_indirect(text: str) -> list[Match]:
    out = _address_after(text, _POB_RE, "Loc naștere", min_len=3)
    for m in out:
        m.note = (m.note + "; " if m.note else "") + "dată indirectă"
    for m in _ISSUER_RE.finditer(text):
        out.append(Match("Emitent act", m.group(1), m.start(1), m.end(1), Confidence.MEDIE,
                         "dată indirectă"))
    for m in _ACT_DATE_RE.finditer(text):
        g = 1 if m.group(1) else 2
        out.append(Match("Dată act", m.group(g), m.start(g), m.end(g), Confidence.MEDIE,
                         "dată indirectă"))
    return out


# ---------------------------------------------------------------- agregare
ALL_DETECTORS = (find_cnp, find_iban, find_id_doc, find_contact, find_dob, find_names, find_address,
                 find_property, find_driving_licence, find_indirect)


def detect(text: str) -> list[Match]:
    """Rulează toate detectoarele și elimină suprapunerile (păstrează potrivirea
    cu încrederea cea mai mare, apoi pe cea mai lungă)."""
    matches: list[Match] = []
    for det in ALL_DETECTORS:
        matches.extend(det(text))
    rank = {Confidence.MARE: 0, Confidence.MEDIE: 1, Confidence.MICA: 2}
    matches.sort(key=lambda m: (rank[m.confidence], -(m.end - m.start)))
    kept: list[Match] = []
    for m in matches:
        if any(not (m.end <= k.start or m.start >= k.end) for k in kept):
            # suprapunere: o adresă care conține un telefon etc. -> păstrăm prima
            continue
        kept.append(m)
    return sorted(kept, key=lambda m: m.start)


def find_literal(text: str, value: str, kind: str, confidence: Confidence, note: str) -> list[Match]:
    """Caută toate aparițiile unei valori deja cunoscute (propagare / adăugare manuală).
    Insensibil la majuscule și la spațiile multiple."""
    if not value.strip():
        return []
    pattern = r"\s+".join(re.escape(w) for w in value.split())
    rx = re.compile(rf"(?<![\w]){pattern}(?![\w])", re.IGNORECASE)
    return [Match(kind, m.group(0), m.start(), m.end(), confidence, note) for m in rx.finditer(text)]
