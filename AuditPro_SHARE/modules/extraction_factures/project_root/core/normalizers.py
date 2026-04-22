from __future__ import annotations

import re
from datetime import datetime
from typing import Optional

# OCR confusion map for numeric contexts.
_OCR_DIGIT_MAP = str.maketrans({
    "O": "0",
    "o": "0",
    "I": "1",
    "l": "1",
    "|": "1",
    "S": "5",
    "s": "5",
    "B": "8",
    "b": "8",
    "Z": "2",
    "z": "2",
})


def normalize_ocr_numeric_text(value: str) -> str:
    """Normalize common OCR confusions for numeric strings.

    Args:
        value: Raw OCR text.

    Returns:
        Sanitized text with common OCR substitutions applied.
    """
    return (value or "").translate(_OCR_DIGIT_MAP)


def normalize_amount(amount_str: str) -> Optional[float]:
    """Parse amount strings across Moroccan and international formats.

    Supported examples:
    - 175 000,00
    - 1 749 118.68
    - 25,000.00
    - 1.749.118,68

    Args:
        amount_str: Amount as text.

    Returns:
        Parsed float or None when parsing fails.
    """
    if not amount_str:
        return None

    s = normalize_ocr_numeric_text(amount_str.strip())
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[A-Za-zÀ-ÿ()$€]", "", s)
    s = re.sub(r"\s+", " ", s).strip()

    if not s:
        return None

    points = s.count(".")
    commas = s.count(",")
    spaces = len(re.findall(r" ", s))

    if spaces >= 1 and commas == 1 and points == 0:
        s = s.replace(" ", "").replace(",", ".")
    elif spaces >= 1 and points == 1 and commas == 0:
        s = s.replace(" ", "")
    elif commas == 1 and points >= 1 and s.rfind(",") > s.rfind("."):
        s = s.replace(" ", "").replace(".", "").replace(",", ".")
    elif commas >= 1 and points == 1 and s.rfind(".") > s.rfind(","):
        s = s.replace(" ", "").replace(",", "")
    elif points >= 2 and commas == 0:
        s = s.replace(" ", "").replace(".", "")
    elif commas == 1 and points == 0:
        s = s.replace(" ", "").replace(",", ".")
    else:
        s = s.replace(" ", "")

    s = re.sub(r"[^0-9.\-]", "", s).strip(".")
    if not s:
        return None

    try:
        parsed = float(s)
        return parsed if parsed >= 0 else None
    except (TypeError, ValueError):
        return None


def normalize_date(value: str, max_year_future: int = 5) -> str:
    """Normalize date into DD/MM/YYYY with century validation.

    Args:
        value: Raw date text.
        max_year_future: Allowed future years compared to current year.

    Returns:
        Normalized date or original value when no safe normalization can be done.
    """
    if not value:
        return ""

    raw = value.strip()

    m = re.match(r"(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})", raw)
    if m:
        day, month, year = m.groups()
        year_i = int(year)
        if year_i < 100:
            year_i += 2000
        return _safe_date(int(day), int(month), year_i, raw, max_year_future)

    m = re.match(r"(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})", raw)
    if m:
        year, month, day = m.groups()
        return _safe_date(int(day), int(month), int(year), raw, max_year_future)

    return raw


def _safe_date(day: int, month: int, year: int, fallback: str, max_year_future: int) -> str:
    now = datetime.now().year
    if year < 1990 or year > now + max_year_future:
        return fallback
    try:
        dt = datetime(year, month, day)
        return dt.strftime("%d/%m/%Y")
    except ValueError:
        return fallback


def parse_french_words_to_number(text: str) -> Optional[float]:
    """Convert French amount-in-words into numeric value.

    Args:
        text: French words amount text.

    Returns:
        Parsed numeric value or None.
    """
    if not text:
        return None

    t = text.lower().strip()
    parts = re.split(r"\b(?:dirham|dirhams|dh)\b", t, maxsplit=1)
    main = parts[0].strip()
    cents = parts[1].strip() if len(parts) > 1 else ""
    cents = re.sub(r"\b(?:centime|centimes|cts?)\b", "", cents).strip()

    main_i = _french_words_to_int(main)
    cents_i = _french_words_to_int(cents)
    total = float(main_i) + (float(cents_i) / 100.0)
    return total if total > 0 else None


def _french_words_to_int(text: str) -> int:
    if not text.strip():
        return 0

    units = {
        "zero": 0,
        "zéro": 0,
        "un": 1,
        "une": 1,
        "deux": 2,
        "trois": 3,
        "quatre": 4,
        "cinq": 5,
        "six": 6,
        "sept": 7,
        "huit": 8,
        "neuf": 9,
        "dix": 10,
        "onze": 11,
        "douze": 12,
        "treize": 13,
        "quatorze": 14,
        "quinze": 15,
        "seize": 16,
        "dix-sept": 17,
        "dix-huit": 18,
        "dix-neuf": 19,
        "vingt": 20,
        "trente": 30,
        "quarante": 40,
        "cinquante": 50,
        "soixante": 60,
    }

    normalized = re.sub(r"[\-–]", " ", text.replace("et ", " "))
    tokens = normalized.split()

    total = 0
    current = 0
    i = 0

    while i < len(tokens):
        w = tokens[i]

        if w == "quatre" and i + 1 < len(tokens) and tokens[i + 1] in {"vingt", "vingts"}:
            base = 80
            i += 2
            if i < len(tokens) and tokens[i] == "dix":
                base = 90
                i += 1
                if i < len(tokens) and tokens[i] in units and units[tokens[i]] < 10:
                    base += units[tokens[i]]
                    i += 1
            elif i < len(tokens) and tokens[i] in units and units[tokens[i]] < 20:
                base += units[tokens[i]]
                i += 1
            current += base
            continue

        if w == "soixante":
            base = 60
            i += 1
            if i < len(tokens) and tokens[i] == "dix":
                base = 70
                i += 1
                if i < len(tokens) and tokens[i] in units and units[tokens[i]] < 10:
                    base += units[tokens[i]]
                    i += 1
            elif i < len(tokens) and tokens[i] in units and units[tokens[i]] < 20:
                base += units[tokens[i]]
                i += 1
            current += base
            continue

        if w in {"cent", "cents"}:
            current = current * 100 if current > 0 else 100
            i += 1
            continue

        if w == "mille":
            current = current * 1000 if current > 0 else 1000
            total += current
            current = 0
            i += 1
            continue

        if w in {"million", "millions"}:
            current = current * 1_000_000 if current > 0 else 1_000_000
            total += current
            current = 0
            i += 1
            continue

        if w in {"milliard", "milliards"}:
            current = current * 1_000_000_000 if current > 0 else 1_000_000_000
            total += current
            current = 0
            i += 1
            continue

        if w in units:
            current += units[w]
        elif w.isdigit():
            current += int(w)

        i += 1

    return total + current
