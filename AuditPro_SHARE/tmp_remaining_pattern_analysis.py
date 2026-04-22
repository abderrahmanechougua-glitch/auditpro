from pathlib import Path
from collections import Counter
import re

import pytesseract
from pdf2image import convert_from_path

from modules.extraction_factures.factextv19 import POPPLER_PATH, TESSERACT_PATH

ROOT = Path(r"C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\3 - Travaux\PBC\Factures")
PDFS = sorted(ROOT.rglob("*.pdf"))[30:]

PATTERNS = {
    "num": re.compile(r"(?i)\b(?:facture|invoice|n[°º]|numero|num[eé]ro|bc|bon\s+de\s+commande|bl\.?\s*no?)\b"),
    "date": re.compile(r"(?i)\b(?:date|du|le)\b"),
    "ht": re.compile(r"(?i)\b(?:ht|h\.t|hors\s+taxe|subtotal|sous[- ]?total|total\s*\(\s*ht|total\s+net\s+h\.t\.?|montant\s+h\.t\.?)\b"),
    "tva": re.compile(r"(?i)\b(?:tva|t\.v\.a|vat|taxe)\b"),
    "ttc": re.compile(r"(?i)\b(?:ttc|t\.t\.c|ftc|net\s+a\s+payer|net\s+apayer|total\s+general|total\s+g[ée]n[ée]ral|montant\s+ttc)\b"),
}


def compact_line(line: str) -> str:
    line = re.sub(r"\s+", " ", line).strip()
    line = re.sub(r"\d", "9", line)
    return line[:160]


def main() -> None:
    if TESSERACT_PATH:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

    hits = {name: Counter() for name in PATTERNS}
    processed = 0
    failed = 0

    for pdf_path in PDFS:
        try:
            images = convert_from_path(
                str(pdf_path),
                dpi=170,
                first_page=1,
                last_page=1,
                poppler_path=POPPLER_PATH or None,
            )
            if not images:
                continue
            text = pytesseract.image_to_string(images[0], lang="fra+eng", config="--psm 6 --oem 3")
            processed += 1
        except Exception:
            failed += 1
            continue

        for line in text.splitlines():
            cleaned = line.strip()
            if not cleaned:
                continue
            for name, regex in PATTERNS.items():
                if regex.search(cleaned):
                    hits[name][compact_line(cleaned)] += 1

    print("REMAINING_PDFS", len(PDFS))
    print("OCR_OK", processed)
    print("OCR_FAIL", failed)
    for name in ["num", "date", "ht", "tva", "ttc"]:
        print(f"--- {name.upper()} TOP ---")
        for line, count in hits[name].most_common(30):
            print(f"{count:3d} | {line}")


if __name__ == "__main__":
    main()
