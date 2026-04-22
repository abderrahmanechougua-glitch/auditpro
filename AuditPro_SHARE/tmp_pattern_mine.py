from pathlib import Path
import re
import collections
import pytesseract
from pdf2image import convert_from_path
from modules.extraction_factures.factextv19 import TESSERACT_PATH, POPPLER_PATH

root = Path(r"C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\3 - Travaux\PBC\Factures")
pdfs = sorted(root.rglob("*.pdf"))[:30]

if TESSERACT_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

pats = {
    "num": re.compile(r"(?i)\b(?:facture|invoice|n[°º]|numero|num[eé]ro|bc|bon\s+de\s+commande|bl\.?\s*no?)\b"),
    "date": re.compile(r"(?i)\b(?:date|du|le)\b"),
    "ht": re.compile(r"(?i)\b(?:ht|h\.t|hors\s+taxe|subtotal|sous[- ]?total|total\s*\(\s*ht)\b"),
    "tva": re.compile(r"(?i)\b(?:tva|t\.v\.a|vat|taxe)\b"),
    "ttc": re.compile(r"(?i)\b(?:ttc|t\.t\.c|ftc|net\s+a\s+payer|total\s+general|total\s+g[ée]n[ée]ral|montant\s+ttc)\b"),
}

hits = {k: collections.Counter() for k in pats}
done = 0
fail = 0
for p in pdfs:
    try:
        imgs = convert_from_path(str(p), dpi=170, first_page=1, last_page=1, poppler_path=POPPLER_PATH or None)
        if not imgs:
            continue
        text = pytesseract.image_to_string(imgs[0], lang="fra+eng", config="--psm 6 --oem 3")
        done += 1
    except Exception:
        fail += 1
        continue

    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines:
        for k, rx in pats.items():
            if rx.search(ln):
                compact = re.sub(r"\d", "9", ln)[:140]
                hits[k][compact] += 1

print("SAMPLED", len(pdfs), "OCR_OK", done, "OCR_FAIL", fail)
for k in ["num", "date", "ht", "tva", "ttc"]:
    print("---", k, "TOP---")
    for line, c in hits[k].most_common(20):
        print(f"{c:3d} | {line}")
