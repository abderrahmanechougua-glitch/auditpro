from pathlib import Path
from collections import Counter

from modules.extraction_factures.factextv19 import process_pdf

root = Path(r"C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\3 - Travaux\PBC\Factures")
pdfs = sorted(root.rglob("*.pdf"))[30:40]

all_inv = []
for pdf_path in pdfs:
    try:
        invoices = process_pdf(pdf_path, extract_juridical=True)
        all_inv.extend(invoices)
    except Exception as exc:
        print("ERR", pdf_path.name, str(exc)[:120])

count = len(all_inv)
print("FILES_TESTED", len(pdfs))
print("INVOICES_EXTRACTED", count)
if count:
    print("NUM_RATE", sum(1 for i in all_inv if i.num_facture.value), "/", count)
    print("DATE_RATE", sum(1 for i in all_inv if i.date_facture.value), "/", count)
    print("HT_RATE", sum(1 for i in all_inv if i.montant_ht.value), "/", count)
    print("TVA_RATE", sum(1 for i in all_inv if i.montant_tva.value), "/", count)
    print("TTC_RATE", sum(1 for i in all_inv if i.montant_ttc.value), "/", count)
    print("ICE_RATE", sum(1 for i in all_inv if i.ice.value), "/", count)
    statuses = Counter(i.tva_control.statut for i in all_inv)
    print("TVA_STATUS", dict(statuses))
