from pathlib import Path
from collections import Counter
from modules.extraction_factures.factextv19 import process_pdf

root = Path(r"C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\3 - Travaux\PBC\Factures")
pdfs = sorted(root.rglob("*.pdf"))[:30]

all_inv = []
for p in pdfs:
    try:
        invs = process_pdf(p, extract_juridical=True)
        all_inv.extend(invs)
    except Exception as e:
        print("ERR", p.name, str(e)[:120])

n = len(all_inv)
print("FILES_TESTED", len(pdfs))
print("INVOICES_EXTRACTED", n)
if n:
    print("NUM_RATE", sum(1 for i in all_inv if i.num_facture.value), "/", n)
    print("DATE_RATE", sum(1 for i in all_inv if i.date_facture.value), "/", n)
    print("HT_RATE", sum(1 for i in all_inv if i.montant_ht.value), "/", n)
    print("TVA_RATE", sum(1 for i in all_inv if i.montant_tva.value), "/", n)
    print("TTC_RATE", sum(1 for i in all_inv if i.montant_ttc.value), "/", n)
    print("FOUR_RATE", sum(1 for i in all_inv if i.fournisseur.value), "/", n)
    print("ICE_RATE", sum(1 for i in all_inv if i.ice.value), "/", n)
    c = Counter(i.tva_control.statut for i in all_inv)
    print("TVA_STATUS", dict(c))
    a = Counter()
    for i in all_inv:
        for x in i.anomalies:
            a[x] += 1
    print("TOP_ANOMALIES")
    for k,v in a.most_common(12):
        print(v, "|", k)
