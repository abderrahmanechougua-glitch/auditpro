import sys
sys.path.insert(0, r'AuditPro_SHARE\dist\AuditPro\_internal')
from modules.reconciliation_bg_liasse.reconciliation import _detect_liasse_format, _extract_liasse

files = {
    'ISSAL_MADINA': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\1- Back up\Controle des comptes\Finalisation\ISSAL MADINA - Liasse Comptable 2024.xlsx',
    'GIM_XLS': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Yassine MOUKRIM - GIM - CI - 2025\Finalisation\PBC\Liasse_Fiscale_GLOBAL-INTERNATIONAL-MOTORS_2025 (2).xls',
    'PDF_GIM': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Yassine MOUKRIM - GIM - CI - 2025\Controle des comptes GIM\PBC\TR_ Acomptes IS 2025\Liasse fiscale SIMPL IS 2024.pdf',
    'PDF_PRESTALYS': r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Khawla ECHINE - Prestalys 2025\PBC\DIVERS\Edition_de_la_liasse PRESTALYS au 31-12-2025 VF DU 31-03-2026.pdf',
}

for name, path in files.items():
    try:
        fmt = _detect_liasse_format(path)
        df = _extract_liasse(path)
        print(f'{name}: format={fmt}, rubriques={len(df)}')
        for _, row in df.iterrows():
            label = str(row['Rubrique'])[:50]
            val = row['Montant_Liasse']
            print(f'  {label:50} : {val:>15,.2f}')
    except Exception as e:
        import traceback
        print(f'{name}: ERREUR - {e}')
        traceback.print_exc()
    print()
