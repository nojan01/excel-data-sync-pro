#!/usr/bin/env python3
"""Vergleiche CF-Regeln in Original und Export"""
import zipfile
import re
import sys

def extract_cf_rules(xlsx_path, label):
    print(f"\n=== {label}: {xlsx_path} ===")
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zf:
            content = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
            
            cf_pattern = r'<conditionalFormatting[^>]*sqref="([^"]+)"'
            matches = re.findall(cf_pattern, content)
            
            if matches:
                print(f"Gefunden: {len(matches)} CF-Regeln")
                for i, sqref in enumerate(matches[:10]):  # Erste 10 zeigen
                    print(f"  {i+1}. sqref=\"{sqref}\"")
                if len(matches) > 10:
                    print(f"  ... und {len(matches) - 10} weitere")
            else:
                print("Keine CF-Regeln gefunden!")
    except Exception as e:
        print(f"Fehler: {e}")

if len(sys.argv) >= 2:
    extract_cf_rules(sys.argv[1], "Datei")
else:
    print("Usage: python3 compare_cf.py <xlsx_path>")
    print("Oder: python3 compare_cf.py <original> <export>")
