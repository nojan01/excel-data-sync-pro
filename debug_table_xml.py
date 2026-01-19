#!/usr/bin/env python3
"""Vergleiche Table-XML zwischen Original und Export"""
import zipfile
import os
import re

ORIG_FILE = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx'

# Finde den neuesten Export
desktop = '/Users/nojan/Desktop'
exports = [f for f in os.listdir(desktop) if f.startswith('Export_') and 'GER' in f and f.endswith('.xlsx')]
if not exports:
    print("Keine Export-Datei gefunden!")
    exit(1)
    
exports.sort(key=lambda x: os.path.getmtime(os.path.join(desktop, x)), reverse=True)
EXPORT_FILE = os.path.join(desktop, exports[0])

print(f"Original: {ORIG_FILE}")
print(f"Export: {EXPORT_FILE}")

def get_table_xml(xlsx_path):
    """Extrahiere table1.xml aus XLSX"""
    with zipfile.ZipFile(xlsx_path, 'r') as zf:
        # Suche table1.xml
        for name in zf.namelist():
            if 'tables/table1.xml' in name:
                return zf.read(name).decode('utf-8')
    return None

orig_table = get_table_xml(ORIG_FILE)
export_table = get_table_xml(EXPORT_FILE)

if not orig_table:
    print("Keine Table in Original gefunden!")
    exit(1)
if not export_table:
    print("Keine Table in Export gefunden!")
    exit(1)

print(f"\n=== Original Table XML (erste 2000 Zeichen) ===")
print(orig_table[:2000])

print(f"\n=== Export Table XML (erste 2000 Zeichen) ===")
print(export_table[:2000])

# Prüfe auf tableStyleInfo
print(f"\n=== TableStyleInfo Vergleich ===")
orig_style = re.search(r'<tableStyleInfo[^/]*/>', orig_table)
export_style = re.search(r'<tableStyleInfo[^/]*/>', export_table)

print(f"Original: {orig_style.group(0) if orig_style else 'NICHT GEFUNDEN!'}")
print(f"Export:   {export_style.group(0) if export_style else 'NICHT GEFUNDEN!'}")
