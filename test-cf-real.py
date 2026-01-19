#!/usr/bin/env python3
"""Test: CF-Anpassung mit echter Testdatei"""

import sys
sys.path.insert(0, '/Users/nojan/Documents/GitHub/mvms-tool-electron/python')

from excel_writer import adjust_conditional_formatting
from openpyxl import load_workbook
import tempfile
import shutil
import os

# Verwende die andere Datei ohne Pivot
src = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx'
tmp = tempfile.mktemp(suffix='.xlsx')
shutil.copy(src, tmp)

print(f"Lade {os.path.basename(src)}...")
wb = load_workbook(tmp)
ws = wb.active

print(f"Sheet: {ws.title}")
print(f"Anzahl CF-Regeln VORHER: {len(ws.conditional_formatting._cf_rules)}")

# Zeige erste 5 CF Bereiche
print("\nErste 5 CF-Bereiche VORHER:")
for i, (cf, rules) in enumerate(list(ws.conditional_formatting._cf_rules.items())[:5]):
    print(f"  {i+1}. {cf.sqref}")

# Lösche Spalte B (Index 1)
print("\n" + "="*50)
print("Lösche Spalte B (Index 1, 0-basiert)...")
print("="*50)

deleted_cols = [1]
adjust_conditional_formatting(ws, deleted_cols, None)
ws.delete_cols(2)  # 1-basiert

print(f"\nAnzahl CF-Regeln NACHHER: {len(ws.conditional_formatting._cf_rules)}")

# Zeige erste 5 CF Bereiche nachher
print("\nErste 5 CF-Bereiche NACHHER:")
for i, (cf, rules) in enumerate(list(ws.conditional_formatting._cf_rules.items())[:5]):
    print(f"  {i+1}. {cf.sqref}")

# Speichern und prüfen
output = '/Users/nojan/Desktop/CF-Test-Output.xlsx'
wb.save(output)
print(f"\nGespeichert als: {output}")
print("Öffne diese Datei in Excel und prüfe die bedingten Formatierungen!")

os.unlink(tmp)
