#!/usr/bin/env python3
"""Test: Spalte verstecken mit xlwings auf macOS - minimal"""

import xlwings as xw
import os
import subprocess
import platform

# Kill Excel first
if platform.system() == 'Darwin':
    subprocess.run(['pkill', '-9', 'Microsoft Excel'], capture_output=True)

import time
time.sleep(1)

# Test mit einer existierenden Datei
test_file = "/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx"
output_file = "/Users/nojan/Desktop/TEST-hidden-column.xlsx"

# Kopiere die Datei
import shutil
if os.path.exists(output_file):
    os.remove(output_file)
shutil.copy2(test_file, output_file)

print(f"1. Datei kopiert: {output_file}")

print("2. Öffne mit xlwings und verstecke Spalte B...")
with xw.App(visible=False, add_book=False) as app:
    app.display_alerts = False
    app.screen_updating = False
    
    wb = app.books.open(output_file)
    ws = wb.sheets[0]
    
    print(f"   Sheet: {ws.name}")
    
    # Spalte B verstecken
    col_b = ws.range('B:B')
    print(f"   Range: {col_b}")
    print(f"   api type: {type(col_b.api)}")
    
    # Teste verschiedene Methoden
    try:
        # Methode 1: api.column_hidden
        col_b.api.column_hidden = True
        print("   ✓ api.column_hidden = True gesetzt")
    except Exception as e:
        print(f"   ✗ api.column_hidden FEHLER: {e}")
    
    # Auch Zeile 3 verstecken testen
    try:
        row_3 = ws.range('3:3')
        row_3.api.row_hidden = True
        print("   ✓ api.row_hidden = True gesetzt für Zeile 3")
    except Exception as e:
        print(f"   ✗ api.row_hidden FEHLER: {e}")
    
    # Speichern
    wb.save()
    print("   ✓ Gespeichert")
    wb.close()
    print("   ✓ Geschlossen")

print(f"\n3. FERTIG! Öffne die Datei:")
print(f"   open '{output_file}'")
print("   Prüfe ob Spalte B und Zeile 3 versteckt sind!")
