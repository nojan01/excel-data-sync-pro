#!/usr/bin/env python3
"""Test: Spalte verstecken mit xlwings auf macOS"""

import xlwings as xw
import os
import shutil

# Test-Datei erstellen
test_file = "/tmp/test-hidden-column.xlsx"

print("1. Erstelle Test-Workbook...")
with xw.App(visible=False, add_book=False) as app:
    app.display_alerts = False
    wb = app.books.add()
    ws = wb.sheets[0]
    
    # Daten einfügen
    ws.range('A1').value = 'Spalte A'
    ws.range('B1').value = 'Spalte B (soll versteckt werden)'
    ws.range('C1').value = 'Spalte C'
    ws.range('A2').value = 'Wert A'
    ws.range('B2').value = 'Wert B'
    ws.range('C2').value = 'Wert C'
    
    wb.save(test_file)
    wb.close()

print(f"2. Datei erstellt: {test_file}")

print("3. Öffne Datei und verstecke Spalte B...")
with xw.App(visible=False, add_book=False) as app:
    app.display_alerts = False
    wb = app.books.open(test_file)
    ws = wb.sheets[0]
    
    # Spalte B verstecken - verschiedene Methoden testen
    col_b = ws.range('B:B')
    
    print(f"   Range: {col_b}")
    print(f"   api type: {type(col_b.api)}")
    
    # Methode 1: api.column_hidden
    try:
        col_b.api.column_hidden = True
        print("   ✓ api.column_hidden = True (kein Fehler)")
    except Exception as e:
        print(f"   ✗ api.column_hidden FEHLER: {e}")
    
    # Prüfe ob es gesetzt wurde
    try:
        is_hidden = col_b.api.column_hidden
        print(f"   Spalte B versteckt? {is_hidden}")
    except Exception as e:
        print(f"   Konnte Status nicht lesen: {e}")
    
    wb.save()
    wb.close()

print("4. Öffne Datei erneut und prüfe...")
with xw.App(visible=False, add_book=False) as app:
    app.display_alerts = False
    wb = app.books.open(test_file)
    ws = wb.sheets[0]
    
    col_b = ws.range('B:B')
    try:
        is_hidden = col_b.api.column_hidden
        print(f"   Nach Wiedereröffnen - Spalte B versteckt? {is_hidden}")
    except Exception as e:
        print(f"   Konnte Status nicht lesen: {e}")
    
    # Zeige alle Spaltenbreiten
    for col in ['A', 'B', 'C']:
        try:
            width = ws.range(f'{col}:{col}').column_width
            hidden = ws.range(f'{col}:{col}').api.column_hidden
            print(f"   Spalte {col}: width={width}, hidden={hidden}")
        except Exception as e:
            print(f"   Spalte {col}: Fehler - {e}")
    
    wb.close()

print(f"\n5. FERTIG! Öffne die Datei manuell:")
print(f"   open '{test_file}'")
print("   Prüfe ob Spalte B versteckt ist!")
