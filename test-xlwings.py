#!/usr/bin/env python3
"""
Test: xlwings für Spalten löschen MIT CF-Erhaltung
xlwings steuert echtes Excel - dadurch werden ALLE Formatierungen korrekt angepasst!
"""

import xlwings as xw
import shutil
import os

# Testdatei kopieren
src = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx'
test_file = '/Users/nojan/Desktop/xlwings-test.xlsx'

print(f"Kopiere {os.path.basename(src)}...")
shutil.copy(src, test_file)

print("Öffne Excel mit xlwings...")
# visible=False = headless mode
app = xw.App(visible=False)

try:
    wb = app.books.open(test_file)
    ws = wb.sheets[0]
    
    print(f"Sheet: {ws.name}")
    print(f"Bereich: {ws.used_range.address}")
    
    # Zeige Header der ersten 5 Spalten
    print("\nHeader VORHER:")
    for col in range(1, 6):
        val = ws.range((1, col)).value
        print(f"  Spalte {col}: {val}")
    
    # Lösche Spalte B (Index 2)
    print("\n--- Lösche Spalte B ---")
    ws.range('B:B').delete()
    
    # Zeige Header nach Löschung
    print("\nHeader NACHHER:")
    for col in range(1, 5):
        val = ws.range((1, col)).value
        print(f"  Spalte {col}: {val}")
    
    # Speichern
    wb.save()
    wb.close()
    
    print(f"\n✓ Gespeichert als: {test_file}")
    print("Öffne diese Datei in Excel und prüfe die bedingten Formatierungen!")
    
finally:
    app.quit()
    print("Excel geschlossen.")
