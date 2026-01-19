#!/usr/bin/env python3
"""Test: Verschiedene Methoden zum Spalten-Verstecken auf macOS"""

import xlwings as xw
import os
import subprocess
import platform
import shutil

# Kill Excel first
if platform.system() == 'Darwin':
    subprocess.run(['pkill', '-9', 'Microsoft Excel'], capture_output=True)

import time
time.sleep(1)

# Test mit einer existierenden Datei
test_file = "/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx"
output_file = "/Users/nojan/Desktop/TEST-hidden-methods.xlsx"

# Kopiere die Datei
if os.path.exists(output_file):
    os.remove(output_file)
shutil.copy2(test_file, output_file)

print(f"1. Datei kopiert: {output_file}")

print("2. Öffne mit xlwings...")
with xw.App(visible=True, add_book=False) as app:  # visible=True für Debugging
    app.display_alerts = False
    
    wb = app.books.open(output_file)
    ws = wb.sheets[0]
    
    print(f"   Sheet: {ws.name}")
    
    # ===== METHODE 1: api.column_hidden (Standard) =====
    print("\n--- METHODE 1: api.column_hidden ---")
    try:
        col_b = ws.range('B:B')
        col_b.api.column_hidden = True
        print(f"   Gesetzt. Prüfe: column_hidden = {col_b.api.column_hidden}")
    except Exception as e:
        print(f"   FEHLER: {e}")
    
    # ===== METHODE 2: column_width = 0 =====
    print("\n--- METHODE 2: column_width = 0 ---")
    try:
        col_c = ws.range('C:C')
        print(f"   Vorher: column_width = {col_c.column_width}")
        col_c.column_width = 0
        print(f"   Nachher: column_width = {col_c.column_width}")
    except Exception as e:
        print(f"   FEHLER: {e}")
    
    # ===== METHODE 3: api.ColumnWidth = 0 =====
    print("\n--- METHODE 3: api.ColumnWidth = 0 ---")
    try:
        col_d = ws.range('D:D')
        col_d.api.column_width = 0
        print(f"   Gesetzt")
    except Exception as e:
        print(f"   FEHLER: {e}")
    
    # ===== METHODE 4: Über EntireColumn =====
    print("\n--- METHODE 4: api.entire_column.hidden ---")
    try:
        cell_e1 = ws.range('E1')
        cell_e1.api.entire_column.hidden = True
        print(f"   Gesetzt")
    except Exception as e:
        print(f"   FEHLER: {e}")
    
    # ===== METHODE 5: Zeige alle verfügbaren Properties =====
    print("\n--- Verfügbare API-Eigenschaften für Range B:B ---")
    col_b = ws.range('B:B')
    api = col_b.api
    
    # Liste interessante Properties
    props_to_check = ['column_hidden', 'hidden', 'column_width', 'width', 
                      'entire_column', 'columns', 'visible']
    for prop in props_to_check:
        try:
            val = getattr(api, prop, 'N/A')
            print(f"   {prop}: {val}")
        except Exception as e:
            print(f"   {prop}: ERROR - {e}")
    
    # ===== METHODE 6: Direkt über appscript =====
    print("\n--- METHODE 6: Direkt über appscript ---")
    try:
        # Hole die Excel-Applikation
        from appscript import app, k
        excel = app('Microsoft Excel')
        
        # Verstecke Spalte F
        col_letter = 'F'
        # Setze column hidden direkt
        excel_wb = excel.workbooks[wb.name]
        excel_ws = excel_wb.sheets[ws.name]
        excel_col = excel_ws.columns[col_letter]
        excel_col.hidden.set(True)
        print(f"   Spalte {col_letter} versteckt via appscript")
    except Exception as e:
        print(f"   FEHLER: {e}")
    
    # Speichern
    print("\n3. Speichern...")
    wb.save()
    wb.close()

print(f"\n4. FERTIG! Öffne die Datei:")
print(f"   open '{output_file}'")
print("\nPrüfe welche Spalten versteckt sind:")
print("   B - api.column_hidden")
print("   C - column_width = 0")
print("   D - api.column_width = 0")
print("   E - api.entire_column.hidden")
print("   F - appscript direkt")
