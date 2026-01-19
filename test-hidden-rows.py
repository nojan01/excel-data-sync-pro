#!/usr/bin/env python3
"""Test: Mehrere Zeilen verstecken mit row_height = 0"""

import xlwings as xw
import os
import subprocess
import shutil

# Kill Excel
subprocess.run(['pkill', '-9', 'Microsoft Excel'], capture_output=True)
import time
time.sleep(1)

test_file = "/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx"
output_file = "/Users/nojan/Desktop/TEST-hidden-rows.xlsx"

if os.path.exists(output_file):
    os.remove(output_file)
shutil.copy2(test_file, output_file)

print(f"1. Datei kopiert")

print("2. Verstecke Zeilen 2, 3, 4 (row_height = 0)...")
with xw.App(visible=False, add_book=False) as app:
    app.display_alerts = False
    
    wb = app.books.open(output_file)
    ws = wb.sheets[0]
    
    for excel_row in [2, 3, 4]:
        row_range = ws.range(f'{excel_row}:{excel_row}')
        print(f"   Zeile {excel_row}: Vorher height = {row_range.row_height}")
        row_range.row_height = 0
        print(f"   Zeile {excel_row}: Nachher height = {row_range.row_height}")
    
    wb.save()
    wb.close()

print(f"\n3. FERTIG! open '{output_file}'")
print("   Prüfe ob Zeilen 2, 3, 4 versteckt sind!")
