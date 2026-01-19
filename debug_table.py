#!/usr/bin/env python3
"""Test-Export mit Spalte löschen - mit echten Daten"""
import sys
sys.path.insert(0, 'python')

from excel_reader import read_sheet
from excel_writer import write_sheet

original_path = "/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx"
output_path = "/Users/nojan/Desktop/DEBUG_Export.xlsx"
sheet_name = 'DEFENCE&SPACE Aug-2025'

# 1. Zuerst die Datei lesen
print("Lese Original-Datei...")
read_result = read_sheet(original_path, sheet_name)
headers = read_result['headers']
data = read_result['data']

print(f"Gelesene Spalten: {len(headers)}")
print(f"Gelesene Zeilen: {len(data)}")
print(f"Erste 5 Header: {headers[:5]}")

# 2. Spalte 2 (Index 2 = "Country") löschen
# Daten anpassen: Spalte 2 aus headers und data entfernen
deleted_col = 2
new_headers = headers[:deleted_col] + headers[deleted_col+1:]
new_data = []
for row in data:
    new_row = row[:deleted_col] + row[deleted_col+1:]
    new_data.append(new_row)

print(f"\nNach Löschen von Spalte {deleted_col} ({headers[deleted_col]}):")
print(f"Neue Spalten: {len(new_headers)}")
print(f"Erste 5 Header: {new_headers[:5]}")

# 3. Export mit korrekten Änderungen
changes = {
    'headers': new_headers,
    'data': new_data,
    'deletedColumns': [deleted_col],
    'editedCells': {},
    'fullRewrite': True
}

print("\nExportiere...")
result = write_sheet(original_path, output_path, sheet_name, changes)
print("Result:", result)
