#!/usr/bin/env python3
"""Debug-Script um CF-Transformation zu testen"""

from openpyxl.utils import get_column_letter, column_index_from_string

# Beispiel-Konfiguration (anpassen an Ihren Fall)
deleted_columns = [3]  # Spalte C gelöscht
inserted_ops = [{'position': 7, 'count': 1}, {'position': 9, 'count': 1}]  # H und J eingefügt

def transform_column(col_letter):
    """Transformiert eine Spalte basierend auf gelöschten/eingefügten Spalten."""
    col_num = column_index_from_string(col_letter)
    original_col = col_num
    
    print(f"  Start: {col_letter} = {col_num}")
    
    # Zuerst: Spalten-Löschungen anwenden
    for del_col in sorted(deleted_columns):
        if col_num > del_col:
            print(f"    Löschung bei {del_col}: {col_num} -> {col_num - 1}")
            col_num -= 1
        elif col_num == del_col:
            print(f"    Spalte {col_num} wurde gelöscht!")
            return None
    
    # Dann: Spalten-Einfügungen anwenden
    for op in sorted(inserted_ops, key=lambda x: x.get('position', 0)):
        pos = op.get('position', 0)
        count = op.get('count', 1)
        if col_num >= pos:
            print(f"    Einfügung bei {pos} (count={count}): {col_num} -> {col_num + count}")
            col_num += count
    
    result = get_column_letter(col_num)
    print(f"  Ergebnis: {col_letter}({original_col}) -> {result}({col_num})")
    return result

# Test: Was passiert mit den CF-Spalten?
print("=== CF-Spalten-Transformation ===")
print(f"deleted_columns = {deleted_columns}")
print(f"inserted_ops = {inserted_ops}")
print()

# Teste verschiedene Spalten
test_columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
print("Transformation aller Spalten:")
for col in test_columns:
    result = transform_column(col)
    print()
