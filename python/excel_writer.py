#!/usr/bin/env python3
"""
Excel Writer für Excel Data Sync Pro
Verwendet openpyxl für bessere Kompatibilität mit Excel-Formaten

Der große Vorteil von openpyxl: 
- Öffnet die Original-Datei und modifiziert nur die geänderten Zellen
- Behält ALLE Formatierungen, bedingte Formatierungen, Tabellen, etc.
"""

import json
import sys
import os
from datetime import datetime, date
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill, Font, Alignment, Border


def hex_to_argb(hex_color):
    """Konvertiert Hex ('#FF0000') zu ARGB ('FFFF0000')"""
    if not hex_color:
        return None
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    if len(hex_color) == 6:
        return 'FF' + hex_color.upper()
    return hex_color.upper()


def apply_cell_value(cell, value):
    """Setzt den Wert einer Zelle mit korrektem Typ"""
    if value is None or value == '':
        cell.value = None
    elif isinstance(value, (int, float)):
        cell.value = value
    elif isinstance(value, str):
        # Versuche Datum zu parsen
        for fmt in ['%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y']:
            try:
                cell.value = datetime.strptime(value, fmt)
                return
            except ValueError:
                pass
        # Versuche Zahl zu parsen
        try:
            if '.' in value and ',' not in value:
                cell.value = float(value)
            elif ',' in value and '.' not in value:
                cell.value = float(value.replace(',', '.'))
            else:
                cell.value = value
        except ValueError:
            cell.value = value
    else:
        cell.value = str(value)


def write_sheet(file_path, output_path, sheet_name, changes):
    """
    Schreibt Änderungen in ein Excel-Sheet
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        changes: Dict mit allen Änderungen:
            - headers: Array mit Header-Namen
            - data: 2D-Array mit Daten
            - editedCells: Dict mit geänderten Zellen {key: value}
            - cellStyles: Dict mit Zell-Styles {key: color}
            - rowHighlights: Dict mit Zeilenfarben {rowIdx: color}
            - deletedColumns: Array mit gelöschten Spalten-Indices
            - insertedColumns: Info über eingefügte Spalten
            - hiddenColumns: Array mit versteckten Spalten
            - hiddenRows: Array mit versteckten Zeilen
    
    Returns:
        Dict mit success und ggf. error
    """
    try:
        # Original-Workbook laden
        wb = load_workbook(file_path)
        
        if sheet_name not in wb.sheetnames:
            return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
        
        ws = wb[sheet_name]
        
        headers = changes.get('headers', [])
        data = changes.get('data', [])
        edited_cells = changes.get('editedCells', {})
        cell_styles = changes.get('cellStyles', {})
        row_highlights = changes.get('rowHighlights', {})
        deleted_columns = changes.get('deletedColumns', [])
        inserted_columns = changes.get('insertedColumns')
        hidden_columns = changes.get('hiddenColumns', [])
        hidden_rows = changes.get('hiddenRows', [])
        row_mapping = changes.get('rowMapping')
        
        # 1. Spalten löschen (von hinten nach vorne)
        if deleted_columns:
            for col_idx in sorted(deleted_columns, reverse=True):
                ws.delete_cols(col_idx + 1)  # 1-basiert
        
        # 2. Spalten einfügen
        if inserted_columns:
            operations = inserted_columns.get('operations', [])
            if not operations and inserted_columns.get('position') is not None:
                # Altes Format
                operations = [{
                    'position': inserted_columns['position'],
                    'count': inserted_columns.get('count', 1),
                    'headers': inserted_columns.get('headers', [])
                }]
            
            # Sortiere aufsteigend nach Position
            operations.sort(key=lambda x: x['position'])
            
            cumulative_shift = 0
            for op in operations:
                pos = op['position'] + cumulative_shift
                count = op.get('count', 1)
                op_headers = op.get('headers', [])
                
                # Spalten einfügen (verschiebt automatisch alle anderen)
                ws.insert_cols(pos + 1, count)  # 1-basiert
                
                # Header setzen
                for i, header in enumerate(op_headers):
                    ws.cell(row=1, column=pos + 1 + i, value=header)
                
                cumulative_shift += count
        
        # 3. Header aktualisieren
        for col_idx, header in enumerate(headers):
            ws.cell(row=1, column=col_idx + 1, value=header)
        
        # 4. Daten aktualisieren (nur geänderte Zellen wenn editedCells vorhanden)
        if edited_cells and '_columnInserted' not in edited_cells and '_columnDeleted' not in edited_cells:
            # Nur geänderte Zellen schreiben
            for key, value in edited_cells.items():
                if key.startswith('_'):
                    continue
                parts = key.split('-')
                if len(parts) != 2:
                    continue
                row_idx = int(parts[0])
                col_idx = int(parts[1])
                cell = ws.cell(row=row_idx + 2, column=col_idx + 1)  # +2 für Header
                apply_cell_value(cell, value)
        else:
            # Alle Daten schreiben
            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
                    apply_cell_value(cell, value)
        
        # 5. Cell Styles anwenden (Hintergrundfarben)
        if cell_styles:
            for key, color in cell_styles.items():
                if not color:
                    continue
                parts = key.split('-')
                if len(parts) != 2:
                    continue
                row_idx = int(parts[0])
                col_idx = int(parts[1])
                cell = ws.cell(row=row_idx + 1, column=col_idx + 1)  # +1 weil Styles 0-basiert
                argb = hex_to_argb(color)
                if argb:
                    cell.fill = PatternFill(start_color=argb, end_color=argb, fill_type='solid')
        
        # 6. Row Highlights anwenden
        if row_highlights:
            highlight_colors = {
                'green': 'FF90EE90',
                'yellow': 'FFFFFF00',
                'orange': 'FFFFA500',
                'red': 'FFFF6B6B',
                'blue': 'FF87CEEB',
                'purple': 'FFDDA0DD'
            }
            
            for row_idx_str, color in row_highlights.items():
                row_idx = int(row_idx_str)
                excel_row = row_idx + 2  # +2 für 1-basiert und Header
                
                if color.startswith('#'):
                    argb = hex_to_argb(color)
                else:
                    argb = highlight_colors.get(color, 'FFFFFF00')
                
                # Alle Zellen in der Zeile färben
                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    cell.fill = PatternFill(start_color=argb, end_color=argb, fill_type='solid')
        
        # 7. Versteckte Spalten setzen
        if hidden_columns is not None:
            hidden_set = set(hidden_columns)
            for col_idx in range(len(headers)):
                col_letter = get_column_letter(col_idx + 1)
                ws.column_dimensions[col_letter].hidden = col_idx in hidden_set
        
        # 8. Versteckte Zeilen setzen
        if hidden_rows is not None:
            hidden_set = set(hidden_rows)
            for row_idx in range(len(data)):
                excel_row = row_idx + 2
                ws.row_dimensions[excel_row].hidden = row_idx in hidden_set
        
        # Speichern
        wb.save(output_path)
        wb.close()
        
        return {'success': True, 'outputPath': output_path}
        
    except Exception as e:
        import traceback
        return {
            'success': False, 
            'error': str(e),
            'traceback': traceback.format_exc()
        }


def main():
    """Hauptfunktion - liest Befehle von stdin oder Argumenten"""
    if len(sys.argv) < 2:
        print(json.dumps({'success': False, 'error': 'Kein Befehl angegeben'}))
        sys.exit(1)
    
    command = sys.argv[1]
    
    if command == 'write_sheet':
        # Daten von stdin lesen (für große Datenmengen)
        input_data = sys.stdin.read()
        try:
            params = json.loads(input_data)
        except json.JSONDecodeError as e:
            print(json.dumps({'success': False, 'error': f'JSON Parse Error: {str(e)}'}))
            sys.exit(1)
        
        result = write_sheet(
            params.get('filePath'),
            params.get('outputPath'),
            params.get('sheetName'),
            params.get('changes', {})
        )
        print(json.dumps(result, ensure_ascii=False))
    
    else:
        print(json.dumps({'success': False, 'error': f'Unbekannter Befehl: {command}'}))
        sys.exit(1)


if __name__ == '__main__':
    main()
