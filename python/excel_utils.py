#!/usr/bin/env python3
"""
Excel Utilities - Hilfsfunktionen für Excel-Operationen

Prüft zur Laufzeit welche Excel-Methode verfügbar ist:
1. xlwings (wenn Excel installiert) - für strukturelle Änderungen mit CF-Erhalt
2. openpyxl (immer) - für reine Datenänderungen
"""

import sys
import os

# Cache für Excel-Verfügbarkeit
_excel_available = None

def is_excel_installed():
    """
    Prüft ob Microsoft Excel installiert und per xlwings erreichbar ist.
    Das Ergebnis wird gecached.
    """
    global _excel_available
    
    if _excel_available is not None:
        return _excel_available
    
    try:
        import xlwings as xw
        # Versuche Excel App zu starten (unsichtbar)
        app = xw.App(visible=False, add_book=False)
        # Wenn das funktioniert, ist Excel verfügbar
        app.quit()
        _excel_available = True
    except Exception:
        _excel_available = False
    
    return _excel_available


def delete_columns_with_xlwings(file_path, output_path, sheet_name, column_indices):
    """
    Löscht Spalten mit xlwings (Excel).
    Dies erhält ALLE Formatierungen inkl. Conditional Formatting!
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        column_indices: Liste der zu löschenden Spalten-Indices (0-basiert)
    
    Returns:
        True bei Erfolg, False bei Fehler
    """
    if not column_indices:
        return True
    
    try:
        import xlwings as xw
        
        # Excel starten (unsichtbar)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        try:
            # Workbook öffnen
            wb = app.books.open(file_path)
            ws = wb.sheets[sheet_name]
            
            # Spalten löschen (von hinten nach vorne, 1-basiert)
            for col_idx in sorted(column_indices, reverse=True):
                excel_col = col_idx + 1  # 1-basiert
                ws.range((1, excel_col)).api.EntireColumn.Delete()
            
            # Speichern
            wb.save(output_path)
            wb.close()
            
            return True
            
        finally:
            app.quit()
            
    except Exception:
        return False


def insert_columns_with_xlwings(file_path, output_path, sheet_name, position, count, headers=None):
    """
    Fügt Spalten mit xlwings (Excel) ein.
    Dies erhält ALLE Formatierungen inkl. Conditional Formatting!
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        position: Position wo eingefügt werden soll (0-basiert)
        count: Anzahl der einzufügenden Spalten
        headers: Liste der Header für neue Spalten
    
    Returns:
        True bei Erfolg, False bei Fehler
    """
    try:
        import xlwings as xw
        
        # Excel starten (unsichtbar)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        try:
            # Workbook öffnen
            wb = app.books.open(file_path)
            ws = wb.sheets[sheet_name]
            
            excel_col = position + 1  # 1-basiert
            
            # Spalten einfügen
            for i in range(count):
                ws.range((1, excel_col + i)).api.EntireColumn.Insert()
            
            # Header setzen
            if headers:
                for i, header in enumerate(headers):
                    ws.range((1, excel_col + i)).value = header
            
            # Speichern
            wb.save(output_path)
            wb.close()
            
            return True
            
        finally:
            app.quit()
            
    except Exception:
        return False


def delete_rows_with_xlwings(file_path, output_path, sheet_name, row_indices):
    """
    Löscht Zeilen mit xlwings (Excel).
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei  
        sheet_name: Name des Sheets
        row_indices: Liste der zu löschenden Zeilen-Indices (0-basiert, ohne Header)
    
    Returns:
        True bei Erfolg, False bei Fehler
    """
    if not row_indices:
        return True
    
    try:
        import xlwings as xw
        
        # Excel starten (unsichtbar)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        try:
            # Workbook öffnen
            wb = app.books.open(file_path)
            ws = wb.sheets[sheet_name]
            
            # Zeilen löschen (von hinten nach vorne, 1-basiert, +2 für Header)
            for row_idx in sorted(row_indices, reverse=True):
                excel_row = row_idx + 2  # +2 für Header (1-basiert)
                ws.range((excel_row, 1)).api.EntireRow.Delete()
            
            # Speichern
            wb.save(output_path)
            wb.close()
            
            return True
            
        finally:
            app.quit()
            
    except Exception:
        return False


def structural_change_with_excel(file_path, output_path, sheet_name, 
                                  deleted_columns=None, inserted_columns=None,
                                  deleted_rows=None):
    """
    Führt strukturelle Änderungen mit Excel durch.
    Erhält ALLE Formatierungen inkl. Conditional Formatting.
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        deleted_columns: Liste der zu löschenden Spalten (0-basiert)
        inserted_columns: Dict mit Insert-Operationen
        deleted_rows: Liste der zu löschenden Zeilen (0-basiert)
    
    Returns:
        True bei Erfolg, False bei Fehler
    """
    if not is_excel_installed():
        return False
    
    try:
        import xlwings as xw
        import shutil
        
        # Kopiere Original-Datei zuerst
        shutil.copy2(file_path, output_path)
        
        # Excel starten (unsichtbar)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        try:
            # Workbook öffnen (die Kopie)
            wb = app.books.open(output_path)
            ws = wb.sheets[sheet_name]
            
            # 1. Spalten löschen (von hinten nach vorne)
            if deleted_columns:
                for col_idx in sorted(deleted_columns, reverse=True):
                    excel_col = col_idx + 1
                    ws.range((1, excel_col)).api.EntireColumn.Delete()
            
            # 2. Spalten einfügen
            if inserted_columns:
                operations = inserted_columns.get('operations', [])
                if not operations and inserted_columns.get('position') is not None:
                    operations = [{
                        'position': inserted_columns['position'],
                        'count': inserted_columns.get('count', 1),
                        'headers': inserted_columns.get('headers', [])
                    }]
                
                # Berücksichtige bereits gelöschte Spalten beim Offset
                deleted_before = len(deleted_columns) if deleted_columns else 0
                
                for op in sorted(operations, key=lambda x: x['position']):
                    pos = op['position'] - deleted_before
                    count = op.get('count', 1)
                    headers = op.get('headers', [])
                    excel_col = pos + 1
                    
                    for i in range(count):
                        ws.range((1, excel_col + i)).api.EntireColumn.Insert()
                    
                    for i, header in enumerate(headers):
                        ws.range((1, excel_col + i)).value = header
            
            # 3. Zeilen löschen (von hinten nach vorne)
            if deleted_rows:
                for row_idx in sorted(deleted_rows, reverse=True):
                    excel_row = row_idx + 2  # +2 für Header
                    ws.range((excel_row, 1)).api.EntireRow.Delete()
            
            # Speichern und schließen
            wb.save()
            wb.close()
            
            return True
            
        finally:
            app.quit()
            
    except Exception:
        return False


def get_excel_status():
    """
    Gibt Status-Information über Excel-Verfügbarkeit zurück.
    
    Returns:
        Dict mit 'available' (bool) und 'message' (str)
    """
    available = is_excel_installed()
    
    if available:
        return {
            'available': True,
            'message': 'Microsoft Excel verfügbar - strukturelle Änderungen mit CF-Erhalt möglich'
        }
    else:
        return {
            'available': False,
            'message': 'Microsoft Excel nicht verfügbar - strukturelle Änderungen können CF beeinträchtigen'
        }


if __name__ == '__main__':
    import json
    # Test: Excel-Status ausgeben
    status = get_excel_status()
    print(json.dumps(status, indent=2))
