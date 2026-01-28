#!/usr/bin/env python3
"""
Excel Utilities - Hilfsfunktionen für Excel-Operationen

Prüft zur Laufzeit welche Excel-Methode verfügbar ist:
1. xlwings (wenn Excel installiert) - für strukturelle Änderungen mit CF-Erhalt
2. openpyxl (immer) - für reine Datenänderungen
"""

import sys
import os
import json
import platform

# Für embedded Python auf Windows: pywin32 DLLs finden
if platform.system() == 'Windows':
    # Methode 1: pywin32_system32 im site-packages
    pywin32_dll = os.path.join(sys.prefix, 'Lib', 'site-packages', 'pywin32_system32')
    if os.path.exists(pywin32_dll):
        os.environ['PATH'] = pywin32_dll + os.pathsep + os.environ.get('PATH', '')
    
    # Methode 2: DLLs im Python-Verzeichnis selbst (embedded Python)
    python_dir = os.path.dirname(sys.executable)
    if os.path.exists(os.path.join(python_dir, 'pythoncom311.dll')):
        os.environ['PATH'] = python_dir + os.pathsep + os.environ.get('PATH', '')
    
    # Methode 3: win32 Verzeichnis hinzufügen (für pywintypes)
    win32_dir = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32')
    if os.path.exists(win32_dir):
        sys.path.insert(0, win32_dir)
    
    # Methode 4: win32/lib Verzeichnis (für win32con etc.)
    win32_lib = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32', 'lib')
    if os.path.exists(win32_lib):
        sys.path.insert(0, win32_lib)

# Cache für Excel-Verfügbarkeit
_excel_available = None

def is_excel_installed():
    """
    Prüft ob Microsoft Excel installiert und per xlwings erreichbar ist.
    Das Ergebnis wird gecached.
    
    Gibt False zurück wenn:
    - xlwings nicht installiert ist
    - Excel nicht installiert ist
    - Excel nicht gestartet werden kann
    """
    global _excel_available
    
    if _excel_available is not None:
        return _excel_available
    
    # Auf Windows: Prüfe zuerst ob pywin32 funktioniert
    if platform.system() == 'Windows':
        try:
            import win32com.client
            print(f"[excel_utils] pywin32 (win32com) geladen", file=sys.stderr)
        except ImportError as e:
            print(f"[excel_utils] pywin32 nicht verfügbar: {e}", file=sys.stderr)
            _excel_available = False
            return False
        except Exception as e:
            print(f"[excel_utils] pywin32 Fehler: {e}", file=sys.stderr)
            _excel_available = False
            return False
    
    # Prüfe zuerst ob xlwings überhaupt importiert werden kann
    try:
        import xlwings as xw
        print(f"[excel_utils] xlwings importiert", file=sys.stderr)
    except ImportError as e:
        print(f"[excel_utils] xlwings nicht installiert: {e}", file=sys.stderr)
        _excel_available = False
        return False
    except Exception as e:
        print(f"[excel_utils] xlwings Import-Fehler: {e}", file=sys.stderr)
        _excel_available = False
        return False
    
    # xlwings ist da, jetzt Excel testen
    try:
        # Versuche Excel App zu starten (unsichtbar)
        app = xw.App(visible=False, add_book=False)
        # Wenn das funktioniert, ist Excel verfügbar
        app.quit()
        _excel_available = True
        print(f"[excel_utils] Microsoft Excel verfügbar", file=sys.stderr)
    except Exception as e:
        # Debug-Ausgabe für Fehlerbehebung
        print(f"[excel_utils] Excel nicht verfügbar: {type(e).__name__}: {e}", file=sys.stderr)
        _excel_available = False
    
    return _excel_available


def reset_excel_cache():
    """Setzt den Excel-Cache zurück (für Tests oder nach Neuinstallation)"""
    global _excel_available
    _excel_available = None


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
            'method': 'xlwings',
            'message': 'Microsoft Excel verfügbar - xlwings wird für optimale Formaterhaltung verwendet'
        }
    else:
        return {
            'available': False,
            'method': 'openpyxl',
            'message': 'Microsoft Excel nicht verfügbar - openpyxl wird als Fallback verwendet'
        }


if __name__ == '__main__':
    # Auf Windows: Stelle sicher dass stdout UTF-8 verwendet
    import io
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # Kommandozeilen-Interface für Excel-Check
    if len(sys.argv) > 1 and sys.argv[1] == 'check_excel':
        status = get_excel_status()
        print(json.dumps(status))
    else:
        # Standard: Status ausgeben
        status = get_excel_status()
        print(json.dumps(status, indent=2))
