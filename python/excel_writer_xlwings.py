#!/usr/bin/env python3
"""
Excel Writer für Excel Data Sync Pro - xlwings Version
Verwendet xlwings für native Excel-Kompatibilität

Der große Vorteil von xlwings:
- Alle Operationen werden von Excel selbst durchgeführt
- Perfekte Erhaltung von CF, Styles, Formeln, Tabellen
- Automatische Anpassung aller Referenzen bei strukturellen Änderungen
"""

import json
import sys
import os
import shutil
import platform
import subprocess
from datetime import datetime, date

# Für embedded Python auf Windows: pywin32 DLLs finden
if platform.system() == 'Windows':
    # pywin32_system32 DLLs
    pywin32_dll = os.path.join(sys.prefix, 'Lib', 'site-packages', 'pywin32_system32')
    if os.path.exists(pywin32_dll):
        os.environ['PATH'] = pywin32_dll + os.pathsep + os.environ.get('PATH', '')
    # DLLs im Python-Verzeichnis (embedded)
    python_dir = os.path.dirname(sys.executable)
    if os.path.exists(os.path.join(python_dir, 'pythoncom311.dll')):
        os.environ['PATH'] = python_dir + os.pathsep + os.environ.get('PATH', '')
    # win32 Module
    win32_dir = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32')
    if os.path.exists(win32_dir):
        sys.path.insert(0, win32_dir)
    win32_lib = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32', 'lib')
    if os.path.exists(win32_lib):
        sys.path.insert(0, win32_lib)

# Import xlwings mit detaillierter Fehlerbehandlung
try:
    import xlwings as xw
except ImportError as e:
    print(f"xlwings import failed: {e}", file=sys.stderr)
    print(json.dumps({"success": False, "error": f"xlwings import failed: {e}"}))
    sys.exit(1)
except Exception as e:
    print(f"xlwings import error: {e}", file=sys.stderr)
    print(json.dumps({"success": False, "error": f"xlwings import error: {e}"}))
    sys.exit(1)


def kill_excel_instances():
    """Beendet alle laufenden Excel-Instanzen - plattformübergreifend"""
    import time
    system = platform.system()
    
    # Methode 1: xlwings Apps beenden (funktioniert auf allen Plattformen)
    try:
        for app in xw.apps:
            try:
                for book in app.books:
                    try:
                        book.close()
                    except:
                        pass
                app.quit()
            except:
                pass
    except Exception:
        pass
    
    time.sleep(0.2)
    
    if system == 'Darwin':  # macOS
        # Methode 2: Mit pkill -9 SOFORT beenden (keine Dialoge!)
        try:
            result = subprocess.run(['pgrep', '-x', 'Microsoft Excel'], 
                                    capture_output=True, timeout=1)
            if result.returncode == 0:
                subprocess.run(['pkill', '-9', 'Microsoft Excel'], 
                               capture_output=True, timeout=2)
                time.sleep(0.3)
        except:
            pass
        
        # Methode 3: Falls immer noch da, mit killall
        try:
            result = subprocess.run(['pgrep', '-x', 'Microsoft Excel'], 
                                    capture_output=True, timeout=1)
            if result.returncode == 0:
                subprocess.run(['killall', '-9', 'Microsoft Excel'], 
                               capture_output=True, timeout=2)
                time.sleep(0.3)
        except:
            pass
    
    elif system == 'Windows':  # Windows
        # Methode 2: Mit taskkill Excel beenden
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], 
                          capture_output=True, timeout=5)
            time.sleep(0.3)
        except:
            pass


def hide_excel():
    """Versteckt Excel - plattformübergreifend"""
    system = platform.system()
    
    if system == 'Darwin':  # macOS
        try:
            subprocess.run(['osascript', '-e', 
                'tell application "System Events" to set visible of process "Microsoft Excel" to false'], 
                capture_output=True, timeout=2)
        except:
            pass
    
    elif system == 'Windows':  # Windows
        # Auf Windows: Excel-Fenster minimieren via xlwings
        try:
            for app in xw.apps:
                try:
                    # visible=False versteckt das Fenster komplett
                    app.visible = False
                except:
                    pass
        except:
            pass


def _get_column_letter(col_idx):
    """Konvertiert Spalten-Index (1-basiert) zu Buchstaben (A, B, ..., Z, AA, ...)"""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def hex_to_rgb(hex_color):
    """Konvertiert Hex ('#FF0000') zu RGB-Tuple (255, 0, 0)"""
    if not hex_color:
        return None
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    if len(hex_color) == 6:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b)
    return None


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


def write_sheet_xlwings(file_path, output_path, sheet_name, changes):
    """
    Schreibt Änderungen in ein Excel-Sheet mit xlwings
    
    VORTEILE von xlwings:
    - Excel macht ALLE strukturellen Änderungen
    - CF-Bereiche werden AUTOMATISCH angepasst
    - Formeln werden AUTOMATISCH aktualisiert
    - Perfekte Kompatibilität
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        changes: Dict mit allen Änderungen
    
    Returns:
        Dict mit success und ggf. error
    """
    print(f"[xlwings_writer] write_sheet_xlwings aufgerufen", file=sys.stderr)
    print(f"[xlwings_writer] file_path: {file_path}", file=sys.stderr)
    print(f"[xlwings_writer] sheet_name: {sheet_name}", file=sys.stderr)
    print(f"[xlwings_writer] fromFile: {changes.get('fromFile', False)}", file=sys.stderr)
    print(f"[xlwings_writer] fullRewrite: {changes.get('fullRewrite', False)}", file=sys.stderr)
    print(f"[xlwings_writer] structuralChange: {changes.get('structuralChange', False)}", file=sys.stderr)
    print(f"[xlwings_writer] headers count: {len(changes.get('headers', []))}", file=sys.stderr)
    print(f"[xlwings_writer] data count: {len(changes.get('data', []))}", file=sys.stderr)
    print(f"[xlwings_writer] editedCells count: {len(changes.get('editedCells', {}))}", file=sys.stderr)
    
    try:
        # Parameter extrahieren
        headers = changes.get('headers', [])
        data = changes.get('data', [])
        edited_cells = changes.get('editedCells', {})
        row_highlights = changes.get('rowHighlights', {})
        deleted_columns = changes.get('deletedColumns', [])
        inserted_columns = changes.get('insertedColumns')
        hidden_columns = changes.get('hiddenColumns', [])
        hidden_rows = changes.get('hiddenRows', [])
        deleted_rows = changes.get('deletedRows', [])
        from_file = changes.get('fromFile', False)
        full_rewrite = changes.get('fullRewrite', False)
        structural_change = changes.get('structuralChange', False)
        cleared_row_highlights = changes.get('clearedRowHighlights', [])
        
        # Kopiere Original-Datei zum Ziel (falls unterschiedlich)
        if file_path != output_path:
            shutil.copy2(file_path, output_path)
        
        import time
        
        # WICHTIG: Beende zuerst alle laufenden Excel-Instanzen
        # um Dialog-Probleme zu vermeiden ("Änderungen speichern?")
        print("[xlwings_writer] Beende Excel-Instanzen...", file=sys.stderr)
        kill_excel_instances()
        time.sleep(0.3)  # Kurz warten bis Excel wirklich beendet ist
        
        # Excel-App starten (ohne with-Kontext, um Cleanup-Probleme zu vermeiden)
        print("[xlwings_writer] Starte Excel-App...", file=sys.stderr)
        app = xw.App(visible=False, add_book=False)
        print("[xlwings_writer] Excel-App gestartet!", file=sys.stderr)
        app.display_alerts = False
        app.screen_updating = False
        
        # Excel verstecken
        hide_excel()
        
        # Workbook öffnen
        print(f"[xlwings_writer] Öffne Workbook: {output_path}", file=sys.stderr)
        wb = app.books.open(output_path)
        print(f"[xlwings_writer] Workbook geöffnet!", file=sys.stderr)
        
        # Sheet finden
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            wb.close()
            kill_excel_instances()
            return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
        
        ws = wb.sheets[sheet_name]
        
        # =====================================================================
        # FALL 1: fromFile - Nur versteckte Spalten/Zeilen setzen
        # =====================================================================
        if from_file:
            _apply_hidden_columns_xlwings(ws, hidden_columns)
            _apply_hidden_rows_xlwings(ws, hidden_rows)
            wb.save()
            wb.close()
            kill_excel_instances()
            return {'success': True, 'outputPath': output_path, 'method': 'xlwings'}
        
        # =====================================================================
        # FALL 2: Strukturelle Änderungen
        # xlwings/Excel macht das PERFEKT - CF wird automatisch angepasst!
        # =====================================================================
        
        # SCHRITT 1: SPALTEN LÖSCHEN (von hinten nach vorne)
        # Verwende xlwings native Syntax (funktioniert auf macOS und Windows)
        if deleted_columns:
            for col_idx in sorted(deleted_columns, reverse=True):
                excel_col = col_idx + 1  # 1-basiert
                col_letter = _get_column_letter(excel_col)
                # Verwende Spalten-Range mit delete() - funktioniert plattformübergreifend
                ws.range(f'{col_letter}:{col_letter}').delete()
        
        # SCHRITT 2: SPALTEN EINFÜGEN
        # WICHTIG: Bei xlwings MUSS die Spalte IMMER in Excel eingefügt werden,
        # damit Excel die Formatierungen (CF, Fills, etc.) automatisch verschiebt!
        # Danach schreiben wir die Daten (die bereits die neue Spalte enthalten).
        if inserted_columns:
            operations = inserted_columns.get('operations', [])
            if not operations and inserted_columns.get('position') is not None:
                # Altes Format
                operations = [{
                    'position': inserted_columns['position'],
                    'count': inserted_columns.get('count', 1),
                    'headers': inserted_columns.get('headers', []),
                    'sourceColumn': inserted_columns.get('sourceColumn')
                }]
            
            # Berücksichtige bereits gelöschte Spalten beim Offset
            deleted_before = len(deleted_columns) if deleted_columns else 0
            
            # Akkumulierter Offset für bereits eingefügte Spalten
            inserted_offset = 0
            
            for op in sorted(operations, key=lambda x: x['position']):
                pos = op['position'] - deleted_before + inserted_offset
                count = op.get('count', 1)
                op_headers = op.get('headers', [])
                source_column = op.get('sourceColumn')  # Referenzspalte für Formatierung
                excel_col = pos + 1  # 1-basiert
                col_letter = _get_column_letter(excel_col)
                
                for i in range(count):
                    insert_letter = _get_column_letter(excel_col + i)
                    
                    # Insert-Befehl: Fügt vor der angegebenen Spalte ein
                    # shift='right' verschiebt existierende Zellen nach rechts
                    ws.range(f'{insert_letter}:{insert_letter}').insert(shift='right')
                    
                    # HINWEIS: Wir kopieren KEINE Formatierung von einer Quellspalte,
                    # da das copy() auch Werte mitkopiert und zu falschen Daten führt.
                    # Die neue Spalte erhält Standard-Formatierung.
                    # Die Daten werden später über editedCells geschrieben.
                
                # Header setzen
                for i, header in enumerate(op_headers):
                    ws.range((1, excel_col + i)).value = header
                
                # Offset für nächste Operation erhöhen
                inserted_offset += count
        
        # SCHRITT 3: ZEILEN LÖSCHEN (von hinten nach vorne)
        if deleted_rows:
            for row_idx in sorted(deleted_rows, reverse=True):
                excel_row = row_idx + 2  # +2 für Header (1-basiert)
                ws.range(f'{excel_row}:{excel_row}').delete()
        
        # SCHRITT 4: Original-Zeilenzahl ermitteln für Löschung überschüssiger Zeilen
        used_range = ws.used_range
        original_row_count = used_range.last_cell.row - 1 if used_range else 0  # Ohne Header
        new_row_count = len(data)
        
        # Überschüssige Zeilen löschen - OPTIMIERUNG: Alle auf einmal statt einzeln!
        if full_rewrite and new_row_count < original_row_count:
            rows_to_delete = original_row_count - new_row_count
            # Lösche alle überschüssigen Zeilen in einem Bereich (VIEL schneller!)
            first_row_to_delete = new_row_count + 2  # +2 für Header (1-basiert)
            last_row_to_delete = original_row_count + 1  # +1 für Header (1-basiert)
            print(f"[xlwings_writer] Lösche Zeilen {first_row_to_delete} bis {last_row_to_delete} ({rows_to_delete} Zeilen)", file=sys.stderr)
            ws.range(f'{first_row_to_delete}:{last_row_to_delete}').delete()
            
            # WICHTIG: Bei Filter werden nicht nur Zeilen gelöscht, sondern die
            # verbleibenden Zeilen enthalten andere Daten (gefilterte Zeilen).
            # Wir müssen alle Daten komplett neu schreiben!
            print(f"[xlwings_writer] Filter aktiv - schreibe {new_row_count} Zeilen neu", file=sys.stderr)
            if data and len(data) > 0 and len(data[0]) > 0:
                num_cols = len(data[0])
                # Schreibe alle Daten als Block (Zeile 2 bis new_row_count+1)
                start_cell = ws.range((2, 1))  # Zeile 2, Spalte A (nach Header)
                end_cell = ws.range((new_row_count + 1, num_cols))
                ws.range(start_cell, end_cell).value = data
                print(f"[xlwings_writer] Alle Daten geschrieben: {new_row_count} x {num_cols}", file=sys.stderr)
        
        # SCHRITT 5: ZELL-EDITS (geänderte Zellen)
        # NUR bei editedCells ohne Filter-Rewrite
        # OPTIMIERUNG: Gruppiere nach Spalten und schreibe spaltenweise statt zellweise
        if edited_cells:
            print(f"[xlwings_writer] Schreibe {len(edited_cells)} geänderte Zellen...", file=sys.stderr)
            
            # Gruppiere Zellen nach Spaltenindex
            columns_data = {}  # col_idx -> [(row_idx, value), ...]
            for key, value in edited_cells.items():
                if key.startswith('_'):
                    continue
                parts = key.split('-')
                if len(parts) != 2:
                    continue
                row_idx = int(parts[0])
                col_idx = int(parts[1])
                if col_idx not in columns_data:
                    columns_data[col_idx] = []
                columns_data[col_idx].append((row_idx, value))
            
            # Schreibe jede Spalte als Block (viel schneller!)
            for col_idx, cells in columns_data.items():
                # Sortiere nach Zeile
                cells.sort(key=lambda x: x[0])
                
                # Prüfe ob es ein zusammenhängender Block ist
                min_row = cells[0][0]
                max_row = cells[-1][0]
                
                if len(cells) == (max_row - min_row + 1):
                    # Zusammenhängender Block - schreibe als Range
                    values = [[c[1]] for c in cells]  # 2D-Array für xlwings
                    start_cell = ws.range((min_row + 2, col_idx + 1))
                    end_cell = ws.range((max_row + 2, col_idx + 1))
                    ws.range(start_cell, end_cell).value = values
                    print(f"[xlwings_writer] Spalte {col_idx}: Block {min_row}-{max_row} ({len(cells)} Zellen)", file=sys.stderr)
                else:
                    # Nicht zusammenhängend - einzeln schreiben
                    for row_idx, value in cells:
                        cell = ws.range((row_idx + 2, col_idx + 1))
                        apply_cell_value(cell, value)
                    print(f"[xlwings_writer] Spalte {col_idx}: {len(cells)} einzelne Zellen", file=sys.stderr)
        
        # SCHRITT 8: VERSTECKTE SPALTEN
        _apply_hidden_columns_xlwings(ws, hidden_columns, len(headers) if headers else None)
        
        # SCHRITT 9: VERSTECKTE ZEILEN
        _apply_hidden_rows_xlwings(ws, hidden_rows, len(data) if data else None)
        
        # SCHRITT 10: ROW HIGHLIGHTS
        if row_highlights and len(row_highlights) > 0:
            _apply_row_highlights_xlwings(ws, row_highlights, len(headers) if headers else ws.used_range.last_cell.column)
        
        # SCHRITT 11: CLEARED ROW HIGHLIGHTS
        if cleared_row_highlights and len(cleared_row_highlights) > 0:
            for row_idx in cleared_row_highlights:
                excel_row = row_idx + 2
                num_cols = len(headers) if headers else ws.used_range.last_cell.column
                for col_idx in range(1, num_cols + 1):
                    cell = ws.range((excel_row, col_idx))
                    cell.color = None  # Farbe entfernen
        
        # Speichern und schließen
        wb.save()
        wb.close()
        
        # SOFORT Excel beenden!
        kill_excel_instances()
        
        return {
            'success': True, 
            'outputPath': output_path,
            'method': 'xlwings',
            'cfPreserved': True
        }
        
    except Exception as e:
        import traceback
        error_msg = str(e)
        tb = traceback.format_exc()
        print(f"[xlwings Writer] ERROR: {error_msg}", file=sys.stderr, flush=True)
        print(f"[xlwings Writer] Traceback: {tb}", file=sys.stderr, flush=True)
        kill_excel_instances()  # Auch bei Fehler beenden
        return {
            'success': False, 
            'error': error_msg,
            'traceback': tb
        }


def _apply_hidden_columns_xlwings(ws, hidden_columns, max_cols=None):
    """Setzt versteckte Spalten mit xlwings
    
    Auf macOS funktioniert api.column_hidden nicht zuverlässig.
    Stattdessen verwenden wir column_width = 0.
    """
    if hidden_columns is None or not hidden_columns:
        return
    
    hidden_set = set(hidden_columns)
    
    for col_idx in hidden_set:
        try:
            col_letter = _get_column_letter(col_idx + 1)
            col_range = ws.range(f'{col_letter}:{col_letter}')
            # macOS: column_width = 0 funktioniert zuverlässig
            col_range.column_width = 0
        except Exception as e:
            print(f"[xlwings] Hidden column {col_idx} FEHLER: {e}", file=sys.stderr, flush=True)


def _apply_hidden_rows_xlwings(ws, hidden_rows, max_rows=None):
    """Setzt versteckte Zeilen mit xlwings
    
    Auf macOS funktioniert api.row_hidden nicht zuverlässig.
    Stattdessen verwenden wir row_height = 0.
    """
    if hidden_rows is None or not hidden_rows:
        return
    
    hidden_set = set(hidden_rows)
    
    for row_idx in hidden_set:
        try:
            excel_row = row_idx + 2  # +2 für Header (1-basiert)
            row_range = ws.range(f'{excel_row}:{excel_row}')
            # macOS: row_height = 0 funktioniert zuverlässig
            row_range.row_height = 0
        except Exception as e:
            print(f"[xlwings] Hidden row {row_idx} FEHLER: {e}", file=sys.stderr, flush=True)


def _apply_row_highlights_xlwings(ws, row_highlights, num_columns):
    """Wendet Zeilen-Highlights an mit xlwings"""
    
    highlight_colors = {
        'green': (144, 238, 144),   # LightGreen
        'yellow': (255, 255, 0),    # Yellow
        'orange': (255, 165, 0),    # Orange
        'red': (255, 107, 107),     # LightRed
        'blue': (135, 206, 235),    # SkyBlue
        'purple': (221, 160, 221)   # Plum
    }
    
    last_col_letter = _get_column_letter(num_columns)
    
    for row_idx_str, color in row_highlights.items():
        row_idx = int(row_idx_str)
        excel_row = row_idx + 2  # +2 für 1-basiert und Header
        
        if isinstance(color, str) and color.startswith('#'):
            rgb = hex_to_rgb(color)
        else:
            rgb = highlight_colors.get(color, (255, 255, 0))
        
        if rgb:
            # Ganze Zeile auf einmal färben (VIEL schneller als einzelne Zellen!)
            try:
                row_range = ws.range(f'A{excel_row}:{last_col_letter}{excel_row}')
                row_range.color = rgb
            except Exception as e:
                print(f"[xlwings] Fehler beim Färben von Zeile {excel_row}: {e}", file=sys.stderr, flush=True)


def check_excel_available():
    """Prüft ob Microsoft Excel verfügbar ist"""
    try:
        app = xw.App(visible=False, add_book=False)
        app.quit()
        return {
            'success': True,
            'available': True,
            'excelAvailable': True,
            'xlwingsAvailable': True,
            'method': 'xlwings',
            'message': 'Microsoft Excel verfügbar - xlwings wird für optimale Formaterhaltung verwendet'
        }
    except Exception as e:
        return {
            'success': True,
            'available': False,
            'excelAvailable': False,
            'xlwingsAvailable': True,
            'method': 'openpyxl',
            'message': f'Microsoft Excel nicht verfügbar: {e} - openpyxl wird als Fallback verwendet'
        }


def main():
    """Hauptfunktion - liest Befehle von stdin oder Argumenten"""
    # Debug: Script gestartet
    print("[xlwings_writer] Script gestartet", file=sys.stderr)
    print(f"[xlwings_writer] Python: {sys.executable}", file=sys.stderr)
    print(f"[xlwings_writer] Platform: {platform.system()}", file=sys.stderr)
    print(f"[xlwings_writer] Args: {sys.argv}", file=sys.stderr)
    
    # Auf Windows: Stelle sicher dass stdin/stdout UTF-8 verwenden
    import io
    if sys.platform == 'win32':
        sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        print("[xlwings_writer] UTF-8 wrapper applied", file=sys.stderr)
    
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
        
        result = write_sheet_xlwings(
            params.get('filePath'),
            params.get('outputPath'),
            params.get('sheetName'),
            params.get('changes', {})
        )
        print(json.dumps(result, ensure_ascii=False))
    
    elif command == 'check_excel':
        result = check_excel_available()
        print(json.dumps(result, ensure_ascii=False))
    
    else:
        print(json.dumps({'success': False, 'error': f'Unbekannter Befehl: {command}'}))
        sys.exit(1)


if __name__ == '__main__':
    main()
