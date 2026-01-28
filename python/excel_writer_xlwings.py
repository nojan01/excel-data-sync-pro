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
        inserted_rows = changes.get('insertedRowInfo')  # Info über eingefügte Zeilen
        row_order = changes.get('rowOrder')  # Zeilen-Reihenfolge bei Verschiebung
        from_file = changes.get('fromFile', False)
        full_rewrite = changes.get('fullRewrite', False)
        structural_change = changes.get('structuralChange', False)
        cleared_row_highlights = changes.get('clearedRowHighlights', [])
        row_mapping = changes.get('rowMapping')  # Filter-Mapping: [originalIndex, ...]
        column_order = changes.get('columnOrder')  # Spalten-Reihenfolge: [oldIdx0, oldIdx1, ...]
        
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
        # 
        # REIHENFOLGE (wie manueller Test):
        # ZEILEN: löschen → verschieben → ausblenden → einfügen → markieren → speichern
        # SPALTEN: löschen → verschieben → einfügen → ausblenden
        # =====================================================================
        
        # =====================================================================
        # ZEILEN-OPERATIONEN ZUERST
        # =====================================================================
        
        # SCHRITT 1: ZEILEN LÖSCHEN (von hinten nach vorne)
        if deleted_rows:
            for row_idx in sorted(deleted_rows, reverse=True):
                excel_row = row_idx + 2  # +2 für Header (1-basiert)
                ws.range(f'{excel_row}:{excel_row}').delete()
        
        # SCHRITT 2: ZEILEN VERSCHIEBEN (rowOrder)
        # Die Daten wurden bereits im Frontend umgeordnet - schreibe als Block
        rows_reordered = False
        if row_order and len(row_order) > 0:
            print(f"[xlwings_writer] rowOrder erkannt - schreibe umgeordnete Daten", file=sys.stderr)
            
            # Alle Daten als Block schreiben (schnell!)
            if data and len(data) > 0:
                num_rows = len(data)
                num_cols = len(data[0]) if data[0] else len(headers)
                ws.range((2, 1), (num_rows + 1, num_cols)).value = data
                print(f"[xlwings_writer] Daten nach Zeilen-Verschiebung geschrieben: {num_rows}x{num_cols}", file=sys.stderr)
            
            rows_reordered = True
        
        # SCHRITT 3: ZEILEN AUSBLENDEN
        _apply_hidden_rows_xlwings(ws, hidden_rows, len(data) if data else None)
        
        # SCHRITT 4: ZEILEN EINFÜGEN
        # Muss VOR dem Schreiben der Daten passieren, damit Excel genug Zeilen hat!
        # Nach dem Einfügen: Alle Daten als Block schreiben (schnell!)
        rows_inserted = False
        if inserted_rows:
            operations = inserted_rows.get('operations', [])
            if operations:
                # Sortiere nach Position aufsteigend
                inserted_offset = 0
                for op in sorted(operations, key=lambda x: x['position']):
                    pos = op['position'] + inserted_offset
                    count = op.get('count', 1)
                    excel_row_start = pos + 2  # +2 für Header und 1-basiert
                    excel_row_end = excel_row_start + count - 1
                    
                    print(f"[xlwings_writer] Füge {count} Zeile(n) bei Zeile {excel_row_start} ein", file=sys.stderr)
                    ws.range(f'{excel_row_start}:{excel_row_end}').insert(shift='down')
                    inserted_offset += count
                
                rows_inserted = True
                
                # Nach Einfügung: ALLE Daten als Block schreiben (schnell!)
                if data and len(data) > 0:
                    num_rows = len(data)
                    num_cols = len(data[0]) if data[0] else len(headers)
                    ws.range((2, 1), (num_rows + 1, num_cols)).value = data
                    print(f"[xlwings_writer] Daten nach Zeilen-Einfügung geschrieben: {num_rows}x{num_cols}", file=sys.stderr)
        
        # SCHRITT 5: ZEILEN MARKIEREN (ROW HIGHLIGHTS)
        if row_highlights and len(row_highlights) > 0:
            _apply_row_highlights_xlwings(ws, row_highlights, len(headers) if headers else ws.used_range.last_cell.column)
        
        # SCHRITT 5b: CLEARED ROW HIGHLIGHTS
        if cleared_row_highlights and len(cleared_row_highlights) > 0:
            for row_idx in cleared_row_highlights:
                excel_row = row_idx + 2
                num_cols = len(headers) if headers else ws.used_range.last_cell.column
                for col_idx in range(1, num_cols + 1):
                    cell = ws.range((excel_row, col_idx))
                    cell.color = None  # Farbe entfernen
        
        # SCHRITT 6: ZWISCHENSPEICHERN nach Zeilen-Operationen
        print(f"[xlwings_writer] Zwischenspeichern nach Zeilen-Operationen...", file=sys.stderr)
        wb.save()
        
        # =====================================================================
        # SPALTEN-OPERATIONEN DANACH
        # =====================================================================
        
        # SCHRITT 7: SPALTEN LÖSCHEN (von hinten nach vorne)
        # Verwende xlwings native Syntax (funktioniert auf macOS und Windows)
        if deleted_columns:
            for col_idx in sorted(deleted_columns, reverse=True):
                excel_col = col_idx + 1  # 1-basiert
                col_letter = _get_column_letter(excel_col)
                # Verwende Spalten-Range mit delete() - funktioniert plattformübergreifend
                ws.range(f'{col_letter}:{col_letter}').delete()
        
        # SCHRITT 8: SPALTEN VERSCHIEBEN (columnOrder)
        # WICHTIG: Wir müssen Spalten WIRKLICH in Excel verschieben, damit Formatierung erhalten bleibt!
        # columnOrder = [alt_idx_für_neue_pos_0, alt_idx_für_neue_pos_1, ...]
        # Beispiel: [2, 0, 1] bedeutet: Spalte C -> A, Spalte A -> B, Spalte B -> C
        column_order_applied = False
        if column_order and len(column_order) > 0:
            original_order = list(range(len(column_order)))
            if column_order != original_order:
                print(f"[xlwings_writer] columnOrder erkannt: {column_order}", file=sys.stderr)
                
                # Finde die Spalten-Paare die getauscht werden müssen
                # Wir müssen von der aktuellen Position zur Ziel-Position
                num_cols = len(column_order)
                
                # Berechne welche Verschiebungen nötig sind
                # columnOrder[i] = j bedeutet: Spalte j (0-basiert) soll an Position i
                # Wir müssen also Spalte j nach Position i verschieben
                
                # Erstelle inverse Mapping: wo ist jede Original-Spalte jetzt?
                current_order = list(range(num_cols))  # Am Anfang: [0, 1, 2, ...]
                
                for target_pos in range(num_cols):
                    source_original_idx = column_order[target_pos]
                    
                    # Wo ist diese Spalte aktuell?
                    current_pos = current_order.index(source_original_idx)
                    
                    if current_pos != target_pos:
                        # Muss verschoben werden
                        # Excel 1-basiert
                        source_col_excel = current_pos + 1
                        target_col_excel = target_pos + 1
                        
                        source_letter = _get_column_letter(source_col_excel)
                        target_letter = _get_column_letter(target_col_excel)
                        
                        print(f"[xlwings_writer] Verschiebe Spalte {source_letter} (Original {source_original_idx}) nach Position {target_letter}", file=sys.stderr)
                        
                        # Methode: Cut und Insert
                        source_range = ws.range(f'{source_letter}:{source_letter}')
                        
                        # Ermittle die letzte verwendete Zeile für begrenzte Bereiche (macOS-Kompatibilität)
                        last_row = ws.used_range.last_cell.row if ws.used_range else 1000
                        
                        if current_pos > target_pos:
                            # Nach links verschieben
                            # 1. Insert leere Spalte bei Ziel
                            ws.range(f'{target_letter}:{target_letter}').insert(shift='right')
                            # Dadurch verschiebt sich source um 1 nach rechts
                            new_source_col = source_col_excel + 1
                            new_source_letter = _get_column_letter(new_source_col)
                            # 2. Kopiere Quell-Spalte zur Ziel-Spalte (nur verwendeten Bereich)
                            source_rng = ws.range(f'{new_source_letter}1:{new_source_letter}{last_row}')
                            dest_rng = ws.range(f'{target_letter}1')
                            if platform.system() == 'Windows':
                                source_rng.api.Copy(Destination=dest_rng.api)
                            else:
                                # macOS: appscript - kopiere in Zelle, nicht ganze Spalte
                                source_rng.api.copy_range(destination=dest_rng.api)
                            # 3. Lösche die alte Spalte
                            ws.range(f'{new_source_letter}:{new_source_letter}').delete()
                        else:
                            # Nach rechts verschieben
                            # 1. Insert leere Spalte NACH dem Ziel (target+1)
                            after_target_letter = _get_column_letter(target_col_excel + 1)
                            ws.range(f'{after_target_letter}:{after_target_letter}').insert(shift='right')
                            # 2. Kopiere Quell-Spalte zur neuen Position (nur verwendeten Bereich)
                            source_rng = ws.range(f'{source_letter}1:{source_letter}{last_row}')
                            dest_rng = ws.range(f'{after_target_letter}1')
                            if platform.system() == 'Windows':
                                source_rng.api.Copy(Destination=dest_rng.api)
                            else:
                                # macOS: appscript - kopiere in Zelle, nicht ganze Spalte
                                source_rng.api.copy_range(destination=dest_rng.api)
                            # 3. Lösche die alte Spalte
                            ws.range(f'{source_letter}:{source_letter}').delete()
                        
                        # Update current_order
                        val = current_order.pop(current_pos)
                        current_order.insert(target_pos, val)
                
                column_order_applied = True
                print(f"[xlwings_writer] Spalten-Verschiebung abgeschlossen", file=sys.stderr)
        
        # SCHRITT 9: SPALTEN EINFÜGEN
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
            
            # Die Positionen vom Frontend sind FINALE Positionen (bereits korrigiert)
            for op in sorted(operations, key=lambda x: x['position']):
                pos = op['position']
                count = op.get('count', 1)
                op_headers = op.get('headers', [])
                excel_col = pos + 1  # 1-basiert
                
                for i in range(count):
                    insert_letter = _get_column_letter(excel_col + i)
                    
                    # Insert-Befehl: Fügt vor der angegebenen Spalte ein
                    # shift='right' verschiebt existierende Zellen nach rechts
                    ws.range(f'{insert_letter}:{insert_letter}').insert(shift='right')
                
                # Header setzen
                for i, header in enumerate(op_headers):
                    ws.range((1, excel_col + i)).value = header
        
        # SCHRITT 10: SPALTEN AUSBLENDEN
        _apply_hidden_columns_xlwings(ws, hidden_columns, len(headers) if headers else None)
        
        # =====================================================================
        # WEITERE OPERATIONEN
        # =====================================================================
        
        # SCHRITT 11: Filter-Handling mit rowMapping
        # NUR bei Filter (nicht bei Zeilen-Verschiebung!) - rowMapping enthält die Original-Indizes
        # Wir löschen die Zeilen die NICHT im Mapping sind (von hinten nach vorne)
        # Das erhält die Formatierung der verbleibenden Zeilen!
        # WICHTIG: Überspringe wenn bereits rows_reordered oder rows_inserted (Daten schon geschrieben)
        used_range = ws.used_range
        original_row_count = used_range.last_cell.row - 1 if used_range else 0  # Ohne Header
        new_row_count = len(data)
        
        if row_mapping and len(row_mapping) > 0 and not rows_reordered and not rows_inserted and not column_order_applied:
            # Filter aktiv: Lösche Zeilen die nicht im Mapping sind
            # rowMapping = [originalIndex1, originalIndex2, ...] (0-basiert)
            rows_to_keep = set(row_mapping)
            rows_to_delete = []
            
            for orig_idx in range(original_row_count):
                if orig_idx not in rows_to_keep:
                    rows_to_delete.append(orig_idx)
            
            if rows_to_delete:
                print(f"[xlwings_writer] Filter: Lösche {len(rows_to_delete)} Zeilen (behalte {len(rows_to_keep)})", file=sys.stderr)
                # OPTIMIERUNG: Finde zusammenhängende Bereiche und lösche diese auf einmal
                # Sortiere absteigend damit Indizes beim Löschen stimmen
                sorted_rows = sorted(rows_to_delete, reverse=True)
                
                # Finde zusammenhängende Bereiche
                ranges_to_delete = []
                if sorted_rows:
                    range_end = sorted_rows[0]
                    range_start = sorted_rows[0]
                    
                    for i in range(1, len(sorted_rows)):
                        if sorted_rows[i] == range_start - 1:
                            # Zusammenhängend
                            range_start = sorted_rows[i]
                        else:
                            # Lücke - speichere aktuellen Bereich
                            ranges_to_delete.append((range_start, range_end))
                            range_end = sorted_rows[i]
                            range_start = sorted_rows[i]
                    
                    # Letzten Bereich hinzufügen
                    ranges_to_delete.append((range_start, range_end))
                
                print(f"[xlwings_writer] Lösche {len(ranges_to_delete)} zusammenhängende Bereiche", file=sys.stderr)
                
                # Lösche Bereiche (bereits absteigend sortiert)
                for range_start, range_end in ranges_to_delete:
                    excel_start = range_start + 2  # +2 für Header und 1-basiert
                    excel_end = range_end + 2
                    ws.range(f'{excel_start}:{excel_end}').delete()
            
            # KEINE Daten neu schreiben - die Zeilen sind schon korrekt!
            
        elif full_rewrite and new_row_count < original_row_count:
            # Kein rowMapping aber weniger Zeilen - einfach am Ende löschen
            rows_to_delete = original_row_count - new_row_count
            first_row_to_delete = new_row_count + 2  # +2 für Header (1-basiert)
            last_row_to_delete = original_row_count + 1  # +1 für Header (1-basiert)
            print(f"[xlwings_writer] Lösche Zeilen {first_row_to_delete} bis {last_row_to_delete} ({rows_to_delete} Zeilen)", file=sys.stderr)
            ws.range(f'{first_row_to_delete}:{last_row_to_delete}').delete()
        
        # SCHRITT 12: ZELL-EDITS (geänderte Zellen)
        # NUR wenn nicht bereits durch Block-Write geschrieben (columnOrder, rowsInserted, rowsReordered)
        # OPTIMIERUNG: Gruppiere nach Spalten und schreibe spaltenweise statt zellweise
        if edited_cells and not column_order_applied and not rows_inserted and not rows_reordered:
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
