#!/usr/bin/env python3
"""
Excel Live Session - Persistente Excel-Verbindung für Live-Editing

Statt alle Operationen am Ende auf einmal auszuführen,
bleibt Excel im Hintergrund offen und jede Operation wird SOFORT ausgeführt.

Vorteile:
- Keine Index-Konflikte bei kombinierten Operationen
- Immer aktueller Zustand
- Schnellere Reaktion (Excel ist bereits offen)
- Formatierung bleibt IMMER erhalten

Kommunikation: JSON über stdin/stdout
"""

import json
import sys
import os
import platform
import time
from typing import Optional, Dict, Any

# Für embedded Python auf Windows: pywin32 DLLs finden
if platform.system() == 'Windows':
    pywin32_dll = os.path.join(sys.prefix, 'Lib', 'site-packages', 'pywin32_system32')
    if os.path.exists(pywin32_dll):
        os.environ['PATH'] = pywin32_dll + os.pathsep + os.environ.get('PATH', '')
    python_dir = os.path.dirname(sys.executable)
    if os.path.exists(os.path.join(python_dir, 'pythoncom311.dll')):
        os.environ['PATH'] = python_dir + os.pathsep + os.environ.get('PATH', '')
    win32_dir = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32')
    if os.path.exists(win32_dir):
        sys.path.insert(0, win32_dir)
    win32_lib = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32', 'lib')
    if os.path.exists(win32_lib):
        sys.path.insert(0, win32_lib)

try:
    import xlwings as xw
except ImportError as e:
    print(json.dumps({"success": False, "error": f"xlwings import failed: {e}"}), flush=True)
    sys.exit(1)


class ExcelLiveSession:
    """Persistente Excel-Session für Live-Editing"""
    
    def __init__(self):
        self.app: Optional[xw.App] = None
        self.workbook: Optional[xw.Book] = None
        self.worksheet: Optional[xw.Sheet] = None
        self.file_path: Optional[str] = None
        self.sheet_name: Optional[str] = None
        self._is_running = True
    
    def _log(self, message: str):
        """Logging zu stderr (nicht stdout, das ist für JSON)"""
        print(f"[LiveSession] {message}", file=sys.stderr, flush=True)
    
    def _respond(self, data: Dict[str, Any]):
        """Sendet JSON-Antwort an stdout"""
        print(json.dumps(data), flush=True)
    
    def _get_column_letter(self, col_idx: int) -> str:
        """Konvertiert Spalten-Index (1-basiert) zu Buchstaben"""
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result
    
    def _hide_excel(self):
        """Versteckt Excel"""
        import subprocess
        if platform.system() == 'Darwin':
            try:
                subprocess.run(['osascript', '-e', 
                    'tell application "System Events" to set visible of process "Microsoft Excel" to false'], 
                    capture_output=True, timeout=2)
            except:
                pass
        elif platform.system() == 'Windows' and self.app:
            try:
                self.app.visible = False
            except:
                pass
    
    # =========================================================================
    # SESSION-MANAGEMENT
    # =========================================================================
    
    def open_file(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """Öffnet eine Excel-Datei und hält sie offen"""
        try:
            self._log(f"Öffne Datei: {file_path}, Sheet: {sheet_name}")
            
            # Falls bereits eine Datei offen ist, schließen
            if self.workbook:
                try:
                    self.workbook.close()
                except:
                    pass
            
            # Neue Excel-App starten falls nötig
            if not self.app:
                self._log("Starte Excel-App...")
                self.app = xw.App(visible=False, add_book=False)
                self.app.display_alerts = False
                self.app.screen_updating = False
            
            self._hide_excel()
            
            # Workbook öffnen
            self.workbook = self.app.books.open(file_path)
            self.file_path = file_path
            
            # Sheet finden
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            self.worksheet = self.workbook.sheets[sheet_name]
            self.sheet_name = sheet_name
            
            self._log(f"Datei geöffnet, Sheet: {sheet_name}")
            return {'success': True, 'sheets': sheet_names}
            
        except Exception as e:
            self._log(f"Fehler beim Öffnen: {e}")
            return {'success': False, 'error': str(e)}
    
    def save_file(self, output_path: Optional[str] = None) -> Dict[str, Any]:
        """Speichert die Datei (optional unter neuem Namen)"""
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if output_path and output_path != self.file_path:
                self._log(f"Speichere unter: {output_path}")
                self.workbook.save(output_path)
                self.file_path = output_path
            else:
                self._log("Speichere...")
                self.workbook.save()
            
            return {'success': True, 'outputPath': self.file_path}
            
        except Exception as e:
            self._log(f"Fehler beim Speichern: {e}")
            return {'success': False, 'error': str(e)}
    
    def close_session(self) -> Dict[str, Any]:
        """Schließt die Session"""
        try:
            self._log("Schließe Session...")
            
            if self.workbook:
                try:
                    self.workbook.close()
                except:
                    pass
                self.workbook = None
            
            if self.app:
                try:
                    self.app.quit()
                except:
                    pass
                self.app = None
            
            self.worksheet = None
            self.file_path = None
            self.sheet_name = None
            
            return {'success': True}
            
        except Exception as e:
            self._log(f"Fehler beim Schließen: {e}")
            return {'success': False, 'error': str(e)}
    
    # =========================================================================
    # ZEILEN-OPERATIONEN
    # =========================================================================
    
    def delete_row(self, row_index: int) -> Dict[str, Any]:
        """Löscht eine Zeile (0-basierter Index, ohne Header)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row = row_index + 2  # +2 für Header (1-basiert)
            self._log(f"Lösche Zeile {excel_row}")
            self.worksheet.range(f'{excel_row}:{excel_row}').delete()
            
            return {'success': True, 'deletedRow': row_index}
            
        except Exception as e:
            self._log(f"Fehler beim Löschen der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def insert_row(self, row_index: int, count: int = 1) -> Dict[str, Any]:
        """Fügt leere Zeilen ein (0-basierter Index, ohne Header)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row_start = row_index + 2
            excel_row_end = excel_row_start + count - 1
            
            self._log(f"Füge {count} Zeile(n) bei {excel_row_start} ein")
            self.worksheet.range(f'{excel_row_start}:{excel_row_end}').insert(shift='down')
            
            return {'success': True, 'insertedAt': row_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Einfügen der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def move_row(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """Verschiebt eine Zeile von from_index nach to_index"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_from = from_index + 2
            excel_to = to_index + 2
            
            self._log(f"Verschiebe Zeile {excel_from} nach {excel_to}")
            
            # Ermittle Spaltenanzahl
            last_col = self.worksheet.used_range.last_cell.column if self.worksheet.used_range else 10
            last_col_letter = self._get_column_letter(last_col)
            
            if from_index > to_index:
                # Nach oben verschieben
                # 1. Insert leere Zeile bei Ziel
                self.worksheet.range(f'{excel_to}:{excel_to}').insert(shift='down')
                # 2. Kopiere Quell-Zeile (jetzt +1) zur Ziel-Zeile
                new_from = excel_from + 1
                source_rng = self.worksheet.range(f'A{new_from}:{last_col_letter}{new_from}')
                dest_rng = self.worksheet.range(f'A{excel_to}')
                if platform.system() == 'Windows':
                    source_rng.api.Copy(Destination=dest_rng.api)
                else:
                    source_rng.api.copy_range(destination=dest_rng.api)
                # 3. Lösche alte Zeile
                self.worksheet.range(f'{new_from}:{new_from}').delete()
            else:
                # Nach unten verschieben
                # 1. Insert leere Zeile nach Ziel
                after_target = excel_to + 1
                self.worksheet.range(f'{after_target}:{after_target}').insert(shift='down')
                # 2. Kopiere Quell-Zeile zur neuen Position
                source_rng = self.worksheet.range(f'A{excel_from}:{last_col_letter}{excel_from}')
                dest_rng = self.worksheet.range(f'A{after_target}')
                if platform.system() == 'Windows':
                    source_rng.api.Copy(Destination=dest_rng.api)
                else:
                    source_rng.api.copy_range(destination=dest_rng.api)
                # 3. Lösche alte Zeile
                self.worksheet.range(f'{excel_from}:{excel_from}').delete()
            
            return {'success': True, 'movedFrom': from_index, 'movedTo': to_index}
            
        except Exception as e:
            self._log(f"Fehler beim Verschieben der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def hide_row(self, row_index: int, hidden: bool = True) -> Dict[str, Any]:
        """Versteckt oder zeigt eine Zeile"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row = row_index + 2
            row_range = self.worksheet.range(f'{excel_row}:{excel_row}')
            
            if hidden:
                row_range.row_height = 0
            else:
                row_range.row_height = None  # Standard-Höhe
            
            self._log(f"Zeile {excel_row} {'versteckt' if hidden else 'angezeigt'}")
            return {'success': True, 'row': row_index, 'hidden': hidden}
            
        except Exception as e:
            self._log(f"Fehler beim Verstecken der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def highlight_row(self, row_index: int, color: Optional[str] = None) -> Dict[str, Any]:
        """Markiert eine Zeile mit Farbe (None = Farbe entfernen)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row = row_index + 2
            last_col = self.worksheet.used_range.last_cell.column if self.worksheet.used_range else 10
            last_col_letter = self._get_column_letter(last_col)
            
            row_range = self.worksheet.range(f'A{excel_row}:{last_col_letter}{excel_row}')
            
            if color is None:
                row_range.color = None
                self._log(f"Zeile {excel_row} Farbe entfernt")
            else:
                # Farben-Mapping
                colors = {
                    'green': (144, 238, 144),
                    'yellow': (255, 255, 0),
                    'orange': (255, 165, 0),
                    'red': (255, 107, 107),
                    'blue': (135, 206, 235),
                    'purple': (221, 160, 221)
                }
                rgb = colors.get(color, (255, 255, 0))
                row_range.color = rgb
                self._log(f"Zeile {excel_row} markiert mit {color}")
            
            return {'success': True, 'row': row_index, 'color': color}
            
        except Exception as e:
            self._log(f"Fehler beim Markieren der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    # =========================================================================
    # SPALTEN-OPERATIONEN
    # =========================================================================
    
    def delete_column(self, col_index: int) -> Dict[str, Any]:
        """Löscht eine Spalte (0-basierter Index)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_col = col_index + 1
            col_letter = self._get_column_letter(excel_col)
            
            self._log(f"Lösche Spalte {col_letter}")
            self.worksheet.range(f'{col_letter}:{col_letter}').delete()
            
            return {'success': True, 'deletedColumn': col_index}
            
        except Exception as e:
            self._log(f"Fehler beim Löschen der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    def insert_column(self, col_index: int, count: int = 1, headers: list = None) -> Dict[str, Any]:
        """Fügt leere Spalten ein"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_col = col_index + 1
            
            for i in range(count):
                insert_letter = self._get_column_letter(excel_col + i)
                self._log(f"Füge Spalte {insert_letter} ein")
                self.worksheet.range(f'{insert_letter}:{insert_letter}').insert(shift='right')
            
            # Header setzen falls vorhanden
            if headers:
                for i, header in enumerate(headers):
                    self.worksheet.range((1, excel_col + i)).value = header
            
            return {'success': True, 'insertedAt': col_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Einfügen der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    def move_column(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """Verschiebt eine Spalte"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_from = from_index + 1
            excel_to = to_index + 1
            
            source_letter = self._get_column_letter(excel_from)
            target_letter = self._get_column_letter(excel_to)
            
            self._log(f"Verschiebe Spalte {source_letter} nach {target_letter}")
            
            last_row = self.worksheet.used_range.last_cell.row if self.worksheet.used_range else 1000
            
            if from_index > to_index:
                # Nach links verschieben
                self.worksheet.range(f'{target_letter}:{target_letter}').insert(shift='right')
                new_source_col = excel_from + 1
                new_source_letter = self._get_column_letter(new_source_col)
                source_rng = self.worksheet.range(f'{new_source_letter}1:{new_source_letter}{last_row}')
                dest_rng = self.worksheet.range(f'{target_letter}1')
                if platform.system() == 'Windows':
                    source_rng.api.Copy(Destination=dest_rng.api)
                else:
                    source_rng.api.copy_range(destination=dest_rng.api)
                self.worksheet.range(f'{new_source_letter}:{new_source_letter}').delete()
            else:
                # Nach rechts verschieben
                after_target_letter = self._get_column_letter(excel_to + 1)
                self.worksheet.range(f'{after_target_letter}:{after_target_letter}').insert(shift='right')
                source_rng = self.worksheet.range(f'{source_letter}1:{source_letter}{last_row}')
                dest_rng = self.worksheet.range(f'{after_target_letter}1')
                if platform.system() == 'Windows':
                    source_rng.api.Copy(Destination=dest_rng.api)
                else:
                    source_rng.api.copy_range(destination=dest_rng.api)
                self.worksheet.range(f'{source_letter}:{source_letter}').delete()
            
            return {'success': True, 'movedFrom': from_index, 'movedTo': to_index}
            
        except Exception as e:
            self._log(f"Fehler beim Verschieben der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    def hide_column(self, col_index: int, hidden: bool = True) -> Dict[str, Any]:
        """Versteckt oder zeigt eine Spalte"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_col = col_index + 1
            col_letter = self._get_column_letter(excel_col)
            col_range = self.worksheet.range(f'{col_letter}:{col_letter}')
            
            if hidden:
                col_range.column_width = 0
            else:
                col_range.column_width = None  # Standard-Breite
            
            self._log(f"Spalte {col_letter} {'versteckt' if hidden else 'angezeigt'}")
            return {'success': True, 'column': col_index, 'hidden': hidden}
            
        except Exception as e:
            self._log(f"Fehler beim Verstecken der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    # =========================================================================
    # ZELL-OPERATIONEN
    # =========================================================================
    
    def set_cell_value(self, row_index: int, col_index: int, value: Any) -> Dict[str, Any]:
        """Setzt den Wert einer Zelle"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row = row_index + 2
            excel_col = col_index + 1
            
            self.worksheet.range((excel_row, excel_col)).value = value
            
            return {'success': True, 'row': row_index, 'col': col_index, 'value': value}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen des Zellwerts: {e}")
            return {'success': False, 'error': str(e)}
    
    def get_data(self) -> Dict[str, Any]:
        """Liest alle Daten aus dem aktuellen Sheet"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            used_range = self.worksheet.used_range
            if not used_range:
                return {'success': True, 'headers': [], 'data': []}
            
            all_data = used_range.value
            if not all_data:
                return {'success': True, 'headers': [], 'data': []}
            
            # Erste Zeile = Header
            headers = all_data[0] if isinstance(all_data[0], list) else [all_data[0]]
            data = all_data[1:] if len(all_data) > 1 else []
            
            return {'success': True, 'headers': headers, 'data': data}
            
        except Exception as e:
            self._log(f"Fehler beim Lesen der Daten: {e}")
            return {'success': False, 'error': str(e)}
    
    # =========================================================================
    # MAIN LOOP
    # =========================================================================
    
    def handle_command(self, cmd: Dict[str, Any]) -> Dict[str, Any]:
        """Verarbeitet einen Befehl"""
        action = cmd.get('action', '')
        
        handlers = {
            'open': lambda: self.open_file(cmd.get('filePath'), cmd.get('sheetName')),
            'save': lambda: self.save_file(cmd.get('outputPath')),
            'close': lambda: self.close_session(),
            'getData': lambda: self.get_data(),
            
            # Zeilen
            'deleteRow': lambda: self.delete_row(cmd.get('rowIndex')),
            'insertRow': lambda: self.insert_row(cmd.get('rowIndex'), cmd.get('count', 1)),
            'moveRow': lambda: self.move_row(cmd.get('fromIndex'), cmd.get('toIndex')),
            'hideRow': lambda: self.hide_row(cmd.get('rowIndex'), cmd.get('hidden', True)),
            'highlightRow': lambda: self.highlight_row(cmd.get('rowIndex'), cmd.get('color')),
            
            # Spalten
            'deleteColumn': lambda: self.delete_column(cmd.get('colIndex')),
            'insertColumn': lambda: self.insert_column(cmd.get('colIndex'), cmd.get('count', 1), cmd.get('headers')),
            'moveColumn': lambda: self.move_column(cmd.get('fromIndex'), cmd.get('toIndex')),
            'hideColumn': lambda: self.hide_column(cmd.get('colIndex'), cmd.get('hidden', True)),
            
            # Zellen
            'setCellValue': lambda: self.set_cell_value(cmd.get('rowIndex'), cmd.get('colIndex'), cmd.get('value')),
            
            # Session
            'ping': lambda: {'success': True, 'pong': True},
            'quit': lambda: self._quit(),
        }
        
        handler = handlers.get(action)
        if handler:
            return handler()
        else:
            return {'success': False, 'error': f'Unbekannte Aktion: {action}'}
    
    def _quit(self) -> Dict[str, Any]:
        """Beendet die Session"""
        self._is_running = False
        self.close_session()
        return {'success': True, 'message': 'Session beendet'}
    
    def run(self):
        """Hauptschleife - liest JSON-Befehle von stdin"""
        self._log("Live Session gestartet, warte auf Befehle...")
        
        while self._is_running:
            try:
                line = sys.stdin.readline()
                if not line:
                    self._log("EOF erreicht, beende...")
                    break
                
                line = line.strip()
                if not line:
                    continue
                
                try:
                    cmd = json.loads(line)
                except json.JSONDecodeError as e:
                    self._respond({'success': False, 'error': f'Ungültiges JSON: {e}'})
                    continue
                
                result = self.handle_command(cmd)
                self._respond(result)
                
            except KeyboardInterrupt:
                self._log("Interrupted, beende...")
                break
            except Exception as e:
                self._log(f"Fehler: {e}")
                self._respond({'success': False, 'error': str(e)})
        
        self.close_session()
        self._log("Session beendet")


def main():
    session = ExcelLiveSession()
    session.run()


if __name__ == '__main__':
    main()
