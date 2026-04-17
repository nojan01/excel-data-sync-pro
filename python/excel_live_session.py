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

import atexit
import json
import sys
import os
import platform
import time
import shutil
from datetime import datetime, date, time as dtime, timedelta
from typing import Optional, Dict, Any, List

# WICHTIG: UTF-8 für stdin/stdout erzwingen
# Auf Windows verwendet embedded Python sonst die System-Codepage (CP1252),
# was bei Umlauten in Sheet-Namen (ß, ä, ö, ü) zu Mojibake führt:
# "Große Tabelle" → "GroÃŸe Tabelle" → Sheet nicht gefunden!
if hasattr(sys.stdin, 'reconfigure'):
    sys.stdin.reconfigure(encoding='utf-8')
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# Für embedded Python auf Windows: pywin32 DLLs finden
if platform.system() == 'Windows':
    # Bei embedded Python: sys.prefix kann falsch sein, daher vom executable ausgehen
    python_dir = os.path.dirname(sys.executable)
    
    # pywin32_system32 DLLs (im site-packages)
    pywin32_dll = os.path.join(python_dir, 'Lib', 'site-packages', 'pywin32_system32')
    if os.path.exists(pywin32_dll):
        os.environ['PATH'] = pywin32_dll + os.pathsep + os.environ.get('PATH', '')
    # Fallback: über sys.prefix
    pywin32_dll_alt = os.path.join(sys.prefix, 'Lib', 'site-packages', 'pywin32_system32')
    if os.path.exists(pywin32_dll_alt) and pywin32_dll_alt != pywin32_dll:
        os.environ['PATH'] = pywin32_dll_alt + os.pathsep + os.environ.get('PATH', '')
    
    # DLLs direkt im Python-Verzeichnis (pythoncom311.dll, pywintypes311.dll)
    if os.path.exists(os.path.join(python_dir, 'pythoncom311.dll')):
        os.environ['PATH'] = python_dir + os.pathsep + os.environ.get('PATH', '')
    
    # win32 Module zum sys.path hinzufügen
    win32_dir = os.path.join(python_dir, 'Lib', 'site-packages', 'win32')
    if os.path.exists(win32_dir):
        sys.path.insert(0, win32_dir)
    win32_lib = os.path.join(python_dir, 'Lib', 'site-packages', 'win32', 'lib')
    if os.path.exists(win32_lib):
        sys.path.insert(0, win32_lib)
    
    # Site-packages zum sys.path (für embedded Python)
    site_packages = os.path.join(python_dir, 'Lib', 'site-packages')
    if os.path.exists(site_packages) and site_packages not in sys.path:
        sys.path.insert(0, site_packages)

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
        self.file_password: Optional[str] = None  # Passwort für geschützte Dateien
        self._is_running = True
        self._active_filter_fields: List[int] = []  # Gespeicherte Filter-Felder (1-basiert)
        
        # Undo-Stack
        self.undo_stack: List[Dict] = []
        self._undo_in_progress = False
        self.MAX_UNDO = 10
        self._last_undo_snapshot_time: float = 0  # Für Drosselung bei schnellen Edits
        
        # Undo-Verzeichnis: Windows Temp (lokal, immer schreibbar)
        import tempfile
        self._undo_dir = os.path.join(tempfile.gettempdir(), 'ExcelDataSyncPro_undo')
        try:
            os.makedirs(self._undo_dir, exist_ok=True)
            self._cleanup_leftover_undo_files()
        except Exception:
            pass
        
        # Recovery-System
        self.backup_path: Optional[str] = None
        self.journal_path: Optional[str] = None
        self.journal_entries: List[Dict] = []
        self.last_auto_save: float = 0
        self.auto_save_interval: int = 120  # 2 Minuten
    
    def _log(self, message: str):
        """Logging zu stderr (nicht stdout, das ist für JSON)"""
        print(f"[LiveSession] {message}", file=sys.stderr, flush=True)
    
    def _respond(self, data: Dict[str, Any]):
        """Sendet JSON-Antwort an stdout"""
        print(json.dumps(data, default=self._json_serialize), flush=True)
    
    @staticmethod
    def _json_serialize(obj):
        """Custom JSON serializer für datetime, date, time, bytes etc."""
        if isinstance(obj, datetime):
            # Datum+Uhrzeit: Nur Datum wenn Uhrzeit 00:00:00
            if obj.hour == 0 and obj.minute == 0 and obj.second == 0:
                return obj.strftime('%d.%m.%Y')
            return obj.strftime('%d.%m.%Y %H:%M:%S')
        if isinstance(obj, date):
            return obj.strftime('%d.%m.%Y')
        if isinstance(obj, dtime):
            return obj.strftime('%H:%M:%S')
        if isinstance(obj, bytes):
            return None
        if isinstance(obj, float):
            import math
            if math.isnan(obj) or math.isinf(obj):
                return None
        return str(obj)
    
    # =========================================================================
    # RECOVERY-SYSTEM
    # =========================================================================
    
    def _get_recovery_dir(self) -> str:
        """Gibt das Recovery-Verzeichnis zurück"""
        base = os.path.join(os.environ.get('APPDATA', ''), 'ExcelDataSyncPro')
        
        recovery_dir = os.path.join(base, 'recovery')
        os.makedirs(recovery_dir, exist_ok=True)
        return recovery_dir
    
    def _create_backup(self, file_path: str) -> Optional[str]:
        """Erstellt eine Backup-Kopie der Originaldatei"""
        try:
            recovery_dir = self._get_recovery_dir()
            file_name = os.path.basename(file_path)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_name = f"{os.path.splitext(file_name)[0]}_{timestamp}.bak.xlsx"
            backup_path = os.path.join(recovery_dir, backup_name)
            
            shutil.copy2(file_path, backup_path)
            self._log(f"Backup erstellt: {backup_path}")
            return backup_path
        except Exception as e:
            self._log(f"Backup-Fehler: {e}")
            return None
    
    def _init_journal(self, file_path: str) -> Optional[str]:
        """Initialisiert das Änderungs-Journal"""
        try:
            recovery_dir = self._get_recovery_dir()
            file_name = os.path.basename(file_path)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            journal_name = f"{os.path.splitext(file_name)[0]}_{timestamp}.journal.json"
            journal_path = os.path.join(recovery_dir, journal_name)
            
            # Journal-Header schreiben
            journal_data = {
                'version': '1.0',
                'originalFile': file_path,
                'created': datetime.now().isoformat(),
                'entries': []
            }
            with open(journal_path, 'w', encoding='utf-8') as f:
                json.dump(journal_data, f, ensure_ascii=False)
            
            self._log(f"Journal initialisiert: {journal_path}")
            return journal_path
        except Exception as e:
            self._log(f"Journal-Init-Fehler: {e}")
            return None
    
    def _journal_add(self, operation: str, details: Dict[str, Any]):
        """Fügt einen Eintrag zum Journal hinzu (schnell, append-only)"""
        if not self.journal_path:
            return
        
        try:
            entry = {
                'timestamp': datetime.now().isoformat(),
                'operation': operation,
                'sheet': self.sheet_name,
                **details
            }
            self.journal_entries.append(entry)
            
            # Alle 10 Einträge: Journal auf Disk schreiben
            if len(self.journal_entries) >= 10:
                self._flush_journal()
        except Exception as e:
            self._log(f"Journal-Add-Fehler: {e}")
    
    def _flush_journal(self):
        """Schreibt ausstehende Journal-Einträge auf Disk"""
        if not self.journal_path or not self.journal_entries:
            return
        
        try:
            # Journal lesen, Einträge anhängen, speichern
            with open(self.journal_path, 'r', encoding='utf-8') as f:
                journal_data = json.load(f)
            
            journal_data['entries'].extend(self.journal_entries)
            journal_data['lastModified'] = datetime.now().isoformat()
            
            with open(self.journal_path, 'w', encoding='utf-8') as f:
                json.dump(journal_data, f, ensure_ascii=False, indent=2)
            
            self.journal_entries.clear()
        except Exception as e:
            self._log(f"Journal-Flush-Fehler: {e}")
    
    def _check_auto_save(self):
        """Auto-Save deaktiviert – User speichert bewusst über Save-Button.
        Methode bleibt als No-Op damit Aufrufe nicht entfernt werden müssen."""
        pass
    
    def _cleanup_recovery(self, success: bool = True):
        """Bereinigt Recovery-Dateien nach erfolgreichem Speichern"""
        if success:
            # Bei Erfolg: Backup und Journal löschen
            try:
                if self.backup_path and os.path.exists(self.backup_path):
                    os.remove(self.backup_path)
                    self._log(f"Backup gelöscht: {self.backup_path}")
                if self.journal_path and os.path.exists(self.journal_path):
                    os.remove(self.journal_path)
                    self._log(f"Journal gelöscht: {self.journal_path}")
            except Exception as e:
                self._log(f"Recovery-Cleanup-Fehler: {e}")
        
        self.backup_path = None
        self.journal_path = None
        self.journal_entries.clear()
    
    def get_recovery_files(self) -> Dict[str, Any]:
        """Gibt verfügbare Recovery-Dateien zurück"""
        try:
            recovery_dir = self._get_recovery_dir()
            files = []
            
            for f in os.listdir(recovery_dir):
                full_path = os.path.join(recovery_dir, f)
                stat = os.stat(full_path)
                age_hours = (time.time() - stat.st_mtime) / 3600
                
                if f.endswith('.bak.xlsx'):
                    files.append({
                        'type': 'backup',
                        'path': full_path,
                        'name': f,
                        'size': stat.st_size,
                        'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                        'ageHours': round(age_hours, 1)
                    })
                elif f.endswith('.journal.json'):
                    # Journal-Details lesen
                    with open(full_path, 'r', encoding='utf-8') as jf:
                        journal_data = json.load(jf)
                    files.append({
                        'type': 'journal',
                        'path': full_path,
                        'name': f,
                        'originalFile': journal_data.get('originalFile'),
                        'entryCount': len(journal_data.get('entries', [])),
                        'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                        'ageHours': round(age_hours, 1)
                    })
            
            # Nach Alter sortieren (neueste zuerst)
            files.sort(key=lambda x: x['ageHours'])
            
            # Alte Dateien (>24h) automatisch löschen
            for f in files[:]:
                if f['ageHours'] > 24:
                    try:
                        os.remove(f['path'])
                        files.remove(f)
                        self._log(f"Alte Recovery-Datei gelöscht: {f['name']}")
                    except:
                        pass
            
            return {'success': True, 'files': files}
        except Exception as e:
            return {'success': False, 'error': str(e), 'files': []}
    
    def delete_recovery_file(self, file_path: str) -> Dict[str, Any]:
        """Löscht eine Recovery-Datei"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return {'success': True}
            return {'success': False, 'error': 'Datei nicht gefunden'}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def _get_column_letter(self, col_idx: int) -> str:
        """Konvertiert Spalten-Index (1-basiert) zu Buchstaben"""
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result
    
    def _hide_excel(self):
        """Versteckt Excel"""
        if self.app:
            try:
                self.app.visible = False
            except:
                pass
    
    def set_visible(self, visible: bool = True) -> Dict[str, Any]:
        """Zeigt oder versteckt das Excel-Fenster.
        Wenn sichtbar: Interactive=False → Read-Only-Vorschau, kein Schließen möglich.
        Wenn versteckt: Interactive=True → normale Steuerung durch xlwings."""
        try:
            if not self.app:
                self._log("Keine Excel-App aktiv")
                return {'success': False, 'error': 'Keine Excel-App aktiv'}
            
            # xlwings visible-Eigenschaft verwenden
            self.app.visible = visible
            
            # Interactive-Schutz: Wenn sichtbar → Benutzerinteraktion sperren
            # Verhindert versehentliches Schließen/Bearbeiten von Excel
            try:
                if platform.system() == 'Windows':
                    self.app.api.Interactive = not visible
                else:
                    self.app.api.interactive.set(not visible)
            except Exception as int_err:
                self._log(f"Interactive-Flag Fehler (ignoriert): {int_err}")
            
            self._log(f"Excel Sichtbarkeit gesetzt: {visible}, Interactive={not visible}")
            
            return {'success': True, 'visible': visible}
        except Exception as e:
            self._log(f"Fehler bei set_visible: {e}")
            return {'success': False, 'error': str(e)}
    
    def _force_screen_refresh(self):
        """Erzwingt einen Screen-Refresh in Excel"""
        try:
            if not self.app:
                return
            
            is_visible = False
            try:
                is_visible = self.app.visible
            except Exception:
                pass
            
            app = self.app.api
            
            # Alle Blocker temporär aufheben
            try:
                if is_visible:
                    app.Interactive = True
                app.EnableEvents = True
                app.ScreenUpdating = True
            except Exception:
                pass
            
            try:
                if self.workbook:
                    self.workbook.activate()
                if self.worksheet:
                    try:
                        app.Goto(self.worksheet.api.Range("A1"))
                    except Exception:
                        self.worksheet.activate()
                app.ScreenUpdating = False
                app.ScreenUpdating = True
                self.app.calculate()
            except Exception as win_err:
                self._log(f"_force_screen_refresh Fehler: {win_err}")
            
            # Blocker wiederherstellen
            try:
                app.EnableEvents = False
                if is_visible:
                    app.Interactive = False
            except Exception:
                pass
                
        except Exception as e:
            self._log(f"Fehler bei screen refresh: {e}")
    
    def check_alive(self) -> Dict[str, Any]:
        """Prüft ob Excel und das Workbook noch aktiv sind"""
        try:
            # Prüfe ob App noch existiert
            if not self.app:
                return {'success': True, 'alive': False, 'reason': 'no_app'}
            
            # Prüfe ob Workbook noch offen ist
            if not self.workbook:
                return {'success': True, 'alive': False, 'reason': 'no_workbook'}
            
            # Versuche auf das Workbook zuzugreifen (wirft Exception wenn geschlossen)
            try:
                _ = self.workbook.name
                _ = self.app.visible  # Teste auch App-Zugriff
            except Exception:
                return {'success': True, 'alive': False, 'reason': 'workbook_closed'}
            
            return {'success': True, 'alive': True}
        except Exception as e:
            self._log(f"Fehler bei check_alive: {e}")
            return {'success': True, 'alive': False, 'reason': str(e)}
    
    # =========================================================================
    # SESSION-MANAGEMENT
    # =========================================================================
    
    def init_app(self) -> Dict[str, Any]:
        """Startet die Excel-App ohne eine Datei zu öffnen.
        Wird beim Programmstart aufgerufen, damit Excel bereit ist."""
        try:
            if self.app:
                self._log("Excel-App bereits gestartet")
                return {'success': True, 'already_running': True}
            
            self._log("Starte Excel-App (ohne Datei)...")
            self.app = xw.App(visible=False, add_book=False)
            self.app.display_alerts = False
            self.app.screen_updating = True
            # Berechnung auf manuell — verhindert Pivot-Cache-Neuberechnung beim Öffnen
            try:
                self.app.calculation = 'manual'
            except:
                pass
            self._log("Excel-App bereit (calculation=manual)")
            return {'success': True}
        except Exception as e:
            self._log(f"Fehler beim Starten der Excel-App: {e}")
            return {'success': False, 'error': str(e)}
    
    def open_file(self, file_path: str, sheet_name: str, password: Optional[str] = None) -> Dict[str, Any]:
        """Öffnet eine Excel-Datei und hält sie offen
        
        Args:
            file_path: Pfad zur Excel-Datei
            sheet_name: Name des zu öffnenden Sheets
            password: Optionales Passwort für geschützte Dateien
        """
        try:
            import time as _time
            _t0 = _time.time()
            self._log(f"Öffne Datei: {file_path}, Sheet: {sheet_name}, Password: {'***' if password else 'None'}")
            
            # Passwort speichern für spätere Verwendung
            self.file_password = password
            
            # Falls bereits eine Datei offen ist, schließen
            if self.workbook:
                try:
                    self.workbook.close()
                except:
                    pass
            
            # Neue Excel-App starten falls nötig
            if not self.app:
                self._log("Starte Excel-App...")
                self.app = xw.App(visible=False, add_book=False)  # visible=False - Excel startet versteckt
                self.app.display_alerts = False
                self.app.screen_updating = True
                self._log(f"Excel-App gestartet in {_time.time() - _t0:.1f}s")
            
            # WICHTIG: Berechnung auf "manuell" setzen BEVOR die Datei geöffnet wird.
            # Dateien mit Pivot-Tabellen lösen beim Öffnen eine automatische
            # Neuberechnung des Pivot-Cache aus — bei großen Dateien (10000+ Zeilen)
            # dauert das auf ARM64 so lange, dass der Timeout greift.
            try:
                self.app.calculation = 'manual'
                self._log("Berechnung auf 'manual' gesetzt (Pivot-Cache-Schutz)")
            except Exception as calc_err:
                self._log(f"Konnte Berechnung nicht auf manual setzen: {calc_err}")
            
            # Screen-Updates ausschalten für schnelleres Öffnen
            try:
                self.app.screen_updating = False
            except:
                pass
            
            # Workbook öffnen (mit optionalem Passwort)
            # update_links=False verhindert blockierende Dialoge auf Windows
            _t1 = _time.time()
            self._log("books.open aufrufen...")
            if password:
                self.workbook = self.app.books.open(file_path, update_links=False, password=password)
            else:
                self.workbook = self.app.books.open(file_path, update_links=False)
            self._log(f"books.open fertig in {_time.time() - _t1:.1f}s")
            self.file_path = file_path
            
            # Screen-Updates wieder einschalten
            try:
                self.app.screen_updating = True
            except:
                pass
            
            # Excel MUSS versteckt bleiben nach books.open
            try:
                if not self.app.visible:
                    self.app.api.Visible = False
            except:
                pass
            
            # Sheet finden
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            self.worksheet = self.workbook.sheets[sheet_name]
            self.sheet_name = sheet_name
            
            # Sheet in Excel aktivieren (damit GUI und Excel synchron sind)
            try:
                self.worksheet.activate()
            except Exception as act_err:
                self._log(f"Sheet activate() fehlgeschlagen: {act_err}")
            
            # Berechnung NACH dem Öffnen erneut auf manual forcieren
            # (Workbooks können ihre eigene Einstellung mitbringen und auf automatic zurücksetzen)
            try:
                self.app.calculation = 'manual'
                self.app.enable_events = False
                self._log("calculation='manual' + enable_events=False für gesamte Session")
            except Exception as calc_err:
                self._log(f"Session-Calc-Setup fehlgeschlagen: {calc_err}")
            
            # Bedingte Formatierung deaktivieren (Performance-Optimierung)
            # CF-Regeln die über die gesamte Datei gehen verursachen massive
            # Recalculations bei JEDER Zelländerung → komplett abschalten
            has_cf = False
            if platform.system() == 'Windows':
                try:
                    cf_count = self.worksheet.api.Cells.FormatConditions.Count
                    has_cf = cf_count > 0
                    self._log(f"FormatConditions.Count = {cf_count}")
                except Exception:
                    pass
                try:
                    self.worksheet.api.EnableFormatConditionsCalculation = False
                    self._log("EnableFormatConditionsCalculation = False")
                except Exception as cf_err:
                    self._log(f"CF-Deaktivierung fehlgeschlagen: {cf_err}")
            
            # Recovery-System initialisieren
            self.backup_path = self._create_backup(file_path)
            self.journal_path = self._init_journal(file_path)
            self.last_auto_save = _time.time()
            
            # Undo-Stack leeren bei neuer Datei
            self.undo_stack.clear()
            
            # Read-Only-Check: Wenn die Datei bereits von einem anderen Prozess
            # geöffnet ist, öffnet Excel sie schreibgeschützt
            is_read_only = False
            try:
                is_read_only = self.workbook.api.ReadOnly
            except Exception:
                try:
                    is_read_only = not self.workbook.api.Saved and self.workbook.api.Path == ''
                except Exception:
                    pass
            
            _total = _time.time() - _t0
            self._log(f"Datei geöffnet in {_total:.1f}s, Sheet: {sheet_name}, Sheets: {sheet_names}, ReadOnly: {is_read_only}, HasCF: {has_cf}")
            return {'success': True, 'sheets': sheet_names, 'backupPath': self.backup_path, 'readOnly': is_read_only, 'hasConditionalFormatting': has_cf}
            
        except Exception as e:
            self._log(f"Fehler beim Öffnen nach {_time.time() - _t0:.1f}s: {e}")
            return {'success': False, 'error': str(e)}
    
    def save_file(self, output_path: Optional[str] = None, password: Optional[str] = None, selected_sheets: Optional[list] = None) -> Dict[str, Any]:
        """Speichert die Datei (optional unter neuem Namen und/oder mit Passwort)
        
        Args:
            output_path: Optionaler neuer Dateipfad
            password: Optionales Passwort (None = kein Passwort, '' = Passwort entfernen, 'xxx' = neues Passwort)
            selected_sheets: Optionale Liste von Sheet-Namen die exportiert werden sollen (None = alle)
        """
        import time as _time
        _t0 = _time.time()
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            # Debug: Zeige was wir bekommen haben
            self._log(f"save_file: output_path={output_path}, password={'KEEP' if password == 'KEEP' else ('***' if password else repr(password))}")
            
            # Passwort-Logik: 
            # 'KEEP' = altes Passwort beibehalten
            # None = kein neues Passwort
            # '' = Passwort entfernen
            # 'xxx' = neues Passwort setzen
            keep_password = (password == 'KEEP')
            
            # Effektives Passwort bestimmen
            if keep_password:
                effective_password = self.file_password
            elif password:
                effective_password = password
            else:
                effective_password = None
            
            # Vor dem Speichern: KEIN calculation='automatic' setzen!
            # Das triggert sofort eine Komplett-Neuberechnung aller Sheets/PivotTables
            # und blockiert den COM-Kanal für Minuten.
            # Excel speichert die Daten so wie sie sind — Formeln werden beim
            # nächsten manuellen Öffnen der Datei automatisch neuberechnet.
            
            # Sheet-Filterung: Nur ausgewählte Sheets exportieren
            all_sheet_names = [s.name for s in self.workbook.sheets]
            needs_sheet_filter = (selected_sheets is not None 
                                  and len(selected_sheets) > 0 
                                  and len(selected_sheets) < len(all_sheet_names))
            
            if needs_sheet_filter:
                return self._save_filtered(output_path, password, effective_password, keep_password, selected_sheets, all_sheet_names)
            
            if output_path and output_path != self.file_path:
                self._log(f"Speichere unter: {output_path}")
                
                # Windows: COM-API für SaveAs mit Passwort
                if effective_password:
                    self.workbook.api.SaveAs(output_path, FileFormat=51, Password=effective_password)
                elif not keep_password and self.file_password:
                    # Passwort entfernen: Erst Password leeren, dann SaveAs
                    self.workbook.api.Password = ''
                    self.workbook.api.SaveAs(output_path, FileFormat=51)
                else:
                    self.workbook.api.SaveAs(output_path, FileFormat=51)
                self.file_path = output_path
            else:
                # Windows: COM-API für Passwort-Änderungen
                if effective_password:
                    self.workbook.api.Password = effective_password
                    self.workbook.api.Save()
                elif not keep_password and password is not None:
                    # password == '' → Passwort entfernen
                    self.workbook.api.Password = ''
                    self.workbook.api.Save()
                else:
                    self.workbook.save()
            
            # Passwort-State aktualisieren
            # 'KEEP' ist ein Sentinel vom main.js-Handler (kein neues Passwort gewollt)
            # und darf NICHT als file_password gespeichert werden, sonst wird beim
            # nächsten Export die Datei mit dem Literal 'KEEP' verschlüsselt.
            if password is not None and password != 'KEEP':
                self.file_password = password if password else None
            
            # Nach dem Speichern: Session-Performance-Modus wiederherstellen
            if platform.system() == 'Windows':
                try:
                    self.app.calculation = 'manual'
                    self.app.enable_events = False
                    if self.worksheet:
                        self.worksheet.api.EnableFormatConditionsCalculation = False
                    self._log("Session-Performance-Modus nach Save wiederhergestellt")
                except Exception as restore_err:
                    self._log(f"Post-Save Restore fehlgeschlagen: {restore_err}")
            
            _t1 = _time.time()
            self._log(f"save_file: Excel save took {(_t1 - _t0)*1000:.0f}ms")
            
            # Recovery-Dateien aufräumen nach erfolgreichem Speichern
            self._cleanup_recovery(success=True)
            _t2 = _time.time()
            self._log(f"save_file: cleanup took {(_t2 - _t1)*1000:.0f}ms, total {(_t2 - _t0)*1000:.0f}ms")
            
            return {'success': True, 'outputPath': output_path or self.file_path, 'hasPassword': bool(self.file_password)}
            
        except Exception as e:
            self._log(f"Fehler beim Speichern: {e}")
            return {'success': False, 'error': str(e)}
    
    def _save_filtered(self, output_path, password, effective_password, keep_password, selected_sheets, all_sheet_names):
        """Speichert nur ausgewählte Sheets via SaveCopyAs + Sheet-Löschung"""
        import time as _time, tempfile, os
        _t0 = _time.time()
        
        final_path = output_path or self.file_path
        sheets_to_delete = [name for name in all_sheet_names if name not in selected_sheets]
        self._log(f"_save_filtered: {len(selected_sheets)} von {len(all_sheet_names)} Sheets, lösche: {sheets_to_delete}")
        
        try:
            # 1. Kopie erstellen mit SaveCopyAs (ändert NICHT die aktuelle Workbook-Zuordnung)
            temp_dir = tempfile.mkdtemp(prefix='edsp_export_')
            temp_path = os.path.join(temp_dir, 'temp_export.xlsx')
            self.workbook.api.SaveCopyAs(temp_path)
            self._log(f"SaveCopyAs -> {temp_path}")
            
            # 2. Kopie öffnen (im selben unsichtbaren Excel-Prozess)
            old_alerts = self.app.display_alerts
            self.app.display_alerts = False
            copy_wb = self.app.books.open(temp_path)
            
            try:
                # 3. Nicht ausgewählte Sheets löschen
                for sheet_name in sheets_to_delete:
                    try:
                        copy_wb.sheets[sheet_name].delete()
                    except Exception as del_err:
                        self._log(f"Sheet '{sheet_name}' löschen fehlgeschlagen: {del_err}")
                
                # 4. Passwort setzen falls nötig
                if effective_password:
                    copy_wb.api.Password = effective_password
                elif not keep_password and self.file_password:
                    copy_wb.api.Password = ''
                
                # 5. Als finale Datei speichern
                if final_path != temp_path:
                    copy_wb.api.SaveAs(final_path, FileFormat=51)
                else:
                    copy_wb.save()
                
                self._log(f"Gefilterte Datei gespeichert: {final_path}")
            finally:
                # 6. Kopie schließen
                copy_wb.close()
                self.app.display_alerts = old_alerts
                
                # Temp-Dateien aufräumen
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    os.rmdir(temp_dir)
                except Exception:
                    pass
            
            # Nach dem Speichern: Session-Performance-Modus wiederherstellen
            if platform.system() == 'Windows':
                try:
                    self.app.calculation = 'manual'
                    self.app.enable_events = False
                    if self.worksheet:
                        self.worksheet.api.EnableFormatConditionsCalculation = False
                except Exception as restore_err:
                    self._log(f"Post-Save Restore fehlgeschlagen: {restore_err}")
            
            _t1 = _time.time()
            self._log(f"_save_filtered: took {(_t1 - _t0)*1000:.0f}ms")
            
            self._cleanup_recovery(success=True)
            
            return {'success': True, 'outputPath': final_path, 'hasPassword': bool(effective_password)}
            
        except Exception as e:
            self._log(f"Fehler bei gefiltertem Speichern: {e}")
            return {'success': False, 'error': str(e)}
    
    def set_password(self, password: Optional[str]) -> Dict[str, Any]:
        """Setzt oder entfernt das Passwort für die Datei
        
        Args:
            password: Neues Passwort (None oder '' zum Entfernen)
        """
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if password:
                self._log("Setze Passwort...")
                # Windows: COM-API direkt verwenden
                self.workbook.api.Password = password
                self.workbook.api.Save()
                self.file_password = password
            else:
                self._log("Entferne Passwort...")
                # Windows: Passwort über COM-API leeren
                self.workbook.api.Password = ''
                self.workbook.api.Save()
                self.file_password = None
            
            return {'success': True, 'hasPassword': bool(self.file_password)}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen des Passworts: {e}")
            return {'success': False, 'error': str(e)}
    
    def get_password_status(self) -> Dict[str, Any]:
        """Gibt zurück ob die Datei ein Passwort hat"""
        return {
            'success': True,
            'hasPassword': bool(self.file_password),
            'passwordKnown': self.file_password is not None
        }
    
    # =========================================================================
    # UNDO-STACK  (Snapshot-basiert: SaveCopyAs vor jeder Operation)
    # =========================================================================
    
    def _push_undo_snapshot(self, label: str):
        """Speichert eine komplette Kopie des Workbooks als Undo-Snapshot.
        
        Verwendet Workbook.SaveCopyAs (Windows) bzw. shutil.copy2 (macOS)
        um den KOMPLETTEN Zustand zu sichern — inkl. Formatierung, Formeln,
        bedingte Formatierung, Datenvalidierung, etc.
        
        Args:
            label: Beschreibung der Operation (wird beim Undo angezeigt)
        """
        if self._undo_in_progress:
            return
        if not self.workbook or not self.file_path:
            return
        
        try:
            # Gleiche Dateiendung wie Original verwenden
            _, ext = os.path.splitext(self.file_path)
            if not ext:
                ext = '.xlsx'
            
            # Lokal im Scriptverzeichnis speichern (nicht im Temp/Netzwerk!)
            filename = f'_undo_{os.getpid()}_{int(time.time() * 1000)}{ext}'
            temp_path = os.path.join(self._undo_dir, filename)
            
            if platform.system() == 'Windows':
                # CutCopyMode zurücksetzen vor SaveCopyAs (verhindert Clipboard-Konflikte/Hänger)
                try:
                    self.app.api.CutCopyMode = False
                except Exception:
                    pass
                # SaveCopyAs: Speichert eine Kopie OHNE das offene Workbook zu ändern
                self.workbook.api.SaveCopyAs(temp_path)
            else:
                # macOS: Kein SaveCopyAs verfügbar.
                # Workbook speichern + Datei kopieren
                self.workbook.save()
                shutil.copy2(self.file_path, temp_path)
            
            # Auf den Stack pushen
            self.undo_stack.append({
                'type': 'snapshot',
                'label': label,
                'temp_path': temp_path,
                'sheet_name': self.sheet_name
            })
            
            # Alte Snapshots aufräumen wenn Stack zu groß
            while len(self.undo_stack) > self.MAX_UNDO:
                old = self.undo_stack.pop(0)
                self._cleanup_undo_entry(old)
            
            self._log(f"Undo-Snapshot: {label} → {os.path.basename(temp_path)}")
            
        except Exception as e:
            self._log(f"Undo-Snapshot Fehler (Operation wird trotzdem ausgeführt): {e}")
    
    def _push_undo_command(self, label: str, action: str, params: Dict[str, Any]):
        """Speichert eine Undo-Operation als Gegenbefehl (kein Disk-I/O).
        
        Statt das gesamte Workbook zu kopieren, wird nur die inverse Operation
        gespeichert. Extrem schnell, auch bei großen Dateien über Netzwerk.
        
        Args:
            label: Beschreibung (wird beim Undo angezeigt)
            action: Name der inversen Methode (z.B. 'delete_row', 'move_column')
            params: Parameter für die inverse Methode
        """
        if self._undo_in_progress:
            return
        
        self.undo_stack.append({
            'type': 'command',
            'label': label,
            'action': action,
            'params': params
        })
        
        while len(self.undo_stack) > self.MAX_UNDO:
            old = self.undo_stack.pop(0)
            self._cleanup_undo_entry(old)
        
        self._log(f"Undo-Command: {label} → {action}({params})")
    
    def _cleanup_undo_entry(self, entry: Dict):
        """Löscht die Temp-Datei eines Undo-Eintrags (nur für Snapshots)."""
        if entry.get('type') != 'snapshot':
            return
        try:
            temp_path = entry.get('temp_path')
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
        except Exception:
            pass
    
    def _cleanup_all_undo_files(self):
        """Löscht alle Undo-Temp-Dateien (beim Beenden der Session)."""
        for entry in self.undo_stack:
            self._cleanup_undo_entry(entry)
        self.undo_stack.clear()
        self._cleanup_leftover_undo_files()
        self._log("Alle Undo-Temp-Dateien aufgeräumt")
    
    def _cleanup_leftover_undo_files(self):
        """Löscht alle Dateien im _undo/ Verzeichnis (Reste von Abstürzen)."""
        try:
            if os.path.isdir(self._undo_dir):
                for f in os.listdir(self._undo_dir):
                    fp = os.path.join(self._undo_dir, f)
                    if os.path.isfile(fp):
                        os.unlink(fp)
        except Exception as e:
            self._log(f"Undo-Cleanup Fehler: {e}")
    
    def undo(self) -> Dict[str, Any]:
        """Macht die letzte Aktion rückgängig.
        
        Hybrides Undo-System:
        - type='command': Führt die inverse Operation aus (kein Disk-I/O, blitzschnell)
        - type='snapshot': Stellt kompletten Workbook-Zustand aus Datei wieder her (nur bei Löschungen)
        """
        try:
            if not self.file_path:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not self.undo_stack:
                return {'success': False, 'error': 'Nichts zum Rückgängig machen'}
            
            entry = self.undo_stack.pop()
            label = entry.get('label', 'Unbekannt')
            entry_type = entry.get('type', 'snapshot')
            
            self._undo_in_progress = True
            
            try:
                if entry_type == 'command':
                    return self._undo_command(entry, label)
                else:
                    return self._undo_snapshot(entry, label)
            finally:
                self._undo_in_progress = False
            
        except Exception as e:
            self._undo_in_progress = False
            self._log(f"Undo Fehler: {e}")
            return {'success': False, 'error': str(e)}
    
    def _undo_command(self, entry: Dict, label: str) -> Dict[str, Any]:
        """Führt eine inverse Operation aus (Command-basiertes Undo)."""
        action = entry.get('action')
        params = entry.get('params', {})
        
        self._log(f"Undo (Command): {label} → {action}({params})")
        
        # Dispatch zur inversen Methode
        method = getattr(self, action, None)
        if not method:
            return {'success': False, 'error': f'Unbekannte Undo-Aktion: {action}'}
        
        result = method(**params)
        
        if result.get('success'):
            self._log(f"Undo erfolgreich: {label} (noch {len(self.undo_stack)} Undo-Schritte)")
            return {'success': True, 'undone': label, 'action': action, 'params': params, 'undoCount': len(self.undo_stack)}
        else:
            return {'success': False, 'error': f'Undo fehlgeschlagen: {result.get("error")}'}
    
    def _undo_snapshot(self, entry: Dict, label: str) -> Dict[str, Any]:
        """Stellt den Workbook-Zustand aus einem Snapshot wieder her (nur für Löschungen)."""
        temp_path = entry.get('temp_path')
        sheet_name = entry.get('sheet_name', self.sheet_name)
        
        if not temp_path or not os.path.exists(temp_path):
            return {'success': False, 'error': 'Undo-Snapshot nicht gefunden'}
        
        original_path = self.file_path
        password = self.file_password
        
        self._log(f"Undo (Snapshot): {label} — Restore von {os.path.basename(temp_path)}")
        
        # 0. Fenster-Zustand sichern
        window_state = None
        app_visible = True
        try:
            if platform.system() == 'Windows':
                window_state = self.app.api.WindowState
                app_visible = self.app.api.Visible
            else:
                try:
                    window_state = self.app.api.bounds.get()
                except Exception:
                    pass
        except Exception as ws_err:
            self._log(f"Undo: WindowState lesen Fehler: {ws_err}")
        
        # 1. Workbook schließen (ohne Speichern)
        try:
            self.app.display_alerts = False
            if platform.system() == 'Windows':
                self.workbook.api.Saved = True
                self.workbook.api.Close(SaveChanges=False)
            else:
                try:
                    self.workbook.api.saved.set(True)
                except Exception:
                    pass
                self.workbook.close()
        except Exception as close_err:
            self._log(f"Undo: Schließen-Fehler (wird ignoriert): {close_err}")
        
        self.workbook = None
        self.worksheet = None
        
        # 2. Snapshot → Original kopieren
        self._log(f"Undo: Kopiere Snapshot → {os.path.basename(original_path)}")
        shutil.copy2(temp_path, original_path)
        
        # 3. Workbook wieder öffnen
        self._log("Undo: Workbook wieder öffnen...")
        try:
            self.app.screen_updating = False
        except Exception:
            pass
        
        if password:
            self.workbook = self.app.books.open(original_path, update_links=False, password=password)
        else:
            self.workbook = self.app.books.open(original_path, update_links=False)
        
        # 4. Sheet aktivieren
        try:
            self.worksheet = self.workbook.sheets[sheet_name]
            self.sheet_name = sheet_name
        except Exception:
            self.worksheet = self.workbook.sheets[0]
            self.sheet_name = self.worksheet.name
            self._log(f"Undo: Sheet '{sheet_name}' nicht gefunden, verwende '{self.sheet_name}'")
        
        self.worksheet.activate()
        self.file_path = original_path
        
        # 5. Fenster-Zustand wiederherstellen
        try:
            if platform.system() == 'Windows':
                if app_visible:
                    self.app.api.Visible = True
                if window_state is not None:
                    self.app.api.WindowState = window_state
            else:
                if window_state is not None:
                    try:
                        self.app.api.bounds.set(window_state)
                    except Exception:
                        pass
        except Exception as ws_err:
            self._log(f"Undo: WindowState wiederherstellen Fehler: {ws_err}")
        
        try:
            self.app.screen_updating = True
        except Exception:
            pass
        
        # 6. Temp-Datei aufräumen
        try:
            os.unlink(temp_path)
        except Exception:
            pass
        
        self._log(f"Undo erfolgreich: {label} (noch {len(self.undo_stack)} Undo-Schritte)")
        return {'success': True, 'undone': label, 'undoCount': len(self.undo_stack)}
    
    def close_session(self, save: bool = False) -> Dict[str, Any]:
        """Schließt die Session
        
        Args:
            save: Wenn True, wird vor dem Schließen gespeichert. Standard: False (ohne Speichern schließen)
        """
        try:
            self._log(f"Schließe Session... (save={save})")
            
            if self.app:
                try:
                    self.app.display_alerts = False
                except Exception as e:
                    self._log(f"display_alerts Fehler: {e}")
            
            # Interactive wieder aktivieren (falls als Read-Only-Vorschau gesperrt)
            if self.app:
                try:
                    if platform.system() == 'Windows':
                        self.app.api.Interactive = True
                    else:
                        self.app.api.interactive.set(True)
                except Exception as int_err:
                    self._log(f"Interactive-Restore Fehler (ignoriert): {int_err}")
            
            # Session-weite Einstellungen zurücksetzen
            if self.app and platform.system() == 'Windows':
                try:
                    self.app.calculation = 'automatic'
                    self.app.enable_events = True
                    self._log("calculation='automatic' + enable_events=True (Session-Ende)")
                except Exception as calc_err:
                    self._log(f"Calc-Restore fehlgeschlagen: {calc_err}")
            if self.worksheet and platform.system() == 'Windows':
                try:
                    self.worksheet.api.EnableFormatConditionsCalculation = True
                    self._log("EnableFormatConditionsCalculation = True (Session-Ende)")
                except Exception as cf_err:
                    self._log(f"CF-Reaktivierung fehlgeschlagen: {cf_err}")
            
            if self.workbook:
                try:
                    if save:
                        self.workbook.save()
                    else:
                        # Workbook als "gespeichert" markieren BEVOR close() aufgerufen wird.
                        # Verhindert den "Speichern?"-Dialog auf macOS.
                        try:
                            if platform.system() == 'Windows':
                                self.workbook.api.Saved = True
                            else:
                                self.workbook.api.saved.set(True)
                        except Exception as e:
                            self._log(f"Saved-Flag Fehler: {e}")
                    
                    # Schließen - plattformspezifisch
                    if platform.system() == 'Windows':
                        self.workbook.api.Close(SaveChanges=False)
                    else:
                        self.workbook.close()
                except Exception as e:
                    self._log(f"Fehler beim Schließen des Workbooks: {e}")
                self.workbook = None
            
            if self.app:
                try:
                    self.app.quit()
                except:
                    pass
                self.app = None
            
            # Undo-Temp-Dateien aufräumen
            self._cleanup_all_undo_files()
            
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
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'Zeile {row_index + 1} gelöscht')
            
            self._log(f"Lösche Zeile {excel_row}")
            self.worksheet.range(f'{excel_row}:{excel_row}').delete()
            
            # Journal-Eintrag
            self._journal_add('deleteRow', {'rowIndex': row_index})
            self._check_auto_save()
            
            return {'success': True, 'deletedRow': row_index}
            
        except Exception as e:
            self._log(f"Fehler beim Löschen der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def delete_rows_range(self, row_index: int, count: int = 1) -> Dict[str, Any]:
        """Löscht mehrere aufeinanderfolgende Zeilen (für Undo von insert_row).
        Kein Undo-Snapshot — wird nur intern von _undo_command aufgerufen."""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row_start = row_index + 2
            excel_row_end = excel_row_start + count - 1
            
            self._log(f"delete_rows_range: Lösche {count} Zeile(n) ab {excel_row_start}")
            self.worksheet.range(f'{excel_row_start}:{excel_row_end}').delete()
            
            return {'success': True, 'deletedAt': row_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Löschen der Zeilen: {e}")
            return {'success': False, 'error': str(e)}
    
    def insert_row(self, row_index: int, count: int = 1) -> Dict[str, Any]:
        """Fügt leere Zeilen ein (0-basierter Index, ohne Header)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row_start = row_index + 2
            excel_row_end = excel_row_start + count - 1
            
            # Undo: Inverse Operation = eingefügte Zeilen wieder löschen
            self._push_undo_command(
                f'{count} Zeile(n) eingefügt',
                'delete_rows_range',
                {'row_index': row_index, 'count': count}
            )
            
            self._log(f"Füge {count} Zeile(n) bei {excel_row_start} ein")
            self.worksheet.range(f'{excel_row_start}:{excel_row_end}').insert(shift='down')
            
            # Formatierung der neuen Zeile(n) löschen
            # Excel übernimmt bei Insert die Formatierung der darüberliegenden Zeile
            # (inkl. Hintergrundfarbe), was zu doppelten Farbmarkierungen führt.
            try:
                new_rows = self.worksheet.range(f'{excel_row_start}:{excel_row_end}')
                new_rows.color = None  # Hintergrundfarbe entfernen
                self._log(f"Formatierung der neuen Zeile(n) {excel_row_start}:{excel_row_end} bereinigt")
            except Exception as fmt_err:
                self._log(f"Formatierung-Bereinigung Fehler (ignoriert): {fmt_err}")
            
            # Journal-Eintrag
            self._journal_add('insertRow', {'rowIndex': row_index, 'count': count})
            self._check_auto_save()
            
            return {'success': True, 'insertedAt': row_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Einfügen der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def move_row(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """Verschiebt eine Zeile von from_index nach to_index"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            # Excel-Zeilen sind 1-basiert, Header ist Zeile 1, Daten beginnen bei Zeile 2
            excel_from = from_index + 2
            excel_to = to_index + 2
            
            # Undo: Inverse Operation = zurück verschieben
            self._push_undo_command(
                f'Zeile verschoben ({from_index + 1} → {to_index + 1})',
                'move_row',
                {'from_index': to_index, 'to_index': from_index}
            )
            
            self._log(f"Verschiebe Zeile {excel_from} nach {excel_to}")
            
            if from_index > to_index:
                # Nach oben verschieben
                source_row = self.worksheet.range(f'{excel_from}:{excel_from}')
                target_row = self.worksheet.range(f'{excel_to}:{excel_to}')
                
                if platform.system() == 'Windows':
                    source_row.api.Cut()
                    target_row.api.Insert(Shift=-4121)  # xlShiftDown
                else:
                    # macOS: Verwende xlwings copy() mit destination
                    # 1. Insert leere Zeile bei Ziel
                    target_row.insert(shift='down')
                    # 2. Quellzeile ist jetzt +1 gerutscht
                    new_from = excel_from + 1
                    source_row_new = self.worksheet.range(f'{new_from}:{new_from}')
                    target_row_new = self.worksheet.range(f'{excel_to}:{excel_to}')
                    # 3. Kopiere ganze Zeile mit xlwings copy()
                    source_row_new.copy(destination=target_row_new)
                    # 4. Lösche alte Zeile
                    source_row_new.delete(shift='up')
            else:
                # Nach unten verschieben
                insert_at = excel_to + 1
                source_row = self.worksheet.range(f'{excel_from}:{excel_from}')
                insert_row = self.worksheet.range(f'{insert_at}:{insert_at}')
                
                if platform.system() == 'Windows':
                    source_row.api.Cut()
                    insert_row.api.Insert(Shift=-4121)  # xlShiftDown
                else:
                    # macOS: Verwende xlwings copy() mit destination
                    # 1. Insert leere Zeile nach Ziel
                    insert_row.insert(shift='down')
                    # 2. Kopiere Quellzeile zur neuen Position
                    source_row_copy = self.worksheet.range(f'{excel_from}:{excel_from}')
                    dest_row = self.worksheet.range(f'{insert_at}:{insert_at}')
                    source_row_copy.copy(destination=dest_row)
                    # 3. Lösche alte Zeile
                    self.worksheet.range(f'{excel_from}:{excel_from}').delete(shift='up')
            
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
            
            # Undo: Inverse Operation = Sichtbarkeit umkehren
            self._push_undo_command(
                f'Zeile {excel_row} {"ausgeblendet" if hidden else "eingeblendet"}',
                'hide_row',
                {'row_index': row_index, 'hidden': not hidden}
            )
            
            # Nutze eine Zelle der Zeile und dann entire_row (analog zu hide_column)
            row_range = self.worksheet.range(f'A{excel_row}')
            
            if platform.system() == 'Windows':
                row_range.api.EntireRow.Hidden = hidden
            else:
                # macOS: Nutze xlwings api
                row_range.api.entire_row.hidden.set(hidden)
            
            self._log(f"Zeile {excel_row} {'versteckt' if hidden else 'angezeigt'}")
            return {'success': True, 'row': row_index, 'hidden': hidden}
            
        except Exception as e:
            self._log(f"Fehler beim Verstecken der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def hide_rows_batch(self, row_indices: list, hidden: bool = True) -> Dict[str, Any]:
        """Versteckt oder zeigt mehrere Zeilen auf einmal (Performance-optimiert)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not row_indices:
                return {'success': True, 'count': 0}
            
            # Undo: Inverse Operation = gleiche Zeilen mit umgekehrter Sichtbarkeit
            self._push_undo_command(
                f'{len(row_indices)} Zeilen {"ausgeblendet" if hidden else "eingeblendet"}',
                'hide_rows_batch',
                {'row_indices': list(row_indices), 'hidden': not hidden}
            )
            
            # Konvertiere zu Excel-Zeilen (0-basiert -> 1-basiert + Header)
            excel_rows = [idx + 2 for idx in row_indices]
            
            self._log(f"Batch {'verstecke' if hidden else 'zeige'} {len(excel_rows)} Zeilen")
            
            # Gruppiere aufeinanderfolgende Zeilen für effizientere Ranges
            # z.B. [2,3,4,7,8,10] -> ["A2:A4", "A7:A8", "A10:A10"]
            excel_rows.sort()
            ranges = []
            start = excel_rows[0]
            end = excel_rows[0]
            
            for row in excel_rows[1:]:
                if row == end + 1:
                    end = row
                else:
                    ranges.append(f'A{start}:A{end}')
                    start = row
                    end = row
            ranges.append(f'A{start}:A{end}')
            
            # Screen-Updating deaktivieren für Performance
            app = self.workbook.app
            screen_updating = app.screen_updating
            app.screen_updating = False
            
            try:
                for range_str in ranges:
                    row_range = self.worksheet.range(range_str)
                    if platform.system() == 'Windows':
                        row_range.api.EntireRow.Hidden = hidden
                    else:
                        # macOS: Nutze xlwings api
                        row_range.api.entire_row.hidden.set(hidden)
            finally:
                app.screen_updating = screen_updating
            
            self._log(f"Batch: {len(excel_rows)} Zeilen {'versteckt' if hidden else 'angezeigt'} ({len(ranges)} Ranges)")
            return {'success': True, 'count': len(excel_rows), 'hidden': hidden}
            
        except Exception as e:
            self._log(f"Fehler beim Batch-Verstecken: {e}")
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
            
            # Workbook NICHT als geändert markieren!
            # Farbmarkierungen sind nur visuell und sollen die
            # Originaldatei nicht verändern (kein Auto-Save/Speichern-Dialog).
            try:
                if platform.system() == 'Windows':
                    self.workbook.api.Saved = True
                else:
                    self.workbook.api.saved.set(True)
            except Exception as saved_err:
                self._log(f"Saved-Flag Fehler (ignoriert): {saved_err}")
            
            return {'success': True, 'row': row_index, 'color': color}
            
        except Exception as e:
            self._log(f"Fehler beim Markieren der Zeile: {e}")
            return {'success': False, 'error': str(e)}
    
    def highlight_rows_batch(self, rows: list, color: Optional[str] = None) -> Dict[str, Any]:
        """Markiert mehrere Zeilen mit Farbe in einem Aufruf"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            if not rows:
                return {'success': True, 'count': 0}
            
            last_col = self.worksheet.used_range.last_cell.column if self.worksheet.used_range else 10
            last_col_letter = self._get_column_letter(last_col)
            
            colors_map = {
                'green': (144, 238, 144),
                'yellow': (255, 255, 0),
                'orange': (255, 165, 0),
                'red': (255, 107, 107),
                'blue': (135, 206, 235),
                'purple': (221, 160, 221)
            }
            rgb = colors_map.get(color, (255, 255, 0))
            
            self.app.screen_updating = False
            try:
                for row_index in rows:
                    excel_row = row_index + 2
                    row_range = self.worksheet.range(f'A{excel_row}:{last_col_letter}{excel_row}')
                    if color is None:
                        row_range.color = None
                    else:
                        row_range.color = rgb
            finally:
                self.app.screen_updating = True
            
            try:
                if platform.system() == 'Windows':
                    self.workbook.api.Saved = True
                else:
                    self.workbook.api.saved.set(True)
            except Exception:
                pass
            
            self._log(f"{len(rows)} Zeilen markiert mit {color}")
            return {'success': True, 'count': len(rows)}
            
        except Exception as e:
            self._log(f"Fehler beim Batch-Markieren: {e}")
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
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'Spalte {col_letter} gelöscht')
            
            self._log(f"Lösche Spalte {col_letter} (Index {col_index})")
            
            # Verwende die native Excel API zum Löschen der gesamten Spalte
            # Dies stellt sicher, dass auch Tabellen-Bereiche korrekt angepasst werden
            if platform.system() == 'Darwin':
                # macOS: Verwende xlwings api direkt für ganze Spalte
                self.worksheet.range(f'{col_letter}:{col_letter}').api.entire_column.delete()
            else:
                # Windows: Verwende api.Delete() mit Shift-Parameter
                self.worksheet.range(f'{col_letter}:{col_letter}').api.Delete()
            
            # Screen refresh erzwingen
            self._force_screen_refresh()
            
            # Journal-Eintrag
            self._journal_add('deleteColumn', {'colIndex': col_index})
            self._check_auto_save()
            
            return {'success': True, 'deletedColumn': col_index}
            
        except Exception as e:
            self._log(f"Fehler beim Löschen der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    def delete_columns_range(self, col_index: int, count: int = 1) -> Dict[str, Any]:
        """Löscht mehrere aufeinanderfolgende Spalten (für Undo von insert_column).
        Kein Undo-Snapshot — wird nur intern von _undo_command aufgerufen."""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            for i in range(count - 1, -1, -1):  # Rückwärts löschen damit Indizes stimmen
                excel_col = col_index + 1 + i
                col_letter = self._get_column_letter(excel_col)
                self._log(f"delete_columns_range: Lösche Spalte {col_letter}")
                if platform.system() == 'Darwin':
                    self.worksheet.range(f'{col_letter}:{col_letter}').api.entire_column.delete()
                else:
                    self.worksheet.range(f'{col_letter}:{col_letter}').api.Delete()
            
            return {'success': True, 'deletedAt': col_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Löschen der Spalten: {e}")
            return {'success': False, 'error': str(e)}
    
    def insert_column(self, col_index: int, count: int = 1, headers: list = None) -> Dict[str, Any]:
        """Fügt leere Spalten ein"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_col = col_index + 1
            
            # Undo: Inverse Operation = eingefügte Spalten wieder löschen
            self._push_undo_command(
                f'{count} Spalte(n) eingefügt',
                'delete_columns_range',
                {'col_index': col_index, 'count': count}
            )
            
            # Performance: Screen-Updates + Berechnung deaktivieren während Spalten-Einfügung
            self.app.screen_updating = False
            old_calculation = self.app.calculation
            old_events = self.app.enable_events
            try:
                self.app.calculation = 'manual'
                self.app.enable_events = False
                if self.worksheet and platform.system() == 'Windows':
                    try:
                        self.worksheet.api.EnableFormatConditionsCalculation = False
                    except:
                        pass
                
                # 1) Spalten einfügen via COM API (PivotTable-sicher)
                start_letter = self._get_column_letter(excel_col)
                end_letter = self._get_column_letter(excel_col + count - 1)
                self._log(f"Füge {count} Spalte(n) {start_letter}:{end_letter} ein (COM API)")
                self.worksheet.api.Columns(f'{start_letter}:{end_letter}').Insert()
                
                # 2) Worksheet-Referenz refreshen nach COM Insert
                ws_name = self.worksheet.name
                self.worksheet = self.workbook.sheets[ws_name]
                
                # 3) Header setzen via xlwings (frische Referenz)
                if headers:
                    for i, header in enumerate(headers):
                        self.worksheet.range((1, excel_col + i)).value = header
            finally:
                # Session-Performance-Modus beibehalten:
                # EnableFormatConditionsCalculation=False, enable_events=False
                self.app.calculation = old_calculation
                self.app.enable_events = old_events
                self.app.screen_updating = True
            
            # Journal-Eintrag
            self._journal_add('insertColumn', {'colIndex': col_index, 'count': count, 'headers': headers})
            self._check_auto_save()
            
            return {'success': True, 'insertedAt': col_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Einfügen der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    def data_join_sync(self, operations: list) -> Dict[str, Any]:
        """Batch-Operation für Data Join: Spalten einfügen + Werte setzen in EINEM Aufruf.
        
        Args:
            operations: Liste von Operationen, jede mit:
                - position: 0-basierter Spaltenindex für Einfügung
                - count: Anzahl der einzufügenden Spalten
                - headers: Liste der Header-Namen
                - columnData: Liste von Listen (pro Spalte eine Werteliste)
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            sorted_ops = sorted(operations, key=lambda x: x.get('position', 0))
            insert_offset = 0
            total_inserted = 0
            total_values = 0
            
            self.app.screen_updating = False
            original_calculation = self.app.calculation
            self.app.calculation = 'manual'
            self.app.enable_events = False
            # Bedingte Formatierung deaktivieren (nur Windows)
            if platform.system() == 'Windows':
                try:
                    self.worksheet.api.EnableFormatConditionsCalculation = False
                except Exception:
                    pass
            try:
                for op in sorted_ops:
                    pos = op.get('position', 0) + insert_offset
                    count = op.get('count', 1)
                    headers = op.get('headers', [])
                    column_data = op.get('columnData', [])
                    
                    self._log(f"DataJoin op: pos={pos}, count={count}, headers={headers}, columnData={len(column_data)} arrays")
                    if column_data:
                        for ci, cd in enumerate(column_data):
                            non_empty = [v for v in cd if v != '' and v is not None][:5]
                            self._log(f"  columnData[{ci}]: {len(cd)} values, {len([v for v in cd if v != '' and v is not None])} non-empty, sample: {non_empty}")
                    
                    excel_col = pos + 1
                    
                    # Undo: Inverse Operation = eingefügte Spalten wieder löschen
                    self._push_undo_command(
                        f'DataJoin: {count} Spalte(n) eingefügt',
                        'delete_columns_range',
                        {'col_index': pos, 'count': count}
                    )
                    
                    # 1) Spalten einfügen via COM API (PivotTable-sicher)
                    if count > 0:
                        start_letter = self._get_column_letter(excel_col)
                        end_letter = self._get_column_letter(excel_col + count - 1)
                        self._log(f"DataJoin: Füge {count} Spalte(n) {start_letter}:{end_letter} ein (COM API)")
                        self.worksheet.api.Columns(f'{start_letter}:{end_letter}').Insert()
                    
                    # 2) Worksheet-Referenz refreshen nach COM Insert
                    # xlwings cached intern den Worksheet-Zustand; nach Columns.Insert()
                    # muss die Referenz erneuert werden, damit range().value korrekt schreibt
                    ws_name = self.worksheet.name
                    self.worksheet = self.workbook.sheets[ws_name]
                    
                    # 3) Header + Daten via xlwings schreiben (frische Referenz)
                    if headers:
                        for i, header in enumerate(headers):
                            self.worksheet.range((1, excel_col + i)).value = header
                    
                    # 4) Spaltendaten setzen — Zelle für Zelle (wie openpyxl)
                    #    Bulk range().value versagt auf PivotTable-Sheets,
                    #    deshalb: nur non-empty Werte einzeln schreiben
                    for i, col_values in enumerate(column_data):
                        if col_values and len(col_values) > 0:
                            target_col = excel_col + i
                            written = 0
                            for idx, val in enumerate(col_values):
                                if val is not None and val != '':
                                    self.worksheet.range((2 + idx, target_col)).value = val
                                    written += 1
                            self._log(f"  Spalte {target_col}: {written}/{len(col_values)} Werte geschrieben")
                            total_values += written
                    
                    total_inserted += count
                    insert_offset += count
                    
                    # Journal-Eintrag pro Operation
                    self._journal_add('insertColumn', {'colIndex': pos, 'count': count, 'headers': headers})
                
            finally:
                # Session-Performance-Modus beibehalten:
                # EnableFormatConditionsCalculation=False, enable_events=False
                self.app.enable_events = False
                self.app.calculation = original_calculation
                self.app.screen_updating = True
            
            # Verification: Gezielt non-empty Zellen zurücklesen
            debug_info = {}
            try:
                for op in sorted_ops:
                    pos = op.get('position', 0)
                    count = op.get('count', 1)
                    column_data = op.get('columnData', [])
                    for i in range(count):
                        if i < len(column_data) and column_data[i]:
                            # Finde die tatsächliche Excel-Spalte (mit Offset)
                            offset = 0
                            for prev_op in sorted_ops:
                                if prev_op.get('position', 0) < pos:
                                    offset += prev_op.get('count', 1)
                            actual_col = pos + offset + 1 + i
                            col_letter = self._get_column_letter(actual_col)
                            
                            # Finde Index des ersten non-empty Werts
                            first_ne_idx = None
                            first_ne_val = None
                            for idx, v in enumerate(column_data[i]):
                                if v != '' and v is not None:
                                    first_ne_idx = idx
                                    first_ne_val = v
                                    break
                            
                            # Lese Header (Zeile 1) zurück
                            header_read = self.worksheet.range((1, actual_col)).value
                            
                            verify_data = {
                                'header_readBack': str(header_read) if header_read else 'None',
                                'sentNonEmptyCount': len([v for v in column_data[i] if v != '' and v is not None]),
                                'sentTotal': len(column_data[i]),
                                'col_letter': col_letter,
                                'actual_col': actual_col,
                                'worksheet': self.worksheet.name
                            }
                            
                            # Gezielt die Zelle mit dem ersten non-empty Wert zurücklesen
                            if first_ne_idx is not None:
                                target_row = 2 + first_ne_idx  # +2 weil Zeile 1=Header, Daten ab Zeile 2
                                cell_read = self.worksheet.range((target_row, actual_col)).value
                                verify_data['firstNonEmpty'] = {
                                    'dataIndex': first_ne_idx,
                                    'excelRow': target_row,
                                    'wrote': str(first_ne_val),
                                    'readBack': str(cell_read) if cell_read is not None else 'None',
                                    'match': str(cell_read) == str(first_ne_val) if cell_read is not None else False
                                }
                                # Auch 2. und 3. non-empty Wert prüfen
                                ne_count = 0
                                for idx2, v2 in enumerate(column_data[i]):
                                    if v2 != '' and v2 is not None:
                                        ne_count += 1
                                        if ne_count in (2, 3):
                                            r = 2 + idx2
                                            rv = self.worksheet.range((r, actual_col)).value
                                            verify_data[f'nonEmpty_{ne_count}'] = {
                                                'excelRow': r, 'wrote': str(v2),
                                                'readBack': str(rv) if rv is not None else 'None'
                                            }
                                        if ne_count >= 3:
                                            break
                            
                            debug_info[f'{col_letter} (pos={pos})'] = verify_data
            except Exception as ve:
                debug_info['verifyError'] = str(ve)
            
            self._log(f"DataJoin sync: {total_inserted} Spalten, {total_values} Werte geschrieben")
            return {'success': True, 'insertedColumns': total_inserted, 'valuesWritten': total_values, 'debug': debug_info}
            
        except Exception as e:
            self.app.screen_updating = True
            self._log(f"Fehler beim DataJoin-Sync: {e}")
            return {'success': False, 'error': str(e)}
    
    def move_column(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """Verschiebt eine Spalte per Insert → Copy → Delete (zuverlässig auf allen Plattformen)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if from_index == to_index:
                return {'success': True, 'movedFrom': from_index, 'movedTo': to_index}
            
            excel_from = from_index + 1
            excel_to = to_index + 1
            
            source_letter = self._get_column_letter(excel_from)
            target_letter = self._get_column_letter(excel_to)
            
            # Undo: Inverse Operation = zurück verschieben
            self._push_undo_command(
                f'Spalte verschoben ({from_index + 1} → {to_index + 1})',
                'move_column',
                {'from_index': to_index, 'to_index': from_index}
            )
            
            self._log(f"Verschiebe Spalte {source_letter} (idx {from_index}) -> idx {to_index}")
            
            last_row = self.worksheet.used_range.last_cell.row if self.worksheet.used_range else 1000
            
            # Screen-Updates + Events aus für atomare Operation (verhindert Excel-Hänger)
            app = self.app
            app.screen_updating = False
            try:
                app.api.EnableEvents = False
            except Exception:
                pass
            
            try:
                # 1. Hidden-Rows merken und ALLE einblenden (wie Excel es erwartet)
                hidden_rows = []
                try:
                    for row_idx in range(1, last_row + 1):
                        if self.worksheet.api.Rows(row_idx).Hidden:
                            hidden_rows.append(row_idx)
                    if hidden_rows:
                        self._log(f"Blende {len(hidden_rows)} versteckte Zeilen ein vor Spaltenverschiebung")
                        # Alle Zeilen auf einmal einblenden
                        self.worksheet.api.Rows.Hidden = False
                except Exception as e:
                    self._log(f"Hidden-Row-State konnte nicht gesichert werden: {e}")
                
                # 2. Spalte verschieben via Cut + Insert (jetzt ohne Hidden-Row-Konflikt)
                source_col = self.worksheet.range(f'{source_letter}:{source_letter}')
                
                if from_index > to_index:
                    target_col = self.worksheet.range(f'{target_letter}:{target_letter}')
                    source_col.api.Cut()
                    target_col.api.Insert(Shift=-4161)  # xlShiftToRight
                else:
                    after_target_letter = self._get_column_letter(excel_to + 1)
                    after_target_col = self.worksheet.range(f'{after_target_letter}:{after_target_letter}')
                    source_col.api.Cut()
                    after_target_col.api.Insert(Shift=-4161)  # xlShiftToRight
                
                try:
                    app.api.CutCopyMode = False
                except Exception:
                    pass
                
                # 3. Hidden-Rows wiederherstellen
                if hidden_rows:
                    try:
                        for row_idx in hidden_rows:
                            self.worksheet.api.Rows(row_idx).Hidden = True
                        self._log(f"{len(hidden_rows)} versteckte Zeilen wiederhergestellt")
                    except Exception as e:
                        self._log(f"Hidden-Row-State konnte nicht wiederhergestellt werden: {e}")
            finally:
                # Events + Screen-Updates immer wieder aktivieren
                try:
                    app.api.EnableEvents = True
                except Exception:
                    pass
                app.screen_updating = True
            
            # Journal-Eintrag
            self._journal_add('moveColumn', {'fromIndex': from_index, 'toIndex': to_index})
            self._check_auto_save()
            
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
            
            # Undo: Inverse Operation = Sichtbarkeit umkehren
            self._push_undo_command(
                f'Spalte {col_letter} {"ausgeblendet" if hidden else "eingeblendet"}',
                'hide_column',
                {'col_index': col_index, 'hidden': not hidden}
            )
            
            # Nutze eine Zelle der Spalte und dann entire_column
            col_range = self.worksheet.range(f'{col_letter}1')
            
            if platform.system() == 'Windows':
                col_range.api.EntireColumn.Hidden = hidden
            else:
                # macOS: Nutze xlwings api
                col_range.api.entire_column.hidden.set(hidden)
            
            self._log(f"Spalte {col_letter} {'versteckt' if hidden else 'angezeigt'}")
            return {'success': True, 'column': col_index, 'hidden': hidden}
            
        except Exception as e:
            self._log(f"Fehler beim Verstecken der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    # =========================================================================
    # ZELL-OPERATIONEN
    # =========================================================================
    
    def restore_cell_value(self, row_index: int, col_index: int, value: Any) -> Dict[str, Any]:
        """Stellt einen einzelnen Zellwert wieder her (nur für Undo-Dispatch)."""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            excel_row = row_index + 2
            excel_col = col_index + 1
            self._log(f"restore_cell_value (Undo): ({excel_row},{excel_col}) = {value!r}")
            if platform.system() == 'Darwin':
                self.worksheet.range((excel_row, excel_col)).value = value
            else:
                self.app.screen_updating = False
                self.worksheet.range((excel_row, excel_col)).value = value
                self.app.screen_updating = True
            return {'success': True, 'row': row_index, 'col': col_index, 'value': value}
        except Exception as e:
            self._log(f"Fehler restore_cell_value: {e}")
            return {'success': False, 'error': str(e)}
    
    def restore_cells_batch(self, cells: list) -> Dict[str, Any]:
        """Stellt mehrere Zellwerte wieder her (nur für Undo-Dispatch)."""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            if not cells:
                return {'success': True, 'count': 0}
            self._log(f"restore_cells_batch (Undo): {len(cells)} Zellen")
            app = self.app
            app.screen_updating = False
            try:
                from collections import defaultdict
                rows_map = defaultdict(dict)
                for cell in cells:
                    r = cell.get('row')
                    c = cell.get('col')
                    if r is not None and c is not None:
                        rows_map[r][c] = cell.get('value')
                for row_index in sorted(rows_map.keys()):
                    cols = rows_map[row_index]
                    excel_row = row_index + 2
                    if len(cols) == 1:
                        col_index = next(iter(cols))
                        self.worksheet.range((excel_row, col_index + 1)).value = cols[col_index]
                    else:
                        sorted_cols = sorted(cols.keys())
                        min_c = sorted_cols[0]
                        max_c = sorted_cols[-1]
                        row_values = [cols.get(c) for c in range(min_c, max_c + 1)]
                        start_letter = self._get_column_letter(min_c + 1)
                        end_letter = self._get_column_letter(max_c + 1)
                        rng = f'{start_letter}{excel_row}:{end_letter}{excel_row}'
                        self.worksheet.range(rng).value = row_values
            finally:
                app.screen_updating = True
            return {'success': True, 'count': len(cells)}
        except Exception as e:
            self._log(f"Fehler restore_cells_batch: {e}")
            return {'success': False, 'error': str(e)}
    
    def set_cell_value(self, row_index: int, col_index: int, value: Any, old_value: Any = None) -> Dict[str, Any]:
        """Setzt den Wert einer einzelnen Zelle (formatierungsschonend)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row = row_index + 2
            excel_col = col_index + 1
            
            self._log(f"set_cell_value: row={row_index}, col={col_index}, excel=({excel_row},{excel_col})")
            
            # Schutz: Zellen mit IMAGE/DISPIMG-Formeln dürfen nicht überschrieben werden
            cell = self.worksheet.range((excel_row, excel_col))
            formula = cell.formula
            if formula and isinstance(formula, str):
                formula_upper = formula.upper()
                if '=DISPIMG(' in formula_upper or '=IMAGE(' in formula_upper or '_xlfn.DISPIMG(' in formula_upper:
                    self._log(f"set_cell_value: Zelle ({excel_row},{excel_col}) enthält Bild-Formel — Überschreiben verhindert")
                    return {'success': True, 'skipped': True, 'reason': 'image_formula'}
            
            # Alten Wert für Undo: vom Frontend übernehmen oder von COM lesen
            if old_value is None:
                old_value = cell.value
            self._push_undo_command(
                f'Zelle ({row_index + 1},{col_index + 1}) geändert',
                'restore_cell_value',
                {'row_index': row_index, 'col_index': col_index, 'value': old_value}
            )
            
            if platform.system() == 'Darwin':
                self.worksheet.range((excel_row, excel_col)).value = value
            else:
                # Windows: screen_updating False→True für zuverlässigen Redraw
                self.app.screen_updating = False
                try:
                    self.worksheet.range((excel_row, excel_col)).value = value
                finally:
                    self.app.screen_updating = True
            
            # Änderung im Journal protokollieren
            self._journal_add('setCellValue', {
                'row': row_index,
                'col': col_index,
                'newValue': str(value) if value else None
            })
            
            # Prüfen ob Auto-Save fällig ist
            self._check_auto_save()
            
            return {'success': True, 'row': row_index, 'col': col_index, 'value': value}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen des Zellwerts: {e}")
            import traceback
            self._log(traceback.format_exc())
            return {'success': False, 'error': str(e)}
    
    def set_column_values(self, col_index: int, values: list, start_row: int = 0) -> Dict[str, Any]:
        """Setzt alle Werte einer Spalte auf einmal (effizienter als einzelne setCellValue Aufrufe)
        
        Args:
            col_index: 0-basierter Spaltenindex
            values: Liste von Werten (für jede Zeile ein Wert)
            start_row: 0-basierter Start-Zeilenindex (Default: 0 = erste Datenzeile)
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_col = col_index + 1
            excel_start_row = start_row + 2  # +2 weil Header in Zeile 1
            
            end_row = excel_start_row + len(values) - 1
            col_letter = self._get_column_letter(excel_col)
            range_addr = f'{col_letter}{excel_start_row}:{col_letter}{end_row}'
            
            # Werte als vertikale Liste formatieren (jeder Wert in eigener Liste)
            vertical_values = [[v] for v in values]
            
            self._log(f"Setze {len(values)} Werte in Spalte {col_letter} (Zeilen {excel_start_row}-{end_row})")
            
            # Performance: Screen-Updates deaktivieren für Bulk-Write
            self.app.screen_updating = False
            try:
                self.worksheet.range(range_addr).value = vertical_values
            finally:
                self.app.screen_updating = True
            
            return {'success': True, 'colIndex': col_index, 'count': len(values)}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen der Spaltenwerte: {e}")
            return {'success': False, 'error': str(e)}
    
    def set_row_values(self, row_index: int, values: list) -> Dict[str, Any]:
        """Setzt alle Werte einer Zeile auf einmal (komplette Zeile übertragen)
        
        Args:
            row_index: 0-basierter Zeilenindex (Datenzeile, ohne Header)
            values: Liste von Werten (für jede Spalte ein Wert)
        """
        try:
            if not self.worksheet:
                self._log(f"set_row_values: Keine Datei geöffnet (row={row_index})")
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not values:
                return {'success': True, 'rowIndex': row_index, 'count': 0}
            
            excel_row = row_index + 2  # +2 weil Header in Zeile 1
            num_cols = len(values)
            
            start_col_letter = self._get_column_letter(1)
            end_col_letter = self._get_column_letter(num_cols)
            range_addr = f'{start_col_letter}{excel_row}:{end_col_letter}{excel_row}'
            
            self._log(f"set_row_values: row={row_index}, excel_row={excel_row}, cols={num_cols}, range={range_addr}")
            
            if platform.system() == 'Darwin':
                # macOS: Direkt über AppleScript für zuverlässigen Display-Refresh
                self._set_row_values_applescript(excel_row, values)
            else:
                # Windows: screen_updating=False vor Write, dann True danach
                # Der Übergang False→True erzwingt einen kompletten Redraw
                self._log(f"set_row_values: Schreibe in Range {range_addr}...")
                self.app.screen_updating = False
                try:
                    self.worksheet.range(range_addr).value = values
                finally:
                    self.app.screen_updating = True
                self._log(f"set_row_values: Range-Write + Refresh abgeschlossen")
            
            self._log(f"set_row_values: Erfolgreich (row={row_index}, cols={num_cols})")
            
            # Journal
            self._journal_add('setRowValues', {
                'row': row_index,
                'count': num_cols
            })
            
            self._check_auto_save()
            
            return {'success': True, 'rowIndex': row_index, 'count': num_cols}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen der Zeilenwerte (row={row_index}): {e}")
            import traceback
            self._log(traceback.format_exc())
            return {'success': False, 'error': str(e)}
    
    def _set_row_values_applescript(self, excel_row: int, values: list):
        """Setzt Zeilenwerte über AppleScript (macOS) - garantiert Display-Refresh
        
        Args:
            excel_row: 1-basierte Excel-Zeile
            values: Liste von Werten
        """
        import subprocess
        
        # AppleScript-Befehle generieren, die jede Zelle direkt setzen
        set_commands = []
        for col_idx, value in enumerate(values):
            col = col_idx + 1
            # Wert für AppleScript vorbereiten
            if value is None or value == '' or value == 'None':
                set_commands.append(f'set value of cell {col} of row {excel_row} of active sheet to ""')
            elif isinstance(value, (int, float)):
                set_commands.append(f'set value of cell {col} of row {excel_row} of active sheet to {value}')
            else:
                # Strings escapen für AppleScript (Backslash und Anführungszeichen)
                escaped = str(value).replace('\\', '\\\\').replace('"', '\\"')
                set_commands.append(f'set value of cell {col} of row {excel_row} of active sheet to "{escaped}"')
        
        # In Batches aufteilen (AppleScript hat Längenlimits)
        batch_size = 50
        for i in range(0, len(set_commands), batch_size):
            batch = set_commands[i:i + batch_size]
            commands_str = '\n'.join(batch)
            
            script = f'''tell application "Microsoft Excel"
{commands_str}
end tell'''
            
            try:
                result = subprocess.run(
                    ['osascript', '-e', script],
                    capture_output=True, text=True, timeout=30
                )
                if result.returncode != 0:
                    self._log(f"AppleScript Fehler: {result.stderr.strip()}")
                    # Fallback auf xlwings
                    self._log("Fallback auf xlwings Range-Write")
                    start_col = self._get_column_letter(1)
                    end_col = self._get_column_letter(len(values))
                    self.worksheet.range(f'{start_col}{excel_row}:{end_col}{excel_row}').value = values
                    return
            except subprocess.TimeoutExpired:
                self._log("AppleScript Timeout - Fallback auf xlwings")
                start_col = self._get_column_letter(1)
                end_col = self._get_column_letter(len(values))
                self.worksheet.range(f'{start_col}{excel_row}:{end_col}{excel_row}').value = values
                return
    
    def set_cells_batch(self, cells: list) -> Dict[str, Any]:
        """Setzt mehrere Zellen auf einmal — zeilenweise gruppiert für Performance."""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not cells or len(cells) == 0:
                return {'success': True, 'count': 0}
            
            self._log(f"set_cells_batch: Setze {len(cells)} Zellen")
            
            updated_count = 0
            app = self.app
            app.screen_updating = False
            
            try:
                # Zellen nach Zeile gruppieren für Range-Writes statt Einzelzellen
                from collections import defaultdict
                rows_map = defaultdict(dict)
                old_values_map = {}
                has_all_old_values = True
                for cell in cells:
                    r = cell.get('row')
                    c = cell.get('col')
                    if r is None or c is None:
                        continue
                    rows_map[r][c] = cell.get('value')
                    if 'oldValue' in cell:
                        old_values_map[(r, c)] = cell['oldValue']
                    else:
                        has_all_old_values = False
                
                # Alte Werte: vom Frontend übernehmen oder von COM lesen
                old_cells = []
                if has_all_old_values and old_values_map:
                    # Frontend hat alte Werte mitgeliefert — kein COM-Zugriff nötig
                    for row_index in sorted(rows_map.keys()):
                        cols = rows_map[row_index]
                        for col_index in sorted(cols.keys()):
                            old_cells.append({'row': row_index, 'col': col_index, 'value': old_values_map.get((row_index, col_index))})
                else:
                    # Fallback: Alte Werte von COM lesen
                    for row_index in sorted(rows_map.keys()):
                        cols = rows_map[row_index]
                        excel_row = row_index + 2
                        for col_index in sorted(cols.keys()):
                            old_val = self.worksheet.range((excel_row, col_index + 1)).value
                            old_cells.append({'row': row_index, 'col': col_index, 'value': old_val})
                
                self._push_undo_command(
                    f'{len(old_cells)} Zelle(n) geändert',
                    'restore_cells_batch',
                    {'cells': old_cells}
                )
                
                for row_index in sorted(rows_map.keys()):
                    cols = rows_map[row_index]
                    excel_row = row_index + 2  # +2 für Header
                    
                    if len(cols) == 1:
                        # Einzelzelle direkt setzen
                        col_index = next(iter(cols))
                        self.worksheet.range((excel_row, col_index + 1)).value = cols[col_index]
                        updated_count += 1
                    else:
                        # Zusammenhängenden Bereich als Range schreiben
                        sorted_cols = sorted(cols.keys())
                        min_c = sorted_cols[0]
                        max_c = sorted_cols[-1]
                        row_values = [cols.get(c) for c in range(min_c, max_c + 1)]
                        start_letter = self._get_column_letter(min_c + 1)
                        end_letter = self._get_column_letter(max_c + 1)
                        rng = f'{start_letter}{excel_row}:{end_letter}{excel_row}'
                        self.worksheet.range(rng).value = row_values
                        updated_count += len(cols)
            finally:
                app.screen_updating = True
            
            # Änderungen im Journal protokollieren (vereinfacht)
            self._journal_add('setCellsBatch', {
                'count': updated_count
            })
            
            # Prüfen ob Auto-Save fällig ist
            self._check_auto_save()
            
            return {'success': True, 'count': updated_count}
            
        except Exception as e:
            self._log(f"Fehler beim Batch-Setzen der Zellwerte: {e}")
            return {'success': False, 'error': str(e)}
    
    def copy_cells(self, source_cells: list, target_row: int, target_col: int) -> Dict[str, Any]:
        """Kopiert Zellen in Excel mit nativer copy()-Funktion.
        
        Nutzt xlwings range.copy(destination=...) - kopiert ALLES:
        Werte, Formatierung, Formeln, Rahmen, Schriftart, Größe, etc.
        Merged Cells werden erkannt und am Ziel explizit mit merge() angelegt.
        
        source_cells: Liste von {row, col} (0-basierte Daten-Indizes)
        target_row: 0-basierter Ziel-Zeilen-Index
        target_col: 0-basierter Ziel-Spalten-Index
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not source_cells or len(source_cells) == 0:
                return {'success': True, 'count': 0}
            
            self._log(f"copy_cells: {len(source_cells)} Zellen kopieren nach ({target_row},{target_col})")
            
            app = self.app
            original_screen_updating = app.screen_updating
            merged_regions = []
            
            try:
                if len(source_cells) > 5:
                    app.screen_updating = False
                
                # Ermittle zusammenhängenden Quellbereich (Rechteck)
                rows = [c.get('row') for c in source_cells if c.get('row') is not None]
                cols = [c.get('col') for c in source_cells if c.get('col') is not None]
                
                if not rows or not cols:
                    return {'success': False, 'error': 'Keine gültigen Zell-Koordinaten'}
                
                min_row, max_row = min(rows), max(rows)
                min_col, max_col = min(cols), max(cols)
                
                # Excel-Koordinaten (1-basiert, +2 für Header-Zeile)
                src_start_row = min_row + 2
                src_start_col = min_col + 1
                src_end_row = max_row + 2
                src_end_col = max_col + 1
                
                dst_start_row = target_row + 2
                dst_start_col = target_col + 1
                
                # Undo-Snapshot: Komplettes Workbook sichern
                dst_end_row = dst_start_row + (max_row - min_row)
                dst_end_col = dst_start_col + (max_col - min_col)
                self._push_undo_snapshot(f'Zellen eingefügt ({max_row - min_row + 1}×{max_col - min_col + 1})')
                
                row_offset = dst_start_row - src_start_row
                col_offset = dst_start_col - src_start_col
                
                # Quellbereich
                source_range = self.worksheet.range(
                    (src_start_row, src_start_col),
                    (src_end_row, src_end_col)
                )
                
                # Merged Cells im Quellbereich erkennen BEVOR wir kopieren
                seen_merges = set()
                for r in range(src_start_row, src_end_row + 1):
                    for c in range(src_start_col, src_end_col + 1):
                        cell = self.worksheet.range((r, c))
                        if cell.merge_cells:
                            ma = cell.merge_area
                            merge_key = ma.address
                            if merge_key not in seen_merges:
                                seen_merges.add(merge_key)
                                merged_regions.append({
                                    'src_start_row': ma.row,
                                    'src_start_col': ma.column,
                                    'src_end_row': ma.row + ma.shape[0] - 1,
                                    'src_end_col': ma.column + ma.shape[1] - 1,
                                    'row_span': ma.shape[0],
                                    'col_span': ma.shape[1]
                                })
                
                self._log(f"copy_cells: {len(merged_regions)} Merged-Bereiche im Quellbereich gefunden")
                
                # Zielbereich: bestehende Merges aufheben
                dest_end_row = dst_start_row + (src_end_row - src_start_row)
                dest_end_col = dst_start_col + (src_end_col - src_start_col)
                dest_full_range = self.worksheet.range(
                    (dst_start_row, dst_start_col),
                    (dest_end_row, dest_end_col)
                )
                try:
                    dest_full_range.unmerge()
                except Exception:
                    pass
                
                # Zielbereich (obere linke Ecke reicht für copy)
                dest_range = self.worksheet.range((dst_start_row, dst_start_col))
                
                # Native xlwings copy - kopiert Werte + alle Formatierungen
                source_range.copy(destination=dest_range)
                
                # CutCopyMode sofort zurücksetzen (verhindert Clipboard-Konflikte bei SaveCopyAs)
                try:
                    app.api.CutCopyMode = False
                except Exception:
                    pass
                
                # Merged Cells am Ziel explizit anlegen
                for merge in merged_regions:
                    try:
                        dst_r1 = merge['src_start_row'] + row_offset
                        dst_c1 = merge['src_start_col'] + col_offset
                        dst_r2 = merge['src_end_row'] + row_offset
                        dst_c2 = merge['src_end_col'] + col_offset
                        merge_range = self.worksheet.range((dst_r1, dst_c1), (dst_r2, dst_c2))
                        merge_range.merge()
                        self._log(f"  Merge angelegt: ({dst_r1},{dst_c1}):({dst_r2},{dst_c2})")
                    except Exception as me:
                        self._log(f"  Merge fehlgeschlagen: {me}")
                
                self._log(f"copy_cells: Kopiert ({src_start_row},{src_start_col}):({src_end_row},{src_end_col}) → ({dst_start_row},{dst_start_col})")
                
                count = (max_row - min_row + 1) * (max_col - min_col + 1)
                
            finally:
                if len(source_cells) > 5:
                    app.screen_updating = original_screen_updating
                
                if platform.system() == 'Darwin':
                    self._force_screen_refresh()
            
            self._journal_add('copyCells', {'count': count, 'merges': len(merged_regions)})
            self._check_auto_save()
            
            # Merge-Info zurückgeben für GUI-Update (0-basierte Daten-Indizes)
            merge_info = []
            for merge in merged_regions:
                row_offset_data = target_row - min_row
                col_offset_data = target_col - min_col
                merge_info.append({
                    'startRow': (merge['src_start_row'] - 2 + row_offset_data) + 1,  # Excel-0-basiert (wie in mergedCells)
                    'startCol': merge['src_start_col'] - 1 + col_offset_data,
                    'endRow': (merge['src_end_row'] - 2 + row_offset_data) + 1,
                    'endCol': merge['src_end_col'] - 1 + col_offset_data,
                    'rowSpan': merge['row_span'],
                    'colSpan': merge['col_span']
                })
            
            return {'success': True, 'count': count, 'mergedCells': merge_info}
            
        except Exception as e:
            self._log(f"Fehler beim Kopieren der Zellen: {e}")
            return {'success': False, 'error': str(e)}
    
    def find_replace(self, search_text: str, replace_text: str, 
                     match_case: bool = False, whole_word: bool = False) -> Dict[str, Any]:
        """Nutzt Excel's native Suchen & Ersetzen Funktion - extrem schnell!
        
        Args:
            search_text: Text der gesucht werden soll
            replace_text: Ersetzungstext
            match_case: Groß-/Kleinschreibung beachten
            whole_word: Nur ganze Wörter ersetzen
        
        Returns:
            Dict mit success und count der ersetzten Vorkommen
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not search_text:
                return {'success': False, 'error': 'Suchtext fehlt'}
            
            self._log(f"find_replace: '{search_text}' -> '{replace_text}' (case={match_case}, whole={whole_word})")
            
            # Undo-Eintrag: Umgekehrtes Suchen & Ersetzen
            if replace_text:
                self._push_undo_command(
                    f'Suchen/Ersetzen: "{search_text}" → "{replace_text}"',
                    'find_replace',
                    {
                        'search_text': replace_text,
                        'replace_text': search_text,
                        'match_case': match_case,
                        'whole_word': whole_word
                    }
                )
            
            import platform
            
            app = self.app
            original_screen_updating = app.screen_updating
            
            try:
                app.screen_updating = False
                
                # Excel's native Replace-Funktion über API aufrufen
                if platform.system() == 'Darwin':
                    # macOS: AppleScript direkt aufrufen (umgeht xlwings Name-Probleme mit Sonderzeichen)
                    ws = self.worksheet
                    
                    # Worksheet aktivieren um sicherzustellen dass es das aktive Sheet ist
                    ws.activate()
                    
                    # AppleScript direkt ausführen - verwendet "active sheet" statt Workbook/Sheet-Namen
                    # Das vermeidet Probleme mit Sonderzeichen wie & im Namen
                    import subprocess
                    
                    # AppleScript Replace-Befehl
                    # look_at: 1=whole, 2=part
                    look_at = 1 if whole_word else 2
                    match_case_val = "true" if match_case else "false"
                    
                    # Escape für AppleScript
                    search_escaped = search_text.replace('\\', '\\\\').replace('"', '\\"')
                    replace_escaped = replace_text.replace('\\', '\\\\').replace('"', '\\"')
                    
                    script = f'''
                    tell application "Microsoft Excel"
                        tell active sheet
                            set usedRng to used range
                            replace usedRng what "{search_escaped}" replacement "{replace_escaped}" look at {look_at} match case {match_case_val}
                        end tell
                    end tell
                    '''
                    
                    result = subprocess.run(
                        ['osascript', '-e', script],
                        capture_output=True,
                        text=True,
                        timeout=30
                    )
                    
                    if result.returncode != 0:
                        self._log(f"AppleScript Replace stderr: {result.stderr}")
                        # Bei Fehler: Fallback auf einzelne Zellen
                        raise Exception(f"AppleScript Replace fehlgeschlagen: {result.stderr}")
                    
                    replaced = True
                    count = -1  # Excel gibt nicht die Anzahl zurück
                else:
                    # Windows: COM-API
                    ws = self.worksheet
                    used_range = ws.used_range
                    
                    # xlReplace-Konstanten
                    xlPart = 2
                    xlWhole = 1
                    look_at = xlWhole if whole_word else xlPart
                    
                    replaced = used_range.api.Replace(
                        What=search_text,
                        Replacement=replace_text,
                        LookAt=look_at,
                        MatchCase=match_case
                    )
                    
                    count = -1  # Excel gibt nicht die Anzahl zurück
                
            finally:
                app.screen_updating = original_screen_updating
            
            # Journal-Eintrag
            self._journal_add('findReplace', {
                'search': search_text,
                'replace': replace_text
            })
            
            self._check_auto_save()
            
            return {'success': True, 'replaced': True}
            
        except Exception as e:
            self._log(f"Fehler bei find_replace: {e}")
            return {'success': False, 'error': str(e)}
    
    # =========================================================================
    # FILTER-OPERATIONEN
    # =========================================================================
    
    def set_autofilter(self, filters: list = None) -> Dict[str, Any]:
        """Setzt AutoFilter auf das Worksheet
        
        Args:
            filters: Liste von Filter-Definitionen:
                     [{ colIndex: 0, criteria: "value" }, ...]
                     Wenn None oder leer, wird AutoFilter entfernt
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            used_range = self.worksheet.used_range
            if not used_range:
                return {'success': False, 'error': 'Keine Daten vorhanden'}
            
            self._log(f"set_autofilter: {len(filters) if filters else 0} Filter")
            
            # === Performance-Guard: COM-Calls blockieren Excel bei großen Dateien ===
            # calculation und enable_events MÜSSEN deaktiviert sein,
            # sonst triggert jeder AutoFilter()-Call eine Neuberechnung.
            # NICHT screen_updating togglen — das kann eine zweite Excel-Instanz öffnen!
            app = self.app
            try:
                app.enable_events = False
                if app.calculation != 'manual':
                    app.calculation = 'manual'
                    self._log("  calculation → 'manual' gesetzt")
            except Exception as perf_err:
                self._log(f"  Performance-Guard Setup Fehler: {perf_err}")
            
            try:
                if filters and len(filters) > 0:
                    # AutoFilter aktivieren falls noch nicht aktiv
                    try:
                        if not self.worksheet.api.AutoFilterMode:
                            used_range.api.AutoFilter()
                    except Exception as e:
                        self._log(f"AutoFilter-Aktivierung Fehler: {e}")
                        try:
                            used_range.api.AutoFilter()
                        except:
                            pass
                    
                    # === Nur APP-EIGENE Filter-Felder zurücksetzen ===
                    # NICHT ShowAllData() — das blendet auch manuell versteckte Zeilen
                    # und vorhandene Excel-Filter ein!
                    # Stattdessen: Nur die Spalten resetten, die WIR zuvor gesetzt haben.
                    if not hasattr(self, '_active_filter_fields_per_sheet'):
                        self._active_filter_fields_per_sheet = {}
                    current_sheet = self.sheet_name or ''
                    old_fields = self._active_filter_fields_per_sheet.get(current_sheet, [])
                    if old_fields:
                        self._log(f"  Resette {len(old_fields)} App-eigene Filter-Felder: {old_fields}")
                        for col_idx in old_fields:
                            try:
                                af = self.worksheet.api.AutoFilter
                                if af:
                                    af.Range.AutoFilter(Field=col_idx)
                            except Exception as e:
                                self._log(f"  Reset Feld {col_idx} Fehler: {e}")
                    
                    # Filter nach Spalte gruppieren (gleiche Spalte = OR-Verknüpfung)
                    from collections import defaultdict
                    filters_by_col = defaultdict(list)
                    for f in filters:
                        col_idx = f.get('colIndex', 0) + 1  # 1-basiert
                        filters_by_col[col_idx].append(f)
                    
                    self._active_filter_fields_per_sheet[current_sheet] = []  # Merken für Clear
                    self._active_filter_fields = []  # Legacy-Kompatibilität
                    for col_idx, col_filters in filters_by_col.items():
                        criteria_list = []
                        is_date = False
                        
                        for f in col_filters:
                            criteria = f.get('criteria', '')
                            operator = f.get('operator', 'equals')
                            date_from = f.get('dateFrom', None)
                            date_to = f.get('dateTo', None)
                            
                            # ---- Text-Filter → Wildcard-Criteria ----
                            if operator == 'contains':
                                criteria_list.append(f'*{criteria}*')
                            elif operator == 'notContains':
                                criteria_list.append(f'<>*{criteria}*')
                            elif operator == 'startsWith':
                                criteria_list.append(f'{criteria}*')
                            elif operator == 'endsWith':
                                criteria_list.append(f'*{criteria}')
                            elif operator == 'isEmpty':
                                criteria_list.append('=')
                            elif operator == 'isNotEmpty':
                                criteria_list.append('<>')
                            elif operator == 'equals':
                                criteria_list.append(criteria)
                            
                            # ---- Datums-Filter ----
                            elif operator in ('dateToday', 'datePast', 'dateFuture',
                                              'dateThisWeek', 'dateThisMonth',
                                              'dateInDays', 'dateOverdueDays', 'dateBetween'):
                                is_date = True
                                self._log(f"Datums-Filter Spalte {col_idx}: op={operator}, from={date_from}, to={date_to}")
                                try:
                                    def _date_to_serial(iso_date_str):
                                        """Konvertiert ISO-Datum zu Excel-Seriennummer (Integer)"""
                                        dt = datetime.strptime(iso_date_str, '%Y-%m-%d')
                                        excel_epoch = datetime(1899, 12, 30)
                                        return (dt - excel_epoch).days
                                    
                                    # Prüfe ob Spalte echte Datumswerte enthält
                                    col_letter = self._get_column_letter(col_idx)
                                    cell_val = self.worksheet.range(f'{col_letter}2').value
                                    is_real_date = isinstance(cell_val, datetime)
                                    self._log(f"  Spalte {col_idx} ({col_letter}): Wert={cell_val}, Typ={type(cell_val).__name__}, is_real_date={is_real_date}")
                                    
                                    if is_real_date:
                                        # ===== Echte Datumsspalte =====
                                        # Strategie 1: xlDynamic (Operator=11) für einfache Fälle
                                        # → Das ist was Excel selbst in der UI nutzt, kein Datumsformat nötig.
                                        dynamic_map = {
                                            'dateToday': 1,      # xlFilterToday
                                            'dateThisWeek': 4,   # xlFilterThisWeek
                                            'dateThisMonth': 7,  # xlFilterThisMonth
                                        }
                                        
                                        if operator in dynamic_map:
                                            dyn_const = dynamic_map[operator]
                                            self._log(f"  xlDynamic: Operator=11, Criteria1={dyn_const}")
                                            used_range.api.AutoFilter(Field=col_idx, Operator=11, Criteria1=dyn_const)
                                            self._active_filter_fields.append(col_idx)
                                            self._active_filter_fields_per_sheet[current_sheet].append(col_idx)
                                            self._log(f"  Datums-Filter Spalte {col_idx} per xlDynamic gesetzt")
                                        else:
                                            # Strategie 2: Locale-formatierte Datums-Strings
                                            # AutoFilter erwartet Datums-Criteria im Format der
                                            # Excel-Application-Locale (Application.International).
                                            # Seriennummern funktionieren NICHT als Criteria.
                                            def _iso_to_filter_date(iso_str):
                                                dt = datetime.strptime(iso_str, '%Y-%m-%d')
                                                try:
                                                    app = self.worksheet.book.app
                                                    date_order = int(app.api.International(32))
                                                    date_sep = str(app.api.International(17))
                                                    self._log(f"  International: order={date_order}, sep='{date_sep}'")
                                                except Exception as ie:
                                                    self._log(f"  International FEHLER: {ie}, fallback DE")
                                                    date_order = 1
                                                    date_sep = '.'
                                                d = f"{dt.day:02d}"
                                                m = f"{dt.month:02d}"
                                                y = str(dt.year)
                                                if date_order == 0:    # MDY (US)
                                                    return f"{m}{date_sep}{d}{date_sep}{y}"
                                                elif date_order == 2:  # YMD
                                                    return f"{y}{date_sep}{m}{date_sep}{d}"
                                                else:                  # DMY (DE, UK, etc.)
                                                    return f"{d}{date_sep}{m}{date_sep}{y}"
                                            
                                            fmt_from = _iso_to_filter_date(date_from) if date_from else None
                                            fmt_to = _iso_to_filter_date(date_to) if date_to else None
                                            self._log(f"  Locale-Datum: from={fmt_from}, to={fmt_to}")
                                            
                                            if fmt_from and fmt_to:
                                                c1 = f">={fmt_from}"
                                                c2 = f"<={fmt_to}"
                                                self._log(f"  AutoFilter: c1='{c1}', Operator=xlAnd, c2='{c2}'")
                                                used_range.api.AutoFilter(Field=col_idx, Criteria1=c1, Operator=1, Criteria2=c2)
                                            elif fmt_from:
                                                c1 = f">={fmt_from}"
                                                self._log(f"  AutoFilter: c1='{c1}'")
                                                used_range.api.AutoFilter(Field=col_idx, Criteria1=c1)
                                            elif fmt_to:
                                                c1 = f"<={fmt_to}"
                                                self._log(f"  AutoFilter: c1='{c1}'")
                                                used_range.api.AutoFilter(Field=col_idx, Criteria1=c1)
                                            else:
                                                self._log(f"  Datums-Filter Spalte {col_idx} übersprungen (kein from/to)")
                                                continue
                                            self._active_filter_fields.append(col_idx)
                                            self._active_filter_fields_per_sheet[current_sheet].append(col_idx)
                                            self._log(f"  Datums-Filter Spalte {col_idx} gesetzt")
                                    else:
                                        # ===== Text-Datumsspalte =====
                                        # Formatierung muss zum Text in den Zellen passen
                                        def _format_date_for_text_col(iso_date_str, col_index):
                                            try:
                                                dt = datetime.strptime(iso_date_str, '%Y-%m-%d')
                                            except:
                                                return iso_date_str
                                            try:
                                                cl = self._get_column_letter(col_index)
                                                sample = str(self.worksheet.range(f'{cl}2').value or '')
                                                self._log(f"  Text-Datum Beispiel: '{sample}'")
                                                if '.' in sample and sample.count('.') == 2:
                                                    return dt.strftime('%d.%m.%Y')
                                                elif '/' in sample:
                                                    return dt.strftime('%m/%d/%Y')
                                                elif '-' in sample:
                                                    return dt.strftime('%Y-%m-%d')
                                            except:
                                                pass
                                            return dt.strftime('%d.%m.%Y')
                                        
                                        fmt_from = _format_date_for_text_col(date_from, col_idx) if date_from else None
                                        fmt_to = _format_date_for_text_col(date_to, col_idx) if date_to else None
                                        self._log(f"  Text-Format: from={fmt_from}, to={fmt_to}")
                                        
                                        c1 = None
                                        c2 = None
                                        if fmt_from and fmt_to:
                                            c1 = f">={fmt_from}"
                                            c2 = f"<={fmt_to}"
                                        elif fmt_from:
                                            c1 = f">={fmt_from}"
                                        elif fmt_to:
                                            c1 = f"<={fmt_to}"
                                        
                                        if not c1:
                                            self._log(f"  Text-Datums-Filter Spalte {col_idx} übersprungen")
                                            continue
                                        
                                        if c2:
                                            used_range.api.AutoFilter(Field=col_idx, Criteria1=c1, Operator=1, Criteria2=c2)
                                        else:
                                            used_range.api.AutoFilter(Field=col_idx, Criteria1=c1)
                                        self._active_filter_fields.append(col_idx)
                                        self._active_filter_fields_per_sheet[current_sheet].append(col_idx)
                                        self._log(f"  Text-Datums-Filter Spalte {col_idx} gesetzt: c1={c1}, c2={c2}")
                                except Exception as e:
                                    self._log(f"FEHLER bei Datums-Filter Spalte {col_idx}: {e}")
                                    import traceback
                                    self._log(traceback.format_exc())
                        
                        # ---- Text-Criteria anwenden (nach der for-Schleife) ----
                        if not is_date and criteria_list:
                            try:
                                # Sobald ein negativer Criteria dabei ist → AND (xlAnd=1), sonst OR (xlOr=2)
                                any_negative = any(c.startswith('<>') for c in criteria_list)
                                xl_operator = 1 if any_negative else 2  # xlAnd=1, xlOr=2
                                
                                if len(criteria_list) == 1:
                                    self._log(f"Filter Spalte {col_idx}: criteria='{criteria_list[0]}'")
                                    used_range.api.AutoFilter(Field=col_idx, Criteria1=criteria_list[0])
                                elif len(criteria_list) == 2:
                                    op_name = 'AND' if any_negative else 'OR'
                                    self._log(f"Filter Spalte {col_idx}: {op_name} '{criteria_list[0]}' | '{criteria_list[1]}'")
                                    used_range.api.AutoFilter(Field=col_idx, Criteria1=criteria_list[0], Operator=xl_operator, Criteria2=criteria_list[1])
                                else:
                                    op_name = 'AND' if any_negative else 'OR'
                                    self._log(f"Filter Spalte {col_idx}: {len(criteria_list)} Criteria — verwende erste 2 mit {op_name}")
                                    used_range.api.AutoFilter(Field=col_idx, Criteria1=criteria_list[0], Operator=xl_operator, Criteria2=criteria_list[1])
                                self._active_filter_fields.append(col_idx)
                                self._active_filter_fields_per_sheet[current_sheet].append(col_idx)
                            except Exception as e:
                                self._log(f"Fehler bei Filter Spalte {col_idx}: {e}")
                else:
                    # ===== Keine App-Filter mehr aktiv =====
                    # NUR die von der App gesetzten Filter-Felder zurücksetzen
                    # Vorhandene Excel-AutoFilter (z.B. beim Dateiöffnen) NICHT antasten!
                    if not hasattr(self, '_active_filter_fields_per_sheet'):
                        self._active_filter_fields_per_sheet = {}
                    current_sheet = self.sheet_name or ''
                    old_fields = self._active_filter_fields_per_sheet.get(current_sheet, [])
                    if not old_fields:
                        old_fields = getattr(self, '_active_filter_fields', [])
                    if old_fields:
                        self._log(f"Entferne {len(old_fields)} App-Filter-Felder: {old_fields}")
                        for col_idx in old_fields:
                            try:
                                # AutoFilter.Range statt used_range → funktioniert auch
                                # wenn alle Zeilen ausgeblendet sind (0 Treffer)
                                af = self.worksheet.api.AutoFilter
                                if af:
                                    af.Range.AutoFilter(Field=col_idx)
                                else:
                                    used_range.api.AutoFilter(Field=col_idx)
                                self._log(f"  App-Filter Feld {col_idx} zurückgesetzt")
                            except Exception as e:
                                self._log(f"  App-Filter Feld {col_idx} Reset-Fehler: {e}")
                        self._active_filter_fields = []
                        self._active_filter_fields_per_sheet[current_sheet] = []
                    else:
                        self._log("Keine App-Filter zum Zurücksetzen")
                        
            except Exception as api_error:
                self._log(f"AutoFilter API Fehler: {api_error}")
                return {'success': False, 'error': str(api_error)}
            finally:
                # === Performance-Guard Restore ===
                try:
                    app.enable_events = False  # Session hält events immer aus
                except Exception as restore_err:
                    self._log(f"  Performance-Guard Restore Fehler: {restore_err}")
            
            self._log(f"AutoFilter abgeschlossen: {len(filters) if filters else 0} Filter, aktive Felder: {self._active_filter_fields}")
            return {'success': True, 'filterCount': len(filters) if filters else 0, 'activeFields': list(self._active_filter_fields)}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen des AutoFilters: {e}")
            return {'success': False, 'error': str(e)}
    
    def clear_autofilter(self) -> Dict[str, Any]:
        """Entfernt alle AutoFilter und blendet alle Zeilen wieder ein.
        
        Eigenständige Methode (nicht über set_autofilter), damit
        das Zurücksetzen robust und unabhängig funktioniert.
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            self._log(f"clear_autofilter: Start")
            return self._clear_autofilter_windows()
        except Exception as e:
            self._log(f"clear_autofilter Fehler: {e}")
            return {'success': False, 'error': str(e)}
    
    def _clear_autofilter_windows(self) -> Dict[str, Any]:
        """Windows: AutoFilter-Kriterien entfernen und alle Zeilen einblenden.
        
        Unterstützt:
        - Worksheet-AutoFilter (normale AutoFilter auf dem Sheet)
        - Table/ListObject-AutoFilter (AutoFilter innerhalb von Excel-Tabellen)
        - Filter die über set_autofilter gesetzt wurden
        - Filter die bereits in der Excel-Datei vorhanden waren
        
        Robustheit: Wenn keine Zeilen dem Filter entsprechen (0 Treffer),
        kann ShowAllData() fehlschlagen. In diesem Fall werden die Filter
        direkt über AutoFilter(Field=N) einzeln entfernt.
        """
        try:
            used_range = self.worksheet.used_range
            if not used_range:
                return {'success': True, 'filterCount': 0}
            
            self._log(f"Windows: clear_autofilter Start, gespeicherte Felder: {self._active_filter_fields}")
            cleared_something = False
            
            # =====================================================
            # SCHRITT 1: Table/ListObject AutoFilter zurücksetzen
            # Excel-Tabellen haben ihren EIGENEN AutoFilter!
            # worksheet.AutoFilterMode zeigt diesen NICHT an.
            # WICHTIG: DataBodyRange kann None sein wenn ALLE Zeilen
            # gefiltert sind (0 Treffer) — daher ShowAllData über
            # Worksheet statt über DataBodyRange.Parent aufrufen.
            # =====================================================
            try:
                tables = self.worksheet.api.ListObjects
                if tables and tables.Count > 0:
                    for i in range(1, tables.Count + 1):
                        table = tables.Item(i)
                        table_name = table.Name
                        self._log(f"Windows: Table '{table_name}' gefunden")
                        try:
                            if table.AutoFilter:
                                af = table.AutoFilter
                                # Einzelne Filter-Felder zuerst zurücksetzen
                                # (robuster als ShowAllData, funktioniert auch bei 0 Treffern)
                                try:
                                    filters = af.Filters
                                    for fi in range(1, filters.Count + 1):
                                        try:
                                            if filters.Item(fi).On:
                                                table.Range.AutoFilter(Field=fi)
                                                self._log(f"Windows: Table '{table_name}' Filter Feld {fi} zurückgesetzt")
                                                cleared_something = True
                                        except:
                                            pass
                                except Exception as e:
                                    self._log(f"Windows: Table '{table_name}' Filter-Iteration Fehler: {e}")
                                
                                # ShowAllData als Sicherheitsnetz
                                try:
                                    if af.FilterMode:
                                        self.worksheet.api.ShowAllData()
                                        self._log(f"Windows: Table '{table_name}' ShowAllData erfolgreich")
                                        cleared_something = True
                                except Exception as e:
                                    self._log(f"Windows: Table '{table_name}' ShowAllData Fehler (ignoriert): {e}")
                        except Exception as e:
                            self._log(f"Windows: Table '{table_name}' AutoFilter-Zugriff Fehler: {e}")
                else:
                    self._log("Windows: Keine Tables/ListObjects vorhanden")
            except Exception as e:
                self._log(f"Windows: ListObjects-Check Fehler: {e}")
            
            # =====================================================
            # SCHRITT 2: Worksheet-AutoFilter zurücksetzen
            # (normaler AutoFilter, nicht Teil einer Tabelle)
            # WICHTIG: ShowAllData kann fehlschlagen wenn 0 Zeilen
            # sichtbar sind. In dem Fall Filter einzeln entfernen.
            # =====================================================
            try:
                if self.worksheet.api.AutoFilterMode:
                    self._log("Windows: Worksheet AutoFilterMode ist aktiv")
                    
                    # ERST: Einzelne Filter-Felder zurücksetzen (funktioniert immer)
                    try:
                        af = self.worksheet.api.AutoFilter
                        if af:
                            filters = af.Filters
                            for fi in range(1, filters.Count + 1):
                                try:
                                    if filters.Item(fi).On:
                                        used_range.api.AutoFilter(Field=fi)
                                        self._log(f"Windows: Worksheet Filter Feld {fi} zurückgesetzt")
                                        cleared_something = True
                                except:
                                    pass
                    except Exception as e:
                        self._log(f"Windows: Worksheet Filter-Iteration Fehler: {e}")
                    
                    # DANN: ShowAllData als Sicherheitsnetz
                    try:
                        self.worksheet.api.ShowAllData()
                        self._log("Windows: Worksheet ShowAllData() erfolgreich")
                        cleared_something = True
                    except Exception as e:
                        self._log(f"Windows: Worksheet ShowAllData Fehler (ignoriert): {e}")
                    
                    # AutoFilterMode deaktivieren
                    try:
                        self.worksheet.api.AutoFilterMode = False
                        self._log("Windows: AutoFilterMode auf False gesetzt")
                        cleared_something = True
                    except Exception as e:
                        self._log(f"Windows: AutoFilterMode=False Fehler: {e}")
                else:
                    self._log("Windows: Kein Worksheet-AutoFilter aktiv")
            except Exception as e:
                self._log(f"Windows: AutoFilter-Check Fehler: {e}")
            
            # =====================================================
            # SCHRITT 3: Fallback — gespeicherte Felder einzeln löschen
            # =====================================================
            if not cleared_something and self._active_filter_fields:
                self._log(f"Windows: Fallback — lösche {len(self._active_filter_fields)} gespeicherte Felder")
                for col_idx in self._active_filter_fields:
                    try:
                        used_range.api.AutoFilter(Field=col_idx)
                        self._log(f"Windows: Filter Feld {col_idx} zurückgesetzt (Fallback)")
                        cleared_something = True
                    except Exception as e2:
                        self._log(f"Windows: Filter Feld {col_idx} Fehler: {e2}")
            
            self._active_filter_fields = []
            
            # =====================================================
            # SCHRITT 4: Alle versteckten Zeilen einblenden (Sicherheitsnetz)
            # Falls ShowAllData nicht alle Zeilen eingeblendet hat
            # WICHTIG: Ein einziger COM-Aufruf statt Zeile-für-Zeile,
            # sonst Timeout bei großen Tabellen (>1000 Zeilen)
            # =====================================================
            try:
                # Methode 1: Gesamten used_range auf einmal einblenden
                used_range.api.EntireRow.Hidden = False
                self._log("Windows: Alle Zeilen eingeblendet (EntireRow.Hidden=False)")
            except Exception as e:
                self._log(f"Windows: EntireRow.Hidden Fehler: {e}")
                # Fallback: Rows-Collection auf einmal
                try:
                    self.worksheet.api.Rows.Hidden = False
                    self._log("Windows: Alle Zeilen eingeblendet (Rows.Hidden=False)")
                except Exception as e2:
                    self._log(f"Windows: Rows.Hidden Fehler: {e2}")
            
            # Schritt 5: Scroll-Position auf A1 zurücksetzen
            # Nach Filter-Reset springt Excel oft auf einen leeren Bereich
            # am Ende der Tabelle. Daher explizit auf A1 scrollen.
            try:
                self.worksheet.api.Range("A1").Select()
                app = self.workbook.app.api
                app.ActiveWindow.ScrollRow = 1
                app.ActiveWindow.ScrollColumn = 1
                self._log("Windows: Scroll-Position auf A1 zurückgesetzt")
            except Exception as e:
                self._log(f"Windows: Scroll-Reset Fehler (ignoriert): {e}")
            
            # Schritt 6: Screen-Refresh
            try:
                app = self.workbook.app.api
                app.ScreenUpdating = True
                app.Calculate()
            except:
                pass
            
            self._log(f"Windows: clear_autofilter OK (cleared_something={cleared_something})")
            return {'success': True, 'filterCount': 0}
            
        except Exception as e:
            self._log(f"Windows: clear_autofilter Fehler: {e}")
            return {'success': False, 'error': str(e)}
    
    def switch_sheet(self, sheet_name: str, include_data: bool = False) -> Dict[str, Any]:
        """Wechselt das aktive Arbeitsblatt in der Live Session
        
        Args:
            sheet_name: Name des Zielblatts
            include_data: Wenn True, werden die Sheet-Daten direkt mitgeliefert
                          (spart einen separaten getData-Roundtrip)
        
        WICHTIG: Der visuelle Excel-Wechsel (activate/Goto) passiert ZULETZT,
        NACH dem Datenlesen. So blockieren die langen COM-Leseoperationen
        nicht das Neuzeichnen von Excels Fenster.
        """
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            target_sheet = self.workbook.sheets[sheet_name]
            app = self.workbook.app.api
            
            # Prüfe ob Sheet ausgeblendet ist
            was_hidden = not target_sheet.visible
            if was_hidden:
                try:
                    app.ScreenUpdating = False
                    app.EnableEvents = False
                    target_sheet.visible = True
                finally:
                    try:
                        app.EnableEvents = False
                        app.ScreenUpdating = False
                    except Exception:
                        pass
                self._log(f"Sheet '{sheet_name}' war ausgeblendet → automatisch eingeblendet")
            
            # Python-Referenz auf Ziel-Sheet setzen (KEIN visueller Wechsel!)
            # get_data() und hidden-rows lesen verwenden self.worksheet als Referenz,
            # brauchen kein aktives Sheet → können VOR dem visuellen Wechsel laufen.
            self.worksheet = target_sheet
            self.sheet_name = sheet_name
            
            # CF-Erkennung (liest nur von der Worksheet-Referenz)
            has_cf = False
            try:
                cf_count = self.worksheet.api.Cells.FormatConditions.Count
                has_cf = cf_count > 0
                if has_cf:
                    self._log(f"Sheet '{sheet_name}' hat {cf_count} Conditional Formatting Regeln")
                    try:
                        self.worksheet.api.EnableFormatConditionsCalculation = False
                    except Exception:
                        pass
            except Exception:
                pass
            
            result = {'success': True, 'sheetName': sheet_name, 'wasHidden': was_hidden, 'hasConditionalFormatting': has_cf}
            
            # === PHASE 1: Daten lesen (VOR dem visuellen Wechsel) ===
            # Alle Lese-Operationen verwenden self.worksheet als Referenz.
            # Excel zeigt noch das alte Sheet → kein Flackern, kein blockierter Repaint.
            if include_data:
                try:
                    data_result = self.get_data()
                    result['headers'] = data_result.get('headers', [])
                    result['data'] = data_result.get('data', [])
                except Exception as data_err:
                    self._log(f"Daten nach Sheet-Wechsel konnten nicht gelesen werden: {data_err}")
                    result['dataError'] = str(data_err)
                
                # Versteckte Zeilen/Spalten direkt von Excel lesen (COM)
                try:
                    hidden_rows = []
                    hidden_cols = []
                    used = self.worksheet.used_range
                    if used:
                        last_row = used.last_cell.row
                        last_col = used.last_cell.column
                        
                        for c in range(1, last_col + 1):
                            try:
                                if self.worksheet.api.Columns(c).Hidden:
                                    hidden_cols.append(c - 1)
                            except Exception:
                                pass
                        
                        if last_row >= 2:
                            try:
                                data_range = self.worksheet.api.Range(
                                    self.worksheet.api.Rows(2),
                                    self.worksheet.api.Rows(last_row)
                                )
                                row_height = data_range.RowHeight
                                if row_height is None or data_range.Hidden is None:
                                    for r in range(2, last_row + 1):
                                        try:
                                            if self.worksheet.api.Rows(r).Hidden:
                                                hidden_rows.append(r - 2)
                                        except Exception:
                                            pass
                                elif data_range.Hidden:
                                    hidden_rows = list(range(0, last_row - 1))
                            except Exception:
                                for r in range(2, last_row + 1):
                                    try:
                                        if self.worksheet.api.Rows(r).Hidden:
                                            hidden_rows.append(r - 2)
                                    except Exception:
                                        pass
                    
                    result['hiddenRows'] = hidden_rows
                    result['hiddenColumns'] = hidden_cols
                    if hidden_rows or hidden_cols:
                        self._log(f"Versteckt: {len(hidden_rows)} Zeilen, {len(hidden_cols)} Spalten (COM)")
                except Exception as vis_err:
                    self._log(f"Versteckte Zeilen/Spalten konnten nicht gelesen werden: {vis_err}")
                    result['hiddenRows'] = []
                    result['hiddenColumns'] = []
            
            # KEIN visueller Wechsel hier! Das macht activate_sheet() NACH dem Laden.
            # Grund: Interactive=False (Read-Only-Schutz) blockiert jedes visuelle Update.
            # activate_sheet() hebt Interactive kurz auf, macht activate(), wartet, setzt zurück.
            self._log(f"Sheet gewechselt zu: {sheet_name} (nur Daten)")
            return result
            
        except Exception as e:
            self._log(f"Fehler beim Sheet-Wechsel: {e}")
            return {'success': False, 'error': str(e)}
    
    def activate_sheet(self, sheet_name: str) -> Dict[str, Any]:
        """Visueller Sheet-Wechsel in Excel — SEPARATER Befehl nach dem Datenladen.
        
        Wird vom Frontend aufgerufen NACHDEM die GUI gerendert wurde.
        Interactive=True ist ZWINGEND nötig — ohne das zeichnet Excel nichts.
        Das Frontend cancelt VOR dem Aufruf alle Batch-Syncs, sodass keine
        parallelen COM-Operationen (setCellsBatch) laufen können.
        """
        import time as _time
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            app = self.workbook.app.api
            
            is_visible = False
            try:
                is_visible = self.app.visible
            except Exception:
                pass
            
            target_sheet = self.workbook.sheets[sheet_name]
            
            if is_visible:
                # Sichtbar → Interactive=True nötig damit Excel den Tab-Wechsel zeichnet
                try:
                    app.Interactive = True
                    target_sheet.activate()
                    _time.sleep(0.3)
                except Exception as act_err:
                    self._log(f"activate_sheet Fehler: {act_err}")
                finally:
                    try:
                        app.Interactive = False
                    except Exception:
                        pass
            else:
                # Versteckt → activate() setzt COM-State, kein Interactive nötig
                try:
                    target_sheet.activate()
                except Exception as act_err:
                    self._log(f"activate_sheet (hidden) Fehler: {act_err}")
            
            self._log(f"activate_sheet: {sheet_name} visuell aktiviert")
            return {'success': True}
            
        except Exception as e:
            self._log(f"Fehler bei activate_sheet: {e}")
            return {'success': False, 'error': str(e)}
    
    def set_sheet_visibility(self, sheet_name: str, visible: bool) -> Dict[str, Any]:
        """Setzt die Sichtbarkeit eines Arbeitsblatts
        
        Args:
            sheet_name: Name des Arbeitsblatts
            visible: True = einblenden, False = ausblenden
        """
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            # Read-Only-Prüfung: Im schreibgeschützten Modus können keine Änderungen gemacht werden
            try:
                if self.workbook.api.ReadOnly:
                    return {'success': False, 'error': 'Die Datei ist schreibgeschützt (wird von einem anderen Prozess verwendet). Bitte schließen Sie die Datei in der Haupt-GUI oder in anderen Programmen.'}
            except Exception:
                pass
            
            # Mindestens ein Sheet muss sichtbar bleiben
            if not visible:
                visibility_info = [(s.name, s.visible) for s in self.workbook.sheets]
                visible_count = sum(1 for _, v in visibility_info if v)
                self._log(f"[DEBUG] set_sheet_visibility: sheet='{sheet_name}', visible={visible}, visibility_info={visibility_info}, visible_count={visible_count}")
                if visible_count <= 1:
                    return {'success': False, 'error': 'Mindestens ein Arbeitsblatt muss sichtbar bleiben'}
            
            target_sheet = self.workbook.sheets[sheet_name]
            app = self.workbook.app.api
            try:
                # ScreenUpdating + EnableEvents deaktivieren um Excel-Hänger zu vermeiden
                app.ScreenUpdating = False
                app.EnableEvents = False
                
                # Beim Ausblenden des aktiven Sheets: erst ein anderes Sheet aktivieren
                if not visible:
                    try:
                        active_name = self.workbook.app.api.ActiveSheet.Name
                    except Exception:
                        active_name = None
                    if active_name == sheet_name:
                        # Erstes sichtbares anderes Sheet aktivieren
                        for s in self.workbook.sheets:
                            if s.name != sheet_name and s.visible:
                                s.activate()
                                self._log(f"Aktives Sheet gewechselt zu '{s.name}' vor Ausblenden von '{sheet_name}'")
                                break
                
                target_sheet.visible = visible
            except Exception as write_err:
                err_msg = str(write_err)
                if 'read-only' in err_msg.lower() or 'schreibgeschützt' in err_msg.lower() or '0x800A03EC' in err_msg:
                    return {'success': False, 'error': 'Die Datei ist schreibgeschützt. Wird sie von einem anderen Programm oder in der Haupt-GUI verwendet?'}
                raise
            finally:
                try:
                    app.EnableEvents = True
                    app.ScreenUpdating = True
                except Exception:
                    pass
            
            action_text = "eingeblendet" if visible else "ausgeblendet"
            self._log(f"Sheet '{sheet_name}' {action_text}")
            return {'success': True, 'sheetName': sheet_name, 'visible': visible}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen der Sichtbarkeit: {e}")
            return {'success': False, 'error': str(e)}
    
    def add_sheet(self, sheet_name: str) -> Dict[str, Any]:
        """F\u00fcgt ein neues Arbeitsblatt hinzu (Live-Session)"""
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            # Pr\u00fcfe ob Name bereits existiert
            existing = [s.name for s in self.workbook.sheets]
            if sheet_name in existing:
                return {'success': False, 'error': 'Ein Arbeitsblatt mit diesem Namen existiert bereits'}
            
            self.workbook.sheets.add(name=sheet_name, after=self.workbook.sheets[-1])
            
            self._log(f"Sheet '{sheet_name}' hinzugef\u00fcgt")
            sheets = [s.name for s in self.workbook.sheets]
            return {'success': True, 'sheets': sheets}
            
        except Exception as e:
            self._log(f"Fehler beim Hinzuf\u00fcgen: {e}")
            return {'success': False, 'error': str(e)}
    
    def delete_sheet(self, sheet_name: str) -> Dict[str, Any]:
        """L\u00f6scht ein Arbeitsblatt (Live-Session)"""
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            if len(self.workbook.sheets) <= 1:
                return {'success': False, 'error': 'Das letzte Arbeitsblatt kann nicht gel\u00f6scht werden'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            # Excel-Warnmeldungen tempor\u00e4r unterdr\u00fccken
            self.app.display_alerts = False
            self.workbook.sheets[sheet_name].delete()
            self.app.display_alerts = True
            
            # Falls das aktive Sheet gel\u00f6scht wurde, zum ersten wechseln
            remaining = [s.name for s in self.workbook.sheets]
            if self.sheet_name == sheet_name and remaining:
                self.worksheet = self.workbook.sheets[remaining[0]]
                self.sheet_name = remaining[0]
            
            self._log(f"Sheet '{sheet_name}' gel\u00f6scht")
            return {'success': True, 'sheets': remaining}
            
        except Exception as e:
            self.app.display_alerts = True
            self._log(f"Fehler beim L\u00f6schen: {e}")
            return {'success': False, 'error': str(e)}
    
    def rename_sheet(self, old_name: str, new_name: str) -> Dict[str, Any]:
        """Benennt ein Arbeitsblatt um (Live-Session)"""
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if old_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{old_name}" nicht gefunden'}
            if new_name in sheet_names:
                return {'success': False, 'error': 'Ein Arbeitsblatt mit diesem Namen existiert bereits'}
            
            self.workbook.sheets[old_name].name = new_name
            
            # Falls das aktive Sheet umbenannt wurde
            if self.sheet_name == old_name:
                self.sheet_name = new_name
                self.worksheet = self.workbook.sheets[new_name]
            
            self._log(f"Sheet '{old_name}' umbenannt zu '{new_name}'")
            sheets = [s.name for s in self.workbook.sheets]
            return {'success': True, 'sheets': sheets}
            
        except Exception as e:
            self._log(f"Fehler beim Umbenennen: {e}")
            return {'success': False, 'error': str(e)}
    
    def clone_sheet(self, sheet_name: str, new_name: str) -> Dict[str, Any]:
        """Kopiert ein Arbeitsblatt (Live-Session)"""
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            if new_name in sheet_names:
                return {'success': False, 'error': 'Ein Arbeitsblatt mit diesem Namen existiert bereits'}
            
            source = self.workbook.sheets[sheet_name]
            source.copy(after=self.workbook.sheets[-1], name=new_name)
            
            self._log(f"Sheet '{sheet_name}' kopiert als '{new_name}'")
            sheets = [s.name for s in self.workbook.sheets]
            return {'success': True, 'sheets': sheets}
            
        except Exception as e:
            self._log(f"Fehler beim Kopieren: {e}")
            return {'success': False, 'error': str(e)}
    
    def move_sheet(self, sheet_name: str, new_index: int) -> Dict[str, Any]:
        """Verschiebt ein Arbeitsblatt an eine neue Position (Live-Session)"""
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            num_sheets = len(self.workbook.sheets)
            if new_index < 0 or new_index >= num_sheets:
                return {'success': False, 'error': f'Ungültiger Index: {new_index}'}
            
            current_index = sheet_names.index(sheet_name)
            if current_index == new_index:
                sheets = [s.name for s in self.workbook.sheets]
                return {'success': True, 'sheets': sheets}
            
            # Direkt die COM-API mit Worksheets(name) verwenden statt gecachte
            # xlwings-Referenzen — diese werden durch Sichtbarkeitsänderungen
            # ungültig (RPC_E_DISCONNECTED).
            wb_api = self.workbook.api
            
            # Versteckte Referenz-Sheets temporär sichtbar machen
            # (Excel COM ignoriert Move(Before/After=hiddenSheet) stillschweigend)
            temporarily_shown_names = []
            try:
                if new_index == 0:
                    ref_name = sheet_names[0]
                    ref_ws = wb_api.Worksheets(ref_name)
                    if ref_ws.Visible != -1:  # xlSheetVisible = -1
                        ref_ws.Visible = -1
                        temporarily_shown_names.append(ref_name)
                    wb_api.Worksheets(sheet_name).Move(Before=wb_api.Worksheets(ref_name))
                else:
                    target_name = sheet_names[new_index]
                    target_ws = wb_api.Worksheets(target_name)
                    if target_ws.Visible != -1:
                        target_ws.Visible = -1
                        temporarily_shown_names.append(target_name)
                    if current_index < new_index:
                        wb_api.Worksheets(sheet_name).Move(After=wb_api.Worksheets(target_name))
                    else:
                        wb_api.Worksheets(sheet_name).Move(Before=wb_api.Worksheets(target_name))
            finally:
                for name in temporarily_shown_names:
                    try:
                        wb_api.Worksheets(name).Visible = 0  # xlSheetHidden = 0
                    except Exception:
                        pass
            
            self._log(f"Sheet '{sheet_name}' verschoben zu Index {new_index}")
            sheets = [s.name for s in self.workbook.sheets]
            return {'success': True, 'sheets': sheets}
            
        except Exception as e:
            self._log(f"Fehler beim Verschieben: {e}")
            import traceback
            self._log(traceback.format_exc())
            return {'success': False, 'error': str(e)}
    
    def _format_datetime_values(self, all_data: list, used_range) -> list:
        """Konvertiert datetime-Werte zu Strings (DD.MM.YYYY).
        
        Kein COM-Aufruf, kein .api.Text, kein Format-Ableitung.
        Pure Python strftime — maximal zuverlässig.
        """
        if not all_data:
            return all_data
        
        # Sonderfall: Nur eine Zeile (keine verschachtelte Liste)
        if not isinstance(all_data[0], list):
            for i, val in enumerate(all_data):
                if isinstance(val, (datetime, date)):
                    all_data[i] = self._default_date_str(val)
            return all_data
        
        for row in all_data:
            if not isinstance(row, list):
                continue
            for i, val in enumerate(row):
                if isinstance(val, (datetime, date)):
                    row[i] = self._default_date_str(val)
        
        return all_data
    
    @staticmethod
    def _default_date_str(val) -> str:
        """Einfache Datum-zu-String Konvertierung."""
        try:
            if isinstance(val, datetime):
                if val.hour == 0 and val.minute == 0 and val.second == 0:
                    return val.strftime('%d.%m.%Y')
                return val.strftime('%d.%m.%Y %H:%M:%S')
            if isinstance(val, date):
                return val.strftime('%d.%m.%Y')
            return str(val)
        except Exception:
            return str(val)
    
    def _apply_number_formats(self, all_data: list, used_range) -> list:
        """Ersetzt Zahlen durch exakt den Text den Excel anzeigt (.Text).
        
        Strategie (kein NumberFormat-Parsing, direkte .Text-Prüfung):
        1. Alle Spalten mit Zahlen auf Breite 255 setzen (verhindert '###')
        2. Pro Spalte: .Text EINER Zelle lesen und mit str(Rohwert) vergleichen
           → Wenn unterschiedlich: Spalte braucht .Text für alle Zellen
        3. .Text für alle numerischen Zellen in erkannten Spalten lesen
        4. Spaltenbreiten wiederherstellen
        
        WICHTIG: Excel darf dabei NICHT sichtbar werden.
        """
        self._log(f"_apply_number_formats START: rows={len(all_data) if all_data else 0}, is_list={isinstance(all_data[0], list) if all_data else 'N/A'}")
        
        if not all_data or not isinstance(all_data[0], list):
            self._log("_apply_number_formats SKIP: Keine verschachtelten Listen")
            return all_data
        
        num_rows = len(all_data)
        num_cols = len(all_data[0]) if all_data else 0
        if num_cols == 0 or num_rows <= 1:
            self._log(f"_apply_number_formats SKIP: cols={num_cols}, rows={num_rows}")
            return all_data
        
        ws_api = self.worksheet.api
        app_api = ws_api.Application
        start_row = used_range.row
        start_col = used_range.column
        
        # Sichtbarkeit VOR allen COM-Operationen sichern und erzwingen
        was_visible = True
        try:
            was_visible = bool(app_api.Visible)
            if not was_visible:
                app_api.Visible = False  # explizit nochmal setzen
        except Exception:
            pass
        
        # ScreenUpdating deaktivieren (verhindert Flackern bei Breitenänderung)
        try:
            app_api.ScreenUpdating = False
        except Exception:
            pass
        
        try:
            # Spalten mit Zahlen finden + erste Zahlenzeile merken
            numeric_cols = {}  # col_idx -> first_number_row
            for col_idx in range(num_cols):
                for row_idx in range(1, num_rows):
                    if isinstance(all_data[row_idx][col_idx], (int, float)):
                        numeric_cols[col_idx] = row_idx
                        break
            
            if not numeric_cols:
                self._log("_apply_number_formats: KEINE numerischen Spalten gefunden!")
                # Zeige Typen der ersten Datenzeile
                if num_rows > 1:
                    types = [(i, type(v).__name__, repr(v)[:50]) for i, v in enumerate(all_data[1]) if v not in ('', None)]
                    self._log(f"  Zeile 1 Typen (nicht-leer): {types[:10]}")
                return all_data
            
            self._log(f"_apply_number_formats: {len(numeric_cols)} numerische Spalten gefunden: {dict(list(numeric_cols.items())[:10])}")
            
            # Phase 1: Spaltenbreiten sichern und ALLE numerischen Spalten auf 255 setzen
            saved_widths = {}
            for col_idx in numeric_cols:
                excel_col = start_col + col_idx
                try:
                    saved_widths[col_idx] = ws_api.Columns(excel_col).ColumnWidth
                    ws_api.Columns(excel_col).ColumnWidth = 255
                except Exception:
                    pass
            
            # Phase 2: Erkennung — .Text vs. Rohwert für EINE Zelle pro Spalte
            formatted_cols = []
            for col_idx, first_row in numeric_cols.items():
                try:
                    excel_row = start_row + first_row
                    excel_col = start_col + col_idx
                    cell_text = str(ws_api.Cells(excel_row, excel_col).Text or '').strip()
                    raw_val = all_data[first_row][col_idx]
                    
                    # Normalisieren: Excel zeigt 1.0 als "1" im General-Format
                    if isinstance(raw_val, float) and raw_val == int(raw_val) and abs(raw_val) < 1e15:
                        normalized = str(int(raw_val))
                    else:
                        normalized = str(raw_val)
                    
                    if cell_text != normalized:
                        formatted_cols.append(col_idx)
                        self._log(f"Spalte {col_idx}: raw='{normalized}' display='{cell_text}' → .Text nötig")
                except Exception as e:
                    self._log(f"Spalte {col_idx}: Erkennungsfehler: {e}")
            
            if not formatted_cols:
                self._log(f"_apply_number_formats: KEINE formatierten Spalten erkannt (alle {len(numeric_cols)} Spalten stimmen überein)")
                # Breiten wiederherstellen
                for col_idx, w in saved_widths.items():
                    try:
                        ws_api.Columns(start_col + col_idx).ColumnWidth = w
                    except Exception:
                        pass
                return all_data
            
            self._log(f"Formatierte Spalten: {formatted_cols} von {num_cols} Gesamt")
            
            # Phase 3: .Text für alle numerischen Zellen in erkannten Spalten lesen
            for col_idx in formatted_cols:
                excel_col = start_col + col_idx
                for row_idx in range(1, num_rows):
                    val = all_data[row_idx][col_idx]
                    if not isinstance(val, (int, float)):
                        continue
                    try:
                        cell_text = ws_api.Cells(start_row + row_idx, excel_col).Text
                        if cell_text and not str(cell_text).startswith('#'):
                            all_data[row_idx][col_idx] = str(cell_text).strip()
                    except Exception:
                        pass  # Bei Fehler Rohwert behalten
            
            # Phase 4: Spaltenbreiten wiederherstellen
            for col_idx, w in saved_widths.items():
                try:
                    ws_api.Columns(start_col + col_idx).ColumnWidth = w
                except Exception:
                    pass
            
            return all_data
        finally:
            try:
                app_api.ScreenUpdating = True
            except Exception:
                pass
            # Excel MUSS versteckt bleiben wenn es vorher versteckt war
            if not was_visible:
                try:
                    app_api.Visible = False
                except Exception:
                    pass
    
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
            
            # Datetime-Werte durch formatierten Text ersetzen
            # Liest number_format pro Spalte (1 COM-Aufruf pro Datumsspalte)
            try:
                all_data = self._format_datetime_values(all_data, used_range)
            except Exception as fmt_err:
                self._log(f"Datum-Formatierung Fehler (Fallback auf str): {fmt_err}")
            
            # Zahlenformatierung (Währung, Prozent, k€ etc.) wird jetzt auf der
            # JS-Seite per SSF erledigt (styles.xml aus ZIP + ssf-Library).
            # _apply_number_formats() mit 200k+ einzelnen COM-.Text-Aufrufen
            # war der Hauptgrund für 10–20 Sek. Verzögerung bei großen Sheets.
            self._log(f"get_data: Zahlenformatierung übersprungen (SSF auf JS-Seite)")
            
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
            'open': lambda: self.open_file(cmd.get('filePath'), cmd.get('sheetName'), cmd.get('password')),
            'save': lambda: self.save_file(cmd.get('outputPath'), cmd.get('password'), cmd.get('selectedSheets')),
            'close': lambda: self.close_session(save=cmd.get('save', False)),
            'getData': lambda: self.get_data(),
            'switchSheet': lambda: self.switch_sheet(cmd.get('sheetName'), cmd.get('includeData', False)),
            'activateSheet': lambda: self.activate_sheet(cmd.get('sheetName')),
            'setSheetVisibility': lambda: self.set_sheet_visibility(cmd.get('sheetName'), cmd.get('visible', True)),
            'addSheet': lambda: self.add_sheet(cmd.get('sheetName')),
            'deleteSheet': lambda: self.delete_sheet(cmd.get('sheetName')),
            'renameSheet': lambda: self.rename_sheet(cmd.get('oldName'), cmd.get('newName')),
            'cloneSheet': lambda: self.clone_sheet(cmd.get('sheetName'), cmd.get('newName')),
            'moveSheet': lambda: self.move_sheet(cmd.get('sheetName'), cmd.get('newIndex')),
            
            # Passwort
            'setPassword': lambda: self.set_password(cmd.get('password')),
            'getPasswordStatus': lambda: self.get_password_status(),
            
            # Zeilen
            'deleteRow': lambda: self.delete_row(cmd.get('rowIndex')),
            'insertRow': lambda: self.insert_row(cmd.get('rowIndex'), cmd.get('count', 1)),
            'moveRow': lambda: self.move_row(cmd.get('fromIndex'), cmd.get('toIndex')),
            'hideRow': lambda: self.hide_row(cmd.get('rowIndex'), cmd.get('hidden', True)),
            'hideRowsBatch': lambda: self.hide_rows_batch(cmd.get('rowIndices', []), cmd.get('hidden', True)),
            'highlightRow': lambda: self.highlight_row(cmd.get('rowIndex'), cmd.get('color')),
            'highlightRowsBatch': lambda: self.highlight_rows_batch(cmd.get('rows', []), cmd.get('color')),
            
            # Spalten
            'deleteColumn': lambda: self.delete_column(cmd.get('colIndex')),
            'insertColumn': lambda: self.insert_column(cmd.get('colIndex'), cmd.get('count', 1), cmd.get('headers')),
            'dataJoinSync': lambda: self.data_join_sync(cmd.get('operations', [])),
            'moveColumn': lambda: self.move_column(cmd.get('fromIndex'), cmd.get('toIndex')),
            'hideColumn': lambda: self.hide_column(cmd.get('colIndex'), cmd.get('hidden', True)),
            
            # Zellen
            'setCellValue': lambda: self.set_cell_value(cmd.get('rowIndex'), cmd.get('colIndex'), cmd.get('value'), old_value=cmd.get('oldValue')),
            'setColumnValues': lambda: self.set_column_values(cmd.get('colIndex'), cmd.get('values', []), cmd.get('startRow', 0)),
            'setRowValues': lambda: self.set_row_values(cmd.get('rowIndex'), cmd.get('values', [])),
            'setCellsBatch': lambda: self.set_cells_batch(cmd.get('cells', [])),
            'copyCells': lambda: self.copy_cells(cmd.get('sourceCells', []), cmd.get('targetRow'), cmd.get('targetCol')),
            'findReplace': lambda: self.find_replace(
                cmd.get('searchText', ''),
                cmd.get('replaceText', ''),
                cmd.get('matchCase', False),
                cmd.get('wholeWord', False)
            ),
            
            # Filter
            'setAutoFilter': lambda: self.set_autofilter(cmd.get('filters')),
            'clearAutoFilter': lambda: self.clear_autofilter(),
            
            # Undo
            'undo': lambda: self.undo(),
            
            # Session
            'ping': lambda: {'success': True, 'pong': True},
            'checkAlive': lambda: self.check_alive(),
            'initApp': lambda: self.init_app(),
            'quit': lambda: self._quit(),
            'setVisible': lambda: self.set_visible(cmd.get('visible', True)),
            
            # Recovery
            'getRecoveryFiles': lambda: self.get_recovery_files(),
            'deleteRecoveryFile': lambda: self.delete_recovery_file(cmd.get('filePath')),
        }
        
        handler = handlers.get(action)
        if handler:
            return handler()
        else:
            return {'success': False, 'error': f'Unbekannte Aktion: {action}'}
    
    def _quit(self) -> Dict[str, Any]:
        """Beendet die Session"""
        self._is_running = False
        # Journal flushen vor dem Beenden
        self._flush_journal()
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
                
                action = cmd.get('action', '?')
                cmd_id = cmd.get('_cmdId')
                start_time = time.time()
                self._log(f"CMD empfangen: {action}" + (f" [cmd#{cmd_id}]" if cmd_id is not None else ""))
                
                result = self.handle_command(cmd)
                
                # Command-ID zurückgeben (für Response-Matching im Bridge)
                if cmd_id is not None:
                    result['_cmdId'] = cmd_id
                
                elapsed = time.time() - start_time
                self._log(f"CMD fertig: {action} in {elapsed:.3f}s (success={result.get('success', '?')})")
                
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
    atexit.register(session._cleanup_all_undo_files)
    session.run()


if __name__ == '__main__':
    main()
