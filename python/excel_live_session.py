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
        if platform.system() == 'Darwin':
            base = os.path.expanduser('~/Library/Application Support/ExcelDataSyncPro')
        elif platform.system() == 'Windows':
            base = os.path.join(os.environ.get('APPDATA', ''), 'ExcelDataSyncPro')
        else:
            base = os.path.expanduser('~/.exceldatasyncpro')
        
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
    
    def set_visible(self, visible: bool = True) -> Dict[str, Any]:
        """Zeigt oder versteckt das Excel-Fenster"""
        try:
            if not self.app:
                self._log("Keine Excel-App aktiv")
                return {'success': False, 'error': 'Keine Excel-App aktiv'}
            
            # xlwings visible-Eigenschaft verwenden (funktioniert auf Mac und Windows)
            self.app.visible = visible
            self._log(f"Excel Sichtbarkeit gesetzt: {visible}")
            
            # Auf macOS: Falls visible=True, Excel in den Vordergrund bringen
            if visible and platform.system() == 'Darwin':
                try:
                    import subprocess
                    subprocess.run(['osascript', '-e', 'tell application "Microsoft Excel" to activate'], 
                                   capture_output=True, timeout=2)
                except:
                    pass
            
            return {'success': True, 'visible': visible}
        except Exception as e:
            self._log(f"Fehler bei set_visible: {e}")
            return {'success': False, 'error': str(e)}
    
    def _force_screen_refresh(self):
        """Erzwingt einen Screen-Refresh in Excel"""
        try:
            if not self.app:
                return
            
            # Screen-Updating sicherstellen
            self.app.screen_updating = True
            
            # Auf macOS: Aggressiveres Refresh nötig
            if platform.system() == 'Darwin':
                try:
                    # Workbook und Worksheet aktivieren, damit Excel die Änderung anzeigt
                    if self.workbook:
                        self.workbook.activate()
                    if self.worksheet:
                        self.worksheet.activate()
                    
                    # Formeln neu berechnen (erzwingt auch Display-Update)
                    self.app.calculate()
                except Exception as mac_err:
                    self._log(f"_force_screen_refresh macOS-Fehler: {mac_err}")
            else:
                # Windows: Aggressiveres Refresh - Toggle allein reicht nicht immer
                try:
                    if self.workbook:
                        self.workbook.activate()
                    if self.worksheet:
                        self.worksheet.activate()
                    self.app.screen_updating = False
                    self.app.screen_updating = True
                    self.app.calculate()
                except Exception as win_err:
                    self._log(f"_force_screen_refresh Windows-Fehler: {win_err}")
                
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
            
            # Sheet finden
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            self.worksheet = self.workbook.sheets[sheet_name]
            self.sheet_name = sheet_name
            
            # Recovery-System initialisieren
            self.backup_path = self._create_backup(file_path)
            self.journal_path = self._init_journal(file_path)
            self.last_auto_save = _time.time()
            
            # Undo-Stack leeren bei neuer Datei
            self.undo_stack.clear()
            
            _total = _time.time() - _t0
            self._log(f"Datei geöffnet in {_total:.1f}s, Sheet: {sheet_name}, Sheets: {sheet_names}")
            return {'success': True, 'sheets': sheet_names, 'backupPath': self.backup_path}
            
        except Exception as e:
            self._log(f"Fehler beim Öffnen nach {_time.time() - _t0:.1f}s: {e}")
            return {'success': False, 'error': str(e)}
    
    def save_file(self, output_path: Optional[str] = None, password: Optional[str] = None) -> Dict[str, Any]:
        """Speichert die Datei (optional unter neuem Namen und/oder mit Passwort)
        
        Args:
            output_path: Optionaler neuer Dateipfad
            password: Optionales Passwort (None = kein Passwort, '' = Passwort entfernen, 'xxx' = neues Passwort)
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
            
            is_windows = platform.system() == 'Windows'
            
            if output_path and output_path != self.file_path:
                self._log(f"Speichere unter: {output_path}")
                
                if platform.system() == 'Darwin':
                    # macOS: SaveAs über Speichern + Kopieren
                    self.workbook.save()
                    
                    shutil.copy2(self.file_path, output_path)
                    
                    # Wenn Quelldatei verschlüsselt war UND wir das Passwort NICHT behalten wollen,
                    # müssen wir die Kopie entschlüsseln
                    if self.file_password and not keep_password:
                        try:
                            import msoffcrypto
                            import tempfile
                            
                            with open(output_path, 'rb') as f:
                                file = msoffcrypto.OfficeFile(f)
                                file.load_key(password=self.file_password)
                                
                                temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
                                with os.fdopen(temp_fd, 'wb') as temp_f:
                                    file.decrypt(temp_f)
                            
                            shutil.move(temp_path, output_path)
                        except Exception as decrypt_err:
                            self._log(f"Fehler beim Entschlüsseln: {decrypt_err}")
                else:
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
                if is_windows:
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
                else:
                    # macOS: xlwings save funktioniert direkt
                    self.workbook.save()
            
            # Passwort aktualisieren
            if password is not None:
                self.file_password = password if password else None
            
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
                if platform.system() == 'Windows':
                    # Windows: COM-API direkt verwenden
                    self.workbook.api.Password = password
                    self.workbook.api.Save()
                else:
                    self.workbook.save(password=password)
                self.file_password = password
            else:
                self._log("Entferne Passwort...")
                if platform.system() == 'Windows':
                    # Windows: Passwort über COM-API leeren
                    self.workbook.api.Password = ''
                    self.workbook.api.Save()
                else:
                    # macOS: Speichern ohne Passwort entfernt das Passwort
                    self.workbook.save()
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
            import tempfile
            # Gleiche Dateiendung wie Original verwenden
            _, ext = os.path.splitext(self.file_path)
            if not ext:
                ext = '.xlsx'
            
            # Eindeutiger Temp-Dateiname
            temp_fd, temp_path = tempfile.mkstemp(
                suffix=ext,
                prefix=f'_undo_{os.getpid()}_'
            )
            os.close(temp_fd)  # Dateihandle schließen, SaveCopyAs braucht nur den Pfad
            
            if platform.system() == 'Windows':
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
    
    def _cleanup_undo_entry(self, entry: Dict):
        """Löscht die Temp-Datei eines Undo-Eintrags."""
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
        self._log("Alle Undo-Temp-Dateien aufgeräumt")
    
    def undo(self) -> Dict[str, Any]:
        """Macht die letzte Aktion rückgängig.
        
        Stellt den kompletten Workbook-Zustand aus dem Snapshot wieder her:
        Schließt das Workbook, kopiert den Snapshot zurück, öffnet es erneut.
        100% Fidelity: Formatierung, Formeln, bedingte Formatierung — alles.
        """
        try:
            if not self.file_path:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not self.undo_stack:
                return {'success': False, 'error': 'Nichts zum Rückgängig machen'}
            
            entry = self.undo_stack.pop()
            label = entry.get('label', 'Unbekannt')
            temp_path = entry.get('temp_path')
            sheet_name = entry.get('sheet_name', self.sheet_name)
            
            if not temp_path or not os.path.exists(temp_path):
                return {'success': False, 'error': 'Undo-Snapshot nicht gefunden'}
            
            self._undo_in_progress = True
            original_path = self.file_path
            password = self.file_password
            
            self._log(f"Undo: {label} — Restore von {os.path.basename(temp_path)}")
            
            try:
                # 0. Fenster-Zustand sichern (damit Excel nicht minimiert wird)
                window_state = None
                app_visible = True
                try:
                    if platform.system() == 'Windows':
                        # xlMaximized=-4137, xlMinimized=-4140, xlNormal=-4143
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
                        self.workbook.api.Saved = True  # Verhindert "Speichern?"-Dialog
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
                
                # 2. Snapshot-Datei → Original-Pfad kopieren
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
                    # Fallback: Erstes Sheet
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
                
                # 5. Temp-Datei aufräumen
                try:
                    os.unlink(temp_path)
                except Exception:
                    pass
                
                self._log(f"Undo erfolgreich: {label} (noch {len(self.undo_stack)} Undo-Schritte)")
                return {'success': True, 'undone': label, 'undoCount': len(self.undo_stack)}
                
            finally:
                self._undo_in_progress = False
            
        except Exception as e:
            self._undo_in_progress = False
            self._log(f"Undo Fehler: {e}")
            return {'success': False, 'error': str(e)}
    
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
    
    def insert_row(self, row_index: int, count: int = 1) -> Dict[str, Any]:
        """Fügt leere Zeilen ein (0-basierter Index, ohne Header)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row_start = row_index + 2
            excel_row_end = excel_row_start + count - 1
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'{count} Zeile(n) eingefügt')
            
            self._log(f"Füge {count} Zeile(n) bei {excel_row_start} ein")
            self.worksheet.range(f'{excel_row_start}:{excel_row_end}').insert(shift='down')
            
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
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'Zeile verschoben ({from_index + 1} → {to_index + 1})')
            
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
    
    def insert_column(self, col_index: int, count: int = 1, headers: list = None) -> Dict[str, Any]:
        """Fügt leere Spalten ein"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_col = col_index + 1
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'{count} Spalte(n) eingefügt')
            
            for i in range(count):
                insert_letter = self._get_column_letter(excel_col + i)
                self._log(f"Füge Spalte {insert_letter} ein")
                self.worksheet.range(f'{insert_letter}:{insert_letter}').insert(shift='right')
            
            # Header setzen falls vorhanden
            if headers:
                for i, header in enumerate(headers):
                    self.worksheet.range((1, excel_col + i)).value = header
            
            # Journal-Eintrag
            self._journal_add('insertColumn', {'colIndex': col_index, 'count': count, 'headers': headers})
            self._check_auto_save()
            
            return {'success': True, 'insertedAt': col_index, 'count': count}
            
        except Exception as e:
            self._log(f"Fehler beim Einfügen der Spalte: {e}")
            return {'success': False, 'error': str(e)}
    
    def move_column(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """Verschiebt eine Spalte per Cut & Insert (verhindert doppelte Header)"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if from_index == to_index:
                return {'success': True, 'movedFrom': from_index, 'movedTo': to_index}
            
            excel_from = from_index + 1
            excel_to = to_index + 1
            
            source_letter = self._get_column_letter(excel_from)
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'Spalte verschoben ({from_index + 1} → {to_index + 1})')
            
            self._log(f"Verschiebe Spalte {source_letter} (idx {from_index}) -> idx {to_index}")
            
            last_row = self.worksheet.used_range.last_cell.row if self.worksheet.used_range else 1000
            
            # Schritt 1: Quelldaten in Zwischenspeicher lesen (Header + Daten)
            source_rng = self.worksheet.range(f'{source_letter}1:{source_letter}{last_row}')
            col_data = source_rng.value
            # Einzelwert in Liste umwandeln
            if not isinstance(col_data, list):
                col_data = [col_data]
            
            # Schritt 2: Quellspalte löschen
            self.worksheet.range(f'{source_letter}:{source_letter}').delete()
            
            # Schritt 3: Zielposition anpassen (nach Löschung verschieben sich Indizes)
            if from_index < to_index:
                # Spalte war links, nach Löschung verschiebt sich Ziel um 1 nach links
                insert_col = excel_to  # excel_to - 1 + 1 = excel_to
            else:
                # Spalte war rechts, Ziel bleibt gleich
                insert_col = excel_to
            
            insert_letter = self._get_column_letter(insert_col)
            
            # Schritt 4: Leere Spalte an Zielposition einfügen
            self.worksheet.range(f'{insert_letter}:{insert_letter}').insert(shift='right')
            
            # Schritt 5: Daten in neue Spalte schreiben
            target_rng = self.worksheet.range(f'{insert_letter}1:{insert_letter}{len(col_data)}')
            # xlwings erwartet für vertikale Ranges eine verschachtelte Liste
            target_rng.value = [[v] for v in col_data]
            
            # Screen refresh erzwingen
            self._force_screen_refresh()
            
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
    
    def set_cell_value(self, row_index: int, col_index: int, value: Any) -> Dict[str, Any]:
        """Setzt den Wert einer Zelle"""
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            excel_row = row_index + 2
            excel_col = col_index + 1
            
            # Alten Wert für Journal holen
            old_value = self.worksheet.range((excel_row, excel_col)).value
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot('Zelle bearbeitet')
            
            self.worksheet.range((excel_row, excel_col)).value = value
            
            # Änderung im Journal protokollieren
            self._journal_add('setCellValue', {
                'row': row_index,
                'col': col_index,
                'oldValue': str(old_value) if old_value else None,
                'newValue': str(value) if value else None
            })
            
            # Prüfen ob Auto-Save fällig ist
            self._check_auto_save()
            
            return {'success': True, 'row': row_index, 'col': col_index, 'value': value}
            
        except Exception as e:
            self._log(f"Fehler beim Setzen des Zellwerts: {e}")
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
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'Spaltenwerte geändert (Spalte {col_letter})')
            
            # Werte als vertikale Liste formatieren (jeder Wert in eigener Liste)
            vertical_values = [[v] for v in values]
            
            self._log(f"Setze {len(values)} Werte in Spalte {col_letter} (Zeilen {excel_start_row}-{end_row})")
            self.worksheet.range(range_addr).value = vertical_values
            
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
            
            # Undo-Snapshot: Drosselung bei schnellen aufeinanderfolgenden Edits
            # Nur Snapshot erstellen wenn letzter > 2 Sekunden her
            now = time.time()
            if now - self._last_undo_snapshot_time >= 2.0:
                self._push_undo_snapshot(f'Zeilenwerte geändert (Zeile {row_index + 1})')
                self._last_undo_snapshot_time = now
            else:
                self._log(f"set_row_values: Undo-Snapshot übersprungen (letzer vor {now - self._last_undo_snapshot_time:.1f}s)")
            
            if platform.system() == 'Darwin':
                # macOS: Direkt über AppleScript für zuverlässigen Display-Refresh
                self._set_row_values_applescript(excel_row, values)
            else:
                # Windows: xlwings Range-Write + expliziter Screen-Refresh
                self.app.screen_updating = True
                self._log(f"set_row_values: Schreibe in Range {range_addr}...")
                self.worksheet.range(range_addr).value = values
                self._log(f"set_row_values: Range-Write abgeschlossen, Screen-Refresh...")
                self._force_screen_refresh()
            
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
        """Setzt mehrere Zellen auf einmal (für Suchen & Ersetzen)
        
        LEGACY: Wird nur noch für Einzelzellen verwendet.
        Für Bulk-Ersetzungen: find_replace() nutzen.
        """
        try:
            if not self.worksheet:
                return {'success': False, 'error': 'Keine Datei geöffnet'}
            
            if not cells or len(cells) == 0:
                return {'success': True, 'count': 0}
            
            self._log(f"set_cells_batch: Setze {len(cells)} Zellen")
            
            # Undo-Snapshot: Komplettes Workbook sichern
            self._push_undo_snapshot(f'{len(cells)} Zelle(n) geändert')
            
            updated_count = 0
            
            # Für kleine Batches (≤5 Zellen): Direkt schreiben ohne screen_updating-Tricks
            # Screen-Updating deaktivieren verursacht auf macOS Probleme bei Einzel-Edits
            if len(cells) <= 5:
                for cell in cells:
                    row_index = cell.get('row')
                    col_index = cell.get('col')
                    value = cell.get('value')
                    
                    if row_index is None or col_index is None:
                        continue
                    
                    excel_row = row_index + 2  # +2 für Header
                    excel_col = col_index + 1
                    
                    self.worksheet.range((excel_row, excel_col)).value = value
                    
                    updated_count += 1
                
                # macOS: Screen-Refresh erzwingen damit Änderung sichtbar wird
                if platform.system() == 'Darwin':
                    self._force_screen_refresh()
            else:
                # Performance-Optimierung nur für große Batches
                app = self.app
                original_screen_updating = app.screen_updating
                original_calculation = app.calculation
                
                try:
                    app.screen_updating = False
                    app.calculation = 'manual'
                    
                    for cell in cells:
                        row_index = cell.get('row')
                        col_index = cell.get('col')
                        value = cell.get('value')
                        
                        if row_index is None or col_index is None:
                            continue
                        
                        excel_row = row_index + 2  # +2 für Header
                        excel_col = col_index + 1
                        
                        self.worksheet.range((excel_row, excel_col)).value = value
                        updated_count += 1
                    
                    # Am Ende: Formeln neu berechnen
                    app.calculate()
                    
                finally:
                    # Ursprüngliche Einstellungen wiederherstellen
                    app.screen_updating = original_screen_updating
                    app.calculation = original_calculation
            
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
            
            is_windows = platform.system() == 'Windows'
            
            try:
                if filters and len(filters) > 0:
                    if is_windows:
                        # ===== WINDOWS =====
                        # AutoFilter aktivieren falls noch nicht aktiv
                        try:
                            if not self.worksheet.api.AutoFilterMode:
                                used_range.api.AutoFilter()
                        except Exception as e:
                            self._log(f"AutoFilter-Aktivierung Fehler: {e}")
                            # Versuche es trotzdem
                            try:
                                used_range.api.AutoFilter()
                            except:
                                pass
                        
                        # Filter für jede Spalte setzen
                        self._active_filter_fields = []  # Merken für Clear
                        for f in filters:
                            col_idx = f.get('colIndex', 0) + 1  # 1-basiert
                            criteria = f.get('criteria', '')
                            operator = f.get('operator', 'equals')
                            date_from = f.get('dateFrom', None)
                            date_to = f.get('dateTo', None)
                            
                            # ---- Text-Filter ----
                            if operator == 'contains':
                                criteria = f'*{criteria}*'
                            elif operator == 'notContains':
                                criteria = f'<>*{criteria}*'
                            elif operator == 'startsWith':
                                criteria = f'{criteria}*'
                            elif operator == 'endsWith':
                                criteria = f'*{criteria}'
                            elif operator == 'isEmpty':
                                criteria = '='
                            elif operator == 'isNotEmpty':
                                criteria = '<>'
                            
                            # ---- Datums-Filter ----
                            elif operator in ('dateToday', 'datePast', 'dateFuture',
                                              'dateThisWeek', 'dateThisMonth',
                                              'dateInDays', 'dateOverdueDays', 'dateBetween'):
                                self._log(f"Windows: Datums-Filter Spalte {col_idx}: op={operator}, from={date_from}, to={date_to}")
                                try:
                                    # Datumsformat der Spalte erkennen und ISO-Daten umwandeln
                                    # Excel AutoFilter erwartet Kriterien im angezeigten Zellformat
                                    def _convert_date_for_excel(iso_date_str, col_index):
                                        """Konvertiert ISO-Datum (YYYY-MM-DD) ins Zellformat der Spalte"""
                                        try:
                                            dt = datetime.strptime(iso_date_str, '%Y-%m-%d')
                                        except:
                                            return iso_date_str  # Fallback: unverändert
                                        
                                        try:
                                            # NumberFormat der ersten Datenzelle dieser Spalte lesen
                                            col_letter = self._get_column_letter(col_index)
                                            cell = self.worksheet.range(f'{col_letter}2')
                                            num_fmt = cell.number_format or ''
                                            self._log(f"  Spalte {col_index} NumberFormat: '{num_fmt}'")
                                            
                                            # Bekannte Excel-Datumsformate → Python strftime
                                            # WICHTIG: Reihenfolge von m/d im Format bestimmt MM.DD vs DD.MM
                                            nf = num_fmt.lower().replace('\\', '').strip()
                                            
                                            # Position von m und d im Format bestimmen
                                            # (Suche erstes 'm' und erstes 'd' das nicht in anderen Tokens steckt)
                                            import re as _re
                                            m_pos = -1
                                            d_pos = -1
                                            for _m in _re.finditer(r'[md]', nf):
                                                ch = _m.group()
                                                if ch == 'm' and m_pos < 0:
                                                    m_pos = _m.start()
                                                elif ch == 'd' and d_pos < 0:
                                                    d_pos = _m.start()
                                            
                                            month_first = m_pos < d_pos if m_pos >= 0 and d_pos >= 0 else False
                                            self._log(f"  Format-Analyse: m_pos={m_pos}, d_pos={d_pos}, month_first={month_first}")
                                            
                                            # Separator erkennen
                                            sep = '.'
                                            for ch in nf:
                                                if ch in './-':
                                                    sep = ch
                                                    break
                                            
                                            # Jahr-Format erkennen
                                            has_4y = 'yyyy' in nf
                                            
                                            if 'yyyy' in nf:
                                                y_pos = nf.index('yyyy')
                                                if y_pos == 0 or (m_pos >= 0 and y_pos < m_pos):
                                                    # YYYY zuerst: YYYY-MM-DD
                                                    return dt.strftime(f'%Y{sep}%m{sep}%d')
                                            
                                            # Padding erkennen: 'dd' = 2-stellig, 'd' = ohne führende Null
                                            has_dd = 'dd' in nf
                                            has_mm = 'mm' in nf
                                            
                                            if month_first:
                                                # MM.DD.YY oder M.DD.YY
                                                m_fmt = '%m' if has_mm else '%#m'
                                                d_fmt = '%d' if has_dd else '%#d'
                                                y_fmt = '%Y' if has_4y else '%#y' if not has_4y else '%Y'
                                                # Windows: %#y gibt keine führende Null, aber strftime hat kein 2-digit year
                                                # Verwende stattdessen manuelles Format
                                                y_str = str(dt.year) if has_4y else str(dt.year % 100).zfill(2)
                                                m_str = str(dt.month).zfill(2) if has_mm else str(dt.month)
                                                d_str = str(dt.day).zfill(2) if has_dd else str(dt.day)
                                                return f'{m_str}{sep}{d_str}{sep}{y_str}'
                                            else:
                                                # DD.MM.YY oder D.M.YYYY
                                                y_str = str(dt.year) if has_4y else str(dt.year % 100).zfill(2)
                                                m_str = str(dt.month).zfill(2) if has_mm else str(dt.month)
                                                d_str = str(dt.day).zfill(2) if has_dd else str(dt.day)
                                                return f'{d_str}{sep}{m_str}{sep}{y_str}'
                                        except Exception as e:
                                            self._log(f"  NumberFormat-Erkennung fehlgeschlagen: {e}")
                                        
                                        # Fallback: Beispielwert der Zelle lesen und Format ableiten
                                        try:
                                            col_letter = self._get_column_letter(col_index)
                                            cell_val = self.worksheet.range(f'{col_letter}2').value
                                            sample = str(cell_val or '')
                                            self._log(f"  Fallback: Beispielwert='{sample}' (Typ: {type(cell_val).__name__})")
                                            
                                            if '.' in sample and sample.count('.') == 2:
                                                # Punkt-Trenner → prüfe ob MM.DD oder DD.MM
                                                # Suche eindeutigen Wert in den Daten
                                                import re as _re2
                                                last_row = min(used_range.last_cell.row, 52)
                                                col_vals = self.worksheet.range(f'{col_letter}2:{col_letter}{last_row}').value
                                                if not isinstance(col_vals, list):
                                                    col_vals = [col_vals]
                                                is_mdy = False
                                                for cv in col_vals:
                                                    s = str(cv or '').strip()
                                                    dm = _re2.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{2,4})', s)
                                                    if dm:
                                                        pp1, pp2 = int(dm.group(1)), int(dm.group(2))
                                                        if pp2 > 12 and pp1 <= 12:
                                                            is_mdy = True
                                                            break
                                                        if pp1 > 12 and pp2 <= 12:
                                                            break  # DD.MM confirmed
                                                if is_mdy:
                                                    self._log(f"  Fallback: MM.DD Format erkannt")
                                                    return f'{dt.month}.{dt.day:02d}.{dt.year % 100:02d}'
                                                else:
                                                    return dt.strftime('%d.%m.%Y')
                                            elif '/' in sample:
                                                return dt.strftime('%#m/%#d/%Y')
                                            elif '-' in sample and not sample[:4].isdigit():
                                                return dt.strftime('%d-%m-%Y')
                                        except Exception as e:
                                            self._log(f"  Fallback-Fehler: {e}")
                                        
                                        # Letzter Fallback: M/D/YYYY (englisch, gängigste COM-Variante)
                                        return dt.strftime('%#m/%#d/%Y')
                                    
                                    c1 = None
                                    c2 = None
                                    xl_op = None  # 1 = xlAnd
                                    
                                    # ISO-Daten ins Zellformat konvertieren
                                    excel_from = _convert_date_for_excel(date_from, col_idx) if date_from else None
                                    excel_to = _convert_date_for_excel(date_to, col_idx) if date_to else None
                                    self._log(f"  Konvertiert: from={date_from} → {excel_from}, to={date_to} → {excel_to}")
                                    
                                    if excel_from and excel_to:
                                        c1 = f">={excel_from}"
                                        c2 = f"<={excel_to}"
                                        xl_op = 1  # xlAnd
                                    elif excel_from:
                                        c1 = f">={excel_from}"
                                    elif excel_to:
                                        c1 = f"<={excel_to}"
                                    else:
                                        self._log(f"Windows: Datums-Filter Spalte {col_idx} übersprungen (keine Daten)")
                                        continue
                                    
                                    if c2 and xl_op:
                                        used_range.api.AutoFilter(Field=col_idx, Criteria1=c1, Operator=xl_op, Criteria2=c2)
                                    else:
                                        used_range.api.AutoFilter(Field=col_idx, Criteria1=c1)
                                    self._active_filter_fields.append(col_idx)
                                    self._log(f"Windows: Datums-Filter Spalte {col_idx} gesetzt: c1={c1}, c2={c2}")
                                except Exception as e:
                                    self._log(f"Fehler bei Datums-Filter Spalte {col_idx}: {e}")
                                continue  # Skip den normalen AutoFilter-Aufruf unten
                            
                            self._log(f"Windows: Setze Filter Spalte {col_idx}: operator={operator}, criteria='{criteria}'")
                            
                            try:
                                # AutoFilter mit Kriterien setzen
                                used_range.api.AutoFilter(Field=col_idx, Criteria1=criteria)
                                self._active_filter_fields.append(col_idx)
                            except Exception as e:
                                self._log(f"Fehler bei Filter Spalte {col_idx}: {e}")
                    else:
                        # ===== macOS =====
                        # Auf macOS funktioniert weder appscript auto_filter()
                        # noch VBA via AppleScript zuverlässig.
                        # Batch-Zeilen-Ausblendung: Spalten-Daten in einem Aufruf
                        # lesen, zusammenhängende Bereiche gruppiert ausblenden.
                        
                        last_row = used_range.last_cell.row
                        rows_to_hide = set()
                        
                        for f in filters:
                            col_idx = f.get('colIndex', 0) + 1
                            criteria = f.get('criteria', '').lower()
                            operator = f.get('operator', 'equals')
                            col_letter = self._get_column_letter(col_idx)
                            
                            # Ganze Spalte auf einmal lesen (1 API-Aufruf)
                            col_range = self.worksheet.range(f'{col_letter}2:{col_letter}{last_row}')
                            col_values = col_range.value
                            if not isinstance(col_values, list):
                                col_values = [col_values]
                            
                            # Datum-Grenzen für Datums-Filter vorbereiten
                            date_from = f.get('dateFrom', None)
                            date_to = f.get('dateTo', None)
                            date_from_dt = None
                            date_to_dt = None
                            is_date_op = operator in ('dateToday', 'datePast', 'dateFuture',
                                                       'dateThisWeek', 'dateThisMonth',
                                                       'dateInDays', 'dateOverdueDays', 'dateBetween')
                            if is_date_op:
                                try:
                                    if date_from:
                                        date_from_dt = datetime.strptime(date_from, '%Y-%m-%d')
                                    if date_to:
                                        date_to_dt = datetime.strptime(date_to, '%Y-%m-%d')
                                except Exception as e:
                                    self._log(f"macOS: Datums-Parse Fehler: {e}")
                            
                            for idx, cell_value in enumerate(col_values):
                                row_num = idx + 2
                                cell_str = str(cell_value).lower() if cell_value is not None else ''
                                matches = False
                                
                                if is_date_op:
                                    # Datums-Vergleich
                                    cell_date = None
                                    if isinstance(cell_value, datetime):
                                        cell_date = cell_value
                                    elif isinstance(cell_value, date):
                                        cell_date = datetime.combine(cell_value, dtime())
                                    elif cell_value:
                                        # Versuche String als Datum zu parsen
                                        for dfmt in ('%Y-%m-%d', '%d.%m.%Y', '%m/%d/%Y', '%d/%m/%Y'):
                                            try:
                                                cell_date = datetime.strptime(str(cell_value).strip(), dfmt)
                                                break
                                            except:
                                                pass
                                    
                                    if cell_date:
                                        cell_date = cell_date.replace(hour=0, minute=0, second=0, microsecond=0)
                                        if date_from_dt and date_to_dt:
                                            matches = date_from_dt <= cell_date <= date_to_dt
                                        elif date_from_dt:
                                            matches = cell_date >= date_from_dt
                                        elif date_to_dt:
                                            matches = cell_date <= date_to_dt
                                    # cell_date == None → matches bleibt False
                                elif operator == 'contains':
                                    matches = criteria in cell_str
                                elif operator == 'notContains':
                                    matches = criteria not in cell_str
                                elif operator == 'startsWith':
                                    matches = cell_str.startswith(criteria)
                                elif operator == 'endsWith':
                                    matches = cell_str.endswith(criteria)
                                elif operator == 'equals':
                                    matches = cell_str == criteria
                                elif operator == 'isEmpty':
                                    matches = cell_value is None or cell_str.strip() == ''
                                elif operator == 'isNotEmpty':
                                    matches = cell_value is not None and cell_str.strip() != ''
                                else:
                                    matches = criteria in cell_str
                                if not matches:
                                    rows_to_hide.add(row_num)
                        
                        self._log(f"Filter: {len(rows_to_hide)} von {last_row - 1} Zeilen ausblenden")
                        
                        # Alle Zeilen einblenden (1 API-Aufruf)
                        try:
                            all_rows = self.worksheet.range(f'A2:A{last_row}')
                            all_rows.api.entire_row.hidden.set(False)
                        except Exception as e:
                            self._log(f"Einblenden-Fehler: {e}")
                        
                        # Zusammenhängende Bereiche gruppiert ausblenden
                        if rows_to_hide:
                            sorted_rows = sorted(rows_to_hide)
                            ranges = []
                            start = end = sorted_rows[0]
                            for row in sorted_rows[1:]:
                                if row == end + 1:
                                    end = row
                                else:
                                    ranges.append((start, end))
                                    start = end = row
                            ranges.append((start, end))
                            
                            hidden_count = 0
                            for start_row, end_row in ranges:
                                try:
                                    rng = self.worksheet.range(f'A{start_row}:A{end_row}')
                                    rng.api.entire_row.hidden.set(True)
                                    hidden_count += (end_row - start_row + 1)
                                except Exception as e:
                                    self._log(f"Hide {start_row}:{end_row} Fehler: {e}")
                            self._log(f"{hidden_count} Zeilen in {len(ranges)} Bereichen ausgeblendet")
                else:
                    # ===== AutoFilter entfernen / Alle Zeilen einblenden =====
                    # Delegiere an clear_autofilter
                    return self.clear_autofilter()
                        
            except Exception as api_error:
                self._log(f"AutoFilter API Fehler: {api_error}")
                return {'success': False, 'error': str(api_error)}
            
            self._log(f"AutoFilter abgeschlossen: {len(filters) if filters else 0} Filter")
            return {'success': True, 'filterCount': len(filters) if filters else 0}
            
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
            
            is_windows = platform.system() == 'Windows'
            self._log(f"clear_autofilter: Start (platform={platform.system()})")
            
            if is_windows:
                # ===== WINDOWS: AutoFilter entfernen =====
                return self._clear_autofilter_windows()
            else:
                # ===== macOS: Alle Zeilen einblenden =====
                return self._clear_autofilter_macos()
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
            
            # Schritt 5: Screen-Refresh
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
    
    def _clear_autofilter_macos(self) -> Dict[str, Any]:
        """macOS: Alle versteckten Zeilen einblenden (AutoFilter wird via row.hidden simuliert)."""
        try:
            used_range = self.worksheet.used_range
            if used_range:
                last_row = used_range.last_cell.row
            else:
                last_row = 1
            
            self._log(f"macOS: Einblenden aller Zeilen (last_row={last_row})")
            
            if last_row <= 1:
                self._log("macOS: Nur Header, nichts zum Einblenden")
                return {'success': True, 'filterCount': 0}
            
            # Methode 1: Alle Zeilen des Sheets auf einmal einblenden
            unhidden = False
            try:
                self.worksheet.api.rows.hidden.set(False)
                self._log("macOS: Methode 1 (sheet.rows.hidden=False) erfolgreich")
                unhidden = True
            except Exception as e1:
                self._log(f"macOS: Methode 1 fehlgeschlagen: {e1}")
            
            # Methode 2: Bereich A2:A{last_row} einblenden
            if not unhidden:
                try:
                    all_rows = self.worksheet.range(f'A2:A{last_row}')
                    all_rows.api.entire_row.hidden.set(False)
                    self._log(f"macOS: Methode 2 (range A2:A{last_row}) erfolgreich")
                    unhidden = True
                except Exception as e2:
                    self._log(f"macOS: Methode 2 fehlgeschlagen: {e2}")
            
            # Methode 3: Zeilen einzeln in Blöcken von 100 einblenden
            if not unhidden:
                self._log("macOS: Methode 3 — blockweise Einblendung")
                block_size = 100
                error_count = 0
                for start in range(2, last_row + 1, block_size):
                    end = min(start + block_size - 1, last_row)
                    try:
                        rng = self.worksheet.range(f'A{start}:A{end}')
                        rng.api.entire_row.hidden.set(False)
                    except Exception as e3:
                        error_count += 1
                        self._log(f"macOS: Block {start}-{end} Fehler: {e3}")
                if error_count == 0:
                    unhidden = True
                    self._log("macOS: Methode 3 erfolgreich")
                else:
                    self._log(f"macOS: Methode 3 mit {error_count} Fehlern")
            
            # Methode 4: Zeile für Zeile (letzter Versuch)
            if not unhidden:
                self._log("macOS: Methode 4 — zeilenweise Einblendung")
                for row_num in range(2, last_row + 1):
                    try:
                        self.worksheet.range(f'A{row_num}').api.entire_row.hidden.set(False)
                    except:
                        pass
                unhidden = True
                self._log("macOS: Methode 4 abgeschlossen")
            
            if not unhidden:
                return {'success': False, 'error': 'Konnte Zeilen nicht einblenden'}
            
            self._log("macOS: clear_autofilter OK")
            return {'success': True, 'filterCount': 0}
            
        except Exception as e:
            self._log(f"macOS: clear_autofilter Fehler: {e}")
            return {'success': False, 'error': str(e)}
    
    def switch_sheet(self, sheet_name: str) -> Dict[str, Any]:
        """Wechselt das aktive Arbeitsblatt in der Live Session
        
        Args:
            sheet_name: Name des Zielblatts
        """
        try:
            if not self.workbook:
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            target_sheet = self.workbook.sheets[sheet_name]
            
            # Pr\u00fcfe ob Sheet ausgeblendet ist
            was_hidden = not target_sheet.visible
            if was_hidden:
                # Sheet muss zuerst eingeblendet werden, sonst schl\u00e4gt activate() fehl
                target_sheet.visible = True
                self._log(f"Sheet '{sheet_name}' war ausgeblendet \u2192 automatisch eingeblendet")
            
            self.worksheet = target_sheet
            self.sheet_name = sheet_name
            self.worksheet.activate()
            
            self._log(f"Sheet gewechselt zu: {sheet_name}")
            return {'success': True, 'sheetName': sheet_name, 'wasHidden': was_hidden}
            
        except Exception as e:
            self._log(f"Fehler beim Sheet-Wechsel: {e}")
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
            
            # Mindestens ein Sheet muss sichtbar bleiben
            if not visible:
                visible_count = sum(1 for s in self.workbook.sheets if s.visible)
                if visible_count <= 1:
                    return {'success': False, 'error': 'Mindestens ein Arbeitsblatt muss sichtbar bleiben'}
            
            target_sheet = self.workbook.sheets[sheet_name]
            target_sheet.visible = visible
            
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
                return {'success': False, 'error': 'Keine Datei ge\u00f6ffnet'}
            
            sheet_names = [s.name for s in self.workbook.sheets]
            if sheet_name not in sheet_names:
                return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
            
            num_sheets = len(self.workbook.sheets)
            if new_index < 0 or new_index >= num_sheets:
                return {'success': False, 'error': f'Ung\u00fcltiger Index: {new_index}'}
            
            sheet = self.workbook.sheets[sheet_name]
            
            if new_index == 0:
                # An den Anfang: vor das erste Sheet
                sheet.api.Move(Before=self.workbook.sheets[0].api)
            else:
                # Nach dem Sheet an Position new_index-1
                # Aber wir m\u00fcssen ber\u00fccksichtigen, dass das Sheet selbst verschoben wird
                target_sheet = self.workbook.sheets[new_index]
                current_index = sheet_names.index(sheet_name)
                if current_index < new_index:
                    sheet.api.Move(After=target_sheet.api)
                else:
                    sheet.api.Move(Before=target_sheet.api)
            
            self._log(f"Sheet '{sheet_name}' verschoben zu Index {new_index}")
            sheets = [s.name for s in self.workbook.sheets]
            return {'success': True, 'sheets': sheets}
            
        except Exception as e:
            self._log(f"Fehler beim Verschieben: {e}")
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
            'save': lambda: self.save_file(cmd.get('outputPath'), cmd.get('password')),
            'close': lambda: self.close_session(save=cmd.get('save', False)),
            'getData': lambda: self.get_data(),
            'switchSheet': lambda: self.switch_sheet(cmd.get('sheetName')),
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
            
            # Spalten
            'deleteColumn': lambda: self.delete_column(cmd.get('colIndex')),
            'insertColumn': lambda: self.insert_column(cmd.get('colIndex'), cmd.get('count', 1), cmd.get('headers')),
            'moveColumn': lambda: self.move_column(cmd.get('fromIndex'), cmd.get('toIndex')),
            'hideColumn': lambda: self.hide_column(cmd.get('colIndex'), cmd.get('hidden', True)),
            
            # Zellen
            'setCellValue': lambda: self.set_cell_value(cmd.get('rowIndex'), cmd.get('colIndex'), cmd.get('value')),
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
                start_time = time.time()
                self._log(f"CMD empfangen: {action}")
                
                result = self.handle_command(cmd)
                
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
    session.run()


if __name__ == '__main__':
    main()
