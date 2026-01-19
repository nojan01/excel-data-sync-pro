#!/usr/bin/env python3
"""Test ob Excel wirklich beendet wird"""

import subprocess
import time
import xlwings as xw
import platform

def check_excel_running():
    """Prüft ob Excel läuft"""
    if platform.system() == 'Darwin':
        result = subprocess.run(['pgrep', '-x', 'Microsoft Excel'], 
                                capture_output=True, timeout=2)
        return result.returncode == 0
    return False

def kill_excel_instances():
    """Beendet alle laufenden Excel-Instanzen auf macOS - sehr aggressiv"""
    if platform.system() == 'Darwin':
        # Methode 1: xlwings Apps beenden
        try:
            for app in xw.apps:
                try:
                    app.quit()
                except:
                    pass
        except:
            pass
        
        # Methode 2: AppleScript quit (ohne waiting)
        try:
            subprocess.run(['osascript', '-e', 
                'tell application "Microsoft Excel" to quit saving no'], 
                capture_output=True, timeout=2)
        except:
            pass
        
        # Kurz warten
        time.sleep(0.2)
        
        # Methode 3: Prüfen ob Excel noch läuft und mit pkill beenden
        try:
            result = subprocess.run(['pgrep', '-x', 'Microsoft Excel'], 
                                    capture_output=True, timeout=2)
            if result.returncode == 0:
                # Excel läuft noch - sofort killen!
                print("Excel läuft noch nach AppleScript quit - verwende pkill -9")
                subprocess.run(['pkill', '-9', 'Microsoft Excel'], 
                               capture_output=True, timeout=2)
                time.sleep(0.2)
        except:
            pass
        
        # Methode 4: Falls immer noch da, nochmal versuchen
        try:
            result = subprocess.run(['pgrep', '-x', 'Microsoft Excel'], 
                                    capture_output=True, timeout=1)
            if result.returncode == 0:
                print("Excel läuft IMMER NOCH - verwende killall -9")
                subprocess.run(['killall', '-9', 'Microsoft Excel'], 
                               capture_output=True, timeout=2)
        except:
            pass


# Zuerst sicherstellen dass Excel nicht läuft
print("=== Excel Kill Test (OHNE with-Block) ===")
print(f"Excel läuft vorher: {check_excel_running()}")

# Kill falls es läuft
if check_excel_running():
    print("Beende Excel zuerst...")
    kill_excel_instances()
    time.sleep(0.5)
    print(f"Excel läuft nach Kill: {check_excel_running()}")

# Jetzt xlwings starten OHNE with-Block
print("\n--- Starte xlwings (ohne with) ---")
app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False
print(f"Excel läuft nach xw.App: {check_excel_running()}")

# Workbook erstellen
wb = app.books.add()
ws = wb.sheets[0]
ws.range('A1').value = 'Test'
print("Workbook erstellt und beschrieben")

# Workbook schließen (ohne speichern)
wb.close()
print(f"Excel läuft nach wb.close(): {check_excel_running()}")

# Kill nach wb.close()
kill_excel_instances()
time.sleep(0.3)
print(f"Excel läuft nach kill_excel_instances(): {check_excel_running()}")

if not check_excel_running():
    print("\n✅ ERFOLG: Excel wurde beendet!")
else:
    print("\n❌ FEHLER: Excel läuft immer noch!")
