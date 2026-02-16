# Prüfe, ob openpyxl installiert ist, und gib eine verständliche Fehlermeldung aus
try:
    import openpyxl
except ImportError:
    import sys
    print("[Fehler] Das Python-Modul 'openpyxl' ist nicht installiert. Bitte führe im python-Verzeichnis 'pip3 install -r requirements.txt' aus.", file=sys.stderr)
    sys.exit(1)
#!/usr/bin/env python3
"""
Excel Writer für Excel Data Sync Pro
Verwendet openpyxl für bessere Kompatibilität mit Excel-Formaten

Der große Vorteil von openpyxl: 
- Öffnet die Original-Datei und modifiziert nur die geänderten Zellen
- Behält ALLE Formatierungen, bedingte Formatierungen, Tabellen, etc.

WICHTIG: openpyxl's delete_cols() aktualisiert CF-Bereiche NICHT automatisch!
Für strukturelle Änderungen (Spalten löschen/einfügen) nutzen wir xlwings wenn
Microsoft Excel installiert ist - das erhält ALLE Formatierungen.
"""

import json
import sys
import os
import re
from datetime import datetime, date
from copy import copy

# ============================================================================
# MONKEY-PATCH: openpyxl PatternFill um extLst zu ignorieren
# Manche Excel-Dateien haben erweiterte Formatierungen die openpyxl nicht kennt
# WICHTIG: Muss VOR dem Import von openpyxl erfolgen!
# ============================================================================
import openpyxl.styles.fills as _fills_module
from openpyxl.styles.colors import Color
from openpyxl.descriptors.base import Typed

# Patch die Typed Descriptor Klasse um None-Werte für Color mit Default zu ersetzen
_original_typed_set = Typed.__set__

def _patched_typed_set(self, instance, value):
    """Gepatchter Typed.__set__ der None für Color-Typen durch Default ersetzt"""
    if value is None and hasattr(self, 'expected_type') and self.expected_type == Color:
        # Statt None einen transparenten Default-Color setzen
        value = Color(rgb='00000000')
    _original_typed_set(self, instance, value)

Typed.__set__ = _patched_typed_set

_OriginalPatternFill = _fills_module.PatternFill
_original_init = _OriginalPatternFill.__init__

def _patched_init(self, patternType=None, fgColor=None, bgColor=None, 
                  fill_type=None, start_color=None, end_color=None, **kwargs):
    """Gepatchter __init__ der unbekannte kwargs wie extLst ignoriert"""
    _original_init(self, patternType=patternType, fgColor=fgColor, bgColor=bgColor,
                   fill_type=fill_type, start_color=start_color, end_color=end_color)

_OriginalPatternFill.__init__ = _patched_init

# Patch auch from_tree um extLst child nodes zu entfernen
_original_from_tree = _OriginalPatternFill.from_tree.__func__

@classmethod  
def _patched_from_tree(cls, node):
    """Gepatchte from_tree die extLst child nodes entfernt"""
    for child in list(node):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'extLst':
            node.remove(child)
        # Wenn fgColor oder bgColor leer ist (keine Attribute), entferne es auch
        elif tag in ('fgColor', 'bgColor') and not child.attrib:
            node.remove(child)
    return _original_from_tree(cls, node)

_OriginalPatternFill.from_tree = _patched_from_tree
# ============================================================================

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import range_boundaries, coordinate_from_string
from openpyxl.styles import PatternFill, Font, Alignment, Border
from openpyxl.styles.colors import Color
from openpyxl.formatting.formatting import ConditionalFormattingList

# Standard Theme-Farben (Office Default Theme)
# Diese werden verwendet wenn Theme-Farben nicht aufgelöst werden können
# ACHTUNG: Die Reihenfolge ist wichtig! Excel speichert Theme-Index 0-9
THEME_COLORS = [
    'FFFFFF',  # 0: lt1 - Light 1 (Background 1, usually white)
    '000000',  # 1: dk1 - Dark 1 (Text 1, usually black)
    'E7E6E6',  # 2: lt2 - Light 2 (Background 2)
    '44546A',  # 3: dk2 - Dark 2 (Text 2)
    '4472C4',  # 4: accent1 - Blue
    'ED7D31',  # 5: accent2 - Orange
    '70AD47',  # 6: accent3 - GREEN (not gray!)
    'FFC000',  # 7: accent4 - Gold
    '5B9BD5',  # 8: accent5 - Light Blue
    '7030A0',  # 9: accent6 - Purple
]


def fix_xlsx_relationships(xlsx_path):
    """
    Repariert openpyxl-gespeicherte XLSX-Dateien.
    
    openpyxl hat mehrere Probleme:
    1. Schreibt absolute Pfade in Relationships (z.B. Target="/xl/tables/table1.xml")
       statt relative Pfade (Target="../tables/table1.xml")
    2. Schreibt XML-Dateien ohne XML-Header (<?xml version="1.0"?>)
    3. Fügt headerRowCount="1" zu Tables hinzu, was Probleme verursachen kann
    4. Setzt xmlns an falsche Position (muss am Anfang des table-Elements sein)
    
    Dies führt dazu, dass Excel die Datei als beschädigt erkennt und Tables/AutoFilter entfernt.
    """
    import zipfile
    import tempfile
    import shutil
    
    XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    
    # Erstelle temporäre Kopie
    temp_dir = tempfile.mkdtemp()
    temp_xlsx = os.path.join(temp_dir, 'fixed.xlsx')
    
    try:
        # Extrahiere die XLSX
        with zipfile.ZipFile(xlsx_path, 'r') as zf:
            zf.extractall(temp_dir)
        
        fixed_count = 0
        
        # Durchsuche alle XML-Dateien
        for root, dirs, files in os.walk(temp_dir):
            for f in files:
                if not f.endswith('.xml') and not f.endswith('.rels'):
                    continue
                    
                full_path = os.path.join(root, f)
                
                with open(full_path, 'r', encoding='utf-8') as fp:
                    content = fp.read()
                
                original_content = content
                
                # FIX 1: Füge XML-Header hinzu wenn er fehlt
                if not content.startswith('<?xml'):
                    content = XML_HEADER + content
                
                # FIX 2: Konvertiere absolute Pfade zu relativen (nur für .rels Dateien)
                if f.endswith('.rels'):
                    rel_root = root.replace(temp_dir, '').lstrip(os.sep)
                    
                    if 'worksheets/_rels' in rel_root or 'worksheets\\_rels' in rel_root:
                        content = content.replace('Target="/xl/tables/', 'Target="../tables/')
                        content = content.replace('Target="/xl/drawings/', 'Target="../drawings/')
                        content = content.replace('Target="/xl/printerSettings/', 'Target="../printerSettings/')
                    elif '_rels' in rel_root:
                        content = content.replace('Target="/xl/', 'Target="')
                
                # FIX 3: Repariere Table-XML (table*.xml Dateien)
                if f.startswith('table') and f.endswith('.xml') and 'tables' in root:
                    # openpyxl setzt xmlns am Ende der Attribute, aber es muss am Anfang sein
                    # Außerdem fügt es headerRowCount="1" hinzu, was Probleme macht
                    import re
                    
                    # Entferne headerRowCount="1" - das Original hat es nicht
                    content = re.sub(r'\s+headerRowCount="1"', '', content)
                    
                    # Stelle sicher, dass xmlns direkt nach <table kommt
                    # Pattern: <table ...andere attribute... xmlns="...">
                    # Ziel:    <table xmlns="..." ...andere attribute...>
                    match = re.search(r'<table\s+([^>]*?)xmlns="([^"]+)"([^>]*)>', content)
                    if match:
                        before_xmlns = match.group(1).strip()
                        xmlns_value = match.group(2)
                        after_xmlns = match.group(3).strip()
                        
                        # Nur umordnen wenn xmlns nicht schon am Anfang ist
                        if before_xmlns:
                            all_attrs = f'{before_xmlns} {after_xmlns}'.strip()
                            new_table_tag = f'<table xmlns="{xmlns_value}" {all_attrs}>'
                            content = content[:match.start()] + new_table_tag + content[match.end():]
                
                # FIX 4: Repariere leere inlineStr Zellen in sheet*.xml
                # openpyxl schreibt <c r="X1" t="inlineStr"></c> ohne <is> Element
                # xlsx-populate erwartet aber <is><t>...</t></is> bei t="inlineStr"
                # Lösung: Entferne t="inlineStr" bei leeren Zellen
                if f.startswith('sheet') and f.endswith('.xml') and 'worksheets' in root:
                    import re
                    # Pattern: <c ... t="inlineStr"></c> oder <c ... t="inlineStr"/>
                    # Diese leeren inlineStr-Zellen müssen repariert werden
                    content = re.sub(
                        r'<c\s+([^>]*?)t="inlineStr"([^>]*?)></c>',
                        r'<c \1\2/>',
                        content
                    )
                    content = re.sub(
                        r'<c\s+([^>]*?)t="inlineStr"([^>]*?)/>', 
                        r'<c \1\2/>',
                        content
                    )
                    # Auch leere Rows entfernen: <row r="2"></row> -> entfernen
                    content = re.sub(r'<row r="\d+"></row>', '', content)
                
                if content != original_content:
                    fixed_count += 1
                    with open(full_path, 'w', encoding='utf-8') as fp:
                        fp.write(content)
        
        if fixed_count > 0:
            
            # Erstelle neue XLSX aus den reparierten Dateien
            with zipfile.ZipFile(temp_xlsx, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(temp_dir):
                    dirs[:] = [d for d in dirs if d != '__MACOSX']
                    for f in files:
                        if f == 'fixed.xlsx' or f == '.DS_Store' or f.startswith('._'):
                            continue
                        full_path = os.path.join(root, f)
                        arc_name = os.path.relpath(full_path, temp_dir).replace('\\', '/')
                        zf.write(full_path, arc_name)
            
            # Ersetze Original mit reparierter Version
            shutil.copy2(temp_xlsx, xlsx_path)
    
    finally:
        # Cleanup
        shutil.rmtree(temp_dir, ignore_errors=True)


def restore_table_xml_from_original(output_path, original_path, table_changes=None):
    """
    Kopiert die Table-XML aus der Original-Datei und passt nur ref/tableColumns an.
    
    openpyxl verliert wichtige XML-Attribute wie xr:uid, xmlns:mc, xmlns:xr etc.
    Diese Funktion stellt die Original-Struktur wieder her und passt nur die
    notwendigen Felder an.
    
    Args:
        output_path: Pfad zur Export-Datei (wird modifiziert)
        original_path: Pfad zur Original-Datei
        table_changes: Dict mit {table_name: {'ref': new_ref, 'columns': [col_names]}}
                       Wenn None oder leer, werden alle Tables vom Original kopiert.
    """
    import zipfile
    import tempfile
    import shutil
    import re
    import sys
    
    # Prüfe ob original_path gültig ist
    if not original_path or os.path.normpath(original_path) == os.path.normpath(output_path):
        sys.stderr.write(f"[restore_table_xml] Übersprungen: original_path={original_path}, output_path={output_path}\n")
        return
    
    if not os.path.exists(original_path):
        sys.stderr.write(f"[restore_table_xml] Original existiert nicht: {original_path}\n")
        return
    
    sys.stderr.write(f"[restore_table_xml] Starte Wiederherstellung von {original_path}\n")
    
    # Bei table_changes=None: Leeres Dict verwenden (alle Tables werden kopiert)
    if table_changes is None:
        table_changes = {}
    
    temp_dir = tempfile.mkdtemp()
    temp_xlsx = os.path.join(temp_dir, 'restored.xlsx')
    orig_temp_dir = tempfile.mkdtemp()
    
    try:
        # Extrahiere beide XLSX-Dateien
        with zipfile.ZipFile(output_path, 'r') as zf:
            zf.extractall(temp_dir)
        with zipfile.ZipFile(original_path, 'r') as zf:
            zf.extractall(orig_temp_dir)
        
        fixed_count = 0
        
        # Finde alle table*.xml Dateien
        tables_dir = os.path.join(temp_dir, 'xl', 'tables')
        orig_tables_dir = os.path.join(orig_temp_dir, 'xl', 'tables')
        
        
        if os.path.exists(tables_dir) and os.path.exists(orig_tables_dir):
            for f in os.listdir(tables_dir):
                if not f.startswith('table') or not f.endswith('.xml'):
                    continue
                
                
                export_table_path = os.path.join(tables_dir, f)
                orig_table_path = os.path.join(orig_tables_dir, f)
                
                if not os.path.exists(orig_table_path):
                    continue
                
                # Lies beide Dateien
                with open(export_table_path, 'r', encoding='utf-8') as fp:
                    export_content = fp.read()
                with open(orig_table_path, 'r', encoding='utf-8') as fp:
                    orig_content = fp.read()
                
                # Extrahiere table name aus Export
                name_match = re.search(r'name="([^"]+)"', export_content)
                if not name_match:
                    continue
                table_name = name_match.group(1)
                
                # Prüfe ob wir Änderungen für diese Table haben
                if table_name not in table_changes:
                    # Keine Änderungen - kopiere einfach das Original
                    with open(export_table_path, 'w', encoding='utf-8') as fp:
                        fp.write(orig_content)
                    fixed_count += 1
                    continue
                
                changes = table_changes[table_name]
                new_ref = changes.get('ref')
                new_columns = changes.get('columns', [])
                
                # Starte mit dem Original-Content
                new_content = orig_content
                
                # Aktualisiere ref in <table> und <autoFilter>
                if new_ref:
                    # Table ref
                    new_content = re.sub(r'(<table[^>]*\s)ref="[^"]+"', f'\\1ref="{new_ref}"', new_content)
                    # AutoFilter ref
                    new_content = re.sub(r'(<autoFilter[^>]*\s)ref="[^"]+"', f'\\1ref="{new_ref}"', new_content)
                
                # Aktualisiere tableColumns
                if new_columns:
                    # Finde den tableColumns-Block
                    tc_match = re.search(r'<tableColumns[^>]*>.*?</tableColumns>', new_content, re.DOTALL)
                    if tc_match:
                        # Extrahiere die Original-Columns
                        orig_columns = re.findall(r'<tableColumn\s[^/]*(?:/>|>.*?</tableColumn>)', tc_match.group(0), re.DOTALL)
                        
                        # Erstelle ein Dict: orig_name -> Liste von (index, xml) für Duplikate
                        orig_by_name = {}
                        for idx, orig_col in enumerate(orig_columns):
                            name_match = re.search(r'name="([^"]+)"', orig_col)
                            if name_match:
                                orig_name = name_match.group(1)
                                if orig_name not in orig_by_name:
                                    orig_by_name[orig_name] = []
                                orig_by_name[orig_name].append((idx, orig_col))
                        
                        # Zähler für bereits verwendete Duplikate pro Name
                        used_count = {}
                        
                        # Baue neue tableColumns
                        new_tc_content = f'<tableColumns count="{len(new_columns)}">'
                        
                        for i, col_name in enumerate(new_columns):
                            matching_orig = None
                            
                            # Suche nach Original-Column mit gleichem Namen
                            if col_name in orig_by_name:
                                # Wie viele mit diesem Namen haben wir schon verwendet?
                                used = used_count.get(col_name, 0)
                                available = orig_by_name[col_name]
                                
                                if used < len(available):
                                    # Nimm die nächste verfügbare mit diesem Namen
                                    matching_orig = available[used][1]
                                    used_count[col_name] = used + 1
                            
                            if matching_orig:
                                # Nutze Original-Column und aktualisiere nur die ID und den Namen
                                col_xml = re.sub(r'id="\d+"', f'id="{i+1}"', matching_orig)
                                # Name auch aktualisieren (für den Fall dass er sich geändert hat)
                                safe_name = col_name.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                                col_xml = re.sub(r'name="[^"]+"', f'name="{safe_name}"', col_xml)
                                new_tc_content += col_xml
                            else:
                                # Neue Spalte ohne xr3:uid
                                # Escape special XML chars in name
                                safe_name = col_name.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                                new_tc_content += f'<tableColumn id="{i+1}" name="{safe_name}"/>'
                        
                        new_tc_content += '</tableColumns>'
                        new_content = new_content[:tc_match.start()] + new_tc_content + new_content[tc_match.end():]
                
                # Schreibe die reparierte Datei
                with open(export_table_path, 'w', encoding='utf-8') as fp:
                    fp.write(new_content)
                fixed_count += 1
        
        
        if fixed_count > 0:
            # Erstelle neue XLSX
            with zipfile.ZipFile(temp_xlsx, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(temp_dir):
                    dirs[:] = [d for d in dirs if d != '__MACOSX']
                    for f in files:
                        if f == 'restored.xlsx' or f == '.DS_Store' or f.startswith('._'):
                            continue
                        full_path = os.path.join(root, f)
                        arc_name = os.path.relpath(full_path, temp_dir).replace('\\', '/')
                        zf.write(full_path, arc_name)
            
            shutil.copy2(temp_xlsx, output_path)
    
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        shutil.rmtree(orig_temp_dir, ignore_errors=True)


# Worksheet-Element-Reihenfolge nach ECMA-376 (OpenXML Schema)
# Diese Elemente müssen am Ende von <worksheet> in genau dieser Reihenfolge stehen.
_WORKSHEET_END_ELEMENTS = [
    'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF',
    'picture', 'oleObjects', 'controls', 'webPublishItems',
    'tableParts', 'extLst'
]


def _insert_ws_element(ws_content, element_xml, element_name):
    """
    Fügt ein Element in Worksheet-XML an der schema-konformen Position ein.
    
    Excel repariert Dateien mit falscher Element-Reihenfolge — dabei können
    Elemente (z.B. <drawing>) verloren gehen. Deshalb muss das Element
    VOR allen Elementen eingefügt werden, die im Schema DANACH kommen.
    """
    try:
        idx = _WORKSHEET_END_ELEMENTS.index(element_name)
    except ValueError:
        # Unbekanntes Element → sicher vor </worksheet> einfügen
        return ws_content.replace('</worksheet>', element_xml + '\n</worksheet>')
    
    # Finde die früheste Position eines nachfolgenden Elements
    for after_elem in _WORKSHEET_END_ELEMENTS[idx + 1:]:
        # Suche nach <elementName (mit Leerzeichen oder >) um Verwechslung zu vermeiden
        import re
        pos_match = re.search(r'<' + re.escape(after_elem) + r'[\s>/]', ws_content)
        if pos_match:
            insert_pos = pos_match.start()
            return ws_content[:insert_pos] + element_xml + '\n' + ws_content[insert_pos:]
    
    # Kein nachfolgendes Element gefunden → vor </worksheet>
    return ws_content.replace('</worksheet>', element_xml + '\n</worksheet>')


def restore_external_links_from_original(output_path, original_path):
    """
    Kopiert die externalLinks-Dateien, slicerCaches und definedNames aus dem Original zurück.
    
    openpyxl verliert wichtige XML-Namespaces wie xmlns:mc, mc:Ignorable, xmlns:x14 etc.,
    vereinfacht definedNames (entfernt localSheetId Attribute) und verliert Slicers komplett.
    """
    import tempfile
    import shutil
    import zipfile
    import re
    
    sys.stderr.write(f"[restore_ext] START: output_path={output_path}, original_path={original_path}\n")
    
    if not original_path:
        sys.stderr.write(f"[restore_ext] SKIPPED: original_path leer\n")
        return
    
    if not os.path.exists(original_path):
        sys.stderr.write(f"[restore_ext] SKIPPED: original_path existiert nicht: {original_path}\n")
        return
    
    # Wenn original_path == output_path ("Speichern" statt "Speichern unter"),
    # müssen wir eine temporäre Backup-Kopie erstellen, da openpyxl die Datei
    # bereits überschrieben hat. Die Backup-Kopie enthält noch die Original-Daten
    # WENN sie VOR dem openpyxl-Save erstellt wurde (siehe _backup_original_path).
    backup_path = None
    # WINDOWS: normpath normalisiert Pfade (Slashes, Groß/Klein, trailing sep)
    if os.path.normpath(original_path) == os.path.normpath(output_path):
        sys.stderr.write(f"[restore_ext] original_path == output_path — verwende _backup_original_path\n")
        # Prüfe ob ein Backup vor dem Save erstellt wurde
        backup_candidate = getattr(restore_external_links_from_original, '_backup_original_path', None)
        if backup_candidate and os.path.exists(backup_candidate):
            original_path = backup_candidate
            backup_path = backup_candidate
            sys.stderr.write(f"[restore_ext] Verwende Backup: {backup_path}\n")
        else:
            sys.stderr.write(f"[restore_ext] SKIPPED: Kein Backup verfügbar und original==output\n")
            return
    
    sys.stderr.write(f"[restore_ext] Original existiert, starte Wiederherstellung...\n")
    
    temp_dir = None
    orig_temp_dir = None
    
    try:
        temp_dir = tempfile.mkdtemp()
        orig_temp_dir = tempfile.mkdtemp()
        temp_xlsx = os.path.join(temp_dir, 'restored.xlsx')
        
        with zipfile.ZipFile(output_path, 'r') as zf:
            zf.extractall(temp_dir)
        with zipfile.ZipFile(original_path, 'r') as zf:
            # DEBUG: Zeige alle Dateien im Original-ZIP die mit Bildern/Drawings zu tun haben
            all_names = zf.namelist()
            drawing_related = [n for n in all_names if any(k in n.lower() for k in ['draw', 'media', 'image', 'picture', 'vml', 'richdata', 'rdrichvalue'])]
            sys.stderr.write(f"[restore_ext] Original-ZIP Dateien (drawing/media-related): {drawing_related}\n")
            sys.stderr.write(f"[restore_ext] Original-ZIP alle xl/ Dateien: {[n for n in all_names if n.startswith('xl/')]}\n")
            zf.extractall(orig_temp_dir)
        
        # DEBUG: Inhalt von sheet1.xml.rels anzeigen
        orig_rels_file = os.path.join(orig_temp_dir, 'xl', 'worksheets', '_rels', 'sheet1.xml.rels')
        if os.path.exists(orig_rels_file):
            with open(orig_rels_file, 'r', encoding='utf-8') as f:
                rels_content = f.read()
            sys.stderr.write(f"[restore_ext] sheet1.xml.rels Inhalt: {rels_content[:500]}\n")
        
        # DEBUG: Prüfe ob <drawing> Element in original sheet1.xml vorhanden
        orig_sheet1 = os.path.join(orig_temp_dir, 'xl', 'worksheets', 'sheet1.xml')
        if os.path.exists(orig_sheet1):
            with open(orig_sheet1, 'r', encoding='utf-8') as f:
                sheet1_content = f.read()
            has_drawing = '<drawing ' in sheet1_content or '<drawing>' in sheet1_content
            has_legacy = '<legacyDrawing ' in sheet1_content
            has_picture = '<picture ' in sheet1_content
            sys.stderr.write(f"[restore_ext] Original sheet1.xml: <drawing>={has_drawing}, <legacyDrawing>={has_legacy}, <picture>={has_picture}\n")
        
        # DEBUG: Prüfe ob <drawing> Element in output sheet1.xml vorhanden
        dest_sheet1 = os.path.join(temp_dir, 'xl', 'worksheets', 'sheet1.xml')
        if os.path.exists(dest_sheet1):
            with open(dest_sheet1, 'r', encoding='utf-8') as f:
                dest_sheet1_content = f.read()
            has_drawing_dest = '<drawing ' in dest_sheet1_content or '<drawing>' in dest_sheet1_content
            sys.stderr.write(f"[restore_ext] Output sheet1.xml: <drawing>={has_drawing_dest}\n")
        
        ext_links_dir = os.path.join(temp_dir, 'xl', 'externalLinks')
        orig_ext_links_dir = os.path.join(orig_temp_dir, 'xl', 'externalLinks')
        
        fixed_count = 0
        
        if os.path.exists(orig_ext_links_dir) and os.path.exists(ext_links_dir):
            # externalLink-Dateien VOM ORIGINAL kopieren (überschreiben!).
            # openpyxl korrumpiert externalLink-XML beim Round-Trip
            # (fehlende Namespaces, falsche Struktur, beschädigte cached values).
            # Das Original enthält die korrekten, von Excel validierten Versionen.
            # Die Reihenfolge/Anzahl der Dateien bleibt gleich (openpyxl ändert sie nicht),
            # daher stimmen die Indizes in workbook.xml/definedNames weiterhin.
            for f in os.listdir(orig_ext_links_dir):
                if f.startswith('externalLink') and f.endswith('.xml'):
                    orig_file = os.path.join(orig_ext_links_dir, f)
                    dest_file = os.path.join(ext_links_dir, f)
                    shutil.copy2(orig_file, dest_file)
                    sys.stderr.write(f"[restore_ext] externalLink aus Original kopiert: {f}\n")
                    fixed_count += 1
            
            # _rels ebenfalls vom Original kopieren (Konsistenz mit externalLink-Dateien)
            orig_rels_dir = os.path.join(orig_ext_links_dir, '_rels')
            dest_rels_dir = os.path.join(ext_links_dir, '_rels')
            if os.path.exists(orig_rels_dir):
                if os.path.exists(dest_rels_dir):
                    shutil.rmtree(dest_rels_dir)
                shutil.copytree(orig_rels_dir, dest_rels_dir)
                sys.stderr.write(f"[restore_ext] externalLinks/_rels komplett aus Original kopiert\n")
                fixed_count += 1
        elif os.path.exists(orig_ext_links_dir) and not os.path.exists(ext_links_dir):
            # openpyxl hat externalLinks komplett verloren → alles kopieren
            shutil.copytree(orig_ext_links_dir, ext_links_dir)
            sys.stderr.write(f"[restore_ext] externalLinks komplett aus Original kopiert (fehlten)\n")
            fixed_count += 1
        
        # Kopiere slicerCaches aus dem Original (openpyxl verliert Slicers komplett)
        orig_slicer_dir = os.path.join(orig_temp_dir, 'xl', 'slicerCaches')
        dest_slicer_dir = os.path.join(temp_dir, 'xl', 'slicerCaches')
        if os.path.exists(orig_slicer_dir):
            if not os.path.exists(dest_slicer_dir):
                os.makedirs(dest_slicer_dir)
            for f in os.listdir(orig_slicer_dir):
                if f.endswith('.xml'):
                    shutil.copy2(os.path.join(orig_slicer_dir, f), os.path.join(dest_slicer_dir, f))
                    fixed_count += 1
        
        # Kopiere slicers Ordner auch (falls vorhanden)
        orig_slicers_dir = os.path.join(orig_temp_dir, 'xl', 'slicers')
        dest_slicers_dir = os.path.join(temp_dir, 'xl', 'slicers')
        if os.path.exists(orig_slicers_dir):
            if os.path.exists(dest_slicers_dir):
                shutil.rmtree(dest_slicers_dir)
            shutil.copytree(orig_slicers_dir, dest_slicers_dir)
            fixed_count += 1
        
        # sharedStrings.xml NICHT vom Original kopieren!
        # openpyxl erzeugt beim Speichern eine NEUE sharedStrings.xml mit neu
        # nummerierten Indizes. Die Worksheet-Zellen referenzieren diese neuen
        # Indizes (<c t="s"><v>N</v></c>). Wenn wir die Original-sharedStrings.xml
        # zurückkopieren, zeigen die Indizes auf falsche Strings →
        # Excel meldet "Reparatur auf Dateiebene".
        # (openpyxl schreibt reguläre Strings als shared strings, NICHT inline!
        #  Nur CellRichText-Objekte werden als inlineStr geschrieben.)
        sys.stderr.write(f"[restore_ext] sharedStrings.xml: NICHT kopiert (openpyxl-Indizes beibehalten)\n")
        
        # Stelle workbook.xml SELEKTIV wieder her
        # NICHT blind kopieren! openpyxl aktualisiert externalReferences cached values.
        # Nur Namespace-Deklarationen und definedNames vom Original übernehmen.
        workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
        orig_workbook_path = os.path.join(orig_temp_dir, 'xl', 'workbook.xml')
        
        if os.path.exists(workbook_path) and os.path.exists(orig_workbook_path):
            with open(workbook_path, 'r', encoding='utf-8') as f:
                dest_wb_content = f.read()
            with open(orig_workbook_path, 'r', encoding='utf-8') as f:
                orig_wb_content = f.read()
            
            wb_modified = False
            
            # 1. Namespace-Deklarationen vom Original übernehmen
            orig_wb_root = re.search(r'(<workbook\b[^>]+>)', orig_wb_content)
            dest_wb_root = re.search(r'(<workbook\b[^>]+>)', dest_wb_content)
            if orig_wb_root and dest_wb_root and orig_wb_root.group(1) != dest_wb_root.group(1):
                dest_wb_content = dest_wb_content.replace(dest_wb_root.group(1), orig_wb_root.group(1), 1)
                wb_modified = True
                sys.stderr.write(f"[restore_ext] workbook.xml: Namespaces vom Original wiederhergestellt\n")
            
            # 2. definedNames NICHT vom Original übernehmen!
            # definedNames können externe Referenzen enthalten wie [7]Sheet1!$A$1
            # wobei [7] ein Index in <externalReferences> ist.
            # Wenn openpyxl die externalReferences anders nummeriert als das Original,
            # verweisen die Indizes auf falsche Workbooks → Excel-Reparaturmodus.
            # openpyxl's eigene definedNames haben korrekte Indizes zu seinen externalReferences.
            sys.stderr.write(f"[restore_ext] workbook.xml: definedNames von openpyxl beibehalten (Index-Konsistenz)\n")
            
            if wb_modified:
                with open(workbook_path, 'w', encoding='utf-8') as f:
                    f.write(dest_wb_content)
                fixed_count += 1
        
        # Stelle workbook.xml.rels SELEKTIV wieder her
        # NICHT blind kopieren! Nur fehlende Relationships ergänzen (z.B. slicerCaches).
        # externalLinks-Relationships NICHT überschreiben (openpyxl hat korrekte rIds).
        rels_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        orig_rels_path = os.path.join(orig_temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        if os.path.exists(rels_path) and os.path.exists(orig_rels_path):
            with open(rels_path, 'r', encoding='utf-8') as f:
                dest_rels_content = f.read()
            with open(orig_rels_path, 'r', encoding='utf-8') as f:
                orig_rels_content = f.read()
            
            # Sammle alle Targets die openpyxl bereits hat
            dest_targets = set(re.findall(r'Target="([^"]+)"', dest_rels_content))
            
            # Finde fehlende Relationships aus dem Original
            missing_rels = []
            for rel_match in re.finditer(r'<Relationship\s[^>]+/>', orig_rels_content):
                rel_el = rel_match.group(0)
                target_m = re.search(r'Target="([^"]+)"', rel_el)
                if target_m and target_m.group(1) not in dest_targets:
                    target_val = target_m.group(1)
                    # externalLinks NICHT ergänzen (openpyxl hat korrekte cached values)
                    if 'externalLinks/' in target_val:
                        sys.stderr.write(f"[restore_ext] workbook.xml.rels: externalLink übersprungen: {target_val}\n")
                        continue
                    # Nur ergänzen wenn die Zieldatei existiert (oder extern ist)
                    is_external = 'TargetMode="External"' in rel_el
                    if is_external:
                        missing_rels.append(rel_el)
                    else:
                        target_file_check = os.path.normpath(os.path.join(temp_dir, 'xl', target_val))
                        orig_target_check = os.path.normpath(os.path.join(orig_temp_dir, 'xl', target_val))
                        if os.path.exists(target_file_check):
                            missing_rels.append(rel_el)
                        elif os.path.exists(orig_target_check):
                            # Datei aus Original kopieren (z.B. slicerCache)
                            os.makedirs(os.path.dirname(target_file_check), exist_ok=True)
                            shutil.copy2(orig_target_check, target_file_check)
                            missing_rels.append(rel_el)
                            sys.stderr.write(f"[restore_ext] workbook.xml.rels: {target_val} aus Original kopiert\n")
            
            if missing_rels:
                # rId-Konflikte vermeiden: bestehende rIds sammeln
                existing_rids = set(re.findall(r'Id="(rId\d+)"', dest_rels_content))
                max_rid = 0
                for rid in existing_rids:
                    num = int(rid.replace('rId', ''))
                    if num > max_rid:
                        max_rid = num
                
                # Fehlende Rels mit konfliktfreien rIds einfügen
                renumbered_rels = []
                for rel_el in missing_rels:
                    rid_m = re.search(r'Id="(rId\d+)"', rel_el)
                    if rid_m and rid_m.group(1) in existing_rids:
                        # Konflikt: neuen rId vergeben
                        max_rid += 1
                        new_rid = f'rId{max_rid}'
                        old_rid = rid_m.group(1)
                        rel_el = rel_el.replace(f'Id="{old_rid}"', f'Id="{new_rid}"', 1)
                        existing_rids.add(new_rid)
                        sys.stderr.write(f"[restore_ext] workbook.xml.rels: rId-Konflikt {old_rid} → {new_rid}\n")
                    elif rid_m:
                        existing_rids.add(rid_m.group(1))
                    renumbered_rels.append(rel_el)
                
                # Vor </Relationships> einfügen
                insert_str = '\n'.join(renumbered_rels)
                dest_rels_content = dest_rels_content.replace('</Relationships>', insert_str + '\n</Relationships>')
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(dest_rels_content)
                fixed_count += 1
                sys.stderr.write(f"[restore_ext] workbook.xml.rels: {len(renumbered_rels)} fehlende Relationships ergänzt\n")
        
        # [Content_Types].xml Merge wird NACH den Datei-Kopien durchgeführt (s.u.)
        # Grund: Override-Einträge prüfen ob die Datei existiert.
        # Wenn der Merge VOR dem Kopieren von drawings/media/richData läuft,
        # fehlen die Dateien noch → Override wird übersprungen → "Entfernter Teil: Zeichnungsform"
        
        # Kopiere xl/media aus Original (Bilder/Images - openpyxl kann diese verlieren)
        orig_media_dir = os.path.join(orig_temp_dir, 'xl', 'media')
        dest_media_dir = os.path.join(temp_dir, 'xl', 'media')
        if os.path.exists(orig_media_dir):
            media_files = os.listdir(orig_media_dir)
            sys.stderr.write(f"[restore_ext] xl/media im Original gefunden: {len(media_files)} Dateien: {media_files}\n")
            if os.path.exists(dest_media_dir):
                shutil.rmtree(dest_media_dir)
            shutil.copytree(orig_media_dir, dest_media_dir)
            fixed_count += 1
        else:
            sys.stderr.write(f"[restore_ext] xl/media im Original NICHT vorhanden\n")
        
        # Kopiere xl/drawings aus Original (Drawing-XML mit Bild-Positionen/Ankern)
        orig_drawings_dir = os.path.join(orig_temp_dir, 'xl', 'drawings')
        dest_drawings_dir = os.path.join(temp_dir, 'xl', 'drawings')
        if os.path.exists(orig_drawings_dir):
            drawings_files = os.listdir(orig_drawings_dir)
            sys.stderr.write(f"[restore_ext] xl/drawings im Original gefunden: {len(drawings_files)} Dateien: {drawings_files}\n")
            if os.path.exists(dest_drawings_dir):
                shutil.rmtree(dest_drawings_dir)
            shutil.copytree(orig_drawings_dir, dest_drawings_dir)
            fixed_count += 1
        else:
            sys.stderr.write(f"[restore_ext] xl/drawings im Original NICHT vorhanden\n")
        
        # KRITISCH: Kopiere xl/richData aus Original (Excel 365 Zellbilder)
        # Moderne Excel 365 Zellbilder verwenden richData statt drawings:
        # - xl/richData/rdrichvalue.xml (Bild-Werte)
        # - xl/richData/rdRichValueStructure.xml (Struktur)
        # - xl/richData/rdRichValueTypes.xml (Typen)
        # - xl/richData/richValueRel.xml (Beziehungen zu media/)
        # - xl/richData/_rels/richValueRel.xml.rels (tatsächliche Datei-Referenzen)
        # openpyxl kennt richData NICHT und entfernt alles komplett!
        orig_richdata_dir = os.path.join(orig_temp_dir, 'xl', 'richData')
        dest_richdata_dir = os.path.join(temp_dir, 'xl', 'richData')
        if os.path.exists(orig_richdata_dir):
            richdata_files = []
            for root_dir, dirs, files_list in os.walk(orig_richdata_dir):
                for ff in files_list:
                    rel_path = os.path.relpath(os.path.join(root_dir, ff), orig_richdata_dir)
                    richdata_files.append(rel_path)
            sys.stderr.write(f"[restore_ext] xl/richData im Original gefunden: {len(richdata_files)} Dateien: {richdata_files}\n")
            if os.path.exists(dest_richdata_dir):
                shutil.rmtree(dest_richdata_dir)
            shutil.copytree(orig_richdata_dir, dest_richdata_dir)
            fixed_count += 1
        else:
            sys.stderr.write(f"[restore_ext] xl/richData im Original NICHT vorhanden\n")
        
        # KRITISCH: Kopiere xl/metadata.xml aus Original (benötigt für richData/vm-Attribute)
        # metadata.xml definiert die Value Metadata Typen die richData-Zellbilder referenzieren
        orig_metadata = os.path.join(orig_temp_dir, 'xl', 'metadata.xml')
        dest_metadata = os.path.join(temp_dir, 'xl', 'metadata.xml')
        if os.path.exists(orig_metadata):
            shutil.copy2(orig_metadata, dest_metadata)
            sys.stderr.write(f"[restore_ext] xl/metadata.xml aus Original kopiert\n")
            fixed_count += 1
        
        # Kopiere xl/printerSettings aus Original (Druckeinstellungen)
        # Worksheet-Rels aus dem Original referenzieren printerSettings-Dateien.
        # Wenn diese fehlen, erzeugt Excel "Reparatur auf Dateiebene".
        orig_printer_dir = os.path.join(orig_temp_dir, 'xl', 'printerSettings')
        dest_printer_dir = os.path.join(temp_dir, 'xl', 'printerSettings')
        if os.path.exists(orig_printer_dir):
            printer_files = os.listdir(orig_printer_dir)
            sys.stderr.write(f"[restore_ext] xl/printerSettings im Original gefunden: {len(printer_files)} Dateien\n")
            if os.path.exists(dest_printer_dir):
                shutil.rmtree(dest_printer_dir)
            shutil.copytree(orig_printer_dir, dest_printer_dir)
            fixed_count += 1
        
        # =====================================================================
        # Worksheet-Relationships: MERGE statt REPLACE
        # openpyxl nummeriert rIds um. Wenn wir die Original-Rels komplett
        # überschreiben, passen die rId-Referenzen im Worksheet-XML nicht mehr
        # (Hyperlinks, OLE-Objects, Comments, etc. → rId-Mismatch → Reparatur).
        # Stattdessen: openpyxl's Rels behalten, nur FEHLENDE Rels ergänzen.
        # =====================================================================
        orig_ws_rels_dir = os.path.join(orig_temp_dir, 'xl', 'worksheets', '_rels')
        dest_ws_rels_dir = os.path.join(temp_dir, 'xl', 'worksheets', '_rels')
        ws_rid_mappings = {}  # {rels_filename: {orig_rId: dest_rId}}
        
        if os.path.exists(orig_ws_rels_dir):
            if not os.path.exists(dest_ws_rels_dir):
                os.makedirs(dest_ws_rels_dir)
            
            rels_files = [f for f in os.listdir(orig_ws_rels_dir) if f.endswith('.rels')]
            sys.stderr.write(f"[restore_ext] xl/worksheets/_rels MERGE: {len(rels_files)} Dateien: {rels_files}\n")
            
            for rels_fn in rels_files:
                orig_rels_fp = os.path.join(orig_ws_rels_dir, rels_fn)
                dest_rels_fp = os.path.join(dest_ws_rels_dir, rels_fn)
                
                with open(orig_rels_fp, 'r', encoding='utf-8') as f:
                    orig_rels_xml = f.read()
                
                # Parse Original-Rels
                orig_rels = {}
                for m in re.finditer(r'(<Relationship\s[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?/>)', orig_rels_xml):
                    orig_rels[m.group(2)] = {'target': m.group(3), 'el': m.group(1)}
                
                # Parse Dest-Rels (openpyxl's)
                dest_rels = {}
                if os.path.exists(dest_rels_fp):
                    with open(dest_rels_fp, 'r', encoding='utf-8') as f:
                        dest_rels_xml = f.read()
                    for m in re.finditer(r'<Relationship\s[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?/>', dest_rels_xml):
                        dest_rels[m.group(1)] = m.group(2)
                else:
                    dest_rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
                
                # Target→dest_rId Lookup (normalisiert)
                target_to_dest_rid = {}
                for rid, target in dest_rels.items():
                    norm = target.replace('\\', '/').lower()
                    target_to_dest_rid[norm] = rid
                
                # Mapping bauen und fehlende Rels finden
                mapping = {}
                missing_rels = []
                existing_rids = set(dest_rels.keys())
                max_rid_num = 0
                for rid in existing_rids:
                    num_m = re.search(r'\d+', rid)
                    if num_m:
                        max_rid_num = max(max_rid_num, int(num_m.group(0)))
                
                for orig_rid, orig_info in orig_rels.items():
                    norm_target = orig_info['target'].replace('\\', '/').lower()
                    if norm_target in target_to_dest_rid:
                        # Rel existiert in dest → mapping
                        mapping[orig_rid] = target_to_dest_rid[norm_target]
                    else:
                        # Rel fehlt in dest → mit neuem rId ergänzen
                        max_rid_num += 1
                        new_rid = f'rId{max_rid_num}'
                        mapping[orig_rid] = new_rid
                        
                        # Neues Relationship-Element mit neuem rId
                        new_el = re.sub(r'Id="[^"]+"', f'Id="{new_rid}"', orig_info['el'])
                        missing_rels.append(new_el)
                        existing_rids.add(new_rid)
                        
                        # Zieldatei aus Original kopieren falls nötig
                        target_path = orig_info['target']
                        dest_file = os.path.normpath(os.path.join(temp_dir, 'xl', 'worksheets', target_path))
                        orig_file = os.path.normpath(os.path.join(orig_temp_dir, 'xl', 'worksheets', target_path))
                        if not os.path.exists(dest_file) and os.path.exists(orig_file):
                            os.makedirs(os.path.dirname(dest_file), exist_ok=True)
                            shutil.copy2(orig_file, dest_file)
                            sys.stderr.write(f"[restore_ext] MERGE {rels_fn}: Zieldatei kopiert: {target_path}\n")
                
                if missing_rels:
                    insert_str = '\n'.join(missing_rels)
                    dest_rels_xml = dest_rels_xml.replace('</Relationships>', insert_str + '\n</Relationships>')
                    with open(dest_rels_fp, 'w', encoding='utf-8') as f:
                        f.write(dest_rels_xml)
                    fixed_count += 1
                    sys.stderr.write(f"[restore_ext] MERGE {rels_fn}: {len(missing_rels)} fehlende Rels ergänzt\n")
                elif not os.path.exists(dest_rels_fp):
                    # Original hat Rels, dest nicht → komplett schreiben
                    with open(dest_rels_fp, 'w', encoding='utf-8') as f:
                        f.write(dest_rels_xml)
                    fixed_count += 1
                
                ws_rid_mappings[rels_fn] = mapping
                sys.stderr.write(f"[restore_ext] MERGE {rels_fn}: Mapping={mapping}\n")
        else:
            sys.stderr.write(f"[restore_ext] xl/worksheets/_rels im Original NICHT vorhanden\n")
        
        # =====================================================================
        # [Content_Types].xml SELEKTIV wiederherstellen
        # MUSS NACH allen Datei-Kopien laufen! Override-Einträge prüfen ob die
        # Datei existiert. Wenn dieser Block VOR dem Kopieren von drawings/media/
        # richData läuft, fehlen die Dateien → Override übersprungen →
        # "Entfernter Teil: Zeichnungsform"
        # =====================================================================
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
        orig_content_types_path = os.path.join(orig_temp_dir, '[Content_Types].xml')
        if os.path.exists(content_types_path) and os.path.exists(orig_content_types_path):
            with open(content_types_path, 'r', encoding='utf-8') as f:
                dest_ct_content = f.read()
            with open(orig_content_types_path, 'r', encoding='utf-8') as f:
                orig_ct_content = f.read()
            
            ct_modified = False
            
            # A. Fehlende <Default Extension="..."> Einträge ergänzen
            # KRITISCH für VML-Zeichnungen (.vml), Metafiles (.emf, .wmf) etc.
            dest_extensions = set(re.findall(r'<Default\s+Extension="([^"]+)"', dest_ct_content))
            missing_defaults = []
            for df_match in re.finditer(r'<Default\s[^>]*/>', orig_ct_content):
                df_el = df_match.group(0)
                ext_m = re.search(r'Extension="([^"]+)"', df_el)
                if ext_m and ext_m.group(1) not in dest_extensions:
                    missing_defaults.append(df_el)
                    sys.stderr.write(f"[restore_ext] ContentTypes Default: Extension=\"{ext_m.group(1)}\" ergänzt\n")
            
            if missing_defaults:
                first_override = re.search(r'<Override\s', dest_ct_content)
                if first_override:
                    insert_str = '\n'.join(missing_defaults) + '\n'
                    dest_ct_content = dest_ct_content[:first_override.start()] + insert_str + dest_ct_content[first_override.start():]
                else:
                    insert_str = '\n'.join(missing_defaults)
                    dest_ct_content = dest_ct_content.replace('</Types>', insert_str + '\n</Types>')
                ct_modified = True
            
            # B. Fehlende <Override PartName="..."> Einträge ergänzen
            dest_parts = set(re.findall(r'PartName="([^"]+)"', dest_ct_content))
            missing_overrides = []
            for ov_match in re.finditer(r'<Override\s[^>]*/>', orig_ct_content):
                ov_el = ov_match.group(0)
                pn_m = re.search(r'PartName="([^"]+)"', ov_el)
                if pn_m and pn_m.group(1) not in dest_parts:
                    part_name = pn_m.group(1)
                    # Nur ergänzen wenn die Datei tatsächlich existiert
                    rel_path_ct = part_name.lstrip('/')
                    dest_file_check = os.path.join(temp_dir, rel_path_ct.replace('/', os.sep))
                    if os.path.exists(dest_file_check):
                        missing_overrides.append(ov_el)
                        sys.stderr.write(f"[restore_ext] ContentTypes Override: {part_name} ergänzt\n")
            
            if missing_overrides:
                insert_str = '\n'.join(missing_overrides)
                dest_ct_content = dest_ct_content.replace('</Types>', insert_str + '\n</Types>')
                ct_modified = True
            
            if ct_modified:
                with open(content_types_path, 'w', encoding='utf-8') as f:
                    f.write(dest_ct_content)
                fixed_count += 1
                sys.stderr.write(f"[restore_ext] [Content_Types].xml: {len(missing_defaults)} Default + {len(missing_overrides)} Override Einträge ergänzt\n")
        
        # KRITISCH: <drawing> und <legacyDrawing> Elemente in Worksheet-XMLs wiederherstellen
        # openpyxl ENTFERNT diese Elemente beim Speichern wenn Pillow nicht installiert ist.
        # Ohne <drawing r:id="..."/> im Worksheet-XML zeigt Excel keine Bilder an,
        # selbst wenn xl/media/, xl/drawings/ und _rels wiederhergestellt wurden.
        orig_ws_dir = os.path.join(orig_temp_dir, 'xl', 'worksheets')
        dest_ws_dir = os.path.join(temp_dir, 'xl', 'worksheets')
        if os.path.exists(orig_ws_dir) and os.path.exists(dest_ws_dir):
            for ws_file in os.listdir(orig_ws_dir):
                if not ws_file.endswith('.xml'):
                    continue
                orig_ws_file = os.path.join(orig_ws_dir, ws_file)
                dest_ws_file = os.path.join(dest_ws_dir, ws_file)
                if not os.path.exists(dest_ws_file):
                    continue
                
                with open(orig_ws_file, 'r', encoding='utf-8') as f:
                    orig_ws_content = f.read()
                with open(dest_ws_file, 'r', encoding='utf-8') as f:
                    dest_ws_content = f.read()
                
                ws_modified = False
                
                # Rels-Mapping für dieses Sheet laden
                rels_fn = f"{ws_file}.rels"
                mapping = ws_rid_mappings.get(rels_fn, {})
                
                def _map_rid(element_str, _mapping=mapping):
                    """Ersetzt rId-Referenzen im Element mit gemappten rIds."""
                    def _rid_replacer(m):
                        orig_rid = m.group(1)
                        mapped = _mapping.get(orig_rid, orig_rid)
                        return f'r:id="{mapped}"'
                    return re.sub(r'r:id="([^"]+)"', _rid_replacer, element_str)
                
                # =====================================================================
                # Worksheet-Elemente: NUR ERGÄNZEN wenn openpyxl sie entfernt hat.
                # openpyxl's Rels wurden beibehalten (nur fehlende ergänzt).
                # Daher sind openpyxl's rId-Referenzen (tableParts, pageSetup,
                # hyperlinks, etc.) korrekt. Wir ersetzen sie NICHT mehr.
                # Nur fehlende Elemente (drawing, legacyDrawing, picture) werden
                # mit gemappten rIds aus dem Original eingefügt.
                # =====================================================================
                
                # <drawing> nur ergänzen wenn openpyxl es entfernt hat
                drawing_match = re.search(r'<drawing\s+[^>]*/\s*>', orig_ws_content)
                if not drawing_match:
                    drawing_match = re.search(r'<drawing\s+[^>]*>.*?</drawing>', orig_ws_content, re.DOTALL)
                if drawing_match:
                    if not re.search(r'<drawing[\s>]', dest_ws_content):
                        # openpyxl hat <drawing> entfernt → mit gemapptem rId ergänzen
                        drawing_el = _map_rid(drawing_match.group(0))
                        dest_ws_content = _insert_ws_element(dest_ws_content, drawing_el, 'drawing')
                        ws_modified = True
                        sys.stderr.write(f"[restore] {ws_file}: <drawing> ergänzt (fehlte): {drawing_el}\n")
                    else:
                        sys.stderr.write(f"[restore] {ws_file}: <drawing> bereits vorhanden (openpyxl-rId beibehalten)\n")
                
                # <legacyDrawing> nur ergänzen wenn fehlend
                legacy_match = re.search(r'<legacyDrawing\s+[^>]*/\s*>', orig_ws_content)
                if not legacy_match:
                    legacy_match = re.search(r'<legacyDrawing\s+[^>]*>.*?</legacyDrawing>', orig_ws_content, re.DOTALL)
                if legacy_match:
                    if not re.search(r'<legacyDrawing[\s>]', dest_ws_content):
                        legacy_el = _map_rid(legacy_match.group(0))
                        dest_ws_content = _insert_ws_element(dest_ws_content, legacy_el, 'legacyDrawing')
                        ws_modified = True
                        sys.stderr.write(f"[restore] {ws_file}: <legacyDrawing> ergänzt (fehlte)\n")
                
                # <picture> nur ergänzen wenn fehlend (Hintergrundbilder)
                picture_match = re.search(r'<picture\s+[^>]*/\s*>', orig_ws_content)
                if picture_match:
                    if not re.search(r'<picture[\s>]', dest_ws_content):
                        picture_el = _map_rid(picture_match.group(0))
                        dest_ws_content = _insert_ws_element(dest_ws_content, picture_el, 'picture')
                        ws_modified = True
                
                # tableParts und pageSetup NICHT ersetzen!
                # openpyxl's Rels sind beibehalten, daher stimmen openpyxl's rId-Referenzen
                # in tableParts, pageSetup, hyperlinks, oleObjects, comments etc.
                # Früher mussten diese ersetzt werden weil Rels komplett überschrieben wurden.
                
                # KRITISCH: Namespace-Deklarationen vom Original-Worksheet wiederherstellen
                # openpyxl schreibt nur minimal: xmlns="..." xmlns:r="..."
                # Excel benötigt aber: xmlns:mc, mc:Ignorable, xmlns:x14ac, xmlns:xr etc.
                # Ohne diese Namespaces erkennt Excel vm-Attribute und andere Erweiterungen nicht.
                # re.search statt re.match weil XML-Header vor <worksheet> stehen kann
                orig_root_match = re.search(r'(<worksheet\b[^>]+>)', orig_ws_content)
                dest_root_match = re.search(r'(<worksheet\b[^>]+>)', dest_ws_content)
                if orig_root_match and dest_root_match:
                    orig_root = orig_root_match.group(1)
                    dest_root = dest_root_match.group(1)
                    if orig_root != dest_root:
                        dest_ws_content = dest_ws_content.replace(dest_root, orig_root, 1)
                        ws_modified = True
                        sys.stderr.write(f"[restore] {ws_file}: Worksheet-Namespaces vom Original wiederhergestellt\n")
                
                # KRITISCH: vm-Attribute auf Zellen wiederherstellen (Excel 365 Zellbilder)
                # openpyxl kennt vm-Attribute NICHT und entfernt sie komplett beim Speichern.
                # vm="N" auf <c>-Elementen verweist auf xl/metadata.xml → xl/richData/ → Bilder
                # Ohne vm-Attribute werden Zellbilder trotz vorhandener richData nicht angezeigt.
                if 'vm=' in orig_ws_content:
                    # Sammle alle Zellen mit vm-Attribut aus dem Original
                    vm_cells = {}
                    # Zwei mögliche Attribut-Reihenfolgen: r="..." vm="..." oder vm="..." r="..."
                    for vm_match in re.finditer(r'<c\s[^>]*?r="([A-Z]+\d+)"[^>]*?\bvm="(\d+)"', orig_ws_content):
                        vm_cells[vm_match.group(1)] = vm_match.group(2)
                    for vm_match in re.finditer(r'<c\s[^>]*?\bvm="(\d+)"[^>]*?r="([A-Z]+\d+)"', orig_ws_content):
                        if vm_match.group(2) not in vm_cells:
                            vm_cells[vm_match.group(2)] = vm_match.group(1)
                    
                    if vm_cells:
                        vm_restored = 0
                        vm_created = 0
                        for cell_ref, vm_val in vm_cells.items():
                            # Finde die Zelle in der Zieldatei und füge vm-Attribut hinzu
                            cell_pattern = re.compile(r'(<c\s[^>]*?r="' + re.escape(cell_ref) + '"[^>]*?)(/?>)')
                            match = cell_pattern.search(dest_ws_content)
                            if match and 'vm=' not in match.group(0):
                                new_attr = match.group(1) + f' vm="{vm_val}"' + match.group(2)
                                dest_ws_content = dest_ws_content[:match.start()] + new_attr + dest_ws_content[match.end():]
                                vm_restored += 1
                                ws_modified = True
                            elif not match:
                                # Zelle existiert nicht in Zieldatei (openpyxl hat sie entfernt)
                                # Erstelle das Zell-Element in der passenden Zeile
                                row_num_match = re.search(r'(\d+)$', cell_ref)
                                if row_num_match:
                                    row_num = row_num_match.group(1)
                                    row_pattern = re.compile(r'(<row\s[^>]*?\br="' + re.escape(row_num) + '"[^>]*?>)')
                                    row_match = row_pattern.search(dest_ws_content)
                                    if row_match:
                                        cell_el = f'<c r="{cell_ref}" vm="{vm_val}"/>'
                                        insert_pos = row_match.end()
                                        dest_ws_content = dest_ws_content[:insert_pos] + cell_el + dest_ws_content[insert_pos:]
                                        vm_created += 1
                                        ws_modified = True
                                    else:
                                        # Zeile existiert auch nicht → Zeile UND Zelle in sheetData erstellen
                                        # Dies passiert wenn openpyxl oder fix_xlsx_relationships leere Zeilen entfernt hat
                                        sheet_data_end = re.search(r'</sheetData>', dest_ws_content)
                                        if sheet_data_end:
                                            row_el = f'<row r="{row_num}"><c r="{cell_ref}" vm="{vm_val}"/></row>'
                                            insert_pos = sheet_data_end.start()
                                            dest_ws_content = dest_ws_content[:insert_pos] + row_el + '\n' + dest_ws_content[insert_pos:]
                                            vm_created += 1
                                            ws_modified = True
                                            sys.stderr.write(f"[restore] {ws_file}: Zeile {row_num} + Zelle {cell_ref} mit vm={vm_val} neu erstellt\n")
                        if vm_restored > 0 or vm_created > 0:
                            sys.stderr.write(f"[restore] {ws_file}: {vm_restored} vm-Attribute wiederhergestellt, {vm_created} vm-Zellen neu erstellt (von {len(vm_cells)} im Original)\n")
                
                if ws_modified:
                    with open(dest_ws_file, 'w', encoding='utf-8') as f:
                        f.write(dest_ws_content)
                    fixed_count += 1
        
        # =====================================================================
        # CONTENT_TYPES KONSISTENZ: Fehlende referenzierte Dateien nachkopieren
        # openpyxl erzeugt nicht alle Dateien des Originals (z.B. calcChain.xml).
        # Wenn [Content_Types].xml vom Original kopiert wird aber Dateien fehlen,
        # löst Excel den "Reparatur"-Modus aus → richData/Bilder werden entfernt!
        # =====================================================================
        ct_file = os.path.join(temp_dir, '[Content_Types].xml')
        if os.path.exists(ct_file):
            with open(ct_file, 'r', encoding='utf-8') as f:
                ct_content = f.read()
            
            def _ct_override_fixer(m):
                nonlocal fixed_count
                part_name = m.group(1)  # z.B. "/xl/calcChain.xml"
                rel_path = part_name.lstrip('/')
                dest_file_ct = os.path.join(temp_dir, rel_path.replace('/', os.sep))
                if os.path.exists(dest_file_ct):
                    return m.group(0)  # Datei existiert, Eintrag behalten
                # externalLinks NIEMALS aus Original kopieren (stale cached values!)
                if 'externalLinks/' in part_name:
                    sys.stderr.write(f"[restore_ext] ContentTypes-Konsistenz: {part_name} entfernt (externalLink, nicht kopieren)\n")
                    return ''
                # Datei fehlt — aus Original kopieren
                orig_file_ct = os.path.join(orig_temp_dir, rel_path.replace('/', os.sep))
                if os.path.exists(orig_file_ct):
                    os.makedirs(os.path.dirname(dest_file_ct), exist_ok=True)
                    shutil.copy2(orig_file_ct, dest_file_ct)
                    sys.stderr.write(f"[restore_ext] ContentTypes-Konsistenz: {part_name} aus Original kopiert\n")
                    fixed_count += 1
                    return m.group(0)  # Datei jetzt vorhanden, Eintrag behalten
                else:
                    # Datei existiert nirgends — Eintrag entfernen
                    sys.stderr.write(f"[restore_ext] ContentTypes-Konsistenz: {part_name} entfernt (nirgends vorhanden)\n")
                    return ''
            
            new_ct = re.sub(r'<Override\s+PartName="(/[^"]+)"[^>]*/>\s*', _ct_override_fixer, ct_content)
            if new_ct != ct_content:
                with open(ct_file, 'w', encoding='utf-8') as f:
                    f.write(new_ct)
                sys.stderr.write(f"[restore_ext] [Content_Types].xml bereinigt\n")
        
        # WORKBOOK.XML.RELS KONSISTENZ: Fehlende referenzierte Dateien nachkopieren
        wb_rels_file = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        if os.path.exists(wb_rels_file):
            with open(wb_rels_file, 'r', encoding='utf-8') as f:
                wb_rels_ct = f.read()
            
            def _rels_fixer(m):
                nonlocal fixed_count
                full_el = m.group(0)
                target = m.group(1)
                # Externe URLs und TargetMode="External" überspringen
                if 'TargetMode="External"' in full_el:
                    return full_el
                if target.startswith('http://') or target.startswith('https://') or target.startswith('mailto:'):
                    return full_el
                # Relativer Pfad: relativ zu xl/
                target_file_r = os.path.normpath(os.path.join(temp_dir, 'xl', target))
                if os.path.exists(target_file_r):
                    return full_el  # Datei existiert
                # externalLinks NIEMALS aus Original kopieren (stale cached values!)
                if 'externalLinks/' in target:
                    sys.stderr.write(f"[restore_ext] Rels-Konsistenz: xl/{target} entfernt (externalLink, nicht kopieren)\n")
                    return ''
                orig_target_r = os.path.normpath(os.path.join(orig_temp_dir, 'xl', target))
                if os.path.exists(orig_target_r):
                    os.makedirs(os.path.dirname(target_file_r), exist_ok=True)
                    shutil.copy2(orig_target_r, target_file_r)
                    sys.stderr.write(f"[restore_ext] Rels-Konsistenz: xl/{target} aus Original kopiert\n")
                    fixed_count += 1
                    return full_el
                else:
                    sys.stderr.write(f"[restore_ext] Rels-Konsistenz: xl/{target} entfernt (nicht gefunden)\n")
                    return ''
            
            new_rels = re.sub(r'<Relationship\s[^>]*?Target="([^"]+)"[^>]*/>\s*', _rels_fixer, wb_rels_ct)
            if new_rels != wb_rels_ct:
                with open(wb_rels_file, 'w', encoding='utf-8') as f:
                    f.write(new_rels)
                sys.stderr.write(f"[restore_ext] workbook.xml.rels bereinigt\n")
        
        # WORKSHEET RELS KONSISTENZ: Fehlende referenzierte Dateien prüfen
        # Worksheet-Rels wurden aus dem Original kopiert und referenzieren evtl.
        # Dateien (printerSettings, comments, ctrlProps etc.) die im Output fehlen.
        ws_rels_check_dir = os.path.join(temp_dir, 'xl', 'worksheets', '_rels')
        if os.path.exists(ws_rels_check_dir):
            for ws_rels_fn in os.listdir(ws_rels_check_dir):
                if not ws_rels_fn.endswith('.rels'):
                    continue
                ws_rels_fp = os.path.join(ws_rels_check_dir, ws_rels_fn)
                with open(ws_rels_fp, 'r', encoding='utf-8') as f:
                    ws_rels_content = f.read()
                
                def _ws_rels_fixer(m, _fn=ws_rels_fn):
                    nonlocal fixed_count
                    full_el = m.group(0)
                    target = m.group(1)
                    if 'TargetMode="External"' in full_el:
                        return full_el
                    if target.startswith('http://') or target.startswith('https://') or target.startswith('mailto:'):
                        return full_el
                    # Targets sind relativ zu xl/worksheets/
                    target_file = os.path.normpath(os.path.join(temp_dir, 'xl', 'worksheets', target))
                    if os.path.exists(target_file):
                        return full_el
                    # Aus Original kopieren
                    orig_target = os.path.normpath(os.path.join(orig_temp_dir, 'xl', 'worksheets', target))
                    if os.path.exists(orig_target):
                        os.makedirs(os.path.dirname(target_file), exist_ok=True)
                        shutil.copy2(orig_target, target_file)
                        sys.stderr.write(f"[restore_ext] WS-Rels-Konsistenz: {_fn} → {target} aus Original kopiert\n")
                        fixed_count += 1
                        return full_el
                    else:
                        sys.stderr.write(f"[restore_ext] WS-Rels-Konsistenz: {_fn} → {target} entfernt (nicht gefunden)\n")
                        return ''
                
                new_ws_rels = re.sub(r'<Relationship\s[^>]*?Target="([^"]+)"[^>]*/>\s*', _ws_rels_fixer, ws_rels_content)
                if new_ws_rels != ws_rels_content:
                    with open(ws_rels_fp, 'w', encoding='utf-8') as f:
                        f.write(new_ws_rels)
                    sys.stderr.write(f"[restore_ext] {ws_rels_fn} bereinigt\n")
        
        sys.stderr.write(f"[restore_ext] fixed_count={fixed_count}\n")
        
        if fixed_count > 0:
            
            # Erstelle neue XLSX
            with zipfile.ZipFile(temp_xlsx, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(temp_dir):
                    dirs[:] = [d for d in dirs if d != '__MACOSX']
                    for f in files:
                        if f == 'restored.xlsx' or f == '.DS_Store' or f.startswith('._'):
                            continue
                        full_path = os.path.join(root, f)
                        arc_name = os.path.relpath(full_path, temp_dir).replace('\\', '/')
                        zf.write(full_path, arc_name)
            
            shutil.copy2(temp_xlsx, output_path)
            sys.stderr.write(f"[restore_ext] XLSX wiederhergestellt und gespeichert\n")
        
        # DIAGNOSE: Dump des endgültigen ZIP-Inhalts für Image-Debugging
        # Läuft IMMER (auch wenn fixed_count == 0), um den Endzustand zu zeigen
        try:
            with zipfile.ZipFile(output_path, 'r') as zf:
                all_names = zf.namelist()
                img_related = [n for n in all_names if any(k in n.lower() for k in 
                    ['draw', 'media', 'image', 'picture', 'vml', 'richdata', 'rdrichvalue', 'metadata', '_rels/sheet'])]
                sys.stderr.write(f"[DIAGNOSE] === FINAL ZIP STATE ===\n")
                sys.stderr.write(f"[DIAGNOSE] Image-relevante Dateien ({len(img_related)}):\n")
                for n in sorted(img_related):
                    size = zf.getinfo(n).file_size
                    sys.stderr.write(f"[DIAGNOSE]   {n} ({size} bytes)\n")
                
                # Sheet XMLs prüfen
                for n in sorted(all_names):
                    if n.startswith('xl/worksheets/') and n.endswith('.xml') and '/_rels/' not in n:
                        sheet_xml = zf.read(n).decode('utf-8', errors='replace')
                        has_drawing = bool(re.search(r'<drawing[\s>]', sheet_xml))
                        has_tp = bool(re.search(r'<tableParts', sheet_xml))
                        has_vm = bool(re.search(r'\bvm=', sheet_xml))
                        has_mc = 'mc:Ignorable' in sheet_xml
                        
                        # rId-Werte extrahieren
                        dr_rid = re.search(r'<drawing r:id="(rId\d+)"', sheet_xml)
                        tp_rids = re.findall(r'<tablePart[^>]*r:id="(rId\d+)"', sheet_xml)
                        ps_rid = re.search(r'<pageSetup[^>]*r:id="(rId\d+)"', sheet_xml)
                        vm_vals = re.findall(r'\bvm="(\d+)"', sheet_xml)
                        
                        # Namespace im Root
                        root_m = re.search(r'(<worksheet\b[^>]{0,100})', sheet_xml)
                        root_preview = root_m.group(1) if root_m else 'N/A'
                        
                        sys.stderr.write(f"[DIAGNOSE] {n}: drawing={has_drawing}({dr_rid.group(1) if dr_rid else '-'}), "
                                        f"tableParts={has_tp}({','.join(tp_rids) if tp_rids else '-'}), "
                                        f"vm={has_vm}({','.join(vm_vals[:3]) if vm_vals else '-'}), "
                                        f"mc:Ignorable={has_mc}, "
                                        f"pageSetup_rid={ps_rid.group(1) if ps_rid else '-'}\n")
                        sys.stderr.write(f"[DIAGNOSE]   root: {root_preview}...\n")
                        
                        # Rels für dieses Sheet prüfen
                        sheet_name_only = n.split('/')[-1]
                        rels_name = f"xl/worksheets/_rels/{sheet_name_only}.rels"
                        if rels_name in all_names:
                            rels_xml = zf.read(rels_name).decode('utf-8', errors='replace')
                            rels_entries = re.findall(r'Id="(rId\d+)"[^>]*Type="[^"]*?/(\w+)"[^>]*Target="([^"]*)"', rels_xml)
                            sys.stderr.write(f"[DIAGNOSE]   rels ({rels_name}):\n")
                            for rid, rtype, target in rels_entries:
                                sys.stderr.write(f"[DIAGNOSE]     {rid} -> {rtype} ({target})\n")
                        else:
                            sys.stderr.write(f"[DIAGNOSE]   KEINE RELS DATEI für {sheet_name_only}\n")
                
                sys.stderr.write(f"[DIAGNOSE] === END ===\n")
        except Exception as diag_err:
            sys.stderr.write(f"[DIAGNOSE] Fehler: {diag_err}\n")
    
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
        if orig_temp_dir:
            shutil.rmtree(orig_temp_dir, ignore_errors=True)


def apply_tint(rgb_hex, tint):
    """
    Wendet einen Tint auf eine RGB-Farbe an.
    Tint > 0: heller (Richtung weiß)
    Tint < 0: dunkler (Richtung schwarz)
    """
    if not rgb_hex or len(rgb_hex) < 6:
        return rgb_hex
    
    # Parse RGB
    r = int(rgb_hex[0:2], 16)
    g = int(rgb_hex[2:4], 16)
    b = int(rgb_hex[4:6], 16)
    
    if tint > 0:
        # Aufhellen (Richtung weiß)
        r = int(r + (255 - r) * tint)
        g = int(g + (255 - g) * tint)
        b = int(b + (255 - b) * tint)
    elif tint < 0:
        # Abdunkeln (Richtung schwarz)
        r = int(r * (1 + tint))
        g = int(g * (1 + tint))
        b = int(b * (1 + tint))
    
    # Clamp to 0-255
    r = max(0, min(255, r))
    g = max(0, min(255, g))
    b = max(0, min(255, b))
    
    return f'{r:02X}{g:02X}{b:02X}'

def theme_color_to_rgb(color, workbook=None):
    """
    Konvertiert eine Theme-Farbe zu RGB.
    
    Args:
        color: openpyxl Color Objekt
        workbook: Workbook für Theme-Lookup (optional)
    
    Returns:
        RGB Hex-String (z.B. 'FF0000') oder None
    """
    if not color:
        return None
    
    color_type = getattr(color, 'type', None)
    
    if color_type == 'rgb':
        rgb = color.rgb
        if isinstance(rgb, str) and len(rgb) >= 6:
            # Entferne Alpha wenn vorhanden (ARGB -> RGB)
            if len(rgb) == 8:
                return rgb[2:]
            return rgb
        return None
    
    if color_type == 'theme':
        theme_idx = color.theme
        tint = getattr(color, 'tint', 0) or 0
        
        # Hole Basis-Farbe aus Theme
        if theme_idx is not None and 0 <= theme_idx < len(THEME_COLORS):
            base_rgb = THEME_COLORS[theme_idx]
            # Wende Tint an
            return apply_tint(base_rgb, tint)
        return None
    
    if color_type == 'indexed':
        # Indexed colors - verwende Standard-Palette
        # Für einfache Fälle
        indexed = getattr(color, 'indexed', None)
        if indexed == 9:  # Weiß
            return 'FFFFFF'
        elif indexed == 8:  # Schwarz
            return '000000'
        # Andere indexed colors erstmal ignorieren
        return None
    
    return None

def convert_fill_to_rgb(fill):
    """
    Konvertiert ein Fill-Objekt mit Theme-Farben zu einem Fill mit RGB-Farben.
    Dies ist nötig weil openpyxl Theme-Farben nicht korrekt schreibt.
    
    WICHTIG: Pattern-Typen wie gray125 mit Theme-Farben werden zu solid konvertiert,
    da das Muster sonst nicht korrekt dargestellt wird.
    """
    if not fill or fill.patternType is None:
        return fill
    
    fg_rgb = None
    bg_rgb = None
    
    if fill.fgColor:
        fg_rgb = theme_color_to_rgb(fill.fgColor)
    if fill.bgColor:
        bg_rgb = theme_color_to_rgb(fill.bgColor)
    
    # Wenn keine Konvertierung nötig (schon RGB und solid), gib Original zurück
    fg_type = getattr(fill.fgColor, 'type', None) if fill.fgColor else None
    bg_type = getattr(fill.bgColor, 'type', None) if fill.bgColor else None
    
    if fg_type == 'rgb' and (bg_type == 'rgb' or bg_type is None) and fill.patternType == 'solid':
        return fill
    
    # Pattern-Typ: gray125 oder andere Muster mit Theme-Farben -> solid
    # Denn das Muster-Rendering hängt von der Theme-Definition ab
    pattern_type = fill.patternType
    if pattern_type and pattern_type != 'solid' and fg_type == 'theme':
        pattern_type = 'solid'  # Konvertiere zu solid fill
    
    # Erstelle neues Fill mit RGB-Farben
    new_fill = PatternFill(
        patternType=pattern_type,
        fgColor=Color(rgb='FF' + fg_rgb) if fg_rgb else None,
        bgColor=Color(rgb='FF' + bg_rgb) if bg_rgb else None
    )
    
    return new_fill

def convert_font_to_rgb(font):
    """
    Konvertiert ein Font-Objekt mit Theme-Farben zu einem Font mit RGB-Farben.
    """
    if not font:
        return font
    
    if not font.color:
        return font
    
    color_type = getattr(font.color, 'type', None)
    if color_type == 'rgb':
        return font  # Schon RGB
    
    rgb = theme_color_to_rgb(font.color)
    if not rgb:
        return font  # Konnte nicht konvertieren
    
    # Erstelle neuen Font mit RGB-Farbe
    new_font = Font(
        name=font.name,
        size=font.size,
        bold=font.bold,
        italic=font.italic,
        underline=font.underline,
        strike=font.strike,
        color=Color(rgb='FF' + rgb)
    )
    
    return new_font

# xlwings-Unterstützung (optional, für strukturelle Änderungen mit CF-Erhalt)
try:
    from excel_utils import is_excel_installed, structural_change_with_excel
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    def is_excel_installed():
        return False
    def structural_change_with_excel(*args, **kwargs):
        return False


def hex_to_argb(hex_color):
    """Konvertiert Hex ('#FF0000') zu ARGB ('FFFF0000')"""
    if not hex_color:
        return None
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    if len(hex_color) == 6:
        return 'FF' + hex_color.upper()
    return hex_color.upper()


def shift_cell_reference(cell_ref, deleted_col_indices, inserted_cols=None):
    """
    Verschiebt eine Zell-Referenz basierend auf gelöschten/eingefügten Spalten.
    
    Args:
        cell_ref: Zell-Referenz wie 'A1' oder 'AB123'
        deleted_col_indices: Liste der gelöschten Spalten-Indices (0-basiert)
        inserted_cols: Dict mit {position: count} für eingefügte Spalten
    
    Returns:
        Neue Zell-Referenz oder None wenn die Zelle gelöscht wurde
    """
    if not cell_ref:
        return cell_ref
    
    # Parse Zell-Referenz
    match = re.match(r'^([A-Z]+)(\d+)$', cell_ref.upper())
    if not match:
        return cell_ref
    
    col_letter = match.group(1)
    row_num = match.group(2)
    col_idx = column_index_from_string(col_letter) - 1  # 0-basiert
    
    # Prüfe ob Spalte gelöscht wurde
    if deleted_col_indices and col_idx in deleted_col_indices:
        return None
    
    # Berechne Verschiebung
    shift = 0
    
    # Verschiebung durch gelöschte Spalten (die VOR dieser Spalte lagen)
    if deleted_col_indices:
        for del_idx in sorted(deleted_col_indices):
            if del_idx < col_idx:
                shift -= 1
    
    # Verschiebung durch eingefügte Spalten
    if inserted_cols:
        for pos, count in inserted_cols.items():
            if pos <= col_idx:
                shift += count
    
    new_col_idx = col_idx + shift
    if new_col_idx < 0:
        return None
    
    new_col_letter = get_column_letter(new_col_idx + 1)
    return f"{new_col_letter}{row_num}"


def shift_range_reference(range_ref, deleted_col_indices, inserted_cols=None):
    """
    Verschiebt einen Bereichs-Referenz wie 'A1:C10'.
    
    Returns:
        Neuen Bereich oder None wenn der Bereich komplett gelöscht wurde
    """
    if not range_ref:
        return range_ref
    
    # Handle mehrere Bereiche (z.B. "A1:B2 C3:D4")
    parts = range_ref.split()
    new_parts = []
    
    for part in parts:
        if ':' in part:
            # Bereich wie A1:C10
            start, end = part.split(':')
            new_start = shift_cell_reference(start, deleted_col_indices, inserted_cols)
            new_end = shift_cell_reference(end, deleted_col_indices, inserted_cols)
            
            if new_start and new_end:
                new_parts.append(f"{new_start}:{new_end}")
        else:
            # Einzelne Zelle
            new_ref = shift_cell_reference(part, deleted_col_indices, inserted_cols)
            if new_ref:
                new_parts.append(new_ref)
    
    return ' '.join(new_parts) if new_parts else None


def adjust_tables(ws, deleted_col_indices, inserted_cols=None, new_headers=None):
    """
    Passt alle Excel-Tabellen (Tables) an wenn Spalten gelöscht/eingefügt werden.
    
    WICHTIG: openpyxl's insert_cols/delete_cols passt Table-Ranges NICHT automatisch an!
    
    Args:
        ws: Worksheet
        deleted_col_indices: Liste der gelöschten Spalten-Indices (0-basiert)
        inserted_cols: Dict mit {position: count} für eingefügte Spalten
        new_headers: Liste der neuen Header (falls vorhanden, für Column-Update)
    """
    if not deleted_col_indices and not inserted_cols:
        return
    
    from openpyxl.worksheet.table import TableColumn
    from openpyxl.utils.cell import range_boundaries
    
    for table_name in ws.tables:
        table = ws.tables[table_name]
        old_ref = table.ref
        
        # Parse die alte Range
        min_col, min_row, max_col, max_row = range_boundaries(old_ref)
        
        # Berechne neue Spaltenanzahl
        old_col_count = max_col - min_col + 1
        deleted_count = len(deleted_col_indices) if deleted_col_indices else 0
        inserted_count = sum(inserted_cols.values()) if inserted_cols else 0
        new_col_count = old_col_count - deleted_count + inserted_count
        
        if new_col_count <= 0:
            continue
        
        # Table startet immer bei Spalte A (openpyxl verschiebt die Daten)
        # Nach delete_cols() ist die erste Spalte immer A1
        new_max_col = min_col + new_col_count - 1
        new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(new_max_col)}{max_row}"
        
        table.ref = new_ref
        
        # Auch den AutoFilter der Tabelle anpassen
        if table.autoFilter and table.autoFilter.ref:
            table.autoFilter.ref = new_ref
        
        # TABLE COLUMNS ANPASSEN
        # Die tableColumns müssen zur neuen Spaltenanzahl passen
        old_columns = list(table.tableColumns)
        
        # SCHRITT 1: Gelöschte Spalten aus tableColumns entfernen
        if deleted_col_indices:
            # Sortiere absteigend um Indexverschiebungen zu vermeiden
            for del_idx in sorted(deleted_col_indices, reverse=True):
                if del_idx < len(old_columns):
                    removed = old_columns.pop(del_idx)
        
        # SCHRITT 2: Neue Spalten einfügen
        if inserted_cols and new_headers:
            for pos, count in sorted(inserted_cols.items()):
                insert_idx = pos
                for i in range(count):
                    new_col_id = len(old_columns) + i + 1
                    new_col_name = new_headers[insert_idx + i] if insert_idx + i < len(new_headers) else f"Column{new_col_id}"
                    new_column = TableColumn(id=new_col_id, name=new_col_name)
                    old_columns.insert(insert_idx + i, new_column)
        
        # SCHRITT 3: Aktualisiere alle Column IDs (müssen 1, 2, 3, ... sein)
        # WICHTIG: Namen NICHT mit new_headers überschreiben bei delete!
        # Die Namen bleiben korrekt wenn wir nur die gelöschte Column entfernen.
        for idx, col in enumerate(old_columns):
            col.id = idx + 1
        
        # Setze die neuen Columns
        table.tableColumns = old_columns


def adjust_conditional_formatting(ws, deleted_col_indices, inserted_cols=None):
    """
    Passt alle bedingten Formatierungen an wenn Spalten gelöscht/eingefügt werden.
    
    WICHTIG: openpyxl's delete_cols() macht das NICHT automatisch!
    
    Args:
        ws: Worksheet
        deleted_col_indices: Liste der gelöschten Spalten-Indices (0-basiert)
        inserted_cols: Dict mit {position: count} für eingefügte Spalten
    """
    if not deleted_col_indices and not inserted_cols:
        return
    
    
    # Sammle alle CF-Regeln
    old_rules = list(ws.conditional_formatting._cf_rules.items())
    
    # Lösche alle CF-Regeln
    ws.conditional_formatting = ConditionalFormattingList()
    
    # Füge angepasste Regeln wieder hinzu
    for cf_obj, rules in old_rules:
        old_sqref = str(cf_obj.sqref)
        new_sqref = shift_range_reference(old_sqref, deleted_col_indices, inserted_cols)
        
        
        if new_sqref:
            # Füge Regel mit neuem Bereich hinzu
            for rule in rules:
                ws.conditional_formatting.add(new_sqref, rule)


def adjust_cf_for_row_changes(ws, row_mapping, original_row_count):
    """
    Passt alle bedingten Formatierungen an wenn Zeilen gelöscht/verschoben werden.
    
    Args:
        ws: Worksheet
        row_mapping: Liste wo row_mapping[new_pos] = original_data_row_idx (0-basiert)
        original_row_count: Ursprüngliche Anzahl der Datenzeilen
    """
    import re
    import sys
    
    if not row_mapping:
        return
    
    new_row_count = len(row_mapping)
    
    # Wenn keine Änderung in der Anzahl, nichts zu tun
    rows_deleted = original_row_count - new_row_count
    if rows_deleted <= 0:
        return
    
    sys.stderr.write(f"[CF ROW ADJUST] {rows_deleted} Zeilen gelöscht, passe CF an...\n")
    
    # Sammle alle CF-Regeln
    old_rules = list(ws.conditional_formatting._cf_rules.items())
    
    # Lösche alle CF-Regeln
    ws.conditional_formatting = ConditionalFormattingList()
    
    def adjust_cell_ref(cell_ref, deleted_count, new_max_row):
        """Passt eine Zellreferenz an (z.B. H2404 -> H2403)"""
        match = re.match(r'^(\$?)([A-Z]+)(\$?)(\d+)$', cell_ref.upper())
        if not match:
            return cell_ref
        
        col_abs = match.group(1)
        col_letter = match.group(2)
        row_abs = match.group(3)
        row_num = int(match.group(4))
        
        # Header-Zeile (1) nicht anpassen
        if row_num == 1:
            return cell_ref
        
        # Datenzeilen: Zeile 2 = Datenzeile 0
        # Nach Löschen: Neue max Zeile = new_max_row + 1 (Header)
        new_row = row_num - deleted_count
        
        # Nicht unter Zeile 2 gehen
        if new_row < 2:
            new_row = 2
        
        # Nicht über die neue maximale Zeile hinaus
        max_excel_row = new_max_row + 1  # +1 für Header
        if new_row > max_excel_row:
            new_row = max_excel_row
        
        return f"{col_abs}{col_letter}{row_abs}{new_row}"
    
    def adjust_range(range_str, deleted_count, new_max_row):
        """Passt einen Bereich an (z.B. H2:H2404 -> H2:H2403)"""
        # Kann mehrere Bereiche enthalten, getrennt durch Leerzeichen
        parts = range_str.split(' ')
        adjusted_parts = []
        
        for part in parts:
            if ':' in part:
                # Bereich wie H2:H2404
                start, end = part.split(':')
                new_start = adjust_cell_ref(start, deleted_count, new_max_row)
                new_end = adjust_cell_ref(end, deleted_count, new_max_row)
                adjusted_parts.append(f"{new_start}:{new_end}")
            else:
                # Einzelne Zelle wie I458
                adjusted_parts.append(adjust_cell_ref(part, deleted_count, new_max_row))
        
        return ' '.join(adjusted_parts)
    
    adjusted_count = 0
    # Füge angepasste Regeln wieder hinzu
    for cf_obj, rules in old_rules:
        old_sqref = str(cf_obj.sqref)
        new_sqref = adjust_range(old_sqref, rows_deleted, new_row_count)
        
        if new_sqref != old_sqref:
            adjusted_count += 1
        
        if new_sqref:
            for rule in rules:
                ws.conditional_formatting.add(new_sqref, rule)
    
    sys.stderr.write(f"[CF ROW ADJUST] {adjusted_count} CF-Bereiche angepasst\n")


def transform_cf_range(range_ref, column_mapping, deleted_set, target_col_count):
    """
    Transformiert CF-Bereiche basierend auf dem Spalten-Mapping.
    
    Args:
        range_ref: Original-Bereich wie 'A1:C10' oder 'A1:B2 C3:D4'
        column_mapping: Dict {new_col_idx: original_col_idx} (-1 für neue Spalten)
        deleted_set: Set der gelöschten Original-Spalten
        target_col_count: Anzahl der Zielspalten
    
    Returns:
        Transformierter Bereich oder None
    """
    if not range_ref:
        return None
    
    # Baue reverse mapping: original_col -> new_col
    reverse_mapping = {}
    for new_col, orig_col in column_mapping.items():
        if orig_col >= 0:  # Nicht neue Spalten
            reverse_mapping[orig_col] = new_col
    
    def transform_cell_ref(cell_ref):
        """Transformiert eine einzelne Zellreferenz"""
        match = re.match(r'^([A-Z]+)(\d+)$', cell_ref.upper())
        if not match:
            return None
        
        col_letter = match.group(1)
        row_num = match.group(2)
        orig_col_idx = column_index_from_string(col_letter) - 1  # 0-basiert
        
        # Spalte gelöscht?
        if orig_col_idx in deleted_set:
            return None
        
        # Finde neue Position
        if orig_col_idx in reverse_mapping:
            new_col_idx = reverse_mapping[orig_col_idx]
            new_col_letter = get_column_letter(new_col_idx + 1)
            return f"{new_col_letter}{row_num}"
        else:
            # Spalte nicht im Mapping - behalte Original (falls im Zielbereich)
            if orig_col_idx < target_col_count:
                return cell_ref
            return None
    
    # Handle mehrere Bereiche
    parts = range_ref.split()
    new_parts = []
    
    for part in parts:
        if ':' in part:
            start, end = part.split(':')
            new_start = transform_cell_ref(start)
            new_end = transform_cell_ref(end)
            
            if new_start and new_end:
                new_parts.append(f"{new_start}:{new_end}")
        else:
            new_ref = transform_cell_ref(part)
            if new_ref:
                new_parts.append(new_ref)
    
    return ' '.join(new_parts) if new_parts else None


def apply_cell_value(cell, value):
    """
    Setzt den Wert einer Zelle mit korrektem Typ.
    OPTIMIERT für Performance bei großen Datenmengen.
    Überspringt MergedCell-Objekte (nur die obere linke Zelle ist beschreibbar).
    Überspringt Bild-Platzhalter ('🖼️ Bild') damit Original-Zellwerte erhalten bleiben.
    """
    from datetime import date
    from openpyxl.cell.cell import MergedCell
    import re
    
    # MergedCell überspringen - nur die obere linke Zelle einer Merged-Region ist beschreibbar
    if isinstance(cell, MergedCell):
        return
    
    # Bild-Platzhalter überspringen — der Reader setzt '🖼️ Bild' für Zellen mit
    # richData/vm-Bildern. Wenn wir diesen Wert schreiben, ändert sich der Zell-Typ
    # von t="e" (error) zu t="inlineStr", was die vm-Attribut-Wiederherstellung und
    # damit die Bildanzeige in Excel zerstört. Original-Zellwert beibehalten!
    if isinstance(value, str) and '🖼️' in value:
        return
    
    # Schnelle Typchecks zuerst
    if value is None or value == '':
        cell.value = None
        return
    
    value_type = type(value)
    
    if value_type is bool:
        cell.value = value
    elif value_type in (int, float):
        cell.value = value
    elif value_type is datetime:
        cell.value = value
    elif value_type is date:
        cell.value = datetime.combine(value, datetime.min.time())
    elif value_type is str:
        # Versuche Datum-Strings zurück zu datetime zu konvertieren
        # Format vom Reader: '30.06.2013 00:00:00' oder '30.06.2013'
        parsed_date = None
        if len(value) >= 10:
            # Versuche verschiedene Datumsformate
            for fmt in ['%d.%m.%Y %H:%M:%S', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']:
                try:
                    parsed_date = datetime.strptime(value, fmt)
                    break
                except ValueError:
                    continue
        
        if parsed_date:
            cell.value = parsed_date
        else:
            cell.value = value
    else:
        cell.value = str(value)


def _direct_xml_cell_edit(file_path, output_path, sheet_name, real_edits,
                          hidden_columns=None, hidden_rows=None):
    """
    Direkte XML-Bearbeitung von Zellwerten OHNE openpyxl-Roundtrip.
    
    Kopiert die Original-XLSX 1:1 und ändert nur die betroffenen Zellwerte
    direkt im Worksheet-XML. Dadurch bleiben ALLE Strukturen intakt:
    - Relationships (rIds) 
    - SharedStrings-Indizes
    - Namespaces (mc:Ignorable, xr, x14ac, etc.)
    - Drawings, Media, RichData
    - Tables, Slicers, External Links
    - Conditional Formatting, Styles, Fonts
    
    Verwendet ZIP-to-ZIP Entry-Kopie: Liest Einträge direkt aus dem Original-ZIP
    und schreibt sie in ein neues ZIP. Nur modifizierte XMLs werden ersetzt.
    Keine Dateisystem-Extraktion nötig → keine .DS_Store, ._files, Permission-Probleme.
    
    Kein fix_xlsx_relationships, kein restore_table_xml, kein restore_external_links nötig.
    
    Args:
        file_path: Quelldatei
        output_path: Zieldatei
        sheet_name: Name des Sheets
        real_edits: Dict mit "rowIdx-colIdx" → value (0-basiert, ohne Header)
        hidden_columns: Liste versteckter Spalten (0-basiert) oder None
        hidden_rows: Liste versteckter Zeilen (0-basiert) oder None
    """
    import zipfile
    import tempfile
    import shutil
    from xml.etree import ElementTree as ET
    
    sys.stderr.write(f"[DIRECT_XML] Start: {len(real_edits)} Edits für Sheet '{sheet_name}'\n")
    
    # Namespaces die in Excel-Dateien vorkommen
    NS = {
        '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
        'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
        'xr3': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
        'xr6': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision6',
        'xr10': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision10',
    }
    MAIN_NS = NS['']
    
    # Registriere alle Namespaces um sie beim Schreiben nicht zu verlieren
    for prefix, uri in NS.items():
        if prefix:
            ET.register_namespace(prefix, uri)
    ET.register_namespace('', MAIN_NS)
    
    # =========================================================================
    # ZIP-to-ZIP Ansatz: Lese direkt aus dem ZIP, modifiziere im Speicher,
    # schreibe in neues ZIP. Keine Dateisystem-Extraktion!
    # =========================================================================
    
    # Temporäre Ausgabedatei (wird am Ende umbenannt)
    temp_output = output_path + '.tmp'
    
    try:
        with zipfile.ZipFile(file_path, 'r') as src_zip:
            # 1. Finde Sheet-XML Pfad aus workbook.xml + workbook.xml.rels
            wb_xml = src_zip.read('xl/workbook.xml').decode('utf-8')
            wb_root = ET.fromstring(wb_xml)
            
            sheet_rid = None
            for sheet_el in wb_root.iter(f'{{{MAIN_NS}}}sheet'):
                if sheet_el.get('name') == sheet_name:
                    sheet_rid = sheet_el.get(f'{{{NS["r"]}}}id')
                    break
            
            if not sheet_rid:
                raise ValueError(f"Sheet '{sheet_name}' nicht in workbook.xml gefunden")
            
            sys.stderr.write(f"[DIRECT_XML] Sheet '{sheet_name}' hat rId={sheet_rid}\n")
            
            rels_xml = src_zip.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            rels_root = ET.fromstring(rels_xml)
            rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
            
            sheet_file = None
            for rel_el in rels_root.iter(f'{{{rels_ns}}}Relationship'):
                if rel_el.get('Id') == sheet_rid:
                    sheet_file = rel_el.get('Target')
                    break
            
            if not sheet_file:
                raise ValueError(f"Relationship {sheet_rid} nicht in workbook.xml.rels gefunden")
            
            # Sheet-Pfad normalisieren (Target ist relativ zu xl/)
            sheet_zip_path = 'xl/' + sheet_file.lstrip('/')
            # Normalisiere ../ etc.
            parts = sheet_zip_path.split('/')
            normalized = []
            for p in parts:
                if p == '..':
                    if normalized:
                        normalized.pop()
                elif p != '.':
                    normalized.append(p)
            sheet_zip_path = '/'.join(normalized)
            
            sys.stderr.write(f"[DIRECT_XML] Sheet-ZIP-Pfad: {sheet_zip_path}\n")
            
            # 2. Lese SharedStrings (nur zum Referenzieren)
            shared_strings = []
            has_shared_strings = 'xl/sharedStrings.xml' in src_zip.namelist()
            ss_content = None
            if has_shared_strings:
                ss_content = src_zip.read('xl/sharedStrings.xml').decode('utf-8')
                ss_root = ET.fromstring(ss_content)
                for si in ss_root.iter(f'{{{MAIN_NS}}}si'):
                    texts = []
                    for t in si.iter(f'{{{MAIN_NS}}}t'):
                        if t.text:
                            texts.append(t.text)
                    shared_strings.append(''.join(texts))
            
            # 3. Lese Sheet-XML
            sheet_content = src_zip.read(sheet_zip_path).decode('utf-8')
            
            # 4. Konvertiere Edits zu Excel-Koordinaten
            edits_by_ref = {}
            for key, value in real_edits.items():
                parts = key.split('-')
                if len(parts) != 2:
                    continue
                row_idx = int(parts[0])
                col_idx = int(parts[1])
                excel_row = row_idx + 2
                excel_col = col_idx + 1
                col_letter = get_column_letter(excel_col)
                cell_ref = f"{col_letter}{excel_row}"
                edits_by_ref[cell_ref] = value
            
            sys.stderr.write(f"[DIRECT_XML] Zell-Referenzen: {list(edits_by_ref.keys())}\n")
            
            # 5. Für jede editierte Zelle: Wert im XML ändern
            modified = False
            for cell_ref, value in edits_by_ref.items():
                if isinstance(value, str) and '🖼️' in value:
                    continue
                
                sheet_content, was_modified = _replace_cell_value_in_xml(
                    sheet_content, cell_ref, value, MAIN_NS, shared_strings, has_shared_strings
                )
                if was_modified:
                    modified = True
            
            # 6. Hidden Columns/Rows direkt im XML setzen
            if hidden_columns is not None:
                sheet_content = _set_hidden_cols_in_xml(sheet_content, hidden_columns, MAIN_NS)
                modified = True
            if hidden_rows is not None:
                sheet_content = _set_hidden_rows_in_xml(sheet_content, hidden_rows, MAIN_NS)
                modified = True
            
            # 7. SharedStrings aktualisieren falls neue Strings hinzugefügt wurden
            ss_modified = False
            if has_shared_strings and hasattr(_replace_cell_value_in_xml, '_new_strings'):
                new_strings = _replace_cell_value_in_xml._new_strings
                if new_strings and ss_content:
                    new_count = len(shared_strings) + len(new_strings)
                    ss_content = re.sub(r'count="\d+"', f'count="{new_count}"', ss_content)
                    ss_content = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{new_count}"', ss_content)
                    
                    new_si_xml = ''
                    for s in new_strings:
                        escaped = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        new_si_xml += f'<si><t>{escaped}</t></si>'
                    ss_content = ss_content.replace('</sst>', new_si_xml + '</sst>')
                    ss_modified = True
                    sys.stderr.write(f"[DIRECT_XML] {len(new_strings)} neue SharedStrings hinzugefügt\n")
                
                _replace_cell_value_in_xml._new_strings = []
            
            if not modified:
                # Keine Änderungen → Original einfach kopieren
                if os.path.normpath(file_path) != os.path.normpath(output_path):
                    shutil.copy2(file_path, output_path)
                sys.stderr.write(f"[DIRECT_XML] Keine Änderungen nötig\n")
                return {'success': True, 'outputPath': output_path, 'method': 'direct-xml'}
            
            # 8. ZIP-to-ZIP: Kopiere alle Einträge, ersetze nur modifizierte
            with zipfile.ZipFile(temp_output, 'w') as dst_zip:
                for item in src_zip.infolist():
                    # Überspringe macOS-Artefakte
                    if item.filename.startswith('__MACOSX') or \
                       item.filename.endswith('.DS_Store') or \
                       '/.DS_Store' in item.filename or \
                       item.filename.split('/')[-1].startswith('._'):
                        continue
                    
                    if item.filename == sheet_zip_path:
                        # Modifiziertes Sheet-XML schreiben
                        item.compress_type = zipfile.ZIP_DEFLATED
                        dst_zip.writestr(item, sheet_content.encode('utf-8'))
                    elif item.filename == 'xl/sharedStrings.xml' and ss_modified:
                        # Modifizierte SharedStrings schreiben
                        item.compress_type = zipfile.ZIP_DEFLATED
                        dst_zip.writestr(item, ss_content.encode('utf-8'))
                    else:
                        # Original-Bytes 1:1 kopieren
                        data = src_zip.read(item.filename)
                        dst_zip.writestr(item, data)
        
        # 9. Temporäre Datei an Zielort verschieben
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(temp_output, output_path)
        sys.stderr.write(f"[DIRECT_XML] Erfolgreich gespeichert: {output_path}\n")
        
        return {'success': True, 'outputPath': output_path, 'method': 'direct-xml'}
    
    except Exception:
        # Aufräumen bei Fehler
        if os.path.exists(temp_output):
            os.remove(temp_output)
        raise


def _replace_cell_value_in_xml(sheet_content, cell_ref, value, main_ns, shared_strings, has_shared_strings):
    """
    Ersetzt den Wert einer einzelnen Zelle im rohen Worksheet-XML.
    
    Strategien:
    - Zelle existiert mit t="s" (SharedString): Neuen String zur SharedStrings-Tabelle
      hinzufügen und Index aktualisieren
    - Zelle existiert mit t="inlineStr": Text direkt ersetzen
    - Zelle existiert ohne Typ (Zahl): Wert direkt ersetzen
    - Zelle existiert nicht: Zelle in passender Zeile einfügen
    
    Returns (new_content, was_modified)
    """
    import re
    from datetime import datetime, date
    
    # Initialisiere _new_strings Tracker (für SharedStrings-Updates)
    if not hasattr(_replace_cell_value_in_xml, '_new_strings'):
        _replace_cell_value_in_xml._new_strings = []
    
    # Bestimme den neuen Wert und Typ
    if value is None or value == '':
        new_type = 'empty'
        new_val = ''
    elif isinstance(value, bool):
        new_type = 'bool'
        new_val = '1' if value else '0'
    elif isinstance(value, (int, float)):
        new_type = 'number'
        new_val = str(value)
        # Ganzzahlen ohne .0 
        if isinstance(value, float) and value == int(value):
            new_val = str(int(value))
    elif isinstance(value, str):
        # Prüfe ob es ein Datum ist
        parsed_date = None
        if len(value) >= 10:
            for fmt in ['%d.%m.%Y %H:%M:%S', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']:
                try:
                    parsed_date = datetime.strptime(value, fmt)
                    break
                except ValueError:
                    continue
        
        if parsed_date:
            # Excel-Datumsserial (Tage seit 1899-12-30)
            from datetime import timedelta
            excel_epoch = datetime(1899, 12, 30)
            delta = parsed_date - excel_epoch
            serial = delta.days + delta.seconds / 86400.0
            new_type = 'number'
            new_val = str(int(serial)) if delta.seconds == 0 else str(serial)
        else:
            new_type = 'string'
            new_val = value
    else:
        new_type = 'string'
        new_val = str(value)
    
    # Suche die Zelle im XML
    # Mögliche Formate:
    # <c r="A1" t="s"><v>5</v></c>
    # <c r="A1" t="inlineStr"><is><t>text</t></is></c>
    # <c r="A1" s="3"><v>42</v></c>
    # <c r="A1"/>
    # <c r="A1" s="3"/>
    
    # Pattern für die Zelle (selbstschließend oder mit Inhalt)
    cell_pattern = re.compile(
        r'(<c\s[^>]*?r="' + re.escape(cell_ref) + r'"[^>]*?)(/\s*>|>(.*?)</c>)',
        re.DOTALL
    )
    
    match = cell_pattern.search(sheet_content)
    
    if match:
        cell_open = match.group(1)   # <c r="A1" t="s" s="3"
        cell_close = match.group(2)  # /> oder >...</c>
        
        # Extrahiere das s="..." Attribut (Style-Index) — muss erhalten bleiben!
        style_attr = ''
        s_match = re.search(r'\bs="(\d+)"', cell_open)
        if s_match:
            style_attr = f' s="{s_match.group(1)}"'
        
        # Extrahiere vm="..." Attribut — muss erhalten bleiben (Zellbilder)
        vm_attr = ''
        vm_match = re.search(r'\bvm="(\d+)"', cell_open)
        if vm_match:
            vm_attr = f' vm="{vm_match.group(1)}"'
        
        # Baue neues Zell-Element
        if new_type == 'empty':
            new_cell = f'<c r="{cell_ref}"{style_attr}{vm_attr}/>'
        elif new_type == 'number':
            new_cell = f'<c r="{cell_ref}"{style_attr}{vm_attr}><v>{new_val}</v></c>'
        elif new_type == 'bool':
            new_cell = f'<c r="{cell_ref}"{style_attr}{vm_attr} t="b"><v>{new_val}</v></c>'
        elif new_type == 'string':
            if has_shared_strings:
                # Neuen SharedString hinzufügen und Index verwenden
                new_idx = len(shared_strings) + len(_replace_cell_value_in_xml._new_strings)
                _replace_cell_value_in_xml._new_strings.append(new_val)
                new_cell = f'<c r="{cell_ref}"{style_attr}{vm_attr} t="s"><v>{new_idx}</v></c>'
            else:
                # Inline String
                escaped = new_val.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                new_cell = f'<c r="{cell_ref}"{style_attr}{vm_attr} t="inlineStr"><is><t>{escaped}</t></is></c>'
        
        sheet_content = sheet_content[:match.start()] + new_cell + sheet_content[match.end():]
        return sheet_content, True
    
    else:
        # Zelle existiert nicht → in passender Zeile einfügen
        row_num = re.search(r'(\d+)$', cell_ref).group(1)
        
        # Finde die Zeile
        row_pattern = re.compile(
            r'(<row\s[^>]*?\br="' + re.escape(row_num) + r'"[^>]*?>)',
            re.DOTALL
        )
        row_match = row_pattern.search(sheet_content)
        
        if row_match:
            # Baue neues Zell-Element
            if new_type == 'empty':
                return sheet_content, False  # Leere Zelle die nicht existiert → nichts tun
            elif new_type == 'number':
                new_cell = f'<c r="{cell_ref}"><v>{new_val}</v></c>'
            elif new_type == 'bool':
                new_cell = f'<c r="{cell_ref}" t="b"><v>{new_val}</v></c>'
            elif new_type == 'string':
                if has_shared_strings:
                    new_idx = len(shared_strings) + len(_replace_cell_value_in_xml._new_strings)
                    _replace_cell_value_in_xml._new_strings.append(new_val)
                    new_cell = f'<c r="{cell_ref}" t="s"><v>{new_idx}</v></c>'
                else:
                    escaped = new_val.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    new_cell = f'<c r="{cell_ref}" t="inlineStr"><is><t>{escaped}</t></is></c>'
            
            insert_pos = row_match.end()
            sheet_content = sheet_content[:insert_pos] + new_cell + sheet_content[insert_pos:]
            return sheet_content, True
        else:
            # Zeile existiert nicht → vor </sheetData> erstellen
            if new_type == 'empty':
                return sheet_content, False
            
            if new_type == 'number':
                new_cell = f'<c r="{cell_ref}"><v>{new_val}</v></c>'
            elif new_type == 'bool':
                new_cell = f'<c r="{cell_ref}" t="b"><v>{new_val}</v></c>'
            elif new_type == 'string':
                if has_shared_strings:
                    new_idx = len(shared_strings) + len(_replace_cell_value_in_xml._new_strings)
                    _replace_cell_value_in_xml._new_strings.append(new_val)
                    new_cell = f'<c r="{cell_ref}" t="s"><v>{new_idx}</v></c>'
                else:
                    escaped = new_val.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    new_cell = f'<c r="{cell_ref}" t="inlineStr"><is><t>{escaped}</t></is></c>'
            
            new_row = f'<row r="{row_num}">{new_cell}</row>\n'
            sheet_content = sheet_content.replace('</sheetData>', new_row + '</sheetData>')
            return sheet_content, True


def _set_hidden_cols_in_xml(sheet_content, hidden_columns, main_ns):
    """Setzt hidden-Attribute auf <col> Elemente im Worksheet-XML."""
    import re
    
    if not hidden_columns:
        return sheet_content
    
    hidden_set = set(hidden_columns)
    
    # Finde alle <col> Elemente und setze/entferne hidden
    def _fix_col(m):
        col_el = m.group(0)
        # Extrahiere min/max
        min_m = re.search(r'min="(\d+)"', col_el)
        max_m = re.search(r'max="(\d+)"', col_el)
        if not min_m or not max_m:
            return col_el
        col_min = int(min_m.group(1))
        col_max = int(max_m.group(1))
        
        # Prüfe ob ALLE Spalten in diesem Range versteckt sein sollen
        # (0-basiert in hidden_columns, 1-basiert in XML)
        all_hidden = all((c - 1) in hidden_set for c in range(col_min, col_max + 1))
        
        if all_hidden:
            if 'hidden="1"' not in col_el and "hidden='1'" not in col_el:
                col_el = col_el.replace('/>', ' hidden="1"/>')
                if not col_el.endswith('/>'):
                    col_el = re.sub(r'>', ' hidden="1">', col_el, count=1)
        else:
            col_el = re.sub(r'\s*hidden="1"', '', col_el)
        
        return col_el
    
    sheet_content = re.sub(r'<col\s[^>]*/>', _fix_col, sheet_content)
    return sheet_content


def _set_hidden_rows_in_xml(sheet_content, hidden_rows, main_ns):
    """Setzt hidden-Attribute auf <row> Elemente im Worksheet-XML."""
    import re
    
    if not hidden_rows:
        return sheet_content
    
    hidden_set = set(hidden_rows)
    
    def _fix_row(m):
        row_tag = m.group(0)
        r_m = re.search(r'\br="(\d+)"', row_tag)
        if not r_m:
            return row_tag
        row_num = int(r_m.group(1))
        # 0-basiert in hidden_rows, Datenzeilen ab row 2 (row 1 = Header)
        row_idx = row_num - 2
        
        if row_idx in hidden_set:
            if 'hidden="1"' not in row_tag:
                row_tag = row_tag.rstrip('>') + ' hidden="1">'
        else:
            row_tag = re.sub(r'\s*hidden="1"', '', row_tag)
        
        return row_tag
    
    # Nur <row ...> Opening-Tags matchen (nicht den Inhalt)
    sheet_content = re.sub(r'<row\s[^>]*?>', _fix_row, sheet_content)
    return sheet_content


def write_sheet(file_path, output_path, sheet_name, changes, original_path=None):
    """
    Schreibt Änderungen in ein Excel-Sheet
    
    WICHTIG: Bei strukturellen Änderungen (fullRewrite=True) werden die 
    NEUEN Daten geschrieben. Die Original-Struktur wird beibehalten wo möglich.
    
    Args:
        file_path: Pfad zur Arbeitsdatei (kopierte Datei)
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        changes: Dict mit allen Änderungen
        original_path: Pfad zur Original-Datei (für restore_table_xml)
    
    Returns:
        Dict mit success und ggf. error
    """
    # Wenn kein original_path gegeben, verwende file_path (Legacy-Kompatibilität)
    if original_path is None:
        original_path = file_path
    
    # KRITISCH: Wenn original_path == file_path == output_path (Speichern in gleicher Datei),
    # muss eine Backup-Kopie erstellt werden BEVOR openpyxl die Datei überschreibt.
    # Sonst kann restore_external_links_from_original nichts wiederherstellen.
    # WINDOWS: normpath normalisiert Pfade (Slashes, Groß/Klein, trailing sep)
    _backup_file = None
    if os.path.normpath(original_path) == os.path.normpath(output_path):
        import tempfile
        _backup_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        _backup_path = _backup_file.name
        _backup_file.close()
        import shutil
        shutil.copy2(original_path, _backup_path)
        sys.stderr.write(f"[WRITE_SHEET] Backup erstellt: {_backup_path} (original==output)\\n")
        # Speichere Backup-Pfad als Funktionsattribut für restore_external_links_from_original
        restore_external_links_from_original._backup_original_path = _backup_path
    
    try:
        # Original-Workbook laden
        # Workaround für openpyxl Bug mit extLst in PatternFill
        # rich_text=True damit CellRichText-Objekte erhalten bleiben
        try:
            wb = load_workbook(file_path, rich_text=True)
        except TypeError as e:
            if 'extLst' in str(e):
                # openpyxl kann diese Datei nicht verarbeiten - Fallback-Fehler
                return {
                    'success': False, 
                    'error': f'Diese Datei enthält erweiterte Formatierungen die openpyxl nicht unterstützt. Bitte Excel/xlwings verwenden.',
                    'requiresXlwings': True
                }
            raise
        
        if sheet_name not in wb.sheetnames:
            return {'success': False, 'error': f'Sheet "{sheet_name}" nicht gefunden'}
        
        ws = wb[sheet_name]
        
        # Parameter extrahieren
        headers = changes.get('headers', [])
        data = changes.get('data', [])
        edited_cells = changes.get('editedCells', {})
        cell_styles = changes.get('cellStyles', {})
        row_highlights = changes.get('rowHighlights', {})
        deleted_columns = changes.get('deletedColumns', [])
        inserted_columns = changes.get('insertedColumns')
        column_order = changes.get('columnOrder')  # [neuIdx] = altIdx
        hidden_columns = changes.get('hiddenColumns', [])
        hidden_rows = changes.get('hiddenRows', [])
        row_mapping = changes.get('rowMapping')
        from_file = changes.get('fromFile', False)
        full_rewrite = changes.get('fullRewrite', False)
        structural_change = changes.get('structuralChange', False)
        frontend_auto_filter = changes.get('autoFilterRange')  # AutoFilter vom Frontend
        
        cleared_row_highlights = changes.get('clearedRowHighlights', [])
        affected_rows = changes.get('affectedRows', [])
        
        # Zeilen-Operationen (analog zu Spalten-Operationen)
        deleted_rows = changes.get('deletedRowIndices', [])
        inserted_rows = changes.get('insertedRowInfo')
        row_order = changes.get('rowOrder')  # [neuIdx] = altIdx
        
        # DEBUG: Zeige alle relevanten Flags
        import sys
        sys.stderr.write(f"[WRITE_SHEET] row_highlights={row_highlights}, cleared_row_highlights={cleared_row_highlights}\n")
        sys.stderr.write(f"[WRITE_SHEET] row_mapping={bool(row_mapping)}, structural_change={structural_change}, full_rewrite={full_rewrite}\n")
        sys.stderr.write(f"[WRITE_SHEET] deleted_rows={deleted_rows}, inserted_rows={bool(inserted_rows)}, row_order={bool(row_order)}\n")
        sys.stderr.write(f"[WRITE_SHEET] deleted_columns={deleted_columns}, inserted_columns={bool(inserted_columns)}, column_order={bool(column_order)}\n")
        
        # DEBUG: mergedCells vom Frontend
        imported_mc_debug = changes.get('mergedCells', [])
        sys.stderr.write(f"[WRITE_SHEET] mergedCells count={len(imported_mc_debug)}\n")
        if imported_mc_debug:
            for i, mc in enumerate(imported_mc_debug[:5]):
                sys.stderr.write(f"[WRITE_SHEET] mergedCells[{i}]: startRow={mc.get('startRow')}, startCol={mc.get('startCol')}, endRow={mc.get('endRow')}, endCol={mc.get('endCol')}\n")
            if len(imported_mc_debug) > 5:
                sys.stderr.write(f"[WRITE_SHEET] ... und {len(imported_mc_debug) - 5} weitere\n")
        
        # =====================================================================
        # FALL 1: fromFile - Nur versteckte Spalten/Zeilen setzen
        # =====================================================================
        if from_file:
            _apply_hidden_columns(ws, hidden_columns)
            _apply_hidden_rows(ws, hidden_rows)
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            # WICHTIG: Auch bei fromFile richData/Bilder/Namespaces wiederherstellen!
            # openpyxl verliert beim Speichern richData, metadata, vm-Attribute etc.
            restore_table_xml_from_original(output_path, original_path, table_changes=None)
            restore_external_links_from_original(output_path, original_path)
            return {'success': True, 'outputPath': output_path}
        
        # =====================================================================
        # FALL 1.X: UNIVERSELLE PIPELINE für Spalten- UND Zeilen-Operationen
        # Führt alle Operationen STRIKT SEQUENTIELL aus:
        # 1-4. Zeilen-Operationen (alle Daten zuerst speichern, dann rekonstruieren)
        #      1. Alle Original-Zeilen speichern
        #      2. Finale Zeilen-Reihenfolge berechnen (Löschen + Verschieben)
        #      3. Überschüssige Zeilen entfernen
        #      4. Zeilen in neuer Reihenfolge schreiben
        # 5. Zeilen einfügen
        # 6. Zeilen verstecken (NACH allen strukturellen Änderungen)
        # 7. Spalten löschen (von hinten nach vorne)
        # 8. Spalten einfügen (von vorne nach hinten)
        # 9. Spalten verschieben/reorder
        # 10. Spalten verstecken
        # 11. Row Highlights
        # 12. Tables reparieren
        # 13. Einmal speichern
        # 14. XML restore
        # =====================================================================
        
        # Prüfe ob rowMapping nur die Identität ist (keine echte Änderung)
        row_mapping_is_identity = True
        if row_mapping:
            for i, val in enumerate(row_mapping):
                if val != i:
                    row_mapping_is_identity = False
                    break
        
        # Prüfe ob wir Zeilen-Operationen haben
        has_row_operations = deleted_rows or inserted_rows or (row_order and len(row_order) > 0)
        
        # Prüfe ob wir den Pipeline-Pfad nutzen können
        # (Spalten- ODER Zeilen-Operationen)
        has_column_operations = deleted_columns or inserted_columns or (column_order and len(column_order) > 0)
        can_use_pipeline = (has_column_operations or has_row_operations) and row_mapping_is_identity and not affected_rows
        
        if can_use_pipeline:
            from openpyxl.worksheet.table import TableColumn
            from openpyxl.utils.cell import range_boundaries
            from openpyxl.cell.cell import MergedCell
            import sys
            
            sys.stderr.write(f"[PIPELINE] Starte: deleted_rows={deleted_rows}, row_order={row_order is not None}, hidden_rows={hidden_rows}, deleted_columns={deleted_columns}, inserted_columns={inserted_columns is not None}, column_order={column_order is not None}\n")
            
            # =====================================================================
            # ZEILEN-OPERATIONEN: Alle Daten ZUERST speichern, dann rekonstruieren
            # =====================================================================
            
            has_any_row_change = deleted_rows or (row_order and len(row_order) > 0)
            
            if has_any_row_change:
                max_col = ws.max_column
                original_max_row = ws.max_row
                
                # SCHRITT 1: Alle Original-Zeilen komplett speichern (vor jeder Änderung!)
                sys.stderr.write(f"[PIPELINE] Schritt 1: Speichere alle {original_max_row - 1} Original-Zeilen\n")
                all_rows_backup = {}
                for excel_row in range(2, original_max_row + 1):  # Ab Zeile 2 (nach Header)
                    row_idx = excel_row - 2  # 0-basierter Index
                    all_rows_backup[row_idx] = {}
                    
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=excel_row, column=col)
                        if isinstance(cell, MergedCell):
                            continue
                        all_rows_backup[row_idx][col] = {
                            'value': cell.value,
                            'fill': copy(cell.fill) if cell.fill else None,
                            'font': copy(cell.font) if cell.font else None,
                            'alignment': copy(cell.alignment) if cell.alignment else None,
                            'border': copy(cell.border) if cell.border else None,
                            'number_format': cell.number_format,
                            'hyperlink': cell.hyperlink.target if cell.hyperlink else None
                        }
                
                # SCHRITT 2: Bestimme finale Zeilen-Reihenfolge
                # row_order enthält: [neuIdx] = altIdx (nach Löschen!)
                # deleted_rows enthält: Original-Indizes der gelöschten Zeilen
                
                deleted_set = set(deleted_rows) if deleted_rows else set()
                
                if row_order and len(row_order) > 0:
                    # row_order gibt die neue Reihenfolge vor
                    # Die Indizes in row_order beziehen sich auf Zeilen NACH dem Löschen
                    # Wir müssen sie zurück auf Original-Indizes mappen
                    
                    # Erstelle Mapping: Index nach Löschen → Original-Index
                    remaining_original_indices = []
                    for orig_idx in range(len(all_rows_backup)):
                        if orig_idx not in deleted_set:
                            remaining_original_indices.append(orig_idx)
                    
                    # row_order[new_pos] = after_delete_idx → wir brauchen original_idx
                    final_row_order = []
                    for new_pos, after_delete_idx in enumerate(row_order):
                        if after_delete_idx < len(remaining_original_indices):
                            original_idx = remaining_original_indices[after_delete_idx]
                            final_row_order.append(original_idx)
                    
                    sys.stderr.write(f"[PIPELINE] Schritt 2: Finale Zeilen-Reihenfolge (Original-Indizes): {final_row_order[:10]}...\n")
                else:
                    # Keine Verschiebung, nur Löschen - behalte Reihenfolge der nicht-gelöschten
                    final_row_order = [idx for idx in range(len(all_rows_backup)) if idx not in deleted_set]
                    sys.stderr.write(f"[PIPELINE] Schritt 2: Nur Löschen, behalte {len(final_row_order)} Zeilen\n")
                
                # SCHRITT 3: Überschüssige Zeilen löschen (von hinten)
                target_row_count = len(final_row_order)
                current_data_rows = original_max_row - 1  # Ohne Header
                
                if current_data_rows > target_row_count:
                    rows_to_delete = current_data_rows - target_row_count
                    sys.stderr.write(f"[PIPELINE] Schritt 3: Lösche {rows_to_delete} überschüssige Zeilen\n")
                    for _ in range(rows_to_delete):
                        ws.delete_rows(ws.max_row, 1)
                
                # SCHRITT 4: Zeilen in neuer Reihenfolge schreiben
                sys.stderr.write(f"[PIPELINE] Schritt 4: Schreibe {len(final_row_order)} Zeilen in neuer Reihenfolge\n")
                for new_idx, original_idx in enumerate(final_row_order):
                    new_excel_row = new_idx + 2
                    
                    if original_idx not in all_rows_backup:
                        continue
                    
                    for col, data_item in all_rows_backup[original_idx].items():
                        cell = ws.cell(row=new_excel_row, column=col)
                        if isinstance(cell, MergedCell):
                            continue
                        cell.value = data_item['value']
                        if data_item['fill']:
                            cell.fill = data_item['fill']
                        if data_item['font']:
                            cell.font = data_item['font']
                        if data_item['alignment']:
                            cell.alignment = data_item['alignment']
                        if data_item['border']:
                            cell.border = data_item['border']
                        if data_item['number_format']:
                            cell.number_format = data_item['number_format']
                        if data_item['hyperlink']:
                            cell.hyperlink = data_item['hyperlink']
            
            # ===== SCHRITT 5: Zeilen EINFÜGEN =====
            if inserted_rows:
                operations = inserted_rows.get('operations', [])
                operations.sort(key=lambda x: x['position'])
                sys.stderr.write(f"[PIPELINE] Schritt 5: Füge Zeilen ein {[op['position'] for op in operations]}\n")
                
                for op in operations:
                    position = op['position']
                    count = op.get('count', 1)
                    excel_row = position + 2
                    
                    for i in range(count):
                        ws.insert_rows(excel_row + i, 1)
                        
                        # Formatierung von Zeile darüber kopieren
                        if excel_row + i > 2:
                            source_row = excel_row + i - 1
                            for col in range(1, ws.max_column + 1):
                                source_cell = ws.cell(row=source_row, column=col)
                                target_cell = ws.cell(row=excel_row + i, column=col)
                                if source_cell.fill:
                                    target_cell.fill = copy(source_cell.fill)
                                if source_cell.font:
                                    target_cell.font = copy(source_cell.font)
                                if source_cell.alignment:
                                    target_cell.alignment = copy(source_cell.alignment)
                                if source_cell.border:
                                    target_cell.border = copy(source_cell.border)
                                if source_cell.number_format:
                                    target_cell.number_format = source_cell.number_format
            
            # ===== SCHRITT 6: Zeilen VERSTECKEN (NACH allen strukturellen Änderungen) =====
            sys.stderr.write(f"[PIPELINE] Schritt 6: Zeilen verstecken, hidden_rows={hidden_rows}\n")
            _apply_hidden_rows(ws, hidden_rows)
            
            # ===== SCHRITT 7: Spalten LÖSCHEN (von hinten nach vorne) =====
            if deleted_columns:
                sorted_deleted = sorted(deleted_columns, reverse=True)
                sys.stderr.write(f"[PIPELINE] Schritt 7: Lösche Spalten {sorted_deleted}\n")
                
                for col_idx in sorted_deleted:
                    excel_col = col_idx + 1
                    max_col = ws.max_column
                    
                    # Spaltenbreiten speichern
                    saved_widths = {}
                    for col in range(excel_col + 1, max_col + 1):
                        col_letter = get_column_letter(col)
                        if col_letter in ws.column_dimensions:
                            saved_widths[col] = ws.column_dimensions[col_letter].width
                    
                    # Spalte löschen
                    ws.delete_cols(excel_col, 1)
                    
                    # Spaltenbreiten wiederherstellen
                    for old_col, width in saved_widths.items():
                        if width:
                            new_letter = get_column_letter(old_col - 1)
                            ws.column_dimensions[new_letter].width = width
                    
                    # CF anpassen
                    adjust_conditional_formatting(ws, [col_idx], None)
            
            # ===== SCHRITT 8: Spalten EINFÜGEN (von vorne nach hinten) =====
            if inserted_columns:
                operations = inserted_columns.get('operations', [])
                if not operations and inserted_columns.get('position') is not None:
                    operations = [{
                        'position': inserted_columns['position'],
                        'count': inserted_columns.get('count', 1),
                        'sourceColumn': inserted_columns.get('sourceColumn')
                    }]
                
                operations.sort(key=lambda x: x['position'])
                sys.stderr.write(f"[PIPELINE] Schritt 8: Füge Spalten ein\n")
                
                for op_idx, op in enumerate(operations):
                    position = op['position']
                    count = op.get('count', 1)
                    source_column = op.get('sourceColumn')
                    excel_col = position + 1
                    
                    for i in range(count):
                        insert_at = excel_col + i
                        
                        # Formatierung der Referenzspalte speichern
                        source_format = {}
                        source_width = None
                        if source_column is not None:
                            source_excel_col = source_column + 1
                            for prev_op in operations[:op_idx]:
                                if source_column >= prev_op['position']:
                                    source_excel_col += prev_op.get('count', 1)
                            
                            col_letter = get_column_letter(source_excel_col)
                            if col_letter in ws.column_dimensions:
                                source_width = ws.column_dimensions[col_letter].width
                            
                            for row in range(1, ws.max_row + 1):
                                cell = ws.cell(row=row, column=source_excel_col)
                                source_format[row] = {
                                    'fill': copy(cell.fill) if cell.fill else None,
                                    'font': copy(cell.font) if cell.font else None,
                                    'alignment': copy(cell.alignment) if cell.alignment else None,
                                    'border': copy(cell.border) if cell.border else None,
                                    'number_format': cell.number_format
                                }
                        
                        # Spaltenbreiten speichern
                        saved_widths = {}
                        for col in range(insert_at, ws.max_column + 1):
                            col_letter = get_column_letter(col)
                            if col_letter in ws.column_dimensions:
                                saved_widths[col] = ws.column_dimensions[col_letter].width
                        
                        # Spalte einfügen
                        ws.insert_cols(insert_at, 1)
                        
                        # Spaltenbreiten wiederherstellen
                        for old_col, width in saved_widths.items():
                            if width:
                                new_letter = get_column_letter(old_col + 1)
                                ws.column_dimensions[new_letter].width = width
                        
                        # CF anpassen
                        inserted_cols_for_cf = {insert_at - 1: 1}
                        adjust_conditional_formatting(ws, [], inserted_cols_for_cf)
                        
                        # Formatierung anwenden
                        if source_width:
                            ws.column_dimensions[get_column_letter(insert_at)].width = source_width
                        
                        for row, fmt in source_format.items():
                            cell = ws.cell(row=row, column=insert_at)
                            if fmt['fill']:
                                cell.fill = fmt['fill']
                            if fmt['font']:
                                cell.font = fmt['font']
                            if fmt['alignment']:
                                cell.alignment = fmt['alignment']
                            if fmt['border']:
                                cell.border = fmt['border']
                            if fmt.get('number_format'):
                                cell.number_format = fmt['number_format']
                    
                    # Header setzen
                    op_headers = op.get('headers', [])
                    for i, header in enumerate(op_headers):
                        ws.cell(row=1, column=excel_col + i).value = header
                    
                    # Daten schreiben
                    if data and headers:
                        for i in range(count):
                            col_idx = position + i
                            if col_idx < len(headers):
                                for row_idx, row_data in enumerate(data):
                                    if col_idx < len(row_data):
                                        cell = ws.cell(row=row_idx + 2, column=excel_col + i)
                                        apply_cell_value(cell, row_data[col_idx])
            
            # ===== SCHRITT 9: Spalten VERSCHIEBEN/REORDER =====
            sys.stderr.write(f"[PIPELINE] Schritt 9: Spalten verschieben\n")
            if column_order and len(column_order) > 0:
                columns_changed = False
                for new_idx, old_idx in enumerate(column_order):
                    if new_idx != old_idx:
                        columns_changed = True
                        break
                
                if columns_changed:
                    num_cols = len(column_order)
                    max_row = ws.max_row
                    
                    # Alle Spalten in temp_columns speichern
                    temp_columns = {}
                    for old_col_idx in range(num_cols):
                        old_excel_col = old_col_idx + 1
                        temp_columns[old_col_idx] = {}
                        
                        for row in range(1, max_row + 1):
                            cell = ws.cell(row=row, column=old_excel_col)
                            if isinstance(cell, MergedCell):
                                continue
                            temp_columns[old_col_idx][row] = {
                                'value': cell.value,
                                'hyperlink': cell.hyperlink.target if cell.hyperlink else None,
                            }
                    
                    # Spalten in neuer Reihenfolge schreiben
                    for new_col_idx, old_col_idx in enumerate(column_order):
                        new_excel_col = new_col_idx + 1
                        
                        if old_col_idx not in temp_columns:
                            continue
                        
                        for row, data_item in temp_columns[old_col_idx].items():
                            cell = ws.cell(row=row, column=new_excel_col)
                            if isinstance(cell, MergedCell):
                                continue
                            cell.value = data_item['value']
                            if data_item['hyperlink']:
                                cell.hyperlink = data_item['hyperlink']
            
            # ===== SCHRITT 10: Versteckte Spalten =====
            sys.stderr.write(f"[PIPELINE] Schritt 10: Spalten verstecken\n")
            _apply_hidden_columns(ws, hidden_columns)
            
            # ===== SCHRITT 11: Row Highlights =====
            sys.stderr.write(f"[PIPELINE] Schritt 11: Row Highlights\n")
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, len(headers) if headers else 0)
            
            # ===== SCHRITT 12: Tables reparieren =====
            sys.stderr.write(f"[PIPELINE] Schritt 12: Tables reparieren\n")
            table_changes = {}
            for table_name in ws.tables:
                table = ws.tables[table_name]
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                
                new_max_col = ws.max_column
                new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(new_max_col)}{max_row}"
                table.ref = new_ref
                if table.autoFilter:
                    table.autoFilter.ref = new_ref
                
                # tableColumns aus Header-Zellen neu aufbauen
                new_columns = []
                for col_idx in range(min_col, new_max_col + 1):
                    header_cell = ws.cell(row=min_row, column=col_idx)
                    col_name = str(header_cell.value) if header_cell.value else f"Column{col_idx}"
                    new_columns.append(TableColumn(id=col_idx - min_col + 1, name=col_name))
                
                table.tableColumns = new_columns
                table_changes[table_name] = {'ref': table.ref, 'columns': [col.name for col in new_columns]}
            
            # ===== SCHRITT 13: EINMAL speichern =====
            sys.stderr.write(f"[PIPELINE] Schritt 13: Speichern\n")
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # ===== SCHRITT 14: XML restore =====
            sys.stderr.write(f"[PIPELINE] Schritt 14: XML restore\n")
            if table_changes:
                restore_table_xml_from_original(output_path, original_path, table_changes)
            
            restore_external_links_from_original(output_path, original_path)
            
            return {'success': True, 'outputPath': output_path, 'method': 'openpyxl-pipeline'}
        
        # =====================================================================
        # LEGACY FALLBACK: Alte Einzel-FÄLLe für Kompatibilität
        # (werden nur noch erreicht wenn can_use_pipeline = False)
        # =====================================================================
        
        # LEGACY: Bei Spalten-Insert IMMER FALL 1.5 verwenden!
        only_column_insert = inserted_columns and not deleted_columns
        
        if only_column_insert:
            
            operations = inserted_columns.get('operations', [])
            if not operations and inserted_columns.get('position') is not None:
                operations = [{
                    'position': inserted_columns['position'],
                    'count': inserted_columns.get('count', 1),
                    'sourceColumn': inserted_columns.get('sourceColumn')
                }]
            
            # Sortiere aufsteigend - so kompensiert jede Einfügung die nächste automatisch
            operations.sort(key=lambda x: x['position'])
            
            from openpyxl.worksheet.table import TableColumn
            from openpyxl.utils.cell import range_boundaries
            
            # Alle Operationen im Speicher durchführen
            # Die Positionen vom Frontend sind die FINALEN Positionen (nach allen Einfügungen)
            # Wenn wir aufsteigend einfügen, brauchen wir keinen Offset!
            
            for op_idx, op in enumerate(operations):
                position = op['position']
                count = op.get('count', 1)
                source_column = op.get('sourceColumn')
                excel_col = position + 1  # 0-basiert → 1-basiert, KEIN Offset nötig!
                
                
                for i in range(count):
                    insert_at = excel_col + i
                    
                    # Speichere Formatierung der Referenzspalte (im aktuellen Zustand des Worksheets)
                    source_format = {}
                    source_width = None
                    if source_column is not None:
                        # source_column muss auch im aktuellen Worksheet-Zustand gefunden werden
                        # Nach vorherigen Einfügungen könnte die Position verschoben sein
                        source_excel_col = source_column + 1
                        # Korrigiere für bereits eingefügte Spalten
                        for prev_op in operations[:op_idx]:
                            if source_column >= prev_op['position']:
                                source_excel_col += prev_op.get('count', 1)
                        
                        col_letter = get_column_letter(source_excel_col)
                        if col_letter in ws.column_dimensions:
                            source_width = ws.column_dimensions[col_letter].width
                        
                        for row in range(1, ws.max_row + 1):
                            cell = ws.cell(row=row, column=source_excel_col)
                            source_format[row] = {
                                'fill': copy(cell.fill) if cell.fill else None,
                                'font': copy(cell.font) if cell.font else None,
                                'alignment': copy(cell.alignment) if cell.alignment else None,
                                'border': copy(cell.border) if cell.border else None,
                                'number_format': cell.number_format
                            }
                    
                    # Spaltenbreiten speichern
                    saved_widths = {}
                    for col in range(insert_at, ws.max_column + 1):
                        col_letter = get_column_letter(col)
                        if col_letter in ws.column_dimensions:
                            saved_widths[col] = ws.column_dimensions[col_letter].width
                    
                    # Spalte einfügen
                    ws.insert_cols(insert_at, 1)
                    
                    # Spaltenbreiten wiederherstellen
                    for old_col, width in saved_widths.items():
                        if width:
                            new_letter = get_column_letter(old_col + 1)
                            ws.column_dimensions[new_letter].width = width
                    
                    # CF anpassen
                    inserted_cols_for_cf = {insert_at - 1: 1}
                    adjust_conditional_formatting(ws, [], inserted_cols_for_cf)
                    
                    # Formatierung auf neue Spalte anwenden
                    if source_width:
                        ws.column_dimensions[get_column_letter(insert_at)].width = source_width
                    
                    for row, fmt in source_format.items():
                        cell = ws.cell(row=row, column=insert_at)
                        if fmt['fill']:
                            cell.fill = fmt['fill']
                        if fmt['font']:
                            cell.font = fmt['font']
                        if fmt['alignment']:
                            cell.alignment = fmt['alignment']
                        if fmt['border']:
                            cell.border = fmt['border']
                        if fmt.get('number_format'):
                            cell.number_format = fmt['number_format']
                
                # Header für neue Spalten setzen
                op_headers = op.get('headers', [])
                for i, header in enumerate(op_headers):
                    ws.cell(row=1, column=excel_col + i).value = header
                
                # Daten für diese Spalten schreiben
                if data and headers:
                    for i in range(count):
                        col_idx = position + i
                        if col_idx < len(headers):
                            for row_idx, row_data in enumerate(data):
                                if col_idx < len(row_data):
                                    cell = ws.cell(row=row_idx + 2, column=excel_col + i)
                                    apply_cell_value(cell, row_data[col_idx])
            
            # Versteckte Spalten/Zeilen
            _apply_hidden_columns(ws, hidden_columns)
            _apply_hidden_rows(ws, hidden_rows)
            
            # Row Highlights (FALL 1.5 - Spalten einfügen)
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, ws.max_column)
            
            # Tables reparieren: Am Ende EINMAL aus Header-Zellen neu aufbauen
            for table_name in ws.tables:
                table = ws.tables[table_name]
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                
                new_max_col = ws.max_column
                new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(new_max_col)}{max_row}"
                table.ref = new_ref
                if table.autoFilter:
                    table.autoFilter.ref = new_ref
                
                # Baue tableColumns aus den Header-Zellen
                new_columns = []
                for col_idx in range(min_col, new_max_col + 1):
                    header_cell = ws.cell(row=min_row, column=col_idx)
                    col_name = str(header_cell.value) if header_cell.value else f"Column{col_idx}"
                    new_columns.append(TableColumn(id=col_idx - min_col + 1, name=col_name))
                
                table.tableColumns = new_columns
            
            # Einmal speichern
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # Table-Infos für XML restore sammeln
            table_changes = {}
            wb_temp = load_workbook(output_path, rich_text=True)
            ws_temp = wb_temp[sheet_name]
            for table_name in ws_temp.tables:
                table = ws_temp.tables[table_name]
                col_names = [col.name for col in table.tableColumns]
                table_changes[table_name] = {'ref': table.ref, 'columns': col_names}
            wb_temp.close()
            
            # Original-Table-XML wiederherstellen (xr:uid etc.)
            if table_changes:
                restore_table_xml_from_original(output_path, original_path, table_changes)
            
            restore_external_links_from_original(output_path, original_path)
            
            return {'success': True, 'outputPath': output_path, 'method': 'openpyxl-insert-only'}
        
        # =====================================================================
        # FALL 1.9: Spalten LÖSCHEN UND EINFÜGEN kombiniert
        # SERIELL im Speicher - so bleibt die Formatierung erhalten!
        # =====================================================================
        column_delete_and_insert = deleted_columns and inserted_columns and row_mapping_is_identity
        
        if column_delete_and_insert:
            
            from openpyxl.worksheet.table import TableColumn
            from openpyxl.utils.cell import range_boundaries
            
            # ===== SCHRITT 1: Erst alle Spalten LÖSCHEN (von hinten nach vorne) =====
            sorted_deleted = sorted(deleted_columns, reverse=True)
            
            for col_idx in sorted_deleted:
                excel_col = col_idx + 1  # 0-basiert → 1-basiert
                max_col = ws.max_column
                
                # Spaltenbreiten speichern (rechts von der zu löschenden Spalte)
                saved_widths = {}
                for col in range(excel_col + 1, max_col + 1):
                    col_letter = get_column_letter(col)
                    if col_letter in ws.column_dimensions:
                        saved_widths[col] = ws.column_dimensions[col_letter].width
                
                # Spalte löschen
                ws.delete_cols(excel_col, 1)
                
                # Spaltenbreiten wiederherstellen (um 1 nach links verschoben)
                for old_col, width in saved_widths.items():
                    if width:
                        new_letter = get_column_letter(old_col - 1)
                        ws.column_dimensions[new_letter].width = width
                
                # CF anpassen
                adjust_conditional_formatting(ws, [col_idx], None)
            
            # ===== SCHRITT 2: Dann alle Spalten EINFÜGEN =====
            operations = inserted_columns.get('operations', [])
            if not operations and inserted_columns.get('position') is not None:
                operations = [{
                    'position': inserted_columns['position'],
                    'count': inserted_columns.get('count', 1),
                    'sourceColumn': inserted_columns.get('sourceColumn')
                }]
            
            # Sortiere aufsteigend
            operations.sort(key=lambda x: x['position'])
            
            for op_idx, op in enumerate(operations):
                position = op['position']
                count = op.get('count', 1)
                source_column = op.get('sourceColumn')
                excel_col = position + 1  # 0-basiert → 1-basiert
                
                for i in range(count):
                    insert_at = excel_col + i
                    
                    # Formatierung der Referenzspalte speichern
                    source_format = {}
                    source_width = None
                    if source_column is not None:
                        source_excel_col = source_column + 1
                        # Korrigiere für bereits eingefügte Spalten
                        for prev_op in operations[:op_idx]:
                            if source_column >= prev_op['position']:
                                source_excel_col += prev_op.get('count', 1)
                        
                        col_letter = get_column_letter(source_excel_col)
                        if col_letter in ws.column_dimensions:
                            source_width = ws.column_dimensions[col_letter].width
                        
                        for row in range(1, ws.max_row + 1):
                            cell = ws.cell(row=row, column=source_excel_col)
                            source_format[row] = {
                                'fill': copy(cell.fill) if cell.fill else None,
                                'font': copy(cell.font) if cell.font else None,
                                'alignment': copy(cell.alignment) if cell.alignment else None,
                                'border': copy(cell.border) if cell.border else None,
                                'number_format': cell.number_format
                            }
                    
                    # Spaltenbreiten speichern
                    saved_widths = {}
                    for col in range(insert_at, ws.max_column + 1):
                        col_letter = get_column_letter(col)
                        if col_letter in ws.column_dimensions:
                            saved_widths[col] = ws.column_dimensions[col_letter].width
                    
                    # Spalte einfügen
                    ws.insert_cols(insert_at, 1)
                    
                    # Spaltenbreiten wiederherstellen
                    for old_col, width in saved_widths.items():
                        if width:
                            new_letter = get_column_letter(old_col + 1)
                            ws.column_dimensions[new_letter].width = width
                    
                    # CF anpassen
                    inserted_cols_for_cf = {insert_at - 1: 1}
                    adjust_conditional_formatting(ws, [], inserted_cols_for_cf)
                    
                    # Formatierung auf neue Spalte anwenden
                    if source_width:
                        ws.column_dimensions[get_column_letter(insert_at)].width = source_width
                    
                    for row, fmt in source_format.items():
                        cell = ws.cell(row=row, column=insert_at)
                        if fmt['fill']:
                            cell.fill = fmt['fill']
                        if fmt['font']:
                            cell.font = fmt['font']
                        if fmt['alignment']:
                            cell.alignment = fmt['alignment']
                        if fmt['border']:
                            cell.border = fmt['border']
                        if fmt.get('number_format'):
                            cell.number_format = fmt['number_format']
                
                # Header für neue Spalten setzen
                op_headers = op.get('headers', [])
                for i, header in enumerate(op_headers):
                    ws.cell(row=1, column=excel_col + i).value = header
                
                # Daten für diese Spalten schreiben
                if data and headers:
                    for i in range(count):
                        col_idx = position + i
                        if col_idx < len(headers):
                            for row_idx, row_data in enumerate(data):
                                if col_idx < len(row_data):
                                    cell = ws.cell(row=row_idx + 2, column=excel_col + i)
                                    apply_cell_value(cell, row_data[col_idx])
            
            # Versteckte Spalten/Zeilen
            _apply_hidden_columns(ws, hidden_columns)
            _apply_hidden_rows(ws, hidden_rows)
            
            # Row Highlights (FALL 1.9 - Spalten löschen und einfügen)
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, ws.max_column)
            
            # Tables reparieren: Am Ende EINMAL aus Header-Zellen neu aufbauen
            for table_name in ws.tables:
                table = ws.tables[table_name]
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                
                new_max_col = ws.max_column
                new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(new_max_col)}{max_row}"
                table.ref = new_ref
                if table.autoFilter:
                    table.autoFilter.ref = new_ref
                
                # Baue tableColumns aus den Header-Zellen
                new_columns = []
                for col_idx in range(min_col, new_max_col + 1):
                    header_cell = ws.cell(row=min_row, column=col_idx)
                    col_name = str(header_cell.value) if header_cell.value else f"Column{col_idx}"
                    new_columns.append(TableColumn(id=col_idx - min_col + 1, name=col_name))
                
                table.tableColumns = new_columns
            
            # Einmal speichern
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # Table-Infos für XML restore sammeln
            table_changes = {}
            wb_temp = load_workbook(output_path, rich_text=True)
            ws_temp = wb_temp[sheet_name]
            for table_name in ws_temp.tables:
                table = ws_temp.tables[table_name]
                col_names = [col.name for col in table.tableColumns]
                table_changes[table_name] = {'ref': table.ref, 'columns': col_names}
            wb_temp.close()
            
            # Original-Table-XML wiederherstellen (xr:uid etc.)
            if table_changes:
                restore_table_xml_from_original(output_path, original_path, table_changes)
            
            restore_external_links_from_original(output_path, original_path)
            
            return {'success': True, 'outputPath': output_path, 'method': 'openpyxl-delete-and-insert'}
        
        # =====================================================================
        # FALL 1.6: Nur Spalten LÖSCHEN (keine anderen strukturellen Änderungen)
        # Analog zu FALL 1.5 - nutzt openpyxl's delete_cols() direkt
        # OHNE alle Daten neu zu schreiben - das erhält Table-Styles!
        # =====================================================================
        only_column_delete = deleted_columns and not inserted_columns and row_mapping_is_identity
        
        if only_column_delete:
            
            # Sortiere absteigend (von hinten nach vorne löschen)
            sorted_deleted = sorted(deleted_columns, reverse=True)
            
            for col_idx in sorted_deleted:
                excel_col = col_idx + 1  # 0-basiert → 1-basiert
                
                max_col = ws.max_column
                
                # 1. SPALTENBREITEN SPEICHERN (rechts von der zu löschenden Spalte)
                saved_widths = {}
                for col in range(excel_col + 1, max_col + 1):
                    col_letter = get_column_letter(col)
                    if col_letter in ws.column_dimensions:
                        saved_widths[col] = ws.column_dimensions[col_letter].width
                
                # 2. SPALTE LÖSCHEN (openpyxl verschiebt alles automatisch)
                ws.delete_cols(excel_col, 1)
                
                # 3. SPALTENBREITEN WIEDERHERSTELLEN (um 1 nach links verschoben)
                for old_col, width in saved_widths.items():
                    if width:
                        new_letter = get_column_letter(old_col - 1)
                        ws.column_dimensions[new_letter].width = width
                
                # 4. CF anpassen
                adjust_conditional_formatting(ws, [col_idx], None)
                
                # 5. Tables anpassen
                adjust_tables(ws, [col_idx], None, headers)
            
            # Versteckte Spalten/Zeilen
            _apply_hidden_columns(ws, hidden_columns)
            _apply_hidden_rows(ws, hidden_rows)
            
            # Row Highlights (FALL 1.6 - Spalten löschen)
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, ws.max_column)
            
            # Sammle Table-Infos für restore
            table_changes = {}
            for table_name in ws.tables:
                table = ws.tables[table_name]
                col_names = [col.name for col in table.tableColumns]
                table_changes[table_name] = {
                    'ref': table.ref,
                    'columns': col_names
                }
            
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # Stelle Original-Table-XML wieder her (mit korrekten xr:uid etc.)
            if table_changes:
                restore_table_xml_from_original(output_path, original_path, table_changes)
            
            # Stelle externalLinks aus Original wieder her (openpyxl verliert Namespaces)
            restore_external_links_from_original(output_path, original_path)
            
            return {'success': True, 'outputPath': output_path, 'method': 'openpyxl-delete-only'}
        
        # =====================================================================
        # FALL 1.7: NUR Spaltenreihenfolge ändern (ohne Insert/Delete)
        # Dieser Pfad ordnet Spalten physisch um, OHNE alle Zellen neu zu schreiben.
        # Das erhält Table-Styles (Zebra-Muster) perfekt!
        # =====================================================================
        only_column_order = (column_order and len(column_order) > 0 and 
                            not inserted_columns and not deleted_columns and 
                            row_mapping_is_identity and not affected_rows)
        
        if only_column_order:
            
            # Prüfe ob sich die Spaltenreihenfolge wirklich geändert hat
            columns_changed = False
            for new_idx, old_idx in enumerate(column_order):
                if new_idx != old_idx:
                    columns_changed = True
                    break
            
            if not columns_changed:
                pass  # Keine Änderung nötig
            else:
                # Physische Spaltenumordnung durch Swap-Operationen
                # column_order[neue_position] = alte_position
                
                from openpyxl.cell.cell import MergedCell
                
                num_cols = len(column_order)
                max_row = ws.max_row
                
                # Temporärer Speicher für alle Spalten (Werte + Hyperlinks)
                temp_columns = {}
                
                # SCHRITT 1: Alle Spalten in temp_columns speichern
                for old_col_idx in range(num_cols):
                    old_excel_col = old_col_idx + 1
                    temp_columns[old_col_idx] = {}
                    
                    for row in range(1, max_row + 1):
                        cell = ws.cell(row=row, column=old_excel_col)
                        if isinstance(cell, MergedCell):
                            continue
                        temp_columns[old_col_idx][row] = {
                            'value': cell.value,
                            'hyperlink': cell.hyperlink.target if cell.hyperlink else None,
                        }
                
                # SCHRITT 2: Spalten in neuer Reihenfolge schreiben
                for new_col_idx, old_col_idx in enumerate(column_order):
                    new_excel_col = new_col_idx + 1
                    
                    if old_col_idx not in temp_columns:
                        continue
                    
                    for row, data_item in temp_columns[old_col_idx].items():
                        cell = ws.cell(row=row, column=new_excel_col)
                        if isinstance(cell, MergedCell):
                            continue
                        
                        # Nur Wert und Hyperlink setzen - KEINE Formatierung!
                        # So bleibt das Table-Style-Zebra-Muster erhalten
                        cell.value = data_item['value']
                        if data_item['hyperlink']:
                            cell.hyperlink = data_item['hyperlink']
                
            
            # Versteckte Spalten/Zeilen anwenden
            _apply_hidden_columns(ws, hidden_columns, len(headers))
            _apply_hidden_rows(ws, hidden_rows, len(data) if data else 0)
            
            # Row Highlights
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, len(headers))
            
            # WICHTIG: Bei Spalten-Verschieben die tableColumns AKTUALISIEREN!
            # Die Spalten wurden physisch umgeordnet, also müssen die Column-Namen
            # aus den Header-Zellen neu gelesen werden.
            from openpyxl.worksheet.table import TableColumn
            from openpyxl.utils.cell import range_boundaries
            
            table_changes = {}
            for table_name in ws.tables:
                table = ws.tables[table_name]
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                
                # Baue tableColumns aus den Header-Zellen (die sind jetzt umgeordnet)
                new_columns = []
                for col_idx in range(min_col, max_col + 1):
                    header_cell = ws.cell(row=min_row, column=col_idx)
                    col_name = str(header_cell.value) if header_cell.value else f"Column{col_idx}"
                    new_columns.append(TableColumn(id=col_idx - min_col + 1, name=col_name))
                
                table.tableColumns = new_columns
                
                col_names = [col.name for col in new_columns]
                table_changes[table_name] = {
                    'ref': table.ref,
                    'columns': col_names
                }
            
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # Stelle Table-XML aus Original wieder her MIT der neuen Spaltenreihenfolge
            if table_changes:
                restore_table_xml_from_original(output_path, original_path, table_changes)
            
            # Stelle externalLinks aus Original wieder her
            restore_external_links_from_original(output_path, original_path)
            
            return {'success': True, 'outputPath': output_path, 'method': 'openpyxl-column-order'}
        
        # =====================================================================
        # FALL 2: Strukturelle Änderungen (fullRewrite)
        # WICHTIG: openpyxl's delete_cols() passt CF-Bereiche NICHT an!
        # Wenn Excel installiert ist, nutzen wir xlwings für perfekten CF-Erhalt.
        # =====================================================================
        if structural_change or full_rewrite:
            import sys
            sys.stderr.write(f"[FALL 2] structural_change={structural_change}, full_rewrite={full_rewrite}, row_mapping={'ja' if row_mapping else 'nein'}\n")
            sys.stderr.write(f"[FALL 2] file_path={file_path}\n")
            sys.stderr.write(f"[FALL 2] output_path={output_path}\n")
            sys.stderr.write(f"[FALL 2] original_path={original_path}\n")
            if row_mapping:
                sys.stderr.write(f"[FALL 2] row_mapping (erste 10): {row_mapping[:10] if len(row_mapping) > 10 else row_mapping}\n")
            
            # OPTION A: Nutze xlwings wenn Excel verfügbar ist
            # Das erhält ALLE Formatierungen inkl. CF perfekt!
            # TEMPORÄR DEAKTIVIERT FÜR FALLBACK-TEST
            use_excel_for_structural = False  # (deleted_columns or inserted_columns) and is_excel_installed()
            if use_excel_for_structural:
                wb.close()  # Workbook schließen, damit Excel es öffnen kann
                
                # Strukturelle Änderungen mit Excel durchführen
                success = structural_change_with_excel(
                    file_path, output_path, sheet_name,
                    deleted_columns=deleted_columns,
                    inserted_columns=inserted_columns,
                    deleted_rows=None  # TODO: deleted_rows implementieren
                )
                
                if success:
                    # Datei erneut öffnen um Daten zu schreiben
                    wb = load_workbook(output_path, rich_text=True)
                    ws = wb[sheet_name]
                    
                    # Header und Daten schreiben (die Struktur ist jetzt korrekt)
                    for col_idx, header in enumerate(headers):
                        ws.cell(row=1, column=col_idx + 1, value=header)
                    
                    for row_idx, row_data in enumerate(data):
                        excel_row = row_idx + 2
                        for col_idx, value in enumerate(row_data):
                            cell = ws.cell(row=excel_row, column=col_idx + 1)
                            apply_cell_value(cell, value)
                    
                    _apply_hidden_columns(ws, hidden_columns, len(headers))
                    _apply_hidden_rows(ws, hidden_rows, len(data))
                    
                    if row_highlights:
                        _apply_row_highlights(ws, row_highlights, len(headers))
                    
                    wb.save(output_path)
                    wb.close()
                    fix_xlsx_relationships(output_path)
                    return {
                        'success': True, 
                        'outputPath': output_path,
                        'method': 'xlwings',
                        'cfPreserved': True
                    }
                else:
                    wb = load_workbook(file_path, rich_text=True)
                    ws = wb[sheet_name]
            
            # ================================================================
            # NEUER ANSATZ FÜR ROW_MAPPING: shutil.copy() + nur Werte ändern
            # ================================================================
            # Wenn Zeilen gelöscht oder eingefügt wurden (row_mapping vorhanden), nutzen wir
            # den shutil-Ansatz: Original kopieren, dann NUR Zeilenreihenfolge ändern.
            # Das erhält ALLE Formatierungen perfekt!
            # ================================================================
            if row_mapping and len(row_mapping) > 0:
                identity_mapping = list(range(len(row_mapping)))
                current_max_row = ws.max_row
                rows_changed = current_max_row - 1 - len(row_mapping)  # -1 für Header (positiv=gelöscht, negativ=eingefügt)
                
                # DEBUG: Zeige alle relevanten Variablen
                sys.stderr.write(f"[ZIP-DEBUG] current_max_row (ws.max_row)={current_max_row}\n")
                sys.stderr.write(f"[ZIP-DEBUG] len(row_mapping)={len(row_mapping)}\n")
                sys.stderr.write(f"[ZIP-DEBUG] rows_changed={rows_changed}\n")
                sys.stderr.write(f"[ZIP-DEBUG] deleted_rows aus Frontend={deleted_rows}\n")
                sys.stderr.write(f"[ZIP-DEBUG] row_mapping[:10]={row_mapping[:10]}\n")
                
                # ZIP-Ansatz aktivieren wenn:
                # - Zeilen gelöscht wurden (rows_changed > 0)
                # - Zeilen eingefügt wurden (rows_changed < 0)
                # - Zeilen umsortiert wurden (row_mapping != identity_mapping)
                if row_mapping != identity_mapping or rows_changed != 0:
                    import shutil
                    import tempfile
                    import zipfile
                    import re
                    from lxml import etree
                    
                    action = "gelöschte" if rows_changed > 0 else "eingefügte" if rows_changed < 0 else "umsortierte"
                    sys.stderr.write(f"[ZIP-ANSATZ] Verwende direkte XML-Manipulation für {abs(rows_changed)} {action} Zeilen\n")
                    
                    # Workbook schließen (ohne zu speichern!)
                    wb.close()
                    
                    # WICHTIG: Wir kopieren die ORIGINAL-Datei (nicht file_path, das ist schon die Export-Datei!)
                    # original_path enthält die unberührte Formatierung
                    basis_datei = original_path if original_path else file_path
                    sys.stderr.write(f"[ZIP-ANSATZ] Basis-Datei: {basis_datei}\n")
                    
                    # Immer die Basis-Datei zur Ausgabe kopieren (erhält ALLE Formatierungen!)
                    shutil.copy2(basis_datei, output_path)
                    sys.stderr.write(f"[ZIP-ANSATZ] Datei kopiert: {basis_datei} -> {output_path}\n")
                    
                    # Jetzt direkt die XML im ZIP manipulieren
                    # xlsx ist ein ZIP mit XML-Dateien drin
                    
                    # Finde das richtige Sheet
                    sheet_xml_path = None
                    with zipfile.ZipFile(output_path, 'r') as zf:
                        # Lese workbook.xml um Sheet-Namen zu finden
                        workbook_xml = zf.read('xl/workbook.xml')
                        wb_tree = etree.fromstring(workbook_xml)
                        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                        
                        for sheet_elem in wb_tree.findall('.//main:sheet', ns):
                            if sheet_elem.get('name') == sheet_name:
                                # rId aus Attribut holen
                                r_id = sheet_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                                
                                # Relationships lesen um Pfad zu finden
                                rels_xml = zf.read('xl/_rels/workbook.xml.rels')
                                rels_tree = etree.fromstring(rels_xml)
                                
                                for rel in rels_tree:
                                    if rel.get('Id') == r_id:
                                        sheet_xml_path = 'xl/' + rel.get('Target')
                                        break
                                break
                    
                    if not sheet_xml_path:
                        sys.stderr.write(f"[ZIP-ANSATZ] Sheet {sheet_name} nicht gefunden, fallback zu openpyxl\n")
                        wb = load_workbook(output_path, rich_text=True)
                        ws = wb[sheet_name]
                    else:
                        sys.stderr.write(f"[ZIP-ANSATZ] Sheet XML: {sheet_xml_path}\n")
                        
                        # Sheet-XML lesen und modifizieren
                        with zipfile.ZipFile(output_path, 'r') as zf:
                            sheet_xml = zf.read(sheet_xml_path)
                        
                        sheet_tree = etree.fromstring(sheet_xml)
                        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                        
                        # sharedStrings.xml lesen (für String-Werte)
                        shared_strings = []
                        try:
                            with zipfile.ZipFile(output_path, 'r') as zf:
                                ss_xml = zf.read('xl/sharedStrings.xml')
                                ss_tree = etree.fromstring(ss_xml)
                                for si in ss_tree.findall('.//main:si', ns):
                                    t_elem = si.find('.//main:t', ns)
                                    if t_elem is not None and t_elem.text:
                                        shared_strings.append(t_elem.text)
                                    else:
                                        shared_strings.append('')
                        except Exception:
                            pass
                        
                        # Finde sheetData Element
                        sheet_data = sheet_tree.find('.//main:sheetData', ns)
                        new_max_row = len(data) + 1  # +1 für Header
                        
                        # Aktualisiere dimension-Element wenn vorhanden
                        dimension = sheet_tree.find('.//main:dimension', ns)
                        if dimension is not None:
                            ref = dimension.get('ref')
                            if ref and ':' in ref:
                                match = re.match(r'([A-Z]+\d+):([A-Z]+)(\d+)', ref)
                                if match:
                                    start_ref, end_col, old_end_row = match.groups()
                                    new_ref = f"{start_ref}:{end_col}{new_max_row}"
                                    dimension.set('ref', new_ref)
                                    sys.stderr.write(f"[ZIP-ANSATZ] Dimension: {ref} -> {new_ref}\n")
                        
                        # Aktualisiere autoFilter wenn vorhanden
                        auto_filter = sheet_tree.find('.//main:autoFilter', ns)
                        if auto_filter is not None:
                            af_ref = auto_filter.get('ref')
                            if af_ref:
                                match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', af_ref)
                                if match:
                                    start_col, start_row, end_col, end_row = match.groups()
                                    new_af_ref = f"{start_col}{start_row}:{end_col}{new_max_row}"
                                    auto_filter.set('ref', new_af_ref)
                                    sys.stderr.write(f"[ZIP-ANSATZ] AutoFilter: {af_ref} -> {new_af_ref}\n")
                        
                        # ================================================================
                        # KORREKTUR: row_mapping[new_idx] = ORIGINAL_idx (NICHT after-delete!)
                        # Das Frontend schickt bereits die Original-Indizes:
                        # - Beim Löschen: originalIdx = i >= rowIndex ? i + 1 : i
                        # - Bei eingefügten Zeilen: -1
                        # 
                        # Kein Zurückmappen nötig!
                        # ================================================================
                        
                        # Verwende deleted_rows aus dem Frontend für CF-Anpassung
                        frontend_deleted_rows = set(deleted_rows) if deleted_rows else set()
                        sys.stderr.write(f"[ZIP-ANSATZ] Frontend deleted_rows: {sorted(frontend_deleted_rows)[:10] if frontend_deleted_rows else 'keine'}\n")
                        
                        # row_mapping enthält bereits ORIGINAL-Indizes!
                        # row_mapping[new_idx] = original_idx
                        row_shift_map = {}  # old_excel_row -> new_excel_row
                        inserted_rows_set = set()  # neue Zeilen die eingefügt wurden
                        
                        for new_idx, original_idx in enumerate(row_mapping):
                            new_excel_row = new_idx + 2
                            if original_idx < 0:
                                # Neue eingefügte Zeile (original_idx = -1)
                                inserted_rows_set.add(new_excel_row)
                            else:
                                # original_idx ist bereits der Original-Index!
                                old_excel_row = original_idx + 2  # +2 für Header
                                row_shift_map[old_excel_row] = new_excel_row
                        
                        # Finde gelöschte Zeilen als Excel-Zeilen
                        deleted_excel_rows = set(idx + 2 for idx in frontend_deleted_rows)
                        
                        # Debug: Zeige die ersten Mappings
                        sys.stderr.write(f"[ZIP-ANSATZ] row_mapping (erste 10): {row_mapping[:10]}\n")
                        first_mappings = list(row_shift_map.items())[:5]
                        sys.stderr.write(f"[ZIP-ANSATZ] row_shift_map (erste 5): {first_mappings}\n")
                        
                        # Bestimme ob nur Verschiebung (keine Löschung/Einfügung)
                        is_pure_reorder = len(frontend_deleted_rows) == 0 and len(inserted_rows_set) == 0
                        
                        if deleted_excel_rows:
                            sys.stderr.write(f"[ZIP-ANSATZ] Gelöschte Zeilen (Excel): {sorted(deleted_excel_rows)[:10]}...\n")
                        if inserted_rows_set:
                            sys.stderr.write(f"[ZIP-ANSATZ] Eingefügte Zeilen: {sorted(inserted_rows_set)[:10]}...\n")
                        if is_pure_reorder:
                            sys.stderr.write(f"[ZIP-ANSATZ] Reine Verschiebung - CF-Bereiche werden NICHT angepasst\n")
                        
                        cf_elements = sheet_tree.findall('.//main:conditionalFormatting', ns)
                        cf_updated = 0
                        cf_removed = 0
                        
                        # Bei reiner Verschiebung: CF nicht anpassen (Excel-Standardverhalten)
                        # Die Zeilen wandern, aber die CF-Regeln bleiben an ihren Positionen
                        # Das bedeutet: Die neue Zeile an Position X bekommt die CF von Position X
                        if not is_pure_reorder:
                            for cf in cf_elements:
                                sqref = cf.get('sqref')
                                if sqref:
                                    new_ranges = []
                                    changed = False
                                    
                                    for range_part in sqref.split():
                                        range_match = re.match(r'([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?', range_part)
                                        if range_match:
                                            start_col, start_row_str, end_col, end_row_str = range_match.groups()
                                            start_row = int(start_row_str)
                                            
                                            if end_row_str:
                                                # Bereich wie L2:L2404
                                                end_row = int(end_row_str)
                                                # Neue Start-Zeile berechnen
                                                if start_row in row_shift_map:
                                                    new_start = row_shift_map[start_row]
                                                elif start_row in deleted_rows:
                                                    # Start wurde gelöscht - finde nächste gültige Zeile
                                                    new_start = None
                                                    for r in range(start_row + 1, end_row + 1):
                                                        if r in row_shift_map:
                                                            new_start = row_shift_map[r]
                                                            break
                                                    if new_start is None:
                                                        # Ganzer Bereich gelöscht - überspringen
                                                        changed = True
                                                        continue
                                                else:
                                                    # Zeile 1 (Header) - bleibt
                                                    new_start = start_row
                                                
                                                # Neue End-Zeile berechnen
                                                if end_row in row_shift_map:
                                                    new_end = row_shift_map[end_row]
                                                elif end_row >= current_max_row:
                                                    new_end = new_max_row
                                                else:
                                                    # Zeile wurde gelöscht - finde nächste gültige davor
                                                    new_end = None
                                                    for r in range(end_row, start_row - 1, -1):
                                                        if r in row_shift_map:
                                                            new_end = row_shift_map[r]
                                                            break
                                                    if new_end is None:
                                                        new_end = new_max_row
                                                
                                                if new_start != start_row or new_end != end_row:
                                                    changed = True
                                                
                                                new_range = f"{start_col}{new_start}:{end_col}{new_end}"
                                                new_ranges.append(new_range)
                                            else:
                                                # Einzelne Zelle wie A5
                                                if start_row in deleted_rows:
                                                    # Zelle wurde gelöscht - überspringen
                                                    changed = True
                                                    continue
                                                
                                                new_row = row_shift_map.get(start_row, start_row)
                                                if new_row != start_row:
                                                    changed = True
                                                new_ranges.append(f"{start_col}{new_row}")
                                        else:
                                            new_ranges.append(range_part)
                                    
                                    if changed:
                                        new_sqref = ' '.join(new_ranges)
                                        cf.set('sqref', new_sqref)
                                        cf_updated += 1
                                        
                                        # Auch die Formeln in den cfRule-Elementen anpassen
                                        for rule in cf.findall('main:cfRule', ns):
                                            for formula in rule.findall('main:formula', ns):
                                                if formula.text:
                                                    # Zellreferenzen in Formel anpassen
                                                    # z.B. $K2 oder K2 oder $K$2
                                                    def adjust_cell_ref(match):
                                                        col = match.group(1)
                                                        row_num = int(match.group(2))
                                                        if row_num in row_shift_map:
                                                            return f"{col}{row_shift_map[row_num]}"
                                                        elif row_num in deleted_rows:
                                                            # Zeile gelöscht - nehme nächste gültige
                                                            for r in range(row_num + 1, current_max_row + 1):
                                                                if r in row_shift_map:
                                                                    return f"{col}{row_shift_map[r]}"
                                                        return match.group(0)
                                                    
                                                    new_formula = re.sub(r'(\$?[A-Z]+\$?)(\d+)', adjust_cell_ref, formula.text)
                                                    if new_formula != formula.text:
                                                        formula.text = new_formula
                        
                        if cf_updated > 0:
                            sys.stderr.write(f"[ZIP-ANSATZ] {cf_updated} CF-Bereiche angepasst\n")
                        
                        if sheet_data is not None:
                            # Zellen aktualisieren basierend auf row_mapping
                            # row_mapping[new_idx] = original_idx (original Excel-Zeile)
                            
                            # Sammle alle Zeilen
                            rows = sheet_data.findall('main:row', ns)
                            row_dict = {}
                            for row_elem in rows:
                                row_num = int(row_elem.get('r'))
                                row_dict[row_num] = row_elem
                            
                            # Strategie: 
                            # row_mapping[new_idx] = original_idx (0-basiert, ohne Header)
                            # Das bedeutet: Datenzeile new_idx sollte die Formatierung von Original-Zeile original_idx+2 haben
                            # 
                            # Wir müssen:
                            # 1. Für jede neue Position new_row (2, 3, 4, ...):
                            #    - Die XML-Zeile von original_row = row_mapping[new_row-2] + 2 nehmen
                            #    - Diese Zeile auf new_row umnummerieren
                            #    - Die Werte aus data[new_row-2] einsetzen
                            
                            # Erstelle neue sheetData mit korrekt angeordneten Zeilen
                            new_rows = []
                            
                            # Header (Zeile 1) bleibt
                            if 1 in row_dict:
                                new_rows.append((1, row_dict[1]))
                            
                            # Finde eine Vorlage-Zeile für neue eingefügte Zeilen
                            # Wir nehmen die erste existierende Datenzeile als Vorlage
                            template_row = None
                            for r in range(2, current_max_row + 1):
                                if r in row_dict:
                                    template_row = row_dict[r]
                                    break
                            
                            # Datenzeilen umsortieren
                            # row_mapping[new_idx] = original_idx (BEREITS Original-Index!)
                            for new_data_idx, original_idx in enumerate(row_mapping):
                                new_excel_row = new_data_idx + 2  # Ziel-Zeile in Excel
                                
                                if original_idx < 0:
                                    # NEUE EINGEFÜGTE ZEILE - muss erstellt werden
                                    if template_row is not None:
                                        from copy import deepcopy
                                        new_row_elem = deepcopy(template_row)
                                        new_row_elem.set('r', str(new_excel_row))
                                        
                                        # Alle Zellen umnummerieren und Werte leeren
                                        cells = new_row_elem.findall('main:c', ns)
                                        for cell in cells:
                                            old_ref = cell.get('r')
                                            if old_ref:
                                                col_match = re.match(r'([A-Z]+)\d+', old_ref)
                                                if col_match:
                                                    col = col_match.group(1)
                                                    cell.set('r', f"{col}{new_excel_row}")
                                                    # Wert leeren für neue Zeile
                                                    v_elem = cell.find('main:v', ns)
                                                    if v_elem is not None:
                                                        cell.remove(v_elem)
                                                    is_elem = cell.find('main:is', ns)
                                                    if is_elem is not None:
                                                        cell.remove(is_elem)
                                        
                                        new_rows.append((new_excel_row, new_row_elem))
                                        sys.stderr.write(f"[ZIP-ANSATZ] Neue Zeile {new_excel_row} erstellt\n")
                                else:
                                    # original_idx ist bereits der Original-Index!
                                    orig_excel_row = original_idx + 2  # Original Excel-Zeile
                                    
                                    # Debug für erste 5 Zeilen
                                    if new_data_idx < 5:
                                        sys.stderr.write(f"[ZIP-ANSATZ] Mapping: neue Pos {new_data_idx} (Excel {new_excel_row}) <- original {original_idx} (Excel {orig_excel_row})\n")
                                    
                                    if orig_excel_row in row_dict:
                                        # WICHTIG: deepcopy machen, damit das Original nicht modifiziert wird!
                                        from copy import deepcopy
                                        row_elem = deepcopy(row_dict[orig_excel_row])
                                        
                                        # Zeile umnummerieren
                                        row_elem.set('r', str(new_excel_row))
                                        
                                        # Alle Zellen in der Zeile umnummerieren
                                        cells = row_elem.findall('main:c', ns)
                                        for cell in cells:
                                            old_ref = cell.get('r')
                                            if old_ref:
                                                col_match = re.match(r'([A-Z]+)\d+', old_ref)
                                                if col_match:
                                                    col = col_match.group(1)
                                                    cell.set('r', f"{col}{new_excel_row}")
                                        
                                        new_rows.append((new_excel_row, row_elem))
                                    else:
                                        sys.stderr.write(f"[ZIP-ANSATZ] WARNUNG: Zeile {orig_excel_row} nicht gefunden für Position {new_excel_row}\n")
                            
                            # Alle alten Zeilen entfernen
                            for row_elem in list(sheet_data):
                                sheet_data.remove(row_elem)
                            
                            # Neue Zeilen in korrekter Reihenfolge einfügen
                            new_rows.sort(key=lambda x: x[0])
                            for row_num, row_elem in new_rows:
                                sheet_data.append(row_elem)
                            
                            sys.stderr.write(f"[ZIP-ANSATZ] {len(new_rows)} Zeilen neu angeordnet\n")
                        
                        # ===== HIDDEN ROWS: Versteckte Zeilen im XML setzen =====
                        # hidden_rows enthält 0-basierte Indizes, Excel-Zeilen sind 1-basiert (+2 für Header)
                        if hidden_rows:
                            sys.stderr.write(f"[ZIP-ANSATZ] Verstecke Zeilen: {hidden_rows}\n")
                            
                            # Finde oder erstelle sheetFormatPr Element
                            sheet_format_pr = sheet_tree.find('.//main:sheetFormatPr', ns)
                            
                            # Für jeden hidden row, setze das hidden-Attribut in der row
                            hidden_set = set(hidden_rows)
                            if sheet_data is not None:
                                rows = sheet_data.findall('main:row', ns)
                                for row_elem in rows:
                                    row_num = int(row_elem.get('r'))
                                    row_idx = row_num - 2  # 0-basierter Index (ohne Header)
                                    
                                    if row_idx in hidden_set:
                                        row_elem.set('hidden', '1')
                                        sys.stderr.write(f"[ZIP-ANSATZ] Zeile {row_num} (idx={row_idx}) versteckt\n")
                                    else:
                                        # Sicherstellen dass nicht-versteckte Zeilen hidden=0 haben
                                        if row_elem.get('hidden') == '1':
                                            row_elem.set('hidden', '0')
                        
                        # Speichere modifizierte Sheet-XML
                        new_sheet_xml = etree.tostring(sheet_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                        
                        # Finde und aktualisiere Table-Definitionen (für Zebra-Style)
                        # Tables sind in xl/tables/table*.xml
                        modified_tables = {}
                        try:
                            with zipfile.ZipFile(output_path, 'r') as zf:
                                for name in zf.namelist():
                                    if name.startswith('xl/tables/table') and name.endswith('.xml'):
                                        table_xml = zf.read(name)
                                        table_tree = etree.fromstring(table_xml)
                                        
                                        # Prüfe ob diese Tabelle zum aktuellen Sheet gehört
                                        # (vereinfacht: wir aktualisieren alle Tables die im richtigen Bereich sind)
                                        ref = table_tree.get('ref')
                                        if ref:
                                            # Parse ref wie "A1:AZ500"
                                            match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', ref)
                                            if match:
                                                start_col, start_row, end_col, end_row = match.groups()
                                                start_row = int(start_row)
                                                end_row = int(end_row)
                                                
                                                # Wenn Tabelle bei Zeile 1 startet, ist es wahrscheinlich unsere Datentabelle
                                                if start_row == 1:
                                                    new_end_row = new_max_row
                                                    new_ref = f"{start_col}{start_row}:{end_col}{new_end_row}"
                                                    table_tree.set('ref', new_ref)
                                                    
                                                    # Auch autoFilter anpassen wenn vorhanden
                                                    af = table_tree.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}autoFilter')
                                                    if af is not None:
                                                        af.set('ref', new_ref)
                                                    
                                                    modified_tables[name] = etree.tostring(table_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                                                    sys.stderr.write(f"[ZIP-ANSATZ] Table {name}: {ref} -> {new_ref}\n")
                        except Exception as e:
                            sys.stderr.write(f"[ZIP-ANSATZ] Table-Anpassung Fehler: {e}\n")
                        
                        # ZIP aktualisieren mit allen Änderungen
                        temp_zip = output_path + '.tmp'
                        with zipfile.ZipFile(output_path, 'r') as zin:
                            with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
                                for item in zin.infolist():
                                    if item.filename == sheet_xml_path:
                                        zout.writestr(item, new_sheet_xml)
                                    elif item.filename in modified_tables:
                                        zout.writestr(item, modified_tables[item.filename])
                                    else:
                                        zout.writestr(item, zin.read(item.filename))
                        
                        shutil.move(temp_zip, output_path)
                        
                        # Row Highlights müssen NACH dem ZIP-Ansatz angewendet werden
                        # Da ZIP nur XML manipuliert, öffnen wir die Datei erneut für Highlights
                        if row_highlights or cleared_row_highlights:
                            wb_hl = load_workbook(output_path, rich_text=True)
                            ws_hl = wb_hl[sheet_name]
                            
                            # Markierungen anwenden
                            if row_highlights:
                                sys.stderr.write(f"[ZIP-ANSATZ] Wende {len(row_highlights)} Row Highlights an\n")
                                _apply_row_highlights(ws_hl, row_highlights, ws_hl.max_column)
                            
                            # Markierungen entfernen
                            if cleared_row_highlights:
                                sys.stderr.write(f"[ZIP-ANSATZ] Entferne {len(cleared_row_highlights)} Row Highlights\n")
                                for row_idx in cleared_row_highlights:
                                    excel_row = row_idx + 2
                                    for col_idx in range(1, ws_hl.max_column + 1):
                                        cell = ws_hl.cell(row=excel_row, column=col_idx)
                                        cell.fill = PatternFill()  # Keine Füllung
                            
                            wb_hl.save(output_path)
                            wb_hl.close()
                            fix_xlsx_relationships(output_path)
                            restore_table_xml_from_original(output_path, original_path, table_changes=None)
                            restore_external_links_from_original(output_path, original_path)
                        
                        sys.stderr.write(f"[ZIP-ANSATZ] Erfolgreich gespeichert\n")
                        return {
                            'success': True,
                            'outputPath': output_path,
                            'method': 'direct-xml-manipulation'
                        }
                    
                    # NUR die Zellwerte überschreiben (Formatierungen bleiben!)
                    # Die Daten werden in neuer Reihenfolge geschrieben
                    for new_row_idx, row_data in enumerate(data):
                        excel_row = new_row_idx + 2  # +2 für Header
                        for col_idx, value in enumerate(row_data):
                            if col_idx < len(headers):  # Nur vorhandene Spalten
                                cell = ws.cell(row=excel_row, column=col_idx + 1)
                                apply_cell_value(cell, value)
                    
                    # Header aktualisieren
                    for col_idx, header in enumerate(headers):
                        ws.cell(row=1, column=col_idx + 1, value=header)
                    
                    # Überschüssige Zeilen am Ende leeren (nur Werte, Formatierung bleibt)
                    new_max_row = len(data) + 1  # +1 für Header
                    old_max_row = ws.max_row
                    if old_max_row > new_max_row:
                        sys.stderr.write(f"[SHUTIL-ANSATZ] Leere Zeilen {new_max_row + 1} bis {old_max_row}\n")
                        for row in range(new_max_row + 1, old_max_row + 1):
                            for col in range(1, len(headers) + 1):
                                cell = ws.cell(row=row, column=col)
                                cell.value = None
                    
                    # CF-Bereiche anpassen (die Zeilennummern müssen angepasst werden)
                    # current_max_row wurde VOR dem Schließen gespeichert
                    adjust_cf_for_row_changes(ws, row_mapping, current_max_row - 1)
                    
                    # Hidden Rows/Columns anwenden
                    _apply_hidden_columns(ws, hidden_columns, len(headers))
                    _apply_hidden_rows(ws, hidden_rows, len(data))
                    
                    # Row Highlights anwenden
                    if row_highlights:
                        _apply_row_highlights(ws, row_highlights, len(headers))
                    
                    # Cleared Row Highlights entfernen
                    if cleared_row_highlights:
                        sys.stderr.write(f"[SHUTIL-ANSATZ] Entferne {len(cleared_row_highlights)} Row Highlights\n")
                        for row_idx in cleared_row_highlights:
                            excel_row = row_idx + 2
                            for col_idx in range(1, len(headers) + 1):
                                cell = ws.cell(row=excel_row, column=col_idx)
                                cell.fill = PatternFill()  # Keine Füllung
                    
                    # AutoFilter setzen
                    if frontend_auto_filter or original_auto_filter:
                        try:
                            af_ref = f"A1:{get_column_letter(len(headers))}{new_max_row}"
                            ws.auto_filter.ref = af_ref
                        except Exception:
                            pass
                    
                    # Speichern und fertig
                    wb.save(output_path)
                    wb.close()
                    fix_xlsx_relationships(output_path)
                    
                    # WICHTIG: Table-XML vom Original wiederherstellen!
                    restore_table_xml_from_original(output_path, original_path, table_changes=None)
                    restore_external_links_from_original(output_path, original_path)
                    
                    sys.stderr.write(f"[SHUTIL-ANSATZ] Erfolgreich gespeichert\n")
                    return {
                        'success': True,
                        'outputPath': output_path,
                        'method': 'openpyxl-shutil-copy'
                    }
            
            # OPTION B: openpyxl mit insert_cols/delete_cols
            # 
            # RICHTIGER ANSATZ: insert_cols() und delete_cols() verwenden!
            # Diese Funktionen verschieben automatisch ALLE Formatierungen mit.
            
            # SCHRITT 0: AUTOFILTER VOR ALLEM SPEICHERN UND ENTFERNEN
            original_auto_filter = ws.auto_filter.ref or frontend_auto_filter
            if ws.auto_filter.ref:
                ws.auto_filter.ref = None  # AutoFilter temporär entfernen
            
            # Speichere Original-Spaltenzahl VOR allen Änderungen
            original_max_col = ws.max_column
            original_max_row = ws.max_row
            target_col_count = len(headers)
            
            # ================================================================
            # SCHRITT 0.5: ZEILEN PHYSISCH UMORDNEN (bei row_mapping)
            # row_mapping[neue_position] = original_daten_row_idx (0-basiert)
            # Kopiert alle Zellen mit Formatierung
            # ================================================================
            if row_mapping and len(row_mapping) > 0:
                from openpyxl.cell.cell import MergedCell
                
                # Prüfe ob tatsächlich eine Umordnung nötig ist
                identity_mapping = list(range(len(row_mapping)))
                needs_reorder = row_mapping != identity_mapping
                
                if needs_reorder:
                    # Speichere alle benötigten Zeilen mit Formatierung
                    # Key = Original-Daten-Index (0-basiert), Value = Zellen-Info
                    row_data_with_styles = {}
                    max_col = ws.max_column
                    
                    # Sammle alle Hyperlinks der Originaldatei
                    original_hyperlinks = {}
                    for row_idx in range(2, ws.max_row + 1):
                        for col_idx in range(1, max_col + 1):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            if cell.hyperlink:
                                if row_idx not in original_hyperlinks:
                                    original_hyperlinks[row_idx] = {}
                                original_hyperlinks[row_idx][col_idx] = cell.hyperlink.target
                    
                    # Prüfe ob openpyxl CellRichText unterstützt
                    try:
                        from openpyxl.cell.rich_text import CellRichText
                        has_rich_text_support = True
                    except ImportError:
                        has_rich_text_support = False
                    
                    # Sammle alle Original-Zeilen die wir brauchen
                    styles_found = 0
                    for orig_data_idx in set(row_mapping):
                        excel_row = orig_data_idx + 2  # +2: Excel 1-basiert + Header
                        row_info = {}
                        for col_idx in range(1, max_col + 1):
                            cell = ws.cell(row=excel_row, column=col_idx)
                            if isinstance(cell, MergedCell):
                                continue
                            
                            # Prüfe ob der Wert RichText ist
                            cell_value = cell.value
                            is_rich_text = has_rich_text_support and isinstance(cell_value, CellRichText) if has_rich_text_support else False
                            
                            # Debug: Prüfe ob Zelle Formatierung hat
                            has_fill = cell.fill and cell.fill.patternType and cell.fill.patternType != 'none'
                            has_font = cell.font and (cell.font.bold or cell.font.italic or cell.font.color)
                            if has_fill or has_font:
                                styles_found += 1
                            
                            row_info[col_idx] = {
                                'value': cell_value,
                                'is_rich_text': is_rich_text,
                                'fill': copy(cell.fill) if cell.fill else None,
                                'font': copy(cell.font) if cell.font else None,
                                'alignment': copy(cell.alignment) if cell.alignment else None,
                                'border': copy(cell.border) if cell.border else None,
                                'number_format': cell.number_format,
                                'hyperlink': original_hyperlinks.get(excel_row, {}).get(col_idx)
                            }
                        row_data_with_styles[orig_data_idx] = row_info
                    
                    # Schreibe die Zeilen in neuer Reihenfolge
                    # Speichere RichText und Hyperlinks für später (werden nach SCHRITT 4 angewendet)
                    rich_text_cells_to_restore = {}  # Key: "excel_row-col_idx", Value: CellRichText
                    hyperlinks_to_restore = {}  # Key: "excel_row-col_idx", Value: hyperlink target
                    
                    styles_applied = 0
                    for new_pos, orig_row_idx in enumerate(row_mapping):
                        excel_row = new_pos + 2  # Zielzeile
                        if orig_row_idx in row_data_with_styles:
                            row_info = row_data_with_styles[orig_row_idx]
                            for col_idx, cell_info in row_info.items():
                                cell = ws.cell(row=excel_row, column=col_idx)
                                if isinstance(cell, MergedCell):
                                    continue
                                # Formatierungen anwenden (Wert wird später durch data[] überschrieben)
                                # WICHTIG: Immer kopieren, auch wenn "leer" - sonst gehen Defaults verloren
                                if cell_info.get('fill'):
                                    cell.fill = cell_info['fill']
                                    styles_applied += 1
                                if cell_info.get('font'):
                                    cell.font = cell_info['font']
                                    styles_applied += 1
                                if cell_info.get('alignment'):
                                    cell.alignment = cell_info['alignment']
                                if cell_info.get('border'):
                                    cell.border = cell_info['border']
                                # number_format: Immer setzen wenn vorhanden (auch 'General')
                                if cell_info.get('number_format'):
                                    cell.number_format = cell_info['number_format']
                                # RichText für später speichern (wird nach data[] Schreiben angewendet)
                                if cell_info.get('is_rich_text') and cell_info.get('value') is not None:
                                    rich_text_cells_to_restore[f"{excel_row}-{col_idx}"] = cell_info['value']
                                # Hyperlink für später speichern
                                if cell_info.get('hyperlink'):
                                    hyperlinks_to_restore[f"{excel_row}-{col_idx}"] = cell_info['hyperlink']
                    
                    # CF-Bereiche anpassen für gelöschte Zeilen
                    adjust_cf_for_row_changes(ws, row_mapping, original_max_row - 1)  # -1 für Header
            
            # ================================================================
            # SCHRITT 0.6: MERGED CELLS ANPASSEN (bei row_mapping)
            # Wenn Zeilen gelöscht/verschoben wurden, müssen Merged Cells angepasst werden
            # ================================================================
            if row_mapping and len(row_mapping) > 0:
                # Erstelle inverses Mapping: original_row -> new_row (oder None wenn gelöscht)
                # row_mapping[new_pos] = orig_data_idx
                orig_to_new = {}
                for new_pos, orig_data_idx in enumerate(row_mapping):
                    # orig_data_idx ist 0-basiert (Datenzeile), Excel-Zeile = orig_data_idx + 2
                    orig_excel_row = orig_data_idx + 2
                    new_excel_row = new_pos + 2
                    orig_to_new[orig_excel_row] = new_excel_row
                
                # Sammle alle Merged Cells und entferne sie
                merged_ranges_to_update = []
                for merged_range in list(ws.merged_cells.ranges):
                    # Nur Merged Cells im Datenbereich (Zeile >= 2) verarbeiten
                    if merged_range.min_row >= 2:
                        merged_ranges_to_update.append({
                            'min_row': merged_range.min_row,
                            'max_row': merged_range.max_row,
                            'min_col': merged_range.min_col,
                            'max_col': merged_range.max_col
                        })
                        try:
                            ws.unmerge_cells(str(merged_range))
                        except Exception:
                            pass
                
                # Füge Merged Cells mit neuen Positionen wieder hinzu
                final_max_data_row = len(row_mapping) + 1  # +1 für Header
                for merge_info in merged_ranges_to_update:
                    old_min_row = merge_info['min_row']
                    old_max_row = merge_info['max_row']
                    
                    # Finde neue Positionen für alle Zeilen des Merge-Bereichs
                    new_min_row = orig_to_new.get(old_min_row)
                    new_max_row = orig_to_new.get(old_max_row)
                    
                    # Nur wenn beide Zeilen noch existieren und im gültigen Bereich sind
                    if new_min_row is not None and new_max_row is not None:
                        if new_min_row <= final_max_data_row and new_max_row <= final_max_data_row:
                            # Prüfe ob alle Zeilen im Bereich noch zusammenhängend sind
                            all_rows_valid = True
                            expected_new_rows = []
                            for old_row in range(old_min_row, old_max_row + 1):
                                new_row = orig_to_new.get(old_row)
                                if new_row is None:
                                    all_rows_valid = False
                                    break
                                expected_new_rows.append(new_row)
                            
                            if all_rows_valid and expected_new_rows:
                                # Prüfe ob die neuen Zeilen zusammenhängend sind
                                expected_new_rows.sort()
                                is_contiguous = True
                                for i in range(1, len(expected_new_rows)):
                                    if expected_new_rows[i] != expected_new_rows[i-1] + 1:
                                        is_contiguous = False
                                        break
                                
                                if is_contiguous:
                                    actual_new_min = expected_new_rows[0]
                                    actual_new_max = expected_new_rows[-1]
                                    try:
                                        ws.merge_cells(
                                            start_row=actual_new_min,
                                            start_column=merge_info['min_col'],
                                            end_row=actual_new_max,
                                            end_column=merge_info['max_col']
                                        )
                                    except Exception:
                                        pass
            
            # ================================================================
            # SCHRITT 1: SPALTEN EINFÜGEN
            # WICHTIG: openpyxl verschiebt NICHT automatisch Formatierungen!
            # Wir müssen das manuell machen.
            # ================================================================
            if inserted_columns:
                operations = inserted_columns.get('operations', [])
                if not operations and inserted_columns.get('position') is not None:
                    operations = [{
                        'position': inserted_columns['position'],
                        'count': inserted_columns.get('count', 1)
                    }]
                
                # Sortiere aufsteigend (von vorne nach hinten)
                operations.sort(key=lambda x: x['position'])
                
                # Akkumulierter Offset für bereits eingefügte Spalten
                inserted_offset = 0
                
                for op in operations:
                    position = op['position']
                    count = op.get('count', 1)
                    source_column = op.get('sourceColumn')  # Referenzspalte für Formatierung
                    
                    # Position und sourceColumn um bereits eingefügte Spalten anpassen
                    excel_col = position + 1 + inserted_offset  # 0-basiert → 1-basiert + Offset
                    
                    
                    
                    # FÜR JEDE NEUE SPALTE einzeln:
                    for i in range(count):
                        insert_at = excel_col + i
                        
                        # 0. FORMATIERUNG DER REFERENZSPALTE SPEICHERN (falls vorhanden)
                        source_format = {}
                        source_width = None
                        if source_column is not None:
                            # sourceColumn auch um Offset anpassen!
                            source_excel_col = source_column + 1 + inserted_offset
                            col_letter = get_column_letter(source_excel_col)
                            if col_letter in ws.column_dimensions:
                                source_width = ws.column_dimensions[col_letter].width
                            
                            # Alle Zeilen der Referenzspalte speichern
                            for row in range(1, ws.max_row + 1):
                                cell = ws.cell(row=row, column=source_excel_col)
                                source_format[row] = {
                                    'fill': copy(cell.fill) if cell.fill else None,
                                    'font': copy(cell.font) if cell.font else None,
                                    'alignment': copy(cell.alignment) if cell.alignment else None,
                                    'border': copy(cell.border) if cell.border else None,
                                    'number_format': cell.number_format
                                }
                        
                        # 1. SPALTENBREITEN SPEICHERN (OPTIMIERT: nur Breiten)
                        # Die Zellenformate werden von openpyxl beim insert_cols beibehalten
                        # für die bestehenden Zellen. Wir verschieben nur die Breiten.
                        saved_widths = {}
                        max_col = ws.max_column
                        
                        for col in range(insert_at, max_col + 1):
                            col_letter = get_column_letter(col)
                            if col_letter in ws.column_dimensions:
                                saved_widths[col] = ws.column_dimensions[col_letter].width
                        
                        # 2. SPALTE EINFÜGEN
                        ws.insert_cols(insert_at, 1)
                        
                        # 3. SPALTENBREITEN WIEDERHERSTELLEN (um 1 nach rechts verschoben)
                        for old_col, width in saved_widths.items():
                            if width:
                                new_letter = get_column_letter(old_col + 1)
                                ws.column_dimensions[new_letter].width = width
                        
                        
                        # 4. CONDITIONAL FORMATTING ANPASSEN
                        # openpyxl verschiebt CF-Bereiche NICHT automatisch!
                        inserted_cols_for_cf = {insert_at - 1: 1}  # 0-basiert für die Funktion
                        adjust_conditional_formatting(ws, [], inserted_cols_for_cf)
                        
                        # 5. TABLES ANPASSEN (inkl. Table Columns)
                        # openpyxl verschiebt Table-Ranges NICHT automatisch!
                        adjust_tables(ws, [], inserted_cols_for_cf, headers)
                        
                        # 6. FORMATIERUNG DER REFERENZSPALTE AUF NEUE SPALTE ANWENDEN
                        if source_format:
                            # Spaltenbreite
                            if source_width:
                                new_letter = get_column_letter(insert_at)
                                ws.column_dimensions[new_letter].width = source_width
                            
                            # Zellenformatierung (überspringe Header-Zeile 1, damit der neue Header-Name erhalten bleibt)
                            for row, fmt in source_format.items():
                                cell = ws.cell(row=row, column=insert_at)
                                if fmt['fill']:
                                    cell.fill = fmt['fill']
                                if fmt['font']:
                                    cell.font = fmt['font']
                                if fmt['alignment']:
                                    cell.alignment = fmt['alignment']
                                if fmt['border']:
                                    cell.border = fmt['border']
                                if fmt.get('number_format'):
                                    cell.number_format = fmt['number_format']
                    
                    # Offset für nächste Operation erhöhen
                    inserted_offset += count
                            
            
            # ================================================================
            # SCHRITT 2: SPALTEN LÖSCHEN
            # WICHTIG: openpyxl verschiebt Zellformate NICHT automatisch!
            # Wir müssen Spaltenbreiten manuell verschieben.
            # Die Zellformate werden aber korrekt verschoben wenn wir die Zellen
            # NACH dem delete_cols neu schreiben (was in SCHRITT 3+4 passiert).
            # ================================================================
            if deleted_columns:
                # Sortiere absteigend (von hinten nach vorne löschen)
                sorted_deleted = sorted(deleted_columns, reverse=True)
                for col_idx in sorted_deleted:
                    excel_col = col_idx + 1  # 0-basiert → 1-basiert
                    
                    max_col = ws.max_column
                    
                    # 1. SPALTENBREITEN SPEICHERN
                    saved_widths = {}
                    for col in range(excel_col + 1, max_col + 1):
                        col_letter = get_column_letter(col)
                        if col_letter in ws.column_dimensions:
                            saved_widths[col] = ws.column_dimensions[col_letter].width
                    
                    # 2. SPALTE LÖSCHEN
                    ws.delete_cols(excel_col, 1)
                    
                    # 3. SPALTENBREITEN WIEDERHERSTELLEN (um 1 nach links verschoben)
                    for old_col, width in saved_widths.items():
                        if width:
                            new_letter = get_column_letter(old_col - 1)
                            ws.column_dimensions[new_letter].width = width
                    
                    # 4. CONDITIONAL FORMATTING ANPASSEN
                    adjust_conditional_formatting(ws, [col_idx], None)
                    
                    # 5. TABLES ANPASSEN (mit headers für korrekte Column-Namen)
                    adjust_tables(ws, [col_idx], None, headers)
            
            # ================================================================
            # SCHRITT 3: HEADER SCHREIBEN (Werte)
            # ================================================================
            from openpyxl.cell.cell import MergedCell
            for col_idx, header in enumerate(headers):
                cell = ws.cell(row=1, column=col_idx + 1)
                if not isinstance(cell, MergedCell):
                    cell.value = header
            
            # ================================================================
            # SCHRITT 3.5: RICHTEXT UND HYPERLINKS VOR DEM SCHREIBEN SAMMELN
            # Wenn kein row_mapping existiert, müssen wir trotzdem RichText
            # und Hyperlinks sammeln, da SCHRITT 4 alle Werte überschreibt
            # ================================================================
            try:
                # Prüfe ob rich_text_cells_to_restore bereits existiert (von SCHRITT 0.5)
                _ = rich_text_cells_to_restore
            except NameError:
                # Kein row_mapping - sammle RichText und Hyperlinks jetzt
                try:
                    from openpyxl.cell.rich_text import CellRichText
                    has_rich_text_support = True
                except ImportError:
                    has_rich_text_support = False
                
                rich_text_cells_to_restore = {}
                hyperlinks_to_restore = {}
                
                # Sammle RichText und Hyperlinks von allen Datenzellen
                for row_idx in range(len(data)):
                    excel_row = row_idx + 2  # +2: Excel 1-basiert + Header
                    for col_idx in range(1, len(headers) + 1):
                        cell = ws.cell(row=excel_row, column=col_idx)
                        if isinstance(cell, MergedCell):
                            continue
                        
                        # RichText prüfen
                        if has_rich_text_support and isinstance(cell.value, CellRichText):
                            rich_text_cells_to_restore[f"{excel_row}-{col_idx}"] = cell.value
                        
                        # Hyperlink prüfen
                        if cell.hyperlink and cell.hyperlink.target:
                            hyperlinks_to_restore[f"{excel_row}-{col_idx}"] = cell.hyperlink.target
            
            # ================================================================
            # SCHRITT 4: DATEN SCHREIBEN (Werte)
            # ================================================================
            for row_idx, row_data in enumerate(data):
                excel_row = row_idx + 2  # +2 für Header (1-basiert)
                for col_idx, value in enumerate(row_data):
                    cell = ws.cell(row=excel_row, column=col_idx + 1)
                    apply_cell_value(cell, value)
            
            # ================================================================
            # SCHRITT 4.5: RICHTEXT UND HYPERLINKS WIEDERHERSTELLEN
            # Diese wurden in SCHRITT 0.5 gespeichert und müssen nach dem
            # Schreiben der Daten wiederhergestellt werden
            # ================================================================
            from openpyxl.cell.cell import MergedCell
            
            # Stelle RichText wieder her (falls vorhanden)
            try:
                if rich_text_cells_to_restore:
                    for key, rich_text_value in rich_text_cells_to_restore.items():
                        parts = key.split('-')
                        excel_row = int(parts[0])
                        col_idx = int(parts[1])
                        try:
                            cell = ws.cell(row=excel_row, column=col_idx)
                            if not isinstance(cell, MergedCell):
                                cell.value = rich_text_value
                        except Exception:
                            pass
            except NameError:
                pass  # Variable nicht definiert (kein row_mapping)
            
            # Stelle Hyperlinks wieder her (falls vorhanden)
            try:
                if hyperlinks_to_restore:
                    for key, hyperlink_target in hyperlinks_to_restore.items():
                        parts = key.split('-')
                        excel_row = int(parts[0])
                        col_idx = int(parts[1])
                        try:
                            cell = ws.cell(row=excel_row, column=col_idx)
                            if not isinstance(cell, MergedCell):
                                cell.hyperlink = hyperlink_target
                        except Exception:
                            pass
            except NameError:
                pass  # Variable nicht definiert (kein row_mapping)
            
            # ================================================================
            # SCHRITT 5: ÜBERSCHÜSSIGE SPALTEN AM ENDE LÖSCHEN
            # ================================================================
            current_max_col = ws.max_column
            if current_max_col > target_col_count:
                cols_to_delete = current_max_col - target_col_count
                ws.delete_cols(target_col_count + 1, cols_to_delete)
            
            # ================================================================
            # SCHRITT 6: VERSTECKTE SPALTEN
            # ================================================================
            _apply_hidden_columns(ws, hidden_columns, len(headers))
            
            # ================================================================
            # SCHRITT 7: VERSTECKTE ZEILEN
            # ================================================================
            _apply_hidden_rows(ws, hidden_rows, len(data))
            
            # ================================================================
            # SCHRITT 8: ROW HIGHLIGHTS
            # ================================================================
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, len(headers))
            
            # ================================================================
            # SCHRITT 8.5: NUMBER FORMATS UND CELL FONTS (für Data Join)
            # ================================================================
            number_formats = changes.get('numberFormats', {})
            cell_fonts = changes.get('cellFonts', {})
            imported_cell_styles = changes.get('cellStyles', {})
            if number_formats:
                _apply_number_formats(ws, number_formats)
            if cell_fonts:
                _apply_cell_fonts(ws, cell_fonts)
            if imported_cell_styles:
                _apply_imported_cell_styles(ws, imported_cell_styles)
            
            # RichText aus Copy-Paste anwenden
            imported_rich_text = changes.get('richTextCells', {})
            if imported_rich_text:
                _apply_imported_rich_text(ws, imported_rich_text)
            
            # ================================================================
            # SCHRITT 9: CLEARED ROW HIGHLIGHTS (Markierungen entfernen)
            # ================================================================
            if cleared_row_highlights:
                for row_idx in cleared_row_highlights:
                    excel_row = row_idx + 2
                    for col_idx in range(1, len(headers) + 1):
                        cell = ws.cell(row=excel_row, column=col_idx)
                        cell.fill = PatternFill()  # Keine Füllung
            
            # ================================================================
            # SCHRITT 9.5: ÜBERSCHÜSSIGE ZEILEN UND MERGED CELLS ENTFERNEN
            # Wenn Zeilen gelöscht wurden, kann die Datei mehr Zeilen haben als
            # wir jetzt Daten haben. Diese müssen entfernt werden.
            # ================================================================
            final_data_row_count = len(data)  # Anzahl der Datenzeilen (ohne Header)
            final_max_row = final_data_row_count + 1  # +1 für Header
            
            # Entferne Merged Cells die außerhalb des neuen Datenbereichs liegen
            merged_to_remove = []
            for merged_range in list(ws.merged_cells.ranges):
                # Wenn die Merged Range außerhalb des neuen Bereichs liegt
                if merged_range.min_row > final_max_row:
                    merged_to_remove.append(str(merged_range))
                # Wenn die Range teilweise außerhalb liegt, auch entfernen
                elif merged_range.max_row > final_max_row:
                    merged_to_remove.append(str(merged_range))
            
            for range_str in merged_to_remove:
                try:
                    ws.unmerge_cells(range_str)
                except Exception:
                    pass
            
            # Leere überschüssige Zeilen (NICHT löschen - ws.delete_rows() beschädigt die Datei!)
            # Stattdessen: Zellen leeren und Formatierung entfernen
            if original_max_row > final_max_row:
                for row in range(final_max_row + 1, original_max_row + 1):
                    for col in range(1, original_max_col + 1):
                        try:
                            cell = ws.cell(row=row, column=col)
                            cell.value = None
                            cell.fill = PatternFill()  # Keine Füllung
                            cell.border = Border()     # Kein Rahmen
                        except Exception:
                            pass
            
            # ================================================================
            # SCHRITT 10: AUTOFILTER SETZEN
            # ================================================================
            af_source = frontend_auto_filter or original_auto_filter
            if af_source:
                try:
                    final_max_row = len(data) + 1  # +1 für Header
                    final_af_ref = f"A1:{get_column_letter(target_col_count)}{final_max_row}"
                    ws.auto_filter.ref = final_af_ref
                except Exception as e:
                    pass
            
            # ================================================================
            # SCHRITT 10.5: MERGED CELLS AUS FRONTEND ANWENDEN
            # Überschreibt alle Merges mit dem vollständigen GUI-Zustand
            # ================================================================
            imported_merged_cells = changes.get('mergedCells', [])
            if imported_merged_cells:
                _apply_imported_merged_cells(ws, imported_merged_cells)
            
            # ================================================================
            # SCHRITT 11: SAMMLE TABLE-INFOS FÜR RESTORE
            # ================================================================
            table_changes = {}
            for table_name in ws.tables:
                table = ws.tables[table_name]
                col_names = [col.name for col in table.tableColumns]
                table_changes[table_name] = {
                    'ref': table.ref,
                    'columns': col_names
                }
            
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # Stelle Original-Table-XML wieder her (mit korrekten xr:uid etc.)
            # WICHTIG: Bei Spalten-INSERT NICHT aufrufen - openpyxl erzeugt saubere XML
            # Bei Spalten-DELETE hingegen schon, um xr:uid/xr3:uid zu erhalten
            if table_changes and not inserted_columns:
                restore_table_xml_from_original(output_path, original_path, table_changes)
            elif table_changes and inserted_columns:
                pass  # Bei INSERT keine XML-Wiederherstellung nötig
            
            # Stelle externalLinks aus Original wieder her (openpyxl verliert Namespaces)
            restore_external_links_from_original(output_path, original_path)
            
            return {'success': True, 'outputPath': output_path, 'method': 'openpyxl'}
        
        # =====================================================================
        # FALL 3: Nur Zell-Edits (keine strukturellen Änderungen)
        # =====================================================================
        
        # Prüfe ob wir echte Zell-Edits haben (nicht nur Highlights)
        real_edits = {k: v for k, v in edited_cells.items() if not k.startswith('_')} if edited_cells else {}
        
        # Prüfe ob zusätzliche Änderungen neben reinen Zell-Edits vorliegen
        imported_cell_styles = changes.get('cellStyles', {})
        cell_fonts = changes.get('cellFonts', {})
        imported_rich_text = changes.get('richTextCells', {})
        imported_merged_cells = changes.get('mergedCells', [])
        
        # WICHTIG: FALL 3a (direkte XML-Bearbeitung) kopiert die Original-Datei 1:1
        # und ändert NUR Zellwerte. Alle Styles, Fonts, MergedCells, Highlights,
        # Drawings etc. bleiben aus dem Original erhalten.
        #
        # Das Frontend sendet IMMER die Original-Daten mit:
        # - cellStyles: Original-Styles der editierten Zellen (zum Wiederherstellen in openpyxl)
        # - cellFonts: Original-Fonts der editierten Zellen
        # - richTextCells: Original-RichText der editierten Zellen
        # - mergedCells: ALLE pre-existierenden Merged Cells
        # - rowHighlights: ALLE aktuellen Highlights (auch pre-existierende)
        #
        # Diese Daten sind KEINE neuen Änderungen - sie sind pre-existierend!
        # FALL 3a preserviert sie automatisch aus dem Original.
        #
        # NUR wenn das Frontend _hasFormatChanges gesetzt hat (Paste-mit-Format,
        # Data Join Styles, oder Highlight-Fill-Löschung), sind echte
        # Formatierungsänderungen vorhanden → FALL 3b nötig.
        has_format_flag = edited_cells.get('_hasFormatChanges', False) if edited_cells else False
        has_extra_changes = has_format_flag
        has_highlight_changes = bool(cleared_row_highlights)  # Nur wenn Highlights explizit entfernt
        
        sys.stderr.write(f"[GATE] _hasFormatChanges={has_format_flag}, has_extra_changes={has_extra_changes}, has_highlight_changes={has_highlight_changes}\n")
        sys.stderr.write(f"[GATE] real_edits={len(real_edits)}, cellStyles={len(imported_cell_styles)}, mergedCells={len(imported_merged_cells)}\n")
        
        # =====================================================================
        # FALL 3a: NUR Zell-Edits → Direkte XML-Bearbeitung (kein openpyxl-Roundtrip)
        # Dies vermeidet das Überschreiben von Rels, Namespaces, SharedStrings etc.
        # Die Original-Datei bleibt zu 100% intakt, nur die Zellwerte werden geändert.
        # =====================================================================
        if real_edits and not has_extra_changes and not has_highlight_changes:
            wb.close()  # openpyxl Workbook nicht mehr benötigt
            sys.stderr.write(f"[FALL 3a] Direkte XML-Bearbeitung für {len(real_edits)} Zell-Edits\n")
            
            try:
                result = _direct_xml_cell_edit(
                    file_path, output_path, sheet_name, real_edits,
                    hidden_columns, hidden_rows
                )
                return result
            except Exception as xml_err:
                sys.stderr.write(f"[FALL 3a] Fehler bei direkter XML-Bearbeitung: {xml_err}\n")
                sys.stderr.write(f"[FALL 3a] Fallback auf openpyxl-Pfad...\n")
                # Fallback: openpyxl-Pfad (FALL 3b)
                wb = load_workbook(file_path, rich_text=True)
                ws = wb[sheet_name]
        
        # =====================================================================
        # FALL 3b: Zell-Edits MIT zusätzlichen Änderungen (Highlights, Styles, etc.)
        # Hier muss openpyxl verwendet werden, da XML-Bearbeitung zu komplex wäre.
        # =====================================================================
        
        # Wenn NUR Highlights (keine echten Edits), lade von Original-Datei neu (falls verfügbar)
        # Das stellt sicher dass alte Highlights nicht erhalten bleiben
        if row_highlights is not None and not real_edits:
            if original_path and original_path != file_path and os.path.exists(original_path):
                wb.close()
                import shutil
                shutil.copy2(original_path, output_path)
                wb = load_workbook(output_path, rich_text=True)
                ws = wb[sheet_name]
            else:
                # Kein Original verfügbar - entferne alle Fills in Zeilen die NICHT markiert sind
                # Das ist nicht perfekt (verliert Zebra-Muster), aber besser als alte Highlights zu behalten
                _clear_all_row_fills_except(ws, row_highlights)
        
        if real_edits:
            for key, value in real_edits.items():
                parts = key.split('-')
                if len(parts) != 2:
                    continue
                row_idx = int(parts[0])
                col_idx = int(parts[1])
                cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
                apply_cell_value(cell, value)
        
        # Versteckte Spalten/Zeilen setzen
        _apply_hidden_columns(ws, hidden_columns)
        _apply_hidden_rows(ws, hidden_rows)
        
        # Kopierte Zell-Hintergründe anwenden (aus Copy-Paste)
        if imported_cell_styles:
            _apply_imported_cell_styles(ws, imported_cell_styles)
        
        # Kopierte Schriftformatierungen anwenden (aus Copy-Paste)
        if cell_fonts:
            _apply_cell_fonts(ws, cell_fonts)
        
        # Kopierte RichText-Formatierung anwenden (aus Copy-Paste)
        if imported_rich_text:
            _apply_imported_rich_text(ws, imported_rich_text)
        
        # Kopierte Merged Cells anwenden (aus Copy-Paste)
        if imported_merged_cells:
            _apply_imported_merged_cells(ws, imported_merged_cells)
        
        # Row Highlights
        if row_highlights:
            _apply_row_highlights(ws, row_highlights, ws.max_column)
        
        # Cleared Row Highlights (Markierungen entfernen)
        if cleared_row_highlights:
            sys.stderr.write(f"[FALL 3b] Entferne {len(cleared_row_highlights)} Row Highlights\n")
            for row_idx in cleared_row_highlights:
                excel_row = row_idx + 2  # 0-basiert nach 1-basiert + Header
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    cell.fill = PatternFill()  # Keine Füllung
        
        wb.save(output_path)
        wb.close()
        fix_xlsx_relationships(output_path)
        
        # WICHTIG: Table-XML vom Original wiederherstellen!
        restore_table_xml_from_original(output_path, original_path, table_changes=None)
        
        # WICHTIG: Auch workbook.xml, slicerCaches, etc. vom Original wiederherstellen!
        restore_external_links_from_original(output_path, original_path)
        
        return {'success': True, 'outputPath': output_path}
        
    except Exception as e:
        import traceback
        error_msg = str(e)
        tb = traceback.format_exc()
        print(f"[Python Writer] ERROR: {error_msg}", file=sys.stderr)
        print(f"[Python Writer] Traceback: {tb}", file=sys.stderr)
        return {
            'success': False, 
            'error': error_msg,
            'traceback': tb
        }
    finally:
        # Backup-Datei aufräumen (erstellt für Speichern-in-gleicher-Datei Szenario)
        if _backup_file is not None:
            try:
                os.unlink(_backup_path)
                sys.stderr.write(f"[WRITE_SHEET] Backup gelöscht: {_backup_path}\\n")
            except Exception:
                pass
            # Funktionsattribut zurücksetzen
            restore_external_links_from_original._backup_original_path = None


def _apply_hidden_columns(ws, hidden_columns, max_cols=None):
    """Setzt versteckte Spalten"""
    if hidden_columns is None:
        return
    
    hidden_set = set(hidden_columns)
    max_col = max_cols if max_cols else ws.max_column
    
    for col_idx in range(max_col):
        col_letter = get_column_letter(col_idx + 1)
        ws.column_dimensions[col_letter].hidden = col_idx in hidden_set


def _apply_hidden_rows(ws, hidden_rows, max_rows=None):
    """Setzt versteckte Zeilen"""
    if hidden_rows is None:
        return
    
    hidden_set = set(hidden_rows)
    max_row = max_rows if max_rows else (ws.max_row - 1)  # Ohne Header
    
    for row_idx in range(max_row):
        excel_row = row_idx + 2  # +2 für Header
        ws.row_dimensions[excel_row].hidden = row_idx in hidden_set


def _clear_all_row_fills_except(ws, row_highlights):
    """
    Entfernt Fills von allen Zeilen AUSSER den in row_highlights angegebenen.
    Wird verwendet wenn kein Original verfügbar ist und Highlights entfernt werden sollen.
    """
    # Sammle die Zeilen die Highlights behalten sollen
    highlighted_rows = set()
    if row_highlights:
        for row_idx_str in row_highlights.keys():
            highlighted_rows.add(int(row_idx_str) + 2)  # +2 für Excel-Row (1-basiert + Header)
    
    # Durchgehe alle Datenzeilen und entferne Fills die nicht Highlights sind
    max_row = ws.max_row


def _apply_number_formats(ws, number_formats):
    """Wendet Zahlenformate aus dem Frontend auf Zellen an"""
    if not number_formats:
        return
    
    for key, fmt in number_formats.items():
        try:
            parts = key.split('-')
            if len(parts) != 2:
                continue
            row_idx = int(parts[0])
            col_idx = int(parts[1])
            cell = ws.cell(row=row_idx + 2, column=col_idx + 1)  # +2 für Header, +1 für 1-basiert
            cell.number_format = fmt
        except Exception:
            pass


def _apply_cell_fonts(ws, cell_fonts):
    """Wendet Font-Formatierungen aus dem Frontend auf Zellen an"""
    if not cell_fonts:
        return
    
    sys.stderr.write(f"[CellFonts] Starte: {len(cell_fonts)} Font-Einträge zu verarbeiten\n")
    applied = 0
    
    for key, font_info in cell_fonts.items():
        try:
            parts = key.split('-')
            if len(parts) != 2:
                continue
            row_idx = int(parts[0])
            col_idx = int(parts[1])
            cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
            
            # Bestehende Font-Eigenschaften beibehalten und nur überschreiben was wir haben
            existing_font = cell.font
            font_kwargs = {
                'name': existing_font.name,
                'size': existing_font.size,
                'bold': existing_font.bold,
                'italic': existing_font.italic,
                'color': existing_font.color,
                'underline': existing_font.underline,
                'strikethrough': existing_font.strikethrough,
            }
            
            # Neue Werte überschreiben
            if font_info.get('name'):
                font_kwargs['name'] = font_info['name']
            if font_info.get('size'):
                font_kwargs['size'] = font_info['size']
            if 'bold' in font_info:
                font_kwargs['bold'] = font_info['bold']
            if 'italic' in font_info:
                font_kwargs['italic'] = font_info['italic']
            if font_info.get('color'):
                color_val = font_info['color']
                # #RRGGBB → RRGGBB konvertieren (openpyxl erwartet kein #)
                if isinstance(color_val, str) and color_val.startswith('#'):
                    color_val = color_val[1:]
                # Zu ARGB konvertieren wenn nötig
                if isinstance(color_val, str) and len(color_val) == 6:
                    color_val = 'FF' + color_val
                font_kwargs['color'] = color_val
            
            cell.font = Font(**font_kwargs)
            applied += 1
        except Exception as e:
            sys.stderr.write(f"[CellFonts] Fehler bei {key}: {e}\n")
    
    sys.stderr.write(f"[CellFonts] Fertig: {applied} von {len(cell_fonts)} Fonts angewendet\n")


def _apply_imported_cell_styles(ws, cell_styles):
    """
    Wendet KOMPLETTE Zell-Formatierung aus dem Frontend auf Zellen an.
    Unterstützt sowohl altes Format (String = nur Hintergrundfarbe)
    als auch neues Format (Object = komplette Formatierung inkl. Font, Fill, Alignment).
    
    Neues Format-Objekt kann enthalten:
    - fill: Hintergrundfarbe (#RRGGBB oder ARGB)
    - bold, italic, underline, strikethrough: Font-Eigenschaften
    - fontSize, fontName, fontColor: Font-Details
    - textAlign: Horizontale Ausrichtung (left, center, right)
    - verticalAlign: Vertikale Ausrichtung (top, center, bottom)
    - wrapText: Textumbruch (true/false)
    """
    if not cell_styles:
        return
    
    applied = 0
    for key, style_data in cell_styles.items():
        try:
            parts = key.split('-')
            if len(parts) != 2:
                continue
            row_idx = int(parts[0])
            col_idx = int(parts[1])
            cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
            
            # Altes Format: String = nur Hintergrundfarbe
            if isinstance(style_data, str):
                color = style_data
                if color.startswith('#'):
                    argb = hex_to_argb(color)
                else:
                    argb = color if len(color) == 8 else f'FF{color}'
                cell.fill = PatternFill(start_color=argb, end_color=argb, fill_type='solid')
                applied += 1
                continue
            
            # Neues Format: Komplettes Style-Objekt
            if not isinstance(style_data, dict):
                continue
            
            # === FILL (Hintergrundfarbe) ===
            fill_color = style_data.get('fill')
            if fill_color and fill_color != 'transparent':
                if isinstance(fill_color, str):
                    if fill_color.startswith('#'):
                        argb = hex_to_argb(fill_color)
                    else:
                        argb = fill_color if len(fill_color) == 8 else f'FF{fill_color}'
                    cell.fill = PatternFill(start_color=argb, end_color=argb, fill_type='solid')
            
            # === FONT (Schriftformatierung) ===
            has_font_info = any(style_data.get(k) is not None for k in 
                              ['bold', 'italic', 'underline', 'strikethrough', 'fontSize', 'fontName', 'fontColor'])
            if has_font_info:
                # Bestehende Font-Eigenschaften als Basis verwenden
                existing_font = cell.font
                font_kwargs = {
                    'name': existing_font.name,
                    'size': existing_font.size,
                    'bold': existing_font.bold,
                    'italic': existing_font.italic,
                    'underline': existing_font.underline,
                    'strikethrough': existing_font.strikethrough,
                    'color': existing_font.color,
                }
                
                if style_data.get('fontName'):
                    font_kwargs['name'] = style_data['fontName']
                if style_data.get('fontSize'):
                    font_kwargs['size'] = style_data['fontSize']
                if 'bold' in style_data:
                    font_kwargs['bold'] = bool(style_data['bold'])
                if 'italic' in style_data:
                    font_kwargs['italic'] = bool(style_data['italic'])
                if 'underline' in style_data:
                    font_kwargs['underline'] = 'single' if style_data['underline'] else None
                if 'strikethrough' in style_data:
                    font_kwargs['strikethrough'] = bool(style_data['strikethrough'])
                if style_data.get('fontColor'):
                    color_val = style_data['fontColor']
                    if isinstance(color_val, str) and color_val.startswith('#'):
                        color_val = color_val[1:]
                    if isinstance(color_val, str) and len(color_val) == 6:
                        color_val = 'FF' + color_val
                    font_kwargs['color'] = color_val
                
                cell.font = Font(**font_kwargs)
            
            # === ALIGNMENT (Ausrichtung) ===
            has_alignment = any(style_data.get(k) is not None for k in 
                              ['textAlign', 'verticalAlign', 'wrapText'])
            if has_alignment:
                align_kwargs = {}
                if style_data.get('textAlign'):
                    align_kwargs['horizontal'] = style_data['textAlign']
                if style_data.get('verticalAlign'):
                    align_kwargs['vertical'] = style_data['verticalAlign']
                if 'wrapText' in style_data:
                    align_kwargs['wrap_text'] = bool(style_data['wrapText'])
                if align_kwargs:
                    cell.alignment = Alignment(**align_kwargs)
            
            applied += 1
        except Exception as e:
            sys.stderr.write(f"[CellStyles] Fehler bei {key}: {e}\n")
    
    if applied > 0:
        sys.stderr.write(f"[CellStyles] {applied} von {len(cell_styles)} Zell-Styles angewendet\n")


def _apply_imported_merged_cells(ws, merged_cells_list):
    """
    Wendet Merged Cells aus dem Frontend auf das Worksheet an.
    Verwendet differentiellen Ansatz: nur Änderungen anwenden statt alles neu zu mergen.
    So bleibt die Textformatierung bestehender Merges erhalten.
    
    Frontend-Format: [{ startRow, startCol, endRow, endCol, rowSpan, colSpan }]
    startRow/startCol sind 0-basiert (0 = Excel-Zeile 1 = Header).
    """
    if not merged_cells_list:
        sys.stderr.write(f"[MergedCells] Liste leer - übersprungen\n")
        return
    
    sys.stderr.write(f"[MergedCells] Starte: {len(merged_cells_list)} Merges zu verarbeiten\n")
    
    # Bestehende Merges als String-Set
    existing_ranges = set(str(mr) for mr in ws.merged_cells.ranges)
    sys.stderr.write(f"[MergedCells] Bestehende Merges im Sheet: {len(existing_ranges)}\n")
    for mr_str in list(existing_ranges)[:5]:
        sys.stderr.write(f"[MergedCells] Bestehend: {mr_str}\n")
    
    # Ziel-Merges als String-Set aufbauen
    target_ranges = set()
    for merge in merged_cells_list:
        start_row = merge.get('startRow', 0) + 1  # 0-basiert -> 1-basiert
        start_col = merge.get('startCol', 0) + 1
        end_row = merge.get('endRow', 0) + 1
        end_col = merge.get('endCol', 0) + 1
        
        if start_row > 0 and start_col > 0 and end_row >= start_row and end_col >= start_col:
            cell_range = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
            target_ranges.add(cell_range)
        else:
            sys.stderr.write(f"[MergedCells] Übersprungen (ungültig): ({start_row},{start_col}):({end_row},{end_col})\n")
    
    # Differenz berechnen
    to_remove = existing_ranges - target_ranges  # Merges die entfernt werden müssen
    to_add = target_ranges - existing_ranges      # Neue Merges die hinzugefügt werden
    unchanged = existing_ranges & target_ranges   # Merges die bleiben (nicht anfassen!)
    
    sys.stderr.write(f"[MergedCells] Differenz: {len(unchanged)} unverändert, {len(to_remove)} entfernen, {len(to_add)} hinzufügen\n")
    
    # Nur nicht mehr benötigte Merges entfernen
    for mr_str in to_remove:
        try:
            ws.unmerge_cells(mr_str)
            sys.stderr.write(f"[MergedCells] Entfernt: {mr_str}\n")
        except Exception as e:
            sys.stderr.write(f"[MergedCells] Fehler beim Unmerge {mr_str}: {e}\n")
    
    # Nur neue Merges hinzufügen
    for mr_str in to_add:
        try:
            ws.merge_cells(mr_str)
            sys.stderr.write(f"[MergedCells] Hinzugefügt: {mr_str}\n")
        except Exception as e:
            sys.stderr.write(f"[MergedCells] Fehler beim Mergen {mr_str}: {e}\n")
    
    sys.stderr.write(f"[MergedCells] Fertig: {len(unchanged) + len(to_add)} Merged Cells im Ergebnis\n")


def _apply_imported_rich_text(ws, rich_text_cells):
    """
    Wendet RichText-Formatierung aus dem Frontend auf Zellen an.
    Wird für kopierte RichText-Zellen verwendet.
    
    Frontend-Format: { "row-col": [{ text: "...", styles: { bold, italic, ... } }] }
    Keys sind 0-basiert (wie editedCells).
    """
    if not rich_text_cells:
        return
    
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
    except ImportError:
        sys.stderr.write("[RichText] openpyxl CellRichText nicht verfügbar - überspringe RichText\n")
        return
    
    applied = 0
    for key, parts in rich_text_cells.items():
        try:
            key_parts = key.split('-')
            if len(key_parts) != 2:
                continue
            row_idx = int(key_parts[0])
            col_idx = int(key_parts[1])
            cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
            
            if not isinstance(parts, list) or len(parts) == 0:
                continue
            
            # Konvertiere Frontend-Format zu CellRichText
            text_blocks = []
            for part in parts:
                text = part.get('text', '')
                styles = part.get('styles', {})
                
                font_kwargs = {}
                if styles.get('bold'):
                    font_kwargs['b'] = True
                if styles.get('italic'):
                    font_kwargs['i'] = True
                if styles.get('underline'):
                    font_kwargs['u'] = 'single'
                if styles.get('strikethrough'):
                    font_kwargs['strike'] = True
                if styles.get('fontSize'):
                    font_kwargs['sz'] = styles['fontSize']
                if styles.get('fontName'):
                    font_kwargs['rFont'] = styles['fontName']
                if styles.get('color'):
                    color_val = styles['color']
                    if isinstance(color_val, str) and color_val.startswith('#'):
                        font_kwargs['color'] = color_val[1:]  # Entferne #
                    elif isinstance(color_val, str):
                        font_kwargs['color'] = color_val
                
                if font_kwargs:
                    inline_font = InlineFont(**font_kwargs)
                    text_blocks.append(TextBlock(inline_font, text))
                else:
                    text_blocks.append(text)
            
            cell.value = CellRichText(*text_blocks)
            applied += 1
        except Exception as e:
            sys.stderr.write(f"[RichText] Fehler bei {key}: {e}\n")
    
    if applied > 0:
        sys.stderr.write(f"[RichText] {applied} RichText-Zellen angewendet\n")


def _apply_row_highlights(ws, row_highlights, num_columns):
    """Wendet Zeilen-Highlights an"""
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
        
        if isinstance(color, str) and color.startswith('#'):
            argb = hex_to_argb(color)
        else:
            argb = highlight_colors.get(color, 'FFFFFF00')
        
        # Alle Zellen in der Zeile färben
        for col_idx in range(1, num_columns + 1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.fill = PatternFill(start_color=argb, end_color=argb, fill_type='solid')


def main():
    """Hauptfunktion - liest Befehle von stdin oder Argumenten"""
    # Auf Windows: Stelle sicher dass stdin/stdout UTF-8 verwenden
    import io
    if sys.platform == 'win32':
        sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
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
            params.get('changes', {}),
            params.get('originalPath')  # NEU: Original-Datei für restore_table_xml
        )
        print(json.dumps(result, ensure_ascii=False))
    
    elif command == 'check_excel':
        # Prüft ob Microsoft Excel verfügbar ist
        excel_available = is_excel_installed()
        result = {
            'success': True,
            'excelAvailable': excel_available,
            'xlwingsAvailable': XLWINGS_AVAILABLE,
            'message': 'Excel verfügbar - strukturelle Änderungen mit CF-Erhalt möglich' if excel_available else 'Excel nicht verfügbar - CF-Erhalt bei strukturellen Änderungen eingeschränkt'
        }
        print(json.dumps(result, ensure_ascii=False))
    
    else:
        print(json.dumps({'success': False, 'error': f'Unbekannter Befehl: {command}'}))
        sys.exit(1)


if __name__ == '__main__':
    main()
