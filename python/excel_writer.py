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
                    for f in files:
                        if f == 'fixed.xlsx':
                            continue
                        full_path = os.path.join(root, f)
                        arc_name = full_path.replace(temp_dir + os.sep, '')
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
                    for f in files:
                        if f == 'restored.xlsx':
                            continue
                        full_path = os.path.join(root, f)
                        arc_name = full_path.replace(temp_dir + os.sep, '')
                        zf.write(full_path, arc_name)
            
            shutil.copy2(temp_xlsx, output_path)
    
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        shutil.rmtree(orig_temp_dir, ignore_errors=True)


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
    
    if not original_path or original_path == output_path:
        return
    
    if not os.path.exists(original_path):
        return
    
    temp_dir = None
    orig_temp_dir = None
    
    try:
        temp_dir = tempfile.mkdtemp()
        orig_temp_dir = tempfile.mkdtemp()
        temp_xlsx = os.path.join(temp_dir, 'restored.xlsx')
        
        with zipfile.ZipFile(output_path, 'r') as zf:
            zf.extractall(temp_dir)
        with zipfile.ZipFile(original_path, 'r') as zf:
            zf.extractall(orig_temp_dir)
        
        ext_links_dir = os.path.join(temp_dir, 'xl', 'externalLinks')
        orig_ext_links_dir = os.path.join(orig_temp_dir, 'xl', 'externalLinks')
        
        fixed_count = 0
        
        if os.path.exists(orig_ext_links_dir) and os.path.exists(ext_links_dir):
            # Kopiere alle externalLink*.xml Dateien
            for f in os.listdir(orig_ext_links_dir):
                if f.startswith('externalLink') and f.endswith('.xml'):
                    orig_file = os.path.join(orig_ext_links_dir, f)
                    dest_file = os.path.join(ext_links_dir, f)
                    if os.path.exists(dest_file):
                        shutil.copy2(orig_file, dest_file)
                        fixed_count += 1
            
            # WICHTIG: Auch die _rels Dateien kopieren (openpyxl verliert Relationships)
            orig_rels_dir = os.path.join(orig_ext_links_dir, '_rels')
            dest_rels_dir = os.path.join(ext_links_dir, '_rels')
            if os.path.exists(orig_rels_dir) and os.path.exists(dest_rels_dir):
                for f in os.listdir(orig_rels_dir):
                    if f.endswith('.xml.rels'):
                        orig_rels = os.path.join(orig_rels_dir, f)
                        dest_rels = os.path.join(dest_rels_dir, f)
                        if os.path.exists(dest_rels):
                            shutil.copy2(orig_rels, dest_rels)
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
        
        # Kopiere sharedStrings.xml (Original verwendet shared strings, openpyxl inline strings)
        orig_shared_strings = os.path.join(orig_temp_dir, 'xl', 'sharedStrings.xml')
        dest_shared_strings = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
        if os.path.exists(orig_shared_strings):
            shutil.copy2(orig_shared_strings, dest_shared_strings)
            fixed_count += 1
        
        # Stelle workbook.xml aus Original wieder her (behält definedNames, externalReferences, slicerCaches-Refs)
        workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
        orig_workbook_path = os.path.join(orig_temp_dir, 'xl', 'workbook.xml')
        
        if os.path.exists(workbook_path) and os.path.exists(orig_workbook_path):
            # Kopiere komplett das Original workbook.xml
            shutil.copy2(orig_workbook_path, workbook_path)
            fixed_count += 1
        
        # Stelle workbook.xml.rels aus Original wieder her (enthält slicerCache Referenzen)
        rels_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        orig_rels_path = os.path.join(orig_temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        if os.path.exists(orig_rels_path):
            shutil.copy2(orig_rels_path, rels_path)
            fixed_count += 1
        
        # Stelle [Content_Types].xml aus Original wieder her (enthält slicerCache ContentTypes)
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
        orig_content_types_path = os.path.join(orig_temp_dir, '[Content_Types].xml')
        if os.path.exists(orig_content_types_path):
            shutil.copy2(orig_content_types_path, content_types_path)
            fixed_count += 1
        
        if fixed_count > 0:
            
            # Erstelle neue XLSX
            with zipfile.ZipFile(temp_xlsx, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(temp_dir):
                    for f in files:
                        if f == 'restored.xlsx':
                            continue
                        full_path = os.path.join(root, f)
                        arc_name = full_path.replace(temp_dir + os.sep, '')
                        zf.write(full_path, arc_name)
            
            shutil.copy2(temp_xlsx, output_path)
    
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
    """
    from datetime import date
    from openpyxl.cell.cell import MergedCell
    import re
    
    # MergedCell überspringen - nur die obere linke Zelle einer Merged-Region ist beschreibbar
    if isinstance(cell, MergedCell):
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
        
        # DEBUG: Zeige alle relevanten Flags
        
        # =====================================================================
        # FALL 1: fromFile - Nur versteckte Spalten/Zeilen setzen
        # =====================================================================
        if from_file:
            _apply_hidden_columns(ws, hidden_columns)
            _apply_hidden_rows(ws, hidden_rows)
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            return {'success': True, 'outputPath': output_path}
        
        # =====================================================================
        # FALL 1.X: UNIVERSELLE PIPELINE für Spalten-Operationen
        # Führt alle Operationen SERIELL in korrekter Reihenfolge aus:
        # 1. Spalten löschen (von hinten nach vorne)
        # 2. Spalten einfügen (von vorne nach hinten) 
        # 3. Spalten verschieben/reorder
        # 4. Zellwerte schreiben
        # 5. Spalten/Zeilen verstecken
        # 6. Row Highlights
        # 7. Tables reparieren
        # 8. Einmal speichern + restore XML
        # =====================================================================
        
        # Prüfe ob rowMapping nur die Identität ist (keine echte Änderung)
        row_mapping_is_identity = True
        if row_mapping:
            for i, val in enumerate(row_mapping):
                if val != i:
                    row_mapping_is_identity = False
                    break
        
        # Prüfe ob wir den Pipeline-Pfad nutzen können
        # (Spalten-Operationen ohne komplexe Zeilen-Änderungen)
        has_column_operations = deleted_columns or inserted_columns or (column_order and len(column_order) > 0)
        can_use_pipeline = has_column_operations and row_mapping_is_identity and not affected_rows
        
        if can_use_pipeline:
            from openpyxl.worksheet.table import TableColumn
            from openpyxl.utils.cell import range_boundaries
            from openpyxl.cell.cell import MergedCell
            
            # ===== SCHRITT 1: Spalten LÖSCHEN (von hinten nach vorne) =====
            if deleted_columns:
                sorted_deleted = sorted(deleted_columns, reverse=True)
                
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
            
            # ===== SCHRITT 2: Spalten EINFÜGEN (von vorne nach hinten) =====
            if inserted_columns:
                operations = inserted_columns.get('operations', [])
                if not operations and inserted_columns.get('position') is not None:
                    operations = [{
                        'position': inserted_columns['position'],
                        'count': inserted_columns.get('count', 1),
                        'sourceColumn': inserted_columns.get('sourceColumn')
                    }]
                
                operations.sort(key=lambda x: x['position'])
                
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
            
            # ===== SCHRITT 3: Spalten VERSCHIEBEN/REORDER =====
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
            
            # ===== SCHRITT 4: Versteckte Spalten/Zeilen =====
            _apply_hidden_columns(ws, hidden_columns)
            _apply_hidden_rows(ws, hidden_rows)
            
            # ===== SCHRITT 5: Row Highlights =====
            if row_highlights:
                _apply_row_highlights(ws, row_highlights, len(headers) if headers else 0)
            
            # ===== SCHRITT 6: Tables reparieren =====
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
            
            # ===== SCHRITT 7: EINMAL speichern =====
            wb.save(output_path)
            wb.close()
            fix_xlsx_relationships(output_path)
            
            # ===== SCHRITT 8: XML restore =====
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
        # Führt erst Delete (wie FALL 1.6) und dann Insert (wie FALL 1.5) aus
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
                        
                        # Aktualisiere Conditional Formatting Bereiche
                        # row_mapping[new_idx] = orig_idx bedeutet:
                        # - Originale Datenzeile orig_idx ist jetzt an Position new_idx
                        # - Excel-Zeile orig_idx+2 wird zu Excel-Zeile new_idx+2
                        # - Bei eingefügten Zeilen ist orig_idx = -1 (neue Zeile ohne Original)
                        
                        # Erstelle Mapping: alte_zeile -> neue_zeile
                        row_shift_map = {}  # old_excel_row -> new_excel_row
                        inserted_rows = set()  # neue Zeilen die eingefügt wurden
                        for new_idx, orig_idx in enumerate(row_mapping):
                            new_excel_row = new_idx + 2
                            if orig_idx < 0:
                                # Neue eingefügte Zeile (orig_idx = -1)
                                inserted_rows.add(new_excel_row)
                            else:
                                old_excel_row = orig_idx + 2  # +2 für Header
                                row_shift_map[old_excel_row] = new_excel_row
                        
                        # Finde gelöschte Zeilen (die nicht im Mapping sind)
                        # WICHTIG: Nur orig_idx aus row_mapping prüfen, nicht alle bis current_max_row
                        # Bei reiner Verschiebung enthält row_mapping alle Indizes 0 bis len-1
                        expected_orig_indices = set(range(len(row_mapping)))  # 0 bis n-1
                        actual_orig_indices = set(idx for idx in row_mapping if idx >= 0)
                        
                        deleted_rows = set()
                        # Prüfe ob alle erwarteten orig_indices vorhanden sind
                        for expected_idx in expected_orig_indices:
                            if expected_idx not in actual_orig_indices:
                                deleted_rows.add(expected_idx + 2)  # +2 für Excel-Zeile
                        
                        # Bestimme ob nur Verschiebung (keine Löschung/Einfügung)
                        is_pure_reorder = len(deleted_rows) == 0 and len(inserted_rows) == 0
                        
                        if deleted_rows:
                            sys.stderr.write(f"[ZIP-ANSATZ] Gelöschte Zeilen: {sorted(deleted_rows)[:10]}...\n")
                        if inserted_rows:
                            sys.stderr.write(f"[ZIP-ANSATZ] Eingefügte Zeilen: {sorted(inserted_rows)[:10]}...\n")
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
                            for new_data_idx, orig_data_idx in enumerate(row_mapping):
                                new_excel_row = new_data_idx + 2  # Ziel-Zeile in Excel
                                
                                if orig_data_idx < 0:
                                    # NEUE EINGEFÜGTE ZEILE - muss erstellt werden
                                    if template_row is not None:
                                        import copy
                                        new_row_elem = copy.deepcopy(template_row)
                                        new_row_elem.set('r', str(new_excel_row))
                                        
                                        # Alle Zellen umnummerieren und Werte setzen
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
                                    orig_excel_row = orig_data_idx + 2  # Quell-Zeile mit der Formatierung
                                    
                                    if orig_excel_row in row_dict:
                                        row_elem = row_dict[orig_excel_row]
                                        
                                        # Zeile umnummerieren
                                        row_elem.set('r', str(new_excel_row))
                                        
                                        # Alle Zellen in der Zeile umnummerieren
                                        cells = row_elem.findall('main:c', ns)
                                        for cell in cells:
                                            old_ref = cell.get('r')
                                            if old_ref:
                                                # Extrahiere Spalte und ersetze Zeile
                                                col_match = re.match(r'([A-Z]+)\d+', old_ref)
                                                if col_match:
                                                    col = col_match.group(1)
                                                    cell.set('r', f"{col}{new_excel_row}")
                                        
                                        new_rows.append((new_excel_row, row_elem))
                            
                            # Alle alten Zeilen entfernen
                            for row_elem in list(sheet_data):
                                sheet_data.remove(row_elem)
                            
                            # Neue Zeilen in korrekter Reihenfolge einfügen
                            new_rows.sort(key=lambda x: x[0])
                            for row_num, row_elem in new_rows:
                                sheet_data.append(row_elem)
                            
                            sys.stderr.write(f"[ZIP-ANSATZ] {len(new_rows)} Zeilen neu angeordnet\n")
                        
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
                
                sys.stderr.write(f"[SCHRITT 0.5] row_mapping Länge: {len(row_mapping)}\n")
                sys.stderr.write(f"[SCHRITT 0.5] identity_mapping erste 5: {identity_mapping[:5]}\n")
                sys.stderr.write(f"[SCHRITT 0.5] row_mapping erste 5: {row_mapping[:5]}\n")
                sys.stderr.write(f"[SCHRITT 0.5] needs_reorder: {needs_reorder}\n")
                
                # DEBUG: Prüfe CF-Regeln
                cf_count = len(list(ws.conditional_formatting._cf_rules.items()))
                sys.stderr.write(f"[DEBUG] Conditional Formatting Regeln: {cf_count}\n")
                
                # DEBUG: Prüfe Tables (Zebra-Muster könnte Table Style sein)
                if hasattr(ws, '_tables'):
                    sys.stderr.write(f"[DEBUG] Tables im Worksheet: {len(ws._tables)}\n")
                    for table in ws._tables:
                        try:
                            if hasattr(table, 'name'):
                                sys.stderr.write(f"[DEBUG] Table: {table.name}, Bereich: {table.ref}, Style: {table.tableStyleInfo.name if table.tableStyleInfo else 'keiner'}\n")
                            else:
                                sys.stderr.write(f"[DEBUG] Table (string): {table}\n")
                        except Exception as e:
                            sys.stderr.write(f"[DEBUG] Table error: {e}\n")
                
                # Zeige nur große Bereiche (Zebra-Muster)
                for cf_obj, rules in list(ws.conditional_formatting._cf_rules.items()):
                    sqref_str = str(cf_obj.sqref)
                    # Prüfe ob es ein großer Bereich ist (mehr als 100 Zeilen)
                    if ':' in sqref_str:
                        parts = sqref_str.split(':')
                        if len(parts) == 2:
                            import re
                            match1 = re.search(r'\d+', parts[0])
                            match2 = re.search(r'\d+', parts[1])
                            if match1 and match2:
                                row1 = int(match1.group())
                                row2 = int(match2.group())
                                if row2 - row1 > 100:
                                    sys.stderr.write(f"[DEBUG] Großer CF Bereich: {sqref_str}, Regeln: {len(rules)}, Typ: {rules[0].type if rules else 'unbekannt'}\n")
                
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
                    
                    sys.stderr.write(f"[SCHRITT 0.5] Zellen mit Formatierung gefunden: {styles_found}\n")
                    sys.stderr.write(f"[SCHRITT 0.5] row_data_with_styles Anzahl: {len(row_data_with_styles)}\n")
                    
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
                    
                    sys.stderr.write(f"[SCHRITT 0.5] Styles angewendet (fill+font): {styles_applied}\n")
                    
                    # DEBUG: Finde eine Zelle mit Formatierung in Zeile 2 nach dem Kopieren
                    for test_col in range(1, min(ws.max_column + 1, 50)):
                        test_cell_after = ws.cell(row=2, column=test_col)
                        if test_cell_after.fill and test_cell_after.fill.patternType and test_cell_after.fill.patternType != 'none':
                            sys.stderr.write(f"[DEBUG] Nach SCHRITT 0.5: Zeile 2 Spalte {test_col}: patternType={test_cell_after.fill.patternType}, fgColor={test_cell_after.fill.fgColor.rgb if test_cell_after.fill.fgColor else None}\n")
                            break
                    else:
                        sys.stderr.write(f"[DEBUG] Nach SCHRITT 0.5: Keine formatierte Zelle in Zeile 2 gefunden\n")
                    
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
            
            # DEBUG: Finde eine Zelle mit Formatierung in Zeile 2 nach SCHRITT 4
            for test_col in range(1, min(ws.max_column + 1, 50)):
                test_cell_after4 = ws.cell(row=2, column=test_col)
                if test_cell_after4.fill and test_cell_after4.fill.patternType and test_cell_after4.fill.patternType != 'none':
                    sys.stderr.write(f"[DEBUG] Nach SCHRITT 4: Zeile 2 Spalte {test_col}: patternType={test_cell_after4.fill.patternType}, fgColor={test_cell_after4.fill.fgColor.rgb if test_cell_after4.fill.fgColor else None}\n")
                    break
            else:
                sys.stderr.write(f"[DEBUG] Nach SCHRITT 4: Keine formatierte Zelle in Zeile 2 gefunden\n")
            
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
        if edited_cells:
            for key, value in edited_cells.items():
                if key.startswith('_'):
                    continue
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
        
        # Row Highlights
        if row_highlights:
            _apply_row_highlights(ws, row_highlights, ws.max_column)
        
        wb.save(output_path)
        wb.close()
        fix_xlsx_relationships(output_path)
        
        # WICHTIG: Table-XML vom Original wiederherstellen!
        # openpyxl verliert beim Speichern xr3:uid Attribute,
        # deshalb müssen wir die Table-XML aus der Original-Datei kopieren.
        restore_table_xml_from_original(output_path, original_path, table_changes=None)
        
        # WICHTIG: Auch workbook.xml, slicerCaches, etc. vom Original wiederherstellen!
        # openpyxl verliert Slicers, Extensions und viele Namespaces
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
