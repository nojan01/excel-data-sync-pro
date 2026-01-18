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
    """
    import zipfile
    import tempfile
    import shutil
    import re
    
    if not table_changes:
        return
    
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
                        # Extrahiere die xr3:uid Attribute aus den Original-Columns (für Wiederverwendung)
                        orig_columns = re.findall(r'<tableColumn\s[^/]*(?:/>|>.*?</tableColumn>)', tc_match.group(0), re.DOTALL)
                        
                        # Extrahiere xr3:uid falls vorhanden
                        def get_uid(col_xml):
                            uid_match = re.search(r'xr3:uid="([^"]+)"', col_xml)
                            return uid_match.group(1) if uid_match else None
                        
                        # Baue neue tableColumns
                        new_tc_content = f'<tableColumns count="{len(new_columns)}">'
                        
                        for i, col_name in enumerate(new_columns):
                            # Versuche eine passende Original-Column zu finden (nach Name)
                            matching_orig = None
                            for orig_col in orig_columns:
                                name_match = re.search(r'name="([^"]+)"', orig_col)
                                if name_match and name_match.group(1) == col_name:
                                    matching_orig = orig_col
                                    break
                            
                            if matching_orig:
                                # Nutze Original-Column und aktualisiere nur die ID
                                col_xml = re.sub(r'id="\d+"', f'id="{i+1}"', matching_orig)
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


def write_sheet(file_path, output_path, sheet_name, changes):
    """
    Schreibt Änderungen in ein Excel-Sheet
    
    WICHTIG: Bei strukturellen Änderungen (fullRewrite=True) werden die 
    NEUEN Daten geschrieben. Die Original-Struktur wird beibehalten wo möglich.
    
    Args:
        file_path: Pfad zur Original-Datei
        output_path: Pfad zur Ausgabe-Datei
        sheet_name: Name des Sheets
        changes: Dict mit allen Änderungen
    
    Returns:
        Dict mit success und ggf. error
    """
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
        # FALL 2: Strukturelle Änderungen (fullRewrite)
        # WICHTIG: openpyxl's delete_cols() passt CF-Bereiche NICHT an!
        # Wenn Excel installiert ist, nutzen wir xlwings für perfekten CF-Erhalt.
        # =====================================================================
        if structural_change or full_rewrite:
            
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
                                if cell_info.get('font'):
                                    cell.font = cell_info['font']
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
                
                for op in operations:
                    position = op['position']
                    count = op.get('count', 1)
                    source_column = op.get('sourceColumn')  # Referenzspalte für Formatierung
                    excel_col = position + 1  # 0-basiert → 1-basiert
                    
                    
                    # FÜR JEDE NEUE SPALTE einzeln:
                    for i in range(count):
                        insert_at = excel_col + i
                        
                        # 0. FORMATIERUNG DER REFERENZSPALTE SPEICHERN (falls vorhanden)
                        source_format = {}
                        source_width = None
                        if source_column is not None:
                            source_excel_col = source_column + 1
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
            if table_changes:
                restore_table_xml_from_original(output_path, file_path, table_changes)
            
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
            params.get('changes', {})
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
