"""
Direkte XML-Spaltenoperationen für Excel .xlsx Dateien (ZIP-to-ZIP).

Kein openpyxl-Roundtrip → alle Strukturen bleiben intakt:
- Namespaces, Slicers, Drawings, Media, RichData, External Links
- Tables, SharedStrings, Styles, Conditional Formatting

Prinzip: Original-ZIP entpacken, nur die nötigen XMLs ändern, wieder einpacken.
Alles was nicht geändert wird, bleibt 1:1 erhalten.
"""
import re
import os
import sys
import zipfile
import shutil
from xml.etree import ElementTree as ET


# =============================================================================
# SPALTEN-HILFSFUNKTIONEN
# =============================================================================

def _col_letter_to_num(letter):
    """Wandelt Spaltenbuchstabe(n) in 1-basierte Nummer um. A=1, B=2, Z=26, AA=27."""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def _num_to_col_letter(num):
    """Wandelt 1-basierte Spaltennummer in Buchstabe(n) um. 1=A, 2=B, 26=Z, 27=AA."""
    result = ''
    while num > 0:
        num -= 1
        result = chr(ord('A') + num % 26) + result
        num //= 26
    return result


def _parse_cell_ref(ref):
    """
    Zerlegt 'AB123' oder '$AB$123' in (dollar1, col_letter, dollar2, row_number).
    Gibt (None, None, None, None) bei ungültigem Ref.
    """
    m = re.match(r'^(\$?)([A-Z]+)(\$?)(\d+)$', ref)
    if m:
        return m.group(1), m.group(2), m.group(3), int(m.group(4))
    return None, None, None, None


def _remap_col_in_ref(ref, col_map):
    """
    Wendet ein Spalten-Mapping auf eine Zellreferenz an.
    col_map: dict von alter 1-basierter Spaltennummer → neuer 1-basierter Spaltennummer.
    Spalten die nicht im Mapping sind → ref wird entfernt (None zurückgegeben).
    """
    dollar1, col_letter, dollar2, row_num = _parse_cell_ref(ref)
    if col_letter is None:
        return ref  # Kein gültiger Cell-Ref
    old_num = _col_letter_to_num(col_letter)
    new_num = col_map.get(old_num)
    if new_num is None:
        return None  # Spalte wurde gelöscht
    new_letter = _num_to_col_letter(new_num)
    return f"{dollar1}{new_letter}{dollar2}{row_num}"


def _remap_range_ref(range_ref, col_map):
    """
    Wendet ein Spalten-Mapping auf eine Range-Referenz an (z.B. 'A1:C10').
    Bei gelöschten Randspalten wird der Bereich auf die verbleibenden Spalten geschrumpft.
    Gibt None zurück nur wenn ALLE Spalten im Bereich gelöscht wurden.
    """
    if ':' not in range_ref:
        return _remap_col_in_ref(range_ref, col_map)
    parts = range_ref.split(':')
    new_start = _remap_col_in_ref(parts[0], col_map)
    new_end = _remap_col_in_ref(parts[1], col_map)
    if new_start is not None and new_end is not None:
        return f"{new_start}:{new_end}"

    # Mindestens eine Randspalte wurde gelöscht →
    # tatsächlichen Bereich aller verbleibenden Spalten ermitteln
    d1s, col_s, d2s, row_s = _parse_cell_ref(parts[0])
    d1e, col_e, d2e, row_e = _parse_cell_ref(parts[1])
    if col_s is None or col_e is None:
        return None

    old_min = _col_letter_to_num(col_s)
    old_max = _col_letter_to_num(col_e)

    mapped_cols = []
    for c in range(old_min, old_max + 1):
        nc = col_map.get(c)
        if nc is not None:
            mapped_cols.append(nc)

    if not mapped_cols:
        return None  # Alle Spalten im Bereich gelöscht

    new_min_col = min(mapped_cols)
    new_max_col = max(mapped_cols)
    new_start_ref = f"{d1s}{_num_to_col_letter(new_min_col)}{d2s}{row_s}"
    new_end_ref = f"{d1e}{_num_to_col_letter(new_max_col)}{d2e}{row_e}"
    return f"{new_start_ref}:{new_end_ref}"


def _remap_sqref(sqref, col_map):
    """
    Wendet ein Spalten-Mapping auf sqref an (Space-separated Ranges).
    Entfernt Ranges deren Spalten gelöscht wurden.
    """
    ranges = sqref.split()
    new_ranges = []
    for r in ranges:
        new_r = _remap_range_ref(r, col_map)
        if new_r is not None:
            new_ranges.append(new_r)
    return ' '.join(new_ranges) if new_ranges else None


# =============================================================================
# SPALTEN-MAPPING BUILDER
# =============================================================================

def _build_col_map_for_delete(deleted_cols_0based, max_col=500):
    """
    Baut ein Spalten-Mapping für Spalten-Löschung.
    deleted_cols_0based: Liste von 0-basierten Spaltenindizes die gelöscht werden.
    Gibt dict zurück: alte 1-basierte Spaltennummer → neue 1-basierte Spaltennummer.
    Gelöschte Spalten fehlen im Mapping (→ None bei Lookup).
    """
    deleted_set = set(c + 1 for c in deleted_cols_0based)  # 1-basiert
    col_map = {}
    new_col = 1
    for old_col in range(1, max_col + 1):
        if old_col in deleted_set:
            continue  # Gelöscht → nicht im Mapping
        col_map[old_col] = new_col
        new_col += 1
    return col_map


def _build_col_map_for_insert(insert_operations, max_col=500):
    """
    Baut ein Spalten-Mapping für Spalten-Einfügung.
    insert_operations: Liste von {position: 0-basiert, count: Anzahl}.
    Gibt dict zurück: alte 1-basierte Spaltennummer → neue 1-basierte Spaltennummer.
    
    WICHTIG: Die Frontend-Positionen sind kumulativ — jede Position bezieht sich
    auf den Zustand NACH allen vorherigen Inserts. Daher müssen sie in
    Original-Positionen konvertiert werden (Anzahl vorheriger Inserts abziehen).
    """
    inserts = sorted([(op['position'] + 1, op.get('count', 1)) for op in insert_operations])
    
    # Kumulative Positionen → Original-Positionen konvertieren
    total_inserted_before = 0
    adjusted = []
    for pos, count in inserts:
        adjusted.append((pos - total_inserted_before, count))
        total_inserted_before += count
    inserts = adjusted
    
    col_map = {}
    shift = 0
    insert_idx = 0
    for old_col in range(1, max_col + 1):
        while insert_idx < len(inserts) and inserts[insert_idx][0] == old_col:
            shift += inserts[insert_idx][1]
            insert_idx += 1
        col_map[old_col] = old_col + shift
    return col_map


def _build_col_map_for_reorder(column_order):
    """
    Baut ein Spalten-Mapping für Spalten-Verschiebung.
    column_order: Liste wo column_order[neuer_0based_idx] = alter_0based_idx.
    Gibt dict zurück: alte 1-basierte Spaltennummer → neue 1-basierte Spaltennummer.
    """
    col_map = {}
    for new_idx, old_idx in enumerate(column_order):
        col_map[old_idx + 1] = new_idx + 1
    # Spalten jenseits der column_order bleiben an ihrem Platz
    max_mapped = len(column_order)
    for old_col in range(max_mapped + 1, max_mapped + 200):
        col_map[old_col] = old_col
    return col_map


# =============================================================================
# XML-TRANSFORMATIONEN
# =============================================================================

def _apply_col_map_to_sheet_xml(sheet_xml, col_map, skip_sort=False):
    """
    Wendet ein Spalten-Mapping auf alle relevanten Elemente im Worksheet-XML an.
    
    Modifiziert:
    - <c r="A1"> Zell-Referenzen in <sheetData>
    - <row spans="1:10">
    - <col min="1" max="1"> Spalten-Definitionen
    - <mergeCell ref="A1:C1">
    - <conditionalFormatting sqref="A2:A100">
    - <autoFilter ref="A1:J100">
    - <hyperlink ref="A2">
    - <dataValidation sqref="B2:B100">
    - <xm:sqref> in <extLst> (x14 CF, Sparklines etc.)
    
    Args:
        skip_sort: Wenn True, wird die Zell-Sortierung übersprungen.
                   Nützlich wenn der Aufrufer selbst sortiert (Performance).
    """
    
    # ---- 0. <dimension ref="A1:J100"> aktualisieren ----
    def _remap_dimension(m):
        ref = m.group(1)
        new_ref = _remap_range_ref(ref, col_map)
        if new_ref is None:
            return ''  # Dimension entfernen wenn alles gelöscht
        return f'<dimension ref="{new_ref}"/>'

    sheet_xml = re.sub(r'<dimension\s+ref="([^"]+)"\s*/>', _remap_dimension, sheet_xml)

    # ---- 1. <c r="..."> Zell-Referenzen ----
    # Ermittle welche Spalten gelöscht werden (nicht im Mapping)
    deleted_col_nums = set()
    max_in_map = max(col_map.keys()) if col_map else 0
    for c in range(1, max_in_map + 1):
        if c not in col_map:
            deleted_col_nums.add(c)
    
    # Entferne zuerst alle <c>-Elemente gelöschter Spalten komplett
    if deleted_col_nums:
        def _remove_deleted_cell(m):
            col_letter = m.group(1)
            col_num = _col_letter_to_num(col_letter)
            if col_num in deleted_col_nums:
                return ''  # Ganzes Element entfernen
            return m.group(0)
        # Selbstschließende <c .../> 
        sheet_xml = re.sub(r'<c\s[^>]*?r="([A-Z]+)\d+"[^>]*/>', _remove_deleted_cell, sheet_xml)
        # <c ...>...</c>
        sheet_xml = re.sub(r'<c\s[^>]*?r="([A-Z]+)\d+"[^>]*>.*?</c>', _remove_deleted_cell, sheet_xml, flags=re.DOTALL)
    
    # Jetzt alle verbleibenden Zell-Referenzen remappen
    def _remap_cell(m):
        prefix = m.group(1)  # z.B. '<c r="' (alles vor dem Ref)
        ref_str = m.group(2)  # z.B. 'C5'
        dollar1, col_letter, dollar2, row_num = _parse_cell_ref(ref_str)
        if col_letter is None:
            return m.group(0)
        old_num = _col_letter_to_num(col_letter)
        new_num = col_map.get(old_num, old_num)
        new_letter = _num_to_col_letter(new_num)
        new_ref = f"{dollar1}{new_letter}{dollar2}{row_num}"
        return f'{prefix}{new_ref}"'
    
    sheet_xml = re.sub(r'(<c\s[^>]*?r=")([A-Z]+\d+)"', _remap_cell, sheet_xml)
    
    # ---- 2. <row spans="1:10"> — entfernen, Excel berechnet sie neu ----
    sheet_xml = re.sub(r' spans="[^"]*"', '', sheet_xml)
    
    # ---- 3. <col min="1" max="1" ...> ----
    cols_match = re.search(r'<cols>(.*?)</cols>', sheet_xml, re.DOTALL)
    if cols_match:
        cols_content = cols_match.group(1)
        col_elements = re.findall(r'<col\s[^>]*/>', cols_content)
        
        # Sammle alle Spalten-Properties nach alter Nummer
        col_props = {}
        for col_el in col_elements:
            min_m = re.search(r'min="(\d+)"', col_el)
            max_m = re.search(r'max="(\d+)"', col_el)
            if not min_m or not max_m:
                continue
            col_min = int(min_m.group(1))
            col_max = int(max_m.group(1))
            
            attrs = {}
            for attr_m in re.finditer(r'(\w+)="([^"]*)"', col_el):
                if attr_m.group(1) not in ('min', 'max'):
                    attrs[attr_m.group(1)] = attr_m.group(2)
            
            for c in range(col_min, col_max + 1):
                col_props[c] = dict(attrs)
        
        # Mappe auf neue Positionen
        new_col_props = {}
        for old_col, props in col_props.items():
            new_col = col_map.get(old_col)
            if new_col is not None:
                new_col_props[new_col] = props
        
        # Baue <cols> neu, gruppiere zusammenhängende gleiche Properties
        if new_col_props:
            sorted_cols = sorted(new_col_props.keys())
            groups = []
            current_start = sorted_cols[0]
            current_attrs = new_col_props[sorted_cols[0]]
            
            for i in range(1, len(sorted_cols)):
                c = sorted_cols[i]
                if c == sorted_cols[i - 1] + 1 and new_col_props[c] == current_attrs:
                    continue
                else:
                    groups.append((current_start, sorted_cols[i - 1], current_attrs))
                    current_start = c
                    current_attrs = new_col_props[c]
            groups.append((current_start, sorted_cols[-1], current_attrs))
            
            new_cols_xml = '<cols>'
            for g_min, g_max, attrs in groups:
                attr_str = ' '.join(f'{k}="{v}"' for k, v in attrs.items())
                new_cols_xml += f'<col min="{g_min}" max="{g_max}" {attr_str}/>'
            new_cols_xml += '</cols>'
            
            sheet_xml = sheet_xml[:cols_match.start()] + new_cols_xml + sheet_xml[cols_match.end():]
        else:
            sheet_xml = sheet_xml[:cols_match.start()] + sheet_xml[cols_match.end():]
    
    # ---- 4. <mergeCell ref="A1:C1"> ----
    def _remap_merge(m):
        ref = m.group(1)
        new_ref = _remap_range_ref(ref, col_map)
        if new_ref is None:
            return ''
        return f'<mergeCell ref="{new_ref}"/>'
    
    sheet_xml = re.sub(r'<mergeCell\s+ref="([^"]+)"\s*/>', _remap_merge, sheet_xml)
    merge_count = len(re.findall(r'<mergeCell\s', sheet_xml))
    sheet_xml = re.sub(r'<mergeCells\s+count="\d+"', f'<mergeCells count="{merge_count}"', sheet_xml)
    sheet_xml = re.sub(r'<mergeCells\s+count="0"\s*>\s*</mergeCells>', '', sheet_xml)
    
    # ---- 5. <conditionalFormatting sqref="..."> ----
    def _remap_cf(m):
        sqref = m.group(1)
        rest = m.group(2)
        new_sqref = _remap_sqref(sqref, col_map)
        if new_sqref is None:
            return ''
        return f'<conditionalFormatting sqref="{new_sqref}"{rest}'
    
    sheet_xml = re.sub(r'<conditionalFormatting\s+sqref="([^"]+)"([^>]*)', _remap_cf, sheet_xml)
    sheet_xml = re.sub(
        r'<conditionalFormatting\s+sqref=""\s*>.*?</conditionalFormatting>',
        '', sheet_xml, flags=re.DOTALL)
    
    # ---- 6. <autoFilter ref="A1:J100"> ----
    def _remap_autofilter(m):
        ref = m.group(1)
        new_ref = _remap_range_ref(ref, col_map)
        if new_ref is None:
            return ''
        return f'<autoFilter ref="{new_ref}"'
    
    sheet_xml = re.sub(r'<autoFilter\s+ref="([^"]+)"', _remap_autofilter, sheet_xml)
    
    # ---- 6b. <filterColumn colId="X"> innerhalb <autoFilter> ----
    # colId ist 0-basiert. Nach Spaltenoperationen müssen die colIds
    # angepasst werden, sonst verwirft Excel den AutoFilter.
    af_match = re.search(r'<autoFilter\s+ref="([^"]+)"', sheet_xml)
    if af_match:
        af_ref = af_match.group(1)
        if ':' in af_ref:
            af_start_cell = af_ref.split(':')[0]
            _, af_start_col, _, _ = _parse_cell_ref(af_start_cell)
            af_start_num = _col_letter_to_num(af_start_col) if af_start_col else 1
        else:
            af_start_num = 1
        
        # Baue Mapping: alter colId → neuer colId
        # colId 0-based, Spalte = af_start_num + colId
        new_positions = sorted(set(col_map.values()))
        
        def _remap_sheet_filter_column(m):
            full = m.group(0)
            old_colid = int(m.group(1))
            old_col = af_start_num + old_colid
            new_col = col_map.get(old_col)
            if new_col is None:
                return ''  # Spalte gelöscht → filterColumn entfernen
            new_colid = new_positions.index(new_col) if new_col in new_positions else old_colid
            return full.replace(f'colId="{old_colid}"', f'colId="{new_colid}"')
        
        sheet_xml = re.sub(
            r'<filterColumn\s[^>]*colId="(\d+)"[^>]*/>\s*',
            _remap_sheet_filter_column, sheet_xml)
        sheet_xml = re.sub(
            r'<filterColumn\s[^>]*colId="(\d+)"[^>]*>.*?</filterColumn>\s*',
            _remap_sheet_filter_column, sheet_xml, flags=re.DOTALL)
        
        # sortState/sortCondition ref-Bereiche anpassen
        def _remap_sort_ref_sheet(m):
            prefix = m.group(1)
            ref = m.group(2)
            new_ref = _remap_range_ref(ref, col_map)
            if new_ref is None:
                return ''
            return f'{prefix}ref="{new_ref}"'
        
        sheet_xml = re.sub(r'(<sortState\s[^>]*?)ref="([^"]+)"', _remap_sort_ref_sheet, sheet_xml)
        sheet_xml = re.sub(r'(<sortCondition\s[^>]*?)ref="([^"]+)"', _remap_sort_ref_sheet, sheet_xml)
        
        # Leere autoFilter bereinigen
        sheet_xml = re.sub(r'(<autoFilter\s[^>]*?)>\s*</autoFilter>', r'\1/>', sheet_xml)
    
    # ---- 7. <hyperlink ref="A2" ...> ----
    def _remap_hyperlink(m):
        full = m.group(0)
        ref = m.group(1)
        new_ref = _remap_col_in_ref(ref, col_map)
        if new_ref is None:
            return ''
        return full.replace(f'ref="{ref}"', f'ref="{new_ref}"')
    
    sheet_xml = re.sub(r'<hyperlink\s[^>]*ref="([^"]+)"[^>]*/>', _remap_hyperlink, sheet_xml)
    
    # ---- 8. <dataValidation sqref="B2:B100" ...> ----
    def _remap_dv(m):
        full = m.group(0)
        sqref = m.group(1)
        new_sqref = _remap_sqref(sqref, col_map)
        if new_sqref is None:
            return ''
        return full.replace(f'sqref="{sqref}"', f'sqref="{new_sqref}"')
    
    sheet_xml = re.sub(r'<dataValidation\s[^>]*sqref="([^"]+)"[^>]*/?>', _remap_dv, sheet_xml)
    
    # ---- 9. Definierte Namen mit Sheet-Referenzen ----
    # definedName-Werte wie "Sheet1!$A$1:$J$100" in workbook.xml
    # werden hier NICHT geändert (workbook.xml wird separat behandelt)
    
    # ---- 9b. <xm:sqref> in <extLst> (x14 CF, Sparklines, Data Bars etc.) ----
    # Excel 2016+ speichert erweiterte Conditional Formatting in <extLst> mit
    # <xm:sqref>-Elementen statt sqref-Attributen. Diese müssen ebenfalls
    # remapped werden, sonst erkennt Excel die Inkonsistenz zwischen Standard-CF
    # und Extended-CF und repariert die Datei.
    def _remap_xm_sqref(m):
        sqref = m.group(1)
        new_sqref = _remap_sqref(sqref, col_map)
        if new_sqref is None:
            return ''  # Alle Spalten gelöscht → Element entfernen
        return f'<xm:sqref>{new_sqref}</xm:sqref>'
    
    sheet_xml = re.sub(r'<xm:sqref>([^<]+)</xm:sqref>', _remap_xm_sqref, sheet_xml)
    
    # Auch xm:f Formeln mit Zellreferenzen in extLst anpassen
    # (z.B. Data Bars: <xm:f>Sheet1!$A$2</xm:f>)
    def _remap_xm_f(m):
        formula = m.group(1)
        # Nur einfache Zellreferenzen/Ranges remappen (Sheet!$A$1:$B$10 oder $A$1:$B$10)
        # Komplexe Formeln mit Funktionen lassen wir unverändert
        if re.match(r'^[^(]+!?\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$', formula):
            # Extrahiere optional Sheet-Prefix
            if '!' in formula:
                sheet_prefix, ref_part = formula.rsplit('!', 1)
                new_ref = _remap_range_ref(ref_part, col_map) if ':' in ref_part else _remap_col_in_ref(ref_part, col_map)
                if new_ref is None:
                    return ''  # Referenz gelöscht
                return f'<xm:f>{sheet_prefix}!{new_ref}</xm:f>'
            else:
                new_ref = _remap_range_ref(formula, col_map) if ':' in formula else _remap_col_in_ref(formula, col_map)
                if new_ref is None:
                    return ''
                return f'<xm:f>{new_ref}</xm:f>'
        return m.group(0)  # Komplexe Formel unverändert lassen
    
    sheet_xml = re.sub(r'<xm:f>([^<]+)</xm:f>', _remap_xm_f, sheet_xml)
    
    # ---- 10. Zellen innerhalb jeder <row> nach Spalte sortieren ----
    if not skip_sort:
        sheet_xml = _sort_cells_in_rows(sheet_xml)
    
    return sheet_xml


def _sort_cells_in_rows(sheet_xml):
    """
    Sortiert <c> Elemente innerhalb jeder <row> nach Spaltennummer.
    Excel erwartet <c> Elemente in aufsteigender Spaltenreihenfolge!
    """
    def _sort_cells_in_row(row_match):
        row_tag = row_match.group(1)  # <row ...>
        row_content = row_match.group(2)  # Alles zwischen <row> und </row>
        
        # Finde alle <c ...>...</c> und <c .../> Elemente
        cells = []
        for cm in re.finditer(r'<c\s[^>]*r="([A-Z]+)(\d+)"[^>]*/>' 
                              r'|<c\s[^>]*r="([A-Z]+)(\d+)"[^>]*>.*?</c>', row_content, re.DOTALL):
            col_letter = cm.group(1) or cm.group(3)
            col_num = _col_letter_to_num(col_letter)
            cells.append((col_num, cm.group(0)))
        
        if not cells:
            return row_match.group(0)
        
        # Sortiere nach Spaltennummer
        cells.sort(key=lambda x: x[0])
        sorted_content = ''.join(cell_xml for _, cell_xml in cells)
        
        return f'{row_tag}{sorted_content}</row>'
    
    return re.sub(
        r'(<row\s[^>]*>)(.*?)</row>',
        _sort_cells_in_row,
        sheet_xml,
        flags=re.DOTALL
    )


def _apply_col_map_to_table_xml(table_xml, col_map, new_headers=None,
                                 inserted_col_info=None):
    """
    Wendet ein Spalten-Mapping auf eine Table-XML an.
    
    Modifiziert:
    - <table ref="A1:J100">
    - <autoFilter ref="A1:J100">
    - <tableColumn id="..." name="..."> Reihenfolge/Anzahl
    
    Args:
        inserted_col_info: Liste von (new_col_1based, header_name) für eingefügte Spalten.
            Dies sind Spalten die NEU sind und nicht aus col_map kommen.
    """
    
    # Alle Spalten die in der neuen Tabelle existieren
    # = gemappte existierende + eingefügte neue
    all_new_cols = set(col_map.values())
    if inserted_col_info:
        for new_col, _ in inserted_col_info:
            all_new_cols.add(new_col)
    
    # Table ref — ermittle die neuen Grenzen korrekt
    def _remap_table_range(m):
        prefix = m.group(1)
        ref = m.group(2)
        if ':' not in ref:
            return m.group(0)
        start, end = ref.split(':')
        d1s, col_s, d2s, row_s = _parse_cell_ref(start)
        d1e, col_e, d2e, row_e = _parse_cell_ref(end)
        if col_s is None or col_e is None:
            return m.group(0)
        old_min = _col_letter_to_num(col_s)
        old_max = _col_letter_to_num(col_e)
        # Sammle alle neuen Spaltenpositionen im Table-Bereich
        new_cols = []
        for oc in range(old_min, old_max + 1):
            nc = col_map.get(oc)
            if nc is not None:
                new_cols.append(nc)
        # Auch eingefügte Spalten berücksichtigen
        if inserted_col_info:
            for new_col, _ in inserted_col_info:
                new_cols.append(new_col)
        if not new_cols:
            return m.group(0)
        new_min = min(new_cols)
        new_max = max(new_cols)
        new_start = f"{d1s}{_num_to_col_letter(new_min)}{d2s}{row_s}"
        new_end = f"{d1e}{_num_to_col_letter(new_max)}{d2e}{row_e}"
        return f'{prefix}ref="{new_start}:{new_end}"'
    
    table_xml = re.sub(r'(<table\s[^>]*?)ref="([^"]+)"', _remap_table_range, table_xml)
    table_xml = re.sub(r'(<autoFilter\s[^>]*?)ref="([^"]+)"', _remap_table_range, table_xml)
    
    # AutoFilter-Kinder anpassen: <filterColumn colId="X"> und <sortState>
    # Nach Spaltenoperationen sind die colId-Werte ungültig → Excel verwirft
    # die gesamte Tabelle! Wir remappen die colId-Werte basierend auf col_map.
    # 
    # colId ist 0-basiert relativ zum Tabellenbereich.
    # Beispiel: Tabelle A1:J100, filterColumn colId="2" → Spalte C (3. Spalte der Tabelle)
    #
    # Bei gelöschten Spalten: filterColumn entfernen
    # Bei verschobenen Spalten: colId neu berechnen
    
    # Ermittle den Tabellen-Startpunkt (alte und neue Spalte)
    table_ref_m = re.search(r'<table\s[^>]*ref="([^"]+)"', table_xml)
    if table_ref_m:
        t_ref = table_ref_m.group(1)
        if ':' in t_ref:
            t_start = t_ref.split(':')[0]
            _, t_start_col, _, _ = _parse_cell_ref(t_start)
            new_table_start = _col_letter_to_num(t_start_col) if t_start_col else 1
        else:
            new_table_start = 1
    else:
        new_table_start = 1
    
    # Alle neuen Spaltenpositionen sortiert für colId-Berechnung
    all_mapped_new = sorted(set(col_map.values()))
    if inserted_col_info:
        for nc, _ in inserted_col_info:
            if nc not in all_mapped_new:
                all_mapped_new.append(nc)
        all_mapped_new.sort()
    
    # Baue Reverse-Map: old_col → new_0based_index_in_table
    old_col_to_new_colid = {}
    for old_col, new_col in col_map.items():
        if new_col in all_mapped_new:
            old_col_to_new_colid[old_col] = all_mapped_new.index(new_col)
    
    # Ermittle alte Tabellen-Startspalte aus dem Original-ref
    # Wir brauchen das um colId (0-basiert) → alte Spalte (1-basiert) umzurechnen
    # Das Original-ref wurde bereits geändert, also müssen wir die alte Startspalte
    # aus den col_map Keys ableiten
    old_table_start = min(col_map.keys()) if col_map else 1
    
    # filterColumn colId remappen
    def _remap_filter_column(m):
        full_match = m.group(0)
        old_colid = int(m.group(1))
        old_col = old_table_start + old_colid  # 1-basiert
        if old_col in old_col_to_new_colid:
            new_colid = old_col_to_new_colid[old_col]
            return full_match.replace(f'colId="{old_colid}"', f'colId="{new_colid}"')
        else:
            # Spalte wurde gelöscht → filterColumn entfernen
            return ''
    
    # Selbstschließende filterColumn UND filterColumn mit Inhalt
    table_xml = re.sub(
        r'<filterColumn\s[^>]*colId="(\d+)"[^>]*/>\s*',
        _remap_filter_column, table_xml)
    table_xml = re.sub(
        r'<filterColumn\s[^>]*colId="(\d+)"[^>]*>.*?</filterColumn>\s*',
        _remap_filter_column, table_xml, flags=re.DOTALL)
    
    # sortState/sortCondition ref-Bereiche anpassen
    # Diese enthalten Zellbereiche die remapped werden müssen
    def _remap_sort_ref(m):
        prefix = m.group(1)
        ref = m.group(2)
        if ':' not in ref:
            return m.group(0)
        start, end = ref.split(':')
        d1s, col_s, d2s, row_s = _parse_cell_ref(start)
        d1e, col_e, d2e, row_e = _parse_cell_ref(end)
        if col_s is None or col_e is None:
            return m.group(0)
        sc = _col_letter_to_num(col_s)
        ec = _col_letter_to_num(col_e)
        nc_s = col_map.get(sc)
        nc_e = col_map.get(ec)
        if nc_s is None:
            # Start-Spalte gelöscht → nimm erste verfügbare
            for c in range(sc, ec + 1):
                nc_s = col_map.get(c)
                if nc_s:
                    break
        if nc_e is None:
            # End-Spalte gelöscht → nimm letzte verfügbare
            for c in range(ec, sc - 1, -1):
                nc_e = col_map.get(c)
                if nc_e:
                    break
        if nc_s is None or nc_e is None:
            return ''  # Alle Spalten gelöscht → Element entfernen
        new_start = f"{d1s}{_num_to_col_letter(nc_s)}{d2s}{row_s}"
        new_end = f"{d1e}{_num_to_col_letter(nc_e)}{d2e}{row_e}"
        return f'{prefix}ref="{new_start}:{new_end}"'
    
    table_xml = re.sub(r'(<sortState\s[^>]*?)ref="([^"]+)"', _remap_sort_ref, table_xml)
    table_xml = re.sub(r'(<sortCondition\s[^>]*?)ref="([^"]+)"', _remap_sort_ref, table_xml)
    
    # Leere autoFilter bereinigen (nur noch ref, keine Kinder)
    # <autoFilter ref="..."></autoFilter> → <autoFilter ref="..."/>
    table_xml = re.sub(r'(<autoFilter\s[^>]*?)>\s*</autoFilter>', r'\1/>', table_xml)
    
    # tableColumns neu aufbauen
    tc_match = re.search(r'<tableColumns\s[^>]*>(.*?)</tableColumns>', table_xml, re.DOTALL)
    if tc_match:
        # Matche selbstschließende UND nicht-selbstschließende tableColumn-Elemente
        col_elements = list(re.finditer(
            r'<tableColumn\s[^>]*/>|<tableColumn\s[^>]*>.*?</tableColumn>',
            tc_match.group(1), re.DOTALL))
        
        old_columns = []
        for i, tc_m in enumerate(col_elements):
            tc_el = tc_m.group(0)
            name_m = re.search(r'name="([^"]*)"', tc_el)
            name = name_m.group(1) if name_m else f'Column{i + 1}'
            old_columns.append({'xml': tc_el, 'name': name, 'old_col': i + 1})
        
        # Existierende Spalten mappen + sortieren
        new_columns = []  # (new_position, xml_string)
        for tc in old_columns:
            new_col = col_map.get(tc['old_col'])
            if new_col is not None:
                new_el = tc['xml']
                new_columns.append((new_col, tc['name'], new_el))
        
        # Eingefügte Spalten hinzufügen (einfache tableColumn-Elemente)
        if inserted_col_info:
            for new_col, header in inserted_col_info:
                escaped_name = str(header).replace('&', '&amp;').replace(
                    '<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                new_el = f'<tableColumn id="0" name="{escaped_name}"/>'
                new_columns.append((new_col, header, new_el))
        
        new_columns.sort(key=lambda x: x[0])
        
        # Baue tableColumns XML neu mit korrekten IDs
        tc_parts = []
        for new_id, (new_col, name, new_el) in enumerate(new_columns, 1):
            if new_headers and new_id - 1 < len(new_headers):
                name = new_headers[new_id - 1]
            new_el = re.sub(r'id="\d+"', f'id="{new_id}"', new_el)
            tc_parts.append(new_el)
        
        new_tc_xml = f'<tableColumns count="{len(tc_parts)}">' + ''.join(tc_parts) + '</tableColumns>'
        table_xml = table_xml[:tc_match.start()] + new_tc_xml + table_xml[tc_match.end():]
    
    return table_xml


def _apply_col_map_to_drawing_xml(drawing_xml, col_map):
    """
    Wendet ein Spalten-Mapping auf Drawing-XML an.
    Passt <xdr:col> Werte in Anchor-Elementen an (0-basiert im XML, 1-basiert im col_map).
    """
    
    def _remap_drawing_col(m):
        old_col_0 = int(m.group(1))
        old_col_1 = old_col_0 + 1  # col_map ist 1-basiert
        new_col_1 = col_map.get(old_col_1, old_col_1)
        new_col_0 = new_col_1 - 1
        return f'<xdr:col>{new_col_0}</xdr:col>'
    
    drawing_xml = re.sub(r'<xdr:col>(\d+)</xdr:col>', _remap_drawing_col, drawing_xml)
    return drawing_xml


def _apply_col_map_to_workbook_xml(wb_xml, col_map, sheet_name):
    """
    Passt definedNames in workbook.xml an die neuen Spalten an.
    Betrifft z.B. _xlnm._FilterDatabase, _xlnm.Print_Area etc.
    """
    
    def _remap_defined_name_value(m):
        full = m.group(0)
        value = m.group(1)
        # Format: Sheet1!$A$1:$J$100 oder 'Sheet Name'!$A$1:$J$100
        # Nur für das betroffene Sheet anpassen
        parts = value.split('!')
        if len(parts) != 2:
            return full
        sname = parts[0].strip("'")
        if sname != sheet_name:
            return full
        range_ref = parts[1]
        new_range = _remap_range_ref(range_ref, col_map)
        if new_range is None:
            return full
        new_value = f"{parts[0]}!{new_range}"
        return full.replace(value, new_value)
    
    wb_xml = re.sub(
        r'<definedName\s[^>]*>([^<]+)</definedName>',
        _remap_defined_name_value, wb_xml)
    
    return wb_xml


# =============================================================================
# HAUPTFUNKTION
# =============================================================================

def direct_xml_column_operations(file_path, output_path, sheet_name,
                                 deleted_columns=None, inserted_columns=None,
                                 column_order=None, hidden_columns=None,
                                 headers=None, data=None):
    """
    Führt Spaltenoperationen direkt auf dem XML durch (ZIP-to-ZIP).
    
    KEIN openpyxl-Roundtrip → alle Strukturen bleiben intakt:
    - Namespaces, Slicers, Drawings, Media, RichData, External Links
    - Tables, SharedStrings, Styles, Conditional Formatting
    
    Args:
        file_path: Quelldatei (.xlsx)
        output_path: Zieldatei (.xlsx)
        sheet_name: Name des Sheets
        deleted_columns: Liste von 0-basierten Spaltenindizes zum Löschen
        inserted_columns: Dict mit operations: [{position, count, headers, sourceColumn}]
        column_order: Liste wo column_order[new_idx] = old_idx
        hidden_columns: Liste von 0-basierten Spaltenindizes zum Verstecken
        headers: Liste der Header (für eingefügte Spalten)
        data: 2D-Liste der Daten (für eingefügte Spalten)
    
    Returns:
        Dict mit success und outputPath
    """
    sys.stderr.write(f"[XML_COL_OPS] Start für Sheet '{sheet_name}'\n")
    sys.stderr.write(f"[XML_COL_OPS] deleted={deleted_columns}, inserted={inserted_columns is not None}, "
                     f"reorder={column_order is not None}, hidden={hidden_columns}\n")
    
    MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    
    temp_output = output_path + '.tmp'
    
    try:
        with zipfile.ZipFile(file_path, 'r') as src_zip:
            # 1. Finde Sheet-XML Pfad
            wb_xml_raw = src_zip.read('xl/workbook.xml').decode('utf-8')
            wb_root = ET.fromstring(wb_xml_raw)
            
            sheet_rid = None
            for sheet_el in wb_root.iter(f'{{{MAIN_NS}}}sheet'):
                if sheet_el.get('name') == sheet_name:
                    sheet_rid = sheet_el.get(f'{{{R_NS}}}id')
                    break
            
            if not sheet_rid:
                raise ValueError(f"Sheet '{sheet_name}' nicht in workbook.xml gefunden")
            
            rels_xml_raw = src_zip.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            rels_root = ET.fromstring(rels_xml_raw)
            
            sheet_file = None
            for rel_el in rels_root.iter(f'{{{RELS_NS}}}Relationship'):
                if rel_el.get('Id') == sheet_rid:
                    sheet_file = rel_el.get('Target')
                    break
            
            if not sheet_file:
                raise ValueError(f"Relationship {sheet_rid} nicht gefunden")
            
            sheet_zip_path = 'xl/' + sheet_file.lstrip('/')
            parts = sheet_zip_path.split('/')
            normalized = []
            for p in parts:
                if p == '..':
                    if normalized:
                        normalized.pop()
                elif p != '.':
                    normalized.append(p)
            sheet_zip_path = '/'.join(normalized)
            
            sys.stderr.write(f"[XML_COL_OPS] Sheet-ZIP-Pfad: {sheet_zip_path}\n")
            
            # 2. Lese Sheet-XML
            sheet_content = src_zip.read(sheet_zip_path).decode('utf-8')
            
            # 3. Maximale Spalte aus Sheet ermitteln (für Mapping-Größe)
            max_col_in_sheet = 1
            for cm in re.finditer(r'<c\s[^>]*r="([A-Z]+)\d+"', sheet_content):
                col_num = _col_letter_to_num(cm.group(1))
                if col_num > max_col_in_sheet:
                    max_col_in_sheet = col_num
            max_col_for_map = max_col_in_sheet + 50  # Puffer
            
            sys.stderr.write(f"[XML_COL_OPS] Max Spalte im Sheet: {max_col_in_sheet} "
                             f"({_num_to_col_letter(max_col_in_sheet)})\n")
            
            # 4. Finde zugehörige Table- und Drawing-Dateien
            sheet_rels_path = sheet_zip_path.replace(
                'worksheets/', 'worksheets/_rels/') + '.rels'
            table_files = {}  # rId → ZIP-Pfad
            drawing_files = {}  # rId → ZIP-Pfad
            sheet_rels_content = None
            
            if sheet_rels_path in src_zip.namelist():
                sheet_rels_content = src_zip.read(sheet_rels_path).decode('utf-8')
                try:
                    rels_el = ET.fromstring(sheet_rels_content)
                    for rel_node in rels_el.iter(f'{{{RELS_NS}}}Relationship'):
                        rid = rel_node.get('Id', '')
                        target = rel_node.get('Target', '')
                        rtype = rel_node.get('Type', '')
                        
                        if not target.startswith('/'):
                            resolved = 'xl/worksheets/' + target
                            rparts = resolved.split('/')
                            norm = []
                            for p in rparts:
                                if p == '..':
                                    if norm:
                                        norm.pop()
                                elif p != '.':
                                    norm.append(p)
                            resolved = '/'.join(norm)
                        else:
                            resolved = target.lstrip('/')
                        
                        if 'table' in rtype.lower():
                            table_files[rid] = resolved
                        elif 'drawing' in rtype.lower():
                            drawing_files[rid] = resolved
                except ET.ParseError as pe:
                    sys.stderr.write(f"[XML_COL_OPS] WARNUNG: Sheet-Rels Parse-Fehler: {pe}\n")
            
            sys.stderr.write(f"[XML_COL_OPS] Tables: {list(table_files.values())}, "
                             f"Drawings: {list(drawing_files.values())}\n")
            
            # 5. Bestimme das Spalten-Mapping
            # Reihenfolge: Delete → Insert → Reorder
            col_map = None
            
            if deleted_columns:
                col_map = _build_col_map_for_delete(deleted_columns, max_col_for_map)
                sys.stderr.write(f"[XML_COL_OPS] Delete-Map für Spalten {deleted_columns}\n")
            
            if inserted_columns:
                ops = inserted_columns.get('operations', [])
                if not ops and inserted_columns.get('position') is not None:
                    ops = [{
                        'position': inserted_columns['position'],
                        'count': inserted_columns.get('count', 1),
                    }]
                if ops:
                    insert_map = _build_col_map_for_insert(ops, max_col_for_map)
                    if col_map:
                        combined = {}
                        for old_col, mid_col in col_map.items():
                            if mid_col is not None:
                                final_col = insert_map.get(mid_col, mid_col)
                                combined[old_col] = final_col
                        col_map = combined
                    else:
                        col_map = insert_map
                    sys.stderr.write(f"[XML_COL_OPS] Insert-Map für {len(ops)} Operation(en)\n")
            
            if column_order and len(column_order) > 0:
                columns_changed = any(
                    new_idx != old_idx
                    for new_idx, old_idx in enumerate(column_order)
                )
                if columns_changed:
                    reorder_map = _build_col_map_for_reorder(column_order)
                    if col_map:
                        combined = {}
                        for old_col, mid_col in col_map.items():
                            if mid_col is not None:
                                final_col = reorder_map.get(mid_col, mid_col)
                                combined[old_col] = final_col
                        col_map = combined
                    else:
                        col_map = reorder_map
                    sys.stderr.write(f"[XML_COL_OPS] Reorder-Map für {len(column_order)} Spalten\n")
            
            if col_map is None and hidden_columns is None:
                if os.path.normpath(file_path) != os.path.normpath(output_path):
                    shutil.copy2(file_path, output_path)
                return {'success': True, 'outputPath': output_path, 'method': 'xml-col-ops-noop'}
            
            # 6. Mapping anwenden
            modified_files = {}  # ZIP-Pfad → neuer Inhalt (bytes)
            
            if col_map:
                # Sheet-XML — skip_sort=True weil wir nach allen Änderungen
                # gezielt sortieren (Performance: vermeidet doppelte Sortierung)
                needs_reorder_sort = bool(column_order and any(
                    new_idx != old_idx for new_idx, old_idx in enumerate(column_order)))
                sheet_content = _apply_col_map_to_sheet_xml(
                    sheet_content, col_map, skip_sort=True)
                sys.stderr.write(f"[XML_COL_OPS] Sheet-XML angepasst\n")
                
                # Neue Zellen bei Insert einfügen
                inserted_col_info = []  # (new_col_1based, header) für Table-XML
                if inserted_columns:
                    ops = inserted_columns.get('operations', [])
                    if not ops and inserted_columns.get('position') is not None:
                        ops = [{
                            'position': inserted_columns['position'],
                            'count': inserted_columns.get('count', 1),
                            'headers': inserted_columns.get('headers', []),
                        }]
                    
                    # Sammle alle neuen Zellen pro Zeile (effizient)
                    new_cells_by_row = {}  # excel_row → list of cell XML strings
                    
                    for op in ops:
                        position = op['position']
                        count = op.get('count', 1)
                        op_headers = op.get('headers', [])
                        
                        # Header-Zellen (Zeile 1)
                        for i, header in enumerate(op_headers):
                            col_num = position + 1 + i
                            col_letter = _num_to_col_letter(col_num)
                            cell_ref = f"{col_letter}1"
                            escaped = str(header).replace(
                                '&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                            new_cell = (f'<c r="{cell_ref}" t="inlineStr">'
                                        f'<is><t>{escaped}</t></is></c>')
                            new_cells_by_row.setdefault(1, []).append(new_cell)
                            inserted_col_info.append((col_num, header))
                        
                        # Falls keine headers, trotzdem die Position merken
                        if not op_headers:
                            for i in range(count):
                                col_num = position + 1 + i
                                inserted_col_info.append((col_num, f'Column{col_num}'))
                        
                        # Daten-Zellen
                        if data:
                            for row_idx, row_data in enumerate(data):
                                excel_row = row_idx + 2
                                for ci in range(count):
                                    col_idx = position + ci
                                    if col_idx < len(row_data) and row_data[col_idx] is not None:
                                        col_num = position + 1 + ci
                                        col_letter = _num_to_col_letter(col_num)
                                        cell_ref = f"{col_letter}{excel_row}"
                                        val = row_data[col_idx]
                                        if isinstance(val, (int, float)):
                                            new_cell = f'<c r="{cell_ref}"><v>{val}</v></c>'
                                        else:
                                            escaped = str(val).replace(
                                                '&', '&amp;').replace(
                                                '<', '&lt;').replace('>', '&gt;')
                                            new_cell = (f'<c r="{cell_ref}" t="inlineStr">'
                                                        f'<is><t>{escaped}</t></is></c>')
                                        new_cells_by_row.setdefault(excel_row, []).append(new_cell)
                    
                    # Alle neuen Zellen in einem Durchgang einfügen
                    if new_cells_by_row:
                        def _insert_cells_into_row(row_match):
                            row_tag = row_match.group(1)
                            row_content = row_match.group(2)
                            row_num_m = re.search(r'r="(\d+)"', row_tag)
                            if row_num_m:
                                row_num = int(row_num_m.group(1))
                                if row_num in new_cells_by_row:
                                    extra = ''.join(new_cells_by_row[row_num])
                                    row_content = extra + row_content
                            return f'{row_tag}{row_content}</row>'
                        
                        sheet_content = re.sub(
                            r'(<row\s[^>]*>)(.*?)</row>',
                            _insert_cells_into_row,
                            sheet_content, flags=re.DOTALL)
                    
                    # Nach dem Einfügen: Zellen sortieren
                    if needs_reorder_sort:
                        # Bei Reorder + Insert müssen ALLE Zellen sortiert werden,
                        # da die Spalten-Positionen sich beliebig ändern können
                        sheet_content = _sort_cells_in_rows(sheet_content)
                        sys.stderr.write(f"[XML_COL_OPS] Volle Zell-Sortierung nach Reorder+Insert\n")
                    elif new_cells_by_row:
                        # Nur betroffene Zeilen sortieren (Performance)
                        # Bei Insert-only sind die existierenden Zellen durch Rechts-Shift
                        # noch in korrekter Reihenfolge — nur Zeilen mit neuen Zellen
                        # müssen sortiert werden.
                        target_rows = set(new_cells_by_row.keys())
                        
                        def _sort_cells_in_target_row(row_match):
                            row_tag = row_match.group(1)
                            row_content = row_match.group(2)
                            row_num_m = re.search(r'r="(\d+)"', row_tag)
                            if not row_num_m or int(row_num_m.group(1)) not in target_rows:
                                return row_match.group(0)  # Unverändert
                            
                            cells = []
                            for cm in re.finditer(
                                r'<c\s[^>]*r="([A-Z]+)(\d+)"[^>]*/>'
                                r'|<c\s[^>]*r="([A-Z]+)(\d+)"[^>]*>.*?</c>',
                                row_content, re.DOTALL):
                                col_letter = cm.group(1) or cm.group(3)
                                col_num = _col_letter_to_num(col_letter)
                                cells.append((col_num, cm.group(0)))
                            if not cells:
                                return row_match.group(0)
                            cells.sort(key=lambda x: x[0])
                            sorted_content = ''.join(cell_xml for _, cell_xml in cells)
                            return f'{row_tag}{sorted_content}</row>'
                        
                        sheet_content = re.sub(
                            r'(<row\s[^>]*>)(.*?)</row>',
                            _sort_cells_in_target_row,
                            sheet_content, flags=re.DOTALL)
                    
                    sys.stderr.write(f"[XML_COL_OPS] {len(new_cells_by_row)} Zeilen mit neuen Zellen eingefügt\n")
                elif needs_reorder_sort:
                    # Bei Reorder (ohne Insert) müssen alle Zellen sortiert werden,
                    # da die Spalten-Positionen sich beliebig ändern können
                    sheet_content = _sort_cells_in_rows(sheet_content)
                    sys.stderr.write(f"[XML_COL_OPS] Volle Zell-Sortierung nach Spalten-Reorder\n")
                
                # Table-XMLs anpassen
                for rid, table_path in table_files.items():
                    if table_path in src_zip.namelist():
                        table_content = src_zip.read(table_path).decode('utf-8')
                        table_content = _apply_col_map_to_table_xml(
                            table_content, col_map, headers,
                            inserted_col_info=inserted_col_info if inserted_col_info else None)
                        modified_files[table_path] = table_content.encode('utf-8')
                        sys.stderr.write(f"[XML_COL_OPS] Table {table_path} angepasst\n")
                
                # Drawing-XMLs anpassen
                for rid, drawing_path in drawing_files.items():
                    if drawing_path in src_zip.namelist():
                        drawing_content = src_zip.read(drawing_path).decode('utf-8')
                        drawing_content = _apply_col_map_to_drawing_xml(
                            drawing_content, col_map)
                        modified_files[drawing_path] = drawing_content.encode('utf-8')
                        sys.stderr.write(f"[XML_COL_OPS] Drawing {drawing_path} angepasst\n")
                
                # workbook.xml: definedNames anpassen
                wb_xml_content = src_zip.read('xl/workbook.xml').decode('utf-8')
                new_wb_xml = _apply_col_map_to_workbook_xml(
                    wb_xml_content, col_map, sheet_name)
                if new_wb_xml != wb_xml_content:
                    modified_files['xl/workbook.xml'] = new_wb_xml.encode('utf-8')
                    sys.stderr.write(f"[XML_COL_OPS] workbook.xml definedNames angepasst\n")
            
            # Hidden Columns
            if hidden_columns is not None:
                hidden_set = set(hidden_columns)
                
                def _fix_col_hidden(m):
                    col_el = m.group(0)
                    min_m2 = re.search(r'min="(\d+)"', col_el)
                    max_m2 = re.search(r'max="(\d+)"', col_el)
                    if not min_m2 or not max_m2:
                        return col_el
                    col_min2 = int(min_m2.group(1))
                    col_max2 = int(max_m2.group(1))
                    all_hidden = all((c - 1) in hidden_set
                                     for c in range(col_min2, col_max2 + 1))
                    if all_hidden:
                        if 'hidden="1"' not in col_el:
                            col_el = col_el.replace('/>', ' hidden="1"/>')
                    else:
                        col_el = re.sub(r'\s*hidden="1"', '', col_el)
                    return col_el
                
                sheet_content = re.sub(r'<col\s[^>]*/>', _fix_col_hidden, sheet_content)
            
            modified_files[sheet_zip_path] = sheet_content.encode('utf-8')
            
            # 7. ZIP-to-ZIP: Original-Einträge 1:1 kopieren, nur modifizierte ersetzen
            with zipfile.ZipFile(temp_output, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    if item.filename.endswith('/'):
                        continue
                    if item.filename.startswith('__MACOSX') or \
                       item.filename.endswith('.DS_Store') or \
                       item.filename.split('/')[-1].startswith('._'):
                        continue
                    
                    if item.filename in modified_files:
                        item.compress_type = zipfile.ZIP_DEFLATED
                        dst_zip.writestr(item, modified_files[item.filename])
                    else:
                        data_bytes = src_zip.read(item.filename)
                        dst_zip.writestr(item, data_bytes)
        
        # An Zielort verschieben
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(temp_output, output_path)
        
        sys.stderr.write(f"[XML_COL_OPS] Erfolgreich: {output_path}\n")
        return {'success': True, 'outputPath': output_path, 'method': 'xml-col-ops'}
    
    except Exception as e:
        if os.path.exists(temp_output):
            os.remove(temp_output)
        sys.stderr.write(f"[XML_COL_OPS] Fehler: {e}\n")
        import traceback
        traceback.print_exc(file=sys.stderr)
        raise
