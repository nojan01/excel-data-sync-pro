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
import io
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
# ZEILEN-HILFSFUNKTIONEN
# =============================================================================

def _remap_row_in_ref(ref, row_map):
    """
    Wendet ein Zeilen-Mapping auf eine Zellreferenz an.
    row_map: dict von alter Excel-Zeilennummer → neuer Excel-Zeilennummer.
    Zeilen die nicht im Mapping sind → ref wird entfernt (None zurückgegeben).
    """
    dollar1, col_letter, dollar2, row_num = _parse_cell_ref(ref)
    if col_letter is None:
        return ref  # Kein gültiger Cell-Ref
    new_row = row_map.get(row_num)
    if new_row is None:
        return None  # Zeile wurde gelöscht
    return f"{dollar1}{col_letter}{dollar2}{new_row}"


def _remap_row_range_ref(range_ref, row_map):
    """
    Wendet ein Zeilen-Mapping auf eine Range-Referenz an (z.B. 'A2:J100').
    Bei gelöschten Randzeilen wird der Bereich auf die verbleibenden Zeilen geschrumpft.
    Gibt None zurück nur wenn ALLE Zeilen im Bereich gelöscht wurden.
    """
    if ':' not in range_ref:
        return _remap_row_in_ref(range_ref, row_map)
    parts = range_ref.split(':')
    new_start = _remap_row_in_ref(parts[0], row_map)
    new_end = _remap_row_in_ref(parts[1], row_map)
    if new_start is not None and new_end is not None:
        return f"{new_start}:{new_end}"

    # Mindestens eine Randzeile wurde gelöscht →
    # tatsächlichen Bereich aller verbleibenden Zeilen ermitteln
    d1s, col_s, d2s, row_s = _parse_cell_ref(parts[0])
    d1e, col_e, d2e, row_e = _parse_cell_ref(parts[1])
    if col_s is None or col_e is None:
        return None

    mapped_rows = []
    for r in range(row_s, row_e + 1):
        nr = row_map.get(r)
        if nr is not None:
            mapped_rows.append(nr)

    if not mapped_rows:
        return None  # Alle Zeilen im Bereich gelöscht

    new_min_row = min(mapped_rows)
    new_max_row = max(mapped_rows)
    new_start_ref = f"{d1s}{col_s}{d2s}{new_min_row}"
    new_end_ref = f"{d1e}{col_e}{d2e}{new_max_row}"
    return f"{new_start_ref}:{new_end_ref}"


def _remap_row_sqref(sqref, row_map):
    """
    Wendet ein Zeilen-Mapping auf sqref an (Space-separated Ranges).
    Entfernt Ranges deren Zeilen alle gelöscht wurden.
    """
    ranges = sqref.split()
    new_ranges = []
    for r in ranges:
        new_r = _remap_row_range_ref(r, row_map)
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
    
    Die Frontend-Positionen sind FINALE Positionen (nach allen Einfügungen).
    Algorithmus: Durchlaufe FINALE Positionen 1..N, überspringe Insert-Positionen,
    weise alten Spalten die verbleibenden Positionen zu.
    """
    # Sammle alle FINALEN Insert-Positionen (1-basiert)
    insert_positions = set()
    for op in insert_operations:
        pos_1based = op['position'] + 1  # 0-basiert → 1-basiert
        count = op.get('count', 1)
        for i in range(count):
            insert_positions.add(pos_1based + i)
    
    total_inserts = len(insert_positions)
    max_final = max_col + total_inserts
    
    col_map = {}
    old_col = 1
    for final_pos in range(1, max_final + 1):
        if final_pos in insert_positions:
            continue  # Position gehört einer neuen Spalte
        if old_col > max_col:
            break
        col_map[old_col] = final_pos
        old_col += 1
    
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
# ZEILEN-MAPPING BUILDER
# =============================================================================

def _build_row_maps(deleted_rows_0based, row_order, max_excel_row, inserted_rows=None):
    """
    Baut Zeilen-Mappings für Zeilen-Löschung, -Verschiebung und -Einfügung.
    
    Args:
        deleted_rows_0based: Liste von 0-basierten Daten-Indizes (0 = Excel-Zeile 2).
        row_order: Liste wo row_order[new_pos] = original_data_idx (0-basiert).
                   -1 = eingefügte Zeile (wird in Schritt 3 behandelt).
                   Gelöschte Indizes werden automatisch ignoriert.
                   Kann None/leer sein bei reiner Löschung.
        max_excel_row: Maximale Excel-Zeilennummer im Sheet.
        inserted_rows: Dict mit 'operations': [{position: int, count: int}]
                       Position ist 0-basierter Daten-Index im FINALEN Grid.
    
    Returns:
        (data_row_map, meta_row_map, insert_positions)
        - data_row_map: {old_excel_row: final_excel_row} — für sheetData (delete + reorder + insert-shift)
        - meta_row_map: {old_excel_row: after_delete_excel_row} — für Metadaten (nur delete)
        - insert_positions: set of final Excel-Zeilennummern für eingefügte leere Zeilen
    """
    deleted_set = set(deleted_rows_0based) if deleted_rows_0based else set()
    
    # Schritt 1: Delete-Map (alte Zeile → Zeile nach Löschung, ohne Reorder)
    meta_row_map = {1: 1}  # Header bleibt immer Zeile 1
    remaining = []  # (original_data_idx, intermediate_excel_row)
    new_row = 2
    for orig_idx in range(max_excel_row - 1):  # Datenzeilen (ohne Header)
        if orig_idx not in deleted_set:
            old_excel = orig_idx + 2
            meta_row_map[old_excel] = new_row
            remaining.append((orig_idx, new_row))
            new_row += 1
    
    # Schritt 2: Reorder auf Delete anwenden
    if row_order and len(row_order) > 0:
        data_row_map = {1: 1}
        # row_order enthält ORIGINAL 0-basierte Daten-Indizes (aus Frontend rowMapping)
        # -1 = eingefügte Zeilen (werden in Schritt 3 per insert-shift behandelt)
        # Lookup: original_data_idx → old_excel_row (nur für überlebende Zeilen)
        orig_to_old_excel = {orig_idx: orig_idx + 2 for orig_idx, _ in remaining}
        
        final_pos = 0  # Konsekutiver Positions-Zähler (ohne -1 Einträge)
        for entry in row_order:
            if entry == -1:
                continue  # Eingefügte Zeile, wird in Schritt 3 behandelt
            if entry in orig_to_old_excel:
                old_excel = orig_to_old_excel[entry]
                final_excel = final_pos + 2
                data_row_map[old_excel] = final_excel
                final_pos += 1
    else:
        # Nur Löschen, kein Reorder → data_row_map = meta_row_map (Kopie)
        data_row_map = dict(meta_row_map)
    
    # Schritt 3: Inserts — existierende Zeilen nach unten verschieben
    insert_positions = set()  # Finale Excel-Zeilennummern der eingefügten Zeilen
    if inserted_rows:
        ops = inserted_rows.get('operations', [])
        if not ops and inserted_rows.get('position') is not None:
            ops = [{'position': inserted_rows['position'], 'count': inserted_rows.get('count', 1)}]
        
        if ops:
            # Sammle alle Insert-Positionen (0-basierte Daten-Indizes im finalen Grid)
            insert_data_indices = set()
            for op in ops:
                pos = op['position']
                cnt = op.get('count', 1)
                for i in range(cnt):
                    insert_data_indices.add(pos + i)
            
            # Berechne finale Excel-Zeilen für alle existierenden Zeilen (Shift durch Inserts)
            # Sortiere die aktuellen finalen Positionen
            existing_final = sorted(set(data_row_map.values()) - {1})  # Ohne Header
            
            # Baue finale Positionen: gehe durch 0-basierte finale Indizes
            # und verteile existierende + eingefügte Zeilen
            new_data_row_map = {1: 1}
            existing_iter = iter(existing_final)
            # Reverse-Map: final_excel → old_excel
            rev_map = {v: k for k, v in data_row_map.items() if k != 1}
            
            final_row = 2  # Excel-Zeile startet bei 2 (Header = 1)
            existing_idx = 0  # Zähler für existierende Zeilen
            total_existing = len(existing_final)
            data_idx = 0  # 0-basierter Daten-Index im finalen Grid
            
            while existing_idx < total_existing or data_idx in insert_data_indices:
                if data_idx in insert_data_indices:
                    insert_positions.add(final_row)
                    final_row += 1
                    data_idx += 1
                elif existing_idx < total_existing:
                    old_final = existing_final[existing_idx]
                    old_excel = rev_map[old_final]
                    new_data_row_map[old_excel] = final_row
                    final_row += 1
                    data_idx += 1
                    existing_idx += 1
                else:
                    break
            
            data_row_map = new_data_row_map
            sys.stderr.write(f"[BUILD_ROW_MAPS] {len(insert_positions)} Insert-Positionen: "
                             f"{sorted(insert_positions)[:10]}...\n")
    
    return data_row_map, meta_row_map, insert_positions


# =============================================================================
# XML-TRANSFORMATIONEN
# =============================================================================

def _cleanup_orphaned_shared_formulas(sheet_xml):
    """
    Entfernt verwaiste Shared-Formula-Slaves und -Masters mit ungültigem ref.
    
    Wenn der Master einer shared formula gelöscht wurde (z.B. Spalte/Zeile gelöscht),
    bleiben Slave-Zellen mit <f t="shared" si="X"/> oder <f t="shared" si="X"></f>
    ohne Master-Definition.
    Excel erkennt das als inkonsistent → "Zellinformationen" Reparatur.
    
    Lösung: Orphaned slaves → <f> Element entfernen (cached <v> bleibt erhalten).
    Ebenfalls: Masters deren ref-Range nicht mehr die eigene Zelle enthält → entfernen.
    """
    # Sammle alle si-Werte die einen Master haben (haben ref= UND si= Attribut)
    defined_si = set()
    for m in re.finditer(r'<f\s[^>]*?\bsi="(\d+)"[^>]*?\bref="', sheet_xml):
        defined_si.add(m.group(1))
    for m in re.finditer(r'<f\s[^>]*?\bref="[^"]*"[^>]*?\bsi="(\d+)"', sheet_xml):
        defined_si.add(m.group(1))
    
    def _check_orphan(m):
        f_tag = m.group(0)
        si_match = re.search(r'\bsi="(\d+)"', f_tag)
        if si_match and si_match.group(1) not in defined_si:
            return ''  # Verwaister Slave → entfernen
        return f_tag
    
    # 1. Selbstschließende Slaves: <f t="shared" si="X"/>
    sheet_xml = re.sub(r'<f\s[^>]*?\bt="shared"[^>]*?/>', _check_orphan, sheet_xml)
    
    # 2. Nicht-selbstschließende Slaves: <f t="shared" si="X"></f> oder <f t="shared" si="X"> </f>
    #    (Manche Excel-Writer erzeugen leere <f>...</f> statt <f/>)
    sheet_xml = re.sub(r'<f\s[^>]*?\bt="shared"[^>]*?>\s*</f>', _check_orphan, sheet_xml)
    
    # 3. Shared-Formula Masters ohne gültige Slaves prüfen:
    #    Wenn nach dem Cleanup ein Master-si existiert aber KEINE Slaves mehr
    #    dafür vorhanden sind UND der ref-Bereich nur eine einzige Zelle ist,
    #    konvertiere den Master in eine normale Formel (ref + si + t entfernen).
    #    → Das verhindert, dass Excel die Shared-Formel als inkonsistent erkennt.
    for si_val in list(defined_si):
        # Zähle verbleibende Nutzungen dieses si (Master + Slaves)
        si_uses = len(re.findall(rf'\bsi="{re.escape(si_val)}"', sheet_xml))
        if si_uses == 1:
            # Nur der Master selbst übrig — shared formula hat keine Slaves mehr
            # → Konvertiere Master in normale Formel: entferne t="shared", si="X", ref="..."
            def _demote_lonely_master(m):
                f_tag = m.group(0)
                si_m = re.search(rf'\bsi="{re.escape(si_val)}"', f_tag)
                if not si_m:
                    return f_tag
                # Entferne t="shared", si="X", ref="..." Attribute
                f_tag = re.sub(r'\s*\bt="shared"', '', f_tag)
                f_tag = re.sub(r'\s*\bsi="\d+"', '', f_tag)
                f_tag = re.sub(r'\s*\bref="[^"]*"', '', f_tag)
                return f_tag
            # Matche Master-<f> (hat Formeltext, also nicht selbstschließend)
            sheet_xml = re.sub(
                r'<f\s[^>]*?\bt="shared"[^>]*?>.*?</f>',
                _demote_lonely_master, sheet_xml, flags=re.DOTALL)
    
    return sheet_xml


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
        # <c ...>...</c> — (?<!/) verhindert dass /> einer self-closing Zelle als > gematcht wird
        sheet_xml = re.sub(r'<c\s[^>]*?r="([A-Z]+)\d+"[^>]*(?<!/)>.*?</c>', _remove_deleted_cell, sheet_xml, flags=re.DOTALL)
    
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
    
    # ---- 1b. <f ref="..."> Shared/Array-Formel-Refs remappen ----
    # Shared formulas: <f t="shared" ref="B2:B100" si="0">formula</f>
    # Array formulas:  <f t="array" ref="A1:C3">{formula}</f>
    # Die ref-Attribute definieren den Gültigkeitsbereich der Formel.
    # Nach Spalten-Ops müssen die Spalten-Referenzen angepasst werden,
    # sonst sieht Excel eine Inkonsistenz zwischen ref-Range und tatsächlichen
    # Zell-Positionen → "Zellinformationen" Reparatur.
    def _remap_f_ref(m):
        prefix = m.group(1)   # z.B. '<f t="shared" ref="'
        ref_val = m.group(2)  # z.B. 'B2:B8450'
        new_ref = _remap_range_ref(ref_val, col_map)
        if new_ref is None:
            # Kompletter Formel-Range gelöscht — ganzes <f> Element wird
            # durch _cleanup_orphaned_shared_formulas bereinigt
            return f'{prefix}{ref_val}"'  # Unverändert lassen, Cleanup kommt
        return f'{prefix}{new_ref}"'
    
    sheet_xml = re.sub(r'(<f\s[^>]*?\bref=")([^"]+)"', _remap_f_ref, sheet_xml)
    
    # ---- 1c. Verwaiste Shared-Formula-Slaves bereinigen ----
    # Wenn der Master einer shared formula in einer gelöschten Spalte lag,
    # bleiben Slave-Zellen mit <f t="shared" si="X"/> ohne Master-Definition.
    # Excel erkennt das als inkonsistent → "Zellinformationen" Reparatur.
    sheet_xml = _cleanup_orphaned_shared_formulas(sheet_xml)
    
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
            # NICHT '' zurückgeben! Das entfernt nur den öffnenden Tag,
            # lässt aber ">...rules...</conditionalFormatting>" als invalides XML.
            # Stattdessen leeres sqref setzen → Cleanup-Regex entfernt das GANZE Element.
            return '<conditionalFormatting sqref=""'
        return f'<conditionalFormatting sqref="{new_sqref}"{rest}'
    
    sheet_xml = re.sub(r'<conditionalFormatting\s+sqref="([^"]+)"([^>]*)', _remap_cf, sheet_xml)
    # Entferne CF-Elemente mit leerem sqref (inkl. ggf. weiterer Attribute wie pivot="0")
    sheet_xml = re.sub(
        r'<conditionalFormatting\s+sqref=""[^>]*>.*?</conditionalFormatting>',
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
        
        # sortCondition/sortState ref-Bereiche anpassen
        # WICHTIG: Ganze Elemente matchen um orphaned XML-Fragmente zu vermeiden!
        # Vorher: Regex matchte nur bis ref="..." und return '' ließ '/>' orphaned zurück.
        
        # Schritt 1: sortCondition — ganzes self-closing Element matchen
        def _remap_sort_condition(m):
            full = m.group(0)
            ref = m.group(1)
            new_ref = _remap_range_ref(ref, col_map)
            if new_ref is None:
                return ''  # Ganzes Element sauber entfernen
            return full.replace(f'ref="{ref}"', f'ref="{new_ref}"')
        
        sheet_xml = re.sub(
            r'<sortCondition\s[^>]*?ref="([^"]+)"[^>]*/>\s*',
            _remap_sort_condition, sheet_xml)
        
        # Schritt 2: sortState — ganzen Block matchen + leeren sortState entfernen
        def _remap_sort_state_block(m):
            full = m.group(0)
            ref = m.group(1)
            new_ref = _remap_range_ref(ref, col_map)
            if new_ref is None:
                return ''  # Ganzes Element sauber entfernen
            result = full.replace(f'ref="{ref}"', f'ref="{new_ref}"')
            # Keine sortConditions mehr → leeren sortState entfernen
            if '<sortCondition' not in result:
                return ''
            return result
        
        # sortState mit Content
        sheet_xml = re.sub(
            r'<sortState\s[^>]*?ref="([^"]+)"[^>]*>.*?</sortState>\s*',
            _remap_sort_state_block, sheet_xml, flags=re.DOTALL)
        # sortState self-closing (defensiv)
        sheet_xml = re.sub(
            r'<sortState\s[^>]*?ref="([^"]+)"[^>]*/>\s*',
            _remap_sort_state_block, sheet_xml)
        
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
    
    # Non-self-closing: <dataValidation ...>..children..</dataValidation>
    sheet_xml = re.sub(
        r'<dataValidation\s[^>]*sqref="([^"]+)"[^>]*>.*?</dataValidation>',
        _remap_dv, sheet_xml, flags=re.DOTALL)
    # Self-closing: <dataValidation .../>
    sheet_xml = re.sub(
        r'<dataValidation\s[^>]*sqref="([^"]+)"[^>]*/\s*>',
        _remap_dv, sheet_xml)
    # dataValidations count aktualisieren (analog zu mergeCells)
    dv_count = len(re.findall(r'<dataValidation\s', sheet_xml))
    sheet_xml = re.sub(r'<dataValidations\s+count="\d+"', f'<dataValidations count="{dv_count}"', sheet_xml)
    sheet_xml = re.sub(r'<dataValidations\s+count="0"\s*>\s*</dataValidations>', '', sheet_xml)
    
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
    
    # ---- DIAGNOSTIK: XML-Validierung nach Col-Ops ----
    try:
        import sys as _sys
        # 1. AutoFilter Block dumpen
        _af_block = re.search(r'<autoFilter[\s>].*?(?:</autoFilter>|/>)', sheet_xml, re.DOTALL)
        if _af_block:
            _text = _af_block.group(0)
            if len(_text) > 3000:
                _text = _text[:1500] + "\n... TRUNCATED ...\n" + _text[-500:]
            _sys.stderr.write(f"[COL_DIAG] autoFilter block ({len(_af_block.group(0))} chars):\n{_text}\n")
        else:
            _sys.stderr.write("[COL_DIAG] Kein <autoFilter> gefunden\n")
        
        # 2. Check: orphaned sortState/sortCondition Fragmente
        for _orphan_pat, _name in [
            (r'(?<![<\w])/>',  'orphaned />'),
            (r'</sortState>', 'closing </sortState> without opening'),
            (r'</sortCondition>', 'closing </sortCondition> without opening'),
        ]:
            for _om in re.finditer(_orphan_pat, sheet_xml):
                _ctx_start = max(0, _om.start() - 80)
                _ctx_end = min(len(sheet_xml), _om.end() + 30)
                _ctx = sheet_xml[_ctx_start:_ctx_end].replace('\n', '\\n')
                _sys.stderr.write(f"[COL_DIAG] WARNUNG: {_name} at pos {_om.start()}: ...{_ctx}...\n")
        
        # 3. Check: Shared-Formula-Konsistenz
        _sf_masters = set()
        for _sfm in re.finditer(r'<f\s[^>]*?t="shared"[^>]*?si="(\d+)"[^>]*?ref="', sheet_xml):
            _sf_masters.add(_sfm.group(1))
        _orphan_slaves = 0
        for _sfs in re.finditer(r'<f\s+t="shared"\s+si="(\d+)"\s*/>', sheet_xml):
            if _sfs.group(1) not in _sf_masters:
                _orphan_slaves += 1
        if _orphan_slaves > 0:
            _sys.stderr.write(f"[COL_DIAG] WARNUNG: {_orphan_slaves} orphaned shared formula slaves!\n")
        else:
            _sys.stderr.write(f"[COL_DIAG] Shared formulas OK (masters: {len(_sf_masters)})\n")
        
        # 4. Erste 3 Zeilen dumpen zur Zell-Prüfung
        _row_count = 0
        for _rm in re.finditer(r'<row\s[^>]*>.*?</row>', sheet_xml, re.DOTALL):
            _row_count += 1
            if _row_count <= 3:
                _row_text = _rm.group(0)
                if len(_row_text) > 500:
                    _row_text = _row_text[:500] + "..."
                _sys.stderr.write(f"[COL_DIAG] Row {_row_count}: {_row_text}\n")
            if _row_count > 3:
                break
        _sys.stderr.write(f"[COL_DIAG] Diagnostik abgeschlossen\n")
    except Exception as _diag_err:
        import sys as _sys
        _sys.stderr.write(f"[COL_DIAG] Diagnostik-Fehler: {_diag_err}\n")
    
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
        r'(<row\s[^>]*(?<!/)>)(.*?)</row>',
        _sort_cells_in_row,
        sheet_xml,
        flags=re.DOTALL
    )


# =============================================================================
# ZEILEN-XML-TRANSFORMATION
# =============================================================================

def _apply_row_map_to_sheet_xml(sheet_xml, data_row_map, meta_row_map, insert_positions=None):
    """
    Wendet Zeilen-Mappings auf alle relevanten Elemente im Worksheet-XML an.
    
    Zwei separate Maps:
    - data_row_map: für <sheetData> Zeilen (delete + reorder)
    - meta_row_map: für Metadaten wie mergeCells, CF, autoFilter (nur delete, kein reorder)
      → Parity mit openpyxl-Pfad, der Merges/CF nicht für Reorder anpasst.
    
    Args:
        data_row_map: {old_excel_row: final_excel_row}
        meta_row_map: {old_excel_row: after_delete_excel_row}
        insert_positions: set of final Excel-Zeilennummern für eingefügte leere Zeilen
    """
    
    # ---- 0. <dimension ref="A1:J100"> aktualisieren ----
    def _remap_dim(m):
        ref = m.group(1)
        new_ref = _remap_row_range_ref(ref, meta_row_map)
        if new_ref is None:
            return ''
        return f'<dimension ref="{new_ref}"/>'
    sheet_xml = re.sub(r'<dimension\s+ref="([^"]+)"\s*/>', _remap_dim, sheet_xml)
    
    # ---- 1. <sheetData> Zeilen verarbeiten ----
    sd_match = re.search(r'(<sheetData[^>]*>)(.*?)(</sheetData>)', sheet_xml, re.DOTALL)
    if sd_match:
        sd_open = sd_match.group(1)
        sd_content = sd_match.group(2)
        sd_close = sd_match.group(3)
        
        new_rows = []
        for row_m in re.finditer(
                r'<row\s[^>]*?r="(\d+)"[^>]*(?:/>|>.*?</row>)',
                sd_content, re.DOTALL):
            old_row = int(row_m.group(1))
            new_row = data_row_map.get(old_row)
            if new_row is None:
                continue  # Zeile gelöscht
            
            row_xml = row_m.group(0)
            
            # r="X" in <row> Tag aktualisieren (nur in der <row>-Eröffnung)
            row_xml = re.sub(
                r'(<row\s[^>]*?)r="\d+"',
                f'\\1r="{new_row}"',
                row_xml, count=1)
            
            # spans entfernen (Excel berechnet neu)
            row_xml = re.sub(r'\s*spans="[^"]*"', '', row_xml)
            
            # Alle <c r="AX"> Zell-Referenzen auf neue Zeilennummer
            row_xml = re.sub(
                r'(<c\s[^>]*?r="[A-Z]+)\d+"',
                f'\\g<1>{new_row}"',
                row_xml)
            
            new_rows.append((new_row, row_xml))
        
        # Sortieren nach neuer Zeilennummer
        new_rows.sort(key=lambda x: x[0])
        
        # Leere Zeilen für Insert-Positionen einfügen
        if insert_positions:
            for ins_row in sorted(insert_positions):
                new_rows.append((ins_row, f'<row r="{ins_row}"></row>'))
            new_rows.sort(key=lambda x: x[0])
        
        new_sd = sd_open + ''.join(xml for _, xml in new_rows) + sd_close
        sheet_xml = sheet_xml[:sd_match.start()] + new_sd + sheet_xml[sd_match.end():]
    
    # ---- 1b. Verwaiste Shared-Formula-Slaves bereinigen ----
    sheet_xml = _cleanup_orphaned_shared_formulas(sheet_xml)
    
    # ---- 2. <mergeCell ref="A1:C5"> ----
    def _remap_row_merge(m):
        ref = m.group(1)
        new_ref = _remap_row_range_ref(ref, meta_row_map)
        if new_ref is None:
            return ''
        return f'<mergeCell ref="{new_ref}"/>'
    sheet_xml = re.sub(r'<mergeCell\s+ref="([^"]+)"\s*/>', _remap_row_merge, sheet_xml)
    merge_count = len(re.findall(r'<mergeCell\s', sheet_xml))
    sheet_xml = re.sub(r'<mergeCells\s+count="\d+"', f'<mergeCells count="{merge_count}"', sheet_xml)
    sheet_xml = re.sub(r'<mergeCells\s+count="0"\s*>\s*</mergeCells>', '', sheet_xml)
    
    # ---- 3. <conditionalFormatting sqref="A2:J100"> ----
    def _remap_row_cf(m):
        sqref = m.group(1)
        rest = m.group(2)
        new_sqref = _remap_row_sqref(sqref, meta_row_map)
        if new_sqref is None:
            return '<conditionalFormatting sqref=""'
        return f'<conditionalFormatting sqref="{new_sqref}"{rest}'
    sheet_xml = re.sub(r'<conditionalFormatting\s+sqref="([^"]+)"([^>]*)', _remap_row_cf, sheet_xml)
    sheet_xml = re.sub(
        r'<conditionalFormatting\s+sqref=""[^>]*>.*?</conditionalFormatting>',
        '', sheet_xml, flags=re.DOTALL)
    
    # ---- 4. <autoFilter ref="A1:J100"> ----
    def _remap_row_af(m):
        ref = m.group(1)
        new_ref = _remap_row_range_ref(ref, meta_row_map)
        if new_ref is None:
            return ''
        return f'<autoFilter ref="{new_ref}"'
    sheet_xml = re.sub(r'<autoFilter\s+ref="([^"]+)"', _remap_row_af, sheet_xml)
    
    # ---- 5. <hyperlink ref="A2"> ----
    def _remap_row_hl(m):
        full = m.group(0)
        ref = m.group(1)
        new_ref = _remap_row_in_ref(ref, meta_row_map)
        if new_ref is None:
            return ''
        return full.replace(f'ref="{ref}"', f'ref="{new_ref}"')
    sheet_xml = re.sub(r'<hyperlink\s[^>]*ref="([^"]+)"[^>]*/>', _remap_row_hl, sheet_xml)
    
    # ---- 6. <dataValidation sqref="B2:B100"> ----
    def _remap_row_dv(m):
        full = m.group(0)
        sqref = m.group(1)
        new_sqref = _remap_row_sqref(sqref, meta_row_map)
        if new_sqref is None:
            return ''
        return full.replace(f'sqref="{sqref}"', f'sqref="{new_sqref}"')
    sheet_xml = re.sub(
        r'<dataValidation\s[^>]*sqref="([^"]+)"[^>]*>.*?</dataValidation>',
        _remap_row_dv, sheet_xml, flags=re.DOTALL)
    sheet_xml = re.sub(
        r'<dataValidation\s[^>]*sqref="([^"]+)"[^>]*/\s*>',
        _remap_row_dv, sheet_xml)
    dv_count = len(re.findall(r'<dataValidation\s', sheet_xml))
    sheet_xml = re.sub(r'<dataValidations\s+count="\d+"', f'<dataValidations count="{dv_count}"', sheet_xml)
    sheet_xml = re.sub(r'<dataValidations\s+count="0"\s*>\s*</dataValidations>', '', sheet_xml)
    
    # ---- 7. <xm:sqref> in <extLst> (x14 CF, Sparklines etc.) ----
    def _remap_row_xm_sqref(m):
        sqref = m.group(1)
        new_sqref = _remap_row_sqref(sqref, meta_row_map)
        if new_sqref is None:
            return ''
        return f'<xm:sqref>{new_sqref}</xm:sqref>'
    sheet_xml = re.sub(r'<xm:sqref>([^<]+)</xm:sqref>', _remap_row_xm_sqref, sheet_xml)
    
    # ---- 7b. <xm:f> Formeln ----
    def _remap_row_xm_f(m):
        formula = m.group(1)
        if re.match(r'^[^(]+!?\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$', formula):
            if '!' in formula:
                sheet_prefix, ref_part = formula.rsplit('!', 1)
                new_ref = (_remap_row_range_ref(ref_part, meta_row_map) if ':' in ref_part
                           else _remap_row_in_ref(ref_part, meta_row_map))
                if new_ref is None:
                    return ''
                return f'<xm:f>{sheet_prefix}!{new_ref}</xm:f>'
            else:
                new_ref = (_remap_row_range_ref(formula, meta_row_map) if ':' in formula
                           else _remap_row_in_ref(formula, meta_row_map))
                if new_ref is None:
                    return ''
                return f'<xm:f>{new_ref}</xm:f>'
        return m.group(0)
    sheet_xml = re.sub(r'<xm:f>([^<]+)</xm:f>', _remap_row_xm_f, sheet_xml)
    
    # ---- 8. <f ref="A2:A100"> (Shared/Array Formula Ranges) ----
    # Diese verwenden meta_row_map weil Formel-Ranges Metadaten sind
    def _remap_f_ref(m):
        prefix = m.group(1)
        ref = m.group(2)
        new_ref = _remap_row_range_ref(ref, meta_row_map)
        if new_ref is None:
            return m.group(0)
        return f'{prefix}ref="{new_ref}"'
    sheet_xml = re.sub(r'(<f\s[^>]*?)ref="([^"]+)"', _remap_f_ref, sheet_xml)
    
    # ---- 9. <sortState> / <sortCondition> ----
    def _remap_row_sort(m):
        prefix = m.group(1)
        ref = m.group(2)
        new_ref = _remap_row_range_ref(ref, meta_row_map)
        if new_ref is None:
            return ''
        return f'{prefix}ref="{new_ref}"'
    sheet_xml = re.sub(r'(<sortState\s[^>]*?)ref="([^"]+)"', _remap_row_sort, sheet_xml)
    sheet_xml = re.sub(r'(<sortCondition\s[^>]*?)ref="([^"]+)"', _remap_row_sort, sheet_xml)
    
    return sheet_xml


def _apply_row_map_to_table_xml(table_xml, meta_row_map):
    """
    Wendet ein Zeilen-Mapping auf eine Table-XML an.
    Aktualisiert ref/autoFilter Ranges (max row).
    tableColumns bleiben unverändert (Spalten-Level).
    """
    def _remap_table_row_range(m):
        prefix = m.group(1)
        ref = m.group(2)
        new_ref = _remap_row_range_ref(ref, meta_row_map)
        if new_ref is None:
            return m.group(0)
        return f'{prefix}ref="{new_ref}"'
    
    table_xml = re.sub(r'(<table\s[^>]*?)ref="([^"]+)"', _remap_table_row_range, table_xml)
    table_xml = re.sub(r'(<autoFilter\s[^>]*?)ref="([^"]+)"', _remap_table_row_range, table_xml)
    return table_xml


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
    
    # sortCondition/sortState ref-Bereiche anpassen
    # WICHTIG: Ganze Elemente matchen um orphaned XML zu vermeiden
    
    # Schritt 1: sortCondition — ganzes self-closing Element
    def _remap_sort_condition_tbl(m):
        full = m.group(0)
        ref_in_el = re.search(r'ref="([^"]+)"', full)
        if not ref_in_el:
            return full
        ref = ref_in_el.group(1)
        if ':' not in ref:
            return full
        start, end = ref.split(':')
        d1s, col_s, d2s, row_s = _parse_cell_ref(start)
        d1e, col_e, d2e, row_e = _parse_cell_ref(end)
        if col_s is None or col_e is None:
            return full
        sc = _col_letter_to_num(col_s)
        ec = _col_letter_to_num(col_e)
        nc_s = col_map.get(sc)
        nc_e = col_map.get(ec)
        if nc_s is None:
            for c in range(sc, ec + 1):
                nc_s = col_map.get(c)
                if nc_s:
                    break
        if nc_e is None:
            for c in range(ec, sc - 1, -1):
                nc_e = col_map.get(c)
                if nc_e:
                    break
        if nc_s is None or nc_e is None:
            return ''  # Ganzes Element sauber entfernen
        new_start = f"{d1s}{_num_to_col_letter(nc_s)}{d2s}{row_s}"
        new_end = f"{d1e}{_num_to_col_letter(nc_e)}{d2e}{row_e}"
        return full.replace(f'ref="{ref}"', f'ref="{new_start}:{new_end}"')
    
    table_xml = re.sub(
        r'<sortCondition\s[^>]*?ref="[^"]*"[^>]*/>\s*',
        _remap_sort_condition_tbl, table_xml)
    
    # Schritt 2: sortState — ganzen Block + leeren sortState entfernen
    def _remap_sort_state_tbl(m):
        full = m.group(0)
        ref_in_el = re.search(r'ref="([^"]+)"', full)
        if not ref_in_el:
            return full
        ref = ref_in_el.group(1)
        if ':' not in ref:
            return full
        start, end = ref.split(':')
        d1s, col_s, d2s, row_s = _parse_cell_ref(start)
        d1e, col_e, d2e, row_e = _parse_cell_ref(end)
        if col_s is None or col_e is None:
            return full
        sc = _col_letter_to_num(col_s)
        ec = _col_letter_to_num(col_e)
        nc_s = col_map.get(sc)
        nc_e = col_map.get(ec)
        if nc_s is None:
            for c in range(sc, ec + 1):
                nc_s = col_map.get(c)
                if nc_s:
                    break
        if nc_e is None:
            for c in range(ec, sc - 1, -1):
                nc_e = col_map.get(c)
                if nc_e:
                    break
        if nc_s is None or nc_e is None:
            return ''
        new_start = f"{d1s}{_num_to_col_letter(nc_s)}{d2s}{row_s}"
        new_end = f"{d1e}{_num_to_col_letter(nc_e)}{d2e}{row_e}"
        result = full.replace(f'ref="{ref}"', f'ref="{new_start}:{new_end}"')
        if '<sortCondition' not in result:
            return ''
        return result
    
    # sortState mit Content
    table_xml = re.sub(
        r'<sortState\s[^>]*?ref="[^"]*"[^>]*>.*?</sortState>\s*',
        _remap_sort_state_tbl, table_xml, flags=re.DOTALL)
    # sortState self-closing
    table_xml = re.sub(
        r'<sortState\s[^>]*?ref="[^"]*"[^>]*/>\s*',
        _remap_sort_state_tbl, table_xml)
    
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
                                 headers=None, data=None, return_bytes=False,
                                 strip_row_hidden=False, hidden_rows=None,
                                 source_bytes=None):
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
        strip_row_hidden: Wenn True, werden hidden-Attribute von <row> Tags
                         VOR den Spalten-Ops entfernt (werden danach separat re-applied)
        hidden_rows: Liste von 0-basierten Zeilenindizes zum Verstecken.
                    Wenn angegeben, werden hidden-Attribute NACH den Spalten-Ops
                    direkt im selben ZIP-Durchgang gesetzt (kein extra ZIP-Pass).
    
    Returns:
        Dict mit success und outputPath
    """
    sys.stderr.write(f"[XML_COL_OPS] Start für Sheet '{sheet_name}'\n")
    sys.stderr.write(f"[XML_COL_OPS] deleted={deleted_columns}, inserted={inserted_columns is not None}, "
                     f"reorder={column_order is not None}, hidden={hidden_columns}\n")
    
    MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    
    temp_output = output_path + '.tmp' if not return_bytes else None
    
    try:
        src_stream = source_bytes if source_bytes is not None else file_path
        with zipfile.ZipFile(src_stream, 'r') as src_zip:
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
            
            # 2b. Hidden-Attribute von <row> Tags entfernen (werden später re-applied)
            # Analog zum manuellen Workaround: Zeilen einblenden → Spalten-Ops → wieder ausblenden
            if strip_row_hidden:
                _before_len = len(sheet_content)
                sheet_content = re.sub(r'(<row\s[^>]*?)\s+hidden="[^"]*"', r'\1', sheet_content)
                _stripped = _before_len - len(sheet_content)
                if _stripped > 0:
                    sys.stderr.write(f"[XML_COL_OPS] Row-hidden Attribute entfernt ({_stripped} Bytes Differenz)\n")
            
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
                        
                        if rtype.endswith('/table'):
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
                return {'success': True, 'outputPath': output_path, 'method': 'xml-col-ops-noop',
                        'has_slicers': False}
            
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
                    
                    # Frontend-Positionen sind FINALE Positionen (nach allen Einfügungen).
                    # Kein kumulativer Offset nötig — position + 1 ist direkt die 1-basierte Spalte.
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
                            r'(<row\s[^>]*(?<!/)>)(.*?)</row>',
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
                            r'(<row\s[^>]*(?<!/)>)(.*?)</row>',
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
                
                # Comments-XMLs (xl/comments*.xml) anpassen:
                # <comment ref="A1"> Zell-Referenzen müssen remapped werden
                for cname in src_zip.namelist():
                    if re.match(r'xl/comments\d*\.xml$', cname):
                        comments_xml = src_zip.read(cname).decode('utf-8')
                        def _remap_comment_ref(m):
                            prefix = m.group(1)
                            ref = m.group(2)
                            new_ref = _remap_col_in_ref(ref, col_map)
                            if new_ref is None:
                                return ''  # Kommentar für gelöschte Spalte entfernen
                            return f'{prefix}ref="{new_ref}"'
                        new_comments = re.sub(
                            r'(<comment\s[^>]*?)ref="([A-Z]+\d+)"',
                            _remap_comment_ref, comments_xml)
                        if new_comments != comments_xml:
                            modified_files[cname] = new_comments.encode('utf-8')
                            sys.stderr.write(f"[XML_COL_OPS] {cname} Kommentar-Refs angepasst\n")
            
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
            
            # 6b. Hidden Rows direkt im selben Durchgang anwenden
            # Statt separatem ZIP-to-ZIP Pass (_apply_hidden_rows_to_xlsx)
            # werden hidden-Attribute hier in-memory gesetzt.
            if hidden_rows:
                _hr_set = set(hidden_rows)
                _hr_counts = [0, 0]  # [added, unchanged]
                
                def _apply_row_hidden(m):
                    row_tag = m.group(0)
                    r_m = re.search(r'\br="(\d+)"', row_tag)
                    if not r_m:
                        return row_tag
                    row_num = int(r_m.group(1))
                    # 0-basiert in hidden_rows, Datenzeilen ab row 2 (row 1 = Header)
                    row_idx = row_num - 2
                    
                    if row_idx not in _hr_set:
                        return row_tag
                    
                    # hidden="1" einfügen (Attribut wurde bereits in 2b gestrippt)
                    if row_tag.rstrip().endswith('/>'):
                        base = row_tag.rstrip()[:-2].rstrip()
                        result = base + ' hidden="1"/>'
                    else:
                        # Opening tag: vor > einfügen
                        result = row_tag.rstrip().rstrip('>') + ' hidden="1">'
                    _hr_counts[0] += 1
                    return result
                
                sheet_content = re.sub(r'<row\s[^>]*?/?>', _apply_row_hidden, sheet_content)
                sys.stderr.write(f"[XML_COL_OPS] Hidden rows applied: {_hr_counts[0]} rows hidden "
                                 f"(von {len(hidden_rows)} angefordert)\n")
            
            modified_files[sheet_zip_path] = sheet_content.encode('utf-8')
            
            # 7. calcChain.xml entfernen: Enthält Zell-Referenzen die nach
            # Spaltenverschiebungen stale werden → Excel-Reparatur.
            # Excel regeneriert calcChain automatisch beim Öffnen.
            skip_files = set()
            namelist = src_zip.namelist()
            if 'xl/calcChain.xml' in namelist:
                skip_files.add('xl/calcChain.xml')
                sys.stderr.write(f"[XML_COL_OPS] calcChain.xml wird übersprungen (wird von Excel regeneriert)\n")
            
            # Content_Types: calcChain-Override entfernen (sonst verweist es auf fehlende Datei)
            # workbook.xml.rels: calcChain-Relationship entfernen
            if skip_files:
                ct_path = '[Content_Types].xml'
                if ct_path in namelist:
                    ct_xml = (modified_files[ct_path].decode('utf-8')
                              if ct_path in modified_files
                              else src_zip.read(ct_path).decode('utf-8'))
                    ct_orig = ct_xml
                    ct_xml = re.sub(
                        r'<Override\s[^>]*PartName="/xl/calcChain\.xml"[^>]*/>\s*',
                        '', ct_xml)
                    if ct_xml != ct_orig:
                        modified_files[ct_path] = ct_xml.encode('utf-8')
                        sys.stderr.write(f"[XML_COL_OPS] [Content_Types].xml: calcChain-Override entfernt\n")
                
                wb_rels_path = 'xl/_rels/workbook.xml.rels'
                if wb_rels_path in namelist:
                    wb_rels = (modified_files[wb_rels_path].decode('utf-8')
                               if wb_rels_path in modified_files
                               else src_zip.read(wb_rels_path).decode('utf-8'))
                    wb_rels_orig = wb_rels
                    wb_rels = re.sub(
                        r'<Relationship\s[^>]*Target="calcChain\.xml"[^>]*/>\s*',
                        '', wb_rels)
                    if wb_rels != wb_rels_orig:
                        modified_files[wb_rels_path] = wb_rels.encode('utf-8')
                        sys.stderr.write(f"[XML_COL_OPS] workbook.xml.rels: calcChain-Relationship entfernt\n")
            
            # 7b. Slicer-Erkennung (mit bereits geladenen Daten — kein extra ZIP-Lesen)
            has_slicers = any(
                n.startswith('xl/slicerCaches/') or n.startswith('xl/slicers/')
                for n in namelist)
            if not has_slicers:
                # Prüfe in bereits geladenen/modifizierten Inhalten
                for check_path in ['[Content_Types].xml', 'xl/workbook.xml']:
                    content = None
                    if check_path in modified_files:
                        content = modified_files[check_path].decode('utf-8')
                    elif check_path in namelist:
                        content = src_zip.read(check_path).decode('utf-8')
                    if content and 'slicer' in content.lower():
                        has_slicers = True
                        break
            if not has_slicers and 'slicer' in sheet_content.lower():
                has_slicers = True
            
            sys.stderr.write(f"[XML_COL_OPS] Slicer erkannt: {has_slicers}\n")
            
            # 8. ZIP-to-ZIP: Original-Einträge 1:1 kopieren, nur modifizierte ersetzen
            # Bei return_bytes: In BytesIO schreiben statt auf Disk
            dst_target = io.BytesIO() if return_bytes else temp_output
            with zipfile.ZipFile(dst_target, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    if item.filename.endswith('/'):
                        continue
                    if item.filename.startswith('__MACOSX') or \
                       item.filename.endswith('.DS_Store') or \
                       item.filename.split('/')[-1].startswith('._'):
                        continue
                    if item.filename in skip_files:
                        continue
                    
                    if item.filename in modified_files:
                        item.compress_type = zipfile.ZIP_DEFLATED
                        dst_zip.writestr(item, modified_files[item.filename])
                    else:
                        data_bytes = src_zip.read(item.filename)
                        dst_zip.writestr(item, data_bytes)
        
        if return_bytes:
            dst_target.seek(0)
            sys.stderr.write(f"[XML_COL_OPS] Erfolgreich (in-memory, {dst_target.getbuffer().nbytes} bytes)\n")
            return {'success': True, 'outputPath': output_path, 'method': 'xml-col-ops',
                    'has_slicers': has_slicers, 'zip_bytes': dst_target}
        
        # An Zielort verschieben
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(temp_output, output_path)
        
        sys.stderr.write(f"[XML_COL_OPS] Erfolgreich: {output_path}\n")
        return {'success': True, 'outputPath': output_path, 'method': 'xml-col-ops',
                'has_slicers': has_slicers}
    
    except Exception as e:
        if temp_output and os.path.exists(temp_output):
            os.remove(temp_output)
        sys.stderr.write(f"[XML_COL_OPS] Fehler: {e}\n")
        import traceback
        traceback.print_exc(file=sys.stderr)
        raise


# =============================================================================
# DIREKTE XML-ZEILENOPERATIONEN (ZIP-to-ZIP)
# =============================================================================

def direct_xml_row_operations(file_path, output_path, sheet_name,
                              deleted_rows=None, row_order=None,
                              hidden_rows=None, inserted_rows=None,
                              source_bytes=None, return_bytes=False):
    """
    Führt Zeilenoperationen direkt auf dem XML durch (ZIP-to-ZIP).
    
    KEIN openpyxl-Roundtrip → alle Strukturen bleiben intakt:
    - Namespaces, Slicers, Drawings, Media, RichData, External Links
    - Tables, SharedStrings, Styles, Conditional Formatting
    
    Analog zu direct_xml_column_operations, aber für Zeilen.
    
    Args:
        file_path: Quelldatei (.xlsx)
        output_path: Zieldatei (.xlsx)
        sheet_name: Name des Sheets
        deleted_rows: Liste von 0-basierten Daten-Indizes zum Löschen
        row_order: Liste wo row_order[new_pos] = old_pos_after_delete (0-basiert)
        hidden_rows: Liste von 0-basierten Daten-Indizes zum Verstecken
        inserted_rows: Dict mit 'operations': [{position: int, count: int}]
        source_bytes: BytesIO mit ZIP-Daten (wenn None, wird file_path gelesen)
        return_bytes: Wenn True, wird BytesIO statt Datei zurückgegeben
    
    Returns:
        Dict mit success, outputPath, method, has_slicers (und zip_bytes bei return_bytes)
    """
    sys.stderr.write(f"[XML_ROW_OPS] Start für Sheet '{sheet_name}'\n")
    sys.stderr.write(f"[XML_ROW_OPS] deleted={len(deleted_rows) if deleted_rows else 0}, "
                     f"reorder={row_order is not None}, "
                     f"hidden={len(hidden_rows) if hidden_rows else 0}\n")
    
    MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    
    temp_output = output_path + '.row_tmp' if not return_bytes else None
    
    try:
        src_stream = source_bytes if source_bytes is not None else file_path
        with zipfile.ZipFile(src_stream, 'r') as src_zip:
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
            
            sys.stderr.write(f"[XML_ROW_OPS] Sheet-ZIP-Pfad: {sheet_zip_path}\n")
            
            # 2. Lese Sheet-XML
            sheet_content = src_zip.read(sheet_zip_path).decode('utf-8')
            
            # 3. Maximale Zeile aus Sheet ermitteln
            max_row_in_sheet = 1
            for rm in re.finditer(r'<row\s[^>]*?r="(\d+)"', sheet_content):
                row_num = int(rm.group(1))
                if row_num > max_row_in_sheet:
                    max_row_in_sheet = row_num
            
            sys.stderr.write(f"[XML_ROW_OPS] Max Zeile im Sheet: {max_row_in_sheet}\n")
            
            # 4. Finde zugehörige Table-Dateien
            sheet_rels_path = sheet_zip_path.replace(
                'worksheets/', 'worksheets/_rels/') + '.rels'
            table_files = {}  # rId → ZIP-Pfad
            
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
                except ET.ParseError as pe:
                    sys.stderr.write(f"[XML_ROW_OPS] WARNUNG: Sheet-Rels Parse-Fehler: {pe}\n")
            
            sys.stderr.write(f"[XML_ROW_OPS] Tables: {list(table_files.values())}\n")
            
            # 5. Zeilen-Mappings bauen
            data_row_map, meta_row_map, insert_positions = _build_row_maps(
                deleted_rows, row_order, max_row_in_sheet, inserted_rows)
            
            surviving_count = len([v for k, v in data_row_map.items() if k != 1])
            if insert_positions:
                sys.stderr.write(f"[XML_ROW_OPS] {len(insert_positions)} Zeilen werden eingefügt\n")
            sys.stderr.write(f"[XML_ROW_OPS] {surviving_count} Zeilen überleben "
                             f"(von {max_row_in_sheet - 1} Datenzeilen)\n")
            
            # 6. Mapping auf Sheet-XML anwenden
            modified_files = {}
            
            sheet_content = _apply_row_map_to_sheet_xml(
                sheet_content, data_row_map, meta_row_map, insert_positions)
            sys.stderr.write(f"[XML_ROW_OPS] Sheet-XML angepasst\n")
            
            # 7. Hidden Rows anwenden (auf die neuen Zeilennummern)
            if hidden_rows:
                # hidden_rows sind 0-basierte Daten-Indizes → Excel-Zeilen = idx + 2
                # Wir müssen sie durch data_row_map auf finale Positionen mappen
                hidden_excel_rows = set()
                for idx in hidden_rows:
                    old_excel = idx + 2
                    new_excel = data_row_map.get(old_excel)
                    if new_excel is not None:
                        hidden_excel_rows.add(new_excel)
                
                if hidden_excel_rows:
                    def _set_row_hidden(m):
                        row_tag = m.group(0)
                        r_m = re.search(r'r="(\d+)"', row_tag)
                        if r_m and int(r_m.group(1)) in hidden_excel_rows:
                            if 'hidden="1"' not in row_tag:
                                # Vor dem schließenden > oder /> einfügen
                                if row_tag.endswith('/>'):
                                    return row_tag[:-2] + ' hidden="1"/>'
                                elif row_tag.endswith('>'):
                                    return row_tag[:-1] + ' hidden="1">'
                            return row_tag
                        else:
                            # Nicht-hidden Zeilen: hidden="1" entfernen falls vorhanden
                            return re.sub(r'\s*hidden="1"', '', row_tag)
                    
                    sheet_content = re.sub(r'<row\s[^>]*?>', _set_row_hidden, sheet_content)
                    sys.stderr.write(f"[XML_ROW_OPS] {len(hidden_excel_rows)} Zeilen versteckt\n")
            
            modified_files[sheet_zip_path] = sheet_content.encode('utf-8')
            
            # 8. Table-XMLs anpassen (Zeilen-Bereich kürzen)
            for rid, table_path in table_files.items():
                if table_path in src_zip.namelist():
                    table_content = src_zip.read(table_path).decode('utf-8')
                    new_table = _apply_row_map_to_table_xml(table_content, meta_row_map)
                    if new_table != table_content:
                        modified_files[table_path] = new_table.encode('utf-8')
                        sys.stderr.write(f"[XML_ROW_OPS] Table {table_path} angepasst\n")
            
            # 9. Comments anpassen (xl/comments*.xml)
            for cname in src_zip.namelist():
                if re.match(r'xl/comments\d*\.xml$', cname):
                    comments_xml = src_zip.read(cname).decode('utf-8')
                    def _remap_comment_row_ref(m):
                        prefix = m.group(1)
                        ref = m.group(2)
                        new_ref = _remap_row_in_ref(ref, data_row_map)
                        if new_ref is None:
                            return ''  # Kommentar für gelöschte Zeile entfernen
                        return f'{prefix}ref="{new_ref}"'
                    new_comments = re.sub(
                        r'(<comment\s[^>]*?)ref="([A-Z]+\d+)"',
                        _remap_comment_row_ref, comments_xml)
                    if new_comments != comments_xml:
                        modified_files[cname] = new_comments.encode('utf-8')
                        sys.stderr.write(f"[XML_ROW_OPS] {cname} Kommentar-Refs angepasst\n")
            
            # 10. workbook.xml: definedNames mit Zeilen-Referenzen anpassen
            wb_xml_content = src_zip.read('xl/workbook.xml').decode('utf-8')
            # definedName-Werte wie "Sheet1!$A$1:$J$100" → Zeilen aktualisieren
            def _remap_defined_name_rows(m):
                full = m.group(0)
                value = m.group(1)
                parts = value.split(',')
                new_parts = []
                changed = False
                for part in parts:
                    part = part.strip()
                    if '!' in part:
                        sheet_prefix, ref_part = part.rsplit('!', 1)
                        if ':' in ref_part:
                            new_ref = _remap_row_range_ref(ref_part, meta_row_map)
                        else:
                            new_ref = _remap_row_in_ref(ref_part, meta_row_map)
                        if new_ref is not None:
                            new_parts.append(f"{sheet_prefix}!{new_ref}")
                            if new_ref != ref_part:
                                changed = True
                        else:
                            changed = True
                    else:
                        new_parts.append(part)
                if changed and new_parts:
                    new_value = ','.join(new_parts)
                    return full.replace(f'>{value}<', f'>{new_value}<')
                return full
            
            new_wb_xml = re.sub(
                r'<definedName\s[^>]*>([^<]+)</definedName>',
                _remap_defined_name_rows, wb_xml_content)
            if new_wb_xml != wb_xml_content:
                modified_files['xl/workbook.xml'] = new_wb_xml.encode('utf-8')
                sys.stderr.write(f"[XML_ROW_OPS] workbook.xml definedNames angepasst\n")
            
            # 11. calcChain.xml entfernen (Excel regeneriert beim Öffnen)
            skip_files = set()
            namelist = src_zip.namelist()
            if 'xl/calcChain.xml' in namelist:
                skip_files.add('xl/calcChain.xml')
                sys.stderr.write(f"[XML_ROW_OPS] calcChain.xml wird übersprungen\n")
            
            if skip_files:
                ct_path = '[Content_Types].xml'
                if ct_path in namelist:
                    ct_xml = (modified_files[ct_path].decode('utf-8')
                              if ct_path in modified_files
                              else src_zip.read(ct_path).decode('utf-8'))
                    ct_orig = ct_xml
                    ct_xml = re.sub(
                        r'<Override\s[^>]*PartName="/xl/calcChain\.xml"[^>]*/>\s*',
                        '', ct_xml)
                    if ct_xml != ct_orig:
                        modified_files[ct_path] = ct_xml.encode('utf-8')
                
                wb_rels_path = 'xl/_rels/workbook.xml.rels'
                if wb_rels_path in namelist:
                    wb_rels = (modified_files[wb_rels_path].decode('utf-8')
                               if wb_rels_path in modified_files
                               else src_zip.read(wb_rels_path).decode('utf-8'))
                    wb_rels_orig = wb_rels
                    wb_rels = re.sub(
                        r'<Relationship\s[^>]*Target="calcChain\.xml"[^>]*/>\s*',
                        '', wb_rels)
                    if wb_rels != wb_rels_orig:
                        modified_files[wb_rels_path] = wb_rels.encode('utf-8')
            
            # 12. Slicer-Erkennung
            has_slicers = any(
                n.startswith('xl/slicerCaches/') or n.startswith('xl/slicers/')
                for n in namelist)
            if not has_slicers:
                for check_path in ['[Content_Types].xml', 'xl/workbook.xml']:
                    content = None
                    if check_path in modified_files:
                        content = modified_files[check_path].decode('utf-8')
                    elif check_path in namelist:
                        content = src_zip.read(check_path).decode('utf-8')
                    if content and 'slicer' in content.lower():
                        has_slicers = True
                        break
            if not has_slicers and 'slicer' in sheet_content.lower():
                has_slicers = True
            
            # 13. ZIP-to-ZIP: Original-Einträge 1:1 kopieren, nur modifizierte ersetzen
            # Bei return_bytes: In BytesIO schreiben statt auf Disk
            dst_target = io.BytesIO() if return_bytes else temp_output
            with zipfile.ZipFile(dst_target, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    if item.filename.endswith('/'):
                        continue
                    if item.filename.startswith('__MACOSX') or \
                       item.filename.endswith('.DS_Store') or \
                       item.filename.split('/')[-1].startswith('._'):
                        continue
                    if item.filename in skip_files:
                        continue
                    
                    if item.filename in modified_files:
                        item.compress_type = zipfile.ZIP_DEFLATED
                        dst_zip.writestr(item, modified_files[item.filename])
                    else:
                        data_bytes = src_zip.read(item.filename)
                        dst_zip.writestr(item, data_bytes)
        
        if return_bytes:
            dst_target.seek(0)
            sys.stderr.write(f"[XML_ROW_OPS] Erfolgreich (in-memory, {dst_target.getbuffer().nbytes} bytes)\n")
            return {'success': True, 'outputPath': output_path, 'method': 'xml-row-ops',
                    'has_slicers': has_slicers, 'zip_bytes': dst_target}
        
        # An Zielort verschieben
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(temp_output, output_path)
        
        sys.stderr.write(f"[XML_ROW_OPS] Erfolgreich: {output_path}\n")
        return {'success': True, 'outputPath': output_path, 'method': 'xml-row-ops',
                'has_slicers': has_slicers}
    
    except Exception as e:
        if temp_output and os.path.exists(temp_output):
            os.remove(temp_output)
        sys.stderr.write(f"[XML_ROW_OPS] Fehler: {e}\n")
        import traceback
        traceback.print_exc(file=sys.stderr)
        raise
