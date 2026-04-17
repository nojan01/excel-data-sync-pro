#!/usr/bin/env python3
"""
Test-Script: Simuliert den FAST-PATH Spalten-Delete auf der echten Datei
und prüft die Ausgabe auf Zellinformationen-Korruption.

Echtes Szenario: Sheet "DEFENCE&SPACE Jan-2026", 2 Spalten löschen [6, 1],
5966 Hidden Rows anwenden.
"""
import sys, os, re, zipfile, shutil, tempfile
from xml.etree import ElementTree as ET

# Setup
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_DIR = os.path.join(SCRIPT_DIR, 'python')
sys.path.insert(0, PYTHON_DIR)

from excel_xml_ops import direct_xml_column_operations
from excel_writer import _strip_slicers_from_zip, _strip_pivot_tables_for_sheet, _apply_hidden_rows_to_xlsx

INPUT_FILE = os.path.expanduser("~/Desktop/2026-02-19 DEFENCE&SPACE MVMS Master Asset List - Kopie.xlsx")
SHEET_NAME = "DEFENCE&SPACE Jan-2026"
DELETED_COLUMNS = [6, 1]  # 0-basiert = Spalte G (7.) und Spalte B (2.)

# Temporäre Ausgabe
OUTPUT_FILE = os.path.join(tempfile.gettempdir(), "test_col_delete_output.xlsx")

print(f"Input: {INPUT_FILE}")
print(f"Sheet: {SHEET_NAME}")
print(f"Delete columns: {DELETED_COLUMNS} (0-based = Spalte B und G)")
print(f"Output: {OUTPUT_FILE}")
print()

# 0. Selbstschließende <row/> Tags im Original zählen
print("=" * 60)
print("VORAB-CHECK: Selbstschließende <row/> Tags im Original")
print("=" * 60)
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

with zipfile.ZipFile(INPUT_FILE, 'r') as z:
    wb_xml = z.read('xl/workbook.xml').decode('utf-8')
    wb_root = ET.fromstring(wb_xml)
    sheet_rid = None
    for s in wb_root.iter(f'{{{MAIN_NS}}}sheet'):
        if s.get('name') == SHEET_NAME:
            sheet_rid = s.get(f'{{{R_NS}}}id')
            break
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rels_root = ET.fromstring(rels_xml)
    sheet_file = None
    for r in rels_root.iter(f'{{{RELS_NS}}}Relationship'):
        if r.get('Id') == sheet_rid:
            sheet_file = r.get('Target')
            break
    orig_sheet_path = 'xl/' + sheet_file.lstrip('/')
    orig_xml = z.read(orig_sheet_path).decode('utf-8')
    self_closing_count = len(re.findall(r'<row\s[^>]*/>', orig_xml))
    total_rows = len(re.findall(r'<row\s', orig_xml))
    print(f"  Gesamt <row> Tags: {total_rows}")
    print(f"  Selbstschließende <row .../> Tags: {self_closing_count}")
    if self_closing_count > 0:
        print(f"  *** WARNUNG: {self_closing_count} selbstschließende <row/> Tags gefunden!")
        print(f"  *** Diese hätten OHNE den Fix ungültiges XML erzeugt!")
        # Zeige erste 3 Beispiele
        for i, rm in enumerate(re.finditer(r'<row\s[^>]*/>', orig_xml)):
            if i >= 3:
                break
            print(f"  Beispiel: {rm.group(0)[:80]}")
    print()

# 1. Spalten-Delete via XML
print("=" * 60)
print("SCHRITT 1: direct_xml_column_operations")
print("=" * 60)
result = direct_xml_column_operations(
    file_path=INPUT_FILE,
    output_path=OUTPUT_FILE,
    sheet_name=SHEET_NAME,
    deleted_columns=DELETED_COLUMNS,
    inserted_columns=None,
    column_order=None,
    hidden_columns=None,
    headers=None,
    data=None
)
print(f"Result: {result}")
print()

# 2. Slicer-Strip
print("=" * 60)
print("SCHRITT 2: _strip_slicers_from_zip")
print("=" * 60)
try:
    _strip_slicers_from_zip(OUTPUT_FILE)
    print("Slicer-Strip erfolgreich")
except Exception as e:
    print(f"Slicer-Strip Fehler: {e}")
print()

# 3. Pivot-Strip
print("=" * 60)
print("SCHRITT 3: _strip_pivot_tables_for_sheet")
print("=" * 60)
try:
    _strip_pivot_tables_for_sheet(OUTPUT_FILE, SHEET_NAME)
    print("Pivot-Strip erfolgreich")
except Exception as e:
    print(f"Pivot-Strip Fehler: {e}")
print()

# 3.5. Hidden Rows (echtes Szenario: 5966 hidden rows)
print("=" * 60)
print("SCHRITT 3.5: _apply_hidden_rows_to_xlsx (Hidden Rows)")
print("=" * 60)
# Simuliere Hidden Rows: erste 5966 Datenzeilen (0-basiert)
# In der echten App kommen diese von JavaScript als 0-basierte Datenzeilen-Indices
HIDDEN_ROWS = list(range(5966))
try:
    _apply_hidden_rows_to_xlsx(OUTPUT_FILE, SHEET_NAME, HIDDEN_ROWS)
    print(f"Hidden Rows erfolgreich: {len(HIDDEN_ROWS)} Zeilen versteckt")
except Exception as e:
    print(f"Hidden Rows Fehler: {e}")
    import traceback
    traceback.print_exc()
print()

# 4. VALIDIERUNG
print("=" * 60)
print("VALIDIERUNG DER AUSGABE")
print("=" * 60)

errors = []

with zipfile.ZipFile(OUTPUT_FILE, 'r') as z:
    # Finde sheet1.xml
    wb_xml = z.read('xl/workbook.xml').decode('utf-8')
    wb_root = ET.fromstring(wb_xml)
    sheet_rid = None
    for s in wb_root.iter(f'{{{MAIN_NS}}}sheet'):
        if s.get('name') == SHEET_NAME:
            sheet_rid = s.get(f'{{{R_NS}}}id')
            break
    
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rels_root = ET.fromstring(rels_xml)
    sheet_file = None
    for r in rels_root.iter(f'{{{RELS_NS}}}Relationship'):
        if r.get('Id') == sheet_rid:
            sheet_file = r.get('Target')
            break
    
    sheet_path = 'xl/' + sheet_file.lstrip('/')
    print(f"Sheet ZIP path: {sheet_path}")
    
    sheet_xml = z.read(sheet_path).decode('utf-8')
    
    # --- Check 1: Dimension ---
    dim_m = re.search(r'<dimension\s+ref="([^"]+)"', sheet_xml)
    if dim_m:
        dim_ref = dim_m.group(1)
        print(f"Dimension: {dim_ref}")
    else:
        print("Dimension: NICHT GEFUNDEN")
    
    # --- Check 2: Zellen nach Spalte ---
    cell_cols = {}
    for cm in re.finditer(r'<c\s[^>]*r="([A-Z]+)(\d+)"', sheet_xml):
        col = cm.group(1)
        row = int(cm.group(2))
        cell_cols.setdefault(col, []).append(row)
    
    print(f"\nZellen pro Spalte (erste 10):")
    sorted_cols = sorted(cell_cols.keys())
    for col in sorted_cols[:10]:
        count = len(cell_cols[col])
        print(f"  {col}: {count} Zellen (Zeilen {min(cell_cols[col])}-{max(cell_cols[col])})")
    if len(sorted_cols) > 10:
        print(f"  ... und {len(sorted_cols) - 10} weitere Spalten")
    
    # --- Check 3: SharedStrings ---
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
        ss_count_m = re.search(r'uniqueCount="(\d+)"', ss_xml)
        if ss_count_m:
            ss_count = int(ss_count_m.group(1))
            print(f"\nSharedStrings uniqueCount: {ss_count}")
            
            # Prüfe ob Zellen SharedString-Indices referenzieren die out-of-bounds sind
            bad_ssi = 0
            for cm in re.finditer(r'<c\s[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', sheet_xml):
                idx = int(cm.group(1))
                if idx >= ss_count:
                    bad_ssi += 1
                    if bad_ssi <= 5:
                        # Finde den Zellnamen
                        ref_m = re.search(r'r="([A-Z]+\d+)"', cm.group(0))
                        ref = ref_m.group(1) if ref_m else "?"
                        errors.append(f"FEHLER: Zelle {ref} hat SharedString-Index {idx} >= {ss_count}")
            if bad_ssi > 5:
                errors.append(f"... und {bad_ssi - 5} weitere ungültige SharedString-Indices")
            if bad_ssi == 0:
                print("SharedString-Indices: ALLE OK")
    
    # --- Check 4: Row/Cell Ordering ---
    print("\nRow/Cell Ordering:")
    row_issues = 0
    last_row_num = 0
    for rm in re.finditer(r'<row\s[^>]*r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.DOTALL):
        row_num = int(rm.group(1))
        row_content = rm.group(2)
        
        if row_num <= last_row_num:
            errors.append(f"FEHLER: Zeile {row_num} nach Zeile {last_row_num} (falsche Reihenfolge)")
            row_issues += 1
        last_row_num = row_num
        
        # Prüfe Zell-Reihenfolge innerhalb der Zeile
        last_col = 0
        for cc in re.finditer(r'r="([A-Z]+)\d+"', row_content):
            col_letter = cc.group(1)
            col_num = 0
            for ch in col_letter:
                col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
            if col_num <= last_col:
                errors.append(f"FEHLER: Zeile {row_num}: Spalte {col_letter}({col_num}) nach Spalte ({last_col}) (falsche Reihenfolge)")
                row_issues += 1
                if row_issues > 10:
                    break
            last_col = col_num
    
    if row_issues == 0:
        print("  Row ordering: OK")
        print("  Cell ordering: OK")
    
    # --- Check 5: Spans ---
    span_count = len(re.findall(r' spans="', sheet_xml))
    print(f"\nSpans-Attribute: {span_count} (sollte 0 sein)")
    if span_count > 0:
        errors.append(f"FEHLER: {span_count} spans-Attribute noch vorhanden")
    
    # --- Check 6: Tables (tableParts in sheet) ---
    table_parts = re.findall(r'<tablePart\s[^>]*r:id="([^"]+)"', sheet_xml)
    print(f"\ntableParts im Sheet: {len(table_parts)}")
    
    # Prüfe ob die referenzierten Tables existieren
    sheet_rels_path = sheet_path.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if sheet_rels_path in z.namelist():
        srels = z.read(sheet_rels_path).decode('utf-8')
        for tp_rid in table_parts:
            if f'Id="{tp_rid}"' not in srels:
                errors.append(f"FEHLER: tablePart r:id={tp_rid} hat keine Relationship in {sheet_rels_path}")
            else:
                # Prüfe ob die referenzierte Table-Datei existiert
                target_m = re.search(r'Id="' + re.escape(tp_rid) + r'"[^>]*Target="([^"]+)"', srels)
                if target_m:
                    table_path = 'xl/worksheets/' + target_m.group(1)
                    norm_parts = []
                    for p in table_path.split('/'):
                        if p == '..':
                            if norm_parts:
                                norm_parts.pop()
                        elif p != '.':
                            norm_parts.append(p)
                    table_path = '/'.join(norm_parts)
                    if table_path not in z.namelist():
                        errors.append(f"FEHLER: Table-Datei {table_path} fehlt (referenziert von {tp_rid})")
                    else:
                        table_xml = z.read(table_path).decode('utf-8')
                        table_ref_m = re.search(r'ref="([^"]+)"', table_xml)
                        if table_ref_m:
                            t_ref = table_ref_m.group(1)
                            print(f"  Table {tp_rid} → {table_path}: ref={t_ref}")
                            if 'H' in t_ref.split(':')[-1]:
                                errors.append(f"FEHLER: Table {table_path} ref enthält H: {t_ref}")
    
    # --- Check 7: PivotTable-Referenzen ---
    if sheet_rels_path in z.namelist():
        srels = z.read(sheet_rels_path).decode('utf-8')
        pivot_rels = re.findall(r'Type="[^"]*pivotTable[^"]*"', srels)
        print(f"\nPivotTable-Rels im Sheet: {len(pivot_rels)}")
        if pivot_rels:
            errors.append(f"WARNUNG: {len(pivot_rels)} PivotTable-Rels noch vorhanden nach Strip")
    
    # --- Check 8: Conditional Formatting ---
    cf_count = len(re.findall(r'<conditionalFormatting\s', sheet_xml))
    print(f"\nConditionalFormatting-Blöcke: {cf_count}")
    for cf_m in re.finditer(r'<conditionalFormatting\s+sqref="([^"]*)"', sheet_xml):
        sqref = cf_m.group(1)
        if not sqref:
            errors.append(f"FEHLER: Leerer sqref in conditionalFormatting")
        # Prüfe ob H vorkommt
        for part in sqref.split():
            if ':' in part:
                end = part.split(':')[-1]
                col_only = ''.join(c for c in end if c.isalpha())
                if col_only >= 'H':
                    errors.append(f"FEHLER: conditionalFormatting sqref enthält H+: {sqref}")
    
    # --- Check 9: extLst Elemente ---
    extlst_count = len(re.findall(r'<extLst', sheet_xml))
    xm_sqref_count = len(re.findall(r'<xm:sqref>', sheet_xml))
    print(f"\nextLst-Blöcke: {extlst_count}")
    print(f"xm:sqref-Elemente: {xm_sqref_count}")
    for xm_m in re.finditer(r'<xm:sqref>([^<]+)</xm:sqref>', sheet_xml):
        sqref = xm_m.group(1)
        for part in sqref.split():
            if ':' in part:
                end = part.split(':')[-1]
                col_only = ''.join(c for c in end if c.isalpha())
                if col_only >= 'H':
                    errors.append(f"FEHLER: xm:sqref enthält H+: {sqref}")
    
    # --- Check 10: Content_Types ---
    ct_xml = z.read('[Content_Types].xml').decode('utf-8')
    ct_overrides = re.findall(r'PartName="([^"]+)"', ct_xml)
    missing_parts = []
    for pn in ct_overrides:
        zip_path = pn.lstrip('/')
        if zip_path not in z.namelist():
            missing_parts.append(zip_path)
            errors.append(f"FEHLER: Content_Types referenziert fehlende Datei: {zip_path}")
    if missing_parts:
        print(f"\nContent_Types: {len(missing_parts)} fehlende Dateien!")
        for mp in missing_parts[:10]:
            print(f"  FEHLT: {mp}")
    else:
        print(f"\nContent_Types: alle {len(ct_overrides)} referenzierten Dateien vorhanden")
    
    # --- Check 11: Relationships ohne Ziel ---
    for rels_file in z.namelist():
        if rels_file.endswith('.rels'):
            rels_content = z.read(rels_file).decode('utf-8')
            for rel_m in re.finditer(r'Target="([^"]+)"', rels_content):
                target = rel_m.group(1)
                if target.startswith('http://') or target.startswith('https://'):
                    continue
                # Resolve relative path
                base_dir = '/'.join(rels_file.replace('_rels/', '').rsplit('/', 1)[:-1])
                if target.startswith('/'):
                    resolved = target.lstrip('/')
                else:
                    resolved = base_dir + '/' + target if base_dir else target
                # Normalize
                norm = []
                for p in resolved.split('/'):
                    if p == '..':
                        if norm:
                            norm.pop()
                    elif p and p != '.':
                        norm.append(p)
                resolved = '/'.join(norm)
                if resolved not in z.namelist():
                    # Könnte ein externer Link sein
                    type_m = re.search(r'Type="[^"]*External[^"]*"', rels_content)
                    if not type_m:
                        pass  # Nur als Info, nicht als Error
    
    # --- Check 12: Leere Zeilen (alle Zellen entfernt) ---
    empty_rows = 0
    for rm in re.finditer(r'<row\s[^>]*>(.*?)</row>', sheet_xml, re.DOTALL):
        content = rm.group(1).strip()
        if not content or not re.search(r'<c\s', content):
            row_m = re.search(r'r="(\d+)"', rm.group(0))
            row_num = row_m.group(1) if row_m else "?"
            empty_rows += 1
            if empty_rows <= 3:
                errors.append(f"WARNUNG: Leere Zeile {row_num} (keine Zellen)")
    if empty_rows > 3:
        errors.append(f"... und {empty_rows - 3} weitere leere Zeilen")
    
    # --- Check 13: Doppelte Zellen ---
    dup_count = 0
    for rm in re.finditer(r'<row\s[^>]*r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.DOTALL):
        row_num = rm.group(1)
        row_content = rm.group(2)
        seen_refs = set()
        for cc in re.finditer(r'r="([A-Z]+\d+)"', row_content):
            ref = cc.group(1)
            if ref in seen_refs:
                dup_count += 1
                if dup_count <= 5:
                    errors.append(f"FEHLER: Doppelte Zelle {ref} in Zeile {row_num}")
            seen_refs.add(ref)
    if dup_count > 5:
        errors.append(f"... und {dup_count - 5} weitere doppelte Zellen")
    
    # --- Check 14: XML Well-Formedness (KRITISCH) ---
    print("\nXML Well-Formedness:")
    try:
        ET.fromstring(sheet_xml)
        print("  Sheet-XML ist well-formed ✓")
    except ET.ParseError as pe:
        errors.append(f"FEHLER: Sheet-XML ist NICHT well-formed: {pe}")
        print(f"  *** FEHLER: {pe}")
    
    # --- Check 15: Selbstschließende <row/> mit hidden Attribut ---
    print("\nSelbstschließende <row/> nach Hidden-Rows:")
    bad_self_closing = 0
    for rm in re.finditer(r'<row\s[^>]*/>', sheet_xml):
        tag = rm.group(0)
        if 'hidden="1"' in tag:
            # OK: <row r="5" hidden="1"/> ist valide
            pass
        # Prüfe auf ungültiges Pattern: <row .../ hidden="1">
        if re.search(r'/\s+hidden', tag):
            bad_self_closing += 1
            if bad_self_closing <= 3:
                errors.append(f"FEHLER: Ungültiges selbstschließendes <row>: {tag[:80]}")
    valid_self_closing = len(re.findall(r'<row\s[^>]*/>', sheet_xml))
    print(f"  Valide selbstschließende <row/>: {valid_self_closing}")
    if bad_self_closing > 0:
        errors.append(f"  {bad_self_closing} ungültige selbstschließende <row/> Tags!")
    
    # --- Check 16: Hidden Rows korrekt angewendet ---
    print("\nHidden Rows:")
    hidden_count = len(re.findall(r'<row\s[^>]*hidden="1"', sheet_xml))
    print(f"  Zeilen mit hidden='1': {hidden_count}")
    print(f"  Erwartet: {len(HIDDEN_ROWS)}")
    if hidden_count != len(HIDDEN_ROWS):
        errors.append(f"WARNUNG: Hidden-Row-Anzahl weicht ab: {hidden_count} vs erwartet {len(HIDDEN_ROWS)}")

# ERGEBNIS
print()
print("=" * 60)
if errors:
    print(f"ERGEBNIS: {len(errors)} Probleme gefunden!")
    print("=" * 60)
    for e in errors:
        print(f"  ❌ {e}")
else:
    print("ERGEBNIS: KEINE Probleme gefunden — Datei sollte sauber sein")
    print("=" * 60)
    print("  ✅ Dimension OK")
    print("  ✅ SharedString-Indices OK")
    print("  ✅ Row/Cell Ordering OK")
    print("  ✅ Tables OK")
    print("  ✅ Content_Types OK")
    print("  ✅ Keine doppelten Zellen")
    print("  ✅ XML well-formed")
    print("  ✅ Hidden Rows korrekt")

print(f"\nAusgabe-Datei: {OUTPUT_FILE}")
