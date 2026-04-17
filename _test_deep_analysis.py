#!/usr/bin/env python3
"""
Tiefe Analyse: Vergleicht Original vs. Output auf Zellebene.
Prüft Formeln, Styles, und XML-Struktur-Konsistenz.
"""
import sys, os, re, zipfile
from xml.etree import ElementTree as ET

INPUT_FILE = os.path.expanduser("~/Desktop/2026-02-19 DEFENCE&SPACE MVMS Master Asset List.xlsx")
OUTPUT_FILE = "/var/folders/8t/qdvgsldx12lf_nyrxjgvwqvc0000gn/T/test_col_delete_output.xlsx"

MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def get_sheet_xml(xlsx_path, sheet_name):
    """Extrahiert sheet XML und Pfad."""
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        wb = z.read('xl/workbook.xml').decode('utf-8')
        # Finde Sheet rId
        for m in re.finditer(r'<sheet\s[^>]*name="([^"]+)"[^>]*/>', wb):
            if m.group(1) == sheet_name:
                rid_m = re.search(r'r:id="([^"]+)"', m.group(0))
                if rid_m:
                    rid = rid_m.group(1)
                    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
                    for rm in re.finditer(r'Id="' + re.escape(rid) + r'"[^>]*Target="([^"]+)"', rels):
                        path = 'xl/' + rm.group(1).lstrip('/')
                        return z.read(path).decode('utf-8'), path
    return None, None

print("=" * 70)
print("ANALYSE 1: ORIGINAL-DATEI - sheet1.xml Struktur")
print("=" * 70)

with zipfile.ZipFile(INPUT_FILE, 'r') as z:
    orig_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
    
    # Top-Level Elemente im Sheet
    # Finde alle direkten Kinder von <worksheet>
    print("\nTop-Level Elemente im Original sheet1.xml:")
    # Suche nach allen Tags direkt nach <worksheet ...>
    for m in re.finditer(r'<(\w+:?\w+)\s', orig_xml):
        tag = m.group(1)
    
    # Besser: Suche spezifische Elemente
    elements = [
        ('dimension', r'<dimension\s[^>]*/>'),
        ('sheetViews', r'<sheetViews'),
        ('sheetFormatPr', r'<sheetFormatPr'),
        ('cols', r'<cols>'),
        ('sheetData', r'<sheetData'),
        ('autoFilter', r'<autoFilter\s'),
        ('mergeCells', r'<mergeCells'),
        ('conditionalFormatting', r'<conditionalFormatting\s'),
        ('dataValidations', r'<dataValidations'),
        ('hyperlinks', r'<hyperlinks'),
        ('pageMargins', r'<pageMargins'),
        ('pageSetup', r'<pageSetup'),
        ('headerFooter', r'<headerFooter'),
        ('drawing', r'<drawing\s'),
        ('tableParts', r'<tableParts'),
        ('extLst', r'<extLst'),
        ('ignoredErrors', r'<ignoredErrors'),
        ('sheetProtection', r'<sheetProtection'),
    ]
    for name, pattern in elements:
        count = len(re.findall(pattern, orig_xml))
        if count > 0:
            # Kurzes Beispiel
            m = re.search(pattern + r'[^>]*', orig_xml)
            snippet = m.group(0)[:120] if m else ""
            print(f"  ✓ {name} ({count}x): {snippet}")
    
    # Formeln die Spalte D oder H referenzieren
    print("\n--- Formeln im Original ---")
    formulas = re.findall(r'<f>([^<]+)</f>', orig_xml)
    print(f"Gesamt Formeln: {len(formulas)}")
    formulas_with_D = [f for f in formulas if re.search(r'\bD\d', f) or 'D:' in f]
    formulas_with_H = [f for f in formulas if re.search(r'\bH\d', f) or 'H:' in f]
    print(f"Formeln mit Spalte D: {len(formulas_with_D)}")
    print(f"Formeln mit Spalte H: {len(formulas_with_H)}")
    if formulas_with_D[:5]:
        for f in formulas_with_D[:5]:
            print(f"  Beispiel D: {f[:100]}")
    if formulas_with_H[:5]:
        for f in formulas_with_H[:5]:
            print(f"  Beispiel H: {f[:100]}")
    
    # styles.xml - max styleId
    styles_xml = z.read('xl/styles.xml').decode('utf-8')
    xf_count = len(re.findall(r'<xf\s', styles_xml))
    # cellXfs count
    cellxfs_m = re.search(r'<cellXfs\s+count="(\d+)"', styles_xml)
    cellxfs_count = int(cellxfs_m.group(1)) if cellxfs_m else 0
    print(f"\nStyles: cellXfs count={cellxfs_count}")
    
    # Prüfe Style-Indices in sheet1
    max_style = -1
    style_issues = 0
    for cm in re.finditer(r'<c\s[^>]*s="(\d+)"', orig_xml):
        s_idx = int(cm.group(1))
        if s_idx > max_style:
            max_style = s_idx
        if s_idx >= cellxfs_count:
            style_issues += 1
    print(f"Max Style-Index in sheet1: {max_style}")
    print(f"Style-Index Fehler (>= cellXfs count): {style_issues}")
    
    # Sheet rels
    if 'xl/worksheets/_rels/sheet1.xml.rels' in z.namelist():
        srels = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
        print(f"\n--- Original sheet1.xml.rels ---")
        for rm in re.finditer(r'<Relationship\s[^>]*/>', srels):
            print(f"  {rm.group(0)[:150]}")

print()
print("=" * 70)
print("ANALYSE 2: OUTPUT-DATEI - Detaillierte Prüfung")
print("=" * 70)

with zipfile.ZipFile(OUTPUT_FILE, 'r') as z:
    out_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
    
    # Formeln im Output
    print("\n--- Formeln im Output ---")
    formulas = re.findall(r'<f>([^<]+)</f>', out_xml)
    print(f"Gesamt Formeln: {len(formulas)}")
    formulas_with_H = [f for f in formulas if re.search(r'[^A-Z]H\d|^H\d|\$H\$|H:', f)]
    formulas_with_D = [f for f in formulas if re.search(r'[^A-Z]D\d|^D\d|\$D\$|D:', f)]
    print(f"Formeln mit Spalte H: {len(formulas_with_H)}")
    if formulas_with_H:
        for f in formulas_with_H[:10]:
            # Finde die Zelle
            idx = out_xml.find(f'<f>{f}</f>')
            if idx > 0:
                # Suche rückwärts nach dem r= Attribut
                before = out_xml[max(0, idx - 200):idx]
                ref_m = re.search(r'r="([A-Z]+\d+)"[^>]*$', before)
                ref = ref_m.group(1) if ref_m else "?"
                print(f"  ❌ Zelle {ref}: {f[:100]}")
    
    formulas_with_excl = [f for f in formulas if '#REF' in f or '#NULL' in f]
    print(f"Formeln mit #REF!/#NULL!: {len(formulas_with_excl)}")
    
    # Style-Index Check
    styles_xml = z.read('xl/styles.xml').decode('utf-8')
    cellxfs_m = re.search(r'<cellXfs\s+count="(\d+)"', styles_xml)
    cellxfs_count = int(cellxfs_m.group(1)) if cellxfs_m else 0
    
    max_style = -1
    bad_style_cells = []
    for cm in re.finditer(r'<c\s[^>]*r="([A-Z]+\d+)"[^>]*s="(\d+)"', out_xml):
        ref = cm.group(1)
        s_idx = int(cm.group(2))
        if s_idx > max_style:
            max_style = s_idx
        if s_idx >= cellxfs_count:
            bad_style_cells.append((ref, s_idx))
    
    # Auch umgekehrt: s="N" r="X"
    for cm in re.finditer(r'<c\s[^>]*s="(\d+)"[^>]*r="([A-Z]+\d+)"', out_xml):
        s_idx = int(cm.group(1))
        ref = cm.group(2)
        if s_idx > max_style:
            max_style = s_idx
        if s_idx >= cellxfs_count:
            bad_style_cells.append((ref, s_idx))
    
    print(f"\nStyles: cellXfs count={cellxfs_count}, max style in sheet={max_style}")
    if bad_style_cells:
        print(f"  ❌ {len(bad_style_cells)} Zellen mit ungültigem Style-Index!")
        for ref, idx in bad_style_cells[:10]:
            print(f"     {ref}: s={idx} (max erlaubt: {cellxfs_count - 1})")
    else:
        print("  ✅ Alle Style-Indices gültig")
    
    # Type/Value Consistency
    print("\n--- Cell Type/Value Konsistenz ---")
    type_issues = 0
    for cm in re.finditer(r'<c\s[^>]*t="s"[^>]*>(.*?)</c>', out_xml, re.DOTALL):
        content = cm.group(1)
        v_m = re.search(r'<v>([^<]*)</v>', content)
        if v_m:
            try:
                int(v_m.group(1))
            except ValueError:
                type_issues += 1
                if type_issues <= 3:
                    ref_m = re.search(r'r="([A-Z]+\d+)"', cm.group(0))
                    ref = ref_m.group(1) if ref_m else "?"
                    print(f"  ⚠️ Zelle {ref}: t='s' aber v='{v_m.group(1)[:50]}'")
    print(f"  Type/Value Probleme: {type_issues}")
    
    # Sheet1.xml.rels im Output
    if 'xl/worksheets/_rels/sheet1.xml.rels' in z.namelist():
        srels = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
        print(f"\n--- Output sheet1.xml.rels ---")
        for rm in re.finditer(r'<Relationship\s[^>]*/>', srels):
            snippet = rm.group(0)[:150]
            print(f"  {snippet}")
            # Prüfe ob Target existiert
            target_m = re.search(r'Target="([^"]+)"', rm.group(0))
            if target_m:
                target = target_m.group(1)
                if not target.startswith('http'):
                    resolved = 'xl/worksheets/' + target
                    norm = []
                    for p in resolved.split('/'):
                        if p == '..':
                            if norm:
                                norm.pop()
                        elif p and p != '.':
                            norm.append(p)
                    resolved = '/'.join(norm)
                    if resolved not in z.namelist():
                        print(f"    ❌ ZIEL FEHLT: {resolved}")
    else:
        print(f"\n--- KEINE sheet1.xml.rels im Output ---")
    
    # Prüfe alle Elemente die r:id referenzieren
    print(f"\n--- Referenzen in sheet1.xml (r:id, relationships) ---")
    for rm in re.finditer(r'r:id="([^"]+)"', out_xml):
        rid = rm.group(1)
        if 'xl/worksheets/_rels/sheet1.xml.rels' in z.namelist():
            srels = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
            if f'Id="{rid}"' not in srels:
                # Finde den Kontext
                start = max(0, rm.start() - 50)
                end = min(len(out_xml), rm.end() + 50)
                context = out_xml[start:end]
                print(f"  ❌ {rid} — in sheet1.xml referenziert aber NICHT in .rels: ...{context}...")
    
    # Prüfe cols Element
    cols_m = re.search(r'<cols>(.*?)</cols>', out_xml, re.DOTALL)
    if cols_m:
        print(f"\n--- <cols> Definitionen ---")
        for cm in re.finditer(r'<col\s[^>]*/>', cols_m.group(1)):
            print(f"  {cm.group(0)[:120]}")
    
    # ignoredErrors
    ie_count = len(re.findall(r'<ignoredError\s', out_xml))
    if ie_count:
        print(f"\n--- ignoredErrors: {ie_count} ---")
        for m in re.finditer(r'<ignoredError\s[^>]*/>', out_xml):
            print(f"  {m.group(0)[:150]}")
            # Prüfe sqref auf H
            sqref_m = re.search(r'sqref="([^"]+)"', m.group(0))
            if sqref_m:
                sqref = sqref_m.group(1)
                if 'H' in sqref:
                    print(f"    ❌ ignoredError sqref enthält H: {sqref}")
    
    # RAW: Zähle ALLE Vorkommen von 'H' als Spaltenreferenz
    print(f"\n--- Suche nach H-Spalten-Referenzen im gesamten sheet1.xml ---")
    # Muster: Spalte H in verschiedenen Kontexten
    h_refs = re.findall(r'(?<![A-Z])H(\d+)(?!\d)', out_xml)
    h_dollar = re.findall(r'\$H\$?(\d+)|\$H(?=:|\$)', out_xml)
    h_range = re.findall(r'H\d+:H|:H\d+', out_xml)
    print(f"  H<number> Muster: {len(h_refs)}")
    print(f"  $H$ Muster: {len(h_dollar)}")
    print(f"  H in Ranges: {len(h_range)}")
    
    if h_refs:
        # Zeige Kontext
        for m in re.finditer(r'(?<![A-Z])H(\d+)(?!\d)', out_xml):
            start = max(0, m.start() - 80)
            end = min(len(out_xml), m.end() + 80)
            context = out_xml[start:end].replace('\n', ' ')
            print(f"  Kontext: ...{context}...")
            if len(h_refs) > 5:
                break
