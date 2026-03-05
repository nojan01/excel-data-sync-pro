# Image Handling Investigation — openpyxl Fallback Mode

## Executive Summary

The openpyxl fallback mode has an **extensive image restoration mechanism** (`restore_external_links_from_original()`, ~1200 lines) that is designed to compensate for openpyxl silently dropping images when Pillow is not installed. This mechanism is correctly invoked on **every** `wb.save()` code path. However, several issues can cause image loss despite the restore mechanism:

1. **xlsx-populate post-processing** strips images after restoration (password-protected files)
2. **Clone operations** permanently strip images with no restore path
3. **Multi-pass operations** create fragile chains where Pass 2's "original" may already lack images
4. **ExcelJS writer** has zero image handling code

---

## 1. Architecture Overview

```
JS Bridge (python_bridge.js)
  ├── Copy source → target
  ├── Decrypt via xlsx-populate (if password)    ← POTENTIAL IMAGE LOSS
  ├── Apply pending sheet ops (JSZip)            ← CLONE STRIPS IMAGES
  ├── For each sheet: call Python writer
  │     └── excel_writer.py write_sheet()
  │           ├── FALL 3a: ZIP-to-ZIP (no openpyxl) → IMAGES PRESERVED
  │           ├── FALL 3b: openpyxl → restore chain → IMAGES RESTORED
  │           ├── FALL 1/1.5/1.9: openpyxl → restore chain → IMAGES RESTORED
  │           └── FALL 2: ZIP-ANSATZ → restore chain → IMAGES RESTORED
  └── Re-encrypt via xlsx-populate (if password) ← POTENTIAL IMAGE LOSS
```

---

## 2. Complete Image Handling Flow

### Phase 1: JS Bridge Pre-Processing (`python_bridge.js`)

**File copy** (line 797):
```js
fs.copyFileSync(sourcePath, targetPath);
```
→ Images preserved (byte-for-byte copy)

**Password decryption** (lines 802–813):
```js
const pwWorkbook = await XlsxPopulate.fromFileAsync(targetPath, { password: ... });
await pwWorkbook.toFileAsync(targetPath); // decrypt
```
→ **RISK**: xlsx-populate re-saves the entire XLSX. May strip ZIP entries it doesn't understand.

**Pending sheet operations** (lines 819–823 → `applyPendingSheetOperations()`):
- Uses JSZip to modify workbook.xml, rels, content types
- **Clone operation** (lines 640–681): Explicitly strips `<drawing>`, `<legacyDrawing>`, and drawing relationships from cloned sheets
- Non-clone operations (rename, move, visibility) preserve all ZIP entries

### Phase 2: Python Writer (`excel_writer.py`)

**Backup mechanism** (lines 4386–4393):
When `original_path == output_path`, a backup copy is created BEFORE openpyxl saves, preserving the original with images:
```python
if os.path.normpath(original_path) == os.path.normpath(output_path):
    shutil.copy2(original_path, _backup_path)
    restore_external_links_from_original._backup_original_path = _backup_path
```

**openpyxl load** (line 4400):
```python
wb = load_workbook(file_path, rich_text=True)
```
→ Without Pillow, openpyxl silently drops ALL image/drawing data from its in-memory model. No error, no warning.

**openpyxl save** (e.g., line 5033):
```python
wb.save(output_path)
```
→ Output XLSX is missing: `xl/media/*`, `xl/drawings/*` content, `<drawing>` elements in sheet XMLs, `xl/richData/*`, `vm` attributes on cells.

### Phase 3: Restore Chain

Called after every `wb.save()`:

1. **`fix_xlsx_relationships(output_path)`** (line 124)
   - Fixes openpyxl path issues (absolute → relative)
   - ZIP-to-ZIP from openpyxl output (images still missing)

2. **`restore_table_xml_from_original(output_path, original_path)`** (line 292)
   - Restores table XML from original
   - ZIP-to-ZIP from openpyxl output (images still missing)

3. **`restore_external_links_from_original(output_path, original_path)`** (line 1146)
   - **THE MAIN IMAGE RESTORATION** — details below

### Phase 4: Image Restoration Detail (`restore_external_links_from_original`)

**Step 1 — Extract** (lines 1207–1222):
```python
with zipfile.ZipFile(output_path, 'r') as zf:
    zf.extractall(temp_dir)         # openpyxl output (no images)
with zipfile.ZipFile(original_path, 'r') as zf:
    zf.extractall(orig_temp_dir)    # original (has images)
```

**Step 2 — Copy media** (lines 1493–1501):
```python
orig_media_dir = os.path.join(orig_temp_dir, 'xl', 'media')
if os.path.exists(orig_media_dir):
    if os.path.exists(dest_media_dir):
        shutil.rmtree(dest_media_dir)
    shutil.copytree(orig_media_dir, dest_media_dir)
```
→ Copies all image files (PNG, JPEG, EMF, etc.) from original

**Step 3 — Copy drawings** (lines 1505–1515):
```python
orig_drawings_dir = os.path.join(orig_temp_dir, 'xl', 'drawings')
if os.path.exists(orig_drawings_dir):
    shutil.copytree(orig_drawings_dir, dest_drawings_dir)
```
→ Copies drawing XMLs (anchor positions, sizes, image references)
→ If `structural_change=True`: strips slicer shapes via `_strip_slicer_shapes_from_drawings()` (images preserved)

**Step 4 — Copy richData** (lines 1530–1546):
```python
orig_richdata_dir = os.path.join(orig_temp_dir, 'xl', 'richData')
if os.path.exists(orig_richdata_dir):
    shutil.copytree(orig_richdata_dir, dest_richdata_dir)
```
→ Copies Excel 365 cell images (rdrichvalue.xml, richValueRel.xml, etc.)

**Step 5 — Copy metadata.xml** (lines 1549–1553):
```python
orig_metadata = os.path.join(orig_temp_dir, 'xl', 'metadata.xml')
if os.path.exists(orig_metadata):
    shutil.copy2(orig_metadata, dest_metadata)
```
→ Required for richData/vm-attribute cell images

**Step 6 — Worksheet rels** (lines 1570–1725):
- **MERGE mode** (`structural_change=True`): Keeps openpyxl's rels, adds missing from original (drawings, printerSettings)
- **REPLACE mode** (`structural_change=False`): Copies original rels wholesale

**Step 7 — Content_Types.xml** (lines 1737–1797):
- Adds missing `<Default Extension="...">` entries (vml, emf, wmf)
- Adds missing `<Override PartName="...">` for drawings/media that exist in temp_dir
- Runs AFTER file copies so existence checks pass

**Step 8 — Worksheet XML elements** (lines 1802–2060):
Restores `<drawing>`, `<legacyDrawing>`, `<picture>` elements in each sheet XML:
- In MERGE mode: with rId mapping from original to openpyxl's numbering
- In REPLACE mode: direct copy from original
- Uses `_insert_ws_element()` to place elements in correct OpenXML schema order

**Step 9 — vm-attributes** (lines 2110–2165):
Restores `vm="N"` attributes on `<c>` elements for Excel 365 cell images. Creates missing cells/rows if openpyxl removed them.

**Step 10 — Namespace restoration** (lines 2100–2110):
Replaces openpyxl's minimal `<worksheet>` root with original's full namespace declarations (xmlns:mc, mc:Ignorable, xmlns:xr, etc.)

**Step 11 — ROBUST RE-ZIP** (lines 2310–2378):
```python
with zipfile.ZipFile(temp_xlsx, 'w', zipfile.ZIP_DEFLATED) as new_zf:
    with zipfile.ZipFile(original_path, 'r') as orig_zf:
        for item in orig_zf.infolist():
            if name in temp_files:
                new_zf.write(temp_files[name], name)  # modified version
            else:
                data = orig_zf.read(name)
                new_zf.writestr(info, data)  # original bytes 1:1
```
→ Uses ORIGINAL as base ZIP, replaces only modified entries
→ Preserves ALL original entries: drawings, media, embeddings, charts, ctrlProps, activeX, diagrams

**Step 12 — Diagnostic logging** (lines 2383–2395):
Dumps all image-related files in the final ZIP to stderr for debugging.

### Phase 5: JS Bridge Post-Processing

**Password re-encryption** (lines 1144–1152):
```js
const pwWorkbook = await XlsxPopulate.fromFileAsync(targetPath);
await pwWorkbook.toFileAsync(targetPath, { password: finalPassword });
```
→ **RISK**: Runs AFTER all Python restoration. If xlsx-populate doesn't preserve unknown ZIP entries, all restored images are lost.

---

## 3. Code Paths Where Images Could Be Lost

### A. xlsx-populate Password Handling (HIGH RISK)

**Location**: `python_bridge.js` lines 802–813 (pre) and 1144–1152 (post)

**Problem**: xlsx-populate opens and re-saves the XLSX for encryption/decryption. This round-trip may not preserve all ZIP entries (especially media, drawings, richData that xlsx-populate doesn't model).

**Impact**: The post-processing encryption runs AFTER the Python restore mechanism has successfully added images back. If xlsx-populate strips them, all restoration work is undone.

**Affected scenarios**: Any export with password protection.

### B. Clone Operations Strip Images (CONFIRMED)

**Location**: `python_bridge.js` lines 642–648

```js
// Strip elements that reference shared resources
clonedXml = clonedXml.replace(/<drawing\s[^>]*\/>/g, '');
clonedXml = clonedXml.replace(/<drawing[\s\S]*?<\/drawing>/g, '');
clonedXml = clonedXml.replace(/<legacyDrawing\s[^>]*\/>/g, '');
```

And lines 680–690: drawing relationships are also stripped from cloned sheet rels.

**Problem**: The subsequent Python restore uses `originalSourcePath` which contains the ORIGINAL sheets, not the cloned sheet. `restore_external_links_from_original` copies drawings/media from original, but the cloned sheet's worksheet XML has no `<drawing>` element. The restore only processes sheets that exist in BOTH original and output — the cloned sheet only exists in the output.

**Impact**: Cloned sheets permanently lose all images.

**Fix needed**: After clone, either:
1. Copy the source sheet's drawing reference into the clone, OR
2. Create a new drawing file for the clone that references the same images

### C. Multi-Pass Operations (`originalPath = targetPath`)

**Location**: `python_bridge.js` lines 896–903, 1021–1028

**Problem**: In combined row+column operations and column-only+cell-edit operations:
- Pass 1 uses `originalPath = originalSourcePath` (pristine original)
- Pass 2 uses `originalPath = targetPath` (already-modified file)

If Pass 1 successfully restores images, targetPath should have them, and Pass 2's backup mechanism (line 4386) creates a backup before openpyxl overwrites. This chain SHOULD work.

However, if any intermediate step in Pass 1 corrupts or fails to restore images, Pass 2 cannot recover them because its "original" (targetPath) already lacks them.

**Impact**: Fragile chain — any failure in Pass 1 cascades to Pass 2.

### D. Pillow Not Installed (ROOT CAUSE of openpyxl behavior)

**Location**: `requirements.txt` — only lists `openpyxl>=3.0.0` and `xlwings>=0.30.0`. Pillow is NOT listed. Neither is it in `python-embed/win-x64/Lib/site-packages/`.

**Impact**: openpyxl's `load_workbook()` silently drops ALL image data. Every `wb.save()` creates an XLSX without images. The entire restore mechanism exists SOLELY to compensate for this.

### E. ExcelJS Writer Has Zero Image Handling

**Location**: `exceljs-writer.js` (2041 lines)

**Finding**: Zero matches for "image", "drawing", "media", "picture", or "addImage" in the entire file. ExcelJS writer does not handle images at all.

**Impact**: If ExcelJS is used as the writer (instead of Python), images are silently lost with no restore mechanism.

---

## 4. FALL 3a: The Only Natively Safe Path

**Location**: `excel_writer.py` lines 7183–7199

```python
if (real_edits or has_visibility_changes or has_add_highlights) and not has_extra_changes:
    wb.close()  # openpyxl not needed
    result = _direct_xml_cell_edit(file_path, output_path, sheet_name, real_edits, ...)
```

FALL 3a uses direct ZIP-to-ZIP copying (line 3405+), modifying only the worksheet XML in memory. ALL other ZIP entries (drawings, media, richData, etc.) are copied byte-for-byte from source. No openpyxl round-trip.

**No `fix_xlsx_relationships`, no `restore_table_xml`, no `restore_external_links` needed.**

This is the ONLY code path that natively preserves images without requiring post-hoc restoration.

---

## 5. Existing Incomplete Image Code

### Diagnostic Functions Already Present

**`_check_zip_drawings()`** (line 103): Checkpoint function that logs drawing/media files at each stage. Called at lines 5034–5040 in the PIPELINE path:
```python
_check_zip_drawings(output_path, "nach wb.save()")
_check_zip_drawings(output_path, "nach fix_xlsx_relationships()")
_check_zip_drawings(output_path, "nach restore_table_xml()")
_check_zip_drawings(output_path, "nach restore_external_links()")
```

**Extensive DEBUG logging** in `restore_external_links_from_original`:
- Lines 1215–1245: Dumps original ZIP drawing/media files, sheet rels content, `<drawing>` presence in sheet XMLs
- Line 2383: `[DIAGNOSE] === FINAL ZIP STATE ===` — dumps all image-related files in final ZIP

This diagnostic infrastructure suggests the developer was actively debugging image loss — the restore mechanism may have been added iteratively in response to discovered issues.

### FALL 3a Row Highlights Extension

`_direct_xml_cell_edit` was extended to handle row highlights (line 3410 parameter `row_highlights`), specifically to AVOID the openpyxl roundtrip that causes image loss:

```python
# AUCH für reine Visibility-Änderungen (hideRow/hideColumn ohne Edits):
# openpyxl-Roundtrip verliert Drawings ohne Pillow → "Zeichnungsform entfernt".
# ZIP-to-ZIP preserviert ALLES aus dem Original.
```

---

## 6. Suggested Fix Approach

### Priority 1: Investigate xlsx-populate Image Preservation

Test whether `xlsx-populate.fromFileAsync()` → `toFileAsync()` preserves media/drawings/richData. If not, the password re-encryption at line 1144 destroys all restored images.

**Fix**: Either:
- Use a ZIP-level encryption approach that preserves all entries
- Re-run `restore_external_links_from_original()` AFTER xlsx-populate encryption
- Use a different encryption library

### Priority 2: Install Pillow

Adding `Pillow` to `requirements.txt` and the embedded Python would allow openpyxl to natively handle images during round-trip, eliminating the need for most of the restore mechanism.

```
# requirements.txt
openpyxl>=3.0.0
xlwings>=0.30.0
Pillow>=9.0.0
```

For the embedded Python (`python-embed/win-x64/`), include the Pillow wheel.

### Priority 3: Fix Clone Image Restoration

After cloning a sheet in `applyPendingSheetOperations`, duplicate the drawing reference:

```js
// After cloning sheet XML:
// 1. Copy source drawing XML to new drawing file
// 2. Copy source drawing rels to new drawing rels
// 3. Keep <drawing> element in cloned sheet XML
// 4. Update rId in cloned sheet to point to new drawing
```

### Priority 4: Add ExcelJS Image Preservation

The ExcelJS writer needs image passthrough support. ExcelJS has `workbook.addImage()` API but it's not used. At minimum, implement ZIP-level preservation similar to FALL 3a.

### Priority 5: Reduce Multi-Pass Fragility

For combined operations, consider saving `originalSourcePath` for ALL passes instead of switching to `targetPath` for Pass 2. The comment at line 896 explains why `targetPath` is used (to preserve Pass 1's changes), but a dedicated backup could serve as the image source:

```js
// Pass 2: Use targetPath for structure, but originalSourcePath for image restoration
const colConfig = {
    originalPath: targetPath,           // structure from Pass 1
    imageSourcePath: originalSourcePath  // always from pristine original
};
```

---

## 7. Summary Table

| Export Path | Images Preserved? | Mechanism | Risk Level |
|---|---|---|---|
| FALL 3a (direct XML) | ✅ Yes | ZIP-to-ZIP copy | None |
| FALL 3b (openpyxl + formats) | ✅ Yes* | restore_external_links | Low |
| FALL 1 (fromFile) | ✅ Yes* | restore_external_links | Low |
| FALL 1.5/1.9 (column ops) | ✅ Yes* | restore_external_links | Low |
| FALL 2 (structural/fullRewrite) | ✅ Yes* | restore_external_links | Low |
| PIPELINE (row+col operations) | ✅ Yes* | restore_external_links | Medium |
| Clone operations | ❌ No | Drawing intentionally stripped | **High** |
| Password-protected export | ⚠️ Maybe | xlsx-populate may strip | **High** |
| ExcelJS writer | ❌ No | No image code exists | **High** |
| Multi-pass (Pass 2) | ⚠️ Depends | Chain from Pass 1 | Medium |

\* Assuming `original_path` points to a file WITH images and no post-processing strips them.

---

## 8. Key Code Locations

| Component | File | Lines | Description |
|---|---|---|---|
| Main writer entry | excel_writer.py | 4357–4440 | `write_sheet()` + backup mechanism |
| Restore function | excel_writer.py | 1146–2400 | `restore_external_links_from_original()` |
| Media/drawing copy | excel_writer.py | 1493–1558 | Copy xl/media, xl/drawings, xl/richData |
| Worksheet XML restore | excel_writer.py | 1802–2060 | `<drawing>`, `<legacyDrawing>`, `<picture>` |
| vm-attribute restore | excel_writer.py | 2110–2165 | Excel 365 cell images |
| ROBUST RE-ZIP | excel_writer.py | 2310–2378 | ZIP-to-ZIP from original as base |
| Content_Types merge | excel_writer.py | 1737–1797 | Drawing/media overrides |
| Diagnostic checkpoint | excel_writer.py | 103–121 | `_check_zip_drawings()` |
| Direct XML (FALL 3a) | excel_writer.py | 3405–3500 | `_direct_xml_cell_edit()` |
| FALL 3a/3b gate | excel_writer.py | 7140–7200 | Decision logic |
| Clone stripping | python_bridge.js | 640–690 | Drawing/legacyDrawing removal |
| xlsx-populate decrypt | python_bridge.js | 802–813 | Pre-processing |
| xlsx-populate encrypt | python_bridge.js | 1144–1152 | Post-processing |
| Multi-pass Pass 2 | python_bridge.js | 896–940 | `originalPath = targetPath` |
| ExcelJS writer | exceljs-writer.js | 1–2041 | Zero image references |
| Slicer stripping | excel_writer.py | 1002–1060 | `_strip_slicer_shapes_from_drawings()` (images safe) |
| Pillow reference | requirements.txt | — | NOT listed |
