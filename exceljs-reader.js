// ============================================================================
// EXCELJS MIGRATION - NEUE READ-FUNKTION
// ============================================================================
// Dieses Modul enthält die ExcelJS-basierte Sheet-Read-Funktion
// Zum Testen der Migration von xlsx-populate zu exceljs

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const AdmZip = require('adm-zip');
let XlsxPopulate = null; // Lazy-load für Passwort-Entschlüsselung

/**
 * Konvertiert Spalten-Buchstaben zu Index (A=0, B=1, ..., Z=25, AA=26, ...)
 * @param {string} letters - Spalten-Buchstaben (z.B. "A", "AA", "BC")
 * @returns {number} 0-basierter Spalten-Index
 */
function colLettersToIndex(letters) {
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
        index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index - 1; // 0-basiert
}

/**
 * Parst einen Range-String (z.B. "A1:H1") zu einem Objekt
 * @param {string} rangeStr - Range im Format "A1:H1"
 * @returns {Object|null} { startRow, startCol, endRow, endCol, rowSpan, colSpan }
 */
function parseRangeString(rangeStr) {
    const match = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
        console.warn(`[ExcelJS] Ungültiger Range-String: ${rangeStr}`);
        return null;
    }
    
    const startCol = colLettersToIndex(match[1]);
    const startRow = parseInt(match[2]) - 1; // 0-basiert
    const endCol = colLettersToIndex(match[3]);
    const endRow = parseInt(match[4]) - 1; // 0-basiert
    
    return {
        startRow,
        startCol,
        endRow,
        endCol,
        rowSpan: endRow - startRow + 1,
        colSpan: endCol - startCol + 1
    };
}

/**
 * Erkennt Zeilenfarben basierend auf cellStyles.
 * Eine Zeile wird als "markiert" erkannt, wenn ALLE Zellen die gleiche Hintergrundfarbe haben
 * und diese Farbe einer der bekannten Highlight-Farben entspricht.
 * 
 * @param {Object} cellStyles - Map von "rowIndex-colIndex" zu Style-Objekt mit fill
 * @param {number} rowCount - Anzahl der Datenzeilen
 * @param {number} colCount - Anzahl der Spalten
 * @returns {Array} Array von [rowIndex, colorName] Paaren
 */
function detectRowHighlights(cellStyles, rowCount, colCount) {
    const highlights = [];
    
    // Mapping von ARGB-Farben zu Highlight-Namen (ohne Alpha-Kanal)
    const colorMapping = {
        '90EE90': 'green',   // Light Green
        'FFFF00': 'yellow',  // Yellow
        'FFA500': 'orange',  // Orange
        'FF6B6B': 'red',     // Light Red
        '87CEEB': 'blue',    // Sky Blue
        'DDA0DD': 'purple',  // Plum
        // Alternative Farben die auch erkannt werden sollen
        '4CAF50': 'green',
        'FFEB3B': 'yellow',
        'FF9800': 'orange',
        'F44336': 'red',
        '2196F3': 'blue',
        '9C27B0': 'purple'
    };
    
    // Für jede Zeile prüfen
    for (let rowIdx = 0; rowIdx < rowCount; rowIdx++) {
        const rowFills = [];
        
        // Alle Zellen in der Zeile durchgehen
        for (let colIdx = 0; colIdx < colCount; colIdx++) {
            // cellStyles-Key Format: "rowIndex-colIndex" wobei rowIndex 1-basiert ist für Datenzeilen
            const styleKey = `${rowIdx + 1}-${colIdx}`;
            const style = cellStyles[styleKey];
            
            if (style && style.fill) {
                // Fill ist im Format "#RRGGBB"
                const fillHex = style.fill.replace('#', '').toUpperCase();
                rowFills.push(fillHex);
            } else {
                rowFills.push(null);
            }
        }
        
        // Prüfen ob alle Zellen die gleiche Farbe haben (und nicht null)
        const nonNullFills = rowFills.filter(f => f !== null);
        if (nonNullFills.length === colCount && nonNullFills.length > 0) {
            const firstFill = nonNullFills[0];
            const allSame = nonNullFills.every(f => f === firstFill);
            
            if (allSame) {
                // Prüfen ob die Farbe einer bekannten Highlight-Farbe entspricht
                const colorName = colorMapping[firstFill];
                if (colorName) {
                    highlights.push([rowIdx, colorName]);
                }
            }
        }
    }
    
    return highlights;
}

/**
 * Extrahiert Fill-Farben direkt aus der XLSX-Datei (ZIP-Format).
 * Dies ist ein Workaround für ExcelJS, das bei bestimmten Excel-Dateien
 * (z.B. von SoftMaker/PlanMaker erstellt) keine Fills erkennt.
 * 
 * @param {string} filePath - Pfad zur XLSX-Datei
 * @param {string} sheetName - Name des Sheets
 * @returns {Object} Map von "rowNumber-colNumber" zu Fill-Farbe (z.B. "#FF0000")
 */
function extractFillsFromXLSX(fileBufferOrPath, sheetName) {
    const cellFills = {};
    
    try {
        // Akzeptiert Buffer oder Dateipfad
        const zip = Buffer.isBuffer(fileBufferOrPath) ? new AdmZip(fileBufferOrPath) : new AdmZip(fileBufferOrPath);
        
        // 1. styles.xml lesen und Fills extrahieren
        const stylesEntry = zip.getEntry('xl/styles.xml');
        if (!stylesEntry) {
            return cellFills;
        }
        
        const stylesXml = stylesEntry.getData().toString('utf8');
        
        // Fills extrahieren (Position im Array = fillId)
        const fills = [];
        const fillsMatch = stylesXml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
        if (fillsMatch) {
            const fillPattern = /<fill[^\/]*>([\s\S]*?)<\/fill>/g;
            let fillMatch;
            while ((fillMatch = fillPattern.exec(fillsMatch[1])) !== null) {
                const fillContent = fillMatch[1];
                // Suche nach fgColor mit rgb-Attribut
                const fgColorMatch = fillContent.match(/<fgColor[^>]*rgb="([A-Fa-f0-9]{8})"[^>]*\/?>/);
                if (fgColorMatch) {
                    const argb = fgColorMatch[1];
                    // ARGB zu RGB (erste 2 Zeichen = Alpha)
                    const rgb = argb.substring(2);
                    fills.push('#' + rgb);
                } else {
                    fills.push(null);
                }
            }
        }
        
        // 2. cellXfs extrahieren (Style ID -> Fill ID Mapping)
        const styleToFill = [];
        const cellXfsMatch = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
        if (cellXfsMatch) {
            const xfPattern = /<xf[^>]*>/g;
            let xfMatch;
            while ((xfMatch = xfPattern.exec(cellXfsMatch[1])) !== null) {
                const xfContent = xfMatch[0];
                const fillIdMatch = xfContent.match(/fillId="(\d+)"/);
                const applyFillMatch = xfContent.match(/applyFill="(\d+)"/);
                
                if (fillIdMatch) {
                    const fillId = parseInt(fillIdMatch[1]);
                    // applyFill muss 1 sein oder nicht vorhanden (dann gilt fillId)
                    if (!applyFillMatch || applyFillMatch[1] === '1') {
                        styleToFill.push(fillId);
                    } else {
                        styleToFill.push(null);
                    }
                } else {
                    styleToFill.push(null);
                }
            }
        }
        
        // 3. Sheet-Daten finden
        // Zuerst workbook.xml lesen um Sheet rId zu finden
        const workbookEntry = zip.getEntry('xl/workbook.xml');
        if (!workbookEntry) {
            return cellFills;
        }
        
        const workbookXml = workbookEntry.getData().toString('utf8');
        const sheetsMatch = workbookXml.match(/<sheets>([\s\S]*?)<\/sheets>/);
        let sheetRId = 'rId1'; // Default
        
        if (sheetsMatch) {
            const sheetPattern = /<sheet[^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"[^>]*\/?>/g;
            let sheetMatch;
            while ((sheetMatch = sheetPattern.exec(sheetsMatch[1])) !== null) {
                if (sheetMatch[1] === sheetName) {
                    sheetRId = sheetMatch[2];
                    break;
                }
            }
        }
        
        // Relationship-Datei lesen um tatsächlichen Sheet-Pfad zu finden
        const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
        if (!relsEntry) {
            return cellFills;
        }
        
        const relsXml = relsEntry.getData().toString('utf8');
        let sheetPath = null;
        
        const relPattern = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*\/?>/g;
        let relMatch;
        while ((relMatch = relPattern.exec(relsXml)) !== null) {
            if (relMatch[1] === sheetRId) {
                sheetPath = relMatch[2];
                break;
            }
        }
        
        if (!sheetPath) {
            return cellFills;
        }
        
        // Sheet-XML laden (Pfad kann relativ sein, z.B. "worksheets/sheet1.xml")
        const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : `xl/${sheetPath}`;
        const sheetEntry = zip.getEntry(fullSheetPath);
        if (!sheetEntry) {
            return cellFills;
        }
        
        const sheetXml = sheetEntry.getData().toString('utf8');
        
        // 4. Zellen mit Style-IDs extrahieren
        const cellPattern = /<c r="([A-Z]+)(\d+)"[^>]*s="(\d+)"[^>]*>/g;
        let cellMatch;
        while ((cellMatch = cellPattern.exec(sheetXml)) !== null) {
            const colLetters = cellMatch[1];
            const rowNum = parseInt(cellMatch[2]);
            const styleId = parseInt(cellMatch[3]);
            
            // Spalten-Buchstaben zu Index konvertieren (A=0, B=1, ...)
            let colIndex = 0;
            for (let i = 0; i < colLetters.length; i++) {
                colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 64);
            }
            colIndex--; // 0-basiert
            
            // Fill-Farbe ermitteln
            const fillId = styleToFill[styleId];
            if (fillId !== null && fillId !== undefined && fills[fillId]) {
                const fillColor = fills[fillId];
                // Ignoriere Weiß
                if (fillColor !== '#FFFFFF') {
                    // Key: "rowNumber-colIndex" (rowNumber ist 1-basiert, colIndex 0-basiert)
                    // Für Daten-Zeilen (ab Zeile 2) wird rowNumber-1 als Key verwendet
                    // da das Frontend 1-basierte Indizes verwendet
                    const dataRowIndex = rowNum - 1; // Zeile 2 = Datenzeile 1
                    const key = `${dataRowIndex}-${colIndex}`;
                    cellFills[key] = fillColor;
                }
            }
        }
        
        return cellFills;
        
    } catch (error) {
        console.error('[XLSX-Extract] Fehler:', error);
        return cellFills;
    }
}



/**
 * Extrahiert Sheet-Metadaten direkt aus der XLSX-Datei (ZIP-Format).
 * Schnell (~10ms), kein Zell-Parsing, blockiert den Event-Loop nicht nennenswert.
 * Liefert: Spaltenanzahl, versteckte Spalten, verbundene Zellen, AutoFilter.
 */
function extractSheetMetadata(fileBufferOrPath, sheetName) {
    const result = {
        columnCount: 1,
        hiddenColumns: [],
        mergedCells: [],
        autoFilterRange: null,
        imageCells: []  // Zellen die Bild-Formeln enthalten (DISPIMG, IMAGE)
    };
    
    try {
        // Akzeptiert Buffer oder Dateipfad
        const zip = Buffer.isBuffer(fileBufferOrPath) ? new AdmZip(fileBufferOrPath) : new AdmZip(fileBufferOrPath);
        
        // workbook.xml lesen um Sheet-rId zu finden
        const workbookEntry = zip.getEntry('xl/workbook.xml');
        if (!workbookEntry) return result;
        
        const workbookXml = workbookEntry.getData().toString('utf8');
        const sheetsMatch = workbookXml.match(/<sheets>([\s\S]*?)<\/sheets>/);
        let sheetRId = 'rId1';
        
        if (sheetsMatch) {
            const sheetPattern = /<sheet[^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"[^>]*\/?>/g;
            let m;
            while ((m = sheetPattern.exec(sheetsMatch[1])) !== null) {
                const name = m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
                if (name === sheetName) {
                    sheetRId = m[2];
                    break;
                }
            }
        }
        
        // Rels lesen für Sheet-Pfad
        const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
        if (!relsEntry) return result;
        
        const relsXml = relsEntry.getData().toString('utf8');
        let sheetPath = null;
        
        const relPattern = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*\/?>/g;
        let relMatch;
        while ((relMatch = relPattern.exec(relsXml)) !== null) {
            if (relMatch[1] === sheetRId) {
                sheetPath = relMatch[2];
                break;
            }
        }
        
        if (!sheetPath) return result;
        
        const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : `xl/${sheetPath}`;
        const sheetEntry = zip.getEntry(fullSheetPath);
        if (!sheetEntry) return result;
        
        const sheetXml = sheetEntry.getData().toString('utf8');
        
        // Dimension (Spaltenanzahl)
        const dimMatch = sheetXml.match(/<dimension ref="[A-Z]+\d+:([A-Z]+)\d+"/);
        if (dimMatch) {
            result.columnCount = colLettersToIndex(dimMatch[1]) + 1;
        }
        
        // Versteckte Spalten aus <cols>
        const colsMatch = sheetXml.match(/<cols>([\s\S]*?)<\/cols>/);
        if (colsMatch) {
            const colPattern = /<col\s([^>]*)\/?\s*>/g;
            let colM;
            while ((colM = colPattern.exec(colsMatch[1])) !== null) {
                const attrs = colM[1];
                if (/hidden="(1|true)"/i.test(attrs)) {
                    const minMatch = attrs.match(/min="(\d+)"/);
                    const maxMatch = attrs.match(/max="(\d+)"/);
                    if (minMatch && maxMatch) {
                        for (let i = parseInt(minMatch[1]) - 1; i < parseInt(maxMatch[1]); i++) {
                            result.hiddenColumns.push(i);
                        }
                    }
                }
            }
        }
        
        // AutoFilter direkt aus Sheet-XML
        const afMatch = sheetXml.match(/<autoFilter[^>]*ref="([^"]*)"/);
        if (afMatch) {
            result.autoFilterRange = afMatch[1];
        }
        
        // Verbundene Zellen aus <mergeCells>
        const mcMatch = sheetXml.match(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/);
        if (mcMatch) {
            const mcPattern = /<mergeCell\s+ref="([^"]*)"\s*\/?>/g;
            let mcM;
            while ((mcM = mcPattern.exec(mcMatch[1])) !== null) {
                const parsed = parseRangeString(mcM[1]);
                if (parsed) result.mergedCells.push(parsed);
            }
        }
        
        // ================================================================
        // Bild-Zellen erkennen
        // Excel 365 speichert Zellbilder auf 2 Arten:
        // 1) DISPIMG/IMAGE Formeln in <f>-Tags (ältere Methode)
        // 2) vm-Attribut (Value Metadata) auf <c>-Tags + richData (moderne Methode)
        // ================================================================
        const cellBlocks = sheetXml.split('</c>');
        
        // Methode 1: DISPIMG/IMAGE Formeln
        for (const block of cellBlocks) {
            if (!/DISPIMG|IMAGE/i.test(block)) continue;
            if (!/<f[\s>]/.test(block)) continue;
            
            const cellRefMatch = block.match(/<c\s[^>]*?r="([A-Z]+\d+)"/);
            if (!cellRefMatch) continue;
            
            const cellRef = cellRefMatch[1];
            const colLetters = cellRef.match(/^([A-Z]+)/)[1];
            const rowNum = parseInt(cellRef.match(/(\d+)$/)[1]);
            result.imageCells.push({
                ref: cellRef,
                col: colLettersToIndex(colLetters),
                row: rowNum - 1
            });
        }
        
        // Shared formulas mit DISPIMG/IMAGE
        if (result.imageCells.length > 0) {
            const sharedImageIds = new Set();
            for (const block of cellBlocks) {
                if (!/DISPIMG|IMAGE/i.test(block)) continue;
                const siMatch = block.match(/<f[^>]*t="shared"[^>]*si="(\d+)"/);
                if (siMatch) sharedImageIds.add(siMatch[1]);
            }
            if (sharedImageIds.size > 0) {
                for (const block of cellBlocks) {
                    const siRef = block.match(/<f[^>]*t="shared"[^>]*si="(\d+)"[^>]*\/?>/);
                    if (!siRef || !sharedImageIds.has(siRef[1])) continue;
                    const cellRefMatch = block.match(/<c\s[^>]*?r="([A-Z]+\d+)"/);
                    if (!cellRefMatch) continue;
                    const cellRef = cellRefMatch[1];
                    const col = colLettersToIndex(cellRef.match(/^([A-Z]+)/)[1]);
                    const row = parseInt(cellRef.match(/(\d+)$/)[1]) - 1;
                    if (!result.imageCells.some(ic => ic.col === col && ic.row === row)) {
                        result.imageCells.push({ ref: cellRef, col, row });
                    }
                }
            }
        }
        
        // Methode 2: vm-Attribut (Value Metadata) — Excel 365 Zellbilder
        // Zellen mit vm="N" verweisen auf xl/richData/rdrichvalue.xml
        // Prüfe zuerst ob richData existiert (= es gibt Zellbilder)
        const hasRichData = zip.getEntries().some(e => /richData/i.test(e.entryName));
        if (hasRichData) {
            // Prüfe welche richData-Einträge Bilder enthalten
            let richDataHasImages = false;
            try {
                // rdrichvalue.xml enthält <rv>-Einträge, Bilder haben type mit Bild-Referenz
                const rdEntry = zip.getEntry('xl/richData/rdrichvalue.xml') || 
                                zip.getEntry('xl/richData/rdRichValue.xml');
                if (rdEntry) {
                    const rdXml = rdEntry.getData().toString('utf8');
                    // Bilder in richData haben typischerweise einen Verweis auf media/ oder image
                    richDataHasImages = /img|image|picture|Bild/i.test(rdXml) || 
                                         rdXml.includes('<v>') || // Hat Werte = wahrscheinlich Bilder
                                         true; // richData existiert = fast immer Bilder
                }
                // Auch rdRichValueStructure.xml prüfen
                const structEntry = zip.getEntry('xl/richData/rdrichvaluestructure.xml') ||
                                    zip.getEntry('xl/richData/rdRichValueStructure.xml');
                if (structEntry) {
                    const structXml = structEntry.getData().toString('utf8');
                    if (/image|img|picture/i.test(structXml)) {
                        richDataHasImages = true;
                    }
                }
            } catch (e) {
                // Fallback: richData existiert = nehme an es sind Bilder
                richDataHasImages = true;
            }
            
            if (richDataHasImages) {
                // Finde alle Zellen mit vm-Attribut
                for (const block of cellBlocks) {
                    // vm-Attribut auf dem <c>-Element
                    const vmMatch = block.match(/<c\s[^>]*?\bvm="(\d+)"[^>]*?r="([A-Z]+\d+)"/);
                    const vmMatch2 = block.match(/<c\s[^>]*?r="([A-Z]+\d+)"[^>]*?\bvm="(\d+)"/);
                    
                    let cellRef = null;
                    if (vmMatch) {
                        cellRef = vmMatch[2];
                    } else if (vmMatch2) {
                        cellRef = vmMatch2[1];
                    }
                    
                    if (!cellRef) continue;
                    
                    const col = colLettersToIndex(cellRef.match(/^([A-Z]+)/)[1]);
                    const row = parseInt(cellRef.match(/(\d+)$/)[1]) - 1;
                    if (!result.imageCells.some(ic => ic.col === col && ic.row === row)) {
                        result.imageCells.push({ ref: cellRef, col, row });
                    }
                }
                console.log(`[SheetMetadata] richData gefunden, ${result.imageCells.length} Bild-Zellen (inkl. vm-Attribut)`);
            }
        }
        
        if (result.imageCells.length > 0) {
            console.log(`[SheetMetadata] ${result.imageCells.length} Bild-Zellen erkannt: ${result.imageCells.map(c => c.ref).join(', ')}`);
        }
        
        // Tables (für AutoFilter, falls kein direkter AutoFilter vorhanden)
        if (!result.autoFilterRange) {
            try {
                const sheetFileName = fullSheetPath.split('/').pop();
                const sheetDir = fullSheetPath.substring(0, fullSheetPath.lastIndexOf('/'));
                const sheetRelsPath = `${sheetDir}/_rels/${sheetFileName}.rels`;
                const sheetRelsEntry = zip.getEntry(sheetRelsPath);
                
                if (sheetRelsEntry) {
                    const sheetRelsXml = sheetRelsEntry.getData().toString('utf8');
                    const tableRefPattern = /Target="([^"]*table[^"]*)"/g;
                    let tableRefMatch;
                    
                    while ((tableRefMatch = tableRefPattern.exec(sheetRelsXml)) !== null) {
                        let tablePath = tableRefMatch[1];
                        if (tablePath.startsWith('../')) {
                            tablePath = 'xl/' + tablePath.replace('../', '');
                        } else if (!tablePath.startsWith('xl/')) {
                            tablePath = `${sheetDir}/${tablePath}`;
                        }
                        
                        const tableEntry = zip.getEntry(tablePath);
                        if (tableEntry) {
                            const tableXml = tableEntry.getData().toString('utf8');
                            const tableAfMatch = tableXml.match(/<autoFilter[^>]*ref="([^"]*)"/);
                            if (tableAfMatch) {
                                result.autoFilterRange = tableAfMatch[1];
                                break;
                            }
                            // Fallback: tableRef als AutoFilter
                            const tableRefAttr = tableXml.match(/<table[^>]*ref="([^"]*)"/);
                            if (tableRefAttr) {
                                result.autoFilterRange = tableRefAttr[1];
                                break;
                            }
                        }
                    }
                }
            } catch (tableErr) {
                // Tables sind optional
            }
        }
        
    } catch (error) {
        console.error('[SheetMetadata] Fehler:', error.message);
    }
    
    return result;
}


/**
 * Liest ein Excel-Sheet mit ExcelJS Streaming Reader (nicht-blockierend)
 * 
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des zu lesenden Sheets
 * @param {string|null} password - Optional: Passwort für geschützte Dateien
 * @returns {Promise<Object>} Sheet-Daten im gleichen Format wie xlsx-populate
 */
async function readSheetWithExcelJS(filePath, sheetName, password = null) {
    const startTime = Date.now();
    let tempFilePath = null; // Für entschlüsselte Dateien
    const timings = {}; // Detaillierte Zeitmessungen
    
    try {
        console.log(`[ExcelJS] === START readSheetWithExcelJS === Datei: ${path.basename(filePath)}, Sheet: ${sheetName}`);
        
        // Bei passwortgeschützten Dateien: xlsx-populate zum Entschlüsseln verwenden
        // ExcelJS hat bekannte Probleme mit Passwort-Entschlüsselung
        let actualFilePath = filePath;
        
        if (password) {
            try {
                // Lazy-load xlsx-populate
                if (!XlsxPopulate) {
                    XlsxPopulate = require('xlsx-populate');
                }
                
                // Datei mit xlsx-populate öffnen (entschlüsseln)
                const pwWorkbook = await XlsxPopulate.fromFileAsync(filePath, { password });
                
                // Als temporäre Datei ohne Passwort speichern
                tempFilePath = path.join(os.tmpdir(), `mvms_decrypt_${crypto.randomUUID()}.xlsx`);
                await pwWorkbook.toFileAsync(tempFilePath);
                
                actualFilePath = tempFilePath;
                
            } catch (pwError) {
                console.error('[ExcelJS] Fehler beim Entschlüsseln:', pwError.message);
                
                // Prüfe ob es ein Passwort-Fehler ist
                if (pwError.message.includes('password') || pwError.message.includes('Password') || 
                    pwError.message.includes('decrypt') || pwError.message.includes('Decrypt')) {
                    return { 
                        success: false, 
                        error: 'Falsches Passwort oder Datei kann nicht entschlüsselt werden',
                        needsPassword: true
                    };
                }
                throw pwError;
            }
        }
        
        // Datei EINMAL async lesen — dieser Buffer wird für ALLES verwendet:
        // 1. Passwort-Check (AdmZip)
        // 2. Sheet-Metadaten (AdmZip)
        // 3. Streaming Reader (Readable.from)
        // Danach ist der File-Handle sofort frei für xlwings/Excel
        const t0 = Date.now();
        const fileBuffer = await fs.promises.readFile(actualFilePath);
        timings.fileRead = Date.now() - t0;
        console.log(`[ExcelJS] Datei gelesen: ${(fileBuffer.length / 1024 / 1024).toFixed(1)} MB in ${timings.fileRead}ms`);
        
        // Prüfe ob die Datei passwortgeschützt ist
        if (!password) {
            try {
                new AdmZip(fileBuffer);
            } catch (zipError) {
                if (zipError.message.includes('password') || zipError.message.includes('Password') ||
                    zipError.message.includes('encrypted') || zipError.message.includes('Encrypted') ||
                    zipError.message.includes('CFB') || zipError.message.includes('Invalid or unsupported zip')) {
                    return { 
                        success: false, 
                        error: 'Diese Datei ist passwortgeschützt. Bitte Passwort eingeben.',
                        needsPassword: true
                    };
                }
            }
        }
        
        // Sheet-Metadaten aus Buffer extrahieren (kein erneuter Dateizugriff!)
        const t1 = Date.now();
        const metadata = extractSheetMetadata(fileBuffer, sheetName);
        timings.metadata = Date.now() - t1;
        let actualColumnCount = metadata.columnCount || 1;
        console.log(`[ExcelJS] Metadaten: ${actualColumnCount} Spalten, ${metadata.mergedCells.length} Merged Cells, ${metadata.hiddenColumns.length} Hidden Cols in ${timings.metadata}ms`);
        
        // Daten-Strukturen initialisieren
        const headers = [];
        const data = [];
        const hiddenRows = [];
        const cellStyles = {};
        const cellFormulas = {};
        const cellHyperlinks = {};
        const richTextCells = {};
        
        // Metadaten aus ZIP (statt aus worksheet-Objekt)
        const autoFilterRange = metadata.autoFilterRange;
        const mergedCells = metadata.mergedCells;
        const hiddenColumns = metadata.hiddenColumns;
        
        // Bild-Zellen als Set für schnellen Lookup (z.B. "1_4" = col 1, row 4)
        const imageCellSet = new Set();
        if (metadata.imageCells && metadata.imageCells.length > 0) {
            for (const ic of metadata.imageCells) {
                imageCellSet.add(`${ic.col}_${ic.row}`);
            }
            console.log(`[ExcelJS] ${imageCellSet.size} Bild-Zellen aus Metadaten geladen`);
        }
        
        // ============================================================
        // STREAMING READER: Liest Zeilen einzeln, blockiert Event-Loop NICHT
        // Buffer wird wiederverwendet (kein erneuter Dateizugriff!)
        // ============================================================
        const t2 = Date.now();
        const { Readable } = require('stream');
        const readStream = Readable.from(fileBuffer);
        const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(readStream, {
            sharedStrings: 'cache',
            hyperlinks: 'cache',
            styles: 'cache',
            worksheets: 'emit'
        });
        
        let sheetFound = false;
        
        for await (const worksheetReader of workbookReader) {
            if (worksheetReader.name !== sheetName) continue;
            sheetFound = true;
            console.log(`[ExcelJS] Sheet "${sheetName}" gefunden, starte Streaming...`);
            
            let dataRowCounter = 0;
            let lastProgressLog = Date.now();
        
        for await (const row of worksheetReader) {
            const rowNumber = row.number;
            
            // Leere Zeilen auffüllen (Streaming überspringt Zeilen ohne Zellen im XML)
            if (rowNumber > 1) {
                const expectedDataRow = rowNumber - 2;
                while (dataRowCounter < expectedDataRow) {
                    data.push(new Array(actualColumnCount).fill(''));
                    dataRowCounter++;
                }
            }
            
            // Erste Zeile = Header
            if (rowNumber === 1) {
                // Initialisiere Header-Array mit leeren Strings für alle Spalten
                for (let i = 0; i < actualColumnCount; i++) {
                    headers.push('');
                }
                row.eachCell((cell, colNumber) => {
                    const colIndex = colNumber - 1;
                    // Header-Array erweitern falls nötig
                    while (colIndex >= headers.length) {
                        headers.push('');
                        actualColumnCount = Math.max(actualColumnCount, headers.length);
                    }
                    // Überschreibe den leeren Wert mit dem tatsächlichen Wert
                    if (!cell.value) {
                        headers[colIndex] = '';
                    } else if (typeof cell.value === 'object') {
                        // Rich Text, Hyperlinks, Bilder etc.
                        if (cell.value.richText) {
                            headers[colIndex] = cell.value.richText.map(part => part.text).join('');
                        } else if (cell.value.text !== undefined) {
                            headers[colIndex] = String(cell.value.text);
                        } else if (cell.value.buffer || cell.value.image || cell.value.imageId) {
                            headers[colIndex] = '🖼️ Bild';
                        } else {
                            headers[colIndex] = '📎 Objekt';
                        }
                    } else {
                        headers[colIndex] = String(cell.value);
                    }
                    
                    // WICHTIG: Auch Header-Styles extrahieren (für Frontend-Kompatibilität)
                    const styleKey = `0-${colIndex}`; // Header = Zeile 0
                    const style = {};
                    
                    if (cell.font) {
                        if (cell.font.bold) style.bold = true;
                        if (cell.font.italic) style.italic = true;
                        if (cell.font.underline) style.underline = true;
                        if (cell.font.strike) style.strikethrough = true;
                        if (cell.font.size) {
                            style.fontSize = cell.font.size;
                        }
                        if (cell.font.name && cell.font.name !== 'Calibri') {
                            style.fontName = cell.font.name;
                        }
                        if (cell.font.color?.argb) {
                            const colorHex = cell.font.color.argb.substring(2);
                            if (colorHex !== '000000') {
                                style.fontColor = `#${colorHex}`;
                            }
                        }
                    }
                    
                    // Alignment extrahieren
                    if (cell.alignment) {
                        if (cell.alignment.horizontal && cell.alignment.horizontal !== 'general') {
                            style.textAlign = cell.alignment.horizontal;
                        }
                        if (cell.alignment.vertical && cell.alignment.vertical !== 'bottom') {
                            style.verticalAlign = cell.alignment.vertical;
                        }
                        if (cell.alignment.wrapText) {
                            style.wrapText = true;
                        }
                    }
                    
                    // Fill extrahieren
                    if (cell.fill) {
                        if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' && cell.fill.fgColor?.argb) {
                            const fillHex = cell.fill.fgColor.argb.substring(2);
                            if (fillHex !== 'FFFFFF') {
                                style.fill = `#${fillHex}`;
                            }
                        }
                    }
                    
                    // Borders extrahieren
                    if (cell.border) {
                        const borders = {};
                        for (const side of ['top', 'bottom', 'left', 'right']) {
                            if (cell.border[side] && cell.border[side].style) {
                                borders[side] = {
                                    style: cell.border[side].style,
                                    color: cell.border[side].color?.argb ? `#${cell.border[side].color.argb.substring(2)}` : null
                                };
                            }
                        }
                        if (Object.keys(borders).length > 0) {
                            style.borders = borders;
                        }
                    }
                    
                    if (Object.keys(style).length > 0) {
                        cellStyles[styleKey] = style;
                    }
                });
                continue; // Weiter zur nächsten Zeile
            }
            
            // Daten-Zeilen
            // Initialisiere rowData mit leeren Strings für alle Spalten
            const rowData = new Array(actualColumnCount).fill('');
            
            // WICHTIG: Style-Key basiert auf dataRowCounter, nicht auf rowNumber!
            // Das stellt sicher, dass leere Zeilen nicht zu Index-Mismatches führen
            const currentDataRowIndex = dataRowCounter;
            
            row.eachCell((cell, colNumber) => {
                const colIndex = colNumber - 1;
                // WICHTIG: Frontend erwartet 1-basierte Indizes (wie xlsx-populate)
                const styleKey = `${currentDataRowIndex + 1}-${colIndex}`;
                
                let cellValue = cell.value;
                
                // Bild-Zellen sofort erkennen (aus XML-Metadaten)
                // ExcelJS kann DISPIMG/IMAGE Formeln nicht parsen, liefert #VALUE! oder Objekte
                // Daher VOR jeder anderen Verarbeitung abfangen
                if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                    cellValue = '🖼️ Bild';
                    rowData[colIndex] = cellValue;
                    return; // Nächste Zelle
                }                
                // Formel extrahieren - WICHTIG: VOR der Objekt-Behandlung!
                // Bei Formeln kann cell.value ein Objekt sein mit { formula, result }
                // oder cell.formula ist direkt verfügbar
                if (cell.formula) {
                    cellFormulas[styleKey] = cell.formula;
                    // Das Ergebnis ist in cell.result (nicht cell.value!)
                    cellValue = cell.result !== undefined ? cell.result : '';
                    // Bild-Formeln erkennen: IMAGE, DISPIMG (mit beliebigen Prefixen wie _xlfn._xlws.)
                    if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(cell.formula)) {
                        cellValue = '🖼️ Bild';
                    }
                } else if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
                    // Formel als Objekt gespeichert: { formula: '...', result: ... }
                    cellFormulas[styleKey] = cell.value.formula;
                    cellValue = cell.value.result !== undefined ? cell.value.result : '';
                    // Bild-Formeln erkennen
                    if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(cell.value.formula)) {
                        cellValue = '🖼️ Bild';
                    }
                }
                
                // Error-Werte behandeln (z.B. { error: '#VALUE!' } aus Formel-Ergebnissen)
                if (cellValue && typeof cellValue === 'object' && cellValue.error) {
                    // Prüfe ob die zugehörige Formel eine Bild-Formel ist
                    const formula = cellFormulas[styleKey] || '';
                    if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(formula)) {
                        cellValue = '🖼️ Bild';
                    } else if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                        // Bild-Formel aus XML-Metadaten erkannt
                        cellValue = '🖼️ Bild';
                    } else {
                        console.log(`[ExcelJS] Error-Wert in Zelle ${styleKey}: ${cellValue.error}, Formel: ${formula || 'keine'}`);
                        cellValue = String(cellValue.error);
                    }
                }
                // String-Error-Werte: #VALUE! ohne Formel = möglicherweise Bild oder nicht-auswertbar
                else if (typeof cellValue === 'string' && cellValue === '#VALUE!') {
                    const formula = cellFormulas[styleKey] || '';
                    if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(formula)) {
                        cellValue = '🖼️ Bild';
                    } else if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                        // Bild-Formel aus XML-Metadaten erkannt
                        cellValue = '🖼️ Bild';
                    } else {
                        console.log(`[ExcelJS] #VALUE! String in Zelle ${styleKey}, Formel: ${formula || 'keine'}, cell.value Typ: ${typeof cell.value}, Keys: ${cell.value && typeof cell.value === 'object' ? Object.keys(cell.value).join(',') : 'N/A'}`);
                    }
                }
                
                // Hyperlink extrahieren
                if (cell.hyperlink) {
                    cellHyperlinks[styleKey] = cell.hyperlink.hyperlink || cell.hyperlink;
                }
                
                // WICHTIG: Datums-Behandlung VOR der allgemeinen Objekt-Behandlung!
                // Date ist auch ein Objekt, würde sonst mit String() konvertiert werden
                if (cellValue instanceof Date) {
                    // Excel-Format aus numFmt extrahieren (falls vorhanden)
                    const numFmt = cell.numFmt || '';
                    
                    // Prüfe ob es ein Zeit-Format ist (h für Stunden, : für Zeit-Separator)
                    // WICHTIG: 'm' allein ist Monat, nicht Minute!
                    // Minute wird nur nach 'h' oder vor 's' verwendet
                    const hasTime = numFmt.includes('h') || numFmt.includes('H') || numFmt.includes(':');
                    
                    if (hasTime) {
                        // Mit Zeit: ISO-Format verwenden
                        cellValue = cellValue.toISOString().replace('T', ' ').substring(0, 19);
                    } else {
                        // Nur Datum: Format aus numFmt ableiten
                        const day = cellValue.getDate();
                        const month = cellValue.getMonth() + 1;
                        const year = cellValue.getFullYear();
                        
                        // Führende Nullen hinzufügen wenn Format es verlangt
                        const dayStr = numFmt.includes('dd') ? String(day).padStart(2, '0') : String(day);
                        const monthStr = numFmt.includes('mm') ? String(month).padStart(2, '0') : String(month);
                        
                        // Jahr-Format: yyyy = 4 Ziffern, yy = 2 Ziffern
                        let yearStr = String(year);
                        if (!numFmt.includes('yyyy') && numFmt.includes('yy')) {
                            yearStr = yearStr.substring(2);
                        }
                        
                        // Separator bestimmen: ., /, oder -
                        if (numFmt.includes('.')) {
                            // Deutsches Format: D.M.YYYY
                            cellValue = `${dayStr}.${monthStr}.${yearStr}`;
                        } else if (numFmt.includes('-')) {
                            // ISO-ähnlich: M-D-YYYY
                            cellValue = `${monthStr}-${dayStr}-${yearStr}`;
                        } else {
                            // Standard: D.M.YYYY (da ursprüngliche Datei deutsches Format hatte)
                            cellValue = `${dayStr}.${monthStr}.${yearStr}`;
                        }
                    }
                }
                
                // Objekt-Werte behandeln (Rich Text, Hyperlinks, etc.)
                // WICHTIG: Nur wenn es KEINE Formel war (die wurde oben schon behandelt)
                // Wir prüfen cell.value (nicht cellValue), um zu sehen ob es ein spezielles Objekt ist
                if (cell.value && typeof cell.value === 'object' && !cell.formula && !cell.value.formula) {
                    // Rich Text extrahieren
                    if (cell.value.richText) {
                        const richText = cell.value.richText.map(part => ({
                            text: part.text,
                            styles: {
                                bold: part.font?.bold || false,
                                italic: part.font?.italic || false,
                                underline: part.font?.underline || false,
                                strikethrough: part.font?.strike || false,
                                color: part.font?.color?.argb ? `#${part.font.color.argb.substring(2)}` : null,
                                fontSize: part.font?.size || null,
                                fontName: part.font?.name || null
                            }
                        }));
                        richTextCells[styleKey] = richText;
                        // Konvertiere zu Plain Text - nimm den text direkt aus dem Original!
                        cellValue = cell.value.richText.map(part => part.text).join('');
                    }
                    // Hyperlink-Objekte (haben text und hyperlink Properties)
                    else if (cell.value.text !== undefined && cell.value.hyperlink !== undefined) {
                        cellValue = cell.value.text;
                        cellHyperlinks[styleKey] = cell.value.hyperlink;
                    }
                    // Andere Objekte - versuche text-Property zu nutzen
                    else if (cell.value.text !== undefined) {
                        cellValue = cell.value.text;
                    }
                    // Fallback: Null oder leerer String
                    else if (cell.value === null) {
                        cellValue = '';
                    }
                    // Bild-Objekte erkennen (Buffer, Base64 oder image-Properties)
                    else if (cell.value.buffer || cell.value.image || cell.value.imageId || 
                             (cell.value.extension && (cell.value.extension === 'png' || cell.value.extension === 'jpeg' || cell.value.extension === 'gif' || cell.value.extension === 'bmp'))) {
                        cellValue = '🖼️ Bild';
                    }
                    // Error-Objekte erkennen
                    else if (cell.value.error) {
                        // Prüfe ob es eine Bild-Zelle ist (aus XML-Metadaten)
                        if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                            cellValue = '🖼️ Bild';
                        } else {
                            cellValue = cell.value.error; // z.B. #REF!, #VALUE!, #DIV/0!
                        }
                    }
                    // Letzter Fallback: Unbekanntes Objekt -> versuche sinnvolle Darstellung
                    else {
                        // Prüfe ob JSON-Serialisierung "[object Object]" vermeidet
                        const keys = Object.keys(cell.value);
                        if (keys.length === 0) {
                            cellValue = '';
                        } else {
                            // Logge das unbekannte Objekt für Debugging
                            console.log(`[ExcelJS] Unbekanntes Objekt in Zelle ${styleKey}:`, JSON.stringify(cell.value).substring(0, 200));
                            cellValue = '📎 Objekt';
                        }
                    }
                }
                
                // Styles extrahieren
                const style = {};
                
                if (cell.font) {
                    if (cell.font.bold) style.bold = true;
                    if (cell.font.italic) style.italic = true;
                    if (cell.font.underline) style.underline = true;
                    if (cell.font.strike) style.strikethrough = true;
                    if (cell.font.size) {
                        style.fontSize = cell.font.size;
                    }
                    if (cell.font.name && cell.font.name !== 'Calibri') {
                        style.fontName = cell.font.name;
                    }
                    if (cell.font.color?.argb) {
                        const colorHex = cell.font.color.argb.substring(2);
                        if (colorHex !== '000000') {
                            style.fontColor = `#${colorHex}`;
                        }
                    }
                }
                
                // Alignment extrahieren
                if (cell.alignment) {
                    if (cell.alignment.horizontal && cell.alignment.horizontal !== 'general') {
                        style.textAlign = cell.alignment.horizontal;
                    }
                    if (cell.alignment.vertical && cell.alignment.vertical !== 'bottom') {
                        style.verticalAlign = cell.alignment.vertical;
                    }
                    if (cell.alignment.wrapText) {
                        style.wrapText = true;
                    }
                }
                
                // Fill extrahieren
                if (cell.fill) {
                    if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' && cell.fill.fgColor?.argb) {
                        const fillHex = cell.fill.fgColor.argb.substring(2);
                        if (fillHex !== 'FFFFFF') {
                            style.fill = `#${fillHex}`;
                        }
                    }
                }
                
                // Borders extrahieren
                if (cell.border) {
                    const borders = {};
                    for (const side of ['top', 'bottom', 'left', 'right']) {
                        if (cell.border[side] && cell.border[side].style) {
                            borders[side] = {
                                style: cell.border[side].style,
                                color: cell.border[side].color?.argb ? `#${cell.border[side].color.argb.substring(2)}` : null
                            };
                        }
                    }
                    if (Object.keys(borders).length > 0) {
                        style.borders = borders;
                    }
                }
                
                if (Object.keys(style).length > 0) {
                    cellStyles[styleKey] = style;
                }
                
                // WICHTIG: Date-Objekte MÜSSEN hier als String formatiert werden
                // da sie sonst bei der IPC-Serialisierung zu "Thu Sep 19 2013..." werden
                if (cellValue instanceof Date) {
                    // Fallback-Formatierung falls oben nicht gegriffen hat
                    const day = cellValue.getDate();
                    const month = cellValue.getMonth() + 1;
                    const year = cellValue.getFullYear();
                    cellValue = `${day}.${month}.${year}`;
                }
                // Auch String-Werte prüfen die wie Date.toString() aussehen
                else if (typeof cellValue === 'string' && /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s/.test(cellValue)) {
                    // Versuche den Date-String zu parsen
                    const parsedDate = new Date(cellValue);
                    if (!isNaN(parsedDate.getTime())) {
                        const day = parsedDate.getDate();
                        const month = parsedDate.getMonth() + 1;
                        const year = parsedDate.getFullYear();
                        cellValue = `${day}.${month}.${year}`;
                    }
                }
                
                // Setze den Wert an der korrekten Position
                // Letzte Absicherung: Objekte die durchgerutscht sind, nie als "[object Object]" speichern
                if (cellValue !== null && cellValue !== undefined && typeof cellValue === 'object') {
                    console.log(`[ExcelJS] Objekt durchgerutscht in Zelle ${styleKey}:`, JSON.stringify(cellValue).substring(0, 200));
                    cellValue = '📎 Objekt';
                }
                rowData[colIndex] = cellValue === null || cellValue === undefined ? '' : cellValue;
            });
            
            // Versteckte Zeilen - verwende dataRowCounter statt rowNumber
            if (row.hidden) {
                hiddenRows.push(currentDataRowIndex); // 0-basierter Index im Daten-Array
            }
            
            data.push(rowData);
            dataRowCounter++; // Zähler für nächste Daten-Zeile erhöhen
            
            // Fortschritt alle 2 Sekunden loggen
            const now = Date.now();
            if (now - lastProgressLog > 2000) {
                console.log(`[ExcelJS] Streaming: ${dataRowCounter} Zeilen verarbeitet (${now - t2}ms)`);
                lastProgressLog = now;
            }
        } // Ende for-await row
        
            console.log(`[ExcelJS] Streaming abgeschlossen: ${dataRowCounter} Datenzeilen in ${Date.now() - t2}ms`);
            timings.streaming = Date.now() - t2;
            break; // Sheet gefunden, keine weiteren Sheets verarbeiten
        } // Ende for-await worksheetReader
        
        if (!sheetFound) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // WICHTIG: Header-Zeile als erste Zeile in data einfügen
        // Das Frontend erwartet data.slice(1) - also Header an Position 0
        data.unshift(headers);
        
        // ============================================================
        // FALLBACK: Fill-Farben direkt aus XLSX extrahieren
        // ExcelJS erkennt bei manchen Dateien (z.B. SoftMaker) keine Fills
        // NUR wenn ExcelJS keine einzige Fill-Farbe gefunden hat
        // SKIP für große Dateien (>5000 Zeilen) — blockiert den Event-Loop zu lange
        // ============================================================
        const excelJSHasFills = Object.values(cellStyles).some(s => s.fill);
        const dataRowCount = data.length - 1; // Minus Header
        
        if (!excelJSHasFills && dataRowCount <= 5000) {
            console.log(`[ExcelJS] Kein ExcelJS-Fill gefunden, verwende ZIP-Fallback (${dataRowCount} Zeilen)`);
            const t3 = Date.now();
            // Buffer statt Dateipfad verwenden (kein erneuter Dateizugriff!)
            const directFills = extractFillsFromXLSX(fileBuffer, sheetName);
            timings.fillFallback = Date.now() - t3;
            console.log(`[ExcelJS] Fill-Fallback: ${Object.keys(directFills).length} Fills in ${timings.fillFallback}ms`);
            
            if (Object.keys(directFills).length > 0) {
                for (const [key, fillColor] of Object.entries(directFills)) {
                    if (cellStyles[key]) {
                        if (!cellStyles[key].fill) {
                            cellStyles[key].fill = fillColor;
                        }
                    } else {
                        cellStyles[key] = { fill: fillColor };
                    }
                }
            }
        } else if (!excelJSHasFills && dataRowCount > 5000) {
            console.log(`[ExcelJS] Fill-Fallback ÜBERSPRUNGEN (${dataRowCount} Zeilen > 5000 — zu langsam)`);
            timings.fillFallback = 0;
        }
        
        const totalTime = Date.now() - startTime;
        
        // Zeilenfarben erkennen (wenn alle Zellen einer Zeile die gleiche Hintergrundfarbe haben)
        const t4 = Date.now();
        const rowHighlights = detectRowHighlights(cellStyles, data.length, headers.length);
        timings.highlights = Date.now() - t4;
        
        console.log(`[ExcelJS] === FERTIG === ${data.length} Zeilen, ${headers.length} Spalten, ${Object.keys(cellStyles).length} Styles in ${totalTime}ms`);
        console.log(`[ExcelJS] Timings: fileRead=${timings.fileRead}ms, metadata=${timings.metadata}ms, streaming=${timings.streaming}ms, fillFallback=${timings.fillFallback || 0}ms, highlights=${timings.highlights}ms`);
        
        return {
            success: true,
            headers,
            data,
            hiddenColumns,
            hiddenRows,
            cellStyles,
            cellFormulas,
            cellHyperlinks,
            richTextCells,
            mergedCells,
            autoFilterRange,
            rowHighlights,  // NEU: Zeilenfarben als Array von [rowIndex, colorName]
            stats: {
                rows: data.length,
                columns: headers.length,
                loadTimeMs: totalTime,
                timings // Detaillierte Zeitmessungen für Diagnose
            }
        };
        
    } catch (error) {
        console.error('[ExcelJS] Fehler beim Laden:', error);
        return { success: false, error: error.message };
    } finally {
        // Temporäre entschlüsselte Datei aufräumen
        if (tempFilePath) {
            try {
                const fs = require('fs');
                if (fs.existsSync(tempFilePath)) {
                    fs.unlinkSync(tempFilePath);
                }
            } catch (cleanupError) {
                console.warn('[ExcelJS] Konnte temporäre Datei nicht löschen:', cleanupError.message);
            }
        }
    }
}

module.exports = {
    readSheetWithExcelJS,
    extractFillsFromXLSX
};
