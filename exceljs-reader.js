// ============================================================================
// EXCELJS MIGRATION - NEUE READ-FUNKTION
// ============================================================================
// Dieses Modul enthält die ExcelJS-basierte Sheet-Read-Funktion
// Zum Testen der Migration von xlsx-populate zu exceljs

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');
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
function extractFillsFromXLSX(filePath, sheetName) {
    const cellFills = {};
    
    try {
        const zip = new AdmZip(filePath);
        
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
 * Liest ein Excel-Sheet mit ExcelJS (Alternative zu xlsx-populate)
 * 
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des zu lesenden Sheets
 * @param {string|null} password - Optional: Passwort für geschützte Dateien
 * @returns {Promise<Object>} Sheet-Daten im gleichen Format wie xlsx-populate
 */
async function readSheetWithExcelJS(filePath, sheetName, password = null) {
    const startTime = Date.now();
    let tempFilePath = null; // Für entschlüsselte Dateien
    
    try {
        const workbook = new ExcelJS.Workbook();
        
        const loadStart = Date.now();
        
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
                tempFilePath = path.join(os.tmpdir(), `mvms_decrypt_${Date.now()}_${path.basename(filePath)}`);
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
        
        // Datei laden (mit oder ohne vorherige Entschlüsselung)
        try {
            await workbook.xlsx.readFile(actualFilePath);
        } catch (readError) {
            // Prüfe ob die Datei passwortgeschützt ist (ohne Passwort versucht zu öffnen)
            if (!password && (
                readError.message.includes('password') || 
                readError.message.includes('Password') ||
                readError.message.includes('encrypted') ||
                readError.message.includes('Encrypted') ||
                readError.message.includes('CFB') // Common error for encrypted files
            )) {
                return { 
                    success: false, 
                    error: 'Diese Datei ist passwortgeschützt. Bitte Passwort eingeben.',
                    needsPassword: true
                };
            }
            throw readError;
        }
        
        // Sheet finden
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Daten-Strukturen initialisieren
        const headers = [];
        const data = [];
        const hiddenColumns = [];
        const hiddenRows = [];
        const cellStyles = {};
        const cellFormulas = {};
        const cellHyperlinks = {};
        const richTextCells = {};
        
        // AutoFilter-Bereich - kann String oder Objekt sein
        // Auch aus Excel-Tabellen (Tables) extrahieren
        let autoFilterRange = null;
        if (worksheet.autoFilter) {
            if (typeof worksheet.autoFilter === 'string') {
                autoFilterRange = worksheet.autoFilter;
            } else if (worksheet.autoFilter.ref) {
                autoFilterRange = worksheet.autoFilter.ref;
            }
        }
        
        // Falls kein direkter AutoFilter, prüfe Excel-Tabellen
        if (!autoFilterRange && worksheet.tables) {
            for (const tableName of Object.keys(worksheet.tables)) {
                const table = worksheet.tables[tableName];
                if (table && table.table) {
                    // Erst autoFilterRef prüfen, dann tableRef als Fallback
                    // Excel-Tabellen haben immer einen AutoFilter über den gesamten Tabellenbereich
                    if (table.table.autoFilterRef) {
                        autoFilterRange = table.table.autoFilterRef;
                        break;
                    } else if (table.table.tableRef) {
                        autoFilterRange = table.table.tableRef;
                        break;
                    }
                }
            }
        }
        
        // Merged Cells (verbundene Zellen) extrahieren und konvertieren
        // ExcelJS gibt Strings wie "A1:H1" zurück, Frontend erwartet Objekte
        const mergedCells = [];
        if (worksheet.model && worksheet.model.merges) {
            worksheet.model.merges.forEach(rangeStr => {
                // Parse "A1:H1" zu { startRow, startCol, endRow, endCol, rowSpan, colSpan }
                const parsed = parseRangeString(rangeStr);
                if (parsed) {
                    mergedCells.push(parsed);
                }
            });
        }
        
        // Versteckte Spalten ermitteln
        worksheet.columns.forEach((col, colIndex) => {
            if (col.hidden) {
                hiddenColumns.push(colIndex);
            }
        });
        
        // Ermittle die tatsächliche Spaltenanzahl (kann mehr sein als in Zeile 1)
        const actualColumnCount = worksheet.columnCount;
        
        // Zähler für tatsächliche Daten-Zeilen (ohne Header)
        // WICHTIG: dataRowCounter ist jetzt gleich rowNumber-2, da wir includeEmpty:true verwenden
        // Das ist wichtig für mergedCells, die auf echten Excel-Zeilen-Indizes basieren
        let dataRowCounter = 0;
        
        // Zeilen durchgehen - WICHTIG: includeEmpty:true damit Zeilen-Indizes mit mergedCells übereinstimmen!
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            // Erste Zeile = Header
            if (rowNumber === 1) {
                // Initialisiere Header-Array mit leeren Strings für alle Spalten
                for (let i = 0; i < actualColumnCount; i++) {
                    headers.push('');
                }
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const colIndex = colNumber - 1;
                    // Überschreibe den leeren Wert mit dem tatsächlichen Wert
                    headers[colIndex] = cell.value ? String(cell.value) : '';
                    
                    // WICHTIG: Auch Header-Styles extrahieren (für Frontend-Kompatibilität)
                    const styleKey = `0-${colIndex}`; // Header = Zeile 0
                    const style = {};
                    
                    if (cell.font) {
                        if (cell.font.bold) style.bold = true;
                        if (cell.font.italic) style.italic = true;
                        if (cell.font.underline) style.underline = true;
                        if (cell.font.strike) style.strikethrough = true;
                        if (cell.font.size && cell.font.size !== 11) {
                            style.fontSize = cell.font.size;
                        }
                        if (cell.font.color?.argb) {
                            const colorHex = cell.font.color.argb.substring(2);
                            // Ignoriere Standard-Schwarz (000000)
                            if (colorHex !== '000000') {
                                style.fontColor = `#${colorHex}`;
                            }
                        }
                    }
                    
                    // Fill extrahieren - prüfe auch pattern === 'solid'
                    if (cell.fill) {
                        if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' && cell.fill.fgColor?.argb) {
                            const fillHex = cell.fill.fgColor.argb.substring(2);
                            if (fillHex !== 'FFFFFF') {
                                style.fill = `#${fillHex}`;
                            }
                        }
                    }
                    
                    if (Object.keys(style).length > 0) {
                        cellStyles[styleKey] = style;
                    }
                });
                return; // Weiter zur nächsten Zeile
            }
            
            // Daten-Zeilen
            // Initialisiere rowData mit leeren Strings für alle Spalten
            const rowData = new Array(actualColumnCount).fill('');
            
            // WICHTIG: Style-Key basiert auf dataRowCounter, nicht auf rowNumber!
            // Das stellt sicher, dass leere Zeilen nicht zu Index-Mismatches führen
            const currentDataRowIndex = dataRowCounter;
            
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const colIndex = colNumber - 1;
                // WICHTIG: Frontend erwartet 1-basierte Indizes (wie xlsx-populate)
                const styleKey = `${currentDataRowIndex + 1}-${colIndex}`;
                
                let cellValue = cell.value;
                
                // Formel extrahieren - WICHTIG: VOR der Objekt-Behandlung!
                // Bei Formeln kann cell.value ein Objekt sein mit { formula, result }
                // oder cell.formula ist direkt verfügbar
                if (cell.formula) {
                    cellFormulas[styleKey] = cell.formula;
                    // Das Ergebnis ist in cell.result (nicht cell.value!)
                    cellValue = cell.result !== undefined ? cell.result : '';
                } else if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
                    // Formel als Objekt gespeichert: { formula: '...', result: ... }
                    cellFormulas[styleKey] = cell.value.formula;
                    cellValue = cell.value.result !== undefined ? cell.value.result : '';
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
                    // Letzter Fallback: Unbekanntes Objekt -> versuche String-Konvertierung
                    else {
                        cellValue = String(cell.value);
                    }
                }
                
                // Styles extrahieren
                const style = {};
                
                if (cell.font) {
                    if (cell.font.bold) style.bold = true;
                    if (cell.font.italic) style.italic = true;
                    // ExcelJS verwendet underline: "single", "double", etc. statt true
                    if (cell.font.underline) style.underline = true;
                    if (cell.font.strike) style.strikethrough = true;
                    if (cell.font.size && cell.font.size !== 11) {
                        style.fontSize = cell.font.size;
                    }
                    if (cell.font.color?.argb) {
                        const colorHex = cell.font.color.argb.substring(2);
                        // Ignoriere Standard-Schwarz (000000)
                        if (colorHex !== '000000') {
                            style.fontColor = `#${colorHex}`;
                        }
                    }
                }
                
                // Fill extrahieren - prüfe auch pattern === 'solid'
                if (cell.fill) {
                    if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' && cell.fill.fgColor?.argb) {
                        const fillHex = cell.fill.fgColor.argb.substring(2);
                        // Ignoriere Weiß (FFFFFF)
                        if (fillHex !== 'FFFFFF') {
                            style.fill = `#${fillHex}`;
                        }
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
                rowData[colIndex] = cellValue === null || cellValue === undefined ? '' : cellValue;
            });
            
            // Versteckte Zeilen - verwende dataRowCounter statt rowNumber
            if (row.hidden) {
                hiddenRows.push(currentDataRowIndex); // 0-basierter Index im Daten-Array
            }
            
            data.push(rowData);
            dataRowCounter++; // Zähler für nächste Daten-Zeile erhöhen
        });
        
        // WICHTIG: Header-Zeile als erste Zeile in data einfügen
        // Das Frontend erwartet data.slice(1) - also Header an Position 0
        data.unshift(headers);
        
        // ============================================================
        // FALLBACK: Fill-Farben direkt aus XLSX extrahieren
        // ExcelJS erkennt bei manchen Dateien (z.B. SoftMaker) keine Fills
        // ============================================================
        const directFills = extractFillsFromXLSX(filePath, sheetName);
        
        if (Object.keys(directFills).length > 0) {
            // Füge fehlende Fills zu cellStyles hinzu
            for (const [key, fillColor] of Object.entries(directFills)) {
                if (cellStyles[key]) {
                    // Style existiert, aber vielleicht fehlt fill
                    if (!cellStyles[key].fill) {
                        cellStyles[key].fill = fillColor;
                    }
                } else {
                    // Neuer Style nur mit fill
                    cellStyles[key] = { fill: fillColor };
                }
            }
        }
        
        const totalTime = Date.now() - startTime;
        
        // Zeilenfarben erkennen (wenn alle Zellen einer Zeile die gleiche Hintergrundfarbe haben)
        const rowHighlights = detectRowHighlights(cellStyles, data.length, headers.length);
        
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
                loadTimeMs: totalTime
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
