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
            console.log('[XLSX-Extract] Keine styles.xml gefunden');
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
        console.log('[XLSX-Extract] Gefundene Fills:', fills);
        
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
        console.log('[XLSX-Extract] Style zu Fill Mapping:', styleToFill);
        
        // 3. Sheet-Daten finden
        // Zuerst workbook.xml lesen um Sheet rId zu finden
        const workbookEntry = zip.getEntry('xl/workbook.xml');
        if (!workbookEntry) {
            console.log('[XLSX-Extract] Keine workbook.xml gefunden');
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
        console.log(`[XLSX-Extract] Sheet "${sheetName}" hat rId: ${sheetRId}`);
        
        // Relationship-Datei lesen um tatsächlichen Sheet-Pfad zu finden
        const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
        if (!relsEntry) {
            console.log('[XLSX-Extract] Keine workbook.xml.rels gefunden');
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
            console.log(`[XLSX-Extract] Kein Pfad für ${sheetRId} gefunden`);
            return cellFills;
        }
        console.log(`[XLSX-Extract] Sheet-Pfad: ${sheetPath}`);
        
        // Sheet-XML laden (Pfad kann relativ sein, z.B. "worksheets/sheet1.xml")
        const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : `xl/${sheetPath}`;
        const sheetEntry = zip.getEntry(fullSheetPath);
        if (!sheetEntry) {
            console.log(`[XLSX-Extract] Sheet ${fullSheetPath} nicht gefunden`);
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
                    console.log(`[XLSX-Extract] Zelle ${colLetters}${rowNum} (Key: ${key}): Fill ${fillColor}`);
                }
            }
        }
        
        console.log('[XLSX-Extract] Extrahierte Fills:', cellFills);
        return cellFills;
        
    } catch (error) {
        console.error('[XLSX-Extract] Fehler:', error);
        return cellFills;
    }
}

// Debug-Log-Datei
const DEBUG_LOG = '/Users/nojan/Desktop/exceljs-debug.log';

function debugLog(message) {
    const timestamp = new Date().toISOString();
    const logLine = `[${timestamp}] ${message}\n`;
    fs.appendFileSync(DEBUG_LOG, logLine);
    console.log(message);
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
        
        console.log('[ExcelJS] Lade Workbook...');
        const loadStart = Date.now();
        
        // Bei passwortgeschützten Dateien: xlsx-populate zum Entschlüsseln verwenden
        // ExcelJS hat bekannte Probleme mit Passwort-Entschlüsselung
        let actualFilePath = filePath;
        
        if (password) {
            console.log('[ExcelJS] Passwortgeschützte Datei - verwende xlsx-populate zum Entschlüsseln...');
            
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
                console.log('[ExcelJS] Datei entschlüsselt, lade mit ExcelJS...');
                
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
                console.log('[ExcelJS] Datei scheint passwortgeschützt zu sein');
                return { 
                    success: false, 
                    error: 'Diese Datei ist passwortgeschützt. Bitte Passwort eingeben.',
                    needsPassword: true
                };
            }
            throw readError;
        }
        
        console.log(`[ExcelJS] Workbook geladen in ${Date.now() - loadStart}ms`);
        
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
                        debugLog('[ExcelJS] AutoFilter aus Tabelle (autoFilterRef):', tableName, autoFilterRange);
                        break;
                    } else if (table.table.tableRef) {
                        autoFilterRange = table.table.tableRef;
                        debugLog('[ExcelJS] AutoFilter aus Tabelle (tableRef):', tableName, autoFilterRange);
                        break;
                    }
                }
            }
        }
        debugLog('[ExcelJS] AutoFilter Range:', autoFilterRange);
        
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
        console.log('[ExcelJS] Merged Cells:', mergedCells);
        
        // Versteckte Spalten ermitteln
        worksheet.columns.forEach((col, colIndex) => {
            if (col.hidden) {
                hiddenColumns.push(colIndex);
            }
        });
        
        // Ermittle die tatsächliche Spaltenanzahl (kann mehr sein als in Zeile 1)
        const actualColumnCount = worksheet.columnCount;
        console.log('[ExcelJS] Actual column count:', actualColumnCount);
        
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
                
                // Formel extrahieren
                if (cell.formula) {
                    cellFormulas[styleKey] = cell.formula;
                    cellValue = cell.result || cellValue;
                }
                
                // Hyperlink extrahieren
                if (cell.hyperlink) {
                    cellHyperlinks[styleKey] = cell.hyperlink.hyperlink || cell.hyperlink;
                }
                
                // Objekt-Werte behandeln (Rich Text, Hyperlinks, etc.)
                if (cell.value && typeof cell.value === 'object') {
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
                
                // Datum formatieren
                if (cellValue instanceof Date) {
                    cellValue = cellValue.toISOString().split('T')[0];
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
        console.log('[ExcelJS] Prüfe auf fehlende Fills...');
        const directFills = extractFillsFromXLSX(filePath, sheetName);
        
        if (Object.keys(directFills).length > 0) {
            console.log(`[ExcelJS] ${Object.keys(directFills).length} Fills aus XLSX extrahiert`);
            
            // Füge fehlende Fills zu cellStyles hinzu
            for (const [key, fillColor] of Object.entries(directFills)) {
                if (cellStyles[key]) {
                    // Style existiert, aber vielleicht fehlt fill
                    if (!cellStyles[key].fill) {
                        cellStyles[key].fill = fillColor;
                        console.log(`[ExcelJS] Fill hinzugefügt für ${key}: ${fillColor}`);
                    }
                } else {
                    // Neuer Style nur mit fill
                    cellStyles[key] = { fill: fillColor };
                    console.log(`[ExcelJS] Neuer Style für ${key}: ${fillColor}`);
                }
            }
        }
        
        const totalTime = Date.now() - startTime;
        debugLog(`[ExcelJS] Sheet geladen in ${totalTime}ms (${data.length} Zeilen)`);
        debugLog(`[ExcelJS] Extrahierte Styles: ${Object.keys(cellStyles).length}`);
        if (Object.keys(cellStyles).length > 0) {
            debugLog('[ExcelJS] Beispiel-Styles: ' + JSON.stringify(Object.entries(cellStyles).slice(0, 10)));
        }
        debugLog('[ExcelJS] Alle Style-Keys: ' + Object.keys(cellStyles).join(', '));
        
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
                    console.log('[ExcelJS] Temporäre Datei gelöscht:', tempFilePath);
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
