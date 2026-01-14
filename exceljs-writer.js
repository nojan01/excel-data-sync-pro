// ============================================================================
// EXCELJS MIGRATION - WRITE-FUNKTION (VOLLSTÄNDIG)
// ============================================================================
// Export/Save-Funktion mit ExcelJS - ersetzt xlsx-populate komplett

const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');
const { extractFillsFromXLSX } = require('./exceljs-reader');

/**
 * Konvertiert Spaltenbuchstaben zu 1-basiertem Index (A=1, B=2, AA=27, etc.)
 */
function colLetterToNumber(letters) {
    let num = 0;
    for (let i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

/**
 * Konvertiert 1-basierten Index zu Spaltenbuchstaben (1=A, 2=B, 27=AA, etc.)
 */
function colNumberToLetter(num) {
    let result = '';
    while (num > 0) {
        const remainder = (num - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        num = Math.floor((num - 1) / 26);
    }
    return result;
}

/**
 * Prüft ob eine CF-Referenz NUR auf eine bestimmte Spalte verweist
 * @param {string} ref - CF-Referenz wie "BI1:BI100" oder "BH1 BI2:BI50"
 * @param {number} colNumber - Spaltennummer die geprüft wird
 * @returns {boolean} true wenn alle Referenzen nur auf diese Spalte zeigen
 */
function refOnlyReferencesColumn(ref, colNumber) {
    const targetCol = colNumberToLetter(colNumber);
    const ranges = ref.split(' ');
    
    for (const range of ranges) {
        const parts = range.split(':');
        for (const part of parts) {
            const match = part.match(/^([A-Z]+)/);
            if (match && match[1] !== targetCol) {
                return false;
            }
        }
    }
    return true;
}

/**
 * Entfernt oder reduziert Referenzen zu einer bestimmten Spalte aus einer Multi-Range-Ref
 * Beispiele:
 * - "BI1:BI50" mit Spalte 61 (BI) → wird komplett entfernt
 * - "BG1:BI50" mit Spalte 61 (BI) → wird zu "BG1:BH50" (Endspaltereduziert)
 * - "BH1:BH50 BI1:BI50" mit Spalte 61 (BI) → wird zu "BH1:BH50"
 * 
 * @param {string} ref - CF-Referenz wie "BH1:BH50 BI1:BI50"
 * @param {number} colNumber - Spaltennummer die entfernt werden soll
 * @returns {string} Bereinigte Referenz
 */
function removeColumnFromRef(ref, colNumber) {
    const targetCol = colNumberToLetter(colNumber);
    const prevCol = colNumber > 1 ? colNumberToLetter(colNumber - 1) : null;
    const ranges = ref.split(' ');
    
    const processedRanges = [];
    
    for (const range of ranges) {
        const parts = range.split(':');
        
        if (parts.length === 2) {
            // Bereich wie "BG1:BI50"
            const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
            const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
            
            if (startMatch && endMatch) {
                const startCol = startMatch[1];
                const startRow = startMatch[2];
                const endCol = endMatch[1];
                const endRow = endMatch[2];
                
                if (startCol === targetCol && endCol === targetCol) {
                    // Kompletter Bereich in der Zielspalte - entfernen
                    continue;
                } else if (endCol === targetCol && prevCol) {
                    // Endspalte ist die Zielspalte - reduzieren
                    processedRanges.push(startCol + startRow + ':' + prevCol + endRow);
                } else if (startCol === targetCol) {
                    // Startspalte ist die Zielspalte - sollte nicht vorkommen nach Verschiebung
                    // aber für Vollständigkeit: behalte den Rest
                    const nextCol = colNumberToLetter(colNumber + 1);
                    processedRanges.push(nextCol + startRow + ':' + endCol + endRow);
                } else {
                    // Bereich enthält die Spalte nicht oder ist ok
                    processedRanges.push(range);
                }
            } else {
                processedRanges.push(range);
            }
        } else {
            // Einzelne Zelle wie "BI5"
            const match = range.match(/^([A-Z]+)/);
            if (match && match[1] !== targetCol) {
                processedRanges.push(range);
            }
            // Wenn es die Zielspalte ist, entfernen (nicht hinzufügen)
        }
    }
    
    return processedRanges.join(' ') || ref; // Fallback auf original wenn alles entfernt würde
}

/**
 * Passt alle Tables nach dem Löschen einer Spalte an
 * Tables haben ihre eigene tableRef und autoFilterRef die angepasst werden müssen.
 * Außerdem muss die entsprechende Spalte aus table.columns entfernt werden.
 * 
 * @param {Worksheet} worksheet - Das ExcelJS Worksheet
 * @param {number} deletedColNumber - 1-basierte Spaltennummer die gelöscht wurde
 */
function adjustTablesAfterColumnDelete(worksheet, deletedColNumber) {
    if (!worksheet.tables || Object.keys(worksheet.tables).length === 0) {
        return;
    }
    
    Object.keys(worksheet.tables).forEach(tableName => {
        const tableEntry = worksheet.tables[tableName];
        const table = tableEntry.table;
        
        if (!table) {
            return;
        }
        
        // Anpassen von tableRef (z.B. "A1:BI2404" -> "A1:BH2404")
        if (table.tableRef) {
            const oldRef = table.tableRef;
            const newRef = reduceAutoFilterRange(oldRef);
            table.tableRef = newRef;
        }
        
        // Anpassen von autoFilterRef (z.B. "A1:BI2404" -> "A1:BH2404")
        if (table.autoFilterRef) {
            const oldRef = table.autoFilterRef;
            const newRef = reduceAutoFilterRange(oldRef);
            table.autoFilterRef = newRef;
        }
        
        // Entferne die gelöschte Spalte aus table.columns
        if (table.columns && Array.isArray(table.columns)) {
            const oldCount = table.columns.length;
            // deletedColNumber ist 1-basiert, Array ist 0-basiert
            const indexToRemove = deletedColNumber - 1;
            if (indexToRemove >= 0 && indexToRemove < table.columns.length) {
                table.columns.splice(indexToRemove, 1);
            }
        }
    });
}

/**
 * Reduziert den AutoFilter-Bereich um 1 Spalte (nach Spaltenlöschung)
 * @param {string} autoFilterRange - z.B. "A1:BJ2404"
 * @returns {string} Reduzierter Bereich z.B. "A1:BI2404"
 */
function reduceAutoFilterRange(autoFilterRange) {
    const match = autoFilterRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) return autoFilterRange;
    
    const startCol = match[1];
    const startRow = match[2];
    const endCol = match[3];
    const endRow = match[4];
    
    const endColNum = colLetterToNumber(endCol);
    if (endColNum > 1) {
        const newEndCol = colNumberToLetter(endColNum - 1);
        return startCol + startRow + ':' + newEndCol + endRow;
    }
    return autoFilterRange;
}

/**
 * Passt eine Zellreferenz (z.B. "AN2135") um deletedCol nach links an
 * @param {string} cellRef - Zellreferenz wie "AN2135"
 * @param {number} deletedColNumber - 1-basierte Spaltennummer die gelöscht wurde
 * @returns {string} Angepasste Referenz
 */
function adjustCellReference(cellRef, deletedColNumber) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return cellRef;
    
    const colLetter = match[1];
    const rowNumber = match[2];
    const colNumber = colLetterToNumber(colLetter);
    
    if (colNumber > deletedColNumber) {
        // Spalte nach rechts vom gelöschten - verschiebe nach links
        return colNumberToLetter(colNumber - 1) + rowNumber;
    } else if (colNumber === deletedColNumber) {
        // Referenziert genau die gelöschte Spalte - bleibt problematisch
        // Hier geben wir die angepasste Referenz zurück (wird ungültig)
        return colNumberToLetter(colNumber) + rowNumber;
    }
    return cellRef;
}

/**
 * Passt einen Bereich (z.B. "AN2135:AY2404") um deletedCol nach links an
 * Unterstützt auch Multi-Range-Referenzen wie "L1811 L1812:M1875" (Leerzeichen-separiert)
 * @param {string} rangeRef - Bereichsreferenz wie "AN2135:AY2404" oder "L1811 L1812:M1875"
 * @param {number} deletedColNumber - 1-basierte Spaltennummer die gelöscht wurde
 * @returns {string} Angepasster Bereich
 */
function adjustRangeReference(rangeRef, deletedColNumber) {
    // Multi-Range-Referenzen sind Leerzeichen-separiert
    const multiRanges = rangeRef.split(' ');
    
    const adjustedRanges = multiRanges.map(singleRange => {
        const parts = singleRange.split(':');
        if (parts.length === 2) {
            // A1:B10 Format
            return adjustCellReference(parts[0], deletedColNumber) + ':' + adjustCellReference(parts[1], deletedColNumber);
        }
        // Einzelne Zelle wie "L1811"
        return adjustCellReference(singleRange, deletedColNumber);
    });
    
    return adjustedRanges.join(' ');
}

/**
 * Passt alle bedingten Formatierungen nach Spaltenlöschung an
 * 
 * EINFACHE LOGIK:
 * 1. Alle CF-Referenzen um 1 nach links verschieben (B→A, C→B, usw.)
 * 2. CF die auf die letzte Spalte (jetzt leer) zeigt → entfernen
 * 
 * @param {Worksheet} worksheet - Das ExcelJS Worksheet
 * @param {number} deletedColNumber - 1-basierte Spaltennummer die gelöscht wurde
 * @param {number} lastColumnBeforeDelete - Die letzte Spalte vor dem Löschen
 */
function adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColNumber, lastColumnBeforeDelete) {
    const cf = worksheet.conditionalFormattings;
    if (!cf || !Array.isArray(cf) || cf.length === 0) {
        return;
    }
    
    // Nach dem Löschen ist die letzte Spalte jetzt leer
    const emptyLastCol = lastColumnBeforeDelete;
    
    let adjustedCount = 0;
    let removedCount = 0;
    const toRemove = [];
    
    cf.forEach((cfEntry, idx) => {
        if (!cfEntry.ref) return;
        
        const oldRef = cfEntry.ref;
        
        // Schritt 1: Verschiebe alle Referenzen nach links (Spalten > deletedCol werden um 1 reduziert)
        const adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
        
        // Schritt 2: Prüfe ob die VERSCHOBENE Referenz NUR auf die leere letzte Spalte zeigt
        if (refOnlyReferencesColumn(adjustedRef, emptyLastCol)) {
            toRemove.push(idx);
            removedCount++;
            return;
        }
        
        // Schritt 3: Entferne die leere letzte Spalte aus Multi-Range-Refs (falls enthalten)
        const cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
        
        // IMMER die angepasste Referenz setzen, nicht nur wenn cleanedRef != oldRef
        if (cleanedRef !== oldRef) {
            cfEntry.ref = cleanedRef;
            adjustedCount++;
        }
        
        // Auch Formeln in den Regeln anpassen
        if (cfEntry.rules) {
            cfEntry.rules.forEach(rule => {
                if (rule.formulae && Array.isArray(rule.formulae)) {
                    rule.formulae = rule.formulae.map(formula => {
                        return formula.replace(/\$?([A-Z]+)\$?(\d+)/g, (match, col, row) => {
                            const colNum = colLetterToNumber(col);
                            if (colNum > deletedColNumber) {
                                const newCol = colNumberToLetter(colNum - 1);
                                return match.replace(col, newCol);
                            }
                            return match;
                        });
                    });
                }
            });
        }
    });
    
    // Entferne CF-Regeln (von hinten nach vorne)
    for (let i = toRemove.length - 1; i >= 0; i--) {
        cf.splice(toRemove[i], 1);
    }
}

/**
 * Passt alle Merged Cells VOR Spaltenlöschung an
 * WICHTIG: Diese Funktion muss VOR spliceColumns aufgerufen werden!
 * 
 * Die Logik:
 * 1. Master-Zellen Werte und Styles speichern
 * 2. Alle Merges entfernen (unmerge)
 * 3. Neue Merge-Bereiche berechnen
 * 4. Merges werden NACH spliceColumns wieder gesetzt mit den gespeicherten Werten
 * 
 * @param {Worksheet} worksheet - Das ExcelJS Worksheet
 * @param {number} deletedColNumber - 1-basierte Spaltennummer die gelöscht wird
 * @returns {Array} Array von Objekten mit {range, value, fill} die NACH spliceColumns gesetzt werden sollen
 */
function prepareMergedCellsForColumnDelete(worksheet, deletedColNumber) {
    const merges = worksheet.model.merges;
    if (!merges || !Array.isArray(merges) || merges.length === 0) {
        return [];
    }
    
    // Alle alten Merges speichern
    const oldMerges = [...merges];
    const newMerges = [];
    
    // Berechne neue Merge-Bereiche und speichere Master-Werte
    oldMerges.forEach(mergeRange => {
        const parts = mergeRange.split(':');
        if (parts.length !== 2) return;
        
        // Parse Start und Ende
        const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
        const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
        if (!startMatch || !endMatch) return;
        
        const startCol = colLetterToNumber(startMatch[1]);
        const startRow = parseInt(startMatch[2]);
        const endCol = colLetterToNumber(endMatch[1]);
        const endRow = parseInt(endMatch[2]);
        
        // Hole den Wert und Style der Master-Zelle (oben-links)
        const masterCell = worksheet.getCell(startRow, startCol);
        const masterValue = masterCell.value;
        const masterFill = masterCell.fill;
        
        // Prüfe ob die gelöschte Spalte innerhalb des Merge-Bereichs liegt
        if (deletedColNumber >= startCol && deletedColNumber <= endCol) {
            // Gelöschte Spalte ist INNERHALB des Merge
            if (startCol === endCol) {
                // Merge ist nur 1 Spalte breit und diese wurde gelöscht
                // → Merge komplett entfernen
            } else if (deletedColNumber === startCol) {
                // Die ERSTE Spalte des Merge wird gelöscht
                // Der neue Master ist die nächste Spalte (die nach spliceColumns an startCol steht)
                const newEndCol = endCol - 1;
                const newRange = colNumberToLetter(startCol) + startRow + ':' + colNumberToLetter(newEndCol) + endRow;
                newMerges.push({ range: newRange, value: masterValue, fill: masterFill });
            } else {
                // Eine Spalte in der MITTE oder am ENDE des Merge wird gelöscht
                const newEndCol = endCol - 1;
                const newRange = colNumberToLetter(startCol) + startRow + ':' + colNumberToLetter(newEndCol) + endRow;
                newMerges.push({ range: newRange, value: masterValue, fill: masterFill });
            }
        } else if (deletedColNumber < startCol) {
            // Gelöschte Spalte ist LINKS vom Merge → verschiebe nach links
            const newStartCol = startCol - 1;
            const newEndCol = endCol - 1;
            const newRange = colNumberToLetter(newStartCol) + startRow + ':' + colNumberToLetter(newEndCol) + endRow;
            newMerges.push({ range: newRange, value: masterValue, fill: masterFill });
        } else {
            // Gelöschte Spalte ist RECHTS vom Merge → keine Änderung nötig
            newMerges.push({ range: mergeRange, value: masterValue, fill: masterFill });
        }
    });
    
    // Alle alten Merges entfernen VOR spliceColumns
    oldMerges.forEach(mergeRange => {
        try {
            worksheet.unMergeCells(mergeRange);
        } catch (e) {
            // Ignorieren wenn Bereich nicht existiert
        }
    });
    
    return newMerges;
}

/**
 * Setzt die Merged Cells nach spliceColumns mit den gespeicherten Werten
 * und entfernt cellStyles für Slave-Zellen im Merged-Bereich
 * @param {Worksheet} worksheet - Das ExcelJS Worksheet
 * @param {Array} mergeInfos - Array von Objekten mit {range, value, fill}
 * @param {Object} cellStyles - Die cellStyles vom Frontend (wird modifiziert!)
 * @returns {Array} Array von Zell-Keys die aus cellStyles entfernt wurden
 */
function applyMergedCellsAfterColumnDelete(worksheet, mergeInfos, cellStyles) {
    const removedKeys = [];
    
    if (!mergeInfos || mergeInfos.length === 0) {
        return removedKeys;
    }
    
    mergeInfos.forEach(mergeInfo => {
        try {
            worksheet.unMergeCells(mergeRange);
        } catch (e) {
            // Ignorieren wenn Bereich nicht existiert
        }
    });
    
    // Alle alten Merges entfernen VOR spliceColumns
    oldMerges.forEach(mergeRange => {
        try {
            const range = mergeInfo.range;
            const value = mergeInfo.value;
            const fill = mergeInfo.fill;
            
            // Parse den neuen Range um Start und Ende zu finden
            const parts = range.split(':');
            if (parts.length !== 2) return;
            
            const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
            const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
            if (!startMatch || !endMatch) return;
            
            const startCol = colLetterToNumber(startMatch[1]);
            const startRow = parseInt(startMatch[2]);
            const endCol = colLetterToNumber(endMatch[1]);
            const endRow = parseInt(endMatch[2]);
            
            // Setze zuerst den Wert auf die neue Master-Zelle
            const masterCell = worksheet.getCell(startRow, startCol);
            if (value !== undefined && value !== null) {
                masterCell.value = value;
            }
            if (fill) {
                masterCell.fill = fill;
            }
            
            // WICHTIG: Entferne cellStyles nur für SLAVE-Zellen im Merged-Bereich
            // Die Master-Zelle (startRow, startCol) behält ihren cellStyle!
            // Dies verhindert, dass applyMissingFills falsche Fills auf Slave-Zellen anwendet,
            // während die Master-Zelle ihren korrekten Fill behält
            //
            // cellStyles Format: "rowIdx-colIdx" wobei:
            // - rowIdx = Excel-Zeile - 1 (Zeile 1 = Index 0)
            // - colIdx = Excel-Spalte - 1 (Spalte A = Index 0)
            if (cellStyles) {
                for (let excelRow = startRow; excelRow <= endRow; excelRow++) {
                    for (let excelCol = startCol; excelCol <= endCol; excelCol++) {
                        // WICHTIG: Master-Zelle NICHT entfernen!
                        if (excelRow === startRow && excelCol === startCol) {
                            continue; // Master-Zelle behalten
                        }
                        
                        // Konvertiere Excel 1-basiert zu cellStyles 0-basiert
                        const rowIdx = excelRow - 1;
                        const colIdx = excelCol - 1;
                        const key = `${rowIdx}-${colIdx}`;
                        
                        if (cellStyles[key]) {
                            delete cellStyles[key];
                            removedKeys.push(key);
                        }
                    }
                }
            }
            
            // Dann merge
            worksheet.mergeCells(range);
        } catch (e) {
            // Merge fehlgeschlagen - ignorieren
        }
    });

    return removedKeys;
}

/**
 * Post-Processing: Kopiert styles.xml und Zeilen-Styles von Original nach Export
 * Dies erhält Theme-Colors und andere Styles, die ExcelJS nicht richtig handhabt.
 * 
 * @param {string} sourcePath - Pfad zur Original-Quelldatei
 * @param {string} targetPath - Pfad zur Export-Datei (wird in-place modifiziert)
 * @param {string} sheetName - Name des Sheets
 * @param {Array<number>} rowMapping - Mapping: rowMapping[newPos] = originalPos (0-basiert, Datenzeilen)
 */
async function postProcessRowStyles(sourcePath, targetPath, sheetName, rowMapping) {
    if (!rowMapping || rowMapping.length === 0) {
        return;
    }

    try {
        // Lade beide Dateien
        const sourceZip = new AdmZip(sourcePath);
        const targetZip = new AdmZip(targetPath);
        
        // 1. Kopiere styles.xml vom Original (enthält Theme-Colors und alle Style-Definitionen)
        const stylesXml = sourceZip.readAsText('xl/styles.xml');
        if (stylesXml) {
            targetZip.updateFile('xl/styles.xml', Buffer.from(stylesXml, 'utf8'));
        }
        
        // 2. Finde das richtige Sheet im XML
        // Erst workbook.xml lesen um Sheet-ID zu finden
        const workbookXml = sourceZip.readAsText('xl/workbook.xml');
        let sheetId = '1';
        const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]*name="${sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*r:id="rId(\\d+)"`));
        if (sheetMatch) {
            sheetId = sheetMatch[1];
        }
        
        // Finde die entsprechende sheet.xml
        const relsXml = sourceZip.readAsText('xl/_rels/workbook.xml.rels');
        let sheetPath = `xl/worksheets/sheet${sheetId}.xml`;
        const relMatch = relsXml.match(new RegExp(`Id="rId${sheetId}"[^>]*Target="([^"]+)"`));
        if (relMatch) {
            sheetPath = 'xl/' + relMatch[1].replace(/^\.\.\//, '').replace(/^\//, '');
        }
        
        // 3. Lese Original-Sheet und extrahiere Style-Indices pro Zelle
        const sourceSheetXml = sourceZip.readAsText(sheetPath);
        const targetSheetXml = targetZip.readAsText(sheetPath);
        
        if (!sourceSheetXml || !targetSheetXml) {
            return;
        }
        
        // 4. Extrahiere alle Zell-Styles aus dem Original (Zeile -> Spalte -> StyleIndex)
        const originalCellStyles = {};
        const rowRegex = /<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
        let rowMatch;
        
        while ((rowMatch = rowRegex.exec(sourceSheetXml)) !== null) {
            const rowNum = parseInt(rowMatch[1]);
            const rowContent = rowMatch[2];
            originalCellStyles[rowNum] = {};
            
            // Extrahiere Style-Index für jede Zelle in dieser Zeile
            const cellRegex = /<c[^>]*r="([A-Z]+)\d+"[^>]*>/g;
            let cellMatch;
            while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
                const cellTag = cellMatch[0];
                const col = cellMatch[1];
                const styleMatch = cellTag.match(/s="(\d+)"/);
                if (styleMatch) {
                    originalCellStyles[rowNum][col] = styleMatch[1];
                }
            }
        }
        
        // 5. Erstelle Mapping von neuen Zeilen zu Original-Styles
        // rowMapping[newDataIdx] = originalDataIdx
        // Excel-Zeile = dataIdx + 2 (Zeile 1 = Header)
        const newRowStyles = {};
        for (let newDataIdx = 0; newDataIdx < rowMapping.length; newDataIdx++) {
            const originalDataIdx = rowMapping[newDataIdx];
            const originalExcelRow = originalDataIdx + 2;
            const newExcelRow = newDataIdx + 2;
            
            if (originalCellStyles[originalExcelRow]) {
                newRowStyles[newExcelRow] = originalCellStyles[originalExcelRow];
            }
        }
        
        // 6. Update Target Sheet XML mit den umgeordneten Styles
        let updatedSheetXml = targetSheetXml;
        let updatedCells = 0;
        
        for (const [newRowNumStr, colStyles] of Object.entries(newRowStyles)) {
            const newRowNum = parseInt(newRowNumStr);
            
            // Finde diese Zeile im Target
            const rowPattern = new RegExp(`(<row[^>]*r="${newRowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
            const targetRowMatch = updatedSheetXml.match(rowPattern);
            
            if (targetRowMatch) {
                let rowContent = targetRowMatch[2];
                
                // Update Style-Index für jede Zelle
                for (const [col, styleIdx] of Object.entries(colStyles)) {
                    // Finde die Zelle in dieser Zeile
                    const cellPattern = new RegExp(`(<c[^>]*r="${col}${newRowNum}"[^>]*)(/?>)`);
                    const cellMatch = rowContent.match(cellPattern);
                    
                    if (cellMatch) {
                        let cellTag = cellMatch[1];
                        const cellEnd = cellMatch[2];
                        
                        // Entferne bestehenden s="X" Attribut
                        cellTag = cellTag.replace(/\s*s="\d+"/g, '');
                        
                        // Füge neuen Style-Index hinzu
                        cellTag = cellTag + ` s="${styleIdx}"`;
                        
                        rowContent = rowContent.replace(cellMatch[0], cellTag + cellEnd);
                        updatedCells++;
                    }
                }
                
                // Ersetze die Zeile im XML
                updatedSheetXml = updatedSheetXml.replace(rowPattern, targetRowMatch[1] + rowContent + targetRowMatch[3]);
            }
        }
        
        // 7. Speichere das aktualisierte Sheet
        targetZip.updateFile(sheetPath, Buffer.from(updatedSheetXml, 'utf8'));
        
        // 8. Schreibe die aktualisierte ZIP-Datei
        targetZip.writeZip(targetPath);

    } catch (error) {
        console.error('[ExcelJS Writer] Post-Processing Fehler:', error);
        // Fehler nicht werfen - die Datei wurde bereits gespeichert
    }
}

/**
 * Exportiert/Speichert mehrere Sheets mit ExcelJS
 * 
 * @param {string} sourcePath - Pfad zur Quelldatei
 * @param {string} targetPath - Pfad zur Zieldatei
 * @param {Array} sheets - Array von Sheet-Daten
 * @param {Object} options - Optionen (password, sourcePassword)
 * @returns {Promise<Object>} Erfolg/Fehler
 */
async function exportMultipleSheetsWithExcelJS(sourcePath, targetPath, sheets, options = {}) {
    const startTime = Date.now();
    
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(sourcePath);
        
        // Liste der ausgewählten Sheet-Namen
        const selectedSheetNames = sheets.map(s => s.sheetName);
        
        // Alle Sheets der Originaldatei
        const allSheetNames = workbook.worksheets.map(ws => ws.name);
        
        // Sheets entfernen, die nicht ausgewählt wurden (von hinten nach vorne)
        for (let i = allSheetNames.length - 1; i >= 0; i--) {
            const sheetName = allSheetNames[i];
            if (!selectedSheetNames.includes(sheetName)) {
                const sheetToDelete = workbook.getWorksheet(sheetName);
                if (sheetToDelete) {
                    workbook.removeWorksheet(sheetToDelete.id);
                }
            }
        }
        
        let sheetsProcessed = 0;
        
        // Jedes Sheet verarbeiten
        for (const sheetData of sheets) {
            const worksheet = workbook.getWorksheet(sheetData.sheetName);
            if (!worksheet) {
                console.warn('[ExcelJS Writer] Sheet "' + sheetData.sheetName + '" nicht gefunden - übersprungen');
                continue;
            }
            
            // WICHTIG: Auch bei fromFile müssen wir Fills anwenden, die ExcelJS nicht korrekt gelesen hat
            // (z.B. bei SoftMaker-Dateien werden manche Fills nicht erkannt)
            // OPTIMIERUNG: Bei großen Dateien überspringen wir das, da ExcelJS die meisten Fills korrekt lädt
            if (sheetData.fromFile) {
                let cellStyles = sheetData.cellStyles;
                
                // Bei sehr großen Dateien: Styles überspringen (ExcelJS hat sie bereits)
                const styleCount = cellStyles ? Object.keys(cellStyles).length : 0;
                if (styleCount > 50000) {
                    // Überspringe Fill-Anwendung bei sehr großen Dateien
                } else if (!cellStyles || styleCount === 0) {
                    // Keine cellStyles vom Frontend - nichts zu tun
                } else {
                    // Moderate Anzahl Styles - anwenden
                    applyMissingFills(worksheet, cellStyles);
                }
                
                // Versteckte Spalten setzen
                if (sheetData.hiddenColumns !== undefined) {
                    const hiddenSet = new Set(sheetData.hiddenColumns || []);
                    const columnCount = worksheet.columnCount || 0;

                    for (let colIdx = 0; colIdx < columnCount; colIdx++) {
                        const col = worksheet.getColumn(colIdx + 1);
                        const shouldBeHidden = hiddenSet.has(colIdx);
                        if (col.hidden !== shouldBeHidden) {
                            col.hidden = shouldBeHidden;
                        }
                    }
                }
                
                // AutoFilter setzen, falls übergeben
                if (sheetData.autoFilterRange) {
                    worksheet.autoFilter = sheetData.autoFilterRange;
                }
                
                sheetsProcessed++;
                continue;
            }
            
            // Verarbeite das Sheet
            await processSheet(worksheet, sheetData);
            sheetsProcessed++;
        }

        // Datei speichern
        await workbook.xlsx.writeFile(targetPath);
        
        // Post-Processing für Zeilen-Verschiebungen DEAKTIVIERT
        // Das Kopieren von styles.xml verursacht Probleme mit den String-Referenzen
        // TODO: Bessere Lösung finden, die nur Style-Indices mapped ohne styles.xml zu ersetzen
        /*
        for (const sheetData of sheets) {
            if (sheetData.rowMapping && sheetData.rowMapping.length > 0) {
                await postProcessRowStyles(sourcePath, targetPath, sheetData.sheetName, sheetData.rowMapping);
            }
        }
        */
        
        // Passwortschutz anwenden falls gewünscht
        // ExcelJS unterstützt keinen Passwortschutz, daher verwenden wir xlsx-populate
        if (options.password) {
            try {
                const XlsxPopulate = require('xlsx-populate');
                const pwWorkbook = await XlsxPopulate.fromFileAsync(targetPath);
                await pwWorkbook.toFileAsync(targetPath, { password: options.password });
            } catch (pwError) {
                console.error('[ExcelJS Writer] Fehler beim Passwortschutz:', pwError.message);
                // Datei wurde bereits gespeichert, nur ohne Passwort
            }
        }
        
        const totalTime = Date.now() - startTime;

        return {
            success: true,
            message: sheetsProcessed + ' Sheet(s) exportiert: ' + targetPath,
            sheetsExported: sheetsProcessed,
            stats: { totalTimeMs: totalTime }
        };
        
    } catch (error) {
        console.error('[ExcelJS Writer] Fehler:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Verarbeitet ein einzelnes Sheet
 */
async function processSheet(worksheet, sheetData) {
    const {
        headers,
        data,
        changedCells,
        visibleColumns,
        hiddenColumns,
        hiddenRows = [],
        cellStyles = {},
        cellFormulas = {},
        cellHyperlinks = {},
        richTextCells = {},
        affectedRows = [],
        fullRewrite = false,
        structuralChange = false,  // NEU: Signalisiert strukturelle Änderung (Spalte gelöscht/eingefügt)
        deletedColumnIndex,  // LEGACY: Einzelner Index der gelöschten Spalte (0-basiert)
        deletedColumnIndices = [],  // NEU: Array von gelöschten Spalten-Indizes (0-basiert)
        columnOrder = null,  // NEU: Neue Spaltenreihenfolge (Array von 0-basierten Indizes)
        rowMapping = null,  // NEU: Mapping für Zeilen-Verschiebung (Array: neue Position -> Original Excel-Zeilen-Index)
        autoFilterRange
    } = sheetData;
    
    // Kompatibilität: Legacy-Format (einzelner Index) in Array konvertieren
    let columnsToDelete = [];
    if (deletedColumnIndices && deletedColumnIndices.length > 0) {
        columnsToDelete = [...deletedColumnIndices];
    } else if (deletedColumnIndex !== undefined) {
        columnsToDelete = [deletedColumnIndex];
    }
    
    const hasRowMapping = rowMapping && Array.isArray(rowMapping) && rowMapping.length > 0;

    // MODUS 1: Nur geänderte Zellen (changedCells) - schnellster Weg
    if (changedCells && !fullRewrite && Object.keys(changedCells).length > 0) {
        for (const [cellKey, newValue] of Object.entries(changedCells)) {
            const [rowStr, colStr] = cellKey.split('-');
            const rowIdx = parseInt(rowStr);
            const colIdx = parseInt(colStr);
            
            // rowIdx ist 0-basiert für Datenzeilen, Excel-Zeile = rowIdx + 2 (Zeile 1 = Header)
            const cell = worksheet.getCell(rowIdx + 2, colIdx + 1);
            cell.value = newValue === null || newValue === undefined ? '' : newValue;
        }
        
        // RichText für geänderte Zellen anwenden
        if (richTextCells && Object.keys(richTextCells).length > 0) {
            applyRichText(worksheet, richTextCells);
        }
        
        // Formeln für geänderte Zellen anwenden
        if (cellFormulas && Object.keys(cellFormulas).length > 0) {
            applyFormulas(worksheet, cellFormulas);
        }
        
        // Hyperlinks für geänderte Zellen anwenden
        if (cellHyperlinks && Object.keys(cellHyperlinks).length > 0) {
            applyHyperlinks(worksheet, cellHyperlinks);
        }
        
        // Versteckte Spalten setzen - WICHTIG: Auch sichtbare Spalten explizit auf hidden=false setzen
        // damit Änderungen im Tool (Spalte ein-/ausblenden) korrekt gespeichert werden
        if (hiddenColumns !== undefined) {
            const hiddenSet = new Set(hiddenColumns || []);
            const columnCount = worksheet.columnCount || 0;

            for (let colIdx = 0; colIdx < columnCount; colIdx++) {
                const col = worksheet.getColumn(colIdx + 1);
                const shouldBeHidden = hiddenSet.has(colIdx);
                if (col.hidden !== shouldBeHidden) {
                    col.hidden = shouldBeHidden;
                }
            }
        }
        
        // AutoFilter setzen/erhalten, falls übergeben
        if (autoFilterRange) {
            worksheet.autoFilter = autoFilterRange;
        }
        
        return;
    }
    
    // MODUS 2: Vollständiges Schreiben (headers + data vorhanden)
    if (headers && data) {
        // Bei strukturellen Änderungen (Spalten gelöscht): ExcelJS spliceColumns verwenden
        // WICHTIG: Mehrere Spalten müssen in ABSTEIGENDER Reihenfolge gelöscht werden,
        // damit die Indizes der verbleibenden Spalten nicht durch vorherige Löschungen verfälscht werden
        if (structuralChange && columnsToDelete.length > 0) {
            // Sortiere absteigend - höchste Indizes zuerst löschen
            const sortedColumns = [...columnsToDelete].sort((a, b) => b - a);
            
            // Letzte Spalte VOR allen Löschungen merken (für CF-Cleanup am Ende)
            const lastColumnBeforeDelete = worksheet.columnCount;
            
            // Jede Spalte einzeln löschen
            for (const deletedColumnIndex of sortedColumns) {
                // deletedColumnIndex ist 0-basiert vom Frontend, Excel-Spalten sind 1-basiert
                const deletedColExcel = deletedColumnIndex + 1;
                
                // WICHTIG: Merged Cells VOR spliceColumns vorbereiten!
                const newMergeRanges = prepareMergedCellsForColumnDelete(worksheet, deletedColExcel);
                
                // spliceColumns löscht die Spalte UND verschiebt alle nachfolgenden nach links
                worksheet.spliceColumns(deletedColExcel, 1);
                
                // Merged Cells NACH spliceColumns wieder setzen
                applyMergedCellsAfterColumnDelete(worksheet, newMergeRanges, cellStyles);
                
                // CF und Tables nach JEDER Löschung anpassen (da sich Referenzen ändern)
                adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColExcel, worksheet.columnCount + 1);
                adjustTablesAfterColumnDelete(worksheet, deletedColExcel);
            }
            
            // WICHTIG: ExcelJS Bug-Workaround nach allen Löschungen!
            // Entferne überschüssige Zellen aus allen Rows
            const actualColCount = worksheet.columnCount;
            worksheet._rows.forEach((row) => {
                if (row && row._cells) {
                    for (let i = actualColCount; i < row._cells.length; i++) {
                        if (row._cells[i]) {
                            delete row._cells[i];
                        }
                    }
                    row._cells.length = actualColCount;
                }
            });
            
            // Auch _columns bereinigen
            if (worksheet._columns) {
                for (let i = actualColCount; i < worksheet._columns.length; i++) {
                    if (worksheet._columns[i]) {
                        delete worksheet._columns[i];
                    }
                }
                worksheet._columns.length = actualColCount;
            }
            
            // AutoFilter anpassen - reduziere den Bereich um die Anzahl gelöschter Spalten
            if (autoFilterRange) {
                let adjustedAutoFilter = autoFilterRange;
                for (let i = 0; i < sortedColumns.length; i++) {
                    adjustedAutoFilter = reduceAutoFilterRange(adjustedAutoFilter);
                }
                worksheet.autoFilter = adjustedAutoFilter;
            }
            
            // Versteckte Spalten setzen
            if (hiddenColumns !== undefined) {
                const hiddenSet = new Set(hiddenColumns || []);
                const columnCount = worksheet.columnCount || 0;
                for (let colIdx = 0; colIdx < columnCount; colIdx++) {
                    const col = worksheet.getColumn(colIdx + 1);
                    col.hidden = hiddenSet.has(colIdx);
                }
            }
            
            // WICHTIG: Nach spliceColumns sind alle Styles bereits korrekt verschoben!
            // spliceColumns verschiebt alle Zelleigenschaften (Werte, Styles, Fills, Fonts, etc.)
            // 
            // Die cellStyles vom Frontend sind BEREITS angepasst!
            // Das Frontend hat die Indices bereits korrigiert (src/index.html, Zeile ~7467)
            // Wir müssen sie hier NICHT nochmal anpassen!
            // 
            // ExcelJS erkennt aber nicht alle Fills aus dem Original (Indexed/Theme Colors).
            // Daher wenden wir die cellStyles direkt an - mit den korrekten Indices vom Frontend.
            
            if (cellStyles && Object.keys(cellStyles).length > 0) {
                applyMissingFills(worksheet, cellStyles, true);
            }
            
            // Nach spliceColumns sind alle Daten bereits korrekt verschoben
            // Die Frontend-Daten werden NICHT geschrieben, nur fehlende Styles ergänzt
            return;
        }
        
        // WICHTIG: Spaltenbreiten VOR dem Schreiben sichern
        // Ermittle die tatsächliche Spaltenanzahl im Worksheet (nicht nur headers.length)
        const columnWidths = {};
        const columnStyles = {};
        const actualColumnCount = Math.max(headers.length, worksheet.columnCount || 0);
        
        // Default-Spaltenbreite aus dem Worksheet (falls vorhanden)
        // Excel-Standard ist 8.43 Zeichen (etwa 64 Pixel)
        const defaultColWidth = worksheet.properties?.defaultColWidth || 8.43;
        
        for (let colIdx = 1; colIdx <= actualColumnCount; colIdx++) {
            const col = worksheet.getColumn(colIdx);
            // Breite immer sichern - mit Fallback auf Default
            if (col.width !== undefined && col.width !== null) {
                columnWidths[colIdx] = col.width;
            } else {
                // Immer eine Breite setzen - verhindert Verlust bei beschädigten Dateien
                columnWidths[colIdx] = defaultColWidth;
            }
            // Auch die column-level properties sichern (outlineLevel, hidden, etc.)
            if (col.hidden !== undefined) {
                columnStyles[colIdx] = columnStyles[colIdx] || {};
                columnStyles[colIdx].hidden = col.hidden;
            }
            if (col.outlineLevel !== undefined) {
                columnStyles[colIdx] = columnStyles[colIdx] || {};
                columnStyles[colIdx].outlineLevel = col.outlineLevel;
            }
        }
        

        
        // Zeilenhöhen VOR dem Schreiben sichern
        const rowHeights = {};
        const actualRowCount = Math.max(data.length + 1, worksheet.rowCount || 0);
        
        for (let rowIdx = 1; rowIdx <= actualRowCount; rowIdx++) {
            const row = worksheet.getRow(rowIdx);
            // Höhe immer sichern (auch customHeight prüfen)
            if (row.height !== undefined && row.height !== null) {
                rowHeights[rowIdx] = row.height;
            }
        }
        

        
        // Bei Zeilen-Verschiebung: ExcelJS-Zeilen physisch umordnen (mit allen Styles!)
        // rowMapping[newPos] = originalPos (0-basiert für Datenzeilen)
        // Excel-Zeilen: Header=1, Datenzeilen ab 2
        if (hasRowMapping) {
            // Neues _rows Array erstellen mit umgeordneten Zeilen
            // _rows ist 0-basiert, Excel-Zeilen sind 1-basiert
            // rowMapping ist für Daten (0-basiert), Excel-Datenzeilen starten bei Index 1 (Zeile 2)
            const headerRow = worksheet._rows[0]; // Header-Zeile bleibt
            const newRows = [headerRow];
            
            // Für jede neue Position: Die originale Excel-Zeile holen
            for (let newDataIdx = 0; newDataIdx < rowMapping.length; newDataIdx++) {
                const originalDataIdx = rowMapping[newDataIdx];
                // Excel-Zeile: originalDataIdx + 2 (Header ist 1, erste Datenzeile ist 2)
                // _rows Index: originalDataIdx + 1 (0-basiert)
                const originalRowsIdx = originalDataIdx + 1;
                const row = worksheet._rows[originalRowsIdx];
                
                if (row) {
                    // Zeilen-Nummer aktualisieren (1-basiert)
                    row._number = newDataIdx + 2;
                    newRows.push(row);
                } else {
                    newRows.push(undefined);
                }
            }
            
            // Worksheet _rows ersetzen
            worksheet._rows = newRows;
        }
        
        // Header aktualisieren (Zeile 1)
        headers.forEach((header, colIndex) => {
            const cell = worksheet.getCell(1, colIndex + 1);
            cell.value = header;
        });
        
        // Daten aktualisieren (ab Zeile 2) - mit Fortschritts-Log für große Dateien
        const totalRows = data.length;
        const logInterval = Math.max(500, Math.floor(totalRows / 10));
        
        // WICHTIG: Bei rowMapping wurden die Zeilen bereits physisch umgeordnet (mit Styles!)
        // Wir müssen nur noch geänderte Werte schreiben, nicht alle Zellen überschreiben
        if (hasRowMapping) {
            
            // Nur editierte Zellen aktualisieren (aus editedCells)
            const editedCellsObj = sheetData.editedCells || {};
            let editedCount = 0;
            
            for (const [key, newValue] of Object.entries(editedCellsObj)) {
                const [rowStr, colStr] = key.split(',');
                const dataRowIdx = parseInt(rowStr);
                const colIdx = parseInt(colStr);
                
                // Finde die neue Position dieser Zeile nach dem Mapping
                // rowMapping[newPos] = originalPos
                let newRowIdx = -1;
                for (let i = 0; i < rowMapping.length; i++) {
                    if (rowMapping[i] === dataRowIdx) {
                        newRowIdx = i;
                        break;
                    }
                }
                
                if (newRowIdx >= 0) {
                    const excelRow = newRowIdx + 2; // +2 weil Header in Zeile 1
                    const excelCol = colIdx + 1;
                    const cell = worksheet.getCell(excelRow, excelCol);
                    
                    // Style sichern
                    const savedStyle = JSON.parse(JSON.stringify(cell.style || {}));
                    
                    // Wert setzen
                    cell.value = newValue === null || newValue === undefined ? '' : newValue;
                    
                    // Style wiederherstellen
                    if (Object.keys(savedStyle).length > 0) {
                        cell.style = savedStyle;
                    }
                    
                    editedCount++;
                }
            }
        } else {
            // Normales Schreiben ohne rowMapping
            data.forEach((row, rowIndex) => {
                row.forEach((value, colIndex) => {
                    const cell = worksheet.getCell(rowIndex + 2, colIndex + 1);
                    cell.value = value === null || value === undefined ? '' : value;
                });
                
            });
        }
        
        // WICHTIG: Überschüssige Zeilen leeren (wenn Zeilen gefiltert wurden)
        // newRowCount = data.length + 1 (Header-Zeile)
        const newRowCount = data.length + 1;
        const originalRowCount = worksheet.rowCount || 0;
        
        if (originalRowCount > newRowCount) {
            // Alle überschüssigen Zeilen leeren
            const colCount = Math.max(headers.length, worksheet.columnCount || 0);
            for (let rowIdx = newRowCount + 1; rowIdx <= originalRowCount; rowIdx++) {
                for (let colIdx = 1; colIdx <= colCount; colIdx++) {
                    const cell = worksheet.getCell(rowIdx, colIdx);
                    cell.value = null;
                    cell.style = {};
                }
            }
            
            // Zeilen aus dem internen Array entfernen
            if (worksheet._rows) {
                for (let i = newRowCount; i < worksheet._rows.length; i++) {
                    if (worksheet._rows[i]) {
                        delete worksheet._rows[i];
                    }
                }
                worksheet._rows.length = newRowCount;
            }
        }
        
        // WICHTIG: Überschüssige Spalten leeren (wenn Spalten gelöscht wurden)
        const originalColumnCount = worksheet.columnCount || 0;
        if (originalColumnCount > headers.length) {
            // Alle Zeilen durchgehen und überschüssige Zellen leeren
            const rowCount = worksheet.rowCount || (data.length + 1);
            for (let rowIdx = 1; rowIdx <= rowCount; rowIdx++) {
                for (let colIdx = headers.length + 1; colIdx <= originalColumnCount; colIdx++) {
                    const cell = worksheet.getCell(rowIdx, colIdx);
                    cell.value = null;
                    cell.style = {};
                }
            }
        }
        
        // Spaltenbreiten wiederherstellen (nur für vorhandene Spalten)
        for (const [colIdx, width] of Object.entries(columnWidths)) {
            const colNum = parseInt(colIdx);
            if (colNum <= headers.length) {
                const col = worksheet.getColumn(colNum);
                col.width = width;
            }
        }
        
        // Column-Styles wiederherstellen (hidden, outlineLevel) - nur für vorhandene Spalten
        for (const [colIdx, styles] of Object.entries(columnStyles)) {
            const colNum = parseInt(colIdx);
            if (colNum <= headers.length) {
                const col = worksheet.getColumn(colNum);
                if (styles.hidden !== undefined) col.hidden = styles.hidden;
                if (styles.outlineLevel !== undefined) col.outlineLevel = styles.outlineLevel;
            }
        }
        
        // Zeilenhöhen wiederherstellen
        for (const [rowIdx, height] of Object.entries(rowHeights)) {
            worksheet.getRow(parseInt(rowIdx)).height = height;
        }
        
        // Bei fullRewrite: Styles anwenden
        // ABER: Bei Spalten-Löschung wurden die Excel-Styles bereits via spliceColumns verschoben
        // ABER: Bei Zeilen-Verschiebung wurden die Excel-Zeilen bereits umgeordnet (mit Styles)
        // Bei anderen fullRewrite-Fällen: Frontend-Styles anwenden
        const usedSpliceColumns = structuralChange && columnsToDelete.length > 0;
        const hasRowMoves = affectedRows && affectedRows.length > 0;
        
        if (hasRowMapping) {
            // Zeilen-Verschiebung: ExcelJS-Zeilen wurden bereits umgeordnet (mit allen Styles!)
            // Keine Frontend-Styles nötig, da die Original-Excel-Styles erhalten bleiben
        } else if (fullRewrite && !usedSpliceColumns) {
            const styleCount = Object.keys(cellStyles || {}).length;

            // Frontend-Styles anwenden
            if (styleCount > 0) {
                applyStyles(worksheet, cellStyles);
            }
        } else if (usedSpliceColumns) {
            // Excel-Styles wurden bereits verschoben - Frontend-Styles überspringen
        } else if (affectedRows && affectedRows.length > 0) {
            // Nur betroffene Zeilen aktualisieren
            const affectedSet = new Set(affectedRows);
            const filteredStyles = {};
            for (const [key, style] of Object.entries(cellStyles)) {
                const rowIdx = parseInt(key.split('-')[0]);
                if (rowIdx === 0 || affectedSet.has(rowIdx)) {
                    filteredStyles[key] = style;
                }
            }
            if (Object.keys(filteredStyles).length > 0) {
                applyStyles(worksheet, filteredStyles);
            }
        }
        
        // Versteckte Spalten setzen - WICHTIG: Spaltenbreite aus gesichertem Object nehmen!
        if (visibleColumns && visibleColumns.length > 0 && visibleColumns.length < headers.length) {
            const visibleSet = new Set(visibleColumns);
            for (let colIdx = 0; colIdx < headers.length; colIdx++) {
                const col = worksheet.getColumn(colIdx + 1);
                col.hidden = !visibleSet.has(colIdx);
                // Breite aus gesichertem Object wiederherstellen
                if (columnWidths[colIdx + 1]) {
                    col.width = columnWidths[colIdx + 1];
                }
            }
        }
        
        // Explizit versteckte Spalten
        if (hiddenColumns && hiddenColumns.length > 0) {
            hiddenColumns.forEach(colIdx => {
                const col = worksheet.getColumn(colIdx + 1);
                col.hidden = true;
                // Breite aus gesichertem Object wiederherstellen
                if (columnWidths[colIdx + 1]) {
                    col.width = columnWidths[colIdx + 1];
                }
            });
        }
        
        // Versteckte Zeilen - WICHTIG: Zeilenhöhe aus gesichertem Object nehmen!
        if (hiddenRows && hiddenRows.length > 0) {
            const hiddenRowSet = new Set(hiddenRows);
            for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
                const row = worksheet.getRow(rowIdx + 2);
                row.hidden = hiddenRowSet.has(rowIdx);
                // Höhe aus gesichertem Object wiederherstellen
                if (rowHeights[rowIdx + 2]) {
                    row.height = rowHeights[rowIdx + 2];
                }
            }
        }
    }
    
    // RichText anwenden
    if (richTextCells && Object.keys(richTextCells).length > 0) {
        applyRichText(worksheet, richTextCells);
    }
    
    // Formeln anwenden
    if (cellFormulas && Object.keys(cellFormulas).length > 0) {
        applyFormulas(worksheet, cellFormulas);
    }
    
    // Hyperlinks anwenden
    if (cellHyperlinks && Object.keys(cellHyperlinks).length > 0) {
        applyHyperlinks(worksheet, cellHyperlinks);
    }
    
    // AutoFilter wiederherstellen
    if (autoFilterRange) {
        worksheet.autoFilter = autoFilterRange;
    }
}

/**
 * Wendet Styles auf Zellen an
 */
function applyStyles(worksheet, cellStyles) {
    const entries = Object.entries(cellStyles);
    const totalCount = entries.length;

    let processedCount = 0;
    const batchSize = 2000;
    
    for (const [styleKey, style] of entries) {
        try {
            const [rowIdx, colIdx] = styleKey.split('-').map(Number);
            // Frontend sendet 1-basierte rowIdx (originalIndex + 1)
            // rowIdx 0 = Header (Excel-Zeile 1), rowIdx 1 = erste Datenzeile (Excel-Zeile 2)
            const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
            
            // Font-Styles - bestehende Eigenschaften beibehalten
            const existingFont = cell.font || {};
            const font = { ...existingFont };
            
            if (style.bold !== undefined) font.bold = style.bold;
            if (style.italic !== undefined) font.italic = style.italic;
            if (style.underline !== undefined) font.underline = style.underline;
            if (style.strikethrough !== undefined) font.strike = style.strikethrough;
            if (style.fontSize) font.size = style.fontSize;
            if (style.fontColor) {
                const hex = style.fontColor.replace('#', '');
                font.color = { argb: 'FF' + hex };
            }
            
            if (Object.keys(font).length > 0) {
                cell.font = font;
            }
            
            // Fill/Background
            if (style.fill) {
                const hex = style.fill.replace('#', '');
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF' + hex }
                };
            }
            
            // Alignment
            if (style.textAlign || style.verticalAlign) {
                const alignment = cell.alignment || {};
                if (style.textAlign) alignment.horizontal = style.textAlign;
                if (style.verticalAlign) alignment.vertical = style.verticalAlign;
                cell.alignment = alignment;
            }
        } catch (error) {
            // Silent fail für einzelne Zellen
        }

        processedCount++;
    }
}

/**
 * Wendet RichText auf Zellen an
 */
function applyRichText(worksheet, richTextCells) {
    for (const [styleKey, richText] of Object.entries(richTextCells)) {
        try {
            if (!Array.isArray(richText) || richText.length === 0) continue;
            
            const [rowIdx, colIdx] = styleKey.split('-').map(Number);
            if (isNaN(rowIdx) || isNaN(colIdx)) continue;
            
            const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
            
            const richTextValue = richText.map(part => {
                if (!part || typeof part !== 'object') return null;
                
                const text = String(part.text || '');
                if (text.length === 0) return null;
                
                const font = {};
                if (part.styles) {
                    if (part.styles.bold) font.bold = true;
                    if (part.styles.italic) font.italic = true;
                    if (part.styles.underline) font.underline = true;
                    if (part.styles.strikethrough) font.strike = true;
                    if (part.styles.color) {
                        const hex = part.styles.color.replace('#', '');
                        font.color = { argb: 'FF' + hex };
                    }
                    if (part.styles.fontSize) font.size = part.styles.fontSize;
                    if (part.styles.fontName) font.name = part.styles.fontName;
                }
                
                return { text, font };
            }).filter(Boolean);
            
            if (richTextValue.length > 0) {
                cell.value = { richText: richTextValue };
            }
        } catch (error) {
            console.warn('[ExcelJS Writer] RichText für ' + styleKey + ' übersprungen:', error.message);
        }
    }
}

/**
 * Wendet Formeln auf Zellen an
 */
function applyFormulas(worksheet, cellFormulas) {
    for (const [styleKey, formula] of Object.entries(cellFormulas)) {
        try {
            const [rowIdx, colIdx] = styleKey.split('-').map(Number);
            if (isNaN(rowIdx) || isNaN(colIdx)) continue;
            
            const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
            const formulaStr = formula.startsWith('=') ? formula.substring(1) : formula;
            cell.value = { formula: formulaStr };
        } catch (error) {
            console.warn('[ExcelJS Writer] Formel für ' + styleKey + ' fehlgeschlagen:', error.message);
        }
    }
}

/**
 * Wendet Hyperlinks auf Zellen an
 */
function applyHyperlinks(worksheet, cellHyperlinks) {
    for (const [styleKey, hyperlink] of Object.entries(cellHyperlinks)) {
        try {
            const [rowIdx, colIdx] = styleKey.split('-').map(Number);
            if (isNaN(rowIdx) || isNaN(colIdx)) continue;
            
            const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
            const currentValue = cell.value;
            const displayText = typeof currentValue === 'string' ? currentValue : hyperlink;
            
            cell.value = {
                text: displayText,
                hyperlink: hyperlink
            };
        } catch (error) {
            console.warn('[ExcelJS Writer] Hyperlink für ' + styleKey + ' fehlgeschlagen:', error.message);
        }
    }
}

/**
 * Legacy-Funktion für Einzelsheet-Export (Rückwärtskompatibilität)
 */
async function exportSheetWithExcelJS(sourcePath, targetPath, sheetData) {
    return exportMultipleSheetsWithExcelJS(sourcePath, targetPath, [sheetData]);
}

/**
 * Wendet fehlende Fills auf Zellen an
 * ExcelJS liest manche Fills nicht korrekt (z.B. bei SoftMaker-Dateien)
 * Diese Funktion prüft ob die Zelle bereits eine Fill hat und setzt sie nur wenn nötig
 * @param {boolean} force - Wenn true, wird das Performance-Limit ignoriert
 */
function applyMissingFills(worksheet, cellStyles, force = false) {
    const entries = Object.entries(cellStyles);
    const totalCount = entries.length;
    
    // Skip wenn zu viele Cells - Performance-Optimierung
    // Bei großen Dateien (>50k Zellen) nur Fills anwenden wenn explizit angefordert
    if (totalCount > 50000 && !force) {
        return;
    }

    let appliedCount = 0;
    let skippedCount = 0;
    let noFillInStyleCount = 0;
    let processedCount = 0;
    const batchSize = 10000;
    
    for (const [styleKey, style] of entries) {
        if (!style.fill) {
            noFillInStyleCount++;
            continue;
        }
        
        try {
            const [rowIdx, colIdx] = styleKey.split('-').map(Number);
            if (isNaN(rowIdx) || isNaN(colIdx)) continue;
            
            // rowIdx 0 = Header (Excel-Zeile 1), rowIdx 1 = erste Datenzeile (Excel-Zeile 2)
            const cell = worksheet.getCell(rowIdx + 1, colIdx + 1);
            
            // Prüfe ob die Zelle bereits eine Fill hat
            const existingFill = cell.fill;
            const hasFill = existingFill && 
                            existingFill.type === 'pattern' && 
                            existingFill.pattern === 'solid' &&
                            existingFill.fgColor?.argb;
            
            // Nur setzen wenn keine Fill vorhanden
            if (!hasFill) {
                const hex = style.fill.replace('#', '');
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF' + hex }
                };
                appliedCount++;
            } else {
                skippedCount++;
            }
        } catch (error) {
            // Silent fail für einzelne Zellen
        }

        processedCount++;
    }
}

module.exports = {
    exportSheetWithExcelJS,
    exportMultipleSheetsWithExcelJS
};
