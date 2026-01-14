// Test: Simuliere den echten Flow aus exceljs-writer.js
const ExcelJS = require('exceljs');
const fs = require('fs');

// Kopie der Funktionen aus exceljs-writer.js
function colLetterToNumber(letters) {
    let num = 0;
    for (let i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

function colNumberToLetter(num) {
    let result = '';
    while (num > 0) {
        const remainder = (num - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        num = Math.floor((num - 1) / 26);
    }
    return result;
}

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

function removeColumnFromRef(ref, colNumber) {
    const targetCol = colNumberToLetter(colNumber);
    const prevCol = colNumber > 1 ? colNumberToLetter(colNumber - 1) : null;
    const ranges = ref.split(' ');
    
    const processedRanges = [];
    
    for (const range of ranges) {
        const parts = range.split(':');
        
        if (parts.length === 2) {
            const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
            const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
            
            if (startMatch && endMatch) {
                const startCol = startMatch[1];
                const startRow = startMatch[2];
                const endCol = endMatch[1];
                const endRow = endMatch[2];
                
                if (startCol === targetCol && endCol === targetCol) {
                    continue;
                } else if (endCol === targetCol && prevCol) {
                    processedRanges.push(startCol + startRow + ':' + prevCol + endRow);
                } else if (startCol === targetCol) {
                    const nextCol = colNumberToLetter(colNumber + 1);
                    processedRanges.push(nextCol + startRow + ':' + endCol + endRow);
                } else {
                    processedRanges.push(range);
                }
            } else {
                processedRanges.push(range);
            }
        } else {
            const match = range.match(/^([A-Z]+)/);
            if (match && match[1] !== targetCol) {
                processedRanges.push(range);
            }
        }
    }
    
    return processedRanges.join(' ') || ref;
}

function adjustCellReference(cellRef, deletedColNumber) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return cellRef;
    
    const colLetter = match[1];
    const rowNumber = match[2];
    const colNumber = colLetterToNumber(colLetter);
    
    if (colNumber > deletedColNumber) {
        return colNumberToLetter(colNumber - 1) + rowNumber;
    } else if (colNumber === deletedColNumber) {
        return colNumberToLetter(colNumber) + rowNumber;
    }
    return cellRef;
}

function adjustRangeReference(rangeRef, deletedColNumber) {
    const multiRanges = rangeRef.split(' ');
    
    const adjustedRanges = multiRanges.map(singleRange => {
        const parts = singleRange.split(':');
        if (parts.length === 2) {
            return adjustCellReference(parts[0], deletedColNumber) + ':' + adjustCellReference(parts[1], deletedColNumber);
        }
        return adjustCellReference(singleRange, deletedColNumber);
    });
    
    return adjustedRanges.join(' ');
}

function adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColNumber, lastColumnBeforeDelete) {
    const cf = worksheet.conditionalFormattings;
    if (!cf || !Array.isArray(cf) || cf.length === 0) {
        console.log('[ExcelJS Writer] Keine bedingten Formatierungen zu aktualisieren');
        return;
    }
    
    const emptyLastCol = lastColumnBeforeDelete;
    const emptyLastColLetter = colNumberToLetter(emptyLastCol);
    
    console.log('[ExcelJS Writer] CF-Anpassung: ' + cf.length + ' Regeln, letzte Spalte (leer): ' + emptyLastColLetter);
    
    let adjustedCount = 0;
    let removedCount = 0;
    const toRemove = [];
    
    cf.forEach((cfEntry, idx) => {
        if (!cfEntry.ref) return;
        
        const oldRef = cfEntry.ref;
        
        const adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
        
        if (refOnlyReferencesColumn(adjustedRef, emptyLastCol)) {
            console.log('[ExcelJS Writer] ENTFERNE CF (zeigt auf leere Spalte ' + emptyLastColLetter + '): ' + oldRef + ' → ' + adjustedRef);
            toRemove.push(idx);
            removedCount++;
            return;
        }
        
        const cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
        
        // HIER IST DER AKTUELLE CODE:
        if (cleanedRef !== oldRef) {
            cfEntry.ref = cleanedRef;
            adjustedCount++;
            console.log('[ExcelJS Writer] CF: ' + oldRef + ' → ' + cleanedRef);
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
    
    for (let i = toRemove.length - 1; i >= 0; i--) {
        cf.splice(toRemove[i], 1);
    }
    
    console.log('[ExcelJS Writer] CF fertig: ' + adjustedCount + ' angepasst, ' + removedCount + ' entfernt');
}

async function testRealFlow() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Test');
    
    // Füge Daten hinzu
    for (let i = 1; i <= 5; i++) {
        ws.getCell(1, i).value = colNumberToLetter(i);
    }
    
    // Füge bedingte Formatierung hinzu
    ws.addConditionalFormatting({
        ref: 'D1:E10',
        rules: [{ type: 'expression', formulae: ['$D1>0'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }]
    });
    
    // Speichere und lade
    const tempPath = '/tmp/test-cf-flow.xlsx';
    await wb.xlsx.writeFile(tempPath);
    
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(tempPath);
    const ws2 = wb2.getWorksheet('Test');
    
    console.log('=== VOR ALLES ===');
    console.log('Zellen:', ws2.getCell('A1').value, ws2.getCell('B1').value, ws2.getCell('C1').value, ws2.getCell('D1').value, ws2.getCell('E1').value);
    console.log('CF ref:', ws2.conditionalFormattings[0]?.ref);
    console.log('columnCount:', ws2.columnCount);
    
    const deletedColExcel = 3; // Spalte C
    const lastColumnBeforeDelete = ws2.columnCount;
    
    console.log('\n=== LÖSCHE SPALTE C (3), lastColumnBeforeDelete=' + lastColumnBeforeDelete + ' ===');
    
    // Rufe spliceColumns auf
    ws2.spliceColumns(deletedColExcel, 1);
    
    console.log('\n=== NACH spliceColumns, VOR adjustConditionalFormattingsAfterColumnDelete ===');
    console.log('Zellen:', ws2.getCell('A1').value, ws2.getCell('B1').value, ws2.getCell('C1').value, ws2.getCell('D1').value, ws2.getCell('E1').value);
    console.log('CF ref:', ws2.conditionalFormattings[0]?.ref);
    
    // Rufe die Anpassungsfunktion auf
    adjustConditionalFormattingsAfterColumnDelete(ws2, deletedColExcel, lastColumnBeforeDelete);
    
    console.log('\n=== NACH adjustConditionalFormattingsAfterColumnDelete ===');
    console.log('CF ref:', ws2.conditionalFormattings[0]?.ref);
    console.log('Erwarteter Wert: C1:D10');
    console.log('Formeln:', ws2.conditionalFormattings[0]?.rules[0]?.formulae);
    console.log('Erwartete Formeln: $C1>0');
    
    // Simuliere das zweite spliceColumns für die leere letzte Spalte
    const newLastColumn = lastColumnBeforeDelete;
    if (ws2.columnCount >= newLastColumn) {
        console.log('\n=== ZWEITES spliceColumns auf Spalte ' + newLastColumn + ' (' + colNumberToLetter(newLastColumn) + ') ===');
        ws2.spliceColumns(newLastColumn, 1);
    }
    
    console.log('\n=== NACH ZWEITEM spliceColumns ===');
    console.log('CF ref:', ws2.conditionalFormattings[0]?.ref);
    console.log('Erwarteter Wert: C1:D10 (sollte gleich bleiben)');
    
    // Speichere und prüfe Ergebnis
    await wb2.xlsx.writeFile(tempPath);
    
    const wb3 = new ExcelJS.Workbook();
    await wb3.xlsx.readFile(tempPath);
    const ws3 = wb3.getWorksheet('Test');
    
    console.log('\n=== NACH SPEICHERN UND NEU LADEN ===');
    console.log('CF ref:', ws3.conditionalFormattings[0]?.ref);
    console.log('Erwarteter Wert: C1:D10');
    
    fs.unlinkSync(tempPath);
}

testRealFlow().catch(console.error);
