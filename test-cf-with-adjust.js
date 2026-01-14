// Test: Prüfe adjustConditionalFormattingsAfterColumnDelete mit echten Daten
const ExcelJS = require('exceljs');
const fs = require('fs');

// Kopiere alle relevanten Funktionen aus exceljs-writer.js
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
    console.log('[ExcelJS Writer] deletedColNumber: ' + deletedColNumber + ' (' + colNumberToLetter(deletedColNumber) + ')');
    
    let adjustedCount = 0;
    let removedCount = 0;
    const toRemove = [];
    
    cf.forEach((cfEntry, idx) => {
        if (!cfEntry.ref) return;
        
        const oldRef = cfEntry.ref;
        
        const adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
        
        if (refOnlyReferencesColumn(adjustedRef, emptyLastCol)) {
            if (removedCount < 3) {
                console.log('[ExcelJS Writer] ENTFERNE CF (zeigt auf leere Spalte ' + emptyLastColLetter + '): ' + oldRef + ' → ' + adjustedRef);
            }
            toRemove.push(idx);
            removedCount++;
            return;
        }
        
        const cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
        
        if (cleanedRef !== oldRef) {
            cfEntry.ref = cleanedRef;
            adjustedCount++;
            if (adjustedCount <= 5) {
                console.log('[ExcelJS Writer] CF: ' + oldRef + ' → ' + cleanedRef);
            }
        }
        
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

const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
const testOutputPath = '/tmp/test-cf-with-adjust.xlsx';

async function testWithAdjust() {
    console.log('=== Test: CF mit adjustConditionalFormattingsAfterColumnDelete ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    console.log('Sheet:', worksheet.name);
    console.log('Spalten VOR Löschen:', worksheet.columnCount);
    
    const cf = worksheet.conditionalFormattings;
    console.log('CF-Regeln:', cf.length);
    
    console.log('\n=== CF VOR Löschen (erste 3) ===');
    cf.slice(0, 3).forEach((cfEntry, idx) => {
        console.log('CF ' + (idx + 1) + ' ref:', cfEntry.ref);
    });
    
    // Lösche Spalte E (5)
    const deletedColExcel = 5;
    const lastColumnBeforeDelete = worksheet.columnCount;
    
    console.log('\n=== LÖSCHE Spalte ' + deletedColExcel + ' (' + colNumberToLetter(deletedColExcel) + ') ===');
    worksheet.spliceColumns(deletedColExcel, 1);
    
    console.log('\n=== Rufe adjustConditionalFormattingsAfterColumnDelete auf ===');
    adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColExcel, lastColumnBeforeDelete);
    
    console.log('\n=== CF NACH Anpassung (erste 3) ===');
    cf.slice(0, 3).forEach((cfEntry, idx) => {
        console.log('CF ' + (idx + 1) + ' ref:', cfEntry.ref);
    });
    
    // Erwartete Werte
    console.log('\n=== Validierung ===');
    console.log('Original: AN2135:AY2404');
    console.log('Nach Anpassung:', cf[0]?.ref);
    console.log('Erwartet: AM2135:AX2404');
    console.log('KORREKT:', cf[0]?.ref === 'AM2135:AX2404' ? 'JA' : 'NEIN');
    
    // Speichere und prüfe
    await workbook.xlsx.writeFile(testOutputPath);
    
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(testOutputPath);
    const worksheet2 = workbook2.getWorksheet(1);
    
    console.log('\n=== Nach Speichern und Neu-Laden ===');
    console.log('CF ref:', worksheet2.conditionalFormattings[0]?.ref);
    console.log('KORREKT:', worksheet2.conditionalFormattings[0]?.ref === 'AM2135:AX2404' ? 'JA' : 'NEIN');
    
    fs.unlinkSync(testOutputPath);
}

testWithAdjust().catch(console.error);
