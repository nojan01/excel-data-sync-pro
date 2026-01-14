// Test: Prüfe ob das zweite spliceColumns die CF wieder kaputt macht
const ExcelJS = require('exceljs');
const fs = require('fs');

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

function adjustCellReference(cellRef, deletedColNumber) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return cellRef;
    
    const colLetter = match[1];
    const rowNumber = match[2];
    const colNumber = colLetterToNumber(colLetter);
    
    if (colNumber > deletedColNumber) {
        return colNumberToLetter(colNumber - 1) + rowNumber;
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

function adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColNumber, lastColumnBeforeDelete) {
    const cf = worksheet.conditionalFormattings;
    if (!cf || !Array.isArray(cf) || cf.length === 0) {
        console.log('  [CF] Keine CF zu aktualisieren');
        return;
    }
    
    const emptyLastCol = lastColumnBeforeDelete;
    const emptyLastColLetter = colNumberToLetter(emptyLastCol);
    
    console.log('  [CF] ' + cf.length + ' Regeln, deletedCol: ' + deletedColNumber + ', emptyLastCol: ' + emptyLastCol);
    
    let adjustedCount = 0;
    let removedCount = 0;
    const toRemove = [];
    
    cf.forEach((cfEntry, idx) => {
        if (!cfEntry.ref) return;
        
        const oldRef = cfEntry.ref;
        const adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
        
        if (refOnlyReferencesColumn(adjustedRef, emptyLastCol)) {
            toRemove.push(idx);
            removedCount++;
            return;
        }
        
        const cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
        
        if (cleanedRef !== oldRef) {
            cfEntry.ref = cleanedRef;
            adjustedCount++;
        }
    });
    
    for (let i = toRemove.length - 1; i >= 0; i--) {
        cf.splice(toRemove[i], 1);
    }
    
    console.log('  [CF] Fertig: ' + adjustedCount + ' angepasst, ' + removedCount + ' entfernt');
}

const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';

async function testSecondSplice() {
    console.log('=== Test: Zweites spliceColumns-Problem ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    console.log('Sheet:', worksheet.name);
    console.log('Spalten VOR Löschen:', worksheet.columnCount, '(' + colNumberToLetter(worksheet.columnCount) + ')');
    
    const cf = worksheet.conditionalFormattings;
    console.log('\n1. CF VOR allem:', cf[0]?.ref);
    
    // Lösche Spalte E (5)
    const deletedColExcel = 5;
    const lastColumnBeforeDelete = worksheet.columnCount; // 61 = BI
    
    console.log('\n2. spliceColumns(' + deletedColExcel + ', 1) - lösche Spalte ' + colNumberToLetter(deletedColExcel));
    worksheet.spliceColumns(deletedColExcel, 1);
    console.log('   Spalten NACH erstem splice:', worksheet.columnCount, '(' + colNumberToLetter(worksheet.columnCount) + ')');
    console.log('   CF nach erstem splice:', cf[0]?.ref, '(ExcelJS ändert CF nicht)');
    
    console.log('\n3. adjustConditionalFormattingsAfterColumnDelete()');
    adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColExcel, lastColumnBeforeDelete);
    console.log('   CF nach adjust:', cf[0]?.ref);
    
    // JETZT das zweite spliceColumns wie im echten Code
    const newLastColumn = lastColumnBeforeDelete; // 61 = BI
    console.log('\n4. Zweites spliceColumns(' + newLastColumn + ', 1) - lösche leere letzte Spalte ' + colNumberToLetter(newLastColumn));
    console.log('   worksheet.columnCount VOR zweitem splice:', worksheet.columnCount);
    
    if (worksheet.columnCount >= newLastColumn) {
        worksheet.spliceColumns(newLastColumn, 1);
        console.log('   Zweites spliceColumns ausgeführt');
    } else {
        console.log('   Zweites spliceColumns ÜBERSPRUNGEN (columnCount < newLastColumn)');
    }
    
    console.log('   Spalten NACH zweitem splice:', worksheet.columnCount, '(' + colNumberToLetter(worksheet.columnCount) + ')');
    console.log('   CF nach zweitem splice:', cf[0]?.ref);
    
    // Prüfe ob die CF noch korrekt ist
    console.log('\n=== Validierung ===');
    console.log('Original CF: AN2135:AY2404');
    console.log('Finale CF:  ', cf[0]?.ref);
    console.log('Erwartet:    AM2135:AX2404');
    console.log('KORREKT:', cf[0]?.ref === 'AM2135:AX2404' ? 'JA' : 'NEIN - PROBLEM!');
}

testSecondSplice().catch(console.error);
