const ExcelJS = require('exceljs');

// Importiere die Anpassungsfunktionen direkt
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

function adjustMergedCellsAfterColumnDelete(worksheet, deletedColNumber) {
    const merges = worksheet.model.merges;
    if (!merges || !Array.isArray(merges) || merges.length === 0) {
        console.log('Keine Merged Cells zu aktualisieren');
        return;
    }
    
    console.log('Passe ' + merges.length + ' Merged Cells an (Spalte ' + deletedColNumber + ' gelöscht)');
    
    const oldMerges = [...merges];
    
    // Unmerge alle
    oldMerges.forEach(mergeRange => {
        try {
            worksheet.unMergeCells(mergeRange);
        } catch (e) {}
    });
    
    let adjustedCount = 0;
    let removedCount = 0;
    
    oldMerges.forEach(mergeRange => {
        const parts = mergeRange.split(':');
        if (parts.length !== 2) return;
        
        const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
        const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
        if (!startMatch || !endMatch) return;
        
        const startCol = colLetterToNumber(startMatch[1]);
        const startRow = parseInt(startMatch[2]);
        const endCol = colLetterToNumber(endMatch[1]);
        const endRow = parseInt(endMatch[2]);
        
        if (deletedColNumber >= startCol && deletedColNumber <= endCol) {
            if (startCol === endCol) {
                removedCount++;
                return;
            } else {
                const newStartCol = startCol;
                const newEndCol = endCol - 1;
                const newRange = colNumberToLetter(newStartCol) + startRow + ':' + colNumberToLetter(newEndCol) + endRow;
                console.log('  ' + mergeRange + ' -> ' + newRange + ' (Spalte innerhalb)');
                try {
                    worksheet.mergeCells(newRange);
                    adjustedCount++;
                } catch (e) {
                    console.log('  Fehler: ' + e.message);
                }
            }
        } else if (deletedColNumber < startCol) {
            const newStartCol = startCol - 1;
            const newEndCol = endCol - 1;
            const newRange = colNumberToLetter(newStartCol) + startRow + ':' + colNumberToLetter(newEndCol) + endRow;
            console.log('  ' + mergeRange + ' -> ' + newRange + ' (verschoben)');
            try {
                worksheet.mergeCells(newRange);
                adjustedCount++;
            } catch (e) {
                console.log('  Fehler: ' + e.message);
            }
        } else {
            console.log('  ' + mergeRange + ' (unverändert)');
            try {
                worksheet.mergeCells(mergeRange);
            } catch (e) {
                console.log('  Fehler: ' + e.message);
            }
        }
    });
    
    console.log(adjustedCount + ' Merged Cells angepasst, ' + removedCount + ' entfernt');
}

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/test-styles-exceljs.xlsx');
    
    const ws = wb.worksheets[0];
    
    console.log('=== VOR ÄNDERUNG ===');
    console.log('Merges:', ws.model.merges);
    console.log('Spalten:', ws.columnCount);
    
    // Lösche Spalte A mit spliceColumns
    ws.spliceColumns(1, 1);
    
    console.log('\n=== NACH spliceColumns ===');
    console.log('Merges (unverändert!):', ws.model.merges);
    
    // Jetzt Merged Cells anpassen
    console.log('\n=== PASSE MERGED CELLS AN ===');
    adjustMergedCellsAfterColumnDelete(ws, 1);
    
    console.log('\n=== NACH ANPASSUNG ===');
    console.log('Merges:', ws.model.merges);
    
    // Speichern
    await wb.xlsx.writeFile('/Users/nojan/Desktop/test-splice-merge-fixed.xlsx');
    console.log('\nDatei gespeichert');
    
    // Neu laden und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile('/Users/nojan/Desktop/test-splice-merge-fixed.xlsx');
    const ws2 = wb2.worksheets[0];
    console.log('\n=== NACH RELOAD ===');
    console.log('Merges:', ws2.model.merges);
    
    // Prüfe Fills
    console.log('\n=== FILL-CHECK Zeile 4 (nach Spalte A löschen) ===');
    console.log('Erwartung: A=Grün (war B), B=Blau (war C), etc.');
    for (let col = 1; col <= 8; col++) {
        const cell = ws2.getCell(4, col);
        const fill = cell.fill;
        let color = '-';
        if (fill && fill.type === 'pattern' && fill.fgColor) {
            color = fill.fgColor.argb || fill.fgColor.theme || 'x';
        }
        console.log('  Spalte ' + col + ': Fill=' + color + ', Wert=' + cell.value);
    }
}

test().catch(e => console.error(e));
