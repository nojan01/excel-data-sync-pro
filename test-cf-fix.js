const ExcelJS = require('exceljs');

// Hilfsfunktionen
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
    const parts = rangeRef.split(':');
    if (parts.length === 2) {
        return adjustCellReference(parts[0], deletedColNumber) + ':' + adjustCellReference(parts[1], deletedColNumber);
    }
    return adjustCellReference(rangeRef, deletedColNumber);
}

function adjustConditionalFormattingsAfterColumnDelete(worksheet, deletedColNumber) {
    const cf = worksheet.conditionalFormattings;
    if (!cf || !Array.isArray(cf) || cf.length === 0) {
        console.log('Keine bedingten Formatierungen zu aktualisieren');
        return;
    }
    
    console.log('Passe ' + cf.length + ' bedingte Formatierungsregeln an (Spalte ' + deletedColNumber + ' gelöscht)');
    
    let adjustedCount = 0;
    cf.forEach(cfEntry => {
        if (cfEntry.ref) {
            const oldRef = cfEntry.ref;
            const newRef = adjustRangeReference(oldRef, deletedColNumber);
            if (oldRef !== newRef) {
                cfEntry.ref = newRef;
                adjustedCount++;
            }
        }
    });
    
    console.log(adjustedCount + ' CF-Referenzen angepasst');
}

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = wb.getWorksheet(1);
    console.log('Sheet:', ws.name);
    console.log('CF Regeln:', ws.conditionalFormattings.length);
    
    // Zeige erste 3 CF refs VOR
    console.log('\n=== VOR spliceColumns ===');
    ws.conditionalFormattings.slice(0, 3).forEach(cf => console.log('  ', cf.ref));
    
    // Lösche Spalte 1 (A)
    console.log('\nLösche Spalte A (1)...');
    ws.spliceColumns(1, 1);
    
    // OHNE Anpassung
    console.log('\n=== NACH spliceColumns (ohne Anpassung) ===');
    ws.conditionalFormattings.slice(0, 3).forEach(cf => console.log('  ', cf.ref));
    
    // MIT Anpassung
    adjustConditionalFormattingsAfterColumnDelete(ws, 1);
    
    console.log('\n=== NACH CF-Anpassung ===');
    ws.conditionalFormattings.slice(0, 3).forEach(cf => console.log('  ', cf.ref));
    
    // Test der Logik
    console.log('\n=== Test der Konvertierung ===');
    console.log('AN2135 -> AM2135:', adjustCellReference('AN2135', 1));
    console.log('AY2404 -> AX2404:', adjustCellReference('AY2404', 1));
    console.log('A1 -> A1:', adjustCellReference('A1', 1));  // A bleibt A (wird nicht verschoben)
    console.log('B1 -> A1:', adjustCellReference('B1', 1));  // B wird zu A
    
    // Speichern zum Testen
    const outPath = '/Users/nojan/Desktop/test-cf-adjusted.xlsx';
    await wb.xlsx.writeFile(outPath);
    console.log('\nGespeichert:', outPath);
}

test().catch(console.error);
