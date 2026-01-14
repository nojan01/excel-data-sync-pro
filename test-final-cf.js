// Finaler Test: Spaltenlöschung mit CF-Anpassung
const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Test');
    
    // Spalten A bis E mit Daten
    ws.getCell('A1').value = 'A';
    ws.getCell('B1').value = 'B';
    ws.getCell('C1').value = 'C';
    ws.getCell('D1').value = 'D';
    ws.getCell('E1').value = 'E';
    
    // CF auf Spalte C (Spalte 3)
    ws.addConditionalFormatting({
        ref: 'C1:C10',
        rules: [{
            type: 'cellIs',
            operator: 'greaterThan',
            priority: 1,
            formulae: ['5'],
            style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } }
        }]
    });
    
    // CF auf Spalte D (Spalte 4)
    ws.addConditionalFormatting({
        ref: 'D1:D10',
        rules: [{
            type: 'cellIs',
            operator: 'lessThan',
            priority: 2,
            formulae: ['3'],
            style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } }
        }]
    });
    
    console.log('=== VOR Spaltenlöschung ===');
    console.log('Spalten:', ws.getCell('A1').value, ws.getCell('B1').value, ws.getCell('C1').value, ws.getCell('D1').value, ws.getCell('E1').value);
    console.log('CF:');
    ws.conditionalFormattings.forEach((cf, i) => console.log('  ' + i + ': ref=' + cf.ref));
    
    // Spalte A löschen (Index 1)
    const deletedCol = 1;
    const lastColBefore = 5; // E ist Spalte 5
    
    console.log('\n=== Lösche Spalte A ===');
    ws.spliceColumns(deletedCol, 1);
    
    console.log('\n=== NACH spliceColumns (OHNE manuelle Anpassung) ===');
    console.log('Spalten:', ws.getCell('A1').value, ws.getCell('B1').value, ws.getCell('C1').value, ws.getCell('D1').value, ws.getCell('E1').value);
    console.log('CF (NICHT angepasst - ExcelJS macht das nicht!):');
    ws.conditionalFormattings.forEach((cf, i) => console.log('  ' + i + ': ref=' + cf.ref));
    
    // Jetzt die manuelle Anpassung simulieren
    console.log('\n=== MANUELLE CF-Anpassung ===');
    
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
    
    // Passe alle CF refs an
    const emptyLastCol = lastColBefore; // Spalte 5 (E) ist jetzt leer
    const toRemove = [];
    
    ws.conditionalFormattings.forEach((cf, i) => {
        const oldRef = cf.ref;
        const adjustedRef = adjustRangeReference(oldRef, deletedCol);
        
        // Prüfe ob die angepasste Ref nur auf die leere Spalte zeigt
        const match = adjustedRef.match(/^([A-Z]+)/);
        if (match && colLetterToNumber(match[1]) === emptyLastCol) {
            console.log('  ENTFERNE: ' + oldRef + ' → ' + adjustedRef + ' (zeigt auf leere Spalte)');
            toRemove.push(i);
        } else {
            console.log('  ANPASSEN: ' + oldRef + ' → ' + adjustedRef);
            cf.ref = adjustedRef;
        }
    });
    
    // Entferne von hinten
    for (let i = toRemove.length - 1; i >= 0; i--) {
        ws.conditionalFormattings.splice(toRemove[i], 1);
    }
    
    console.log('\n=== NACH manueller Anpassung ===');
    console.log('Spalten:', ws.getCell('A1').value, ws.getCell('B1').value, ws.getCell('C1').value, ws.getCell('D1').value);
    console.log('CF:');
    ws.conditionalFormattings.forEach((cf, i) => console.log('  ' + i + ': ref=' + cf.ref));
    
    console.log('\n=== ERWARTETES ERGEBNIS ===');
    console.log('Daten: B wird zu A, C wird zu B, D wird zu C, E wird zu D');
    console.log('CF auf C1:C10 (altes D) → sollte jetzt C1:C10 sein');
    console.log('CF auf D1:D10 (altes E, leer) → sollte entfernt werden');
    
    // Speichern
    await workbook.xlsx.writeFile('/tmp/test-cf-final.xlsx');
    console.log('\n✓ Gespeichert in /tmp/test-cf-final.xlsx');
}

test().catch(console.error);
