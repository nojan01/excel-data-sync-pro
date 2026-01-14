const ExcelJS = require('exceljs');

async function check() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = workbook.worksheets[0];
    
    console.log('=== EXPORT ANALYSE ===');
    console.log('columnCount:', ws.columnCount);
    
    // Zeige letzte Spalten
    console.log('\n=== LETZTE SPALTEN ===');
    for (let col = 58; col <= 62; col++) {
        const header = ws.getRow(1).getCell(col).value;
        const letter = colNumberToLetter(col);
        console.log(`Spalte ${col} (${letter}): "${header || '(leer)'}"`);
    }
    
    // Suche CF die noch hohe Spalten referenzieren
    console.log('\n=== CF MIT HOHEN SPALTEN (BG+) ===');
    const cf = ws.conditionalFormattings || [];
    let highColCount = 0;
    
    for (const cfEntry of cf) {
        const ref = cfEntry.ref || '';
        // Suche nach Spalten >= BG (59)
        const match = ref.match(/B[G-Z]|C[A-Z]/g);
        if (match) {
            highColCount++;
            if (highColCount <= 10) {
                console.log('CF ref:', ref.substring(0, 100) + (ref.length > 100 ? '...' : ''));
            }
        }
    }
    console.log('Gesamt CF mit BG+ Spalten:', highColCount);
    console.log('Gesamt CF:', cf.length);
    
    // Prüfe Zelle in letzter Spalte
    const lastCol = ws.columnCount;
    console.log('\n=== LETZTE SPALTE', lastCol, '(' + colNumberToLetter(lastCol) + ') ===');
    console.log('Header:', ws.getRow(1).getCell(lastCol).value);
    console.log('Zeile 2:', ws.getRow(2).getCell(lastCol).value);
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

check().catch(console.error);
