// Debug-Test für Spaltenbreiten
const ExcelJS = require('exceljs');
const path = require('path');

async function debugColumnWidths() {
    // Ersetze diesen Pfad mit deiner Test-Datei
    const testFile = process.argv[2];
    
    if (!testFile) {
        console.log('Verwendung: node test-column-widths-debug.js <excel-datei>');
        process.exit(1);
    }
    
    console.log('Lade Datei:', testFile);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(testFile);
    
    for (const worksheet of workbook.worksheets) {
        console.log('\n========================================');
        console.log('Sheet:', worksheet.name);
        console.log('Spaltenanzahl (columnCount):', worksheet.columnCount);
        console.log('Zeilenanzahl (rowCount):', worksheet.rowCount);
        console.log('----------------------------------------');
        
        // Spaltenbreiten anzeigen
        console.log('\nSpaltenbreiten:');
        for (let colIdx = 1; colIdx <= Math.min(worksheet.columnCount, 30); colIdx++) {
            const col = worksheet.getColumn(colIdx);
            const letter = getColumnLetter(colIdx);
            const width = col.width;
            const hidden = col.hidden;
            
            if (width !== undefined || hidden) {
                console.log(`  Spalte ${letter} (${colIdx}): width=${width}, hidden=${hidden}`);
            } else {
                console.log(`  Spalte ${letter} (${colIdx}): KEINE BREITE GESETZT`);
            }
        }
        
        // Zeilenhöhen anzeigen (erste 10 Zeilen)
        console.log('\nZeilenhöhen (erste 10):');
        for (let rowIdx = 1; rowIdx <= Math.min(worksheet.rowCount, 10); rowIdx++) {
            const row = worksheet.getRow(rowIdx);
            const height = row.height;
            const hidden = row.hidden;
            
            if (height !== undefined || hidden) {
                console.log(`  Zeile ${rowIdx}: height=${height}, hidden=${hidden}`);
            } else {
                console.log(`  Zeile ${rowIdx}: KEINE HÖHE GESETZT`);
            }
        }
    }
}

function getColumnLetter(colNum) {
    let letter = '';
    while (colNum > 0) {
        const remainder = (colNum - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        colNum = Math.floor((colNum - 1) / 26);
    }
    return letter;
}

debugColumnWidths().catch(err => {
    console.error('Fehler:', err);
    process.exit(1);
});
