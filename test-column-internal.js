const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie.xlsx');
    
    const ws = workbook.getWorksheet('DEFENCE&SPACE Aug-2025');
    
    console.log('columnCount:', ws.columnCount);
    console.log('columns:', ws.columns?.length);
    console.log('');
    
    // Direkter Zugriff auf _columns
    console.log('_columns Anzahl:', ws._columns?.length);
    
    // Prüfe interne Struktur
    if (ws._columns) {
        console.log('\nSpalten mit Breite in _columns:');
        ws._columns.forEach((col, i) => {
            if (col && col.width) {
                console.log('  _columns[' + i + ']: width=' + col.width);
            }
        });
    }
    
    // Vergleiche mit getColumn
    console.log('\nPrüfe Spalten 1-20 mit getColumn():');
    for (let i = 1; i <= 20; i++) {
        const col = ws.getColumn(i);
        const letter = String.fromCharCode(64 + (i > 26 ? (Math.floor((i-1)/26) + 64) : 0)) + 
                       String.fromCharCode(64 + ((i-1) % 26) + 1);
        console.log('Spalte ' + i + ' (' + letter + '): width=' + col.width);
    }
    
    // Prüfe ob es eine cols-XML gibt
    console.log('\nPrüfe columns property:');
    if (ws.columns) {
        console.log('columns.length:', ws.columns.length);
        for (let i = 0; i < Math.min(10, ws.columns.length); i++) {
            const col = ws.columns[i];
            if (col) {
                console.log('columns[' + i + ']: width=' + col.width);
            }
        }
    }
}

test().catch(console.error);
