const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie.xlsx');
    
    const ws = workbook.getWorksheet('DEFENCE&SPACE Aug-2025');
    
    console.log('=== VOR dem Schreiben ===');
    console.log('_columns Anzahl:', ws._columns?.length);
    console.log('');
    
    // Spaltenbreiten VOR dem Schreiben
    console.log('Spaltenbreiten VOR:');
    for (let i = 1; i <= 10; i++) {
        console.log('  Spalte ' + i + ': width=' + ws.getColumn(i).width);
    }
    
    // Sichere die Spaltenbreiten JETZT
    const savedWidths = {};
    for (let i = 1; i <= 61; i++) {
        const col = ws.getColumn(i);
        if (col.width !== undefined && col.width !== null) {
            savedWidths[i] = col.width;
        }
    }
    console.log('\nGesicherte Breiten:', Object.keys(savedWidths).length);
    
    // Jetzt schreiben wir Daten - nur 10 Zeilen zum Test
    console.log('\n=== Schreibe Daten ===');
    for (let row = 1; row <= 10; row++) {
        for (let col = 1; col <= 61; col++) {
            const cell = ws.getCell(row, col);
            cell.value = cell.value; // Gleiches Value setzen
        }
    }
    
    console.log('\n=== NACH dem Schreiben ===');
    console.log('_columns Anzahl:', ws._columns?.length);
    
    // Prüfe Spaltenbreiten NACH dem Schreiben
    console.log('\nSpaltenbreiten NACH:');
    let missingCount = 0;
    for (let i = 1; i <= 10; i++) {
        const width = ws.getColumn(i).width;
        console.log('  Spalte ' + i + ': width=' + width);
        if (width === undefined) missingCount++;
    }
    
    // Zähle wie viele Spalten jetzt keine Breite haben
    for (let i = 1; i <= 61; i++) {
        if (ws.getColumn(i).width === undefined) missingCount++;
    }
    console.log('\nFehlende Breiten (von 61):', missingCount);
    
    // Stelle die Breiten wieder her
    console.log('\n=== Stelle Spaltenbreiten wieder her ===');
    for (const [colIdx, width] of Object.entries(savedWidths)) {
        ws.getColumn(parseInt(colIdx)).width = width;
    }
    
    console.log('\nSpaltenbreiten NACH Wiederherstellung:');
    for (let i = 1; i <= 10; i++) {
        console.log('  Spalte ' + i + ': width=' + ws.getColumn(i).width);
    }
}

test().catch(console.error);
