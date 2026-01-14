const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = wb.getWorksheet(1);
    console.log('Sheet:', ws.name);
    
    // Speichere Styles für Zeilen 2-10, Spalten A-E VOR dem Löschen
    console.log('\n=== VOR spliceColumns ===');
    console.log('Zeile 2, Spalten A-E:');
    for (let col = 1; col <= 5; col++) {
        const cell = ws.getCell(2, col);
        const fill = cell.fill;
        const value = cell.value;
        const colLetter = String.fromCharCode(64 + col);
        console.log(`  ${colLetter}2: Value="${String(value).substring(0, 20)}", Fill=${fill?.fgColor?.argb || fill?.bgColor?.argb || 'none'}`);
    }
    
    console.log('\nZeile 5, Spalten A-E:');
    for (let col = 1; col <= 5; col++) {
        const cell = ws.getCell(5, col);
        const fill = cell.fill;
        const value = cell.value;
        const colLetter = String.fromCharCode(64 + col);
        console.log(`  ${colLetter}5: Value="${String(value).substring(0, 20)}", Fill=${fill?.fgColor?.argb || fill?.bgColor?.argb || 'none'}`);
    }
    
    // Lösche Spalte A
    console.log('\n--- Lösche Spalte A ---');
    ws.spliceColumns(1, 1);
    
    console.log('\n=== NACH spliceColumns ===');
    console.log('Zeile 2, neue Spalten A-D (ursprünglich B-E):');
    for (let col = 1; col <= 4; col++) {
        const cell = ws.getCell(2, col);
        const fill = cell.fill;
        const value = cell.value;
        const colLetter = String.fromCharCode(64 + col);
        console.log(`  ${colLetter}2: Value="${String(value).substring(0, 20)}", Fill=${fill?.fgColor?.argb || fill?.bgColor?.argb || 'none'}`);
    }
    
    console.log('\nZeile 5, neue Spalten A-D (ursprünglich B-E):');
    for (let col = 1; col <= 4; col++) {
        const cell = ws.getCell(5, col);
        const fill = cell.fill;
        const value = cell.value;
        const colLetter = String.fromCharCode(64 + col);
        console.log(`  ${colLetter}5: Value="${String(value).substring(0, 20)}", Fill=${fill?.fgColor?.argb || fill?.bgColor?.argb || 'none'}`);
    }
    
    // Prüfe ob Column-Styles existieren
    console.log('\n=== Column-Level Styles ===');
    for (let col = 1; col <= 5; col++) {
        const column = ws.getColumn(col);
        console.log(`Spalte ${col}: width=${column.width}, style=${JSON.stringify(column.style || {}).substring(0, 100)}`);
    }
    
    // Prüfe Row-Level Styles
    console.log('\n=== Row-Level Styles (Zeilen 2-5) ===');
    for (let row = 2; row <= 5; row++) {
        const rowObj = ws.getRow(row);
        console.log(`Zeile ${row}: style=${JSON.stringify(rowObj.style || {}).substring(0, 100)}`);
    }
}

test().catch(console.error);
