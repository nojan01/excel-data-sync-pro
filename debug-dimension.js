const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('VOR splice:');
    console.log('  dimensions:', ws.dimensions.range);
    
    // Lösche Spalte 1
    ws.spliceColumns(1, 1);
    
    console.log('');
    console.log('NACH splice:');
    console.log('  dimensions:', ws.dimensions.range);
    
    // Prüfe welche Zeilen noch Spalte 61 haben
    console.log('');
    console.log('Zeilen mit Daten in Spalte 61 (BI):');
    let count = 0;
    ws.eachRow((row, rowNumber) => {
        const cell = row.getCell(61);
        if (cell.value !== null && cell.value !== undefined) {
            if (count < 5) {
                console.log('  Zeile', rowNumber, ':', cell.value, '(type:', typeof cell.value, ')');
            }
            count++;
        }
    });
    console.log('  Gesamt:', count, 'Zeilen');
    
    // Prüfe rowDims für erste paar Zeilen
    console.log('');
    console.log('Row dimensions für erste Zeilen:');
    for (let i = 1; i <= 5; i++) {
        const row = ws.getRow(i);
        console.log('  Zeile', i, ':', row.dimensions);
    }
}

test().catch(console.error);
