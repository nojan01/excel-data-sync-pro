#!/usr/bin/env node
const ExcelJS = require('exceljs');

async function inspectCell() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const sheetName = 'DEFENCE&SPACE Aug-2025';
    
    console.log('Inspiziere Zelle direkt mit ExcelJS...\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    
    // Prüfe Zeile 243, Spalte 8 (hat fontColor: #FFFFFF, sollte Fill haben)
    console.log('Zeile 244 (243+1), Spalte 8:');
    const cell = worksheet.getCell(244, 8);
    console.log('  Wert:', cell.value);
    console.log('  Font:', JSON.stringify(cell.font, null, 2));
    console.log('  Fill:', JSON.stringify(cell.fill, null, 2));
    
    // Prüfe noch ein paar andere Zellen
    console.log('\nZeile 4, Spalte 2:');
    const cell2 = worksheet.getCell(4, 2);
    console.log('  Wert:', cell2.value);
    console.log('  Font:', JSON.stringify(cell2.font, null, 2));
    console.log('  Fill:', JSON.stringify(cell2.fill, null, 2));
    
    // Suche nach Zellen mit Fill
    console.log('\nSuche nach Zellen mit Fill in den ersten 100 Zeilen...');
    let fillCount = 0;
    for (let row = 1; row <= 100; row++) {
        for (let col = 1; col <= 61; col++) {
            const c = worksheet.getCell(row, col);
            if (c.fill && c.fill.type === 'pattern' && c.fill.fgColor) {
                fillCount++;
                if (fillCount <= 5) {
                    console.log(`  Zeile ${row}, Spalte ${col}: Fill =`, JSON.stringify(c.fill));
                }
            }
        }
    }
    console.log(`\nGesamt ${fillCount} Zellen mit Fill gefunden in ersten 100 Zeilen`);
}

inspectCell();
