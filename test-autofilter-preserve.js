const ExcelJS = require('exceljs');
const fs = require('fs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie 2.xlsx');
    
    const ws = workbook.getWorksheet(1);
    console.log('VOR Speichern:');
    console.log('autoFilter:', ws.autoFilter);
    console.log('tables:', Object.keys(ws.tables || {}));
    
    if (ws.tables && Object.keys(ws.tables).length > 0) {
        const tableName = Object.keys(ws.tables)[0];
        console.log('Table autoFilterRef:', ws.tables[tableName].table?.autoFilterRef);
    }
    
    // Speichern ohne Änderungen
    const testPath = '/tmp/test-autofilter.xlsx';
    await workbook.xlsx.writeFile(testPath);
    
    // Neu laden und prüfen
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(testPath);
    const ws2 = workbook2.getWorksheet(1);
    
    console.log('\nNACH Speichern:');
    console.log('autoFilter:', ws2.autoFilter);
    console.log('tables:', Object.keys(ws2.tables || {}));
    
    if (ws2.tables && Object.keys(ws2.tables).length > 0) {
        const tableName = Object.keys(ws2.tables)[0];
        console.log('Table autoFilterRef:', ws2.tables[tableName].table?.autoFilterRef);
    }
    
    fs.unlinkSync(testPath);
}

test().catch(console.error);
