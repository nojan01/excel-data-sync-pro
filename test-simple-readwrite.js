const ExcelJS = require('exceljs');

async function testSimpleReadWrite() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const testPath = '/tmp/test-simple-readwrite.xlsx';
    
    console.log('=== TEST: Einfaches Lesen + Schreiben ohne Änderungen ===\n');
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(originalPath);
    
    const ws = wb.getWorksheet(1);
    
    // Prüfe Styles vor dem Schreiben
    console.log('Styles vor dem Schreiben:');
    for (let row = 2; row <= 5; row++) {
        const cell = ws.getCell(row, 1);
        console.log(`  Zeile ${row}, Spalte A:`);
        console.log(`    Fill:`, JSON.stringify(cell.style?.fill || 'none'));
        console.log(`    Font:`, JSON.stringify(cell.style?.font || 'none'));
    }
    
    // Schreiben ohne Änderungen
    await wb.xlsx.writeFile(testPath);
    console.log('\nDatei geschrieben:', testPath);
    
    // Neu lesen und prüfen
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(testPath);
    const ws2 = wb2.getWorksheet(1);
    
    console.log('\nStyles nach dem Schreiben:');
    for (let row = 2; row <= 5; row++) {
        const cell = ws2.getCell(row, 1);
        console.log(`  Zeile ${row}, Spalte A:`);
        console.log(`    Fill:`, JSON.stringify(cell.style?.fill || 'none'));
        console.log(`    Font:`, JSON.stringify(cell.style?.font || 'none'));
    }
    
    // XML-Vergleich
    const AdmZip = require('adm-zip');
    const zipOrig = new AdmZip(originalPath);
    const zipTest = new AdmZip(testPath);
    
    const stylesOrig = zipOrig.readAsText('xl/styles.xml');
    const stylesTest = zipTest.readAsText('xl/styles.xml');
    
    const fillsOrigCount = (stylesOrig.match(/<fill>/g) || []).length;
    const fillsTestCount = (stylesTest.match(/<fill>/g) || []).length;
    
    console.log('\n=== XML Analyse ===');
    console.log('Fills Original:', fillsOrigCount);
    console.log('Fills Test:', fillsTestCount);
}

testSimpleReadWrite().catch(console.error);
