// Test: Prüfe ob zweiter spliceColumns die CF wieder kaputt macht
const ExcelJS = require('exceljs');

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

const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';

async function test() {
    console.log('=== Test: Effekt des zweiten spliceColumns auf CF ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    const cf = worksheet.conditionalFormattings;
    
    console.log('1. Original CF[0].ref:', cf[0]?.ref);
    console.log('   columnCount:', worksheet.columnCount);
    
    // Erster spliceColumns (Spalte 1 = A löschen)
    console.log('\n2. Führe ersten spliceColumns(1, 1) aus...');
    worksheet.spliceColumns(1, 1);
    console.log('   CF[0].ref nach spliceColumns:', cf[0]?.ref);
    console.log('   columnCount:', worksheet.columnCount);
    
    // Manuelle Anpassung (wie unsere Funktion)
    console.log('\n3. Manuelle CF-Anpassung (simuliert adjustConditionalFormattingsAfterColumnDelete)...');
    const originalRef = cf[0].ref; // z.B. AN2135:AY2404
    cf[0].ref = 'AM2135:AX2404';
    console.log('   CF[0].ref nach manueller Anpassung:', cf[0].ref);
    
    // Prüfe VOR dem zweiten splice
    console.log('\n4. VOR zweitem spliceColumns:');
    console.log('   CF[0].ref:', cf[0]?.ref);
    console.log('   columnCount:', worksheet.columnCount);
    
    // Zweiter spliceColumns (letzte Spalte löschen)
    const lastCol = worksheet.columnCount;
    console.log('\n5. Führe zweiten spliceColumns(' + lastCol + ', 1) aus (Spalte ' + colNumberToLetter(lastCol) + ')...');
    worksheet.spliceColumns(lastCol, 1);
    
    console.log('   CF[0].ref NACH zweitem spliceColumns:', cf[0]?.ref);
    console.log('   columnCount:', worksheet.columnCount);
    
    if (cf[0]?.ref === 'AM2135:AX2404') {
        console.log('\n✅ CF ist korrekt: AM2135:AX2404');
    } else {
        console.log('\n❌ CF wurde geändert! Erwartet: AM2135:AX2404, Gefunden:', cf[0]?.ref);
    }
    
    // Speichere und prüfe ob Änderungen persistiert werden
    const testPath = '/tmp/test-second-splice.xlsx';
    await workbook.xlsx.writeFile(testPath);
    console.log('\n6. Datei gespeichert, lade neu...');
    
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(testPath);
    const ws2 = workbook2.getWorksheet(1);
    console.log('   CF[0].ref nach Reload:', ws2.conditionalFormattings[0]?.ref);
    
    require('fs').unlinkSync(testPath);
}

test().catch(console.error);
