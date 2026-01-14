// Test: Prüfe ob CF-Änderungen beim Speichern erhalten bleiben
const ExcelJS = require('exceljs');
const fs = require('fs');

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
const testOutputPath = '/tmp/test-cf-save.xlsx';

async function testCFSave() {
    console.log('=== Test: CF beim Speichern ===\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    const cf = worksheet.conditionalFormattings;
    
    console.log('Original CF[0].ref:', cf[0]?.ref);
    
    // Ändere die erste CF-Referenz manuell
    if (cf[0]) {
        cf[0].ref = 'AM2135:AX2404';
        console.log('Geändert CF[0].ref zu:', cf[0].ref);
    }
    
    // Prüfe VOR dem Speichern
    console.log('\nVOR Speichern - CF[0].ref:', worksheet.conditionalFormattings[0]?.ref);
    
    // Speichere
    await workbook.xlsx.writeFile(testOutputPath);
    console.log('Datei gespeichert nach:', testOutputPath);
    
    // Lade neu und prüfe
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(testOutputPath);
    const worksheet2 = workbook2.getWorksheet(1);
    
    console.log('\nNACH Speichern und Neu-Laden:');
    console.log('CF[0].ref:', worksheet2.conditionalFormattings[0]?.ref);
    console.log('Erwartet: AM2135:AX2404');
    console.log('KORREKT:', worksheet2.conditionalFormattings[0]?.ref === 'AM2135:AX2404' ? 'JA' : 'NEIN - PROBLEM!');
    
    fs.unlinkSync(testOutputPath);
}

testCFSave().catch(console.error);
