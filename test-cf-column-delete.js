// Test: Simuliere Spaltenlöschung und prüfe CF-Anpassung
const ExcelJS = require('exceljs');
const fs = require('fs');

// Importiere die Funktionen aus exceljs-writer
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
const testOutputPath = '/tmp/test-cf-after-delete.xlsx';

async function testCFAfterColumnDelete() {
    console.log('=== Test: CF nach Spaltenlöschung ===');
    console.log('Datei:', filePath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    console.log('\nSheet:', worksheet.name);
    console.log('Spalten VOR Löschen:', worksheet.columnCount);
    
    const cf = worksheet.conditionalFormattings;
    console.log('CF-Regeln:', cf.length);
    
    // Zeige die ersten 3 CF vor Löschen
    console.log('\n=== CF VOR Löschen (erste 3) ===');
    cf.slice(0, 3).forEach((cfEntry, idx) => {
        console.log('CF ' + (idx + 1) + ' ref:', cfEntry.ref);
    });
    
    // Welche Spalte löschen wir? Lass uns Spalte E (5) löschen als Beispiel
    const deletedColExcel = 5; // Spalte E
    console.log('\n=== LÖSCHE Spalte ' + deletedColExcel + ' (' + colNumberToLetter(deletedColExcel) + ') ===');
    
    const lastColumnBeforeDelete = worksheet.columnCount;
    console.log('lastColumnBeforeDelete:', lastColumnBeforeDelete, '(' + colNumberToLetter(lastColumnBeforeDelete) + ')');
    
    // spliceColumns aufrufen
    worksheet.spliceColumns(deletedColExcel, 1);
    
    console.log('\n=== CF NACH spliceColumns (VOR adjustConditionalFormattings) ===');
    cf.slice(0, 3).forEach((cfEntry, idx) => {
        console.log('CF ' + (idx + 1) + ' ref:', cfEntry.ref);
    });
    
    // Hier sollte adjustConditionalFormattingsAfterColumnDelete aufgerufen werden
    // Aber lass uns prüfen, ob die CF schon angepasst wurde oder nicht
    
    // Prüfe ob die Spaltenreferenzen korrekt angepasst wurden
    console.log('\n=== Analyse: Wurden CF-Refs angepasst? ===');
    const firstCF = cf[0];
    if (firstCF) {
        const expectedNewRef = 'AM2135:AX2404'; // AN → AM, AY → AX nach Löschen von E
        console.log('Original ref:', 'AN2135:AY2404');
        console.log('Aktuelle ref:', firstCF.ref);
        console.log('Erwartet ref:', expectedNewRef);
        console.log('ExcelJS hat CF automatisch angepasst:', firstCF.ref !== 'AN2135:AY2404');
    }
    
    // Speichere und prüfe nach Neu-Laden
    await workbook.xlsx.writeFile(testOutputPath);
    
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(testOutputPath);
    const worksheet2 = workbook2.getWorksheet(1);
    
    console.log('\n=== Nach Speichern und Neu-Laden ===');
    console.log('CF-Regeln:', worksheet2.conditionalFormattings.length);
    worksheet2.conditionalFormattings.slice(0, 3).forEach((cfEntry, idx) => {
        console.log('CF ' + (idx + 1) + ' ref:', cfEntry.ref);
    });
    
    fs.unlinkSync(testOutputPath);
}

testCFAfterColumnDelete().catch(console.error);
