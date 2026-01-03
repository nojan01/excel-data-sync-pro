/**
 * Test: xlsx-populate vs ExcelJS - Formatierungserhalt
 * 
 * Dieses Skript testet, ob xlsx-populate die Formatierung besser erhält als ExcelJS.
 * 
 * Ausführen: node test-xlsx-populate.js
 */

const XlsxPopulate = require('xlsx-populate');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Testdatei - bitte anpassen!
const TEST_FILE = process.argv[2] || 'test-input.xlsx';
const OUTPUT_XLSXPOPULATE = 'test-output-xlsxpopulate.xlsx';
const OUTPUT_EXCELJS = 'test-output-exceljs.xlsx';

async function testXlsxPopulate(inputFile, outputFile) {
    console.log('\n=== Test: xlsx-populate ===');
    
    try {
        // Datei laden
        const workbook = await XlsxPopulate.fromFileAsync(inputFile);
        
        // Sheets auflisten
        const sheets = workbook.sheets();
        console.log(`Sheets gefunden: ${sheets.length}`);
        sheets.forEach(sheet => {
            console.log(`  - ${sheet.name()}`);
        });
        
        // Erstes Sheet
        const sheet = sheets[0];
        
        // Letzte Zeile finden
        const usedRange = sheet.usedRange();
        const lastRow = usedRange ? usedRange.endCell().rowNumber() : 1;
        console.log(`Letzte Zeile: ${lastRow}`);
        
        // Neue Zeile hinzufügen
        const newRowNum = lastRow + 1;
        sheet.cell(`A${newRowNum}`).value('TEST');
        sheet.cell(`B${newRowNum}`).value('xlsx-populate Test');
        sheet.cell(`C${newRowNum}`).value(new Date().toISOString());
        
        console.log(`Neue Zeile ${newRowNum} hinzugefügt`);
        
        // Speichern
        await workbook.toFileAsync(outputFile);
        console.log(`? Gespeichert als: ${outputFile}`);
        
        return true;
    } catch (error) {
        console.error(`? Fehler: ${error.message}`);
        return false;
    }
}

async function testExcelJS(inputFile, outputFile) {
    console.log('\n=== Test: ExcelJS ===');
    
    try {
        // Datei laden
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(inputFile);
        
        // Sheets auflisten
        console.log(`Sheets gefunden: ${workbook.worksheets.length}`);
        workbook.worksheets.forEach(sheet => {
            console.log(`  - ${sheet.name}`);
        });
        
        // Erstes Sheet
        const sheet = workbook.worksheets[0];
        
        // Letzte Zeile finden
        const lastRow = sheet.rowCount;
        console.log(`Letzte Zeile: ${lastRow}`);
        
        // Neue Zeile hinzufügen
        const newRowNum = lastRow + 1;
        const newRow = sheet.getRow(newRowNum);
        newRow.getCell(1).value = 'TEST';
        newRow.getCell(2).value = 'ExcelJS Test';
        newRow.getCell(3).value = new Date().toISOString();
        newRow.commit();
        
        console.log(`Neue Zeile ${newRowNum} hinzugefügt`);
        
        // Speichern
        await workbook.xlsx.writeFile(outputFile);
        console.log(`? Gespeichert als: ${outputFile}`);
        
        return true;
    } catch (error) {
        console.error(`? Fehler: ${error.message}`);
        return false;
    }
}

async function main() {
    console.log('========================================');
    console.log('Formatierungserhalt-Test');
    console.log('========================================');
    
    // Prüfen ob Testdatei existiert
    if (!fs.existsSync(TEST_FILE)) {
        console.log(`\n??  Testdatei nicht gefunden: ${TEST_FILE}`);
        console.log('\nBitte eine Excel-Datei mit Formatierungen bereitstellen:');
        console.log('  node test-xlsx-populate.js "pfad/zur/datei.xlsx"');
        console.log('\nOder eine test-input.xlsx im aktuellen Ordner erstellen.');
        return;
    }
    
    console.log(`\nTestdatei: ${TEST_FILE}`);
    console.log(`Dateigröße: ${(fs.statSync(TEST_FILE).size / 1024).toFixed(1)} KB`);
    
    // Tests ausführen
    await testXlsxPopulate(TEST_FILE, OUTPUT_XLSXPOPULATE);
    await testExcelJS(TEST_FILE, OUTPUT_EXCELJS);
    
    console.log('\n========================================');
    console.log('Test abgeschlossen!');
    console.log('========================================');
    console.log('\nBitte die Ausgabedateien in Excel öffnen und vergleichen:');
    console.log(`  1. ${OUTPUT_XLSXPOPULATE} (xlsx-populate)`);
    console.log(`  2. ${OUTPUT_EXCELJS} (ExcelJS)`);
    console.log('\nPrüfen Sie:');
    console.log('  - Zellformatierungen (Farben, Schriftarten)');
    console.log('  - Bedingte Formatierungen');
    console.log('  - Spaltenbreiten');
    console.log('  - Rahmen und Linien');
    console.log('  - Formeln');
}

main().catch(console.error);
