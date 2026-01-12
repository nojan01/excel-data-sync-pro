#!/usr/bin/env node

/**
 * ExcelJS vs xlsx-populate Performance-Test
 * 
 * Verwendung:
 *   node test-exceljs.js <excel-datei> <sheet-name>
 * 
 * Beispiel:
 *   node test-exceljs.js test.xlsx "Sheet1"
 */

const { readSheetWithExcelJS } = require('./exceljs-reader');
const XlsxPopulate = require('xlsx-populate');
const path = require('path');

async function runTest(filePath, sheetName) {
    console.log('\n╔══════════════════════════════════════════════════════╗');
    console.log('║     ExcelJS vs xlsx-populate Performance-Test       ║');
    console.log('╚══════════════════════════════════════════════════════╝\n');
    console.log(`Datei: ${path.basename(filePath)}`);
    console.log(`Sheet: ${sheetName}\n`);
    
    try {
        // Test 1: xlsx-populate (aktuell)
        console.log('► Test 1: xlsx-populate (aktuell)...');
        const populateStart = Date.now();
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const worksheet = workbook.sheet(sheetName);
        
        if (!worksheet) {
            console.error(`❌ Sheet "${sheetName}" nicht gefunden!`);
            process.exit(1);
        }
        
        const usedRange = worksheet.usedRange();
        const populateRows = usedRange ? (usedRange.endCell().rowNumber() - 1) : 0;
        const populateCols = usedRange ? usedRange.endCell().columnNumber() : 0;
        const populateTime = Date.now() - populateStart;
        
        console.log(`   ✓ ${populateTime}ms`);
        console.log(`   ✓ ${populateRows} Zeilen × ${populateCols} Spalten`);
        console.log(`   ✓ ${(populateRows * populateCols).toLocaleString()} Zellen\n`);
        
        // Test 2: ExcelJS (neu)
        console.log('► Test 2: ExcelJS (neu)...');
        const exceljsStart = Date.now();
        const exceljsResult = await readSheetWithExcelJS(filePath, sheetName);
        const exceljsTime = Date.now() - exceljsStart;
        
        if (!exceljsResult.success) {
            console.error(`❌ ExcelJS Fehler: ${exceljsResult.error}`);
            process.exit(1);
        }
        
        console.log(`   ✓ ${exceljsTime}ms`);
        console.log(`   ✓ ${exceljsResult.data.length} Zeilen × ${exceljsResult.headers.length} Spalten`);
        console.log(`   ✓ ${(exceljsResult.data.length * exceljsResult.headers.length).toLocaleString()} Zellen\n`);
        
        // Vergleich
        console.log('╔══════════════════════════════════════════════════════╗');
        console.log('║                     ERGEBNIS                         ║');
        console.log('╚══════════════════════════════════════════════════════╝\n');
        
        const speedup = ((populateTime - exceljsTime) / populateTime * 100);
        const faster = speedup > 0 ? 'schneller' : 'langsamer';
        const icon = speedup > 0 ? '🚀' : '🐌';
        
        console.log(`${icon} ExcelJS ist ${Math.abs(speedup).toFixed(1)}% ${faster}`);
        console.log(`   xlsx-populate: ${populateTime}ms`);
        console.log(`   ExcelJS:       ${exceljsTime}ms`);
        console.log(`   Differenz:     ${Math.abs(populateTime - exceljsTime)}ms\n`);
        
        // Daten-Qualität
        console.log('📊 Daten-Qualität:');
        console.log(`   Styles:     ${Object.keys(exceljsResult.cellStyles).length} Zellen`);
        console.log(`   Formeln:    ${Object.keys(exceljsResult.cellFormulas).length} Zellen`);
        console.log(`   Hyperlinks: ${Object.keys(exceljsResult.cellHyperlinks).length} Zellen`);
        console.log(`   RichText:   ${Object.keys(exceljsResult.richTextCells).length} Zellen`);
        console.log(`   Versteckte Spalten: ${exceljsResult.hiddenColumns.length}`);
        console.log(`   Versteckte Zeilen:  ${exceljsResult.hiddenRows.length}\n`);
        
    } catch (error) {
        console.error('❌ Fehler:', error.message);
        process.exit(1);
    }
}

// Kommandozeilen-Argumente
const args = process.argv.slice(2);

if (args.length < 2) {
    console.log('Verwendung: node test-exceljs.js <excel-datei> <sheet-name>');
    console.log('Beispiel:   node test-exceljs.js test.xlsx "Sheet1"');
    process.exit(1);
}

const [filePath, sheetName] = args;
runTest(filePath, sheetName);
