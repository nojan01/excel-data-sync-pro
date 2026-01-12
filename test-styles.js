#!/usr/bin/env node
const { readSheetWithExcelJS } = require('./exceljs-reader');

async function testStyles() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const sheetName = 'DEFENCE&SPACE Aug-2025';
    
    console.log('Teste Style-Extraktion...\n');
    
    const result = await readSheetWithExcelJS(filePath, sheetName);
    
    if (!result.success) {
        console.error('Fehler:', result.error);
        process.exit(1);
    }
    
    console.log('✅ Sheet geladen');
    console.log(`   Zeilen: ${result.data.length}`);
    console.log(`   Spalten: ${result.headers.length}`);
    console.log(`   Styles gefunden: ${Object.keys(result.cellStyles).length}`);
    console.log(`   RichText: ${Object.keys(result.richTextCells).length}`);
    console.log(`   Formulas: ${Object.keys(result.cellFormulas).length}`);
    console.log(`   Hyperlinks: ${Object.keys(result.cellHyperlinks).length}`);
    console.log(`   AutoFilter: ${result.autoFilterRange || 'Nicht vorhanden'}\n`);
    
    // Zeige erste 5 Styles
    console.log('Erste 5 Styles:');
    const styleEntries = Object.entries(result.cellStyles).slice(0, 5);
    styleEntries.forEach(([key, style]) => {
        console.log(`   ${key}:`, JSON.stringify(style));
    });
}

testStyles();
