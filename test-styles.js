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
    
    // Zähle Styles mit Fill
    const stylesWithFill = Object.entries(result.cellStyles).filter(([k, v]) => v.fill);
    console.log(`   Styles mit Fill-Farbe: ${stylesWithFill.length}`);
    
    // Zeige erste 5 Styles
    console.log('\nErste 10 Styles:');
    const styleEntries = Object.entries(result.cellStyles).slice(0, 10);
    styleEntries.forEach(([key, style]) => {
        console.log(`   ${key}:`, JSON.stringify(style));
    });
    
    // Zeige RichText Details
    if (Object.keys(result.richTextCells).length > 0) {
        console.log('\nRichText Zellen:');
        Object.entries(result.richTextCells).forEach(([key, fragments]) => {
            console.log(`   ${key}:`, JSON.stringify(fragments));
        });
    }
    
    // Zeige Styles mit Fill
    if (stylesWithFill.length > 0) {
        console.log('\nStyles mit Fill (erste 5):');
        stylesWithFill.slice(0, 5).forEach(([key, style]) => {
            console.log(`   ${key}:`, JSON.stringify(style));
        });
    }
}

testStyles();
