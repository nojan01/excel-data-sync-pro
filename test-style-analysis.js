#!/usr/bin/env node
const { readSheetWithExcelJS } = require('./exceljs-reader');

async function analyzeStyles() {
    const result = await readSheetWithExcelJS('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx', 'DEFENCE&SPACE Aug-2025');
    
    const allKeys = Object.keys(result.cellStyles);
    const headerKeys = allKeys.filter(k => k.startsWith('0-'));
    const dataKeys = allKeys.filter(k => !k.startsWith('0-'));
    
    console.log('Style-Analyse:');
    console.log(`  Gesamt: ${allKeys.length}`);
    console.log(`  Header (0-*): ${headerKeys.length}`);
    console.log(`  Daten (1-*): ${dataKeys.length}\n`);
    
    console.log('Daten-Styles (alle):');
    dataKeys.forEach(key => {
        console.log(`  ${key}: ${JSON.stringify(result.cellStyles[key])}`);
    });
}

analyzeStyles();
