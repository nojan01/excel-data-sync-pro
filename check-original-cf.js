const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function checkOriginalCF() {
    const files = fs.readdirSync('/Users/nojan/Desktop').filter(f => f.endsWith('.xlsx'));
    const mvmsFile = files.find(f => f.includes('MVMS') || f.includes('Defence') || f.includes('DEFENCE'));
    
    if (!mvmsFile) {
        console.log('Keine MVMS-Datei gefunden auf Desktop');
        console.log('Verfügbare .xlsx Dateien:', files);
        return;
    }
    
    const filePath = path.join('/Users/nojan/Desktop', mvmsFile);
    console.log('Lade:', filePath);
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    
    const ws = wb.worksheets[0];
    console.log('\nSheet:', ws.name);
    console.log('Column Count:', ws.columnCount);
    
    const cf = ws.conditionalFormattings;
    if (!cf || cf.length === 0) {
        console.log('Keine CF gefunden!');
        return;
    }
    
    console.log('\n=== CF-Regeln die auf Spalte A zeigen ===');
    let countA = 0;
    cf.forEach((cfEntry, idx) => {
        if (cfEntry.ref && cfEntry.ref.match(/\bA\d/)) {
            console.log(`CF[${idx}]: ${cfEntry.ref}`);
            countA++;
        }
    });
    console.log(`\nGesamt: ${countA} CF-Regeln mit Spalte A`);
    
    console.log('\n=== CF-Regeln die NUR auf Spalte A zeigen ===');
    function refOnlyReferencesColumn(ref, colLetter) {
        const ranges = ref.split(' ');
        for (const range of ranges) {
            const parts = range.split(':');
            for (const part of parts) {
                const match = part.match(/^([A-Z]+)/);
                if (match && match[1] !== colLetter) {
                    return false;
                }
            }
        }
        return true;
    }
    
    let countOnlyA = 0;
    cf.forEach((cfEntry, idx) => {
        if (cfEntry.ref && refOnlyReferencesColumn(cfEntry.ref, 'A')) {
            console.log(`CF[${idx}]: ${cfEntry.ref}`);
            countOnlyA++;
        }
    });
    console.log(`\nGesamt: ${countOnlyA} CF-Regeln die NUR auf Spalte A zeigen`);
    
    // Erste 5 CF anzeigen
    console.log('\n=== Erste 5 CF-Regeln ===');
    cf.slice(0, 5).forEach((cfEntry, idx) => {
        console.log(`CF[${idx}]: ref="${cfEntry.ref}"`);
    });
}

checkOriginalCF().catch(console.error);
