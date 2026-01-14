const ExcelJS = require('exceljs');

// Vergleiche CF zwischen Original und Export
async function compareCF() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('=== VERGLEICH CF ORIGINAL vs EXPORT ===\n');
    
    // Original
    const wbOrig = new ExcelJS.Workbook();
    await wbOrig.xlsx.readFile(originalPath);
    const wsOrig = wbOrig.worksheets.find(w => w.name.includes('DEFENCE'));
    
    // Export
    const wbExp = new ExcelJS.Workbook();
    await wbExp.xlsx.readFile(exportPath);
    const wsExp = wbExp.worksheets.find(w => w.name.includes('DEFENCE'));
    
    console.log('Original CF-Anzahl:', wsOrig.conditionalFormattings?.length || 0);
    console.log('Export CF-Anzahl:', wsExp.conditionalFormattings?.length || 0);
    
    // Zeige die ersten 5 CFs zum Vergleich
    console.log('\n=== ERSTE 5 CF IM ORIGINAL ===');
    const cfOrig = wsOrig.conditionalFormattings || [];
    cfOrig.slice(0, 5).forEach((cf, i) => {
        console.log(`\nCF #${i+1}: ref="${cf.ref}"`);
        if (cf.rules) {
            cf.rules.forEach((r, ri) => {
                console.log(`  Rule ${ri+1}: type=${r.type}, formulae=${JSON.stringify(r.formulae)}`);
                if (r.style?.fill) {
                    console.log(`    Fill: ${JSON.stringify(r.style.fill)}`);
                }
            });
        }
    });
    
    console.log('\n=== ERSTE 5 CF IM EXPORT ===');
    const cfExp = wsExp.conditionalFormattings || [];
    cfExp.slice(0, 5).forEach((cf, i) => {
        console.log(`\nCF #${i+1}: ref="${cf.ref}"`);
        if (cf.rules) {
            cf.rules.forEach((r, ri) => {
                console.log(`  Rule ${ri+1}: type=${r.type}, formulae=${JSON.stringify(r.formulae)}`);
                if (r.style?.fill) {
                    console.log(`    Fill: ${JSON.stringify(r.style.fill)}`);
                }
            });
        }
    });
    
    // Suche nach CF die eine bestimmte Zelle betreffen
    console.log('\n=== SUCHE CF FÜR SPALTE B (ehemals C) ===');
    cfExp.filter(cf => cf.ref && cf.ref.includes('B')).slice(0, 3).forEach((cf, i) => {
        console.log(`\nCF: ref="${cf.ref}"`);
        if (cf.rules) {
            cf.rules.forEach((r, ri) => {
                console.log(`  Rule ${ri+1}: type=${r.type}`);
                console.log(`    formulae=${JSON.stringify(r.formulae)}`);
            });
        }
    });
}

compareCF().catch(console.error);
