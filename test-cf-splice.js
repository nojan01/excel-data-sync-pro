const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = wb.getWorksheet(1);
    console.log('Sheet:', ws.name);
    console.log('Anzahl CF Regeln:', ws.conditionalFormattings.length);
    
    // Prüfe welche Spalten in den CF Regeln referenziert werden
    const refCols = new Set();
    ws.conditionalFormattings.forEach(cf => {
        const ref = cf.ref;
        // Extrahiere Spaltenbuchstaben
        const match = ref.match(/([A-Z]+)/g);
        if (match) {
            match.forEach(col => refCols.add(col));
        }
    });
    
    console.log('\nReferenzierte Spalten in CF:', Array.from(refCols).sort().join(', '));
    
    // Zeige die ersten CF Regeln mit Spalte A oder B
    console.log('\n=== CF Regeln die Spalte A oder B referenzieren ===');
    let count = 0;
    ws.conditionalFormattings.forEach(cf => {
        if (cf.ref.startsWith('A') || cf.ref.startsWith('B')) {
            if (count < 5) {
                console.log('\nRef:', cf.ref);
                cf.rules.forEach(rule => {
                    console.log('  Type:', rule.type);
                    console.log('  Style Fill:', JSON.stringify(rule.style?.fill?.fgColor || rule.style?.fill?.bgColor));
                });
            }
            count++;
        }
    });
    console.log('\nGesamt CF Regeln mit A/B:', count);
    
    // Test: Was passiert mit spliceColumns?
    console.log('\n=== Test: spliceColumns und CF ===');
    console.log('CF Regeln vor spliceColumns:', ws.conditionalFormattings.length);
    
    // Zeige erste 3 CF refs VOR
    console.log('Erste 3 CF refs VOR:');
    ws.conditionalFormattings.slice(0, 3).forEach(cf => console.log('  ', cf.ref));
    
    // Lösche Spalte 1 (A)
    ws.spliceColumns(1, 1);
    
    console.log('\nCF Regeln nach spliceColumns:', ws.conditionalFormattings.length);
    
    // Zeige erste 3 CF refs NACH
    console.log('Erste 3 CF refs NACH:');
    ws.conditionalFormattings.slice(0, 3).forEach(cf => console.log('  ', cf.ref));
}

test().catch(console.error);
