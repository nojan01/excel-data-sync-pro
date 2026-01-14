const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = wb.getWorksheet(1);
    console.log('Sheet:', ws.name);
    
    const cf = ws.conditionalFormattings;
    console.log('CF existiert:', !!cf);
    console.log('CF Typ:', typeof cf);
    console.log('CF Konstruktor:', cf.constructor.name);
    console.log('CF Keys:', Object.keys(cf));
    console.log('CF eigene Properties:', Object.getOwnPropertyNames(cf));
    
    // Prüfe forEach
    if (typeof cf.forEach === 'function') {
        console.log('\nforEach verfügbar, iteriere...');
        let count = 0;
        cf.forEach((rules, ref) => {
            if (count < 5) {
                console.log('\n  Bereich:', ref);
                console.log('  Rules Typ:', typeof rules);
                if (Array.isArray(rules)) {
                    console.log('  Anzahl Regeln:', rules.length);
                    rules.slice(0, 2).forEach((rule, j) => {
                        console.log('    Regel', j+1, ':', JSON.stringify(rule).substring(0, 200));
                    });
                } else {
                    console.log('  Rules:', JSON.stringify(rules).substring(0, 200));
                }
            }
            count++;
        });
        console.log('\nGesamt CF Bereiche:', count);
    }
}

test().catch(console.error);
