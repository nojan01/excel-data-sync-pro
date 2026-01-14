const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = wb.getWorksheet(1);
    console.log('Sheet:', ws.name);
    
    console.log('\n=== Bedingte Formatierungen ===');
    const cf = ws.conditionalFormattings;
    console.log('Typ:', typeof cf, cf ? 'existiert' : 'null');
    
    if (cf && cf.model) {
        console.log('Model keys:', Object.keys(cf.model));
        const entries = Object.entries(cf.model);
        console.log('Anzahl Bereiche:', entries.length);
        entries.slice(0, 5).forEach(([ref, rules]) => {
            console.log('\nBereich:', ref);
            console.log('Regeln:', JSON.stringify(rules).substring(0, 300));
        });
    }
}

test().catch(console.error);
