const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = wb.getWorksheet(1);
    console.log('Sheet:', ws.name);
    
    const cf = ws.conditionalFormattings;
    console.log('CF existiert:', !!cf);
    console.log('CF model:', cf.model ? 'ja' : 'nein');
    
    if (cf.model) {
        const entries = Object.entries(cf.model);
        console.log('Anzahl CF Bereiche:', entries.length);
        
        // Zeige die ersten 5
        entries.slice(0, 5).forEach(([ref, rules], i) => {
            console.log('\n--- CF Bereich', i+1, '---');
            console.log('Referenz:', ref);
            console.log('Regeln Anzahl:', rules.length);
            rules.forEach((rule, j) => {
                console.log('  Regel', j+1 + ':');
                console.log('    type:', rule.type);
                console.log('    priority:', rule.priority);
                if (rule.style && rule.style.fill) {
                    console.log('    fill:', JSON.stringify(rule.style.fill));
                }
                if (rule.formulae) {
                    console.log('    formulae:', JSON.stringify(rule.formulae).substring(0, 100));
                }
            });
        });
    }
}

test().catch(console.error);
