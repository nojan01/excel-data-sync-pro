const ExcelJS = require('exceljs');

async function checkCF() {
    const filePath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];
    
    console.log('=== BEDINGTE FORMATIERUNG DIE BI ENTHÄLT ===\n');
    
    const cfs = sheet.conditionalFormattings;
    let cfWithBI = 0;
    
    cfs.forEach((cf, index) => {
        const refs = cf.ref || '';
        // Prüfe ob BI in der Referenz vorkommt
        if (refs.includes('BI') || refs.includes(':BI') || refs.match(/B[I-Z]/)) {
            cfWithBI++;
            console.log(`CF #${index}: ref="${refs}"`);
            if (cf.rules && cf.rules.length > 0) {
                console.log(`  Rules: ${cf.rules.length}`);
                cf.rules.forEach((rule, ri) => {
                    console.log(`    Rule ${ri}: type=${rule.type}, style=${JSON.stringify(rule.style)}`);
                });
            }
        }
    });
    
    console.log(`\nAnzahl CF-Regeln mit BI: ${cfWithBI}`);
    console.log(`Gesamt CF-Regeln: ${cfs.length}`);
    
    // Zeige auch alle CF die Spalten >= 60 referenzieren
    console.log('\n=== CF MIT SPALTEN >= BH (60) ===\n');
    const highColPattern = /B[H-Z]|C[A-Z]/;
    cfs.forEach((cf, index) => {
        const refs = cf.ref || '';
        if (highColPattern.test(refs)) {
            console.log(`CF #${index}: ref="${refs}"`);
        }
    });
}

checkCF().catch(console.error);
