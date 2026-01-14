// Debug: Analysiere bedingte Formatierungen in der echten Datei
const ExcelJS = require('exceljs');

const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';

async function analyzeCF() {
    console.log('=== Analysiere bedingte Formatierungen ===');
    console.log('Datei:', filePath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    for (const worksheet of workbook.worksheets) {
        console.log('\n========================================');
        console.log('Sheet:', worksheet.name);
        console.log('Spalten:', worksheet.columnCount);
        console.log('Zeilen:', worksheet.rowCount);
        
        const cf = worksheet.conditionalFormattings;
        if (!cf || cf.length === 0) {
            console.log('Keine bedingten Formatierungen');
            continue;
        }
        
        console.log('Bedingte Formatierungen:', cf.length);
        
        // Zeige die ersten 10 CF-Regeln
        cf.slice(0, 10).forEach((cfEntry, idx) => {
            console.log('\n[CF ' + (idx + 1) + ']');
            console.log('  ref:', cfEntry.ref);
            if (cfEntry.rules && cfEntry.rules.length > 0) {
                cfEntry.rules.forEach((rule, rIdx) => {
                    console.log('  rule[' + rIdx + '].type:', rule.type);
                    if (rule.formulae) {
                        console.log('  rule[' + rIdx + '].formulae:', rule.formulae);
                    }
                });
            }
        });
        
        if (cf.length > 10) {
            console.log('\n... und ' + (cf.length - 10) + ' weitere CF-Regeln');
        }
        
        // Analysiere welche Spalten von CF betroffen sind
        const colsInCF = new Set();
        cf.forEach(cfEntry => {
            if (cfEntry.ref) {
                const matches = cfEntry.ref.match(/([A-Z]+)/g);
                if (matches) {
                    matches.forEach(col => colsInCF.add(col));
                }
            }
        });
        console.log('\nSpalten in CF-Referenzen:', Array.from(colsInCF).sort().join(', '));
    }
}

analyzeCF().catch(console.error);
