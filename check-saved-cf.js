// Prüfe die CF in der gespeicherten Datei
const ExcelJS = require('exceljs');

const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';

async function checkSavedCF() {
    console.log('=== Prüfe bedingte Formatierungen in der gespeicherten Datei ===');
    console.log('Datei:', filePath);
    console.log('');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    console.log('Sheet:', worksheet.name);
    console.log('Spalten:', worksheet.columnCount);
    
    const cf = worksheet.conditionalFormattings;
    console.log('CF-Regeln:', cf?.length || 0);
    
    if (cf && cf.length > 0) {
        console.log('\nErste 5 CF-Referenzen:');
        cf.slice(0, 5).forEach((cfEntry, idx) => {
            console.log('  CF ' + (idx + 1) + ':', cfEntry.ref);
        });
        
        // Analysiere Spaltenbuchstaben
        const colsInCF = new Set();
        cf.forEach(cfEntry => {
            if (cfEntry.ref) {
                const matches = cfEntry.ref.match(/([A-Z]+)/g);
                if (matches) {
                    matches.forEach(col => colsInCF.add(col));
                }
            }
        });
        console.log('\nSpalten in CF:', Array.from(colsInCF).sort().join(', '));
    }
}

checkSavedCF().catch(console.error);
