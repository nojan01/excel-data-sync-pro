const ExcelJS = require('exceljs');

async function check() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.worksheets[0];
    
    const cf = ws.conditionalFormattings || [];
    console.log('Anzahl CF:', cf.length);
    
    // Finde alle Spalten mit CF
    const colsWithCF = new Set();
    
    cf.forEach(entry => {
        const ref = entry.ref || '';
        // Extrahiere alle Spaltenbuchstaben
        const matches = ref.match(/[A-Z]+/g);
        if (matches) {
            matches.forEach(col => colsWithCF.add(col));
        }
    });
    
    // Sortiere nach Spaltenposition
    const sorted = Array.from(colsWithCF).sort((a, b) => {
        const numA = colLetterToNumber(a);
        const numB = colLetterToNumber(b);
        return numA - numB;
    });
    
    console.log('Erste 10 Spalten mit CF:', sorted.slice(0, 10).join(', '));
    console.log('Letzte 10 Spalten mit CF:', sorted.slice(-10).join(', '));
    
    // Zeige erste CF-Referenzen
    console.log('\nErste 5 CF refs:');
    cf.slice(0, 5).forEach((entry, i) => {
        console.log('  ' + (i+1) + ': ' + entry.ref);
    });
}

function colLetterToNumber(col) {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
        num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num;
}

check().catch(console.error);
