const ExcelJS = require('exceljs');

function colLetterToNumber(letters) {
    let num = 0;
    for (let i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

async function check() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('=== Suche CF-Regeln die Spalte BI (61) betreffen ===');
    console.log('Spalte BI = Nummer', colLetterToNumber('BI'));
    console.log('Spaltenanzahl:', ws.columnCount);
    console.log('');
    
    const cfBI = ws.conditionalFormattings.filter(cf => {
        const ref = cf.ref;
        const match = ref.match(/([A-Z]+)\d+:([A-Z]+)\d+/);
        if (match) {
            const endCol = colLetterToNumber(match[2]);
            return endCol >= 61;
        }
        return false;
    });
    
    console.log('Gefunden:', cfBI.length, 'CF-Regeln die bis Spalte BI oder weiter reichen');
    if (cfBI.length > 0) {
        console.log('Beispiele:');
        for (let i = 0; i < Math.min(5, cfBI.length); i++) {
            console.log('  ', cfBI[i].ref);
        }
    }
    
    console.log('');
    console.log('=== Alle einzigartigen CF-Referenzen ===');
    const uniqueRefs = [...new Set(ws.conditionalFormattings.map(cf => cf.ref))];
    uniqueRefs.forEach(ref => console.log('  ', ref));
    
    // Prüfe auch Zellen in Spalte BI
    console.log('');
    console.log('=== Zellen in Spalte BI (61) ===');
    const biCol = ws.getColumn(61);
    console.log('Header BI:', ws.getCell(1, 61).value);
    console.log('Zelle BI2:', ws.getCell(2, 61).value);
}

check().catch(console.error);
