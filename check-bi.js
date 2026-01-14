const ExcelJS = require('exceljs');

async function check() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = workbook.worksheets[0];
    const cf = ws.conditionalFormattings;
    
    console.log('Suche nach BI in CF-Referenzen...');
    
    let biRefs = [];
    if (cf) {
        cf.forEach((e, i) => {
            if (e.ref && e.ref.includes('BI')) {
                biRefs.push({idx: i, ref: e.ref});
            }
        });
    }
    
    console.log('CF mit BI:', biRefs.length);
    if (biRefs.length > 0) {
        console.log('Beispiele:');
        biRefs.slice(0, 10).forEach(r => console.log('  [' + r.idx + ']', r.ref));
    }
    
    // Prüfe auch Spalte BI direkt
    const biCol = 61;
    console.log('\nSpalte BI (61) existiert?', ws.columnCount >= biCol);
    if (ws.columnCount >= biCol) {
        console.log('BI1 Wert:', ws.getCell(1, biCol).value);
    }
}
check();
