const ExcelJS = require('exceljs');

async function check() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    const ws = workbook.worksheets[0];
    const cf = ws.conditionalFormattings;
    
    console.log('Spaltenanzahl:', ws.columnCount);
    
    // Letzte Spalte mit Daten
    let lastDataCol = 0;
    for (let c = 1; c <= ws.columnCount; c++) {
        if (ws.getCell(1, c).value) lastDataCol = c;
    }
    console.log('Letzte Spalte mit Header:', lastDataCol);
    console.log('Leere Spalten am Ende:', ws.columnCount - lastDataCol);
    
    // Header der letzten 5 Spalten
    console.log('\nLetzte Spalten:');
    for (let c = Math.max(1, ws.columnCount - 4); c <= ws.columnCount; c++) {
        const cell = ws.getCell(1, c);
        console.log('  Spalte', c, ':', cell.value || '(leer)');
    }
    
    console.log('\nAutoFilter:', ws.autoFilter || 'Nicht gesetzt');
    console.log('CF gesamt:', cf ? cf.length : 0);
    
    // CF die auf Spalten jenseits der Daten zeigen
    if (cf && cf.length > 0) {
        const colsInCF = new Set();
        cf.forEach(e => {
            if (e.ref) e.ref.match(/[A-Z]+/g)?.forEach(c => colsInCF.add(c));
        });
        
        const empty = [];
        for (const col of colsInCF) {
            let n = 0;
            for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
            if (n > lastDataCol) empty.push(col + '(' + n + ')');
        }
        if (empty.length) console.log('\n⚠️ CF auf leere Spalten:', empty.join(', '));
        else console.log('\n✓ Keine CF auf leere Spalten');
    }
}
check();
