const ExcelJS = require('exceljs');

async function analyzeZebraPattern() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Excel-Datei...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    
    // Prüfe verschiedene Spalten auf Zebra-Muster
    const columnsToCheck = [1, 10, 30, 60, 61]; // A, J, AD, BH, BI
    
    for (const colNum of columnsToCheck) {
        const colLetter = worksheet.getColumn(colNum).letter;
        console.log(`\n=== Spalte ${colLetter} (${colNum}) - Zeilen 2-10 ===`);
        
        for (let row = 2; row <= 10; row++) {
            const cell = worksheet.getCell(row, colNum);
            const fill = cell.style?.fill;
            
            let fillInfo = 'KEINE';
            if (fill) {
                if (fill.type === 'pattern' && fill.pattern !== 'none') {
                    fillInfo = `pattern=${fill.pattern}`;
                    if (fill.fgColor) {
                        if (fill.fgColor.argb) fillInfo += `, fgColor=${fill.fgColor.argb}`;
                        if (fill.fgColor.theme !== undefined) fillInfo += `, theme=${fill.fgColor.theme}, tint=${fill.fgColor.tint}`;
                        if (fill.fgColor.indexed !== undefined) fillInfo += `, indexed=${fill.fgColor.indexed}`;
                    }
                    if (fill.bgColor) {
                        if (fill.bgColor.argb) fillInfo += `, bgColor=${fill.bgColor.argb}`;
                        if (fill.bgColor.indexed !== undefined) fillInfo += `, bgIndexed=${fill.bgColor.indexed}`;
                    }
                }
            }
            
            console.log(`${colLetter}${row}: ${fillInfo}`);
        }
    }
    
    // Prüfe ob es eine Excel-Tabelle (ListObject) gibt
    console.log('\n=== Excel Tabellen ===');
    if (worksheet.tables && Object.keys(worksheet.tables).length > 0) {
        console.log(`Tabellen gefunden: ${Object.keys(worksheet.tables).length}`);
        for (const [name, table] of Object.entries(worksheet.tables)) {
            console.log(`  Tabelle: ${name}, Ref: ${table.ref || table.tableRef}`);
        }
    } else {
        console.log('Keine Excel-Tabellen gefunden');
    }
    
    // Zeige Zeilenfüllungen für verschiedene Zeilen
    console.log('\n=== Vergleich Zeile 2 vs Zeile 3 (Spalten A-E) ===');
    for (let col = 1; col <= 5; col++) {
        const cell2 = worksheet.getCell(2, col);
        const cell3 = worksheet.getCell(3, col);
        
        const fill2 = cell2.style?.fill;
        const fill3 = cell3.style?.fill;
        
        console.log(`Spalte ${col}:`);
        console.log(`  Zeile 2: ${JSON.stringify(fill2)}`);
        console.log(`  Zeile 3: ${JSON.stringify(fill3)}`);
    }
}

analyzeZebraPattern().catch(console.error);
