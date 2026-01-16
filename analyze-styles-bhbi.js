const ExcelJS = require('exceljs');

async function analyzeStyles() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('Lade Excel-Datei...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    
    // Spalten BH (60) und BI (61)
    const colBH = 60;
    const colBI = 61;
    
    console.log('=== Styles in Spalte BH (erste 20 Zeilen) ===');
    for (let row = 1; row <= 20; row++) {
        const cell = worksheet.getCell(row, colBH);
        const fill = cell.style?.fill;
        const hasFill = fill && fill.type === 'pattern' && fill.pattern !== 'none';
        console.log(`BH${row}: value="${cell.value}", fill=${hasFill ? JSON.stringify(fill) : 'none'}`);
    }
    
    console.log('\n=== Styles in Spalte BI (erste 20 Zeilen) ===');
    for (let row = 1; row <= 20; row++) {
        const cell = worksheet.getCell(row, colBI);
        const fill = cell.style?.fill;
        const hasFill = fill && fill.type === 'pattern' && fill.pattern !== 'none';
        console.log(`BI${row}: value="${cell.value}", fill=${hasFill ? JSON.stringify(fill) : 'none'}`);
    }
    
    // Zähle wie viele Zellen in BH/BI eine Füllfarbe haben
    console.log('\n=== Statistik: Zellen mit Füllung in BH/BI ===');
    let bhFillCount = 0;
    let biFillCount = 0;
    
    for (let row = 1; row <= worksheet.rowCount; row++) {
        const cellBH = worksheet.getCell(row, colBH);
        const cellBI = worksheet.getCell(row, colBI);
        
        const fillBH = cellBH.style?.fill;
        const fillBI = cellBI.style?.fill;
        
        if (fillBH && fillBH.type === 'pattern' && fillBH.pattern !== 'none') {
            bhFillCount++;
        }
        if (fillBI && fillBI.type === 'pattern' && fillBI.pattern !== 'none') {
            biFillCount++;
        }
    }
    
    console.log(`BH: ${bhFillCount} von ${worksheet.rowCount} Zellen haben Füllung`);
    console.log(`BI: ${biFillCount} von ${worksheet.rowCount} Zellen haben Füllung`);
    
    // Finde Beispiele für Zellen mit Füllung
    if (bhFillCount > 0 || biFillCount > 0) {
        console.log('\n=== Beispiele für Zellen mit Füllung ===');
        let found = 0;
        for (let row = 1; row <= worksheet.rowCount && found < 5; row++) {
            const cellBH = worksheet.getCell(row, colBH);
            const cellBI = worksheet.getCell(row, colBI);
            
            const fillBH = cellBH.style?.fill;
            const fillBI = cellBI.style?.fill;
            
            if (fillBH && fillBH.type === 'pattern' && fillBH.pattern !== 'none') {
                console.log(`BH${row}: ${JSON.stringify(fillBH)}`);
                found++;
            }
            if (fillBI && fillBI.type === 'pattern' && fillBI.pattern !== 'none') {
                console.log(`BI${row}: ${JSON.stringify(fillBI)}`);
                found++;
            }
        }
    }
}

analyzeStyles().catch(console.error);
