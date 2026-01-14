const ExcelJS = require('exceljs');

async function test() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws = wb.getWorksheet(1);
    
    console.log('=== Export-Datei Spaltenanalyse ===');
    console.log('Spaltenanzahl:', ws.columnCount);
    console.log('');
    
    // Letzte 5 Spalten
    for (let col = ws.columnCount - 4; col <= ws.columnCount; col++) {
        const colObj = ws.getColumn(col);
        const letter = colObj.letter;
        const header = ws.getCell(1, col).value;
        console.log(`Spalte ${col} (${letter}): "${header}"`);
    }
    
    console.log('');
    console.log('=== CF-Referenzen die bis zur letzten Spalte reichen ===');
    
    const cf = ws.conditionalFormattings;
    const lastColLetter = ws.getColumn(ws.columnCount).letter; // BH
    
    console.log('Letzte Spalte:', lastColLetter);
    console.log('');
    
    // Finde CFs die BH enthalten
    const bhRefs = cf.filter(c => c.ref && c.ref.includes(lastColLetter));
    console.log(`CF mit "${lastColLetter}":`, bhRefs.length);
    bhRefs.slice(0, 5).forEach(c => console.log('  -', c.ref));
    
    // Finde CFs die über die letzte Spalte hinausgehen (BI, BJ, etc.)
    console.log('');
    console.log('=== CF-Referenzen die ÜBER die letzte Spalte hinausgehen ===');
    const beyondRefs = cf.filter(c => {
        if (!c.ref) return false;
        // Suche nach BI, BJ, BK etc.
        return /\bBI\d|\bBJ\d|\bBK\d/.test(c.ref);
    });
    console.log('Gefunden:', beyondRefs.length);
    beyondRefs.slice(0, 10).forEach(c => console.log('  -', c.ref));
    
    // Prüfe die AY-Referenzen nochmal
    console.log('');
    console.log('=== AY1:AY2404 Analyse ===');
    console.log('AY = Spalte 51');
    console.log('Nach Löschung von A sollte alte AZ (52) jetzt AY (51) sein');
    
    const cell51 = ws.getCell(1, 51);
    console.log('Header Spalte 51 (AY):', cell51.value);
    
    // Was war vorher in Spalte 52 (AZ)?
    console.log('');
    console.log('Im Original:');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile('/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const ws2 = wb2.getWorksheet(1);
    
    const cell52orig = ws2.getCell(1, 52);
    console.log('Original Header Spalte 52 (AZ):', cell52orig.value);
    
    // Vergleich
    console.log('');
    console.log('Vergleich:');
    console.log('Export Spalte 51 (AY):', cell51.value);
    console.log('Original Spalte 52 (AZ):', cell52orig.value);
    console.log('Sind gleich:', cell51.value === cell52orig.value);
}

test().catch(console.error);
