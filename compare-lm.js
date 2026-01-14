const ExcelJS = require('exceljs');

// Vergleiche Spalten L und M zwischen Original und Export
async function compareLM() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    console.log('=== VERGLEICH SPALTEN L/M ===\n');
    
    // Original
    const wbOrig = new ExcelJS.Workbook();
    await wbOrig.xlsx.readFile(originalPath);
    const wsOrig = wbOrig.worksheets.find(w => w.name.includes('DEFENCE'));
    
    // Export
    const wbExp = new ExcelJS.Workbook();
    await wbExp.xlsx.readFile(exportPath);
    const wsExp = wbExp.worksheets.find(w => w.name.includes('DEFENCE'));
    
    // Im Original: Spalte A wurde gelöscht
    // Also: Original Spalte M = Export Spalte L, Original Spalte N = Export Spalte M
    
    console.log('HINWEIS: Spalte A wurde gelöscht');
    console.log('Original Spalte M (13) -> Export Spalte L (12)');
    console.log('Original Spalte N (14) -> Export Spalte M (13)\n');
    
    // Zeige Zeilen 2-20 zum Vergleich
    console.log('=== ZEILEN 2-20 ===\n');
    
    for (let row = 2; row <= 20; row++) {
        // Original: Spalte M = 13, N = 14
        const origM = wsOrig.getCell(row, 13);
        const origN = wsOrig.getCell(row, 14);
        
        // Export: Spalte L = 12, M = 13
        const expL = wsExp.getCell(row, 12);
        const expM = wsExp.getCell(row, 13);
        
        // Fills extrahieren
        const getArgb = (cell) => {
            if (!cell.fill || cell.fill.type !== 'pattern') return 'none';
            return cell.fill.fgColor?.argb || cell.fill.bgColor?.argb || 'none';
        };
        
        const origMFill = getArgb(origM);
        const origNFill = getArgb(origN);
        const expLFill = getArgb(expL);
        const expMFill = getArgb(expM);
        
        // Nur zeigen wenn unterschiedlich
        const mDiff = origMFill !== expLFill;
        const nDiff = origNFill !== expMFill;
        
        if (mDiff || nDiff) {
            console.log(`Zeile ${row}:`);
            if (mDiff) {
                console.log(`  M->L: Orig=${origMFill} vs Exp=${expLFill} ${mDiff ? '❌' : '✓'}`);
            }
            if (nDiff) {
                console.log(`  N->M: Orig=${origNFill} vs Exp=${expMFill} ${nDiff ? '❌' : '✓'}`);
            }
        }
    }
    
    console.log('\n=== ERSTE 50 UNTERSCHIEDE SPALTE M->L ===\n');
    let diffCount = 0;
    for (let row = 2; row <= 2404 && diffCount < 50; row++) {
        const origM = wsOrig.getCell(row, 13);
        const expL = wsExp.getCell(row, 12);
        
        const getArgb = (cell) => {
            if (!cell.fill || cell.fill.type !== 'pattern') return 'none';
            return cell.fill.fgColor?.argb || cell.fill.bgColor?.argb || 'none';
        };
        
        const origFill = getArgb(origM);
        const expFill = getArgb(expL);
        
        if (origFill !== expFill) {
            console.log(`Zeile ${row}: Orig M=${origFill}, Exp L=${expFill}, Value="${origM.value}"`);
            diffCount++;
        }
    }
    
    console.log(`\n${diffCount} Unterschiede gefunden (max 50 angezeigt)`);
}

compareLM().catch(console.error);
