// Finde alle Zellen mit rosa Farbe FFE0E0 im Original
const ExcelJS = require('exceljs');

async function findRosa() {
    const originalPath = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(originalPath);
    const ws = wb.worksheets[0];
    
    console.log('=== SUCHE ZELLEN MIT FFE0E0 (rosa) ===');
    
    // Durchsuche alle Zellen
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const fill = cell.fill;
            if (fill && fill.type === 'pattern' && fill.fgColor?.argb) {
                const argb = fill.fgColor.argb;
                if (argb.includes('E0E0') || argb.includes('e0e0')) {
                    console.log('  ' + cell.address + ': Fill=' + argb + ', Wert=' + cell.value);
                }
            }
        });
    });
    
    // Zeige auch alle Fills in styles.xml
    console.log('\n=== FILLS aus styles.xml ===');
    console.log('Index 17 ist #FFE0E0 laut extractFillsFromXLSX');
    
    // Prüfe welche Zellen Style-ID haben die auf Fill 17 mappt
    const AdmZip = require('adm-zip');
    const zip = new AdmZip(originalPath);
    
    const stylesEntry = zip.getEntry('xl/styles.xml');
    const stylesXml = stylesEntry.getData().toString('utf8');
    
    // Finde cellXfs - welche Style IDs haben fillId > 0?
    const cellXfsMatch = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
    if (cellXfsMatch) {
        const xfPattern = /<xf[^>]*fillId="(\d+)"[^>]*>/g;
        let xfMatch;
        let styleIdx = 0;
        console.log('\nStyle IDs mit Fill:');
        while ((xfMatch = xfPattern.exec(cellXfsMatch[1])) !== null) {
            const fillId = parseInt(xfMatch[1]);
            if (fillId === 17) {
                console.log('  Style ID ' + styleIdx + ' hat fillId=17 (FFE0E0)');
            }
            styleIdx++;
        }
    }
    
    // Welche Zellen haben Style ID 42 (oder was auch immer auf fillId 17 mappt)?
    console.log('\n=== SUCHE ZELLEN MIT fillId=17 ===');
    
    // Sheet XML lesen
    const sheetEntry = zip.getEntry('xl/worksheets/sheet1.xml');
    const sheetXml = sheetEntry.getData().toString('utf8');
    
    // Finde Zellen mit s="42" oder ähnlich
    const cellPattern = /<c r="([A-Z]+\d+)"[^>]*s="42"[^>]*>/g;
    let cellMatch;
    while ((cellMatch = cellPattern.exec(sheetXml)) !== null) {
        console.log('  Zelle ' + cellMatch[1] + ' hat Style ID 42');
    }
}

findRosa().catch(e => console.error(e));
