const XlsxPopulate = require('xlsx-populate');
const JSZip = require('jszip');
const fs = require('fs');

async function check() {
    const files = fs.readdirSync('/Users/nojan/Desktop').filter(f => f.includes('export') || f.includes('Export'));
    const exportFile = files.find(f => f.endsWith('.xlsx'));
    
    const data = fs.readFileSync('/Users/nojan/Desktop/' + exportFile);
    const zip = await JSZip.loadAsync(data);
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml').async('string');
    
    // Zähle Spalten in cols
    const colsMatch = sheet1.match(/<cols>([\s\S]*?)<\/cols>/);
    if (colsMatch) {
        const colCount = (colsMatch[1].match(/<col /g) || []).length;
        console.log('Anzahl <col> Definitionen:', colCount);
        
        // Finde die höchste max-Spalte
        const maxVals = [...colsMatch[1].matchAll(/max="(\d+)"/g)].map(m => parseInt(m[1]));
        console.log('Höchste max-Spalte:', Math.max(...maxVals));
    }
    
    // Zähle Spalten in Zeile 1
    const row1 = sheet1.match(/<row r="1"[^>]*>([\s\S]*?)<\/row>/);
    if (row1) {
        const cellRefs = [...row1[1].matchAll(/r="([A-Z]+)1"/g)].map(m => m[1]);
        console.log('\nZellen in Zeile 1:', cellRefs.length);
        console.log('Letzte Zelle:', cellRefs[cellRefs.length - 1]);
    }
    
    // Prüfe spans-Attribut
    const spans = sheet1.match(/spans="([^"]+)"/);
    if (spans) {
        console.log('\nSpans:', spans[1]);
    }
}

check().catch(console.error);
