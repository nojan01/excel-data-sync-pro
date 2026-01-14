const AdmZip = require('adm-zip');

async function compareStyleIndices() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    const zipOrig = new AdmZip(originalPath);
    const zipExp = new AdmZip(exportPath);
    
    const sheetOrig = zipOrig.readAsText('xl/worksheets/sheet1.xml');
    const sheetExp = zipExp.readAsText('xl/worksheets/sheet1.xml');
    
    // Style-Indices für Zeile 2-20 extrahieren
    console.log('=== STYLE INDICES VERGLEICH (Zeile 2-20) ===\n');
    
    function getRowStyles(sheetXml, rowNum) {
        const rowRegex = new RegExp(`<row[^>]*r="${rowNum}"[^>]*>([\\s\\S]*?)</row>`);
        const match = sheetXml.match(rowRegex);
        if (!match) return {};
        
        const styles = {};
        const cellRegex = /<c[^>]*r="([A-Z]+)\d+"[^>]*>/g;
        let cellMatch;
        while ((cellMatch = cellRegex.exec(match[1])) !== null) {
            const col = cellMatch[1];
            const styleMatch = cellMatch[0].match(/s="(\d+)"/);
            styles[col] = styleMatch ? styleMatch[1] : '0';
        }
        return styles;
    }
    
    for (let row = 2; row <= 10; row++) {
        const origStyles = getRowStyles(sheetOrig, row);
        const expStyles = getRowStyles(sheetExp, row);
        
        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
        const origStr = cols.map(c => `${c}:${origStyles[c] || '-'}`).join(' ');
        const expStr = cols.map(c => `${c}:${expStyles[c] || '-'}`).join(' ');
        
        console.log(`Zeile ${row}:`);
        console.log(`  Orig:   ${origStr}`);
        console.log(`  Export: ${expStr}`);
        
        // Prüfe ob sie übereinstimmen
        const match = cols.every(c => origStyles[c] === expStyles[c]);
        if (!match) {
            console.log(`  → UNTERSCHIEDLICH`);
        }
        console.log('');
    }
    
    // Zeige die CellXf Definitionen
    console.log('\n=== CELLXF DEFINITIONEN ===\n');
    
    const stylesOrig = zipOrig.readAsText('xl/styles.xml');
    const stylesExp = zipExp.readAsText('xl/styles.xml');
    
    // Parse CellXf
    function parseCellXfs(stylesXml) {
        const cellXfsMatch = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
        if (!cellXfsMatch) return [];
        
        const xfs = [];
        const xfRegex = /<xf([^>]*)(?:\/>|>[\s\S]*?<\/xf>)/g;
        let match;
        while ((match = xfRegex.exec(cellXfsMatch[1])) !== null) {
            xfs.push(match[1].trim());
        }
        return xfs;
    }
    
    const origXfs = parseCellXfs(stylesOrig);
    const expXfs = parseCellXfs(stylesExp);
    
    console.log('Original CellXfs:', origXfs.length);
    console.log('Export CellXfs:', expXfs.length);
    
    // Zeige erste 5 von jedem
    console.log('\nOriginal (erste 10):');
    origXfs.slice(0, 10).forEach((xf, i) => console.log(`  [${i}] ${xf.substring(0, 80)}`));
    
    console.log('\nExport (erste 10):');
    expXfs.slice(0, 10).forEach((xf, i) => console.log(`  [${i}] ${xf.substring(0, 80)}`));
}

compareStyleIndices().catch(console.error);
