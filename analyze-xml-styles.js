const AdmZip = require('adm-zip');

async function analyzeXML() {
    const originalPath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    // Original
    const zipOrig = new AdmZip(originalPath);
    const sheetOrig = zipOrig.readAsText('xl/worksheets/sheet1.xml');
    const stylesOrig = zipOrig.readAsText('xl/styles.xml');
    
    // Export
    const zipExp = new AdmZip(exportPath);
    const sheetExp = zipExp.readAsText('xl/worksheets/sheet1.xml');
    const stylesExp = zipExp.readAsText('xl/styles.xml');
    
    console.log('=== STYLES.XML VERGLEICH ===\n');
    
    // Parse fills
    const fillsOrigMatch = stylesOrig.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
    const fillsExpMatch = stylesExp.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
    
    console.log('Original fills count:', (fillsOrigMatch?.[1]?.match(/<fill>/g) || []).length);
    console.log('Export fills count:', (fillsExpMatch?.[1]?.match(/<fill>/g) || []).length);
    
    // Zeige einige Zeilen aus dem Worksheet
    console.log('\n=== WORKSHEET ROWS ===\n');
    
    // Extrahiere Zeile 2-5 aus Original
    const rowsOrig = sheetOrig.match(/<row[^>]*r="([2-9]|[1-2][0-9])"[^>]*>[\s\S]*?<\/row>/g) || [];
    const rowsExp = sheetExp.match(/<row[^>]*r="([2-9]|[1-2][0-9])"[^>]*>[\s\S]*?<\/row>/g) || [];
    
    console.log('Zeilen 2-29 im Original:', rowsOrig.length);
    console.log('Zeilen 2-29 im Export:', rowsExp.length);
    
    // Zeige Style-Index (s="X") für einige Zellen
    console.log('\n=== STYLE INDICES VERGLEICH ===\n');
    
    // Funktion um s-Attribute aus einer Zeile zu extrahieren
    function extractStyles(rowXml) {
        const cells = rowXml.match(/<c[^>]*>/g) || [];
        return cells.map(c => {
            const sMatch = c.match(/s="(\d+)"/);
            const rMatch = c.match(/r="([A-Z]+\d+)"/);
            return { cell: rMatch?.[1], style: sMatch?.[1] || '0' };
        });
    }
    
    // Vergleiche die ersten 5 Zeilen
    for (let i = 0; i < Math.min(5, rowsOrig.length); i++) {
        const rowNumOrig = rowsOrig[i].match(/r="(\d+)"/)?.[1];
        const rowNumExp = rowsExp[i]?.match(/r="(\d+)"/)?.[1];
        
        console.log(`\n--- Original Zeile ${rowNumOrig} ---`);
        const stylesO = extractStyles(rowsOrig[i]);
        console.log(stylesO.slice(0, 8).map(s => `${s.cell}:s${s.style}`).join(', '));
        
        if (rowsExp[i]) {
            console.log(`--- Export Zeile ${rowNumExp} ---`);
            const stylesE = extractStyles(rowsExp[i]);
            console.log(stylesE.slice(0, 8).map(s => `${s.cell}:s${s.style}`).join(', '));
        }
    }
    
    // Zeige styles.xml cellXfs (die eigentlichen Style-Definitionen)
    console.log('\n=== CELLXFS (Style-Definitionen) ===\n');
    
    const cellXfsOrig = stylesOrig.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/)?.[1] || '';
    const cellXfsExp = stylesExp.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/)?.[1] || '';
    
    const xfCountOrig = (cellXfsOrig.match(/<xf/g) || []).length;
    const xfCountExp = (cellXfsExp.match(/<xf/g) || []).length;
    
    console.log('Original cellXf count:', xfCountOrig);
    console.log('Export cellXf count:', xfCountExp);
    
    // Prüfe ob Zeile 2 die gleichen Styles hat wie in Original
    console.log('\n=== DETAILLIERTER ZEILEN-VERGLEICH ===\n');
    
    // Original Zeile 2
    const row2Orig = sheetOrig.match(/<row[^>]*r="2"[^>]*>[\s\S]*?<\/row>/)?.[0] || '';
    const row2Exp = sheetExp.match(/<row[^>]*r="2"[^>]*>[\s\S]*?<\/row>/)?.[0] || '';
    
    console.log('Original Zeile 2 (erste 500 chars):');
    console.log(row2Orig.substring(0, 500));
    console.log('\nExport Zeile 2 (erste 500 chars):');
    console.log(row2Exp.substring(0, 500));
}

analyzeXML().catch(console.error);
