const AdmZip = require('adm-zip');

async function debugRowMapping() {
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    // Export
    const zipExp = new AdmZip(exportPath);
    const sheetExp = zipExp.readAsText('xl/worksheets/sheet1.xml');
    
    // Extrahiere die ersten 20 Zeilen mit Werten und Style-Indizes
    console.log('=== EXPORT DATEI ANALYSE ===\n');
    
    function extractRowData(sheetXml, rowNum) {
        const rowRegex = new RegExp(`<row[^>]*r="${rowNum}"[^>]*>([\\s\\S]*?)<\\/row>`);
        const rowMatch = sheetXml.match(rowRegex);
        if (!rowMatch) return null;
        
        const rowXml = rowMatch[0];
        const cells = [];
        
        // Extrahiere alle Zellen
        const cellRegex = /<c[^>]*r="([A-Z]+)\d+"[^>]*(?:s="(\d+)")?[^>]*>(?:<v>([^<]*)<\/v>)?/g;
        let match;
        while ((match = cellRegex.exec(rowXml)) !== null) {
            const col = match[1];
            const styleMatch = rowXml.match(new RegExp(`r="${col}${rowNum}"[^>]*s="(\\d+)"`));
            const style = styleMatch ? styleMatch[1] : '0';
            cells.push({
                col,
                style,
                value: match[3] || ''
            });
        }
        
        return cells.slice(0, 5);  // Erste 5 Spalten
    }
    
    // Zeige Zeilen 2-15
    for (let row = 2; row <= 15; row++) {
        const cells = extractRowData(sheetExp, row);
        if (cells) {
            console.log(`Zeile ${row}: ` + cells.map(c => `${c.col}:s${c.style}`).join(', '));
        }
    }
    
    // Zähle verschiedene Style-Indizes
    console.log('\n=== STYLE-INDEX VERTEILUNG ===\n');
    
    const allStyleIndices = sheetExp.match(/s="(\d+)"/g) || [];
    const styleCounts = {};
    allStyleIndices.forEach(s => {
        const idx = s.match(/s="(\d+)"/)[1];
        styleCounts[idx] = (styleCounts[idx] || 0) + 1;
    });
    
    console.log('Style-Index Häufigkeiten:');
    Object.entries(styleCounts)
        .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
        .forEach(([idx, count]) => {
            console.log(`  s="${idx}": ${count} Zellen`);
        });
}

debugRowMapping().catch(console.error);
