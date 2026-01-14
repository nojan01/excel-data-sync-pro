const AdmZip = require('adm-zip');

function showColsElement(filePath) {
    console.log('=== <cols> Element Analyse ===\n');
    console.log('Datei:', filePath, '\n');
    
    const zip = new AdmZip(filePath);
    const sheet1Entry = zip.getEntry('xl/worksheets/sheet1.xml');
    
    if (sheet1Entry) {
        const sheetXml = sheet1Entry.getData().toString('utf8');
        
        // Suche nach <cols> Element
        const colsMatch = sheetXml.match(/<cols>([\s\S]*?)<\/cols>/);
        if (colsMatch) {
            console.log('Vollständiges <cols> Element:\n');
            
            // Parse einzelne <col> Elemente
            const colPattern = /<col\s+([^>]*)\/>/g;
            let colMatch;
            let colCount = 0;
            
            while ((colMatch = colPattern.exec(colsMatch[1])) !== null) {
                colCount++;
                const attrs = colMatch[1];
                
                // Extrahiere alle Attribute
                const minMatch = attrs.match(/min="(\d+)"/);
                const maxMatch = attrs.match(/max="(\d+)"/);
                const widthMatch = attrs.match(/width="([^"]+)"/);
                const hiddenMatch = attrs.match(/hidden="([^"]+)"/);
                const styleMatch = attrs.match(/style="([^"]+)"/);
                const customWidthMatch = attrs.match(/customWidth="([^"]+)"/);
                
                const min = minMatch ? minMatch[1] : '?';
                const max = maxMatch ? maxMatch[1] : '?';
                const width = widthMatch ? widthMatch[1] : 'nicht gesetzt';
                const hidden = hiddenMatch ? hiddenMatch[1] : 'false';
                const style = styleMatch ? styleMatch[1] : '-';
                const customWidth = customWidthMatch ? customWidthMatch[1] : '-';
                
                const minInt = parseInt(min);
                const maxInt = parseInt(max);
                const minLetter = minInt <= 26 ? String.fromCharCode(64 + minInt) : `A${String.fromCharCode(64 + minInt - 26)}`;
                const maxLetter = maxInt <= 26 ? String.fromCharCode(64 + maxInt) : `A${String.fromCharCode(64 + maxInt - 26)}`;
                const range = min === max ? minLetter : `${minLetter}-${maxLetter}`;
                
                console.log(`<col min="${min}" max="${max}"> (${range})`);
                console.log(`    width="${width}", hidden="${hidden}", customWidth="${customWidth}"`);
            }
            
            console.log(`\nGesamt: ${colCount} <col> Elemente`);
            
        } else {
            console.log('Kein <cols> Element gefunden!');
        }
    } else {
        console.log('sheet1.xml nicht gefunden');
    }
}

const testFile = process.argv[2] || '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie 2.xlsx';
showColsElement(testFile);
