const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');

async function analyzeHiddenColumns(filePath) {
    console.log('=== Analyse versteckter Spalten/Zeilen ===\n');
    console.log('Datei:', filePath, '\n');
    
    // 1. ExcelJS-Analyse
    console.log('--- ExcelJS Analyse ---');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const ws = workbook.getWorksheet(1);
    console.log('Sheet:', ws.name);
    console.log('columnCount:', ws.columnCount);
    
    // Prüfe versteckte Spalten laut ExcelJS
    console.log('\nVersteckte Spalten (ExcelJS):');
    let hiddenColsExcelJS = [];
    for (let i = 1; i <= Math.min(30, ws.columnCount); i++) {
        const col = ws.getColumn(i);
        if (col.hidden) {
            hiddenColsExcelJS.push(i);
            const letter = String.fromCharCode(64 + i);
            console.log(`  Spalte ${letter} (${i}): hidden=true, width=${col.width}`);
        }
    }
    if (hiddenColsExcelJS.length === 0) {
        console.log('  Keine versteckten Spalten gefunden');
    }
    
    // Prüfe versteckte Zeilen laut ExcelJS
    console.log('\nVersteckte Zeilen (ExcelJS, erste 100):');
    let hiddenRowsExcelJS = [];
    for (let i = 1; i <= Math.min(100, ws.rowCount); i++) {
        const row = ws.getRow(i);
        if (row.hidden) {
            hiddenRowsExcelJS.push(i);
        }
    }
    console.log(`  ${hiddenRowsExcelJS.length} versteckte Zeilen: ${hiddenRowsExcelJS.slice(0, 10).join(', ')}${hiddenRowsExcelJS.length > 10 ? '...' : ''}`);
    
    // 2. Direkte XML-Analyse
    console.log('\n--- Direkte XML Analyse ---');
    const zip = new AdmZip(filePath);
    
    // Finde das richtige Sheet-XML
    const workbookXml = zip.getEntry('xl/workbook.xml');
    if (workbookXml) {
        const wbContent = workbookXml.getData().toString('utf8');
        // Finde Sheet-IDs
        const sheetMatch = wbContent.match(/<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"[^>]*\/>/g);
        if (sheetMatch) {
            console.log('Sheets in Workbook:', sheetMatch.length);
        }
    }
    
    // Prüfe sheet1.xml
    const sheet1Entry = zip.getEntry('xl/worksheets/sheet1.xml');
    if (sheet1Entry) {
        const sheetXml = sheet1Entry.getData().toString('utf8');
        
        // Suche nach <cols> Element (Spaltenformatierungen)
        const colsMatch = sheetXml.match(/<cols>([\s\S]*?)<\/cols>/);
        if (colsMatch) {
            console.log('\n<cols> Element gefunden:');
            const colPattern = /<col\s+([^>]*)\/>/g;
            let colMatch;
            let hiddenColsXML = [];
            while ((colMatch = colPattern.exec(colsMatch[1])) !== null) {
                const attrs = colMatch[1];
                const minMatch = attrs.match(/min="(\d+)"/);
                const maxMatch = attrs.match(/max="(\d+)"/);
                const hiddenMatch = attrs.match(/hidden="(\d+|true)"/i);
                const widthMatch = attrs.match(/width="([^"]+)"/);
                
                if (minMatch && maxMatch) {
                    const min = parseInt(minMatch[1]);
                    const max = parseInt(maxMatch[1]);
                    const hidden = hiddenMatch ? (hiddenMatch[1] === '1' || hiddenMatch[1].toLowerCase() === 'true') : false;
                    const width = widthMatch ? parseFloat(widthMatch[1]) : null;
                    
                    if (hidden) {
                        for (let i = min; i <= max; i++) {
                            hiddenColsXML.push(i);
                        }
                        const letters = min === max 
                            ? String.fromCharCode(64 + min)
                            : `${String.fromCharCode(64 + min)}-${String.fromCharCode(64 + max)}`;
                        console.log(`  Spalte ${letters}: hidden=true, width=${width}`);
                    }
                }
            }
            
            if (hiddenColsXML.length === 0) {
                console.log('  Keine versteckten Spalten im XML');
            } else {
                console.log(`\nVersteckte Spalten laut XML: ${hiddenColsXML.join(', ')}`);
            }
            
            // Vergleiche
            console.log('\n--- Vergleich ---');
            console.log('ExcelJS findet:', hiddenColsExcelJS.length, 'versteckte Spalten');
            console.log('XML enthält:', hiddenColsXML.length, 'versteckte Spalten');
            
            if (hiddenColsExcelJS.length !== hiddenColsXML.length) {
                console.log('\n⚠️ DISKREPANZ! ExcelJS erkennt nicht alle versteckten Spalten!');
            }
        } else {
            console.log('Kein <cols> Element im Sheet-XML gefunden');
        }
        
        // Suche nach versteckten Zeilen im XML
        console.log('\nVersteckte Zeilen im XML:');
        const rowPattern = /<row\s+[^>]*hidden="(1|true)"[^>]*r="(\d+)"[^>]*>/gi;
        let rowMatch;
        let hiddenRowsXML = [];
        while ((rowMatch = rowPattern.exec(sheetXml)) !== null) {
            hiddenRowsXML.push(parseInt(rowMatch[2]));
        }
        // Auch anderes Format prüfen
        const rowPattern2 = /<row\s+[^>]*r="(\d+)"[^>]*hidden="(1|true)"[^>]*>/gi;
        while ((rowMatch = rowPattern2.exec(sheetXml)) !== null) {
            if (!hiddenRowsXML.includes(parseInt(rowMatch[1]))) {
                hiddenRowsXML.push(parseInt(rowMatch[1]));
            }
        }
        console.log(`  ${hiddenRowsXML.length} versteckte Zeilen im XML`);
        
    } else {
        console.log('sheet1.xml nicht gefunden');
    }
}

// Teste mit der Datei
const testFile = process.argv[2] || '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER Kopie 2.xlsx';
analyzeHiddenColumns(testFile).catch(console.error);
