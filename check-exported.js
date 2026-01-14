const ExcelJS = require('exceljs');

async function checkExportedFile() {
    // Prüfe die EXPORTIERTE Datei
    const exportedPath = '/Users/nojan/Desktop/test-export.xlsx';
    
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(exportedPath);
        
        const worksheet = workbook.worksheets[0];
        const cf = worksheet.conditionalFormattings;
        
        console.log('=== Analyse der EXPORTIERTEN Datei ===');
        console.log('Worksheet columnCount:', worksheet.columnCount);
        
        // Finde die letzte Spalte mit Daten in Zeile 1
        let lastColWithData = 0;
        for (let col = 1; col <= worksheet.columnCount; col++) {
            const cell = worksheet.getCell(1, col);
            if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
                lastColWithData = col;
            }
        }
        
        console.log('Letzte Spalte mit Daten in Zeile 1:', lastColWithData);
        console.log('Differenz (leere Spalten am Ende):', worksheet.columnCount - lastColWithData);
        
        // Zeige die letzten 5 Spalten
        console.log('\nLetzte 5 Spalten (Header):');
        for (let col = Math.max(1, worksheet.columnCount - 4); col <= worksheet.columnCount; col++) {
            const cell = worksheet.getCell(1, col);
            const colLetter = cell.address.replace(/\d+/, '');
            console.log('  ' + colLetter + ': "' + (cell.value || '') + '"');
        }
        
        // Prüfe CF auf die letzte Spalte
        if (cf && cf.length > 0) {
            console.log('\nCF Regeln gesamt:', cf.length);
            
            // Finde CF die auf Spalten nach lastColWithData zeigen
            const colsInCF = new Set();
            cf.forEach(cfEntry => {
                if (cfEntry.ref) {
                    const matches = cfEntry.ref.match(/[A-Z]+/g);
                    if (matches) {
                        matches.forEach(col => colsInCF.add(col));
                    }
                }
            });
            
            console.log('Alle Spalten in CF-Referenzen:', Array.from(colsInCF).sort().join(', '));
            
            // Welche Spalten in CF sind jenseits der Daten?
            const emptyColsWithCF = [];
            for (const colLetter of colsInCF) {
                let colNum = 0;
                for (let i = 0; i < colLetter.length; i++) {
                    colNum = colNum * 26 + (colLetter.charCodeAt(i) - 64);
                }
                if (colNum > lastColWithData) {
                    emptyColsWithCF.push(colLetter);
                }
            }
            
            if (emptyColsWithCF.length > 0) {
                console.log('\n⚠️ CF-Referenzen auf Spalten OHNE Daten:', emptyColsWithCF.join(', '));
            } else {
                console.log('\n✓ Keine CF-Referenzen auf leere Spalten');
            }
        }
        
        // Prüfe AutoFilter
        console.log('\nAutoFilter:', worksheet.autoFilter || 'Nicht gesetzt');
        
    } catch (err) {
        console.log('Fehler:', err.message);
        console.log('Bitte exportiere zuerst eine Datei nach /Users/nojan/Desktop/test-export.xlsx');
    }
}

checkExportedFile();
