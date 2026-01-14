const ExcelJS = require('exceljs');

// Hilfsfunktionen
function colLetterToNumber(col) {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
        num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num;
}

function colNumberToLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

async function analyzeCF() {
    const filePath = '/Users/nojan/Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    const cf = worksheet.conditionalFormattings;
    
    console.log('=== CF Analyse für letzte Spalte ===');
    console.log('Worksheet columnCount:', worksheet.columnCount);
    console.log('Letzte Spalte:', colNumberToLetter(worksheet.columnCount));
    console.log('CF Regeln gesamt:', cf ? cf.length : 0);
    
    if (!cf || cf.length === 0) return;
    
    // Finde alle Spalten die in CF referenziert werden
    const colsInCF = new Set();
    const lastColRefs = [];
    
    cf.forEach((cfEntry, idx) => {
        if (cfEntry.ref) {
            // Parse alle Spalten aus der Referenz
            const matches = cfEntry.ref.match(/[A-Z]+/g);
            if (matches) {
                matches.forEach(col => colsInCF.add(col));
            }
            
            // Prüfe ob die letzte Spalte referenziert wird
            const lastCol = colNumberToLetter(worksheet.columnCount);
            if (cfEntry.ref.includes(lastCol)) {
                lastColRefs.push({idx, ref: cfEntry.ref});
            }
        }
    });
    
    console.log('\nAlle Spalten in CF-Referenzen:', Array.from(colsInCF).sort().join(', '));
    console.log('\nCF-Regeln die letzte Spalte (' + colNumberToLetter(worksheet.columnCount) + ') referenzieren:');
    console.log('Anzahl:', lastColRefs.length);
    
    if (lastColRefs.length > 0) {
        console.log('\nBeispiele (max 10):');
        lastColRefs.slice(0, 10).forEach(r => {
            console.log('  [' + r.idx + '] ' + r.ref);
        });
    }
    
    // Simuliere was nach Spaltenlöschung passiert
    console.log('\n=== Simulation: Spalte L (12) löschen ===');
    const deletedCol = 12;
    const lastColNum = worksheet.columnCount; // z.B. 62 = BJ
    
    // Nach spliceColumns: Inhalte verschieben sich, aber CF-Refs bleiben
    // Die "neue" leere letzte Spalte ist immer noch auf Position lastColNum (BJ)
    // Aber die CF-Refs zeigen auf die ORIGINALEN Positionen
    
    console.log('Original letzte Spalte:', colNumberToLetter(lastColNum), '(', lastColNum, ')');
    console.log('Nach spliceColumns: Letzte Spalte mit Daten ist jetzt:', colNumberToLetter(lastColNum - 1));
    console.log('Die leere Spalte ist auf Position:', colNumberToLetter(lastColNum));
    
    // Welche CF-Referenzen NACH der Anpassung auf die leere Spalte zeigen würden
    console.log('\n=== CF nach Anpassung die auf leere Spalte zeigen ===');
    
    let wouldBeEmpty = [];
    cf.forEach((cfEntry, idx) => {
        if (cfEntry.ref) {
            // Simuliere adjustRangeReference
            const adjustedRef = cfEntry.ref.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
                const colNum = colLetterToNumber(col);
                if (colNum > deletedCol) {
                    return colNumberToLetter(colNum - 1) + row;
                }
                return match;
            });
            
            // Prüfe ob die angepasste Referenz auf lastColNum - 1 zeigt
            // NEIN - das ist falsch. Nach dem Verschieben ist die LEERE Spalte auf Position lastColNum
            // Die Referenzen die VOR Anpassung auf lastColNum zeigten, zeigen NACH Anpassung auf lastColNum-1
            // Aber diese Spalte hat jetzt Daten (von der nächsten Spalte)
            
            // Das richtige: CF-Refs die NACH Anpassung auf die neue letzte Spalte (lastColNum-1) zeigen
            // UND deren Original-Ref auf Spalten > deletedCol zeigten
            const lastColLetter = colNumberToLetter(lastColNum - 1); // z.B. BI
            if (adjustedRef.includes(lastColLetter)) {
                // Prüfe ob diese Referenz NUR auf diese Spalte zeigt
                const matches = adjustedRef.match(/[A-Z]+/g);
                const allLastCol = matches && matches.every(m => m === lastColLetter);
                if (allLastCol) {
                    wouldBeEmpty.push({idx, original: cfEntry.ref, adjusted: adjustedRef});
                }
            }
        }
    });
    
    console.log('CF die nach Anpassung nur auf ' + colNumberToLetter(lastColNum - 1) + ' zeigen:', wouldBeEmpty.length);
    wouldBeEmpty.slice(0, 5).forEach(r => {
        console.log('  Original:', r.original, '-> Angepasst:', r.adjusted);
    });
}

analyzeCF().catch(console.error);
