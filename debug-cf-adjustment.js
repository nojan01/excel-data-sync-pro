const ExcelJS = require('exceljs');

// Kopiere die Hilfsfunktionen aus exceljs-writer.js
function colLetterToNumber(letters) {
    let num = 0;
    for (let i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

function colNumberToLetter(num) {
    let result = '';
    while (num > 0) {
        const remainder = (num - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        num = Math.floor((num - 1) / 26);
    }
    return result;
}

function adjustCellReference(cellRef, deletedColNumber) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return cellRef;
    
    const colLetter = match[1];
    const rowNumber = match[2];
    const colNumber = colLetterToNumber(colLetter);
    
    if (colNumber > deletedColNumber) {
        return colNumberToLetter(colNumber - 1) + rowNumber;
    } else if (colNumber === deletedColNumber) {
        return colNumberToLetter(colNumber) + rowNumber;
    }
    return cellRef;
}

function adjustRangeReference(rangeRef, deletedColNumber) {
    const multiRanges = rangeRef.split(' ');
    const adjustedRanges = multiRanges.map(singleRange => {
        const parts = singleRange.split(':');
        if (parts.length === 2) {
            return adjustCellReference(parts[0], deletedColNumber) + ':' + adjustCellReference(parts[1], deletedColNumber);
        }
        return adjustCellReference(singleRange, deletedColNumber);
    });
    return adjustedRanges.join(' ');
}

function refOnlyReferencesColumn(ref, colNumber) {
    const colLetter = colNumberToLetter(colNumber);
    const ranges = ref.split(' ');
    
    for (const range of ranges) {
        const parts = range.split(':');
        for (const part of parts) {
            const match = part.match(/^([A-Z]+)/);
            if (match && match[1] !== colLetter) {
                return false;
            }
        }
    }
    return true;
}

function removeColumnFromRef(ref, colNumber) {
    const targetCol = colNumberToLetter(colNumber);
    const prevCol = colNumber > 1 ? colNumberToLetter(colNumber - 1) : null;
    const ranges = ref.split(' ');
    
    const processedRanges = [];
    
    for (const range of ranges) {
        const parts = range.split(':');
        
        if (parts.length === 2) {
            const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
            const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
            
            if (startMatch && endMatch) {
                const startCol = startMatch[1];
                const startRow = startMatch[2];
                const endCol = endMatch[1];
                const endRow = endMatch[2];
                
                if (startCol === targetCol && endCol === targetCol) {
                    continue;
                } else if (endCol === targetCol && prevCol) {
                    processedRanges.push(startCol + startRow + ':' + prevCol + endRow);
                } else if (startCol === targetCol) {
                    const nextCol = colNumberToLetter(colNumber + 1);
                    processedRanges.push(nextCol + startRow + ':' + endCol + endRow);
                } else {
                    processedRanges.push(range);
                }
            } else {
                processedRanges.push(range);
            }
        } else {
            const match = range.match(/^([A-Z]+)/);
            if (match && match[1] !== targetCol) {
                processedRanges.push(range);
            }
        }
    }
    
    return processedRanges.join(' ') || ref;
}

async function main() {
    const exportPath = '/Users/nojan/Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx';
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(exportPath);
    
    const worksheet = workbook.worksheets[0];
    const cf = worksheet.conditionalFormattings;
    
    console.log('=== Analysiere CF in Export-Datei ===');
    console.log('Gesamt CF-Regeln:', cf.length);
    console.log('');
    
    // Suche nach CF mit AY
    console.log('=== CF-Regeln die "AY" enthalten ===');
    const ayRefs = cf.filter(c => c.ref && c.ref.includes('AY'));
    console.log('Gefunden:', ayRefs.length);
    ayRefs.forEach(c => console.log('  -', c.ref));
    
    console.log('');
    console.log('=== Simulation: Was sollte passiert sein? ===');
    
    // Simuliere die Anpassung für AY1:AY2404
    const testRef = 'AY1:AY2404';
    const deletedColNumber = 1;  // Spalte A
    const lastColumnBeforeDelete = 61;  // War ursprünglich BK
    const emptyLastCol = lastColumnBeforeDelete;  // Nach Löschung ist BK leer
    
    console.log('Test-Referenz:', testRef);
    console.log('Gelöschte Spalte:', deletedColNumber, '(' + colNumberToLetter(deletedColNumber) + ')');
    console.log('Letzte Spalte vor Löschung:', lastColumnBeforeDelete, '(' + colNumberToLetter(lastColumnBeforeDelete) + ')');
    console.log('');
    
    // Schritt 1: Verschiebe nach links
    const adjustedRef = adjustRangeReference(testRef, deletedColNumber);
    console.log('Schritt 1 - adjustRangeReference:', testRef, '→', adjustedRef);
    
    // Schritt 2: Prüfe ob nur leere Spalte
    const onlyEmptyCol = refOnlyReferencesColumn(adjustedRef, emptyLastCol);
    console.log('Schritt 2 - refOnlyReferencesColumn(' + adjustedRef + ', ' + emptyLastCol + '):', onlyEmptyCol);
    
    if (onlyEmptyCol) {
        console.log('  → WÜRDE ENTFERNT WERDEN (zeigt nur auf leere Spalte)');
    } else {
        // Schritt 3: Entferne leere Spalte aus Ref
        const cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
        console.log('Schritt 3 - removeColumnFromRef(' + adjustedRef + ', ' + emptyLastCol + '):', cleanedRef);
        
        console.log('');
        console.log('Vergleich:');
        console.log('  oldRef:', testRef);
        console.log('  cleanedRef:', cleanedRef);
        console.log('  cleanedRef !== oldRef:', cleanedRef !== testRef);
        
        if (cleanedRef !== testRef) {
            console.log('  → cfEntry.ref würde auf', cleanedRef, 'gesetzt werden');
        } else {
            console.log('  → PROBLEM: cfEntry.ref würde NICHT geändert!');
        }
    }
    
    console.log('');
    console.log('=== Prüfe echte AY-Referenzen ===');
    for (const entry of ayRefs.slice(0, 5)) {
        const oldRef = entry.ref;
        const adjusted = adjustRangeReference(oldRef, deletedColNumber);
        const cleaned = removeColumnFromRef(adjusted, emptyLastCol);
        console.log('');
        console.log('Ref:', oldRef);
        console.log('  → adjusted:', adjusted);
        console.log('  → cleaned:', cleaned);
        console.log('  → würde aktualisiert:', cleaned !== oldRef);
    }
}

main().catch(console.error);
