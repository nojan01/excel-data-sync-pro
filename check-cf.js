/**
 * Prüfe CF (bedingte Formatierung) nach Spaltenlöschung
 */
const ExcelJS = require('exceljs');
const path = require('path');

async function checkCF() {
    // Original-Datei
    const originalFile = path.join(process.env.HOME, 'Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    // Exportierte Datei
    const exportedFile = path.join(process.env.HOME, 'Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    console.log('=== CF-VERGLEICH: Original vs Export ===\n');
    
    // Lade beide Dateien
    const wb1 = new ExcelJS.Workbook();
    await wb1.xlsx.readFile(originalFile);
    const ws1 = wb1.worksheets[0];
    
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(exportedFile);
    const ws2 = wb2.worksheets[0];
    
    // CF-Regeln auslesen
    const cf1 = ws1.conditionalFormattings || [];
    const cf2 = ws2.conditionalFormattings || [];
    
    console.log(`Original: ${cf1.length} CF-Bereiche`);
    console.log(`Export: ${cf2.length} CF-Bereiche`);
    
    // Erste 10 CF-Regeln vergleichen
    console.log('\n=== ERSTE 10 CF-REGELN ===\n');
    
    for (let i = 0; i < Math.min(10, cf1.length); i++) {
        const orig = cf1[i];
        const exp = cf2[i];
        
        console.log(`--- CF #${i + 1} ---`);
        console.log(`Original: ref="${orig?.ref}", rules: ${orig?.rules?.length || 0}`);
        if (orig?.rules?.[0]) {
            const rule = orig.rules[0];
            console.log(`  Type: ${rule.type}, Priority: ${rule.priority}`);
            if (rule.formulae) console.log(`  Formulae: ${JSON.stringify(rule.formulae)}`);
        }
        
        console.log(`Export:   ref="${exp?.ref}", rules: ${exp?.rules?.length || 0}`);
        if (exp?.rules?.[0]) {
            const rule = exp.rules[0];
            console.log(`  Type: ${rule.type}, Priority: ${rule.priority}`);
            if (rule.formulae) console.log(`  Formulae: ${JSON.stringify(rule.formulae)}`);
        }
        
        // Prüfen: Export ref sollte = Original ref mit verschobenen Spalten sein
        if (orig?.ref && exp?.ref) {
            const expected = adjustRange(orig.ref);
            if (exp.ref === expected) {
                console.log(`  ✅ CF korrekt verschoben`);
            } else {
                console.log(`  ❌ CF FALSCH: erwartet "${expected}", ist "${exp.ref}"`);
            }
        }
        console.log();
    }
}

// Simuliert die Verschiebung: Spalte A wird gelöscht, B->A, C->B, etc.
function adjustRange(ref) {
    // Beispiel: "B2:D10" -> "A2:C10"
    const match = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
        // Einzelne Zelle oder komplexer Range
        return ref.replace(/([A-Z]+)(\d+)/g, (m, col, row) => {
            const newCol = adjustCol(col);
            return newCol ? newCol + row : m;
        });
    }
    
    const [, startCol, startRow, endCol, endRow] = match;
    const newStartCol = adjustCol(startCol);
    const newEndCol = adjustCol(endCol);
    
    if (!newStartCol || !newEndCol) return null; // Spalte war A und wird gelöscht
    
    return `${newStartCol}${startRow}:${newEndCol}${endRow}`;
}

function adjustCol(col) {
    const num = colLetterToNumber(col);
    if (num <= 1) return null; // Spalte A wird gelöscht
    return colNumberToLetter(num - 1);
}

function colLetterToNumber(letters) {
    let result = 0;
    for (let i = 0; i < letters.length; i++) {
        result = result * 26 + (letters.charCodeAt(i) - 64);
    }
    return result;
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

checkCF().catch(console.error);
