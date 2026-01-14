/**
 * Vergleich: Original vs exportierte Datei nach Spaltenlöschung
 */
const ExcelJS = require('exceljs');
const path = require('path');

async function compareFiles() {
    // Original-Datei
    const originalFile = path.join(process.env.HOME, 'Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    // Exportierte Datei (nach Spaltenlöschung in der App)
    const exportedFile = path.join(process.env.HOME, 'Desktop/Export_2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    
    console.log('=== VERGLEICH: Original vs Export ===\n');
    console.log('Original:', originalFile);
    console.log('Export:', exportedFile);
    
    // Lade beide Dateien
    const wb1 = new ExcelJS.Workbook();
    await wb1.xlsx.readFile(originalFile);
    const ws1 = wb1.worksheets[0];
    
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(exportedFile);
    const ws2 = wb2.worksheets[0];
    
    console.log(`\nOriginal: ${ws1.rowCount} Zeilen, ${ws1.columnCount} Spalten`);
    console.log(`Export: ${ws2.rowCount} Zeilen, ${ws2.columnCount} Spalten`);
    
    // Header vergleichen
    console.log('\n=== HEADER VERGLEICH ===');
    console.log('Original Zeile 1:');
    for (let col = 1; col <= 10; col++) {
        const cell = ws1.getCell(1, col);
        console.log(`  ${getColLetter(col)}: "${cell.value}"`);
    }
    console.log('\nExport Zeile 1 (sollte = Original Zeile 1, Spalte B-J):');
    for (let col = 1; col <= 9; col++) {
        const cell = ws2.getCell(1, col);
        console.log(`  ${getColLetter(col)}: "${cell.value}"`);
    }
    
    // Prüfen: Export Col A sollte = Original Col B sein
    console.log('\n=== WERT-VERGLEICH (Original B-J vs Export A-I) ===');
    let correctValues = 0;
    let wrongValues = 0;
    for (let row = 1; row <= Math.min(100, ws1.rowCount); row++) {
        for (let col = 2; col <= 10; col++) {
            const origVal = String(ws1.getCell(row, col).value || '');
            const expVal = String(ws2.getCell(row, col - 1).value || '');
            if (origVal === expVal) {
                correctValues++;
            } else {
                wrongValues++;
                if (wrongValues <= 5) {
                    console.log(`❌ Zeile ${row}: Orig[${getColLetter(col)}]="${origVal.substring(0, 20)}" != Exp[${getColLetter(col-1)}]="${expVal.substring(0, 20)}"`);
                }
            }
        }
    }
    console.log(`Korrekt: ${correctValues}, Falsch: ${wrongValues}`);
    
    // Fill-Vergleich
    console.log('\n=== FILL-VERGLEICH ===');
    let correctFills = 0;
    let wrongFills = 0;
    let lostFills = 0;
    
    for (let row = 1; row <= Math.min(100, ws1.rowCount); row++) {
        for (let col = 2; col <= 20; col++) {
            const origFill = getFillColor(ws1.getCell(row, col));
            const expFill = getFillColor(ws2.getCell(row, col - 1));
            
            if (origFill) {
                if (origFill === expFill) {
                    correctFills++;
                } else if (!expFill) {
                    lostFills++;
                    if (lostFills <= 5) {
                        console.log(`❌ FILL VERLOREN: Zeile ${row}, Orig[${getColLetter(col)}]="${origFill}" -> Exp[${getColLetter(col-1)}]=(kein Fill)`);
                    }
                } else {
                    wrongFills++;
                    if (wrongFills <= 5) {
                        console.log(`❌ FILL FALSCH: Zeile ${row}, Orig[${getColLetter(col)}]="${origFill}" != Exp[${getColLetter(col-1)}]="${expFill}"`);
                    }
                }
            }
        }
    }
    console.log(`Fills: ${correctFills} korrekt, ${wrongFills} falsch, ${lostFills} verloren`);
    
    // Font-Vergleich
    console.log('\n=== FONT-VERGLEICH ===');
    let correctFonts = 0;
    let wrongFonts = 0;
    
    for (let row = 1; row <= Math.min(100, ws1.rowCount); row++) {
        for (let col = 2; col <= 20; col++) {
            const origFont = getFontSummary(ws1.getCell(row, col));
            const expFont = getFontSummary(ws2.getCell(row, col - 1));
            
            if (origFont) {
                if (origFont === expFont) {
                    correctFonts++;
                } else {
                    wrongFonts++;
                    if (wrongFonts <= 5) {
                        console.log(`❌ FONT FALSCH: Zeile ${row}, Orig[${getColLetter(col)}]="${origFont}" != Exp[${getColLetter(col-1)}]="${expFont}"`);
                    }
                }
            }
        }
    }
    console.log(`Fonts: ${correctFonts} korrekt, ${wrongFonts} falsch`);
    
    console.log('\n=== ZUSAMMENFASSUNG ===');
    const totalErrors = wrongValues + wrongFills + lostFills + wrongFonts;
    const totalChecks = correctValues + wrongValues + correctFills + wrongFills + lostFills + correctFonts + wrongFonts;
    console.log(`Fehler: ${totalErrors} von ${totalChecks} (${(totalErrors / totalChecks * 100).toFixed(1)}%)`);
}

function getFillColor(cell) {
    if (cell.fill?.type === 'pattern' && cell.fill.fgColor) {
        return cell.fill.fgColor.argb || cell.fill.fgColor.theme || JSON.stringify(cell.fill.fgColor);
    }
    return null;
}

function getFontSummary(cell) {
    if (!cell.font) return null;
    const parts = [];
    if (cell.font.bold) parts.push('Bold');
    if (cell.font.italic) parts.push('Italic');
    if (cell.font.color?.argb) parts.push(cell.font.color.argb);
    if (cell.font.size) parts.push(`${cell.font.size}`);
    return parts.length ? parts.join(',') : null;
}

function getColLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

compareFiles().catch(console.error);
