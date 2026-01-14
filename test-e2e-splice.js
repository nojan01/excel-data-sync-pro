/**
 * Vollständiger End-to-End Test: 
 * Öffne Datei -> Lösche Spalte A -> Speichere -> Vergleiche
 */
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function fullE2ETest() {
    const inputFile = path.join(process.env.HOME, 'Desktop/2025-08-31 DEFENCE&SPACE MVMS Master Asset List GER.xlsx');
    const outputFile = path.join(process.env.HOME, 'Desktop/E2E_Test_Output.xlsx');
    
    console.log('=== END-TO-END TEST ===\n');
    console.log('Input:', inputFile);
    console.log('Output:', outputFile);
    
    // 1. Öffne die Datei
    console.log('\n1. Lade Datei...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFile);
    
    const worksheet = workbook.worksheets[0];
    console.log(`   Sheet: ${worksheet.name}`);
    console.log(`   Rows: ${worksheet.rowCount}, Cols: ${worksheet.columnCount}`);
    
    // 2. Snapshot VOR dem Splice
    console.log('\n2. Erstelle Snapshot vor spliceColumns...');
    const beforeSnapshot = createSnapshot(worksheet);
    console.log(`   ${Object.keys(beforeSnapshot.values).length} Werte gespeichert`);
    console.log(`   ${Object.keys(beforeSnapshot.fills).length} Fills gespeichert`);
    console.log(`   ${Object.keys(beforeSnapshot.fonts).length} Fonts gespeichert`);
    
    // 3. Lösche Spalte A
    console.log('\n3. Lösche Spalte A mit spliceColumns(1, 1)...');
    worksheet.spliceColumns(1, 1);
    console.log('   Spalte gelöscht.');
    
    // 4. Speichere die Datei
    console.log('\n4. Speichere Datei...');
    await workbook.xlsx.writeFile(outputFile);
    console.log('   Gespeichert.');
    
    // 5. Lade die gespeicherte Datei neu
    console.log('\n5. Lade gespeicherte Datei neu...');
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(outputFile);
    const worksheet2 = workbook2.worksheets[0];
    console.log(`   Sheet: ${worksheet2.name}`);
    console.log(`   Rows: ${worksheet2.rowCount}, Cols: ${worksheet2.columnCount}`);
    
    // 6. Erstelle Snapshot NACH dem Neuladen
    console.log('\n6. Erstelle Snapshot nach Neuladen...');
    const afterSnapshot = createSnapshot(worksheet2);
    console.log(`   ${Object.keys(afterSnapshot.values).length} Werte gelesen`);
    console.log(`   ${Object.keys(afterSnapshot.fills).length} Fills gelesen`);
    console.log(`   ${Object.keys(afterSnapshot.fonts).length} Fonts gelesen`);
    
    // 7. Vergleich
    console.log('\n7. VERGLEICH (erwarte: Spalte B->A, C->B, etc.)...\n');
    
    let valueCorrect = 0;
    let valueWrong = 0;
    let fillCorrect = 0;
    let fillWrong = 0;
    let fillLost = 0;
    let fontCorrect = 0;
    let fontWrong = 0;
    
    // Für jede Zelle die vorher in Spalte 2+ war (nach Verschiebung: Spalte 1+)
    for (let row = 1; row <= Math.min(200, beforeSnapshot.maxRow); row++) {
        for (let col = 2; col <= Math.min(20, beforeSnapshot.maxCol); col++) {
            const beforeKey = `${row}-${col}`;
            const afterKey = `${row}-${col - 1}`; // Nach Verschiebung
            
            // Werte
            const beforeVal = beforeSnapshot.values[beforeKey];
            const afterVal = afterSnapshot.values[afterKey];
            if (beforeVal && beforeVal === afterVal) {
                valueCorrect++;
            } else if (beforeVal && beforeVal !== afterVal) {
                valueWrong++;
                if (valueWrong <= 5) {
                    console.log(`❌ WERT: ${beforeKey} -> ${afterKey}: "${beforeVal}" !== "${afterVal}"`);
                }
            }
            
            // Fills
            const beforeFill = beforeSnapshot.fills[beforeKey];
            const afterFill = afterSnapshot.fills[afterKey];
            if (beforeFill) {
                if (beforeFill === afterFill) {
                    fillCorrect++;
                } else if (!afterFill) {
                    fillLost++;
                    if (fillLost <= 5) {
                        console.log(`❌ FILL VERLOREN: ${beforeKey} -> ${afterKey}: "${beforeFill}" -> (kein Fill)`);
                    }
                } else {
                    fillWrong++;
                    if (fillWrong <= 5) {
                        console.log(`❌ FILL FALSCH: ${beforeKey} -> ${afterKey}: "${beforeFill}" !== "${afterFill}"`);
                    }
                }
            }
            
            // Fonts
            const beforeFont = beforeSnapshot.fonts[beforeKey];
            const afterFont = afterSnapshot.fonts[afterKey];
            if (beforeFont) {
                if (beforeFont === afterFont) {
                    fontCorrect++;
                } else {
                    fontWrong++;
                    if (fontWrong <= 5) {
                        console.log(`⚠️  FONT: ${beforeKey} -> ${afterKey}: "${beforeFont}" !== "${afterFont}"`);
                    }
                }
            }
        }
    }
    
    console.log('\n=== ZUSAMMENFASSUNG ===');
    console.log(`Werte: ${valueCorrect} korrekt, ${valueWrong} falsch`);
    console.log(`Fills: ${fillCorrect} korrekt, ${fillWrong} falsch, ${fillLost} verloren`);
    console.log(`Fonts: ${fontCorrect} korrekt, ${fontWrong} falsch`);
    
    const totalChecks = valueCorrect + valueWrong + fillCorrect + fillWrong + fillLost + fontCorrect + fontWrong;
    const totalWrong = valueWrong + fillWrong + fillLost + fontWrong;
    const percentWrong = totalChecks > 0 ? (totalWrong / totalChecks * 100).toFixed(1) : 0;
    console.log(`\nGesamt: ${percentWrong}% falsch`);
}

function createSnapshot(worksheet) {
    const values = {};
    const fills = {};
    const fonts = {};
    let maxRow = 0;
    let maxCol = 0;
    
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        maxRow = Math.max(maxRow, rowNumber);
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            maxCol = Math.max(maxCol, colNumber);
            const key = `${rowNumber}-${colNumber}`;
            
            // Wert
            values[key] = String(cell.value || '').substring(0, 30);
            
            // Fill
            if (cell.fill?.type === 'pattern' && cell.fill.fgColor) {
                fills[key] = cell.fill.fgColor.argb || cell.fill.fgColor.theme || JSON.stringify(cell.fill.fgColor);
            }
            
            // Font (vereinfacht)
            if (cell.font) {
                const parts = [];
                if (cell.font.bold) parts.push('Bold');
                if (cell.font.color?.argb) parts.push(cell.font.color.argb);
                if (cell.font.size) parts.push(`${cell.font.size}`);
                if (parts.length) fonts[key] = parts.join(',');
            }
        });
    });
    
    return { values, fills, fonts, maxRow, maxCol };
}

fullE2ETest().catch(console.error);
