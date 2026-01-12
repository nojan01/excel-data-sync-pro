#!/usr/bin/env node

/**
 * ExcelJS VISUELLER Row-Move Test
 * 
 * Macht die verschobene Zeile DEUTLICH SICHTBAR durch:
 * - Gelber Hintergrund auf der verschobenen Zeile
 * - Report welche Zeile verschoben wurde
 */

const { readSheetWithExcelJS } = require('./exceljs-reader');
const { exportSheetWithExcelJS } = require('./exceljs-writer');
const ExcelJS = require('exceljs');
const path = require('path');

async function visualRowMoveTest(filePath, sheetName) {
    console.log('\n╔══════════════════════════════════════════════════════╗');
    console.log('║     VISUELLER Row-Move Test mit ExcelJS             ║');
    console.log('╚══════════════════════════════════════════════════════╝\n');
    
    try {
        // Original laden
        console.log('► Schritt 1: Lade Original...');
        const original = await readSheetWithExcelJS(filePath, sheetName);
        
        if (!original.success) {
            console.error(`❌ ${original.error}`);
            process.exit(1);
        }
        
        console.log(`   ✓ ${original.data.length} Zeilen\n`);
        
        // Zeige welche Zeile wir verschieben (mit erstem Wert)
        const sourceRowIdx = 4; // Zeile 5 (Daten-Zeile, 0-basiert)
        const targetRowIdx = 9; // Zeile 10
        
        const sourceRowData = original.data[sourceRowIdx];
        const firstValue = sourceRowData[0] || sourceRowData[1] || sourceRowData[2] || 'leer';
        
        console.log('► Schritt 2: Row-Move Info');
        console.log(`   Von:   Zeile ${sourceRowIdx + 2} (Daten-Zeile ${sourceRowIdx + 1})`);
        console.log(`   Nach:  Zeile ${targetRowIdx + 2} (Daten-Zeile ${targetRowIdx + 1})`);
        console.log(`   Inhalt: "${String(firstValue).substring(0, 50)}..."\n`);
        
        // Workbook direkt mit ExcelJS laden und modifizieren
        console.log('► Schritt 3: Verschiebe Zeile...');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            console.error(`❌ Sheet nicht gefunden`);
            process.exit(1);
        }
        
        // Zeile verschieben (Excel-Zeilen sind 1-basiert, +2 wegen Header)
        const sourceExcelRow = sourceRowIdx + 2;
        const targetExcelRow = targetRowIdx + 2;
        
        // Zeile kopieren
        const sourceRow = worksheet.getRow(sourceExcelRow);
        const targetRow = worksheet.getRow(targetExcelRow);
        
        // Backup der Quellzeile
        const sourceValues = [];
        const sourceCells = [];
        sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            sourceValues[colNumber] = {
                value: cell.value,
                style: {
                    font: cell.font ? { ...cell.font } : {},
                    fill: cell.fill ? { ...cell.fill } : {},
                    alignment: cell.alignment ? { ...cell.alignment } : {},
                    border: cell.border ? { ...cell.border } : {}
                }
            };
        });
        
        // Zeilen zwischen source und target nach oben/unten schieben
        if (targetRowIdx > sourceRowIdx) {
            // Nach unten verschieben: Zeilen dazwischen nach oben
            for (let i = sourceRowIdx + 1; i <= targetRowIdx; i++) {
                const fromRow = worksheet.getRow(i + 2);
                const toRow = worksheet.getRow(i + 1);
                
                fromRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const targetCell = toRow.getCell(colNumber);
                    targetCell.value = cell.value;
                    targetCell.style = {
                        font: cell.font ? { ...cell.font } : {},
                        fill: cell.fill ? { ...cell.fill } : {},
                        alignment: cell.alignment ? { ...cell.alignment } : {},
                        border: cell.border ? { ...cell.border } : {}
                    };
                });
            }
        }
        
        // Verschobene Zeile an Zielposition setzen + GELB markieren
        Object.keys(sourceValues).forEach(colNumber => {
            const col = parseInt(colNumber);
            const targetCell = targetRow.getCell(col);
            const source = sourceValues[col];
            
            targetCell.value = source.value;
            
            // Style übernehmen ABER Hintergrund GELB für Sichtbarkeit
            targetCell.font = source.style.font;
            targetCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' } // GELB!
            };
            targetCell.alignment = source.style.alignment;
            targetCell.border = source.style.border;
        });
        
        console.log(`   ✓ Zeile ${sourceExcelRow} → Zeile ${targetExcelRow}`);
        console.log(`   ✓ Verschobene Zeile GELB markiert\n`);
        
        // RichText separat prüfen
        console.log('► Schritt 4: Prüfe RichText...');
        let richTextFound = 0;
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value && typeof cell.value === 'object' && cell.value.richText) {
                    richTextFound++;
                    console.log(`   ✓ RichText in Zeile ${rowNumber}, Spalte ${colNumber}`);
                    console.log(`     Inhalt: "${cell.value.richText.map(r => r.text).join('')}"`);
                }
            });
        });
        
        if (richTextFound === 0) {
            console.log(`   ⚠️  Keine RichText-Zellen gefunden im Original\n`);
        } else {
            console.log(`   ✓ ${richTextFound} RichText-Zellen gefunden\n`);
        }
        
        // Speichern
        console.log('► Schritt 5: Speichere Test-Datei...');
        const outputPath = filePath.replace('.xlsx', '_VISUAL_ROWMOVE_TEST.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        
        console.log(`   ✓ Gespeichert: ${path.basename(outputPath)}\n`);
        
        // Finale Anweisungen
        console.log('╔══════════════════════════════════════════════════════╗');
        console.log('║              BITTE JETZT PRÜFEN                      ║');
        console.log('╚══════════════════════════════════════════════════════╝\n');
        
        console.log('📂 Öffnen Sie die Datei in Excel:');
        console.log(`   ${outputPath}\n`);
        
        console.log('👀 Prüfen Sie visuell:');
        console.log(`   1. Zeile ${targetExcelRow} hat GELBEN Hintergrund`);
        console.log(`   2. Inhalt der gelben Zeile: "${String(firstValue).substring(0, 40)}..."`);
        console.log(`   3. Alle anderen Formatierungen intakt?`);
        console.log(`   4. RichText-Zellen korrekt?\n`);
        
    } catch (error) {
        console.error('❌ Fehler:', error.message);
        console.error(error.stack);
        process.exit(1);
    }
}

// Kommandozeilen-Argumente
const args = process.argv.slice(2);

if (args.length < 2) {
    console.log('Verwendung: node test-visual-rowmove.js <excel-datei> <sheet-name>');
    console.log('Beispiel:   node test-visual-rowmove.js test.xlsx "Sheet1"');
    process.exit(1);
}

const [filePath, sheetName] = args;
visualRowMoveTest(filePath, sheetName);
