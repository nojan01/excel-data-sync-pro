#!/usr/bin/env node

/**
 * ExcelJS Row-Move + Formatierungs-Test
 * 
 * Der WICHTIGSTE Test: Prüft ob Formatierung bei Row-Moves erhalten bleibt!
 * 
 * Verwendung:
 *   node test-row-move.js <excel-datei> <sheet-name>
 * 
 * Was wird getestet:
 * 1. Datei laden mit ExcelJS
 * 2. Zeile verschieben (simuliert Row-Move)
 * 3. Mit fullRewrite speichern
 * 4. Prüfen ob Formatierung erhalten bleibt
 */

const { readSheetWithExcelJS } = require('./exceljs-reader');
const { exportSheetWithExcelJS } = require('./exceljs-writer');
const path = require('path');
const fs = require('fs');

async function testRowMove(filePath, sheetName) {
    console.log('\n╔══════════════════════════════════════════════════════╗');
    console.log('║   ExcelJS Row-Move + Formatierungs-Test             ║');
    console.log('╚══════════════════════════════════════════════════════╝\n');
    console.log(`Datei: ${path.basename(filePath)}`);
    console.log(`Sheet: ${sheetName}\n`);
    
    try {
        // Schritt 1: Original-Daten laden
        console.log('► Schritt 1: Lade Original-Daten...');
        const originalData = await readSheetWithExcelJS(filePath, sheetName);
        
        if (!originalData.success) {
            console.error(`❌ Fehler: ${originalData.error}`);
            process.exit(1);
        }
        
        console.log(`   ✓ ${originalData.data.length} Zeilen geladen`);
        console.log(`   ✓ ${Object.keys(originalData.cellStyles).length} formatierte Zellen gefunden\n`);
        
        // Wichtige formatierte Zellen merken
        const originalStyles = JSON.parse(JSON.stringify(originalData.cellStyles));
        const originalRichText = JSON.parse(JSON.stringify(originalData.richTextCells));
        
        // Schritt 2: Row-Move simulieren (Zeile 5 nach Zeile 10 verschieben)
        console.log('► Schritt 2: Simuliere Row-Move (Zeile 5 → Zeile 10)...');
        
        if (originalData.data.length < 10) {
            console.error('❌ Datei muss mindestens 10 Daten-Zeilen haben!');
            process.exit(1);
        }
        
        // Zeile aus Array entfernen und an neuer Position einfügen
        const movedRow = originalData.data.splice(4, 1)[0]; // Zeile 5 (0-basiert = 4)
        originalData.data.splice(9, 0, movedRow); // An Position 10 einfügen
        
        // Styles für verschobene Zeilen aktualisieren
        const newCellStyles = {};
        const affectedRows = [4, 5, 6, 7, 8, 9]; // Alle betroffenen Zeilen
        
        for (const [key, style] of Object.entries(originalStyles)) {
            const [rowIdx, colIdx] = key.split('-').map(Number);
            
            if (rowIdx === 4) {
                // Zeile 5 → Zeile 10 (rowIdx 4 → 9)
                newCellStyles[`9-${colIdx}`] = style;
            } else if (rowIdx >= 5 && rowIdx <= 9) {
                // Zeilen 6-10 → Zeilen 5-9 (nach oben schieben)
                newCellStyles[`${rowIdx - 1}-${colIdx}`] = originalStyles[`${rowIdx}-${colIdx}`];
            } else {
                // Alle anderen Zeilen unverändert
                newCellStyles[key] = style;
            }
        }
        
        console.log(`   ✓ Zeile verschoben`);
        console.log(`   ✓ ${Object.keys(newCellStyles).length} Styles neu zugeordnet\n`);
        
        // Schritt 3: Mit fullRewrite speichern
        console.log('► Schritt 3: Speichere mit fullRewrite...');
        
        const tempFile = filePath.replace('.xlsx', '_ROWMOVE_TEST.xlsx');
        
        const sheetData = {
            sheetName: sheetName,
            headers: originalData.headers,
            data: originalData.data,
            cellStyles: newCellStyles,
            richTextCells: originalData.richTextCells,
            cellFormulas: originalData.cellFormulas,
            cellHyperlinks: originalData.cellHyperlinks,
            hiddenColumns: originalData.hiddenColumns,
            hiddenRows: originalData.hiddenRows,
            fullRewrite: true // WICHTIG!
        };
        
        const writeResult = await exportSheetWithExcelJS(filePath, tempFile, sheetData);
        
        if (!writeResult.success) {
            console.error(`❌ Fehler beim Speichern: ${writeResult.error}`);
            process.exit(1);
        }
        
        console.log(`   ✓ Gespeichert: ${path.basename(tempFile)}`);
        console.log(`   ✓ Zeit: ${writeResult.stats.totalTimeMs}ms\n`);
        
        // Schritt 4: Gespeicherte Datei neu laden und Formatierung prüfen
        console.log('► Schritt 4: Prüfe gespeicherte Datei...');
        
        const savedData = await readSheetWithExcelJS(tempFile, sheetName);
        
        if (!savedData.success) {
            console.error(`❌ Fehler beim Laden der gespeicherten Datei: ${savedData.error}`);
            process.exit(1);
        }
        
        console.log(`   ✓ ${savedData.data.length} Zeilen geladen`);
        console.log(`   ✓ ${Object.keys(savedData.cellStyles).length} formatierte Zellen gefunden\n`);
        
        // Schritt 5: Formatierung vergleichen
        console.log('╔══════════════════════════════════════════════════════╗');
        console.log('║                  ERGEBNIS                            ║');
        console.log('╚══════════════════════════════════════════════════════╝\n');
        
        // Vergleiche Anzahl formatierter Zellen
        const originalStyleCount = Object.keys(originalStyles).length;
        const newStyleCount = Object.keys(newCellStyles).length;
        const savedStyleCount = Object.keys(savedData.cellStyles).length;
        
        console.log('📊 Formatierte Zellen:');
        console.log(`   Original:     ${originalStyleCount}`);
        console.log(`   Nach Move:    ${newStyleCount}`);
        console.log(`   Gespeichert:  ${savedStyleCount}\n`);
        
        // Prüfe ob wichtige Styles erhalten sind
        let stylesPreserved = 0;
        let stylesLost = 0;
        
        for (const [key, originalStyle] of Object.entries(newCellStyles)) {
            const savedStyle = savedData.cellStyles[key];
            
            if (savedStyle) {
                // Prüfe ob alle Style-Properties erhalten sind
                const propsMatch = 
                    (originalStyle.bold === savedStyle.bold || (!originalStyle.bold && !savedStyle.bold)) &&
                    (originalStyle.italic === savedStyle.italic || (!originalStyle.italic && !savedStyle.italic)) &&
                    (originalStyle.fill === savedStyle.fill || (!originalStyle.fill && !savedStyle.fill));
                
                if (propsMatch) {
                    stylesPreserved++;
                } else {
                    stylesLost++;
                }
            } else {
                stylesLost++;
            }
        }
        
        const preserveRate = (stylesPreserved / newStyleCount * 100).toFixed(1);
        
        console.log('✨ Formatierungs-Erhaltung:');
        console.log(`   Erhalten:  ${stylesPreserved} (${preserveRate}%)`);
        console.log(`   Verloren:  ${stylesLost}\n`);
        
        // RichText prüfen
        const richTextCount = Object.keys(originalRichText).length;
        const savedRichTextCount = Object.keys(savedData.richTextCells).length;
        
        console.log('📝 RichText:');
        console.log(`   Original:     ${richTextCount}`);
        console.log(`   Gespeichert:  ${savedRichTextCount}\n`);
        
        // Finale Bewertung
        console.log('╔══════════════════════════════════════════════════════╗');
        console.log('║                  BEWERTUNG                           ║');
        console.log('╚══════════════════════════════════════════════════════╝\n');
        
        if (preserveRate >= 95) {
            console.log('✅ BESTANDEN - Formatierung wird sehr gut erhalten!');
            console.log('   ExcelJS ist für die Migration geeignet.\n');
        } else if (preserveRate >= 80) {
            console.log('⚠️  TEILWEISE - Formatierung wird größtenteils erhalten');
            console.log('   Weitere Tests empfohlen.\n');
        } else {
            console.log('❌ DURCHGEFALLEN - Zu viel Formatierung geht verloren!');
            console.log('   xlsx-populate bleibt die bessere Wahl.\n');
        }
        
        console.log(`💾 Test-Datei: ${tempFile}`);
        console.log('   Öffnen Sie die Datei in Excel um die Formatierung visuell zu prüfen.\n');
        
        // Dateigröße vergleichen
        const originalSize = fs.statSync(filePath).size;
        const savedSize = fs.statSync(tempFile).size;
        const sizeDiff = ((savedSize - originalSize) / originalSize * 100).toFixed(1);
        
        console.log('📁 Dateigröße:');
        console.log(`   Original: ${(originalSize / 1024 / 1024).toFixed(2)} MB`);
        console.log(`   Test:     ${(savedSize / 1024 / 1024).toFixed(2)} MB (${sizeDiff > 0 ? '+' : ''}${sizeDiff}%)\n`);
        
    } catch (error) {
        console.error('❌ Fehler:', error.message);
        console.error(error.stack);
        process.exit(1);
    }
}

// Kommandozeilen-Argumente
const args = process.argv.slice(2);

if (args.length < 2) {
    console.log('Verwendung: node test-row-move.js <excel-datei> <sheet-name>');
    console.log('Beispiel:   node test-row-move.js test.xlsx "Sheet1"');
    process.exit(1);
}

const [filePath, sheetName] = args;
testRowMove(filePath, sheetName);
