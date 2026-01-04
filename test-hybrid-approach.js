/**
 * Hybrid-Ansatz Test:
 * - Daten mit xlsx-populate ändern
 * - Conditional Formatting XML aus Original-Datei erhalten
 */

const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function testHybridApproach() {
    // Test mit einer Beispieldatei
    const testSourcePath = process.argv[2];
    
    if (!testSourcePath) {
        console.log('Verwendung: node test-hybrid-approach.js <excel-datei>');
        console.log('');
        console.log('Dieser Test zeigt, wie Conditional Formatting erhalten werden kann.');
        return;
    }
    
    if (!fs.existsSync(testSourcePath)) {
        console.error('Datei nicht gefunden:', testSourcePath);
        return;
    }
    
    console.log('=== Hybrid-Ansatz Test ===\n');
    console.log('Quelldatei:', testSourcePath);
    
    // 1. Original-ZIP lesen und Conditional Formatting extrahieren
    console.log('\n1. Lese Original-ZIP und extrahiere Conditional Formatting...');
    
    const originalBuffer = fs.readFileSync(testSourcePath);
    const originalZip = await JSZip.loadAsync(originalBuffer);
    
    // Conditional Formatting ist in xl/worksheets/sheet*.xml
    const conditionalFormattingData = {};
    
    for (const [filename, file] of Object.entries(originalZip.files)) {
        if (filename.match(/xl\/worksheets\/sheet\d+\.xml$/)) {
            const content = await file.async('string');
            
            // Extrahiere <conditionalFormatting> Blöcke
            const cfMatches = content.match(/<conditionalFormatting[^>]*>[\s\S]*?<\/conditionalFormatting>/g);
            
            if (cfMatches && cfMatches.length > 0) {
                conditionalFormattingData[filename] = cfMatches;
                console.log(`   ${filename}: ${cfMatches.length} Conditional Formatting Regeln gefunden`);
            }
        }
    }
    
    if (Object.keys(conditionalFormattingData).length === 0) {
        console.log('   Keine Conditional Formatting Regeln in der Datei gefunden.');
        console.log('   (Das ist normal für Dateien ohne bedingte Formatierung)');
    }
    
    // 2. Datei mit xlsx-populate öffnen und modifizieren
    console.log('\n2. Öffne Datei mit xlsx-populate und mache Test-Änderung...');
    
    const workbook = await XlsxPopulate.fromFileAsync(testSourcePath);
    const sheet = workbook.sheet(0);
    
    // Test: Füge eine Zeile mit Daten hinzu
    const lastRow = sheet.usedRange() ? sheet.usedRange().endCell().rowNumber() : 1;
    console.log(`   Letzte Zeile: ${lastRow}`);
    
    // Schreibe Test-Daten in nächste Zeile
    sheet.cell(lastRow + 1, 1).value('TEST-HYBRID');
    sheet.cell(lastRow + 1, 2).value(new Date().toISOString());
    console.log(`   Test-Daten in Zeile ${lastRow + 1} geschrieben`);
    
    // 3. Speichere mit xlsx-populate
    const outputPath = testSourcePath.replace('.xlsx', '_hybrid_test.xlsx');
    await workbook.toFileAsync(outputPath);
    console.log(`\n3. Zwischendatei gespeichert: ${outputPath}`);
    
    // 4. Jetzt das Conditional Formatting wieder einfügen
    console.log('\n4. Füge Conditional Formatting wieder ein...');
    
    const modifiedBuffer = fs.readFileSync(outputPath);
    const modifiedZip = await JSZip.loadAsync(modifiedBuffer);
    
    let cfRestored = 0;
    
    for (const [filename, cfBlocks] of Object.entries(conditionalFormattingData)) {
        if (modifiedZip.files[filename]) {
            let content = await modifiedZip.files[filename].async('string');
            
            // Prüfe ob CF schon existiert (xlsx-populate könnte es behalten haben)
            const existingCF = content.match(/<conditionalFormatting[^>]*>[\s\S]*?<\/conditionalFormatting>/g);
            
            if (!existingCF || existingCF.length === 0) {
                // CF fehlt - füge es vor </worksheet> ein
                const insertPoint = content.lastIndexOf('</worksheet>');
                if (insertPoint > -1) {
                    const cfXml = cfBlocks.join('\n');
                    content = content.slice(0, insertPoint) + cfXml + '\n' + content.slice(insertPoint);
                    modifiedZip.file(filename, content);
                    cfRestored += cfBlocks.length;
                    console.log(`   ${filename}: ${cfBlocks.length} CF-Regeln wiederhergestellt`);
                }
            } else {
                console.log(`   ${filename}: CF bereits vorhanden (${existingCF.length} Regeln)`);
            }
        }
    }
    
    // 5. Finale Datei speichern
    const finalBuffer = await modifiedZip.generateAsync({ type: 'nodebuffer' });
    const finalPath = testSourcePath.replace('.xlsx', '_hybrid_final.xlsx');
    fs.writeFileSync(finalPath, finalBuffer);
    
    console.log(`\n5. Finale Datei gespeichert: ${finalPath}`);
    console.log(`   CF-Regeln wiederhergestellt: ${cfRestored}`);
    
    // Aufräumen - Zwischendatei löschen
    fs.unlinkSync(outputPath);
    
    console.log('\n=== Test abgeschlossen ===');
    console.log('\nBitte prüfe die Datei:', finalPath);
    console.log('- Sind die Conditional Formatting Regeln noch aktiv?');
    console.log('- Werden die Farben korrekt angezeigt?');
}

testHybridApproach().catch(console.error);
