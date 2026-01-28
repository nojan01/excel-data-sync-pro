#!/usr/bin/env node
/**
 * Test-Skript für Excel Live Session
 * 
 * Testet den kompletten Workflow:
 * 1. Datei öffnen
 * 2. Zeile löschen
 * 3. Zeile verschieben
 * 4. Zeile verstecken
 * 5. Zeile einfügen
 * 6. Zeile markieren
 * 7. Spalten-Operationen
 * 8. Speichern & Exportieren
 */

const { getLiveSession } = require('./python/excel_live_bridge');
const path = require('path');
const fs = require('fs');

async function runTest() {
    console.log('=== Excel Live Session Test ===\n');
    
    const session = getLiveSession();
    
    try {
        // 1. Session starten
        console.log('1. Starte Live Session...');
        const startResult = await session.start();
        console.log('   Ergebnis:', startResult);
        
        // Prüfe ob eine Test-Datei angegeben wurde
        const testFile = process.argv[2];
        const testSheet = process.argv[3] || 'Sheet1';
        
        if (!testFile) {
            console.log('\nUsage: node test-live-session.js <excel-file.xlsx> [sheet-name]');
            console.log('Example: node test-live-session.js ~/Documents/test.xlsx Tabelle1');
            await session.quit();
            return;
        }
        
        // 2. Datei öffnen
        console.log(`\n2. Öffne Datei: ${testFile}, Sheet: ${testSheet}`);
        const openResult = await session.openFile(testFile, testSheet);
        console.log('   Ergebnis:', openResult);
        
        if (!openResult.success) {
            console.error('   Fehler beim Öffnen:', openResult.error);
            await session.quit();
            return;
        }
        
        // 3. Daten lesen
        console.log('\n3. Lese Daten...');
        const dataResult = await session.getData();
        console.log(`   Headers: ${dataResult.headers?.length || 0} Spalten`);
        console.log(`   Daten: ${dataResult.data?.length || 0} Zeilen`);
        
        // 4. Test-Operationen (kommentiert für Sicherheit)
        console.log('\n4. Test-Operationen (deaktiviert für Sicherheit)');
        console.log('   Um Operationen zu testen, entkommentieren Sie den Code:');
        
        /*
        // Zeile löschen (z.B. Zeile 3 = Index 2)
        console.log('\n   4a. Lösche Zeile 3...');
        const deleteResult = await session.deleteRow(2);
        console.log('   Ergebnis:', deleteResult);
        
        // Zeile verschieben (z.B. Zeile 5 nach Zeile 2)
        console.log('\n   4b. Verschiebe Zeile 5 nach Position 2...');
        const moveResult = await session.moveRow(4, 1);
        console.log('   Ergebnis:', moveResult);
        
        // Zeile verstecken
        console.log('\n   4c. Verstecke Zeile 4...');
        const hideResult = await session.hideRow(3, true);
        console.log('   Ergebnis:', hideResult);
        
        // Zeile einfügen
        console.log('\n   4d. Füge neue Zeile bei Position 5 ein...');
        const insertResult = await session.insertRow(4, 1);
        console.log('   Ergebnis:', insertResult);
        
        // Zeile markieren
        console.log('\n   4e. Markiere Zeile 6 grün...');
        const highlightResult = await session.highlightRow(5, 'green');
        console.log('   Ergebnis:', highlightResult);
        */
        
        // 5. Speichern (als neue Datei)
        const outputFile = testFile.replace('.xlsx', '_live_test.xlsx');
        console.log(`\n5. Speichere unter: ${outputFile}`);
        const saveResult = await session.saveFile(outputFile);
        console.log('   Ergebnis:', saveResult);
        
        // 6. Session schließen
        console.log('\n6. Schließe Session...');
        const closeResult = await session.close();
        console.log('   Ergebnis:', closeResult);
        
        console.log('\n=== Test abgeschlossen ===');
        
    } catch (err) {
        console.error('Fehler:', err);
    } finally {
        await session.quit();
    }
}

// Test ausführen
runTest().catch(console.error);
