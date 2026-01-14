// Test der neuen Style-Key Berechnung
const { readSheetWithExcelJS } = require('./exceljs-reader.js');

async function testNewIndexing() {
    const result = await readSheetWithExcelJS('/Users/nojan/Desktop/test-styles.xlsx', 'Style Tests');
    
    if (!result.success) {
        console.error('Fehler:', result.error);
        return;
    }
    
    console.log('=== NEUE STYLE-KEYS ===');
    console.log('Anzahl Styles:', Object.keys(result.cellStyles).length);
    
    // Zeige alle Style-Keys
    console.log('\nAlle Style-Keys:', Object.keys(result.cellStyles).join(', '));
    
    // Prüfe spezifisch Zeile 7 (0-basiert = 8 im 1-basierten System, wo Fett/Kursiv/etc. sein sollte)
    console.log('\n=== PRÜFE SCHRIFTFORMATIERUNGEN ===');
    for (let i = 0; i < 10; i++) {
        const key = `7-${i}`;
        if (result.cellStyles[key]) {
            console.log(`[${key}]:`, JSON.stringify(result.cellStyles[key]));
        }
    }
    
    // Prüfe die Daten um zu sehen, wo Fett/Kursiv ist
    console.log('\n=== DATEN-ZEILEN ===');
    result.data.forEach((row, index) => {
        if (index > 0 && index <= 10) { // Erste 10 Daten-Zeilen (nach Header)
            console.log(`data[${index}] (originalIndex ${index-1}):`, row.slice(0, 5).join(' | '));
        }
    });
}

testNewIndexing();
