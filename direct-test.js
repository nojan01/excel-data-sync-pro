// Direkter Test: Was kommt aus dem Reader und was erwartet das Frontend?
const { readSheetWithExcelJS } = require('./exceljs-reader');

async function directTest() {
    console.log('=== DIREKTER INDEX-TEST ===\n');
    
    const result = await readSheetWithExcelJS('/Users/nojan/Desktop/test-styles.xlsx', 'Style Tests');
    
    console.log('1. result.data Länge:', result.data.length);
    console.log('2. result.data[0] (sollte Header sein):', result.data[0]?.slice(0, 3));
    console.log('3. result.data[1] (sollte erste Datenzeile sein):', result.data[1]?.slice(0, 3));
    
    // Simuliere Frontend: data.slice(1)
    const frontendData = result.data.slice(1);
    console.log('\n4. Nach slice(1) - frontendData[0]:', frontendData[0]?.slice(0, 3));
    
    // Was sucht das Frontend für frontendData[0]?
    const originalIndex = 0;
    const colIndex = 0;
    const frontendKey = `${originalIndex + 1}-${colIndex}`;
    console.log('\n5. Frontend sucht Key:', frontendKey);
    console.log('6. Style für diesen Key:', result.cellStyles[frontendKey]);
    
    // Was sind die tatsächlichen Keys?
    console.log('\n7. Alle Style-Keys:', Object.keys(result.cellStyles).slice(0, 15));
    
    // Test für Zeile mit bekanntem Style (A3 = "Hintergrundfarben:", bold)
    console.log('\n8. Suche "Hintergrundfarben:" in Daten...');
    frontendData.forEach((row, idx) => {
        if (row[0] && row[0].toString().includes('Hintergrund')) {
            console.log(`   Gefunden bei frontendData[${idx}], Wert: "${row[0]}"`);
            console.log(`   Frontend würde Key suchen: ${idx + 1}-0`);
            console.log(`   Style dafür:`, result.cellStyles[`${idx + 1}-0`]);
        }
    });
}

directTest().catch(console.error);
