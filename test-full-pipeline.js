// Test: Voller Pipeline-Test - simuliert was main.js macht
const { readSheetWithExcelJS } = require('./exceljs-reader');

async function testFullPipeline() {
    const filePath = '/Users/nojan/Desktop/test-styles.xlsx';
    const sheetName = 'Style Tests'; // Anpassen an echten Sheet-Namen
    
    console.log('='.repeat(60));
    console.log('TEST: Volle Pipeline (wie main.js)');
    console.log('='.repeat(60));
    
    try {
        // 1. ExcelJS Reader aufrufen (wie in main.js)
        const result = await readSheetWithExcelJS(filePath, sheetName, null);
        
        if (!result.success) {
            console.log('FEHLER:', result.error);
            return;
        }
        
        console.log('\n1. ExcelJS Reader Ergebnis:');
        console.log('   - headers:', result.headers?.length);
        console.log('   - data:', result.data?.length, 'Zeilen');
        console.log('   - cellStyles:', Object.keys(result.cellStyles || {}).length, 'Einträge');
        
        // 2. Simuliere main.js Handler (spread + override)
        const mainJsResult = {
            ...result,
            dataValidations: {},
            mergedCells: []
        };
        
        console.log('\n2. Nach main.js Handler ({...result, ...}):');
        console.log('   - cellStyles:', Object.keys(mainJsResult.cellStyles || {}).length, 'Einträge');
        console.log('   - Hat cellStyles?', 'cellStyles' in mainJsResult);
        
        // 3. Zeige einige Style-Keys
        if (mainJsResult.cellStyles) {
            const keys = Object.keys(mainJsResult.cellStyles);
            console.log('\n3. Style-Keys (erste 10):', keys.slice(0, 10));
            
            // Zeige einige Styles
            console.log('\n4. Beispiel-Styles:');
            keys.slice(0, 5).forEach(key => {
                console.log(`   ${key}:`, mainJsResult.cellStyles[key]);
            });
        }
        
        // 4. Simuliere Frontend-Lookup
        console.log('\n5. Frontend-Simulation:');
        console.log('   explorerState.data = result.data.slice(1)');
        const frontendData = mainJsResult.data.slice(1);
        console.log('   frontendData hat', frontendData.length, 'Zeilen');
        
        // Simuliere Rendering der ersten Zeilen
        console.log('\n6. Render-Simulation (erste 3 Zeilen, erste 3 Spalten):');
        for (let originalIndex = 0; originalIndex < Math.min(3, frontendData.length); originalIndex++) {
            for (let colIndex = 0; colIndex < Math.min(3, frontendData[originalIndex]?.length || 0); colIndex++) {
                // Genau wie im Frontend:
                const cellStyleKey = `${originalIndex + 1}-${colIndex}`;
                const cellStyle = mainJsResult.cellStyles[cellStyleKey];
                const cellValue = frontendData[originalIndex][colIndex];
                
                console.log(`   Row ${originalIndex}, Col ${colIndex}: Key="${cellStyleKey}", Value="${cellValue}", Style:`, cellStyle || 'KEINE');
            }
        }
        
        console.log('\n' + '='.repeat(60));
        console.log('ERGEBNIS: Styles werden ' + (Object.keys(mainJsResult.cellStyles).length > 0 ? 'KORREKT übertragen' : 'NICHT übertragen'));
        console.log('='.repeat(60));
        
    } catch (error) {
        console.error('TEST FEHLER:', error);
    }
}

testFullPipeline();
