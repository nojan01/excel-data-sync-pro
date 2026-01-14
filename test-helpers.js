// Test der Hilfsfunktionen

function colNumberToLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

function refOnlyReferencesColumn(ref, colNumber) {
    const targetCol = colNumberToLetter(colNumber);
    const ranges = ref.split(' ');
    
    for (const range of ranges) {
        const parts = range.split(':');
        for (const part of parts) {
            const match = part.match(/^([A-Z]+)/);
            if (match && match[1] !== targetCol) {
                return false;
            }
        }
    }
    return true;
}

// Tests
console.log('Test 1: refOnlyReferencesColumn("A2:A10", 1)');
console.log('Erwartet: true, Ergebnis:', refOnlyReferencesColumn('A2:A10', 1));

console.log('\nTest 2: refOnlyReferencesColumn("B2:B10", 1)');
console.log('Erwartet: false, Ergebnis:', refOnlyReferencesColumn('B2:B10', 1));

console.log('\nTest 3: refOnlyReferencesColumn("A2:B10", 1)');
console.log('Erwartet: false, Ergebnis:', refOnlyReferencesColumn('A2:B10', 1));

console.log('\nTest 4: colNumberToLetter(1)');
console.log('Erwartet: A, Ergebnis:', colNumberToLetter(1));

console.log('\nTest 5: colNumberToLetter(61)');
console.log('Erwartet: BI, Ergebnis:', colNumberToLetter(61));
