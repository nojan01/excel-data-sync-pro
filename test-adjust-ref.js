function colLetterToNumber(letters) {
    let num = 0;
    for (let i = 0; i < letters.length; i++) {
        num = num * 26 + (letters.charCodeAt(i) - 64);
    }
    return num;
}

function colNumberToLetter(num) {
    let result = '';
    while (num > 0) {
        const remainder = (num - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        num = Math.floor((num - 1) / 26);
    }
    return result;
}

function adjustCellReference(cellRef, deletedColNumber) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return cellRef;
    
    const colLetter = match[1];
    const rowNumber = match[2];
    const colNumber = colLetterToNumber(colLetter);
    
    if (colNumber > deletedColNumber) {
        return colNumberToLetter(colNumber - 1) + rowNumber;
    } else if (colNumber === deletedColNumber) {
        return colNumberToLetter(colNumber) + rowNumber;
    }
    return cellRef;
}

function adjustRangeReference(rangeRef, deletedColNumber) {
    const multiRanges = rangeRef.split(' ');
    const adjustedRanges = multiRanges.map(singleRange => {
        const parts = singleRange.split(':');
        if (parts.length === 2) {
            return adjustCellReference(parts[0], deletedColNumber) + ':' + adjustCellReference(parts[1], deletedColNumber);
        }
        return adjustCellReference(singleRange, deletedColNumber);
    });
    return adjustedRanges.join(' ');
}

console.log('Test: AY1:AY2404 bei Löschung von Spalte 1 (A)');
console.log('AY = Spalte', colLetterToNumber('AY'));
console.log('');
console.log('Ergebnis:', adjustRangeReference('AY1:AY2404', 1));
console.log('Erwartet: AX1:AX2404');
