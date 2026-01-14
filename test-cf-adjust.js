// Test der CF-Anpassungslogik
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

function removeColumnFromRef(ref, colNumber) {
    const targetCol = colNumberToLetter(colNumber);
    const prevCol = colNumber > 1 ? colNumberToLetter(colNumber - 1) : null;
    const ranges = ref.split(' ');
    
    const processedRanges = [];
    
    for (const range of ranges) {
        const parts = range.split(':');
        
        if (parts.length === 2) {
            const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
            const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
            
            if (startMatch && endMatch) {
                const startCol = startMatch[1];
                const startRow = startMatch[2];
                const endCol = endMatch[1];
                const endRow = endMatch[2];
                
                if (startCol === targetCol && endCol === targetCol) {
                    continue;
                } else if (endCol === targetCol && prevCol) {
                    processedRanges.push(startCol + startRow + ':' + prevCol + endRow);
                } else if (startCol === targetCol) {
                    const nextCol = colNumberToLetter(colNumber + 1);
                    processedRanges.push(nextCol + startRow + ':' + endCol + endRow);
                } else {
                    processedRanges.push(range);
                }
            } else {
                processedRanges.push(range);
            }
        } else {
            const match = range.match(/^([A-Z]+)/);
            if (match && match[1] !== targetCol) {
                processedRanges.push(range);
            }
        }
    }
    
    return processedRanges.join(' ') || ref;
}

// Test: Spalte 3 (C) wird gelöscht, CF ist auf D1:E10
console.log('=== Test 1: CF rechts von gelöschter Spalte ===');
let oldRef = 'D1:E10';
let deletedColNumber = 3; // C
let lastColumnBeforeDelete = 10; // J
let emptyLastCol = lastColumnBeforeDelete;

console.log('oldRef:', oldRef);
console.log('deletedColNumber:', deletedColNumber, '(' + colNumberToLetter(deletedColNumber) + ')');
console.log('lastColumnBeforeDelete:', lastColumnBeforeDelete, '(' + colNumberToLetter(lastColumnBeforeDelete) + ')');

let adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
console.log('adjustedRef:', adjustedRef);

let cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
console.log('cleanedRef:', cleanedRef);

console.log('cleanedRef !== oldRef:', cleanedRef !== oldRef);
console.log('=> cfEntry.ref wird gesetzt auf:', cleanedRef !== oldRef ? cleanedRef : 'NICHT GEÄNDERT');

console.log('\n=== Test 2: CF auf gleicher Spalte wie gelöschte ===');
oldRef = 'C1:C10';
deletedColNumber = 3; // C
lastColumnBeforeDelete = 10;
emptyLastCol = lastColumnBeforeDelete;

console.log('oldRef:', oldRef);
adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
console.log('adjustedRef:', adjustedRef);
cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
console.log('cleanedRef:', cleanedRef);
console.log('cleanedRef !== oldRef:', cleanedRef !== oldRef);

console.log('\n=== Test 3: Realistische CF-Referenz (BH:BI bei Löschung von Spalte 5) ===');
oldRef = 'BH1:BI100';
deletedColNumber = 5; // E
lastColumnBeforeDelete = 62; // BJ
emptyLastCol = lastColumnBeforeDelete;

console.log('oldRef:', oldRef);
console.log('deletedColNumber:', deletedColNumber, '(' + colNumberToLetter(deletedColNumber) + ')');
console.log('emptyLastCol:', emptyLastCol, '(' + colNumberToLetter(emptyLastCol) + ')');

adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
console.log('adjustedRef:', adjustedRef);
cleanedRef = removeColumnFromRef(adjustedRef, emptyLastCol);
console.log('cleanedRef:', cleanedRef);
console.log('cleanedRef !== oldRef:', cleanedRef !== oldRef);

console.log('\n=== Test 4: CF links von gelöschter Spalte (sollte nicht geändert werden) ===');
oldRef = 'A1:B10';
deletedColNumber = 5; // E

console.log('oldRef:', oldRef);
adjustedRef = adjustRangeReference(oldRef, deletedColNumber);
console.log('adjustedRef:', adjustedRef);
console.log('Wurde adjustedRef geändert?', adjustedRef !== oldRef);
