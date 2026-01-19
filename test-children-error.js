const xlsx = require('xlsx-populate');
const fs = require('fs');

// Teste die zuletzt exportierte Datei
const testFile = '/Users/nojan/Desktop/test-styles-exceljs.xlsx';

console.log('Testing xlsx-populate reading...');
console.log('File exists:', fs.existsSync(testFile));
console.log('File size:', fs.statSync(testFile).size, 'bytes');
console.log('Last modified:', fs.statSync(testFile).mtime);

xlsx.fromFileAsync(testFile)
    .then(workbook => {
        console.log('\n✅ File loaded successfully!');
        const sheet = workbook.sheet(0);
        console.log('Sheet name:', sheet.name());
        console.log('Used range:', sheet.usedRange()?.address() || 'undefined');
        
        // Check _node structure
        console.log('\n--- Internal structure ---');
        console.log('sheet._node exists:', !!sheet._node);
        if (sheet._node) {
            console.log('sheet._node.children exists:', !!sheet._node.children);
            console.log('sheet._node.children type:', typeof sheet._node.children);
            if (sheet._node.children) {
                console.log('sheet._node.children length:', sheet._node.children.length);
                // Check each child
                sheet._node.children.forEach((child, i) => {
                    console.log(`  child[${i}]:`, child ? child.name : 'undefined');
                });
            }
        }
        
        // Try reading some cells
        console.log('\n--- Cell values ---');
        try {
            console.log('A1:', sheet.cell('A1').value());
            console.log('A2:', sheet.cell('A2').value());
        } catch (e) {
            console.log('Error reading cells:', e.message);
        }
    })
    .catch(err => {
        console.log('\n❌ Error loading file:');
        console.log('Message:', err.message);
        console.log('\n--- Full Stack ---');
        console.log(err.stack);
    });
