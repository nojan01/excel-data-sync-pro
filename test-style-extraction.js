const ExcelJS = require('exceljs');
const path = require('path');

async function analyzeStyleExtraction() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/nojan/Desktop/test-styles.xlsx');
    
    const worksheet = workbook.getWorksheet('Style Tests');
    
    const cellStyles = {};
    
    // Simuliere die Extraktion wie im echten Code
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const colIndex = colNumber - 1;
            const rowIndex = rowNumber - 1;
            const style = {};
            let hasStyle = false;
            
            // Font-Styles
            if (cell.font) {
                if (cell.font.bold) { style.bold = true; hasStyle = true; }
                if (cell.font.italic) { style.italic = true; hasStyle = true; }
                if (cell.font.underline) { style.underline = true; hasStyle = true; }
                if (cell.font.strike) { style.strikethrough = true; hasStyle = true; }
                if (cell.font.size && cell.font.size !== 11) {
                    style.fontSize = cell.font.size;
                    hasStyle = true;
                }
                if (cell.font.color?.argb && cell.font.color.argb !== 'FF000000') {
                    style.fontColor = '#' + cell.font.color.argb.substring(2);
                    hasStyle = true;
                }
            }
            
            // Fill
            if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
                if (cell.fill.fgColor?.argb) {
                    style.fill = '#' + cell.fill.fgColor.argb.substring(2);
                    hasStyle = true;
                }
            }
            
            if (hasStyle) {
                const key = `${rowIndex}-${colIndex}`;
                cellStyles[key] = style;
                
                // Debug: Zeige Styles für Zeilen 8-11 (0-basiert)
                if (rowIndex >= 8 && rowIndex <= 10) {
                    console.log(`[${key}] "${cell.value}" -> Style:`, JSON.stringify(style));
                }
            }
        });
    });
    
    console.log('\n=== ZUSAMMENFASSUNG ===');
    console.log('Gesamt Styles:', Object.keys(cellStyles).length);
    
    // Prüfe spezifisch Zeile 9 (0-basiert = Excel Zeile 10)
    console.log('\n=== ZEILE 9 STYLES (Excel Zeile 10) ===');
    for (let i = 0; i < 6; i++) {
        const key = `9-${i}`;
        if (cellStyles[key]) {
            console.log(`[${key}]:`, JSON.stringify(cellStyles[key]));
        } else {
            console.log(`[${key}]: NICHT VORHANDEN!`);
        }
    }
}

analyzeStyleExtraction();
