// ============================================================================
// EXCELJS MIGRATION - WRITE-FUNKTION
// ============================================================================
// Export/Save-Funktion mit ExcelJS - Test für Formatierungs-Erhaltung

const ExcelJS = require('exceljs');

/**
 * Exportiert/Speichert ein Sheet mit ExcelJS
 * 
 * @param {string} sourcePath - Pfad zur Quelldatei
 * @param {string} targetPath - Pfad zur Zieldatei
 * @param {Object} sheetData - Sheet-Daten mit headers, data, cellStyles, etc.
 * @returns {Promise<Object>} Erfolg/Fehler
 */
async function exportSheetWithExcelJS(sourcePath, targetPath, sheetData) {
    const startTime = Date.now();
    
    try {
        console.log('[ExcelJS Writer] Lade Workbook...');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(sourcePath);
        
        const worksheet = workbook.getWorksheet(sheetData.sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetData.sheetName}" nicht gefunden` };
        }
        
        console.log('[ExcelJS Writer] Aktualisiere Daten...');
        
        // Header aktualisieren (Zeile 1)
        sheetData.headers.forEach((header, colIndex) => {
            const cell = worksheet.getCell(1, colIndex + 1);
            cell.value = header;
        });
        
        // Daten aktualisieren (ab Zeile 2)
        sheetData.data.forEach((row, rowIndex) => {
            row.forEach((value, colIndex) => {
                const cell = worksheet.getCell(rowIndex + 2, colIndex + 1);
                cell.value = value === null || value === undefined ? '' : value;
            });
        });
        
        // Bei fullRewrite: ALLE Styles anwenden
        if (sheetData.fullRewrite === true && sheetData.cellStyles) {
            console.log('[ExcelJS Writer] Full Rewrite - setze alle Styles...');
            
            // Erst alle Styles zurücksetzen
            for (let rowIndex = 0; rowIndex < sheetData.data.length; rowIndex++) {
                for (let colIndex = 0; colIndex < sheetData.headers.length; colIndex++) {
                    const cell = worksheet.getCell(rowIndex + 2, colIndex + 1);
                    
                    // Reset auf Standard
                    cell.font = {};
                    cell.fill = { type: 'pattern', pattern: 'none' };
                }
            }
            
            // Dann neue Styles anwenden
            for (const [styleKey, style] of Object.entries(sheetData.cellStyles)) {
                const [rowIdx, colIdx] = styleKey.split('-').map(Number);
                const cell = worksheet.getCell(rowIdx + 2, colIdx + 1);
                
                // Font-Styles
                const font = {};
                if (style.bold) font.bold = true;
                if (style.italic) font.italic = true;
                if (style.underline) font.underline = true;
                if (style.strikethrough) font.strike = true;
                if (style.fontColor) {
                    // CSS hex zu ARGB konvertieren
                    const hex = style.fontColor.replace('#', '');
                    font.color = { argb: 'FF' + hex };
                }
                
                if (Object.keys(font).length > 0) {
                    cell.font = font;
                }
                
                // Fill/Background
                if (style.fill) {
                    const hex = style.fill.replace('#', '');
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF' + hex }
                    };
                }
            }
        }
        
        // RichText anwenden
        if (sheetData.richTextCells) {
            for (const [styleKey, richText] of Object.entries(sheetData.richTextCells)) {
                const [rowIdx, colIdx] = styleKey.split('-').map(Number);
                const cell = worksheet.getCell(rowIdx + 2, colIdx + 1);
                
                cell.value = {
                    richText: richText.map(part => ({
                        text: part.text,
                        font: {
                            bold: part.styles?.bold || false,
                            italic: part.styles?.italic || false,
                            underline: part.styles?.underline || false,
                            color: part.styles?.color ? { argb: 'FF' + part.styles.color.replace('#', '') } : undefined
                        }
                    }))
                };
            }
        }
        
        // Formeln anwenden
        if (sheetData.cellFormulas) {
            for (const [styleKey, formula] of Object.entries(sheetData.cellFormulas)) {
                const [rowIdx, colIdx] = styleKey.split('-').map(Number);
                const cell = worksheet.getCell(rowIdx + 2, colIdx + 1);
                cell.value = { formula };
            }
        }
        
        // Hyperlinks anwenden
        if (sheetData.cellHyperlinks) {
            for (const [styleKey, hyperlink] of Object.entries(sheetData.cellHyperlinks)) {
                const [rowIdx, colIdx] = styleKey.split('-').map(Number);
                const cell = worksheet.getCell(rowIdx + 2, colIdx + 1);
                cell.value = {
                    text: cell.value || hyperlink,
                    hyperlink: hyperlink
                };
            }
        }
        
        // Versteckte Spalten
        if (sheetData.hiddenColumns && sheetData.hiddenColumns.length > 0) {
            sheetData.hiddenColumns.forEach(colIdx => {
                const col = worksheet.getColumn(colIdx + 1);
                col.hidden = true;
            });
        }
        
        // Versteckte Zeilen
        if (sheetData.hiddenRows && sheetData.hiddenRows.length > 0) {
            sheetData.hiddenRows.forEach(rowIdx => {
                const row = worksheet.getRow(rowIdx + 2); // +2 wegen Header
                row.hidden = true;
            });
        }
        
        // AutoFilter wiederherstellen
        if (sheetData.autoFilterRange) {
            console.log('[ExcelJS Writer] Setze AutoFilter:', sheetData.autoFilterRange);
            worksheet.autoFilter = sheetData.autoFilterRange;
        }
        
        console.log('[ExcelJS Writer] Speichere Datei...');
        await workbook.xlsx.writeFile(targetPath);
        
        const totalTime = Date.now() - startTime;
        console.log(`[ExcelJS Writer] Erfolgreich gespeichert in ${totalTime}ms`);
        
        return {
            success: true,
            message: `Export erfolgreich: ${targetPath}`,
            stats: {
                totalTimeMs: totalTime,
                rows: sheetData.data.length,
                columns: sheetData.headers.length
            }
        };
        
    } catch (error) {
        console.error('[ExcelJS Writer] Fehler:', error);
        return { success: false, error: error.message };
    }
}

module.exports = {
    exportSheetWithExcelJS
};
