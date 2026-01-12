// ============================================================================
// EXCELJS MIGRATION - NEUE READ-FUNKTION
// ============================================================================
// Dieses Modul enthält die ExcelJS-basierte Sheet-Read-Funktion
// Zum Testen der Migration von xlsx-populate zu exceljs

const ExcelJS = require('exceljs');

/**
 * Liest ein Excel-Sheet mit ExcelJS (Alternative zu xlsx-populate)
 * 
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des zu lesenden Sheets
 * @param {string|null} password - Optional: Passwort für geschützte Dateien
 * @returns {Promise<Object>} Sheet-Daten im gleichen Format wie xlsx-populate
 */
async function readSheetWithExcelJS(filePath, sheetName, password = null) {
    const startTime = Date.now();
    
    try {
        const workbook = new ExcelJS.Workbook();
        
        console.log('[ExcelJS] Lade Workbook...');
        const loadStart = Date.now();
        
        // Datei laden (mit oder ohne Passwort)
        if (password) {
            await workbook.xlsx.readFile(filePath, { password });
        } else {
            await workbook.xlsx.readFile(filePath);
        }
        
        console.log(`[ExcelJS] Workbook geladen in ${Date.now() - loadStart}ms`);
        
        // Sheet finden
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Daten-Strukturen initialisieren
        const headers = [];
        const data = [];
        const hiddenColumns = [];
        const hiddenRows = [];
        const cellStyles = {};
        const cellFormulas = {};
        const cellHyperlinks = {};
        const richTextCells = {};
        
        // AutoFilter-Bereich
        let autoFilterRange = worksheet.autoFilter ? worksheet.autoFilter.ref : null;
        
        // Versteckte Spalten ermitteln
        worksheet.columns.forEach((col, colIndex) => {
            if (col.hidden) {
                hiddenColumns.push(colIndex);
            }
        });
        
        // Zeilen durchgehen
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            // Erste Zeile = Header
            if (rowNumber === 1) {
                row.eachCell({ includeEmpty: true }, (cell) => {
                    headers.push(cell.value ? String(cell.value) : '');
                });
                return; // Weiter zur nächsten Zeile
            }
            
            // Daten-Zeilen
            const rowData = [];
            
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const colIndex = colNumber - 1;
                const dataRowIndex = rowNumber - 2; // -1 für Header, -1 für 0-basiert
                // WICHTIG: Frontend erwartet 1-basierte Indizes (wie xlsx-populate)
                const styleKey = `${dataRowIndex + 1}-${colIndex}`;
                
                let cellValue = cell.value;
                
                // Formel extrahieren
                if (cell.formula) {
                    cellFormulas[styleKey] = cell.formula;
                    cellValue = cell.result || cellValue;
                }
                
                // Hyperlink extrahieren
                if (cell.hyperlink) {
                    cellHyperlinks[styleKey] = cell.hyperlink.hyperlink || cell.hyperlink;
                }
                
                // Rich Text extrahieren
                if (cell.value && typeof cell.value === 'object' && cell.value.richText) {
                    const richText = cell.value.richText.map(part => ({
                        text: part.text,
                        styles: {
                            bold: part.font?.bold || false,
                            italic: part.font?.italic || false,
                            underline: part.font?.underline || false,
                            color: part.font?.color?.argb ? `#${part.font.color.argb.substring(2)}` : null
                        }
                    }));
                    richTextCells[styleKey] = richText;
                    cellValue = richText.map(r => r.text).join('');
                }
                
                // Styles extrahieren
                const style = {};
                
                if (cell.font) {
                    if (cell.font.bold) style.bold = true;
                    if (cell.font.italic) style.italic = true;
                    if (cell.font.underline) style.underline = true;
                    if (cell.font.strike) style.strikethrough = true;
                    if (cell.font.color?.argb) {
                        style.fontColor = `#${cell.font.color.argb.substring(2)}`;
                    }
                }
                
                if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor?.argb) {
                    style.fill = `#${cell.fill.fgColor.argb.substring(2)}`;
                }
                
                if (Object.keys(style).length > 0) {
                    cellStyles[styleKey] = style;
                }
                
                // Datum formatieren
                if (cellValue instanceof Date) {
                    cellValue = cellValue.toISOString().split('T')[0];
                }
                
                rowData.push(cellValue === null || cellValue === undefined ? '' : cellValue);
            });
            
            // Versteckte Zeilen
            if (row.hidden) {
                hiddenRows.push(rowNumber - 2); // -2 wegen Header und 0-basiert
            }
            
            data.push(rowData);
        });
        
        const totalTime = Date.now() - startTime;
        console.log(`[ExcelJS] Sheet geladen in ${totalTime}ms (${data.length} Zeilen)`);
        
        return {
            success: true,
            headers,
            data,
            hiddenColumns,
            hiddenRows,
            cellStyles,
            cellFormulas,
            cellHyperlinks,
            richTextCells,
            autoFilterRange,
            stats: {
                rows: data.length,
                columns: headers.length,
                loadTimeMs: totalTime
            }
        };
        
    } catch (error) {
        console.error('[ExcelJS] Fehler beim Laden:', error);
        return { success: false, error: error.message };
    }
}

module.exports = {
    readSheetWithExcelJS
};
