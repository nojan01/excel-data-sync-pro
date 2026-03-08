// ============================================================================
// EXCELJS MIGRATION - NEUE READ-FUNKTION
// ============================================================================
// Dieses Modul enthält die ExcelJS-basierte Sheet-Read-Funktion
// Zum Testen der Migration von xlsx-populate zu exceljs

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const AdmZip = require('adm-zip');
let XlsxPopulate = null; // Lazy-load für Passwort-Entschlüsselung

// Excel Standard-Theme-Farben (Office-Theme)
// Index 0-9 entspricht den Theme-Colors dk1, lt1, dk2, lt2, accent1-6
const THEME_COLORS = [
    '000000', // 0: dk1 (schwarz)
    'FFFFFF', // 1: lt1 (weiß)
    '44546A', // 2: dk2
    'E7E6E6', // 3: lt2
    '4472C4', // 4: accent1 (blau)
    'ED7D31', // 5: accent2 (orange)
    'A5A5A5', // 6: accent3 (grau)
    'FFC000', // 7: accent4 (gold)
    '5B9BD5', // 8: accent5 (hellblau)
    '70AD47', // 9: accent6 (grün)
];

/**
 * Löst eine ExcelJS-Farbangabe zu einem Hex-String auf.
 * Unterstützt ARGB, Theme-Farben (mit Tint) und Indexed-Farben.
 * @param {Object} color - ExcelJS color object { argb, theme, tint, indexed }
 * @returns {string|null} Hex-Farbstring (z.B. '#FF0000') oder null
 */
function resolveColor(color) {
    if (!color) return null;
    
    // 1. ARGB-Wert (häufigstes Format)
    if (color.argb) {
        const hex = color.argb.length === 8 ? color.argb.substring(2) : color.argb;
        if (hex !== '000000') return `#${hex}`;
        return null;
    }
    
    // 2. Theme-Farbe (z.B. { theme: 4, tint: -0.25 })
    if (color.theme !== undefined && color.theme !== null) {
        let hex = THEME_COLORS[color.theme] || '000000';
        
        // Tint anwenden (aufhellen/abdunkeln)
        if (color.tint) {
            hex = applyTint(hex, color.tint);
        }
        
        if (hex !== '000000') return `#${hex}`;
        return null;
    }
    
    // 3. Indexed-Farbe (Legacy-Format)
    if (color.indexed !== undefined && color.indexed !== null) {
        const INDEXED_COLORS = [
            '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
            '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
            '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
            '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
            '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
            '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
            '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
            '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
        ];
        const idx = color.indexed;
        if (idx >= 0 && idx < INDEXED_COLORS.length) {
            const hex = INDEXED_COLORS[idx];
            if (hex !== '000000') return `#${hex}`;
        }
        return null;
    }
    
    return null;
}

/**
 * Wendet einen Tint-Wert auf eine Hex-Farbe an.
 * Positiver Tint = aufhellen (Richtung Weiß), negativer Tint = abdunkeln (Richtung Schwarz).
 * @param {string} hex - 6-stelliger Hex-String (ohne #)
 * @param {number} tint - Wert zwischen -1.0 und 1.0
 * @returns {string} Angepasster 6-stelliger Hex-String
 */
function applyTint(hex, tint) {
    let r = parseInt(hex.substring(0, 2), 16);
    let g = parseInt(hex.substring(2, 4), 16);
    let b = parseInt(hex.substring(4, 6), 16);
    
    if (tint > 0) {
        // Aufhellen: Richtung Weiß
        r = Math.round(r + (255 - r) * tint);
        g = Math.round(g + (255 - g) * tint);
        b = Math.round(b + (255 - b) * tint);
    } else {
        // Abdunkeln: Richtung Schwarz
        r = Math.round(r * (1 + tint));
        g = Math.round(g * (1 + tint));
        b = Math.round(b * (1 + tint));
    }
    
    r = Math.max(0, Math.min(255, r));
    g = Math.max(0, Math.min(255, g));
    b = Math.max(0, Math.min(255, b));
    
    return r.toString(16).padStart(2, '0') + g.toString(16).padStart(2, '0') + b.toString(16).padStart(2, '0');
}

/**
 * Formatiert ein JS Date-Objekt gemäß einem Excel numFmt-String.
 * Parst die Excel-Format-Tokens (d, dd, m, mm, mmm, mmmm, yy, yyyy, h, hh, s, ss)
 * und gibt den formatierten String zurück — exakt wie Excel es anzeigen würde.
 * 
 * @param {Date} dt - Das Date-Objekt
 * @param {string} numFmt - Excel-Format-String (z.B. "dd.mm.yyyy", "m/d/yy h:mm")
 * @returns {string} Formatierter Datum-String
 */
function formatDateWithNumFmt(dt, numFmt) {
    if (!numFmt || numFmt === 'General') {
        // Kein Format: deutsches Standard-Datum
        const d = dt.getDate(), m = dt.getMonth() + 1, y = dt.getFullYear();
        return `${String(d).padStart(2, '0')}.${String(m).padStart(2, '0')}.${y}`;
    }

    // Nur den Datum/Zeit-Teil verwenden (vor Semikolon = positiver Format-Teil)
    let fmt = numFmt.split(';')[0];
    
    // Locale-Code aus [$-XXX] extrahieren BEVOR wir Klammern entfernen
    // z.B. [$-407] = Deutsch, [$-409] = US-Englisch, [$-809] = UK-Englisch
    let dateSeparator = '.'; // Standard: deutsch
    const localeMatch = fmt.match(/\[\$-([0-9A-Fa-f]+)\]/);
    if (localeMatch) {
        const localeId = parseInt(localeMatch[1], 16) & 0xFFFF; // Nur Language-ID
        // US-English (0x409), Filipino (0x464), etc. verwenden /
        const slashLocales = [0x409, 0x809, 0x0C09, 0x1009, 0x1409, 0x464];
        // Französisch etc. verwenden /
        const frenchLocales = [0x40C, 0x80C, 0x100C];
        if (slashLocales.includes(localeId) || frenchLocales.includes(localeId)) {
            dateSeparator = '/';
        } else if (localeId === 0x410) { // Italienisch
            dateSeparator = '/';
        }
        // Deutsch (0x407, 0x807, 0xC07) und die meisten anderen → '.'
    }
    
    // Escaped Text in eckigen Klammern entfernen (z.B. [$-407], [$€-de-DE])
    fmt = fmt.replace(/\[([^\]]*)\]/g, '');
    
    // In Excel ist '/' ein Platzhalter für den lokalen Datumstrenner — KEIN Literal!
    // Ersetze '/' durch den erkannten lokalen Separator
    fmt = fmt.replace(/\//g, dateSeparator);
    
    // Literal-Strings in Anführungszeichen durch Platzhalter ersetzen
    const literals = [];
    fmt = fmt.replace(/"([^"]*)"/g, (_, text) => {
        literals.push(text);
        return `\x00LIT${literals.length - 1}\x00`;
    });
    // Backslash-Escapes: \X → Literal X
    fmt = fmt.replace(/\\(.)/g, (_, ch) => {
        literals.push(ch);
        return `\x00LIT${literals.length - 1}\x00`;
    });

    const day = dt.getDate();
    const month = dt.getMonth() + 1;
    const year = dt.getFullYear();
    const hours24 = dt.getHours();
    const minutes = dt.getMinutes();
    const seconds = dt.getSeconds();

    // AM/PM-Erkennung
    const hasAMPM = /am\/pm|AM\/PM|a\/p|A\/P/i.test(fmt);
    const hours12 = hasAMPM ? (hours24 % 12 || 12) : hours24;
    const ampm = hours24 < 12 ? 'AM' : 'PM';

    // Token-basiertes Ersetzen (längste Matches zuerst)
    // 'm' nach 'h' oder vor 's' = Minute, sonst = Monat
    // Wir müssen die Reihenfolge der Tokens im Format-String beachten
    const monthNames = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];
    const monthNamesFull = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];

    // Tokenize: Finde alle Format-Tokens und deren Positionen
    const tokenRegex = /mmmm|mmm|mm|m|dddd|ddd|dd|d|yyyy|yy|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p/gi;
    let result = '';
    let lastIndex = 0;
    let afterHour = false; // Ist das nächste 'm' eine Minute?
    
    // Erst mal alle Tokens sammeln um den Kontext (h vor m) zu kennen
    const tokens = [];
    let match;
    while ((match = tokenRegex.exec(fmt)) !== null) {
        tokens.push({ token: match[0], index: match.index, end: match.index + match[0].length });
    }

    // Bestimme für jedes 'm'-Token ob es Monat oder Minute ist
    for (let i = 0; i < tokens.length; i++) {
        const t = tokens[i].token.toLowerCase();
        if (t === 'm' || t === 'mm') {
            // Minute wenn: vorheriges Token ist h/hh ODER nächstes Token ist s/ss
            const prev = i > 0 ? tokens[i - 1].token.toLowerCase() : '';
            const next = i < tokens.length - 1 ? tokens[i + 1].token.toLowerCase() : '';
            tokens[i].isMinute = (prev === 'h' || prev === 'hh' || next === 's' || next === 'ss');
        }
    }

    // Jetzt den Format-String zusammenbauen
    for (let i = 0; i < tokens.length; i++) {
        const t = tokens[i];
        // Text vor diesem Token
        result += fmt.substring(lastIndex, t.index);
        lastIndex = t.end;

        const lower = t.token.toLowerCase();
        switch (lower) {
            case 'yyyy': result += String(year); break;
            case 'yy':   result += String(year).substring(2); break;
            case 'mmmm': result += monthNamesFull[month - 1]; break;
            case 'mmm':  result += monthNames[month - 1]; break;
            case 'mm':
                if (t.isMinute) {
                    result += String(minutes).padStart(2, '0');
                } else {
                    result += String(month).padStart(2, '0');
                }
                break;
            case 'm':
                if (t.isMinute) {
                    result += String(minutes);
                } else {
                    result += String(month);
                }
                break;
            case 'dddd': {
                const dayNames = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
                result += dayNames[dt.getDay()];
                break;
            }
            case 'ddd': {
                const dayNamesShort = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
                result += dayNamesShort[dt.getDay()];
                break;
            }
            case 'dd':   result += String(day).padStart(2, '0'); break;
            case 'd':    result += String(day); break;
            case 'hh':   result += String(hasAMPM ? hours12 : hours24).padStart(2, '0'); break;
            case 'h':    result += String(hasAMPM ? hours12 : hours24); break;
            case 'ss':   result += String(seconds).padStart(2, '0'); break;
            case 's':    result += String(seconds); break;
            case 'am/pm': case 'a/p':
                result += t.token === t.token.toUpperCase() ? ampm : ampm.toLowerCase();
                break;
            default:     result += t.token; break;
        }
    }
    // Rest nach dem letzten Token
    result += fmt.substring(lastIndex);

    // Literal-Platzhalter zurückersetzen
    result = result.replace(/\x00LIT(\d+)\x00/g, (_, idx) => literals[parseInt(idx)]);

    // Bereinigung: Führende/angehängte Leerzeichen trimmen
    return result.trim();
}

/**
 * Formatiert einen numerischen Wert gemäß dem Excel numFmt-String.
 * Buchhaltungs-/Währungsformate wie _-* #.##0,00 "€"_- oder #,##0.00 speichern
 * intern volle Float-Präzision (z.B. 95.89556867501796), zeigen aber gerundet an (95,20 €).
 * Diese Funktion erkennt Nachkommastellen, Tausendertrenner, Währungssymbole und
 * gibt einen fertig formatierten String zurück.
 *
 * @param {number} value - Der numerische Zellwert
 * @param {string} numFmt - Excel-Format-String (z.B. '#,##0.00', '_-* #.##0,00 "€"_-')
 * @returns {string|number} Formatierter Wert als String (z.B. "95,20 €") oder Originalwert
 */
function roundNumericByFormat(value, numFmt) {
    if (!numFmt || numFmt === 'General' || typeof value !== 'number' || !isFinite(value)) {
        return value;
    }

    // --- 1. Währungssymbol extrahieren (VOR dem Bereinigen!) ---
    let currencySymbol = '';
    let currencyPosition = 'suffix'; // 'prefix' oder 'suffix'
    
    // Nur den positiven Format-Teil verwenden (vor erstem Semikolon)
    const originalFmt = numFmt.split(';')[0];
    
    // Währungssymbol aus "€", "CHF", "$", etc. in Anführungszeichen
    const quotedMatch = originalFmt.match(/"([^"]*[€$£¥₹₽CHFkr].*?)"|"(.*?[€$£¥₹₽].*?)"/i);
    if (quotedMatch) {
        currencySymbol = (quotedMatch[1] || quotedMatch[2]).trim();
    }
    // Währungssymbol aus [$€-de-DE] oder [$€] oder [$CHF-...] Locale-Codes
    if (!currencySymbol) {
        const localeMatch = originalFmt.match(/\[\$([^\-\]]+)/);
        if (localeMatch) {
            currencySymbol = localeMatch[1].trim();
        }
    }
    // Einzelnes Währungszeichen ohne Quotes (z.B. #,##0.00€ oder €#,##0.00)
    if (!currencySymbol) {
        const bareMatch = originalFmt.match(/([€$£¥₹₽])/);
        if (bareMatch) {
            currencySymbol = bareMatch[1];
        }
    }
    
    // Position bestimmen: Symbol VOR oder NACH der Zahl?
    if (currencySymbol) {
        // Finde Position des Symbols relativ zu den Ziffernplatzhaltern
        const symbolPos = originalFmt.indexOf(currencySymbol.charAt(0));
        const firstDigit = originalFmt.search(/[0#]/);
        if (symbolPos >= 0 && firstDigit >= 0 && symbolPos < firstDigit) {
            currencyPosition = 'prefix';
        }
    }

    // --- 2. Format bereinigen für Dezimalstellen-Erkennung ---
    let fmt = originalFmt;
    fmt = fmt.replace(/"[^"]*"/g, '');
    fmt = fmt.replace(/\\./g, '');
    fmt = fmt.replace(/\[[^\]]*\]/g, '');
    fmt = fmt.replace(/_./g, '');
    fmt = fmt.replace(/\*./g, '');

    // --- 3. Tausendertrenner erkennen ---
    // International: #,##0 (Komma = Tausender, Punkt = Dezimal)
    // Deutsch:       #.##0 (Punkt = Tausender, Komma = Dezimal)
    let useThousandSep = false;
    let isGermanFormat = false;
    
    // Deutsches Format: Punkt als Tausendertrenner VOR Komma als Dezimaltrenner
    if (/[0#]\.##[0#]/.test(fmt) || /[0#]\.[0#]{3}/.test(fmt)) {
        // z.B. #.##0,00 → Punkt ist Tausender
        if (fmt.includes(',')) {
            isGermanFormat = true;
            useThousandSep = true;
        }
    }
    // Internationales Format: Komma als Tausendertrenner
    if (/[0#],##[0#]/.test(fmt) || /[0#],[0#]{3}/.test(fmt)) {
        if (!isGermanFormat) {
            useThousandSep = true;
        }
    }

    // --- 4. Prozent-Format ---
    if (fmt.includes('%')) {
        const percentMatch = fmt.match(/[0#]\.(0+)\s*%/) || fmt.match(/\.(0+)/);
        if (percentMatch) {
            const decimals = percentMatch[1].length;
            const factor = Math.pow(10, decimals + 2);
            return Math.round(value * factor) / factor;
        }
        return value;
    }

    // --- 5. Dezimalstellen erkennen und Wert formatieren ---
    let decimals = -1;
    
    // Standard-Format (Punkt als Dezimaltrenner): 0.00, #,##0.00
    let decimalMatch = fmt.match(/\.(0+)(?:[^0#]|$)/);
    if (decimalMatch && !isGermanFormat) {
        decimals = decimalMatch[1].length;
    }

    // Deutsches Format (Komma als Dezimaltrenner): #.##0,00
    if (decimals < 0) {
        decimalMatch = fmt.match(/,(0+)(?:[^0#]|$)/);
        if (decimalMatch) {
            const commaPos = fmt.indexOf(',');
            const dotBefore = fmt.lastIndexOf('.', commaPos);
            if (dotBefore >= 0) {
                isGermanFormat = true;
                useThousandSep = true;
                decimals = decimalMatch[1].length;
            }
        }
    }
    
    // Ganzzahl-Format (#,##0 oder 0)
    if (decimals < 0 && /[0#]/.test(fmt) && /0/.test(fmt)) {
        if (!fmt.includes('.') && !/,(0+)/.test(fmt)) {
            decimals = 0;
        }
    }

    if (decimals < 0) {
        return value;
    }

    // --- 6. Zahl formatieren ---
    const isNegative = value < 0;
    const absValue = Math.abs(value);
    const rounded = absValue.toFixed(decimals);
    
    let result;
    if (useThousandSep) {
        // Tausendertrenner einfügen
        const parts = rounded.split('.');
        const intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, isGermanFormat ? '.' : ',');
        if (decimals > 0) {
            const decSep = isGermanFormat ? ',' : '.';
            result = intPart + decSep + parts[1];
        } else {
            result = intPart;
        }
    } else {
        // Ohne Tausendertrenner, aber ggf. deutschen Dezimaltrenner
        if (isGermanFormat && decimals > 0) {
            result = rounded.replace('.', ',');
        } else {
            result = rounded;
        }
    }
    
    // Vorzeichen
    if (isNegative) {
        result = '-' + result;
    }
    
    // Währungssymbol anfügen
    if (currencySymbol) {
        if (currencyPosition === 'prefix') {
            result = currencySymbol + ' ' + result;
        } else {
            result = result + ' ' + currencySymbol;
        }
    }
    
    return result;
}

/**
 * Konvertiert Spalten-Buchstaben zu Index (A=0, B=1, ..., Z=25, AA=26, ...)
 * @param {string} letters - Spalten-Buchstaben (z.B. "A", "AA", "BC")
 * @returns {number} 0-basierter Spalten-Index
 */
function colLettersToIndex(letters) {
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
        index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index - 1; // 0-basiert
}

/**
 * Parst einen Range-String (z.B. "A1:H1") zu einem Objekt
 * @param {string} rangeStr - Range im Format "A1:H1"
 * @returns {Object|null} { startRow, startCol, endRow, endCol, rowSpan, colSpan }
 */
function parseRangeString(rangeStr) {
    const match = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
        console.warn(`[ExcelJS] Ungültiger Range-String: ${rangeStr}`);
        return null;
    }
    
    const startCol = colLettersToIndex(match[1]);
    const startRow = parseInt(match[2]) - 1; // 0-basiert
    const endCol = colLettersToIndex(match[3]);
    const endRow = parseInt(match[4]) - 1; // 0-basiert
    
    return {
        startRow,
        startCol,
        endRow,
        endCol,
        rowSpan: endRow - startRow + 1,
        colSpan: endCol - startCol + 1
    };
}

/**
 * Erkennt Zeilenfarben basierend auf cellStyles.
 * Eine Zeile wird als "markiert" erkannt, wenn ALLE Zellen die gleiche Hintergrundfarbe haben
 * und diese Farbe einer der bekannten Highlight-Farben entspricht.
 * 
 * @param {Object} cellStyles - Map von "rowIndex-colIndex" zu Style-Objekt mit fill
 * @param {number} rowCount - Anzahl der Datenzeilen
 * @param {number} colCount - Anzahl der Spalten
 * @returns {Array} Array von [rowIndex, colorName] Paaren
 */
function detectRowHighlights(cellStyles, rowCount, colCount) {
    const highlights = [];
    
    // Mapping von ARGB-Farben zu Highlight-Namen (ohne Alpha-Kanal)
    const colorMapping = {
        '90EE90': 'green',   // Light Green
        'FFFF00': 'yellow',  // Yellow
        'FFA500': 'orange',  // Orange
        'FF6B6B': 'red',     // Light Red
        '87CEEB': 'blue',    // Sky Blue
        'DDA0DD': 'purple',  // Plum
        // Alternative Farben die auch erkannt werden sollen
        '4CAF50': 'green',
        'FFEB3B': 'yellow',
        'FF9800': 'orange',
        'F44336': 'red',
        '2196F3': 'blue',
        '9C27B0': 'purple'
    };
    
    // Für jede Zeile prüfen
    for (let rowIdx = 0; rowIdx < rowCount; rowIdx++) {
        const rowFills = [];
        
        // Alle Zellen in der Zeile durchgehen
        for (let colIdx = 0; colIdx < colCount; colIdx++) {
            // cellStyles-Key Format: "rowIndex-colIndex" wobei rowIndex 1-basiert ist für Datenzeilen
            const styleKey = `${rowIdx + 1}-${colIdx}`;
            const style = cellStyles[styleKey];
            
            if (style && style.fill) {
                // Fill ist im Format "#RRGGBB"
                const fillHex = style.fill.replace('#', '').toUpperCase();
                rowFills.push(fillHex);
            } else {
                rowFills.push(null);
            }
        }
        
        // Prüfen ob alle Zellen die gleiche Farbe haben (und nicht null)
        const nonNullFills = rowFills.filter(f => f !== null);
        if (nonNullFills.length === colCount && nonNullFills.length > 0) {
            const firstFill = nonNullFills[0];
            const allSame = nonNullFills.every(f => f === firstFill);
            
            if (allSame) {
                // Prüfen ob die Farbe einer bekannten Highlight-Farbe entspricht
                const colorName = colorMapping[firstFill];
                if (colorName) {
                    highlights.push([rowIdx, colorName]);
                }
            }
        }
    }
    
    return highlights;
}

/**
 * Extrahiert Fill-Farben direkt aus der XLSX-Datei (ZIP-Format).
 * Dies ist ein Workaround für ExcelJS, das bei bestimmten Excel-Dateien
 * (z.B. von SoftMaker/PlanMaker erstellt) keine Fills erkennt.
 * 
 * @param {string} filePath - Pfad zur XLSX-Datei
 * @param {string} sheetName - Name des Sheets
 * @returns {Object} Map von "rowNumber-colNumber" zu Fill-Farbe (z.B. "#FF0000")
 */
function extractFillsFromXLSX(fileBufferOrPath, sheetName, existingZip = null) {
    const cellFills = {};
    
    try {
        // Existierendes ZIP-Objekt wiederverwenden oder neues erstellen
        const zip = existingZip || (Buffer.isBuffer(fileBufferOrPath) ? new AdmZip(fileBufferOrPath) : new AdmZip(fileBufferOrPath));
        
        // 1. styles.xml lesen und Fills extrahieren
        const stylesEntry = zip.getEntry('xl/styles.xml');
        if (!stylesEntry) {
            return cellFills;
        }
        
        const stylesXml = stylesEntry.getData().toString('utf8');
        
        // Fills extrahieren (Position im Array = fillId)
        const fills = [];
        const fillsMatch = stylesXml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/);
        if (fillsMatch) {
            const fillPattern = /<fill[^\/]*>([\s\S]*?)<\/fill>/g;
            let fillMatch;
            while ((fillMatch = fillPattern.exec(fillsMatch[1])) !== null) {
                const fillContent = fillMatch[1];
                // Suche nach fgColor mit rgb-Attribut
                const fgColorMatch = fillContent.match(/<fgColor[^>]*rgb="([A-Fa-f0-9]{8})"[^>]*\/?>/);
                if (fgColorMatch) {
                    const argb = fgColorMatch[1];
                    // ARGB zu RGB (erste 2 Zeichen = Alpha)
                    const rgb = argb.substring(2);
                    fills.push('#' + rgb);
                } else {
                    fills.push(null);
                }
            }
        }
        
        // 2. cellXfs extrahieren (Style ID -> Fill ID Mapping)
        const styleToFill = [];
        const cellXfsMatch = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
        if (cellXfsMatch) {
            const xfPattern = /<xf[^>]*>/g;
            let xfMatch;
            while ((xfMatch = xfPattern.exec(cellXfsMatch[1])) !== null) {
                const xfContent = xfMatch[0];
                const fillIdMatch = xfContent.match(/fillId="(\d+)"/);
                const applyFillMatch = xfContent.match(/applyFill="(\d+)"/);
                
                if (fillIdMatch) {
                    const fillId = parseInt(fillIdMatch[1]);
                    // applyFill muss 1 sein oder nicht vorhanden (dann gilt fillId)
                    if (!applyFillMatch || applyFillMatch[1] === '1') {
                        styleToFill.push(fillId);
                    } else {
                        styleToFill.push(null);
                    }
                } else {
                    styleToFill.push(null);
                }
            }
        }
        
        // 3. Sheet-Daten finden
        // Zuerst workbook.xml lesen um Sheet rId zu finden
        const workbookEntry = zip.getEntry('xl/workbook.xml');
        if (!workbookEntry) {
            return cellFills;
        }
        
        const workbookXml = workbookEntry.getData().toString('utf8');
        const sheetsMatch = workbookXml.match(/<sheets>([\s\S]*?)<\/sheets>/);
        let sheetRId = 'rId1'; // Default
        
        if (sheetsMatch) {
            const sheetPattern = /<sheet[^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"[^>]*\/?>/g;
            let sheetMatch;
            while ((sheetMatch = sheetPattern.exec(sheetsMatch[1])) !== null) {
                if (sheetMatch[1] === sheetName) {
                    sheetRId = sheetMatch[2];
                    break;
                }
            }
        }
        
        // Relationship-Datei lesen um tatsächlichen Sheet-Pfad zu finden
        const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
        if (!relsEntry) {
            return cellFills;
        }
        
        const relsXml = relsEntry.getData().toString('utf8');
        let sheetPath = null;
        
        // Attribut-Reihenfolge-unabhängig: openpyxl schreibt Target vor Id
        const relElements = relsXml.match(/<Relationship\s[^>]*\/?>/g) || [];
        for (const relEl of relElements) {
            const idMatch = relEl.match(/Id="([^"]*)"/);
            const targetMatch = relEl.match(/Target="([^"]*)"/);
            if (idMatch && targetMatch && idMatch[1] === sheetRId) {
                sheetPath = targetMatch[1];
                break;
            }
        }
        
        if (!sheetPath) {
            return cellFills;
        }
        
        // Sheet-XML laden (Pfad kann relativ sein, z.B. "worksheets/sheet1.xml")
        const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : `xl/${sheetPath}`;
        const sheetEntry = zip.getEntry(fullSheetPath);
        if (!sheetEntry) {
            return cellFills;
        }
        
        const sheetXml = sheetEntry.getData().toString('utf8');
        
        // 4. Zellen mit Style-IDs extrahieren
        const cellPattern = /<c r="([A-Z]+)(\d+)"[^>]*s="(\d+)"[^>]*>/g;
        let cellMatch;
        while ((cellMatch = cellPattern.exec(sheetXml)) !== null) {
            const colLetters = cellMatch[1];
            const rowNum = parseInt(cellMatch[2]);
            const styleId = parseInt(cellMatch[3]);
            
            // Spalten-Buchstaben zu Index konvertieren (A=0, B=1, ...)
            let colIndex = 0;
            for (let i = 0; i < colLetters.length; i++) {
                colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 64);
            }
            colIndex--; // 0-basiert
            
            // Fill-Farbe ermitteln
            const fillId = styleToFill[styleId];
            if (fillId !== null && fillId !== undefined && fills[fillId]) {
                const fillColor = fills[fillId];
                // Ignoriere Weiß
                if (fillColor !== '#FFFFFF') {
                    // Key: "rowNumber-colIndex" (rowNumber ist 1-basiert, colIndex 0-basiert)
                    // Für Daten-Zeilen (ab Zeile 2) wird rowNumber-1 als Key verwendet
                    // da das Frontend 1-basierte Indizes verwendet
                    const dataRowIndex = rowNum - 1; // Zeile 2 = Datenzeile 1
                    const key = `${dataRowIndex}-${colIndex}`;
                    cellFills[key] = fillColor;
                }
            }
        }
        
        return cellFills;
        
    } catch (error) {
        console.error('[XLSX-Extract] Fehler:', error);
        return cellFills;
    }
}



/**
 * Extrahiert Sheet-Metadaten direkt aus der XLSX-Datei (ZIP-Format).
 * Schnell (~10ms), kein Zell-Parsing, blockiert den Event-Loop nicht nennenswert.
 * Liefert: Spaltenanzahl, versteckte Spalten, verbundene Zellen, AutoFilter.
 */
function extractSheetMetadata(fileBufferOrPath, sheetName, existingZip = null) {
    const result = {
        columnCount: 1,
        hiddenColumns: [],
        hiddenRows: [],  // 0-basierte Daten-Indizes (ohne Header)
        mergedCells: [],
        autoFilterRange: null,
        imageCells: []  // Zellen die Bild-Formeln enthalten (DISPIMG, IMAGE)
    };
    
    try {
        // Existierendes ZIP-Objekt wiederverwenden oder neues erstellen
        const zip = existingZip || (Buffer.isBuffer(fileBufferOrPath) ? new AdmZip(fileBufferOrPath) : new AdmZip(fileBufferOrPath));
        
        // workbook.xml lesen um Sheet-rId zu finden
        const workbookEntry = zip.getEntry('xl/workbook.xml');
        if (!workbookEntry) return result;
        
        const workbookXml = workbookEntry.getData().toString('utf8');
        const sheetsMatch = workbookXml.match(/<sheets>([\s\S]*?)<\/sheets>/);
        let sheetRId = 'rId1';
        
        if (sheetsMatch) {
            const sheetPattern = /<sheet[^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"[^>]*\/?>/g;
            let m;
            while ((m = sheetPattern.exec(sheetsMatch[1])) !== null) {
                const name = m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
                if (name === sheetName) {
                    sheetRId = m[2];
                    break;
                }
            }
        }
        
        // Rels lesen für Sheet-Pfad
        const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
        if (!relsEntry) return result;
        
        const relsXml = relsEntry.getData().toString('utf8');
        let sheetPath = null;
        
        // Attribut-Reihenfolge-unabhängig: openpyxl schreibt Target vor Id
        const relElements = relsXml.match(/<Relationship\s[^>]*\/?>/g) || [];
        for (const relEl of relElements) {
            const idMatch = relEl.match(/Id="([^"]*)"/);
            const targetMatch = relEl.match(/Target="([^"]*)"/);
            if (idMatch && targetMatch && idMatch[1] === sheetRId) {
                sheetPath = targetMatch[1];
                break;
            }
        }
        
        if (!sheetPath) return result;
        
        const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : `xl/${sheetPath}`;
        const sheetEntry = zip.getEntry(fullSheetPath);
        if (!sheetEntry) return result;
        
        const sheetXml = sheetEntry.getData().toString('utf8');
        
        // Dimension (Spaltenanzahl)
        const dimMatch = sheetXml.match(/<dimension ref="[A-Z]+\d+:([A-Z]+)\d+"/);
        if (dimMatch) {
            result.columnCount = colLettersToIndex(dimMatch[1]) + 1;
        }
        
        // Versteckte Spalten aus <cols>
        const colsMatch = sheetXml.match(/<cols>([\s\S]*?)<\/cols>/);
        if (colsMatch) {
            const colPattern = /<col\s([^>]*)\/?\s*>/g;
            let colM;
            while ((colM = colPattern.exec(colsMatch[1])) !== null) {
                const attrs = colM[1];
                if (/hidden="(1|true)"/i.test(attrs)) {
                    const minMatch = attrs.match(/min="(\d+)"/);
                    const maxMatch = attrs.match(/max="(\d+)"/);
                    if (minMatch && maxMatch) {
                        for (let i = parseInt(minMatch[1]) - 1; i < parseInt(maxMatch[1]); i++) {
                            result.hiddenColumns.push(i);
                        }
                    }
                }
            }
        }
        
        // Versteckte Zeilen aus <row> Elementen in <sheetData>
        // WICHTIG: ExcelJS Streaming Reader ignoriert row.hidden, daher hier extrahieren
        const rowPattern = /<row\s([^>]*)>/g;
        let rowM;
        while ((rowM = rowPattern.exec(sheetXml)) !== null) {
            const attrs = rowM[1];
            if (/hidden="(1|true)"/i.test(attrs)) {
                const rMatch = attrs.match(/\br="(\d+)"/);
                if (rMatch) {
                    const excelRow = parseInt(rMatch[1]);
                    if (excelRow >= 2) {  // Zeile 1 = Header, ab Zeile 2 = Daten
                        result.hiddenRows.push(excelRow - 2);  // 0-basierter Daten-Index
                    }
                }
            }
        }

        // AutoFilter direkt aus Sheet-XML
        const afMatch = sheetXml.match(/<autoFilter[^>]*ref="([^"]*)"/);
        if (afMatch) {
            result.autoFilterRange = afMatch[1];
        }
        
        // Verbundene Zellen aus <mergeCells>
        const mcMatch = sheetXml.match(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/);
        if (mcMatch) {
            const mcPattern = /<mergeCell\s+ref="([^"]*)"\s*\/?>/g;
            let mcM;
            while ((mcM = mcPattern.exec(mcMatch[1])) !== null) {
                const parsed = parseRangeString(mcM[1]);
                if (parsed) result.mergedCells.push(parsed);
            }
        }
        
        // ================================================================
        // Bild-Zellen erkennen
        // Excel 365 speichert Zellbilder auf 2 Arten:
        // 1) DISPIMG/IMAGE Formeln in <f>-Tags (ältere Methode)
        // 2) vm-Attribut (Value Metadata) auf <c>-Tags + richData (moderne Methode)
        // PERFORMANCE: Schnell-Check ob überhaupt Bild-Marker vorhanden sind
        // → Vermeidet teures sheetXml.split('</c>') bei großen Sheets ohne Bilder
        // ================================================================
        const hasImageFormulas = /DISPIMG|IMAGE/i.test(sheetXml);
        const hasVmAttributes = /\bvm="/.test(sheetXml);
        
        if (hasImageFormulas || hasVmAttributes) {
        const cellBlocks = sheetXml.split('</c>');
        
        // Methode 1: DISPIMG/IMAGE Formeln
        if (hasImageFormulas) {
        for (const block of cellBlocks) {
            if (!/DISPIMG|IMAGE/i.test(block)) continue;
            if (!/<f[\s>]/.test(block)) continue;
            
            const cellRefMatch = block.match(/<c\s[^>]*?r="([A-Z]+\d+)"/);
            if (!cellRefMatch) continue;
            
            const cellRef = cellRefMatch[1];
            const colLetters = cellRef.match(/^([A-Z]+)/)[1];
            const rowNum = parseInt(cellRef.match(/(\d+)$/)[1]);
            result.imageCells.push({
                ref: cellRef,
                col: colLettersToIndex(colLetters),
                row: rowNum - 1
            });
        }
        
        // Shared formulas mit DISPIMG/IMAGE
        if (result.imageCells.length > 0) {
            const sharedImageIds = new Set();
            for (const block of cellBlocks) {
                if (!/DISPIMG|IMAGE/i.test(block)) continue;
                const siMatch = block.match(/<f[^>]*t="shared"[^>]*si="(\d+)"/);
                if (siMatch) sharedImageIds.add(siMatch[1]);
            }
            if (sharedImageIds.size > 0) {
                for (const block of cellBlocks) {
                    const siRef = block.match(/<f[^>]*t="shared"[^>]*si="(\d+)"[^>]*\/?>/);
                    if (!siRef || !sharedImageIds.has(siRef[1])) continue;
                    const cellRefMatch = block.match(/<c\s[^>]*?r="([A-Z]+\d+)"/);
                    if (!cellRefMatch) continue;
                    const cellRef = cellRefMatch[1];
                    const col = colLettersToIndex(cellRef.match(/^([A-Z]+)/)[1]);
                    const row = parseInt(cellRef.match(/(\d+)$/)[1]) - 1;
                    if (!result.imageCells.some(ic => ic.col === col && ic.row === row)) {
                        result.imageCells.push({ ref: cellRef, col, row });
                    }
                }
            }
        }
        } // Ende hasImageFormulas
        
        // Methode 2: vm-Attribut (Value Metadata) — Excel 365 Zellbilder
        // Zellen mit vm="N" verweisen auf xl/richData/rdrichvalue.xml
        // Prüfe zuerst ob richData existiert (= es gibt Zellbilder)
        if (hasVmAttributes) {
        const hasRichData = zip.getEntries().some(e => /richData/i.test(e.entryName));
        if (hasRichData) {
            // Prüfe welche richData-Einträge Bilder enthalten
            let richDataHasImages = false;
            try {
                // rdrichvalue.xml enthält <rv>-Einträge, Bilder haben type mit Bild-Referenz
                const rdEntry = zip.getEntry('xl/richData/rdrichvalue.xml') || 
                                zip.getEntry('xl/richData/rdRichValue.xml');
                if (rdEntry) {
                    const rdXml = rdEntry.getData().toString('utf8');
                    // Bilder in richData haben typischerweise einen Verweis auf media/ oder image
                    richDataHasImages = /img|image|picture|Bild/i.test(rdXml) || 
                                         rdXml.includes('<v>') || // Hat Werte = wahrscheinlich Bilder
                                         true; // richData existiert = fast immer Bilder
                }
                // Auch rdRichValueStructure.xml prüfen
                const structEntry = zip.getEntry('xl/richData/rdrichvaluestructure.xml') ||
                                    zip.getEntry('xl/richData/rdRichValueStructure.xml');
                if (structEntry) {
                    const structXml = structEntry.getData().toString('utf8');
                    if (/image|img|picture/i.test(structXml)) {
                        richDataHasImages = true;
                    }
                }
            } catch (e) {
                // Fallback: richData existiert = nehme an es sind Bilder
                richDataHasImages = true;
            }
            
            if (richDataHasImages) {
                // Finde alle Zellen mit vm-Attribut
                for (const block of cellBlocks) {
                    // vm-Attribut auf dem <c>-Element
                    const vmMatch = block.match(/<c\s[^>]*?\bvm="(\d+)"[^>]*?r="([A-Z]+\d+)"/);
                    const vmMatch2 = block.match(/<c\s[^>]*?r="([A-Z]+\d+)"[^>]*?\bvm="(\d+)"/);
                    
                    let cellRef = null;
                    let vmValue = null;
                    if (vmMatch) {
                        cellRef = vmMatch[2];
                        vmValue = vmMatch[1];
                    } else if (vmMatch2) {
                        cellRef = vmMatch2[1];
                        vmValue = vmMatch2[2];
                    }
                    
                    if (!cellRef) continue;
                    
                    const col = colLettersToIndex(cellRef.match(/^([A-Z]+)/)[1]);
                    const row = parseInt(cellRef.match(/(\d+)$/)[1]) - 1;
                    if (!result.imageCells.some(ic => ic.col === col && ic.row === row)) {
                        result.imageCells.push({ ref: cellRef, col, row, vmValue });
                    }
                }
                console.log(`[SheetMetadata] richData gefunden, ${result.imageCells.length} Bild-Zellen (inkl. vm-Attribut)`);
            }
        }
        } // Ende hasVmAttributes
        } // Ende hasImageFormulas || hasVmAttributes
        
        if (result.imageCells.length > 0) {
            console.log(`[SheetMetadata] ${result.imageCells.length} Bild-Zellen erkannt: ${result.imageCells.map(c => c.ref).join(', ')}`);
        }
        
        // Tables (für AutoFilter, falls kein direkter AutoFilter vorhanden)
        if (!result.autoFilterRange) {
            try {
                const sheetFileName = fullSheetPath.split('/').pop();
                const sheetDir = fullSheetPath.substring(0, fullSheetPath.lastIndexOf('/'));
                const sheetRelsPath = `${sheetDir}/_rels/${sheetFileName}.rels`;
                const sheetRelsEntry = zip.getEntry(sheetRelsPath);
                
                if (sheetRelsEntry) {
                    const sheetRelsXml = sheetRelsEntry.getData().toString('utf8');
                    const tableRefPattern = /Target="([^"]*table[^"]*)"/g;
                    let tableRefMatch;
                    
                    while ((tableRefMatch = tableRefPattern.exec(sheetRelsXml)) !== null) {
                        let tablePath = tableRefMatch[1];
                        if (tablePath.startsWith('../')) {
                            tablePath = 'xl/' + tablePath.replace('../', '');
                        } else if (!tablePath.startsWith('xl/')) {
                            tablePath = `${sheetDir}/${tablePath}`;
                        }
                        
                        const tableEntry = zip.getEntry(tablePath);
                        if (tableEntry) {
                            const tableXml = tableEntry.getData().toString('utf8');
                            const tableAfMatch = tableXml.match(/<autoFilter[^>]*ref="([^"]*)"/);
                            if (tableAfMatch) {
                                result.autoFilterRange = tableAfMatch[1];
                                break;
                            }
                            // Fallback: tableRef als AutoFilter
                            const tableRefAttr = tableXml.match(/<table[^>]*ref="([^"]*)"/);
                            if (tableRefAttr) {
                                result.autoFilterRange = tableRefAttr[1];
                                break;
                            }
                        }
                    }
                }
            } catch (tableErr) {
                // Tables sind optional
            }
        }
        
    } catch (error) {
        console.error('[SheetMetadata] Fehler:', error.message);
    }
    
    return result;
}


/**
 * Parst die Shared Strings aus der Excel-ZIP-Datei.
 * ExcelJS Streaming löst RichText-SharedStrings nicht korrekt auf,
 * daher müssen wir sie manuell parsen.
 * 
 * @param {Buffer} fileBuffer - Der Datei-Buffer
 * @returns {Array} Array von SharedString-Objekten: { text, richText?, font? }
 */
function parseSharedStrings(fileBuffer, existingZip = null) {
    try {
        const zip = existingZip || new AdmZip(fileBuffer);
        const ssEntry = zip.getEntry('xl/sharedStrings.xml');
        if (!ssEntry) return [];
        
        const xml = ssEntry.getData().toString('utf8');
        const strings = [];
        
        // Parse jedes <si>...</si> Element
        const siRegex = /<si>([\s\S]*?)<\/si>/g;
        let siMatch;
        
        while ((siMatch = siRegex.exec(xml)) !== null) {
            const siContent = siMatch[1];
            
            // Prüfe ob es RichText ist (mehrere <r>-Elemente)
            const runs = [];
            const runRegex = /<r>([\s\S]*?)<\/r>/g;
            let runMatch;
            
            while ((runMatch = runRegex.exec(siContent)) !== null) {
                const runContent = runMatch[1];
                
                // Text extrahieren
                const textMatch = runContent.match(/<t[^>]*>([\s\S]*?)<\/t>/);
                const text = textMatch ? textMatch[1]
                    .replace(/&amp;/g, '&')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&apos;/g, "'")
                    : '';
                
                // Font/Run-Properties extrahieren
                const rprMatch = runContent.match(/<rPr>([\s\S]*?)<\/rPr>/);
                const font = {};
                
                if (rprMatch) {
                    const rpr = rprMatch[1];
                    if (/<b\s*\/>|<b>/.test(rpr)) font.bold = true;
                    if (/<i\s*\/>|<i>/.test(rpr)) font.italic = true;
                    if (/<u\s*\/>|<u>|<u [^>]*\/>/.test(rpr)) font.underline = true;
                    if (/<strike\s*\/>|<strike>/.test(rpr)) font.strike = true;
                    
                    // Schriftgröße
                    const szMatch = rpr.match(/<sz\s+val="([^"]+)"/);
                    if (szMatch) font.size = parseFloat(szMatch[1]);
                    
                    // Schriftname
                    const nameMatch = rpr.match(/<rFont\s+val="([^"]+)"/);
                    if (nameMatch) font.name = nameMatch[1];
                    
                    // Farbe - ARGB
                    const colorMatch = rpr.match(/<color\s+([^/>]*)\/?>/);
                    if (colorMatch) {
                        const colorAttrs = colorMatch[1];
                        const rgbMatch = colorAttrs.match(/rgb="([^"]+)"/);
                        const themeMatch = colorAttrs.match(/theme="([^"]+)"/);
                        const tintMatch = colorAttrs.match(/tint="([^"]+)"/);
                        const indexedMatch = colorAttrs.match(/indexed="([^"]+)"/);
                        
                        if (rgbMatch) {
                            font.color = { argb: rgbMatch[1] };
                        } else if (themeMatch) {
                            font.color = { theme: parseInt(themeMatch[1]) };
                            if (tintMatch) font.color.tint = parseFloat(tintMatch[1]);
                        } else if (indexedMatch) {
                            font.color = { indexed: parseInt(indexedMatch[1]) };
                        }
                    }
                }
                
                runs.push({ text, font: Object.keys(font).length > 0 ? font : undefined });
            }
            
            if (runs.length > 0) {
                // RichText: mehrere Runs
                strings.push({ richText: runs });
            } else {
                // Einfacher Text (kein RichText)
                const tMatch = siContent.match(/<t[^>]*>([\s\S]*?)<\/t>/);
                const plainText = tMatch ? tMatch[1]
                    .replace(/&amp;/g, '&')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&apos;/g, "'")
                    : '';
                strings.push({ text: plainText });
            }
        }
        
        return strings;
    } catch (e) {
        console.error('[ExcelJS] Shared Strings Parse-Fehler:', e.message);
        return [];
    }
}


/**
 * Extrahiert Cell-Styles aus einer ExcelJS-Zelle.
 * Wird sowohl im Non-Streaming als auch im Streaming Reader verwendet.
 * @param {Object} cell - ExcelJS Cell-Objekt
 * @returns {Object} Style-Objekt
 */
function extractCellStyle(cell) {
    const style = {};
    
    if (cell.font) {
        if (cell.font.bold) style.bold = true;
        if (cell.font.italic) style.italic = true;
        if (cell.font.underline) style.underline = true;
        if (cell.font.strike) style.strikethrough = true;
        if (cell.font.size) style.fontSize = cell.font.size;
        if (cell.font.name && cell.font.name !== 'Calibri') style.fontName = cell.font.name;
        const fontColor = resolveColor(cell.font.color);
        if (fontColor) style.fontColor = fontColor;
    }
    
    if (cell.alignment) {
        if (cell.alignment.horizontal && cell.alignment.horizontal !== 'general') {
            style.textAlign = cell.alignment.horizontal;
        }
        if (cell.alignment.vertical && cell.alignment.vertical !== 'bottom') {
            style.verticalAlign = cell.alignment.vertical;
        }
        if (cell.alignment.wrapText) style.wrapText = true;
    }
    
    if (cell.fill) {
        if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
            const fillColor = resolveColor(cell.fill.fgColor);
            if (fillColor) style.fill = fillColor;
        }
    }
    
    if (cell.border) {
        const borders = {};
        for (const side of ['top', 'bottom', 'left', 'right']) {
            if (cell.border[side] && cell.border[side].style) {
                borders[side] = {
                    style: cell.border[side].style,
                    color: resolveColor(cell.border[side].color)
                };
            }
        }
        if (Object.keys(borders).length > 0) style.borders = borders;
    }
    
    return style;
}


/**
/**
 * Interne Hilfsfunktion: Liest ein Sheet via ExcelJS Streaming Reader.
 * Wird von readSheetWithExcelJS aufgerufen. Wenn der Streaming Reader fehlschlägt
 * (z.B. bei ImportExcel/EPPlus-Dateien mit ZIP Data Descriptors), wirft die Funktion
 * einen Fehler, damit der Aufrufer auf Non-Streaming zurückfallen kann.
 */
async function _readSheetStreaming(
    ExcelJS, fileBuffer, sheetName, actualColumnCount,
    sharedStrings, imageCellSet, cellFormulas, cellHyperlinks, cellStyles, richTextCells
) {
    const headers = [];
    const data = [];
    let sheetFound = false;
    
    console.log(`[ExcelJS] Verwende Streaming Reader (kein RichText)`);
    
    const { Readable } = require('stream');
    const readStream = Readable.from(fileBuffer);
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(readStream, {
        sharedStrings: 'cache',
        hyperlinks: 'cache',
        styles: 'cache',
        worksheets: 'emit'
    });
    
    const t2 = Date.now();
    
    for await (const worksheetReader of workbookReader) {
        if (worksheetReader.name !== sheetName) continue;
        sheetFound = true;
        console.log(`[ExcelJS] Sheet "${sheetName}" gefunden, starte Streaming...`);
        
        let dataRowCounter = 0;
        let lastProgressLog = Date.now();
        let headerRowNumber = null; // Dynamisch: erste nicht-leere Zeile wird Header
    
    for await (const row of worksheetReader) {
        const rowNumber = row.number;
        
        // Leere Zeilen auffüllen (Streaming überspringt Zeilen ohne Zellen im XML)
        if (headerRowNumber !== null && rowNumber > headerRowNumber + 1) {
            const expectedDataRow = rowNumber - headerRowNumber - 1;
            while (dataRowCounter < expectedDataRow) {
                data.push(new Array(actualColumnCount).fill(''));
                dataRowCounter++;
            }
        }
        
        // Erste nicht-leere Zeile = Header (ImportExcel-Dateien können leere Zeile 1 haben)
        if (headerRowNumber === null) {
            headerRowNumber = rowNumber;
            if (rowNumber !== 1) {
                console.log(`[ExcelJS] Header nicht in Zeile 1, sondern in Zeile ${rowNumber} gefunden (z.B. ImportExcel)`);
            }
            // Initialisiere Header-Array mit leeren Strings für alle Spalten
            for (let i = 0; i < actualColumnCount; i++) {
                headers.push('');
            }
            row.eachCell((cell, colNumber) => {
                const colIndex = colNumber - 1;
                // Header-Array erweitern falls nötig
                while (colIndex >= headers.length) {
                    headers.push('');
                    actualColumnCount = Math.max(actualColumnCount, headers.length);
                }
                // Überschreibe den leeren Wert mit dem tatsächlichen Wert
                // SharedString-Referenz auflösen (Streaming löst RichText-SharedStrings nicht auf)
                if (cell.value && typeof cell.value === 'object' && cell.value.sharedString !== undefined && sharedStrings.length > 0) {
                    const ssIdx = cell.value.sharedString;
                    const ss = sharedStrings[ssIdx];
                    if (ss) {
                        if (ss.richText) {
                            cell.value = { richText: ss.richText.map(r => ({ text: r.text, font: r.font })) };
                        } else {
                            cell.value = ss.text || '';
                        }
                    }
                }
                if (!cell.value) {
                    headers[colIndex] = '';
                } else if (typeof cell.value === 'object') {
                    // Rich Text, Hyperlinks, Bilder etc.
                    if (cell.value.richText) {
                        headers[colIndex] = cell.value.richText.map(part => part.text).join('');
                    } else if (cell.value.text !== undefined) {
                        headers[colIndex] = String(cell.value.text);
                    } else if (cell.value.buffer || cell.value.image || cell.value.imageId) {
                        headers[colIndex] = '🖼️ Bild';
                    } else {
                        headers[colIndex] = '📎 Objekt';
                    }
                } else {
                    headers[colIndex] = String(cell.value);
                }
                
                // WICHTIG: Auch Header-Styles extrahieren (für Frontend-Kompatibilität)
                const styleKey = `0-${colIndex}`; // Header = Zeile 0
                const style = {};
                
                if (cell.font) {
                    if (cell.font.bold) style.bold = true;
                    if (cell.font.italic) style.italic = true;
                    if (cell.font.underline) style.underline = true;
                    if (cell.font.strike) style.strikethrough = true;
                    if (cell.font.size) {
                        style.fontSize = cell.font.size;
                    }
                    if (cell.font.name && cell.font.name !== 'Calibri') {
                        style.fontName = cell.font.name;
                    }
                const fontColor = resolveColor(cell.font.color);
                if (fontColor) {
                    style.fontColor = fontColor;
                }
            }
            
            // Alignment extrahieren
            if (cell.alignment) {
                if (cell.alignment.horizontal && cell.alignment.horizontal !== 'general') {
                    style.textAlign = cell.alignment.horizontal;
                }
                if (cell.alignment.vertical && cell.alignment.vertical !== 'bottom') {
                    style.verticalAlign = cell.alignment.vertical;
                }
                if (cell.alignment.wrapText) {
                    style.wrapText = true;
                }
            }
            
            // Fill extrahieren
            if (cell.fill) {
                if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
                    const fillColor = resolveColor(cell.fill.fgColor);
                    if (fillColor) {
                        style.fill = fillColor;
                    }
                }
            }
            
            // Borders extrahieren
            if (cell.border) {
                const borders = {};
                for (const side of ['top', 'bottom', 'left', 'right']) {
                    if (cell.border[side] && cell.border[side].style) {
                        borders[side] = {
                            style: cell.border[side].style,
                            color: resolveColor(cell.border[side].color)
                        };
                    }
                }
                if (Object.keys(borders).length > 0) {
                    style.borders = borders;
                }
            }
            
            if (Object.keys(style).length > 0) {
                cellStyles[styleKey] = style;
            }
            });
            continue; // Weiter zur nächsten Zeile
        }
        
        // Daten-Zeilen
        // Initialisiere rowData mit leeren Strings für alle Spalten
        const rowData = new Array(actualColumnCount).fill('');
        
        // WICHTIG: Style-Key basiert auf dataRowCounter, nicht auf rowNumber!
        // Das stellt sicher, dass leere Zeilen nicht zu Index-Mismatches führen
        const currentDataRowIndex = dataRowCounter;
        
        row.eachCell((cell, colNumber) => {
            const colIndex = colNumber - 1;
            // WICHTIG: Frontend erwartet 1-basierte Indizes (wie xlsx-populate)
            const styleKey = `${currentDataRowIndex + 1}-${colIndex}`;
            
            let cellValue = cell.value;
            
            // Bild-Zellen sofort erkennen (aus XML-Metadaten)
            // ExcelJS kann DISPIMG/IMAGE Formeln nicht parsen, liefert #VALUE! oder Objekte
            // Daher VOR jeder anderen Verarbeitung abfangen
            if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                cellValue = '🖼️ Bild';
                // Style trotzdem extrahieren (textAlign, verticalAlign etc.)
                const style = extractCellStyle(cell);
                if (Object.keys(style).length > 0) {
                    cellStyles[styleKey] = style;
                }
                rowData[colIndex] = cellValue;
                return; // Nächste Zelle
            }                
            // Formel extrahieren - WICHTIG: VOR der Objekt-Behandlung!
            // Bei Formeln kann cell.value ein Objekt sein mit { formula, result }
            // oder cell.formula ist direkt verfügbar
            if (cell.formula) {
                cellFormulas[styleKey] = cell.formula;
                // Das Ergebnis ist in cell.result (nicht cell.value!)
                cellValue = cell.result !== undefined ? cell.result : '';
                // Bild-Formeln erkennen: IMAGE, DISPIMG (mit beliebigen Prefixen wie _xlfn._xlws.)
                if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(cell.formula)) {
                    cellValue = '🖼️ Bild';
                }
            } else if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
                // Formel als Objekt gespeichert: { formula: '...', result: ... }
                cellFormulas[styleKey] = cell.value.formula;
                cellValue = cell.value.result !== undefined ? cell.value.result : '';
                // Bild-Formeln erkennen
                if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(cell.value.formula)) {
                    cellValue = '🖼️ Bild';
                }
            }
            
            // Error-Werte behandeln (z.B. { error: '#VALUE!' } aus Formel-Ergebnissen)
            if (cellValue && typeof cellValue === 'object' && cellValue.error) {
                // Prüfe ob die zugehörige Formel eine Bild-Formel ist
                const formula = cellFormulas[styleKey] || '';
                if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(formula)) {
                    cellValue = '🖼️ Bild';
                } else if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                    // Bild-Formel aus XML-Metadaten erkannt
                    cellValue = '🖼️ Bild';
                } else {
                    console.log(`[ExcelJS] Error-Wert in Zelle ${styleKey}: ${cellValue.error}, Formel: ${formula || 'keine'}`);
                    cellValue = String(cellValue.error);
                }
            }
            // String-Error-Werte: #VALUE! ohne Formel = möglicherweise Bild oder nicht-auswertbar
            else if (typeof cellValue === 'string' && cellValue === '#VALUE!') {
                const formula = cellFormulas[styleKey] || '';
                if (/IMAGE\s*\(|DISPIMG\s*\(/i.test(formula)) {
                    cellValue = '🖼️ Bild';
                } else if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                    // Bild-Formel aus XML-Metadaten erkannt
                    cellValue = '🖼️ Bild';
                } else {
                    console.log(`[ExcelJS] #VALUE! String in Zelle ${styleKey}, Formel: ${formula || 'keine'}, cell.value Typ: ${typeof cell.value}, Keys: ${cell.value && typeof cell.value === 'object' ? Object.keys(cell.value).join(',') : 'N/A'}`);
                }
            }
            
            // Hyperlink extrahieren
            if (cell.hyperlink) {
                cellHyperlinks[styleKey] = cell.hyperlink.hyperlink || cell.hyperlink;
            }
            
            // WICHTIG: Datums-Behandlung VOR der allgemeinen Objekt-Behandlung!
            // Date ist auch ein Objekt, würde sonst mit String() konvertiert werden
            if (cellValue instanceof Date) {
                cellValue = formatDateWithNumFmt(cellValue, cell.numFmt || '');
            }
            
            // Numerische Werte gemäß numFmt runden (z.B. Buchhaltungsformat 95,90 € → 2 Nachkommastellen)
            if (typeof cellValue === 'number' && cell.numFmt) {
                cellValue = roundNumericByFormat(cellValue, cell.numFmt);
            }
            
            // SharedString-Referenz auflösen (Streaming löst RichText-SharedStrings nicht auf)
            if (cell.value && typeof cell.value === 'object' && cell.value.sharedString !== undefined && sharedStrings.length > 0) {
                const ssIdx = cell.value.sharedString;
                const ss = sharedStrings[ssIdx];
                if (ss) {
                    if (ss.richText) {
                        cell.value = { richText: ss.richText.map(r => ({ text: r.text, font: r.font })) };
                    } else {
                        cell.value = ss.text || '';
                    }
                    cellValue = typeof cell.value === 'string' ? cell.value : (cell.value.richText ? cell.value.richText.map(p => p.text).join('') : '');
                }
            }
            
            // Objekt-Werte behandeln (Rich Text, Hyperlinks, etc.)
            // WICHTIG: Nur wenn es KEINE Formel war (die wurde oben schon behandelt)
            // Wir prüfen cell.value (nicht cellValue), um zu sehen ob es ein spezielles Objekt ist
            // Date-Objekte NICHT als Objekte behandeln (wurden oben bereits formatiert)
            if (cell.value && typeof cell.value === 'object' && !(cell.value instanceof Date) && !cell.formula && !cell.value.formula) {
                // Rich Text extrahieren
                if (cell.value.richText) {
                    const richText = cell.value.richText.map(part => ({
                        text: part.text,
                        styles: {
                            bold: part.font?.bold || false,
                            italic: part.font?.italic || false,
                            underline: part.font?.underline || false,
                            strikethrough: part.font?.strike || false,
                        color: resolveColor(part.font?.color),
                            fontSize: part.font?.size || null,
                            fontName: part.font?.name || null
                        }
                    }));
                    richTextCells[styleKey] = richText;
                    // Konvertiere zu Plain Text - nimm den text direkt aus dem Original!
                    cellValue = cell.value.richText.map(part => part.text).join('');
                }
                // Hyperlink-Objekte (haben text und hyperlink Properties)
                else if (cell.value.text !== undefined && cell.value.hyperlink !== undefined) {
                    cellValue = cell.value.text;
                    cellHyperlinks[styleKey] = cell.value.hyperlink;
                }
                // Andere Objekte - versuche text-Property zu nutzen
                else if (cell.value.text !== undefined) {
                    cellValue = cell.value.text;
                }
                // Fallback: Null oder leerer String
                else if (cell.value === null) {
                    cellValue = '';
                }
                // Bild-Objekte erkennen (Buffer, Base64 oder image-Properties)
                else if (cell.value.buffer || cell.value.image || cell.value.imageId || 
                         (cell.value.extension && (cell.value.extension === 'png' || cell.value.extension === 'jpeg' || cell.value.extension === 'gif' || cell.value.extension === 'bmp'))) {
                    cellValue = '🖼️ Bild';
                }
                // Error-Objekte erkennen
                else if (cell.value.error) {
                    // Prüfe ob es eine Bild-Zelle ist (aus XML-Metadaten)
                    if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                        cellValue = '🖼️ Bild';
                    } else {
                        cellValue = cell.value.error; // z.B. #REF!, #VALUE!, #DIV/0!
                    }
                }
                // Letzter Fallback: Unbekanntes Objekt -> versuche sinnvolle Darstellung
                else {
                    // Prüfe ob JSON-Serialisierung "[object Object]" vermeidet
                    const keys = Object.keys(cell.value);
                    if (keys.length === 0) {
                        cellValue = '';
                    } else {
                        // Logge das unbekannte Objekt für Debugging
                        console.log(`[ExcelJS] Unbekanntes Objekt in Zelle ${styleKey}:`, JSON.stringify(cell.value).substring(0, 200));
                        cellValue = '📎 Objekt';
                    }
                }
            }
            
            // Styles extrahieren
            const style = {};
            
            if (cell.font) {
                if (cell.font.bold) style.bold = true;
                if (cell.font.italic) style.italic = true;
                if (cell.font.underline) style.underline = true;
                if (cell.font.strike) style.strikethrough = true;
                if (cell.font.size) {
                    style.fontSize = cell.font.size;
                }
                if (cell.font.name && cell.font.name !== 'Calibri') {
                    style.fontName = cell.font.name;
                }
                const fontColor = resolveColor(cell.font.color);
                if (fontColor) {
                    style.fontColor = fontColor;
                }
            }
            
            // Alignment extrahieren
            if (cell.alignment) {
                if (cell.alignment.horizontal && cell.alignment.horizontal !== 'general') {
                    style.textAlign = cell.alignment.horizontal;
                }
                if (cell.alignment.vertical && cell.alignment.vertical !== 'bottom') {
                    style.verticalAlign = cell.alignment.vertical;
                }
                if (cell.alignment.wrapText) {
                    style.wrapText = true;
                }
            }
            
            // Fill extrahieren
            if (cell.fill) {
                if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
                    const fillColor = resolveColor(cell.fill.fgColor);
                    if (fillColor) {
                        style.fill = fillColor;
                    }
                }
            }
            
            // Borders extrahieren
            if (cell.border) {
                const borders = {};
                for (const side of ['top', 'bottom', 'left', 'right']) {
                    if (cell.border[side] && cell.border[side].style) {
                        borders[side] = {
                            style: cell.border[side].style,
                            color: resolveColor(cell.border[side].color)
                        };
                    }
                }
                if (Object.keys(borders).length > 0) {
                    style.borders = borders;
                }
            }
            
            if (Object.keys(style).length > 0) {
                cellStyles[styleKey] = style;
            }
            
            // WICHTIG: Date-Objekte MÜSSEN hier als String formatiert werden
            // da sie sonst bei der IPC-Serialisierung zu "Thu Sep 19 2013..." werden
            if (cellValue instanceof Date) {
                // Fallback-Formatierung falls oben nicht gegriffen hat
                const day = cellValue.getDate();
                const month = cellValue.getMonth() + 1;
                const year = cellValue.getFullYear();
                cellValue = `${day}.${month}.${year}`;
            }
            // Auch String-Werte prüfen die wie Date.toString() aussehen
            else if (typeof cellValue === 'string' && /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s/.test(cellValue)) {
                // Versuche den Date-String zu parsen
                const parsedDate = new Date(cellValue);
                if (!isNaN(parsedDate.getTime())) {
                    const day = parsedDate.getDate();
                    const month = parsedDate.getMonth() + 1;
                    const year = parsedDate.getFullYear();
                    cellValue = `${day}.${month}.${year}`;
                }
            }
            
            // Setze den Wert an der korrekten Position
            // Letzte Absicherung: Objekte die durchgerutscht sind, nie als "[object Object]" speichern
            if (cellValue !== null && cellValue !== undefined && typeof cellValue === 'object') {
                console.log(`[ExcelJS] Objekt durchgerutscht in Zelle ${styleKey}:`, JSON.stringify(cellValue).substring(0, 200));
                cellValue = '📎 Objekt';
            }
            rowData[colIndex] = cellValue === null || cellValue === undefined ? '' : cellValue;
        });
        
        // Hidden Rows werden aus Metadaten (XML) geladen, nicht aus row.hidden
        
        data.push(rowData);
        dataRowCounter++; // Zähler für nächste Daten-Zeile erhöhen
        
        // Fortschritt alle 2 Sekunden loggen
        const now = Date.now();
        if (now - lastProgressLog > 2000) {
            console.log(`[ExcelJS] Streaming: ${dataRowCounter} Zeilen verarbeitet (${now - t2}ms)`);
            lastProgressLog = now;
        }
    } // Ende for-await row
    
        console.log(`[ExcelJS] Streaming abgeschlossen: ${dataRowCounter} Datenzeilen in ${Date.now() - t2}ms`);
        break; // Sheet gefunden, keine weiteren Sheets verarbeiten
    } // Ende for-await worksheetReader
    
    return { headers, data, actualColumnCount, sheetFound };
}

/**
 * Liest ein Excel-Sheet mit ExcelJS Streaming Reader (nicht-blockierend)
 * 
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des zu lesenden Sheets
 * @param {string|null} password - Optional: Passwort für geschützte Dateien
 * @returns {Promise<Object>} Sheet-Daten im gleichen Format wie xlsx-populate
 */
async function readSheetWithExcelJS(filePath, sheetName, password = null, cachedBuffer = null) {
    const startTime = Date.now();
    let tempFilePath = null; // Für entschlüsselte Dateien
    const timings = {}; // Detaillierte Zeitmessungen
    
    try {
        console.log(`[ExcelJS] === START readSheetWithExcelJS === Datei: ${path.basename(filePath)}, Sheet: ${sheetName}`);
        
        // Bei passwortgeschützten Dateien: xlsx-populate zum Entschlüsseln verwenden
        // ExcelJS hat bekannte Probleme mit Passwort-Entschlüsselung
        let actualFilePath = filePath;
        
        if (password) {
            try {
                // Lazy-load xlsx-populate
                if (!XlsxPopulate) {
                    XlsxPopulate = require('xlsx-populate');
                }
                
                // Datei mit xlsx-populate öffnen (entschlüsseln)
                // Gecachten Buffer nutzen falls vorhanden (vermeidet erneuten Netzwerk-Read)
                let pwWorkbook;
                if (cachedBuffer) {
                    pwWorkbook = await XlsxPopulate.fromDataAsync(cachedBuffer, { password });
                } else {
                    pwWorkbook = await XlsxPopulate.fromFileAsync(filePath, { password });
                }
                
                // Als temporäre Datei ohne Passwort speichern
                tempFilePath = path.join(os.tmpdir(), `mvms_decrypt_${crypto.randomUUID()}.xlsx`);
                await pwWorkbook.toFileAsync(tempFilePath);
                
                actualFilePath = tempFilePath;
                
            } catch (pwError) {
                console.error('[ExcelJS] Fehler beim Entschlüsseln:', pwError.message);
                
                // Prüfe ob es ein Passwort-Fehler ist
                if (pwError.message.includes('password') || pwError.message.includes('Password') || 
                    pwError.message.includes('decrypt') || pwError.message.includes('Decrypt')) {
                    return { 
                        success: false, 
                        error: 'Falsches Passwort oder Datei kann nicht entschlüsselt werden',
                        needsPassword: true
                    };
                }
                throw pwError;
            }
        }
        
        // Datei EINMAL async lesen — dieser Buffer wird für ALLES verwendet:
        // 1. Passwort-Check (AdmZip)
        // 2. Sheet-Metadaten (AdmZip)
        // 3. Streaming Reader (Readable.from)
        // Danach ist der File-Handle sofort frei für xlwings/Excel
        // PERFORMANCE: Gecachten Buffer wiederverwenden (spart Netzwerk-I/O bei erneuten Reads)
        const t0 = Date.now();
        const fileBuffer = (cachedBuffer && actualFilePath === filePath) ? cachedBuffer : await fs.promises.readFile(actualFilePath);
        timings.fileRead = Date.now() - t0;
        console.log(`[ExcelJS] Datei gelesen: ${(fileBuffer.length / 1024 / 1024).toFixed(1)} MB in ${timings.fileRead}ms${(cachedBuffer && actualFilePath === filePath) ? ' (aus Cache)' : ''}`);
        
        // ZIP einmalig erstellen und für Passwort-Check, Metadaten + SharedStrings wiederverwenden
        // Spart 2-3 redundante AdmZip-Instanziierungen (~200-500ms bei 5MB-Dateien)
        let zip = null;
        if (!password) {
            try {
                zip = new AdmZip(fileBuffer);
            } catch (zipError) {
                if (zipError.message.includes('password') || zipError.message.includes('Password') ||
                    zipError.message.includes('encrypted') || zipError.message.includes('Encrypted') ||
                    zipError.message.includes('CFB') || zipError.message.includes('Invalid or unsupported zip')) {
                    return { 
                        success: false, 
                        error: 'Diese Datei ist passwortgeschützt. Bitte Passwort eingeben.',
                        needsPassword: true
                    };
                }
            }
        }
        // Für Passwort-Dateien: ZIP aus entschlüsseltem Buffer erstellen
        if (!zip) {
            zip = new AdmZip(fileBuffer);
        }
        
        // Sheet-Metadaten aus Buffer extrahieren (ZIP-Instanz wiederverwenden!)
        const t1 = Date.now();
        const metadata = extractSheetMetadata(fileBuffer, sheetName, zip);
        timings.metadata = Date.now() - t1;
        let actualColumnCount = metadata.columnCount || 1;
        console.log(`[ExcelJS] Metadaten: ${actualColumnCount} Spalten, ${metadata.mergedCells.length} Merged Cells, ${metadata.hiddenColumns.length} Hidden Cols, ${metadata.hiddenRows.length} Hidden Rows in ${timings.metadata}ms`);
        
        // Daten-Strukturen initialisieren
        const headers = [];
        const data = [];
        // Hidden Rows aus Metadaten verwenden (ExcelJS Streaming Reader ignoriert row.hidden!)
        const hiddenRows = [...metadata.hiddenRows];
        const cellStyles = {};
        const cellFormulas = {};
        const cellHyperlinks = {};
        const richTextCells = {};
        
        // Metadaten aus ZIP (statt aus worksheet-Objekt)
        const autoFilterRange = metadata.autoFilterRange;
        const mergedCells = metadata.mergedCells;
        const hiddenColumns = metadata.hiddenColumns;
        
        // Bild-Zellen als Set für schnellen Lookup (z.B. "1_4" = col 1, row 4)
        const imageCellSet = new Set();
        if (metadata.imageCells && metadata.imageCells.length > 0) {
            for (const ic of metadata.imageCells) {
                imageCellSet.add(`${ic.col}_${ic.row}`);
            }
            console.log(`[ExcelJS] ${imageCellSet.size} Bild-Zellen aus Metadaten geladen`);
        }
        
        // ============================================================
        // SHARED STRINGS: Prüfe ob RichText vorhanden ist
        // Streaming löst RichText-SharedStrings nicht auf und liefert keine Styles
        // ============================================================
        const sharedStrings = parseSharedStrings(fileBuffer, zip);
        const hasRichTextSharedStrings = sharedStrings.some(ss => ss.richText);
        if (sharedStrings.length > 0) {
            console.log(`[ExcelJS] ${sharedStrings.length} Shared Strings, davon ${sharedStrings.filter(s => s.richText).length} mit RichText`);
        }
        
        // ============================================================
        // READER-AUSWAHL: Non-Streaming wenn RichText vorhanden (für korrekte Styles)
        // Streaming wenn kein RichText (für Performance bei großen Dateien)
        // FALLBACK: Wenn Streaming fehlschlägt (z.B. bei ImportExcel/EPPlus-Dateien
        //   mit ZIP Data Descriptors), automatisch auf Non-Streaming wechseln
        // ============================================================
        const t2 = Date.now();
        let useNonStreaming = hasRichTextSharedStrings;
        let streamingFailed = false;
        
        if (!useNonStreaming) {
            // Versuche Streaming Reader zuerst
            try {
                const streamingResult = await _readSheetStreaming(
                    ExcelJS, fileBuffer, sheetName, actualColumnCount,
                    sharedStrings, imageCellSet, cellFormulas, cellHyperlinks, cellStyles, richTextCells
                );
                // Streaming erfolgreich — Daten übernehmen
                headers.push(...streamingResult.headers);
                data.push(...streamingResult.data);
                actualColumnCount = streamingResult.actualColumnCount;
                if (!streamingResult.sheetFound) {
                    return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
                }
                timings.streaming = Date.now() - t2;
                console.log(`[ExcelJS] Streaming abgeschlossen: ${data.length} Datenzeilen in ${timings.streaming}ms`);
            } catch (streamErr) {
                // Streaming fehlgeschlagen (z.B. ImportExcel/EPPlus: "invalid signature: 0x8074b50")
                // → Fallback auf Non-Streaming Reader
                console.warn(`[ExcelJS] Streaming fehlgeschlagen: ${streamErr.message} → Fallback auf Non-Streaming`);
                useNonStreaming = true;
                streamingFailed = true;
            }
        }
        
        if (useNonStreaming) {
            // NON-STREAMING: Löst alle SharedStrings, RichText und Styles korrekt auf
            const reason = streamingFailed ? 'Streaming-Fallback' : 'RichText erkannt';
            console.log(`[ExcelJS] Verwende Non-Streaming Reader (${reason})`);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(fileBuffer);
            
            const worksheet = workbook.getWorksheet(sheetName);
            if (!worksheet) {
                return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
            }
            
            let dataRowCounter = 0;
            let headerRowNumber = null; // Dynamisch: erste nicht-leere Zeile wird Header
            
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                if (headerRowNumber === null) {
                    headerRowNumber = rowNumber;
                    if (rowNumber !== 1) {
                        console.log(`[ExcelJS] Header nicht in Zeile 1, sondern in Zeile ${rowNumber} gefunden (z.B. ImportExcel)`);
                    }
                    // Header
                    for (let i = 0; i < actualColumnCount; i++) {
                        headers.push('');
                    }
                    row.eachCell((cell, colNumber) => {
                        const colIndex = colNumber - 1;
                        while (colIndex >= headers.length) {
                            headers.push('');
                            actualColumnCount = Math.max(actualColumnCount, headers.length);
                        }
                        if (!cell.value) {
                            headers[colIndex] = '';
                        } else if (typeof cell.value === 'object') {
                            if (cell.value.richText) {
                                headers[colIndex] = cell.value.richText.map(part => part.text).join('');
                            } else if (cell.value.text !== undefined) {
                                headers[colIndex] = String(cell.value.text);
                            } else if (cell.value.buffer || cell.value.image || cell.value.imageId) {
                                headers[colIndex] = '🖼️ Bild';
                            } else {
                                headers[colIndex] = '📎 Objekt';
                            }
                        } else {
                            headers[colIndex] = String(cell.value);
                        }
                        
                        // Header-Styles
                        const styleKey = `0-${colIndex}`;
                        const style = extractCellStyle(cell);
                        if (Object.keys(style).length > 0) {
                            cellStyles[styleKey] = style;
                        }
                    });
                    return; // Nächste Zeile
                }
                
                // Leere Zeilen auffüllen
                const expectedDataRow = rowNumber - headerRowNumber - 1;
                while (dataRowCounter < expectedDataRow) {
                    data.push(new Array(actualColumnCount).fill(''));
                    dataRowCounter++;
                }
                
                const currentDataRowIndex = dataRowCounter;
                const rowData = new Array(actualColumnCount).fill('');
                
                row.eachCell((cell, colNumber) => {
                    const colIndex = colNumber - 1;
                    while (colIndex >= rowData.length) {
                        rowData.push('');
                        actualColumnCount = Math.max(actualColumnCount, rowData.length);
                    }
                    
                    const styleKey = `${currentDataRowIndex + 1}-${colIndex}`;
                    let cellValue = cell.value;
                    
                    // Bild-Zellen sofort erkennen (aus XML-Metadaten)
                    // Muss VOR allen anderen Checks stehen (wie im Streaming Reader)
                    if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                        cellValue = '🖼️ Bild';
                        // Style trotzdem extrahieren
                        const style = extractCellStyle(cell);
                        if (Object.keys(style).length > 0) {
                            cellStyles[styleKey] = style;
                        }
                        rowData[colIndex] = cellValue;
                        return; // Nächste Zelle
                    }
                    
                    // Formeln
                    if (cell.formula) {
                        cellFormulas[styleKey] = cell.formula;
                        cellValue = cell.result !== undefined ? cell.result : cell.value;
                    } else if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
                        cellFormulas[styleKey] = cell.value.formula;
                        cellValue = cell.value.result !== undefined ? cell.value.result : '';
                    }
                    
                    // Hyperlinks
                    if (cell.hyperlink) {
                        cellHyperlinks[styleKey] = cell.hyperlink.hyperlink || cell.hyperlink;
                    }
                    
                    // Datum
                    if (cellValue instanceof Date) {
                        cellValue = formatDateWithNumFmt(cellValue, cell.numFmt || '');
                    }
                    
                    // Numerische Werte gemäß numFmt runden (z.B. Buchhaltungsformat 95,90 € → 2 Nachkommastellen)
                    if (typeof cellValue === 'number' && cell.numFmt) {
                        cellValue = roundNumericByFormat(cellValue, cell.numFmt);
                    }
                    
                    // Objekte (RichText, Hyperlinks, etc.)
                    if (cell.value && typeof cell.value === 'object' && !(cell.value instanceof Date) && !cell.formula && !cell.value.formula) {
                        if (cell.value.richText) {
                            const richText = cell.value.richText.map(part => ({
                                text: part.text,
                                styles: {
                                    bold: part.font?.bold || false,
                                    italic: part.font?.italic || false,
                                    underline: part.font?.underline || false,
                                    strikethrough: part.font?.strike || false,
                                    color: resolveColor(part.font?.color),
                                    fontSize: part.font?.size || null,
                                    fontName: part.font?.name || null
                                }
                            }));
                            richTextCells[styleKey] = richText;
                            cellValue = cell.value.richText.map(part => part.text).join('');
                        } else if (cell.value.text !== undefined && cell.value.hyperlink !== undefined) {
                            cellValue = cell.value.text;
                            cellHyperlinks[styleKey] = cell.value.hyperlink;
                        } else if (cell.value.text !== undefined) {
                            cellValue = cell.value.text;
                        } else if (cell.value === null) {
                            cellValue = '';
                        } else if (cell.value.error) {
                            if (imageCellSet.has(`${colIndex}_${rowNumber - 1}`)) {
                                cellValue = '🖼️ Bild';
                            } else {
                                cellValue = cell.value.error;
                            }
                        } else if (cell.value.buffer || cell.value.image || cell.value.imageId) {
                            cellValue = '🖼️ Bild';
                        } else {
                            cellValue = '';
                        }
                    }
                    
                    // Styles
                    const style = extractCellStyle(cell);
                    if (Object.keys(style).length > 0) {
                        cellStyles[styleKey] = style;
                    }
                    
                    // Date Fallback
                    if (cellValue instanceof Date) {
                        cellValue = cellValue.toLocaleDateString('de-DE');
                    }
                    
                    if (typeof cellValue === 'object' && cellValue !== null) {
                        cellValue = '';
                    }
                    
                    rowData[colIndex] = cellValue === null || cellValue === undefined ? '' : cellValue;
                });
                
                // Hidden Rows werden aus Metadaten (XML) geladen, nicht aus row.hidden
                
                data.push(rowData);
                dataRowCounter++;
            });
            
            console.log(`[ExcelJS] Non-Streaming abgeschlossen: ${dataRowCounter} Datenzeilen in ${Date.now() - t2}ms`);
            timings.streaming = Date.now() - t2;
        } // Ende Non-Streaming-Block
        
        // WICHTIG: Header-Zeile als erste Zeile in data einfügen
        // Das Frontend erwartet data.slice(1) - also Header an Position 0
        data.unshift(headers);
        
        // ============================================================
        // FALLBACK: Fill-Farben direkt aus XLSX extrahieren
        // ExcelJS erkennt bei manchen Dateien (z.B. SoftMaker) keine Fills
        // NUR wenn ExcelJS keine einzige Fill-Farbe gefunden hat
        // SKIP für große Dateien (>5000 Zeilen) — blockiert den Event-Loop zu lange
        // ============================================================
        const excelJSHasFills = Object.values(cellStyles).some(s => s.fill);
        const dataRowCount = data.length - 1; // Minus Header
        
        if (!excelJSHasFills && dataRowCount <= 5000) {
            console.log(`[ExcelJS] Kein ExcelJS-Fill gefunden, verwende ZIP-Fallback (${dataRowCount} Zeilen)`);
            const t3 = Date.now();
            // Buffer statt Dateipfad verwenden (kein erneuter Dateizugriff!)
            const directFills = extractFillsFromXLSX(fileBuffer, sheetName, zip);
            timings.fillFallback = Date.now() - t3;
            console.log(`[ExcelJS] Fill-Fallback: ${Object.keys(directFills).length} Fills in ${timings.fillFallback}ms`);
            
            if (Object.keys(directFills).length > 0) {
                for (const [key, fillColor] of Object.entries(directFills)) {
                    if (cellStyles[key]) {
                        if (!cellStyles[key].fill) {
                            cellStyles[key].fill = fillColor;
                        }
                    } else {
                        cellStyles[key] = { fill: fillColor };
                    }
                }
            }
        } else if (!excelJSHasFills && dataRowCount > 5000) {
            console.log(`[ExcelJS] Fill-Fallback ÜBERSPRUNGEN (${dataRowCount} Zeilen > 5000 — zu langsam)`);
            timings.fillFallback = 0;
        }
        
        const totalTime = Date.now() - startTime;
        
        // Zeilenfarben erkennen (wenn alle Zellen einer Zeile die gleiche Hintergrundfarbe haben)
        const t4 = Date.now();
        const rowHighlights = detectRowHighlights(cellStyles, data.length, headers.length);
        timings.highlights = Date.now() - t4;
        
        console.log(`[ExcelJS] === FERTIG === ${data.length} Zeilen, ${headers.length} Spalten, ${Object.keys(cellStyles).length} Styles in ${totalTime}ms`);
        console.log(`[ExcelJS] Timings: fileRead=${timings.fileRead}ms, metadata=${timings.metadata}ms, streaming=${timings.streaming}ms, fillFallback=${timings.fillFallback || 0}ms, highlights=${timings.highlights}ms`);
        
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
            mergedCells,
            autoFilterRange,
            rowHighlights,  // NEU: Zeilenfarben als Array von [rowIndex, colorName]
            imageCells: metadata.imageCells || [],  // NEU: Bild-Zellen mit vm-Werten für Copy&Paste
            stats: {
                rows: data.length,
                columns: headers.length,
                loadTimeMs: totalTime,
                timings // Detaillierte Zeitmessungen für Diagnose
            }
        };
        
    } catch (error) {
        console.error('[ExcelJS] Fehler beim Laden:', error);
        return { success: false, error: error.message };
    } finally {
        // Temporäre entschlüsselte Datei aufräumen
        if (tempFilePath) {
            try {
                const fs = require('fs');
                if (fs.existsSync(tempFilePath)) {
                    fs.unlinkSync(tempFilePath);
                }
            } catch (cleanupError) {
                console.warn('[ExcelJS] Konnte temporäre Datei nicht löschen:', cleanupError.message);
            }
        }
    }
}

module.exports = {
    readSheetWithExcelJS,
    extractFillsFromXLSX,
    extractSheetMetadata
};
