const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const XlsxPopulate = require('xlsx-populate'); // LEGACY: Nur noch für Backup/Fallback
const { readSheetWithExcelJS } = require('./exceljs-reader'); // PRIMÄR: ExcelJS Reader
const { exportSheetWithExcelJS } = require('./exceljs-writer'); // PRIMÄR: ExcelJS Writer
const pythonBridge = require('./python/python_bridge'); // NEU: Python/openpyxl für bessere Kompatibilität
const fs = require('fs');
const os = require('os');

// Erhöhe V8 Heap-Größe für große Dateien (muss vor app.ready gesetzt werden)
// 16GB für sehr große Excel-Dateien
app.commandLine.appendSwitch('js-flags', '--max-old-space-size=16384');

// WORKBOOK-CACHE: Halte geladene Workbooks im Speicher für schnelleren Sheet-Wechsel
const workbookCache = new Map();
const CACHE_MAX_SIZE = 3; // Maximal 3 Workbooks cachen
const CACHE_MAX_AGE = 5 * 60 * 1000; // 5 Minuten

function getCachedWorkbook(filePath, password = null) {
    const key = `${filePath}:${password || ''}`;
    const cached = workbookCache.get(key);
    
    if (cached && (Date.now() - cached.timestamp < CACHE_MAX_AGE)) {
        console.log('[Cache] Workbook aus Cache geladen');
        cached.timestamp = Date.now(); // Refresh timestamp
        return cached.workbook;
    }
    
    return null;
}

function setCachedWorkbook(filePath, password = null, workbook) {
    const key = `${filePath}:${password || ''}`;
    
    // LRU: Ältesten Cache-Eintrag entfernen wenn voll
    if (workbookCache.size >= CACHE_MAX_SIZE) {
        let oldestKey = null;
        let oldestTime = Date.now();
        
        for (const [k, v] of workbookCache.entries()) {
            if (v.timestamp < oldestTime) {
                oldestTime = v.timestamp;
                oldestKey = k;
            }
        }
        
        if (oldestKey) {
            console.log('[Cache] Entferne ältesten Cache-Eintrag');
            workbookCache.delete(oldestKey);
        }
    }
    
    workbookCache.set(key, {
        workbook,
        timestamp: Date.now()
    });
    console.log('[Cache] Workbook gecacht');
}

function clearWorkbookCache() {
    workbookCache.clear();
    console.log('[Cache] Cache geleert');
}

/**
 * Direkte ZIP-Manipulation für Partial-Cell-Mode
 * Ändert nur die betroffenen Zellen im sheet.xml ohne die gesamte Datei zu laden
 * VIEL speichersparender als xlsx-populate für kleine Änderungen!
 */
async function savePartialCellChangesDirectly(filePath, cellChangesBySheet) {
    const JSZip = require('jszip');
    
    try {
        console.log('[DirectSave] Starte direkte ZIP-Manipulation');
        
        // 1. Excel-Datei als ZIP laden
        const fileData = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(fileData);
        
        // 2. workbook.xml lesen um Sheet-IDs zu bekommen
        const workbookXml = await zip.file('xl/workbook.xml').async('string');
        const sheetIdMap = {};
        
        // Parse Sheet-Namen zu IDs: <sheet name="SheetName" sheetId="1" r:id="rId1"/>
        const sheetMatches = workbookXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="(rId\d+)"/g);
        for (const match of sheetMatches) {
            // XML-Entities dekodieren (&amp; -> &, &lt; -> <, etc.)
            const decodedName = match[1]
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'");
            sheetIdMap[decodedName] = match[2];
        }
        
        // 3. _rels/workbook.xml.rels lesen um rId zu Dateinamen zu mappen
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
        const relMap = {};
        const relMatches = relsXml.matchAll(/<Relationship[^>]+Id="(rId\d+)"[^>]+Target="([^"]+)"/g);
        for (const match of relMatches) {
            relMap[match[1]] = match[2];
        }
        
        let totalChanges = 0;
        
        // 4. Für jedes Sheet die Änderungen anwenden
        for (const [sheetName, changedCells] of Object.entries(cellChangesBySheet)) {
            
            if (Object.keys(changedCells).length === 0) continue;
            
            // Sheet-Datei finden
            const rId = sheetIdMap[sheetName];
            if (!rId) {
                console.warn(`[DirectSave] Sheet "${sheetName}" nicht gefunden`);
                continue;
            }
            
            const sheetFile = 'xl/' + relMap[rId];
            const sheetXml = await zip.file(sheetFile).async('string');
            
            // 5. XML parsen und Zellen ändern
            let modifiedXml = sheetXml;
            
            for (const [cellKey, newValue] of Object.entries(changedCells)) {
                const [rowStr, colStr] = cellKey.split('-');
                const rowNum = parseInt(rowStr) + 1; // 0-basiert -> 1-basiert
                const colNum = parseInt(colStr) + 1;
                
                // Spalten-Buchstabe berechnen (1=A, 2=B, etc.)
                const colLetter = numberToColumnLetter(colNum);
                const cellRef = `${colLetter}${rowNum}`;
                
                // Suche nach der Zelle im XML
                const cellRegex = new RegExp(`<c r="${cellRef}"[^>]*>.*?</c>`, 's');
                const match = modifiedXml.match(cellRegex);
                
                if (match) {
                    // Zelle existiert - Wert ändern
                    const oldCell = match[0];
                    const newCell = updateCellValue(oldCell, newValue, cellRef);
                    modifiedXml = modifiedXml.replace(oldCell, newCell);
                } else {
                    // Zelle existiert noch nicht - einfügen
                    modifiedXml = insertNewCell(modifiedXml, cellRef, rowNum, newValue);
                }
                
                totalChanges++;
            }
            
            // 6. Geändertes XML zurückschreiben
            zip.file(sheetFile, modifiedXml);
            console.log(`[DirectSave] Sheet "${sheetName}": ${Object.keys(changedCells).length} Zellen aktualisiert`);
        }
        
        // 7. ZIP zurückschreiben
        const outputBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE'
        });
        
        fs.writeFileSync(filePath, outputBuffer);
        
        // Cache invalidieren da Datei geändert wurde
        clearWorkbookCache();
        
        console.log(`[DirectSave] Erfolgreich ${totalChanges} Zellen gespeichert`);
        
        return {
            success: true,
            message: `${totalChanges} Änderung(en) gespeichert (Direkt-Methode)`,
            totalChanges
        };
        
    } catch (error) {
        console.error('[DirectSave] Fehler:', error);
        // Fallback auf normale Methode
        return { success: false, error: error.message, needsFallback: true };
    }
}

/**
 * Hilfsfunktion: Spalten-Nummer zu Buchstabe (1=A, 27=AA, etc.)
 */
function numberToColumnLetter(num) {
    let letter = '';
    while (num > 0) {
        const remainder = (num - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        num = Math.floor((num - 1) / 26);
    }
    return letter;
}

/**
 * Hilfsfunktion: Zell-Wert im XML aktualisieren
 * WICHTIG: Style-Attribute (s="...") müssen erhalten bleiben!
 */
function updateCellValue(cellXml, newValue, cellRef) {
    // Parse das <c> Tag und extrahiere alle Attribute
    const cTagMatch = cellXml.match(/<c\s+([^>]+)>/);
    if (!cTagMatch) return cellXml;
    
    const attributes = cTagMatch[1];
    
    // Extrahiere existierende Attribute (r, s, t, etc.)
    const attrMap = {};
    const attrRegex = /(\w+)="([^"]*)"/g;
    let match;
    while ((match = attrRegex.exec(attributes)) !== null) {
        attrMap[match[1]] = match[2];
    }
    
    // Bestimme neuen Typ und Wert
    let cellType = 'n'; // Default: Number
    let valueToInsert = String(newValue || '');
    
    if (typeof newValue === 'number' || (!isNaN(newValue) && newValue !== '' && newValue !== null)) {
        cellType = 'n';
        valueToInsert = String(newValue);
    } else if (newValue === '' || newValue === null || newValue === undefined) {
        // Leere Zelle - nur Style behalten, kein Wert
        cellType = null;
        valueToInsert = '';
    } else {
        // String - inline string verwenden (nicht SharedString wegen Komplexität)
        cellType = 'str';
        valueToInsert = escapeXml(String(newValue));
    }
    
    // Attribute neu zusammenbauen - WICHTIG: s="..." (Style) beibehalten!
    attrMap['r'] = cellRef;
    if (cellType) {
        attrMap['t'] = cellType;
    } else {
        delete attrMap['t'];
    }
    
    // Attribute-String neu bauen
    const newAttrs = Object.entries(attrMap)
        .map(([key, val]) => `${key}="${val}"`)
        .join(' ');
    
    // Neue Zelle bauen
    if (valueToInsert === '') {
        // Leere Zelle - nur Tag mit Attributen
        return `<c ${newAttrs}/>`;
    } else if (cellType === 'str') {
        // Inline String
        return `<c ${newAttrs}><v>${valueToInsert}</v></c>`;
    } else {
        // Numerischer Wert
        return `<c ${newAttrs}><v>${valueToInsert}</v></c>`;
    }
}


/**
 * Hilfsfunktion: Neue Zelle ins XML einfügen
 */
function insertNewCell(sheetXml, cellRef, rowNum, value) {
    // Finde oder erstelle die Row
    const rowRegex = new RegExp(`<row r="${rowNum}"[^>]*>.*?</row>`, 's');
    const rowMatch = sheetXml.match(rowRegex);
    
    const cellType = (typeof value === 'number' || !isNaN(value)) ? 'n' : 's';
    const newCellXml = `<c r="${cellRef}" t="${cellType}"><v>${escapeXml(String(value || ''))}</v></c>`;
    
    if (rowMatch) {
        // Row existiert - Zelle einfügen
        const oldRow = rowMatch[0];
        const newRow = oldRow.replace('</row>', newCellXml + '</row>');
        return sheetXml.replace(oldRow, newRow);
    } else {
        // Row existiert nicht - erstelle sie
        const newRow = `<row r="${rowNum}">${newCellXml}</row>`;
        // Finde <sheetData> und füge ein
        return sheetXml.replace('</sheetData>', newRow + '</sheetData>');
    }
}

/**
 * Hilfsfunktion: XML-Escaping
 */
function escapeXml(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}


// ============================================
// JSDoc TYPE DEFINITIONS
// ============================================

/**
 * @typedef {Object} FileDialogOptions
 * @property {string} [title] - Dialog title
 * @property {string} [defaultPath] - Default file path
 * @property {Array<{name: string, extensions: string[]}>} [filters] - File type filters
 */

/**
 * @typedef {Object} ExcelReadResult
 * @property {boolean} success - Whether the operation succeeded
 * @property {string} [filePath] - Path to the file
 * @property {string[]} [sheets] - List of sheet names
 * @property {string} [error] - Error message if failed
 */

/**
 * @typedef {Object} ExcelSheetData
 * @property {boolean} success - Whether the operation succeeded
 * @property {string[]} [headers] - Column headers
 * @property {Array<Array<string|number|Date>>} [data] - Row data
 * @property {number} [rowCount] - Total number of rows
 * @property {string} [error] - Error message if failed
 */

/**
 * @typedef {Object} TransferRow
 * @property {string} flag - Row flag (A/D/C or empty)
 * @property {string} comment - Row comment
 * @property {Object<number, string>} data - Column index to value mapping
 */

/**
 * @typedef {Object} InsertRowsParams
 * @property {string} filePath - Path to target Excel file
 * @property {string} sheetName - Target sheet name
 * @property {TransferRow[]} rows - Rows to insert
 * @property {number} startColumn - Starting column index
 */

/**
 * @typedef {Object} ConfigData
 * @property {string} [file1Path] - Source file path
 * @property {string} [file2Path] - Target file path
 * @property {string} [templatePath] - Template file path
 * @property {string} [sheet1Name] - Source sheet name
 * @property {string} [sheet2Name] - Target sheet name
 * @property {number} [startColumn] - Start column for insertion
 * @property {number} [checkColumn] - Column for duplicate checking
 * @property {number[]} [sourceColumns] - Source columns to copy
 */

/**
 * @typedef {Object} ExportParams
 * @property {string} filePath - Path to save the file
 * @property {string[]} headers - Column headers
 * @property {Array<Array<string|number>>} data - Row data
 */

let mainWindow = null;

// ============================================
// SICHERHEITSFUNKTIONEN
// ============================================

/**
 * Security-Logger für sicherheitsrelevante Ereignisse
 * Protokolliert kritische Operationen mit Zeitstempel
 * Speichert manipulationssicher in Log-Datei mit HMAC-Signaturen
 */
const crypto = require('crypto');

const securityLog = {
    entries: [],
    maxEntries: 1000,
    logFilePath: null,
    secretKey: null,
    lastHash: 'GENESIS',

    /**
     * Initialisiert den Logger mit Dateipfad und generiert/lädt den Secret Key
     */
    init() {
        const userDataPath = app.getPath('userData');
        this.logFilePath = path.join(userDataPath, 'security.log');
        const keyPath = path.join(userDataPath, '.security-key');

        // Secret Key laden oder generieren (einmalig pro Installation)
        try {
            if (fs.existsSync(keyPath)) {
                this.secretKey = fs.readFileSync(keyPath, 'utf8');
            } else {
                this.secretKey = crypto.randomBytes(32).toString('hex');
                fs.writeFileSync(keyPath, this.secretKey, { mode: 0o600 });
            }
        } catch (e) {
            // Fallback: Session-basierter Key (weniger sicher, aber funktional)
            this.secretKey = crypto.randomBytes(32).toString('hex');
        }

        // Letzten Hash aus existierender Log-Datei laden
        this.loadLastHash();
    },

    /**
     * Lädt den letzten Hash aus der existierenden Log-Datei
     */
    loadLastHash() {
        try {
            if (fs.existsSync(this.logFilePath)) {
                const content = fs.readFileSync(this.logFilePath, 'utf8').trim();
                if (content) {
                    const lines = content.split('\n');
                    const lastLine = lines[lines.length - 1];
                    if (lastLine) {
                        const entry = JSON.parse(lastLine);
                        if (entry.hash) {
                            this.lastHash = entry.hash;
                        }
                    }
                }
            }
        } catch (e) {
            // Bei Fehler mit GENESIS starten
            this.lastHash = 'GENESIS';
        }
    },

    /**
     * Berechnet HMAC-Signatur für einen Eintrag
     */
    calculateHMAC(data) {
        return crypto.createHmac('sha256', this.secretKey)
            .update(JSON.stringify(data))
            .digest('hex');
    },

    /**
     * Berechnet verketteten Hash (enthält vorherigen Hash)
     */
    calculateChainHash(entry, prevHash) {
        const dataToHash = JSON.stringify(entry) + prevHash;
        return crypto.createHash('sha256').update(dataToHash).digest('hex');
    },

    /**
     * Protokolliert ein sicherheitsrelevantes Ereignis
     * @param {'INFO'|'WARN'|'ERROR'|'SECURITY'} level - Log-Level
     * @param {string} action - Durchgeführte Aktion
     * @param {Object} details - Zusätzliche Details
     */
    log(level, action, details = {}) {
        const entry = {
            timestamp: new Date().toISOString(),
            level,
            action,
            details: { ...details, pid: process.pid }
        };

        this.entries.push(entry);

        // Älteste Einträge entfernen wenn Limit erreicht
        if (this.entries.length > this.maxEntries) {
            this.entries.shift();
        }

        // Auf Konsole ausgeben (im Entwicklungsmodus)
        const logMessage = `[${entry.timestamp}] [${level}] ${action}`;
        if (level === 'ERROR' || level === 'SECURITY') {
            console.error(logMessage, details);
        } else if (level === 'WARN') {
            console.warn(logMessage, details);
        } else if (process.argv.includes('--dev')) {
            console.log(logMessage, details);
        }

        // In Datei schreiben (manipulationssicher)
        this.writeToFile(entry);
    },

    /**
     * Schreibt Eintrag manipulationssicher in Log-Datei
     */
    writeToFile(entry) {
        if (!this.logFilePath || !this.secretKey) return;

        try {
            // Verketteten Hash berechnen (enthält vorherigen Hash)
            const chainHash = this.calculateChainHash(entry, this.lastHash);

            // HMAC-Signatur für Integrität
            const signedEntry = {
                ...entry,
                prevHash: this.lastHash,
                hash: chainHash,
                signature: this.calculateHMAC({ ...entry, prevHash: this.lastHash, hash: chainHash })
            };

            this.lastHash = chainHash;

            // An Datei anhängen
            fs.appendFileSync(this.logFilePath, JSON.stringify(signedEntry) + '\n');
        } catch (e) {
            // Fehler beim Schreiben ignorieren (Logging sollte App nicht crashen)
        }
    },

    /**
     * Liest alle Logs aus der Datei
     * @returns {Array} Log-Einträge
     */
    readFromFile() {
        if (!this.logFilePath) return [];

        try {
            if (!fs.existsSync(this.logFilePath)) return [];

            const content = fs.readFileSync(this.logFilePath, 'utf8').trim();
            if (!content) return [];

            return content.split('\n').map(line => {
                try {
                    return JSON.parse(line);
                } catch {
                    return null;
                }
            }).filter(Boolean);
        } catch (e) {
            return [];
        }
    },

    /**
     * Verifiziert die Integrität der Log-Datei
     * Prüft HMAC-Signaturen und Hash-Kette
     * @returns {{valid: boolean, errors: string[], totalEntries: number, verifiedEntries: number}}
     */
    verifyIntegrity() {
        const entries = this.readFromFile();
        const errors = [];
        let prevHash = 'GENESIS';
        let verifiedCount = 0;

        for (let i = 0; i < entries.length; i++) {
            const entry = entries[i];

            // 1. Prüfe ob prevHash zum vorherigen Eintrag passt
            if (entry.prevHash !== prevHash) {
                errors.push(`Zeile ${i + 1}: Hash-Kette unterbrochen (erwartet: ${prevHash.substring(0, 8)}..., gefunden: ${entry.prevHash?.substring(0, 8)}...)`);
            }

            // 2. Prüfe ob der Hash korrekt berechnet wurde
            const entryWithoutMeta = {
                timestamp: entry.timestamp,
                level: entry.level,
                action: entry.action,
                details: entry.details
            };
            const expectedHash = this.calculateChainHash(entryWithoutMeta, entry.prevHash);
            if (entry.hash !== expectedHash) {
                errors.push(`Zeile ${i + 1}: Hash-Manipulation erkannt bei "${entry.action}"`);
            }

            // 3. Prüfe HMAC-Signatur
            const dataToSign = { ...entryWithoutMeta, prevHash: entry.prevHash, hash: entry.hash };
            const expectedSig = this.calculateHMAC(dataToSign);
            if (entry.signature !== expectedSig) {
                errors.push(`Zeile ${i + 1}: Signatur ungültig bei "${entry.action}"`);
            } else {
                verifiedCount++;
            }

            prevHash = entry.hash;
        }

        return {
            valid: errors.length === 0,
            errors,
            totalEntries: entries.length,
            verifiedEntries: verifiedCount
        };
    },

    /**
     * Gibt alle Log-Einträge zurück (aus Speicher)
     * @returns {Array} Log-Einträge
     */
    getEntries() {
        return [...this.entries];
    },

    /**
     * Gibt Log-Einträge eines bestimmten Levels zurück
     * @param {'INFO'|'WARN'|'ERROR'|'SECURITY'} level
     * @returns {Array} Gefilterte Log-Einträge
     */
    getByLevel(level) {
        return this.entries.filter(e => e.level === level);
    },

    /**
     * Löscht die Log-Datei (mit neuem GENESIS-Eintrag)
     */
    clearLogs() {
        try {
            if (this.logFilePath && fs.existsSync(this.logFilePath)) {
                fs.unlinkSync(this.logFilePath);
            }
            this.lastHash = 'GENESIS';
            this.entries = [];

            // Neuen GENESIS-Eintrag erstellen
            this.log('SECURITY', 'LOGS_CLEARED', { reason: 'User initiated clear' });
            return { success: true };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }
};

/**
 * Network-Log für Aktionen auf Netzlaufwerken
 * Ermöglicht Nachverfolgung wer (Rechnername) wann was wo gemacht hat
 * DSGVO-konform: Nur Hostname, keine persönlichen Daten
 */
const networkLog = {
    hostname: null,
    lockTimeout: 5000, // 5 Sekunden Lock-Timeout
    networkDrives: null, // Set von Netzlaufwerk-Buchstaben (Windows)
    initialScanDone: false, // Flag für einmaliges Logging

    /**
     * Initialisiert den Network-Logger
     */
    init() {
        this.hostname = os.hostname();
        this.networkDrives = new Set();
        this.initialScanDone = false;

        // Windows: Netzlaufwerke erkennen
        if (process.platform === 'win32') {
            this.updateNetworkDrives(true); // true = initial scan, loggen
            // Periodisch aktualisieren (alle 30 Sekunden), aber ohne Logging
            setInterval(() => this.updateNetworkDrives(false), 30000);
        }
    },

    /**
     * Prüft ob ein Pfad auf einem Netzlaufwerk liegt
     * @param {string} filePath - Der zu prüfende Pfad
     * @returns {boolean} true wenn Netzlaufwerk
     */
    isNetworkPath(filePath) {
        if (!filePath || typeof filePath !== 'string') return false;

        // macOS: /Volumes/ (außer Macintosh HD)
        if (process.platform === 'darwin') {
            if (filePath.startsWith('/Volumes/')) {
                // Lokale Festplatte ausschließen
                const volumeName = filePath.split('/')[2];
                // Typische lokale Volume-Namen
                const localVolumes = ['Macintosh HD', 'Macintosh HD - Data', 'System'];
                return !localVolumes.includes(volumeName);
            }
            // SMB/AFP-Mounts in anderen Pfaden
            if (filePath.includes('/net/') || filePath.includes('/Network/')) {
                return true;
            }
        }

        // Windows: UNC-Pfade (\\server\share) oder gemappte Laufwerke prüfen
        if (process.platform === 'win32') {
            // UNC-Pfad
            if (filePath.startsWith('\\\\') || filePath.startsWith('//')) {
                return true;
            }

            // Gemapptes Netzlaufwerk prüfen (z.B. Y:, Z:)
            // Prüfe synchron mit bekannten Netzlaufwerken aus Cache
            const driveLetter = filePath.match(/^([A-Za-z]):/);
            if (driveLetter) {
                const drive = driveLetter[1].toUpperCase();
                // Prüfe ob in bekannten Netzlaufwerken
                if (this.networkDrives && this.networkDrives.has(drive)) {
                    return true;
                }
            }
        }

        return false;
    },

    /**
     * Aktualisiert die Liste der Netzlaufwerke (Windows)
     * Wird beim Start und periodisch aufgerufen
     * Erkennt: Netzlaufwerke (DriveType=4), gemappte Laufwerke (net use), VMware/VirtualBox Shared Folders
     * @param {boolean} logResults - Nur beim ersten Aufruf true, um Log-Spam zu vermeiden
     */
    async updateNetworkDrives(logResults = false) {
        if (process.platform !== 'win32') return;

        try {
            const { execSync } = require('child_process');
            this.networkDrives = new Set();

            // Methode 1: wmic für DriveType=4 (klassische Netzlaufwerke)
            try {
                const wmicOutput = execSync('wmic logicaldisk where drivetype=4 get deviceid', {
                    encoding: 'utf8',
                    timeout: 5000,
                    windowsHide: true
                });

                const wmicLines = wmicOutput.split('\n');
                for (const line of wmicLines) {
                    const match = line.trim().match(/^([A-Z]):?$/i);
                    if (match) {
                        this.networkDrives.add(match[1].toUpperCase());
                    }
                }
            } catch (e) {
                // wmic fehlgeschlagen, ignorieren
            }

            // Methode 2: net use für gemappte Netzlaufwerke (inkl. VMware Shared Folders)
            try {
                const netUseOutput = execSync('net use', {
                    encoding: 'utf8',
                    timeout: 5000,
                    windowsHide: true
                });

                // Suche nach Zeilen mit Laufwerksbuchstaben (z.B. "OK           Y:        \\vmware-host\...")
                const netUseLines = netUseOutput.split('\n');
                for (const line of netUseLines) {
                    const match = line.match(/\s+([A-Z]):\s+/i);
                    if (match) {
                        this.networkDrives.add(match[1].toUpperCase());
                    }
                }
            } catch (e) {
                // net use fehlgeschlagen, ignorieren
            }

            // Methode 3: Prüfe alle Laufwerke auf Remote-Eigenschaft
            try {
                const allDrivesOutput = execSync('wmic logicaldisk get deviceid,drivetype,providername', {
                    encoding: 'utf8',
                    timeout: 5000,
                    windowsHide: true
                });

                const driveLines = allDrivesOutput.split('\n');
                for (const line of driveLines) {
                    // Prüfe auf VMware, VirtualBox oder andere Netzwerk-Provider
                    if (line.toLowerCase().includes('vmware') ||
                        line.toLowerCase().includes('virtualbox') ||
                        line.toLowerCase().includes('vboxsvr') ||
                        line.toLowerCase().includes('\\\\')) {
                        const match = line.match(/([A-Z]):/i);
                        if (match) {
                            this.networkDrives.add(match[1].toUpperCase());
                        }
                    }
                }
            } catch (e) {
                // Fallback fehlgeschlagen, ignorieren
            }

            // Log gefundene Netzlaufwerke nur beim ersten Scan
            if (logResults) {
                securityLog.log('INFO', 'NETWORK_DRIVES_DETECTED', {
                    detected: this.networkDrives.size > 0 ? Array.from(this.networkDrives).join(', ') : 'none',
                    count: this.networkDrives.size
                });
            }

        } catch (err) {
            console.log('Network drive detection error:', err.message);
            // Nur beim ersten Scan Fehler loggen
            if (logResults) {
                securityLog.log('WARN', 'NETWORK_DRIVES_SCAN_FAILED', {
                    error: err.message
                });
            }
        }
    },

    /**
     * Ermittelt den Log-Dateipfad für ein Netzlaufwerk
     * @param {string} filePath - Pfad zur bearbeiteten Datei
     * @returns {string|null} Pfad zur Log-Datei oder null
     */
    getNetworkLogPath(filePath) {
        if (!this.isNetworkPath(filePath)) return null;

        const dir = path.dirname(filePath);
        return path.join(dir, '.excel-sync-audit.log');
    },

    /**
     * Schreibt einen Eintrag ins Netzwerk-Log mit File-Locking
     * @param {string} filePath - Pfad zur bearbeiteten Datei
     * @param {string} action - Durchgeführte Aktion
     * @param {Object} details - Zusätzliche Details
     * @returns {Promise<{success: boolean, error?: string}>}
     */
    async log(filePath, action, details = {}) {
        const logPath = this.getNetworkLogPath(filePath);
        if (!logPath) return { success: true, skipped: true };

        const entry = {
            timestamp: new Date().toISOString(),
            hostname: this.hostname,
            action,
            file: path.basename(filePath),
            details
        };

        // File-Locking: Lock-Datei erstellen
        const lockPath = logPath + '.lock';
        const lockContent = `${this.hostname}:${Date.now()}`;

        try {
            // Versuche Lock zu erwerben
            let lockAcquired = false;
            const startTime = Date.now();

            while (!lockAcquired && (Date.now() - startTime) < this.lockTimeout) {
                try {
                    // Exclusive Schreibzugriff (wx = write exclusive, fails if exists)
                    fs.writeFileSync(lockPath, lockContent, { flag: 'wx' });
                    lockAcquired = true;
                } catch (lockErr) {
                    if (lockErr.code === 'EEXIST') {
                        // Lock existiert - prüfe ob er abgelaufen ist
                        try {
                            const existingLock = fs.readFileSync(lockPath, 'utf8');
                            const [, lockTime] = existingLock.split(':');
                            if (Date.now() - parseInt(lockTime, 10) > this.lockTimeout) {
                                // Lock ist abgelaufen - entfernen und neu versuchen
                                fs.unlinkSync(lockPath);
                            } else {
                                // Kurz warten und erneut versuchen
                                await new Promise(r => setTimeout(r, 100));
                            }
                        } catch {
                            await new Promise(r => setTimeout(r, 100));
                        }
                    } else {
                        throw lockErr;
                    }
                }
            }

            if (!lockAcquired) {
                return { success: false, error: 'Could not acquire file lock' };
            }

            // In Log-Datei schreiben
            try {
                fs.appendFileSync(logPath, JSON.stringify(entry) + '\n');
            } finally {
                // Lock immer freigeben
                try {
                    fs.unlinkSync(lockPath);
                } catch {
                    // Ignore unlock errors
                }
            }

            // Vollständige Replikation in lokale Security-Logs
            // Dies ermöglicht Nachverfolgung auch wenn Netzwerk nicht verfügbar
            securityLog.log('INFO', `NETWORK_${action}`, {
                hostname: this.hostname,
                file: path.basename(filePath),
                logPath: logPath.substring(0, 100),
                ...details
            });

            return { success: true };
        } catch (err) {
            securityLog.log('WARN', 'NETWORK_LOG_FAILED', {
                error: err.message,
                action
            });
            return { success: false, error: err.message };
        }
    },

    /**
     * Liest das Netzwerk-Log für einen bestimmten Ordner
     * @param {string} filePath - Pfad zu einer Datei im Ordner
     * @returns {Array} Log-Einträge
     */
    readLogs(filePath) {
        const logPath = this.getNetworkLogPath(filePath);
        if (!logPath) return [];

        try {
            if (!fs.existsSync(logPath)) return [];

            const content = fs.readFileSync(logPath, 'utf8').trim();
            if (!content) return [];

            return content.split('\n').map(line => {
                try {
                    return JSON.parse(line);
                } catch {
                    return null;
                }
            }).filter(Boolean);
        } catch {
            return [];
        }
    },

    /**
     * Erstellt eine Session-Lock-Datei für eine geöffnete Datei
     * @param {string} filePath - Pfad zur Excel-Datei
     * @returns {{success: boolean, error?: string}}
     */
    createSessionLock(filePath) {
        if (!this.isNetworkPath(filePath)) return { success: true, skipped: true };

        const lockPath = this.getSessionLockPath(filePath);
        if (!lockPath) return { success: true, skipped: true };

        const lockData = {
            hostname: this.hostname,
            timestamp: new Date().toISOString(),
            pid: process.pid
        };

        try {
            fs.writeFileSync(lockPath, JSON.stringify(lockData));
            return { success: true, lockPath };
        } catch (err) {
            return { success: false, error: err.message };
        }
    },

    /**
     * Entfernt die Session-Lock-Datei
     * @param {string} filePath - Pfad zur Excel-Datei
     */
    removeSessionLock(filePath) {
        if (!this.isNetworkPath(filePath)) return;

        const lockPath = this.getSessionLockPath(filePath);
        if (!lockPath) return;

        try {
            if (fs.existsSync(lockPath)) {
                // Nur löschen wenn es unser Lock ist
                const content = fs.readFileSync(lockPath, 'utf8');
                const lockData = JSON.parse(content);
                if (lockData.hostname === this.hostname) {
                    fs.unlinkSync(lockPath);
                }
            }
        } catch {
            // Ignore errors
        }
    },

    /**
     * Ermittelt den Pfad zur Session-Lock-Datei
     * @param {string} filePath - Pfad zur Excel-Datei
     * @returns {string|null}
     */
    getSessionLockPath(filePath) {
        if (!filePath) return null;
        const dir = path.dirname(filePath);
        const filename = path.basename(filePath);
        return path.join(dir, `.~lock.${filename}`);
    },

    /**
     * Prüft ob eine Datei kürzlich von einem anderen Rechner bearbeitet wurde
     * @param {string} filePath - Pfad zur Excel-Datei
     * @param {number} minutesThreshold - Schwellwert in Minuten (Standard: 5)
     * @returns {{conflict: boolean, details?: Object}}
     */
    checkConflict(filePath, minutesThreshold = 5) {
        if (!this.isNetworkPath(filePath)) {
            return { conflict: false, isNetworkPath: false };
        }

        const result = {
            conflict: false,
            isNetworkPath: true,
            recentActivity: null,
            activeLock: null,
            hostname: this.hostname
        };

        // 1. Prüfe Session-Lock-Datei
        const lockPath = this.getSessionLockPath(filePath);
        if (lockPath && fs.existsSync(lockPath)) {
            try {
                const content = fs.readFileSync(lockPath, 'utf8');
                const lockData = JSON.parse(content);

                // Lock von anderem Rechner?
                if (lockData.hostname !== this.hostname) {
                    const lockTime = new Date(lockData.timestamp);
                    const ageMinutes = (Date.now() - lockTime.getTime()) / 60000;

                    // Lock gilt als aktiv wenn jünger als 30 Minuten
                    if (ageMinutes < 30) {
                        result.conflict = true;
                        result.activeLock = {
                            hostname: lockData.hostname,
                            timestamp: lockData.timestamp,
                            ageMinutes: Math.round(ageMinutes)
                        };
                    }
                }
            } catch {
                // Ignore parse errors
            }
        }

        // 2. Prüfe Netzwerk-Log nach kürzlicher Aktivität
        const logs = this.readLogs(filePath);
        const fileName = path.basename(filePath);
        const threshold = Date.now() - (minutesThreshold * 60000);

        // Finde den letzten Eintrag für diese Datei von einem anderen Rechner
        for (let i = logs.length - 1; i >= 0; i--) {
            const entry = logs[i];
            if (entry.file !== fileName) continue;
            if (entry.hostname === this.hostname) continue;

            const entryTime = new Date(entry.timestamp).getTime();
            if (entryTime > threshold) {
                result.conflict = true;
                result.recentActivity = {
                    hostname: entry.hostname,
                    timestamp: entry.timestamp,
                    action: entry.action,
                    ageMinutes: Math.round((Date.now() - entryTime) / 60000)
                };
                break;
            }
        }

        return result;
    }
};

/**
 * Config-Schema-Validierung
 * Prüft ob die geladene Konfiguration gültige Typen und Werte hat
 */
/**
 * Ermittelt den Computernamen für die Computer-spezifische Konfiguration
 * @returns {string} Der Computername in Großbuchstaben
 */
function getComputerName() {
    return (os.hostname() || 'UNKNOWN').toUpperCase();
}

/**
 * Führt die default-Konfiguration mit der Computer-spezifischen zusammen
 * @param {Object} rawConfig - Die rohe Config mit default und Computer-Abschnitten
 * @returns {{mergedConfig: Object, computerName: string, hasComputerSection: boolean}}
 */
function mergeComputerConfig(rawConfig) {
    const computerName = getComputerName();
    
    // Prüfe ob es eine verschachtelte Struktur ist (hat 'default' Abschnitt)
    if (!rawConfig.default && !rawConfig[computerName]) {
        // Alte flache Struktur - direkt zurückgeben
        return {
            mergedConfig: rawConfig,
            computerName: computerName,
            hasComputerSection: false,
            isLegacyFormat: true
        };
    }
    
    // Neue verschachtelte Struktur
    const defaultConfig = rawConfig.default || {};
    const computerConfig = rawConfig[computerName] || {};
    
    // Deep merge: Computer-spezifische Werte überschreiben default
    const mergedConfig = { ...defaultConfig, ...computerConfig };
    
    return {
        mergedConfig,
        computerName,
        hasComputerSection: !!rawConfig[computerName],
        isLegacyFormat: false
    };
}

const configSchema = {
    // Liste der erlaubten Config-Felder
    configFields: [
        'file1Path', 'file2Path', 'templatePath', 'sheet1Name', 'sheet2Name',
        'startColumn', 'checkColumn', 'flagColumn', 'commentColumn',
        'sourceColumns', 'enableFlag', 'enableComment', 'workingDir',
        'theme', 'language', 'file1SheetName', 'file2SheetName', 'mapping',
        'exportDate', 'extraColumns', 'file1Name', 'file2Name', 'templateName',
        'excelEngine'  // 'auto', 'xlwings' oder 'openpyxl'
    ],

    /**
     * Prüft ob ein Schlüssel ein Computer-Name sein könnte (für verschachtelte Configs)
     * @param {string} key - Der zu prüfende Schlüssel
     * @returns {boolean}
     */
    isComputerSection(key) {
        // 'default' ist speziell, alle anderen unbekannten Keys könnten Computer-Namen sein
        if (key === 'default') return true;
        // Bekannte Config-Felder sind keine Computer-Abschnitte
        return !this.configFields.includes(key);
    },

    /**
     * Validiert ein Config-Objekt gegen das erwartete Schema
     * @param {Object} config - Das zu validierende Config-Objekt (bereits gemergt)
     * @returns {{valid: boolean, errors: string[]}} Validierungsergebnis
     */
    validate(config) {
        const errors = [];

        if (!config || typeof config !== 'object' || Array.isArray(config)) {
            return { valid: false, errors: ['Config muss ein Objekt sein'] };
        }

        // Definiere erwartete Typen für jedes Feld
        const fieldTypes = {
            file1Path: 'string',
            file2Path: 'string',
            templatePath: 'string',
            sheet1Name: 'string',
            sheet2Name: 'string',
            startColumn: 'number',
            checkColumn: 'number',
            flagColumn: 'number',
            commentColumn: 'number',
            sourceColumns: 'array',
            enableFlag: 'boolean',
            enableComment: 'boolean',
            workingDir: 'string',
            theme: 'string',
            language: 'string',
            // Zusätzliche Felder aus config.json
            file1SheetName: 'string',
            file2SheetName: 'string',
            mapping: 'object',
            exportDate: 'string',
            extraColumns: 'object',
            file1Name: 'string',
            file2Name: 'string',
            templateName: 'string',
            excelEngine: 'string'  // 'auto', 'xlwings' oder 'openpyxl'
        };

        // Erlaubte Werte für bestimmte Felder
        const allowedValues = {
            theme: ['dark', 'light'],
            language: ['de', 'en'],
            excelEngine: ['auto', 'xlwings', 'openpyxl']
        };

        // Prüfe jeden bekannten Schlüssel
        for (const [key, value] of Object.entries(config)) {
            // Überspringe null/undefined Werte (optional)
            if (value === null || value === undefined) {
                continue;
            }

            const expectedType = fieldTypes[key];

            // Unbekannter Schlüssel - könnte ein Computer-Abschnitt sein (bei verschachtelter Config)
            // oder ein wirklich unbekannter Key
            if (!expectedType) {
                // Nur warnen wenn es kein Objekt ist (Computer-Abschnitte sind Objekte)
                if (typeof value !== 'object' || Array.isArray(value)) {
                    securityLog.log('WARN', 'CONFIG_UNKNOWN_KEY', { key });
                }
                continue;
            }

            // Typ-Prüfung
            if (expectedType === 'array') {
                if (!Array.isArray(value)) {
                    errors.push(`Feld '${key}' muss ein Array sein, ist aber ${typeof value}`);
                } else if (key === 'sourceColumns') {
                    // Array-Elemente müssen Zahlen sein
                    const invalidElements = value.filter(v => typeof v !== 'number' || !Number.isInteger(v));
                    if (invalidElements.length > 0) {
                        errors.push(`Feld '${key}' enthält ungültige Elemente (müssen Ganzzahlen sein)`);
                    }
                }
            } else if (typeof value !== expectedType) {
                errors.push(`Feld '${key}' muss vom Typ '${expectedType}' sein, ist aber '${typeof value}'`);
            }

            // Werte-Prüfung für enum-artige Felder
            if (allowedValues[key] && !allowedValues[key].includes(value)) {
                errors.push(`Feld '${key}' hat ungültigen Wert '${value}'. Erlaubt: ${allowedValues[key].join(', ')}`);
            }

            // Zahlen müssen positiv sein (für Spalten-Indizes)
            if (expectedType === 'number' && typeof value === 'number') {
                if (value < 0 || !Number.isFinite(value)) {
                    errors.push(`Feld '${key}' muss eine positive Zahl sein`);
                }
            }

            // Pfad-Validierung für Dateipfade
            if (key.endsWith('Path') && typeof value === 'string' && value.length > 0) {
                if (value.includes('\0')) {
                    errors.push(`Feld '${key}' enthält ungültige Zeichen (Null-Byte)`);
                    securityLog.log('SECURITY', 'CONFIG_NULL_BYTE_IN_PATH', { key, value: '[REDACTED]' });
                }
            }
        }

        return { valid: errors.length === 0, errors };
    },

    /**
     * Bereinigt ein Config-Objekt von ungültigen oder gefährlichen Werten
     * @param {Object} config - Das zu bereinigende Config-Objekt
     * @returns {Object} Bereinigtes Config-Objekt
     */
    sanitize(config) {
        if (!config || typeof config !== 'object') {
            return {};
        }

        const sanitized = {};
        const safeKeys = [
            'file1Path', 'file2Path', 'templatePath', 'sheet1Name', 'sheet2Name',
            'startColumn', 'checkColumn', 'flagColumn', 'commentColumn',
            'sourceColumns', 'enableFlag', 'enableComment', 'workingDir',
            'theme', 'language',
            // Zusätzliche Keys aus config.json
            'file1SheetName', 'file2SheetName', 'mapping', 'exportDate',
            'extraColumns', 'file1Name', 'file2Name', 'templateName'
        ];

        for (const key of safeKeys) {
            if (config.hasOwnProperty(key) && config[key] !== undefined) {
                sanitized[key] = config[key];
            }
        }

        return sanitized;
    }
};

/**
 * Prüft ob ein Dateipfad sicher ist (keine Path Traversal-Angriffe)
 * @param {string} filePath - Der zu prüfende Pfad
 * @returns {boolean} true wenn der Pfad sicher ist
 */
function isValidFilePath(filePath) {
    if (!filePath || typeof filePath !== 'string') {
        securityLog.log('WARN', 'INVALID_PATH_TYPE', { type: typeof filePath });
        return false;
    }

    // Normalisiere den Pfad
    const normalized = path.normalize(filePath);

    // Prüfe auf Path Traversal-Muster
    if (normalized.includes('..')) {
        securityLog.log('SECURITY', 'PATH_TRAVERSAL_ATTEMPT', {
            path: filePath.substring(0, 100) + (filePath.length > 100 ? '...' : '')
        });
        return false;
    }

    // Prüfe auf null-bytes (kann Sicherheitsprüfungen umgehen)
    if (filePath.includes('\0')) {
        securityLog.log('SECURITY', 'NULL_BYTE_IN_PATH', {
            pathLength: filePath.length
        });
        return false;
    }

    return true;
}

// ============================================
// FENSTER ERSTELLEN
// ============================================
function createWindow() {
    // Plattformspezifisches Icon
    const iconFile = process.platform === 'darwin' ? 'icon.icns' :
                     process.platform === 'win32' ? 'icon.ico' : 'icon.png';

    mainWindow = new BrowserWindow({
        width: 1600,
        height: 1000,
        minWidth: 800,
        minHeight: 600,
        resizable: true,
        maximizable: true,
        fullscreenable: true,
        fullscreen: false,  // Explizit kein Vollbild
        simpleFullscreen: false,  // Kein einfacher Vollbildmodus
        title: 'Excel Data Sync Pro',
        icon: path.join(__dirname, 'assets', iconFile),
        frame: true,
        // Wichtig für korrekte Dialog-Darstellung
        useContentSize: false,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        }
    });

    mainWindow.loadFile('src/index.html');

    // Anwendungsmenü mit Edit-Befehlen (für Ctrl+C, Ctrl+V, Ctrl+X, Ctrl+A)
    const appMenuTemplate = [
        {
            label: 'Edit',
            submenu: [
                { role: 'undo', accelerator: 'CmdOrCtrl+Z' },
                { role: 'redo', accelerator: 'CmdOrCtrl+Shift+Z' },
                { type: 'separator' },
                { role: 'cut', accelerator: 'CmdOrCtrl+X' },
                { role: 'copy', accelerator: 'CmdOrCtrl+C' },
                { role: 'paste', accelerator: 'CmdOrCtrl+V' },
                { role: 'selectAll', accelerator: 'CmdOrCtrl+A' }
            ]
        },
        {
            label: 'View',
            submenu: [
                { role: 'reload', accelerator: 'CmdOrCtrl+R' },
                { role: 'forceReload', accelerator: 'CmdOrCtrl+Shift+R' },
                { role: 'toggleDevTools', accelerator: 'CmdOrCtrl+Shift+I' },
                { type: 'separator' },
                { role: 'resetZoom' },
                { role: 'zoomIn' },
                { role: 'zoomOut' },
                { type: 'separator' },
                { role: 'togglefullscreen' }
            ]
        }
    ];

    // Auf macOS muss das App-Menü zuerst kommen
    if (process.platform === 'darwin') {
        appMenuTemplate.unshift({
            label: app.name,
            submenu: [
                { role: 'about' },
                { type: 'separator' },
                { role: 'hide' },
                { role: 'hideOthers' },
                { role: 'unhide' },
                { type: 'separator' },
                { role: 'quit' }
            ]
        });
    }

    const appMenu = Menu.buildFromTemplate(appMenuTemplate);
    Menu.setApplicationMenu(appMenu);

    // Kontextmenü für alle Bereiche aktivieren
    mainWindow.webContents.on('context-menu', (event, params) => {
        const { isEditable, selectionText, editFlags } = params;

        // Immer ein vollständiges Bearbeiten-Menü anzeigen
        const menuTemplate = [
            {
                label: 'Ausschneiden',
                role: 'cut',
                enabled: editFlags.canCut || isEditable
            },
            {
                label: 'Kopieren',
                role: 'copy',
                enabled: editFlags.canCopy || selectionText
            },
            {
                label: 'Einfügen',
                role: 'paste',
                enabled: true  // Immer aktiv - System prüft ob Zwischenablage Inhalt hat
            },
            { type: 'separator' },
            {
                label: 'Alles auswählen',
                role: 'selectAll',
                enabled: true
            }
        ];

        const menu = Menu.buildFromTemplate(menuTemplate);
        menu.popup({ window: mainWindow });
    });

    // DevTools oeffnen (nur waehrend Entwicklung)
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }

    // Menueleiste ausblenden (optional)
    mainWindow.setMenuBarVisibility(false);

    // Schließen-Anfrage abfangen für Warteschlangen-Prüfung
    let closeConfirmed = false;

    mainWindow.on('close', (e) => {
        if (!closeConfirmed) {
            e.preventDefault();
            // Renderer fragen, ob Warteschlange leer ist
            mainWindow.webContents.send('app:beforeClose');
        }
    });

    ipcMain.on('app:confirmClose', (event, canClose) => {
        if (canClose) {
            closeConfirmed = true;
            mainWindow.close();
        }
    });

    // Fenster-Referenz aufräumen wenn geschlossen
    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(async () => {
    // Security-Logger initialisieren (für Datei-basiertes Logging)
    securityLog.init();
    securityLog.log('INFO', 'APP_STARTED', { version: app.getVersion() });

    // Network-Logger initialisieren (für Netzlaufwerk-Protokollierung)
    networkLog.init();

    // Excel-Verfügbarkeit prüfen und loggen
    try {
        const excelStatus = await pythonBridge.checkExcelAvailable();
        const engine = pythonBridge.getExcelEngine();
        
        if (excelStatus.excelAvailable) {
            console.log(`[App] ✓ Microsoft Excel erkannt - xlwings wird verwendet (Engine: ${engine})`);
            securityLog.log('INFO', 'EXCEL_DETECTED', { 
                method: 'xlwings', 
                configuredEngine: engine,
                message: excelStatus.message 
            });
        } else {
            console.log(`[App] ✗ Microsoft Excel nicht verfügbar - openpyxl wird verwendet (Engine: ${engine})`);
            securityLog.log('INFO', 'EXCEL_NOT_DETECTED', { 
                method: 'openpyxl', 
                configuredEngine: engine,
                message: excelStatus.message 
            });
        }
    } catch (error) {
        console.log('[App] ⚠ Excel-Prüfung fehlgeschlagen:', error.message);
        securityLog.log('WARN', 'EXCEL_CHECK_FAILED', { error: error.message });
    }

    createWindow();
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    // Nur wenn App bereit ist und kein Fenster offen ist
    if (app.isReady() && BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// ============================================
// DATEI-DIALOGE - Windows-Workaround f�r Dialog-Problem
// ============================================

// Datei oeffnen Dialog
ipcMain.handle('dialog:openFile', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verf�gbar f�r Dialog');
        return null;
    }

    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();

    try {
        // Standard-Pfad setzen (hilft bei Dialog-Gr��enproblemen unter Windows)
        const defaultPath = options.defaultPath || app.getPath('documents');

        const result = await dialog.showOpenDialog({
            title: options.title || 'Datei oeffnen',
            defaultPath: defaultPath,
            filters: options.filters || [
                { name: 'Excel-Dateien', extensions: ['xlsx', 'xls'] },
                { name: 'Alle Dateien', extensions: ['*'] }
            ],
            properties: ['openFile']
        });

        // Nach Dialog: Hauptfenster wieder fokussieren
        if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.focus();
        }

        if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
            return null;
        }
        return result.filePaths[0];
    } catch (err) {
        console.error('Dialog Fehler:', err);
        return null;
    }
});

// Datei speichern Dialog
ipcMain.handle('dialog:saveFile', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verf�gbar f�r Dialog');
        return null;
    }

    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();

    try {
        // Standard-Pfad setzen falls nicht angegeben
        const defaultPath = options.defaultPath || app.getPath('documents');

        const result = await dialog.showSaveDialog({
            title: options.title || 'Datei speichern',
            defaultPath: defaultPath,
            filters: options.filters || [
                { name: 'Excel-Dateien', extensions: ['xlsx'] }
            ]
        });

        // Nach Dialog: Hauptfenster wieder fokussieren
        if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.focus();
        }

        if (result.canceled || !result.filePath) {
            return null;
        }
        return result.filePath;
    } catch (err) {
        console.error('Dialog Fehler:', err);
        return null;
    }
});

// Ordner oeffnen Dialog (fuer Arbeitsordner)
ipcMain.handle('dialog:openFolder', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verfuegbar fuer Dialog');
        return null;
    }

    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();

    try {
        const defaultPath = options.defaultPath || app.getPath('documents');

        const result = await dialog.showOpenDialog({
            title: options.title || 'Ordner auswaehlen',
            defaultPath: defaultPath,
            properties: ['openDirectory']
        });

        // Nach Dialog: Hauptfenster wieder fokussieren
        if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.focus();
        }

        if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
            return null;
        }
        return result.filePaths[0];
    } catch (err) {
        console.error('Dialog Fehler:', err);
        return null;
    }
});

// ============================================
// DATEISYSTEM OPERATIONEN
// ============================================

// Prüfen ob Datei existiert
ipcMain.handle('fs:checkFileExists', async (event, filePath) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { exists: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        const fs = require('fs');
        const exists = fs.existsSync(filePath);
        return { exists };
    } catch (err) {
        console.error('Dateiprüfung Fehler:', err);
        return { exists: false, error: err.message };
    }
});

// ============================================
// EXCEL OPERATIONEN (xlsx-populate - erhaelt Formatierung!)
// ============================================

/**
 * Optimierte Speicher-Funktion für große Excel-Dateien
 * Verwendet toFileAsync mit GC für bessere Memory-Verwaltung
 * @param {Object} workbook - XlsxPopulate Workbook Objekt
 * @param {string} filePath - Zielpfad zum Speichern
 * @param {Object} saveOptions - Optionen (z.B. password)
 * @param {string} sourcePath - Optionaler Quellpfad für Größenprüfung
 */
async function saveWorkbookOptimized(workbook, filePath, saveOptions = {}, sourcePath = null) {
    try {
        // Prüfe Dateigröße für Logging
        const pathToCheck = sourcePath || filePath;
        let fileSizeMB = 0;
        
        if (fs.existsSync(pathToCheck)) {
            const stats = fs.statSync(pathToCheck);
            fileSizeMB = stats.size / (1024 * 1024);
        }
        
        if (fileSizeMB > 10) {
            console.log(`[SaveOptimized] Große Datei (${fileSizeMB.toFixed(1)}MB) - speichere mit GC`);
        }
        
        // toFileAsync ist bereits intern optimiert und verwendet Streams
        await workbook.toFileAsync(filePath, saveOptions);
        
        // Cache invalidieren da Datei geändert wurde
        clearWorkbookCache();
        
        // Garbage Collection nach dem Speichern für große Dateien
        if (fileSizeMB > 10 && global.gc) {
            global.gc();
            console.log('[SaveOptimized] Garbage Collection durchgeführt');
        }
    } catch (saveError) {
        console.error('[SaveOptimized] Speichern fehlgeschlagen:', saveError.message);
        throw saveError;
    }
}

// Hilfsfunktion: Prüft ob eine Excel-Datei Pivot-Tabellen enthält
async function checkForPivotTables(filePath) {
    const JSZip = require('jszip');
    try {
        const fileData = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(fileData);

        // Suche nach Pivot-Tabellen-Dateien im ZIP
        const pivotFiles = Object.keys(zip.files).filter(name =>
            name.includes('pivotTable') ||
            name.includes('pivotCache') ||
            name.includes('PivotTable') ||
            name.includes('PivotCache')
        );

        return pivotFiles.length > 0;
    } catch (error) {
        console.error('Fehler beim Prüfen auf Pivot-Tabellen:', error);
        return false;
    }
}

// Hilfsfunktion: Entfernt nicht verwendete Spalten aus dem Worksheet (Formatierung, Breite etc.)
function removeUnusedColumns(worksheet, usedColumnCount, originalColumnCount) {
    // 1. Zuerst alle Row-Objekte und deren XML-Nodes bereinigen
    if (worksheet._rows) {
        for (const row of Object.values(worksheet._rows)) {
            // Zell-Objekte entfernen
            if (row && row._cells) {
                for (const cellCol of Object.keys(row._cells)) {
                    if (parseInt(cellCol) > usedColumnCount) {
                        delete row._cells[cellCol];
                    }
                }
            }

            // XML Cell-Nodes entfernen
            if (row && row._node && row._node.children) {
                for (let i = row._node.children.length - 1; i >= 0; i--) {
                    const cellNode = row._node.children[i];
                    if (cellNode && cellNode.attributes && cellNode.attributes.r) {
                        const cellRef = cellNode.attributes.r;
                        const colLetters = cellRef.replace(/\d+/g, '');
                        const colNum = columnLetterToNumber(colLetters);
                        if (colNum > usedColumnCount) {
                            row._node.children.splice(i, 1);
                        }
                    }
                }
            }

            // Spans-Attribut korrigieren
            if (row && row._node && row._node.attributes && row._node.attributes.spans) {
                row._node.attributes.spans = `1:${usedColumnCount}`;
            }
        }
    }

    // 2. Auch die sheetData direkt durchgehen (für Rows die nicht in _rows sind)
    if (worksheet._node && worksheet._node.children) {
        const sheetDataNode = worksheet._node.children.find(c => c && c.name === 'sheetData');
        if (sheetDataNode && sheetDataNode.children) {
            for (const rowNode of sheetDataNode.children) {
                if (rowNode && rowNode.name === 'row') {
                    // Spans korrigieren
                    if (rowNode.attributes && rowNode.attributes.spans) {
                        rowNode.attributes.spans = `1:${usedColumnCount}`;
                    }

                    // Zellen außerhalb des Bereichs entfernen
                    if (rowNode.children) {
                        for (let i = rowNode.children.length - 1; i >= 0; i--) {
                            const cellNode = rowNode.children[i];
                            if (cellNode && cellNode.name === 'c' && cellNode.attributes && cellNode.attributes.r) {
                                const cellRef = cellNode.attributes.r;
                                const colLetters = cellRef.replace(/\d+/g, '');
                                const colNum = columnLetterToNumber(colLetters);
                                if (colNum > usedColumnCount) {
                                    rowNode.children.splice(i, 1);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // 3. Column-Objekte entfernen
    if (worksheet._columns) {
        for (const colNum of Object.keys(worksheet._columns)) {
            if (parseInt(colNum) > usedColumnCount) {
                delete worksheet._columns[colNum];
            }
        }
    }

    // Alle <col> XML-Nodes bearbeiten, die über usedColumnCount hinausgehen
    if (worksheet._colsNode && worksheet._colsNode.children) {
        const colsToRemove = [];
        const colsToModify = [];

        for (let i = 0; i < worksheet._colsNode.children.length; i++) {
            const colNode = worksheet._colsNode.children[i];
            if (colNode && colNode.attributes) {
                const min = parseInt(colNode.attributes.min);
                const max = parseInt(colNode.attributes.max);

                if (min > usedColumnCount) {
                    // Gesamter col-Bereich liegt außerhalb - komplett entfernen
                    colsToRemove.push(i);
                } else if (max > usedColumnCount) {
                    // Col-Bereich geht über usedColumnCount hinaus - auf usedColumnCount kürzen
                    colsToModify.push({ node: colNode, newMax: usedColumnCount });
                }
            }
        }

        // Von hinten entfernen
        for (let i = colsToRemove.length - 1; i >= 0; i--) {
            worksheet._colsNode.children.splice(colsToRemove[i], 1);
        }

        // Modifizieren
        for (const mod of colsToModify) {
            mod.node.attributes.max = mod.newMax;
        }
    }

    // ColNodes-Referenzen aufräumen
    if (worksheet._colNodes) {
        for (const colNum of Object.keys(worksheet._colNodes)) {
            if (parseInt(colNum) > usedColumnCount) {
                delete worksheet._colNodes[colNum];
            }
        }
    }

    // Merged Cells entfernen, die in gelöschten Spalten liegen
    if (worksheet._mergeCells) {
        const keysToRemove = [];
        for (const key of Object.keys(worksheet._mergeCells)) {
            // Key ist im Format "A1:B2"
            const match = key.match(/([A-Z]+)\d+:([A-Z]+)\d+/);
            if (match) {
                // Spalten-Buchstaben zu Nummern konvertieren
                const startColNum = columnLetterToNumber(match[1]);
                const endColNum = columnLetterToNumber(match[2]);

                // Wenn der Merge-Bereich in einer gelöschten Spalte liegt
                if (startColNum > usedColumnCount || endColNum > usedColumnCount) {
                    keysToRemove.push(key);
                }
            }
        }
        for (const key of keysToRemove) {
            delete worksheet._mergeCells[key];
        }
    }

    // Data Validations entfernen, die in gelöschten Spalten liegen
    if (worksheet._dataValidations) {
        const keysToRemove = [];
        for (const key of Object.keys(worksheet._dataValidations)) {
            const match = key.match(/([A-Z]+)\d+/);
            if (match) {
                const colNum = columnLetterToNumber(match[1]);
                if (colNum > usedColumnCount) {
                    keysToRemove.push(key);
                }
            }
        }
        for (const key of keysToRemove) {
            delete worksheet._dataValidations[key];
        }
    }
}

// Hilfsfunktion: Spalten-Buchstabe zu Nummer (A=1, B=2, AA=27, etc.)
function columnLetterToNumber(letters) {
    let result = 0;
    for (let i = 0; i < letters.length; i++) {
        result = result * 26 + (letters.charCodeAt(i) - 64);
    }
    return result;
}

// Hilfsfunktion: Entfernt nicht verwendete Zeilen aus dem Worksheet (Formatierung, Höhe etc.)
function removeUnusedRows(worksheet, usedRowCount, originalRowCount) {
    // Zeilen von hinten nach vorne entfernen (ab usedRowCount+2 bis originalRowCount+1, +1 wegen Header)
    // usedRowCount = Anzahl der Datenzeilen, also usedRowCount+1 = letzte Datenzeile (1-basiert)
    // originalRowCount = ursprüngliche Anzahl der Datenzeilen
    const lastUsedRow = usedRowCount + 1; // +1 für Header
    const lastOriginalRow = originalRowCount + 1; // +1 für Header

    for (let row = lastOriginalRow; row > lastUsedRow; row--) {
        // Row-Objekt entfernen
        if (worksheet._rows && worksheet._rows[row]) {
            delete worksheet._rows[row];
        }
    }

    // Merged Cells entfernen, die in gelöschten Zeilen liegen
    if (worksheet._mergeCells) {
        const keysToRemove = [];
        for (const key of Object.keys(worksheet._mergeCells)) {
            // Key ist im Format "A1:B2"
            const match = key.match(/[A-Z]+(\d+):[A-Z]+(\d+)/);
            if (match) {
                const startRowNum = parseInt(match[1]);
                const endRowNum = parseInt(match[2]);

                // Wenn der Merge-Bereich in einer gelöschten Zeile liegt
                if (startRowNum > lastUsedRow || endRowNum > lastUsedRow) {
                    keysToRemove.push(key);
                }
            }
        }
        for (const key of keysToRemove) {
            delete worksheet._mergeCells[key];
        }
    }

    // Data Validations entfernen, die in gelöschten Zeilen liegen
    if (worksheet._dataValidations) {
        const keysToRemove = [];
        for (const key of Object.keys(worksheet._dataValidations)) {
            const match = key.match(/[A-Z]+(\d+)/);
            if (match) {
                const rowNum = parseInt(match[1]);
                if (rowNum > lastUsedRow) {
                    keysToRemove.push(key);
                }
            }
        }
        for (const key of keysToRemove) {
            delete worksheet._dataValidations[key];
        }
    }
}

// Hilfsfunktion: Entfernt AutoFilter, bedingte Formatierung und andere Filter-Elemente beim Export
function removeFiltersAndConditionalFormatting(worksheet) {
    // Elemente, die beim Export entfernt werden sollen
    const nodesToRemove = [
        'autoFilter',              // AutoFilter (Dropdown-Pfeile in Kopfzeile)
        'conditionalFormatting',   // Bedingte Formatierung
        'tableParts',              // Tabellenteile (verweisen auf Filter)
        'filterColumn',            // Filter-Spalten-Definitionen
        'colorScale',              // Farbskala (Teil der bedingten Formatierung)
        'dataBar',                 // Datenbalken (Teil der bedingten Formatierung)
        'iconSet'                  // Icon-Set (Teil der bedingten Formatierung)
    ];

    // 1. Haupt-Worksheet-Node bereinigen
    if (worksheet._node && worksheet._node.children) {
        for (let i = worksheet._node.children.length - 1; i >= 0; i--) {
            const child = worksheet._node.children[i];
            if (child && child.name && nodesToRemove.includes(child.name)) {
                worksheet._node.children.splice(i, 1);
            }
        }
    }

    // 2. Interne Referenzen löschen
    if (worksheet._autoFilter) {
        worksheet._autoFilter = null;
    }

    // 3. Alle Zell-Styles auf Hintergrund "none" setzen für nicht-Header Zellen
    // (Optional - macht die Datei "sauber")
}

// Excel-Datei lesen
ipcMain.handle('excel:readFile', async (event, filePath, password = null) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        // Prüfe auf Pivot-Tabellen
        const hasPivotTables = await checkForPivotTables(filePath);

        const options = password ? { password } : {};
        const workbook = await XlsxPopulate.fromFileAsync(filePath, options);
        const sheets = workbook.sheets().map(ws => ws.name());

        return {
            success: true,
            fileName: path.basename(filePath),
            filePath: filePath,
            sheets: sheets,
            isPasswordProtected: !!password,
            hasPivotTables: hasPivotTables
        };
    } catch (error) {
        // Prüfe ob es sich um eine passwortgeschützte Datei handelt
        if (error.message.includes("Can't find end of central directory") ||
            error.message.includes("Encrypted file")) {
            return {
                success: false,
                error: 'Passwort erforderlich',
                isPasswordProtected: true,
                needsPassword: true
            };
        }
        return { success: false, error: error.message };
    }
});

// A/B TEST: ExcelJS vs xlsx-populate
ipcMain.handle('excel:readSheetTest', async (event, filePath, sheetName, password = null) => {
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        console.log('\n========== PERFORMANCE-VERGLEICH ==========');
        
        // Test 1: xlsx-populate (aktuell)
        const populateStart = Date.now();
        const workbook = await XlsxPopulate.fromFileAsync(filePath, password ? { password } : {});
        const populateLoadTime = Date.now() - populateStart;
        
        const worksheet = workbook.sheet(sheetName);
        const usedRange = worksheet?.usedRange();
        const populateRows = usedRange ? (usedRange.endCell().rowNumber() - 1) : 0;
        
        console.log(`[xlsx-populate] ${populateLoadTime}ms für ${populateRows} Zeilen`);
        
        // Test 2: ExcelJS (neu)
        const exceljsStart = Date.now();
        const exceljsResult = await readSheetWithExcelJS(filePath, sheetName, password);
        const exceljsLoadTime = Date.now() - exceljsStart;
        
        if (exceljsResult.success) {
            console.log(`[ExcelJS] ${exceljsLoadTime}ms für ${exceljsResult.data.length} Zeilen`);
            
            const speedup = ((populateLoadTime - exceljsLoadTime) / populateLoadTime * 100).toFixed(1);
            console.log(`[Vergleich] ExcelJS ist ${speedup}% ${speedup > 0 ? 'schneller' : 'langsamer'}`);
            console.log('==========================================\n');
            
            return {
                success: true,
                comparison: {
                    populate: { time: populateLoadTime, rows: populateRows },
                    exceljs: { time: exceljsLoadTime, rows: exceljsResult.data.length },
                    speedupPercent: parseFloat(speedup)
                },
                data: exceljsResult
            };
        } else {
            return { success: false, error: exceljsResult.error };
        }
        
    } catch (error) {
        console.error('[Test] Fehler:', error);
        return { success: false, error: error.message };
    }
});

// Sheet-Daten lesen - NEU MIT EXCELJS (50% schneller, besser gewartet)
ipcMain.handle('excel:readSheet', async (event, filePath, sheetName, password = null, quickLoad = false) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        // Nutze ExcelJS Reader
        const result = await readSheetWithExcelJS(filePath, sheetName, password);
        
        if (!result.success) {
            return result;
        }
        
        // Füge zusätzliche Felder hinzu für Kompatibilität
        const finalResult = {
            ...result,
            dataValidations: {}, // TODO: Später implementieren wenn benötigt
            // mergedCells kommt bereits vom ExcelJS-Reader
            mergedCells: result.mergedCells || [],
            // richTextCells kommt bereits vom ExcelJS-Reader
            richTextCells: result.richTextCells || {},
            // autoFilterRange kommt bereits vom ExcelJS-Reader
            autoFilterRange: result.autoFilterRange || null
        };
        
        return finalResult;
        
    } catch (error) {
        // Prüfe ob es sich um eine passwortgeschützte Datei handelt
        if (error.message.includes("Can't find end of central directory") ||
            error.message.includes("Encrypted file") ||
            error.message.includes("password")) {
            return {
                success: false,
                error: 'Passwort erforderlich',
                isPasswordProtected: true,
                needsPassword: true
            };
        }
        return { success: false, error: error.message };
    }
});

// ======================================================================
// Python/openpyxl Export - Behält ALLE Formatierungen bei
// Vorteile gegenüber ExcelJS:
// - Conditional Formatting bleibt vollständig erhalten (2812+ Regeln)
// - Tabellen bleiben erhalten
// - Theme-Farben werden korrekt behandelt
// - Identisches Verhalten wie MS Excel
// ======================================================================
ipcMain.handle('python:exportMultipleSheets', async (event, { sourcePath, originalSourcePath, targetPath, sheets, password = null, sourcePassword = null }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        const startTime = Date.now();
        
        // Export mit Python/openpyxl durchführen
        const result = await pythonBridge.exportMultipleSheets(sourcePath, targetPath, sheets, { password, sourcePassword, originalSourcePath });
        
        const duration = Date.now() - startTime;
        
        if (!result.success) {
            securityLog.log('ERROR', 'PYTHON_EXPORT_FAILED', {
                sourceFile: path.basename(sourcePath),
                targetFile: path.basename(targetPath),
                error: result.error
            });
            return { success: false, error: result.error };
        }
        
        securityLog.log('INFO', 'PYTHON_EXPORT_COMPLETED', {
            sourceFile: path.basename(sourcePath),
            targetFile: path.basename(targetPath),
            sheetsExported: result.sheetsExported,
            method: result.method || 'openpyxl',
            timeMs: duration
        });

        // Netzwerk-Log für Quelldatei (falls auf Netzlaufwerk)
        await networkLog.log(sourcePath, 'PYTHON_EXPORT_SOURCE', {
            targetFile: path.basename(targetPath),
            sheetsExported: result.sheetsExported
        });

        // Netzwerk-Log für Zieldatei (falls auf Netzlaufwerk)
        await networkLog.log(targetPath, 'PYTHON_EXPORT_TARGET', {
            sourceFile: path.basename(sourcePath),
            sheetsExported: result.sheetsExported
        });

        return {
            success: true,
            message: result.message,
            sheetsExported: result.sheetsExported,
            passwordProtected: !!password,
            stats: { totalTimeMs: duration }
        };
    } catch (error) {
        console.error('[Python Export] Exception:', error);
        securityLog.log('ERROR', 'PYTHON_EXPORT_EXCEPTION', {
            sourceFile: path.basename(sourcePath),
            targetFile: path.basename(targetPath),
            error: error.message
        });
        return { success: false, error: error.message };
    }
});

// ======================================================================
// EXCEL ENGINE STEUERUNG
// Ermöglicht das Umschalten zwischen xlwings und openpyxl
// ======================================================================
ipcMain.handle('excel:checkAvailable', async () => {
    try {
        return await pythonBridge.checkExcelAvailable();
    } catch (error) {
        return { success: false, excelAvailable: false, error: error.message };
    }
});

ipcMain.handle('excel:setEngine', async (event, engine) => {
    try {
        pythonBridge.setExcelEngine(engine);
        console.log(`[App] Excel-Engine auf "${engine}" gesetzt`);
        securityLog.log('INFO', 'EXCEL_ENGINE_CHANGED', { engine });
        
        // Status nach Änderung zurückgeben
        return await pythonBridge.checkExcelAvailable();
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('excel:getEngine', async () => {
    try {
        return { success: true, engine: pythonBridge.getExcelEngine() };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

/*
// ======================================================================
// ALTE XLSX-POPULATE VERSION - BACKUP (wird nicht mehr verwendet)
// ======================================================================
// Sheet-Daten lesen (xlsx-populate - Standard)
ipcMain.handle('excel:readSheet', async (event, filePath, sheetName, password = null, quickLoad = false) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        console.log(`[Load] Sheet "${sheetName}" - Quick Load: ${quickLoad}`);
        const loadStart = Date.now();
        
        // Versuche Workbook aus Cache zu laden
        let workbook = getCachedWorkbook(filePath, password);
        
        if (!workbook) {
            // Nicht im Cache - neu laden
            console.log('[Load] Workbook von Datei laden...');
            const parseStart = Date.now();
            const options = password ? { password } : {};
            workbook = await XlsxPopulate.fromFileAsync(filePath, options);
            const parseTime = Date.now() - parseStart;
            console.log(`[Load] Workbook geparst in ${parseTime}ms`);
            
            // Im Cache speichern
            setCachedWorkbook(filePath, password, workbook);
        }
        
        const worksheet = workbook.sheet(sheetName);

        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }

        // Benutzte Range ermitteln
        const usedRange = worksheet.usedRange();
        if (!usedRange) {
            return { success: true, headers: [], data: [] };
        }

        const startRow = usedRange.startCell().rowNumber();
        const endRow = usedRange.endCell().rowNumber();
        const startCol = usedRange.startCell().columnNumber();
        const endCol = usedRange.endCell().columnNumber();

        const data = [];
        const headers = [];
        const hiddenColumns = []; // Indices der versteckten Spalten
        const hiddenRows = []; // Indices der versteckten Zeilen (0-basiert, ohne Header)
        const cellStyles = {}; // Styles für jede Zelle: "row-col" -> { bold, italic, fill, fontColor, ... }
        const cellFormulas = {}; // Formeln für jede Zelle: "row-col" -> "=FORMULA"
        const cellHyperlinks = {}; // Hyperlinks für jede Zelle: "row-col" -> "https://..."
        const richTextCells = {}; // Rich Text für Zellen: "row-col" -> [{ text, styles: { bold, italic, ... } }, ...]
        let autoFilterRange = null; // AutoFilter-Bereich falls vorhanden

        // Hilfsfunktion: Farbe zu CSS konvertieren
        function colorToCSS(color) {
            if (!color) return null;

            // xlsx-populate gibt Farben in verschiedenen Formaten zurück
            if (typeof color === 'string') {
                // Bereits ein Hex-String
                if (color.match(/^[0-9A-Fa-f]{6,8}$/)) {
                    // ARGB oder RGB Format
                    if (color.length === 8) {
                        // ARGB - ignoriere Alpha
                        return '#' + color.substring(2);
                    }
                    return '#' + color;
                }
                return color;
            }

            if (typeof color === 'object') {
                // Objekt mit rgb oder theme
                if (color.rgb) {
                    const rgb = color.rgb;
                    if (rgb.length === 8) {
                        return '#' + rgb.substring(2);
                    }
                    return '#' + rgb;
                }
                if (color.theme !== undefined) {
                    // Theme-Farben - verwende Standard-Farben
                    const themeColors = [
                        '#000000', // 0 - dark1
                        '#FFFFFF', // 1 - light1
                        '#44546A', // 2 - dark2
                        '#E7E6E6', // 3 - light2
                        '#4472C4', // 4 - accent1
                        '#ED7D31', // 5 - accent2
                        '#A5A5A5', // 6 - accent3
                        '#FFC000', // 7 - accent4
                        '#5B9BD5', // 8 - accent5
                        '#70AD47'  // 9 - accent6
                    ];
                    return themeColors[color.theme] || null;
                }
            }

            return null;
        }

        // Hilfsfunktion: Excel-Datum zu lesbarem String konvertieren
        function excelDateToString(excelDate) {
            // Excel-Datum: Tage seit 1.1.1900 (mit falschem Schaltjahr 1900)
            // JavaScript: Millisekunden seit 1.1.1970
            const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 30.12.1899 UTC
            const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);

            // Pr�fen ob es ein reines Datum oder Datum mit Uhrzeit ist
            const hasTime = (excelDate % 1) !== 0;

            if (hasTime) {
                // Datum mit Uhrzeit
                const day = String(jsDate.getUTCDate()).padStart(2, '0');
                const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                const year = jsDate.getUTCFullYear();
                const hours = String(jsDate.getUTCHours()).padStart(2, '0');
                const minutes = String(jsDate.getUTCMinutes()).padStart(2, '0');
                return `${day}.${month}.${year} ${hours}:${minutes}`;
            } else {
                // Nur Datum
                const day = String(jsDate.getUTCDate()).padStart(2, '0');
                const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                const year = jsDate.getUTCFullYear();
                return `${day}.${month}.${year}`;
            }
        }

        // Hilfsfunktion: Prüfen ob ein Zellwert ein Datum ist
        function isExcelDate(cell, value) {
            if (typeof value !== 'number') return false;

            // Prüfe das Zahlenformat der Zelle
            let numFmt;
            try {
                numFmt = cell.style('numberFormat');
            } catch (e) {
                numFmt = null;
            }

            if (numFmt && typeof numFmt === 'string') {
                // Standard-Excel-Datumsformate (Format-IDs als Strings)
                // Diese werden von xlsx-populate oft als Strings zur�ckgegeben
                const dateFormatIds = [
                    '14', '15', '16', '17', '18', '19', '20', '21', '22',
                    '45', '46', '47', '27', '30', '36', '50', '57'
                ];

                // Pr�fe auf numerische Format-ID
                if (dateFormatIds.includes(String(numFmt))) {
                    return true;
                }

                // Explizite Nicht-Datum-Formate
                const nonDatePatterns = [
                    /^General$/i,
                    /^[#0,]+(\.[#0]+)?$/,        // Zahlenformat wie #,##0.00
                    /^[#0,]+(\.[#0]+)?%$/,       // Prozent
                    /%/,                          // Prozentzeichen
                    /�|EUR|\$/,                   // W�hrung
                    /^@$/,                        // Text
                    /^\[.*?\][#0]/,               // Buchhaltungsformat
                ];

                for (const pattern of nonDatePatterns) {
                    if (pattern.test(numFmt)) {
                        return false;
                    }
                }

                // Typische Datumsformate erkennen (Strings)
                const datePatterns = [
                    /d+[\/\-.\s]m+[\/\-.\s]y+/i,     // d.m.y, d/m/y, d-m-y, d m y
                    /m+[\/\-.\s]d+[\/\-.\s]y+/i,     // m/d/y (US-Format)
                    /y+[\/\-.\s]m+[\/\-.\s]d+/i,     // y-m-d (ISO-Format)
                    /dd\.mm\.yyyy/i,                  // Deutsches Format
                    /dd\/mm\/yyyy/i,                  // Britisches Format
                    /mm\/dd\/yyyy/i,                  // US-Format
                    /yyyy-mm-dd/i,                    // ISO-Format
                    /\[.*?\]dd/i,                     // Benutzerdefinierte Formate
                    /mmm/i,                           // Monatsname (mmm, mmmm)
                    /^d+$/i,                          // Nur "d" oder "dd"
                    /^[$-].*d.*m.*y/i,                // Locale-spezifische Formate
                    /[$-F800]/,                       // Windows Locale Format
                    /[$-407]/,                        // Deutsches Locale
                ];

                for (const pattern of datePatterns) {
                    if (pattern.test(numFmt)) {
                        return true;
                    }
                }
            }

            // Heuristik f�r Werte ohne explizites Format oder mit "General"
            // Excel-Datum: 1 = 1.1.1900, 44197 = 1.1.2021, 47848 = 1.1.2031
            // Typischer Bereich f�r aktuelle Daten: 35000 (1995) bis 55000 (2050)
            if (value >= 1 && value <= 73050) {
                // Nur ganzzahlige Werte oder Werte mit Zeitanteil pr�fen
                // Sehr kleine Zahlen (< 365) sind wahrscheinlich keine Daten
                if (value < 365) {
                    return false; // Wahrscheinlich eine normale Zahl (Tage im Jahr etc.)
                }

                // Pr�fe ob es vern�nftig aussieht
                // Moderne Daten liegen zwischen 30000 (1982) und 55000 (2050)
                if (value >= 30000 && value <= 55000) {
                    // Wenn kein explizites Nicht-Datum-Format, k�nnte es ein Datum sein
                    if (!numFmt || numFmt === 'General' || numFmt === 'general') {
                        // Zus�tzliche Heuristik: Ganzzahlige Werte in diesem Bereich
                        // sind sehr wahrscheinlich Daten
                        if (Number.isInteger(value)) {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        const rowCount = endRow - startRow + 1;
        const colCount = endCol - startCol + 1;
        console.log(`[Load] Sheet-Größe: ${rowCount} Zeilen x ${colCount} Spalten`);
        
        const loopStart = Date.now();
        
        for (let row = startRow; row <= endRow; row++) {
            const rowData = [];

            // Prüfe ob die Zeile in Excel versteckt ist (nur für Datenzeilen, nicht Header)
            if (row > startRow) {
                try {
                    const rowObj = worksheet.row(row);
                    if (rowObj && rowObj.hidden()) {
                        // Zeilen-Index 0-basiert, ohne Header
                        hiddenRows.push(row - startRow - 1);
                    }
                } catch (e) {
                    // Zeile existiert möglicherweise nicht explizit
                }
            }

            for (let col = startCol; col <= endCol; col++) {
                const cell = worksheet.cell(row, col);
                const value = cell.value();

                let textValue = '';
                if (value !== undefined && value !== null) {
                    // Prüfe ob es ein RichText-Objekt ist
                    if (value && typeof value === 'object' && value.constructor && value.constructor.name === 'RichText') {
                        // RichText: Extrahiere Fragmente mit Styles
                        textValue = value.text(); // Gesamttext für die Anzeige

                        // Speichere Fragmente für Datenzeilen (row > startRow)
                        // SKIP bei quickLoad für Performance
                        if (!quickLoad && row > startRow) {
                            const fragments = [];
                            for (let i = 0; i < value.length; i++) {
                                const fragment = value.get(i);
                                const fragmentStyles = {};

                                // Styles aus dem Fragment extrahieren
                                try {
                                    if (fragment.style('bold')) fragmentStyles.bold = true;
                                    if (fragment.style('italic')) fragmentStyles.italic = true;
                                    if (fragment.style('underline')) fragmentStyles.underline = true;
                                    if (fragment.style('strikethrough')) fragmentStyles.strikethrough = true;
                                    if (fragment.style('subscript')) fragmentStyles.subscript = true;
                                    if (fragment.style('superscript')) fragmentStyles.superscript = true;

                                    const fontColor = fragment.style('fontColor');
                                    if (fontColor) {
                                        const cssColor = colorToCSS(fontColor);
                                        if (cssColor && cssColor !== '#000000') {
                                            fragmentStyles.fontColor = cssColor;
                                        }
                                    }

                                    const fontSize = fragment.style('fontSize');
                                    if (fontSize && fontSize !== 11) {
                                        fragmentStyles.fontSize = fontSize;
                                    }
                                } catch (e) {
                                    // Style nicht verfügbar
                                }

                                fragments.push({
                                    text: fragment.value(),
                                    styles: Object.keys(fragmentStyles).length > 0 ? fragmentStyles : null
                                });
                            }

                            // Nur speichern wenn es tatsächlich unterschiedliche Formatierungen gibt
                            const hasVariedStyles = fragments.some(f => f.styles !== null);
                            if (hasVariedStyles) {
                                const richTextKey = `${row - startRow}-${col - 1}`;
                                richTextCells[richTextKey] = fragments;
                            }
                        }
                    } else if (value instanceof Date) {
                        // Falls xlsx-populate bereits ein Date-Objekt zurückgibt
                        const day = String(value.getDate()).padStart(2, '0');
                        const month = String(value.getMonth() + 1).padStart(2, '0');
                        const year = value.getFullYear();
                        textValue = `${day}.${month}.${year}`;
                    } else if (isExcelDate(cell, value)) {
                        textValue = excelDateToString(value);
                    } else {
                        textValue = String(value);
                    }
                }

                // Styles auslesen (nur für Datenzeilen, nicht für Header)
                // SKIP bei quickLoad für Performance
                if (!quickLoad && row > startRow) {
                    try {
                        const style = {};
                        let hasStyle = false;

                        // Bold
                        const bold = cell.style('bold');
                        if (bold) {
                            style.bold = true;
                            hasStyle = true;
                        }

                        // Italic
                        const italic = cell.style('italic');
                        if (italic) {
                            style.italic = true;
                            hasStyle = true;
                        }

                        // Underline
                        const underline = cell.style('underline');
                        if (underline) {
                            style.underline = true;
                            hasStyle = true;
                        }

                        // Strikethrough
                        const strikethrough = cell.style('strikethrough');
                        if (strikethrough) {
                            style.strikethrough = true;
                            hasStyle = true;
                        }

                        // Font Color
                        const fontColor = cell.style('fontColor');
                        if (fontColor) {
                            const cssColor = colorToCSS(fontColor);
                            if (cssColor && cssColor !== '#000000') {
                                style.fontColor = cssColor;
                                hasStyle = true;
                            }
                        }

                        // Fill/Background Color
                        const fill = cell.style('fill');
                        if (fill) {
                            if (typeof fill === 'object') {
                                // xlsx-populate fill Struktur: { type: "solid", color: { rgb: "AARRGGBB" } }
                                let fillColor = null;

                                if (fill.color) {
                                    // color ist ein Objekt mit rgb Property
                                    fillColor = colorToCSS(fill.color);
                                } else if (fill.foreground) {
                                    // Foreground bei manchen Patterns
                                    fillColor = colorToCSS(fill.foreground);
                                }

                                if (fillColor && fillColor !== '#FFFFFF') {
                                    style.fill = fillColor;
                                    hasStyle = true;
                                }
                            }
                        }

                        // Font Size
                        const fontSize = cell.style('fontSize');
                        if (fontSize && fontSize !== 11) { // 11 ist Standard
                            style.fontSize = fontSize;
                            hasStyle = true;
                        }

                        // Horizontal Alignment
                        const hAlign = cell.style('horizontalAlignment');
                        if (hAlign && hAlign !== 'general') {
                            style.textAlign = hAlign;
                            hasStyle = true;
                        }

                        // Speichere nur wenn Style vorhanden
                        if (hasStyle) {
                            const rowIndex = row - startRow; // 0-basiert, inkl. Header
                            cellStyles[`${rowIndex}-${col - 1}`] = style;
                        }
                    } catch (e) {
                        // Style konnte nicht gelesen werden
                    }

                    // Formel auslesen (nur für Datenzeilen)
                    // SKIP bei quickLoad
                    if (!quickLoad) {
                        try {
                            const formula = cell.formula();
                            if (formula) {
                                const rowIndex = row - startRow; // 0-basiert, inkl. Header
                                cellFormulas[`${rowIndex}-${col - 1}`] = formula;
                            }
                        } catch (e) {
                            // Formel konnte nicht gelesen werden
                        }

                        // Hyperlink auslesen (nur für Datenzeilen)
                        try {
                            const hyperlink = cell.hyperlink();
                            if (hyperlink) {
                                const rowIndex = row - startRow; // 0-basiert, inkl. Header
                                cellHyperlinks[`${rowIndex}-${col - 1}`] = hyperlink;
                            }
                        } catch (e) {
                            // Hyperlink konnte nicht gelesen werden
                        }
                    }
                }

                // Header-Zeile (erste Zeile)
                if (row === startRow) {
                    headers[col - 1] = textValue || `Spalte ${col}`;
                    // Prüfe ob die Spalte in Excel versteckt ist
                    try {
                        const column = worksheet.column(col);
                        if (column && column.hidden()) {
                            hiddenColumns.push(col - 1); // 0-basierter Index
                        }
                    } catch (e) {
                        // Spalte existiert möglicherweise nicht explizit
                    }
                }
                rowData[col - 1] = textValue;
            }

            // Zeilen auffuellen bis zur maximalen Spaltenanzahl
            while (rowData.length < headers.length) {
                rowData.push('');
            }

            data.push(rowData);
        }

        // Data Validations (Dropdown-Listen) auslesen
        const dataValidations = {};
        try {
            // xlsx-populate speichert Data Validations im Sheet-Objekt
            // Wir iterieren über alle Zellen und prüfen auf dataValidation
            for (let col = startCol; col <= endCol; col++) {
                const colValidations = [];
                let hasValidation = false;

                for (let row = startRow; row <= endRow; row++) {
                    const cell = worksheet.cell(row, col);
                    try {
                        const validation = cell.dataValidation();
                        if (validation && validation.type === 'list') {
                            hasValidation = true;
                            let allowedValues = [];

                            // Explizite Werte-Liste
                            if (validation.formula1) {
                                const formula = validation.formula1;
                                // Prüfe ob es eine Referenz oder eine Liste ist
                                if (formula.startsWith('"') && formula.endsWith('"')) {
                                    // Explizite Liste: "Wert1,Wert2,Wert3"
                                    allowedValues = formula.slice(1, -1).split(',').map(v => v.trim());
                                } else if (formula.includes(':')) {
                                    // Bereichsreferenz: Sheet1!$A$1:$A$10 oder $A$1:$A$10
                                    try {
                                        // Versuche den Bereich aufzulösen
                                        const rangeValues = [];
                                        let targetSheet = worksheet;
                                        let rangeRef = formula;

                                        // Prüfe auf Sheet-Referenz
                                        if (formula.includes('!')) {
                                            const parts = formula.split('!');
                                            const refSheetName = parts[0].replace(/'/g, ''); // Entferne Anführungszeichen
                                            rangeRef = parts[1];
                                            targetSheet = workbook.sheet(refSheetName);
                                        }

                                        if (targetSheet) {
                                            // Entferne $ Zeichen und parse den Bereich
                                            const cleanRef = rangeRef.replace(/\$/g, '');
                                            const range = targetSheet.range(cleanRef);
                                            if (range) {
                                                range.forEach(c => {
                                                    const val = c.value();
                                                    if (val !== undefined && val !== null && val !== '') {
                                                        rangeValues.push(String(val));
                                                    }
                                                });
                                            }
                                        }
                                        allowedValues = rangeValues;
                                    } catch (e) {
                                        // Bereich konnte nicht aufgelöst werden
                                    }
                                } else {
                                    // Einfache Formel oder Liste ohne Anführungszeichen
                                    allowedValues = formula.split(',').map(v => v.trim());
                                }
                            }

                            if (allowedValues.length > 0) {
                                // Speichere für diese Zeile (0-basiert, -1 weil startRow = Header)
                                const rowIndex = row - startRow;
                                colValidations.push({
                                    row: rowIndex,
                                    values: allowedValues,
                                    allowBlank: validation.allowBlank !== false
                                });
                            }
                        }
                    } catch (e) {
                        // Zelle hat keine Validation oder Fehler beim Lesen
                    }
                }

                if (hasValidation && colValidations.length > 0) {
                    // Prüfe ob alle Zeilen die gleichen Werte haben (spaltenweite Validation)
                    const firstValues = JSON.stringify(colValidations[0].values);
                    const allSame = colValidations.every(v => JSON.stringify(v.values) === firstValues);

                    if (allSame && colValidations.length > 1) {
                        // Spaltenweite Validation - alle Zeilen haben gleiche Optionen
                        dataValidations[col - 1] = {
                            type: 'column',
                            values: colValidations[0].values,
                            allowBlank: colValidations[0].allowBlank
                        };
                    } else {
                        // Zeilenspezifische Validations
                        dataValidations[col - 1] = {
                            type: 'rows',
                            rows: colValidations.reduce((acc, v) => {
                                acc[v.row] = { values: v.values, allowBlank: v.allowBlank };
                                return acc;
                            }, {})
                        };
                    }
                }
            }
        } catch (e) {
            // Data Validations konnten nicht gelesen werden
        }

        // AutoFilter auslesen
        try {
            const sheetNode = worksheet._node;
            if (sheetNode && sheetNode.children) {
                for (const child of sheetNode.children) {
                    if (child && child.name === 'autoFilter' && child.attributes && child.attributes.ref) {
                        autoFilterRange = child.attributes.ref;
                        break;
                    }
                }
            }
        } catch (e) {
            // AutoFilter konnte nicht gelesen werden
        }

        // Merged Cells auslesen
        const mergedCells = [];
        try {
            // xlsx-populate speichert mergeCells in sheet._mergeCells
            const mergeCellsMap = worksheet._mergeCells;
            if (mergeCellsMap && typeof mergeCellsMap === 'object') {
                // Konvertiere Excel-Referenzen zu 0-basierten Indizes
                const parseRef = (cellRef) => {
                    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
                    if (match) {
                        let col = 0;
                        for (let i = 0; i < match[1].length; i++) {
                            col = col * 26 + (match[1].charCodeAt(i) - 64);
                        }
                        return { row: parseInt(match[2]), col: col };
                    }
                    return null;
                };

                for (const ref of Object.keys(mergeCellsMap)) {
                    // ref ist z.B. "A1:C3"
                    const parts = ref.split(':');
                    if (parts.length === 2) {
                        const start = parseRef(parts[0]);
                        const end = parseRef(parts[1]);

                        if (start && end) {
                            // Speichere als 0-basierte Excel-Zeilen-Indizes
                            // Excel-Zeile 1 → Index 0, Excel-Zeile 2 → Index 1, etc.
                            mergedCells.push({
                                startRow: start.row - 1, // Excel 1-basiert → 0-basiert
                                startCol: start.col - 1, // 0-basiert
                                endRow: end.row - 1,     // Excel 1-basiert → 0-basiert
                                endCol: end.col - 1,
                                rowSpan: end.row - start.row + 1,
                                colSpan: end.col - start.col + 1
                            });
                        }
                    }
                }
            }
        } catch (e) {
            // Merged Cells konnten nicht gelesen werden
        }

        const loopTime = Date.now() - loopStart;
        console.log(`[Load] Zellen-Loop: ${loopTime}ms`);

        const loadTime = Date.now() - loadStart;
        console.log(`[Load] GESAMT: ${loadTime}ms (${data.length} Zeilen, QuickLoad: ${quickLoad})`);

        return {
            success: true,
            headers: headers,
            data: data,
            hiddenColumns: hiddenColumns,
            hiddenRows: hiddenRows,
            dataValidations: dataValidations,
            cellStyles: cellStyles,
            cellFormulas: cellFormulas,
            cellHyperlinks: cellHyperlinks,
            richTextCells: richTextCells,
            autoFilterRange: autoFilterRange,
            mergedCells: mergedCells
        };
    } catch (error) {
        // Prüfe ob es sich um eine passwortgeschützte Datei handelt
        if (error.message.includes("Can't find end of central directory") ||
            error.message.includes("Encrypted file")) {
            return {
                success: false,
                error: 'Passwort erforderlich',
                isPasswordProtected: true,
                needsPassword: true
            };
        }
        return { success: false, error: error.message };
    }
});
*/

// ==================== SHEET-VERWALTUNG ====================

// Neues Arbeitsblatt hinzufügen
ipcMain.handle('excel:addSheet', async (event, { filePath, sheetName }) => {
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);

        // Prüfe ob Name bereits existiert
        const existingSheet = workbook.sheet(sheetName);
        if (existingSheet) {
            return { success: false, error: 'Ein Arbeitsblatt mit diesem Namen existiert bereits' };
        }

        workbook.addSheet(sheetName);
        await saveWorkbookOptimized(workbook, filePath, {}, filePath);

        securityLog.log('INFO', 'SHEET_ADDED', {
            file: path.basename(filePath),
            sheet: sheetName
        });

        return {
            success: true,
            sheets: workbook.sheets().map(s => s.name())
        };
    } catch (error) {
        securityLog.log('ERROR', 'SHEET_ADD_FAILED', { error: error.message });
        return { success: false, error: error.message };
    }
});

// Arbeitsblatt löschen
ipcMain.handle('excel:deleteSheet', async (event, { filePath, sheetName }) => {
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);

        // Mindestens ein Blatt muss bleiben
        if (workbook.sheets().length <= 1) {
            return { success: false, error: 'Das letzte Arbeitsblatt kann nicht gelöscht werden' };
        }

        workbook.deleteSheet(sheetName);
        await saveWorkbookOptimized(workbook, filePath, {}, filePath);

        securityLog.log('INFO', 'SHEET_DELETED', {
            file: path.basename(filePath),
            sheet: sheetName
        });

        return {
            success: true,
            sheets: workbook.sheets().map(s => s.name())
        };
    } catch (error) {
        securityLog.log('ERROR', 'SHEET_DELETE_FAILED', { error: error.message });
        return { success: false, error: error.message };
    }
});

// Arbeitsblatt umbenennen
ipcMain.handle('excel:renameSheet', async (event, { filePath, oldName, newName }) => {
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);

        // Prüfe ob neuer Name bereits existiert
        const existingSheet = workbook.sheet(newName);
        if (existingSheet) {
            return { success: false, error: 'Ein Arbeitsblatt mit diesem Namen existiert bereits' };
        }

        const sheet = workbook.sheet(oldName);
        if (!sheet) {
            return { success: false, error: 'Arbeitsblatt nicht gefunden' };
        }

        sheet.name(newName);
        await saveWorkbookOptimized(workbook, filePath, {}, filePath);

        return {
            success: true,
            sheets: workbook.sheets().map(s => s.name())
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Arbeitsblatt kopieren/klonen
ipcMain.handle('excel:cloneSheet', async (event, { filePath, sheetName, newName }) => {
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);

        // Prüfe ob neuer Name bereits existiert
        const existingSheet = workbook.sheet(newName);
        if (existingSheet) {
            return { success: false, error: 'Ein Arbeitsblatt mit diesem Namen existiert bereits' };
        }

        const sheetToClone = workbook.sheet(sheetName);
        if (!sheetToClone) {
            return { success: false, error: 'Arbeitsblatt nicht gefunden' };
        }

        workbook.cloneSheet(sheetToClone, newName);
        await saveWorkbookOptimized(workbook, filePath, {}, filePath);

        return {
            success: true,
            sheets: workbook.sheets().map(s => s.name())
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Arbeitsblatt verschieben (Reihenfolge ändern)
ipcMain.handle('excel:moveSheet', async (event, { filePath, sheetName, newIndex }) => {
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);

        workbook.moveSheet(sheetName, newIndex);
        await saveWorkbookOptimized(workbook, filePath, {}, filePath);

        return {
            success: true,
            sheets: workbook.sheets().map(s => s.name())
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Zeilen in Excel einfuegen (MIT Formatierungserhalt dank xlsx-populate!)
ipcMain.handle('excel:insertRows', async (event, { filePath, sheetName, rows, startColumn, enableFlag = true, enableComment = true, flagColumn = 1, commentColumn = 2, sourceFilePath = null, sourceSheetName = null, sourceColumns = [] }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const worksheet = workbook.sheet(sheetName);

        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }

        // Hilfsfunktion: Deutsches Datum zu Excel-Datum konvertieren
        function parseGermanDateToExcel(dateStr) {
            if (!dateStr || typeof dateStr !== 'string') return null;

            // Deutsches Datum: dd.mm.yyyy oder dd.mm.yyyy hh:mm
            const dateTimeMatch = dateStr.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})$/);
            const dateMatch = dateStr.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);

            if (dateTimeMatch) {
                const day = parseInt(dateTimeMatch[1], 10);
                const month = parseInt(dateTimeMatch[2], 10);
                const year = parseInt(dateTimeMatch[3], 10);
                const hours = parseInt(dateTimeMatch[4], 10);
                const minutes = parseInt(dateTimeMatch[5], 10);

                // Validierung
                if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900 || year > 2100) {
                    return null;
                }

                // Excel-Datum berechnen (UTC um Zeitzonenproblemen vorzubeugen)
                const jsDate = Date.UTC(year, month - 1, day, hours, minutes);
                const excelEpoch = Date.UTC(1899, 11, 30);
                const excelDate = (jsDate - excelEpoch) / 86400000;
                return excelDate;
            } else if (dateMatch) {
                const day = parseInt(dateMatch[1], 10);
                const month = parseInt(dateMatch[2], 10);
                const year = parseInt(dateMatch[3], 10);

                // Validierung
                if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900 || year > 2100) {
                    return null;
                }

                // Excel-Datum berechnen (UTC)
                const jsDate = Date.UTC(year, month - 1, day);
                const excelEpoch = Date.UTC(1899, 11, 30);
                const excelDate = Math.floor((jsDate - excelEpoch) / 86400000);
                return excelDate;
            }

            return null;
        }

        // Hilfsfunktion: Wert intelligent konvertieren
        function convertValue(value, targetCell) {
            if (value === null || value === undefined || value === '') {
                return { value: value, isDate: false };
            }

            // Pr�fe ob es ein deutsches Datum ist
            const excelDate = parseGermanDateToExcel(value);
            if (excelDate !== null) {
                return { value: excelDate, isDate: true };
            }

            // Pr�fe ob es eine Zahl ist (mit deutschen Dezimaltrennzeichen)
            if (typeof value === 'string') {
                // Deutsche Zahlen: 1.234,56 -> 1234.56
                const germanNumberMatch = value.match(/^-?\d{1,3}(\.\d{3})*(,\d+)?$/);
                if (germanNumberMatch) {
                    const normalized = value.replace(/\./g, '').replace(',', '.');
                    const num = parseFloat(normalized);
                    if (!isNaN(num)) {
                        return { value: num, isDate: false };
                    }
                }

                // Englische Zahlen oder einfache Zahlen
                const simpleNumber = value.match(/^-?\d+(\.\d+)?$/);
                if (simpleNumber) {
                    const num = parseFloat(value);
                    if (!isNaN(num)) {
                        return { value: num, isDate: false };
                    }
                }
            }

            // Als String belassen
            return { value: value, isDate: false };
        }

        // Formatvorlage aus Header-Zeile oder vorhandenen Zeilen ermitteln
        function getColumnFormat(colNumber) {
            // Suche nach dem ersten nicht-leeren Wert in dieser Spalte (ab Zeile 2)
            const usedRange = worksheet.usedRange();
            if (!usedRange) return null;

            const endRow = Math.min(usedRange.endCell().rowNumber(), 100); // Max 100 Zeilen pr�fen

            for (let row = 2; row <= endRow; row++) {
                const cell = worksheet.cell(row, colNumber);
                const value = cell.value();
                if (value !== undefined && value !== null && value !== '') {
                    const numFmt = cell.style('numberFormat');
                    if (numFmt && numFmt !== 'General') {
                        return numFmt;
                    }
                }
            }
            return null;
        }

        // Erste leere Zeile finden (ab Zeile 2, da Zeile 1 = Header)
        let insertRow = 2;  // Standard: direkt nach Header

        const usedRange = worksheet.usedRange();
        if (usedRange) {
            const endRow = usedRange.endCell().rowNumber();

            // Pr�fe ab Zeile 2, ob es Daten gibt
            for (let row = 2; row <= endRow; row++) {
                // Prüfe Flag-Spalte (wenn aktiviert) oder Datenspalte
                const checkCol = enableFlag ? flagColumn : startColumn;
                const flagCell = worksheet.cell(row, checkCol).value();
                const dataCell = worksheet.cell(row, startColumn).value();

                // Zeile ist leer wenn beide Zellen leer sind
                const flagEmpty = flagCell === undefined || flagCell === null || flagCell === '';
                const dataEmpty = dataCell === undefined || dataCell === null || dataCell === '';

                if (flagEmpty && dataEmpty) {
                    // Erste leere Zeile gefunden
                    insertRow = row;
                    break;
                }
                // Wenn wir hier sind, ist die Zeile nicht leer - gehe zur nächsten
                insertRow = row + 1;
            }
        }

        // Spaltenformate vorab ermitteln
        const columnFormats = {};

        // Formatierungsvorlage: Letzte belegte Zeile vor der Einfügeposition
        const templateRow = insertRow > 2 ? insertRow - 1 : 2;

        // Hilfsfunktion: Kopiere Zellformatierung von Template-Zeile der Zieldatei
        function copyStyleFromTemplate(targetCell, colNumber) {
            try {
                const templateCell = worksheet.cell(templateRow, colNumber);
                if (!templateCell) return;

                // Verfügbare Styles in xlsx-populate
                const styles = [
                    'bold', 'italic', 'underline', 'strikethrough',
                    'fontFamily', 'fontSize', 'fontColor',
                    'horizontalAlignment', 'verticalAlignment',
                    'wrapText', 'numberFormat'
                ];

                styles.forEach(styleName => {
                    try {
                        const styleValue = templateCell.style(styleName);
                        if (styleValue !== undefined && styleValue !== null) {
                            targetCell.style(styleName, styleValue);
                        }
                    } catch (e) {
                        // Ignoriere Fehler bei einzelnen Styles
                    }
                });

                // Fill (Hintergrundfarbe) separat behandeln - ist ein komplexes Objekt
                try {
                    const fillValue = templateCell.style('fill');
                    if (fillValue && typeof fillValue === 'object') {
                        // Deep copy des Fill-Objekts um Referenzprobleme zu vermeiden
                        const fillCopy = JSON.parse(JSON.stringify(fillValue));
                        targetCell.style('fill', fillCopy);
                    }
                } catch (e) {
                    // Ignoriere Fehler beim Kopieren der Fill-Formatierung
                }
            } catch (e) {
                // Ignoriere Fehler
            }
        }

        // Neue Zeilen einfuegen
        let insertedCount = 0;
        for (const row of rows) {
            const newRowNum = insertRow + insertedCount;

            // Bei Leerzeile: Zeile als "belegt" markieren
            if (row.flag === 'leer') {
                // Flag-Spalte setzen wenn aktiviert
                if (enableFlag) {
                    const flagCell = worksheet.cell(newRowNum, flagColumn);
                    copyStyleFromTemplate(flagCell, flagColumn);
                    flagCell.value(' ');
                } else {
                    // Wenn Flag deaktiviert, Leerzeichen in erste Datenspalte setzen
                    // damit die Zeile als "belegt" gilt und nicht überschrieben wird
                    const firstDataCell = worksheet.cell(newRowNum, startColumn);
                    copyStyleFromTemplate(firstDataCell, startColumn);
                    firstDataCell.value(' ');
                }
                // Kommentar trotzdem schreiben wenn vorhanden und aktiviert
                if (enableComment && row.comment) {
                    const commentCell = worksheet.cell(newRowNum, commentColumn);
                    copyStyleFromTemplate(commentCell, commentColumn);
                    commentCell.value(row.comment);
                }
                insertedCount++;
                continue;
            }

            // Flag in konfigurierter Spalte (nur wenn aktiviert)
            if (enableFlag && row.flag) {
                const flagCell = worksheet.cell(newRowNum, flagColumn);
                copyStyleFromTemplate(flagCell, flagColumn);
                flagCell.value(row.flag);
            }

            // Kommentar in konfigurierter Spalte (nur wenn aktiviert)
            if (enableComment && row.comment) {
                const commentCell = worksheet.cell(newRowNum, commentColumn);
                copyStyleFromTemplate(commentCell, commentColumn);
                commentCell.value(row.comment);
            }

            // Daten ab Startspalte - row.data ist ein Objekt mit Index als Key
            if (row.data) {
                const dataKeys = Object.keys(row.data);

                dataKeys.forEach(key => {
                    const index = parseInt(key);
                    const value = row.data[key];
                    if (value !== null && value !== undefined && value !== '') {
                        const colNumber = startColumn + index;
                        const targetCell = worksheet.cell(newRowNum, colNumber);

                        // Formatierung von Template-Zeile der Zieldatei kopieren
                        // (Quelldatei kann bedingte Formatierungen nicht liefern)
                        copyStyleFromTemplate(targetCell, colNumber);

                        const converted = convertValue(value, targetCell);
                        targetCell.value(converted.value);

                        // Wenn es ein Datum ist, prüfe ob schon ein Format von Template kopiert wurde
                        if (converted.isDate) {
                            const currentFormat = targetCell.style('numberFormat');
                            // Nur setzen wenn kein Format vorhanden (Template hatte auch keins)
                            if (!currentFormat || currentFormat === 'General') {
                                // Spaltenformat aus Cache oder neu ermitteln
                                if (!(colNumber in columnFormats)) {
                                    columnFormats[colNumber] = getColumnFormat(colNumber);
                                }

                                const colFormat = columnFormats[colNumber];
                                if (colFormat) {
                                    targetCell.style('numberFormat', colFormat);
                                } else {
                                    // Standard deutsches Datumsformat setzen
                                    targetCell.style('numberFormat', 'DD.MM.YYYY');
                                }
                            }
                            // Wenn Format vom Template/Source kopiert wurde, behalte es!
                        }
                    }
                });
            }

            insertedCount++;
        }

        // Speichern (xlsx-populate erhält die originale Formatierung!)
        await saveWorkbookOptimized(workbook, filePath, {}, filePath);

        // Netzwerk-Log für Datenübertragung (falls auf Netzlaufwerk)
        await networkLog.log(filePath, 'DATA_TRANSFER', {
            file: path.basename(filePath),
            sheet: sheetName,
            rowsInserted: insertedCount,
            startRow: insertRow,
            sourceFile: sourceFilePath ? path.basename(sourceFilePath) : null,
            sourceSheet: sourceSheetName
        });

        // Auch Quelldatei loggen (falls auf anderem Netzlaufwerk)
        if (sourceFilePath) {
            await networkLog.log(sourceFilePath, 'DATA_TRANSFER_SOURCE', {
                targetFile: path.basename(filePath),
                targetSheet: sheetName,
                rowsTransferred: insertedCount
            });
        }

        return {
            success: true,
            message: `${insertedCount} Zeile(n) ab Zeile ${insertRow} eingefuegt`,
            insertedCount: insertedCount,
            startRow: insertRow
        };
    } catch (error) {
        console.error('Fehler beim Einf�gen:', error);
        return { success: false, error: error.message };
    }
});

// Datei kopieren (fuer "Neuer Monat") - BINÄRE KOPIE erhält 100% Formatierung!
ipcMain.handle('excel:copyFile', async (event, { sourcePath, targetPath, sheetName, keepHeader }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        // Prüfe ob Quelldatei existiert
        if (!fs.existsSync(sourcePath)) {
            return { success: false, error: `Quelldatei nicht gefunden: ${sourcePath}` };
        }

        // Wenn keepHeader false ist oder nicht gesetzt, einfach binär kopieren
        // Das erhält 100% der Formatierung!
        if (!keepHeader) {
            try {
                fs.copyFileSync(sourcePath, targetPath);
            } catch (copyErr) {
                return { success: false, error: `Kopieren fehlgeschlagen: ${copyErr.message}` };
            }
            return { success: true, message: `Datei kopiert: ${targetPath}` };
        }

        // Wenn keepHeader true ist, müssen wir die Daten löschen
        // Aber zuerst: Binär kopieren, dann nur die Werte löschen
        try {
            fs.copyFileSync(sourcePath, targetPath);
        } catch (copyErr) {
            return { success: false, error: `Kopieren fehlgeschlagen: ${copyErr.message}` };
        }

        // Jetzt die kopierte Datei �ffnen und nur die Datenwerte l�schen
        const workbook = await XlsxPopulate.fromFileAsync(targetPath);

        // Wenn sheetName angegeben und existiert, nutze dieses Sheet
        // Ansonsten nimm das erste Sheet
        let worksheet = null;
        if (sheetName) {
            worksheet = workbook.sheet(sheetName);
        }
        if (!worksheet) {
            // Erstes verfuegbares Sheet nehmen
            worksheet = workbook.sheets()[0];
        }

        if (!worksheet) {
            return { success: false, error: 'Keine Worksheets in der Template-Datei gefunden' };
        }

        // Nur die Werte ab Zeile 2 l�schen (Header in Zeile 1 bleibt)
        // Formatierung bleibt erhalten!
        const usedRange = worksheet.usedRange();
        if (usedRange) {
            const endRow = usedRange.endCell().rowNumber();
            const endCol = usedRange.endCell().columnNumber();

            // Alle Datenwerte ab Zeile 2 l�schen
            for (let row = 2; row <= endRow; row++) {
                for (let col = 1; col <= endCol; col++) {
                    worksheet.cell(row, col).value(undefined);
                }
            }
        }

        await saveWorkbookOptimized(workbook, targetPath, {});

        return { success: true, message: `Datei erstellt: ${targetPath}` };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Daten exportieren (fuer Datenexplorer) - nur ein Sheet
ipcMain.handle('excel:exportData', async (event, { filePath, headers, data }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        // Neue leere Workbook erstellen
        const workbook = await XlsxPopulate.fromBlankAsync();

        // Erstes Sheet umbenennen
        const worksheet = workbook.sheet(0);
        worksheet.name('Export');

        // Header-Zeile
        headers.forEach((header, colIndex) => {
            worksheet.cell(1, colIndex + 1).value(header);
        });

        // Daten-Zeilen (data ist ein Array von Arrays)
        data.forEach((row, rowIndex) => {
            row.forEach((value, colIndex) => {
                worksheet.cell(rowIndex + 2, colIndex + 1).value(value || '');
            });
        });

        await saveWorkbookOptimized(workbook, filePath, {});

        return { success: true, message: `Export erstellt: ${filePath}` };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Daten exportieren MIT allen Sheets (fuer Datenexplorer - Vollexport)
ipcMain.handle('excel:exportWithAllSheets', async (event, { sourcePath, targetPath, sheetName, headers, data, visibleColumns }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        // Originaldatei laden (mit allen Sheets und Formatierung)
        const workbook = await XlsxPopulate.fromFileAsync(sourcePath);
        const allSheets = workbook.sheets().map(s => s.name());

        // Das aktive Sheet finden
        const worksheet = workbook.sheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }

        // Ursprüngliche Zeilenanzahl ermitteln
        const usedRange = worksheet.usedRange();
        let originalRowCount = 0;
        if (usedRange) {
            originalRowCount = usedRange.endCell().rowNumber() - 1; // -1 für Header
            usedRange.clear();
        }

        // ALLE Spalten exportieren (auch ausgeblendete) und Hidden-Attribute setzen
        // Header-Zeile
        headers.forEach((header, colIndex) => {
            worksheet.cell(1, colIndex + 1).value(header);
        });

        // Daten-Zeilen
        data.forEach((row, rowIndex) => {
            row.forEach((value, colIndex) => {
                worksheet.cell(rowIndex + 2, colIndex + 1).value(value || '');
            });
        });

        // Hidden-Attribute für ausgeblendete Spalten setzen
        if (visibleColumns && visibleColumns.length > 0 && visibleColumns.length < headers.length) {
            // Set mit sichtbaren Spalten-Indizes erstellen
            const visibleSet = new Set(visibleColumns);

            // Alle Spalten durchgehen und hidden setzen wo nötig
            for (let colIdx = 0; colIdx < headers.length; colIdx++) {
                const column = worksheet.column(colIdx + 1);
                if (!visibleSet.has(colIdx)) {
                    column.hidden(true);
                } else {
                    column.hidden(false);
                }
            }
        }

        // Hidden-Attribute für ausgeblendete Zeilen setzen (aus hiddenRows)
        // Die hiddenRows werden vom Frontend mitgeschickt

        // Nicht verwendete Zeilen als hidden markieren (wenn weniger Zeilen als ursprünglich)
        if (data.length < originalRowCount) {
            for (let rowIdx = data.length + 2; rowIdx <= originalRowCount + 1; rowIdx++) {
                worksheet.row(rowIdx).hidden(true);
            }
        }

        // Speichern (alle anderen Sheets bleiben unverändert)
        await saveWorkbookOptimized(workbook, targetPath, {}, sourcePath);

        return { success: true, message: `Export erstellt: ${targetPath}`, sheets: allSheets };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Export mit Auswahl der Arbeitsblätter (für Datenexplorer) - behält Formatierung bei
// ============================================================================
// UNIFIED EXCELJS EXPORT - Ersetzt xlsx-populate komplett
// ============================================================================
ipcMain.handle('excel:exportMultipleSheets', async (event, { sourcePath, targetPath, sheets, password = null, sourcePassword = null }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        console.log('[Export] Starte Export mit ExcelJS...');
        
        // Dateigröße für Logging
        const stats = fs.statSync(sourcePath);
        const fileSizeMB = stats.size / (1024 * 1024);
        console.log(`[Export] Dateigröße: ${fileSizeMB.toFixed(2)} MB`);
        
        // Import der neuen ExcelJS-Funktion
        const { exportMultipleSheetsWithExcelJS } = require('./exceljs-writer');
        
        // Export mit ExcelJS durchführen
        const result = await exportMultipleSheetsWithExcelJS(sourcePath, targetPath, sheets, { password, sourcePassword });
        
        if (!result.success) {
            securityLog.log('ERROR', 'EXCEL_EXPORT_FAILED', {
                sourceFile: path.basename(sourcePath),
                targetFile: path.basename(targetPath),
                error: result.error
            });
            return { success: false, error: result.error };
        }
        
        securityLog.log('INFO', 'EXCEL_EXPORT_COMPLETED', {
            sourceFile: path.basename(sourcePath),
            targetFile: path.basename(targetPath),
            sheetsExported: result.sheetsExported,
            method: 'exceljs',
            timeMs: result.stats?.totalTimeMs
        });

        // Netzwerk-Log für Quelldatei (falls auf Netzlaufwerk)
        await networkLog.log(sourcePath, 'EXCEL_EXPORT_SOURCE', {
            targetFile: path.basename(targetPath),
            sheetsExported: result.sheetsExported
        });

        // Netzwerk-Log für Zieldatei (falls auf Netzlaufwerk)
        await networkLog.log(targetPath, 'EXCEL_EXPORT_TARGET', {
            sourceFile: path.basename(sourcePath),
            sheetsExported: result.sheetsExported
        });

        return {
            success: true,
            message: result.message,
            sheetsExported: result.sheetsExported,
            passwordProtected: !!password
        };
    } catch (error) {
        securityLog.log('ERROR', 'EXCEL_EXPORT_FAILED', {
            sourceFile: path.basename(sourcePath),
            targetFile: path.basename(targetPath),
            error: error.message
        });
        return { success: false, error: error.message };
    }
});

// Änderungen direkt in die Originaldatei speichern (für Datenexplorer)
// ============================================================================
// SAVE FILE - Speichert Änderungen direkt in die Quelldatei
// ============================================================================
ipcMain.handle('excel:saveFile', async (event, { filePath, sheets, password = null, sourcePassword = null }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        console.log('[Save] Starte Speichern mit ExcelJS...');
        console.log(`[Save] Datei: ${filePath}`);
        console.log(`[Save] Sheets: ${sheets.length}`);
        
        // Zeige was gesendet wurde
        for (const sheet of sheets) {
            if (sheet.changedCells) {
                console.log(`[Save] Sheet "${sheet.sheetName}": ${Object.keys(sheet.changedCells).length} geänderte Zellen`);
            } else if (sheet.fullRewrite) {
                console.log(`[Save] Sheet "${sheet.sheetName}": Full Rewrite (${sheet.data?.length || 0} Zeilen)`);
            }
        }
        
        // Speichern = Exportieren in dieselbe Datei
        // Nutze exportMultipleSheetsWithExcelJS mit targetPath = sourcePath
        const { exportMultipleSheetsWithExcelJS } = require('./exceljs-writer');
        
        const result = await exportMultipleSheetsWithExcelJS(filePath, filePath, sheets, { 
            password, 
            sourcePassword: sourcePassword || password 
        });
        
        if (!result.success) {
            console.error('[Save] Fehler:', result.error);
            securityLog.log('ERROR', 'EXCEL_SAVE_FAILED', {
                file: path.basename(filePath),
                error: result.error
            });
            return { success: false, error: result.error };
        }
        
        console.log(`[Save] Erfolgreich in ${result.stats?.totalTimeMs || 0}ms`);
        
        securityLog.log('INFO', 'EXCEL_SAVE_COMPLETED', {
            file: path.basename(filePath),
            sheetsModified: result.sheetsExported,
            method: 'exceljs',
            timeMs: result.stats?.totalTimeMs
        });

        // Netzwerk-Log für Datei (falls auf Netzlaufwerk)
        await networkLog.log(filePath, 'EXCEL_SAVE', {
            sheetsModified: result.sheetsExported
        });

        return {
            success: true,
            message: result.message,
            sheetsModified: result.sheetsExported,
            passwordProtected: !!password
        };
    } catch (error) {
        console.error('[Save] Exception:', error);
        securityLog.log('ERROR', 'EXCEL_SAVE_EXCEPTION', {
            file: path.basename(filePath),
            error: error.message
        });
        return { success: false, error: error.message };
    }
});

// ============================================
// KONFIGURATION
// ============================================

// Config speichern (mit Unterstützung für Computer-spezifische Abschnitte)
ipcMain.handle('config:save', async (event, { filePath, config, mergeMode = 'auto' }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        // Schema-Validierung vor dem Speichern
        const validation = configSchema.validate(config);
        if (!validation.valid) {
            securityLog.log('WARN', 'CONFIG_VALIDATION_FAILED_ON_SAVE', {
                errors: validation.errors,
                path: path.basename(filePath)
            });
            // Trotzdem speichern, aber warnen (Rückwärtskompatibilität)
        }

        // Config bereinigen (nur bekannte Felder speichern)
        const sanitizedConfig = configSchema.sanitize(config);
        const computerName = getComputerName();
        
        let finalConfig = sanitizedConfig;
        let savedToSection = null;
        let convertedToNested = false;
        
        // Prüfen ob die Datei existiert
        if (fs.existsSync(filePath) && mergeMode !== 'overwrite') {
            try {
                const existingContent = fs.readFileSync(filePath, 'utf8');
                const existingConfig = JSON.parse(existingContent);
                
                // Prüfe ob es bereits eine verschachtelte Struktur ist (hat 'default' Abschnitt)
                const hasNestedStructure = existingConfig.default !== undefined;
                
                if (hasNestedStructure) {
                    // Bereits verschachtelt: Nur den eigenen Computer-Abschnitt aktualisieren
                    finalConfig = { ...existingConfig };
                    finalConfig[computerName] = sanitizedConfig;
                    savedToSection = computerName;
                    
                    securityLog.log('INFO', 'CONFIG_MERGED_TO_SECTION', {
                        computerName: computerName,
                        path: path.basename(filePath)
                    });
                } else {
                    // Flache Struktur existiert -> In verschachtelte Struktur konvertieren!
                    // Die bestehende flache Config wird zu "default"
                    // Die neue Config wird unter dem Computernamen gespeichert
                    finalConfig = {
                        default: existingConfig,  // Bestehende Config wird default
                        [computerName]: sanitizedConfig  // Neue Config unter Computer-Name
                    };
                    savedToSection = computerName;
                    convertedToNested = true;
                    
                    securityLog.log('INFO', 'CONFIG_CONVERTED_TO_NESTED', {
                        computerName: computerName,
                        path: path.basename(filePath)
                    });
                }
            } catch (parseError) {
                // Datei existiert aber ist kein gültiges JSON -> überschreiben
                securityLog.log('WARN', 'CONFIG_EXISTING_INVALID', {
                    path: path.basename(filePath),
                    error: parseError.message
                });
            }
        }

        fs.writeFileSync(filePath, JSON.stringify(finalConfig, null, 2), 'utf8');
        
        securityLog.log('INFO', 'CONFIG_SAVED', { 
            path: path.basename(filePath),
            computerName: savedToSection || 'flat',
            mode: savedToSection ? 'merged' : 'overwrite',
            convertedToNested: convertedToNested
        });
        
        return { 
            success: true,
            savedToSection: savedToSection,
            computerName: computerName,
            convertedToNested: convertedToNested
        };
    } catch (error) {
        securityLog.log('ERROR', 'CONFIG_SAVE_FAILED', { error: error.message });
        
        // Benutzerfreundliche Fehlermeldungen
        let userMessage = error.message;
        if (error.code === 'EPERM' || error.code === 'EACCES' || error.code === 'EROFS') {
            userMessage = 'Config-Datei ist schreibgeschützt oder Sie haben keine Schreibberechtigung.';
        } else if (error.code === 'ENOENT') {
            userMessage = 'Zielordner existiert nicht.';
        } else if (error.code === 'ENOSPC') {
            userMessage = 'Kein Speicherplatz verfügbar.';
        }
        
        return { success: false, error: userMessage };
    }
});

// Config laden
ipcMain.handle('config:load', async (event, filePath) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }

    try {
        if (!fs.existsSync(filePath)) {
            return { success: false, error: 'Datei nicht gefunden' };
        }
        const content = fs.readFileSync(filePath, 'utf8');
        let config;
        try {
            config = JSON.parse(content);
        } catch (parseError) {
            securityLog.log('ERROR', 'CONFIG_INVALID_JSON', {
                path: path.basename(filePath),
                error: parseError.message
            });
            return { success: false, error: 'Ungültige JSON-Syntax' };
        }

        // Computer-spezifische Config zusammenführen
        const { mergedConfig, computerName, hasComputerSection, isLegacyFormat } = mergeComputerConfig(config);

        // Schema-Validierung (auf der gemergten Config)
        const validation = configSchema.validate(mergedConfig);
        if (!validation.valid) {
            securityLog.log('WARN', 'CONFIG_VALIDATION_FAILED', {
                path: path.basename(filePath),
                errors: validation.errors
            });
            // Config trotzdem laden, aber Warnungen zurückgeben
            return {
                success: true,
                config: configSchema.sanitize(mergedConfig),
                computerName: computerName,
                hasComputerSection: hasComputerSection,
                isLegacyFormat: isLegacyFormat,
                warnings: validation.errors
            };
        }

        securityLog.log('INFO', 'CONFIG_LOADED', { 
            path: path.basename(filePath),
            computerName: computerName,
            hasComputerSection: hasComputerSection
        });
        return { 
            success: true, 
            config: configSchema.sanitize(mergedConfig),
            computerName: computerName,
            hasComputerSection: hasComputerSection,
            isLegacyFormat: isLegacyFormat
        };
    } catch (error) {
        securityLog.log('ERROR', 'CONFIG_LOAD_FAILED', { error: error.message });
        return { success: false, error: error.message };
    }
});

// App-Pfad ermitteln (fuer Config im Programmordner)
ipcMain.handle('app:getPath', async (event) => {
    const exePath = app.getPath('exe');
    const exeDir = path.dirname(exePath);

    return {
        appPath: app.getAppPath(),
        userData: app.getPath('userData'),
        exe: exePath,
        exeDir: exeDir
    };
});

// Security-Logs abrufen
ipcMain.handle('security:getLogs', async (event, { fromFile = true, limit = 500 } = {}) => {
    try {
        let entries;
        if (fromFile) {
            entries = securityLog.readFromFile();
        } else {
            entries = securityLog.getEntries();
        }

        // Neueste zuerst, mit Limit
        const limited = entries.slice(-limit).reverse();

        return {
            success: true,
            entries: limited,
            totalCount: entries.length,
            logFilePath: securityLog.logFilePath
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Security-Log Integrität prüfen
ipcMain.handle('security:verifyLogs', async (event) => {
    try {
        const result = securityLog.verifyIntegrity();
        return { success: true, ...result };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Security-Logs löschen (nur mit Bestätigung)
ipcMain.handle('security:clearLogs', async (event) => {
    try {
        const result = securityLog.clearLogs();
        return result;
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Computername abrufen (für Computer-spezifische Config)
ipcMain.handle('system:getComputerName', async (event) => {
    return {
        computerName: getComputerName(),
        hostname: os.hostname(),
        platform: os.platform(),
        username: os.userInfo().username
    };
});

// Prüfen ob ein Pfad auf einem Netzlaufwerk liegt
ipcMain.handle('network:isNetworkPath', async (event, filePath) => {
    try {
        return {
            success: true,
            isNetwork: networkLog.isNetworkPath(filePath),
            hostname: networkLog.hostname
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Netzwerk-Logs für einen Ordner abrufen
ipcMain.handle('network:getLogs', async (event, filePath) => {
    try {
        const entries = networkLog.readLogs(filePath);
        const logPath = networkLog.getNetworkLogPath(filePath);

        return {
            success: true,
            entries: entries.reverse(), // Neueste zuerst
            totalCount: entries.length,
            logFilePath: logPath,
            isNetworkPath: networkLog.isNetworkPath(filePath)
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Konfliktprüfung: Prüft ob Datei kürzlich von anderem Rechner bearbeitet wurde
ipcMain.handle('network:checkConflict', async (event, filePath, minutesThreshold = 5) => {
    try {
        const result = networkLog.checkConflict(filePath, minutesThreshold);
        return { success: true, ...result };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Session-Lock erstellen (beim Öffnen einer Datei)
ipcMain.handle('network:createSessionLock', async (event, filePath) => {
    try {
        const result = networkLog.createSessionLock(filePath);
        if (result.success) {
            securityLog.log('INFO', 'SESSION_LOCK_CREATED', {
                file: path.basename(filePath),
                isNetwork: networkLog.isNetworkPath(filePath)
            });
        }
        return result;
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Session-Lock entfernen (beim Schließen einer Datei)
ipcMain.handle('network:removeSessionLock', async (event, filePath) => {
    try {
        networkLog.removeSessionLock(filePath);
        securityLog.log('INFO', 'SESSION_LOCK_REMOVED', {
            file: path.basename(filePath)
        });
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Externe URL im Standard-Browser öffnen
ipcMain.handle('shell:openExternal', async (event, url) => {
    const { shell } = require('electron');
    try {
        // Sicherheitsprüfung: Nur http, https und mailto erlauben
        if (url && (url.startsWith('http://') || url.startsWith('https://') || url.startsWith('mailto:'))) {
            securityLog.log('INFO', 'EXTERNAL_URL_OPENED', {
                protocol: url.split(':')[0],
                domain: url.includes('://') ? url.split('://')[1].split('/')[0] : 'mailto'
            });
            await shell.openExternal(url);
            return { success: true };
        } else {
            return { success: false, error: 'Nur http, https und mailto URLs sind erlaubt' };
        }
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Loading-State für config:loadFromAppDir um Race Conditions zu verhindern
let configLoadingState = {
    isLoading: false,
    pendingPromise: null
};

// Automatisch config.json im Programmordner oder Benutzerordnern suchen
ipcMain.handle('config:loadFromAppDir', async (event, workingDir) => {
    // Race Condition verhindern: Wenn bereits geladen wird, auf das Ergebnis warten
    if (configLoadingState.isLoading && configLoadingState.pendingPromise) {
        return configLoadingState.pendingPromise;
    }

    // Ladevorgang starten
    configLoadingState.isLoading = true;

    const loadConfigAsync = async () => {
        try {
            const exePath = app.getPath('exe');
            const exeDir = path.dirname(exePath);
            const documentsDir = app.getPath('documents');
            const downloadsDir = app.getPath('downloads');

            // PORTABLE EXE: Der Ordner wo die portable EXE gestartet wurde
            const portableDir = process.env.PORTABLE_EXECUTABLE_DIR || '';

            // Suchpfade in Prioritätsreihenfolge
            const possiblePaths = [];

            // 1. ARBEITSORDNER (höchste Priorität - vom Benutzer festgelegt)
            if (workingDir && typeof workingDir === 'string') {
                possiblePaths.push(path.join(workingDir, 'config.json'));
            }

            // 2. Portable EXE: Neben der EXE (höchste Priorität für portable Version)
            if (portableDir) {
                possiblePaths.push(path.join(portableDir, 'config.json'));
            }

            // 3. Installationsordner (neben der EXE)
            possiblePaths.push(path.join(exeDir, 'config.json'));

            // 4. Dokumente-Ordner des Benutzers
            possiblePaths.push(path.join(documentsDir, 'config.json'));
            possiblePaths.push(path.join(documentsDir, 'Excel-Data-Sync-Pro', 'config.json'));

            // 5. Downloads-Ordner des Benutzers
            possiblePaths.push(path.join(downloadsDir, 'config.json'));

            // 6. Im Entwicklungsmodus: Projektordner
            if (process.argv.includes('--dev') || !app.isPackaged) {
                possiblePaths.push(path.join(__dirname, 'config.json'));
                possiblePaths.push(path.join(process.cwd(), 'config.json'));
            }

            // Schnelle Suche - bei erstem Treffer abbrechen
            for (const configPath of possiblePaths) {
                if (fs.existsSync(configPath)) {
                    const content = fs.readFileSync(configPath, 'utf8');
                    let config;
                    try {
                        config = JSON.parse(content);
                    } catch (parseError) {
                        securityLog.log('WARN', 'CONFIG_PARSE_ERROR', {
                            path: path.basename(configPath),
                            error: parseError.message
                        });
                        continue; // Nächsten Pfad probieren
                    }

                    // Computer-spezifische Config zusammenführen
                    const { mergedConfig, computerName, hasComputerSection, isLegacyFormat } = mergeComputerConfig(config);

                    // Schema-Validierung (auf der gemergten Config)
                    const validation = configSchema.validate(mergedConfig);
                    if (!validation.valid) {
                        securityLog.log('WARN', 'CONFIG_VALIDATION_FAILED', {
                            path: path.basename(configPath),
                            errors: validation.errors
                        });
                    }

                    const source = configPath.includes(workingDir || '') ? 'workingDir' :
                                   configPath.includes(portableDir) ? 'portable' :
                                   configPath.includes(exeDir) ? 'exeDir' : 'userDir';

                    securityLog.log('INFO', 'CONFIG_AUTO_LOADED', {
                        path: path.basename(configPath),
                        source: source,
                        computerName: computerName,
                        hasComputerSection: hasComputerSection,
                        isLegacyFormat: isLegacyFormat
                    });

                    // Excel-Engine aus Config setzen (falls vorhanden)
                    if (mergedConfig.excelEngine) {
                        pythonBridge.setExcelEngine(mergedConfig.excelEngine);
                        console.log(`[Config] Excel-Engine aus config.json: ${mergedConfig.excelEngine}`);
                    }

                    return {
                        success: true,
                        config: configSchema.sanitize(mergedConfig),
                        path: configPath,
                        computerName: computerName,
                        hasComputerSection: hasComputerSection,
                        isLegacyFormat: isLegacyFormat,
                        warnings: validation.valid ? undefined : validation.errors
                    };
                }
            }

            // Keine config.json gefunden
            return {
                success: false,
                error: 'Keine config.json gefunden',
                searchedPaths: possiblePaths
            };
        } catch (error) {
            securityLog.log('ERROR', 'CONFIG_AUTO_LOAD_FAILED', { error: error.message });
            return { success: false, error: error.message };
        } finally {
            // Loading-State zurücksetzen
            configLoadingState.isLoading = false;
            configLoadingState.pendingPromise = null;
        }
    };

    // Promise speichern für parallele Aufrufe
    configLoadingState.pendingPromise = loadConfigAsync();
    return configLoadingState.pendingPromise;
});

// ============================================
// TEMPLATE AUS QUELLDATEI ERSTELLEN
// ============================================

/**
 * Hilfsfunktion: Konvertiert Spaltennummer zu Spaltenbuchstabe (1=A, 2=B, 27=AA, etc.)
 */
function numberToColumnLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

/**
 * Hilfsfunktion: Konvertiert Spaltenbuchstabe zu Nummer (A=1, B=2, AA=27, etc.)
 */
function columnLetterToNumber(col) {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
        result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
}

/**
 * Hilfsfunktion: Verschiebt alle Spaltenreferenzen in einer Zellreferenz um n Spalten
 * z.B. "A1" + 2 -> "C1", "H1:H100" + 2 -> "J1:J100"
 */
function shiftColumnReference(ref, shiftBy) {
    return ref.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
        const colNum = columnLetterToNumber(col);
        const newCol = numberToColumnLetter(colNum + shiftBy);
        return newCol + row;
    });
}

/**
 * Hilfsfunktion: Dekodiert XML-Entities (z.B. &amp; -> &)
 */
function decodeXmlEntities(str) {
    return str
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
}

/**
 * Hilfsfunktion: Enkodiert XML-Entities (z.B. & -> &amp;)
 */
function encodeXmlEntities(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Erstellt ein leeres Template aus einer Quelldatei.
 * - Behält nur ausgewählte Sheets
 * - Fügt optional Flag- und Kommentar-Spalten am Anfang ein
 * - Behält die erste Zeile (Header) pro Sheet
 * - Löscht alle Datenzeilen
 * - Erweitert Conditional Formatting Ranges auf ganze Spalten
 * - Erhält alle Formatierungen und Stile
 */
ipcMain.handle('excel:createTemplateFromSource', async (event, { sourcePath, outputPath, selectedSheets, addFlagColumn, addCommentColumn }) => {
    // Sicherheitsprüfungen
    if (!isValidFilePath(sourcePath)) {
        return { success: false, error: 'Ungültiger Quellpfad' };
    }
    if (!isValidFilePath(outputPath)) {
        return { success: false, error: 'Ungültiger Ausgabepfad' };
    }

    // Berechne wie viele Spalten eingefügt werden
    const extraColumnsCount = (addFlagColumn ? 1 : 0) + (addCommentColumn ? 1 : 0);

    try {
        const JSZip = require('jszip');

        // 1. Quelldatei als ZIP lesen
        const sourceBuffer = fs.readFileSync(sourcePath);
        const zip = await JSZip.loadAsync(sourceBuffer);

        // 2. workbook.xml lesen um Sheet-Zuordnungen zu bekommen
        const workbookXml = await zip.file('xl/workbook.xml').async('string');

        // Sheet-Namen und rId extrahieren (Namen sind XML-encoded in der Datei)
        const sheetMatches = [...workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*>/g)];
        const sheetRels = {};        // XML-encoded Namen -> rId
        const sheetRelsDecoded = {}; // Dekodierte Namen -> rId
        const encodedToDecoded = {}; // Mapping encoded -> decoded
        sheetMatches.forEach(match => {
            const encodedName = match[1];
            const decodedName = decodeXmlEntities(encodedName);
            sheetRels[encodedName] = match[2];       // encoded name -> rId
            sheetRelsDecoded[decodedName] = match[2]; // decoded name -> rId
            encodedToDecoded[encodedName] = decodedName;
        });

        // 3. workbook.xml.rels lesen um rId -> sheetX.xml Zuordnung zu bekommen
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
        const rIdToFile = {};
        const relMatches = [...relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)];
        relMatches.forEach(match => {
            rIdToFile[match[1]] = match[2].replace(/^\//, ''); // rId -> worksheets/sheetX.xml
        });

        // 4. Mapping: SheetName (dekodiert) -> Dateiname erstellen
        const sheetToFile = {};
        const sheetToFileEncoded = {}; // Für das Entfernen aus workbook.xml
        for (const [encodedName, rId] of Object.entries(sheetRels)) {
            const target = rIdToFile[rId];
            const decodedName = encodedToDecoded[encodedName];
            if (target) {
                const filePath = 'xl/' + target.replace(/^xl\//, '');
                sheetToFile[decodedName] = filePath;  // Dekodierter Name -> Datei
                sheetToFileEncoded[encodedName] = filePath; // Encoded Name -> Datei (für XML-Operationen)
            }
        }

        // 5. Sheets identifizieren, die NICHT ausgewählt wurden (vergleiche mit dekodierten Namen)
        const allDecodedSheetNames = Object.keys(sheetToFile);
        const sheetsToRemove = allDecodedSheetNames.filter(name => !selectedSheets.includes(name));

        // 6. Nicht ausgewählte Sheets aus workbook.xml entfernen (verwende encoded Namen für XML)
        let modifiedWorkbookXml = workbookXml;
        for (const decodedName of sheetsToRemove) {
            // XML-encoded Name für Regex verwenden
            const encodedName = encodeXmlEntities(decodedName);
            const sheetRegex = new RegExp(`<sheet[^>]*name="${encodedName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`, 'g');
            modifiedWorkbookXml = modifiedWorkbookXml.replace(sheetRegex, '');
        }
        zip.file('xl/workbook.xml', modifiedWorkbookXml);

        // 7. Sheet-Dateien der nicht ausgewählten Sheets entfernen
        for (const decodedName of sheetsToRemove) {
            const sheetFile = sheetToFile[decodedName];
            if (sheetFile && zip.files[sheetFile]) {
                zip.remove(sheetFile);
            }
        }

        // 8. Ausgewählte Worksheets verarbeiten
        let totalCfRules = 0;
        let processedSheets = 0;

        for (const sheetName of selectedSheets) {
            const sheetFile = sheetToFile[sheetName];
            if (!sheetFile || !zip.files[sheetFile]) {
                console.warn(`Sheet-Datei nicht gefunden: ${sheetName} -> ${sheetFile}`);
                continue;
            }

            let sheetXml = await zip.file(sheetFile).async('string');

            // Anzahl der CF-Regeln zählen
            const cfMatches = sheetXml.match(/<conditionalFormatting[^>]*>/g) || [];
            totalCfRules += cfMatches.length;

            // Wenn Extra-Spalten hinzugefügt werden sollen
            if (extraColumnsCount > 0) {
                // A) Alle Zellreferenzen in <c r="..."> verschieben
                sheetXml = sheetXml.replace(
                    /<c\s+r="([A-Z]+)(\d+)"/g,
                    (match, col, row) => {
                        const newCol = numberToColumnLetter(columnLetterToNumber(col) + extraColumnsCount);
                        return `<c r="${newCol}${row}"`;
                    }
                );

                // B) Merge-Cells verschieben (falls vorhanden)
                sheetXml = sheetXml.replace(
                    /<mergeCell\s+ref="([A-Z]+\d+):([A-Z]+\d+)"/g,
                    (match, start, end) => {
                        const newStart = shiftColumnReference(start, extraColumnsCount);
                        const newEnd = shiftColumnReference(end, extraColumnsCount);
                        return `<mergeCell ref="${newStart}:${newEnd}"`;
                    }
                );

                // C) Dimension verschieben (falls vorhanden)
                sheetXml = sheetXml.replace(
                    /<dimension\s+ref="([A-Z]+\d+):([A-Z]+\d+)"/g,
                    (match, start, end) => {
                        // Start bei A1 lassen wenn Extra-Spalten, Ende verschieben
                        const newEnd = shiftColumnReference(end, extraColumnsCount);
                        return `<dimension ref="A1:${newEnd}"`;
                    }
                );

                // D) Neue Header-Zellen am Anfang von Zeile 1 einfügen
                const newCells = [];
                if (addFlagColumn) {
                    newCells.push('<c r="A1" t="inlineStr"><is><t>Flag</t></is></c>');
                }
                if (addCommentColumn) {
                    const col = addFlagColumn ? 'B' : 'A';
                    newCells.push(`<c r="${col}1" t="inlineStr"><is><t>Kommentar</t></is></c>`);
                }

                // Füge neue Zellen in Zeile 1 ein
                if (newCells.length > 0) {
                    sheetXml = sheetXml.replace(
                        /(<row[^>]*r="1"[^>]*>)/,
                        `$1${newCells.join('')}`
                    );
                }
            }

            // E) Datenzeilen löschen (behalte nur Zeile 1 = Header)
            sheetXml = sheetXml.replace(
                /(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/,
                (match, open, content, close) => {
                    const headerRowMatch = content.match(/<row[^>]*r="1"[^>]*>[\s\S]*?<\/row>/);
                    const headerRow = headerRowMatch ? headerRowMatch[0] : '';
                    return open + headerRow + close;
                }
            );

            // F) Conditional Formatting Ranges auf ganze Spalten erweitern UND verschieben
            sheetXml = sheetXml.replace(
                /sqref="([^"]+)"/g,
                (match, sqref) => {
                    const ranges = sqref.split(/\s+/);
                    const columns = new Set();

                    for (const range of ranges) {
                        const colMatch = range.match(/^([A-Z]+)/);
                        if (colMatch) {
                            // Spalte um extraColumnsCount verschieben
                            const originalCol = colMatch[1];
                            const shiftedCol = extraColumnsCount > 0
                                ? numberToColumnLetter(columnLetterToNumber(originalCol) + extraColumnsCount)
                                : originalCol;
                            columns.add(shiftedCol);
                        }
                    }

                    if (columns.size > 0) {
                        const newSqref = Array.from(columns).map(col => `${col}:${col}`).join(' ');
                        return `sqref="${newSqref}"`;
                    }
                    return match;
                }
            );

            zip.file(sheetFile, sheetXml);
            processedSheets++;
        }

        // 9. Template speichern
        const outputBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });

        fs.writeFileSync(outputPath, outputBuffer);

        securityLog.log('INFO', 'TEMPLATE_CREATED', {
            sourceFile: path.basename(sourcePath),
            outputFile: path.basename(outputPath),
            sheetsProcessed: processedSheets,
            addFlagColumn,
            addCommentColumn
        });

        return {
            success: true,
            message: 'Template erfolgreich erstellt',
            fileName: path.basename(outputPath),
            stats: {
                sheetsProcessed: processedSheets,
                cfRulesPreserved: totalCfRules,
                extraColumnsAdded: extraColumnsCount
            }
        };

    } catch (error) {
        console.error('Fehler beim Erstellen des Templates:', error);
        securityLog.log('ERROR', 'TEMPLATE_CREATION_FAILED', {
            sourceFile: path.basename(sourcePath),
            outputFile: path.basename(outputPath),
            error: error.message
        });
        return { success: false, error: error.message };
    }
});

// =============================================================================
// LIVE SESSION - Excel bleibt offen für sofortige Operationen
// =============================================================================

const { getLiveSession } = require('./python/excel_live_bridge');

// Live-Session starten
ipcMain.handle('liveSession:start', async () => {
    console.log('[LiveSession] IPC: start');
    try {
        const session = getLiveSession();
        return await session.start();
    } catch (error) {
        console.error('[LiveSession] start error:', error);
        return { success: false, error: error.message };
    }
});

// Datei öffnen
ipcMain.handle('liveSession:openFile', async (event, filePath, sheetName) => {
    console.log('[LiveSession] IPC: openFile', filePath, sheetName);
    try {
        const session = getLiveSession();
        return await session.openFile(filePath, sheetName);
    } catch (error) {
        console.error('[LiveSession] openFile error:', error);
        return { success: false, error: error.message };
    }
});

// Datei speichern
ipcMain.handle('liveSession:saveFile', async (event, outputPath) => {
    console.log('[LiveSession] IPC: saveFile', outputPath);
    try {
        const session = getLiveSession();
        return await session.saveFile(outputPath);
    } catch (error) {
        console.error('[LiveSession] saveFile error:', error);
        return { success: false, error: error.message };
    }
});

// Session schließen
ipcMain.handle('liveSession:close', async () => {
    console.log('[LiveSession] IPC: close');
    try {
        const session = getLiveSession();
        return await session.close();
    } catch (error) {
        console.error('[LiveSession] close error:', error);
        return { success: false, error: error.message };
    }
});

// Daten lesen
ipcMain.handle('liveSession:getData', async () => {
    try {
        const session = getLiveSession();
        return await session.getData();
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// === ZEILEN-OPERATIONEN ===

ipcMain.handle('liveSession:deleteRow', async (event, rowIndex) => {
    console.log('[LiveSession] IPC: deleteRow', rowIndex);
    try {
        const session = getLiveSession();
        return await session.deleteRow(rowIndex);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:insertRow', async (event, rowIndex, count = 1) => {
    console.log('[LiveSession] IPC: insertRow', rowIndex, count);
    try {
        const session = getLiveSession();
        return await session.insertRow(rowIndex, count);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:moveRow', async (event, fromIndex, toIndex) => {
    console.log('[LiveSession] IPC: moveRow', fromIndex, '->', toIndex);
    try {
        const session = getLiveSession();
        return await session.moveRow(fromIndex, toIndex);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:hideRow', async (event, rowIndex, hidden = true) => {
    console.log('[LiveSession] IPC: hideRow', rowIndex, hidden);
    try {
        const session = getLiveSession();
        return await session.hideRow(rowIndex, hidden);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:highlightRow', async (event, rowIndex, color) => {
    console.log('[LiveSession] IPC: highlightRow', rowIndex, color);
    try {
        const session = getLiveSession();
        return await session.highlightRow(rowIndex, color);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// === SPALTEN-OPERATIONEN ===

ipcMain.handle('liveSession:deleteColumn', async (event, colIndex) => {
    console.log('[LiveSession] IPC: deleteColumn', colIndex);
    try {
        const session = getLiveSession();
        return await session.deleteColumn(colIndex);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:insertColumn', async (event, colIndex, count = 1, headers = null) => {
    console.log('[LiveSession] IPC: insertColumn', colIndex, count);
    try {
        const session = getLiveSession();
        return await session.insertColumn(colIndex, count, headers);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:moveColumn', async (event, fromIndex, toIndex) => {
    console.log('[LiveSession] IPC: moveColumn', fromIndex, '->', toIndex);
    try {
        const session = getLiveSession();
        return await session.moveColumn(fromIndex, toIndex);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

ipcMain.handle('liveSession:hideColumn', async (event, colIndex, hidden = true) => {
    console.log('[LiveSession] IPC: hideColumn', colIndex, hidden);
    try {
        const session = getLiveSession();
        return await session.hideColumn(colIndex, hidden);
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// === ZELL-OPERATIONEN ===

ipcMain.handle('liveSession:setCellValue', async (event, rowIndex, colIndex, value) => {
    try {
        const session = getLiveSession();
        return await session.setCellValue(rowIndex, colIndex, value);
    } catch (error) {
        return { success: false, error: error.message };
    }
});
