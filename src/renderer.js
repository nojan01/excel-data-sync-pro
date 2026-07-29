        // ============================================
        // ELECTRON-ONLY MODE
        // ============================================
        console.log('🚀 Electron-Modus aktiv');
        
        // ============================================
        // INTERNATIONALIZATION (i18n)
        // ============================================
        const translations = {
            de: {
                // Header
                appTitle: 'Excel Data Sync Pro',
                loadConfig: '📂 config.json laden',
                saveConfig: '💾 config.json speichern',
                help: '❓ Hilfe',
                
                // Sidebar
                configuration: 'Konfiguration',
                settings: '⚙️ Einstellungen',
                language: 'Sprache',
                theme: 'Design',
                themeDark: '🌙 Dunkel',
                themeLight: '☀️ Hell',
                
                // Working Directory
                workingDirectory: '📁 Arbeitsordner',
                selectWorkingDir: '📂 Ordner auswählen',
                noWorkingDirSet: 'Kein Ordner gewählt',
                clearWorkingDir: '✖️ Arbeitsordner löschen',
                workingDirSet: '✓ ',
                workingDirCleared: 'Arbeitsordner gelöscht',
                
                // Files
                file1Source: '📄 Datei 1 (Quelle)',
                file2Target: '📄 Datei 2 (Ziel)',
                loadSourceFile: '📂 Quelldatei laden',
                loadTargetFile: '📂 Zieldatei laden',
                noFileLoaded: 'Keine Datei geladen',
                worksheet: 'Arbeitsblatt:',
                loadFileFirst: '-- Erst Datei laden --',
                
                // Mapping
                columnMapping: '🔗 Spalten-Zuordnung',
                configureColumns: '⚙️ Spalten konfigurieren',
                loadBothFiles: 'Laden Sie beide Dateien',
                columnsConfigured: 'Spalte(n) konfiguriert',
                
                // Template
                monthTemplate: '📄 Monats-Template',
                loadTemplate: '📂 Template laden',
                createTemplateFromSource: '🔧 Template aus Quelldatei',
                noTemplateLoaded: 'Kein Template geladen',
                templateHint: 'Leere Excel-Vorlage für "📅 Neue Monatsdatei"',
                createTemplateTitle: '🔧 Template aus Quelldatei erstellen',
                createTemplateDesc: 'Erstellt ein leeres Template mit allen Formatierungen und bedingten Formatierungen aus der Quelldatei.',
                sourceFileLabel: 'Quelldatei:',
                selectSheets: 'Arbeitsblätter auswählen:',
                loadSourceFirst: 'Laden Sie zuerst eine Quelldatei',
                selectAll: '✓ Alle auswählen',
                deselectAll: '✗ Alle abwählen',
                templateInfoText: '💡 Die Header-Zeile wird behalten, alle Datenzeilen werden gelöscht. CF-Regeln werden auf ganze Spalten erweitert.',
                createAndSave: '🔧 Template erstellen & speichern',
                templateCreated: 'Template erfolgreich erstellt',
                sheetsProcessed: 'Arbeitsblätter verarbeitet',
                cfRulesPreserved: 'CF-Regeln erhalten',
                extraColumnsInTemplate: 'Extra-Spalten im Template erstellen:',
                flagColumn: 'Flag-Spalte (A)',
                commentColumn: 'Kommentar-Spalte (B)',
                extraColumnsHint: 'Aktivieren, wenn in "Spalten konfigurieren" die entsprechenden Optionen genutzt werden.',
                
                // History
                lastTransfers: '📋 Letzte Übertragungen',
                noTransfersYet: 'Noch keine Übertragungen',
                
                // Search
                searchPlaceholder: 'Suche... (Platzhalter: * = beliebig, ? = ein Zeichen)',
                search: '🔍 Suchen',
                newRow: '➕ Neue Zeile',
                
                // Results
                readyToStart: 'Bereit zum Starten',
                instructions: '1. Laden Sie die Quelldatei (Datei 1)<br>2. Laden Sie die Zieldatei (Datei 2)<br>3. Konfigurieren Sie die Spalten-Zuordnung<br>4. Suchen Sie nach Zeilen und übertragen Sie diese',
                noResults: 'Keine Treffer für',
                results: 'Treffer für',
                
                // Transfer
                prepareTransfer: '📤 Zeile(n) zur Übertragung vorbereiten',
                selected: 'ausgewählt',
                flag: 'Spalte 1 - Flag:',
                comment: 'Spalte 2 - Kommentar:',
                commentPlaceholder: 'Freier Text...',
                addToQueue: '➕ Markierte zur Warteschlange',
                transferDirect: '➡️ Markierte direkt übertragen',
                selectAll: '✓ Alle auswählen',
                deselectAll: '✗ Alles abwählen',
                
                // Queue
                queue: '📋 Warteschlange',
                rows: 'Zeilen',
                clear: '🗑️ Leeren',
                preview: '👁️ Vorschau',
                exportToTarget: '📤 Export zur Zieldatei',
                directTransfer: '✅ Direkt übertragen',
                dataExplorer: 'Datenexplorer',
                newMonthFile: '📅 Neue Monatsdatei',
                noRowsInQueue: 'Keine Zeilen in der Warteschlange',
                
                // New Row
                createNewRow: '✏️ Neue Zeile erstellen',
                close: '✕ Schließen',
                toQueue: '➕ Zur Warteschlange',
                transferDirectly: '➡️ Direkt übertragen',
                keepFields: 'Felder beibehalten',
                keepFieldsTooltip: 'Wenn aktiviert, bleiben die Feldwerte nach dem Übertragen erhalten – praktisch wenn sich nur ein Feld ändert',
                
                // Messages
                noTargetFile: 'Keine Zieldatei geladen',
                selectAtLeastOne: 'Bitte wählen Sie mindestens eine Zeile aus',
                rowsTransferred: 'Zeile(n) direkt übertragen!',
                rowsAdded: 'Zeile(n) zur Warteschlange hinzugefügt',
                edited: 'bearbeitet',
                duplicates: 'Duplikat(e)',
                configFirst: 'Bitte zuerst Spalten konfigurieren',
                
                // Data Explorer
                explorerTitle: 'Datenexplorer',
                noFileLoadedExplorer: 'Keine Datei geladen',
                openFile: '📂 Datei öffnen',
                fullTextSearch: 'Volltextsuche:',
                searchPlaceholderExplorer: 'Suchbegriff eingeben...',
                columns: '❙❙ Spalten',
                exportXlsx: '💾 Speichern',
                filter: '🔍 Filter',
                addFilter: '➕ Filter hinzufügen',
                resetFilters: '🗑️ Filter zurücksetzen',
                showAll: 'Alle anzeigen',
                hideAll: 'Alle ausblenden',
                noDataLoaded: 'No data loaded.',
                pleaseLoadFile: '📂 Please load an Excel file',
                contains: 'enthält',
                equals: 'gleich',
                startsWith: 'beginnt mit',
                endsWith: 'endet mit',
                selectColumn: 'Spalte wählen',
                filterValue: 'Filterwert',
                findReplace: '🔄 Ersetzen',
                findReplaceTitle: 'Suchen & Ersetzen (Strg+H)',
                addColumns: '🔗 Spalten hinzufügen',
                addColumnsTitle: 'Spalten aus einer anderen Excel-Datei hinzufügen (basierend auf gemeinsamer Seriennummer)',
                
                // Data Join Modal
                joinModalTitle: '🔗 Spalten aus Datei hinzufügen',
                joinModalInfo: '<strong style="color: var(--text);">📋 Funktion:</strong> Fügt Spalten aus einer anderen Excel-Datei hinzu, basierend auf einem gemeinsamen Schlüssel (z.B. Seriennummer). Die Zeilen werden automatisch anhand des Schlüssels zugeordnet.',
                joinStep1: '1️⃣ Datenquelle auswählen',
                joinSelectFile: '📂 Datei auswählen',
                joinNoFileSelected: 'Keine Datei ausgewählt',
                joinDropZoneText: 'Excel-Datei hierher ziehen',
                joinWorksheet: 'Arbeitsblatt:',
                joinLoadFile: '-- Datei laden --',
                joinStep2: '2️⃣ Schlüsselspalte (z.B. Seriennummer)',
                joinStep2Desc: 'Wählen Sie die Spalte, die in beiden Dateien als gemeinsamer Schlüssel dient.',
                joinTargetFile: 'Aktuelle Datei (Ziel):',
                joinSourceFile: 'Quelldatei:',
                joinSelectColumn: '-- Spalte wählen --',
                joinStep3: '3️⃣ Spalten zum Hinzufügen auswählen',
                joinStep3Desc: 'Wählen Sie die Spalten aus der Quelldatei, die zur aktuellen Datei hinzugefügt werden sollen.',
                joinLoadSourceFirst: 'Laden Sie zuerst eine Quelldatei',
                joinStep4: '4️⃣ Optionen',
                joinInsertAfterKeyLabel: 'Neue Spalten direkt nach der Schlüsselspalte einfügen',
                joinMarkNotFoundLabel: 'Nicht gefundene Zeilen markieren (leere Zellen = kein Match)',
                joinPreviewTitle: '📊 Vorschau',
                joinStatTargetRowsLabel: 'Zeilen in Ziel:',
                joinStatSourceRowsLabel: 'Zeilen in Quelle:',
                joinStatMatchesLabel: 'Matches gefunden:',
                joinStatNoMatchLabel: 'Ohne Match:',
                joinPreviewBtn: '👁️ Vorschau berechnen',
                joinExecuteBtn: '✓ Spalten hinzufügen',
                
                // Row/Column Actions
                insertRowAbove: 'Zeile darüber einfügen',
                insertRowBelow: 'Zeile darunter einfügen',
                deleteRow: 'Zeile löschen',
                hideRow: 'Zeile ausblenden',
                showRow: 'Zeile einblenden',
                showAllRows: 'Alle Zeilen einblenden',
                hiddenRows: 'versteckte Zeilen',
                hiddenColumns: 'versteckte Spalten',
                rowHidden: 'Zeile ausgeblendet',
                rowShown: 'Zeile eingeblendet',
                allRowsShown: 'Alle Zeilen eingeblendet',
                columnHidden: 'Spalte ausgeblendet',
                columnShown: 'Spalte eingeblendet',
                allColumnsShown: 'Alle Spalten eingeblendet',
                showColumn: 'Spalte einblenden',
                showAllColumns: 'Alle Spalten einblenden',
                insertColumnBefore: 'Spalte links einfügen',
                insertColumnAfter: 'Spalte rechts einfügen',
                deleteColumn: 'Spalte löschen',
                hideColumn: 'Spalte ausblenden',
                newColumn: 'Neue Spalte',
                enterColumnName: 'Spaltenname eingeben:',
                rowInserted: 'Zeile eingefügt',
                rowDeleted: 'Zeile gelöscht',
                columnInserted: 'Spalte eingefügt',
                columnDeleted: 'Spalte gelöscht',
                deleteRowTitle: 'Zeile löschen?',
                deleteColumnTitle: 'Spalte löschen?',
                deleteRowConfirm: 'Möchten Sie diese Zeile wirklich löschen?',
                deleteColumnConfirm: 'Möchten Sie diese Spalte wirklich löschen?',
                deleteColumnWarning: '⚠️ Alle Daten in dieser Spalte gehen verloren!',
                
                // New Month Modal
                createNewMonthFile: '📅 Neue Monatsdatei erstellen',
                newMonthDescription: 'Erschafft eine Kopie der Template-Datei unter neuem Namen und setzt sie als Datei 2.',
                newFilename: 'Neuer Dateiname:',
                filenamePlaceholder: 'z.B. Vertragsliste_2025-01.xlsx',
                templateLabel: '💡 Template:',
                cancel: 'Abbrechen',
                createAndLoad: '📅 Erstellen & Laden',
                
                // Footer
                copyright: '© Norbert Jander 2026 · v{version}',
                
                // Extra Columns
                extraColumns: 'Extra-Spalten',
                enableFlag: 'Flag-Spalte (A/D/C)',
                enableComment: 'Kommentar-Spalte',
                
                // License
                license: 'Lizenz',
                licenseTitle: 'Lizenzinformationen',
                thirdPartyLicenses: 'Drittanbieter-Lizenzen',
                thirdPartyDesc: 'Excel Data Sync Pro verwendet folgende Open-Source-Bibliotheken:',
                allLicensesMIT: 'Alle verwendeten Bibliotheken sind unter der MIT-Lizenz oder kompatiblen Open-Source-Lizenzen lizenziert.',
                packagesTotal: 'Pakete gesamt',
                
                // Security Logs
                logsButton: 'Protokolle',
                logsTitle: 'Protokolle',
                localLogs: 'Lokale Security-Logs',
                securityLogs: 'Security-Logs',
                securityLogsTitle: 'Security-Logs (Manipulationssicher)',
                integrityStatus: 'Integritätsstatus:',
                refresh: 'Aktualisieren',
                verify: 'Verifizieren',
                clearLogs: 'Logs löschen',
                
                // Network Logs
                networkLogs: 'Netzwerk-Logs',
                networkLogsTitle: 'Netzwerk-Protokoll',
                currentComputer: 'Aktueller Rechner:',
                networkLogPath: 'Log-Pfad:',
                entries: 'Einträge:',
                allComputers: 'Alle Rechner',
                timestamp: 'Zeitstempel',
                computer: 'Rechner',
                action: 'Aktion',
                file: 'Datei',
                details: 'Details',
                noNetworkLogs: 'Keine Netzwerk-Logs vorhanden',
                networkLogsNote: 'Das Netzwerk-Protokoll wird nur für Dateien auf Netzlaufwerken geführt. Es zeigt welcher Rechner wann welche Aktion durchgeführt hat (DSGVO-konform, keine persönlichen Daten).',
                
                // Live Mode
                liveMode: 'Live-Modus',
                liveModeWindows: 'Live-Modus (Windows)',
                liveOpen: '🔴 Live öffnen',
                liveSessionActive: '🔴 Live-Session aktiv',
                liveSessionEnded: 'Live-Session beendet',
                liveOffline: 'Offline',
                liveReady: 'Bereit',
                liveConnecting: 'Verbinde...',
                endLiveSession: '❌ Live-Session beenden',
                toggleExcel: '👁️ Excel',
                sendToExcel: '▶ An Excel',
                resetFilter: 'Filter zurücksetzen',
                moveRowUp: 'Zeile nach oben',
                moveRowDown: 'Zeile nach unten',
                moveColumnLeft: 'Spalte nach links',
                moveColumnRight: 'Spalte nach rechts',
                highlightRow: 'Zeile markieren',
                clearHighlight: 'Markierung entfernen',
                highlightGreen: '🟢 Grün',
                highlightYellow: '🟡 Gelb',
                highlightOrange: '🟠 Orange',
                highlightRed: '🔴 Rot',
                highlightBlue: '🔵 Blau',
                highlightPurple: '🟣 Lila',
                
                // Data Explorer additional
                saveChanges: '💾 Speichern',
                previewChanges: '👁️ Vorschau',
                worksheetLabel: 'Arbeitsblatt:',
                loadFileFirst2: '-- Datei laden --',
                filterSection: '🔍 Filter',
                searchFor: 'Suchen nach:',
                replaceWith: 'Ersetzen durch:',
                textToSearch: 'Text zum Suchen...',
                replacementText: 'Ersatztext...',
                caseSensitive: 'Groß-/Kleinschreibung',
                wholeWord: 'Ganzes Wort',
                findNext: '▼ Nächster',
                replaceOne: 'Ersetzen',
                replaceAll: 'Alle ersetzen',
                undo: 'Rückgängig',
                manageSheets: 'Arbeitsblätter verwalten',
                
                // Find & Replace Panel
                searchAndReplace: 'Suchen & Ersetzen',
                
                // Pagination & Status
                showingRows: 'Zeige',
                ofRows: 'von',
                rowsLabel: 'Zeilen',
                totalLabel: 'gesamt',
                editedLabel: 'bearbeitet',
                pageLabel: 'Seite',
                ofLabel: 'von',
                perPage: '/ Seite',
                rowsLoaded: 'Zeilen geladen',
                paginationActive: 'Pagination aktiv',
                loadingFile: 'Lade Datei...',
                loadingData: 'Lade Daten...',
                noFileLoaded: 'Keine Datei geladen',
                errorLabel: 'Fehler',
                cellsEdited: 'Zelle(n) bearbeitet',
                loadedFromCache: 'Aus Cache geladen',
                changes: 'Änderungen',
                editingsRestored: 'Bearbeitungen wiederhergestellt',
                
                // Column panel
                showAllColumns: 'Alle anzeigen',
                hideAllColumns: 'Alle ausblenden',
                
                // Drop Zone
                dropZoneText: 'Excel-Datei hier ablegen',
                dropZoneHint: 'oder klicken zum Öffnen',
                closeButton: 'Schließen'
            },
            en: {
                // Header
                appTitle: 'Excel Data Sync Pro',
                loadConfig: '📂 Load config.json',
                saveConfig: '💾 Save config.json',
                help: '❓ Help',
                
                // Sidebar
                configuration: 'Configuration',
                settings: '⚙️ Settings',
                language: 'Language',
                theme: 'Theme',
                themeDark: '🌙 Dark',
                themeLight: '☀️ Light',
                
                // Working Directory
                workingDirectory: '📁 Working Directory',
                selectWorkingDir: '📂 Select Folder',
                noWorkingDirSet: 'No folder selected',
                clearWorkingDir: '✖️ Clear Working Directory',
                workingDirSet: '✓ ',
                workingDirCleared: 'Working directory cleared',
                
                // Files
                file1Source: '📄 File 1 (Source)',
                file2Target: '📄 File 2 (Target)',
                loadSourceFile: '📂 Load Source File',
                loadTargetFile: '📂 Load Target File',
                noFileLoaded: 'No file loaded',
                worksheet: 'Worksheet:',
                loadFileFirst: '-- Load file first --',
                
                // Mapping
                columnMapping: '🔗 Column Mapping',
                configureColumns: '⚙️ Configure Columns',
                loadBothFiles: 'Load both files',
                columnsConfigured: 'column(s) configured',
                
                // Template
                monthTemplate: '📄 Month Template',
                loadTemplate: '📂 Load Template',
                createTemplateFromSource: '🔧 Template from Source',
                noTemplateLoaded: 'No template loaded',
                templateHint: 'Empty Excel template for "📅 New Month File"',
                createTemplateTitle: '🔧 Create Template from Source',
                createTemplateDesc: 'Creates an empty template with all formatting and conditional formatting from the source file.',
                sourceFileLabel: 'Source file:',
                selectSheets: 'Select worksheets:',
                loadSourceFirst: 'Load a source file first',
                selectAll: '✓ Select all',
                deselectAll: '✗ Deselect all',
                templateInfoText: '💡 The header row is kept, all data rows are deleted. CF rules are extended to entire columns.',
                createAndSave: '🔧 Create & save template',
                templateCreated: 'Template created successfully',
                sheetsProcessed: 'Worksheets processed',
                cfRulesPreserved: 'CF rules preserved',
                extraColumnsInTemplate: 'Create extra columns in template:',
                flagColumn: 'Flag column (A)',
                commentColumn: 'Comment column (B)',
                extraColumnsHint: 'Enable if these options are used in "Configure columns".',
                
                // History
                lastTransfers: '📋 Recent Transfers',
                noTransfersYet: 'No transfers yet',
                
                // Search
                searchPlaceholder: 'Search... (Wildcards: * = any, ? = one character)',
                search: '🔍 Search',
                newRow: '➕ New Row',
                
                // Results
                readyToStart: 'Ready to Start',
                instructions: '1. Load the source file (File 1)<br>2. Load the target file (File 2)<br>3. Configure the column mapping<br>4. Search for rows and transfer them',
                noResults: 'No results for',
                results: 'results for',
                
                // Transfer
                prepareTransfer: '📤 Prepare row(s) for transfer',
                selected: 'selected',
                flag: 'Column 1 - Flag:',
                comment: 'Column 2 - Comment:',
                commentPlaceholder: 'Free text...',
                addToQueue: '➕ Add to Queue',
                transferDirect: '➡️ Transfer Directly',
                selectAll: '✓ Select All',
                deselectAll: '✗ Deselect All',
                
                // Queue
                queue: '📋 Queue',
                rows: 'rows',
                clear: '🗑️ Clear',
                preview: '👁️ Preview',
                exportToTarget: '📤 Export to Target',
                directTransfer: '✅ Transfer Directly',
                dataExplorer: 'Data Explorer',
                newMonthFile: '📅 New Month File',
                noRowsInQueue: 'No rows in queue',
                
                // New Row
                createNewRow: '✏️ Create New Row',
                close: '✕ Close',
                toQueue: '➕ To Queue',
                transferDirectly: '➡️ Transfer Directly',
                keepFields: 'Keep fields',
                keepFieldsTooltip: 'When enabled, field values are kept after transfer – useful when only one field changes',
                
                // Messages
                noTargetFile: 'No target file loaded',
                selectAtLeastOne: 'Please select at least one row',
                rowsTransferred: 'row(s) transferred directly!',
                rowsAdded: 'row(s) added to queue',
                edited: 'edited',
                duplicates: 'duplicate(s)',
                configFirst: 'Please configure columns first',
                
                // Data Explorer
                explorerTitle: 'Data Explorer',
                noFileLoadedExplorer: 'No file loaded',
                openFile: '📂 Open File',
                fullTextSearch: 'Full-text search:',
                searchPlaceholderExplorer: 'Enter search term...',
                columns: '❙❙ Columns',
                exportXlsx: '📊 Export as XLSX',
                saveFile: '💾 Save',
                filter: '🔍 Filter',
                addFilter: '➕ Add Filter',
                resetFilters: '🗑️ Reset Filters',
                showAll: 'Show all',
                hideAll: 'Hide all',
                noDataLoaded: 'No data loaded.',
                pleaseLoadFile: '📂 Please load an Excel file',
                contains: 'contains',
                equals: 'equals',
                startsWith: 'starts with',
                endsWith: 'ends with',
                selectColumn: 'Select column',
                filterValue: 'Filter value',
                findReplace: '🔄 Replace',
                findReplaceTitle: 'Find & Replace (Ctrl+H)',
                addColumns: '🔗 Add Columns',
                addColumnsTitle: 'Add columns from another Excel file (based on matching serial number)',
                
                // Data Join Modal
                joinModalTitle: '🔗 Add Columns from File',
                joinModalInfo: '<strong style="color: var(--text);">📋 Function:</strong> Adds columns from another Excel file based on a common key (e.g. serial number). Rows are automatically matched by the key.',
                joinStep1: '1️⃣ Select Data Source',
                joinSelectFile: '📂 Select File',
                joinNoFileSelected: 'No file selected',
                joinDropZoneText: 'Drop Excel file here',
                joinWorksheet: 'Worksheet:',
                joinLoadFile: '-- Load file --',
                joinStep2: '2️⃣ Key Column (e.g. Serial Number)',
                joinStep2Desc: 'Select the column that serves as a common key in both files.',
                joinTargetFile: 'Current File (Target):',
                joinSourceFile: 'Source File:',
                joinSelectColumn: '-- Select column --',
                joinStep3: '3️⃣ Select Columns to Add',
                joinStep3Desc: 'Select the columns from the source file to add to the current file.',
                joinLoadSourceFirst: 'Load a source file first',
                joinStep4: '4️⃣ Options',
                joinInsertAfterKeyLabel: 'Insert new columns directly after the key column',
                joinMarkNotFoundLabel: 'Mark rows not found (empty cells = no match)',
                joinPreviewTitle: '📊 Preview',
                joinStatTargetRowsLabel: 'Rows in target:',
                joinStatSourceRowsLabel: 'Rows in source:',
                joinStatMatchesLabel: 'Matches found:',
                joinStatNoMatchLabel: 'No match:',
                joinPreviewBtn: '👁️ Calculate Preview',
                joinExecuteBtn: '✓ Add Columns',
                
                // Row/Column Actions
                insertRowAbove: 'Insert row above',
                insertRowBelow: 'Insert row below',
                deleteRow: 'Delete row',
                hideRow: 'Hide row',
                showRow: 'Show row',
                showAllRows: 'Show all rows',
                hiddenRows: 'hidden rows',
                hiddenColumns: 'hidden columns',
                rowHidden: 'Row hidden',
                rowShown: 'Row shown',
                allRowsShown: 'All rows shown',
                columnHidden: 'Column hidden',
                columnShown: 'Column shown',
                allColumnsShown: 'All columns shown',
                showColumn: 'Show column',
                showAllColumns: 'Show all columns',
                insertColumnBefore: 'Insert column left',
                insertColumnAfter: 'Insert column right',
                deleteColumn: 'Delete column',
                hideColumn: 'Hide column',
                newColumn: 'New Column',
                enterColumnName: 'Enter column name:',
                rowInserted: 'Row inserted',
                rowDeleted: 'Row deleted',
                columnInserted: 'Column inserted',
                columnDeleted: 'Column deleted',
                deleteRowTitle: 'Delete Row?',
                deleteColumnTitle: 'Delete Column?',
                deleteRowConfirm: 'Do you really want to delete this row?',
                deleteColumnConfirm: 'Do you really want to delete this column?',
                deleteColumnWarning: '⚠️ All data in this column will be lost!',
                
                // New Month Modal
                createNewMonthFile: '📅 Create New Month File',
                newMonthDescription: 'Creates a copy of the template file with a new name and sets it as File 2.',
                newFilename: 'New filename:',
                filenamePlaceholder: 'e.g. ContractList_2025-01.xlsx',
                templateLabel: '💡 Template:',
                cancel: 'Cancel',
                createAndLoad: '📅 Create & Load',
                
                // Footer
                copyright: '© Norbert Jander 2026 · v{version}',
                
                // Extra Columns
                extraColumns: 'Extra Columns',
                enableFlag: 'Flag Column (A/D/C)',
                enableComment: 'Comment Column',
                
                // License
                license: 'License',
                licenseTitle: 'License Information',
                thirdPartyLicenses: 'Third-Party Licenses',
                thirdPartyDesc: 'Excel Data Sync Pro uses the following open-source libraries:',
                allLicensesMIT: 'All libraries used are licensed under the MIT License or compatible open-source licenses.',
                packagesTotal: 'packages total',
                
                // Security Logs
                logsButton: 'Logs',
                logsTitle: 'Logs',
                localLogs: 'Local Security Logs',
                securityLogs: 'Security Logs',
                securityLogsTitle: 'Security Logs (Tamper-Proof)',
                integrityStatus: 'Integrity Status:',
                refresh: 'Refresh',
                verify: 'Verify',
                clearLogs: 'Clear Logs',
                
                // Network Logs
                networkLogs: 'Network Logs',
                networkLogsTitle: 'Network Protocol',
                currentComputer: 'Current Computer:',
                networkLogPath: 'Log Path:',
                entries: 'Entries:',
                allComputers: 'All Computers',
                timestamp: 'Timestamp',
                computer: 'Computer',
                action: 'Action',
                file: 'File',
                details: 'Details',
                noNetworkLogs: 'No network logs available',
                networkLogsNote: 'The network log is only kept for files on network drives. It shows which computer performed which action and when (GDPR compliant, no personal data).',
                
                // Live Mode
                liveMode: 'Live Mode',
                liveModeWindows: 'Live Mode (Windows)',
                liveOpen: '🔴 Live Open',
                liveSessionActive: '🔴 Live Session Active',
                liveSessionEnded: 'Live session ended',
                liveOffline: 'Offline',
                liveReady: 'Ready',
                liveConnecting: 'Connecting...',
                endLiveSession: '❌ End Live Session',
                toggleExcel: '👁️ Excel',
                sendToExcel: '▶ To Excel',
                resetFilter: 'Reset Filter',
                moveRowUp: 'Move row up',
                moveRowDown: 'Move row down',
                moveColumnLeft: 'Move column left',
                moveColumnRight: 'Move column right',
                highlightRow: 'Highlight row',
                clearHighlight: 'Clear highlight',
                highlightGreen: '🟢 Green',
                highlightYellow: '🟡 Yellow',
                highlightOrange: '🟠 Orange',
                highlightRed: '🔴 Red',
                highlightBlue: '🔵 Blue',
                highlightPurple: '🟣 Purple',
                
                // Data Explorer additional
                saveChanges: '💾 Save',
                previewChanges: '👁️ Preview',
                worksheetLabel: 'Worksheet:',
                loadFileFirst2: '-- Load file --',
                filterSection: '🔍 Filter',
                searchFor: 'Search for:',
                replaceWith: 'Replace with:',
                textToSearch: 'Text to search...',
                replacementText: 'Replacement text...',
                caseSensitive: 'Case sensitive',
                wholeWord: 'Whole word',
                findNext: '▼ Next',
                replaceOne: 'Replace',
                replaceAll: 'Replace All',
                undo: 'Undo',
                manageSheets: 'Manage worksheets',
                
                // Find & Replace Panel
                searchAndReplace: 'Search & Replace',
                
                // Pagination & Status
                showingRows: 'Showing',
                ofRows: 'of',
                rowsLabel: 'rows',
                totalLabel: 'total',
                editedLabel: 'edited',
                pageLabel: 'Page',
                ofLabel: 'of',
                perPage: '/ Page',
                rowsLoaded: 'rows loaded',
                paginationActive: 'Pagination active',
                loadingFile: 'Loading file...',
                loadingData: 'Loading data...',
                noFileLoaded: 'No file loaded',
                errorLabel: 'Error',
                cellsEdited: 'cell(s) edited',
                loadedFromCache: 'Loaded from cache',
                changes: 'changes',
                editingsRestored: 'edits restored',
                
                // Column panel
                showAllColumns: 'Show all',
                hideAllColumns: 'Hide all',
                
                // Drop Zone
                dropZoneText: 'Drop Excel file here',
                dropZoneHint: 'or click to open',
                closeButton: 'Close'
            }
        };
        
        let currentLanguage = localStorage.getItem('excelSyncLanguage') || 'de';
        let currentTheme = localStorage.getItem('excelSyncTheme') || 'dark';
        
        function t(key) {
            return translations[currentLanguage]?.[key] || translations['de'][key] || key;
        }
        
        function setLanguage(lang) {
            currentLanguage = lang;
            localStorage.setItem('excelSyncLanguage', lang);
            applyTranslations();
        }
        
        function setTheme(theme) {
            currentTheme = theme;
            localStorage.setItem('excelSyncTheme', theme);
            if (theme === 'light') {
                document.body.classList.add('light-theme');
            } else {
                document.body.classList.remove('light-theme');
            }
        }
        
        function applyTranslations() {
            // Update all elements with data-i18n attribute
            document.querySelectorAll('[data-i18n]').forEach(el => {
                const key = el.getAttribute('data-i18n');
                const text = t(key);
                if (el.tagName === 'INPUT' && el.type === 'text') {
                    el.placeholder = text;
                } else if (el.tagName === 'OPTION') {
                    el.textContent = text;
                } else {
                    // Keep existing content if it's a dynamic element (like file info)
                    const dynamicElements = ['file1Info', 'file2Info', 'templateInfo', 'mappingInfo'];
                    if (!dynamicElements.includes(el.id)) {
                        el.innerHTML = text;
                    }
                }
                // Update title attribute if data-i18n-title is present
                const titleKey = el.getAttribute('data-i18n-title');
                if (titleKey) {
                    el.title = t(titleKey);
                }
            });
            
            // Header
            document.querySelector('.logo span').textContent = t('appTitle');
            document.getElementById('btnImportConfig').innerHTML = t('loadConfig');
            document.getElementById('btnExportConfig').innerHTML = t('saveConfig');
            document.getElementById('btnHelp').innerHTML = t('help');
            document.querySelector('.sidebar-header-text').textContent = t('configuration');
            
            // Search section
            document.getElementById('searchInput').placeholder = t('searchPlaceholder');
            document.getElementById('btnSearch').innerHTML = t('search');
            document.getElementById('btnNewRow').innerHTML = t('newRow');
            
            // Empty state
            document.querySelector('.empty-state-title').textContent = t('readyToStart');
            document.querySelector('.empty-state-text').innerHTML = t('instructions');
            
            // Footer
            document.querySelector('footer').innerHTML = t('copyright').replace('{version}', window.__appVersion || window.electronAPI?.appVersion || '');
            
            // Queue panel
            document.getElementById('btnClearQueue').innerHTML = t('clear');
            document.getElementById('btnPreviewTransfer').innerHTML = t('preview');
            document.getElementById('btnExportPS').innerHTML = t('exportToTarget');
            document.getElementById('btnDataExplorer').innerHTML = '📊 ' + t('dataExplorer');
            document.getElementById('btnNewMonth').innerHTML = t('newMonthFile');
            
            // Queue title - update the text before and after the count span
            const queueTitleEl = document.querySelector('#queuePanel .transfer-title');
            if (queueTitleEl) {
                const countSpan = document.getElementById('queueCount');
                const count = countSpan ? countSpan.textContent : '0';
                queueTitleEl.innerHTML = `${t('queue')} (<span id="queueCount">${count}</span> ${t('rows')})`;
            }
            
            // Queue empty message
            const queueEmpty = document.querySelector('.queue-empty');
            if (queueEmpty) {
                queueEmpty.textContent = t('noRowsInQueue');
            }
            
            // Transfer panel
            document.getElementById('btnAddToQueue').innerHTML = t('addToQueue');
            document.getElementById('btnTransferDirect').innerHTML = t('transferDirect');
            document.getElementById('btnSelectAll').innerHTML = t('selectAll');
            document.getElementById('btnDeselectAll').innerHTML = t('deselectAll');
            
            // Theme dropdown options
            const darkOpt = document.querySelector('#selectTheme option[value="dark"]');
            const lightOpt = document.querySelector('#selectTheme option[value="light"]');
            if (darkOpt) darkOpt.textContent = t('themeDark');
            if (lightOpt) lightOpt.textContent = t('themeLight');
            
            // Sidebar buttons
            const btnLoadFile1 = document.getElementById('btnLoadFile1');
            const btnLoadFile2 = document.getElementById('btnLoadFile2');
            const btnConfigMapping = document.getElementById('btnConfigMapping');
            const btnLoadTemplate = document.getElementById('btnLoadTemplate');
            if (btnLoadFile1) btnLoadFile1.innerHTML = t('loadSourceFile');
            if (btnLoadFile2) btnLoadFile2.innerHTML = t('loadTargetFile');
            if (btnConfigMapping) btnConfigMapping.innerHTML = t('configureColumns');
            if (btnLoadTemplate) btnLoadTemplate.innerHTML = t('loadTemplate');
            
            // New row panel
            const newRowTitle = document.querySelector('.new-row-title');
            const btnCloseNewRow = document.getElementById('btnCloseNewRow');
            const btnAddNewRowToQueue = document.getElementById('btnAddNewRowToQueue');
            const btnTransferNewRowDirect = document.getElementById('btnTransferNewRowDirect');
            if (newRowTitle) newRowTitle.textContent = t('createNewRow');
            if (btnCloseNewRow) btnCloseNewRow.innerHTML = t('close');
            if (btnAddNewRowToQueue) btnAddNewRowToQueue.innerHTML = t('toQueue');
            if (btnTransferNewRowDirect) btnTransferNewRowDirect.innerHTML = t('transferDirectly');
            const keepFieldsLabel = document.querySelector('#keepFieldsCheckbox')?.parentElement;
            if (keepFieldsLabel) {
                keepFieldsLabel.title = t('keepFieldsTooltip');
                keepFieldsLabel.childNodes.forEach(n => { if (n.nodeType === 3 && n.textContent.trim()) n.textContent = ' ' + t('keepFields'); });
            }
            
            // Data Explorer Modal
            const explorerTitle = document.querySelector('#dataExplorerModal .modal-title');
            if (explorerTitle) {
                // Nur den Text vor dem Span aktualisieren, nicht das ganze HTML ersetzen
                const fileNameSpan = document.getElementById('explorerFileName');
                const currentFileName = fileNameSpan ? fileNameSpan.textContent : t('noFileLoadedExplorer');
                // Verwende einen TextNode für den Titel-Prefix
                explorerTitle.childNodes[0].textContent = `📊 ${t('explorerTitle')} - `;
                // Falls der Span nicht existiert oder entfernt wurde, neu erstellen
                if (!explorerTitle.querySelector('#explorerFileName')) {
                    const newSpan = document.createElement('span');
                    newSpan.id = 'explorerFileName';
                    newSpan.textContent = currentFileName;
                    explorerTitle.appendChild(newSpan);
                }
            }
            const btnExplorerOpenFile = document.getElementById('btnExplorerOpenFile');
            const btnToggleColumns = document.getElementById('btnToggleColumns');
            const btnExplorerExport = document.getElementById('btnExplorerExport');
            const btnAddExplorerFilter = document.getElementById('btnAddExplorerFilter');
            const btnClearExplorerFilters = document.getElementById('btnClearExplorerFilters');
            const btnShowAllColumns = document.getElementById('btnShowAllColumns');
            const btnHideAllColumns = document.getElementById('btnHideAllColumns');
            const btnCloseExplorerFooter = document.getElementById('btnCloseExplorerFooter');
            const explorerResultCount = document.getElementById('explorerResultCount');
            const explorerSearch = document.getElementById('explorerSearch');
            
            if (btnExplorerOpenFile) btnExplorerOpenFile.innerHTML = t('openFile');
            if (btnToggleColumns) btnToggleColumns.innerHTML = t('columns');
            if (btnExplorerExport) btnExplorerExport.innerHTML = t('exportXlsx');
            if (btnAddExplorerFilter) btnAddExplorerFilter.innerHTML = t('addFilter');
            if (btnClearExplorerFilters) btnClearExplorerFilters.innerHTML = t('resetFilters');
            if (btnShowAllColumns) btnShowAllColumns.textContent = t('showAll');
            if (btnHideAllColumns) btnHideAllColumns.textContent = t('hideAll');
            if (btnCloseExplorerFooter) btnCloseExplorerFooter.textContent = t('close');
            if (explorerResultCount && (explorerResultCount.textContent === 'Keine Daten geladen.' || explorerResultCount.textContent === 'No data loaded.')) {
                explorerResultCount.textContent = t('noDataLoaded');
            }
            if (explorerSearch) explorerSearch.placeholder = t('searchPlaceholderExplorer');
            
            // Data Explorer Preview button
            const btnExplorerPreview = document.getElementById('btnExplorerPreview');
            if (btnExplorerPreview) btnExplorerPreview.innerHTML = t('preview');
            
            // Explorer filter section label
            const filterLabel = document.querySelector('#explorerFilterControls > div > span');
            if (filterLabel) filterLabel.textContent = t('filter');
            
            // Explorer worksheet label and full-text search label
            const worksheetLabels = document.querySelectorAll('#dataExplorerModal .transfer-field label');
            worksheetLabels.forEach(label => {
                if (label.textContent === 'Arbeitsblatt:' || label.textContent === 'Worksheet:') {
                    label.textContent = t('worksheet');
                }
                if (label.textContent === 'Volltextsuche:' || label.textContent === 'Full-text search:') {
                    label.textContent = t('fullTextSearch');
                }
            });
            
            // Explorer filter operator options
            document.querySelectorAll('#explorerFilters .filter-operator').forEach(select => {
                const options = select.querySelectorAll('option');
                options.forEach(opt => {
                    if (opt.value === 'contains') opt.textContent = t('contains');
                    if (opt.value === 'equals') opt.textContent = t('equals');
                    if (opt.value === 'startsWith') opt.textContent = t('startsWith');
                    if (opt.value === 'endsWith') opt.textContent = t('endsWith');
                });
            });
            
            // Data Join Modal - Update select default options
            const joinSourceSheet = document.getElementById('joinSourceSheet');
            const joinTargetKeyColumn = document.getElementById('joinTargetKeyColumn');
            const joinSourceKeyColumn = document.getElementById('joinSourceKeyColumn');
            const joinSourceFileName = document.getElementById('joinSourceFileName');
            const joinColumnsContainer = document.getElementById('joinColumnsContainer');
            
            // Update default options if they exist and are in initial state
            if (joinSourceSheet && joinSourceSheet.options.length === 1) {
                joinSourceSheet.options[0].textContent = t('joinLoadFile');
            }
            if (joinTargetKeyColumn && joinTargetKeyColumn.options.length > 0 && joinTargetKeyColumn.options[0].value === '') {
                joinTargetKeyColumn.options[0].textContent = t('joinSelectColumn');
            }
            if (joinSourceKeyColumn && joinSourceKeyColumn.options.length > 0 && joinSourceKeyColumn.options[0].value === '') {
                joinSourceKeyColumn.options[0].textContent = t('joinSelectColumn');
            }
            // Update "no file selected" text if in initial state
            if (joinSourceFileName && (joinSourceFileName.textContent === 'Keine Datei ausgewählt' || joinSourceFileName.textContent === 'No file selected')) {
                joinSourceFileName.textContent = t('joinNoFileSelected');
            }
            // Update "load source first" text if in initial state
            if (joinColumnsContainer) {
                const initialMsg = joinColumnsContainer.querySelector('div[style*="text-align: center"]');
                if (initialMsg && (initialMsg.textContent.includes('Laden Sie zuerst') || initialMsg.textContent.includes('Load a source file'))) {
                    initialMsg.textContent = t('joinLoadSourceFirst');
                }
            }
            
            // New Month Modal
            const newMonthTitle = document.querySelector('#newMonthModal .modal-title');
            const newMonthDesc = document.querySelector('#newMonthModal .modal-body > p:first-of-type');
            const newMonthLabel = document.querySelector('#newMonthModal .config-label');
            const newMonthFilename = document.getElementById('newMonthFilename');
            const newMonthTemplateLabel = document.querySelector('#newMonthModal .modal-body > p:last-of-type');
            const btnCancelNewMonth = document.getElementById('btnCancelNewMonth');
            const btnConfirmNewMonth = document.getElementById('btnConfirmNewMonth');
            
            if (newMonthTitle) newMonthTitle.textContent = t('createNewMonthFile');
            if (newMonthDesc) newMonthDesc.textContent = t('newMonthDescription');
            if (newMonthLabel) newMonthLabel.textContent = t('newFilename');
            if (newMonthFilename) newMonthFilename.placeholder = t('filenamePlaceholder');
            if (newMonthTemplateLabel) {
                const templateName = document.getElementById('newMonthTemplateName');
                newMonthTemplateLabel.innerHTML = `${t('templateLabel')} <strong id="newMonthTemplateName">${templateName ? templateName.textContent : '-'}</strong>`; 
            }
            if (btnCancelNewMonth) btnCancelNewMonth.textContent = t('cancel');
            if (btnConfirmNewMonth) btnConfirmNewMonth.innerHTML = t('createAndLoad');
            
            // Help Modal - switch content based on language
            const helpContentDe = document.getElementById('helpContentDe');
            const helpContentEn = document.getElementById('helpContentEn');
            const helpModalTitle = document.getElementById('helpModalTitle');
            const btnCloseHelp = document.getElementById('btnCloseHelp');
            
            if (helpContentDe && helpContentEn) {
                if (currentLanguage === 'en') {
                    helpContentDe.style.display = 'none';
                    helpContentEn.style.display = 'block';
                } else {
                    helpContentDe.style.display = 'block';
                    helpContentEn.style.display = 'none';
                }
            }
            if (helpModalTitle) {
                helpModalTitle.textContent = currentLanguage === 'en' ? '❓ Help - Excel Data Sync Pro' : '❓ Hilfe - Excel Data Sync Pro';
            }
            if (btnCloseHelp) {
                btnCloseHelp.textContent = currentLanguage === 'en' ? 'Close' : 'Schließen';
            }
            
            // Column Context Menu - translate items
            const columnContextMenu = document.getElementById('columnContextMenu');
            if (columnContextMenu) {
                const menuItems = columnContextMenu.querySelectorAll('.context-menu-item');
                menuItems.forEach(item => {
                    const action = item.dataset.action;
                    if (action === 'sort-alpha-asc') item.innerHTML = `<span>🔤</span> ${currentLanguage === 'en' ? 'Alphabetical A → Z' : 'Alphabetisch A → Z'}`;
                    if (action === 'sort-alpha-desc') item.innerHTML = `<span>🔤</span> ${currentLanguage === 'en' ? 'Alphabetical Z → A' : 'Alphabetisch Z → A'}`;
                    if (action === 'sort-num-asc') item.innerHTML = `<span>🔢</span> ${currentLanguage === 'en' ? 'Numeric 1 → 9' : 'Numerisch 1 → 9'}`;
                    if (action === 'sort-num-desc') item.innerHTML = `<span>🔢</span> ${currentLanguage === 'en' ? 'Numeric 9 → 1' : 'Numerisch 9 → 1'}`;
                    if (action === 'sort-date-asc') item.innerHTML = `<span>📅</span> ${currentLanguage === 'en' ? 'Date Old → New' : 'Datum Alt → Neu'}`;
                    if (action === 'sort-date-desc') item.innerHTML = `<span>📅</span> ${currentLanguage === 'en' ? 'Date New → Old' : 'Datum Neu → Alt'}`;
                    if (action === 'hide-column') item.innerHTML = `<span>👁️‍🗨️</span> ${t('hideColumn')}`;
                    if (action === 'delete-column') item.innerHTML = `<span>🗑️</span> ${t('deleteColumn')}`;
                    if (action === 'insert-column-before') item.innerHTML = `<span>⬅️</span> ${t('insertColumnBefore')}`;
                    if (action === 'insert-column-after') item.innerHTML = `<span>➡️</span> ${t('insertColumnAfter')}`;
                });
                // Update header
                const header = columnContextMenu.querySelector('.context-menu-header');
                if (header) {
                    const colName = document.getElementById('contextMenuColumnName');
                    header.innerHTML = `${currentLanguage === 'en' ? 'Column' : 'Spalte'}: <span id="contextMenuColumnName">${colName ? colName.textContent : '-'}</span>`;
                }
            }
            
            // Update Live Mode indicator if visible
            updateLiveModeIndicator();
        }
        
        // ==================== JSDoc Type Definitions ====================
        /**
         * @typedef {Object} FileState
         * @property {string|null} name - File name
         * @property {Object|null} workbook - Workbook object
         * @property {string[]} sheets - List of sheet names
         * @property {string|null} selectedSheet - Currently selected sheet
         * @property {Array<Array<string|number>>} data - Sheet data
         * @property {string[]} headers - Column headers
         * @property {string|null} filePath - Full file path
         */

        /**
         * @typedef {Object} MappingConfig
         * @property {number[]} sourceColumns - Source column indices to copy
         * @property {number} targetStartColumn - Target start column (default 1)
         * @property {number} duplicateCheckColumn - Column index for duplicate check
         */

        /**
         * @typedef {Object} TransferQueueItem
         * @property {Array<string|number>} data - Row data array
         * @property {string} flag - Flag value (A/D/C or empty)
         * @property {string} comment - Comment text
         * @property {string} checkValue - Value used for duplicate checking
         * @property {boolean} [wasEdited] - Whether the row was edited
         */

        /**
         * @typedef {Object} TemplateState
         * @property {string|null} name - Template file name
         * @property {Object|null} data - Template data
         * @property {string|null} filePath - Template file path
         */

        /**
         * @typedef {Object} PaginationState
         * @property {number} currentPage - Current page number (1-based)
         * @property {number} pageSize - Items per page
         * @property {number[]} pageSizeOptions - Available page size options
         */

        /**
         * @typedef {Object} AppState
         * @property {FileState} file1 - Source file state
         * @property {FileState} file2 - Target file state
         * @property {MappingConfig} mapping - Column mapping configuration
         * @property {number|null} selectedRow - Currently selected row index
         * @property {number[]} selectedRows - Array of selected row indices
         * @property {Array<Array<string|number>>} searchResults - Search results
         * @property {Array<Object>} history - Transfer history
         * @property {TransferQueueItem[]} transferQueue - Items queued for transfer
         * @property {TemplateState} template - Template file state
         * @property {Object|null} lastDirectoryHandle - Last used directory handle
         * @property {PaginationState} searchPagination - Search results pagination
         */

        // ==================== Undo/Redo System ====================
        const undoRedoState = {
            // Suchergebnisse
            searchUndoStack: [],
            searchRedoStack: [],
            // Datenexplorer
            explorerUndoStack: [],
            explorerRedoStack: [],
            maxStackSize: 50
        };
        
        function pushSearchUndo(action) {
            undoRedoState.searchUndoStack.push(action);
            if (undoRedoState.searchUndoStack.length > undoRedoState.maxStackSize) {
                undoRedoState.searchUndoStack.shift();
            }
            undoRedoState.searchRedoStack = []; // Redo-Stack leeren bei neuer Aktion
        }
        
        function pushExplorerUndo(action) {
            undoRedoState.explorerUndoStack.push(action);
            if (undoRedoState.explorerUndoStack.length > undoRedoState.maxStackSize) {
                undoRedoState.explorerUndoStack.shift();
            }
            undoRedoState.explorerRedoStack = []; // Redo-Stack leeren bei neuer Aktion
        }
        
        function undoSearch() {
            if (undoRedoState.searchUndoStack.length === 0) return false;
            
            const action = undoRedoState.searchUndoStack.pop();
            undoRedoState.searchRedoStack.push(action);
            
            // Ursprünglichen Wert wiederherstellen
            const { rowIndex, colIndex, oldValue, newValue } = action;
            state.searchResults[rowIndex].data[colIndex] = oldValue;
            
            // UI aktualisieren
            const cell = document.querySelector(`#resultsTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
            if (cell) {
                cell.textContent = oldValue;
                cell.classList.toggle('edited', oldValue !== cell.dataset.original);
            }
            return true;
        }
        
        function redoSearch() {
            if (undoRedoState.searchRedoStack.length === 0) return false;
            
            const action = undoRedoState.searchRedoStack.pop();
            undoRedoState.searchUndoStack.push(action);
            
            // Neuen Wert wiederherstellen
            const { rowIndex, colIndex, oldValue, newValue } = action;
            state.searchResults[rowIndex].data[colIndex] = newValue;
            
            // UI aktualisieren
            const cell = document.querySelector(`#resultsTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
            if (cell) {
                cell.textContent = newValue;
                cell.classList.toggle('edited', newValue !== cell.dataset.original);
            }
            return true;
        }
        
        async function undoExplorer() {
            if (undoRedoState.explorerUndoStack.length === 0) return false;
            
            const action = undoRedoState.explorerUndoStack.pop();
            undoRedoState.explorerRedoStack.push(action);
            
            // Prüfe auf moveRows Aktion
            if (action.type === 'moveRows') {
                // HINWEIS: Undo für moveRows aktuell nicht unterstützt (Performance-Optimierung)
                showNotification('Undo für Zeilen-Verschiebung nicht möglich. Bitte Datei neu laden.', 'warning');
                return false;
            }
            
            // Prüfe auf deleteRows Aktion – gelöschte Zeilen wiederherstellen
            if (action.type === 'deleteRows') {
                // Zeilen in aufsteigender Reihenfolge wieder einfügen
                const sorted = [...action.deletedRows].sort((a, b) => a.index - b.index);
                for (const entry of sorted) {
                    explorerState.data.splice(entry.index, 0, entry.data);
                    if (explorerState.originalData) {
                        explorerState.originalData.splice(entry.index, 0, entry.originalData);
                    }
                }
                // Highlights wiederherstellen
                if (action.previousHighlights) {
                    explorerState.rowHighlights = new Map(action.previousHighlights);
                }
                // EditedCells wiederherstellen
                if (action.previousEditedCells) {
                    explorerState.editedCells = new Map(action.previousEditedCells);
                }
                // FilteredData neu erstellen
                explorerState.filteredData = explorerState.data.map((row, idx) => ({
                    row: row,
                    originalIndex: idx
                }));
                updateExplorerEditStatus();
                renderExplorerTable();
                const msg = currentLanguage === 'en'
                    ? `${sorted.length} row(s) restored`
                    : `${sorted.length} Zeile(n) wiederhergestellt`;
                showNotification(msg, 'info');
                return true;
            }
            
            const liveActive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            
            // Prüfe auf Multi-Aktion (mehrere Zellen gleichzeitig)
            if (action.type === 'multi') {
                const cellsToSync = [];
                
                action.actions.forEach(subAction => {
                    const { rowIndex, colIndex, oldValue, originalValue } = subAction;
                    explorerState.data[rowIndex][colIndex] = oldValue;
                    
                    // Auch filteredData aktualisieren
                    const filteredItem = explorerState.filteredData.find(item => item.originalIndex === rowIndex);
                    if (filteredItem && filteredItem.row) {
                        filteredItem.row[colIndex] = oldValue;
                    }
                    
                    const cellKey = `${rowIndex}-${colIndex}`;
                    if (oldValue === originalValue) {
                        explorerState.editedCells.delete(cellKey);
                    } else {
                        explorerState.editedCells.set(cellKey, oldValue);
                    }
                    
                    // Für Live-Session sammeln (oldValue = was aktuell in Excel steht = newValue der ursprünglichen Aktion)
                    cellsToSync.push({ row: rowIndex, col: colIndex, value: oldValue, oldValue: subAction.newValue });
                    
                    // UI aktualisieren
                    const cell = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                    if (cell) {
                        cell.textContent = oldValue;
                        cell.dataset.lastValue = oldValue;
                        cell.classList.toggle('edited', oldValue !== cell.dataset.original);
                    }
                });
                
                // Live-Session Sync
                if (liveActive && cellsToSync.length > 0) {
                    // Prüfe ob es ein einfaches Suchen/Ersetzen war (alle actions haben gleichen newValue/oldValue Unterschied)
                    // Falls ja, nutze native findReplace in umgekehrter Richtung
                    const firstAction = action.actions[0];
                    const canUseNativeReplace = action.searchText && action.replaceText;
                    
                    if (canUseNativeReplace) {
                        // Umgekehrtes Suchen & Ersetzen
                        try {
                            await window.electronAPI.liveSessionFindReplace(
                                action.replaceText,  // Was wir jetzt suchen
                                action.searchText,   // Was wir zurücksetzen
                                action.matchCase || false,
                                action.wholeWord || false
                            );
                        } catch (error) {
                            console.error('[Undo] Native Replace fehlgeschlagen:', error);
                        }
                    } else {
                        // Fallback: Zellen einzeln setzen
                        try {
                            await window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(cellsToSync));
                        } catch (error) {
                            console.error('[Undo] Batch-Sync fehlgeschlagen:', error);
                        }
                    }
                }
                
                updateExplorerEditStatus();
                renderExplorerTable();
                showNotification(`${action.actions.length} Zelle(n) wiederhergestellt`, 'info');
                return true;
            }
            
            // Ursprünglichen Wert wiederherstellen (Standard-Zellbearbeitung)
            const { rowIndex, colIndex, oldValue, newValue } = action;
            explorerState.data[rowIndex][colIndex] = oldValue;
            
            // Auch filteredData aktualisieren
            const filteredItem = explorerState.filteredData.find(item => item.originalIndex === rowIndex);
            if (filteredItem && filteredItem.row) {
                filteredItem.row[colIndex] = oldValue;
            }
            
            const cellKey = `${rowIndex}-${colIndex}`;
            if (oldValue === action.originalValue) {
                explorerState.editedCells.delete(cellKey);
            } else {
                explorerState.editedCells.set(cellKey, oldValue);
            }
            
            // Live-Session Sync
            if (liveActive) {
                try {
                    await window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols([
                        { row: rowIndex, col: colIndex, value: oldValue, oldValue: newValue }
                    ]));
                } catch (error) {
                    console.error('[Undo] Live-Sync fehlgeschlagen:', error);
                }
            }
            
            // UI aktualisieren
            updateExplorerEditStatus();
            renderExplorerTable();
            return true;
        }
        
        async function redoExplorer() {
            if (undoRedoState.explorerRedoStack.length === 0) return false;
            
            const action = undoRedoState.explorerRedoStack.pop();
            undoRedoState.explorerUndoStack.push(action);
            
            // Prüfe auf moveRows Aktion
            if (action.type === 'moveRows') {
                // HINWEIS: Redo für moveRows aktuell nicht unterstützt (Performance-Optimierung)
                showNotification('Redo für Zeilen-Verschiebung nicht möglich.', 'warning');
                return false;
            }
            
            // Prüfe auf deleteRows Aktion – Zeilen erneut löschen
            if (action.type === 'deleteRows') {
                // Zeilen in absteigender Reihenfolge entfernen
                const sorted = [...action.deletedRows].sort((a, b) => b.index - a.index);
                for (const entry of sorted) {
                    explorerState.data.splice(entry.index, 1);
                    if (explorerState.originalData) {
                        explorerState.originalData.splice(entry.index, 1);
                    }
                }
                // Highlights/EditedCells für gelöschte Zeilen entfernen
                for (const entry of sorted) {
                    explorerState.rowHighlights.delete(entry.index);
                    for (const cellKey of [...explorerState.editedCells.keys()]) {
                        if (cellKey.startsWith(entry.index + '-')) {
                            explorerState.editedCells.delete(cellKey);
                        }
                    }
                }
                // FilteredData neu erstellen
                explorerState.filteredData = explorerState.data.map((row, idx) => ({
                    row: row,
                    originalIndex: idx
                }));
                updateExplorerEditStatus();
                renderExplorerTable();
                const msg = currentLanguage === 'en'
                    ? `${sorted.length} row(s) deleted again`
                    : `${sorted.length} Zeile(n) erneut gelöscht`;
                showNotification(msg, 'info');
                return true;
            }
            
            const liveActive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            
            // Prüfe auf Multi-Aktion (mehrere Zellen gleichzeitig)
            if (action.type === 'multi') {
                const cellsToSync = [];
                
                action.actions.forEach(subAction => {
                    const { rowIndex, colIndex, newValue, originalValue } = subAction;
                    explorerState.data[rowIndex][colIndex] = newValue;
                    
                    // Auch filteredData aktualisieren
                    const filteredItem = explorerState.filteredData.find(item => item.originalIndex === rowIndex);
                    if (filteredItem && filteredItem.row) {
                        filteredItem.row[colIndex] = newValue;
                    }
                    
                    const cellKey = `${rowIndex}-${colIndex}`;
                    if (newValue === originalValue) {
                        explorerState.editedCells.delete(cellKey);
                    } else {
                        explorerState.editedCells.set(cellKey, newValue);
                    }
                    
                    // Für Live-Session sammeln (oldValue = was aktuell in Excel steht = oldValue der Aktion)
                    cellsToSync.push({ row: rowIndex, col: colIndex, value: newValue, oldValue: subAction.oldValue });
                    
                    // UI aktualisieren
                    const cell = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                    if (cell) {
                        cell.textContent = newValue;
                        cell.dataset.lastValue = newValue;
                        cell.classList.toggle('edited', newValue !== cell.dataset.original);
                    }
                });
                
                // Live-Session Sync
                if (liveActive && cellsToSync.length > 0) {
                    const canUseNativeReplace = action.searchText && action.replaceText;
                    
                    if (canUseNativeReplace) {
                        try {
                            await window.electronAPI.liveSessionFindReplace(
                                action.searchText,
                                action.replaceText,
                                action.matchCase || false,
                                action.wholeWord || false
                            );
                        } catch (error) {
                            console.error('[Redo] Native Replace fehlgeschlagen:', error);
                        }
                    } else {
                        try {
                            await window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(cellsToSync));
                        } catch (error) {
                            console.error('[Redo] Batch-Sync fehlgeschlagen:', error);
                        }
                    }
                }
                
                updateExplorerEditStatus();
                renderExplorerTable();
                showNotification(`${action.actions.length} Zelle(n) wiederhergestellt`, 'info');
                return true;
            }
            
            // Neuen Wert wiederherstellen (Standard-Zellbearbeitung)
            const { rowIndex, colIndex, oldValue, newValue, originalValue } = action;
            explorerState.data[rowIndex][colIndex] = newValue;
            
            // Auch filteredData aktualisieren
            const filteredItem = explorerState.filteredData.find(item => item.originalIndex === rowIndex);
            if (filteredItem && filteredItem.row) {
                filteredItem.row[colIndex] = newValue;
            }
            
            const cellKey = `${rowIndex}-${colIndex}`;
            if (newValue === originalValue) {
                explorerState.editedCells.delete(cellKey);
            } else {
                explorerState.editedCells.set(cellKey, newValue);
            }
            
            // Live-Session Sync
            if (liveActive) {
                try {
                    await window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols([
                        { row: rowIndex, col: colIndex, value: newValue, oldValue: oldValue }
                    ]));
                } catch (error) {
                    console.error('[Redo] Live-Sync fehlgeschlagen:', error);
                }
            }
            
            // UI aktualisieren
            updateExplorerEditStatus();
            renderExplorerTable();
            return true;
        }
        
        function showUndoRedoFeedback(action) {
            // Kurzes visuelles Feedback
            const toast = document.createElement('div');
            toast.style.cssText = `
                position: fixed;
                bottom: 80px;
                left: 50%;
                transform: translateX(-50%);
                background: var(--bg-lighter);
                color: #ff9800;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 13px;
                z-index: 10000;
                box-shadow: 0 2px 8px rgba(0,0,0,0.3);
                animation: fadeInOut 1.5s ease-in-out;
            `;
            toast.textContent = action;
            document.body.appendChild(toast);
            setTimeout(() => toast.remove(), 1500);
        }
        
        // Globale Notification-Funktion für Erfolgs-, Warn- und Fehlermeldungen
        function showNotification(message, type = 'info') {
            const colors = {
                success: { bg: '#217346', border: '#2d9a5d' },
                error: { bg: '#d32f2f', border: '#f44336' },
                warning: { bg: '#f57c00', border: '#ff9800' },
                info: { bg: '#1976d2', border: '#2196f3' }
            };
            const color = colors[type] || colors.info;
            
            const notification = document.createElement('div');
            notification.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                background: ${color.bg};
                border: 1px solid ${color.border};
                color: white;
                padding: 12px 20px;
                border-radius: 6px;
                font-size: 14px;
                z-index: 100000;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                animation: slideInRight 0.3s ease-out;
                max-width: 400px;
                word-wrap: break-word;
            `;
            notification.textContent = message;
            document.body.appendChild(notification);
            
            setTimeout(() => {
                notification.style.animation = 'slideOutRight 0.3s ease-in';
                setTimeout(() => notification.remove(), 300);
            }, 3000);
        }

        // ==================== Auto-Save System ====================
        const AUTO_SAVE_KEY = 'excelsync_autosave';
        const CLEAN_SHUTDOWN_KEY = 'excelsync_clean_shutdown';
        const AUTO_SAVE_INTERVAL = 30000; // 30 Sekunden
        let autoSaveTimer = null;
        
        function getAutoSaveData() {
            // Sammle alle bearbeiteten Daten
            const data = {
                timestamp: Date.now(),
                version: '1.1',
                // Quelldatei (Datei 1)
                file1: {
                    filePath: state.file1.filePath || null,
                    name: state.file1.name || null,
                    selectedSheet: state.file1.selectedSheet || null
                },
                // Zieldatei (Datei 2)
                file2: {
                    filePath: state.file2.filePath || null,
                    name: state.file2.name || null,
                    selectedSheet: state.file2.selectedSheet || null
                },
                // Mapping
                mapping: state.mapping,
                // Datenexplorer
                explorer: {
                    filePath: explorerState.filePath,
                    fileName: explorerState.fileName,
                    selectedSheet: explorerState.selectedSheet,
                    editedCells: Array.from(explorerState.editedCells.entries())
                },
                // Warteschlange
                transferQueue: state.transferQueue,
                // Suchergebnisse-Bearbeitungen (nur wenn Suche aktiv)
                searchEdits: []
            };
            
            // Suchergebnis-Bearbeitungen sammeln
            document.querySelectorAll('#resultsTableBody td.edited').forEach(td => {
                data.searchEdits.push({
                    row: parseInt(td.dataset.row),
                    col: parseInt(td.dataset.col),
                    original: td.dataset.original,
                    current: td.textContent
                });
            });
            
            return data;
        }
        
        let _lastAutoSaveFingerprint = '';
        function autoSave() {
            // Schnelle Prüfung: Hat sich seit letztem Save etwas geändert?
            const editCount = explorerState.editedCells.size;
            const queueCount = state.transferQueue.length;
            const fp = `${editCount}|${queueCount}`;
            
            if (editCount === 0 && queueCount === 0) {
                // Prüfe ob noch DOM-Edits vorhanden sind (searchEdits)
                const editedCells = document.querySelectorAll('#resultsTableBody td.edited');
                if (editedCells.length === 0) {
                    if (_lastAutoSaveFingerprint !== '') {
                        localStorage.removeItem(AUTO_SAVE_KEY);
                        _lastAutoSaveFingerprint = '';
                    }
                    return;
                }
            }
            
            if (fp === _lastAutoSaveFingerprint) return;
            
            const data = getAutoSaveData();
            
            // Nur speichern wenn es etwas zu speichern gibt
            const hasExplorerEdits = data.explorer.editedCells.length > 0;
            const hasQueueItems = data.transferQueue.length > 0;
            const hasSearchEdits = data.searchEdits.length > 0;
            
            if (hasExplorerEdits || hasQueueItems || hasSearchEdits) {
                try {
                    localStorage.setItem(AUTO_SAVE_KEY, JSON.stringify(data));
                    _lastAutoSaveFingerprint = fp;
                    console.log('Auto-Save: Daten gesichert', {
                        explorerEdits: data.explorer.editedCells.length,
                        queueItems: data.transferQueue.length,
                        searchEdits: data.searchEdits.length
                    });
                } catch (e) {
                    console.warn('Auto-Save fehlgeschlagen:', e);
                }
            }
        }
        
        function clearAutoSave() {
            localStorage.removeItem(AUTO_SAVE_KEY);
            localStorage.removeItem(CLEAN_SHUTDOWN_KEY);
            _lastAutoSaveFingerprint = '';
        }
        
        // Öffnet den Recovery-Ordner im Dateimanager
        async function openRecoveryFolder() {
            try {
                const result = await window.electronAPI.liveSessionOpenRecoveryFolder();
                if (result.success) {
                    console.log('[Recovery] Ordner geöffnet:', result.path);
                } else {
                    console.error('[Recovery] Fehler:', result.error);
                }
            } catch (error) {
                console.error('[Recovery] Fehler:', error);
            }
        }
        
        // Prüft ob Live-Session Recovery-Dateien vorhanden sind
        async function checkLiveSessionRecovery() {
            try {
                const result = await window.electronAPI.liveSessionGetRecoveryFiles();
                if (!result.success || !result.files || result.files.length === 0) {
                    return;
                }
                
                // Gruppiere nach Original-Datei
                const backups = result.files.filter(f => f.type === 'backup');
                const journals = result.files.filter(f => f.type === 'journal');
                
                if (backups.length === 0 && journals.length === 0) {
                    return;
                }
                
                // Zusammenfassung erstellen
                let summary = '🛡️ Live-Modus Recovery-Dateien gefunden:\n\n';
                
                for (const backup of backups) {
                    summary += `📁 Backup: ${backup.name}\n`;
                    summary += `   Größe: ${Math.round(backup.size / 1024)} KB\n`;
                    summary += `   Alter: ${backup.ageHours} Stunden\n\n`;
                }
                
                for (const journal of journals) {
                    summary += `📝 Journal: ${journal.name}\n`;
                    summary += `   Original: ${journal.originalFile || 'unbekannt'}\n`;
                    summary += `   Einträge: ${journal.entryCount}\n`;
                    summary += `   Alter: ${journal.ageHours} Stunden\n\n`;
                }
                
                summary += 'Diese Dateien können zur Wiederherstellung verwendet werden.\n\n';
                summary += 'Möchten Sie den Recovery-Ordner öffnen?';
                
                if (confirm(summary)) {
                    await openRecoveryFolder();
                }
            } catch (error) {
                console.error('[Recovery] Prüfung fehlgeschlagen:', error);
            }
        }
        
        // Markiert einen sauberen App-Shutdown (kein Crash)
        function markCleanShutdown() {
            localStorage.setItem(CLEAN_SHUTDOWN_KEY, 'true');
            // AutoSave-Daten löschen bei normalem Beenden
            localStorage.removeItem(AUTO_SAVE_KEY);
        }
        
        // Prüft ob der letzte Shutdown sauber war
        function wasCleanShutdown() {
            return localStorage.getItem(CLEAN_SHUTDOWN_KEY) === 'true';
        }
        
        async function checkAutoSaveRecovery() {
            try {
                // === 1. Live-Session Recovery-Dateien prüfen ===
                await checkLiveSessionRecovery();
                
                // === 2. LocalStorage Auto-Save prüfen ===
                const saved = localStorage.getItem(AUTO_SAVE_KEY);
                const cleanShutdown = wasCleanShutdown();
                
                // Clean-Shutdown-Flag löschen für nächsten Start
                localStorage.removeItem(CLEAN_SHUTDOWN_KEY);
                
                // Wenn sauberer Shutdown oder keine Daten -> nichts tun
                if (!saved || cleanShutdown) {
                    if (saved) clearAutoSave(); // Alte Daten löschen falls vorhanden
                    return;
                }
                
                const data = JSON.parse(saved);
                const age = Date.now() - data.timestamp;
                const ageMinutes = Math.round(age / 60000);
                
                // Nur wiederherstellen wenn weniger als 24 Stunden alt
                if (age > 24 * 60 * 60 * 1000) {
                    clearAutoSave();
                    return;
                }
                
                const hasExplorerEdits = data.explorer?.editedCells?.length > 0;
                const hasQueueItems = data.transferQueue?.length > 0;
                const hasSearchEdits = data.searchEdits?.length > 0;
                const hasFile1 = data.file1?.filePath;
                const hasFile2 = data.file2?.filePath;
                
                if (!hasExplorerEdits && !hasQueueItems && !hasSearchEdits && !hasFile1 && !hasFile2) {
                    clearAutoSave();
                    return;
                }
                
                // Zusammenfassung erstellen
                let summary = 'Ungespeicherte Daten gefunden:\n\n';
                if (hasFile1) {
                    summary += `• Quelldatei: ${data.file1.name || data.file1.filePath}\n`;
                }
                if (hasFile2) {
                    summary += `• Zieldatei: ${data.file2.name || data.file2.filePath}\n`;
                }
                if (hasQueueItems) {
                    summary += `• ${data.transferQueue.length} Einträge in der Warteschlange\n`;
                }
                if (hasExplorerEdits) {
                    summary += `• ${data.explorer.editedCells.length} bearbeitete Zellen im Datenexplorer\n`;
                    if (data.explorer.fileName) {
                        summary += `  (Datei: ${data.explorer.fileName})\n`;
                    }
                }
                if (hasSearchEdits) {
                    summary += `• ${data.searchEdits.length} bearbeitete Suchergebnisse\n`;
                }
                summary += `\nGespeichert vor ${ageMinutes} Minuten.\n\nMöchten Sie diese Daten wiederherstellen?`;
                
                if (confirm(summary)) {
                    await restoreAutoSave(data);
                } else {
                    clearAutoSave();
                }
            } catch (e) {
                console.warn('Auto-Save Recovery fehlgeschlagen:', e);
                clearAutoSave();
            }
        }
        
        async function restoreAutoSave(data) {
            console.log('[Auto-Save] Starte Wiederherstellung...', data);
            
            // Mapping wiederherstellen (vor den Dateien, da es für die Anzeige benötigt wird)
            if (data.mapping) {
                state.mapping = data.mapping;
                updateMappingPreview();
            }
            
            // Die Wiederherstellung beider Dateien ist unabhängig. Parallel
            // laden, damit die Startzeit nicht die Summe beider Dateien ist.
            const restoreTasks = [];
            if (data.file1?.filePath) {
                console.log('[Auto-Save] Lade Quelldatei:', data.file1.filePath);
                restoreTasks.push(
                    loadMainFileFromPath(1, data.file1.filePath, data.file1.selectedSheet)
                        .then(result => result.success
                            ? console.log('[Auto-Save] Quelldatei wiederhergestellt')
                            : console.warn('[Auto-Save] Quelldatei konnte nicht geladen werden:', result.error))
                        .catch(error => console.warn('[Auto-Save] Quelldatei konnte nicht geladen werden:', error))
                );
            }
            if (data.file2?.filePath) {
                console.log('[Auto-Save] Lade Zieldatei:', data.file2.filePath);
                restoreTasks.push(
                    loadMainFileFromPath(2, data.file2.filePath, data.file2.selectedSheet)
                        .then(result => result.success
                            ? console.log('[Auto-Save] Zieldatei wiederhergestellt')
                            : console.warn('[Auto-Save] Zieldatei konnte nicht geladen werden:', result.error))
                        .catch(error => console.warn('[Auto-Save] Zieldatei konnte nicht geladen werden:', error))
                );
            }
            await Promise.all(restoreTasks);
            
            // Warteschlange wiederherstellen
            if (data.transferQueue?.length > 0) {
                state.transferQueue = data.transferQueue;
                updateQueueDisplay();
                showUndoRedoFeedback(`${data.transferQueue.length} Einträge wiederhergestellt`);
            }
            
            // Datenexplorer-Bearbeitungen werden beim Öffnen der Datei wiederhergestellt
            if (data.explorer?.editedCells?.length > 0 && data.explorer.filePath) {
                // Speichere für spätere Wiederherstellung
                window._pendingExplorerRestore = data.explorer;
            }
            
            // Button-Status nochmal aktualisieren (nach allen async Operationen)
            updateQueueDisplay();
            
            console.log('[Auto-Save] Wiederherstellung abgeschlossen. file2.filePath:', state.file2.filePath);
            
            // Auto-Save nach Wiederherstellung nicht löschen (wird bei nächster Aktion überschrieben)
        }
        
        function startAutoSave() {
            if (autoSaveTimer) clearInterval(autoSaveTimer);
            autoSaveTimer = setInterval(autoSave, AUTO_SAVE_INTERVAL);
        }
        
        function stopAutoSave() {
            if (autoSaveTimer) {
                clearInterval(autoSaveTimer);
                autoSaveTimer = null;
            }
        }

        // ==================== Such-Historie System ====================
        const SEARCH_HISTORY_KEY = 'excelsync_search_history';
        const SEARCH_HISTORY_MAX = 15;
        let searchHistorySelectedIndex = -1;
        let _searchHistoryCache = null; // In-memory Cache
        
        function getSearchHistory() {
            if (_searchHistoryCache !== null) return _searchHistoryCache;
            try {
                const saved = localStorage.getItem(SEARCH_HISTORY_KEY);
                _searchHistoryCache = saved ? JSON.parse(saved) : [];
                return _searchHistoryCache;
            } catch (e) {
                _searchHistoryCache = [];
                return [];
            }
        }
        
        function saveSearchHistory(history) {
            _searchHistoryCache = history; // Cache aktualisieren
            try {
                localStorage.setItem(SEARCH_HISTORY_KEY, JSON.stringify(history));
            } catch (e) {
                console.warn('Such-Historie speichern fehlgeschlagen:', e);
            }
        }
        
        function addToSearchHistory(term, resultCount) {
            if (!term || term.trim().length === 0) return;
            
            const trimmed = term.trim();
            let history = getSearchHistory();
            
            // Existierenden Eintrag entfernen (wird oben neu eingefügt)
            history = history.filter(item => item.term.toLowerCase() !== trimmed.toLowerCase());
            
            // Neuen Eintrag am Anfang einfügen
            history.unshift({
                term: trimmed,
                count: resultCount,
                timestamp: Date.now()
            });
            
            // Auf Maximum begrenzen
            if (history.length > SEARCH_HISTORY_MAX) {
                history = history.slice(0, SEARCH_HISTORY_MAX);
            }
            
            saveSearchHistory(history);
        }
        
        function removeFromSearchHistory(term) {
            let history = getSearchHistory();
            history = history.filter(item => item.term !== term);
            saveSearchHistory(history);
            renderSearchHistoryDropdown();
        }
        
        function clearSearchHistory() {
            _searchHistoryCache = []; // Cache leeren
            localStorage.removeItem(SEARCH_HISTORY_KEY);
            renderSearchHistoryDropdown();
            hideSearchHistoryDropdown();
        }
        
        function renderSearchHistoryDropdown(filterText = '') {
            const dropdown = document.getElementById('searchHistoryDropdown');
            if (!dropdown) return;
            
            let history = getSearchHistory();
            
            // Nach Filter filtern
            if (filterText) {
                const lower = filterText.toLowerCase();
                history = history.filter(item => item.term.toLowerCase().includes(lower));
            }
            
            if (history.length === 0) {
                dropdown.innerHTML = '';
                dropdown.classList.remove('show');
                return;
            }
            
            let html = `
                <div class="search-history-header">
                    <span>🕐 Letzte Suchen</span>
                    <button class="search-history-clear" onclick="clearSearchHistory()">Alle löschen</button>
                </div>
            `;
            
            history.forEach((item, index) => {
                const selected = index === searchHistorySelectedIndex ? ' selected' : '';
                html += `
                    <div class="search-history-item${selected}" 
                         data-index="${index}"
                         data-term="${escapeHtml(item.term)}">
                        <span class="search-history-text">${escapeHtml(item.term)}</span>
                        <span class="search-history-count">${item.count} Treffer</span>
                        <button class="search-history-delete" 
                                onclick="event.stopPropagation(); removeFromSearchHistory('${escapeHtml(item.term).replace(/'/g, "\\'")}')">✕</button>
                    </div>
                `;
            });
            
            dropdown.innerHTML = html;
            
            // Click-Handler für Items
            dropdown.querySelectorAll('.search-history-item').forEach(item => {
                item.addEventListener('click', () => {
                    const term = item.dataset.term;
                    elements.searchInput.value = term;
                    hideSearchHistoryDropdown();
                    search();
                });
            });
        }
        
        function showSearchHistoryDropdown() {
            const dropdown = document.getElementById('searchHistoryDropdown');
            const history = getSearchHistory();
            
            if (history.length === 0) return;
            
            searchHistorySelectedIndex = -1;
            renderSearchHistoryDropdown(elements.searchInput.value);
            dropdown.classList.add('show');
        }
        
        function hideSearchHistoryDropdown() {
            const dropdown = document.getElementById('searchHistoryDropdown');
            if (dropdown) {
                dropdown.classList.remove('show');
            }
            searchHistorySelectedIndex = -1;
        }
        
        function navigateSearchHistory(direction) {
            const history = getSearchHistory();
            const filterText = elements.searchInput.value;
            const filtered = filterText 
                ? history.filter(item => item.term.toLowerCase().includes(filterText.toLowerCase()))
                : history;
            
            if (filtered.length === 0) return;
            
            if (direction === 'down') {
                searchHistorySelectedIndex = Math.min(searchHistorySelectedIndex + 1, filtered.length - 1);
            } else if (direction === 'up') {
                searchHistorySelectedIndex = Math.max(searchHistorySelectedIndex - 1, -1);
            }
            
            renderSearchHistoryDropdown(filterText);
            
            // Bei Auswahl den Text ins Feld setzen
            if (searchHistorySelectedIndex >= 0 && filtered[searchHistorySelectedIndex]) {
                elements.searchInput.value = filtered[searchHistorySelectedIndex].term;
            }
        }

        // ==================== State ====================
        /** @type {AppState} */
        const state = {
            file1: {
                name: null,
                workbook: null,
                sheets: [],
                selectedSheet: null,
                data: [],
                headers: [],
                filePath: null
            },
            file2: {
                name: null,
                workbook: null,
                sheets: [],
                selectedSheet: null,
                data: [],
                headers: [],
                filePath: null
            },
            mapping: {
                sourceColumns: [],
                targetStartColumn: 1,
                duplicateCheckColumn: 0,
                dropdownColumns: [],  // Array von Spalten-Indizes, die als Dropdown angezeigt werden sollen
                changeRequestPattern: '*_Change_Request*.*'  // Suchmuster für Duplikat-Prüfung in Verzeichnis-Dateien
            },
            selectedRow: null,
            selectedRows: [],
            searchResults: [],
            history: [],
            transferQueue: [],
            template: {
                name: null,
                data: null
            },
            lastDirectoryHandle: null,
            // Pagination für Suchergebnisse
            searchPagination: {
                currentPage: 1,
                pageSize: 100,
                pageSizeOptions: [50, 100, 250, 500]
            },
            // Arbeitsordner
            workingDirectory: null,
            // Cache für Change-Request-Dateien (für Duplikatprüfung)
            changeRequestCache: {
                files: [],           // Liste der gefundenen Dateien
                data: new Map(),     // Map<filePath, {headers, data, checkColumn}>
                lastUpdate: null,    // Zeitstempel der letzten Aktualisierung
                directory: null      // Verzeichnis das gescannt wurde
            }
        };
        
        // ==================== DOM Elements ====================
        const elements = {
            // Working Directory
            btnSelectWorkingDir: document.getElementById('btnSelectWorkingDir'),
            workingDirInfo: document.getElementById('workingDirInfo'),
            btnClearWorkingDir: document.getElementById('btnClearWorkingDir'),
            
            // File 1
            btnLoadFile1: document.getElementById('btnLoadFile1'),
            fileInput1: document.getElementById('fileInput1'),
            file1Info: document.getElementById('file1Info'),
            selectSheet1: document.getElementById('selectSheet1'),
            
            // File 2
            btnLoadFile2: document.getElementById('btnLoadFile2'),
            fileInput2: document.getElementById('fileInput2'),
            file2Info: document.getElementById('file2Info'),
            selectSheet2: document.getElementById('selectSheet2'),
            
            // Mapping
            btnConfigMapping: document.getElementById('btnConfigMapping'),
            mappingInfo: document.getElementById('mappingInfo'),
            mappingModal: document.getElementById('mappingModal'),
            mappingList: document.getElementById('mappingList'),
            
            // Search
            searchInput: document.getElementById('searchInput'),
            searchHistoryDropdown: document.getElementById('searchHistoryDropdown'),
            btnSearch: document.getElementById('btnSearch'),
            btnNewRow: document.getElementById('btnNewRow'),
            searchResultsInfo: document.getElementById('searchResultsInfo'),
            
            // New Row
            newRowPanel: document.getElementById('newRowPanel'),
            newRowForm: document.getElementById('newRowForm'),
            newRowFlag: document.getElementById('newRowFlag'),
            newRowComment: document.getElementById('newRowComment'),
            btnCloseNewRow: document.getElementById('btnCloseNewRow'),
            btnAddNewRowToQueue: document.getElementById('btnAddNewRowToQueue'),
            btnTransferNewRowDirect: document.getElementById('btnTransferNewRowDirect'),
            newRowStatus: document.getElementById('newRowStatus'),
            
            // Results
            emptyState: document.getElementById('emptyState'),
            resultsTableContainer: document.getElementById('resultsTableContainer'),
            resultsTable: document.getElementById('resultsTable'),
            resultsTableHead: document.getElementById('resultsTableHead'),
            resultsTableBody: document.getElementById('resultsTableBody'),
            
            // Transfer
            transferPanel: document.getElementById('transferPanel'),
            transferFlag: document.getElementById('transferFlag'),
            transferComment: document.getElementById('transferComment'),
            btnAddToQueue: document.getElementById('btnAddToQueue'),
            btnTransferDirect: document.getElementById('btnTransferDirect'),
            btnSelectAll: document.getElementById('btnSelectAll'),
            btnDeselectAll: document.getElementById('btnDeselectAll'),
            transferStatus: document.getElementById('transferStatus'),
            
            // Queue
            queuePanel: document.getElementById('queuePanel'),
            queueList: document.getElementById('queueList'),
            queueCount: document.getElementById('queueCount'),
            btnClearQueue: document.getElementById('btnClearQueue'),
            btnExportPS: document.getElementById('btnExportPS'),
            btnPreviewTransfer: document.getElementById('btnPreviewTransfer'),
            btnNewMonth: document.getElementById('btnNewMonth'),
            btnDataExplorer: document.getElementById('btnDataExplorer'),
            
            // Diff Preview Modal
            diffPreviewModal: document.getElementById('diffPreviewModal'),
            
            // Template
            btnLoadTemplate: document.getElementById('btnLoadTemplate'),
            btnCreateTemplate: document.getElementById('btnCreateTemplate'),
            templateInput: document.getElementById('templateInput'),
            templateInfo: document.getElementById('templateInfo'),
            
            // Create Template Modal
            createTemplateModal: document.getElementById('createTemplateModal'),
            createTemplateSourceName: document.getElementById('createTemplateSourceName'),
            createTemplateSheetList: document.getElementById('createTemplateSheetList'),
            
            // New Month Modal
            newMonthModal: document.getElementById('newMonthModal'),
            newMonthFilename: document.getElementById('newMonthFilename'),
            newMonthTemplateName: document.getElementById('newMonthTemplateName'),
            
            // History
            historyList: document.getElementById('historyList'),
            
            // Config
            btnExportConfig: document.getElementById('btnExportConfig'),
            btnImportConfig: document.getElementById('btnImportConfig'),
            configInput: document.getElementById('configInput'),
            
            // Data Explorer
            dataExplorerModal: document.getElementById('dataExplorerModal'),
            explorerFileName: document.getElementById('explorerFileName'),
            explorerSheetSelect: document.getElementById('explorerSheetSelect'),
            explorerSearch: document.getElementById('explorerSearch'),
            explorerResultCount: document.getElementById('explorerResultCount'),
            explorerStatus: document.getElementById('explorerStatus'),
            explorerTableHead: document.getElementById('explorerTableHead'),
            explorerTableBody: document.getElementById('explorerTableBody'),
            btnExplorerExport: document.getElementById('btnExplorerExport'),
            btnExplorerOpenFile: document.getElementById('btnExplorerOpenFile'),
            btnExplorerSearch: document.getElementById('btnExplorerSearch'),
            btnToggleExcel: document.getElementById('btnToggleExcel'),
            btnToggleExcelInteractive: document.getElementById('btnToggleExcelInteractive'),
            btnExplorerFullscreen: document.getElementById('btnExplorerFullscreen'),
            btnCloseExplorerX: document.getElementById('btnCloseExplorerX'),
            btnCloseExplorerFooter: document.getElementById('btnCloseExplorerFooter'),
            btnToggleColumns: document.getElementById('btnToggleColumns'),
            btnShowAllColumns: document.getElementById('btnShowAllColumns'),
            btnHideAllColumns: document.getElementById('btnHideAllColumns'),
            btnAddExplorerFilter: document.getElementById('btnAddExplorerFilter'),
            btnClearExplorerFilters: document.getElementById('btnClearExplorerFilters'),
            explorerDropZone: document.getElementById('explorerDropZone'),
            explorerPageSize: document.getElementById('explorerPageSize'),
            btnExplorerFirstPage: document.getElementById('btnExplorerFirstPage'),
            btnExplorerPrevPage: document.getElementById('btnExplorerPrevPage'),
            btnExplorerNextPage: document.getElementById('btnExplorerNextPage'),
            btnExplorerLastPage: document.getElementById('btnExplorerLastPage'),
            
            // Data Join Modal
            dataJoinModal: document.getElementById('dataJoinModal'),
            btnDataJoin: document.getElementById('btnDataJoin'),
            btnCloseDataJoin: document.getElementById('btnCloseDataJoin'),
            btnJoinSelectFile: document.getElementById('btnJoinSelectFile'),
            joinSourceFileName: document.getElementById('joinSourceFileName'),
            joinSourceSheet: document.getElementById('joinSourceSheet'),
            joinTargetKeyColumn: document.getElementById('joinTargetKeyColumn'),
            joinSourceKeyColumn: document.getElementById('joinSourceKeyColumn'),
            joinColumnsContainer: document.getElementById('joinColumnsContainer'),
            joinInsertAfterKey: document.getElementById('joinInsertAfterKey'),
            joinMarkNotFound: document.getElementById('joinMarkNotFound'),
            joinPreviewContainer: document.getElementById('joinPreviewContainer'),
            joinStatTargetRows: document.getElementById('joinStatTargetRows'),
            joinStatSourceRows: document.getElementById('joinStatSourceRows'),
            joinStatMatches: document.getElementById('joinStatMatches'),
            joinStatNoMatch: document.getElementById('joinStatNoMatch'),
            btnCancelDataJoin: document.getElementById('btnCancelDataJoin'),
            btnPreviewDataJoin: document.getElementById('btnPreviewDataJoin'),
            btnExecuteDataJoin: document.getElementById('btnExecuteDataJoin'),
            
            // Help
            btnHelp: document.getElementById('btnHelp'),
            helpModal: document.getElementById('helpModal')
        };

        // ==================== Utility Functions ====================
        function debounce(func, wait) {
            let timeout;
            return function executedFunction(...args) {
                const later = () => {
                    clearTimeout(timeout);
                    func(...args);
                };
                clearTimeout(timeout);
                timeout = setTimeout(later, wait);
            };
        }

        // ==================== Local Storage ====================
        const STORAGE_KEY = 'mvmcVertragslistenConfig';
        const LAST_EXPORT_KEY = 'mvmcVertragslistenLastExport';
        const DB_NAME = 'MVMCVertragsListenDB';
        const DB_VERSION = 1;
        let db = null;
        
        // ==================== Helper Functions ====================
        function showStatus(element, message, type = 'info') {
            if (!element) return;
            element.innerHTML = `<div class="status ${type}">${message}</div>`;
            // Auto-clear after 10 seconds for success messages
            if (type === 'success') {
                setTimeout(() => {
                    if (element.querySelector('.status.success')) {
                        element.innerHTML = '';
                    }
                }, 10000);
            }
        }
        
        // Im Electron-Modus werden Datei-Downloads via electronAPI gehandhabt
        
        function updateWorkbook() {
            // Im Electron-Modus nicht benötigt - Änderungen gehen direkt in die Datei
            return;
        }
        
        // Formatiert Datum und Uhrzeit für History-Einträge
        function formatHistoryDateTime() {
            const now = new Date();
            const date = now.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: '2-digit' });
            const time = now.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
            return `${date} ${time}`;
        }
        
        function updateHistoryDisplay() {
            if (!elements.historyList) return;
            
            if (state.history.length === 0) {
                elements.historyList.innerHTML = `

                    <div style="color: var(--text-muted); font-size: 13px; text-align: center; padding: 20px;">
                        Noch keine Übertragungen
                    </div>`;
                return;
            }
            
            let html = '';
            state.history.forEach(entry => {
                html += `
                    <div class="history-item">
                        <span><strong>[${entry.flag}]</strong> ${escapeHtml(entry.preview || entry.searchValue || '')}</span>
                        <span class="history-time">${entry.time}</span>
                    </div>`;
            });
            elements.historyList.innerHTML = html;
        }
        
        // IndexedDB für große Dateien
        function initDB() {
            return new Promise((resolve, reject) => {
                const request = indexedDB.open(DB_NAME, DB_VERSION);
                
                request.onerror = () => reject(request.error);
                request.onsuccess = () => {
                    db = request.result;
                    resolve(db);
                };
                
                request.onupgradeneeded = (event) => {
                    const database = event.target.result;
                    if (!database.objectStoreNames.contains('files')) {
                        database.createObjectStore('files', { keyPath: 'id' });
                    }
                    if (!database.objectStoreNames.contains('config')) {
                        database.createObjectStore('config', { keyPath: 'id' });
                    }
                };
            });
        }
        
        function saveToIndexedDB(storeName, id, data) {
            return new Promise((resolve, reject) => {
                if (!db) { reject('DB not initialized'); return; }
                const transaction = db.transaction(storeName, 'readwrite');
                const store = transaction.objectStore(storeName);
                const request = store.put({ id, data, timestamp: Date.now() });
                request.onsuccess = () => resolve();
                request.onerror = () => reject(request.error);
            });
        }
        
        function loadFromIndexedDB(storeName, id) {
            return new Promise((resolve, reject) => {
                if (!db) { reject('DB not initialized'); return; }
                const transaction = db.transaction(storeName, 'readonly');
                const store = transaction.objectStore(storeName);
                const request = store.get(id);
                request.onsuccess = () => resolve(request.result?.data);
                request.onerror = () => reject(request.error);
            });
        }
        
        function saveConfig() {
            const config = {
                file1SheetName: state.file1.selectedSheet,
                file2SheetName: state.file2.selectedSheet,
                mapping: state.mapping,
                history: state.history.slice(-100),  // Keep last 100 entries
                // Extra-Spalten Konfiguration
                extraColumns: {
                    enableFlag: isFlagEnabled(),
                    enableComment: isCommentEnabled(),
                    flagColumn: getFlagColumn(),
                    flagValues: getFlagValues().join(','),
                    commentColumn: getCommentColumn(),
                    commentPlaceholder: getCommentPlaceholder()
                }
            };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
        }
        
        function loadConfig() {
            try {
                const saved = localStorage.getItem(STORAGE_KEY);
                if (saved) {
                    const config = JSON.parse(saved);
                    state.mapping = config.mapping || state.mapping;
                    // Sicherstellen dass dropdownColumns existiert (für Abwärtskompatibilität)
                    if (!state.mapping.dropdownColumns) {
                        state.mapping.dropdownColumns = [];
                    }
                    if (state.mapping.changeRequestPattern === undefined) {
                        state.mapping.changeRequestPattern = '*_Change_Request*.*';
                    }
                    state.history = config.history || [];
                    
                    // Extra-Spalten Konfiguration laden
                    if (config.extraColumns) {
                        const ec = config.extraColumns;
                        if (ec.enableFlag !== undefined) {
                            document.getElementById('enableFlagColumn').checked = ec.enableFlag;
                            localStorage.setItem('excelSyncEnableFlag', String(ec.enableFlag));
                        }
                        if (ec.enableComment !== undefined) {
                            document.getElementById('enableCommentColumn').checked = ec.enableComment;
                            localStorage.setItem('excelSyncEnableComment', String(ec.enableComment));
                        }
                        if (ec.flagValues) {
                            document.getElementById('flagValues').value = ec.flagValues;
                            localStorage.setItem('excelSyncFlagValues', ec.flagValues);
                        }
                        if (ec.commentPlaceholder) {
                            document.getElementById('commentPlaceholder').value = ec.commentPlaceholder;
                            localStorage.setItem('excelSyncCommentPlaceholder', ec.commentPlaceholder);
                        }
                        // UI aktualisieren
                        updateFlagDropdownOptions();
                        updateCommentPlaceholders();
                        updateFlagCommentVisibility();
                        updateColumnDisplays();
                    }
                    
                    updateHistoryDisplay();
                    return config;
                }
            } catch (e) {
                console.error('Error loading config:', e);
            }
            return null;
        }
        
        async function exportConfig() {
            const config = {
                file1SheetName: state.file1.selectedSheet,
                file2SheetName: state.file2.selectedSheet,
                mapping: state.mapping,
                exportDate: new Date().toISOString(),
                // Extra-Spalten Konfiguration
                extraColumns: {
                    enableFlag: isFlagEnabled(),
                    enableComment: isCommentEnabled(),
                    flagColumn: getFlagColumn(),
                    flagValues: getFlagValues().join(','),
                    commentColumn: getCommentColumn(),
                    commentPlaceholder: getCommentPlaceholder()
                }
            };
            
            // Im Electron-Modus: Nur Mapping und Dateipfade speichern
            if (state.file1.filePath) {
                config.file1Path = state.file1.filePath;
                config.file1Name = state.file1.name;
            }
            if (state.file2.filePath) {
                config.file2Path = state.file2.filePath;
                config.file2Name = state.file2.name;
            }
            // Template-Pfad speichern
            if (state.template.filePath) {
                config.templatePath = state.template.filePath;
                config.templateName = state.template.name;
            }
            
            // Template speichern (falls vorhanden) - für Browser-Modus als Base64
            if (state.template.data) {
                config.templateData = state.template.data;
                config.templateName = state.template.name;
            }
            
            // In IndexedDB speichern (für große Dateien)
            if (db) {
                saveToIndexedDB('config', 'lastExport', config)
                    .then(() => console.log('Konfig in IndexedDB gespeichert'))
                    .catch(e => console.error('IndexedDB Fehler:', e));
            }
            
            // Kleine Konfig (ohne Dateien) auch in LocalStorage für Fallback
            const configSmall = { ...config };
            delete configSmall.file1Data;
            delete configSmall.file2Data;
            try {
                localStorage.setItem(LAST_EXPORT_KEY, JSON.stringify(configSmall));
            } catch (e) {
                console.warn('LocalStorage Fehler:', e);
            }
            
            const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
            
            // Electron-Modus: Verwende Electron-API zum Speichern
            try {
                const savePath = await window.electronAPI.saveFileDialog({
                    title: 'Konfiguration speichern',
                    defaultPath: getWorkingDirectoryPath() ? (getWorkingDirectoryPath() + '/config.json') : 'config.json',
                    filters: [{ name: 'JSON Dateien', extensions: ['json'] }]
                });
                if (savePath) {
                    const result = await window.electronAPI.saveConfig(savePath, config);
                    if (result && result.success === false) {
                        showStatus(elements.transferStatus, `⚠️ ${result.error}`, 'error');
                        alert(`❌ Speichern fehlgeschlagen:\n\n${result.error}`);
                    } else {
                        // Zeige Info über Computer-Abschnitt
                        let statusMsg = `✓ config.json gespeichert: ${savePath}`;
                        if (result.savedToSection) {
                            if (result.convertedToUserProfiles) {
                                statusMsg = `✓ Config für Benutzer „${result.savedToSection}“ gespeichert (Datei wurde für Mehrbenutzer konvertiert)`;
                            } else {
                                statusMsg = `✓ Config für Benutzer „${result.savedToSection}“ gespeichert`;
                            }
                        }
                        showStatus(elements.transferStatus, statusMsg, 'success');
                    }
                }
            } catch (e) {
                console.error('Fehler beim Speichern:', e);
                showStatus(elements.transferStatus, `Fehler: ${e.message}`, 'error');
                alert(`❌ Speichern fehlgeschlagen:\n\n${e.message}`);
            }
        }
        
        function importConfig(file) {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const config = JSON.parse(e.target.result);
                    state.mapping = config.mapping || state.mapping;
                    
                    // Extra-Spalten Konfiguration importieren
                    if (config.extraColumns) {
                        const ec = config.extraColumns;
                        if (ec.enableFlag !== undefined) {
                            document.getElementById('enableFlagColumn').checked = ec.enableFlag;
                            localStorage.setItem('excelSyncEnableFlag', String(ec.enableFlag));
                        }
                        if (ec.enableComment !== undefined) {
                            document.getElementById('enableCommentColumn').checked = ec.enableComment;
                            localStorage.setItem('excelSyncEnableComment', String(ec.enableComment));
                        }
                        if (ec.flagValues) {
                            document.getElementById('flagValues').value = ec.flagValues;
                            localStorage.setItem('excelSyncFlagValues', ec.flagValues);
                        }
                        if (ec.commentPlaceholder) {
                            document.getElementById('commentPlaceholder').value = ec.commentPlaceholder;
                            localStorage.setItem('excelSyncCommentPlaceholder', ec.commentPlaceholder);
                        }
                        // UI aktualisieren
                        updateFlagDropdownOptions();
                        updateCommentPlaceholders();
                        updateFlagCommentVisibility();
                    }
                    
                    // Speichere in IndexedDB (kann große Dateien speichern)
                    if (db) {
                        try {
                            await saveToIndexedDB('config', 'lastExport', config);
                            console.log('Importierte Konfig in IndexedDB gespeichert');
                        } catch (dbErr) {
                            console.warn('IndexedDB Fehler:', dbErr);
                        }
                    }
                    
                    // Fallback: Speichere auch in LocalStorage (ohne Dateiinhalt wenn zu groß)
                    try {
                        localStorage.setItem(LAST_EXPORT_KEY, e.target.result);
                    } catch (lsErr) {
                        // LocalStorage zu klein - speichere ohne Dateien
                        const smallConfig = { ...config };
                        delete smallConfig.file1Data;
                        delete smallConfig.file2Data;
                        localStorage.setItem(LAST_EXPORT_KEY, JSON.stringify(smallConfig));
                        console.log('LocalStorage zu klein für Dateien, nur Mapping gespeichert');
                    }
                    
                    // Lade eingebettete Excel-Dateien
                    if (config.file1Data) {
                        await loadWorkbookFromBase64(config.file1Data, config.file1Name, 1, config.file1SheetName);
                    } else if (config.file1SheetName) {
                        state.file1.pendingSheet = config.file1SheetName;
                        elements.file1Info.textContent = `⏳ Arbeitsblatt: ${config.file1SheetName}`;
                        elements.file1Info.style.color = 'var(--warning)';
                    }
                    
                    if (config.file2Data) {
                        await loadWorkbookFromBase64(config.file2Data, config.file2Name, 2, config.file2SheetName);
                    } else if (config.file2SheetName) {
                        state.file2.pendingSheet = config.file2SheetName;
                        elements.file2Info.textContent = `⏳ Arbeitsblatt: ${config.file2SheetName}`;
                        elements.file2Info.style.color = 'var(--warning)';
                    }
                    
                    // Template laden (falls vorhanden)
                    if (config.templateData) {
                        state.template.data = config.templateData;
                        state.template.name = config.templateName;
                        elements.templateInfo.textContent = config.templateName;
                        elements.btnNewMonth.disabled = false;
                    }
                    
                    updateMappingInfo();
                    updateMappingPreview();
                    saveConfig();
                    
                    // Zeige geladene Konfig-Details
                    const date = config.exportDate ? new Date(config.exportDate).toLocaleString('de-DE') : 'unbekannt';
                    const mappingCount = state.mapping.sourceColumns?.length || 0;
                    const filesIncluded = (config.file1Data ? 1 : 0) + (config.file2Data ? 1 : 0);

                    showStatus(elements.transferStatus, 
                        `✓ Konfiguration importiert! (${mappingCount} Spalten, ${filesIncluded} Datei(en), vom ${date})`, 'success');

                    elements.mappingInfo.textContent = `${mappingCount} Spalte(n) konfiguriert`;
                    
                } catch (err) {
                    showStatus(elements.transferStatus, 'Fehler beim Importieren: ' + err.message, 'error');
                }
            };
            reader.readAsText(file);
        }
        
        // ==================== File Loading (Electron-Modus) ====================
        // Dateien werden über electronAPI.openFileDialog() geladen
        // Die Browser-Funktionen loadFile(), processFile(), loadWorkbookFromBase64() wurden entfernt
        // selectSheet1/2 im Browser-Modus wurden entfernt (XLSX-basiert)
        
        function checkReadyState() {
            const hasFile1 = state.file1.filePath;
            const hasFile2 = state.file2.filePath;
            const bothLoaded = hasFile1 && hasFile2;
            
            // Mapping-Button nur deaktivieren wenn KEINE Konfig geladen AND kein Dateien
            const hasMapping = state.mapping.sourceColumns && state.mapping.sourceColumns.length > 0;
            elements.btnConfigMapping.disabled = !bothLoaded && !hasMapping;
            
            elements.searchInput.disabled = !state.file1.selectedSheet;
            elements.btnSearch.disabled = !state.file1.selectedSheet;
            
            // "Neue Zeile" Button aktivieren wenn Mapping vorhanden und Datei 2 geladen
            elements.btnNewRow.disabled = !hasMapping || !hasFile2;
            
            // Datenexplorer ist immer aktiviert (kann eigene Dateien öffnen)
            elements.btnDataExplorer.disabled = false;
            
            updateMappingInfo();
            updateMappingPreview();
        }
        
        function updateMappingPreview() {
            const preview = document.getElementById('mappingPreview');
            if (!preview) return;
            
            if (!state.mapping.sourceColumns || state.mapping.sourceColumns.length === 0) {
                preview.innerHTML = '';
                return;
            }
            
            let html = '<strong>Aktuelle Konfig:</strong><br>';
            html += `Start-Spalte: ${getColumnLetter(getDataStartColumn())}<br>`;
            html += `Duplikat-Check: Spalte ${state.mapping.duplicateCheckColumn + 1}<br>`;
            html += `Request-Muster: ${state.mapping.changeRequestPattern || '(keins)'}<br>`;
            html += `Spalten: ${state.mapping.sourceColumns.length}`;
            
            preview.innerHTML = html;
        }
        
        function updateMappingInfo() {
            const hasFile1 = state.file1.filePath;
            const hasFile2 = state.file2.filePath;
            
            if (!hasFile1 || !hasFile2) {
                if (state.mapping.sourceColumns && state.mapping.sourceColumns.length > 0) {
                    elements.mappingInfo.textContent = `${state.mapping.sourceColumns.length} Spalte(n) konfiguriert (Dateien laden)`;
                    elements.mappingInfo.style.color = 'var(--warning)';
                } else {
                    elements.mappingInfo.textContent = 'Laden Sie beide Dateien';
                    elements.mappingInfo.style.color = '';
                }
                return;
            }
            
            elements.mappingInfo.style.color = '';
            if (state.mapping.sourceColumns.length === 0) {
                elements.mappingInfo.textContent = 'Klicken Sie auf "Spalten konfigurieren"';
            } else {
                elements.mappingInfo.textContent = `${state.mapping.sourceColumns.length} Spalte(n) konfiguriert ✓`;
            }
        }
        
        function openMappingModal() {
            elements.mappingModal.classList.remove('hidden');
            renderMappingList();
        }
        
        function closeMappingModal() {
            elements.mappingModal.classList.add('hidden');
        }
        
        function renderMappingList() {
            elements.mappingList.innerHTML = '';
            
            // If no mappings yet, add all columns by default (skip first column - usually row numbers)
            if (state.mapping.sourceColumns.length === 0 && state.file1.headers.length > 1) {
                state.mapping.sourceColumns = state.file1.headers.slice(1).map((_, i) => i + 1);
            } else if (state.mapping.sourceColumns.length === 0 && state.file1.headers.length === 1) {
                state.mapping.sourceColumns = [0];
            }
            
            // Populate duplicate check column dropdown
            const dupSelect = document.getElementById('duplicateCheckColumn');
            dupSelect.innerHTML = '';
            state.file1.headers.forEach((header, idx) => {
                const option = document.createElement('option');
                option.value = idx;
                option.textContent = `${getColumnLetter(idx + 1)}: ${header || '(leer)'}`;
                if (idx === state.mapping.duplicateCheckColumn) option.selected = true;
                dupSelect.appendChild(option);
            });
            
            // Change-Request Pattern Feld befüllen
            const patternInput = document.getElementById('changeRequestPattern');
            if (patternInput) {
                patternInput.value = state.mapping.changeRequestPattern || '*_Change_Request*.*';
            }
            
            state.mapping.sourceColumns.forEach((colIndex, i) => {
                const item = document.createElement('div');
                item.className = 'mapping-item';
                
                const select = document.createElement('select');
                state.file1.headers.forEach((header, idx) => {
                    const option = document.createElement('option');
                    option.value = idx;
                    option.textContent = `${getColumnLetter(idx + 1)}: ${header || '(leer)'}`;
                    if (idx === colIndex) option.selected = true;
                    select.appendChild(option);
                });
                select.addEventListener('change', () => {
                    state.mapping.sourceColumns[i] = parseInt(select.value);
                });
                
                const arrow = document.createElement('span');
                arrow.className = 'mapping-arrow';
                arrow.textContent = '→';
                
                const target = document.createElement('span');
                target.className = 'mapping-target';
                const targetCol = getDataStartColumn() + i;
                target.textContent = `Spalte ${getColumnLetter(targetCol)} in Datei 2`;
                
                const removeBtn = document.createElement('button');
                removeBtn.className = 'mapping-remove';
                removeBtn.textContent = '✕';
                removeBtn.onclick = () => {
                    state.mapping.sourceColumns.splice(i, 1);
                    renderMappingList();
                };
                
                item.appendChild(select);
                item.appendChild(arrow);
                item.appendChild(target);
                item.appendChild(removeBtn);
                elements.mappingList.appendChild(item);
            });
            
            // Update column displays
            updateColumnDisplays();
            
            // Dropdown-Spalten Liste aktualisieren
            renderDropdownColumnsList();
        }
        
        function renderDropdownColumnsList() {
            const container = document.getElementById('dropdownColumnsList');
            if (!container) return;
            
            if (!state.mapping.sourceColumns || state.mapping.sourceColumns.length === 0) {
                container.innerHTML = '<div style="color: var(--text-muted); text-align: center; font-size: 13px;">Laden Sie erst Dateien und konfigurieren Sie Spalten</div>';
                return;
            }
            
            if (!state.file1.data || state.file1.data.length === 0) {
                container.innerHTML = '<div style="color: var(--text-muted); text-align: center; font-size: 13px;">Keine Daten in Quelldatei geladen</div>';
                return;
            }
            
            let html = '';
            state.mapping.sourceColumns.forEach((colIndex) => {
                const headerName = state.file1.headers[colIndex] || `Spalte ${colIndex + 1}`;
                
                // Eindeutige Werte zählen (case-insensitive, trimmed)
                const valueSet = new Set();
                state.file1.data.forEach(row => {
                    const val = row[colIndex];
                    if (val !== null && val !== undefined && val !== '') {
                        const strVal = String(val).trim().toLowerCase();
                        if (strVal) valueSet.add(strVal);
                    }
                });
                const valueCount = valueSet.size;
                
                // Nur Spalten mit sinnvoller Anzahl (1-50) anzeigen
                const isChecked = state.mapping.dropdownColumns && state.mapping.dropdownColumns.includes(colIndex);
                const recommendation = valueCount > 0 && valueCount <= 20 ? '✓' : (valueCount > 50 ? '⚠️ viele Werte' : '');
                
                html += `
                    <div class="dropdown-column-item">
                        <input type="checkbox" class="green-checkbox dropdown-col-checkbox" 
                               id="dropdownCol_${colIndex}" 
                               data-col-index="${colIndex}" 
                               ${isChecked ? 'checked' : ''}>
                        <label for="dropdownCol_${colIndex}">${headerName}</label>
                        <span class="value-count">${valueCount} Werte ${recommendation}</span>
                    </div>`;
            });
            
            container.innerHTML = html;
            
            // Event-Handler für Checkboxen
            container.querySelectorAll('.dropdown-col-checkbox').forEach(cb => {
                cb.addEventListener('change', (e) => {
                    const colIndex = parseInt(e.target.dataset.colIndex);
                    if (!state.mapping.dropdownColumns) {
                        state.mapping.dropdownColumns = [];
                    }
                    
                    if (e.target.checked) {
                        if (!state.mapping.dropdownColumns.includes(colIndex)) {
                            state.mapping.dropdownColumns.push(colIndex);
                        }
                    } else {
                        state.mapping.dropdownColumns = state.mapping.dropdownColumns.filter(c => c !== colIndex);
                    }
                });
            });
        }
        
        function addMappingColumn() {
            const nextCol = state.mapping.sourceColumns.length;
            if (nextCol < state.file1.headers.length) {
                state.mapping.sourceColumns.push(nextCol);
                renderMappingList();
            }
        }
        
        function saveMapping() {
            state.mapping.duplicateCheckColumn = parseInt(document.getElementById('duplicateCheckColumn').value);
            const patternInput = document.getElementById('changeRequestPattern');
            const newPattern = (patternInput.value || '').trim();
            if (newPattern !== state.mapping.changeRequestPattern) {
                state.mapping.changeRequestPattern = newPattern;
                invalidateChangeRequestCache();
            }
            saveConfig();
            updateMappingInfo();
            closeMappingModal();
        }
        
        function getColumnLetter(num) {
            let result = '';
            while (num > 0) {
                num--;
                result = String.fromCharCode(65 + (num % 26)) + result;
                num = Math.floor(num / 26);
            }
            return result;
        }
        
        // ==================== Search Functions ====================
        function wildcardToRegex(pattern) {
            let escaped = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&');
            escaped = escaped.replace(/\*/g, '.*');
            escaped = escaped.replace(/\?/g, '.');
            return new RegExp('^' + escaped + '$', 'i');
        }
        
        /**
         * Prüft ob ein Text einem Suchterm entspricht (mit Platzhalter-Unterstützung)
         */
        function matchesTerm(text, term, hasWildcards) {
            if (!text) return false;
            const str = String(text);
            if (hasWildcards) {
                return wildcardToRegex(term).test(str);
            }
            return str.toLowerCase().includes(term.toLowerCase());
        }
        
        /**
         * Parst eine Suchanfrage mit AND/OR Operatoren
         * Beispiele: "Projekt AND 2025", "Alpha OR Beta", "Item AND (2024 OR 2025)"
         */
        function parseSearchQuery(query) {
            const trimmed = query.trim();
            
            // Prüfen ob AND oder OR Operatoren vorhanden sind
            const hasAnd = / AND /i.test(trimmed);
            const hasOr = / OR /i.test(trimmed);
            
            if (!hasAnd && !hasOr) {
                // Einfache Suche
                return { type: 'simple', term: trimmed };
            }
            
            if (hasAnd && !hasOr) {
                // Nur AND
                const terms = trimmed.split(/ AND /i).map(t => t.trim()).filter(t => t);
                return { type: 'and', terms };
            }
            
            if (hasOr && !hasAnd) {
                // Nur OR
                const terms = trimmed.split(/ OR /i).map(t => t.trim()).filter(t => t);
                return { type: 'or', terms };
            }
            
            // Gemischt: OR hat niedrigere Priorität, AND wird zuerst ausgewertet
            // "A AND B OR C AND D" → (A AND B) OR (C AND D)
            const orParts = trimmed.split(/ OR /i).map(part => {
                const andTerms = part.split(/ AND /i).map(t => t.trim()).filter(t => t);
                if (andTerms.length === 1) {
                    return { type: 'simple', term: andTerms[0] };
                }
                return { type: 'and', terms: andTerms };
            });
            
            return { type: 'or', parts: orParts };
        }
        
        /**
         * Prüft ob eine Zeile der geparsten Suchanfrage entspricht
         */
        function rowMatchesQuery(row, parsed) {
            const rowStr = row.join(' ').toLowerCase();
            
            function termHasWildcards(term) {
                return term.includes('*') || term.includes('?');
            }
            
            function rowContainsTerm(term) {
                const hasWc = termHasWildcards(term);
                for (const cell of row) {
                    if (matchesTerm(cell, term, hasWc)) return true;
                }
                return false;
            }
            
            if (parsed.type === 'simple') {
                return rowContainsTerm(parsed.term);
            }
            
            if (parsed.type === 'and') {
                // Alle Terme müssen in der Zeile vorkommen
                return parsed.terms.every(term => rowContainsTerm(term));
            }
            
            if (parsed.type === 'or') {
                if (parsed.terms) {
                    // Einfache OR-Verknüpfung
                    return parsed.terms.some(term => rowContainsTerm(term));
                }
                if (parsed.parts) {
                    // Gemischte Anfrage mit AND-Gruppen
                    return parsed.parts.some(part => rowMatchesQuery(row, part));
                }
            }
            
            return false;
        }
        
        function search() {
            const query = elements.searchInput.value.trim();
            if (!query || !state.file1.data.length) return;
            
            const hasWildcards = query.includes('*') || query.includes('?');
            const hasOperators = / (AND|OR) /i.test(query);
            state.searchResults = [];
            
            if (hasOperators) {
                // Erweiterte Suche mit AND/OR
                const parsed = parseSearchQuery(query);
                state.file1.data.forEach((row, rowIndex) => {
                    if (rowMatchesQuery(row, parsed)) {
                        state.searchResults.push({ rowIndex: rowIndex, data: row });
                    }
                });
            } else if (hasWildcards) {
                const regex = wildcardToRegex(query);
                state.file1.data.forEach((row, rowIndex) => {
                    for (let col of row) {
                        if (col && regex.test(String(col))) {
                            state.searchResults.push({ rowIndex: rowIndex, data: row });
                            break;
                        }
                    }
                });
            } else {
                const lowerQuery = query.toLowerCase();
                state.file1.data.forEach((row, rowIndex) => {
                    for (let col of row) {
                        if (col && String(col).toLowerCase().includes(lowerQuery)) {
                            state.searchResults.push({ rowIndex: rowIndex, data: row });
                            break;
                        }
                    }
                });
            }
            
            // Such-Historie aktualisieren
            addToSearchHistory(query, state.searchResults.length);
            hideSearchHistoryDropdown();
            
            displaySearchResults(query, hasWildcards || hasOperators);
        }
        
        function displaySearchResults(query, hasWildcards = false) {
            state.selectedRows = [];
            state.searchPagination.currentPage = 1; // Bei neuer Suche zur ersten Seite
            
            if (state.searchResults.length === 0) {
                elements.searchResultsInfo.innerHTML = `Keine Treffer für "<strong>${escapeHtml(query)}</strong>"`;
                elements.emptyState.style.display = 'flex';
                elements.resultsTableContainer.style.display = 'none';
                elements.transferPanel.classList.add('hidden');
                document.getElementById('searchPagination').style.display = 'none';
                return;
            }
            
            renderSearchResultsPage(query, hasWildcards);
        }
        
        // ==================== Event Delegation für Search Results ====================
        let _searchResultsDelegationSetup = false;
        
        function setupSearchResultsDelegation() {
            if (_searchResultsDelegationSetup) return;
            _searchResultsDelegationSetup = true;
            
            const tbody = elements.resultsTableBody;
            
            tbody.addEventListener('blur', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                
                const rowIndex = parseInt(td.dataset.row);
                const colIndex = parseInt(td.dataset.col);
                const original = td.dataset.original;
                const lastValue = td.dataset.lastValue;
                const current = td.textContent;
                
                if (lastValue !== current) {
                    pushSearchUndo({
                        rowIndex, colIndex,
                        oldValue: lastValue, newValue: current,
                        originalValue: original
                    });
                    state.searchResults[rowIndex].data[colIndex] = current;
                    td.dataset.lastValue = current;
                }
                
                td.classList.toggle('edited', original !== current);
            }, true);
            
            tbody.addEventListener('input', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                const original = td.dataset.original;
                const current = td.textContent;
                td.classList.toggle('edited', original !== current);
            });
            
            tbody.addEventListener('focus', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                const rowIndex = parseInt(td.dataset.row);
                if (!state.selectedRows.includes(rowIndex)) {
                    toggleRowSelection(rowIndex, true);
                }
            }, true);
        }

        function renderSearchResultsPage(query = '', hasWildcards = false) {
            const totalResults = state.searchResults.length;
            const pageSize = state.searchPagination.pageSize;
            const totalPages = Math.max(1, Math.ceil(totalResults / pageSize));
            
            // Sicherstellen, dass currentPage gültig ist
            if (state.searchPagination.currentPage > totalPages) {
                state.searchPagination.currentPage = totalPages;
            }
            if (state.searchPagination.currentPage < 1) {
                state.searchPagination.currentPage = 1;
            }
            
            const startIndex = (state.searchPagination.currentPage - 1) * pageSize;
            const endIndex = Math.min(startIndex + pageSize, totalResults);
            const pageResults = state.searchResults.slice(startIndex, endIndex);
            
            const wildcardInfo = hasWildcards ? ' (Platzhalter)' : '';
            if (totalResults > pageSize) {
                elements.searchResultsInfo.innerHTML = 
                    `Zeige <strong>${startIndex + 1}-${endIndex}</strong> von <strong>${totalResults}</strong> Treffern für "<strong>${escapeHtml(query)}</strong>"${wildcardInfo}`;
            } else {
                elements.searchResultsInfo.innerHTML = 
                    `<strong>${totalResults}</strong> Treffer für "<strong>${escapeHtml(query)}</strong>"${wildcardInfo}`;
            }
            
            elements.emptyState.style.display = 'none';
            elements.resultsTableContainer.style.display = 'block';
            
            let headerHtml = '<tr><th style="width: 40px;"><input type="checkbox" id="selectAllCheckbox" title="Alle auf dieser Seite auswählen"></th>';
            state.file1.headers.forEach((header, i) => {
                headerHtml += `<th>${escapeHtml(header || `Spalte ${getColumnLetter(i + 1)}`)}</th>`;
            });
            headerHtml += '</tr>';
            elements.resultsTableHead.innerHTML = headerHtml;
            
            // Nur die aktuelle Seite rendern
            let bodyHtml = '';
            pageResults.forEach((result, pageIndex) => {
                const globalIndex = startIndex + pageIndex;
                bodyHtml += `<tr data-index="${globalIndex}">`;
                bodyHtml += `<td><input type="checkbox" class="row-checkbox" data-index="${globalIndex}" onclick="event.stopPropagation()"></td>`;
                state.file1.headers.forEach((_, colIndex) => {
                    const cell = result.data[colIndex];
                    const cellStr = String(cell ?? '');
                    bodyHtml += `<td contenteditable="true" data-row="${globalIndex}" data-col="${colIndex}" data-original="${escapeHtml(cellStr)}" onclick="event.stopPropagation()">${escapeHtml(cellStr)}</td>`;
                });
                bodyHtml += '</tr>';
            });
            elements.resultsTableBody.innerHTML = bodyHtml;
            
            // lastValue initialisieren
            document.querySelectorAll('#resultsTableBody td[contenteditable]').forEach(td => {
                td.dataset.lastValue = td.textContent;
            });
            
            // Event Delegation für Search Results (einmalig)
            setupSearchResultsDelegation();
            
            document.getElementById('selectAllCheckbox').addEventListener('change', (e) => {
                selectAllRows(e.target.checked);
            });
            
            document.querySelectorAll('.row-checkbox').forEach(cb => {
                cb.addEventListener('change', (e) => {
                    toggleRowSelection(parseInt(e.target.dataset.index), e.target.checked);
                });
            });
            
            // Pagination UI aktualisieren
            updateSearchPagination(totalPages);
            
            // Erste Zeile auf der aktuellen Seite auswählen
            if (pageResults.length > 0) {
                toggleRowSelection(startIndex, true);
            }
            elements.transferPanel.classList.remove('hidden');
        }
        
        function updateSearchPagination(totalPages) {
            const paginationEl = document.getElementById('searchPagination');
            const pageInfoEl = document.getElementById('searchPageInfo');
            const firstBtn = document.getElementById('btnSearchFirstPage');
            const prevBtn = document.getElementById('btnSearchPrevPage');
            const nextBtn = document.getElementById('btnSearchNextPage');
            const lastBtn = document.getElementById('btnSearchLastPage');
            
            // Pagination nur anzeigen wenn mehr als eine Seite
            if (state.searchResults.length > state.searchPagination.pageSize) {
                paginationEl.style.display = 'flex';
                pageInfoEl.textContent = `${t('pageLabel')} ${state.searchPagination.currentPage} ${t('ofLabel')} ${totalPages}`;
                
                // Buttons aktivieren/deaktivieren
                firstBtn.disabled = state.searchPagination.currentPage === 1;
                prevBtn.disabled = state.searchPagination.currentPage === 1;
                nextBtn.disabled = state.searchPagination.currentPage === totalPages;
                lastBtn.disabled = state.searchPagination.currentPage === totalPages;
            } else {
                paginationEl.style.display = 'none';
            }
        }
        
        function searchGoToPage(page) {
            const totalPages = Math.ceil(state.searchResults.length / state.searchPagination.pageSize);
            state.searchPagination.currentPage = Math.max(1, Math.min(page, totalPages));
            renderSearchResultsPage(elements.searchInput.value.trim());
            
            // Zum Tabellenanfang scrollen
            elements.resultsTableContainer.scrollTop = 0;
        }
        
        function searchChangePageSize(newSize) {
            state.searchPagination.pageSize = parseInt(newSize);
            state.searchPagination.currentPage = 1;
            renderSearchResultsPage(elements.searchInput.value.trim());
        }
        
        function getEditedRowData(rowIndex) {
            const cells = document.querySelectorAll(`#resultsTableBody td[data-row="${rowIndex}"]`);
            const data = [];
            cells.forEach(cell => {
                const colIndex = parseInt(cell.dataset.col);
                data[colIndex] = cell.textContent;
            });
            return data;
        }
        
        function isRowEdited(rowIndex) {
            const cells = document.querySelectorAll(`#resultsTableBody td[data-row="${rowIndex}"].edited`);
            return cells.length > 0;
        }
        
        function selectAllRows(selected) {
            const checkboxes = document.querySelectorAll('.row-checkbox');
            checkboxes.forEach(cb => {
                const index = parseInt(cb.dataset.index);
                toggleRowSelection(index, selected);
            });
        }
        
        function toggleRowSelection(index, forceState) {
            const checkbox = document.querySelector(`.row-checkbox[data-index="${index}"]`);
            const tr = document.querySelector(`#resultsTableBody tr[data-index="${index}"]`);
            if (!checkbox || !tr) return;
            
            const isSelected = forceState !== undefined ? forceState : !checkbox.checked;
            checkbox.checked = isSelected;
            tr.classList.toggle('selected', isSelected);
            
            if (isSelected) {
                if (!state.selectedRows.includes(index)) {
                    state.selectedRows.push(index);
                }
            } else {
                state.selectedRows = state.selectedRows.filter(i => i !== index);
            }
            
            document.getElementById('selectedCount').textContent = state.selectedRows.length;
            
            const allCheckbox = document.getElementById('selectAllCheckbox');
            const allCheckboxes = document.querySelectorAll('.row-checkbox');
            const checkedCount = document.querySelectorAll('.row-checkbox:checked').length;
            if (allCheckbox) {
                allCheckbox.checked = checkedCount === allCheckboxes.length;
                allCheckbox.indeterminate = checkedCount > 0 && checkedCount < allCheckboxes.length;
            }
            
            if (state.selectedRows.length > 0) {
                state.selectedRow = state.searchResults[state.selectedRows[state.selectedRows.length - 1]];
            } else {
                state.selectedRow = null;
            }
        }

        function removeFromQueue(index) {
            if (typeof index !== 'number' || index < 0 || index >= state.transferQueue.length) return;
            state.transferQueue.splice(index, 1);
            updateQueueDisplay();
        }
        
        // Sichere globale Funktionen mit Object.defineProperty (nicht überschreibbar)
        Object.defineProperty(window, 'removeFromQueue', {
            value: removeFromQueue,
            writable: false,
            configurable: false
        });
        Object.defineProperty(window, 'toggleRowSelection', {
            value: toggleRowSelection,
            writable: false,
            configurable: false
        });

        // ==================== New Row Functions ====================
        function openNewRowPanel() {
            if (!state.mapping.sourceColumns || state.mapping.sourceColumns.length === 0) {
                showStatus(elements.transferStatus, 'Bitte zuerst Spalten konfigurieren', 'error');
                return;
            }
            
            // Formular mit konfigurierten Spalten aufbauen
            let html = '';
            const comboboxData = [];
            state.mapping.sourceColumns.forEach((colIndex, i) => {
                const headerName = state.file1.headers[colIndex] || `Spalte ${colIndex + 1}`;
                const isDropdown = state.mapping.dropdownColumns && state.mapping.dropdownColumns.includes(colIndex);
                
                if (isDropdown && state.file1.data && state.file1.data.length > 0) {
                    // Eindeutige Werte aus Spalte extrahieren (case-insensitive, trimmed)
                    const valueMap = new Map(); // Für case-insensitive Deduplizierung
                    state.file1.data.forEach(row => {
                        const val = row[colIndex];
                        if (val !== null && val !== undefined && val !== '') {
                            const strVal = String(val).trim();
                            if (strVal) {
                                const lowerKey = strVal.toLowerCase();
                                if (!valueMap.has(lowerKey) || strVal.length < valueMap.get(lowerKey).length) {
                                    valueMap.set(lowerKey, strVal);
                                }
                            }
                        }
                    });
                    const uniqueValues = [...valueMap.values()].sort((a, b) => a.localeCompare(b));
                    
                    if (uniqueValues.length > 0) {
                        // Custom Dropdown-Combobox mit Scrollbar
                        const optionItemsHtml = uniqueValues.map(v => 
                            `<div class="new-row-combobox-option" data-value="${v.replace(/"/g, '&quot;')}">${v.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div>`
                        ).join('');
                        html += `
                            <div class="new-row-field">
                                <label title="${headerName}">${headerName} 📋</label>
                                <div class="new-row-combobox" id="newRowCombo_${i}">
                                    <div class="new-row-combobox-input-wrap">
                                        <input type="text" id="newRowField_${i}" data-col-index="${colIndex}" placeholder="Eingeben oder auswählen..." autocomplete="off">
                                        <button type="button" class="new-row-combobox-toggle" tabindex="-1">▼</button>
                                    </div>
                                    <div class="new-row-combobox-dropdown">
                                        ${optionItemsHtml}
                                    </div>
                                </div>
                            </div>`;
                        comboboxData.push({ index: i, values: uniqueValues });
                    } else {
                        // Fallback auf Input wenn keine Werte gefunden
                        html += `
                            <div class="new-row-field">
                                <label title="${headerName}">${headerName}</label>
                                <input type="text" id="newRowField_${i}" data-col-index="${colIndex}" placeholder="${headerName}">
                            </div>`;
                    }
                } else {
                    // Standard: Text-Input
                    html += `
                        <div class="new-row-field">
                            <label title="${headerName}">${headerName}</label>
                            <input type="text" id="newRowField_${i}" data-col-index="${colIndex}" placeholder="${headerName}">
                        </div>`;
                }
            });
            
            elements.newRowForm.innerHTML = html;
            elements.newRowPanel.classList.remove('hidden');
            
            // Custom Dropdown Combobox Setup
            comboboxData.forEach(cb => {
                const combo = document.getElementById(`newRowCombo_${cb.index}`);
                const inp = document.getElementById(`newRowField_${cb.index}`);
                if (!combo || !inp) return;
                const dropdown = combo.querySelector('.new-row-combobox-dropdown');
                const toggle = combo.querySelector('.new-row-combobox-toggle');
                const allOptions = dropdown.querySelectorAll('.new-row-combobox-option');
                let highlightIdx = -1;

                function openDropdown() {
                    combo.classList.add('open');
                    filterOptions(inp.value);
                }
                function closeDropdown() {
                    combo.classList.remove('open');
                    highlightIdx = -1;
                    allOptions.forEach(o => o.classList.remove('highlighted'));
                }
                function filterOptions(query) {
                    const q = query.toLowerCase().trim();
                    let visibleCount = 0;
                    allOptions.forEach(opt => {
                        const val = opt.dataset.value.toLowerCase();
                        const match = !q || val.includes(q);
                        opt.style.display = match ? '' : 'none';
                        opt.classList.remove('highlighted');
                        if (match) visibleCount++;
                    });
                    // Wenn keine Treffer, Hinweis anzeigen
                    let noMatch = dropdown.querySelector('.no-match');
                    if (visibleCount === 0 && q) {
                        if (!noMatch) {
                            noMatch = document.createElement('div');
                            noMatch.className = 'new-row-combobox-option no-match';
                            noMatch.textContent = 'Kein Treffer';
                            dropdown.appendChild(noMatch);
                        }
                        noMatch.style.display = '';
                    } else if (noMatch) {
                        noMatch.style.display = 'none';
                    }
                    highlightIdx = -1;
                }
                function getVisibleOptions() {
                    return [...allOptions].filter(o => o.style.display !== 'none');
                }
                function setHighlight(idx) {
                    const visible = getVisibleOptions();
                    visible.forEach(o => o.classList.remove('highlighted'));
                    if (idx >= 0 && idx < visible.length) {
                        highlightIdx = idx;
                        visible[idx].classList.add('highlighted');
                        visible[idx].scrollIntoView({ block: 'nearest' });
                    } else {
                        highlightIdx = -1;
                    }
                }

                toggle.addEventListener('mousedown', (e) => {
                    e.preventDefault(); // Kein Fokus-Verlust
                    if (combo.classList.contains('open')) {
                        closeDropdown();
                    } else {
                        openDropdown();
                        inp.focus();
                    }
                });

                inp.addEventListener('focus', () => openDropdown());
                inp.addEventListener('input', () => {
                    if (!combo.classList.contains('open')) openDropdown();
                    filterOptions(inp.value);
                });
                inp.addEventListener('keydown', (e) => {
                    const visible = getVisibleOptions();
                    if (e.key === 'ArrowDown') {
                        e.preventDefault();
                        if (!combo.classList.contains('open')) openDropdown();
                        setHighlight(Math.min((highlightIdx < 0 ? -1 : highlightIdx) + 1, visible.length - 1));
                    } else if (e.key === 'ArrowUp') {
                        e.preventDefault();
                        setHighlight(Math.max((highlightIdx < 0 ? 0 : highlightIdx) - 1, 0));
                    } else if (e.key === 'Enter') {
                        e.preventDefault();
                        if (highlightIdx >= 0 && highlightIdx < visible.length) {
                            inp.value = visible[highlightIdx].dataset.value;
                            closeDropdown();
                        }
                    } else if (e.key === 'Escape') {
                        closeDropdown();
                    } else if (e.key === 'Tab') {
                        closeDropdown();
                    }
                });

                dropdown.addEventListener('mousedown', (e) => {
                    e.preventDefault(); // Kein Fokus-Verlust
                    const opt = e.target.closest('.new-row-combobox-option:not(.no-match)');
                    if (opt) {
                        inp.value = opt.dataset.value;
                        closeDropdown();
                    }
                });

                // Außerhalb klicken → schließen
                document.addEventListener('mousedown', (e) => {
                    if (!combo.contains(e.target)) closeDropdown();
                });
            });
            
            // Fokus auf erstes Feld
            const firstField = document.getElementById('newRowField_0');
            if (firstField) firstField.focus();
        }
        
        // Select-Overlay Setup erfolgt in showNewRowPanel()
        
        function closeNewRowPanel() {
            elements.newRowPanel.classList.add('hidden');
            // Felder leeren
            elements.newRowForm.innerHTML = '';
            elements.newRowComment.value = '';
            elements.newRowFlag.value = 'A';
        }
        
        function getNewRowData() {
            // Sammle alle Werte aus den Eingabefeldern
            const data = new Array(Math.max(...state.mapping.sourceColumns) + 1).fill(''); 
            
            state.mapping.sourceColumns.forEach((colIndex, i) => {
                const field = document.getElementById(`newRowField_${i}`);
                if (field) {
                    data[colIndex] = field.value;
                }
            });
            
            return data;
        }
        
        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
        
        function escapeRegex(string) {
            return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        }
        
        // ==================== New Row Functions ====================
        async function addNewRowToQueue() {
            // Change-Request-Dateien vor der Duplikatprüfung laden. Dies ist
            // insbesondere für manuell eingegebene Seriennummern notwendig.
            await loadChangeRequestFiles();

            const data = getNewRowData();
            const flag = elements.newRowFlag.value;
            const comment = elements.newRowComment.value;
            const checkValue = data[state.mapping.duplicateCheckColumn] || '';
            
            // Prüfe ob bereits in Zieldatei (Datei 2)
            if (checkValue) {
                const duplicate = checkForDuplicate(checkValue);
                if (duplicate) {
                    const rowInfo = duplicate.firstMatch ? duplicate.firstMatch.rowIndex : '?';
                    const source = duplicate.inTarget
                        ? 'Zieldatei'
                        : `Change-Request-Datei (${duplicate.inChangeRequests.join(', ')})`;
                    const errorMsg = `⚠️ Zeile bereits in ${source} vorhanden (Zeile ${rowInfo})`;
                    showStatus(elements.transferStatus, errorMsg, 'warning');
                    showStatus(elements.newRowStatus, errorMsg, 'warning');
                    return;
                }
            }
            
            state.transferQueue.push({
                data: data,
                flag: flag,
                comment: comment,
                checkValue: checkValue,
                isManual: true
            });
            
            updateQueueDisplay();
            const successMsg = '✅ Neue Zeile zur Warteschlange hinzugefügt!';
            showStatus(elements.transferStatus, successMsg, 'success');
            showStatus(elements.newRowStatus, successMsg, 'success');
            
            // Felder leeren für nächste Eingabe (falls nicht beibehalten)
            const keepFields = document.getElementById('keepFieldsCheckbox');
            if (!keepFields || !keepFields.checked) {
                state.mapping.sourceColumns.forEach((colIndex, i) => {
                    const field = document.getElementById(`newRowField_${i}`);
                    if (field) field.value = '';
                    const sel = document.getElementById(`newRowSelect_${i}`);
                    if (sel) sel.selectedIndex = 0;
                });
                elements.newRowComment.value = '';
            }
            
            const firstField = document.getElementById('newRowField_0');
            if (firstField) firstField.focus();
        }
        
        // Leerzeile hinzufügen - unabhängig von Flag-Spalte
        function addEmptyRowToQueue() {
            const comment = elements.newRowComment?.value || '';
            
            // Leeres data-Array erstellen
            const data = new Array(Math.max(...state.mapping.sourceColumns) + 1).fill('');
            
            state.transferQueue.push({
                data: data,
                flag: 'leer',
                comment: comment,
                checkValue: '(Leerzeile)',
                isManual: true
            });
            
            updateQueueDisplay();
            const successMsg = '✅ Leerzeile zur Warteschlange hinzugefügt!';
            showStatus(elements.transferStatus, successMsg, 'success');
            showStatus(elements.newRowStatus, successMsg, 'success');
            
            // Kommentar leeren
            if (elements.newRowComment) {
                elements.newRowComment.value = '';
            }
        }

        async function transferNewRowDirect() {
            const data = getNewRowData();
            const flag = elements.newRowFlag.value;
            const comment = elements.newRowComment.value;
            
            // Prüfe ob Zieldatei verfügbar ist
            const hasTargetFile = state.file2.filePath;
            if (!hasTargetFile) {
                const errorMsg = 'Keine Zieldatei geladen';
                showStatus(elements.transferStatus, errorMsg, 'error');
                showStatus(elements.newRowStatus, errorMsg, 'error');
                return;
            }

            // Change-Request-Dateien vor der Duplikatprüfung laden. Dies ist
            // insbesondere für manuell eingegebene Seriennummern notwendig.
            await loadChangeRequestFiles();
            
            // Zur Warteschlange hinzufügen und direkt übertragen
            if (state.file2.filePath) {
                const checkValue = data[state.mapping.duplicateCheckColumn] || '';
                
                // Prüfe ob bereits in Zieldatei (Datei 2)
                if (checkValue) {
                    const duplicate = checkForDuplicate(checkValue);
                    if (duplicate) {
                        const rowInfo = duplicate.firstMatch ? duplicate.firstMatch.rowIndex : '?';
                        const source = duplicate.inTarget
                            ? 'Zieldatei'
                            : `Change-Request-Datei (${duplicate.inChangeRequests.join(', ')})`;
                        const errorMsg = `⚠️ Zeile bereits in ${source} vorhanden (Zeile ${rowInfo})`;
                        showStatus(elements.transferStatus, errorMsg, 'warning');
                        showStatus(elements.newRowStatus, errorMsg, 'warning');
                        return;
                    }
                }
                
                state.transferQueue.push({
                    data: data,
                    flag: flag,
                    comment: comment,
                    checkValue: checkValue,
                    isManual: true
                });
                
                updateQueueDisplay();
                
                try {
                    // Direkt übertragen
                    await transferQueueToExcel();
                    
                    // Felder leeren für nächste Eingabe (falls nicht beibehalten)
                    const keepFields = document.getElementById('keepFieldsCheckbox');
                    if (!keepFields || !keepFields.checked) {
                        state.mapping.sourceColumns.forEach((colIndex, i) => {
                            const field = document.getElementById(`newRowField_${i}`);
                            if (field) field.value = '';
                            const sel = document.getElementById(`newRowSelect_${i}`);
                            if (sel) sel.selectedIndex = 0;
                        });
                        elements.newRowComment.value = '';
                    }
                    
                    const firstField = document.getElementById('newRowField_0');
                    if (firstField) firstField.focus();
                    
                    // Erfolgsmeldung auch im Modal anzeigen
                    showStatus(elements.newRowStatus, '✅ Zeile erfolgreich übertragen!', 'success');
                } catch (err) {
                    console.error('Fehler bei Übertragung:', err);
                    const errorMsg = `Fehler: ${err.message}`;
                    showStatus(elements.transferStatus, errorMsg, 'error');
                    showStatus(elements.newRowStatus, errorMsg, 'error');
                }
                
                return;
            }
            
            // Browser-Modus: Original-Logik
            const dataStartCol = getDataStartColumn();
            const newRow = new Array(Math.max(
                state.file2.headers.length,
                dataStartCol + state.mapping.sourceColumns.length
            )).fill('');
            
            // Flag und Kommentar in automatisch berechneten Spalten
            if (isFlagEnabled()) {
                newRow[getFlagColumn() - 1] = flag;
            }
            if (isCommentEnabled()) {
                newRow[getCommentColumn() - 1] = comment;
            }
            
            state.mapping.sourceColumns.forEach((srcColIndex, i) => {
                const targetColIndex = dataStartCol - 1 + i;
                newRow[targetColIndex] = data[srcColIndex] || '';
            });
            
            state.file2.data.push(newRow);
            updateWorkbook();
            
            state.history.unshift({
                time: formatHistoryDateTime(),
                flag: flag,
                searchValue: '(Manuelle Eingabe)',
                preview: data[state.mapping.duplicateCheckColumn] || '(Neue Zeile)'
            });
            if (state.history.length > 100) state.history = state.history.slice(0, 100);
            updateHistoryDisplay();
            
            showStatus(elements.transferStatus, '✅ Neue Zeile direkt übertragen!', 'success');
            
            state.mapping.sourceColumns.forEach((colIndex, i) => {
                const field = document.getElementById(`newRowField_${i}`);
                if (field) field.value = '';
            });
            elements.newRowComment.value = '';
            
            const firstField = document.getElementById('newRowField_0');
            if (firstField) firstField.focus();
            
            saveConfig();
        }
        
        // ==================== Transfer ====================
        
        // Pattern für Change-Request-Dateien (konfigurierbar über Spalten konfigurieren)
        function getChangeRequestPattern() {
            return state.mapping.changeRequestPattern || '*_Change_Request*.*';
        }
        
        /**
         * Lädt alle Change-Request-Dateien aus dem Zieldatei-Verzeichnis
         * und cached deren Daten für schnelle Duplikatprüfung
         */
        async function loadChangeRequestFiles() {
            if (!state.file2.filePath) {
                return;
            }
            
            // Verzeichnis aus Zieldatei-Pfad extrahieren (unterstützt / und \)
            const lastSep = Math.max(state.file2.filePath.lastIndexOf('/'), state.file2.filePath.lastIndexOf('\\'));
            const directory = lastSep > 0 ? state.file2.filePath.substring(0, lastSep) : state.file2.filePath;
            
            // Prüfe ob Cache aktuell ist (gleiches Verzeichnis, nicht älter als 5 Minuten)
            const cacheAge = state.changeRequestCache.lastUpdate 
                ? Date.now() - state.changeRequestCache.lastUpdate 
                : Infinity;
            
            if (state.changeRequestCache.directory === directory && cacheAge < 5 * 60 * 1000) {
                return;
            }
            
            try {
                const pattern = getChangeRequestPattern();
                if (!pattern) {
                    return;
                }
                // Die aktuelle Zieldatei wird unten separat durchsucht. Die
                // übrigen Change Requests liegen sowohl im Arbeitsordner als
                // auch optional im Archivunterordner.
                const pathSeparator = directory.includes('\\') ? '\\' : '/';
                const archiveDirectory = directory.endsWith('/') || directory.endsWith('\\')
                    ? `${directory}Change_Requests_old`
                    : `${directory}${pathSeparator}Change_Requests_old`;
                const [currentDirectoryResult, archiveDirectoryResult] = await Promise.all([
                    window.electronAPI.findFiles(directory, pattern),
                    window.electronAPI.findFiles(archiveDirectory, pattern)
                ]);
                
                if (!currentDirectoryResult.success) {
                    console.error('[ChangeRequest] Fehler bei Dateisuche:', currentDirectoryResult.error);
                    return;
                }

                // Der Archivordner ist optional. Andere Fehler (z. B. fehlende
                // Berechtigung auf einem Netzlaufwerk) werden weiterhin protokolliert.
                if (!archiveDirectoryResult.success && archiveDirectoryResult.error !== 'Verzeichnis nicht gefunden') {
                    console.warn('[ChangeRequest] Fehler bei Suche im Archivordner:', archiveDirectoryResult.error);
                }
                
                // Zieldatei selbst nicht erneut als Change Request laden: Sie
                // wird in checkForDuplicate() separat und immer geprüft.
                const files = [...new Set([
                    ...currentDirectoryResult.files,
                    ...(archiveDirectoryResult.success ? archiveDirectoryResult.files : [])
                ])].filter(f => f !== state.file2.filePath);
                
                console.log(`[ChangeRequest] Suchmuster "${pattern}" → ${files.length} Dateien gefunden in ${directory} und ${archiveDirectory}`);
                
                // Cache zurücksetzen
                state.changeRequestCache.files = files;
                state.changeRequestCache.data.clear();
                state.changeRequestCache.directory = directory;
                state.changeRequestCache.lastUpdate = Date.now();
                
                // Dateien laden (nur Excel-Dateien)
                for (const filePath of files) {
                    if (!filePath.match(/\.(xlsx|xls)$/i)) continue;
                    
                    try {
                        // Schnelles Laden: nur Sheets auflisten
                        const fileResult = await window.electronAPI.readExcelFile(filePath);
                        
                        if (fileResult.success && fileResult.sheets.length > 0) {
                            // Dasselbe Sheet wie in der Zieldatei verwenden, falls vorhanden
                            const targetSheet = state.file2.selectedSheet || fileResult.sheets[0];
                            const sheetToLoad = fileResult.sheets.includes(targetSheet) ? targetSheet : fileResult.sheets[0];
                            
                            const sheetResult = await window.electronAPI.readExcelSheet(
                                filePath, 
                                sheetToLoad
                            );
                            
                            if (sheetResult.success) {
                                state.changeRequestCache.data.set(filePath, {
                                    fileName: filePath.substring(Math.max(filePath.lastIndexOf('/'), filePath.lastIndexOf('\\')) + 1),
                                    headers: sheetResult.headers || [],
                                    data: sheetResult.data || [],
                                    sheetName: sheetToLoad
                                });
                            }
                        }
                    } catch (err) {
                        console.warn(`[ChangeRequest] Fehler beim Laden von ${filePath}:`, err.message);
                    }
                }
                

                
            } catch (err) {
                console.error('[ChangeRequest] Fehler:', err);
            }
        }
        
        /**
         * Prüft ob ein Wert bereits in der Zieldatei oder in Change-Request-Dateien existiert
         * @param {string} value - Der zu prüfende Wert
         * @returns {Object|null} - {rowIndex, value, source} wenn gefunden, sonst null
         */
        /**
         * Prüft ob ein Wert bereits in der Zieldatei oder Change-Request-Dateien vorkommt.
         * Gibt ein Objekt mit allen Fundorten zurück.
         * @param {string} value - Der zu prüfende Wert
         * @returns {Object|null} - { inTarget: boolean, inChangeRequests: string[], firstMatch: {...} } oder null
         */
        function checkForDuplicate(value) {
            if (!value) return null;
            
            const valueStr = String(value).toLowerCase().trim();
            const dataStartCol = getDataStartColumn();
            
            let result = {
                inTarget: false,
                inChangeRequests: [], // Liste der Dateinamen
                firstMatch: null
            };
            
            // 1. Prüfe in Zieldatei (Datei 2)
            if (state.file2.data.length > 0) {
                for (let i = 0; i < state.file2.data.length; i++) {
                    const row = state.file2.data[i];
                    for (let j = dataStartCol - 1; j < row.length; j++) {
                        if (row[j] && String(row[j]).toLowerCase().trim() === valueStr) {
                            result.inTarget = true;
                            if (!result.firstMatch) {
                                result.firstMatch = { 
                                    rowIndex: i + 2, 
                                    value: row[j], 
                                    source: 'Zieldatei' 
                                };
                            }
                            break; // Keine weitere Suche in dieser Datei nötig
                        }
                    }
                    if (result.inTarget) break;
                }
            }
            
            // 2. Prüfe in Change-Request-Dateien (Cache)
            for (const [filePath, fileData] of state.changeRequestCache.data) {
                if (!fileData.data) continue;
                
                let foundInFile = false;
                for (let i = 0; i < fileData.data.length && !foundInFile; i++) {
                    const row = fileData.data[i];
                    // Prüfe alle Spalten (Change-Request-Dateien haben möglicherweise anderes Layout)
                    for (let j = 0; j < row.length; j++) {
                        const cellValue = row[j] ? String(row[j]).toLowerCase().trim() : '';
                        if (cellValue === valueStr) {
                            foundInFile = true;
                            result.inChangeRequests.push(fileData.fileName || 'Change-Request');
                            if (!result.firstMatch) {
                                result.firstMatch = { 
                                    rowIndex: i + 2, 
                                    value: row[j], 
                                    source: fileData.fileName || 'Change-Request'
                                };
                            }
                            break;
                        }
                    }
                }
            }
            
            // Nur zurückgeben wenn mindestens ein Treffer
            if (result.inTarget || result.inChangeRequests.length > 0) {
                return result;
            }
            return null;
        }
        
        /**
         * Cache der Change-Request-Dateien invalidieren
         */
        function invalidateChangeRequestCache() {
            state.changeRequestCache.lastUpdate = null;
            state.changeRequestCache.directory = null;
        }
        
        async function transferSelectedDirect() {
            if (state.selectedRows.length === 0) {
                showStatus(elements.transferStatus, 'Bitte wählen Sie mindestens eine Zeile aus', 'error');
                return;
            }
            
            // Prüfe ob Zieldatei verfügbar ist
            const hasTargetFile = state.file2.filePath;
            if (!hasTargetFile) {
                showStatus(elements.transferStatus, 'Keine Zieldatei geladen', 'error');
                return;
            }
            
            // Change-Request-Dateien laden (für erweiterte Duplikatprüfung)
            await loadChangeRequestFiles();
            
            const flag = elements.transferFlag.value;
            const comment = elements.transferComment.value;
            
            // Zur Warteschlange hinzufügen und direkt übertragen
            if (state.file2.filePath) {
                // Zeilen zur Warteschlange hinzufügen
                let addedCount = 0;
                let skippedQueue = 0;
                let skippedTarget = 0;
                let skippedChangeRequest = 0;
                let alsoInChangeRequest = 0;
                
                for (const rowIndex of state.selectedRows) {
                    const row = state.searchResults[rowIndex];
                    if (!row) continue;
                    
                    const rowData = getEditedRowData(rowIndex);
                    const checkValue = rowData[state.mapping.duplicateCheckColumn];
                    
                    // Prüfe ob bereits in Warteschlange
                    const rowAlreadyInQueue = state.transferQueue.some(item => item.rowIndex === row.rowIndex);
                    if (rowAlreadyInQueue) {
                        skippedQueue++;
                        continue;
                    }
                    
                    // Prüfe ob bereits in Zieldatei oder Change-Request-Dateien
                    if (checkValue) {
                        const duplicate = checkForDuplicate(checkValue);
                        if (duplicate) {
                            if (duplicate.inTarget) {
                                skippedTarget++;
                                if (duplicate.inChangeRequests.length > 0) {
                                    alsoInChangeRequest++;
                                }
                            } else if (duplicate.inChangeRequests.length > 0) {
                                skippedChangeRequest++;
                            }
                            continue;
                        }
                    }
                    
                    state.transferQueue.push({
                        data: [...rowData],
                        rowIndex: row.rowIndex,
                        flag: flag,
                        comment: comment,
                        checkValue: checkValue,
                        wasEdited: isRowEdited(rowIndex),
                        sourceRowIndex: row.rowIndex + 2 // +2: rowIndex ist 0-basiert, Excel ist 1-basiert + Header-Zeile
                    });
                    addedCount++;
                }
                
                // Statusmeldung für übersprungene Zeilen
                if (skippedQueue > 0 || skippedTarget > 0 || skippedChangeRequest > 0) {
                    let skipMsg = '';
                    if (skippedQueue > 0) skipMsg += `${skippedQueue} bereits in Warteschlange, `;
                    if (skippedTarget > 0) {
                        skipMsg += `${skippedTarget} bereits in Zieldatei`;
                        if (alsoInChangeRequest > 0) skipMsg += ` (${alsoInChangeRequest} auch in CR)`;
                        skipMsg += ', ';
                    }
                    if (skippedChangeRequest > 0) skipMsg += `${skippedChangeRequest} nur in Change-Request`;
                    showStatus(elements.transferStatus, `⚠️ Übersprungen: ${skipMsg}`, 'warning');
                    await new Promise(r => setTimeout(r, 1500));
                }
                
                if (addedCount === 0) {
                    showStatus(elements.transferStatus, 'Keine neuen Zeilen zum Übertragen', 'warning');
                    return;
                }
                
                updateQueueDisplay();
                
                try {
                    // Direkt übertragen
                    await transferQueueToExcel();
                } catch (err) {
                    console.error('Fehler bei Übertragung:', err);
                    showStatus(elements.transferStatus, `Fehler: ${err.message}`, 'error');
                }
                
                selectAllRows(false);
                elements.transferComment.value = '';
                elements.searchInput.value = '';
                elements.searchInput.focus();
                return;
            }
            
            // Browser-Modus: Original-Logik
            let transferredCount = 0;
            let editedCount = 0;
            let duplicatesFound = [];
            const dataStartCol = getDataStartColumn();
            
            for (const rowIndex of state.selectedRows) {
                const row = state.searchResults[rowIndex];
                if (!row) continue;
                
                const rowData = getEditedRowData(rowIndex);
                const wasEdited = isRowEdited(rowIndex);
                if (wasEdited) editedCount++;
                
                const checkValue = rowData[state.mapping.duplicateCheckColumn];
                
                if (checkValue) {
                    const duplicate = checkForDuplicate(checkValue);
                    if (duplicate && duplicate.firstMatch) {
                        duplicatesFound.push({ 
                            value: checkValue, 
                            row: duplicate.firstMatch.rowIndex,
                            inTarget: duplicate.inTarget,
                            inChangeRequests: duplicate.inChangeRequests
                        });
                    }
                }
                
                const newRow = new Array(Math.max(
                    state.file2.headers.length,
                    dataStartCol + state.mapping.sourceColumns.length
                )).fill('');
                
                // Flag und Kommentar in automatisch berechneten Spalten
                if (isFlagEnabled()) {
                    newRow[getFlagColumn() - 1] = flag;
                }
                if (isCommentEnabled()) {
                    newRow[getCommentColumn() - 1] = comment;
                }
            
                state.mapping.sourceColumns.forEach((srcColIndex, i) => {
                    const targetColIndex = dataStartCol - 1 + i;
                    newRow[targetColIndex] = rowData[srcColIndex] || '';
                });
                
                state.file2.data.push(newRow);
                transferredCount++;
                
                state.history.unshift({
                    time: formatHistoryDateTime(),
                    flag: flag,
                    searchValue: checkValue,
                    preview: rowData[state.mapping.duplicateCheckColumn] || '(Neue Zeile)'
                });
            }
            
            updateWorkbook();
            
            if (state.history.length > 100) state.history.length = 100;
            updateHistoryDisplay();
            saveConfig();
            
            let message = `✓ ${transferredCount} Zeile(n) direkt übertragen!`;
            let status = 'success';
            if (editedCount > 0) {
                message += ` (${editedCount} bearbeitet ✏️)`;
            }
            if (duplicatesFound.length > 0) {
                message += ` ⚠️ ${duplicatesFound.length} Duplikat(e)!`;
                status = 'warning';
            }
            
            showStatus(elements.transferStatus, message, status);
            
            selectAllRows(false);
            elements.transferComment.value = '';
            elements.searchInput.value = '';
            elements.searchInput.focus();
        }
        
        // ==================== Queue Functions ====================
        async function addToQueue() {
            if (state.selectedRows.length === 0) {
                showStatus(elements.transferStatus, 'Bitte wählen Sie mindestens eine Zeile aus', 'error');
                return;
            }
            
            // Change-Request-Dateien laden (für erweiterte Duplikatprüfung)
            await loadChangeRequestFiles();
            
            const flag = elements.transferFlag.value;
            const comment = elements.transferComment.value;
            
            let addedCount = 0;
            let editedCount = 0;
            let skippedQueue = 0;
            let skippedTarget = 0;
            let skippedChangeRequest = 0;
            let alsoInChangeRequest = 0; // Zusätzlich in Change-Request (wenn auch in Zieldatei)
            
            for (const rowIndex of state.selectedRows) {
                const row = state.searchResults[rowIndex];
                if (!row) continue;
                
                const rowData = getEditedRowData(rowIndex);
                const wasEdited = isRowEdited(rowIndex);
                
                const checkValue = rowData[state.mapping.duplicateCheckColumn];
                
                // Prüfe ob bereits in Warteschlange
                const rowAlreadyInQueue = state.transferQueue.some(item => item.rowIndex === row.rowIndex);
                if (rowAlreadyInQueue) {
                    skippedQueue++;
                    continue;
                }
                
                // Prüfe ob bereits in Zieldatei oder Change-Request-Dateien
                if (checkValue) {
                    const duplicate = checkForDuplicate(checkValue);
                    if (duplicate) {
                        if (duplicate.inTarget) {
                            skippedTarget++;
                            // Auch in Change-Requests?
                            if (duplicate.inChangeRequests.length > 0) {
                                alsoInChangeRequest++;
                            }
                        } else if (duplicate.inChangeRequests.length > 0) {
                            skippedChangeRequest++;
                        }
                        continue;
                    }
                }
                
                if (wasEdited) editedCount++;
                
                state.transferQueue.push({
                    data: [...rowData],
                    rowIndex: row.rowIndex,
                    flag: flag,
                    comment: comment,
                    checkValue: checkValue,
                    wasEdited: wasEdited,
                    sourceRowIndex: row.rowIndex + 2 // +2: rowIndex ist 0-basiert, Excel ist 1-basiert + Header-Zeile
                });
                addedCount++;
            }
            
            updateQueueDisplay();
            elements.transferComment.value = '';
            
            let message = `✓ ${addedCount} Zeile(n) zur Warteschlange hinzugefügt`;
            let status = 'success';
            
            if (editedCount > 0) {
                message += ` (${editedCount} bearbeitet ✏️)`;
            }
            if (skippedQueue > 0) {
                message += ` (${skippedQueue} bereits in Warteschlange)`;
            }
            if (skippedTarget > 0) {
                let targetMsg = `⚠️ ${skippedTarget} bereits in Zieldatei`;
                if (alsoInChangeRequest > 0) {
                    targetMsg += ` (${alsoInChangeRequest} davon auch in Change-Request)`;
                }
                message += ` ${targetMsg}!`;
                status = 'warning';
            }
            if (skippedChangeRequest > 0) {
                message += ` 📋 ${skippedChangeRequest} nur in Change-Request!`;
                status = 'warning';
            }
            
            showStatus(elements.transferStatus, message, status);
            
            selectAllRows(false);
            elements.searchInput.value = '';
            elements.searchInput.focus();
        }
        
        function clearQueue() {
            if (state.transferQueue.length > 0 && 
                !confirm(`Warteschlange mit ${state.transferQueue.length} Zeilen wirklich leeren?`)) {
                return;
            }
            state.transferQueue = [];
            updateQueueDisplay();
            autoSave(); // Auto-Save aktualisieren nach Leeren
        }
        
        function updateQueueDisplay() {
            const hasFile2 = state.file2.filePath;
            
            elements.queueCount.textContent = state.transferQueue.length;
            elements.btnClearQueue.disabled = state.transferQueue.length === 0;
            elements.btnExportPS.disabled = state.transferQueue.length === 0 || !hasFile2;
            elements.btnPreviewTransfer.disabled = state.transferQueue.length === 0 || !hasFile2;
            
            if (state.transferQueue.length === 0) {
                elements.queueList.innerHTML = '<div class="queue-empty">Keine Zeilen in der Warteschlange</div>';
                return;
            }
            
            let html = '';
            state.transferQueue.forEach((item, index) => {
                const preview = String(item.checkValue || item.data[0] || '').substring(0, 40);
                const editedBadge = item.wasEdited ? '<span class="queue-item-edited" title="Bearbeitet">✏️</span>' : '';
                html += `
                    <div class="queue-item" style="display: flex; gap: 10px; padding: 8px; background: var(--bg-light); border-radius: 4px; margin-bottom: 5px; align-items: center;">
                        <span style="background: var(--primary); color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold;">${item.flag}</span>
                        ${editedBadge}
                        <span style="flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(String(item.data))}">${escapeHtml(preview)}</span>
                        ${item.comment ? `<span style="color: var(--text-muted); font-size: 12px;" title="${escapeHtml(item.comment)}">${escapeHtml(item.comment.substring(0, 20))}</span>` : '' }
                        <button class="btn btn-secondary btn-sm" data-remove-index="${index}" title="Entfernen">✕</button>
                    </div>`;
            });
            elements.queueList.innerHTML = html;
            
            // Event-Delegation für Remove-Buttons
            elements.queueList.querySelectorAll('[data-remove-index]').forEach(btn => {
                btn.onclick = () => removeFromQueue(parseInt(btn.dataset.removeIndex, 10));
            });
        }
        
        async function transferQueueToExcel() {
            if (state.transferQueue.length === 0) {
                showStatus(elements.transferStatus, 'Keine Zeilen in der Warteschlange', 'error');
                return;
            }
            
            if (!state.file2.filePath) {
                showStatus(elements.transferStatus, 'Keine Zieldatei geladen', 'error');
                return;
            }
            
            const rows = state.transferQueue.map(item => {
                // Konvertiere item.data (Array) zu rowData (Objekt mit Index als Key)
                const rowData = {};
                state.mapping.sourceColumns.forEach((srcColIndex, i) => {
                    rowData[i] = item.data[srcColIndex] || '';
                });
                return {
                    flag: isFlagEnabled() ? item.flag : null,
                    comment: isCommentEnabled() ? item.comment : null,
                    data: rowData,
                    sourceRowIndex: item.sourceRowIndex || null, // Zeilen-Index aus Quelldatei für Formatierung
                    isManual: item.isManual || false
                };
            });
            
            const result = await window.electronAPI.insertExcelRows({
                filePath: state.file2.filePath,
                sheetName: state.file2.selectedSheet,
                rows: rows,
                startColumn: getDataStartColumn(),
                enableFlag: isFlagEnabled(),
                enableComment: isCommentEnabled(),
                flagColumn: getFlagColumn(),
                commentColumn: getCommentColumn(),
                // Quelldatei-Infos für Formatierungskopie
                sourceFilePath: state.file1.filePath || null,
                sourceSheetName: state.file1.selectedSheet || null,
                sourceColumns: state.mapping.sourceColumns || []
            });
            
            if (result.success) {
                state.transferQueue.forEach(item => {
                    state.history.unshift({
                        time: formatHistoryDateTime(),
                        flag: item.flag,
                        searchValue: item.checkValue,
                        preview: String(item.checkValue || item.data[0] || '').substring(0, 30)
                    });
                });
                if (state.history.length > 100) state.history = state.history.slice(0, 100);
                updateHistoryDisplay();
                
                state.transferQueue = [];
                updateQueueDisplay();
                
                // Auto-Save löschen nach erfolgreicher Übertragung
                clearAutoSave();
                
                await loadSheet2Electron(state.file2.selectedSheet);
                
                const successMsg = `✅ ${result.insertedCount} Zeile(n) direkt in Excel eingefügt!`;
                showStatus(elements.transferStatus, successMsg, 'success');
                showStatus(elements.newRowStatus, successMsg, 'success');
            } else {
                const errorMsg = `❌ Fehler: ${result.error}`;
                showStatus(elements.transferStatus, errorMsg, 'error');
                showStatus(elements.newRowStatus, errorMsg, 'error');
            }
        }
        
        // Diff-Vorschau vor Transfer anzeigen
        function showDiffPreview() {
            if (state.transferQueue.length === 0 || !state.file2.filePath) {
                showStatus(elements.transferStatus, 'Keine Zeilen in der Warteschlange oder keine Zieldatei geladen', 'error');
                return;
            }
            
            const modal = elements.diffPreviewModal;
            const targetFileName = state.file2.name || 'Zieldatei';
            const targetSheet = state.file2.selectedSheet || 'Sheet1';
            const currentRowCount = state.file2.data ? state.file2.data.length : 0;
            const startRow = currentRowCount + 1; // 1-basiert für Anzeige
            
            // Info-Bereich aktualisieren
            document.getElementById('diffTargetFile').textContent = targetFileName;
            document.getElementById('diffTargetSheet').textContent = targetSheet;
            document.getElementById('diffTargetRow').textContent = startRow;
            document.getElementById('diffPreviewCount').textContent = state.transferQueue.length;
            
            // Tabelle rendern
            const tableContainer = document.getElementById('diffPreviewTable');
            
            // Ermittle ob Flag/Kommentar aktiv sind
            const flagEnabled = isFlagEnabled();
            const commentEnabled = isCommentEnabled();
            const flagCol = getFlagColumn();      // 1-basiert
            const commentCol = getCommentColumn(); // 1-basiert
            
            // Header erstellen - Zeile zuerst, dann Spalten wie sie in der Zieldatei erscheinen
            let headerHtml = '<tr><th style="width: 50px; text-align: center;">#</th>';
            
            // Spaltenüberschriften aus der Zieldatei verwenden
            const targetHeaders = state.file2.headers || [];
            // Automatisch berechnete Startspalte (0-basiert)
            const targetStartCol = getDataStartColumn() - 1;
            
            // Berechne welche Spalten in der Vorschau angezeigt werden sollen
            // Zeige alle Spalten von der niedrigsten bis zur höchsten belegten Spalte
            const previewColumns = [];
            
            // Sammle alle belegten Spaltenindizes
            const usedColumns = new Map(); // index -> {type, sourceIndex?, name}
            
            // Flag-Spalte (wenn aktiv)
            if (flagEnabled) {
                usedColumns.set(flagCol - 1, {
                    type: 'flag',
                    name: targetHeaders[flagCol - 1] || `Flag`
                });
            }
            
            // Kommentar-Spalte (wenn aktiv)
            if (commentEnabled) {
                usedColumns.set(commentCol - 1, {
                    type: 'comment',
                    name: targetHeaders[commentCol - 1] || `Kommentar`
                });
            }
            
            // Daten-Spalten aus dem Mapping
            state.mapping.sourceColumns.forEach((srcIdx, i) => {
                const targetColIdx = targetStartCol + i;
                usedColumns.set(targetColIdx, {
                    type: 'data',
                    sourceIndex: srcIdx,
                    name: targetHeaders[targetColIdx] || `Spalte ${String.fromCharCode(65 + targetColIdx)}`
                });
            });
            
            // Finde min und max Spaltenindex
            const colIndices = Array.from(usedColumns.keys());
            const minCol = Math.min(...colIndices);
            const maxCol = Math.max(...colIndices);
            
            // Erstelle Spalten von min bis max (inkl. leerer Spalten)
            for (let i = minCol; i <= maxCol; i++) {
                if (usedColumns.has(i)) {
                    previewColumns.push({
                        index: i,
                        ...usedColumns.get(i)
                    });
                } else {
                    // Leere Spalte
                    previewColumns.push({
                        index: i,
                        type: 'empty',
                        name: targetHeaders[i] || `-`
                    });
                }
            }
            
            // Header-Zeile erstellen
            previewColumns.forEach(col => {
                const colLetter = String.fromCharCode(65 + col.index);
                const style = col.type === 'flag' ? 'background: rgba(33, 115, 70, 0.3);' : 
                              col.type === 'comment' ? 'background: rgba(33, 115, 70, 0.2);' : 
                              col.type === 'empty' ? 'background: rgba(128, 128, 128, 0.1); color: var(--text-muted);' : '';
                headerHtml += `<th style="${style}"><small style="color: var(--text-muted);">${colLetter}</small><br>${escapeHtml(col.name)}</th>`;
            });
            headerHtml += '</tr>';
            
            // Zeilen rendern
            let rowsHtml = '';
            state.transferQueue.forEach((item, idx) => {
                const rowNum = startRow + idx;
                const flagClass = item.flag === 'A' ? 'diff-row-add' : 
                                  item.flag === 'D' ? 'diff-row-delete' : 
                                  item.flag === 'C' ? 'diff-row-change' : '';
                
                rowsHtml += `<tr class="${flagClass}">`;
                rowsHtml += `<td style="text-align: center; font-weight: bold; color: var(--text-muted);">${rowNum}</td>`;
                
                // Zellen in der richtigen Reihenfolge
                previewColumns.forEach(col => {
                    if (col.type === 'flag') {
                        const flagValue = item.flag || '';
                        const flagStyle = 'font-weight: bold; text-align: center; background: rgba(33, 115, 70, 0.15);';
                        rowsHtml += `<td style="${flagStyle}"><span class="diff-flag">${flagValue}</span></td>`;
                    } else if (col.type === 'comment') {
                        const commentValue = item.comment || '';
                        const commentStyle = 'font-style: italic; background: rgba(33, 115, 70, 0.1);';
                        rowsHtml += `<td style="${commentStyle}">${escapeHtml(String(commentValue))}</td>`;
                    } else if (col.type === 'empty') {
                        // Leere Spalte - wird nicht beschrieben
                        rowsHtml += `<td style="color: var(--text-muted); text-align: center;">-</td>`;
                    } else {
                        const cellValue = item.data[col.sourceIndex] !== undefined ? item.data[col.sourceIndex] : '';
                        rowsHtml += `<td>${escapeHtml(String(cellValue))}</td>`;
                    }
                });
                
                rowsHtml += '</tr>';
            });
            
            tableContainer.innerHTML = `
                <table class="diff-table">
                    <thead>${headerHtml}</thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            `;
            
            // Modal anzeigen
            modal.classList.remove('hidden');
        }
        
        function closeDiffPreview() {
            elements.diffPreviewModal.classList.add('hidden');
        }
        
        function confirmTransferFromDiff() {
            closeDiffPreview();
            transferQueueToExcel();
        }
        
        // ==================== Working Directory Functions ====================
        async function selectWorkingDirectory() {
            try {
                const folderPath = await window.electronAPI.openFolderDialog({
                    title: t('selectWorkingDir') || 'Arbeitsordner auswählen'
                });
                
                if (!folderPath) return;
                
                state.workingDirectory = folderPath;
                updateWorkingDirectoryUI();
                
                // Im localStorage speichern für Persistenz
                localStorage.setItem('workingDirectory', folderPath);
                
                showStatus(elements.transferStatus, `✓ ${t('workingDirSet')}${folderPath}`, 'success');
            } catch (err) {
                console.error('Fehler beim Auswählen des Arbeitsordners:', err);
                showStatus(elements.transferStatus, `Fehler: ${err.message}`, 'error');
            }
        }
        
        function clearWorkingDirectory() {
            state.workingDirectory = null;
            localStorage.removeItem('workingDirectory');
            updateWorkingDirectoryUI();
            showStatus(elements.transferStatus, `✓ ${t('workingDirCleared')}`, 'success');
        }
        
        function updateWorkingDirectoryUI() {
            if (state.workingDirectory) {
                // Zeige den Ordnernamen (letzter Teil des Pfads)
                const folderName = state.workingDirectory.split(/[/\\]/).pop();
                elements.workingDirInfo.textContent = `✓ ${folderName}`;
                elements.workingDirInfo.title = state.workingDirectory;
                elements.workingDirInfo.classList.add('loaded');
                elements.btnClearWorkingDir.style.display = 'block';
            } else {
                elements.workingDirInfo.textContent = t('noWorkingDirSet');
                elements.workingDirInfo.title = '';
                elements.workingDirInfo.classList.remove('loaded');
                elements.btnClearWorkingDir.style.display = 'none';
            }
        }
        
        function getWorkingDirectoryPath() {
            return state.workingDirectory || undefined;
        }
        
        /**
         * Bestimmt den Standard-Pfad für den Explorer-Dateidialog.
         * Priorität: 1) Arbeitsordner  2) Ordner der zuletzt geöffneten Datei  3) Ordner aus config
         */
        function getExplorerDefaultPath() {
            // 1. Arbeitsordner (vom Benutzer explizit gesetzt)
            if (state.workingDirectory) return state.workingDirectory;
            
            // 2. Ordner der aktuell/zuletzt geöffneten Explorer-Datei
            if (explorerState.filePath) {
                const sep = explorerState.filePath.includes('\\') ? '\\' : '/';
                const dir = explorerState.filePath.substring(0, explorerState.filePath.lastIndexOf(sep));
                if (dir) return dir;
            }
            
            // 3. Ordner der file1 oder file2 aus dem Transfer-Bereich
            const transferFile = state.file1?.filePath || state.file2?.filePath;
            if (transferFile) {
                const sep = transferFile.includes('\\') ? '\\' : '/';
                const dir = transferFile.substring(0, transferFile.lastIndexOf(sep));
                if (dir) return dir;
            }
            
            return undefined;
        }
        
        // Arbeitsordner beim Start laden
        function loadWorkingDirectoryFromStorage() {
            const savedPath = localStorage.getItem('workingDirectory');
            if (savedPath) {
                state.workingDirectory = savedPath;
                updateWorkingDirectoryUI();
            }
        }
        
        // ==================== Electron-Specific Functions ====================
        async function loadMainFileFromPath(fileNumber, filePath, preferredSheetName = null, { createSessionLock = false } = {}) {
            const result = await window.electronAPI.readExcelFile(filePath);
            if (!result.success) return result;

            const isSource = fileNumber === 1;
            const fileState = isSource ? state.file1 : state.file2;
            const sheetSelect = isSource ? elements.selectSheet1 : elements.selectSheet2;
            const fileInfo = isSource ? elements.file1Info : elements.file2Info;

            fileState.name = result.fileName;
            fileState.filePath = filePath;
            fileState.sheets = result.sheets;
            fileState.workbook = { SheetNames: result.sheets };

            if (!isSource) invalidateChangeRequestCache();

            sheetSelect.innerHTML = result.sheets.map(s => `<option value="${s}">${s}</option>`).join('');
            sheetSelect.disabled = false;
            fileInfo.textContent = `✓ ${result.fileName}`;
            fileInfo.classList.add('loaded');

            if (createSessionLock) {
                await window.electronAPI.createSessionLock(filePath);
            }

            const sheetToLoad = preferredSheetName && result.sheets.includes(preferredSheetName)
                ? preferredSheetName
                : result.sheets[0];
            sheetSelect.value = sheetToLoad;
            const sheetLoaded = isSource
                ? await loadSheet1Electron(sheetToLoad)
                : await loadSheet2Electron(sheetToLoad);

            return sheetLoaded
                ? { success: true, fileName: result.fileName }
                : { success: false, error: 'Arbeitsblatt konnte nicht geladen werden' };
        }

        async function loadFile1Electron() {
            const filePath = await window.electronAPI.openFileDialog({
                title: 'Quelldatei öffnen',
                filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
                defaultPath: getWorkingDirectoryPath()
            });
            if (!filePath) return;
            
            // Konfliktprüfung für Netzlaufwerke
            const conflictCheck = await checkAndWarnNetworkConflict(filePath);
            if (!conflictCheck.proceed) return;
            
            const result = await loadMainFileFromPath(1, filePath, null, { createSessionLock: true });
            if (!result.success) {
                showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
            }
        }
        
        async function loadFile2Electron() {
            const filePath = await window.electronAPI.openFileDialog({
                title: 'Zieldatei öffnen',
                filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
                defaultPath: getWorkingDirectoryPath()
            });
            if (!filePath) return;
            
            // Konfliktprüfung für Netzlaufwerke
            const conflictCheck = await checkAndWarnNetworkConflict(filePath);
            if (!conflictCheck.proceed) return;
            
            const result = await loadMainFileFromPath(2, filePath, null, { createSessionLock: true });
            if (!result.success) {
                showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
            }
        }
        
        async function loadSheet1Electron(sheetName) {
            if (!state.file1.filePath || !sheetName) return;
            
            const result = await window.electronAPI.readExcelSheet(state.file1.filePath, sheetName, null, { dataOnly: true });
            if (!result.success) {
                showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
                return false;
            }
            
            state.file1.selectedSheet = sheetName;
            state.file1.headers = result.headers;
            state.file1.data = result.data.slice(1);
            
            saveConfig();
            checkReadyState();
            return true;
        }
        
        async function loadSheet2Electron(sheetName) {
            if (!state.file2.filePath || !sheetName) return;
            
            const result = await window.electronAPI.readExcelSheet(state.file2.filePath, sheetName, null, { dataOnly: true });
            if (!result.success) {
                showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
                return false;
            }
            
            state.file2.selectedSheet = sheetName;
            state.file2.headers = result.headers;
            state.file2.data = result.data.slice(1);
            
            saveConfig();
            checkReadyState();
            return true;
        }
        
        async function loadTemplateElectron() {
            const filePath = await window.electronAPI.openFileDialog({
                title: 'Template-Datei öffnen',
                filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
                defaultPath: getWorkingDirectoryPath()
            });
            if (!filePath) return;
            
            const result = await window.electronAPI.readExcelFile(filePath);
            if (!result.success) {
                showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
                return;
            }
            
            state.template.filePath = filePath;
            state.template.name = result.fileName;
            
            elements.templateInfo.textContent = `✓ ${result.fileName}`;
            elements.templateInfo.classList.add('loaded');
            elements.btnNewMonth.disabled = false;
            
            showStatus(elements.transferStatus, `✓ Template geladen: ${result.fileName}`, 'success');
        }
        
        /**
         * Erstellt ein Template aus einer Quelldatei mit allen Formatierungen
         * - Öffnet Quelldatei-Dialog
         * - Zeigt Sheet-Auswahl Modal
         * - Öffnet Speicher-Dialog für neues Template
         * - Erstellt leeres Template mit erweiterten CF-Ranges
         */
        
        // State für Template-Erstellung
        let createTemplateState = {
            sourcePath: null,
            sheets: []
        };
        
        async function createTemplateFromSourceElectron() {
            const lang = localStorage.getItem('excelSyncLanguage') || 'de';
            const isDE = lang === 'de';
            
            // 1. Quelldatei auswählen
            const sourcePath = await window.electronAPI.openFileDialog({
                title: isDE ? 'Quelldatei auswählen (mit Formatierungen)' : 'Select source file (with formatting)',
                filters: [{ name: 'Excel', extensions: ['xlsx'] }],
                defaultPath: getWorkingDirectoryPath()
            });
            if (!sourcePath) return;
            
            // 2. Datei lesen um Sheets zu bekommen
            const fileResult = await window.electronAPI.readExcelFile(sourcePath);
            if (!fileResult.success) {
                showStatus(elements.transferStatus, `❌ ${isDE ? 'Fehler' : 'Error'}: ${fileResult.error}`, 'error');
                return;
            }
            
            // State speichern
            createTemplateState.sourcePath = sourcePath;
            createTemplateState.sheets = fileResult.sheets;
            
            // 3. Modal mit Sheet-Auswahl anzeigen
            elements.createTemplateSourceName.textContent = fileResult.fileName;
            
            // Sheet-Liste aufbauen
            const sheetListHtml = fileResult.sheets.map((sheetName, index) => `
                <label style="display: flex; align-items: center; padding: 6px 8px; cursor: pointer; border-radius: 4px; transition: background 0.2s;" 
                       onmouseover="this.style.background='var(--bg-medium)'" 
                       onmouseout="this.style.background='transparent'">
                    <input type="checkbox" class="template-sheet-checkbox" value="${sheetName}" checked 
                           style="margin-right: 10px; width: 16px; height: 16px; cursor: pointer;">
                    <span style="flex: 1;">${sheetName}</span>
                    <span style="color: var(--text-muted); font-size: 11px;">Sheet ${index + 1}</span>
                </label>
            `).join('');
            
            elements.createTemplateSheetList.innerHTML = sheetListHtml;
            
            // Modal anzeigen
            elements.createTemplateModal.classList.remove('hidden');
        }
        
        async function confirmCreateTemplate() {
            const lang = localStorage.getItem('excelSyncLanguage') || 'de';
            const isDE = lang === 'de';
            
            // Ausgewählte Sheets sammeln
            const checkboxes = document.querySelectorAll('.template-sheet-checkbox:checked');
            const selectedSheets = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedSheets.length === 0) {
                showStatus(elements.transferStatus, isDE ? '⚠️ Bitte mindestens ein Arbeitsblatt auswählen' : '⚠️ Please select at least one worksheet', 'error');
                return;
            }
            
            // Extra-Spalten Optionen lesen
            const addFlagColumn = document.getElementById('templateFlagColumn').checked;
            const addCommentColumn = document.getElementById('templateCommentColumn').checked;
            
            // Modal schließen
            elements.createTemplateModal.classList.add('hidden');
            
            // Speicherort für Template wählen
            const outputPath = await window.electronAPI.saveFileDialog({
                title: isDE ? 'Template speichern als' : 'Save template as',
                defaultPath: createTemplateState.sourcePath.replace('.xlsx', '_Template.xlsx'),
                filters: [{ name: 'Excel', extensions: ['xlsx'] }]
            });
            if (!outputPath) return;
            
            // Template erstellen
            showStatus(elements.transferStatus, isDE ? '⏳ Template wird erstellt...' : '⏳ Creating template...', 'pending');
            
            const result = await window.electronAPI.createTemplateFromSource({
                sourcePath: createTemplateState.sourcePath,
                outputPath,
                selectedSheets,
                addFlagColumn,
                addCommentColumn
            });
            
            if (!result.success) {
                showStatus(elements.transferStatus, `❌ ${isDE ? 'Fehler' : 'Error'}: ${result.error}`, 'error');
                return;
            }
            
            // Erfolgsmeldung mit Stats
            const stats = result.stats;
            const extraInfo = stats.extraColumnsAdded > 0 
                ? (isDE ? `, ${stats.extraColumnsAdded} Extra-Spalte(n)` : `, ${stats.extraColumnsAdded} extra column(s)`)
                : '';
            const msg = isDE 
                ? `✓ Template erstellt: ${result.fileName}\n   (${stats.sheetsProcessed} Sheet(s), ${stats.cfRulesPreserved} CF-Regeln${extraInfo})`
                : `✓ Template created: ${result.fileName}\n   (${stats.sheetsProcessed} sheet(s), ${stats.cfRulesPreserved} CF rules${extraInfo})`;
            
            showStatus(elements.transferStatus, msg, 'success');
            
            // Template automatisch laden
            const loadResult = await window.electronAPI.readExcelFile(outputPath);
            if (loadResult.success) {
                state.template.filePath = outputPath;
                state.template.name = loadResult.fileName;
                
                elements.templateInfo.textContent = `✓ ${loadResult.fileName}`;
                elements.templateInfo.classList.add('loaded');
                elements.btnNewMonth.disabled = false;
            }
        }
        
        function closeCreateTemplateModal() {
            elements.createTemplateModal.classList.add('hidden');
        }
        
        async function applyLoadedConfig(config, skipFiles = false) {
            state.mapping = config.mapping || state.mapping;
            
            // Extra-Spalten Konfiguration laden
            if (config.extraColumns) {
                const ec = config.extraColumns;
                const flagCheckbox = document.getElementById('enableFlagColumn');
                const commentCheckbox = document.getElementById('enableCommentColumn');
                const flagValuesInput = document.getElementById('flagValues');
                const commentPlaceholderInput = document.getElementById('commentPlaceholder');
                
                if (ec.enableFlag !== undefined && flagCheckbox) {
                    flagCheckbox.checked = ec.enableFlag;
                    localStorage.setItem('excelSyncEnableFlag', String(ec.enableFlag));
                }
                if (ec.enableComment !== undefined && commentCheckbox) {
                    commentCheckbox.checked = ec.enableComment;
                    localStorage.setItem('excelSyncEnableComment', String(ec.enableComment));
                }
                if (ec.flagValues && flagValuesInput) {
                    flagValuesInput.value = ec.flagValues;
                    localStorage.setItem('excelSyncFlagValues', ec.flagValues);
                }
                if (ec.commentPlaceholder && commentPlaceholderInput) {
                    commentPlaceholderInput.value = ec.commentPlaceholder;
                    localStorage.setItem('excelSyncCommentPlaceholder', ec.commentPlaceholder);
                }
                
                // UI aktualisieren
                updateFlagDropdownOptions();
                updateCommentPlaceholders();
                updateFlagCommentVisibility();
                updateColumnDisplays();
            }
            
            // Bei "Öffnen mit...": Nur Mapping/Settings laden, Dateien überspringen
            if (skipFiles) {
                console.log('[Config] skipFiles=true → Quell-/Zieldatei/Template werden NICHT geladen');
                updateMappingInfo();
                updateMappingPreview();
                saveConfig();
                return;
            }
            
            // Datei 1 und Datei 2 sind unabhängig. Beide Metadaten- und
            // Sheet-Ladevorgänge dürfen deshalb gleichzeitig laufen.
            const mainFileLoads = [];
            if (config.file1Path) {
                mainFileLoads.push(
                    loadMainFileFromPath(1, config.file1Path, config.file1SheetName)
                        .catch(error => console.warn('Konnte Datei 1 nicht laden:', error))
                );
            }
            if (config.file2Path) {
                mainFileLoads.push(
                    loadMainFileFromPath(2, config.file2Path, config.file2SheetName)
                        .catch(error => console.warn('Konnte Datei 2 nicht laden:', error))
                );
            }
            await Promise.all(mainFileLoads);
            
            if (config.templatePath) {
                try {
                    const result = await window.electronAPI.readExcelFile(config.templatePath);
                    if (result.success) {
                        state.template.filePath = config.templatePath;
                        state.template.name = result.fileName;
                        elements.templateInfo.textContent = `✓ ${result.fileName}`;
                        elements.templateInfo.classList.add('loaded');
                        elements.btnNewMonth.disabled = false;
                    }
                } catch (e) {
                    console.warn('Konnte Template nicht laden:', e);
                }
            }
            
            updateMappingInfo();
            updateMappingPreview();
            saveConfig();
        }
        
        async function loadConfigFromAppDirOrDialog() {
            try {
                // Zuerst Dialog öffnen
                const filePath = await window.electronAPI.openFileDialog({
                    title: 'config.json laden',
                    filters: [{ name: 'JSON', extensions: ['json'] }],
                    defaultPath: getWorkingDirectoryPath()
                });
                
                if (filePath) {
                    const result = await window.electronAPI.loadConfig(filePath);
                    if (result.success && result.config) {
                        await applyLoadedConfig(result.config);
                        // Zeige Computer-spezifische Info
                        let statusMsg = `✓ config.json geladen: ${filePath}`;
                        if (result.userId && !result.isLegacyFormat) {
                            statusMsg = result.hasUserSection
                                ? `✓ Config für Benutzer „${result.userId}“ geladen`
                                : `✓ Config geladen (Standard, kein Abschnitt für Benutzer „${result.userId}“)`;
                        }
                        showStatus(elements.transferStatus, statusMsg, 'success');
                    } else {
                        showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
                    }
                }
            } catch (e) {
                showStatus(elements.transferStatus, `Fehler: ${e.message}`, 'error');
            }
        }
        
        // ==================== Data Explorer Functions ====================
        const EXPLORER_RECOVERY_KEY = 'excelSyncExplorerRecovery';
        let explorerAutoSaveInterval = null;
        
        // Zwischenspeicher für kopierte Zellen mit Formatierung
        let copiedCellsWithFormat = null;  // { cells: [{row, col, value, style, formula, hyperlink}], minRow, minCol }
        
        // ==================== Live Session Helper Functions ====================
        // Live Session = Excel bleibt im Hintergrund offen, Operationen werden SOFORT ausgeführt
        
        /**
         * Aktualisiert den Live-Mode-Indikator in der Explorer-Toolbar
         */
        function updateLiveModeIndicator() {
            const indicator = document.getElementById('liveModeIndicator');
            const undoBtn = document.getElementById('btnExplorerUndo');
            
            if (!indicator) return;
            
            const isActive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            const isExcelReady = explorerState.engineMode === 'live';
            
            // Undo-Button nur bei aktiver Live-Session anzeigen
            if (undoBtn) {
                undoBtn.style.display = isActive ? '' : 'none';
            }
            
            if (isActive) {
                // Grüner Punkt = Live aktiv (Datei geöffnet)
                indicator.style.background = 'rgba(76, 175, 80, 0.2)';
                indicator.style.border = '1px solid #4CAF50';
                indicator.innerHTML = `<span style="color: #4CAF50;">🟢</span> <span style="color: #4CAF50;">${t('liveReady')}</span>`;
            } else if (isExcelReady) {
                // Grüner Punkt = Excel bereit (keine Datei geladen)
                indicator.style.background = 'rgba(76, 175, 80, 0.15)';
                indicator.style.border = '1px solid #4CAF50';
                indicator.innerHTML = `<span style="color: #4CAF50;">🟢</span> <span style="color: #4CAF50;">Online</span>`;
            } else {
                // Grauer Punkt = Offline/openpyxl
                indicator.style.background = 'var(--bg-lighter)';
                indicator.style.border = '1px solid var(--border)';
                indicator.innerHTML = `<span style="color: var(--text-muted);">⚫</span> <span style="color: var(--text-muted);">${t('liveOffline')}</span>`;
            }
            
            // Sync-Filters-Button aktualisieren
            updateSyncFiltersButton();
        }
        
        /**
         * Startet die Live-Session und öffnet die Datei in Excel
         * @returns {Promise<boolean>} true wenn erfolgreich
         */
        async function startLiveSession() {
            const _liveStart = Date.now();
            const _liveLog = (msg) => console.log(`[LiveSession ${Date.now() - _liveStart}ms] ${msg}`);
            
            if (!explorerState.filePath || !explorerState.selectedSheet) {
                _liveLog('ABBRUCH: Keine Datei/Sheet geladen');
                return false;
            }
            
            try {
                _liveLog(`=== START === Datei: ${explorerState.filePath}, Sheet: ${explorerState.selectedSheet}`);
                
                // Python-Prozess starten
                _liveLog('liveSessionStart aufrufen...');
                const startResult = await window.electronAPI.liveSessionStart();
                _liveLog(`liveSessionStart: success=${startResult.success}${startResult.error ? ', error=' + startResult.error : ''}`);
                if (!startResult.success) {
                    showFloatingStatus('Live-Session Start fehlgeschlagen: ' + startResult.error, true);
                    return false;
                }
                
                // Datei öffnen - bei Fehlschlag Retry nach kurzer Pause (Windows File-Locking)
                _liveLog('liveSessionOpenFile aufrufen (1. Versuch)...');
                let openResult = await window.electronAPI.liveSessionOpenFile(
                    explorerState.filePath, 
                    explorerState.selectedSheet,
                    explorerState.filePassword || null
                );
                _liveLog(`liveSessionOpenFile (1): success=${openResult.success}${openResult.error ? ', error=' + openResult.error : ''}`);
                
                if (!openResult.success) {
                    _liveLog('Warte 2s für Retry...');
                    await new Promise(r => setTimeout(r, 2000));
                    _liveLog('liveSessionOpenFile aufrufen (2. Versuch)...');
                    openResult = await window.electronAPI.liveSessionOpenFile(
                        explorerState.filePath, 
                        explorerState.selectedSheet,
                        explorerState.filePassword || null
                    );
                    _liveLog(`liveSessionOpenFile (2): success=${openResult.success}${openResult.error ? ', error=' + openResult.error : ''}`);
                }
                
                if (!openResult.success) {
                    _liveLog(`FEHLGESCHLAGEN: ${openResult.error}`);
                    showFloatingStatus('Excel konnte Datei nicht öffnen: ' + openResult.error, true);
                    return false;
                }
                
                explorerState.liveSessionActive = true;
                explorerState.liveSessionReady = true;
                explorerState.excelVisible = false;
                explorerState.excelInteractive = false;
                
                // Read-Only-Warnung wenn Datei schreibgeschützt geöffnet wurde
                if (openResult.readOnly) {
                    explorerState.fileReadOnly = true;
                    _liveLog('⚠️ Session AKTIV aber SCHREIBGESCHÜTZT');
                    showNotification('Live-Modus aktiv, aber Datei ist schreibgeschützt (wird sie in der Haupt-GUI oder einem anderen Programm verwendet?)', 'warning');
                } else {
                    explorerState.fileReadOnly = false;
                    _liveLog('✓ Session AKTIV (Excel versteckt)');
                    showFloatingStatus(t('liveSessionActive'));
                }
                
                // Bedingte Formatierung erkennen → Paste mit Format blockieren
                if (openResult.hasConditionalFormatting) {
                    explorerState.hasConditionalFormatting = true;
                    _liveLog('⚠️ Sheet hat bedingte Formatierung → Einfügen mit Formatierung deaktiviert');
                } else {
                    explorerState.hasConditionalFormatting = false;
                }
                
                // Live-Mode-Indikator und Excel-Button aktualisieren
                updateLiveModeIndicator();
                updateExcelToggleButton();
                
                return true;
            } catch (error) {
                _liveLog(`EXCEPTION: ${error.message}`);
                console.error('[LiveSession] Stack:', error);
                showFloatingStatus('Live-Session Fehler: ' + error.message, true);
                explorerState.liveSessionActive = false;
                explorerState.liveSessionReady = false;
                return false;
            }
        }
        
        /**
         * Beendet die Live-Session
         */
        async function stopLiveSession() {
            if (!explorerState.liveSessionActive) return;
            
            try {
                await window.electronAPI.liveSessionClose();
                console.log('[LiveSession] Session beendet');
            } catch (error) {
                console.error('[LiveSession] Fehler beim Beenden:', error);
            } finally {
                // Flags IMMER zurücksetzen, auch bei Fehler
                explorerState.liveSessionActive = false;
                explorerState.liveSessionReady = false;
                explorerState.excelVisible = false;
                explorerState.excelInteractive = false;
                updateLiveModeIndicator();
                updateExcelToggleButton();
            }
        }
        
        /**
         * Zeigt eine Warnung an wenn Excel unerwartet geschlossen wurde
         */
        function showExcelClosedWarning() {
            const dialog = document.createElement('div');
            dialog.className = 'modal-overlay';
            dialog.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.6); display: flex; align-items: center; justify-content: center; z-index: 10000;';
            
            dialog.innerHTML = `
                <div style="background: var(--bg-dark); border: 1px solid var(--accent); border-radius: 8px; padding: 24px; max-width: 400px; text-align: center;">
                    <div style="font-size: 48px; margin-bottom: 16px;">⚠️</div>
                    <h3 style="margin: 0 0 12px 0; color: var(--accent);">Excel wurde geschlossen</h3>
                    <p style="margin: 0 0 20px 0; color: var(--text-muted);">
                        Die Excel-Anwendung wurde unerwartet beendet.<br>
                        Möchten Sie die Live-Session neu starten?
                    </p>
                    <div style="display: flex; gap: 12px; justify-content: center;">
                        <button id="btnRestartSession" class="btn btn-primary" style="padding: 8px 20px;">
                            🔄 Neu starten
                        </button>
                        <button id="btnDismissWarning" class="btn btn-secondary" style="padding: 8px 20px;">
                            Schließen
                        </button>
                    </div>
                </div>
            `;
            
            document.body.appendChild(dialog);
            
            document.getElementById('btnRestartSession').onclick = async () => {
                dialog.remove();
                showFloatingStatus('Starte Live-Session neu...', false);
                await startLiveSession();
            };
            
            document.getElementById('btnDismissWarning').onclick = () => {
                dialog.remove();
            };
            
            // Auch bei Klick außerhalb schließen
            dialog.onclick = (e) => {
                if (e.target === dialog) dialog.remove();
            };
        }
        
        /**
         * Schaltet die Sichtbarkeit des Excel-Fensters um
         */
        async function toggleExcelVisibility() {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) {
                showFloatingStatus('Keine Live-Session aktiv', true);
                return;
            }
            // Guard: verhindere parallele Toggle-Aufrufe (Excel COM verträgt
            // rasche Folge-Calls für visible/Interactive schlecht → Deadlock)
            if (explorerState._visibilityTogglePending) {
                showFloatingStatus('⏳ Umschalten läuft bereits…');
                return;
            }
            explorerState._visibilityTogglePending = true;
            
            try {
                const newVisible = !explorerState.excelVisible;
                const result = await window.electronAPI.liveSessionSetVisible(newVisible);
                
                if (result && result.success) {
                    explorerState.excelVisible = newVisible;
                    // Beim Ausblenden: Interactive-Modus immer zurücksetzen
                    if (!newVisible && explorerState.excelInteractive) {
                        explorerState.excelInteractive = false;
                    }
                    updateExcelToggleButton();
                    showFloatingStatus(newVisible ? '👁️ Excel eingeblendet' : '🙈 Excel ausgeblendet');
                } else {
                    showFloatingStatus('Fehler beim Umschalten', true);
                }
            } catch (error) {
                console.error('[LiveSession] Fehler bei toggleExcelVisibility:', error);
                showFloatingStatus('Fehler: ' + error.message, true);
            } finally {
                explorerState._visibilityTogglePending = false;
            }
        }
        
        /**
         * Schaltet Excel temporär in den bedienbaren Modus (Scrollen/Klicken).
         * Achtung: Änderungen die der User direkt in Excel macht, landen NICHT im Datenmodell.
         */
        async function toggleExcelInteractive() {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) {
                showFloatingStatus('Keine Live-Session aktiv', true);
                return;
            }
            if (!explorerState.excelVisible) {
                showFloatingStatus('Excel muss eingeblendet sein', true);
                return;
            }
            if (explorerState._interactiveTogglePending) {
                showFloatingStatus('⏳ Umschalten läuft bereits…');
                return;
            }
            explorerState._interactiveTogglePending = true;
            
            try {
                const newInteractive = !explorerState.excelInteractive;
                const result = await window.electronAPI.liveSessionSetInteractive(newInteractive);
                
                if (result && result.success) {
                    explorerState.excelInteractive = newInteractive;
                    updateExcelToggleButton();
                    if (newInteractive) {
                        showFloatingStatus('🔓 Excel bedienbar – manuelle Änderungen werden NICHT übernommen!', true);
                    } else {
                        showFloatingStatus('🔒 Excel wieder gesperrt');
                    }
                } else {
                    showFloatingStatus('Fehler beim Umschalten: ' + (result?.error || ''), true);
                }
            } catch (error) {
                console.error('[LiveSession] Fehler bei toggleExcelInteractive:', error);
                showFloatingStatus('Fehler: ' + error.message, true);
            } finally {
                explorerState._interactiveTogglePending = false;
            }
        }
        
        /**
         * Aktualisiert den Excel-Toggle-Button basierend auf dem State
         */
        function updateExcelToggleButton() {
            const btn = document.getElementById('btnToggleExcel');
            if (!btn) return;
            
            const isLiveActive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            btn.disabled = !isLiveActive;
            
            if (isLiveActive) {
                if (explorerState.excelVisible) {
                    btn.innerHTML = '🙈 Excel';
                    btn.title = 'Excel-Fenster ausblenden';
                } else {
                    btn.innerHTML = '👁️ Excel';
                    btn.title = 'Excel-Fenster einblenden';
                }
            } else {
                btn.innerHTML = '👁️ Excel';
                btn.title = 'Live-Session nicht aktiv';
            }
            
            // Interactive-Button: nur sichtbar wenn Excel-Fenster eingeblendet
            const btnInt = document.getElementById('btnToggleExcelInteractive');
            if (btnInt) {
                const showInt = isLiveActive && explorerState.excelVisible;
                btnInt.style.display = showInt ? '' : 'none';
                btnInt.disabled = !showInt;
                if (explorerState.excelInteractive) {
                    btnInt.innerHTML = '🔓 Bedienbar';
                    btnInt.title = 'Klicken um Excel wieder zu sperren';
                    btnInt.classList.add('btn-warning');
                    btnInt.classList.remove('btn-secondary');
                } else {
                    btnInt.innerHTML = '🔒 Bedienbar';
                    btnInt.title = 'Excel temporär bedienbar machen (Scrollen / Klicken). Achtung: Manuelle Änderungen werden NICHT übernommen!';
                    btnInt.classList.add('btn-secondary');
                    btnInt.classList.remove('btn-warning');
                }
            }
        }
        
        /**
         * Führt eine Operation in der Live-Session aus (wenn aktiv)
         * @param {string} operation - Name der Operation (z.B. 'deleteRow')
         * @param {Array} args - Argumente für die Operation
         * @returns {Promise<object|null>} Ergebnis oder null wenn nicht aktiv
         */
        async function liveSessionExecute(operation, ...args) {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) {
                return null;
            }
            
            try {
                const methodName = `liveSession${operation.charAt(0).toUpperCase() + operation.slice(1)}`;
                if (typeof window.electronAPI[methodName] === 'function') {
                    const result = await window.electronAPI[methodName](...args);
                    if (!result.success) {
                        console.warn(`[LiveSession] ${operation} fehlgeschlagen:`, result.error);
                    }
                    return result;
                } else {
                    console.warn(`[LiveSession] Unbekannte Operation: ${operation}`);
                    return null;
                }
            } catch (error) {
                console.error(`[LiveSession] ${operation} Fehler:`, error);
                return null;
            }
        }
        
        /**
         * Speichert die Live-Session als neue Datei (Export)
         * @param {string} outputPath - Pfad für die Export-Datei
         * @param {string|null} password - Optionales Passwort
         * @returns {Promise<object>} Ergebnis
         */
        async function liveSessionExport(outputPath, password = null) {
            if (!explorerState.liveSessionActive) {
                return { success: false, error: 'Live-Session nicht aktiv' };
            }
            
            try {
                // Filter vor dem Speichern an Excel senden
                if (explorerState.filters.some(f => f.column && f.value)) {
                    await syncFiltersToExcel();
                }
                
                const result = await window.electronAPI.liveSessionSaveFile(outputPath, password);
                if (result.success) {
                    showFloatingStatus(`✓ Export via Live-Session: ${outputPath.split('/').pop()}`, 'success');
                    // Passwort-Status aktualisieren
                    if (password) {
                        explorerState.filePassword = password;
                    }
                }
                return result;
            } catch (error) {
                console.error('[LiveSession] Export Fehler:', error);
                return { success: false, error: error.message };
            }
        }
        
        /**
         * Aktualisiert die UI-Anzeige des Live-Session-Status
         * Legacy-Funktion - leitet jetzt an updateLiveModeIndicator() weiter
         */
        function updateLiveSessionIndicator() {
            updateLiveModeIndicator();
        }

        const explorerState = {
            filePath: null,
            originalFilePath: null,  // Ursprüngliche Datei (wird nie durch Export überschrieben)
            fileName: null,
            sheets: [],
            selectedSheet: null,
            headers: [],
            data: [],
            originalData: [],  // Kopie der Originaldaten für Vorschau-Vergleich
            filteredData: [],  // Enthält jetzt {originalIndex, row} Objekte
            searchTerm: '',
            filters: [],
            visibleColumns: [],
            columnOrder: [],  // Benutzerdefinierte Spaltenreihenfolge
            editedCells: new Map(),  // Speichert editierte Zellen: "rowIndex-colIndex" -> neuer Wert
            // Pagination
            currentPage: 1,
            pageSize: 100, // Zeilen pro Seite (konfigurierbar)
            pageSizeOptions: [50, 100, 250, 500, 1000],
            // Sortierung
            sortColumn: null,
            sortDirection: null,  // 'asc', 'desc' oder null
            sortType: 'auto',  // 'auto', 'alpha-asc', 'alpha-desc', 'num-asc', 'num-desc', 'date-asc', 'date-desc'
            // Zeilen-Markierungen
            rowHighlights: new Map(),  // rowIndex -> 'green'|'yellow'|'orange'|'red'|'blue'|'purple'
            originalRowHighlights: new Map(),  // Original-Highlights beim Laden (für Erkennung entfernter Markierungen)
            // Zeilen-Auswahl für Verschiebung
            selectedRows: new Set(),  // Set von originalIndex Werten
            moveMode: false,  // Ob Verschiebe-Modus aktiv ist
            // Drag & Drop State
            draggedColumn: null,
            // Zellen-Auswahl für Mehrfach-Bearbeitung
            selectedCells: new Set(),  // Set von "rowIndex-colIndex" Strings
            selectionAnchor: null,  // {row, col} - Startpunkt der Auswahl
            isSelecting: false,  // Ob gerade eine Auswahl gezogen wird
            // Cache für Sheet-Änderungen (bleibt bei Wechsel erhalten)
            sheetDataCache: new Map(),  // sheetName -> { data, editedCells, rowHighlights, originalData }
            // Data Validations (Dropdown-Listen)
            dataValidations: {},  // colIndex -> { type: 'column'|'rows', values: [], rows: {} }
            // Bedingte Formatierung (Sheet-weit)
            hasConditionalFormatting: false,  // true wenn Sheet CF-Regeln hat → Paste mit Format deaktiviert
            // Blattschutz (Sheet-weit) — Hide Spalten/Zeilen wirkt nicht solange aktiv
            sheetProtected: false,
            // Cell Styles (Formatierungen aus Excel)
            cellStyles: {},  // "rowIndex-colIndex" -> { bold, italic, fontColor, fill, fontSize, textAlign, ... }
            // Cell Formulas (Formeln aus Excel)
            cellFormulas: {},  // "rowIndex-colIndex" -> "=FORMULA"
            // Cell Hyperlinks (Links aus Excel)
            cellHyperlinks: {},  // "rowIndex-colIndex" -> "https://..."
            // Rich Text Cells (formatierter Text mit mehreren Styles)
            richTextCells: {},  // "rowIndex-colIndex" -> [{ text, styles: { bold, italic, ... } }, ...]
            // Hidden Rows (ausgeblendete Zeilen)
            hiddenRows: new Set(),  // Set von 0-basierten Zeilen-Indices
            // AutoFilter Range (falls vorhanden)
            autoFilterRange: null,  // z.B. "A1:D10"
            // Merged Cells (verbundene Zellen)
            mergedCells: [],  // Array von { startRow, startCol, endRow, endCol, rowSpan, colSpan }
            // Ausgeblendete Arbeitsblätter
            hiddenSheets: new Set(),  // Set von Sheet-Namen die ausgeblendet sind
            // Passwort für passwortgeschützte Dateien
            filePassword: null,
            // Pivot-Tabellen erkannt (Warnung vor Fallback-Export)
            hasPivotTables: false,
            // Operations Queues für serielle Abarbeitung
            columnOperationsQueue: [],  // [{type: 'delete', originalIndex: X}, {type: 'insert', position: X, headerName: ''}]
            rowOperationsQueue: [],  // [{type: 'delete', originalIndex: X}, {type: 'insert', position: X}]
            // Pending Sheet-Operationen (Offline-Modus: erst beim Export anwenden)
            pendingSheetOperations: [],  // [{type: 'add'|'delete'|'rename'|'clone'|'move'|'visibility', ...params}]
            // Mapping: aktueller Sheet-Name → Original-Name auf Disk (für Disk-Reads nach Umbenennung)
            sheetDiskNameMap: new Map(),
            // LIVE SESSION MODE - Operationen werden sofort in Excel ausgeführt
            liveSheetChanges: 0,       // Zähler für Sheet-Ops im Live-Modus (move/rename/add/delete/visibility)
            liveSessionActive: false,  // true wenn Live-Session aktiv ist
            liveSessionReady: false,   // true wenn Session bereit für Operationen
            excelVisible: false,       // true wenn Excel-Fenster sichtbar ist
            excelInteractive: false,   // true wenn Excel temporär bedienbar (Interactive=True)
            // VM-Map für Bild-Zellen (Copy&Paste von Zell-Bildern)
            cellVmMap: {},  // "styleKeyRow-col" -> vmValue (z.B. "19-6" -> "1")
            // Virtual Scrolling
            virtualRowHeight: 30,     // Geschätzte Zeilenhöhe in px
            virtualBufferRows: 50,    // Puffer-Zeilen ober-/unterhalb des Viewports
            virtualVisibleStart: -1,  // Erster gerenderter Index in filteredData
            virtualVisibleEnd: -1,    // Letzter gerenderter Index in filteredData
            isLoadingSheet: false,      // true während loadExplorerSheet() läuft → verhindert doppelten Wechsel
            pendingSheetSwitch: null    // Wenn während eines Sheet-Wechsels ein neuer Wechsel kommt → merken
        };
        
        // ==================== Data Join State & Functions ====================
        const dataJoinState = {
            sourceFilePath: null,
            sourceFileName: null,
            sourceSheets: [],
            sourceData: [],
            sourceHeaders: [],
            sourceCellStyles: {},
            sourceCellFonts: {},
            sourceNumberFormats: {},
            selectedSourceSheet: null,
            targetKeyColumnIndex: null,
            sourceKeyColumnIndex: null,
            selectedColumns: [],  // Indices der zu kopierenden Spalten
            columnPositions: [],  // Array von {sourceIndex, targetPosition}
            previewCalculated: false,
            matchStats: {
                targetRows: 0,
                sourceRows: 0,
                matches: 0,
                noMatch: 0
            }
        };
        
        function resetDataJoinState() {
            dataJoinState.sourceFilePath = null;
            dataJoinState.sourceFileName = null;
            dataJoinState.sourceSheets = [];
            dataJoinState.sourceData = [];
            dataJoinState.sourceHeaders = [];
            dataJoinState.sourceCellStyles = {};
            dataJoinState.sourceCellFonts = {};
            dataJoinState.sourceNumberFormats = {};
            dataJoinState.selectedSourceSheet = null;
            dataJoinState.targetKeyColumnIndex = null;
            dataJoinState.sourceKeyColumnIndex = null;
            dataJoinState.selectedColumns = [];
            dataJoinState.columnPositions = [];
            dataJoinState.previewCalculated = false;
            dataJoinState.matchStats = { targetRows: 0, sourceRows: 0, matches: 0, noMatch: 0 };
        }
        
        // =====================================================================
        // Serial-Check: Fehlende Seriennummern gegen mehrere Ist-Listen
        // =====================================================================
        const serialCheckState = {
            target: { filePath: null, fileName: null, sheets: [], sheetName: null, headers: [], columnIndex: -1 },
            sources: [] // je: { id, filePath, fileName, sheets, sheetName, headers, columnIndex, label }
        };
        let _scNextId = 1;
        
        function _scNormalize(v) {
            if (v == null) return '';
            let s = String(v).trim().toLowerCase();
            // führende Nullen ignorieren — aber komplett-Null/leer behandeln wir als leer
            s = s.replace(/^0+/, '');
            return s;
        }
        
        function _scSetStatus(msg, kind) {
            const el = document.getElementById('scStatus');
            if (!el) return;
            el.textContent = msg || '';
            el.style.color = kind === 'error' ? 'var(--danger, #d9534f)'
                : kind === 'success' ? 'var(--success, #5cb85c)'
                : 'var(--text-muted)';
        }
        
        async function openSerialCheckModal() {
            // State zurücksetzen
            serialCheckState.target = { filePath: null, fileName: null, sheets: [], sheetName: null, headers: [], columnIndex: -1 };
            serialCheckState.sources = [];
            _scSetStatus('');
            
            document.getElementById('scTargetFileName').textContent = 'Keine Datei gewählt';
            const scSheet = document.getElementById('scTargetSheet');
            const scCol = document.getElementById('scTargetColumn');
            scSheet.innerHTML = ''; scSheet.disabled = true;
            scCol.innerHTML = ''; scCol.disabled = true;
            
            document.getElementById('scSourcesContainer').innerHTML = '';
            
            // Vorbelegen mit aktueller Explorer-Datei, falls vorhanden
            if (explorerState && explorerState.filePath) {
                try {
                    await _scLoadTargetFile(explorerState.filePath);
                    // Wenn aktuelles Sheet existiert → übernehmen
                    if (explorerState.currentSheet && serialCheckState.target.sheets.includes(explorerState.currentSheet)) {
                        document.getElementById('scTargetSheet').value = explorerState.currentSheet;
                        await _scLoadTargetSheet(explorerState.currentSheet);
                    }
                } catch (e) {
                    console.warn('[SerialCheck] Vorbelegen fehlgeschlagen:', e);
                }
            }
            
            // Eine erste Ist-Zeile zufügen
            _scAddSourceRow();
            
            // Modal-Buttons & Handler (idempotent via onclick-Setzer)
            document.getElementById('btnCloseSerialCheck').onclick = closeSerialCheckModal;
            document.getElementById('btnCancelSerialCheck').onclick = closeSerialCheckModal;
            document.getElementById('btnSCTargetFile').onclick = _scPickTargetFile;
            document.getElementById('btnSCAddSource').onclick = _scAddSourceRow;
            document.getElementById('btnRunSerialCheck').onclick = runSerialCheck;
            document.getElementById('scTargetSheet').onchange = async (e) => {
                try {
                    await _scLoadTargetSheet(e.target.value);
                } catch (err) {
                    _scSetStatus('Fehler: ' + (err && err.message ? err.message : err), 'error');
                }
            };
            document.getElementById('scTargetColumn').onchange = (e) => {
                const v = parseInt(e.target.value, 10);
                serialCheckState.target.columnIndex = isNaN(v) ? -1 : v;
            };
            
            document.getElementById('serialCheckModal').classList.remove('hidden');
        }
        
        function closeSerialCheckModal() {
            document.getElementById('serialCheckModal').classList.add('hidden');
        }
        
        async function _scLoadTargetFile(filePath) {
            const res = await window.electronAPI.readExcelFile(filePath);
            if (!res || !res.success) {
                throw new Error(res && res.error ? res.error : 'Datei konnte nicht gelesen werden');
            }
            const sheets = Array.isArray(res.sheets) ? res.sheets : [];
            serialCheckState.target.filePath = filePath;
            serialCheckState.target.fileName = filePath.split(/[\\/]/).pop();
            serialCheckState.target.sheets = sheets;
            serialCheckState.target.sheetName = null;
            serialCheckState.target.headers = [];
            serialCheckState.target.columnIndex = -1;
            
            document.getElementById('scTargetFileName').textContent = serialCheckState.target.fileName;
            const scSheet = document.getElementById('scTargetSheet');
            scSheet.innerHTML = '<option value="">-- Tabellenblatt --</option>' +
                sheets.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
            scSheet.disabled = sheets.length === 0;
            
            const scCol = document.getElementById('scTargetColumn');
            scCol.innerHTML = '';
            scCol.disabled = true;
        }
        
        async function _scLoadTargetSheet(sheetName) {
            if (!sheetName || !serialCheckState.target.filePath) return;
            const res = await window.electronAPI.readExcelSheet(serialCheckState.target.filePath, sheetName, null);
            if (!res || !res.success) {
                throw new Error(res && res.error ? res.error : 'Blatt konnte nicht gelesen werden');
            }
            serialCheckState.target.sheetName = sheetName;
            serialCheckState.target.headers = Array.isArray(res.headers) ? res.headers : [];
            serialCheckState.target.columnIndex = -1;
            
            const scCol = document.getElementById('scTargetColumn');
            scCol.innerHTML = '<option value="">-- Spalte wählen --</option>' +
                serialCheckState.target.headers.map((h, i) => `<option value="${i}">${escapeHtml(String(h ?? ''))}</option>`).join('');
            scCol.disabled = serialCheckState.target.headers.length === 0;
            // Heuristik: erste Spalte mit "seriennummer" / "serial" / "sn" im Namen vorauswählen
            const idx = serialCheckState.target.headers.findIndex(h => /serien|serial|\bsn\b/i.test(String(h ?? '')));
            if (idx >= 0) {
                scCol.value = String(idx);
                serialCheckState.target.columnIndex = idx;
            }
        }
        
        function _scAddSourceRow() {
            const id = _scNextId++;
            const src = { id, filePath: null, fileName: null, sheets: [], sheetName: null, headers: [], columnIndex: -1, label: '' };
            serialCheckState.sources.push(src);
            
            const container = document.getElementById('scSourcesContainer');
            const row = document.createElement('div');
            row.dataset.scSourceId = String(id);
            row.style.cssText = 'display: grid; grid-template-columns: 1.2fr 1.5fr 1fr 1fr auto; gap: 8px; align-items: end; padding: 8px; background: var(--bg-light); border-radius: 4px;';
            row.innerHTML = `
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Abteilung / Label</label>
                    <input type="text" class="form-control sc-label" placeholder="z.B. Abteilung A" style="width: 100%;">
                </div>
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Datei</label>
                    <div style="display: grid; grid-template-columns: 1fr auto; gap: 4px;">
                        <div class="sc-filename" style="font-size: 12px; padding: 6px 8px; background: var(--bg); border: 1px solid var(--border); border-radius: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">Keine Datei</div>
                        <button class="btn btn-primary sc-select-file" style="height: 30px; padding: 0 8px;">📂</button>
                    </div>
                </div>
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Blatt</label>
                    <select class="form-control sc-sheet" disabled style="width: 100%;"></select>
                </div>
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">SN-Spalte</label>
                    <select class="form-control sc-column" disabled style="width: 100%;"></select>
                </div>
                <div>
                    <button class="btn btn-danger sc-remove" title="Entfernen" style="height: 30px; padding: 0 10px;">✕</button>
                </div>
            `;
            container.appendChild(row);
            
            const qs = sel => row.querySelector(sel);
            qs('.sc-label').oninput = e => { src.label = e.target.value; };
            qs('.sc-select-file').onclick = async () => {
                try {
                    const pickedPath = await window.electronAPI.openFileDialog({
                        title: 'Ist-Liste wählen',
                        filters: [{ name: 'Excel-Dateien', extensions: ['xlsx', 'xlsm', 'xls'] }]
                    });
                    if (!pickedPath) return;
                    _scSetStatus('Lade Ist-Datei…');
                    const res = await window.electronAPI.readExcelFile(pickedPath);
                    if (!res || !res.success) throw new Error(res && res.error ? res.error : 'Datei konnte nicht gelesen werden');
                    src.filePath = pickedPath;
                    src.fileName = pickedPath.split(/[\\/]/).pop();
                    src.sheets = Array.isArray(res.sheets) ? res.sheets : [];
                    src.sheetName = null;
                    src.headers = [];
                    src.columnIndex = -1;
                    if (!src.label) {
                        src.label = src.fileName.replace(/\.(xlsx|xlsm|xls)$/i, '');
                        qs('.sc-label').value = src.label;
                    }
                    qs('.sc-filename').textContent = src.fileName;
                    qs('.sc-filename').title = src.filePath;
                    const sheetSel = qs('.sc-sheet');
                    sheetSel.innerHTML = '<option value="">-- Blatt --</option>' +
                        src.sheets.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
                    sheetSel.disabled = src.sheets.length === 0;
                    qs('.sc-column').innerHTML = '';
                    qs('.sc-column').disabled = true;
                    _scSetStatus('');
                } catch (e) {
                    _scSetStatus('Fehler: ' + (e && e.message ? e.message : e), 'error');
                }
            };
            qs('.sc-sheet').onchange = async (e) => {
                const sn = e.target.value;
                if (!sn || !src.filePath) return;
                try {
                    _scSetStatus('Lade Blatt…');
                    const res = await window.electronAPI.readExcelSheet(src.filePath, sn, null);
                    if (!res || !res.success) throw new Error(res && res.error ? res.error : 'Blatt konnte nicht gelesen werden');
                    src.sheetName = sn;
                    src.headers = Array.isArray(res.headers) ? res.headers : [];
                    src.columnIndex = -1;
                    const colSel = qs('.sc-column');
                    colSel.innerHTML = '<option value="">-- Spalte --</option>' +
                        src.headers.map((h, i) => `<option value="${i}">${escapeHtml(String(h ?? ''))}</option>`).join('');
                    colSel.disabled = src.headers.length === 0;
                    const idx = src.headers.findIndex(h => /serien|serial|\bsn\b/i.test(String(h ?? '')));
                    if (idx >= 0) { colSel.value = String(idx); src.columnIndex = idx; }
                    _scSetStatus('');
                } catch (err) {
                    _scSetStatus('Fehler: ' + (err && err.message ? err.message : err), 'error');
                }
            };
            qs('.sc-column').onchange = e => { src.columnIndex = parseInt(e.target.value, 10); if (isNaN(src.columnIndex)) src.columnIndex = -1; };
            qs('.sc-remove').onclick = () => {
                const i = serialCheckState.sources.findIndex(s => s.id === id);
                if (i >= 0) serialCheckState.sources.splice(i, 1);
                row.remove();
            };
        }
        
        async function _scPickTargetFile() {
            const pickedPath = await window.electronAPI.openFileDialog({
                title: 'Soll-Liste wählen',
                filters: [{ name: 'Excel-Dateien', extensions: ['xlsx', 'xlsm', 'xls'] }]
            });
            if (!pickedPath) return;
            try {
                _scSetStatus('Lade Soll-Datei…');
                await _scLoadTargetFile(pickedPath);
                _scSetStatus('');
                // Sheet-Select-Handler (auto-load) wird zentral beim Modal-Open gesetzt
            } catch (e) {
                _scSetStatus('Fehler: ' + (e && e.message ? e.message : e), 'error');
            }
        }
        
        async function runSerialCheck() {
            try {
                // Validierung
                const t = serialCheckState.target;
                if (!t.filePath || !t.sheetName || t.columnIndex < 0) {
                    _scSetStatus('Bitte Soll-Datei, Blatt und Seriennummern-Spalte wählen', 'error');
                    return;
                }
                const srcs = serialCheckState.sources.filter(s => s.filePath && s.sheetName && s.columnIndex >= 0);
                if (srcs.length === 0) {
                    _scSetStatus('Bitte mindestens eine vollständige Ist-Liste angeben', 'error');
                    return;
                }
                
                _scSetStatus('Lade Soll-Daten…');
                const tRes = await window.electronAPI.readExcelSheet(t.filePath, t.sheetName, null);
                if (!tRes || !tRes.success) throw new Error(tRes && tRes.error ? tRes.error : 'Soll-Blatt nicht lesbar');
                const tHeaders = Array.isArray(tRes.headers) ? tRes.headers : [];
                const tData = Array.isArray(tRes.data) ? tRes.data : [];
                
                // Normalisierte SN-Mengen pro Ist-Liste
                const sourceSets = [];
                for (let i = 0; i < srcs.length; i++) {
                    const s = srcs[i];
                    _scSetStatus(`Lade Ist-Liste ${i + 1}/${srcs.length}: ${s.label || s.fileName}…`);
                    const sRes = await window.electronAPI.readExcelSheet(s.filePath, s.sheetName, null);
                    if (!sRes || !sRes.success) throw new Error(`${s.label || s.fileName}: nicht lesbar`);
                    const sData = Array.isArray(sRes.data) ? sRes.data : [];
                    const set = new Set();
                    for (const row of sData) {
                        if (!Array.isArray(row)) continue;
                        const key = _scNormalize(row[s.columnIndex]);
                        if (key) set.add(key);
                    }
                    sourceSets.push({ label: s.label || s.fileName, set });
                }
                
                // Fehlende finden
                _scSetStatus('Vergleiche…');
                const missingRows = [];
                let emptyCount = 0;
                for (const row of tData) {
                    if (!Array.isArray(row)) continue;
                    const raw = row[t.columnIndex];
                    const key = _scNormalize(raw);
                    if (!key) { emptyCount++; continue; }
                    let found = false;
                    for (const src of sourceSets) {
                        if (src.set.has(key)) { found = true; break; }
                    }
                    if (!found) {
                        // Volle Zeile übernehmen, am Ende Info-Spalte
                        const out = tHeaders.map((_, i) => row[i] != null ? row[i] : '');
                        out.push('Nicht in: ' + sourceSets.map(s => s.label).join(', '));
                        missingRows.push(out);
                    }
                }
                
                if (missingRows.length === 0) {
                    _scSetStatus(`✅ Keine fehlenden Seriennummern gefunden (Soll: ${tData.length} Zeilen, davon ${emptyCount} leer).`, 'success');
                    showNotification('Alle Seriennummern der Soll-Liste sind in mindestens einer Ist-Liste vorhanden.', 'success');
                    return;
                }
                
                // Export-Dialog
                const defaultName = (t.fileName || 'SollListe').replace(/\.(xlsx|xlsm|xls)$/i, '') + '_Fehlend.xlsx';
                const savePath = await window.electronAPI.saveFileDialog({
                    title: 'Fehlende Seriennummern speichern',
                    defaultPath: defaultName,
                    filters: [{ name: 'Excel-Dateien', extensions: ['xlsx'] }]
                });
                if (!savePath) {
                    _scSetStatus('Abgebrochen.', 'error');
                    return;
                }
                
                _scSetStatus('Schreibe Export-Datei…');
                const outHeaders = [...tHeaders.map(h => (h == null ? '' : String(h))), 'Status'];
                const writeRes = await window.electronAPI.serialCheckExportXlsx({
                    outputPath: savePath,
                    sheetName: 'Fehlende Seriennummern',
                    headers: outHeaders,
                    rows: missingRows
                });
                if (!writeRes || !writeRes.success) {
                    throw new Error(writeRes && writeRes.error ? writeRes.error : 'Export fehlgeschlagen');
                }
                
                _scSetStatus(`✅ ${missingRows.length} fehlende Seriennummer(n) exportiert.`, 'success');
                showNotification(`${missingRows.length} fehlende Seriennummer(n) nach ${savePath.split(/[\\/]/).pop()} exportiert.`, 'success');
            } catch (e) {
                console.error('[SerialCheck] Fehler:', e);
                _scSetStatus('Fehler: ' + (e && e.message ? e.message : e), 'error');
                showNotification('Fehler beim Abgleich: ' + (e && e.message ? e.message : e), 'error');
            }
        }
        
        // =====================================================================
        // Value-Count: Werte einer gemeinsamen Spalte über mehrere Dateien zählen
        // =====================================================================
        const valueCountState = {
            sources: [],          // { id, filePath, fileName, sheets, sheetName, headers, columnIndex, label }
            results: null,        // { totals: Map<displayValue, {total, perSource: {id: count}}>, fileLabels: [{id,label}] }
            sortMode: 'count-desc' // 'count-desc' | 'count-asc' | 'value-asc' | 'value-desc'
        };
        let _vcNextId = 1;
        
        function _vcNormalize(v, caseInsensitive, trimOn) {
            if (v == null) return '';
            let s = String(v);
            if (trimOn) s = s.trim();
            if (caseInsensitive) s = s.toLowerCase();
            return s;
        }
        
        function _vcSetStatus(msg, kind) {
            const el = document.getElementById('vcStatus');
            if (!el) return;
            el.textContent = msg || '';
            el.style.color = kind === 'error' ? 'var(--danger, #d9534f)'
                : kind === 'success' ? 'var(--success, #5cb85c)'
                : 'var(--text-muted)';
        }
        
        async function openValueCountModal() {
            valueCountState.sources = [];
            valueCountState.results = null;
            valueCountState.sortMode = 'count-desc';
            _vcSetStatus('');
            document.getElementById('vcSourcesContainer').innerHTML = '';
            document.getElementById('vcResultContainer').style.display = 'none';
            document.getElementById('vcResultFilter').value = '';
            const btnExport = document.getElementById('btnExportValueCount');
            if (btnExport) btnExport.disabled = true;
            
            // Vorbelegen mit aktueller Explorer-Datei
            _vcAddSourceRow();
            if (explorerState && explorerState.filePath) {
                try {
                    const first = valueCountState.sources[0];
                    await _vcLoadSourceFile(first, explorerState.filePath);
                    if (explorerState.currentSheet && first.sheets.includes(explorerState.currentSheet)) {
                        await _vcLoadSourceSheet(first, explorerState.currentSheet);
                    }
                } catch (e) {
                    console.warn('[ValueCount] Vorbelegen fehlgeschlagen:', e);
                }
            }
            
            document.getElementById('btnCloseValueCount').onclick = closeValueCountModal;
            document.getElementById('btnCancelValueCount').onclick = closeValueCountModal;
            document.getElementById('btnVCAddSource').onclick = _vcAddSourceRow;
            document.getElementById('btnRunValueCount').onclick = runValueCount;
            document.getElementById('btnExportValueCount').onclick = exportValueCount;
            document.getElementById('vcResultFilter').oninput = _vcRenderResults;
            document.getElementById('btnVCSortValue').onclick = () => {
                valueCountState.sortMode = valueCountState.sortMode === 'value-asc' ? 'value-desc' : 'value-asc';
                _vcRenderResults();
            };
            document.getElementById('btnVCSortCount').onclick = () => {
                valueCountState.sortMode = valueCountState.sortMode === 'count-desc' ? 'count-asc' : 'count-desc';
                _vcRenderResults();
            };
            
            document.getElementById('valueCountModal').classList.remove('hidden');
        }
        
        function closeValueCountModal() {
            document.getElementById('valueCountModal').classList.add('hidden');
        }
        
        async function _vcLoadSourceFile(src, filePath) {
            const res = await window.electronAPI.readExcelFile(filePath);
            if (!res || !res.success) throw new Error(res && res.error ? res.error : 'Datei nicht lesbar');
            src.filePath = filePath;
            src.fileName = filePath.split(/[\\/]/).pop();
            src.sheets = Array.isArray(res.sheets) ? res.sheets : [];
            src.sheetName = null;
            src.headers = [];
            src.columnIndex = -1;
            if (!src.label) src.label = src.fileName.replace(/\.(xlsx|xlsm|xls)$/i, '');
            _vcSyncRowUI(src);
        }
        
        async function _vcLoadSourceSheet(src, sheetName) {
            const res = await window.electronAPI.readExcelSheet(src.filePath, sheetName, null);
            if (!res || !res.success) throw new Error(res && res.error ? res.error : 'Blatt nicht lesbar');
            src.sheetName = sheetName;
            src.headers = Array.isArray(res.headers) ? res.headers : [];
            src.columnIndex = -1;
            _vcSyncRowUI(src);
        }
        
        function _vcSyncRowUI(src) {
            const row = document.querySelector(`#vcSourcesContainer [data-vc-source-id="${src.id}"]`);
            if (!row) return;
            const qs = sel => row.querySelector(sel);
            qs('.vc-label').value = src.label || '';
            qs('.vc-filename').textContent = src.fileName || 'Keine Datei';
            qs('.vc-filename').title = src.filePath || '';
            const sheetSel = qs('.vc-sheet');
            sheetSel.innerHTML = '<option value="">-- Blatt --</option>' +
                src.sheets.map(s => `<option value="${escapeHtml(s)}"${s === src.sheetName ? ' selected' : ''}>${escapeHtml(s)}</option>`).join('');
            sheetSel.disabled = src.sheets.length === 0;
            const colSel = qs('.vc-column');
            colSel.innerHTML = '<option value="">-- Spalte --</option>' +
                src.headers.map((h, i) => `<option value="${i}"${i === src.columnIndex ? ' selected' : ''}>${escapeHtml(String(h ?? ''))}</option>`).join('');
            colSel.disabled = src.headers.length === 0;
        }
        
        function _vcAddSourceRow() {
            const id = _vcNextId++;
            const src = { id, filePath: null, fileName: null, sheets: [], sheetName: null, headers: [], columnIndex: -1, label: '', skipFirstRow: true };
            valueCountState.sources.push(src);
            
            const container = document.getElementById('vcSourcesContainer');
            const row = document.createElement('div');
            row.dataset.vcSourceId = String(id);
            row.style.cssText = 'display: grid; grid-template-columns: 1.2fr 1.5fr 1fr 1.2fr auto auto; gap: 8px; align-items: end; padding: 8px; background: var(--bg-light); border-radius: 4px;';
            row.innerHTML = `
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Label</label>
                    <input type="text" class="form-control vc-label" placeholder="z.B. Bestand A" style="width: 100%;">
                </div>
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Datei</label>
                    <div style="display: grid; grid-template-columns: 1fr auto; gap: 4px;">
                        <div class="vc-filename" style="font-size: 12px; padding: 6px 8px; background: var(--bg); border: 1px solid var(--border); border-radius: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">Keine Datei</div>
                        <button class="btn btn-primary vc-select-file" style="height: 30px; padding: 0 8px;">📂</button>
                    </div>
                </div>
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Blatt</label>
                    <select class="form-control vc-sheet" disabled style="width: 100%;"></select>
                </div>
                <div>
                    <label style="font-size: 11px; color: var(--text-muted);">Spalte</label>
                    <select class="form-control vc-column" disabled style="width: 100%;"></select>
                </div>
                <div title="Erste Zeile ist Überschrift — beim Zählen überspringen">
                    <label style="font-size: 11px; color: var(--text-muted); display: block;">Header</label>
                    <label style="height: 30px; display: flex; align-items: center; gap: 4px; cursor: pointer; font-size: 12px;">
                        <input type="checkbox" class="vc-skip-header green-checkbox" checked>
                        <span>ignorieren</span>
                    </label>
                </div>
                <div>
                    <button class="btn btn-danger vc-remove" title="Entfernen" style="height: 30px; padding: 0 10px;">✕</button>
                </div>
            `;
            container.appendChild(row);
            
            const qs = sel => row.querySelector(sel);
            qs('.vc-label').oninput = e => { src.label = e.target.value; };
            qs('.vc-select-file').onclick = async () => {
                try {
                    const pickedPath = await window.electronAPI.openFileDialog({
                        title: 'Bestandsdatei wählen',
                        filters: [{ name: 'Excel-Dateien', extensions: ['xlsx', 'xlsm', 'xls'] }]
                    });
                    if (!pickedPath) return;
                    _vcSetStatus('Lade Datei…');
                    await _vcLoadSourceFile(src, pickedPath);
                    _vcSetStatus('');
                } catch (e) {
                    _vcSetStatus('Fehler: ' + (e && e.message ? e.message : e), 'error');
                }
            };
            qs('.vc-sheet').onchange = async (e) => {
                const sn = e.target.value;
                if (!sn || !src.filePath) return;
                try {
                    _vcSetStatus('Lade Blatt…');
                    await _vcLoadSourceSheet(src, sn);
                    _vcSetStatus('');
                } catch (err) {
                    _vcSetStatus('Fehler: ' + (err && err.message ? err.message : err), 'error');
                }
            };
            qs('.vc-column').onchange = e => {
                const v = parseInt(e.target.value, 10);
                src.columnIndex = isNaN(v) ? -1 : v;
            };
            qs('.vc-skip-header').onchange = e => {
                src.skipFirstRow = !!e.target.checked;
            };
            qs('.vc-remove').onclick = () => {
                const i = valueCountState.sources.findIndex(s => s.id === id);
                if (i >= 0) valueCountState.sources.splice(i, 1);
                row.remove();
            };
        }
        
        async function runValueCount() {
            try {
                const srcs = valueCountState.sources.filter(s => s.filePath && s.sheetName && s.columnIndex >= 0);
                if (srcs.length === 0) {
                    _vcSetStatus('Bitte mindestens eine vollständige Datei (Datei + Blatt + Spalte) angeben', 'error');
                    return;
                }
                const caseInsensitive = document.getElementById('vcOptCaseInsensitive').checked;
                const trimOn = document.getElementById('vcOptTrim').checked;
                const includeEmpty = document.getElementById('vcOptIncludeEmpty').checked;
                
                // Aggregation: keyNorm -> { display, perSource: Map<id,count>, total }
                const agg = new Map();
                const fileLabels = [];
                let scannedRows = 0;
                
                for (let i = 0; i < srcs.length; i++) {
                    const s = srcs[i];
                    const label = s.label || s.fileName;
                    fileLabels.push({ id: s.id, label });
                    _vcSetStatus(`Lade ${i + 1}/${srcs.length}: ${label}…`);
                    const res = await window.electronAPI.readExcelSheet(s.filePath, s.sheetName, null);
                    if (!res || !res.success) throw new Error(`${label}: nicht lesbar`);
                    const allData = Array.isArray(res.data) ? res.data : [];
                    const data = s.skipFirstRow ? allData.slice(1) : allData;
                    for (const row of data) {
                        if (!Array.isArray(row)) continue;
                        scannedRows++;
                        const raw = row[s.columnIndex];
                        const keyNorm = _vcNormalize(raw, caseInsensitive, trimOn);
                        if (!keyNorm && !includeEmpty) continue;
                        // Anzeige-Wert: erste Fundstelle (nicht-normalisiert, aber ggf. getrimmt)
                        const display = keyNorm === ''
                            ? '(leer)'
                            : (trimOn ? String(raw ?? '').trim() : String(raw ?? ''));
                        let entry = agg.get(keyNorm);
                        if (!entry) {
                            entry = { display, perSource: {}, total: 0 };
                            agg.set(keyNorm, entry);
                        }
                        entry.perSource[s.id] = (entry.perSource[s.id] || 0) + 1;
                        entry.total++;
                    }
                }
                
                valueCountState.results = { agg, fileLabels, scannedRows };
                _vcSetStatus(`✅ ${agg.size} unterschiedliche Werte in ${scannedRows} Zeilen aus ${srcs.length} Datei(en).`, 'success');
                document.getElementById('vcResultContainer').style.display = 'flex';
                document.getElementById('btnExportValueCount').disabled = agg.size === 0;
                _vcRenderResults();
            } catch (e) {
                console.error('[ValueCount] Fehler:', e);
                _vcSetStatus('Fehler: ' + (e && e.message ? e.message : e), 'error');
                showNotification('Fehler beim Auswerten: ' + (e && e.message ? e.message : e), 'error');
            }
        }
        
        function _vcRenderResults() {
            const r = valueCountState.results;
            if (!r) return;
            const head = document.getElementById('vcResultHead');
            const body = document.getElementById('vcResultBody');
            const summary = document.getElementById('vcResultSummary');
            const filter = (document.getElementById('vcResultFilter').value || '').toLowerCase();
            
            // Header
            const fileCols = r.fileLabels.map(f =>
                `<th style="padding: 6px 10px; text-align: right; border-bottom: 1px solid var(--border); white-space: nowrap;" title="${escapeHtml(f.label)}">${escapeHtml(f.label)}</th>`
            ).join('');
            head.innerHTML = `
                <tr>
                    <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid var(--border);">#</th>
                    <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid var(--border);">Wert</th>
                    <th style="padding: 6px 10px; text-align: right; border-bottom: 1px solid var(--border);">Gesamt</th>
                    ${fileCols}
                </tr>
            `;
            
            // Sort + filter
            const rows = Array.from(r.agg.entries()).map(([key, v]) => ({ key, ...v }));
            const sort = valueCountState.sortMode;
            rows.sort((a, b) => {
                if (sort === 'count-desc') return b.total - a.total || a.display.localeCompare(b.display);
                if (sort === 'count-asc') return a.total - b.total || a.display.localeCompare(b.display);
                if (sort === 'value-asc') return a.display.localeCompare(b.display);
                if (sort === 'value-desc') return b.display.localeCompare(a.display);
                return 0;
            });
            const filtered = filter
                ? rows.filter(x => x.display.toLowerCase().includes(filter))
                : rows;
            
            const totalSum = filtered.reduce((a, x) => a + x.total, 0);
            summary.textContent = `${filtered.length} / ${rows.length} Werte · ${totalSum} Treffer`;
            
            // Body
            body.innerHTML = filtered.map((x, i) => {
                const perCols = r.fileLabels.map(f => {
                    const c = x.perSource[f.id] || 0;
                    const muted = c === 0 ? ' style="color: var(--text-muted);"' : '';
                    return `<td style="padding: 5px 10px; text-align: right; border-bottom: 1px solid var(--border);"${muted}>${c}</td>`;
                }).join('');
                return `
                    <tr style="${i % 2 ? 'background: rgba(255,255,255,0.02);' : ''}">
                        <td style="padding: 5px 10px; color: var(--text-muted); border-bottom: 1px solid var(--border);">${i + 1}</td>
                        <td style="padding: 5px 10px; border-bottom: 1px solid var(--border); word-break: break-word;">${escapeHtml(x.display)}</td>
                        <td style="padding: 5px 10px; text-align: right; font-weight: 600; border-bottom: 1px solid var(--border);">${x.total}</td>
                        ${perCols}
                    </tr>
                `;
            }).join('');
        }
        
        async function exportValueCount() {
            try {
                const r = valueCountState.results;
                if (!r || r.agg.size === 0) return;
                const defaultName = 'Werte_Zaehlung.xlsx';
                const savePath = await window.electronAPI.saveFileDialog({
                    title: 'Wert-Zählung exportieren',
                    defaultPath: defaultName,
                    filters: [{ name: 'Excel-Dateien', extensions: ['xlsx'] }]
                });
                if (!savePath) return;
                
                const headers = ['Wert', 'Gesamt', ...r.fileLabels.map(f => f.label)];
                const rows = Array.from(r.agg.values())
                    .sort((a, b) => b.total - a.total || a.display.localeCompare(b.display))
                    .map(v => [v.display, v.total, ...r.fileLabels.map(f => v.perSource[f.id] || 0)]);
                
                _vcSetStatus('Schreibe Export…');
                const writeRes = await window.electronAPI.serialCheckExportXlsx({
                    outputPath: savePath,
                    sheetName: 'Werte-Zählung',
                    headers,
                    rows
                });
                if (!writeRes || !writeRes.success) {
                    throw new Error(writeRes && writeRes.error ? writeRes.error : 'Export fehlgeschlagen');
                }
                _vcSetStatus(`✅ ${rows.length} Werte nach ${savePath.split(/[\\/]/).pop()} exportiert.`, 'success');
                showNotification(`Export erfolgreich: ${rows.length} Werte`, 'success');
            } catch (e) {
                console.error('[ValueCount] Export-Fehler:', e);
                _vcSetStatus('Export-Fehler: ' + (e && e.message ? e.message : e), 'error');
                showNotification('Export-Fehler: ' + (e && e.message ? e.message : e), 'error');
            }
        }
        
        function openDataJoinModal() {
            if (!explorerState.filePath || explorerState.data.length === 0) {
                showNotification('Bitte zuerst eine Datei im Datenexplorer laden', 'warning');
                return;
            }
            
            resetDataJoinState();
            
            // Ziel-Spalten (aktuelle Datei) befüllen
            const targetKeySelect = elements.joinTargetKeyColumn;
            targetKeySelect.innerHTML = `<option value="">${t('joinSelectColumn')}</option>`;
            explorerState.headers.forEach((header, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = `${getColumnLetter(index + 1)}: ${header || '(leer)'}`;
                targetKeySelect.appendChild(option);
            });
            targetKeySelect.disabled = false;
            
            // Reset andere Felder
            elements.joinSourceFileName.textContent = t('joinNoFileSelected');
            elements.joinSourceSheet.innerHTML = `<option value="">${t('joinLoadFile')}</option>`;
            elements.joinSourceSheet.disabled = true;
            elements.joinSourceKeyColumn.innerHTML = `<option value="">${t('joinSelectColumn')}</option>`;
            elements.joinSourceKeyColumn.disabled = true;
            elements.joinColumnsContainer.innerHTML = `<div style="color: var(--text-muted); font-size: 13px; text-align: center; padding: 20px;">${t('joinLoadSourceFirst')}</div>`;
            elements.joinPreviewContainer.style.display = 'none';
            elements.btnPreviewDataJoin.disabled = true;
            elements.btnExecuteDataJoin.disabled = true;
            
            // Modal öffnen
            elements.dataJoinModal.classList.remove('hidden');
            document.body.classList.add('modal-open');
            
            // Button-Farbe umschalten (aktiv)
            const btn = document.getElementById('btnDataJoin');
            if (btn) {
                btn.classList.remove('btn-primary');
                btn.classList.add('btn-info');
            }
        }
        
        function closeDataJoinModal() {
            elements.dataJoinModal.classList.add('hidden');
            document.body.classList.remove('modal-open');
            
            // Button-Farbe zurücksetzen (inaktiv)
            const btn = document.getElementById('btnDataJoin');
            if (btn) {
                btn.classList.remove('btn-info');
                btn.classList.add('btn-primary');
            }
        }
        
        // Hilfsfunktion: Quelldatei laden (von Pfad)
        async function loadDataJoinSourceFromPath(filePath) {
            try {
                const fileName = filePath.split(/[/\\]/).pop();
                
                // Datei laden um Sheets zu bekommen
                const fileResult = await window.electronAPI.readExcelFile(filePath);
                
                if (!fileResult.success) {
                    showNotification('Fehler beim Laden: ' + fileResult.error, 'error');
                    return;
                }
                
                dataJoinState.sourceFilePath = filePath;
                dataJoinState.sourceFileName = fileName;
                dataJoinState.sourceSheets = fileResult.sheets;
                
                elements.joinSourceFileName.textContent = fileName;
                elements.joinSourceFileName.style.color = 'var(--success)';
                
                // Sheets-Dropdown befüllen
                elements.joinSourceSheet.innerHTML = '';
                fileResult.sheets.forEach(sheetName => {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    elements.joinSourceSheet.appendChild(option);
                });
                elements.joinSourceSheet.disabled = false;
                
                // Erstes Sheet automatisch laden
                if (fileResult.sheets.length > 0) {
                    await loadDataJoinSourceSheet(fileResult.sheets[0]);
                }
                
            } catch (error) {
                console.error('Fehler beim Laden der Quelldatei:', error);
                showNotification('Fehler beim Laden: ' + error.message, 'error');
            }
        }
        
        async function loadDataJoinSourceFile() {
            try {
                const filePath = await window.electronAPI.openFileDialog({
                    title: 'Quelldatei für Spalten-Join auswählen',
                    filters: [{ name: 'Excel-Dateien', extensions: ['xlsx', 'xls'] }],
                    defaultPath: getWorkingDirectoryPath()
                });
                
                if (!filePath) {
                    return;
                }
                
                await loadDataJoinSourceFromPath(filePath);
                
            } catch (error) {
                console.error('Fehler beim Laden der Quelldatei:', error);
                showNotification('Fehler beim Laden: ' + error.message, 'error');
            }
        }
        
        async function loadDataJoinSourceSheet(sheetName) {
            if (!dataJoinState.sourceFilePath || !sheetName) return;
            
            try {
                const sheetResult = await window.electronAPI.readExcelSheet(
                    dataJoinState.sourceFilePath, 
                    sheetName, 
                    null
                );
                
                if (!sheetResult.success) {
                    showNotification('Fehler beim Laden des Sheets: ' + sheetResult.error, 'error');
                    return;
                }
                
                dataJoinState.selectedSourceSheet = sheetName;
                dataJoinState.sourceHeaders = sheetResult.headers || [];
                dataJoinState.sourceData = sheetResult.data || [];
                dataJoinState.sourceCellStyles = sheetResult.cellStyles || {};
                dataJoinState.sourceCellFonts = sheetResult.cellFonts || {};
                dataJoinState.sourceNumberFormats = sheetResult.numberFormats || {};
                
                // Schlüsselspalten-Dropdown befüllen
                const sourceKeySelect = elements.joinSourceKeyColumn;
                sourceKeySelect.innerHTML = `<option value="">${t('joinSelectColumn')}</option>`;
                dataJoinState.sourceHeaders.forEach((header, index) => {
                    const option = document.createElement('option');
                    option.value = index;
                    option.textContent = `${getColumnLetter(index + 1)}: ${header || '(leer)'}`;
                    sourceKeySelect.appendChild(option);
                });
                sourceKeySelect.disabled = false;
                
                // Spalten-Auswahl befüllen
                updateJoinColumnsSelection();
                
                // Preview zurücksetzen
                dataJoinState.previewCalculated = false;
                elements.joinPreviewContainer.style.display = 'none';
                updateJoinButtons();
                
            } catch (error) {
                console.error('Fehler beim Laden des Source-Sheets:', error);
                showNotification('Fehler beim Laden: ' + error.message, 'error');
            }
        }
        
        function updateJoinColumnsSelection() {
            const container = elements.joinColumnsContainer;
            container.innerHTML = '';
            
            if (dataJoinState.sourceHeaders.length === 0) {
                container.innerHTML = '<div style="color: var(--text-muted); font-size: 13px; text-align: center; padding: 20px;">Keine Spalten gefunden</div>';
                return;
            }
            
            // Erstelle Optionen für Ziel-Positionen
            const targetPositionOptions = explorerState.headers.map((h, idx) => 
                `<option value="${idx}">nach ${getColumnLetter(idx + 1)}: ${h || '(leer)'}</option>`
            ).join('');
            
            dataJoinState.sourceHeaders.forEach((header, index) => {
                const row = document.createElement('div');
                row.className = 'join-column-row';
                row.style.cssText = 'display: flex; align-items: center; gap: 10px; padding: 8px 12px; background: var(--bg-light); border-radius: 4px; transition: background 0.2s;';
                row.innerHTML = `
                    <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex: 0 0 auto;">
                        <input type="checkbox" class="green-checkbox join-column-checkbox" data-index="${index}">
                        <span style="font-weight: 500; color: var(--primary-light); min-width: 30px;">${getColumnLetter(index + 1)}</span>
                        <span style="color: var(--text); min-width: 120px; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${header || '(leer)'}">${header || '(leer)'}</span>
                    </label>
                    <span style="color: var(--text-muted); font-size: 12px;">→</span>
                    <select class="config-select join-target-position" data-source-index="${index}" style="flex: 1; min-width: 150px; font-size: 12px;" disabled>
                        <option value="end">Am Ende einfügen</option>
                        ${targetPositionOptions}
                    </select>
                `;
                row.onmouseenter = () => row.style.background = 'var(--bg-lighter)';
                row.onmouseleave = () => row.style.background = 'var(--bg-light)';
                
                const checkbox = row.querySelector('.join-column-checkbox');
                const positionSelect = row.querySelector('.join-target-position');
                
                checkbox.onchange = () => {
                    positionSelect.disabled = !checkbox.checked;
                    refreshPositionDropdowns();
                    updateSelectedJoinColumns();
                    dataJoinState.previewCalculated = false;
                    elements.joinPreviewContainer.style.display = 'none';
                    updateJoinButtons();
                };
                
                positionSelect.onchange = () => {
                    refreshPositionDropdowns();
                    updateSelectedJoinColumns();
                    dataJoinState.previewCalculated = false;
                    elements.joinPreviewContainer.style.display = 'none';
                };
                
                container.appendChild(row);
            });
        }
        
        function updateSelectedJoinColumns() {
            const checkboxes = document.querySelectorAll('.join-column-checkbox:checked');
            
            // Speichere ausgewählte Spalten mit ihren Zielpositionen
            dataJoinState.selectedColumns = [];
            dataJoinState.columnPositions = []; // Array von {sourceIndex, targetPosition}
            
            checkboxes.forEach(cb => {
                const sourceIndex = parseInt(cb.dataset.index);
                const positionSelect = document.querySelector(`.join-target-position[data-source-index="${sourceIndex}"]`);
                const rawValue = positionSelect ? positionSelect.value : 'end';
                
                // Parse zusammengesetzte Werte: "end", "end+3", "2", "2+5" → Basis-Position extrahieren
                const basePos = rawValue.split('+')[0];
                const targetPosition = basePos === 'end' ? 'end' : parseInt(basePos);
                
                dataJoinState.selectedColumns.push(sourceIndex);
                dataJoinState.columnPositions.push({
                    sourceIndex: sourceIndex,
                    targetPosition: targetPosition
                });
            });
        }
        
        /**
         * Aktualisiert die Positions-Dropdowns aller Spalten, um bereits ausgewählte
         * neue Spalten in der Positionsliste anzuzeigen.
         */
        function refreshPositionDropdowns() {
            // Sammle alle aktuell ausgewählten Quellspalten und ihre Positionen
            const selectedCols = [];
            document.querySelectorAll('.join-column-checkbox:checked').forEach(cb => {
                const sourceIndex = parseInt(cb.dataset.index);
                const posSelect = document.querySelector(`.join-target-position[data-source-index="${sourceIndex}"]`);
                if (!posSelect) return;
                const rawValue = posSelect.value;
                const basePos = rawValue.split('+')[0]; // "end" oder z.B. "3"
                const name = dataJoinState.sourceHeaders[sourceIndex] || `Spalte ${getColumnLetter(sourceIndex + 1)}`;
                selectedCols.push({ sourceIndex, basePos, name });
            });
            
            // Erstelle virtuelle Header-Liste: Original-Header + ausgewählte neue Spalten
            let virtualItems = explorerState.headers.map((h, idx) => ({
                label: h || '(leer)',
                value: String(idx),
                isNew: false
            }));
            
            // Gruppiere ausgewählte Spalten nach Basis-Position
            const groups = new Map();
            selectedCols.forEach(sc => {
                if (!groups.has(sc.basePos)) groups.set(sc.basePos, []);
                groups.get(sc.basePos).push(sc);
            });
            
            // Positionierte Spalten einfügen (von hinten nach vorne für korrekte Indizes)
            const numericKeys = Array.from(groups.keys())
                .filter(k => k !== 'end')
                .map(Number)
                .sort((a, b) => b - a);
            
            numericKeys.forEach(pos => {
                const cols = groups.get(String(pos));
                const insertAt = Math.min(pos + 1, virtualItems.length);
                // Umgekehrt einfügen, damit erste Spalte oben steht
                [...cols].reverse().forEach(sc => {
                    virtualItems.splice(insertAt, 0, {
                        label: `${sc.name} (neu)`,
                        value: `${pos}+${sc.sourceIndex}`,
                        isNew: true,
                        sourceIndex: sc.sourceIndex
                    });
                });
            });
            
            // "Am Ende"-Spalten anhängen
            if (groups.has('end')) {
                groups.get('end').forEach(sc => {
                    virtualItems.push({
                        label: `${sc.name} (neu)`,
                        value: `end+${sc.sourceIndex}`,
                        isNew: true,
                        sourceIndex: sc.sourceIndex
                    });
                });
            }
            
            // Aktualisiere alle Position-Dropdowns
            document.querySelectorAll('.join-target-position').forEach(select => {
                const currentSourceIndex = parseInt(select.dataset.sourceIndex);
                const currentValue = select.value;
                
                let optionsHtml = '<option value="end">Am Ende einfügen</option>';
                
                virtualItems.forEach((item, vIdx) => {
                    // Eigenen Eintrag überspringen
                    if (item.isNew && item.sourceIndex === currentSourceIndex) return;
                    
                    const letter = getColumnLetter(vIdx + 1);
                    const style = item.isNew ? 'color: #4ec9b0; font-style: italic;' : '';
                    optionsHtml += `<option value="${item.value}" style="${style}">nach ${letter}: ${item.label}</option>`;
                });
                
                select.innerHTML = optionsHtml;
                
                // Auswahl wiederherstellen
                let found = false;
                for (const opt of select.options) {
                    if (opt.value === currentValue) {
                        select.value = currentValue;
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    // Fallback: Basis-Position versuchen
                    const basePos = currentValue.split('+')[0];
                    for (const opt of select.options) {
                        if (opt.value === basePos) {
                            select.value = basePos;
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        select.value = 'end';
                    }
                }
            });
        }
        
        function updateJoinButtons() {
            const hasTargetKey = elements.joinTargetKeyColumn.value !== '';
            const hasSourceKey = elements.joinSourceKeyColumn.value !== '';
            const hasSelectedColumns = dataJoinState.selectedColumns.length > 0;
            
            elements.btnPreviewDataJoin.disabled = !(hasTargetKey && hasSourceKey && hasSelectedColumns);
            elements.btnExecuteDataJoin.disabled = !dataJoinState.previewCalculated;
        }
        
        function calculateDataJoinPreview() {
            const targetKeyIndex = parseInt(elements.joinTargetKeyColumn.value);
            const sourceKeyIndex = parseInt(elements.joinSourceKeyColumn.value);
            
            if (isNaN(targetKeyIndex) || isNaN(sourceKeyIndex)) {
                showNotification('Bitte Schlüsselspalten auswählen', 'warning');
                return;
            }
            
            dataJoinState.targetKeyColumnIndex = targetKeyIndex;
            dataJoinState.sourceKeyColumnIndex = sourceKeyIndex;
            
            // Source-Daten in Map für schnellen Lookup
            const sourceMap = new Map();
            dataJoinState.sourceData.forEach((row, idx) => {
                const keyValue = String(row[sourceKeyIndex] || '').trim();
                if (keyValue) {
                    sourceMap.set(keyValue, row);
                }
            });
            
            // Matches zählen
            let matches = 0;
            let noMatch = 0;
            
            explorerState.data.forEach((row, idx) => {
                const keyValue = String(row[targetKeyIndex] || '').trim();
                if (keyValue && sourceMap.has(keyValue)) {
                    matches++;
                } else if (keyValue) {
                    noMatch++;
                }
            });
            
            // Stats speichern und anzeigen
            dataJoinState.matchStats = {
                targetRows: explorerState.data.length,
                sourceRows: dataJoinState.sourceData.length,
                matches: matches,
                noMatch: noMatch
            };
            
            elements.joinStatTargetRows.textContent = dataJoinState.matchStats.targetRows;
            elements.joinStatSourceRows.textContent = dataJoinState.matchStats.sourceRows;
            elements.joinStatMatches.textContent = dataJoinState.matchStats.matches;
            elements.joinStatNoMatch.textContent = dataJoinState.matchStats.noMatch;
            
            elements.joinPreviewContainer.style.display = 'block';
            dataJoinState.previewCalculated = true;
            updateJoinButtons();
            
            showNotification(`Vorschau berechnet: ${matches} Matches, ${noMatch} ohne Match`, 'success');
        }
        
        async function executeDataJoin() {
            if (!dataJoinState.previewCalculated) {
                showNotification('Bitte zuerst Vorschau berechnen', 'warning');
                return;
            }
            
            const originalTargetKeyIndex = dataJoinState.targetKeyColumnIndex;
            const sourceKeyIndex = dataJoinState.sourceKeyColumnIndex;
            const columnPositions = dataJoinState.columnPositions || [];
            const markNotFound = elements.joinMarkNotFound.checked;
            
            if (columnPositions.length === 0) {
                showNotification('Bitte mindestens eine Spalte auswählen', 'warning');
                return;
            }
            
            // Source-Daten in Map für schnellen Lookup
            const sourceMap = new Map();
            dataJoinState.sourceData.forEach(row => {
                const keyValue = String(row[sourceKeyIndex] || '').trim();
                if (keyValue) {
                    sourceMap.set(keyValue, row);
                }
            });
            
            // Gruppiere Spalten nach Zielposition und sortiere absteigend
            // (von hinten nach vorne einfügen, damit Indices nicht durcheinander kommen)
            const positionGroups = new Map();
            
            columnPositions.forEach(cp => {
                const pos = cp.targetPosition === 'end' ? Infinity : cp.targetPosition + 1; // +1 weil "nach Spalte X"
                if (!positionGroups.has(pos)) {
                    positionGroups.set(pos, []);
                }
                positionGroups.get(pos).push(cp.sourceIndex);
            });
            
            // Sortiere Positionen absteigend (höchste zuerst)
            const sortedPositions = Array.from(positionGroups.keys()).sort((a, b) => b - a);
            
            // Strukturelle Änderungen für Export sammeln
            const insertOperations = [];
            let totalInsertCount = 0;
            
            // Berechne Original-Positionen BEVOR wir etwas einfügen
            // (für den Export - die Positionen müssen die Original-Indizes sein)
            const originalHeaderCount = explorerState.headers.length;
            const originalPositions = new Map();
            sortedPositions.forEach(insertPos => {
                const sourceIndices = positionGroups.get(insertPos);
                const count = sourceIndices.length;
                // Original-Position ist entweder die konkrete Position oder "am Ende"
                const origPos = insertPos === Infinity ? originalHeaderCount : insertPos;
                originalPositions.set(insertPos, origPos);
            });
            
            // Dynamischer Key-Index: Wird angepasst wenn Spalten VOR der Schlüsselspalte eingefügt werden
            let currentKeyIndex = originalTargetKeyIndex;
            
            // Für jede Position (von hinten nach vorne)
            sortedPositions.forEach(insertPos => {
                const sourceIndices = positionGroups.get(insertPos);
                
                // Tatsächliche Einfügeposition (berücksichtige Infinity für "am Ende")
                let actualInsertPos = insertPos === Infinity ? explorerState.headers.length : insertPos;
                
                // Neue Header für diese Position - mit Duplikatprüfung
                const newHeaders = sourceIndices.map(idx => {
                    let headerName = dataJoinState.sourceHeaders[idx] || `Spalte ${getColumnLetter(idx + 1)}`;
                    let finalName = headerName;
                    let counter = 2;
                    // Prüfe ob Name bereits existiert (in explorerState.headers oder bereits hinzugefügten newHeaders)
                    while (explorerState.headers.includes(finalName)) {
                        finalName = `${headerName}_${counter}`;
                        counter++;
                    }
                    return finalName;
                });
                
                // Header einfügen
                explorerState.headers.splice(actualInsertPos, 0, ...newHeaders);
                
                // Visible Columns aktualisieren
                explorerState.visibleColumns = explorerState.visibleColumns.map(idx => 
                    idx >= actualInsertPos ? idx + newHeaders.length : idx
                );
                for (let i = 0; i < newHeaders.length; i++) {
                    let arrayPos = 0;
                    while (arrayPos < explorerState.visibleColumns.length && 
                           explorerState.visibleColumns[arrayPos] < actualInsertPos + i) {
                        arrayPos++;
                    }
                    explorerState.visibleColumns.splice(arrayPos, 0, actualInsertPos + i);
                }
                
                // Column Order aktualisieren wenn vorhanden
                if (explorerState.columnOrder.length > 0) {
                    explorerState.columnOrder = explorerState.columnOrder.map(idx => 
                        idx >= actualInsertPos ? idx + newHeaders.length : idx
                    );
                    for (let i = 0; i < newHeaders.length; i++) {
                        let arrayPos = 0;
                        while (arrayPos < explorerState.columnOrder.length && 
                               explorerState.columnOrder[arrayPos] < actualInsertPos + i) {
                            arrayPos++;
                        }
                        explorerState.columnOrder.splice(arrayPos, 0, actualInsertPos + i);
                    }
                }
                
                // WICHTIG: ZUERST cellStyles etc. verschieben, BEVOR wir neue Daten einfügen!
                // Sonst werden die gerade eingefügten editedCells auch verschoben.
                const shiftColumnIndices = (obj) => {
                    const newObj = {};
                    for (const [key, value] of Object.entries(obj)) {
                        const [rowStr, colStr] = key.split('-');
                        const colIdx = parseInt(colStr);
                        if (colIdx >= actualInsertPos) {
                            newObj[`${rowStr}-${colIdx + newHeaders.length}`] = value;
                        } else {
                            newObj[key] = value;
                        }
                    }
                    return newObj;
                };
                
                explorerState.cellStyles = shiftColumnIndices(explorerState.cellStyles);
                explorerState.cellFormulas = shiftColumnIndices(explorerState.cellFormulas);
                explorerState.cellHyperlinks = shiftColumnIndices(explorerState.cellHyperlinks);
                explorerState.richTextCells = shiftColumnIndices(explorerState.richTextCells);
                
                // EditedCells: Nur bestehende Keys verschieben (die von vorherigen Einfügungen)
                // Muss VOR dem Setzen neuer Werte passieren!
                const newEditedCells = new Map();
                explorerState.editedCells.forEach((value, key) => {
                    if (key.startsWith('_')) {
                        newEditedCells.set(key, value);
                        return;
                    }
                    const parts = key.split('-');
                    if (parts.length !== 2) {
                        newEditedCells.set(key, value);
                        return;
                    }
                    const colIdx = parseInt(parts[1]);
                    if (isNaN(colIdx)) {
                        newEditedCells.set(key, value);
                        return;
                    }
                    if (colIdx >= actualInsertPos) {
                        newEditedCells.set(`${parts[0]}-${colIdx + newHeaders.length}`, value);
                    } else {
                        newEditedCells.set(key, value);
                    }
                });
                explorerState.editedCells = newEditedCells;
                
                // Daten für jede Zeile hinzufügen
                // WICHTIG: currentKeyIndex verwenden, da sich der Index nach jeder Einfügung verschieben kann
                explorerState.data.forEach((row, rowIndex) => {
                    const keyValue = String(row[currentKeyIndex] || '').trim();
                    const sourceRow = sourceMap.get(keyValue);
                    
                    const newValues = sourceIndices.map(colIdx => {
                        if (sourceRow) {
                            return sourceRow[colIdx] || '';
                        } else {
                            return '';
                        }
                    });
                    
                    row.splice(actualInsertPos, 0, ...newValues);
                    
                    // Als bearbeitet markieren - Keys sind jetzt korrekt weil Verschiebung schon passiert ist
                    newValues.forEach((val, i) => {
                        const cellKey = `${rowIndex}-${actualInsertPos + i}`;
                        explorerState.editedCells.set(cellKey, val);
                    });
                });
                
                // Original-Daten auch aktualisieren
                explorerState.originalData.forEach((row, rowIndex) => {
                    const keyValue = String(row[currentKeyIndex] || '').trim();
                    const sourceRow = sourceMap.get(keyValue);
                    
                    const newValues = sourceIndices.map(colIdx => {
                        return sourceRow ? (sourceRow[colIdx] || '') : '';
                    });
                    
                    row.splice(actualInsertPos, 0, ...newValues);
                });
                
                // Source-Styles für die neuen Spalten übernehmen
                // Wir müssen die Source-Zeilen-Indizes zu Target-Zeilen matchen
                const sourceRowIndexMap = new Map(); // keyValue -> sourceRowIndex
                dataJoinState.sourceData.forEach((row, srcIdx) => {
                    const keyValue = String(row[sourceKeyIndex] || '').trim();
                    if (keyValue && !sourceRowIndexMap.has(keyValue)) {
                        sourceRowIndexMap.set(keyValue, srcIdx);
                    }
                });
                
                // Für jede Target-Zeile die Styles der gematchten Source-Zeile übernehmen
                explorerState.data.forEach((row, targetRowIdx) => {
                    const keyValue = String(row[currentKeyIndex] || '').trim();
                    const sourceRowIdx = sourceRowIndexMap.get(keyValue);
                    
                    if (sourceRowIdx !== undefined) {
                        // Für jede eingefügte Spalte
                        sourceIndices.forEach((srcColIdx, i) => {
                            const targetColIdx = actualInsertPos + i;
                            // Source-Key (beachte: Source-Daten sind 0-basiert ohne Header)
                            // +1 weil Header in Source Zeile 0 ist, Daten ab Zeile 1
                            const srcStyleKey = `${sourceRowIdx + 1}-${srcColIdx}`;
                            // Target-Key: +1 weil Rendering mit rowIndex+1 arbeitet (1-basiert)
                            const targetStyleKey = `${targetRowIdx + 1}-${targetColIdx}`;
                            
                            // CellStyles (Hintergrundfarbe)
                            if (dataJoinState.sourceCellStyles[srcStyleKey]) {
                                explorerState.cellStyles[targetStyleKey] = dataJoinState.sourceCellStyles[srcStyleKey];
                                explorerState.editedCells.set('_hasFormatChanges', true);
                            }
                            
                            // CellFonts (Schriftart)
                            if (dataJoinState.sourceCellFonts[srcStyleKey]) {
                                explorerState.cellFonts = explorerState.cellFonts || {};
                                explorerState.cellFonts[targetStyleKey] = dataJoinState.sourceCellFonts[srcStyleKey];
                            }
                            
                            // NumberFormats (Zahlenformate)
                            if (dataJoinState.sourceNumberFormats[srcStyleKey]) {
                                explorerState.numberFormats = explorerState.numberFormats || {};
                                explorerState.numberFormats[targetStyleKey] = dataJoinState.sourceNumberFormats[srcStyleKey];
                            }
                        });
                    }
                });
                
                // Auch Header-Styles übernehmen (Zeile -1 im Source entspricht den Headers)
                sourceIndices.forEach((srcColIdx, i) => {
                    const targetColIdx = actualInsertPos + i;
                    // Header ist in Source Zeile 0 (0-basiert für Style-Keys)
                    // Aber unser headerRowIdx ist -1 weil wir die Header separat haben
                    // In der Quelldatei ist die Header-Zeile Index 0 im Style-System
                    // Aber da sourceData ab Zeile 2 anfängt, ist der Header-Style-Key "-1-colIdx"
                    // Tatsächlich: die Styles haben Zeile 0 = Header
                    const srcHeaderStyleKey = `0-${srcColIdx}`;
                    
                    // In explorerState sind Headers nicht in cellStyles
                    // Wir müssen prüfen ob headerStyles existiert, sonst erstellen
                    explorerState.headerStyles = explorerState.headerStyles || {};
                    if (dataJoinState.sourceCellStyles[srcHeaderStyleKey]) {
                        explorerState.headerStyles[targetColIdx] = dataJoinState.sourceCellStyles[srcHeaderStyleKey];
                    }
                });
                
                // DataValidations anpassen (Keys sind Spalten-Indizes)
                const newValidations = {};
                for (const [colStr, validation] of Object.entries(explorerState.dataValidations)) {
                    const colIdx = parseInt(colStr);
                    if (colIdx >= actualInsertPos) {
                        newValidations[colIdx + newHeaders.length] = validation;
                    } else {
                        newValidations[colIdx] = validation;
                    }
                }
                explorerState.dataValidations = newValidations;
                
                // Key-Index anpassen wenn Spalten VOR der Schlüsselspalte eingefügt wurden
                if (actualInsertPos <= currentKeyIndex) {
                    currentKeyIndex += newHeaders.length;
                }
                
                // Operation für Export speichern mit ORIGINAL-Position (vor allen Einfügungen)
                insertOperations.push({
                    position: originalPositions.get(insertPos),
                    count: newHeaders.length,
                    headers: newHeaders
                });
                
                totalInsertCount += newHeaders.length;
            });
            
            // Strukturelle Änderung markieren für korrekten Export (mit allen Operationen)
            // Sortiere aufsteigend nach Position für den Export
            insertOperations.sort((a, b) => a.position - b.position);
            
            // Snapshot der ORIGINAL-Positionen für die Live-Session:
            // Python's data_join_sync verwaltet den insert_offset selbst und
            // erwartet daher Original-Positionen. Ohne diesen Snapshot würde
            // der Offset doppelt addiert → 2. Spalte landet eine Position zu weit rechts.
            const liveOperations = insertOperations.map(op => ({
                position: op.position,
                count: op.count,
                headers: op.headers
            }));
            
            // Original-Positionen → FINALE Positionen umrechnen (kumulativer Offset)
            // Backend (_build_col_map_for_insert, nur-Export-Pfad) erwartet FINALE Positionen.
            let cumulativeInsertOffset = 0;
            for (const op of insertOperations) {
                op.position += cumulativeInsertOffset;
                cumulativeInsertOffset += op.count;
            }
            
            explorerState.editedCells.set('_columnInserted', {
                operations: insertOperations,
                totalCount: totalInsertCount
            });
            
            // Markierung für nicht gefundene Zeilen
            // WICHTIG: currentKeyIndex enthält den korrekten Index nach allen Einfügungen!
            if (markNotFound) {
                explorerState.data.forEach((row, rowIndex) => {
                    const keyValue = String(row[currentKeyIndex] || '').trim();
                    if (keyValue && !sourceMap.has(keyValue)) {
                        explorerState.rowHighlights.set(rowIndex, 'yellow');
                    }
                });
            }
            
            // UI SOFORT aktualisieren (vor Live-Session-Sync)
            filterExplorerData();
            closeDataJoinModal();
            
            // Matches zählen
            let matchCount = 0;
            explorerState.data.forEach(row => {
                const keyValue = String(row[currentKeyIndex] || '').trim();
                if (keyValue && sourceMap.has(keyValue)) matchCount++;
            });
            
            showNotification(
                `✓ ${totalInsertCount} Spalte(n) hinzugefügt! ${matchCount} von ${explorerState.data.length} Zeilen mit Daten gefüllt.`,
                'success',
                5000
            );
            
            // Live-Session: Spalten in Excel im Hintergrund synchronisieren
            // WICHTIG: AWAIT statt fire-and-forget, damit beim Speichern
            // die Spalten bereits in Excel eingefügt sind!
            // liveOperations enthält ORIGINAL-Positionen (Python addiert insert_offset selbst)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    await _syncDataJoinToLiveSession(liveOperations, explorerState, sourceMap, currentKeyIndex, markNotFound);
                } catch (err) {
                    console.error('[LiveSession] DataJoin sync error:', err);
                    showFloatingStatus('⚠️ Excel-Sync fehlgeschlagen — beim Speichern wird erneut versucht', 'warning');
                }
            }
        }
        
        /**
         * Synchronisiert DataJoin-Ergebnis im Hintergrund mit der Live-Session.
         * Läuft asynchron NACH der UI-Aktualisierung.
         */
        async function _syncDataJoinToLiveSession(insertOperations, state, sourceMap, currentKeyIndex, markNotFound) {
            const sortedOps = insertOperations.sort((a, b) => a.position - b.position);
            
            const totalCols = insertOperations.reduce((sum, op) => sum + op.count, 0);
            showFloatingStatus(`🔄 Synchronisiere ${totalCols} Spalte${totalCols > 1 ? 'n' : ''} mit Excel...`, 'info');
            
            // Batch-Operationen vorbereiten: Spalten + Daten in EINEM Aufruf
            const batchOps = [];
            let insertOffset = 0;
            
            for (const op of sortedOps) {
                const actualInsertPos = op.position + insertOffset;
                
                // Spaltendaten sammeln
                const columnData = [];
                for (let i = 0; i < op.count; i++) {
                    // Daten aus state.data extrahieren (UI wurde bereits aktualisiert)
                    const colValues = state.data.map(row => {
                        const val = row[actualInsertPos + i];
                        return val !== undefined && val !== null ? val : '';
                    });
                    columnData.push(colValues);
                }
                
                // WICHTIG: Original-Position senden, Python verwaltet insert_offset selbst
                batchOps.push({
                    position: op.position,
                    count: op.count,
                    headers: op.headers,
                    columnData: columnData
                });
                
                insertOffset += op.count;
            }
            
            // Alles in EINEM Aufruf an Python/Excel senden
            console.log('[LiveSession] DataJoin batchOps:', batchOps.length, 'ops,',
                'columnData lengths:', batchOps.map(o => (o.columnData || []).map(cd => cd.length)),
                'headers:', batchOps.map(o => o.headers),
                'positions:', batchOps.map(o => o.position));
            if (batchOps.length > 0 && batchOps[0].columnData && batchOps[0].columnData[0]) {
                const sample = batchOps[0].columnData[0].filter(v => v !== '').slice(0, 5);
                console.log('[LiveSession] Sample non-empty values:', sample);
            }
            const result = await window.electronAPI.liveSessionDataJoinSync(batchOps);
            console.log('[LiveSession] DataJoin result:', JSON.stringify(result));
            
            if (!result || !result.success) {
                const errMsg = result ? (result.error || JSON.stringify(result)) : 'Kein Ergebnis';
                console.error('[LiveSession] DataJoin sync failed:', errMsg, result);
                showFloatingStatus('⚠️ Excel-Sync fehlgeschlagen: ' + errMsg, 'warning');
                return;
            }
            
            // Nicht gefundene Zeilen markieren (Batch)
            if (markNotFound) {
                const notFoundRows = [];
                state.data.forEach((row, rowIndex) => {
                    const keyValue = String(row[currentKeyIndex] || '').trim();
                    if (keyValue && !sourceMap.has(keyValue)) {
                        notFoundRows.push(rowIndex);
                    }
                });
                
                if (notFoundRows.length > 0) {
                    showFloatingStatus(`🔄 Markiere ${notFoundRows.length} nicht gefundene Zeilen...`, 'info');
                    const excelRows = notFoundRows.map(r => getExcelRowPosition(r));
                    await window.electronAPI.liveSessionHighlightRowsBatch(excelRows, 'yellow');
                    console.log('[LiveSession] Batch-Highlight:', notFoundRows.length, 'Zeilen gelb markiert');
                }
            }
            
            showFloatingStatus(`✓ Excel synchronisiert — ${totalCols} Spalte${totalCols > 1 ? 'n' : ''} übertragen`, 'success');
        }
        
        // ==================== Sheet Management Functions ====================
        let selectedSheetForManagement = null;
        
        function openSheetManageModal() {
            if (!explorerState.filePath || explorerState.sheets.length === 0) {
                showNotification('Bitte zuerst eine Datei laden', 'warning');
                return;
            }
            
            selectedSheetForManagement = null;
            updateSheetManageList();
            updateSheetManageButtons();
            
            // Hinweis je nach Modus anpassen
            const hint = document.getElementById('sheetManageHint');
            if (hint) {
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    hint.innerHTML = '<strong>Hinweis:</strong> Änderungen werden direkt in Excel ausgeführt.';
                } else {
                    hint.innerHTML = '<strong>Hinweis:</strong> Änderungen werden sofort in der Datei gespeichert.';
                }
            }
            
            document.getElementById('sheetManageModal').classList.remove('hidden');
        }
        
        function closeSheetManageModal() {
            document.getElementById('sheetManageModal').classList.add('hidden');
        }
        
        function updateSheetManageList() {
            const listContainer = document.getElementById('sheetManageList');
            if (!listContainer) return;
            
            if (explorerState.sheets.length === 0) {
                listContainer.innerHTML = '<div style="padding: 20px; text-align: center; color: var(--text-muted);">Keine Arbeitsblätter vorhanden</div>';
                return;
            }
            
            listContainer.innerHTML = explorerState.sheets.map((sheetName, index) => {
                const isSelected = sheetName === selectedSheetForManagement;
                const isActive = sheetName === explorerState.selectedSheet;
                const isHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(sheetName);
                return `<div class="sheet-list-item${isSelected ? ' selected' : ''}" data-sheet="${escapeHtml(sheetName)}" data-index="${index}">
                    <span class="sheet-index">${index + 1}.</span>
                    <span class="sheet-name">${escapeHtml(sheetName)}</span>
                    ${isHidden ? '<span class="sheet-hidden-badge">Ausgeblendet</span>' : ''}
                    ${isActive ? '<span class="sheet-active-badge">Aktiv</span>' : ''}
                </div>`;
            }).join('');
            
            // Event-Listener für Klick auf Sheet-Items
            listContainer.querySelectorAll('.sheet-list-item').forEach(item => {
                item.addEventListener('click', () => {
                    selectedSheetForManagement = item.dataset.sheet;
                    updateSheetManageList();
                    updateSheetManageButtons();
                });
            });
        }
        
        function updateSheetManageButtons() {
            const hasSelection = selectedSheetForManagement !== null;
            const selectedIndex = explorerState.sheets.indexOf(selectedSheetForManagement);
            const isFirst = selectedIndex === 0;
            const isLast = selectedIndex === explorerState.sheets.length - 1;
            const canDelete = explorerState.sheets.length > 1;
            const isLive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            
            document.getElementById('btnSheetRename').disabled = !hasSelection;
            document.getElementById('btnSheetClone').disabled = !hasSelection;
            document.getElementById('btnSheetDelete').disabled = !hasSelection || !canDelete;
            document.getElementById('btnSheetMoveUp').disabled = !hasSelection || isFirst;
            document.getElementById('btnSheetMoveDown').disabled = !hasSelection || isLast;
            
            // Visibility-Toggle auch im Offline-Modus aktivierbar
            const btnToggle = document.getElementById('btnSheetToggleVisibility');
            if (btnToggle) {
                btnToggle.disabled = !hasSelection;
                if (hasSelection && explorerState.hiddenSheets && explorerState.hiddenSheets.has(selectedSheetForManagement)) {
                    btnToggle.innerHTML = '👁️ Einblenden';
                    btnToggle.title = 'Ausgewähltes Arbeitsblatt einblenden';
                } else {
                    btnToggle.innerHTML = '👁️‍🗨️ Ausblenden';
                    btnToggle.title = 'Ausgewähltes Arbeitsblatt ausblenden';
                }
            }
        }
        
        async function addNewSheet() {
            const name = await showPromptDialog('Neues Arbeitsblatt', 'Name für das neue Arbeitsblatt:', 'Neues Blatt');
            if (!name) return;
            
            try {
                // Prüfe ob Name bereits existiert
                if (explorerState.sheets.includes(name)) {
                    showNotification('Ein Arbeitsblatt mit diesem Namen existiert bereits', 'error');
                    return;
                }
                
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    // Live-Session: direkt in Excel
                    const result = await window.electronAPI.addSheet({
                        filePath: explorerState.filePath,
                        sheetName: name
                    });
                    if (result.success) {
                        explorerState.sheets = result.sheets;
                        explorerState.liveSheetChanges++;
                    } else {
                        showNotification(result.error, 'error');
                        return;
                    }
                } else {
                    // Offline: nur im Speicher
                    explorerState.sheets.push(name);
                    explorerState.pendingSheetOperations.push({ type: 'add', sheetName: name });
                    // Leeres Sheet im Cache anlegen
                    explorerState.sheetDataCache.set(name, {
                        headers: [],
                        data: [],
                        originalData: [],
                        editedCells: new Map(),
                        rowHighlights: new Map(),
                        visibleColumns: [],
                        columnOrder: [],
                        cellStyles: {},
                        cellFormulas: {},
                        cellHyperlinks: {},
                        richTextCells: {},
                        hiddenRows: new Set(),
                        autoFilterRange: null,
                        mergedCells: [],
                        dataValidations: {}
                    });
                }
                
                updateSheetDropdown();
                updateSheetManageList();
                showNotification(`Arbeitsblatt "${name}" hinzugefügt`, 'success');
                
                // Automatisch zum neuen Sheet wechseln
                await loadExplorerSheet(name);
                updateSheetManageList();
            } catch (error) {
                showNotification('Fehler beim Hinzufügen: ' + error.message, 'error');
            }
        }
        
        async function renameSelectedSheet() {
            if (!selectedSheetForManagement) return;
            
            const newName = await showPromptDialog('Arbeitsblatt umbenennen', `Neuer Name für "${selectedSheetForManagement}":`, selectedSheetForManagement);
            if (!newName || newName === selectedSheetForManagement) return;
            
            try {
                // Prüfe ob neuer Name bereits existiert
                if (explorerState.sheets.includes(newName)) {
                    showNotification('Ein Arbeitsblatt mit diesem Namen existiert bereits', 'error');
                    return;
                }
                
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    // Live-Session: direkt in Excel
                    const result = await window.electronAPI.renameSheet({
                        filePath: explorerState.filePath,
                        oldName: selectedSheetForManagement,
                        newName: newName
                    });
                    if (!result.success) {
                        showNotification(result.error, 'error');
                        return;
                    }
                    explorerState.sheets = result.sheets;
                    explorerState.liveSheetChanges++;
                } else {
                    // Offline: nur im Speicher
                    const idx = explorerState.sheets.indexOf(selectedSheetForManagement);
                    if (idx >= 0) explorerState.sheets[idx] = newName;
                    explorerState.pendingSheetOperations.push({ type: 'rename', oldName: selectedSheetForManagement, newName: newName });
                    // Cache-Key wird unten zentral umgehängt (für beide Modi)
                    // hiddenSheets aktualisieren
                    if (explorerState.hiddenSheets.has(selectedSheetForManagement)) {
                        explorerState.hiddenSheets.delete(selectedSheetForManagement);
                        explorerState.hiddenSheets.add(newName);
                    }
                }
                
                // Cache-Key umhängen (für beide Modi: live & offline)
                const cachedEntry = explorerState.sheetDataCache.get(selectedSheetForManagement);
                if (cachedEntry) {
                    explorerState.sheetDataCache.delete(selectedSheetForManagement);
                    explorerState.sheetDataCache.set(newName, cachedEntry);
                }
                
                // Disk-Name-Mapping aktualisieren (für beide Modi: live & offline)
                // Falls das Sheet bereits vorher umbenannt wurde, den Original-Disk-Namen weiterverwenden
                const diskName = explorerState.sheetDiskNameMap.get(selectedSheetForManagement) || selectedSheetForManagement;
                explorerState.sheetDiskNameMap.delete(selectedSheetForManagement);
                explorerState.sheetDiskNameMap.set(newName, diskName);
                
                const wasActive = explorerState.selectedSheet === selectedSheetForManagement;
                if (wasActive) {
                    explorerState.selectedSheet = newName;
                    // Aktive Sheet-Daten sofort im Cache sichern (auch ohne Edits),
                    // damit beim Zurückwechseln die Daten nicht von Disk gelesen werden müssen
                    if (explorerState.data.length > 0) {
                        saveCurrentSheetToCache();
                    }
                }
                selectedSheetForManagement = newName;
                updateSheetDropdown();
                updateSheetManageList();
                showNotification(`Arbeitsblatt umbenannt zu "${newName}"`, 'success');
            } catch (error) {
                showNotification('Fehler beim Umbenennen: ' + error.message, 'error');
            }
        }
        
        async function cloneSelectedSheet() {
            if (!selectedSheetForManagement) return;
            
            const newName = await showPromptDialog('Arbeitsblatt kopieren', `Name für die Kopie von "${selectedSheetForManagement}":`, selectedSheetForManagement + ' (Kopie)');
            if (!newName) return;
            
            try {
                // Prüfe ob Name bereits existiert
                if (explorerState.sheets.includes(newName)) {
                    showNotification('Ein Arbeitsblatt mit diesem Namen existiert bereits', 'error');
                    return;
                }
                
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    // Live-Session: direkt in Excel
                    const result = await window.electronAPI.cloneSheet({
                        filePath: explorerState.filePath,
                        sheetName: selectedSheetForManagement,
                        newName: newName
                    });
                    if (!result.success) {
                        showNotification(result.error, 'error');
                        return;
                    }
                    explorerState.sheets = result.sheets;
                    explorerState.liveSheetChanges++;
                } else {
                    // Offline: nur im Speicher
                    // Sheet hinter dem Original einfügen
                    const srcIdx = explorerState.sheets.indexOf(selectedSheetForManagement);
                    explorerState.sheets.splice(srcIdx + 1, 0, newName);
                    explorerState.pendingSheetOperations.push({ type: 'clone', sourceSheet: selectedSheetForManagement, newName: newName });
                    // Cache-Daten kopieren (falls vorhanden)
                    const srcCache = explorerState.sheetDataCache.get(selectedSheetForManagement);
                    if (srcCache) {
                        explorerState.sheetDataCache.set(newName, {
                            headers: [...srcCache.headers],
                            data: srcCache.data.map(row => [...row]),
                            originalData: srcCache.originalData.map(row => [...row]),
                            editedCells: new Map(),
                            rowHighlights: new Map(srcCache.rowHighlights),
                            visibleColumns: [...(srcCache.visibleColumns || [])],
                            columnOrder: [...(srcCache.columnOrder || [])],
                            cellStyles: { ...srcCache.cellStyles },
                            cellFormulas: { ...srcCache.cellFormulas },
                            cellHyperlinks: { ...srcCache.cellHyperlinks },
                            richTextCells: { ...srcCache.richTextCells },
                            hiddenRows: new Set(srcCache.hiddenRows || []),
                            autoFilterRange: srcCache.autoFilterRange,
                            mergedCells: [...(srcCache.mergedCells || [])],
                            dataValidations: { ...srcCache.dataValidations }
                        });
                    }
                    // Sichtbarkeit kopieren
                    if (explorerState.hiddenSheets.has(selectedSheetForManagement)) {
                        explorerState.hiddenSheets.add(newName);
                    }
                }
                
                updateSheetDropdown();
                updateSheetManageList();
                showNotification(`Arbeitsblatt kopiert als "${newName}"`, 'success');
                
                // Automatisch zum kopierten Sheet wechseln
                await loadExplorerSheet(newName);
                updateSheetManageList();
            } catch (error) {
                showNotification('Fehler beim Kopieren: ' + error.message, 'error');
            }
        }
        
        async function moveSelectedSheet(direction) {
            if (!selectedSheetForManagement) return;
            
            const currentIndex = explorerState.sheets.indexOf(selectedSheetForManagement);
            const newIndex = direction === 'up' ? currentIndex - 1 : currentIndex + 1;
            
            if (newIndex < 0 || newIndex >= explorerState.sheets.length) return;
            
            try {
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    // Live-Session: direkt in Excel
                    const result = await window.electronAPI.moveSheet({
                        filePath: explorerState.filePath,
                        sheetName: selectedSheetForManagement,
                        newIndex: newIndex
                    });
                    if (!result.success) {
                        showNotification(result.error, 'error');
                        return;
                    }
                    explorerState.sheets = result.sheets;
                    explorerState.liveSheetChanges++;
                } else {
                    // Offline: nur im Speicher
                    explorerState.sheets.splice(currentIndex, 1);
                    explorerState.sheets.splice(newIndex, 0, selectedSheetForManagement);
                    explorerState.pendingSheetOperations.push({ type: 'move', sheetName: selectedSheetForManagement, newIndex: newIndex });
                }
                
                updateSheetDropdown();
                updateSheetManageList();
                updateSheetManageButtons();
                showNotification(`Arbeitsblatt verschoben`, 'success');
            } catch (error) {
                showNotification('Fehler beim Verschieben: ' + error.message, 'error');
            }
        }
        
        async function deleteSelectedSheet() {
            if (!selectedSheetForManagement) return;
            if (explorerState.sheets.length <= 1) {
                showNotification('Das letzte Arbeitsblatt kann nicht gelöscht werden', 'warning');
                return;
            }
            
            const confirmed = await showConfirmDialog('Arbeitsblatt löschen', `Möchten Sie das Arbeitsblatt "${selectedSheetForManagement}" wirklich löschen?\n\nDiese Aktion kann nicht rückgängig gemacht werden!`);
            if (!confirmed) return;
            
            try {
                const wasActive = explorerState.selectedSheet === selectedSheetForManagement;
                
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    // Live-Session: direkt in Excel
                    const result = await window.electronAPI.deleteSheet({
                        filePath: explorerState.filePath,
                        sheetName: selectedSheetForManagement
                    });
                    if (!result.success) {
                        showNotification(result.error, 'error');
                        return;
                    }
                    explorerState.sheets = result.sheets;
                    explorerState.liveSheetChanges++;
                } else {
                    // Offline: nur im Speicher
                    explorerState.sheets = explorerState.sheets.filter(s => s !== selectedSheetForManagement);
                    explorerState.pendingSheetOperations.push({ type: 'delete', sheetName: selectedSheetForManagement });
                    // Aus Cache entfernen
                    explorerState.sheetDataCache.delete(selectedSheetForManagement);
                    explorerState.hiddenSheets.delete(selectedSheetForManagement);
                }
                
                selectedSheetForManagement = null;
                
                // Wenn das aktive Sheet gelöscht wurde, das erste laden
                if (wasActive && explorerState.sheets.length > 0) {
                    // selectedSheet auf null setzen, damit loadExplorerSheet
                    // die alten Daten des gelöschten Sheets NICHT im Cache unter
                    // dem Namen des Ziel-Sheets speichert (saveCurrentSheetToCache-Guard)
                    explorerState.selectedSheet = null;
                    await loadExplorerSheet(explorerState.sheets[0]);
                }
                
                updateSheetDropdown();
                updateSheetManageList();
                updateSheetManageButtons();
                showNotification(`Arbeitsblatt gelöscht`, 'success');
            } catch (error) {
                showNotification('Fehler beim Löschen: ' + error.message, 'error');
            }
        }
        
        async function toggleSheetVisibility() {
            if (!selectedSheetForManagement) return;
            
            const isCurrentlyHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(selectedSheetForManagement);
            const newVisible = isCurrentlyHidden; // toggle: hidden → visible, visible → hidden
            
            // Prüfe ob es das letzte sichtbare Sheet wäre
            if (!newVisible) {
                const visibleCount = explorerState.sheets.filter(s => !explorerState.hiddenSheets.has(s)).length;
                if (visibleCount <= 1) {
                    showNotification('Mindestens ein Arbeitsblatt muss sichtbar bleiben', 'warning');
                    return;
                }
            }
            
            try {
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    // Live-Session: über xlwings
                    const result = await window.electronAPI.liveSessionSetSheetVisibility(selectedSheetForManagement, newVisible);
                    if (!result.success) {
                        showNotification(result.error, 'error');
                        return;
                    }
                    explorerState.liveSheetChanges++;
                } else {
                    // Offline: nur im Speicher (wird beim Export angewendet)
                    explorerState.pendingSheetOperations.push({ type: 'visibility', sheetName: selectedSheetForManagement, visible: newVisible });
                }
                
                if (newVisible) {
                    explorerState.hiddenSheets.delete(selectedSheetForManagement);
                    showNotification(`Arbeitsblatt "${selectedSheetForManagement}" eingeblendet`, 'success');
                } else {
                    explorerState.hiddenSheets.add(selectedSheetForManagement);
                    showNotification(`Arbeitsblatt "${selectedSheetForManagement}" ausgeblendet`, 'success');
                }
                updateSheetDropdown();
                updateSheetManageList();
                updateSheetManageButtons();
                
                // Bei Einblenden: zum eingeblendeten Sheet wechseln
                if (newVisible) {
                    await loadExplorerSheet(selectedSheetForManagement);
                    updateSheetManageList();
                }
            } catch (error) {
                showNotification('Fehler beim Ein-/Ausblenden: ' + error.message, 'error');
            }
        }
        
        function updateSheetDropdown() {
            if (elements.explorerSheetSelect) {
                elements.explorerSheetSelect.innerHTML = explorerState.sheets.map(s => {
                    const isHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(s);
                    const label = isHidden ? `👁️‍🗨️ ${s} (ausgeblendet)` : s;
                    return `<option value="${escapeHtml(s)}"${s === explorerState.selectedSheet ? ' selected' : ''}>${escapeHtml(label)}</option>`;
                }).join('');
            }
        }
        
        // Explorer-State komplett zurücksetzen
        function resetExplorerState() {
            // Session-State zurücksetzen (verhindert stale engineMode nach Close)
            explorerState.liveSessionActive = false;
            explorerState.liveSessionReady = false;
            explorerState.excelVisible = false;
            explorerState.excelInteractive = false;
            explorerState.engineMode = 'openpyxl';
            explorerState.fileReadOnly = false;
            
            explorerState.filePath = null;
            explorerState.fileName = null;
            explorerState.sheets = [];
            explorerState.selectedSheet = null;
            explorerState.headers = [];
            explorerState.data = [];
            explorerState.originalData = [];
            explorerState.filteredData = [];
            explorerState.searchTerm = '';
            explorerState.filters = [];
            explorerState.visibleColumns = [];
            explorerState.columnOrder = [];
            explorerState.editedCells.clear();
            explorerState.currentPage = 1;
            explorerState.sortColumn = null;
            explorerState.sortDirection = null;
            explorerState.sortType = 'auto';
            explorerState.rowHighlights.clear();
            explorerState.selectedRows.clear();
            explorerState.moveMode = false;
            explorerState.draggedColumn = null;
            explorerState.selectedCells.clear();
            explorerState.selectionAnchor = null;
            explorerState.isSelecting = false;
            explorerState.sheetDataCache.clear();
            explorerState.dataValidations = {};
            explorerState.cellStyles = {};
            explorerState.cellFormulas = {};

            explorerState.cellHyperlinks = {};
            explorerState.richTextCells = {};
            explorerState.hiddenRows.clear();
            explorerState.autoFilterRange = null;
            explorerState.mergedCells = [];
            explorerState.hiddenSheets = new Set();
            explorerState.filePassword = null;
            explorerState.rowMapping = null;  // Mapping: neue Position -> Original Excel-Zeile
            explorerState.columnOperationsQueue = [];  // Reset Queue
            explorerState.rowOperationsQueue = [];  // Reset Queue
            explorerState.pendingSheetOperations = [];  // Reset Pending Sheet Ops
            explorerState.liveSheetChanges = 0;  // Reset Live-Sheet-Änderungen
            explorerState.sheetDiskNameMap = new Map();  // Reset Disk-Name-Mapping
            // Preload-Token invalidieren (stoppt laufenden Hintergrund-Preload)
            if (explorerState._preloadToken) {
                explorerState._preloadToken.cancelled = true;
            }
            explorerState._preloadToken = null;
            
            // UI zurücksetzen
            if (elements.explorerSheetSelect) {
                elements.explorerSheetSelect.innerHTML = '<option value="">-- Sheet wählen --</option>';
            }
            if (elements.explorerSearch) {
                elements.explorerSearch.value = '';
            }
            const explorerFileName = document.getElementById('explorerFileName');
            if (explorerFileName) {
                explorerFileName.textContent = t('noFileLoaded');
            }
            const btnFileInfo = document.getElementById('btnFileInfo');
            if (btnFileInfo) btnFileInfo.style.display = 'none';
            // Nur thead und tbody leeren, nicht die gesamte Tabelle (sonst werden die Element-Referenzen ungültig)
            if (elements.explorerTableHead) {
                elements.explorerTableHead.innerHTML = '';
            }
            if (elements.explorerTableBody) {
                elements.explorerTableBody.innerHTML = '';
            }
            // Drop-Zone wieder anzeigen
            showExplorerDropZone(true);
            
            const explorerResultCount = document.getElementById('explorerResultCount');
            if (explorerResultCount) {
                explorerResultCount.textContent = currentLanguage === 'en' ? 'No data loaded.' : 'Keine Daten geladen.';
            }
            
            // Indikatoren zurücksetzen (verstecken statt entfernen)
            const hiddenRowsIndicator = document.getElementById('hiddenRowsIndicator');
            if (hiddenRowsIndicator) hiddenRowsIndicator.remove();
            const hiddenColumnsIndicator = document.getElementById('hiddenColumnsIndicator');
            if (hiddenColumnsIndicator) hiddenColumnsIndicator.remove();
            const autoFilterIndicator = document.getElementById('autoFilterIndicator');
            if (autoFilterIndicator) autoFilterIndicator.style.display = 'none';
            const passwordIndicator = document.getElementById('passwordIndicator');
            if (passwordIndicator) passwordIndicator.remove();
            
            // Filter-UI zurücksetzen
            const explorerFiltersEl = document.getElementById('explorerFilters');
            if (explorerFiltersEl) {
                explorerFiltersEl.innerHTML = '';
                explorerFiltersEl.style.display = 'flex'; // Filter-Bereich wieder einblenden (falls eingeklappt)
            }
            const btnClearFilters = document.getElementById('btnClearExplorerFilters');
            if (btnClearFilters) btnClearFilters.disabled = true;
            // Filter-Badge und Toggle-Icon zurücksetzen
            const filterCountBadge = document.getElementById('filterCountBadge');
            if (filterCountBadge) filterCountBadge.style.display = 'none';
            const filterToggleIcon = document.getElementById('filterToggleIcon');
            if (filterToggleIcon) filterToggleIcon.textContent = '▼';
            
            // Spalten-Panel zurücksetzen und schließen
            const columnToggles = document.getElementById('columnToggles');
            if (columnToggles) columnToggles.innerHTML = '';
            const columnTogglePanel = document.getElementById('columnTogglePanel');
            if (columnTogglePanel) columnTogglePanel.style.display = 'none';
            
            // Filter-Panel (explorerFilterControls) zurücksetzen und schließen
            const filterPanel = document.getElementById('explorerFilterControls');
            if (filterPanel) filterPanel.style.display = 'none';
            
            // Ersetzen-Panel zurücksetzen und schließen
            const findReplacePanel = document.getElementById('findReplacePanel');
            if (findReplacePanel) findReplacePanel.style.display = 'none';
            findReplaceState.isOpen = false;
            
            // Button-Farben auf Standard (dunkelgrün) zurücksetzen
            const btnFilter = document.getElementById('btnToggleFilterPanel');
            if (btnFilter) { btnFilter.classList.remove('btn-info'); btnFilter.classList.add('btn-primary'); }
            const btnColumns = document.getElementById('btnToggleColumns');
            if (btnColumns) { btnColumns.classList.remove('btn-info'); btnColumns.classList.add('btn-primary'); }
            const btnFindReplace = document.getElementById('btnToggleFindReplace');
            if (btnFindReplace) { btnFindReplace.classList.remove('btn-info'); btnFindReplace.classList.add('btn-primary'); }
            const btnDataJoin = document.getElementById('btnDataJoin');
            if (btnDataJoin) { btnDataJoin.classList.remove('btn-info'); btnDataJoin.classList.add('btn-primary'); }
            
            // Status zurücksetzen
            if (elements.explorerStatus) {
                elements.explorerStatus.textContent = '';
                elements.explorerStatus.style.color = '#ff9800';
            }
            
            // Pagination zurücksetzen
            const explorerPagination = document.getElementById('explorerPagination');
            if (explorerPagination) explorerPagination.style.display = 'none';
            
            // Auto-Save Intervall stoppen
            stopExplorerAutoSave();
            
            // Suchen & Ersetzen zurücksetzen
            resetFindReplaceState();
        }
        
        // Auto-Save für Crash-Recovery starten
        // Intervall: 60s (vorher 30s) — bei großen Dateien reduziert das die Last
        const EXPLORER_AUTOSAVE_INTERVAL_MS = 60000;
        function startExplorerAutoSave() {
            if (explorerAutoSaveInterval) {
                clearInterval(explorerAutoSaveInterval);
            }
            explorerAutoSaveInterval = setInterval(saveExplorerRecoveryData, EXPLORER_AUTOSAVE_INTERVAL_MS);
            // Sofort einmal speichern
            saveExplorerRecoveryData();
        }
        
        // Auto-Save stoppen
        function stopExplorerAutoSave() {
            if (explorerAutoSaveInterval) {
                clearInterval(explorerAutoSaveInterval);
                explorerAutoSaveInterval = null;
            }
            _lastRecoveryFingerprint = '';
        }
        
        // Recovery-Daten speichern (DIFF-ONLY ab v2.0)
        // Speichert nur die Änderungen (editedCells, rowHighlights, pendingOps),
        // NICHT die kompletten Datensätze. Bei Wiederherstellung wird die Datei neu
        // geladen und die Diffs werden auf die Sheets angewendet.
        let _recoveryIdleHandle = null;
        let _lastRecoveryFingerprint = '';
        function saveExplorerRecoveryData() {
            if (!explorerState.filePath) {
                return;
            }

            // Schnelle Prüfung: Gibt es überhaupt Änderungen?
            const totalChanges = countAllChanges();
            const hasPendingOps = explorerState.pendingSheetOperations.length > 0;
            if (totalChanges === 0 && !hasPendingOps) {
                if (_lastRecoveryFingerprint !== '') {
                    localStorage.removeItem(EXPLORER_RECOVERY_KEY);
                    _lastRecoveryFingerprint = '';
                }
                return;
            }

            // Fingerprint zur Erkennung von Änderungen seit letztem Save
            let cacheFingerprint = '';
            for (const [name, c] of explorerState.sheetDataCache) {
                cacheFingerprint += `${name}:${c.editedCells.size}:${c.rowHighlights.size},`;
            }
            const fingerprint = `${totalChanges}|${explorerState.pendingSheetOperations.length}|${explorerState.selectedSheet}|${cacheFingerprint}|cur:${explorerState.editedCells.size}:${explorerState.rowHighlights.size}`;
            if (fingerprint === _lastRecoveryFingerprint) {
                return;
            }

            if (_recoveryIdleHandle !== null) {
                (window.cancelIdleCallback || clearTimeout).call(window, _recoveryIdleHandle);
                _recoveryIdleHandle = null;
            }

            const doSave = () => {
                _recoveryIdleHandle = null;
                try {
                    // Aktuelles Sheet erst in Cache schreiben (damit alle Sheets gleich behandelt werden)
                    saveCurrentSheetToCache();

                    // Pro Sheet nur die Diffs sammeln (KEINE data/originalData/headers)
                    const sheetDeltas = {};
                    for (const [sheetName, c] of explorerState.sheetDataCache) {
                        const edits = c.editedCells && c.editedCells.size > 0
                            ? Array.from(c.editedCells.entries()) : [];
                        const highlights = c.rowHighlights && c.rowHighlights.size > 0
                            ? Array.from(c.rowHighlights.entries()) : [];
                        if (edits.length === 0 && highlights.length === 0) continue;
                        sheetDeltas[sheetName] = { editedCells: edits, rowHighlights: highlights };
                    }

                    const recoveryData = {
                        version: '2.0',
                        timestamp: Date.now(),
                        filePath: explorerState.filePath,
                        fileName: explorerState.fileName,
                        selectedSheet: explorerState.selectedSheet,
                        sheets: sheetDeltas,
                        pendingSheetOperations: explorerState.pendingSheetOperations,
                        hiddenSheets: Array.from(explorerState.hiddenSheets || [])
                    };

                    localStorage.setItem(EXPLORER_RECOVERY_KEY, JSON.stringify(recoveryData));
                    _lastRecoveryFingerprint = fingerprint;
                } catch (e) {
                    console.warn('Explorer Recovery-Speicherung fehlgeschlagen:', e);
                }
            };

            // Serialisierung in Idle-Zeit verschieben (auch wenn der Payload jetzt klein ist)
            if (window.requestIdleCallback) {
                _recoveryIdleHandle = requestIdleCallback(doSave, { timeout: 5000 });
            } else {
                _recoveryIdleHandle = setTimeout(doSave, 50);
            }
        }
        
        // Recovery-Daten laden
        function loadExplorerRecoveryData() {
            try {
                const saved = localStorage.getItem(EXPLORER_RECOVERY_KEY);
                if (!saved) return null;
                
                const data = JSON.parse(saved);
                
                // Prüfe ob Daten nicht älter als 24 Stunden sind
                const maxAge = 24 * 60 * 60 * 1000; // 24 Stunden
                if (Date.now() - data.timestamp > maxAge) {
                    localStorage.removeItem(EXPLORER_RECOVERY_KEY);
                    return null;
                }
                
                return data;
            } catch (e) {
                console.warn('Explorer Recovery-Laden fehlgeschlagen:', e);
                localStorage.removeItem(EXPLORER_RECOVERY_KEY);
                return null;
            }
        }
        
        // Recovery-Daten anwenden
        // v2.0 (Diff-only): Datei muss vorher per loadExplorerFileByPath geladen sein.
        //   Diese Funktion wendet danach die gespeicherten Diffs auf alle Sheets an.
        // v1.x (Legacy):    Vollständige Datenwiederherstellung aus dem Snapshot.
        async function applyExplorerRecoveryData(data) {
            if (data && data.version === '2.0') {
                await applyExplorerRecoveryDataV2(data);
                return;
            }
            // ----- Legacy v1.x: vollständige State-Wiederherstellung -----
            explorerState.filePath = data.filePath;
            explorerState.fileName = data.fileName;
            explorerState.sheets = data.sheets || [];
            explorerState.selectedSheet = data.selectedSheet;
            explorerState.headers = data.headers;
            explorerState.data = data.data;
            explorerState.originalData = data.originalData || data.data.map(row => [...row]);
            explorerState.visibleColumns = data.visibleColumns;
            explorerState.columnOrder = data.columnOrder || [];
            
            // Map/Set aus Arrays wiederherstellen
            explorerState.editedCells = new Map(data.editedCells || []);
            explorerState.rowHighlights = new Map(data.rowHighlights || []);
            explorerState.pendingSheetOperations = data.pendingSheetOperations || [];
            explorerState.hiddenSheets = new Set(data.hiddenSheets || []);
            
            // Sheet-Cache wiederherstellen
            explorerState.sheetDataCache.clear();
            if (data.sheetDataCache) {
                for (const cached of data.sheetDataCache) {
                    explorerState.sheetDataCache.set(cached.sheetName, {
                        headers: cached.headers,
                        data: cached.data,
                        originalData: cached.originalData || cached.data.map(row => [...row]),
                        visibleColumns: cached.visibleColumns,
                        editedCells: new Map(cached.editedCells || []),
                        rowHighlights: new Map(cached.rowHighlights || []),
                        hiddenRows: new Set(cached.hiddenRows || [])
                    });
                }
            }
            
            // UI aktualisieren
            const explorerFileName = document.getElementById('explorerFileName');
            if (explorerFileName) {
                explorerFileName.textContent = explorerState.fileName || '';
            }
            
            // Sheet-Dropdown füllen (mit Markierung für ausgeblendete Sheets)
            if (elements.explorerSheetSelect && explorerState.sheets.length > 0) {
                elements.explorerSheetSelect.innerHTML = explorerState.sheets.map(s => {
                    const isHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(s);
                    const label = isHidden ? `👁️‍🗨️ ${s} (ausgeblendet)` : s;
                    return `<option value="${s}" ${s === explorerState.selectedSheet ? 'selected' : ''}>${label}</option>`;
                }).join('');
            }
            
            // Daten filtern und anzeigen
            filterExplorerData();
            
            // Auto-Save starten
            startExplorerAutoSave();
        }

        // Wendet v2.0-Recovery (Diffs) an: lädt die Datei und schreibt die gespeicherten
        // editedCells/rowHighlights pro Sheet wieder in den Cache.
        async function applyExplorerRecoveryDataV2(data) {
            // 1. Datei laden (öffnet ggf. Live-Session, fragt Passwort ab etc.)
            await loadExplorerFileByPath(data.filePath);

            // Wenn das Laden fehlgeschlagen ist (Pfad/Pw), abbrechen
            if (!explorerState.filePath) {
                showNotification(
                    currentLanguage === 'en' ? 'Recovery failed: file could not be loaded' : 'Wiederherstellung fehlgeschlagen: Datei konnte nicht geladen werden',
                    'error'
                );
                return;
            }

            // 2. Pending-Sheet-Operations & hidden sheets übernehmen
            explorerState.pendingSheetOperations = data.pendingSheetOperations || [];
            explorerState.hiddenSheets = new Set(data.hiddenSheets || explorerState.hiddenSheets || []);

            // 3. Diffs pro Sheet anwenden — Sheets werden bei Bedarf sequenziell geladen
            const originalSheet = explorerState.selectedSheet;
            const targetSheet = data.selectedSheet || originalSheet;
            const sheetNames = data.sheets ? Object.keys(data.sheets) : [];

            let appliedSheets = 0;
            for (const sheetName of sheetNames) {
                if (!explorerState.sheets.includes(sheetName)) continue; // Sheet existiert nicht mehr
                const diff = data.sheets[sheetName];

                // Sheet aktivieren, damit die Daten geladen werden
                if (explorerState.selectedSheet !== sheetName) {
                    // Aktuelles Sheet in den Cache speichern (damit Edits nicht verloren gehen)
                    saveCurrentSheetToCache();
                    await loadExplorerSheet(sheetName);
                }

                // Diffs auf den aktuellen State anwenden
                if (Array.isArray(diff.editedCells)) {
                    for (const [key, value] of diff.editedCells) {
                        explorerState.editedCells.set(key, value);
                        // Wert auch in data-Array setzen, damit die UI ihn anzeigt
                        const [rStr, cStr] = String(key).split(',');
                        const r = parseInt(rStr, 10), c = parseInt(cStr, 10);
                        if (!isNaN(r) && !isNaN(c) && explorerState.data[r]) {
                            explorerState.data[r][c] = value;
                        }
                    }
                }
                if (Array.isArray(diff.rowHighlights)) {
                    for (const [idx, color] of diff.rowHighlights) {
                        explorerState.rowHighlights.set(parseInt(idx, 10), color);
                    }
                }
                saveCurrentSheetToCache();
                appliedSheets++;
            }

            // 4. Zurück zum ursprünglich aktiven Sheet wechseln
            if (targetSheet && explorerState.selectedSheet !== targetSheet && explorerState.sheets.includes(targetSheet)) {
                saveCurrentSheetToCache();
                await loadExplorerSheet(targetSheet);
                if (elements.explorerSheetSelect) {
                    elements.explorerSheetSelect.value = targetSheet;
                }
            }

            // 5. UI aktualisieren
            filterExplorerData();
            // Auto-Save ist bereits durch loadExplorerFileByPath gestartet
        }
        
        // Recovery-Daten löschen
        function clearExplorerRecoveryData() {
            localStorage.removeItem(EXPLORER_RECOVERY_KEY);
        }
        
        // Kontextmenü für Zeilen-Markierung
        const highlightColors = [
            { name: 'Grün', value: 'green', hex: '#4CAF50' },
            { name: 'Gelb', value: 'yellow', hex: '#FFEB3B' },
            { name: 'Orange', value: 'orange', hex: '#FF9800' },
            { name: 'Rot', value: 'red', hex: '#F44336' },
            { name: 'Blau', value: 'blue', hex: '#2196F3' },
            { name: 'Lila', value: 'purple', hex: '#9C27B0' }
        ];
        
        function showRowContextMenu(e, rowIndex) {
            e.preventDefault();
            
            // Altes Zeilen-Menü entfernen (nur row-context-menu, nicht das column-context-menu)
            const oldMenu = document.querySelector('.row-context-menu');
            if (oldMenu) oldMenu.remove();
            
            const currentColor = explorerState.rowHighlights.get(rowIndex);
            
            let menuHtml = '<div class="context-menu row-context-menu">';
            menuHtml += `<div class="context-menu-item" style="font-weight: 600; color: var(--text-muted); cursor: default;">🎨 ${t('highlightRow')}</div>`;
            menuHtml += '<div class="context-menu-divider"></div>';
            
            highlightColors.forEach(color => {
                const selected = currentColor === color.value ? ' ✓' : '';
                const colorKey = 'highlight' + color.value.charAt(0).toUpperCase() + color.value.slice(1);
                menuHtml += `
                    <div class="context-menu-item" data-action="highlight" data-color="${color.value}" data-row="${rowIndex}">
                        <span class="color-dot" style="background: ${color.hex};"></span>
                        <span>${t(colorKey)}${selected}</span>
                    </div>
                `;
            });
            
            // Markierung entfernen Option (nur anzeigen wenn Zeile markiert ist)
            if (currentColor) {
                menuHtml += '<div class="context-menu-divider"></div>';
                menuHtml += `<div class="context-menu-item" data-action="clear-highlight" data-row="${rowIndex}" style="color: var(--text-muted);">✖️ ${t('clearHighlight')}</div>`;
            }
            
            menuHtml += '<div class="context-menu-divider"></div>';
            menuHtml += `<div class="context-menu-item" data-action="insert-row-above" data-row="${rowIndex}">⬆️ ${t('insertRowAbove')}</div>`;
            menuHtml += `<div class="context-menu-item" data-action="insert-row-below" data-row="${rowIndex}">⬇️ ${t('insertRowBelow')}</div>`;
            menuHtml += '<div class="context-menu-divider"></div>';
            menuHtml += `<div class="context-menu-item" data-action="hide-row" data-row="${rowIndex}">👁️‍🗨️ ${t('hideRow')}</div>`;
            menuHtml += `<div class="context-menu-item" data-action="delete-row" data-row="${rowIndex}" style="color: #F44336;">🗑️ ${t('deleteRow')}</div>`;
            menuHtml += '</div>';
            
            const menu = document.createElement('div');
            menu.innerHTML = menuHtml;
            document.body.appendChild(menu.firstElementChild);
            
            const menuEl = document.querySelector('.row-context-menu');
            
            // Positionierung
            let x = e.clientX;
            let y = e.clientY;
            
            // Sicherstellen, dass das Menü im Viewport bleibt
            const menuRect = menuEl.getBoundingClientRect();
            if (x + menuRect.width > window.innerWidth) {
                x = window.innerWidth - menuRect.width - 5;
            }
            if (y + menuRect.height > window.innerHeight) {
                y = window.innerHeight - menuRect.height - 5;
            }
            
            menuEl.style.left = x + 'px';
            menuEl.style.top = y + 'px';
            
            // Event-Handler
            menuEl.querySelectorAll('[data-action="highlight"]').forEach(item => {
                item.addEventListener('click', () => {
                    const color = item.dataset.color;
                    const row = parseInt(item.dataset.row);
                    setRowHighlight(row, color);
                    menuEl.remove();
                });
            });
            
            // Markierung entfernen Handler
            menuEl.querySelector('[data-action="clear-highlight"]')?.addEventListener('click', () => {
                const row = parseInt(menuEl.querySelector('[data-action="clear-highlight"]').dataset.row);
                clearRowHighlight(row);
                menuEl.remove();
            });
            
            // Neue Zeilen-Aktionen
            menuEl.querySelector('[data-action="insert-row-above"]')?.addEventListener('click', async () => {
                const row = parseInt(menuEl.querySelector('[data-action="insert-row-above"]').dataset.row);
                menuEl.remove();
                await insertExplorerRow(row, 'above');
            });
            
            menuEl.querySelector('[data-action="insert-row-below"]')?.addEventListener('click', async () => {
                const row = parseInt(menuEl.querySelector('[data-action="insert-row-below"]').dataset.row);
                menuEl.remove();
                await insertExplorerRow(row, 'below');
            });
            
            menuEl.querySelector('[data-action="delete-row"]')?.addEventListener('click', async () => {
                const row = parseInt(menuEl.querySelector('[data-action="delete-row"]').dataset.row);
                menuEl.remove();
                await deleteExplorerRow(row);
            });
            
            menuEl.querySelector('[data-action="hide-row"]')?.addEventListener('click', async () => {
                const row = parseInt(menuEl.querySelector('[data-action="hide-row"]').dataset.row);
                menuEl.remove();
                await hideExplorerRow(row);
            });
            
            // Menü bei Klick außerhalb schließen
            const closeHandler = (event) => {
                if (!menuEl.contains(event.target)) {
                    menuEl.remove();
                    document.removeEventListener('click', closeHandler);
                }
            };
            setTimeout(() => document.addEventListener('click', closeHandler), 10);
        }
        
        function setRowHighlight(rowIndex, color) {
            // Toggle: Wenn dieselbe Farbe nochmal gewählt wird, Markierung entfernen
            if (explorerState.rowHighlights.get(rowIndex) === color) {
                // Gleiche Farbe - nichts tun
                return;
            }
            
            // LIVE SESSION: Markiere Zeile in Excel (nur visuell, Datei wird nicht gespeichert)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                liveSessionExecute('highlightRow', getExcelRowPosition(rowIndex), color);
            }
            
            explorerState.rowHighlights.set(rowIndex, color);
            // Strukturelle Änderung markieren für Export
            explorerState.editedCells.set('_rowHighlightChanged', true);
            // KEIN _hasFormatChanges hier! Row-Highlights werden via direktem XML-Pfad
            // (FALL 3a) gesetzt — kein openpyxl-Roundtrip nötig → Slicers bleiben intakt.
            // Nur clearRowHighlight setzt _hasFormatChanges (Fill-Löschung braucht openpyxl).
            applyRowHighlights();
        }
        
        function clearRowHighlight(rowIndex) {
            // LIVE SESSION: Entferne Markierung in Excel (nur visuell, Datei wird nicht gespeichert)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                liveSessionExecute('highlightRow', getExcelRowPosition(rowIndex), null);
            }
            
            explorerState.rowHighlights.delete(rowIndex);
            // Strukturelle Änderung markieren für Export
            explorerState.editedCells.set('_rowHighlightChanged', true);
            explorerState.editedCells.set('_hasFormatChanges', true);
            
            // Fill aus cellStyles für alle Zellen dieser Zeile entfernen
            // (damit die inline-Styles nicht die Entfernung überschreiben)
            // WICHTIG: cellStyles verwendet 1-basierte rowIndex Keys!
            const colCount = explorerState.headers.length;
            for (let colIdx = 0; colIdx < colCount; colIdx++) {
                const cellKey = `${rowIndex + 1}-${colIdx}`;
                if (explorerState.cellStyles[cellKey] && explorerState.cellStyles[cellKey].fill) {
                    delete explorerState.cellStyles[cellKey].fill;
                }
            }
            
            applyRowHighlights();
            // Tabelle neu rendern um die Styles zu aktualisieren
            renderExplorerTable();
        }
        
        function applyRowHighlights() {
            // Alle Highlights entfernen
            document.querySelectorAll('#explorerTableBody tr').forEach(tr => {
                tr.classList.remove('row-highlight-green', 'row-highlight-yellow', 'row-highlight-orange', 
                                   'row-highlight-red', 'row-highlight-blue', 'row-highlight-purple');
            });
            
            // Neue Highlights setzen
            explorerState.rowHighlights.forEach((color, rowIndex) => {
                const tr = document.querySelector(`#explorerTableBody tr[data-original-index="${rowIndex}"]`);
                if (tr) {
                    tr.classList.add(`row-highlight-${color}`);
                }
            });
        }
        
        // ==================== Zeilen/Spalten Einfügen/Löschen ====================
        
        // Neue Zeile einfügen
        async function insertExplorerRow(rowIndex, position = 'below') {
            const insertIndex = position === 'above' ? rowIndex : rowIndex + 1;
            
            // Live-Session: Zeile sofort in Excel einfügen
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // In der Live-Session sind GUI-Daten und Excel synchron.
                // insertIndex ist die korrekte 0-basierte Datenposition.
                // Python rechnet +2 für die Excel-Zeile (Header + 1-basiert).
                try {
                    const result = await window.electronAPI.liveSessionInsertRow(insertIndex, 1);
                    if (!result || !result.success) {
                        console.error('[LiveSession] insertRow failed:', result);
                        showFloatingStatus('❌ Fehler beim Einfügen in Excel', 'error');
                        return;
                    }
                    console.log(`[LiveSession] Row inserted: dataIndex=${insertIndex}`);
                } catch (error) {
                    console.error('[LiveSession] insertRow error:', error);
                    showFloatingStatus('❌ Fehler beim Einfügen in Excel', 'error');
                    return;
                }
            }
            
            // Leere Zeile erstellen mit gleicher Anzahl Spalten
            const emptyRow = new Array(explorerState.headers.length).fill('');
            
            // Zeile in data einfügen
            explorerState.data.splice(insertIndex, 0, emptyRow);
            explorerState.originalData.splice(insertIndex, 0, [...emptyRow]);
            
            // Highlights anpassen (Indizes verschieben)
            const newHighlights = new Map();
            explorerState.rowHighlights.forEach((color, idx) => {
                if (idx >= insertIndex) {
                    newHighlights.set(idx + 1, color);
                } else {
                    newHighlights.set(idx, color);
                }
            });
            explorerState.rowHighlights = newHighlights;
            
            // CellStyles anpassen (Zeilen-Indizes verschieben)
            // Style-Keys: "styleRowIdx-colIdx" wobei styleRowIdx = dataRowIdx + 1 (Header = 0)
            const newCellStyles = {};
            for (const [key, style] of Object.entries(explorerState.cellStyles || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row === 0) {
                    // Header behalten
                    newCellStyles[key] = style;
                } else if (row - 1 >= insertIndex) {
                    // Zeile nach Einfügepunkt: Index + 1
                    newCellStyles[`${row + 1}-${col}`] = style;
                } else {
                    newCellStyles[key] = style;
                }
            }
            explorerState.cellStyles = newCellStyles;
            
            // CellFormulas anpassen
            const newCellFormulas = {};
            for (const [key, formula] of Object.entries(explorerState.cellFormulas || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row === 0) {
                    newCellFormulas[key] = formula;
                } else if (row - 1 >= insertIndex) {
                    newCellFormulas[`${row + 1}-${col}`] = formula;
                } else {
                    newCellFormulas[key] = formula;
                }
            }
            explorerState.cellFormulas = newCellFormulas;
            
            // CellHyperlinks anpassen
            const newCellHyperlinks = {};
            for (const [key, link] of Object.entries(explorerState.cellHyperlinks || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row === 0) {
                    newCellHyperlinks[key] = link;
                } else if (row - 1 >= insertIndex) {
                    newCellHyperlinks[`${row + 1}-${col}`] = link;
                } else {
                    newCellHyperlinks[key] = link;
                }
            }
            explorerState.cellHyperlinks = newCellHyperlinks;
            
            // RichTextCells anpassen
            const newRichTextCells = {};
            for (const [key, rt] of Object.entries(explorerState.richTextCells || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row === 0) {
                    newRichTextCells[key] = rt;
                } else if (row - 1 >= insertIndex) {
                    newRichTextCells[`${row + 1}-${col}`] = rt;
                } else {
                    newRichTextCells[key] = rt;
                }
            }
            explorerState.richTextCells = newRichTextCells;
            
            // HiddenRows anpassen
            const newHiddenRows = new Set();
            explorerState.hiddenRows.forEach(idx => {
                if (idx >= insertIndex) {
                    newHiddenRows.add(idx + 1);
                } else {
                    newHiddenRows.add(idx);
                }
            });
            explorerState.hiddenRows = newHiddenRows;
            
            // EditedCells anpassen
            const newEditedCells = new Map();
            explorerState.editedCells.forEach((value, key) => {
                if (key.startsWith('_')) {
                    newEditedCells.set(key, value);
                    return;
                }
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row >= insertIndex) {
                    newEditedCells.set(`${row + 1}-${col}`, value);
                } else {
                    newEditedCells.set(key, value);
                }
            });
            explorerState.editedCells = newEditedCells;
            
            // Alle Zellen der neuen Zeile als bearbeitet markieren
            for (let col = 0; col < explorerState.headers.length; col++) {
                explorerState.editedCells.set(`${insertIndex}-${col}`, '');
            }
            
            // rowMapping aktualisieren - trackt welche Original-Excel-Zeile an welcher Position ist
            // Bei neuen Zeilen wird -1 als Marker verwendet (keine Original-Zeile)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Live Session: Daten und Excel sind synchron, kein Mapping nötig
                // Excel hat die Zeile bereits eingefügt, alle Indizes stimmen überein
                explorerState.rowMapping = null;
            } else if (explorerState.rowMapping && explorerState.rowMapping.length > 0) {
                // Bestehendes Mapping: -1 an insertIndex einfügen
                // Die bestehenden Mappings bleiben unverändert (sie referenzieren immer noch ihre Original-Zeilen)
                explorerState.rowMapping.splice(insertIndex, 0, -1);  // -1 = neue Zeile
            } else {
                // Neues Mapping erstellen - jeder Index mappt auf sich selbst (Original-Position)
                explorerState.rowMapping = [];
                const originalDataLength = explorerState.data.length;
                for (let i = 0; i < originalDataLength; i++) {
                    if (i === insertIndex) {
                        explorerState.rowMapping.push(-1);  // -1 = neue Zeile
                    } else if (i < insertIndex) {
                        explorerState.rowMapping.push(i);  // Original-Index
                    } else {
                        // Zeilen nach der eingefügten Position: Original-Index ist i-1
                        // (weil die Original-Datei diese Zeile nicht hatte)
                        explorerState.rowMapping.push(i - 1);
                    }
                }
            }
            
            // WICHTIG: Strukturelle Änderung markieren für Export (Full Rewrite nötig)
            explorerState.editedCells.set('_rowInserted', true);
            
            // Tracke eingefügte Zeilen separat (analog zu _columnInserted)
            const existingInsert = explorerState.editedCells.get('_insertedRowInfo');
            if (existingInsert && existingInsert.operations) {
                // Füge neue Operation hinzu
                existingInsert.operations.push({
                    position: insertIndex,  // 0-basiert
                    count: 1
                });
                existingInsert.totalCount = (existingInsert.totalCount || 0) + 1;
                explorerState.editedCells.set('_insertedRowInfo', existingInsert);
            } else {
                // Neue insertedRowInfo mit korrektem Format
                explorerState.editedCells.set('_insertedRowInfo', { 
                    operations: [{
                        position: insertIndex,  // 0-basiert
                        count: 1
                    }],
                    totalCount: 1
                });
            }
            
            // UI aktualisieren
            filterExplorerData();
            showFloatingStatus(t('rowInserted'));
        }
        
        // Zeile löschen
        async function deleteExplorerRow(rowIndex) {
            // Bestätigung bei einzelner Zeile
            const rowPreview = explorerState.data[rowIndex]?.slice(0, 3).join(', ') || '';
            const rowLabel = currentLanguage === 'en' ? 'Row' : 'Zeile';
            const confirmed = await showConfirmDialog(
                t('deleteRowTitle'),
                `${t('deleteRowConfirm')}\n\n${rowLabel} ${rowIndex + 2}: ${rowPreview}...`,
                currentLanguage === 'en' ? 'Delete' : 'Löschen',
                t('cancel')
            );
            
            if (!confirmed) return;
            
            // Live-Session: Zeile sofort in Excel löschen
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                const excelRowPos = getExcelRowPosition(rowIndex);
                await liveSessionExecute('deleteRow', excelRowPos);
                console.log(`[LiveSession] Row deleted: dataIndex=${rowIndex}, excelPos=${excelRowPos}`);
            }
            
            // Zeile aus data entfernen
            explorerState.data.splice(rowIndex, 1);
            explorerState.originalData.splice(rowIndex, 1);
            
            // Highlights anpassen
            const newHighlights = new Map();
            explorerState.rowHighlights.forEach((color, idx) => {
                if (idx > rowIndex) {
                    newHighlights.set(idx - 1, color);
                } else if (idx < rowIndex) {
                    newHighlights.set(idx, color);
                }
                // idx === rowIndex wird nicht übernommen (gelöscht)
            });
            explorerState.rowHighlights = newHighlights;
            
            // EditedCells anpassen
            const newEditedCells = new Map();
            explorerState.editedCells.forEach((value, key) => {
                if (key.startsWith('_')) {
                    newEditedCells.set(key, value);
                    return;
                }
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row > rowIndex) {
                    newEditedCells.set(`${row - 1}-${col}`, value);
                } else if (row < rowIndex) {
                    newEditedCells.set(key, value);
                }
                // row === rowIndex wird nicht übernommen
            });
            explorerState.editedCells = newEditedCells;
            
            // CellStyles anpassen (analog zu deleteExplorerColumn)
            // Style-Keys: "styleRowIdx-colIdx" wobei styleRowIdx = dataRowIdx + 1 (Header = 0)
            // Gelöschte Zeile hat styleRowIdx = rowIndex + 1
            const deleteStyleRow = rowIndex + 1;
            const newCellStyles = {};
            for (const [key, value] of Object.entries(explorerState.cellStyles || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row > deleteStyleRow) {
                    newCellStyles[`${row - 1}-${col}`] = value;
                } else if (row < deleteStyleRow) {
                    newCellStyles[key] = value;
                }
                // row === deleteStyleRow wird nicht übernommen (gelöscht)
            }
            explorerState.cellStyles = newCellStyles;
            
            // CellFormulas anpassen
            const newCellFormulas = {};
            for (const [key, value] of Object.entries(explorerState.cellFormulas || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row > deleteStyleRow) {
                    newCellFormulas[`${row - 1}-${col}`] = value;
                } else if (row < deleteStyleRow) {
                    newCellFormulas[key] = value;
                }
            }
            explorerState.cellFormulas = newCellFormulas;
            
            // CellHyperlinks anpassen
            const newCellHyperlinks = {};
            for (const [key, value] of Object.entries(explorerState.cellHyperlinks || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row > deleteStyleRow) {
                    newCellHyperlinks[`${row - 1}-${col}`] = value;
                } else if (row < deleteStyleRow) {
                    newCellHyperlinks[key] = value;
                }
            }
            explorerState.cellHyperlinks = newCellHyperlinks;
            
            // RichTextCells anpassen
            const newRichTextCells = {};
            for (const [key, value] of Object.entries(explorerState.richTextCells || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (row > deleteStyleRow) {
                    newRichTextCells[`${row - 1}-${col}`] = value;
                } else if (row < deleteStyleRow) {
                    newRichTextCells[key] = value;
                }
            }
            explorerState.richTextCells = newRichTextCells;
            
            // HiddenRows anpassen
            const newHiddenRows = new Set();
            explorerState.hiddenRows.forEach(idx => {
                if (idx > rowIndex) {
                    newHiddenRows.add(idx - 1);
                } else if (idx < rowIndex) {
                    newHiddenRows.add(idx);
                }
            });
            explorerState.hiddenRows = newHiddenRows;
            
            // rowMapping aktualisieren - trackt welche Original-Excel-Zeile an welcher Position ist
            // Analog zu _columnDeleted bei Spalten
            
            // WICHTIG: Zuerst den Original-Index erfassen BEVOR wir rowMapping ändern
            let originalRowIndex;
            if (explorerState.rowMapping && explorerState.rowMapping.length > rowIndex) {
                originalRowIndex = explorerState.rowMapping[rowIndex];
            } else {
                // Ohne Mapping ist der aktuelle Index = Original-Index
                originalRowIndex = rowIndex;
            }
            
            // Tracke gelöschte Original-Zeilen (analog zu _columnDeleted)
            const existingDeleted = explorerState.editedCells.get('_deletedRowIndices');
            let deletedOriginalIndices = [];
            if (existingDeleted && Array.isArray(existingDeleted.originalIndices)) {
                deletedOriginalIndices = existingDeleted.originalIndices.slice();
            }
            // -1 = eingefügte Zeile, diese gibt es nicht im Original
            if (originalRowIndex >= 0) {
                deletedOriginalIndices.push(originalRowIndex);
            }
            explorerState.editedCells.set('_deletedRowIndices', { 
                originalIndices: deletedOriginalIndices,  // Array der ORIGINAL-Zeilen-Indices (0-basiert)
                count: deletedOriginalIndices.length
            });
            
            // rowMapping aktualisieren
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Live Session: Daten und Excel sind synchron, kein Mapping nötig
                explorerState.rowMapping = null;
            } else if (explorerState.rowMapping && explorerState.rowMapping.length > 0) {
                // Bestehendes Mapping: Element an rowIndex entfernen
                const newRowMapping = [...explorerState.rowMapping];
                newRowMapping.splice(rowIndex, 1);
                explorerState.rowMapping = newRowMapping;
            } else {
                // Neues Mapping erstellen: neue Position -> Original-Zeile (vor dem Löschen)
                explorerState.rowMapping = [];
                for (let i = 0; i < explorerState.data.length; i++) {
                    // Zeilen nach der gelöschten Position referenzieren ihre Original-Position + 1
                    const originalIdx = i >= rowIndex ? i + 1 : i;
                    explorerState.rowMapping.push(originalIdx);
                }
            }
            
            // Markierung, dass etwas gelöscht wurde (für Änderungszählung)
            explorerState.editedCells.set('_rowDeleted', true);
            
            // Zur Operations-Queue hinzufügen
            explorerState.rowOperationsQueue.push({
                type: 'delete',
                originalIndex: originalRowIndex
            });
            
            // UI aktualisieren
            filterExplorerData();
            showFloatingStatus(t('rowDeleted'));
        }
        
        // Zeile ausblenden
        async function hideExplorerRow(rowIndex) {
            // Live-Session: Zeile sofort in Excel ausblenden
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    const excelRowPos = getExcelRowPosition(rowIndex);
                    const result = await window.electronAPI.liveSessionHideRow(excelRowPos, true);
                    if (!result || !result.success) {
                        console.error('[LiveSession] hideRow failed:', result);
                    } else {
                        console.log(`[LiveSession] Row hidden: dataIndex=${rowIndex}, excelPos=${excelRowPos}`);
                    }
                } catch (error) {
                    console.error('[LiveSession] hideRow error:', error);
                }
            }
            
            explorerState.hiddenRows.add(rowIndex);
            // Markierung, dass Zeilen-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_rowVisibilityChanged', true);
            filterExplorerData();
            updateExplorerEditStatus();
            updateHiddenRowsIndicator();
            showFloatingStatus(currentLanguage === 'en' ? 'Row hidden' : 'Zeile ausgeblendet');
        }
        
        // Zeile wieder einblenden
        async function showExplorerRow(rowIndex) {
            // Live-Session: Zeile sofort in Excel einblenden
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    const excelRowPos = getExcelRowPosition(rowIndex);
                    const result = await window.electronAPI.liveSessionHideRow(excelRowPos, false);
                    if (!result || !result.success) {
                        console.error('[LiveSession] showRow failed:', result);
                    } else {
                        console.log(`[LiveSession] Row shown: dataIndex=${rowIndex}, excelPos=${excelRowPos}`);
                    }
                } catch (error) {
                    console.error('[LiveSession] showRow error:', error);
                }
            }
            
            explorerState.hiddenRows.delete(rowIndex);
            // Markierung, dass Zeilen-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_rowVisibilityChanged', true);
            filterExplorerData();
            updateExplorerEditStatus();
            updateHiddenRowsIndicator();
            showFloatingStatus(currentLanguage === 'en' ? 'Row shown' : 'Zeile eingeblendet');
        }
        
        // Alle versteckten Zeilen einblenden
        function showAllHiddenRows() {
            // LIVE SESSION: Zeilen in Excel wieder einblenden (Batch für Performance)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady && explorerState.hiddenRows.size > 0) {
                const indicesToShow = Array.from(explorerState.hiddenRows).map(idx => getExcelRowPosition(idx));
                window.electronAPI.liveSessionHideRowsBatch(indicesToShow, false)
                    .then(result => {
                        if (result && result.success) {
                            console.log(`[LiveSession] ${indicesToShow.length} Zeilen in Excel eingeblendet (Batch)`);
                        }
                    })
                    .catch(error => console.error('[LiveSession] showRowsBatch error:', error));
            }
            
            explorerState.hiddenRows.clear();
            explorerState.editedCells.set('_rowVisibilityChanged', true);
            filterExplorerData();
            updateExplorerEditStatus();
            updateHiddenRowsIndicator();
            showFloatingStatus(currentLanguage === 'en' ? 'All rows shown' : 'Alle Zeilen eingeblendet');
        }
        
        // Indikator für versteckte Zeilen aktualisieren
        function updateHiddenRowsIndicator() {
            let indicator = document.getElementById('hiddenRowsIndicator');
            const count = explorerState.hiddenRows.size;
            
            if (count === 0) {
                if (indicator) indicator.remove();
                return;
            }
            
            if (!indicator) {
                // Erstelle Indikator-Button
                indicator = document.createElement('button');
                indicator.id = 'hiddenRowsIndicator';
                indicator.className = 'btn btn-warning btn-sm';
                indicator.title = currentLanguage === 'en' ? 'Click to show hidden rows' : 'Klicken um versteckte Zeilen anzuzeigen';
                
                // Füge in den festen Container ein
                const container = document.getElementById('hiddenIndicatorsContainer');
                if (container) {
                    container.appendChild(indicator);
                }
            }
            
            indicator.innerHTML = `☰ ${count} ${currentLanguage === 'en' ? 'rows hidden' : 'Zeilen ausgeblendet'}`;
            indicator.onclick = showHiddenRowsMenu;
        }
        
        // Menü für versteckte Zeilen anzeigen
        function showHiddenRowsMenu(e) {
            e.preventDefault();
            e.stopPropagation();
            
            // Altes Menü entfernen
            const oldMenu = document.querySelector('.hidden-rows-menu');
            if (oldMenu) oldMenu.remove();
            
            const hiddenArray = Array.from(explorerState.hiddenRows).sort((a, b) => a - b);
            
            let menuHtml = '<div class="context-menu hidden-rows-menu" style="display: flex; flex-direction: column; max-height: 400px;">';
            // Sticky header
            menuHtml += `<div class="context-menu-item" style="font-weight: 600; color: var(--text-muted); cursor: default; flex-shrink: 0;">☰ ${currentLanguage === 'en' ? 'Hidden Rows' : 'Versteckte Zeilen'} (${hiddenArray.length})</div>`;
            menuHtml += '<div class="context-menu-divider" style="flex-shrink: 0;"></div>';
            // Sticky "Alle einblenden" button at top
            menuHtml += `<div class="context-menu-item" data-action="show-all" style="color: var(--primary); font-weight: 600; flex-shrink: 0;">✅ ${currentLanguage === 'en' ? 'Show all' : 'Alle einblenden'}</div>`;
            menuHtml += '<div class="context-menu-divider" style="flex-shrink: 0;"></div>';
            // Scrollable list
            menuHtml += '<div style="overflow-y: auto; flex: 1; min-height: 0;">';
            
            hiddenArray.forEach(rowIndex => {
                const rowPreview = explorerState.data[rowIndex]?.slice(0, 2).join(', ').substring(0, 30) || '...';
                menuHtml += `<div class="context-menu-item" data-action="show-row" data-row="${rowIndex}">
                    👁️ ${currentLanguage === 'en' ? 'Row' : 'Zeile'} ${rowIndex + 2}: ${escapeHtml(rowPreview)}...
                </div>`;
            });
            
            menuHtml += '</div>'; // end scrollable
            menuHtml += '</div>';
            
            const menu = document.createElement('div');
            menu.innerHTML = menuHtml;
            document.body.appendChild(menu.firstElementChild);
            
            const menuEl = document.querySelector('.hidden-rows-menu');
            
            // Positionierung
            const rect = e.target.getBoundingClientRect();
            menuEl.style.left = rect.left + 'px';
            menuEl.style.top = (rect.bottom + 5) + 'px';
            
            // Event-Handler für einzelne Zeilen (Event-Delegation für Performance bei vielen Einträgen)
            const scrollContainer = menuEl.querySelector('[style*="overflow-y"]');
            if (scrollContainer) {
                scrollContainer.addEventListener('click', (evt) => {
                    const item = evt.target.closest('[data-action="show-row"]');
                    if (item) {
                        const row = parseInt(item.dataset.row);
                        showExplorerRow(row);
                        // Menü nach Einzelklick aktualisieren statt schließen
                        if (explorerState.hiddenRows.size > 0) {
                            menuEl.remove();
                            showHiddenRowsMenu(e);
                        } else {
                            menuEl.remove();
                        }
                    }
                });
            }
            
            menuEl.querySelector('[data-action="show-all"]')?.addEventListener('click', () => {
                showAllHiddenRows();
                menuEl.remove();
            });
            
            // Menü bei Klick außerhalb schließen
            const closeHandler = (event) => {
                if (!menuEl.contains(event.target)) {
                    menuEl.remove();
                    document.removeEventListener('click', closeHandler);
                }
            };
            setTimeout(() => document.addEventListener('click', closeHandler), 10);
        }
        
        // Indikator für versteckte Spalten aktualisieren
        function updateHiddenColumnsIndicator() {
            let indicator = document.getElementById('hiddenColumnsIndicator');
            // Berechne Anzahl versteckter Spalten (alle Spalten die nicht in visibleColumns sind)
            const totalColumns = explorerState.headers.length;
            const hiddenCount = totalColumns - explorerState.visibleColumns.length;
            
            if (hiddenCount === 0) {
                if (indicator) indicator.remove();
                return;
            }
            
            if (!indicator) {
                // Erstelle Indikator-Button
                indicator = document.createElement('button');
                indicator.id = 'hiddenColumnsIndicator';
                indicator.className = 'btn btn-warning btn-sm';
                indicator.title = currentLanguage === 'en' ? 'Click to show hidden columns' : 'Klicken um versteckte Spalten anzuzeigen';
                
                // Füge in den festen Container ein
                const container = document.getElementById('hiddenIndicatorsContainer');
                if (container) {
                    container.appendChild(indicator);
                }
            }
            
            indicator.innerHTML = `📊 ${hiddenCount} ${currentLanguage === 'en' ? 'hidden' : 'ausgeblendet'}`;
            indicator.onclick = showHiddenColumnsMenu;
        }
        
        // Menü für versteckte Spalten anzeigen
        function showHiddenColumnsMenu(e) {
            e.preventDefault();
            e.stopPropagation();
            
            // Altes Menü entfernen
            const oldMenu = document.querySelector('.hidden-columns-menu');
            if (oldMenu) oldMenu.remove();
            
            // Versteckte Spalten ermitteln (alle die nicht in visibleColumns sind)
            const hiddenColumns = explorerState.headers
                .map((header, i) => ({ index: i, header }))
                .filter(({ index }) => !explorerState.visibleColumns.includes(index));
            
            let menuHtml = '<div class="context-menu hidden-columns-menu" style="display: flex; flex-direction: column; max-height: 400px;">';
            // Sticky header
            menuHtml += `<div class="context-menu-item" style="font-weight: 600; color: var(--text-muted); cursor: default; flex-shrink: 0;">📊 ${currentLanguage === 'en' ? 'Hidden Columns' : 'Versteckte Spalten'} (${hiddenColumns.length})</div>`;
            menuHtml += '<div class="context-menu-divider" style="flex-shrink: 0;"></div>';
            // Sticky "Alle einblenden" button at top
            menuHtml += `<div class="context-menu-item" data-action="show-all" style="color: var(--primary); font-weight: 600; flex-shrink: 0;">✅ ${currentLanguage === 'en' ? 'Show all' : 'Alle einblenden'}</div>`;
            menuHtml += '<div class="context-menu-divider" style="flex-shrink: 0;"></div>';
            // Scrollable list
            menuHtml += '<div style="overflow-y: auto; flex: 1; min-height: 0;">';
            
            hiddenColumns.forEach(({ index, header }) => {
                const displayHeader = header ? header.substring(0, 30) : `Spalte ${index + 1}`;
                menuHtml += `<div class="context-menu-item" data-action="show-column" data-col="${index}">
                    👁️ ${escapeHtml(displayHeader)}${header && header.length > 30 ? '...' : ''}
                </div>`;
            });
            
            menuHtml += '</div>'; // end scrollable
            menuHtml += '</div>';
            
            const menu = document.createElement('div');
            menu.innerHTML = menuHtml;
            document.body.appendChild(menu.firstElementChild);
            
            const menuEl = document.querySelector('.hidden-columns-menu');
            
            // Positionierung
            const rect = e.target.getBoundingClientRect();
            menuEl.style.left = rect.left + 'px';
            menuEl.style.top = (rect.bottom + 5) + 'px';
            
            // Event-Handler für einzelne Spalten (Event-Delegation für Performance)
            const scrollContainer = menuEl.querySelector('[style*="overflow-y"]');
            if (scrollContainer) {
                scrollContainer.addEventListener('click', (evt) => {
                    const item = evt.target.closest('[data-action="show-column"]');
                    if (item) {
                        const col = parseInt(item.dataset.col);
                        showExplorerColumn(col);
                        // Menü nach Einzelklick aktualisieren statt schließen
                        const remainingHidden = explorerState.headers.length - explorerState.visibleColumns.length;
                        if (remainingHidden > 0) {
                            menuEl.remove();
                            showHiddenColumnsMenu(e);
                        } else {
                            menuEl.remove();
                        }
                    }
                });
            }
            
            menuEl.querySelector('[data-action="show-all"]')?.addEventListener('click', () => {
                showAllExplorerColumns();
                menuEl.remove();
            });
            
            // Menü bei Klick außerhalb schließen
            const closeHandler = (event) => {
                if (!menuEl.contains(event.target)) {
                    menuEl.remove();
                    document.removeEventListener('click', closeHandler);
                }
            };
            setTimeout(() => document.addEventListener('click', closeHandler), 10);
        }
        
        // Einzelne Spalte wieder einblenden
        function showExplorerColumn(colIndex) {
            if (!explorerState.visibleColumns.includes(colIndex)) {
                explorerState.visibleColumns.push(colIndex);
                explorerState.visibleColumns.sort((a, b) => a - b);
            }
            // Markierung, dass Spalten-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_columnVisibilityChanged', true);
            renderExplorerTable();
            updateColumnToggles();
            updateHiddenColumnsIndicator();
            updateExplorerEditStatus();
            showFloatingStatus(currentLanguage === 'en' ? 'Column shown' : 'Spalte eingeblendet');
        }
        
        // Indikator für AutoFilter aktualisieren
        function updateAutoFilterIndicator() {
            const indicator = document.getElementById('autoFilterIndicator');
            if (!indicator) return;
            
            if (!explorerState.autoFilterRange) {
                indicator.style.display = 'none';
                return;
            }
            
            // Zeige Indikator
            let displayText = `▼ AutoFilter: ${explorerState.autoFilterRange}`;
            indicator.textContent = displayText;
            indicator.title = `AutoFilter aktiv: ${explorerState.autoFilterRange}\n(wird beim Speichern erhalten)`;
            indicator.style.display = 'inline-flex';
        }
        
        // Indikator für Passwortschutz aktualisieren
        function updatePasswordIndicator() {
            let indicator = document.getElementById('passwordIndicator');
            
            if (!explorerState.filePassword) {
                if (indicator) indicator.remove();
                return;
            }
            
            if (!indicator) {
                // Erstelle Indikator-Badge
                indicator = document.createElement('span');
                indicator.id = 'passwordIndicator';
                indicator.style.cssText = 'margin-left: 8px; background: #FF9800; color: white; font-size: 11px; padding: 3px 8px; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px; vertical-align: middle; cursor: pointer;';
                
                // Klick-Handler für Passwort-Änderung im Live-Modus
                indicator.onclick = () => showPasswordManagementDialog();
                
                // Füge nach dem Dateinamen ein
                const fileNameElement = document.getElementById('explorerFileName');
                if (fileNameElement && fileNameElement.parentElement) {
                    fileNameElement.parentElement.appendChild(indicator);
                }
            }
            
            indicator.innerHTML = `🔐 Passwortgeschützt`;
            indicator.title = 'Klicken um Passwort zu verwalten.\nDas Passwort wird beim Speichern beibehalten.';
        }
        
        // Dialog zur Passwortverwaltung
        async function showPasswordManagementDialog() {
            return new Promise((resolve) => {
                const dialogHTML = `
                    <div id="passwordManagementDialog" style="position: fixed; top: 0; left: 0; right: 0; bottom: 0; 
                        background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 10000;">
                        <div style="background: var(--background); border-radius: 12px; padding: 24px; width: 400px; 
                            box-shadow: 0 20px 60px rgba(0,0,0,0.4); border: 1px solid var(--border);">
                            <h3 style="margin: 0 0 20px 0; color: var(--text);">🔐 Passwort verwalten</h3>
                            
                            <div style="margin-bottom: 16px;">
                                <p style="color: var(--text-muted); margin-bottom: 12px;">Diese Datei ist passwortgeschützt.</p>
                            </div>
                            
                            <div style="margin-bottom: 16px;">
                                <label style="display: block; margin-bottom: 6px; color: var(--text);">Neues Passwort (leer lassen zum Entfernen):</label>
                                <input type="password" id="newPasswordInput" placeholder="Neues Passwort" 
                                    style="width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; 
                                    background: var(--background-light); color: var(--text); box-sizing: border-box;">
                            </div>
                            
                            <div style="margin-bottom: 20px;">
                                <label style="display: block; margin-bottom: 6px; color: var(--text);">Passwort bestätigen:</label>
                                <input type="password" id="confirmPasswordInput" placeholder="Passwort bestätigen" 
                                    style="width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; 
                                    background: var(--background-light); color: var(--text); box-sizing: border-box;">
                            </div>
                            
                            <p id="passwordError" style="color: #f44336; font-size: 12px; margin-bottom: 12px; display: none;"></p>
                            
                            <div style="display: flex; gap: 12px; justify-content: flex-end;">
                                <button id="cancelPasswordBtn" style="padding: 10px 20px; border: 1px solid var(--border); 
                                    background: var(--background-light); color: var(--text); border-radius: 6px; cursor: pointer;">
                                    Abbrechen
                                </button>
                                <button id="removePasswordBtn" style="padding: 10px 20px; border: none; 
                                    background: #f44336; color: white; border-radius: 6px; cursor: pointer;">
                                    🔓 Passwort entfernen
                                </button>
                                <button id="changePasswordBtn" style="padding: 10px 20px; border: none; 
                                    background: var(--primary); color: white; border-radius: 6px; cursor: pointer;">
                                    🔐 Passwort ändern
                                </button>
                            </div>
                        </div>
                    </div>
                `;
                
                document.body.insertAdjacentHTML('beforeend', dialogHTML);
                const dialog = document.getElementById('passwordManagementDialog');
                const newPasswordInput = dialog.querySelector('#newPasswordInput');
                const confirmPasswordInput = dialog.querySelector('#confirmPasswordInput');
                const passwordError = dialog.querySelector('#passwordError');
                
                newPasswordInput.focus();
                
                const closeDialog = (result) => {
                    dialog.remove();
                    resolve(result);
                };
                
                dialog.querySelector('#cancelPasswordBtn').onclick = () => closeDialog(null);
                dialog.onclick = (e) => { if (e.target === dialog) closeDialog(null); };
                
                dialog.querySelector('#removePasswordBtn').onclick = async () => {
                    if (confirm('Passwortschutz wirklich entfernen?')) {
                        if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                            const result = await window.electronAPI.liveSessionSetPassword(null);
                            if (result.success) {
                                explorerState.filePassword = null;
                                updatePasswordIndicator();
                                showFloatingStatus('🔓 Passwortschutz entfernt', 'success');
                                closeDialog({ removed: true });
                            } else {
                                passwordError.textContent = 'Fehler: ' + result.error;
                                passwordError.style.display = 'block';
                            }
                        } else {
                            explorerState.filePassword = null;
                            updatePasswordIndicator();
                            showFloatingStatus('🔓 Passwort wird beim Speichern entfernt', 'info');
                            closeDialog({ removed: true });
                        }
                    }
                };
                
                dialog.querySelector('#changePasswordBtn').onclick = async () => {
                    const newPass = newPasswordInput.value;
                    const confirmPass = confirmPasswordInput.value;
                    
                    if (!newPass) {
                        passwordError.textContent = 'Bitte ein neues Passwort eingeben';
                        passwordError.style.display = 'block';
                        return;
                    }
                    
                    if (newPass !== confirmPass) {
                        passwordError.textContent = 'Passwörter stimmen nicht überein';
                        passwordError.style.display = 'block';
                        return;
                    }
                    
                    if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                        const result = await window.electronAPI.liveSessionSetPassword(newPass);
                        if (result.success) {
                            explorerState.filePassword = newPass;
                            updatePasswordIndicator();
                            showFloatingStatus('🔐 Passwort geändert', 'success');
                            closeDialog({ changed: true });
                        } else {
                            passwordError.textContent = 'Fehler: ' + result.error;
                            passwordError.style.display = 'block';
                        }
                    } else {
                        explorerState.filePassword = newPass;
                        updatePasswordIndicator();
                        showFloatingStatus('🔐 Passwort wird beim Speichern angewendet', 'info');
                        closeDialog({ changed: true });
                    }
                };
            });
        }
        
        // Hilfsfunktion: Generiert einen eindeutigen Spaltennamen
        function getUniqueColumnName(baseName, headers) {
            // Prüfe ob der Name bereits existiert
            if (!headers.includes(baseName)) {
                return baseName;
            }
            
            // Suche nach dem höchsten Suffix für diesen Basisnamen
            let maxSuffix = 1;
            const basePattern = baseName.replace(/\d+$/, ''); // Entferne trailing numbers
            
            headers.forEach(h => {
                if (h === basePattern || h.startsWith(basePattern)) {
                    const suffix = h.slice(basePattern.length);
                    if (suffix === '') {
                        // Basisname ohne Suffix gefunden
                        maxSuffix = Math.max(maxSuffix, 1);
                    } else {
                        const num = parseInt(suffix, 10);
                        if (!isNaN(num)) {
                            maxSuffix = Math.max(maxSuffix, num);
                        }
                    }
                }
            });
            
            return basePattern + (maxSuffix + 1);
        }
        
        /**
         * Berechnet die physische Excel-Position aus dem logischen colIndex.
         * Nach moveColumn weichen logischer Index (headers[]) und physische Excel-Position ab.
         * columnOrder spiegelt die physische Reihenfolge wider.
         */
        function getExcelColumnPosition(colIndex) {
            if (explorerState.columnOrder.length > 0) {
                const pos = explorerState.columnOrder.indexOf(colIndex);
                if (pos !== -1) return pos;
            }
            return colIndex;
        }
        
        /**
         * Übersetzt einen Daten-Array-Index in die physische Excel-Zeilenposition.
         * Nach Zeilen-Verschiebungen (moveRows) kann der Index in explorerState.data
         * von der physischen Position in Excel abweichen.
         * @param {number} rowIndex - Index in explorerState.data (0-basiert)
         * @returns {number} Physische Zeilenposition (0-basiert, ohne Header)
         */
        function getExcelRowPosition(rowIndex) {
            if (explorerState.rowMapping && explorerState.rowMapping.length > rowIndex) {
                return explorerState.rowMapping[rowIndex];
            }
            return rowIndex;
        }
        
        // Neue Spalte einfügen
        async function insertExplorerColumn(colIndex, position = 'after') {
            const insertIndex = position === 'before' ? colIndex : colIndex + 1;
            
            // Generiere einen eindeutigen Default-Namen
            const defaultName = getUniqueColumnName(t('newColumn'), explorerState.headers);
            
            // Header-Namen abfragen
            const headerName = await showPromptDialog(
                t('newColumn'),
                t('enterColumnName'),
                defaultName
            );
            
            if (headerName === null) return; // Abgebrochen
            
            // Stelle sicher, dass der eingegebene Name eindeutig ist
            const uniqueHeaderName = getUniqueColumnName(headerName || defaultName, explorerState.headers);
            
            // Live-Session: Spalte sofort in Excel einfügen
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    const excelInsertIndex = getExcelColumnPosition(insertIndex);
                    const result = await window.electronAPI.liveSessionInsertColumn(excelInsertIndex, 1, [uniqueHeaderName]);
                    if (!result || !result.success) {
                        console.error('[LiveSession] insertColumn failed:', result);
                        showFloatingStatus('❌ Fehler beim Einfügen der Spalte in Excel', 'error');
                        return;
                    }
                    console.log('[LiveSession] Column inserted at index:', insertIndex);
                } catch (error) {
                    console.error('[LiveSession] insertColumn error:', error);
                    showFloatingStatus('❌ Fehler beim Einfügen der Spalte in Excel', 'error');
                    return;
                }
            }
            
            // Header einfügen
            explorerState.headers.splice(insertIndex, 0, uniqueHeaderName);
            
            // Alle Zeilen erweitern
            explorerState.data.forEach(row => {
                row.splice(insertIndex, 0, '');
            });
            explorerState.originalData.forEach(row => {
                row.splice(insertIndex, 0, '');
            });
            
            // VisibleColumns anpassen
            const newVisibleColumns = [];
            explorerState.visibleColumns.forEach(idx => {
                if (idx >= insertIndex) {
                    newVisibleColumns.push(idx + 1);
                } else {
                    newVisibleColumns.push(idx);
                }
            });
            newVisibleColumns.push(insertIndex); // Neue Spalte ist sichtbar
            newVisibleColumns.sort((a, b) => a - b);
            explorerState.visibleColumns = newVisibleColumns;
            
            // ColumnOrder anpassen (wenn aktiv)
            if (explorerState.columnOrder.length > 0) {
                explorerState.columnOrder = explorerState.columnOrder.map(idx => 
                    idx >= insertIndex ? idx + 1 : idx
                );
                // Neue Spalte an der richtigen Position einfügen
                let arrayPos = 0;
                while (arrayPos < explorerState.columnOrder.length && 
                       explorerState.columnOrder[arrayPos] < insertIndex) {
                    arrayPos++;
                }
                explorerState.columnOrder.splice(arrayPos, 0, insertIndex);
            }
            
            // EditedCells anpassen
            const newEditedCells = new Map();
            explorerState.editedCells.forEach((value, key) => {
                if (key.startsWith('_')) {
                    newEditedCells.set(key, value);
                    return;
                }
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col >= insertIndex) {
                    newEditedCells.set(`${row}-${col + 1}`, value);
                } else {
                    newEditedCells.set(key, value);
                }
            });
            explorerState.editedCells = newEditedCells;
            
            // CellStyles anpassen (Spalten-Indizes verschieben)
            const newCellStyles = {};
            for (const [key, style] of Object.entries(explorerState.cellStyles || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col >= insertIndex) {
                    newCellStyles[`${row}-${col + 1}`] = style;
                } else {
                    newCellStyles[key] = style;
                }
            }
            explorerState.cellStyles = newCellStyles;
            
            // CellFormulas anpassen
            const newCellFormulas = {};
            for (const [key, formula] of Object.entries(explorerState.cellFormulas || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col >= insertIndex) {
                    newCellFormulas[`${row}-${col + 1}`] = formula;
                } else {
                    newCellFormulas[key] = formula;
                }
            }
            explorerState.cellFormulas = newCellFormulas;
            
            // CellHyperlinks anpassen
            const newCellHyperlinks = {};
            for (const [key, link] of Object.entries(explorerState.cellHyperlinks || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col >= insertIndex) {
                    newCellHyperlinks[`${row}-${col + 1}`] = link;
                } else {
                    newCellHyperlinks[key] = link;
                }
            }
            explorerState.cellHyperlinks = newCellHyperlinks;
            
            // Markierung dass Spalte hinzugefügt wurde
            // Prüfe ob bereits eine insertedColumns-Info existiert (z.B. von Data Join)
            const existingInsert = explorerState.editedCells.get('_columnInserted');
            if (existingInsert && existingInsert.operations) {
                // Aktualisiere bestehende Operationen: Wenn die neue Spalte VOR einer bestehenden 
                // eingefügt wird, müssen deren position und sourceColumn erhöht werden
                existingInsert.operations.forEach(op => {
                    if (op.position >= insertIndex) {
                        op.position += 1;
                    }
                    if (op.sourceColumn >= insertIndex) {
                        op.sourceColumn += 1;
                    }
                });
                
                // Füge neue Operation hinzu
                existingInsert.operations.push({
                    position: insertIndex,
                    count: 1,
                    headers: [uniqueHeaderName],
                    sourceColumn: colIndex >= insertIndex ? colIndex + 1 : colIndex  // Referenzspalte für Formatierung (angepasst)
                });
                existingInsert.totalCount = (existingInsert.totalCount || 0) + 1;
                explorerState.editedCells.set('_columnInserted', existingInsert);
            } else {
                // Neue insertedColumns-Info mit korrektem Format
                explorerState.editedCells.set('_columnInserted', { 
                    operations: [{
                        position: insertIndex,
                        count: 1,
                        headers: [uniqueHeaderName],
                        sourceColumn: colIndex  // Referenzspalte für Formatierung
                    }],
                    totalCount: 1
                });
            }
            
            // UI aktualisieren
            filterExplorerData();
            updateColumnToggles();
            showFloatingStatus(t('columnInserted'));
        }
        
        // Spalte löschen
        async function deleteExplorerColumn(colIndex) {
            const columnName = explorerState.headers[colIndex] || `${currentLanguage === 'en' ? 'Column' : 'Spalte'} ${colIndex + 1}`;
            const columnLabel = currentLanguage === 'en' ? 'Column' : 'Spalte';
            
            const confirmed = await showConfirmDialog(
                t('deleteColumnTitle'),
                `${t('deleteColumnConfirm')}\n\n${columnLabel}: "${columnName}"\n\n${t('deleteColumnWarning')}`,
                currentLanguage === 'en' ? 'Delete' : 'Löschen',
                t('cancel')
            );
            
            if (!confirmed) return;
            
            // LIVE SESSION: Lösche Spalte sofort in Excel
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                const excelColIndex = getExcelColumnPosition(colIndex);
                await liveSessionExecute('deleteColumn', excelColIndex);
                console.log(`[LiveSession] Spalte ${colIndex} (Excel-Position ${excelColIndex}) gelöscht`);
            }
            
            // Header entfernen
            explorerState.headers.splice(colIndex, 1);
            
            // Alle Zeilen anpassen
            explorerState.data.forEach(row => {
                row.splice(colIndex, 1);
            });
            explorerState.originalData.forEach(row => {
                row.splice(colIndex, 1);
            });
            
            // VisibleColumns anpassen
            explorerState.visibleColumns = explorerState.visibleColumns
                .filter(idx => idx !== colIndex)
                .map(idx => idx > colIndex ? idx - 1 : idx);
            
            // EditedCells anpassen
            const newEditedCells = new Map();
            explorerState.editedCells.forEach((value, key) => {
                if (key.startsWith('_')) {
                    newEditedCells.set(key, value);
                    return;
                }
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col > colIndex) {
                    newEditedCells.set(`${row}-${col - 1}`, value);
                } else if (col < colIndex) {
                    newEditedCells.set(key, value);
                }
                // col === colIndex wird nicht übernommen
            });
            explorerState.editedCells = newEditedCells;
            
            // CellStyles anpassen
            const newCellStyles = {};
            for (const [key, value] of Object.entries(explorerState.cellStyles || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col > colIndex) {
                    newCellStyles[`${row}-${col - 1}`] = value;
                } else if (col < colIndex) {
                    newCellStyles[key] = value;
                }
            }
            explorerState.cellStyles = newCellStyles;
            
            // CellFormulas anpassen
            const newCellFormulas = {};
            for (const [key, value] of Object.entries(explorerState.cellFormulas || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col > colIndex) {
                    newCellFormulas[`${row}-${col - 1}`] = value;
                } else if (col < colIndex) {
                    newCellFormulas[key] = value;
                }
            }
            explorerState.cellFormulas = newCellFormulas;
            
            // CellHyperlinks anpassen
            const newCellHyperlinks = {};
            for (const [key, value] of Object.entries(explorerState.cellHyperlinks || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col > colIndex) {
                    newCellHyperlinks[`${row}-${col - 1}`] = value;
                } else if (col < colIndex) {
                    newCellHyperlinks[key] = value;
                }
            }
            explorerState.cellHyperlinks = newCellHyperlinks;
            
            // RichTextCells anpassen
            const newRichTextCells = {};
            for (const [key, value] of Object.entries(explorerState.richTextCells || {})) {
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (col > colIndex) {
                    newRichTextCells[`${row}-${col - 1}`] = value;
                } else if (col < colIndex) {
                    newRichTextCells[key] = value;
                }
            }
            explorerState.richTextCells = newRichTextCells;
            
            // ColumnOrder anpassen (wenn aktiv)
            if (explorerState.columnOrder.length > 0) {
                explorerState.columnOrder = explorerState.columnOrder
                    .filter(idx => idx !== colIndex)
                    .map(idx => idx > colIndex ? idx - 1 : idx);
            }
            
            // SortColumn anpassen
            if (explorerState.sortColumn !== null) {
                if (explorerState.sortColumn === colIndex) {
                    explorerState.sortColumn = null;
                    explorerState.sortDirection = null;
                    explorerState.sortType = 'auto';
                } else if (explorerState.sortColumn > colIndex) {
                    explorerState.sortColumn--;
                }
            }
            
            // DataValidations anpassen (Index-basiert)
            if (explorerState.dataValidations) {
                const newValidations = {};
                for (const [key, value] of Object.entries(explorerState.dataValidations)) {
                    const idx = parseInt(key);
                    if (idx > colIndex) {
                        newValidations[idx - 1] = value;
                    } else if (idx < colIndex) {
                        newValidations[key] = value;
                    }
                    // idx === colIndex wird nicht übernommen
                }
                explorerState.dataValidations = newValidations;
            }
            
            // Filter anpassen (Spalten-Referenzen)
            if (explorerState.filters && explorerState.filters.length > 0) {
                explorerState.filters = explorerState.filters.filter(f => {
                    if (!f.column) return true;
                    return parseInt(f.column) !== colIndex;
                }).map(f => {
                    if (!f.column) return f;
                    const col = parseInt(f.column);
                    if (col > colIndex) {
                        return { ...f, column: String(col - 1) };
                    }
                    return f;
                });
            }
            
            // ========== _columnInserted anpassen und prüfen ob eingefügte Spalte gelöscht wird ==========
            const insertedInfo = explorerState.editedCells.get('_columnInserted');
            let isDeletingInsertedColumn = false;
            
            if (insertedInfo && insertedInfo.operations) {
                // Prüfe ob die gelöschte Spalte eine EINGEFÜGTE Spalte ist (nicht aus der Originaldatei)
                const insertedOpIndex = insertedInfo.operations.findIndex(op => op.position === colIndex);
                
                if (insertedOpIndex !== -1) {
                    isDeletingInsertedColumn = true;
                    // Entferne die Operation für die gelöschte eingefügte Spalte
                    insertedInfo.operations.splice(insertedOpIndex, 1);
                    console.log(`[deleteColumn] Eingefügte Spalte an Position ${colIndex} wird gelöscht (kein Original-Index nötig)`);
                }
                
                // Positionen aller verbleibenden Operationen anpassen (die nach der gelöschten liegen)
                insertedInfo.operations.forEach(op => {
                    if (op.position > colIndex) {
                        op.position -= 1;
                    }
                    if (op.sourceColumn > colIndex) {
                        op.sourceColumn -= 1;
                    }
                });
                
                insertedInfo.totalCount = insertedInfo.operations.length;
                
                if (insertedInfo.operations.length === 0) {
                    // Keine eingefügten Spalten mehr übrig
                    explorerState.editedCells.delete('_columnInserted');
                } else {
                    explorerState.editedCells.set('_columnInserted', insertedInfo);
                }
            }
            
            if (!isDeletingInsertedColumn) {
                // ========== Original-Spalte gelöscht: _columnDeleted aktualisieren ==========
                const existingDeleted = explorerState.editedCells.get('_columnDeleted');
                let deletedOriginalIndices = [];
                if (existingDeleted && Array.isArray(existingDeleted.originalIndices)) {
                    deletedOriginalIndices = existingDeleted.originalIndices.slice();
                }
                
                // Berechne den Original-Index:
                // 1. Eingefügte Spalten VOR colIndex abziehen (die existieren nicht in der Originaldatei)
                let insertsBefore = 0;
                if (insertedInfo && insertedInfo.operations) {
                    // Positionen < colIndex sind nach der Anpassung unverändert (nur > colIndex wurde dekrementiert)
                    insertsBefore = insertedInfo.operations.filter(op => op.position < colIndex).length;
                }
                
                let originalIndex = colIndex - insertsBefore;
                
                // 2. Für jeden vorherigen gelöschten Original-Index, der <= aktuellem+Offset war, erhöhe den Offset
                for (const prevOrigIdx of deletedOriginalIndices.sort((a, b) => a - b)) {
                    if (prevOrigIdx <= originalIndex) {
                        originalIndex++;
                    }
                }
                
                deletedOriginalIndices.push(originalIndex);
                
                explorerState.editedCells.set('_columnDeleted', { 
                    originalIndices: deletedOriginalIndices,  // Array der ORIGINAL-Spalten-Indices (0-basiert)
                    count: deletedOriginalIndices.length,
                    originalHeaderCount: (existingDeleted?.originalHeaderCount) || (explorerState.headers.length + 1)
                });
                
                // Zur Operations-Queue hinzufügen
                explorerState.columnOperationsQueue.push({
                    type: 'delete',
                    originalIndex: originalIndex
                });
            }
            
            // UI aktualisieren
            filterExplorerData();
            updateColumnToggles();
            showFloatingStatus(t('columnDeleted'));
        }
        
        // Prompt-Dialog für Texteingabe
        function showPromptDialog(title, message, defaultValue = '', inputType = 'text') {
            return new Promise((resolve) => {
                const existingDialog = document.querySelector('.prompt-dialog-overlay');
                if (existingDialog) existingDialog.remove();
                
                const overlay = document.createElement('div');
                overlay.className = 'prompt-dialog-overlay modal-overlay';
                overlay.style.zIndex = '99999';
                
                const closeDialog = (result) => {
                    overlay.remove();
                    resolve(result);
                };
                
                overlay.innerHTML = `
                    <div class="modal" style="max-width: 400px; width: 90%; min-width: 300px; min-height: auto;">
                        <div class="modal-header">
                            <div class="modal-title">${title}</div>
                        </div>
                        <div class="modal-body" style="padding: 20px;">
                            <p style="margin: 0 0 15px 0; color: var(--text-muted); white-space: pre-line;">${message}</p>
                            <input type="${inputType}" id="promptInput" value="${defaultValue}" style="width: 100%; padding: 12px; box-sizing: border-box; font-size: 14px; background: var(--bg-dark); border: 2px solid var(--border); border-radius: 6px; color: var(--text);">
                        </div>
                        <div class="modal-footer">
                            <button class="btn btn-secondary" id="promptCancel">${currentLanguage === 'en' ? 'Cancel' : 'Abbrechen'}</button>
                            <button class="btn btn-primary" id="promptConfirm">${currentLanguage === 'en' ? 'OK' : 'OK'}</button>
                        </div>
                    </div>
                `;
                
                document.body.appendChild(overlay);
                
                const input = document.getElementById('promptInput');
                input.focus();
                input.select();
                
                input.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') {
                        closeDialog(input.value);
                    } else if (e.key === 'Escape') {
                        closeDialog(null);
                    }
                });
                
                document.getElementById('promptCancel').onclick = () => closeDialog(null);
                document.getElementById('promptConfirm').onclick = () => closeDialog(input.value);
            });
        }
        
        // ==================== Zellen-Auswahl Funktionen ====================
        
        // Zellen-Auswahl visuell aktualisieren
        function updateCellSelectionUI() {
            // Alle Auswahl-Markierungen entfernen
            document.querySelectorAll('#explorerTableBody td.cell-selected').forEach(td => {
                td.classList.remove('cell-selected', 'cell-selection-anchor');
            });
            
            // Ausgewählte Zellen markieren
            explorerState.selectedCells.forEach(cellKey => {
                const [rowIndex, colIndex] = cellKey.split('-').map(Number);
                const td = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                if (td) {
                    td.classList.add('cell-selected');
                }
            });
            
            // Anker-Zelle hervorheben
            if (explorerState.selectionAnchor) {
                const anchorTd = document.querySelector(
                    `#explorerTableBody td[data-row="${explorerState.selectionAnchor.row}"][data-col="${explorerState.selectionAnchor.col}"]`
                );
                if (anchorTd) {
                    anchorTd.classList.add('cell-selection-anchor');
                }
            }
        }
        
        // Zellen-Auswahl leeren
        function clearCellSelection() {
            explorerState.selectedCells.clear();
            explorerState.selectionAnchor = null;
            explorerState.isSelecting = false;
            updateCellSelectionUI();
        }
        
        // Zelle zur Auswahl hinzufügen/entfernen
        function toggleCellSelection(rowIndex, colIndex, addToSelection = false) {
            const cellKey = `${rowIndex}-${colIndex}`;
            
            if (!addToSelection) {
                // Ohne Modifier: Alte Auswahl leeren
                explorerState.selectedCells.clear();
            }
            
            if (explorerState.selectedCells.has(cellKey)) {
                explorerState.selectedCells.delete(cellKey);
            } else {
                explorerState.selectedCells.add(cellKey);
            }
            
            explorerState.selectionAnchor = { row: rowIndex, col: colIndex };
            updateCellSelectionUI();
        }
        
        // Bereich von Zellen auswählen (Shift-Klick)
        function selectCellRange(toRow, toCol) {
            if (!explorerState.selectionAnchor) {
                explorerState.selectionAnchor = { row: toRow, col: toCol };
            }
            
            const fromRow = explorerState.selectionAnchor.row;
            const fromCol = explorerState.selectionAnchor.col;
            
            const minRow = Math.min(fromRow, toRow);
            const maxRow = Math.max(fromRow, toRow);
            const minCol = Math.min(fromCol, toCol);
            const maxCol = Math.max(fromCol, toCol);
            
            // Alte Auswahl leeren und neuen Bereich auswählen
            explorerState.selectedCells.clear();
            
            // Nur sichtbare Spalten berücksichtigen
            const displayColumns = explorerState.columnOrder.length > 0 
                ? explorerState.columnOrder.filter(col => explorerState.visibleColumns.includes(col))
                : explorerState.visibleColumns;
            
            // Alle Zeilen im Bereich finden
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    // Nur auswählen wenn die Spalte sichtbar ist
                    if (displayColumns.includes(c)) {
                        explorerState.selectedCells.add(`${r}-${c}`);
                    }
                }
            }
            
            updateCellSelectionUI();
        }
        
        // Kontextmenü für Zellen-Auswahl
        function showCellContextMenu(e, rowIndex, colIndex) {
            e.preventDefault();
            e.stopPropagation();
            
            // Altes Menü entfernen (nur cell-context-menu, nicht das column-context-menu)
            const oldMenu = document.querySelector('.cell-context-menu');
            if (oldMenu) oldMenu.remove();
            
            const cellKey = `${rowIndex}-${colIndex}`;
            
            // Wenn die angeklickte Zelle nicht in der Auswahl ist, diese zur Auswahl machen
            if (!explorerState.selectedCells.has(cellKey)) {
                explorerState.selectedCells.clear();
                explorerState.selectedCells.add(cellKey);
                explorerState.selectionAnchor = { row: rowIndex, col: colIndex };
                updateCellSelectionUI();
            }
            
            const selectedCount = explorerState.selectedCells.size;
            const cellLabel = selectedCount === 1 ? 'Zelle' : `${selectedCount} Zellen`;
            
            let menuHtml = '<div class="context-menu cell-context-menu">';
            menuHtml += `<div class="context-menu-item" style="font-weight: 600; color: var(--text-muted); cursor: default;">📋 ${cellLabel} ausgewählt</div>`;
            menuHtml += '<div class="context-menu-divider"></div>';
            menuHtml += `<div class="context-menu-item" data-action="copy-content">📋 Kopieren</div>`;
            menuHtml += '<div class="context-menu-divider"></div>';
            menuHtml += `<div class="context-menu-item" data-action="paste-content">📥 Einfügen (nur Werte)</div>`;
            const hasFormatClipboard = copiedCellsWithFormat !== null;
            const cfBlocked = explorerState.hasConditionalFormatting;
            const formatDisabled = !hasFormatClipboard || cfBlocked;
            const formatLabel = cfBlocked ? '🎨 Einfügen mit Formatierung (CF → deaktiviert)' : hasFormatClipboard ? '🎨 Einfügen mit Formatierung' : '🎨 Einfügen mit Formatierung (leer)';
            menuHtml += `<div class="context-menu-item${formatDisabled ? ' disabled' : ''}" data-action="paste-with-format" style="${formatDisabled ? 'opacity: 0.5; cursor: not-allowed;' : ''}">${formatLabel}</div>`;
            menuHtml += '<div class="context-menu-divider"></div>';
            menuHtml += `<div class="context-menu-item" data-action="delete-content">🗑️ Inhalt löschen</div>`;
            menuHtml += '<div class="context-menu-divider"></div>';
            menuHtml += `<div class="context-menu-item" data-action="clear-selection">❌ Auswahl aufheben</div>`;
            menuHtml += '</div>';
            
            const menu = document.createElement('div');
            menu.innerHTML = menuHtml;
            document.body.appendChild(menu.firstElementChild);
            
            const menuEl = document.querySelector('.cell-context-menu');
            
            // Positionierung
            let x = e.clientX;
            let y = e.clientY;
            
            // Sicherstellen, dass das Menü im Viewport bleibt
            const menuRect = menuEl.getBoundingClientRect();
            if (x + menuRect.width > window.innerWidth) {
                x = window.innerWidth - menuRect.width - 5;
            }
            if (y + menuRect.height > window.innerHeight) {
                y = window.innerHeight - menuRect.height - 5;
            }
            
            menuEl.style.left = x + 'px';
            menuEl.style.top = y + 'px';
            
            // Event-Handler: Inhalt löschen
            menuEl.querySelector('[data-action="delete-content"]').addEventListener('click', () => {
                deleteSelectedCellsContent();
                menuEl.remove();
            });
            
            // Event-Handler: Kopieren (immer mit Formatierung)
            menuEl.querySelector('[data-action="copy-content"]').addEventListener('click', () => {
                copySelectedCellsWithFormat();
                menuEl.remove();
            });
            
            // Event-Handler: Einfügen (nur Werte)
            menuEl.querySelector('[data-action="paste-content"]').addEventListener('click', async () => {
                await pasteToSelectedCells();
                menuEl.remove();
            });
            
            // Event-Handler: Einfügen mit Formatierung
            const pasteFormatItem = menuEl.querySelector('[data-action="paste-with-format"]');
            if (pasteFormatItem && copiedCellsWithFormat !== null && !explorerState.hasConditionalFormatting) {
                pasteFormatItem.addEventListener('click', () => {
                    pasteSelectedCellsWithFormat();
                    menuEl.remove();
                });
            }
            
            // Event-Handler: Auswahl aufheben
            menuEl.querySelector('[data-action="clear-selection"]').addEventListener('click', () => {
                clearCellSelection();
                menuEl.remove();
            });
            
            // Menü bei Klick außerhalb schließen
            const closeHandler = (event) => {
                if (!menuEl.contains(event.target)) {
                    menuEl.remove();
                    document.removeEventListener('click', closeHandler);
                }
            };
            setTimeout(() => document.addEventListener('click', closeHandler), 10);
        }
        
        // Ausgewählte Zelleninhalte löschen
        function deleteSelectedCellsContent() {
            if (explorerState.selectedCells.size === 0) return;
            
            const undoActions = [];
            
            explorerState.selectedCells.forEach(cellKey => {
                const [rowIndex, colIndex] = cellKey.split('-').map(Number);
                const td = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                
                if (td) {
                    const oldValue = explorerState.data[rowIndex][colIndex];
                    const originalValue = td.dataset.original;
                    
                    // Undo-Aktion speichern
                    undoActions.push({
                        rowIndex,
                        colIndex,
                        oldValue: String(oldValue ?? ''),
                        newValue: '',
                        originalValue
                    });
                    
                    // Daten aktualisieren
                    explorerState.data[rowIndex][colIndex] = '';
                    explorerState.editedCells.set(cellKey, '');
                    
                    // UI aktualisieren
                    td.textContent = '';
                    td.dataset.lastValue = '';
                    td.classList.add('edited');
                }
            });
            
            // Alle Änderungen als eine Undo-Aktion speichern
            if (undoActions.length > 0) {
                pushExplorerUndo({
                    type: 'multi',
                    actions: undoActions
                });
            }
            
            // Live-Session: Nur betroffene Zellen an Excel syncen (nicht ganze Zeilen)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady && undoActions.length > 0) {
                const cells = undoActions.map(a => ({ row: a.rowIndex, col: a.colIndex, value: '', oldValue: a.oldValue }));
                console.log('[CellEdit] deleteSelectedCellsContent: Sync', cells.length, 'Zellen via Batch');
                window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(cells))
                    .then(res => console.log('[CellEdit] Delete-Batch Ergebnis:', JSON.stringify(res)))
                    .catch(err => console.error('[CellEdit] Delete-Batch fehlgeschlagen:', err));
            }
            
            // Status aktualisieren
            updateExplorerEditStatus();
            
            // Auswahl leeren
            clearCellSelection();
            
            // Info anzeigen
            showFloatingStatus(`${undoActions.length} Zelle(n) gelöscht`);
        }
        
        // Ausgewählte Zelleninhalte kopieren (immer inkl. Formatierung)
        function copySelectedCellsWithFormat() {
            if (explorerState.selectedCells.size === 0) return;
            
            // Auswahl erweitern: Wenn ein Merge-Bereich berührt wird, alle Zellen einschließen
            const expandedSelection = new Set(explorerState.selectedCells);
            for (const key of explorerState.selectedCells) {
                const [row, col] = key.split('-').map(Number);
                const excelRow0 = row + 1;  // Daten-Index → Excel-0-basiert
                for (const merge of explorerState.mergedCells) {
                    if (excelRow0 >= merge.startRow && excelRow0 <= merge.endRow &&
                        col >= merge.startCol && col <= merge.endCol) {
                        // Diese Zelle ist Teil eines Merges → alle Zellen des Merges einschließen
                        for (let r = merge.startRow; r <= merge.endRow; r++) {
                            for (let c = merge.startCol; c <= merge.endCol; c++) {
                                const dataRow = r - 1;  // Excel-0-basiert → Daten-Index
                                if (dataRow >= 0 && dataRow < explorerState.data.length) {
                                    expandedSelection.add(`${dataRow}-${c}`);
                                }
                            }
                        }
                    }
                }
            }
            
            // Zellen nach Position sortieren und mit Formatierung erfassen
            const cells = Array.from(expandedSelection).map(key => {
                const [row, col] = key.split('-').map(Number);
                // WICHTIG: cellStyles, cellFormulas etc. verwenden "originalIndex+1"-Format (1-basiert wegen Header)
                const styleKey = `${row + 1}-${col}`;
                let style = explorerState.cellStyles[styleKey] ? { ...explorerState.cellStyles[styleKey] } : null;
                const richText = explorerState.richTextCells[styleKey] || null;
                
                // Wenn kein cellStyle vorhanden aber richText, dann Style aus richText ableiten
                // (ExcelJS speichert Font-Infos bei Rich Text nur in den Runs, nicht im cellStyles)
                if (!style && richText && richText.length > 0) {
                    // Dominanten Style ermitteln (längster Text-Run oder erster Run mit Styles)
                    const dominantRun = richText.reduce((best, run) => 
                        (run.text || '').length > (best.text || '').length ? run : best
                    , richText[0]);
                    
                    if (dominantRun && dominantRun.styles) {
                        const s = dominantRun.styles;
                        style = {};
                        if (s.bold) style.bold = true;
                        if (s.italic) style.italic = true;
                        if (s.underline) style.underline = true;
                        if (s.strikethrough) style.strikethrough = true;
                        if (s.color && s.color !== '#000000') style.fontColor = s.color;
                        if (s.fontSize) style.fontSize = s.fontSize;
                        if (s.fontName) style.fontName = s.fontName;
                        // Nur wenn mindestens eine Style-Eigenschaft gefunden wurde
                        if (Object.keys(style).length === 0) style = null;
                    }
                }
                
                const vmVal = (explorerState.cellVmMap && explorerState.cellVmMap[styleKey]) || null;
                const cellValue = explorerState.data[row][col] ?? '';
                return {
                    row,
                    col,
                    value: cellValue,
                    style,
                    formula: explorerState.cellFormulas[styleKey] || null,
                    hyperlink: explorerState.cellHyperlinks[styleKey] || null,
                    richText,
                    vm: vmVal
                };
            }).sort((a, b) => a.row - b.row || a.col - b.col);
            
            // Minimale Position ermitteln (für relatives Einfügen)
            const minRow = Math.min(...cells.map(c => c.row));
            const minCol = Math.min(...cells.map(c => c.col));
            
            // Relative Positionen berechnen
            copiedCellsWithFormat = {
                cells: cells.map(c => ({
                    ...c,
                    relRow: c.row - minRow,
                    relCol: c.col - minCol
                })),
                minRow,
                minCol,
                sourceSheet: explorerState.selectedSheet,
                // Merged Cells im kopierten Bereich erfassen (relativ zur Auswahl)
                mergedCells: explorerState.mergedCells
                    .filter(m => {
                        // mergedCells.startRow ist 0-basierter Excel-Index (0 = Zeile 1 = Header)
                        // Daten-Zeile 0 = Excel-Index 1
                        const mDataStartRow = m.startRow - 1;
                        const mDataEndRow = m.endRow - 1;
                        const maxRow = Math.max(...cells.map(c => c.row));
                        const maxCol = Math.max(...cells.map(c => c.col));
                        return mDataStartRow >= minRow && mDataEndRow <= maxRow &&
                               m.startCol >= minCol && m.endCol <= maxCol;
                    })
                    .map(m => ({
                        relStartRow: (m.startRow - 1) - minRow,
                        relStartCol: m.startCol - minCol,
                        relEndRow: (m.endRow - 1) - minRow,
                        relEndCol: m.endCol - minCol,
                        rowSpan: m.rowSpan,
                        colSpan: m.colSpan
                    }))
            };
            
            // Auch als Text in Zwischenablage kopieren
            const rows = new Map();
            cells.forEach(cell => {
                if (!rows.has(cell.row)) {
                    rows.set(cell.row, []);
                }
                rows.get(cell.row).push(String(cell.value));
            });
            
            const text = Array.from(rows.values())
                .map(rowCells => rowCells.join('\t'))
                .join('\n');
            
            navigator.clipboard.writeText(text).then(() => {
                showFloatingStatus(`${explorerState.selectedCells.size} Zelle(n) kopiert`);
            }).catch(err => {
                console.error('Kopieren fehlgeschlagen:', err);
                showFloatingStatus('Kopieren fehlgeschlagen', true);
            });
        }
        
        // Zellen MIT Formatierung einfügen
        function pasteSelectedCellsWithFormat() {
            if (!copiedCellsWithFormat || copiedCellsWithFormat.cells.length === 0) {
                showFloatingStatus('Keine formatierten Zellen zum Einfügen', true);
                return;
            }
            
            // Zielposition ermitteln (erste ausgewählte Zelle oder Anker)
            let targetRow, targetCol;
            if (explorerState.selectedCells.size > 0) {
                const firstCell = explorerState.selectedCells.values().next().value;
                [targetRow, targetCol] = firstCell.split('-').map(Number);
            } else if (explorerState.selectionAnchor) {
                targetRow = explorerState.selectionAnchor.row;
                targetCol = explorerState.selectionAnchor.col;
            } else {
                showFloatingStatus('Keine Zielzelle ausgewählt', true);
                return;
            }
            
            const undoActions = [];
            let pastedCount = 0;
            
            copiedCellsWithFormat.cells.forEach(cell => {
                const newRow = targetRow + cell.relRow;
                const newCol = targetCol + cell.relCol;
                const cellKey = `${newRow}-${newCol}`;  // Für editedCells (0-basiert)
                const styleKey = `${newRow + 1}-${newCol}`;  // Für Styles (1-basiert wegen Header)
                
                // Prüfen ob Zielzelle existiert
                if (newRow >= 0 && newRow < explorerState.data.length && 
                    newCol >= 0 && newCol < explorerState.headers.length) {
                    
                    const oldValue = explorerState.data[newRow][newCol];
                    const oldStyle = explorerState.cellStyles[styleKey];
                    const oldFormula = explorerState.cellFormulas[styleKey];
                    const oldHyperlink = explorerState.cellHyperlinks[styleKey];
                    const oldRichText = explorerState.richTextCells[styleKey];
                    
                    // Undo-Aktion speichern
                    undoActions.push({
                        rowIndex: newRow,
                        colIndex: newCol,
                        oldValue: oldValue ?? '',
                        newValue: cell.value,
                        oldStyle: oldStyle ? { ...oldStyle } : null,
                        newStyle: cell.style ? { ...cell.style } : null,
                        oldFormula,
                        newFormula: cell.formula,
                        oldHyperlink,
                        newHyperlink: cell.hyperlink,
                        oldRichText: oldRichText ? [...oldRichText] : null,
                        newRichText: cell.richText ? [...cell.richText] : null
                    });
                    
                    // Wert setzen (bei vm-Zellen immer Platzhalter verwenden)
                    const pasteValue = cell.vm ? '🖼️ Bild' : cell.value;
                    explorerState.data[newRow][newCol] = pasteValue;
                    explorerState.editedCells.set(cellKey, pasteValue);
                    
                    // Style kopieren (mit styleKey = 1-basiert)
                    if (cell.style) {
                        explorerState.cellStyles[styleKey] = { ...cell.style };
                        
                        // Font-Infos auch in cellFonts speichern (Backend erwartet diese separat)
                        const fontInfo = {};
                        if (cell.style.fontName) fontInfo.name = cell.style.fontName;
                        if (cell.style.fontSize) fontInfo.size = cell.style.fontSize;
                        if (cell.style.bold) fontInfo.bold = cell.style.bold;
                        if (cell.style.italic) fontInfo.italic = cell.style.italic;
                        if (cell.style.fontColor) fontInfo.color = cell.style.fontColor;
                        if (Object.keys(fontInfo).length > 0) {
                            if (!explorerState.cellFonts) explorerState.cellFonts = {};
                            explorerState.cellFonts[styleKey] = fontInfo;
                        }
                    } else {
                        delete explorerState.cellStyles[styleKey];
                        if (explorerState.cellFonts) delete explorerState.cellFonts[styleKey];
                    }
                    
                    // Formel kopieren
                    if (cell.formula) {
                        explorerState.cellFormulas[styleKey] = cell.formula;
                    } else {
                        delete explorerState.cellFormulas[styleKey];
                    }
                    
                    // Hyperlink kopieren
                    if (cell.hyperlink) {
                        explorerState.cellHyperlinks[styleKey] = cell.hyperlink;
                    } else {
                        delete explorerState.cellHyperlinks[styleKey];
                    }
                    
                    // Rich Text kopieren
                    if (cell.richText) {
                        explorerState.richTextCells[styleKey] = [...cell.richText];
                    } else {
                        delete explorerState.richTextCells[styleKey];
                    }
                    
                    // VM-Attribut kopieren (Bild-Referenz für Zell-Bilder)
                    if (cell.vm) {
                        if (!explorerState.cellVmMap) explorerState.cellVmMap = {};
                        explorerState.cellVmMap[styleKey] = cell.vm;
                        // Bild-Zellen brauchen center-Alignment (Excel zentriert Zellbilder)
                        if (!explorerState.cellStyles[styleKey]) {
                            explorerState.cellStyles[styleKey] = {};
                        }
                        if (!explorerState.cellStyles[styleKey].textAlign) {
                            explorerState.cellStyles[styleKey].textAlign = 'center';
                            explorerState.cellStyles[styleKey].verticalAlign = 'middle';
                        }
                    }
                    
                    pastedCount++;
                }
            });
            
            // Merged Cells am Ziel anlegen (GUI-seitig)
            // Alte Merges im Zielbereich entfernen
            const maxRelRow = Math.max(...copiedCellsWithFormat.cells.map(c => c.relRow));
            const maxRelCol = Math.max(...copiedCellsWithFormat.cells.map(c => c.relCol));
            explorerState.mergedCells = explorerState.mergedCells.filter(m => {
                const mDataStartRow = m.startRow - 1;
                const mDataEndRow = m.endRow - 1;
                const overlaps = mDataStartRow <= targetRow + maxRelRow && mDataEndRow >= targetRow &&
                                 m.startCol <= targetCol + maxRelCol && m.endCol >= targetCol;
                return !overlaps;
            });
            // Kopierte Merges am Ziel einfügen (aus der Kopier-Quelle)
            if (copiedCellsWithFormat.mergedCells && copiedCellsWithFormat.mergedCells.length > 0) {
                copiedCellsWithFormat.mergedCells.forEach(rm => {
                    explorerState.mergedCells.push({
                        startRow: (targetRow + rm.relStartRow) + 1,
                        startCol: targetCol + rm.relStartCol,
                        endRow: (targetRow + rm.relEndRow) + 1,
                        endCol: targetCol + rm.relEndCol,
                        rowSpan: rm.rowSpan,
                        colSpan: rm.colSpan
                    });
                });
            }
            
            // Formatierungsänderung markieren (Style/Merge/RichText von Paste)
            explorerState.editedCells.set('_hasFormatChanges', true);
            
            // Undo speichern
            if (undoActions.length > 0) {
                pushExplorerUndo({
                    type: 'multi-format',
                    actions: undoActions
                });
            }
            
            // Live Session Sync - Native Excel Copy + Merged Cells
            if (explorerState.liveSessionActive && explorerState.liveSessionReady && pastedCount > 0) {
                const sourceCells = copiedCellsWithFormat.cells.map(c => ({ row: c.row, col: c.col }));
                window.electronAPI.liveSessionCopyCells(sourceCells, targetRow, targetCol)
                    .then(res => {
                        if (res && res.success) {
                            // Excel hat die Merges nativ angelegt - GUI-Merges mit Excel-Antwort aktualisieren
                            if (res.mergedCells && res.mergedCells.length > 0) {
                                // Entferne vorherige GUI-Merges im Zielbereich (wurden oben schon bereinigt)
                                // und ersetze mit den tatsächlichen Excel-Merges
                                explorerState.mergedCells = explorerState.mergedCells.filter(m => {
                                    const mDataStartRow = m.startRow - 1;
                                    const mDataEndRow = m.endRow - 1;
                                    const overlaps = mDataStartRow <= targetRow + maxRelRow && mDataEndRow >= targetRow &&
                                                     m.startCol <= targetCol + maxRelCol && m.endCol >= targetCol;
                                    return !overlaps;
                                });
                                res.mergedCells.forEach(em => {
                                    explorerState.mergedCells.push(em);
                                });
                                renderExplorerTable();
                            }
                            showFloatingStatus(`Excel-Sync: ${res.count} Zellen kopiert, ${res.mergedCells ? res.mergedCells.length : 0} Merges ✓`);
                        }
                    })
                    .catch(err => console.error('[Paste+Format] Copy fehlgeschlagen:', err));
            }
            
            // Tabelle neu rendern um Formatierung anzuzeigen
            renderExplorerTable();
            updateExplorerEditStatus();
            
            showFloatingStatus(`${pastedCount} Zelle(n) mit Formatierung eingefügt 🎨`);
        }
        
        // Aus Zwischenablage in ausgewählte Zelle(n) einfügen
        async function pasteToSelectedCells() {
            try {
                const text = await navigator.clipboard.readText();
                if (!text) {
                    showFloatingStatus('Zwischenablage ist leer', true);
                    return;
                }
                
                // Wenn nur eine Zelle ausgewählt ist, dort einfügen
                if (explorerState.selectedCells.size === 1) {
                    const cellKey = explorerState.selectedCells.values().next().value;
                    const [rowIndex, colIndex] = cellKey.split('-').map(Number);
                    const td = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                    
                    if (td) {
                        const original = td.dataset.original;
                        const oldValue = explorerState.data[rowIndex][colIndex];
                        
                        // Wert setzen
                        explorerState.data[rowIndex][colIndex] = text;
                        explorerState.editedCells.set(cellKey, text);
                        
                        // UI aktualisieren
                        const contentSpan = td.querySelector('.cell-content');
                        if (contentSpan) {
                            contentSpan.textContent = text;
                        } else {
                            td.textContent = text;
                        }
                        
                        if (text !== original) {
                            td.classList.add('edited');
                        } else {
                            td.classList.remove('edited');
                        }
                        
                        // Undo speichern
                        pushExplorerUndo({
                            rowIndex,
                            colIndex,
                            oldValue: oldValue,
                            newValue: text,
                            originalValue: original
                        });
                        
                        // Live Session Sync (wenn aktiv)
                        if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                            console.log(`[Paste] Live-Sync: Zeile ${rowIndex}, Spalte ${colIndex}`);
                            window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols([{
                                row: rowIndex,
                                col: colIndex,
                                value: text,
                                oldValue: oldValue
                            }]))
                                .then(res => console.log('[Paste] Sync Ergebnis:', JSON.stringify(res)))
                                .catch(err => console.error('[Paste] Sync fehlgeschlagen:', err));
                        }
                        
                        updateExplorerEditStatus();
                        showFloatingStatus('Eingefügt');
                    }
                } else if (explorerState.selectedCells.size > 1) {
                    // Bei mehreren Zellen: gleichen Wert in alle einfügen
                    let count = 0;
                    const cellsToSync = [];
                    explorerState.selectedCells.forEach(cellKey => {
                        const [rowIndex, colIndex] = cellKey.split('-').map(Number);
                        const td = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                        
                        if (td) {
                            const original = td.dataset.original;
                            const prevValue = explorerState.data[rowIndex][colIndex];
                            explorerState.data[rowIndex][colIndex] = text;
                            explorerState.editedCells.set(cellKey, text);
                            
                            const contentSpan = td.querySelector('.cell-content');
                            if (contentSpan) {
                                contentSpan.textContent = text;
                            } else {
                                td.textContent = text;
                            }
                            
                            if (text !== original) {
                                td.classList.add('edited');
                            }
                            cellsToSync.push({ row: rowIndex, col: colIndex, value: text, oldValue: prevValue });
                            count++;
                        }
                    });
                    
                    // Live Session Sync (wenn aktiv)
                    if (explorerState.liveSessionActive && explorerState.liveSessionReady && cellsToSync.length > 0) {
                        console.log(`[Paste] Live-Sync: ${cellsToSync.length} Zellen`);
                        window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(cellsToSync))
                            .then(res => console.log('[Paste] Batch-Sync Ergebnis:', JSON.stringify(res)))
                            .catch(err => console.error('[Paste] Batch-Sync fehlgeschlagen:', err));
                    }
                    
                    updateExplorerEditStatus();
                    showFloatingStatus(`In ${count} Zelle(n) eingefügt`);
                } else {
                    showFloatingStatus('Keine Zelle ausgewählt', true);
                }
            } catch (err) {
                console.error('Einfügen fehlgeschlagen:', err);
                showFloatingStatus('Einfügen fehlgeschlagen', true);
            }
        }
        
        // ============ SUCHEN & ERSETZEN (Find & Replace) ============
        
        // Find & Replace State
        const findReplaceState = {
            matches: [],
            currentMatchIndex: -1,
            lastSearchTerm: '',
            isOpen: false
        };
        
        // Find & Replace State komplett zurücksetzen
        function resetFindReplaceState() {
            findReplaceState.matches = [];
            findReplaceState.currentMatchIndex = -1;
            findReplaceState.lastSearchTerm = '';
            findReplaceState.isOpen = false;
            
            // Panel verstecken
            const panel = document.getElementById('findReplacePanel');
            if (panel) panel.style.display = 'none';
            
            // Input-Felder leeren
            const findText = document.getElementById('findText');
            if (findText) findText.value = '';
            const replaceText = document.getElementById('replaceText');
            if (replaceText) replaceText.value = '';
            
            // Checkboxen zurücksetzen
            const findCaseSensitive = document.getElementById('findCaseSensitive');
            if (findCaseSensitive) findCaseSensitive.checked = false;
            const findWholeWord = document.getElementById('findWholeWord');
            if (findWholeWord) findWholeWord.checked = false;
            const findRegex = document.getElementById('findRegex');
            if (findRegex) findRegex.checked = false;
            
            // Counter zurücksetzen
            const counter = document.getElementById('findMatchCounter');
            if (counter) counter.textContent = '0/0';
        }
        
        // Toggle Find & Replace Panel
        function toggleFindReplacePanel() {
            const panel = document.getElementById('findReplacePanel');
            findReplaceState.isOpen = !findReplaceState.isOpen;
            panel.style.display = findReplaceState.isOpen ? 'flex' : 'none';
            
            // Button-Farbe umschalten
            const btn = document.getElementById('btnToggleFindReplace');
            if (btn) {
                btn.classList.toggle('btn-primary', !findReplaceState.isOpen);
                btn.classList.toggle('btn-info', findReplaceState.isOpen);
            }
            
            if (findReplaceState.isOpen) {
                document.getElementById('findText').focus();
            } else {
                clearFindHighlights();
            }
        }
        
        // Find-Highlights entfernen
        function clearFindHighlights() {
            document.querySelectorAll('#explorerTableBody td.find-match').forEach(td => {
                td.classList.remove('find-match', 'find-current');
            });
            findReplaceState.matches = [];
            findReplaceState.currentMatchIndex = -1;
            updateFindMatchCounter();
        }
        
        // Match-Counter aktualisieren
        function updateFindMatchCounter() {
            const counter = document.getElementById('findMatchCounter');
            if (!counter) return;
            if (findReplaceState.matches.length === 0) {
                counter.textContent = '';
            } else {
                counter.textContent = `${findReplaceState.currentMatchIndex + 1}/${findReplaceState.matches.length}`;
            }
        }
        
        // Suche durchführen
        function performFind() {
            const searchTerm = document.getElementById('findText').value;
            if (!searchTerm) {
                clearFindHighlights();
                return;
            }
            
            const caseSensitive = document.getElementById('findCaseSensitive').checked;
            const wholeWord = document.getElementById('findWholeWord').checked;
            const useRegex = document.getElementById('findRegex').checked;
            
            // Alte Highlights entfernen
            clearFindHighlights();
            
            // Pattern erstellen
            let pattern;
            try {
                if (useRegex) {
                    pattern = new RegExp(searchTerm, caseSensitive ? 'g' : 'gi');
                } else {
                    const escaped = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                    const term = wholeWord ? `\\b${escaped}\\b` : escaped;
                    pattern = new RegExp(term, caseSensitive ? 'g' : 'gi');
                }
            } catch (e) {
                showFloatingStatus('Ungültiger regulärer Ausdruck', true);
                return;
            }
            
            // ALLE Daten durchsuchen (nicht nur sichtbare Zeilen!)
            const matches = [];
            
            // explorerState.data enthält die Rohdaten als 2D-Array
            const dataToSearch = explorerState.data || [];
            
            dataToSearch.forEach((row, rowIndex) => {
                if (!row || !Array.isArray(row)) return;
                row.forEach((cellValue, colIndex) => {
                    const cellText = String(cellValue ?? '');
                    if (pattern.test(cellText)) {
                        matches.push({ rowIndex, colIndex, td: null }); // td wird später gesetzt
                    }
                    // Reset pattern lastIndex für globales Matching
                    pattern.lastIndex = 0;
                });
            });
            
            findReplaceState.matches = matches;
            findReplaceState.lastSearchTerm = searchTerm;
            
            if (matches.length > 0) {
                findReplaceState.currentMatchIndex = 0;
                highlightCurrentMatch();
                showFloatingStatus(`${matches.length} Treffer gefunden`);
            } else {
                showFloatingStatus('Keine Treffer gefunden', true);
            }
            
            updateFindMatchCounter();
        }
        
        // Aktuellen Treffer hervorheben und scrollen
        function highlightCurrentMatch() {
            // Vorherigen aktuellen Treffer zurücksetzen
            document.querySelectorAll('#explorerTableBody td.find-current, #explorerTableBody td.find-match').forEach(td => {
                td.classList.remove('find-current', 'find-match');
            });
            
            if (findReplaceState.matches.length === 0 || findReplaceState.currentMatchIndex < 0) return;
            
            const match = findReplaceState.matches[findReplaceState.currentMatchIndex];
            
            // Virtual Scrolling: Zur Zeile scrollen und re-rendern
            // Finde filteredData-Index für die Zeile
            const filteredIdx = explorerState.filteredData.findIndex(item => item.originalIndex === match.rowIndex);
            if (filteredIdx >= 0) {
                // Prüfe ob Zeile bereits im sichtbaren Bereich liegt
                const needsScroll = filteredIdx < explorerState.virtualVisibleStart || 
                                    filteredIdx >= explorerState.virtualVisibleEnd;
                if (needsScroll) {
                    scrollToVirtualRow(filteredIdx);
                    // Force Re-Render damit die Zeile im DOM ist
                    explorerState.virtualVisibleStart = -1;
                    renderExplorerTable(true);
                }
            }
            
            // Jetzt nach dem Rendern die TD-Zelle finden und markieren
            setTimeout(() => {
                const td = document.querySelector(`#explorerTableBody td[data-row="${match.rowIndex}"][data-col="${match.colIndex}"]`);
                if (td) {
                    // Nur sichtbare Treffer markieren (nur im aktuellen Viewport)
                    // Statt alle 3000+ Matches per querySelector zu suchen,
                    // nur die im sichtbaren Bereich (virtualVisibleStart..End) markieren
                    const visStart = explorerState.virtualVisibleStart;
                    const visEnd = explorerState.virtualVisibleEnd;
                    
                    // Set für schnelle Lookup der sichtbaren originalIndex-Werte
                    const visibleOriginalIndices = new Set();
                    for (let i = visStart; i < visEnd && i < explorerState.filteredData.length; i++) {
                        visibleOriginalIndices.add(explorerState.filteredData[i].originalIndex);
                    }
                    
                    // Nur Matches im sichtbaren Bereich highlighten
                    for (let i = 0; i < findReplaceState.matches.length; i++) {
                        const m = findReplaceState.matches[i];
                        if (!visibleOriginalIndices.has(m.rowIndex)) continue;
                        const mtd = document.querySelector(`#explorerTableBody td[data-row="${m.rowIndex}"][data-col="${m.colIndex}"]`);
                        if (mtd) mtd.classList.add('find-match');
                    }
                    
                    td.classList.add('find-current');
                    td.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
                }
                updateFindMatchCounter();
            }, 50);
        }
        
        // Nächsten Treffer finden
        function findNext() {
            const searchTerm = document.getElementById('findText').value;
            
            // Wenn Suchbegriff geändert wurde, neu suchen
            if (searchTerm !== findReplaceState.lastSearchTerm) {
                performFind();
                return;
            }
            
            if (findReplaceState.matches.length === 0) {
                performFind();
                return;
            }
            
            // Zum nächsten Treffer
            findReplaceState.currentMatchIndex = 
                (findReplaceState.currentMatchIndex + 1) % findReplaceState.matches.length;
            highlightCurrentMatch();
        }
        
        // Vorherigen Treffer finden
        function findPrevious() {
            if (findReplaceState.matches.length === 0) return;
            
            findReplaceState.currentMatchIndex = 
                (findReplaceState.currentMatchIndex - 1 + findReplaceState.matches.length) % findReplaceState.matches.length;
            highlightCurrentMatch();
        }
        
        // Einzelne Ersetzung
        function replaceOne() {
            if (findReplaceState.matches.length === 0 || findReplaceState.currentMatchIndex < 0) {
                findNext();
                return;
            }
            
            const replaceText = document.getElementById('replaceText').value;
            const match = findReplaceState.matches[findReplaceState.currentMatchIndex];
            
            // Zelle ersetzen (synchron, Live-Sync im Hintergrund)
            replaceCellContent(match.rowIndex, match.colIndex, replaceText);
            
            // Treffer aus Liste entfernen
            findReplaceState.matches.splice(findReplaceState.currentMatchIndex, 1);
            
            // Index anpassen
            if (findReplaceState.matches.length === 0) {
                findReplaceState.currentMatchIndex = -1;
                updateFindMatchCounter();
                showFloatingStatus('Alle Treffer ersetzt');
            } else {
                if (findReplaceState.currentMatchIndex >= findReplaceState.matches.length) {
                    findReplaceState.currentMatchIndex = 0;
                }
                highlightCurrentMatch();
            }
        }
        
        // Alle ersetzen
        async function replaceAll() {
            const searchTerm = document.getElementById('findText').value;
            if (!searchTerm) return;
            
            // Immer frische Suche, damit nach Undo/Einzelersetzung alle Treffer korrekt sind
            performFind();
            
            if (findReplaceState.matches.length === 0) return;
            
            const replaceText = document.getElementById('replaceText').value;
            const caseSensitive = document.getElementById('findCaseSensitive').checked;
            const wholeWord = document.getElementById('findWholeWord').checked;
            const useRegex = document.getElementById('findRegex').checked;
            
            // Pattern erstellen
            let pattern;
            try {
                if (useRegex) {
                    pattern = new RegExp(searchTerm, caseSensitive ? 'g' : 'gi');
                } else {
                    const escaped = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                    const term = wholeWord ? `\\b${escaped}\\b` : escaped;
                    pattern = new RegExp(term, caseSensitive ? 'g' : 'gi');
                }
            } catch (e) {
                showFloatingStatus('Ungültiger regulärer Ausdruck', true);
                return;
            }
            
            const undoActions = [];
            const cellsToSync = [];  // Für Live Session Sync
            let replacedCount = 0;
            
            // Alle Treffer ersetzen (arbeitet über Daten, nicht DOM)
            findReplaceState.matches.forEach(match => {
                const oldValue = String(explorerState.data[match.rowIndex]?.[match.colIndex] ?? '');
                const newValue = oldValue.replace(pattern, replaceText);
                
                if (oldValue !== newValue) {
                    // Original-Wert merken (für Undo)
                    const originalValue = explorerState.data[match.rowIndex]?.[match.colIndex];
                    
                    undoActions.push({
                        rowIndex: match.rowIndex,
                        colIndex: match.colIndex,
                        oldValue,
                        newValue,
                        originalValue: originalValue
                    });
                    
                    // Für Live Session Sync merken
                    cellsToSync.push({
                        row: match.rowIndex,
                        col: match.colIndex,
                        value: newValue,
                        oldValue: oldValue
                    });
                    
                    // Daten aktualisieren (explorerState.data)
                    if (explorerState.data[match.rowIndex]) {
                        explorerState.data[match.rowIndex][match.colIndex] = newValue;
                    }
                    
                    // Auch filteredData aktualisieren (für Anzeige)
                    const filteredItem = explorerState.filteredData.find(item => item.originalIndex === match.rowIndex);
                    if (filteredItem && filteredItem.row) {
                        filteredItem.row[match.colIndex] = newValue;
                    }
                    
                    explorerState.editedCells.set(`${match.rowIndex}-${match.colIndex}`, newValue);
                    
                    replacedCount++;
                }
            });
            
            // Undo-Aktion speichern (mit Suchparametern für native Undo/Redo)
            if (undoActions.length > 0) {
                pushExplorerUndo({
                    type: 'multi',
                    actions: undoActions,
                    // Suchparameter für native Excel Find/Replace bei Undo
                    searchText: useRegex ? null : searchTerm,
                    replaceText: useRegex ? null : replaceText,
                    matchCase: caseSensitive,
                    wholeWord: wholeWord
                });
            }
            
            console.log('[ReplaceAll] Ersetzt:', replacedCount, 'Zellen');
            console.log('[ReplaceAll] undoActions:', undoActions.length);
            console.log('[ReplaceAll] cellsToSync:', cellsToSync.length);
            
            const liveActive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            console.log('[ReplaceAll] Live aktiv:', liveActive);
            
            // Live Session Sync (wenn aktiv)
            if (liveActive && replacedCount > 0) {
                try {
                    // Für einfache Suchen (kein Regex): Excel's native Replace nutzen - SCHNELL!
                    if (!useRegex) {
                        console.log(`[ReplaceAll] Nutze Excel native Replace`);
                        const result = await window.electronAPI.liveSessionFindReplace(
                            searchTerm, 
                            replaceText, 
                            caseSensitive, 
                            wholeWord
                        );
                        if (result.success) {
                            console.log('[ReplaceAll] Excel Replace erfolgreich');
                        } else {
                            console.error('[ReplaceAll] Excel Replace fehlgeschlagen:', result.error);
                        }
                    } else {
                        // Bei Regex: Fallback auf Zellenweise (langsam, aber korrekt)
                        console.log(`[ReplaceAll] Regex-Modus: Sync ${cellsToSync.length} Zellen einzeln`);
                        const result = await window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(cellsToSync));
                        if (result.success) {
                            console.log(`[ReplaceAll] Batch-Sync: ${result.count} Zellen`);
                        } else {
                            console.error('[ReplaceAll] Batch-Sync fehlgeschlagen:', result.error);
                        }
                    }
                } catch (error) {
                    console.error('[ReplaceAll] Live-Sync Fehler:', error);
                }
            }
            
            // Tabelle neu rendern und Status aktualisieren
            clearFindHighlights();
            renderExplorerTable();
            updateExplorerEditStatus();
            
            showFloatingStatus(`${replacedCount} Ersetzung(en) durchgeführt`);
        }
        
        // Zelleninhalt mit Regex ersetzen
        function replaceCellContent(rowIndex, colIndex, replaceText) {
            const searchTerm = document.getElementById('findText').value;
            const caseSensitive = document.getElementById('findCaseSensitive').checked;
            const wholeWord = document.getElementById('findWholeWord').checked;
            const useRegex = document.getElementById('findRegex').checked;
            
            let pattern;
            if (useRegex) {
                pattern = new RegExp(searchTerm, caseSensitive ? '' : 'i');
            } else {
                const escaped = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const term = wholeWord ? `\\b${escaped}\\b` : escaped;
                pattern = new RegExp(term, caseSensitive ? '' : 'i');
            }
            
            const cellKey = `${rowIndex}-${colIndex}`;
            
            // Immer explorerState.data verwenden (enthält die rohen Daten-Arrays)
            if (!explorerState.data[rowIndex]) return;
            
            const oldValue = String(explorerState.data[rowIndex][colIndex] ?? '');
            const newValue = oldValue.replace(pattern, replaceText);
            
            // Undo-Aktion speichern
            pushExplorerUndo({
                type: 'single',
                rowIndex,
                colIndex,
                oldValue,
                newValue,
                originalValue: oldValue
            });
            
            // Daten aktualisieren
            explorerState.data[rowIndex][colIndex] = newValue;
            explorerState.editedCells.set(cellKey, newValue);
            
            // Live Session Sync (wenn aktiv) — Fire-and-forget, nicht blockierend
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols([{
                    row: rowIndex,
                    col: colIndex,
                    value: newValue,
                    oldValue: oldValue
                }])).catch(error => {
                    console.error('[Ersetzen] Live-Sync Fehler:', error);
                });
            }
            
            // Falls Zelle sichtbar ist, UI aktualisieren
            const td = document.querySelector(`#explorerTableBody td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
            if (td) {
                td.textContent = newValue;
                td.dataset.lastValue = newValue;
                td.classList.add('edited');
                td.classList.remove('find-match', 'find-current');
            }
            
            updateExplorerEditStatus();
        }
        
        // Event Listeners für Find & Replace
        document.addEventListener('DOMContentLoaded', () => {
            // App-Version im Header anzeigen (synchron aus preload, async als Fallback)
            try {
                const vSync = window.electronAPI?.appVersion;
                if (vSync) window.__appVersion = vSync;
            } catch (_) { /* ignore */ }
            const applyVersion = (v) => {
                if (!v) return;
                window.__appVersion = v;
                const el = document.getElementById('appVersionBadge');
                if (el) el.textContent = 'v' + v;
                const footer = document.querySelector('footer');
                if (footer) {
                    // Footer immer aus Template neu schreiben, damit Version auch dann
                    // erscheint, wenn updateLanguage() bereits vorher lief.
                    try {
                        const tpl = (typeof t === 'function') ? t('copyright') : footer.innerHTML;
                        footer.innerHTML = tpl.replace('{version}', v);
                    } catch (_) {
                        footer.innerHTML = footer.innerHTML.replace('{version}', v);
                    }
                }
            };
            applyVersion(window.__appVersion);
            (async () => {
                try {
                    const v = await window.electronAPI?.getAppVersion?.();
                    applyVersion(v);
                } catch (_) { /* ignore */ }
            })();

            // Toggle Button
            document.getElementById('btnToggleFindReplace')?.addEventListener('click', toggleFindReplacePanel);
            
            // Close Button
            document.getElementById('btnCloseFindReplace')?.addEventListener('click', toggleFindReplacePanel);
            
            // Find Next Button
            document.getElementById('btnFindNext')?.addEventListener('click', findNext);
            
            // Replace One Button
            document.getElementById('btnReplaceOne')?.addEventListener('click', replaceOne);
            
            // Replace All Button
            document.getElementById('btnReplaceAll')?.addEventListener('click', replaceAll);
            
            // Undo Button
            document.getElementById('btnFindReplaceUndo')?.addEventListener('click', async () => {
                if (await undoExplorer()) {
                    showFloatingStatus('Rückgängig gemacht');
                    // Suchergebnisse aktualisieren, damit rückgängig gemachte Zellen wieder in matches sind
                    if (findReplaceState.lastSearchTerm) {
                        performFind();
                    }
                } else {
                    showFloatingStatus('Nichts zum Rückgängigmachen', true);
                }
            });
            
            // Enter-Taste im Suchfeld
            document.getElementById('findText')?.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        findPrevious();
                    } else {
                        findNext();
                    }
                }
            });
            
            // Escape zum Schließen
            document.getElementById('findReplacePanel')?.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    toggleFindReplacePanel();
                }
            });
        });
        
        // Keyboard Shortcut: Ctrl+H für Find & Replace
        document.addEventListener('keydown', (e) => {
            // F11 für Vollbild im Datenexplorer
            if (e.key === 'F11' && !elements.dataExplorerModal.classList.contains('hidden')) {
                e.preventDefault();
                toggleExplorerFullscreen();
            }
            // Escape beendet Vollbild
            if (e.key === 'Escape' && !elements.dataExplorerModal.classList.contains('hidden')) {
                const modal = document.querySelector('#dataExplorerModal .modal');
                if (modal.classList.contains('modal-fullscreen')) {
                    e.preventDefault();
                    toggleExplorerFullscreen();
                    return;
                }
            }
            if (e.ctrlKey && e.key === 'h' && !elements.dataExplorerModal.classList.contains('hidden')) {
                e.preventDefault();
                if (!findReplaceState.isOpen) {
                    toggleFindReplacePanel();
                }
            }
            // F3 für nächsten Treffer
            if (e.key === 'F3' && findReplaceState.isOpen) {
                e.preventDefault();
                if (e.shiftKey) {
                    findPrevious();
                } else {
                    findNext();
                }
            }
        });
        
        // ============ ENDE SUCHEN & ERSETZEN ============
        
        // Floating Status Nachricht anzeigen
        function showFloatingStatus(message, type = false) {
            // Alte Nachricht entfernen
            const oldStatus = document.querySelector('.floating-status');
            if (oldStatus) oldStatus.remove();
            
            // Typ-Mapping: boolean backward-compat + string-Typen
            const colors = {
                error: '#F44336',
                warning: '#FF9800',
                success: '#4CAF50',
                info: '#2196F3'
            };
            let bg, duration;
            if (type === true) {
                bg = colors.error; duration = 4000;
            } else if (typeof type === 'string' && colors[type]) {
                bg = colors[type];
                duration = type === 'success' ? 3500 : type === 'info' ? 0 : 4000;
            } else {
                bg = 'var(--primary)'; duration = 2000;
            }
            
            const status = document.createElement('div');
            status.className = 'floating-status';
            status.style.cssText = `
                position: fixed;
                bottom: 80px;
                left: 50%;
                transform: translateX(-50%);
                background: ${bg};
                color: white;
                padding: 10px 20px;
                border-radius: 6px;
                z-index: 10001;
                font-size: 13px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                animation: fadeInUp 0.3s ease;
            `;
            status.textContent = message;
            document.body.appendChild(status);
            
            // duration=0 → bleibt stehen (für info/progress), wird vom nächsten Aufruf entfernt
            if (duration > 0) {
                setTimeout(() => {
                    status.style.animation = 'fadeOut 0.3s ease';
                    setTimeout(() => status.remove(), 300);
                }, duration);
            }
        }

        async function openDataExplorer() {
            // State immer zurücksetzen beim Öffnen (außer bei Recovery)
            const recoveryData = loadExplorerRecoveryData();
            if (!recoveryData) {
                resetExplorerState();
            }
            
            elements.dataExplorerModal.classList.remove('hidden');
            document.body.classList.add('modal-open');
            
            // Live-Indikator sofort aktualisieren
            updateLiveModeIndicator();
            
            // Prüfe auf Recovery-Daten (von vorherigem Crash)
            if (recoveryData && !explorerState.filePath) {
                // Anzahl Änderungen ermitteln (v2.0 = Diffs pro Sheet, v1.x = flache Liste)
                let totalRecoveryChanges = 0;
                if (recoveryData.version === '2.0' && recoveryData.sheets) {
                    for (const s of Object.values(recoveryData.sheets)) {
                        totalRecoveryChanges += (s.editedCells?.length || 0);
                    }
                } else {
                    totalRecoveryChanges = (recoveryData.editedCells?.length || 0);
                }
                
                // Prüfe ob die Datei noch existiert
                let fileExists = false;
                try {
                    const checkResult = await window.electronAPI.checkFileExists(recoveryData.filePath);
                    fileExists = checkResult.exists;
                } catch (e) {
                    fileExists = false;
                }
                
                if (fileExists && totalRecoveryChanges > 0) {
                    const restore = await showConfirmDialog(
                        currentLanguage === 'en' ? 'Restore Data?' : 'Daten wiederherstellen?',
                        currentLanguage === 'en' 
                            ? `Recovery data found from a previous session:\n\nFile: ${recoveryData.fileName}\nChanges: ${totalRecoveryChanges}\nTime: ${new Date(recoveryData.timestamp).toLocaleString()}\n\nDo you want to restore these changes?`
                            : `Es wurden Wiederherstellungsdaten aus einer vorherigen Sitzung gefunden:\n\nDatei: ${recoveryData.fileName}\nÄnderungen: ${totalRecoveryChanges}\nZeitpunkt: ${new Date(recoveryData.timestamp).toLocaleString()}\n\nMöchten Sie diese Änderungen wiederherstellen?`,
                        currentLanguage === 'en' ? 'Restore' : 'Wiederherstellen',
                        currentLanguage === 'en' ? 'Discard' : 'Verwerfen'
                    );
                    
                    if (restore) {
                        await applyExplorerRecoveryData(recoveryData);
                        showNotification(
                            currentLanguage === 'en' ? 'Data restored successfully' : 'Daten erfolgreich wiederhergestellt', 
                            'success'
                        );
                    } else {
                        clearExplorerRecoveryData();
                    }
                } else {
                    // Datei existiert nicht mehr oder keine Änderungen - Recovery-Daten löschen
                    clearExplorerRecoveryData();
                }
            }
        }
        
        let _isClosingExplorer = false;  // Verhindert mehrfaches Schließen
        
        async function closeDataExplorer() {
            // Verhindere mehrfaches Aufrufen (z.B. durch schnelles Doppelklicken)
            if (_isClosingExplorer) return;
            _isClosingExplorer = true;
            
            try {
            // Prüfe auf ungespeicherte Änderungen ODER aktive Filter im Live-Modus
            const totalChanges = countAllChanges();
            const hasActiveFilters = explorerState.liveSessionActive && 
                explorerState.filters.some(f => f.column && f.value);
            
            if (totalChanges > 0 || hasActiveFilters) {
                let message;
                let confirmLabel;
                let dialogTitle;
                if (totalChanges > 0 && hasActiveFilters) {
                    message = currentLanguage === 'en'
                        ? `You have ${totalChanges} unsaved change(s) and active filters.\n\nDo you really want to close the Data Explorer without saving?`
                        : `Sie haben ${totalChanges} ungespeicherte Änderung(en) und aktive Filter.\n\nMöchten Sie den Datenexplorer wirklich ohne Speichern schließen?`;
                    confirmLabel = currentLanguage === 'en' ? 'Close without saving' : 'Ohne Speichern schließen';
                    dialogTitle = currentLanguage === 'en' ? 'Unsaved Changes' : 'Ungespeicherte Änderungen';
                } else if (hasActiveFilters) {
                    // Nur Filter aktiv, keine ungespeicherten Daten → sanftere Warnung
                    message = currentLanguage === 'en'
                        ? `Filters are still active in Excel. They will be cleared when closing.\n\nClose Data Explorer?`
                        : `In Excel sind noch Filter aktiv. Diese werden beim Schließen zurückgesetzt.\n\nDatenexplorer schließen?`;
                    confirmLabel = currentLanguage === 'en' ? 'Close' : 'Schließen';
                    dialogTitle = currentLanguage === 'en' ? 'Active Filters' : 'Aktive Filter';
                } else {
                    message = currentLanguage === 'en' 
                        ? `You have ${totalChanges} unsaved change(s).\n\nDo you really want to close the Data Explorer without saving?`
                        : `Sie haben ${totalChanges} ungespeicherte Änderung(en).\n\nMöchten Sie den Datenexplorer wirklich ohne Speichern schließen?`;
                    confirmLabel = currentLanguage === 'en' ? 'Close without saving' : 'Ohne Speichern schließen';
                    dialogTitle = currentLanguage === 'en' ? 'Unsaved Changes' : 'Ungespeicherte Änderungen';
                }
                
                const confirmed = await showConfirmDialog(
                    dialogTitle,
                    message,
                    confirmLabel,
                    currentLanguage === 'en' ? 'Cancel' : 'Abbrechen'
                );
                
                if (!confirmed) {
                    _isClosingExplorer = false;
                    return; // Abbruch - Datenexplorer bleibt offen
                }
            }
            
            // Live-Session beenden wenn aktiv
            if (explorerState.liveSessionActive) {
                // Filter vor dem Schließen zurücksetzen — mit Timeout damit Close nicht hängt
                try {
                    if (explorerState.filters.some(f => f.column && f.value)) {
                        const timeout = new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 5000));
                        await Promise.race([
                            window.electronAPI.liveSessionClearAutoFilter(),
                            timeout
                        ]);
                        console.log('[CloseExplorer] AutoFilter in Excel zurückgesetzt');
                    }
                } catch (e) {
                    console.warn('[CloseExplorer] AutoFilter-Reset fehlgeschlagen:', e);
                }
                
                await stopLiveSession();
                updateLiveSessionIndicator();
            }
            
            // State komplett zurücksetzen
            resetExplorerState();
            
            // Recovery-Daten löschen (normales Schließen)
            clearExplorerRecoveryData();
            
            elements.dataExplorerModal.classList.add('hidden');
            document.body.classList.remove('modal-open');
            } finally {
                _isClosingExplorer = false;
            }
        }
        
        // Explorer-Vorschau zeigen (zeigt genau das was exportiert wird)
        function showExplorerPreview() {
            if (!explorerState.filteredData || explorerState.filteredData.length === 0) {
                return;
            }
            
            const modal = document.getElementById('explorerPreviewModal');
            const tableContainer = document.getElementById('previewTableContainer');
            
            // Gleiche Daten wie beim Export verwenden
            const exportData = explorerState.filteredData.map(item => item.row);
            const visibleColumns = explorerState.visibleColumns;
            
            // Info aktualisieren
            document.getElementById('previewFileName').textContent = explorerState.fileName || 'Unbekannt';
            document.getElementById('previewSheetName').textContent = explorerState.selectedSheet || 'Sheet1';
            document.getElementById('previewRowCount').textContent = exportData.length;
            document.getElementById('previewColCount').textContent = visibleColumns.length;
            document.getElementById('previewEditCount').textContent = explorerState.editedCells.size;
            
            // Header erstellen (nur sichtbare Spalten)
            let headerHtml = '<tr><th style="width: 50px; text-align: center; position: sticky; left: 0; background: var(--bg-medium); z-index: 2;">#</th>';
            visibleColumns.forEach(colIdx => {
                const colLetter = String.fromCharCode(65 + colIdx);
                const header = explorerState.headers[colIdx] || '-';
                headerHtml += `<th style="min-width: 120px;"><small style="color: var(--text-muted);">${colLetter}</small><br>${escapeHtml(header)}</th>`;
            });
            headerHtml += '</tr>';
            
            // Zeilen erstellen (mit bearbeiteten Werten - genau wie beim Export)
            let rowsHtml = '';
            explorerState.filteredData.forEach((item, displayIdx) => {
                const originalIndex = item.originalIndex;
                const row = item.row;
                
                // Prüfe ob diese Zeile bearbeitete Zellen hat
                let hasEdits = false;
                for (const colIdx of visibleColumns) {
                    if (explorerState.editedCells.has(`${originalIndex}-${colIdx}`)) {
                        hasEdits = true;
                        break;
                    }
                }
                
                const rowStyle = hasEdits ? 'background: rgba(255, 193, 7, 0.05);' : '';
                rowsHtml += `<tr style="${rowStyle}">`;
                rowsHtml += `<td style="text-align: center; font-weight: bold; color: var(--text-muted); position: sticky; left: 0; background: var(--bg-dark); z-index: 1;">${displayIdx + 1}</td>`;
                
                visibleColumns.forEach(colIdx => {
                    const cellKey = `${originalIndex}-${colIdx}`;
                    const isEdited = explorerState.editedCells.has(cellKey);
                    
                    // Hole den Wert direkt aus der Zeile (enthält bereits Bearbeitungen)
                    const cellValue = row[colIdx] !== undefined ? String(row[colIdx]) : '';
                    
                    const cellStyle = isEdited 
                        ? 'background: rgba(255, 193, 7, 0.3); border: 2px solid #FFC107;' 
                        : '';
                    
                    rowsHtml += `<td style="${cellStyle}">${escapeHtml(cellValue)}</td>`;
                });
                
                rowsHtml += '</tr>';
            });
            
            tableContainer.innerHTML = `
                <table class="results-table" style="width: max-content; min-width: 100%;">
                    <thead style="position: sticky; top: 0; z-index: 3;">${headerHtml}</thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            `;
            
            // Modal anzeigen
            modal.classList.remove('hidden');
        }
        
        function closeExplorerPreview() {
            document.getElementById('explorerPreviewModal').classList.add('hidden');
        }
        
        // Vollbild-Modus für Datenexplorer umschalten
        function toggleExplorerFullscreen() {
            const modal = document.querySelector('#dataExplorerModal .modal');
            const btn = document.getElementById('btnExplorerFullscreen');
            
            if (modal.classList.contains('modal-fullscreen')) {
                modal.classList.remove('modal-fullscreen');
                btn.innerHTML = '⛶';
                btn.title = 'Vollbild (F11)';
            } else {
                modal.classList.add('modal-fullscreen');
                btn.innerHTML = '⛶';
                btn.title = 'Vollbild beenden (F11 oder Esc)';
            }
        }
        
        // Drop-Zone anzeigen/verstecken
        function showExplorerDropZone(show) {
            const dropZone = elements.explorerDropZone;
            if (dropZone) {
                dropZone.style.display = show ? 'flex' : 'none';
            }
        }
        
        // Setup Drag & Drop Zone for Explorer
        function setupExplorerDropZone() {
            const dropZone = elements.explorerDropZone;
            if (!dropZone) return;
            
            // Klick öffnet Datei-Dialog
            dropZone.onclick = loadExplorerFile;
            
            // Drag-Events für die Drop-Zone
            dropZone.ondragover = (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.style.background = 'rgba(0, 122, 204, 0.15)';
                dropZone.style.borderColor = 'var(--primary)';
            };
            
            dropZone.ondragleave = (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.style.background = 'transparent';
                dropZone.style.borderColor = 'transparent';
            };
            
            dropZone.ondrop = async (e) => {
                e.preventDefault();
                e.stopPropagation();
                
                // Styling zurücksetzen
                dropZone.style.background = 'transparent';
                dropZone.style.borderColor = 'transparent';
                
                const files = e.dataTransfer.files;
                if (files.length === 0) return;
                
                const file = files[0];
                const fileName = file.name.toLowerCase();
                
                // Nur Excel-Dateien erlauben
                if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
                    showFloatingStatus('❌ Nur Excel-Dateien (.xlsx, .xls) werden unterstützt', 'error');
                    return;
                }
                
                // Dateipfad über Electron API abrufen (contextIsolation-sicher)
                const filePath = window.electronAPI.getPathForFile(file);
                if (!filePath) {
                    showFloatingStatus('❌ Dateipfad konnte nicht ermittelt werden', 'error');
                    return;
                }
                
                // Datei laden
                await loadExplorerFileByPath(filePath);
            };
        }
        
        // Datei über Pfad laden (für Drag & Drop)
        async function loadExplorerFileByPath(filePath) {
            if (!filePath) return;
            const _loadStart = Date.now();
            const _log = (msg) => console.log(`[LOAD ${Date.now() - _loadStart}ms] ${msg}`);
            _log('=== START loadExplorerFileByPath ===');
            
            // Prüfe auf ungespeicherte Änderungen
            if (hasUnsavedChanges()) {
                const totalChanges = countAllChanges();
                const confirmed = await showConfirmDialog(
                    'Ungespeicherte Änderungen',
                    `Sie haben ${totalChanges} ungespeicherte Änderung(en).\n\n` +
                    `Möchten Sie trotzdem eine neue Datei öffnen?\n\n` +
                    `⚠️ Alle Änderungen gehen verloren!`,
                    'Neue Datei öffnen',
                    'Abbrechen'
                );
                if (!confirmed) return;
            }
            
            // State komplett zurücksetzen bevor neue Datei geladen wird
            resetExplorerState();
            
            // Warnung wenn die gleiche Datei bereits in der Haupt-GUI geladen ist
            const normalizedPath = filePath.replace(/\\/g, '/').toLowerCase();
            const file1Path = (state.file1?.filePath || '').replace(/\\/g, '/').toLowerCase();
            const file2Path = (state.file2?.filePath || '').replace(/\\/g, '/').toLowerCase();
            if (normalizedPath === file1Path || normalizedPath === file2Path) {
                const which = normalizedPath === file1Path ? 'Quelldatei (Datei 1)' : 'Zieldatei (Datei 2)';
                showNotification(`Diese Datei ist auch als ${which} geladen — im Live-Modus kann es zu Schreibkonflikten kommen`, 'warning');
                _log(`WARNUNG: Datei ist auch als ${which} in der Haupt-GUI geladen`);
            }
            
            elements.explorerStatus.textContent = '⏳ Sheet-Namen laden...';
            _log('readExcelFile aufrufen (Sheet-Namen aus ZIP)...');
            
            // Versuche Datei zu öffnen
            let result = await window.electronAPI.readExcelFile(filePath);
            _log(`readExcelFile fertig: success=${result.success}, sheets=${result.sheets?.length}`);
            
            // Passwortgeschützte Datei?
            if (!result.success && result.needsPassword) {
                const password = await showPromptDialog(
                    '🔐 Passwort erforderlich',
                    'Diese Excel-Datei ist passwortgeschützt.\nBitte geben Sie das Passwort ein:',
                    '',
                    'password'
                );
                
                if (password === null) return;
                
                result = await window.electronAPI.readExcelFile(filePath, password);
                
                if (!result.success) {
                    if (result.needsPassword) {
                        showFloatingStatus('❌ Falsches Passwort', 'error');
                    } else {
                        elements.explorerStatus.textContent = `Fehler: ${result.error}`;
                    }
                    return;
                }
                
                explorerState.filePassword = password;
                showFloatingStatus('🔓 Datei entsperrt');
            } else if (!result.success) {
                elements.explorerStatus.textContent = `Fehler: ${result.error}`;
                return;
            } else {
                explorerState.filePassword = null;
            }
            
            // Cache leeren bei neuer Datei
            explorerState.sheetDataCache.clear();
            explorerState.editedCells.clear();
            explorerState.rowHighlights.clear();
            
            explorerState.filePath = filePath;
            explorerState.originalFilePath = filePath;  // Ursprüngliche Datei merken (für Export)
            explorerState.fileName = result.fileName;
            explorerState.sheets = result.sheets;
            explorerState.hiddenSheets = new Set(result.hiddenSheets || []);
            explorerState.hasPivotTables = result.hasPivotTables || false;
            console.log('[LOAD] hasPivotTables =', explorerState.hasPivotTables, '(from result:', result.hasPivotTables, ')');
            
            // Pivot-Warnung wird erst NACH der Engine-Erkennung ausgegeben (siehe unten),
            // damit sie im xlwings-/Live-Mode unterdrückt werden kann.
            
            // UI aktualisieren
            document.getElementById('explorerFileName').textContent = explorerState.fileName;
            
            // FileInfo-Button anzeigen
            const btnFileInfo = document.getElementById('btnFileInfo');
            if (btnFileInfo) btnFileInfo.style.display = '';
            
            // Sheet-Dropdown füllen (mit Markierung für ausgeblendete Sheets)
            elements.explorerSheetSelect.innerHTML = explorerState.sheets
                .map(s => {
                    const isHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(s);
                    const label = isHidden ? `👁️‍🗨️ ${s} (ausgeblendet)` : s;
                    return `<option value="${s}">${label}</option>`;
                })
                .join('');
            
            // selectedSheet VOR dem Laden setzen (für Live-Session-Start)
            if (explorerState.sheets.length > 0) {
                explorerState.selectedSheet = explorerState.sheets[0];
            }
            
            // Sheet-Daten ZUERST laden (Datei wird danach komplett freigegeben)
            // Streaming Reader: blockiert den Event-Loop NICHT, UI bleibt responsiv
            if (explorerState.sheets.length > 0) {
                _log(`loadExplorerSheet starten: "${explorerState.sheets[0]}"`);
                elements.explorerStatus.textContent = '⏳ Sheet-Daten laden (Streaming)...';
                await loadExplorerSheet(explorerState.sheets[0]);
                _log('loadExplorerSheet fertig');
            }
            
            // DANACH Live-Session starten (Datei ist jetzt frei für xlwings/Excel)
            // Auf Windows hält jeder Dateizugriff einen exklusiven Lock —
            // deshalb MUSS das sequenziell sein (v1.2.0-Reihenfolge)
            try {
                const engineSetting = localStorage.getItem('excelSyncEngine') || 'auto';
                _log(`checkExcelAvailable starten (Engine: ${engineSetting})...`);
                elements.explorerStatus.textContent = '⏳ Excel-Verfügbarkeit prüfen...';
                const status = await window.electronAPI.checkExcelAvailable();
                _log(`checkExcelAvailable: excelAvailable=${status?.excelAvailable}`);
                
                const shouldUseLive = (engineSetting === 'xlwings') || 
                                      (engineSetting === 'auto' && status && status.excelAvailable);
                
                if (shouldUseLive && status && status.excelAvailable) {
                    _log('startLiveSession starten...');
                    elements.explorerStatus.textContent = '⏳ Live-Session starten (Excel öffnet Datei)...';
                    const liveOk = await startLiveSession();
                    _log(`startLiveSession Ergebnis: ${liveOk}`);
                    if (liveOk) {
                        explorerState.engineMode = 'live';
                        _log(`✓ Live-Session AKTIV`);
                    } else {
                        explorerState.engineMode = 'openpyxl';
                        _log('✗ Live-Session FEHLGESCHLAGEN → openpyxl Fallback');
                        updateLiveModeIndicator();
                    }
                } else {
                    explorerState.engineMode = 'openpyxl';
                    _log(`openpyxl Modus (Einstellung: ${engineSetting}, Excel: ${status?.excelAvailable})`);
                    updateLiveModeIndicator();
                }
            } catch (e) {
                explorerState.engineMode = 'openpyxl';
                _log(`FEHLER bei Live-Session: ${e.message}`);
                console.error('[Engine] Fallback: openpyxl (Fehler)', e);
                updateLiveModeIndicator();
            }
            
            _log(`=== FERTIG === Gesamt: ${Date.now() - _loadStart}ms, Engine: ${explorerState.engineMode}`);
            elements.explorerStatus.textContent = '';
            showFloatingStatus(`📂 ${result.fileName} geladen (${explorerState.engineMode})`);

            // Hintergrund-Preload der restlichen Sheets (nur openpyxl-Modus)
            if (explorerState.engineMode !== 'live') {
                // fire-and-forget; läuft mit requestIdleCallback
                startBackgroundSheetPreload().catch(err => console.warn('[Preload] Exception:', err));
            }
            
            // Pivot-Warnung NUR im openpyxl-Fallback (im Live-/xlwings-Mode bleiben
            // Pivots beim Speichern erhalten — keine Warnung nötig)
            if (explorerState.hasPivotTables && explorerState.engineMode !== 'live') {
                const isEn = currentLanguage === 'en';
                showConfirmDialog(
                    '⚠️ ' + (isEn ? 'Pivot Tables detected' : 'Pivot-Tabellen erkannt'),
                    isEn
                        ? 'This file contains pivot tables!\n\nWithout Live Mode, pivot tables may be lost or corrupted when saving.\n\nRecommendation: Use Live Mode or create a backup copy.'
                        : 'Diese Datei enthält Pivot-Tabellen!\n\nOhne Live-Modus können Pivot-Tabellen beim Speichern verloren gehen oder beschädigt werden.\n\nEmpfehlung: Verwenden Sie den Live Modus oder erstellen Sie eine Sicherheitskopie.',
                    'OK',
                    null
                );
            }
        }
        
        async function loadExplorerFile() {
            const _loadStart2 = Date.now();
            const _log2 = (msg) => console.log(`[LOAD ${Date.now() - _loadStart2}ms] ${msg}`);
            _log2('=== START loadExplorerFile ===');
            // Prüfe auf ungespeicherte Änderungen
            if (hasUnsavedChanges()) {
                const totalChanges = countAllChanges();
                const confirmed = await showConfirmDialog(
                    'Ungespeicherte Änderungen',
                    `Sie haben ${totalChanges} ungespeicherte Änderung(en).\n\n` +
                    `Möchten Sie trotzdem eine neue Datei öffnen?\n\n` +
                    `⚠️ Alle Änderungen gehen verloren!`,
                    'Neue Datei öffnen',
                    'Abbrechen'
                );
                if (!confirmed) return;
            }
            
            // State komplett zurücksetzen bevor neue Datei geladen wird
            resetExplorerState();
            
            const filePath = await window.electronAPI.openFileDialog({
                title: 'Excel-Datei öffnen',
                filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
                defaultPath: getExplorerDefaultPath()
            });
            if (!filePath) return;
            
            // Warnung wenn die gleiche Datei bereits in der Haupt-GUI geladen ist
            const normalizedPath2 = filePath.replace(/\\/g, '/').toLowerCase();
            const file1Path2 = (state.file1?.filePath || '').replace(/\\/g, '/').toLowerCase();
            const file2Path2 = (state.file2?.filePath || '').replace(/\\/g, '/').toLowerCase();
            if (normalizedPath2 === file1Path2 || normalizedPath2 === file2Path2) {
                const which2 = normalizedPath2 === file1Path2 ? 'Quelldatei (Datei 1)' : 'Zieldatei (Datei 2)';
                showNotification(`Diese Datei ist auch als ${which2} geladen — im Live-Modus kann es zu Schreibkonflikten kommen`, 'warning');
                _log2(`WARNUNG: Datei ist auch als ${which2} in der Haupt-GUI geladen`);
            }
            
            // Versuche Datei zu öffnen (mit Passwort-Retry bei geschützten Dateien)
            let result = await window.electronAPI.readExcelFile(filePath);
            
            // Passwortgeschützte Datei?
            if (!result.success && result.needsPassword) {
                const password = await showPromptDialog(
                    '🔐 Passwort erforderlich',
                    'Diese Excel-Datei ist passwortgeschützt.\nBitte geben Sie das Passwort ein:',
                    '',
                    'password'
                );
                
                if (password === null) {
                    // Abgebrochen
                    return;
                }
                
                // Erneut mit Passwort versuchen
                result = await window.electronAPI.readExcelFile(filePath, password);
                
                if (!result.success) {
                    if (result.needsPassword) {
                        showFloatingStatus('❌ Falsches Passwort', 'error');
                    } else {
                        elements.explorerStatus.textContent = `Fehler: ${result.error}`;
                    }
                    return;
                }
                
                // Passwort für späteres Speichern merken
                explorerState.filePassword = password;
                showFloatingStatus('🔓 Datei entsperrt');
            } else if (!result.success) {
                elements.explorerStatus.textContent = `Fehler: ${result.error}`;
                return;
            } else {
                // Kein Passwort nötig
                explorerState.filePassword = null;
            }
            
            // Cache leeren bei neuer Datei
            explorerState.sheetDataCache.clear();
            explorerState.editedCells.clear();
            explorerState.rowHighlights.clear();
            
            explorerState.filePath = filePath;
            explorerState.originalFilePath = filePath;
            explorerState.fileName = result.fileName;
            explorerState.sheets = result.sheets;
            explorerState.hiddenSheets = new Set(result.hiddenSheets || []);
            explorerState.hasPivotTables = result.hasPivotTables || false;
            console.log('[LOAD] hasPivotTables =', explorerState.hasPivotTables, '(from result:', result.hasPivotTables, ')');
            
            // Pivot-Warnung wird erst NACH der Engine-Erkennung ausgegeben (siehe unten),
            // damit sie im xlwings-/Live-Mode unterdrückt werden kann.
            
            // Element frisch abfragen (kann durch setLanguage ersetzt worden sein)
            const explorerFileNameEl = document.getElementById('explorerFileName');
            if (explorerFileNameEl) explorerFileNameEl.textContent = result.fileName;
            elements.explorerSheetSelect.innerHTML = explorerState.sheets
                .map(s => {
                    const isHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(s);
                    const label = isHidden ? `👁️‍🗨️ ${s} (ausgeblendet)` : s;
                    return `<option value="${s}">${label}</option>`;
                })
                .join('');
            
            // Passwort-Indikator aktualisieren
            updatePasswordIndicator();
            
            // FileInfo-Button anzeigen
            const btnFileInfo = document.getElementById('btnFileInfo');
            if (btnFileInfo) btnFileInfo.style.display = '';
            
            
            // selectedSheet VOR dem Laden setzen
            explorerState.selectedSheet = result.sheets[0];
            
            // Sheet-Daten ZUERST laden (Datei wird danach komplett freigegeben)
            _log2(`loadExplorerSheet starten: "${result.sheets[0]}"`);
            elements.explorerStatus.textContent = '⏳ Sheet-Daten laden (Streaming)...';
            await loadExplorerSheet(result.sheets[0]);
            _log2('loadExplorerSheet fertig');
            
            // DANACH Live-Session starten (Datei ist jetzt frei für xlwings/Excel)
            try {
                const engineSetting = localStorage.getItem('excelSyncEngine') || 'auto';
                _log2(`checkExcelAvailable starten (Engine: ${engineSetting})...`);
                elements.explorerStatus.textContent = '⏳ Excel-Verfügbarkeit prüfen...';
                const status = await window.electronAPI.checkExcelAvailable();
                _log2(`checkExcelAvailable: excelAvailable=${status?.excelAvailable}`);
                
                const shouldUseLive = (engineSetting === 'xlwings') || 
                                      (engineSetting === 'auto' && status && status.excelAvailable);
                
                if (shouldUseLive && status && status.excelAvailable) {
                    _log2('startLiveSession starten...');
                    elements.explorerStatus.textContent = '⏳ Live-Session starten (Excel öffnet Datei)...';
                    const liveOk = await startLiveSession();
                    _log2(`startLiveSession Ergebnis: ${liveOk}`);
                    if (liveOk) {
                        explorerState.engineMode = 'live';
                        _log2(`✓ Live-Session AKTIV`);
                    } else {
                        explorerState.engineMode = 'openpyxl';
                        _log2('✗ Live-Session FEHLGESCHLAGEN → openpyxl Fallback');
                        updateLiveModeIndicator();
                    }
                } else {
                    explorerState.engineMode = 'openpyxl';
                    _log2(`openpyxl Modus (Einstellung: ${engineSetting}, Excel: ${status?.excelAvailable})`);
                    updateLiveModeIndicator();
                }
            } catch (e) {
                explorerState.engineMode = 'openpyxl';
                _log2(`FEHLER bei Live-Session: ${e.message}`);
                console.error('[Engine] Fallback: openpyxl (Fehler)', e);
                updateLiveModeIndicator();
            }

            _log2(`=== FERTIG === Gesamt: ${Date.now() - _loadStart2}ms, Engine: ${explorerState.engineMode}`);
            elements.explorerStatus.textContent = '';
            showFloatingStatus(`📂 ${result.fileName} geladen (${explorerState.engineMode})`);
            
            // Pivot-Warnung NUR im openpyxl-Fallback (im Live-/xlwings-Mode bleiben
            // Pivots beim Speichern erhalten — keine Warnung nötig)
            if (explorerState.hasPivotTables && explorerState.engineMode !== 'live') {
                const isEn = currentLanguage === 'en';
                showConfirmDialog(
                    '⚠️ ' + (isEn ? 'Pivot Tables detected' : 'Pivot-Tabellen erkannt'),
                    isEn
                        ? 'This file contains pivot tables!\n\nWithout Live Mode, pivot tables may be lost or corrupted when saving.\n\nRecommendation: Use Live Mode or create a backup copy.'
                        : 'Diese Datei enthält Pivot-Tabellen!\n\nOhne Live-Modus können Pivot-Tabellen beim Speichern verloren gehen oder beschädigt werden.\n\nEmpfehlung: Verwenden Sie den Live Modus oder erstellen Sie eine Sicherheitskopie.',
                    'OK',
                    null
                );
            }
            
            // Auto-Save für Crash-Recovery starten
            startExplorerAutoSave();
        }
        
        async function loadExplorerSheet(sheetName) {
            if (!explorerState.filePath || !sheetName) return;
            
            // Guard: Verhindert parallele Sheet-Wechsel — Excel muss erst synchronisiert werden
            // Bei schnellem Klicken: Letztes gewünschtes Sheet merken → wird nach Abschluss automatisch geladen
            if (explorerState.isLoadingSheet) {
                console.log(`[loadExplorerSheet] Pending: Sheet "${sheetName}" vorgemerkt (vorheriger Ladevorgang läuft)`);
                explorerState.pendingSheetSwitch = sheetName;
                elements.explorerSheetSelect.value = sheetName; // Dropdown zeigt bereits das Ziel-Sheet
                return;
            }
            
            explorerState.isLoadingSheet = true;
            elements.explorerSheetSelect.disabled = true;
            
            // === UI KOMPLETT SPERREN während Sheet-Wechsel ===
            // Overlay über die Tabelle legen
            const switchOverlay = document.getElementById('explorerSheetSwitchOverlay');
            if (switchOverlay) switchOverlay.classList.remove('hidden');
            // Alle Buttons in der Explorer-Toolbar deaktivieren
            const _disabledExplorerButtons = [];
            document.querySelectorAll('#dataExplorerModal .btn').forEach(btn => {
                if (!btn.disabled) {
                    btn.disabled = true;
                    _disabledExplorerButtons.push(btn);
                }
            });
            // Suche deaktivieren
            if (elements.explorerSearch) elements.explorerSearch.disabled = true;
            
            // === ALLE ausstehenden Sync-Operationen ABBRECHEN (nicht flushen!) ===
            // Cell-Batch-Timer stoppen und Queue leeren (NICHT an Excel senden!)
            if (_cellSyncBatchTimer) {
                clearTimeout(_cellSyncBatchTimer);
                _cellSyncBatchTimer = null;
            }
            _pendingCellSyncs.clear();
            // Per-Cell Debounce-Timer abbrechen
            for (const [key, timer] of _syncCellTimers) {
                clearTimeout(timer);
            }
            _syncCellTimers.clear();
            // Alle laufenden Row-Sync-Timer abbrechen
            for (const [rowIdx, timer] of _syncRowTimers) {
                clearTimeout(timer);
            }
            _syncRowTimers.clear();
            
            try {
            // Prüfe ob das Sheet ausgeblendet ist
            if (explorerState.hiddenSheets && explorerState.hiddenSheets.has(sheetName)) {
                // Nur im Live-Modus kann man einblenden
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    const confirmed = await showConfirmDialog(
                        'Ausgeblendetes Arbeitsblatt',
                        `Das Arbeitsblatt "${sheetName}" ist in Excel ausgeblendet.\n\n` +
                        `Möchten Sie es einblenden und dorthin wechseln?`,
                        'Einblenden & Wechseln',
                        'Abbrechen'
                    );
                    if (!confirmed) {
                        // Dropdown zurücksetzen auf vorheriges Sheet
                        if (explorerState.selectedSheet) {
                            elements.explorerSheetSelect.value = explorerState.selectedSheet;
                        }
                        return;
                    }
                    // Sheet wird automatisch eingeblendet beim switchSheet (Python-Seite)
                }
                // Ohne Live-Session: Daten können trotzdem gelesen werden (ExcelJS liest auch hidden sheets)
            }
            
            // VOR dem Wechsel: Aktuelles Sheet im Cache speichern (inkl. formatierte Daten)
            // WICHTIG: Immer cachen, nicht nur bei Änderungen! Sonst gehen ssf-formatierte
            // Werte verloren wenn Live-Session beim Zurückwechseln Rohwerte liefert.
            if (explorerState.selectedSheet && explorerState.data.length > 0) {
                saveCurrentSheetToCache();
            }
            
            // Live-Session: Arbeitsblatt in Excel wechseln
            // PARALLEL zum Datenlesen starten (switchSheet + readExcelSheet gleichzeitig)
            // switchSheet betrifft nur die Excel-COM-Verbindung, readExcelSheet liest von Disk/Buffer
            console.log('[loadExplorerSheet] Live-Check:', {
                liveSessionActive: explorerState.liveSessionActive,
                liveSessionReady: explorerState.liveSessionReady,
                engineMode: explorerState.engineMode,
                targetSheet: sheetName
            });
            
            // switchSheet-Promise vorbereiten
            // Im Live-Modus (liveSessionActive+Ready) wird switchSheet MIT Daten
            // direkt im Datenlade-Block gemacht → hier nur für Cache-Fall und Session-Neustart
            // WICHTIG: Lazy Factory statt IIFE — verhindert doppelte switchSheet COM-Calls
            let switchSheetPromise = null;
            let switchSheetFactory = null;
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Für den Cache-Fall: switchSheet ohne Daten (nur Excel-Position synchronisieren)
                switchSheetFactory = () => (async () => {
                    try {
                        const switchResult = await window.electronAPI.liveSessionSwitchSheet(sheetName, false);
                        if (switchResult.success) {
                            console.log('[LiveSession] Sheet gewechselt zu:', sheetName, '(Cache-Sync)');
                            if (switchResult.hasConditionalFormatting !== undefined) {
                                explorerState.hasConditionalFormatting = !!switchResult.hasConditionalFormatting;
                            }
                            if (switchResult.wasHidden && explorerState.hiddenSheets) {
                                explorerState.hiddenSheets.delete(sheetName);
                                updateSheetDropdown();
                            }
                            return { ok: true };
                        }
                        console.error('[LiveSession] Cache-Sync Sheet-Wechsel fehlgeschlagen:', switchResult.error);
                    } catch (err) {
                        console.error('[LiveSession] Cache-Sync Sheet-Wechsel Fehler:', err);
                    }
                    return { ok: false };
                })();
            } else if (explorerState.engineMode === 'live') {
                // Live-Modus aktiv aber Session nicht bereit
                // KEIN Session-Neustart hier — loadExplorerFileByPath startet die Session
                // nach dem ersten loadExplorerSheet. Doppeltes Starten vermeiden!
                console.log('[loadExplorerSheet] engineMode=live aber Session nicht aktiv → warte auf loadExplorerFileByPath');
            }
            
            // Loading-Status anzeigen
            elements.explorerStatus.textContent = t('loadingData');
            
            // Prüfe ob dieses Sheet bereits im Cache ist
            const cachedSheet = explorerState.sheetDataCache.get(sheetName);
            
            if (cachedSheet) {
                // Aus Cache laden
                explorerState.selectedSheet = sheetName;
                elements.explorerSheetSelect.value = sheetName;
                explorerState.headers = cachedSheet.headers;
                explorerState.data = cachedSheet.data.map(row => [...row]); // Deep copy
                explorerState.originalData = cachedSheet.originalData.map(row => [...row]);
                explorerState.editedCells = new Map(cachedSheet.editedCells);
                explorerState.rowHighlights = new Map(cachedSheet.rowHighlights);
                explorerState.visibleColumns = [...cachedSheet.visibleColumns];
                explorerState.columnOrder = [...(cachedSheet.columnOrder || [])];
                explorerState.dataValidations = { ...(cachedSheet.dataValidations || {}) };
                explorerState.cellStyles = { ...(cachedSheet.cellStyles || {}) };
                explorerState.cellFormulas = { ...(cachedSheet.cellFormulas || {}) };
                explorerState.cellHyperlinks = { ...(cachedSheet.cellHyperlinks || {}) };
                explorerState.richTextCells = { ...(cachedSheet.richTextCells || {}) };
                explorerState.hiddenRows = new Set(cachedSheet.hiddenRows || []);
                explorerState.autoFilterRange = cachedSheet.autoFilterRange || null;
                explorerState.mergedCells = [...(cachedSheet.mergedCells || [])];
                explorerState.rowMapping = cachedSheet.rowMapping ? [...cachedSheet.rowMapping] : null;
                explorerState.columnOperationsQueue = [...(cachedSheet.columnOperationsQueue || [])];
                
                // UI-State zurücksetzen
                explorerState.filteredData = explorerState.data.map((row, index) => ({ originalIndex: index, row: row }));
                explorerState.currentPage = 1;
                explorerState.searchTerm = '';
                explorerState.sortColumn = null;
                explorerState.sortDirection = null;
                elements.explorerSearch.value = '';
                
                const editCount = explorerState.editedCells.size;
                if (editCount > 0) {
                    elements.explorerStatus.textContent = `${t('loadedFromCache')} (${editCount} ${t('changes')})`;
                } else {
                    elements.explorerStatus.textContent = '';
                }
                
                // WICHTIG: filterExplorerData() statt renderExplorerTable() verwenden,
                // damit versteckte Zeilen (hiddenRows) korrekt ausgefiltert werden!
                filterExplorerData();
                updateColumnToggles();
                updateExplorerEditStatus();
                updateHiddenRowsIndicator();
                updateHiddenColumnsIndicator();
                updateAutoFilterIndicator();
                // SCHRITT 1: switchSheet ohne Daten (nur Python-Referenz setzen)
                // SCHRITT 2: activateSheet (visueller Wechsel in Excel NACH dem GUI-Rendern)
                // AWAIT — Buttons bleiben gesperrt bis COM-Operationen fertig!
                if (switchSheetFactory) {
                    try {
                        const res = await switchSheetFactory();
                        if (!res.ok) {
                            showFloatingStatus('⚠️ Excel Sheet-Wechsel fehlgeschlagen', 'error');
                        } else {
                            try {
                                await window.electronAPI.liveSessionActivateSheet(sheetName);
                            } catch (e) {
                                console.error('[LiveSession] activateSheet Fehler:', e);
                            }
                        }
                    } catch (err) {
                        console.error('[LiveSession] Cache-Sync Switch Fehler:', err);
                    }
                } else if (switchSheetPromise) {
                    try {
                        const res = await switchSheetPromise;
                        if (!res.ok) {
                            showFloatingStatus('⚠️ Excel Sheet-Wechsel fehlgeschlagen', 'error');
                        } else {
                            try {
                                await window.electronAPI.liveSessionActivateSheet(sheetName);
                            } catch (e) {
                                console.error('[LiveSession] activateSheet Fehler:', e);
                            }
                        }
                    } catch (err) {
                        console.error('[LiveSession] Switch Fehler:', err);
                    }
                }
                return;
            }
            
            // Prüfe ob das Sheet umbenannt wurde (Live & Offline)
            // Die Datei auf Disk hat noch den alten Namen → wir müssen den Original-Namen verwenden
            let diskSheetName = sheetName;
            if (explorerState.sheetDiskNameMap.has(sheetName)) {
                diskSheetName = explorerState.sheetDiskNameMap.get(sheetName);
                console.log(`[loadExplorerSheet] Sheet "${sheetName}" wurde umbenannt → lese als "${diskSheetName}" von Disk`);
            }
            
            // ================================================================
            // DATEN-STRATEGIE:
            // Live-Session aktiv → Excel ist die einzige Quelle (xlwings):
            //   switchSheet + getData in EINEM Roundtrip (kein ExcelJS für Zelldaten)
            //   ZIP-Metadaten (hiddenRows, mergedCells, etc.) parallel aus gecachtem Buffer
            // Kein Live-Session → Fallback auf ExcelJS (wie bisher)
            // ================================================================
            
            let result;
            const _sheetStart = Date.now();
            
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // === LIVE-MODUS: Excel ist die einzige Datenquelle ===
                console.log(`[loadExplorerSheet] LIVE-MODUS: switchSheet+getData für "${sheetName}"...`);
                
                // switchSheet MIT Daten + ZIP-Metadaten parallel holen
                const [switchResult, metadataResult] = await Promise.all([
                    // switchSheet + getData in einem Roundtrip (3 Versuche)
                    (async () => {
                        for (let attempt = 1; attempt <= 3; attempt++) {
                            try {
                                if (attempt > 1) {
                                    const delay = attempt === 2 ? 2000 : 4000;
                                    console.log(`[LiveSession] switchSheet+getData Retry #${attempt} (nach ${delay}ms)...`);
                                    await new Promise(r => setTimeout(r, delay));
                                }
                                const res = await window.electronAPI.liveSessionSwitchSheet(sheetName, true);
                                if (res.success && res.headers) {
                                    // CF-Flag aktualisieren
                                    if (res.hasConditionalFormatting !== undefined) {
                                        explorerState.hasConditionalFormatting = !!res.hasConditionalFormatting;
                                    }
                                    // Sheet wurde automatisch eingeblendet → aus hiddenSheets entfernen
                                    if (res.wasHidden && explorerState.hiddenSheets) {
                                        explorerState.hiddenSheets.delete(sheetName);
                                        updateSheetDropdown();
                                    }
                                    return res;
                                } else if (res.success && res.dataError) {
                                    // Switch ok, aber Daten konnten nicht gelesen werden
                                    console.warn(`[LiveSession] Switch ok, aber getData fehlgeschlagen: ${res.dataError}`);
                                } else {
                                    console.error(`[LiveSession] switchSheet+getData fehlgeschlagen (Versuch ${attempt}):`, res.error);
                                }
                            } catch (err) {
                                console.error(`[LiveSession] switchSheet+getData Fehler (Versuch ${attempt}):`, err);
                            }
                        }
                        return null; // Alle Versuche fehlgeschlagen
                    })(),
                    // ZIP-Metadaten parallel aus gecachtem Buffer (schnell, ~10-30ms)
                    window.electronAPI.readSheetMetadata(explorerState.filePath, diskSheetName).catch(err => {
                        console.warn('[loadExplorerSheet] Metadaten-Extraktion fehlgeschlagen:', err);
                        return { success: false };
                    })
                ]);
                
                const _sheetMs = Date.now() - _sheetStart;
                
                if (switchResult && switchResult.success && switchResult.headers) {
                    // Erfolg: Daten von Excel, Metadaten aus ZIP
                    console.log(`[loadExplorerSheet] LIVE-MODUS fertig in ${_sheetMs}ms — ${switchResult.data?.length} Zeilen, ${switchResult.headers?.length} Spalten`);
                    
                    // SSF-Zahlenformatierung auf Rohdaten anwenden (ersetzt Python _apply_number_formats)
                    const _fmtStart = Date.now();
                    try {
                        const fmtResult = await window.electronAPI.applyNumFmtToLiveData(
                            explorerState.filePath, diskSheetName, switchResult.headers, switchResult.data
                        );
                        if (fmtResult.success && fmtResult.data) {
                            switchResult.data = fmtResult.data;
                            console.log(`[loadExplorerSheet] SSF-Formatierung: ${fmtResult.formattedCols} Spalten in ${Date.now() - _fmtStart}ms`);
                        }
                    } catch (fmtErr) {
                        console.warn('[loadExplorerSheet] SSF-Formatierung fehlgeschlagen (Rohwerte):', fmtErr);
                    }
                    
                    // xlwings liefert headers und data getrennt (nicht data[0]=headers wie ExcelJS)
                    // Hidden Rows/Cols: Bevorzuge COM-Daten aus switchResult (funktioniert auch bei Passwort-Dateien!)
                    // Fallback auf ZIP-Metadaten wenn COM keine Daten liefert
                    const comHiddenRows = switchResult.hiddenRows;
                    const comHiddenCols = switchResult.hiddenColumns;
                    const useComVisibility = Array.isArray(comHiddenRows); // COM hat Daten geliefert
                    
                    if (useComVisibility) {
                        console.log(`[loadExplorerSheet] Versteckte Zeilen/Spalten von COM: ${comHiddenRows.length} Zeilen, ${comHiddenCols.length} Spalten`);
                    }
                    
                    result = {
                        success: true,
                        headers: switchResult.headers,
                        data: switchResult.data, // Bereits ohne Header-Zeile
                        _fromLiveSession: true,  // Marker für unterschiedliche data-Behandlung
                        // Versteckte Zeilen/Spalten: COM-Daten bevorzugen, ZIP als Fallback
                        hiddenColumns: useComVisibility ? comHiddenCols : (metadataResult?.success ? (metadataResult.hiddenColumns || []) : []),
                        hiddenRows: useComVisibility ? comHiddenRows : (metadataResult?.success ? (metadataResult.hiddenRows || []) : []),
                        // Restliche Metadaten weiterhin aus ZIP
                        mergedCells: metadataResult?.success ? (metadataResult.mergedCells || []) : [],
                        autoFilterRange: metadataResult?.success ? (metadataResult.autoFilterRange || null) : null,
                        imageCells: metadataResult?.success ? (metadataResult.imageCells || []) : [],
                        // Zellenbasierte Metadaten nicht verfügbar im Live-Modus (kein ExcelJS-Zellenparse)
                        cellStyles: {},
                        cellFormulas: {},
                        cellHyperlinks: {},
                        richTextCells: {},
                        rowHighlights: [],
                        dataValidations: {}
                    };
                } else {
                    // Live-Session fehlgeschlagen → Fallback auf ExcelJS
                    console.warn(`[loadExplorerSheet] LIVE-MODUS fehlgeschlagen nach ${_sheetMs}ms → Fallback auf ExcelJS`);
                    showFloatingStatus('⚠️ Excel-Daten nicht verfügbar — lade von Disk...', 'warning');
                    result = await window.electronAPI.readExcelSheet(explorerState.filePath, diskSheetName, explorerState.filePassword);
                    const _fallbackMs = Date.now() - _sheetStart;
                    console.log(`[loadExplorerSheet] ExcelJS-Fallback fertig in ${_fallbackMs}ms`);
                }
            } else {
                // === FALLBACK-MODUS: Kein Excel → ExcelJS für alles ===
                console.log(`[loadExplorerSheet] FALLBACK-MODUS: readExcelSheet für "${diskSheetName}"...`);
                
                const [excelResult, switchStatus] = await Promise.all([
                    window.electronAPI.readExcelSheet(explorerState.filePath, diskSheetName, explorerState.filePassword),
                    switchSheetFactory ? (switchSheetPromise = switchSheetFactory()) : (switchSheetPromise || Promise.resolve({ ok: true }))
                ]);
                result = excelResult;
                
                const _sheetMs2 = Date.now() - _sheetStart;
                
                if (!switchStatus.ok && switchSheetPromise) {
                    showFloatingStatus('⚠️ Excel Sheet-Wechsel fehlgeschlagen — Daten wurden von Disk geladen', 'error');
                }
                
                console.log(`[loadExplorerSheet] ExcelJS fertig in ${_sheetMs2}ms — ${result.data?.length} Zeilen, ${result.headers?.length} Spalten, cellStyles: ${Object.keys(result.cellStyles || {}).length}`);
            }
            
            if (!result.success) {
                const _sheetMs = Date.now() - _sheetStart;
                console.log(`[loadExplorerSheet] FEHLER nach ${_sheetMs}ms: ${result.error}`);
                elements.explorerStatus.textContent = `Fehler: ${result.error}`;
                return;
            }
            
            explorerState.selectedSheet = sheetName;
            elements.explorerSheetSelect.value = sheetName;
            explorerState.headers = result.headers;
            // CF-Flag auch im Fallback-/openpyxl-Modus setzen (ExcelJS liefert es aus Sheet-XML)
            if (result.hasConditionalFormatting !== undefined) {
                explorerState.hasConditionalFormatting = !!result.hasConditionalFormatting;
            }
            // Blattschutz-Flag (warnt User, dass Hide nicht wirkt)
            if (result.sheetProtected !== undefined) {
                explorerState.sheetProtected = !!result.sheetProtected;
            }
            explorerState._sheetProtectedWarned = false;
            // xlwings liefert data bereits ohne Header-Zeile, ExcelJS hat Header in data[0]
            explorerState.data = result._fromLiveSession ? result.data : result.data.slice(1);
            // Kopie der Originaldaten speichern (deep copy)
            explorerState.originalData = explorerState.data.map(row => [...row]);
            // Spalten-Sichtbarkeit: Berücksichtige hiddenColumns aus Excel
            if (result.hiddenColumns && result.hiddenColumns.length > 0) {
                // Nur Spalten anzeigen, die nicht in hiddenColumns sind
                const hiddenSet = new Set(result.hiddenColumns);
                explorerState.visibleColumns = explorerState.headers
                    .map((_, i) => i)
                    .filter(i => !hiddenSet.has(i));
            } else {
                // Alle Spalten standardmäßig sichtbar machen
                explorerState.visibleColumns = explorerState.headers.map((_, i) => i);
            }
            explorerState.columnOrder = []; // Reset column order
            // filteredData mit originalIndex initialisieren
            explorerState.filteredData = explorerState.data.map((row, index) => ({ originalIndex: index, row: row }));
            explorerState.currentPage = 1; // Pagination zurücksetzen
            explorerState.searchTerm = ''; // Suche zurücksetzen
            explorerState.sortColumn = null; // Sortierung zurücksetzen
            explorerState.sortDirection = null;
            explorerState.editedCells.clear(); // Bearbeitungen zurücksetzen
            // Zeilenfarben aus Excel laden (oder leeren wenn keine vorhanden)
            explorerState.rowHighlights.clear();
            explorerState.originalRowHighlights.clear();
            if (result.rowHighlights && result.rowHighlights.length > 0) {
                for (const [rowIdx, colorName] of result.rowHighlights) {
                    explorerState.rowHighlights.set(rowIdx, colorName);
                    explorerState.originalRowHighlights.set(rowIdx, colorName);  // Original-Zustand merken
                }
                console.log(`[Explorer] ${result.rowHighlights.length} Zeilenfarben aus Excel geladen`);
            }
            // Data Validations (Dropdown-Listen) laden
            explorerState.dataValidations = result.dataValidations || {};
            // Cell Styles (Formatierungen) laden
            explorerState.cellStyles = result.cellStyles || {};
            // Cell Formulas (Formeln) laden
            explorerState.cellFormulas = result.cellFormulas || {};
            // Cell Hyperlinks (Links) laden
            explorerState.cellHyperlinks = result.cellHyperlinks || {};
            // Rich Text Cells (formatierter Text) laden
            explorerState.richTextCells = result.richTextCells || {};
            // Hidden Rows (ausgeblendete Zeilen) laden
            explorerState.hiddenRows = new Set(result.hiddenRows || []);
            // AutoFilter Range laden
            explorerState.autoFilterRange = result.autoFilterRange || null;
            // Row Mapping initialisieren: [0, 1, 2, ...] = Identity
            // Bedeutet: Position 0 im Frontend = Excel-Zeile 2 (erste Datenzeile)
            explorerState.rowMapping = explorerState.data.map((_, i) => i);
            // Merged Cells laden
            explorerState.mergedCells = result.mergedCells || [];
            // Image Cells VM-Map laden (für Bild-Kopie bei Copy&Paste)
            explorerState.cellVmMap = {};
            if (result.imageCells && result.imageCells.length > 0) {
                for (const ic of result.imageCells) {
                    if (ic.vmValue) {
                        // ic.row = rowNum - 1 = dataRowIndex + 1 (bei Standard-Dateien ohne Lücken)
                        // Das entspricht exakt styleKey-Format: "${dataRowIndex + 1}-${col}"
                        const vmKey = `${ic.row}-${ic.col}`;
                        explorerState.cellVmMap[vmKey] = ic.vmValue;
                    }
                }
                if (Object.keys(explorerState.cellVmMap).length > 0) {
                    console.log(`[Explorer] ${Object.keys(explorerState.cellVmMap).length} vm-Bild-Zellen geladen`);
                }
            }
            elements.explorerSearch.value = '';
            
            // Prüfe ob es gespeicherte Bearbeitungen gibt (Auto-Save Recovery)
            if (window._pendingExplorerRestore && 
                window._pendingExplorerRestore.filePath === explorerState.filePath &&
                window._pendingExplorerRestore.selectedSheet === sheetName) {
                
                const restore = window._pendingExplorerRestore;
                let restoredCount = 0;
                
                restore.editedCells.forEach(([key, value]) => {
                    const [rowStr, colStr] = key.split('-');
                    const rowIndex = parseInt(rowStr);
                    const colIndex = parseInt(colStr);
                    
                    if (rowIndex < explorerState.data.length && colIndex < explorerState.headers.length) {
                        explorerState.editedCells.set(key, value);
                        explorerState.data[rowIndex][colIndex] = value;
                        restoredCount++;
                    }
                });
                
                if (restoredCount > 0) {
                    showUndoRedoFeedback(`${restoredCount} ${t('editingsRestored')}`);
                }
                
                delete window._pendingExplorerRestore;
            }
            
            // Status-Meldung für große Dateien
            if (explorerState.data.length > 1000) {
                elements.explorerStatus.textContent = `${explorerState.data.length} ${t('rowsLoaded')} (${t('paginationActive')})`;
            } else {
                elements.explorerStatus.textContent = '';
            }
            
            // WICHTIG: filterExplorerData() statt renderExplorerTable() verwenden,
            // damit versteckte Zeilen (hiddenRows) korrekt ausgefiltert werden!
            filterExplorerData();
            updateColumnToggles();
            updateHiddenRowsIndicator();
            updateHiddenColumnsIndicator();
            updateAutoFilterIndicator();
            
            // GUI ist gerendert → jetzt Excel visuell wechseln (separater Befehl)
            // AWAIT — Buttons bleiben gesperrt bis activate() fertig (hat 3s Timeout in Python)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    await window.electronAPI.liveSessionActivateSheet(sheetName);
                } catch (err) {
                    console.error('[LiveSession] activateSheet Fehler:', err);
                }
            }
            } finally {
                // === UI ENTSPERREN ===
                const switchOverlayEnd = document.getElementById('explorerSheetSwitchOverlay');
                if (switchOverlayEnd) switchOverlayEnd.classList.add('hidden');
                // Alle vorher gesperrten Buttons wieder aktivieren
                if (typeof _disabledExplorerButtons !== 'undefined') {
                    _disabledExplorerButtons.forEach(btn => { btn.disabled = false; });
                }
                if (elements.explorerSearch) elements.explorerSearch.disabled = false;
                
                // Guard aufheben — Sheet-Wechsel wieder erlauben
                explorerState.isLoadingSheet = false;
                elements.explorerSheetSelect.disabled = false;
                
                // Wurde während des Ladevorgangs ein neues Sheet angefordert?
                // → Automatisch zum letzten gewünschten Sheet wechseln
                if (explorerState.pendingSheetSwitch && explorerState.pendingSheetSwitch !== explorerState.selectedSheet) {
                    const pendingSheet = explorerState.pendingSheetSwitch;
                    explorerState.pendingSheetSwitch = null;
                    console.log(`[loadExplorerSheet] Pending Sheet-Wechsel ausführen: "${pendingSheet}"`);
                    loadExplorerSheet(pendingSheet); // Nicht await — wird async gestartet
                } else {
                    explorerState.pendingSheetSwitch = null;
                }
            }
        }
        
        // Speichert das aktuelle Sheet im Cache
        function saveCurrentSheetToCache() {
            if (!explorerState.selectedSheet) return;
            
            explorerState.sheetDataCache.set(explorerState.selectedSheet, {
                headers: [...explorerState.headers],
                data: explorerState.data.map(row => [...row]),
                originalData: explorerState.originalData.map(row => [...row]),
                editedCells: new Map(explorerState.editedCells),
                rowHighlights: new Map(explorerState.rowHighlights),
                visibleColumns: [...explorerState.visibleColumns],
                columnOrder: [...explorerState.columnOrder],
                dataValidations: { ...explorerState.dataValidations },
                cellStyles: { ...explorerState.cellStyles },
                cellFormulas: { ...explorerState.cellFormulas },
                cellHyperlinks: { ...explorerState.cellHyperlinks },
                richTextCells: { ...explorerState.richTextCells },
                hiddenRows: new Set(explorerState.hiddenRows),
                autoFilterRange: explorerState.autoFilterRange,
                mergedCells: [...explorerState.mergedCells],
                rowMapping: explorerState.rowMapping ? [...explorerState.rowMapping] : null,
                columnOperationsQueue: [...explorerState.columnOperationsQueue]
            });
        }

        /**
         * Lädt alle noch nicht gecachten Sheets im Hintergrund (nur openpyxl-Modus).
         * Läuft sequenziell mit requestIdleCallback/setTimeout zwischen Sheets, damit
         * die UI responsiv bleibt. Kann jederzeit via preloadToken.cancelled gestoppt
         * werden (Datei-Wechsel, Export, Close).
         */
        async function startBackgroundSheetPreload() {
            // Nur im openpyxl-Fallback (Live-Mode hat alle Daten schon in Excel)
            if (explorerState.engineMode === 'live') return;
            // Opt-Out per LocalStorage
            if (localStorage.getItem('preloadSheetsInBackground') === 'false') return;
            if (!explorerState.filePath || !Array.isArray(explorerState.sheets)) return;

            const token = { cancelled: false };
            explorerState._preloadToken = token;

            const fileAtStart = explorerState.filePath;
            const filePassword = explorerState.filePassword;

            const toPreload = explorerState.sheets.filter(s => {
                if (s === explorerState.selectedSheet) return false;
                if (explorerState.sheetDataCache.has(s)) return false;
                return true;
            });

            if (toPreload.length === 0) return;
            console.log(`[Preload] Starte Hintergrund-Preload für ${toPreload.length} Sheets`);

            const idle = (cb) => {
                if (typeof window.requestIdleCallback === 'function') {
                    window.requestIdleCallback(cb, { timeout: 2000 });
                } else {
                    setTimeout(cb, 50);
                }
            };

            for (const sheetName of toPreload) {
                if (token.cancelled) { console.log('[Preload] abgebrochen'); return; }
                // Datei gewechselt? → stoppen
                if (explorerState.filePath !== fileAtStart) { console.log('[Preload] Datei hat gewechselt, stoppe'); return; }
                // Sheet wurde inzwischen manuell geladen oder ist aktuell ausgewählt
                if (explorerState.sheetDataCache.has(sheetName)) continue;
                if (sheetName === explorerState.selectedSheet) continue;

                // Auf Idle warten
                await new Promise(res => idle(res));
                if (token.cancelled || explorerState.filePath !== fileAtStart) return;

                // Disk-Sheet-Namen berücksichtigen (bei umbenannten Sheets)
                let diskSheetName = sheetName;
                if (explorerState.sheetDiskNameMap && explorerState.sheetDiskNameMap.has(sheetName)) {
                    diskSheetName = explorerState.sheetDiskNameMap.get(sheetName);
                }

                try {
                    const t0 = performance.now();
                    const res = await window.electronAPI.readExcelSheet(fileAtStart, diskSheetName, filePassword);
                    if (token.cancelled || explorerState.filePath !== fileAtStart) return;
                    if (!res || !res.success) {
                        console.warn(`[Preload] "${sheetName}" fehlgeschlagen: ${res?.error || 'unbekannt'}`);
                        continue;
                    }
                    // Nicht cachen, falls User inzwischen manuell geladen hat
                    if (explorerState.sheetDataCache.has(sheetName)) continue;

                    // Cache-Eintrag aufbauen (gleiche Struktur wie saveCurrentSheetToCache)
                    const data = res.data.slice(1); // ExcelJS: data[0] = headers
                    const headers = res.headers;
                    const hiddenColsSet = new Set(res.hiddenColumns || []);
                    const visibleColumns = headers.map((_, i) => i).filter(i => !hiddenColsSet.has(i));
                    const rowHighlights = new Map();
                    if (res.rowHighlights && res.rowHighlights.length > 0) {
                        for (const [rowIdx, color] of res.rowHighlights) rowHighlights.set(rowIdx, color);
                    }

                    explorerState.sheetDataCache.set(sheetName, {
                        headers: [...headers],
                        data: data.map(row => [...row]),
                        originalData: data.map(row => [...row]),
                        editedCells: new Map(),
                        rowHighlights,
                        visibleColumns,
                        columnOrder: [],
                        dataValidations: res.dataValidations || {},
                        cellStyles: res.cellStyles || {},
                        cellFormulas: res.cellFormulas || {},
                        cellHyperlinks: res.cellHyperlinks || {},
                        richTextCells: res.richTextCells || {},
                        hiddenRows: new Set(res.hiddenRows || []),
                        autoFilterRange: res.autoFilterRange || null,
                        mergedCells: res.mergedCells || [],
                        rowMapping: data.map((_, i) => i),
                        columnOperationsQueue: []
                    });
                    console.log(`[Preload] "${sheetName}" gecacht (${data.length} Zeilen, ${(performance.now() - t0).toFixed(0)}ms)`);
                } catch (err) {
                    if (token.cancelled) return;
                    console.warn(`[Preload] "${sheetName}" Fehler:`, err?.message || err);
                }
            }
            console.log('[Preload] fertig');
        }
        
        // Prüft ob es ungespeicherte Änderungen in irgendeinem Sheet gibt
        function hasUnsavedChanges() {
            // Pending Sheet-Operationen prüfen
            if (explorerState.pendingSheetOperations.length > 0) {
                console.log('[hasUnsavedChanges] pendingSheetOperations:', explorerState.pendingSheetOperations.length);
                return true;
            }
            
            // Aktuelles Sheet prüfen
            if (explorerState.editedCells.size > 0) {
                const keys = Array.from(explorerState.editedCells.keys()).slice(0, 5);
                console.log(`[hasUnsavedChanges] editedCells: ${explorerState.editedCells.size} Einträge, Keys: ${keys.join(', ')}`);
                return true;
            }
            
            // Cache prüfen
            for (const [sheetName, cached] of explorerState.sheetDataCache) {
                if (cached.editedCells.size > 0) {
                    const keys = Array.from(cached.editedCells.keys()).slice(0, 5);
                    console.log(`[hasUnsavedChanges] Cache "${sheetName}": ${cached.editedCells.size} Einträge, Keys: ${keys.join(', ')}`);
                    return true;
                }
            }
            
            return false;
        }
        
        // Zählt alle Änderungen über alle Sheets
        function countAllChanges() {
            let total = explorerState.editedCells.size;
            
            for (const [sheetName, cached] of explorerState.sheetDataCache) {
                if (sheetName !== explorerState.selectedSheet) {
                    total += cached.editedCells.size;
                }
            }
            
            // Pending Sheet-Operationen zählen (Offline + Live)
            total += explorerState.pendingSheetOperations.length;
            total += explorerState.liveSheetChanges;
            
            return total;
        }
        
        /**
         * Parst einen Datumswert aus verschiedenen Formaten
         * Unterstützt: Excel-Seriennummern, ISO-Daten, deutsche Datumsformate
         */
        /**
         * Konvertiert 2-stellige Jahreszahl zu 4-stellig (00-49 → 2000-2049, 50-99 → 1950-1999)
         */
        function expandYear(y) {
            if (y < 100) return y < 50 ? 2000 + y : 1900 + y;
            return y;
        }
        
        /**
         * Monatsname (deutsch/englisch, kurz/lang) zu Monatsnummer (0-basiert)
         */
        function monthNameToIndex(name) {
            const n = name.toLowerCase().replace(/[äö]/g, m => m === 'ä' ? 'a' : 'o');
            const months = {
                'jan': 0, 'januar': 0, 'january': 0,
                'feb': 1, 'februar': 1, 'february': 1,
                'mar': 2, 'marz': 2, 'mär': 2, 'märz': 2, 'march': 2,
                'apr': 3, 'april': 3,
                'mai': 4, 'may': 4,
                'jun': 5, 'juni': 5, 'june': 5,
                'jul': 6, 'juli': 6, 'july': 6,
                'aug': 7, 'august': 7,
                'sep': 8, 'sept': 8, 'september': 8,
                'okt': 9, 'oct': 9, 'oktober': 9, 'october': 9,
                'nov': 10, 'november': 10,
                'dez': 11, 'dec': 11, 'dezember': 11, 'december': 11
            };
            return months[name.toLowerCase()] ?? months[n] ?? -1;
        }
        
        function parseDateValue(value, dateOrder) {
            if (!value) return null;
            
            // Bereits ein Date-Objekt
            if (value instanceof Date && !isNaN(value)) return value;
            
            // Excel-Seriennummer (Zahl zwischen 1 und 100000)
            if (typeof value === 'number' && value > 0 && value < 100000) {
                // Excel-Datum: Tage seit 1899-12-30 (mit dem berühmten Leap-Year-Bug)
                const excelEpoch = new Date(1899, 11, 30);
                const date = new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
                if (!isNaN(date)) return date;
            }
            
            const str = String(value).trim();
            if (!str) return null;
            
            // ISO-Format: 2026-01-08 (auch mit optionaler Zeit)
            const isoMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
            if (isoMatch) {
                const date = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
                if (!isNaN(date)) return date;
            }
            
            // Punkt-Trenner: DD.MM.YYYY (deutsch) oder MM.DD.YYYY (US mit Punkt)
            // Auto-Erkennung + optionaler dateOrder-Hint aus Spaltendaten
            const dotMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
            if (dotMatch) {
                const p1 = parseInt(dotMatch[1]);
                const p2 = parseInt(dotMatch[2]);
                const year = expandYear(parseInt(dotMatch[3]));
                let day, month;
                if (p2 > 12 && p1 <= 12) {
                    // p2 kann kein Monat sein → MM.DD Format (z.B. 1.31.26)
                    month = p1; day = p2;
                } else if (p1 > 12 && p2 <= 12) {
                    // p1 kann kein Monat sein → DD.MM Format (z.B. 31.01.26)
                    day = p1; month = p2;
                } else if (dateOrder === 'mdy') {
                    // Spaltenformat sagt MM.DD (z.B. 1.05.26 = Jan 5)
                    month = p1; day = p2;
                } else {
                    // Default oder dateOrder === 'dmy' → DD.MM (deutsch)
                    day = p1; month = p2;
                }
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    const date = new Date(year, month - 1, day);
                    // Validierung: JS Date kann ungültige Werte wrappen (z.B. Tag 31 bei 30-Tage-Monat)
                    if (!isNaN(date) && date.getMonth() === month - 1 && date.getDate() === day) {
                        return date;
                    }
                }
            }
            
            // US/Internationales Format mit Schrägstrich: 01/08/2026, 1/8/26
            const usMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
            if (usMatch) {
                const year = expandYear(parseInt(usMatch[3]));
                const date = new Date(year, parseInt(usMatch[1]) - 1, parseInt(usMatch[2]));
                if (!isNaN(date)) return date;
            }
            
            // Bindestrich-Format: 08-01-2026, 8-1-26 (Tag-Monat-Jahr)
            const dashMatch = str.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})(?:\s|$)/);
            if (dashMatch) {
                const year = expandYear(parseInt(dashMatch[3]));
                const date = new Date(year, parseInt(dashMatch[2]) - 1, parseInt(dashMatch[1]));
                if (!isNaN(date)) return date;
            }
            
            // Monatsname-Formate: "8-Jan-26", "08-Jan-2026", "8. Jan 2026", "8. Januar 2026"
            const monthNameMatch = str.match(/^(\d{1,2})[\.\-/\s]+([A-Za-zÄäÖöÜü]+)[\.\-/\s]+(\d{2,4})/);
            if (monthNameMatch) {
                const mi = monthNameToIndex(monthNameMatch[2]);
                if (mi >= 0) {
                    const year = expandYear(parseInt(monthNameMatch[3]));
                    const date = new Date(year, mi, parseInt(monthNameMatch[1]));
                    if (!isNaN(date)) return date;
                }
            }
            
            // Umgekehrtes Monatsname-Format: "Jan 8, 2026", "Januar 8 2026"
            const monthFirstMatch = str.match(/^([A-Za-zÄäÖöÜü]+)[\.\-/\s]+(\d{1,2})[,.\s]+(\d{2,4})/);
            if (monthFirstMatch) {
                const mi = monthNameToIndex(monthFirstMatch[1]);
                if (mi >= 0) {
                    const year = expandYear(parseInt(monthFirstMatch[3]));
                    const date = new Date(year, mi, parseInt(monthFirstMatch[2]));
                    if (!isNaN(date)) return date;
                }
            }
            
            // Fallback: JavaScript Date-Parser
            const fallback = new Date(str);
            if (!isNaN(fallback)) return fallback;
            
            return null;
        }
        
        /**
         * Erkennt ein Beispiel-Datum aus einer Spalte für den Placeholder-Text
         * Gibt den ersten nicht-leeren Datumswert der Spalte als String zurück
         */
        function getDateSampleFromColumn(colIndex) {
            if (!explorerState.data || colIndex < 0) return '';
            for (let i = 0; i < Math.min(explorerState.data.length, 50); i++) {
                const val = explorerState.data[i][colIndex];
                if (val && String(val).trim() !== '') {
                    const parsed = parseDateValue(val);
                    if (parsed) return String(val).trim();
                }
            }
            return '';
        }
        
        /**
         * Erkennt ob eine Spalte MM.DD oder DD.MM Reihenfolge bei Punkt-Format verwendet.
         * Sucht in den Spaltendaten nach eindeutigen Werten (Teil > 12).
         * @returns 'mdy' | 'dmy'
         */
        function detectColumnDateOrder(colIndex) {
            if (!explorerState.data || colIndex < 0) return 'dmy';
            for (let i = 0; i < Math.min(explorerState.data.length, 200); i++) {
                const val = explorerState.data[i]?.[colIndex];
                if (!val) continue;
                const str = String(val).trim();
                const m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
                if (m) {
                    const p1 = parseInt(m[1]);
                    const p2 = parseInt(m[2]);
                    if (p2 > 12 && p1 <= 12) return 'mdy'; // z.B. 1.31.26 → MM.DD
                    if (p1 > 12 && p2 <= 12) return 'dmy'; // z.B. 31.01.26 → DD.MM
                }
            }
            return 'dmy'; // Default: deutsch
        }
        
        /**
         * Aktualisiert die Placeholders der Von/Bis-Felder basierend auf Spalten-Daten
         */
        function updateDatePlaceholders(row) {
            const colSelect = row.querySelector('.filter-column');
            const dateFromInput = row.querySelector('.filter-date-from');
            const dateToInput = row.querySelector('.filter-date-to');
            const colIndex = parseInt(colSelect.value);
            
            if (isNaN(colIndex) || colIndex < 0) {
                dateFromInput.placeholder = 'Von...';
                dateToInput.placeholder = 'Bis...';
                return;
            }
            
            const sample = getDateSampleFromColumn(colIndex);
            if (sample) {
                dateFromInput.placeholder = `Von z.B. ${sample}`;
                dateToInput.placeholder = `Bis z.B. ${sample}`;
            } else {
                dateFromInput.placeholder = 'Von...';
                dateToInput.placeholder = 'Bis...';
            }
        }
        
        /**
         * Prüft ob ein Datum den Filter-Bedingungen entspricht
         */
        function matchDateFilter(cellValue, operator, days = 0, dateFrom = null, dateTo = null, dateOrder = null) {
            const cellDate = parseDateValue(cellValue, dateOrder);
            if (!cellDate) return false;
            
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            cellDate.setHours(0, 0, 0, 0);
            
            const diffMs = cellDate.getTime() - today.getTime();
            const diffDays = Math.round(diffMs / (24 * 60 * 60 * 1000));
            
            switch (operator) {
                case 'dateInDays':
                    // Datum liegt in der Zeitspanne von heute bis +X Tage
                    // Beispiel: days=7 zeigt alles was heute bis in 7 Tagen fällig wird
                    return diffDays >= 0 && diffDays <= days;
                    
                case 'dateOverdueDays':
                    // Datum liegt in der Zeitspanne von heute bis -X Tage
                    // Beispiel: days=7 zeigt alles was in den letzten 7 Tagen überfällig wurde
                    return diffDays < 0 && Math.abs(diffDays) <= days;
                    
                case 'dateBetween': {
                    // Datum liegt im angegebenen Zeitraum von-bis
                    if (!dateFrom && !dateTo) return false;
                    
                    let fromDate = null;
                    let toDate = null;
                    
                    // WICHTIG: parseDateValue mit dateOrder verwenden!
                    // Gleiche Spaltenformat-Erkennung für Filtereingabe und Zellwerte
                    if (dateFrom) {
                        fromDate = parseDateValue(dateFrom, dateOrder);
                        if (fromDate) fromDate.setHours(0, 0, 0, 0);
                    }
                    if (dateTo) {
                        toDate = parseDateValue(dateTo, dateOrder);
                        if (toDate) toDate.setHours(23, 59, 59, 999); // Ende des Tages
                    }
                    
                    // Eingabe vorhanden, aber nicht parsebar (z.B. unvollständige Eingabe)
                    // → nicht filtern, alle Zeilen anzeigen
                    if (!fromDate && !toDate) return true;
                    
                    // Wenn nur Von-Datum: alles ab diesem Datum
                    if (fromDate && !toDate) {
                        return cellDate >= fromDate;
                    }
                    // Wenn nur Bis-Datum: alles bis zu diesem Datum
                    if (!fromDate && toDate) {
                        return cellDate <= toDate;
                    }
                    // Beide Daten: Bereich prüfen
                    return cellDate >= fromDate && cellDate <= toDate;
                }
                    
                case 'dateToday':
                    return diffDays === 0;
                    
                case 'datePast':
                    return diffDays < 0;
                    
                case 'dateFuture':
                    return diffDays > 0;
                    
                case 'dateThisWeek':
                    // Diese Woche = 0 bis 6 Tage in der Zukunft oder Vergangenheit ab Montag
                    const dayOfWeek = today.getDay(); // 0=So, 1=Mo, ..., 6=Sa
                    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
                    const monday = new Date(today);
                    monday.setDate(today.getDate() + mondayOffset);
                    const sunday = new Date(monday);
                    sunday.setDate(monday.getDate() + 6);
                    return cellDate >= monday && cellDate <= sunday;
                    
                case 'dateThisMonth':
                    return cellDate.getMonth() === today.getMonth() && 
                           cellDate.getFullYear() === today.getFullYear();
                    
                default:
                    return false;
            }
        }
        
        function filterExplorerData() {
            // Erstelle Array mit originalem Index für jede Zeile
            let filtered = explorerState.data.map((row, index) => ({ originalIndex: index, row: row }));
            
            // Versteckte Zeilen ausfiltern
            if (explorerState.hiddenRows.size > 0) {
                filtered = filtered.filter(item => !explorerState.hiddenRows.has(item.originalIndex));
            }
            
            // Volltextsuche mit Platzhalter-Unterstützung (* und ?)
            if (explorerState.searchTerm) {
                const term = explorerState.searchTerm;
                const hasWildcards = term.includes('*') || term.includes('?');
                
                if (hasWildcards) {
                    // Platzhalter-Suche: * = beliebig viele Zeichen, ? = ein Zeichen
                    const regex = wildcardToRegex(term);
                    filtered = filtered.filter(item => 
                        item.row.some(cell => cell && regex.test(String(cell)))
                    );
                } else {
                    // Normale Suche (case-insensitive, enthält)
                    const lowerTerm = term.toLowerCase();
                    filtered = filtered.filter(item => 
                        item.row.some(cell => cell && String(cell).toLowerCase().includes(lowerTerm))
                    );
                }
            }
            
            // Spaltenfilter — gleiche Spalte = ODER, verschiedene Spalten = UND
            const validFilters = explorerState.filters.filter(f => f.column);
            const filtersByColumn = new Map();
            validFilters.forEach(filter => {
                const col = filter.column;
                if (!filtersByColumn.has(col)) filtersByColumn.set(col, []);
                filtersByColumn.get(col).push(filter);
            });
            
            // Pro Spalten-Gruppe: positive Filter (contains, equals, ...) = OR,
            // negative Filter (notContains) = AND. Über alle Spalten-Gruppen: AND.
            filtersByColumn.forEach((columnFilters, col) => {
                const colIndex = parseInt(col);
                
                // Aufteilen in positive und negative Filter
                const positiveFilters = columnFilters.filter(f => f.operator !== 'notContains');
                const negativeFilters = columnFilters.filter(f => f.operator === 'notContains');
                
                filtered = filtered.filter(item => {
                    const cellValue = String(item.row[colIndex] || '').toLowerCase();
                    const rawCellValue = item.row[colIndex];
                    
                    // Positive Filter: OR (mindestens einer muss matchen)
                    const positiveMatch = positiveFilters.length === 0 || positiveFilters.some(filter => {
                        const value = (filter.value || '').toLowerCase();
                        switch (filter.operator) {
                            case 'contains': return cellValue.includes(value);
                            case 'equals': return cellValue === value;
                            case 'startsWith': return cellValue.startsWith(value);
                            case 'endsWith': return cellValue.endsWith(value);
                            case 'isEmpty': return !rawCellValue || String(rawCellValue).trim() === '';
                            case 'isNotEmpty': return rawCellValue && String(rawCellValue).trim() !== '';
                            case 'dateBetween': {
                                const colDateOrder = detectColumnDateOrder(colIndex);
                                return matchDateFilter(rawCellValue, filter.operator, 0, filter.dateFrom, filter.dateTo, colDateOrder);
                            }
                            case 'dateInDays': case 'dateOverdueDays': case 'dateToday':
                            case 'datePast': case 'dateFuture': case 'dateThisWeek': case 'dateThisMonth': {
                                const days = parseInt(filter.value) || 0;
                                const colDateOrder2 = detectColumnDateOrder(colIndex);
                                return matchDateFilter(rawCellValue, filter.operator, days, null, null, colDateOrder2);
                            }
                            default: return true;
                        }
                    });
                    
                    // Negative Filter: AND (ALLE müssen matchen = Zeile darf KEINEN der Begriffe enthalten)
                    const negativeMatch = negativeFilters.every(filter => {
                        const value = (filter.value || '').toLowerCase();
                        return !cellValue.includes(value);
                    });
                    
                    return positiveMatch && negativeMatch;
                });
            });
            
            explorerState.filteredData = filtered;
            explorerState.currentPage = 1; // Bei Filteränderung zurück zur ersten Seite
            
            // Live-Session: Button anzeigen wenn aktiv (Filter werden erst bei Enter/Button-Klick gesendet)
            updateSyncFiltersButton();
            
            // Sortierung anwenden, wenn eine gesetzt ist
            if (explorerState.sortColumn !== null && explorerState.sortDirection !== null) {
                applyExplorerSort();
            }
            
            renderExplorerTable();
        }
        
        /**
         * Aktualisiert die Sichtbarkeit des "An Excel senden" Buttons
         */
        function updateSyncFiltersButton() {
            const btn = document.getElementById('btnSyncFiltersToExcel');
            if (!btn) return;
            
            const hasFilters = explorerState.filters.some(f => f.column && f.value);
            const isLiveActive = explorerState.liveSessionActive && explorerState.liveSessionReady;
            
            btn.style.display = isLiveActive ? 'inline-block' : 'none';
            btn.disabled = !hasFilters;
        }
        
        /**
         * Synchronisiert die aktuellen Filter mit Excel's AutoFilter
         */
        async function syncFiltersToExcel() {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) {
                return;
            }
            
            // Concurrency Guard: Nur ein COM-Call gleichzeitig
            if (syncFiltersToExcel._running) {
                syncFiltersToExcel._pending = true;
                return;
            }
            syncFiltersToExcel._running = true;
            syncFiltersToExcel._pending = false;
            
            try {
                const noValueOps = ['isEmpty', 'isNotEmpty', 'dateToday', 'datePast', 'dateFuture', 'dateThisWeek', 'dateThisMonth'];
                const dateOps = ['dateInDays', 'dateOverdueDays', 'dateBetween', 'dateToday', 'datePast', 'dateFuture', 'dateThisWeek', 'dateThisMonth'];
                
                // Konvertiere Explorer-Filter zu Excel-AutoFilter-Format
                const excelFilters = explorerState.filters
                    .filter(f => {
                        if (!f.column) return false;
                        // Operatoren die keinen Wert brauchen
                        if (noValueOps.includes(f.operator)) return true;
                        // dateBetween braucht dateFrom oder dateTo
                        if (f.operator === 'dateBetween') return !!(f.dateFrom || f.dateTo);
                        // dateInDays/dateOverdueDays brauchen days
                        if (f.operator === 'dateInDays' || f.operator === 'dateOverdueDays') return true;
                        return !!f.value;
                    })
                    .map(f => {
                        const op = f.operator || 'contains';
                        
                        // Datumsfilter: Berechne konkrete Datumsgrenzen
                        if (dateOps.includes(op)) {
                            const today = new Date();
                            today.setHours(0, 0, 0, 0);
                            const fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
                            
                            let dateFrom = null, dateTo = null;
                            
                            switch (op) {
                                case 'dateToday':
                                    dateFrom = fmt(today);
                                    dateTo = fmt(today);
                                    break;
                                case 'datePast': {
                                    dateTo = fmt(new Date(today.getTime() - 86400000)); // gestern
                                    break;
                                }
                                case 'dateFuture': {
                                    dateFrom = fmt(new Date(today.getTime() + 86400000)); // morgen
                                    break;
                                }
                                case 'dateThisWeek': {
                                    const dow = today.getDay();
                                    const mondayOff = dow === 0 ? -6 : 1 - dow;
                                    const mon = new Date(today);
                                    mon.setDate(today.getDate() + mondayOff);
                                    const sun = new Date(mon);
                                    sun.setDate(mon.getDate() + 6);
                                    dateFrom = fmt(mon);
                                    dateTo = fmt(sun);
                                    break;
                                }
                                case 'dateThisMonth': {
                                    const first = new Date(today.getFullYear(), today.getMonth(), 1);
                                    const last = new Date(today.getFullYear(), today.getMonth() + 1, 0);
                                    dateFrom = fmt(first);
                                    dateTo = fmt(last);
                                    break;
                                }
                                case 'dateInDays': {
                                    const days = parseInt(f.value) || parseInt(f.days) || 0;
                                    dateFrom = fmt(today);
                                    dateTo = fmt(new Date(today.getTime() + days * 86400000));
                                    break;
                                }
                                case 'dateOverdueDays': {
                                    const days = parseInt(f.value) || parseInt(f.days) || 0;
                                    dateFrom = fmt(new Date(today.getTime() - days * 86400000));
                                    dateTo = fmt(new Date(today.getTime() - 86400000)); // bis gestern
                                    break;
                                }
                                case 'dateBetween': {
                                    // User tippt Datum im Tabellenformat (z.B. "1.31.26" oder "14.02.2026")
                                    // Spalten-Datumsformat erkennen für korrekte Interpretation
                                    const colIdx = parseInt(f.column);
                                    const dateOrd = detectColumnDateOrder(colIdx);
                                    if (f.dateFrom) {
                                        const parsed = parseDateValue(f.dateFrom, dateOrd);
                                        dateFrom = parsed ? fmt(parsed) : null;
                                    }
                                    if (f.dateTo) {
                                        const parsed = parseDateValue(f.dateTo, dateOrd);
                                        dateTo = parsed ? fmt(parsed) : null;
                                    }
                                    console.log(`[syncFilters] dateBetween col ${colIdx} order=${dateOrd}: "${f.dateFrom}" → ${dateFrom}, "${f.dateTo}" → ${dateTo}`);
                                    break;
                                }
                            }
                            
                            return {
                                colIndex: parseInt(f.column),
                                operator: op,
                                criteria: '',
                                dateFrom: dateFrom,
                                dateTo: dateTo
                            };
                        }
                        
                        // isEmpty/isNotEmpty: keinen Criteria-Wert senden
                        const criteria = (op === 'isEmpty' || op === 'isNotEmpty') ? '' : f.value;
                        return {
                            colIndex: parseInt(f.column),
                            criteria: criteria,
                            operator: op
                        };
                    });
                
                // setAutoFilter mit leerer Liste = nur App-eigene Filter entfernen
                // (pre-existierende Excel-Filter bleiben erhalten)
                console.log('[syncFilters] Sende', excelFilters.length, 'Filter an Excel:', JSON.stringify(excelFilters));
                const result = await window.electronAPI.liveSessionSetAutoFilter(excelFilters);
                console.log('[syncFilters] Ergebnis:', JSON.stringify(result));
                if (result && !result.success) {
                    console.error('[LiveSession] Filter fehlgeschlagen:', result);
                }
                return result;
            } catch (error) {
                console.error('[LiveSession] Filter-Sync Fehler:', error);
                return { success: false, error: error.message };
            } finally {
                syncFiltersToExcel._running = false;
                // Falls während des Laufs ein neuer Call reinkam, jetzt nachholen
                if (syncFiltersToExcel._pending) {
                    syncFiltersToExcel._pending = false;
                    syncFiltersToExcel();
                }
            }
        }
        
        /**
         * Sortiert die Explorer-Daten nach einer Spalte
         */
        function sortExplorerByColumn(colIndex, sortType = 'auto') {
            // Toggle Sortierrichtung bei 'auto'
            if (sortType === 'auto') {
                if (explorerState.sortColumn === colIndex) {
                    if (explorerState.sortDirection === 'asc') {
                        explorerState.sortDirection = 'desc';
                    } else if (explorerState.sortDirection === 'desc') {
                        // Dritter Klick: Sortierung aufheben
                        explorerState.sortColumn = null;
                        explorerState.sortDirection = null;
                        explorerState.sortType = 'auto';
                        filterExplorerData(); // Neu filtern ohne Sortierung
                        return;
                    }
                } else {
                    explorerState.sortColumn = colIndex;
                    explorerState.sortDirection = 'asc';
                }
                explorerState.sortType = 'auto';
            } else {
                // Expliziter Sortiertyp vom Kontextmenü
                explorerState.sortColumn = colIndex;
                explorerState.sortType = sortType;
                // Direction aus dem sortType ableiten
                explorerState.sortDirection = sortType.endsWith('-asc') ? 'asc' : 'desc';
            }
            
            applyExplorerSort();
            renderExplorerTable();
        }
        
        /**
         * Wendet die aktuelle Sortierung auf filteredData an
         */
        function applyExplorerSort() {
            if (explorerState.sortColumn === null) return;
            
            const colIndex = explorerState.sortColumn;
            const direction = explorerState.sortDirection;
            const sortType = explorerState.sortType || 'auto';
            
            explorerState.filteredData.sort((a, b) => {
                const valA = a.row[colIndex];
                const valB = b.row[colIndex];
                
                // Null-Werte ans Ende
                if (valA == null && valB == null) return 0;
                if (valA == null) return 1;
                if (valB == null) return -1;
                
                let comparison = 0;
                
                if (sortType === 'alpha-asc' || sortType === 'alpha-desc') {
                    // Alphabetische Sortierung
                    const strA = String(valA).toLowerCase();
                    const strB = String(valB).toLowerCase();
                    comparison = strA.localeCompare(strB, 'de');
                } else if (sortType === 'num-asc' || sortType === 'num-desc') {
                    // Numerische Sortierung
                    const numA = parseNumericValue(valA);
                    const numB = parseNumericValue(valB);
                    if (isNaN(numA) && isNaN(numB)) comparison = 0;
                    else if (isNaN(numA)) comparison = 1;
                    else if (isNaN(numB)) comparison = -1;
                    else comparison = numA - numB;
                } else if (sortType === 'date-asc' || sortType === 'date-desc') {
                    // Datums-Sortierung
                    const dateA = parseDateValue(valA);
                    const dateB = parseDateValue(valB);
                    if (!dateA && !dateB) comparison = 0;
                    else if (!dateA) comparison = 1;
                    else if (!dateB) comparison = -1;
                    else comparison = dateA.getTime() - dateB.getTime();
                } else {
                    // Auto: Numerisch wenn möglich, sonst alphabetisch
                    const numA = parseFloat(valA);
                    const numB = parseFloat(valB);
                    
                    if (!isNaN(numA) && !isNaN(numB)) {
                        comparison = numA - numB;
                    } else {
                        const strA = String(valA).toLowerCase();
                        const strB = String(valB).toLowerCase();
                        if (strA < strB) comparison = -1;
                        else if (strA > strB) comparison = 1;
                        else comparison = 0;
                    }
                }
                
                return direction === 'asc' ? comparison : -comparison;
            });
        }
        
        /**
         * Parst einen Wert als Zahl (berücksichtigt deutsche Zahlenformate)
         */
        function parseNumericValue(value) {
            if (typeof value === 'number') return value;
            if (value == null) return NaN;
            const str = String(value).trim();
            // Deutsche Zahlen: 1.234,56 → 1234.56
            const normalized = str.replace(/\./g, '').replace(',', '.');
            return parseFloat(normalized);
        }
        
        // ==================== Column Context Menu ====================
        let contextMenuColumn = null;
        
        function showColumnContextMenu(e, colIndex) {
            e.preventDefault();
            
            const menu = document.getElementById('columnContextMenu');
            const columnName = explorerState.headers[colIndex] || `Spalte ${colIndex + 1}`;
            
            document.getElementById('contextMenuColumnName').textContent = columnName;
            contextMenuColumn = colIndex;
            
            // Positionierung
            let x = e.clientX;
            let y = e.clientY;
            
            // Sicherstellen, dass das Menü im Viewport bleibt
            menu.classList.remove('hidden');
            menu.style.display = 'block';
            const menuRect = menu.getBoundingClientRect();
            
            if (x + menuRect.width > window.innerWidth) {
                x = window.innerWidth - menuRect.width - 10;
            }
            if (y + menuRect.height > window.innerHeight) {
                y = window.innerHeight - menuRect.height - 10;
            }
            
            menu.style.left = x + 'px';
            menu.style.top = y + 'px';
        }
        
        function hideColumnContextMenu() {
            const menu = document.getElementById('columnContextMenu');
            if (menu) {
                menu.classList.add('hidden');
                menu.style.display = 'none';
            }
            contextMenuColumn = null;
        }
        
        function handleContextMenuAction(action) {
            if (contextMenuColumn === null) return;
            
            switch (action) {
                case 'sort-alpha-asc':
                    sortExplorerByColumn(contextMenuColumn, 'alpha-asc');
                    break;
                case 'sort-alpha-desc':
                    sortExplorerByColumn(contextMenuColumn, 'alpha-desc');
                    break;
                case 'sort-num-asc':
                    sortExplorerByColumn(contextMenuColumn, 'num-asc');
                    break;
                case 'sort-num-desc':
                    sortExplorerByColumn(contextMenuColumn, 'num-desc');
                    break;
                case 'sort-date-asc':
                    sortExplorerByColumn(contextMenuColumn, 'date-asc');
                    break;
                case 'sort-date-desc':
                    sortExplorerByColumn(contextMenuColumn, 'date-desc');
                    break;
                case 'filter-date-soon':
                    showDateFilterDialog(contextMenuColumn, 'dateInDays');
                    break;
                case 'filter-date-overdue':
                    showDateFilterDialog(contextMenuColumn, 'dateOverdueDays');
                    break;
                case 'hide-column':
                    toggleExplorerColumn(contextMenuColumn, false);
                    break;
                case 'delete-column':
                    deleteExplorerColumn(contextMenuColumn);
                    break;
                case 'insert-column-before':
                    insertExplorerColumn(contextMenuColumn, 'before');
                    break;
                case 'insert-column-after':
                    insertExplorerColumn(contextMenuColumn, 'after');
                    break;
            }
            
            hideColumnContextMenu();
        }
        
        /**
         * Zeigt Dialog für Datum-Filter mit Tage-Eingabe
         */
        function showDateFilterDialog(colIndex, filterType) {
            const columnName = explorerState.headers[colIndex] || `Spalte ${colIndex + 1}`;
            const title = filterType === 'dateInDays' 
                ? `⏰ Fällig in X Tagen - ${columnName}`
                : `⚠️ Überfällig seit X Tagen - ${columnName}`;
            const placeholder = filterType === 'dateInDays'
                ? 'Zeige Einträge die in den nächsten X Tagen fällig werden'
                : 'Zeige Einträge die seit X Tagen überfällig sind';
            
            const days = prompt(title + '\n\n' + placeholder + '\n\nAnzahl Tage eingeben:', '7');
            
            if (days !== null && !isNaN(parseInt(days))) {
                // Filter hinzufügen
                addDateFilterForColumn(colIndex, filterType, parseInt(days));
            }
        }
        
        /**
         * Fügt einen Datum-Filter für eine Spalte hinzu
         */
        function addDateFilterForColumn(colIndex, operator, days) {
            // Filter-Panel öffnen falls nicht sichtbar
            const filterPanel = document.getElementById('explorerFilterPanel');
            if (filterPanel && filterPanel.style.display === 'none') {
                filterPanel.style.display = 'block';
            }
            
            // Neuen Filter erstellen
            const template = document.getElementById('explorerFilterTemplate');
            const clone = template.content.cloneNode(true);
            const row = clone.querySelector('.explorer-filter-row');
            
            // Spalten-Dropdown befüllen und auswählen
            const colSelect = row.querySelector('.filter-column');
            colSelect.innerHTML = `<option value="">${t('selectColumn')}</option>` + 
                explorerState.headers.map((h, i) => `<option value="${i}">${escapeHtml(h || `Spalte ${i + 1}`)}</option>`).join('');
            colSelect.value = colIndex;
            
            // Operator setzen
            const operatorSelect = row.querySelector('.filter-operator');
            operatorSelect.value = operator;
            
            // Tage-Feld konfigurieren
            const valueInput = row.querySelector('.filter-value');
            const daysInput = row.querySelector('.filter-days');
            const dateFromInput = row.querySelector('.filter-date-from');
            const dateToInput = row.querySelector('.filter-date-to');
            valueInput.style.display = 'none';
            daysInput.style.display = 'block';
            daysInput.value = days;
            
            // Event-Listener
            operatorSelect.onchange = () => {
                const op = operatorSelect.value;
                const needsDays = op === 'dateInDays' || op === 'dateOverdueDays';
                const needsDateRange = op === 'dateBetween';
                const needsNoValue = ['dateToday', 'datePast', 'dateFuture', 'dateThisWeek', 'dateThisMonth', 'isEmpty', 'isNotEmpty'].includes(op);
                
                daysInput.style.display = needsDays ? 'block' : 'none';
                dateFromInput.style.display = needsDateRange ? 'block' : 'none';
                dateToInput.style.display = needsDateRange ? 'block' : 'none';
                valueInput.style.display = (needsNoValue || needsDays || needsDateRange) ? 'none' : 'block';
                
                if (needsNoValue) {
                    valueInput.value = '_no_value_required_';
                } else if (needsDays || needsDateRange) {
                    valueInput.value = '';
                }
                
                // Placeholders aktualisieren wenn dateBetween gewählt
                if (needsDateRange) {
                    updateDatePlaceholders(row);
                }
                
                updateFiltersFromDOM(true);
            };
            
            row.querySelector('.sync-filter').onclick = () => {
                syncFiltersToExcel();
            };
            row.querySelector('.remove-filter').onclick = () => {
                row.remove();
                updateFiltersFromDOM(true);
            };
            row.querySelector('.filter-column').onchange = () => {
                const op = operatorSelect.value;
                if (op === 'dateBetween') {
                    updateDatePlaceholders(row);
                }
                updateFiltersFromDOM(true);
            };
            const debouncedUpdateFilters = debounce(updateFiltersFromDOM, 300);
            // Datumsfelder: Nur GUI filtern, KEIN Excel-Sync per Debounce.
            // Excel-Sync erfolgt nur per Enter (siehe onkeydown unten).
            const debouncedUpdateFiltersGuiOnly = debounce(() => {
                updateFiltersFromDOM(); // ohne syncNow → geht in den Debounce-Pfad
                // Debounce-Timer für Excel-Sync abbrechen
                if (updateFiltersFromDOM._syncTimer) {
                    clearTimeout(updateFiltersFromDOM._syncTimer);
                    updateFiltersFromDOM._syncTimer = null;
                }
            }, 300);
            daysInput.oninput = debouncedUpdateFilters;
            dateFromInput.oninput = debouncedUpdateFiltersGuiOnly;
            dateToInput.oninput = debouncedUpdateFiltersGuiOnly;
            valueInput.oninput = debouncedUpdateFilters;
            
            // Enter-Handler für Datums-Textfelder
            dateFromInput.onkeydown = (e) => {
                if (e.key === 'Enter') { e.preventDefault(); syncFiltersToExcel(); }
            };
            dateToInput.onkeydown = (e) => {
                if (e.key === 'Enter') { e.preventDefault(); syncFiltersToExcel(); }
            };
            
            // Enter-Handler für Filter-Value
            valueInput.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    syncFiltersToExcel();
                }
            };
            daysInput.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    syncFiltersToExcel();
                }
            };
            
            document.getElementById('explorerFilters').appendChild(row);
            document.getElementById('btnClearExplorerFilters').disabled = false;
            
            // Filter anwenden
            updateFiltersFromDOM();
        }
        
        // Event-Listener für Kontextmenü
        document.addEventListener('DOMContentLoaded', () => {
            // Kontextmenü-Items
            document.querySelectorAll('#columnContextMenu .context-menu-item').forEach(item => {
                item.addEventListener('click', () => {
                    handleContextMenuAction(item.dataset.action);
                });
            });
            
            // Klick außerhalb schließt Menü
            document.addEventListener('click', (e) => {
                if (!e.target.closest('#columnContextMenu')) {
                    hideColumnContextMenu();
                }
            });
            
            // Escape schließt Menü
            document.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    hideColumnContextMenu();
                }
            });
            
            // Event-Delegation für Spalten-Kontextmenü (robuster als pro-Element Listener)
            const explorerTableHead = document.getElementById('explorerTableHead');
            if (explorerTableHead) {
                explorerTableHead.addEventListener('contextmenu', (e) => {
                    const th = e.target.closest('.sortable-header');
                    if (th && th.dataset.col !== undefined) {
                        e.preventDefault();
                        e.stopPropagation();
                        const colIndex = parseInt(th.dataset.col);
                        showColumnContextMenu(e, colIndex);
                    }
                });
            }
        });
        
        /**
         * Verschiebt eine Spalte per Drag & Drop
         */
        function moveExplorerColumn(fromColIndex, toColIndex) {
            // Initialisiere columnOrder wenn nötig - mit ALLEN Spalten (auch versteckte)
            if (explorerState.columnOrder.length === 0) {
                // Start mit allen Spalten in Original-Reihenfolge
                explorerState.columnOrder = explorerState.headers.map((_, i) => i);
            }
            
            // Finde die Positionen in columnOrder
            const fromPos = explorerState.columnOrder.indexOf(fromColIndex);
            const toPos = explorerState.columnOrder.indexOf(toColIndex);
            
            if (fromPos === -1 || toPos === -1) return;
            
            // LIVE SESSION: Verschiebe Spalte sofort in Excel
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // fromPos/toPos sind bereits die physischen Excel-Positionen (positions in columnOrder)
                showFloatingStatus('🔄 Excel wird synchronisiert...', 'info');
                liveSessionExecute('moveColumn', fromPos, toPos).then(() => {
                    console.log(`[LiveSession] Spalte ${fromColIndex} (Excel-Pos ${fromPos}) nach ${toColIndex} (Excel-Pos ${toPos}) verschoben`);
                    showFloatingStatus('✓ Spalte verschoben', 'success');
                }).catch((err) => {
                    console.error('[LiveSession] moveColumn error:', err);
                    showFloatingStatus('⚠️ Spaltenverschiebung fehlgeschlagen', 'warning');
                });
            }
            
            // Entferne die Spalte von der alten Position
            explorerState.columnOrder.splice(fromPos, 1);
            
            // Füge sie an der neuen Position ein
            explorerState.columnOrder.splice(toPos, 0, fromColIndex);
            
            // Markiere als strukturelle Änderung
            const existingMoves = explorerState.editedCells.get('_columnMoved') || [];
            existingMoves.push({ from: fromColIndex, to: toColIndex });
            explorerState.editedCells.set('_columnMoved', existingMoves);
            
            // Tabelle neu rendern
            renderExplorerTable();
        }
        
        /**
         * Setzt die Spaltenreihenfolge zurück
         */
        function resetExplorerColumnOrder() {
            explorerState.columnOrder = [];
            renderExplorerTable();
        }
        
        // Hilfsfunktion: Prüft ob eine Zelle Teil eines Merged-Bereichs ist
        // Gibt zurück: null (nicht merged), 'master' (Hauptzelle), 'hidden' (Teil eines Merges, ausblenden)
        function getMergedCellInfo(rowIndex, colIndex) {
            for (const merge of explorerState.mergedCells) {
                // rowIndex ist der Index im data-Array (0 = erste Datenzeile = Excel-Zeile 2)
                // merge.startRow ist 0-basierter Excel-Zeilen-Index (0 = Excel-Zeile 1)
                // Datenzeile 0 entspricht Excel-Zeile 2 → 0-basierter Index 1
                const excelRowIndex0Based = rowIndex + 1;
                
                if (excelRowIndex0Based >= merge.startRow && excelRowIndex0Based <= merge.endRow &&
                    colIndex >= merge.startCol && colIndex <= merge.endCol) {
                    // Diese Zelle ist Teil eines Merges
                    if (excelRowIndex0Based === merge.startRow && colIndex === merge.startCol) {
                        // Master-Zelle (oben links)
                        return {
                            type: 'master',
                            rowSpan: merge.rowSpan,
                            colSpan: merge.colSpan
                        };
                    } else {
                        // Versteckte Zelle
                        return { type: 'hidden' };
                    }
                }
            }
            return null; // Nicht Teil eines Merges
        }
        
        // Hilfsfunktion: Prüft ob eine Header-Zelle (Excel-Zeile 1) Teil eines Merged-Bereichs ist
        // excelRow0Based = 0 für Header (Excel-Zeile 1)
        function getHeaderMergedCellInfo(colIndex) {
            const excelRowIndex0Based = 0; // Header = Excel-Zeile 1 = 0-basiert Index 0
            for (const merge of explorerState.mergedCells) {
                if (excelRowIndex0Based >= merge.startRow && excelRowIndex0Based <= merge.endRow &&
                    colIndex >= merge.startCol && colIndex <= merge.endCol) {
                    // Diese Zelle ist Teil eines Merges
                    if (excelRowIndex0Based === merge.startRow && colIndex === merge.startCol) {
                        // Master-Zelle (oben links)
                        return {
                            type: 'master',
                            rowSpan: merge.rowSpan,
                            colSpan: merge.colSpan
                        };
                    } else {
                        // Versteckte Zelle
                        return { type: 'hidden' };
                    }
                }
            }
            return null; // Nicht Teil eines Merges
        }
        
        // ==================== Event Delegation für Explorer Table ====================
        // Einmalig registrierte Event-Listener auf dem tbody statt auf jeder einzelnen Zelle.
        // Reduziert die Anzahl der Listener von O(Zellen) auf O(1) pro Event-Typ.
        let _explorerDelegationSetup = false;
        let _explorerMouseUpHandler = null;
        // Virtual Scrolling
        let _virtualScrollSetup = false;
        let _virtualScrollRAF = null;
        
        // Hilfsfunktion: Komplette Zeile an Excel übertragen (für Cell-Edit + Dropdown)
        // Debounced per Row: Schnelle Edits an derselben Zeile werden zusammengefasst
        const _syncRowTimers = new Map();
        const _syncCellTimers = new Map();
        let _isPastingSyncLock = false;  // Unterdrückt syncCell während Paste
        // Batch-Buffer für Zell-Syncs: sammelt Änderungen und sendet als ein setCellsBatch
        const _pendingCellSyncs = new Map(); // key: "row-col" → {row, col, value, oldValue}
        let _cellSyncBatchTimer = null;
        
        // Hilfsfunktion: Reinen Text-Wert einer Zelle lesen (ohne Icons wie ƒ, 🔗, ⊞)
        // Zellen können <span class="formula-icon">, <span class="hyperlink-icon">,
        // <span class="merged-icon"> und <span class="cell-content"> enthalten.
        function getCellTextValue(td) {
            // Wenn cell-content span vorhanden → dessen textContent verwenden
            const contentSpan = td.querySelector('.cell-content');
            if (contentSpan) return contentSpan.textContent;
            // Sonst: alle Icon-Spans ignorieren und nur Text-Nodes + andere Spans lesen
            let text = '';
            for (const node of td.childNodes) {
                if (node.nodeType === Node.TEXT_NODE) {
                    text += node.textContent;
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                    // Icon-Spans überspringen (position: absolute, pointer-events: none)
                    if (node.classList && (node.classList.contains('formula-icon') || 
                        node.classList.contains('hyperlink-icon') || 
                        node.classList.contains('merged-icon'))) {
                        continue;
                    }
                    text += node.textContent;
                }
            }
            return text;
        }
        
        // Mappt einen logischen Spaltenindex auf die physische Excel-Spaltenposition.
        // Nach Spaltenverschiebungen weichen logische und physische Positionen voneinander ab.
        function _logicalToExcelCol(logicalColIndex) {
            if (explorerState.columnOrder.length === 0) return logicalColIndex;
            const physicalPos = explorerState.columnOrder.indexOf(logicalColIndex);
            return physicalPos >= 0 ? physicalPos : logicalColIndex;
        }

        // Ordnet ein Row-Array von logischer in physische Spaltenreihenfolge um
        function _reorderRowForExcel(rowData) {
            if (explorerState.columnOrder.length === 0) {
                return rowData.map(v => (v === null || v === undefined) ? '' : v);
            }
            return explorerState.columnOrder.map(logIdx => {
                const v = logIdx < rowData.length ? rowData[logIdx] : null;
                return (v === null || v === undefined) ? '' : v;
            });
        }

        // Mappt col-Felder in einem cells-Array für liveSessionSetCellsBatch
        function _mapCellsBatchCols(cells) {
            if (explorerState.columnOrder.length === 0) return cells;
            return cells.map(c => ({ ...c, col: _logicalToExcelCol(c.col) }));
        }

        function syncRowToExcel(rowIndex) {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) return;
            // Vorherigen Timer für diese Zeile abbrechen
            if (_syncRowTimers.has(rowIndex)) {
                clearTimeout(_syncRowTimers.get(rowIndex));
            }
            _syncRowTimers.set(rowIndex, setTimeout(() => {
                _syncRowTimers.delete(rowIndex);
                const rowData = explorerState.data[rowIndex];
                if (!rowData) return;
                const values = _reorderRowForExcel(rowData);
                console.log('[CellEdit] syncRowToExcel: Zeile', rowIndex, '→', values.length, 'Spalten');
                window.electronAPI.liveSessionSetRowValues(rowIndex, values)
                    .then(res => console.log('[CellEdit] Row-Sync Ergebnis:', JSON.stringify(res)))
                    .catch(err => console.error('[CellEdit] Row-Sync fehlgeschlagen:', err));
            }, 500));
        }
        // Sofort-Sync für blur/Enter (überspringt den Debounce-Timer)
        function syncRowToExcelImmediate(rowIndex) {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) return;
            // Laufenden Debounce-Timer abbrechen, da wir jetzt sofort senden
            if (_syncRowTimers.has(rowIndex)) {
                clearTimeout(_syncRowTimers.get(rowIndex));
                _syncRowTimers.delete(rowIndex);
            }
            const rowData = explorerState.data[rowIndex];
            if (!rowData) return;
            const values = _reorderRowForExcel(rowData);
            console.log('[CellEdit] syncRowToExcelImmediate: Zeile', rowIndex, '→', values.length, 'Spalten');
            window.electronAPI.liveSessionSetRowValues(rowIndex, values)
                .then(res => console.log('[CellEdit] Row-Sync Ergebnis:', JSON.stringify(res)))
                .catch(err => console.error('[CellEdit] Row-Sync fehlgeschlagen:', err));
        }
        
        // Prüft ob eine Zelle ein Bild (vm) oder Rich Text enthält — solche Zellen dürfen nicht
        // per setCellValue überschrieben werden, da das die IMAGE/DISPIMG-Formel bzw. Formatierung zerstört.
        function _isCellProtected(rowIndex, colIndex) {
            // cellVmMap und richTextCells verwenden "row+1"-Format (inkl. Header-Zeile)
            const styleKey = `${rowIndex + 1}-${colIndex}`;
            if (explorerState.cellVmMap && explorerState.cellVmMap[styleKey]) return true;
            if (explorerState.richTextCells && explorerState.richTextCells[styleKey]) return true;
            return false;
        }

        // Batch-Flush: Alle gesammelten Zell-Änderungen als ein setCellsBatch senden
        function _flushCellSyncBatch() {
            if (_pendingCellSyncs.size === 0) return;
            const cells = Array.from(_pendingCellSyncs.values());
            _pendingCellSyncs.clear();
            _cellSyncBatchTimer = null;
            console.log('[CellEdit] Batch-Flush:', cells.length, 'Zellen via setCellsBatch');
            window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(cells))
                .then(res => console.log('[CellEdit] Batch-Sync Ergebnis:', JSON.stringify(res)))
                .catch(err => console.error('[CellEdit] Batch-Sync fehlgeschlagen:', err));
        }

        // Einzelzellen-Sync (formatierungsschonend — sammelt in Batch-Buffer)
        function syncCellToExcel(rowIndex, colIndex) {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) return;
            if (_isPastingSyncLock) return;
            if (_isCellProtected(rowIndex, colIndex)) return;
            const cellKey = `${rowIndex}-${colIndex}`;
            // Per-Cell Debounce abbrechen
            if (_syncCellTimers.has(cellKey)) {
                clearTimeout(_syncCellTimers.get(cellKey));
            }
            _syncCellTimers.set(cellKey, setTimeout(() => {
                _syncCellTimers.delete(cellKey);
                const value = explorerState.data[rowIndex] ? explorerState.data[rowIndex][colIndex] : undefined;
                const sendValue = (value === null || value === undefined) ? '' : value;
                const excelCol = _logicalToExcelCol(colIndex);
                // In Batch-Buffer aufnehmen (Excel-Spaltenindex verwenden)
                _pendingCellSyncs.set(cellKey, { row: rowIndex, col: colIndex, value: sendValue });
                // Batch-Timer (re)starten
                if (_cellSyncBatchTimer) clearTimeout(_cellSyncBatchTimer);
                _cellSyncBatchTimer = setTimeout(_flushCellSyncBatch, 200);
            }, 500));
        }
        function syncCellToExcelImmediate(rowIndex, colIndex) {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) return;
            if (_isPastingSyncLock) return;
            if (_isCellProtected(rowIndex, colIndex)) return;
            const cellKey = `${rowIndex}-${colIndex}`;
            if (_syncCellTimers.has(cellKey)) {
                clearTimeout(_syncCellTimers.get(cellKey));
                _syncCellTimers.delete(cellKey);
            }
            // Auch Row-Timer für diese Zeile abbrechen (falls noch ausstehend)
            if (_syncRowTimers.has(rowIndex)) {
                clearTimeout(_syncRowTimers.get(rowIndex));
                _syncRowTimers.delete(rowIndex);
            }
            const value = explorerState.data[rowIndex] ? explorerState.data[rowIndex][colIndex] : undefined;
            const sendValue = (value === null || value === undefined) ? '' : value;
            const excelCol = _logicalToExcelCol(colIndex);
            // In Batch-Buffer aufnehmen und sofort flushen
            _pendingCellSyncs.set(cellKey, { row: rowIndex, col: colIndex, value: sendValue });
            if (_cellSyncBatchTimer) {
                clearTimeout(_cellSyncBatchTimer);
                _cellSyncBatchTimer = null;
            }
            _flushCellSyncBatch();
        }
        
        function setupExplorerTableDelegation() {
            if (_explorerDelegationSetup) return;
            _explorerDelegationSetup = true;
            
            const tbody = elements.explorerTableBody;
            
            // --- Checkbox change (Zeilen-Auswahl) ---
            tbody.addEventListener('change', function(e) {
                const checkbox = e.target.closest('.row-select-checkbox');
                if (checkbox) {
                    const rowIndex = parseInt(checkbox.dataset.rowIndex);
                    if (checkbox.checked) {
                        explorerState.selectedRows.add(rowIndex);
                        checkbox.closest('tr').classList.add('row-selected');
                        checkbox.parentElement.classList.add('selected');
                    } else {
                        explorerState.selectedRows.delete(rowIndex);
                        checkbox.closest('tr').classList.remove('row-selected');
                        checkbox.parentElement.classList.remove('selected');
                    }
                    updateRowMoveToolbar();
                    const selectAllCheckbox = document.getElementById('selectAllRows');
                    if (selectAllCheckbox) {
                        const pageRows = document.querySelectorAll('#explorerTableBody .row-select-checkbox');
                        const allChecked = Array.from(pageRows).every(cb => cb.checked);
                        selectAllCheckbox.checked = allChecked && pageRows.length > 0;
                    }
                    return;
                }
                
                // Dropdown-Zellen (Data Validation)
                const select = e.target.closest('.cell-dropdown');
                if (select) {
                    const rowIndex = parseInt(select.dataset.row);
                    const colIndex = parseInt(select.dataset.col);
                    const td = select.closest('td');
                    const original = td.dataset.original;
                    const lastValue = select.dataset.lastValue;
                    const current = select.value;
                    const cellKey = `${rowIndex}-${colIndex}`;
                    
                    if (original !== current) {
                        explorerState.editedCells.set(cellKey, current);
                        explorerState.data[rowIndex][colIndex] = current;
                        td.classList.add('edited');
                    } else {
                        explorerState.editedCells.delete(cellKey);
                        td.classList.remove('edited');
                    }
                    
                    if (lastValue !== current) {
                        pushExplorerUndo({
                            rowIndex, colIndex,
                            oldValue: lastValue, newValue: current,
                            originalValue: original
                        });
                        select.dataset.lastValue = current;
                        syncCellToExcelImmediate(rowIndex, colIndex);
                    }
                    updateExplorerEditStatus();
                }
            });
            
            // --- Input (Cell-Edit) ---
            tbody.addEventListener('input', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                
                const rowIndex = parseInt(td.dataset.row);
                const colIndex = parseInt(td.dataset.col);
                
                // Bild-/RichText-Zellen: Änderungen ignorieren
                if (_isCellProtected(rowIndex, colIndex)) return;
                const original = td.dataset.original;
                const current = getCellTextValue(td);
                const cellKey = `${rowIndex}-${colIndex}`;
                
                if (original !== current) {
                    explorerState.editedCells.set(cellKey, current);
                    td.classList.add('edited');
                } else {
                    explorerState.editedCells.delete(cellKey);
                    td.classList.remove('edited');
                }
                // explorerState.data IMMER aktualisieren (auch wenn == original)
                explorerState.data[rowIndex][colIndex] = current;
                
                updateExplorerEditStatus();
                
                // Live-Session Sync (debounced — nur geänderte Zelle)
                syncCellToExcel(rowIndex, colIndex);
            });
            
            // --- Blur (Undo speichern + Live-Sync) ---
            tbody.addEventListener('blur', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                
                td.parentElement.classList.remove('editing-row');
                
                const rowIndex = parseInt(td.dataset.row);
                const colIndex = parseInt(td.dataset.col);
                
                // Bild-/RichText-Zellen: NICHT überschreiben (zerstört IMAGE-Formel)
                if (_isCellProtected(rowIndex, colIndex)) return;
                
                const original = td.dataset.original;
                const lastValue = td.dataset.lastValue;
                const current = getCellTextValue(td);
                
                // Immer explorerState.data aktualisieren (falls input-Event
                // den letzten Tastendruck noch nicht erfasst hat)
                explorerState.data[rowIndex][colIndex] = current;
                
                if (lastValue !== current) {
                    pushExplorerUndo({
                        rowIndex, colIndex,
                        oldValue: lastValue, newValue: current,
                        originalValue: original
                    });
                    td.dataset.lastValue = current;
                }
                // Immer sofort syncen bei blur (cancelt auch pendenden Debounce)
                syncCellToExcelImmediate(rowIndex, colIndex);
            }, true); // useCapture für blur!
            
            // --- Focus ---
            tbody.addEventListener('focus', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                td.parentElement.classList.add('editing-row');
            }, true); // useCapture für focus!
            
            // --- Paste ---
            tbody.addEventListener('paste', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                e.preventDefault();
                const text = (e.clipboardData || window.clipboardData).getData('text/plain');
                document.execCommand('insertText', false, text);
            });
            
            // --- Keydown (Enter, Escape, Delete) ---
            tbody.addEventListener('keydown', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    
                    const rowIndex = parseInt(td.dataset.row);
                    const colIndex = parseInt(td.dataset.col);
                    const current = getCellTextValue(td);
                    const lastValue = td.dataset.lastValue;
                    
                    // explorerState.data immer aktualisieren
                    explorerState.data[rowIndex][colIndex] = current;
                    
                    if (lastValue !== current) {
                        pushExplorerUndo({
                            rowIndex, colIndex,
                            oldValue: lastValue, newValue: current,
                            originalValue: td.dataset.original
                        });
                        td.dataset.lastValue = current;
                        syncCellToExcelImmediate(rowIndex, colIndex);
                    }
                    
                    const nextRow = td.parentElement.nextElementSibling;
                    if (nextRow) {
                        const nextCell = nextRow.querySelector(`td[data-col="${colIndex}"]`);
                        if (nextCell) nextCell.focus();
                    }
                } else if (e.key === 'Escape') {
                    const rowIndex = parseInt(td.dataset.row);
                    const colIndex = parseInt(td.dataset.col);
                    const original = td.dataset.original;
                    td.textContent = original;
                    const cellKey = `${rowIndex}-${colIndex}`;
                    explorerState.editedCells.delete(cellKey);
                    explorerState.data[rowIndex][colIndex] = original;
                    td.classList.remove('edited');
                    td.dataset.lastValue = original;
                    
                    syncCellToExcelImmediate(rowIndex, colIndex);
                    td.blur();
                    updateExplorerEditStatus();
                } else if (e.key === 'Delete' || e.key === 'Backspace') {
                    if (explorerState.selectedCells.size > 1 && !document.activeElement.isContentEditable) {
                        e.preventDefault();
                        deleteSelectedCellsContent();
                    }
                }
            });
            
            // --- Mousedown (Zellen-Auswahl) ---
            tbody.addEventListener('mousedown', function(e) {
                const td = e.target.closest('td[contenteditable]');
                if (!td || e.button !== 0) return;
                
                const rowIndex = parseInt(td.dataset.row);
                const colIndex = parseInt(td.dataset.col);
                
                if (e.shiftKey && explorerState.selectionAnchor) {
                    e.preventDefault();
                    selectCellRange(rowIndex, colIndex);
                } else if (e.ctrlKey || e.metaKey) {
                    e.preventDefault();
                    toggleCellSelection(rowIndex, colIndex, true);
                } else {
                    if (!explorerState.selectedCells.has(`${rowIndex}-${colIndex}`)) {
                        explorerState.selectedCells.clear();
                    }
                    explorerState.selectionAnchor = { row: rowIndex, col: colIndex };
                    explorerState.isSelecting = true;
                }
            });
            
            // --- Mouseenter (Drag-Bereichsauswahl, throttled via rAF) ---
            let _dragRafPending = false;
            tbody.addEventListener('mouseenter', function(e) {
                if (!explorerState.isSelecting || e.buttons !== 1) return;
                const td = e.target.closest('td[contenteditable]');
                if (!td) return;
                if (_dragRafPending) return;
                _dragRafPending = true;
                requestAnimationFrame(() => {
                    _dragRafPending = false;
                    const rowIndex = parseInt(td.dataset.row);
                    const colIndex = parseInt(td.dataset.col);
                    selectCellRange(rowIndex, colIndex);
                });
            }, true);
            
            // --- Contextmenu (Rechtsklick) ---
            tbody.addEventListener('contextmenu', function(e) {
                // Zellen-Kontextmenü
                const td = e.target.closest('td[contenteditable]');
                if (td) {
                    const rowIndex = parseInt(td.dataset.row);
                    const colIndex = parseInt(td.dataset.col);
                    showCellContextMenu(e, rowIndex, colIndex);
                    return;
                }
                // Zeilen-Kontextmenü (Checkbox-Spalte)
                const checkboxCell = e.target.closest('.row-checkbox-cell');
                if (checkboxCell) {
                    const tr = checkboxCell.closest('tr');
                    const rowIndex = parseInt(tr.dataset.originalIndex);
                    if (!isNaN(rowIndex)) {
                        showRowContextMenu(e, rowIndex);
                    }
                }
            });
            
            // --- Click (Hyperlinks mit Ctrl/Cmd) ---
            tbody.addEventListener('click', function(e) {
                const span = e.target.closest('td.has-hyperlink .cell-content');
                if (span && (e.ctrlKey || e.metaKey)) {
                    e.preventDefault();
                    e.stopPropagation();
                    const td = span.closest('td');
                    const hyperlink = td.dataset.hyperlink;
                    if (hyperlink) window.electronAPI.openExternal(hyperlink);
                }
            });
            
            // --- Dblclick (Hyperlinks) ---
            tbody.addEventListener('dblclick', function(e) {
                const span = e.target.closest('td.has-hyperlink .cell-content');
                if (span) {
                    e.preventDefault();
                    e.stopPropagation();
                    const td = span.closest('td');
                    const hyperlink = td.dataset.hyperlink;
                    if (hyperlink) window.electronAPI.openExternal(hyperlink);
                }
            });
            
            // --- Mouseup (einmalig auf document) ---
            if (!_explorerMouseUpHandler) {
                _explorerMouseUpHandler = function() {
                    explorerState.isSelecting = false;
                };
                document.addEventListener('mouseup', _explorerMouseUpHandler);
            }
        }

        function renderExplorerTable(bodyOnly) {
            const btnToggleExcel = document.getElementById('btnToggleExcel');
            const btnDataJoin = document.getElementById('btnDataJoin');
            
            if (!explorerState.headers.length) {
                elements.explorerTableHead.innerHTML = '';
                elements.explorerTableBody.innerHTML = '';
                elements.explorerResultCount.textContent = t('noDataLoaded');
                document.getElementById('explorerPagination').style.display = 'none';
                if (btnToggleExcel) btnToggleExcel.disabled = true;
                if (btnDataJoin) btnDataJoin.disabled = true;
                // Drop-Zone anzeigen wenn keine Daten
                showExplorerDropZone(true);
                return;
            }
            
            if (!bodyOnly) {
                // Drop-Zone ausblenden wenn Daten vorhanden
                showExplorerDropZone(false);
                
                // Excel-Toggle-Button nur aktivieren, wenn Live-Session aktiv ist
                if (btnToggleExcel) {
                    btnToggleExcel.disabled = !(explorerState.liveSessionActive && explorerState.liveSessionReady);
                }
                // Data Join Button aktivieren
                if (btnDataJoin) btnDataJoin.disabled = false;
            }
            
            // Virtual Scrolling: Sichtbaren Bereich berechnen
            const container = document.getElementById('explorerTableContainer');
            const totalRows = explorerState.filteredData.length;
            const rowHeight = explorerState.virtualRowHeight;
            const buffer = explorerState.virtualBufferRows;
            
            // Bei vollem Re-Render: Virtual Scroll State zurücksetzen
            if (!bodyOnly) {
                explorerState.virtualVisibleStart = -1;
                explorerState.virtualVisibleEnd = -1;
            }
            
            const scrollTop = container.scrollTop;
            // Thead-Höhe abziehen (sticky header nimmt Platz ein)
            const theadHeight = elements.explorerTableHead ? elements.explorerTableHead.offsetHeight : 0;
            const viewportHeight = Math.max(container.clientHeight - theadHeight, 300); // Fallback 300px
            
            const firstVisible = Math.max(0, Math.floor((scrollTop - theadHeight) / rowHeight));
            const lastVisible = Math.max(0, Math.ceil((scrollTop - theadHeight + viewportHeight) / rowHeight));
            
            const startIndex = Math.max(0, firstVisible - buffer);
            const endIndex = Math.min(totalRows, lastVisible + buffer);
            
            // Bei bodyOnly (Scroll): Prüfen ob Re-Render nötig
            if (bodyOnly && explorerState.virtualVisibleStart >= 0) {
                const bufferThreshold = Math.floor(buffer / 3);
                if (firstVisible >= explorerState.virtualVisibleStart + bufferThreshold &&
                    lastVisible <= explorerState.virtualVisibleEnd - bufferThreshold) {
                    return; // Noch innerhalb des Puffers
                }
            }
            
            explorerState.virtualVisibleStart = startIndex;
            explorerState.virtualVisibleEnd = endIndex;
            
            const pageData = explorerState.filteredData.slice(startIndex, endIndex);
            
            // Header mit Sortierung und Drag & Drop
            // Verwende columnOrder wenn gesetzt, sonst visibleColumns
            const displayColumns = explorerState.columnOrder.length > 0 
                ? explorerState.columnOrder.filter(col => explorerState.visibleColumns.includes(col))
                : explorerState.visibleColumns;
            
            if (!bodyOnly) {
            // Prüfe ob alle sichtbaren Zeilen ausgewählt sind
            const allVisibleSelected = explorerState.filteredData.length > 0 && explorerState.filteredData.every(item => explorerState.selectedRows.has(item.originalIndex));
            
            // Erste Zeile: Excel-Spaltenbuchstaben (A, B, C, ...)
            let headerHtml = '<tr class="column-letter-row">';
            // Leere Zelle für Zeilennummer-Spalte
            headerHtml += `<th style="width: 40px; min-width: 40px; max-width: 40px; text-align: center; padding: 2px 4px; background: var(--bg-lighter); color: var(--text-muted); font-size: 10px; font-weight: normal; position: sticky; left: 0; z-index: 2; border-bottom: none;"></th>`;
            // Leere Zelle für Checkbox-Spalte
            headerHtml += `<th style="width: 40px; min-width: 40px; max-width: 40px; text-align: center; padding: 2px 4px; background: var(--bg-lighter); position: sticky; left: 40px; z-index: 2; border-bottom: none;"></th>`;
            displayColumns.forEach(colIndex => {
                // Prüfe ob dieser Spaltenbuchstabe Teil eines Header-Merges ist
                const headerMergeInfo = getHeaderMergedCellInfo(colIndex);
                if (headerMergeInfo && headerMergeInfo.type === 'hidden') {
                    // Diese Spalte gehört zu einem Header-Merge, wird nicht gerendert
                    return;
                }
                
                const colLetter = getColumnLetter(colIndex + 1);
                
                // Colspan für Merged Header berechnen
                let colspanAttr = '';
                if (headerMergeInfo && headerMergeInfo.type === 'master' && headerMergeInfo.colSpan > 1) {
                    let visibleColSpan = 0;
                    for (let i = 0; i < headerMergeInfo.colSpan; i++) {
                        if (displayColumns.includes(colIndex + i)) {
                            visibleColSpan++;
                        }
                    }
                    if (visibleColSpan > 1) {
                        // Zeige Bereich wie A-H
                        const endColLetter = getColumnLetter(colIndex + headerMergeInfo.colSpan);
                        colspanAttr = ` colspan="${visibleColSpan}"`;
                        headerHtml += `<th style="text-align: center; padding: 2px 4px; background: var(--bg-lighter); color: var(--text-muted); font-size: 10px; font-weight: normal; border-bottom: none;"${colspanAttr} title="Excel-Spalten ${colLetter}-${endColLetter}">${colLetter}-${endColLetter}</th>`;
                        return;
                    }
                }
                
                headerHtml += `<th style="text-align: center; padding: 2px 4px; background: var(--bg-lighter); color: var(--text-muted); font-size: 10px; font-weight: normal; border-bottom: none;" title="Excel-Spalte ${colLetter}">${colLetter}</th>`;
            });
            headerHtml += '</tr>';
            
            // Zweite Zeile: Spaltenüberschriften mit Sortierung
            headerHtml += '<tr>';
            // Zeilennummer-Spalte (fest, nicht verschiebbar) — zeigt Excel-Zeilennummern
            headerHtml += `<th style="width: 40px; min-width: 40px; max-width: 40px; text-align: center; padding: 2px 4px; background: var(--bg-lighter); color: var(--text-muted); font-size: 10px; font-weight: normal; position: sticky; left: 0; z-index: 2; border-bottom: none;" title="Excel-Zeilennummer">1</th>`;
            // Checkbox-Spalte für Zeilenauswahl
            headerHtml += `<th style="width: 40px; min-width: 40px; max-width: 40px; text-align: center; padding: 4px; position: sticky; left: 40px; z-index: 2; background: var(--bg-lighter);">
                <input type="checkbox" id="selectAllRows" ${allVisibleSelected ? 'checked' : ''} 
                    title="Alle sichtbaren Zeilen auswählen" style="cursor: pointer; width: 16px; height: 16px;">
            </th>`;
            
            displayColumns.forEach(colIndex => {
                // Prüfe ob dieser Header Teil eines Merged-Bereichs ist
                const headerMergeInfo = getHeaderMergedCellInfo(colIndex);
                if (headerMergeInfo && headerMergeInfo.type === 'hidden') {
                    // Diese Spalte gehört zu einem Header-Merge, wird nicht gerendert (colspan übernimmt)
                    return;
                }
                
                const headerText = escapeHtml(explorerState.headers[colIndex] || `Spalte ${colIndex + 1}`);
                const colLetter = getColumnLetter(colIndex + 1);
                let sortIcon = '';
                if (explorerState.sortColumn === colIndex) {
                    sortIcon = explorerState.sortDirection === 'asc' ? ' ▲' : ' ▼';
                }
                
                // Header-Styles anwenden (Zeile 0)
                const headerStyleKey = `0-${colIndex}`;
                const headerStyle = explorerState.cellStyles[headerStyleKey];
                let inlineStyle = '';
                if (headerStyle) {
                    const styles = [];
                    if (headerStyle.bold) styles.push('font-weight: bold');
                    if (headerStyle.italic) styles.push('font-style: italic');
                    if (headerStyle.fontColor) styles.push(`color: ${headerStyle.fontColor}`);
                    if (headerStyle.fill) styles.push(`background-color: ${headerStyle.fill}`);
                    if (headerStyle.textAlign) styles.push(`text-align: ${headerStyle.textAlign}`);
                    if (headerStyle.verticalAlign) {
                        const vAlign = headerStyle.verticalAlign === 'middle' ? 'middle' : headerStyle.verticalAlign === 'top' ? 'top' : 'bottom';
                        styles.push(`vertical-align: ${vAlign}`);
                    }
                    if (styles.length > 0) {
                        inlineStyle = ` style="${styles.join('; ')}"`;
                    }
                }
                
                // Merged Cell Attribute für Header
                let mergedAttrs = '';
                let mergedClass = '';
                let mergedIcon = '';
                if (headerMergeInfo && headerMergeInfo.type === 'master') {
                    mergedClass = ' merged-cell merged-cell-master';
                    // colspan nur für sichtbare Spalten berechnen
                    const colSpan = headerMergeInfo.colSpan;
                    // Berechne wie viele der Merge-Spalten tatsächlich in displayColumns sind
                    let visibleColSpan = 0;
                    for (let i = 0; i < colSpan; i++) {
                        if (displayColumns.includes(colIndex + i)) {
                            visibleColSpan++;
                        }
                    }
                    if (visibleColSpan > 1) {
                        mergedAttrs = ` colspan="${visibleColSpan}"`;
                    }
                    // Icon anzeigen wenn merged
                    if (headerMergeInfo.colSpan > 1) {
                        mergedIcon = `<span class="merged-icon" title="Verbundene Header-Zellen: 1×${headerMergeInfo.colSpan}">⊞</span>`;
                    }
                }
                
                // Merged Headers sind NICHT verschiebbar (zu komplex)
                const isMergedHeader = headerMergeInfo && headerMergeInfo.type === 'master' && headerMergeInfo.colSpan > 1;
                const draggableAttr = isMergedHeader ? 'draggable="false"' : 'draggable="true"';
                const titleText = isMergedHeader 
                    ? `Spalte ${colLetter} - Verbundene Spalten können nicht verschoben werden` 
                    : `Spalte ${colLetter} - Klicken zum Sortieren, Ziehen zum Verschieben`;
                const notDraggableStyle = isMergedHeader ? ' style="cursor: not-allowed;"' : '';
                
                headerHtml += `<th class="sortable-header${mergedClass}${isMergedHeader ? ' not-draggable' : ''}" data-col="${colIndex}" ${draggableAttr} title="${titleText}"${inlineStyle}${mergedAttrs}${notDraggableStyle}>${mergedIcon}${headerText}${sortIcon}</th>`;
            });
            headerHtml += '</tr>';
            elements.explorerTableHead.innerHTML = headerHtml;
            
            // Dynamisch top-Offset der zweiten Header-Zeile berechnen
            const firstHeaderRow = elements.explorerTableHead.querySelector('tr:first-child');
            if (firstHeaderRow) {
                const firstRowHeight = firstHeaderRow.offsetHeight;
                elements.explorerTableHead.querySelectorAll('tr:nth-child(2) th').forEach(th => {
                    th.style.top = firstRowHeight + 'px';
                });
            }
            // Select-All Checkbox Event-Listener
            const selectAllCheckbox = document.getElementById('selectAllRows');
            if (selectAllCheckbox) {
                selectAllCheckbox.addEventListener('change', function() {
                    // Alle gefilterten Zeilen auswählen/abwählen (nicht nur sichtbare)
                    explorerState.filteredData.forEach(item => {
                        if (this.checked) {
                            explorerState.selectedRows.add(item.originalIndex);
                        } else {
                            explorerState.selectedRows.delete(item.originalIndex);
                        }
                    });
                    updateRowSelectionUI();
                    updateRowMoveToolbar();
                });
            }
            
            // Sortier- und Drag-Event-Listener
            document.querySelectorAll('#explorerTableHead .sortable-header').forEach(th => {
                // Klick zum Sortieren
                th.addEventListener('click', (e) => {
                    // Nur sortieren wenn nicht gerade gedraggt wurde
                    if (!th.classList.contains('was-dragged')) {
                        const colIndex = parseInt(th.dataset.col);
                        sortExplorerByColumn(colIndex);
                    }
                    th.classList.remove('was-dragged');
                });
                
                // Rechtsklick für Kontextmenü wird jetzt via Event-Delegation auf thead behandelt
                
                // Drag Start
                th.addEventListener('dragstart', (e) => {
                    explorerState.draggedColumn = parseInt(th.dataset.col);
                    th.classList.add('dragging');
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', th.dataset.col);
                });
                
                // Drag Over - OPTIMIERT: nur preventDefault ohne DOM-Manipulation
                th.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                });
                
                // Drag Enter - einmalig statt kontinuierlich
                th.addEventListener('dragenter', (e) => {
                    e.preventDefault();
                    if (parseInt(th.dataset.col) !== explorerState.draggedColumn) {
                        th.classList.add('drag-over');
                    }
                });
                
                // Drag Leave
                th.addEventListener('dragleave', (e) => {
                    // Nur entfernen wenn wirklich verlassen (nicht bei Child-Elements)
                    if (e.target === th) {
                        th.classList.remove('drag-over');
                    }
                });
                
                // Drag End
                th.addEventListener('dragend', () => {
                    th.classList.remove('dragging');
                    document.querySelectorAll('.sortable-header').forEach(h => h.classList.remove('drag-over'));
                });
                
                // Drop
                th.addEventListener('drop', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    th.classList.remove('drag-over');
                    
                    const fromCol = explorerState.draggedColumn;
                    const toCol = parseInt(th.dataset.col);
                    
                    if (fromCol !== toCol && fromCol !== null) {
                        // Markiere dass gedraggt wurde (verhindert Sortierung beim Klick)
                        th.classList.add('was-dragged');
                        
                        // Aktualisiere die Spaltenreihenfolge
                        moveExplorerColumn(fromCol, toCol);
                    }
                    
                    explorerState.draggedColumn = null;
                });
            });
            } // end if (!bodyOnly)
            
            // Body - Virtual Scrolling (sichtbare Zeilen + Puffer)
            // Spacer oben für korrekte Scroll-Höhe
            const colCount = displayColumns.length + 2; // +2 für Zeilennummer + Checkbox
            const topSpacerHeight = startIndex * rowHeight;
            const bottomSpacerHeight = Math.max(0, (totalRows - endIndex) * rowHeight);
            
            let bodyHtml = '';
            if (topSpacerHeight > 0) {
                bodyHtml += `<tr class="virtual-spacer"><td colspan="${colCount}" style="height:${topSpacerHeight}px;padding:0;border:none;"></td></tr>`;
            }
            pageData.forEach(item => {
                const originalIndex = item.originalIndex;
                const row = item.row;
                const isSelected = explorerState.selectedRows.has(originalIndex);
                const highlightColor = explorerState.rowHighlights.get(originalIndex);
                const trClasses = [];
                if (isSelected) trClasses.push('row-selected');
                if (highlightColor) trClasses.push(`row-highlight-${highlightColor}`);
                const trClassAttr = trClasses.length > 0 ? ` class="${trClasses.join(' ')}"` : '';
                // Excel-Zeilennummer = originalIndex + 2 (Header = Zeile 1, Daten ab Zeile 2)
                const dataRowNumber = originalIndex + 2;
                bodyHtml += `<tr data-original-index="${originalIndex}"${trClassAttr}>`;
                // Zeilennummer-Zelle (fest, sticky)
                bodyHtml += `<td class="row-number-cell" style="width: 40px; min-width: 40px; max-width: 40px; text-align: center; padding: 4px; background: var(--bg-lighter); color: var(--text-muted); font-size: 11px; font-family: monospace; position: sticky; left: 0; z-index: 1; user-select: none;" title="Excel-Zeile ${dataRowNumber}">${dataRowNumber}</td>`;
                // Checkbox-Zelle für Zeilenauswahl
                bodyHtml += `<td class="row-checkbox-cell ${isSelected ? 'selected' : ''}" style="width: 40px; min-width: 40px; max-width: 40px; position: sticky; left: 40px; z-index: 1; background: var(--bg-secondary);">
                    <input type="checkbox" class="row-select-checkbox" data-row-index="${originalIndex}" 
                        ${isSelected ? 'checked' : ''}>
                </td>`;
                displayColumns.forEach(colIndex => {
                    // Prüfe ob diese Zelle Teil eines Merged-Bereichs ist
                    const mergeInfo = getMergedCellInfo(originalIndex, colIndex);
                    if (mergeInfo && mergeInfo.type === 'hidden') {
                        // Diese Zelle gehört zu einem Merge, wird nicht gerendert (colspan/rowspan übernimmt)
                        return;
                    }
                    
                    const cellValue = String(row[colIndex] ?? '');
                    const cellKey = `${originalIndex}-${colIndex}`;
                    const isEdited = explorerState.editedCells.has(cellKey);
                    const editedClass = isEdited ? ' edited' : '';
                    // Hole den echten Original-Wert aus originalData
                    const originalRow = explorerState.originalData[originalIndex];
                    const originalValue = originalRow ? String(originalRow[colIndex] ?? '') : cellValue;
                    
                    // Merged Cell Attribute
                    let mergedClass = '';
                    let mergedAttrs = '';
                    let mergedIcon = '';
                    let mergedMasterRow = originalIndex;  // Für Style-Lookup bei Merged Cells
                    let mergedMasterCol = colIndex;
                    if (mergeInfo && mergeInfo.type === 'master') {
                        mergedClass = ' merged-cell merged-cell-master';
                        // colspan nur wenn Spalte sichtbar
                        const visibleColSpan = Math.min(mergeInfo.colSpan, displayColumns.length - displayColumns.indexOf(colIndex));
                        if (visibleColSpan > 1) {
                            mergedAttrs += ` colspan="${visibleColSpan}"`;
                        }
                        // rowspan - berechne wie viele Zeilen im sichtbaren Bereich sind
                        if (mergeInfo.rowSpan > 1) {
                            // Berechne wie viele Zeilen des Merges im aktuellen Virtual-Scroll-Bereich sichtbar sind
                            const visibleRowSpan = Math.min(mergeInfo.rowSpan, endIndex - originalIndex);
                            if (visibleRowSpan > 1) {
                                mergedAttrs += ` rowspan="${visibleRowSpan}"`;
                            }
                        }
                        // Icon anzeigen wenn merged
                        if (mergeInfo.rowSpan > 1 || mergeInfo.colSpan > 1) {
                            mergedIcon = `<span class="merged-icon" title="Verbundene Zellen: ${mergeInfo.rowSpan}×${mergeInfo.colSpan}">⊞</span>`;
                        }
                    }
                    
                    // Cell Styles aus Excel
                    // Bei Merged Cells: Style der Master-Zelle verwenden
                    const cellStyleKey = `${mergedMasterRow + 1}-${mergedMasterCol}`; // +1 weil Styles inkl. Header-Zeile gespeichert sind
                    const cellStyle = explorerState.cellStyles[cellStyleKey];
                    let inlineStyle = '';
                    
                    // Formel prüfen (gleicher Key wie cellStyles)
                    const cellFormula = explorerState.cellFormulas[cellStyleKey];
                    const hasFormula = !!cellFormula;
                    const formulaTooltip = hasFormula ? ` title="Formel: =${escapeHtml(cellFormula)}"` : '';
                    const formulaClass = hasFormula ? ' has-formula' : '';
                    const formulaIcon = hasFormula ? '<span class="formula-icon" title="Diese Zelle enthält eine Formel">ƒ</span>' : '';
                    
                    // Hyperlink prüfen (gleicher Key wie cellStyles)
                    const cellHyperlink = explorerState.cellHyperlinks[cellStyleKey];
                    const hasHyperlink = !!cellHyperlink;
                    const hyperlinkClass = hasHyperlink ? ' has-hyperlink' : '';
                    const hyperlinkIcon = hasHyperlink ? '<span class="hyperlink-icon" title="Link öffnen: ' + escapeHtml(cellHyperlink) + '">🔗</span>' : '';
                    
                    // Rich Text prüfen (gleicher Key wie cellStyles)
                    const richTextFragments = explorerState.richTextCells[cellStyleKey];
                    const hasRichText = richTextFragments && richTextFragments.length > 0;
                    let richTextHtml = '';
                    if (hasRichText) {
                        // Generiere HTML für Rich Text Fragmente
                        richTextHtml = richTextFragments.map(fragment => {
                            const styles = [];
                            if (fragment.styles) {
                                if (fragment.styles.bold) styles.push('font-weight: bold');
                                if (fragment.styles.italic) styles.push('font-style: italic');
                                if (fragment.styles.underline) styles.push('text-decoration: underline');
                                if (fragment.styles.strikethrough) styles.push('text-decoration: line-through');
                                if (fragment.styles.subscript) styles.push('vertical-align: sub; font-size: smaller');
                                if (fragment.styles.superscript) styles.push('vertical-align: super; font-size: smaller');
                                // Unterstütze sowohl "fontColor" als auch "color"
                                const fragmentColor = fragment.styles.fontColor || fragment.styles.color;
                                if (fragmentColor) styles.push(`color: ${fragmentColor}`);
                                if (fragment.styles.fontSize) styles.push(`font-size: ${fragment.styles.fontSize}px`);
                                if (fragment.styles.fontName) styles.push(`font-family: '${fragment.styles.fontName}', sans-serif`);
                            }
                            const styleAttr = styles.length > 0 ? ` style="${styles.join('; ')}"` : '';
                            return `<span${styleAttr}>${escapeHtml(fragment.text)}</span>`;
                        }).join('');
                    }
                    const richTextClass = hasRichText ? ' has-rich-text' : '';
                    
                    // Echte Style-Anwendung
                    if (cellStyle) {
                        const styles = [];
                        if (cellStyle.bold) styles.push('font-weight: bold');
                        if (cellStyle.italic) styles.push('font-style: italic');
                        if (cellStyle.underline) styles.push('text-decoration: underline');
                        if (cellStyle.strikethrough) styles.push('text-decoration: line-through');
                        if (cellStyle.fontColor) styles.push(`color: ${cellStyle.fontColor}`);
                        if (cellStyle.fill) styles.push(`background-color: ${cellStyle.fill}`);
                        if (cellStyle.fontSize && cellStyle.fontSize !== 11) styles.push(`font-size: ${cellStyle.fontSize}px`);
                        if (cellStyle.fontName) styles.push(`font-family: '${cellStyle.fontName}', sans-serif`);
                        if (cellStyle.textAlign) styles.push(`text-align: ${cellStyle.textAlign}`);
                        if (cellStyle.verticalAlign) {
                            const vAlign = cellStyle.verticalAlign === 'middle' ? 'middle' : cellStyle.verticalAlign === 'top' ? 'top' : 'bottom';
                            styles.push(`vertical-align: ${vAlign}`);
                        }
                        if (cellStyle.wrapText) styles.push('white-space: pre-wrap');
                        if (styles.length > 0) {
                            inlineStyle = ` style="${styles.join('; ')}"`;
                        }
                    }
                    
                    // Bild-Platzhalter immer zentrieren (Fallback falls Style fehlt)
                    if (cellValue === '🖼️ Bild' && (!cellStyle || !cellStyle.textAlign)) {
                        const fallbackStyles = ['text-align: center', 'vertical-align: middle'];
                        if (inlineStyle) {
                            // Bestehende Styles ergänzen
                            const existing = inlineStyle.replace(/^ style="/, '').replace(/"$/, '');
                            inlineStyle = ` style="${existing}; ${fallbackStyles.join('; ')}"`;
                        } else {
                            inlineStyle = ` style="${fallbackStyles.join('; ')}"`;
                        }
                    }
                    
                    // Prüfe auf Data Validation (Dropdown-Liste)
                    const validation = explorerState.dataValidations[colIndex];
                    let validationValues = null;
                    
                    if (validation) {
                        if (validation.type === 'column') {
                            // Spaltenweite Validation
                            validationValues = validation.values;
                        } else if (validation.type === 'rows' && validation.rows[originalIndex + 1]) {
                            // Zeilenspezifische Validation (+1 weil Header-Zeile abgezogen wurde)
                            validationValues = validation.rows[originalIndex + 1].values;
                        }
                    }
                    
                    if (validationValues && validationValues.length > 0) {
                        // Zelle mit Dropdown rendern
                        const options = validationValues.map(v => {
                            const escaped = escapeHtml(v);
                            const selected = v === cellValue ? ' selected' : '';
                            return `<option value="${escaped}"${selected}>${escaped}</option>`;
                        }).join('');
                        // Füge leere Option hinzu wenn allowBlank
                        const emptyOption = validation.allowBlank !== false ? '<option value=""></option>' : '';
                        bodyHtml += `<td class="has-dropdown${editedClass}${formulaClass}${hyperlinkClass}${mergedClass}" data-row="${originalIndex}" data-col="${colIndex}" data-original="${escapeHtml(originalValue)}"${inlineStyle}${formulaTooltip}${hasHyperlink ? ` data-hyperlink="${escapeHtml(cellHyperlink)}"` : ''}${mergedAttrs}>
                            ${mergedIcon}${hyperlinkIcon}${formulaIcon}<select class="cell-dropdown" data-row="${originalIndex}" data-col="${colIndex}"${inlineStyle}>
                                ${emptyOption}${options}
                            </select>
                        </td>`;
                    } else {
                        // Normale editierbare Zelle
                        // Bestimme den anzuzeigenden Inhalt (Rich Text oder normaler Text)
                        const cellContent = hasRichText ? richTextHtml : escapeHtml(cellValue);
                        
                        if (hasHyperlink) {
                            // Zelle mit klickbarem Hyperlink
                            bodyHtml += `<td contenteditable="true" data-row="${originalIndex}" data-col="${colIndex}" data-original="${escapeHtml(originalValue)}" data-hyperlink="${escapeHtml(cellHyperlink)}" class="${editedClass}${formulaClass}${hyperlinkClass}${richTextClass}${mergedClass}"${inlineStyle}${formulaTooltip}${mergedAttrs}>${mergedIcon}${hyperlinkIcon}${formulaIcon}<span class="cell-content">${cellContent}</span></td>`;
                        } else if (hasRichText) {
                            // Zelle mit Rich Text (gemischte Formatierung)
                            bodyHtml += `<td contenteditable="true" data-row="${originalIndex}" data-col="${colIndex}" data-original="${escapeHtml(originalValue)}" class="${editedClass}${formulaClass}${richTextClass}${mergedClass}"${inlineStyle}${formulaTooltip}${mergedAttrs}>${mergedIcon}${formulaIcon}<span class="cell-content">${cellContent}</span></td>`;
                        } else {
                            bodyHtml += `<td contenteditable="true" data-row="${originalIndex}" data-col="${colIndex}" data-original="${escapeHtml(originalValue)}" class="${editedClass}${formulaClass}${mergedClass}"${inlineStyle}${formulaTooltip}${mergedAttrs}>${mergedIcon}${formulaIcon}${cellContent}</td>`;
                        }
                    }
                });
                bodyHtml += '</tr>';
            });
            // Spacer unten für korrekte Scroll-Höhe
            if (bottomSpacerHeight > 0) {
                bodyHtml += `<tr class="virtual-spacer"><td colspan="${colCount}" style="height:${bottomSpacerHeight}px;padding:0;border:none;"></td></tr>`;
            }
            elements.explorerTableBody.innerHTML = bodyHtml || '<tr><td colspan="100" style="text-align: center; padding: 20px;">Keine Daten gefunden</td></tr>';
            
            // Zeilen-Highlights werden jetzt direkt im HTML-Build gesetzt (keine nachträgliche DOM-Traversierung mehr)
            
            // Event Delegation einmalig registrieren (nicht bei jedem Render!)
            setupExplorerTableDelegation();
            
            // Event-Listener für Zeilen-Checkboxen (delegiert auf tbody)
            // Event-Listener für Hyperlinks (delegiert auf tbody)
            // Event-Listener für editierbare Zellen (delegiert auf tbody)
            // Event-Listener für Dropdown-Zellen (delegiert auf tbody)
            // → Alle über setupExplorerTableDelegation() registriert (einmalig)
            
            // lastValue für alle editierbaren Zellen initialisieren (ohne Icons)
            document.querySelectorAll('#explorerTableBody td[contenteditable]').forEach(td => {
                td.dataset.lastValue = getCellTextValue(td);
            });
            // lastValue für Dropdown-Zellen initialisieren
            document.querySelectorAll('#explorerTableBody .cell-dropdown').forEach(select => {
                select.dataset.lastValue = select.value;
            });
            
            // Ergebnis-Info
            if (totalRows > 0) {
                const editedCount = explorerState.editedCells.size;
                const editedInfo = editedCount > 0 ? ` (${editedCount} ${t('editedLabel')})` : '';
                elements.explorerResultCount.textContent = `${totalRows} ${t('rowsLabel')} (${t('totalLabel')}: ${explorerState.data.length})${editedInfo}`;
            } else {
                elements.explorerResultCount.textContent = `0 ${t('ofRows')} ${explorerState.data.length} ${t('rowsLabel')}`;
            }
            
            // Pagination UI ausblenden (Virtual Scrolling ersetzt Pagination)
            document.getElementById('explorerPagination').style.display = 'none';
            
            // Virtual Scroll Listener einmalig registrieren
            if (!_virtualScrollSetup) {
                _virtualScrollSetup = true;
                container.addEventListener('scroll', _onVirtualScroll);
                // Bei Größenänderung des Containers neu rendern
                new ResizeObserver(() => _onVirtualScroll()).observe(container);
            }
            
            // Row Move Toolbar aktualisieren
            updateRowMoveToolbar();
        }
        
        // Virtual Scroll Handler: wird bei Scroll auf dem Table-Container aufgerufen
        function _onVirtualScroll() {
            if (_virtualScrollRAF) return;
            _virtualScrollRAF = requestAnimationFrame(() => {
                _virtualScrollRAF = null;
                renderExplorerTable(true); // bodyOnly
            });
        }
        
        // Scrollt den Table-Container so, dass eine bestimmte Zeile sichtbar ist
        function scrollToVirtualRow(rowIndex) {
            const container = document.getElementById('explorerTableContainer');
            const theadHeight = elements.explorerTableHead ? elements.explorerTableHead.offsetHeight : 0;
            const targetTop = theadHeight + rowIndex * explorerState.virtualRowHeight;
            const viewportHeight = container.clientHeight;
            // Zeile in die Mitte des Viewports scrollen
            container.scrollTop = Math.max(0, targetTop - viewportHeight / 2);
        }
        
        // Aktualisiert die Row Move Toolbar basierend auf der Auswahl
        function updateRowMoveToolbar() {
            const toolbar = document.getElementById('rowMoveToolbar');
            const countSpan = document.getElementById('selectedRowCount');
            const selectedCount = explorerState.selectedRows.size;
            
            if (toolbar && countSpan) {
                countSpan.textContent = selectedCount;
                toolbar.style.display = selectedCount > 0 ? 'flex' : 'none';
            }
        }
        
        // Aktualisiert die visuelle Auswahl der Zeilen
        function updateRowSelectionUI() {
            document.querySelectorAll('.row-select-checkbox').forEach(checkbox => {
                const rowIndex = parseInt(checkbox.dataset.rowIndex);
                const isSelected = explorerState.selectedRows.has(rowIndex);
                checkbox.checked = isSelected;
                const tr = checkbox.closest('tr');
                const td = checkbox.parentElement;
                if (tr) {
                    if (isSelected) {
                        tr.classList.add('row-selected');
                        td.classList.add('selected');
                    } else {
                        tr.classList.remove('row-selected');
                        td.classList.remove('selected');
                    }
                }
            });
        }
        
        // Führt das Verschieben der ausgewählten Zeilen aus
        async function executeRowMove() {
            const selectedCount = explorerState.selectedRows.size;
            if (selectedCount === 0) {
                showNotification('Keine Zeilen ausgewählt', 'warning');
                return;
            }
            
            const movePosition = document.getElementById('movePosition').value; // 'before' oder 'after'
            const targetRowInput = document.getElementById('moveTargetRow');
            const targetRow = parseInt(targetRowInput.value);
            
            const maxExcelRow = explorerState.data.length + 1;
            if (isNaN(targetRow) || targetRow < 2 || targetRow > maxExcelRow) {
                showNotification(`Ungültige Zielzeile. Bitte eine Zahl zwischen 2 und ${maxExcelRow} eingeben.`, 'warning');
                return;
            }
            
            // Zielindex (0-basiert): Excel-Zeile 2 = Index 0
            let targetIndex = targetRow - 2;
            
            // Sortiere die ausgewählten Indizes
            const selectedIndices = Array.from(explorerState.selectedRows).sort((a, b) => a - b);
            
            // Prüfe ob Zielzeile in der Auswahl enthalten ist
            if (explorerState.selectedRows.has(targetIndex)) {
                showNotification('Die Zielzeile kann nicht in der Auswahl enthalten sein.', 'warning');
                return;
            }
            
            // LIVE SESSION: Verschiebe Zeilen sofort in Excel
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Berechne Zielposition (gleiche Logik wie die Daten-Manipulation unten)
                let liveNewTarget = targetIndex;
                selectedIndices.forEach(idx => {
                    if (idx < targetIndex) liveNewTarget--;
                });
                if (movePosition === 'after') liveNewTarget++;
                
                const bundleCount = selectedIndices.length;
                const movingUp = selectedIndices[0] > liveNewTarget;
                
                if (movingUp) {
                    // Aufsteigend verarbeiten (niedrigster Index zuerst)
                    for (let i = 0; i < bundleCount; i++) {
                        const fromIdx = selectedIndices[i];
                        const toIdx = liveNewTarget + i;
                        if (fromIdx !== toIdx) {
                            await liveSessionExecute('moveRow', fromIdx, toIdx);
                        }
                    }
                } else {
                    // Absteigend verarbeiten (höchster Index zuerst)
                    for (let i = bundleCount - 1; i >= 0; i--) {
                        const fromIdx = selectedIndices[i];
                        const toIdx = liveNewTarget + i;
                        if (fromIdx !== toIdx) {
                            await liveSessionExecute('moveRow', fromIdx, toIdx);
                        }
                    }
                }
                console.log(`[LiveSession] ${bundleCount} Zeile(n) verschoben (target=${liveNewTarget})`);
            }
            
            // Speichere den aktuellen Zustand für Undo - EFFIZIENT: nur betroffene Zeilen
            // Wir speichern nur die Indizes und lassen den State unverändert im Undo-System
            const previousDataOrder = explorerState.data.map((_, i) => i);
            const previousCellStyles = { ...explorerState.cellStyles }; // Flache Kopie reicht
            const previousRichTextCells = { ...explorerState.richTextCells }; // Flache Kopie reicht
            
            // WICHTIG: Style-Keys haben Format "rowIndex-colIndex" wobei rowIndex 1 = erste Datenzeile
            // (rowIndex 0 = Header)
            // Daten-Index 0 entspricht Style-Index 1
            
            // OPTIMIERUNG: Gruppiere alle Metadaten einmalig nach Zeile (O(n) statt O(n²))
            // WICHTIG: Tiefe Kopien erstellen um Referenzprobleme zu vermeiden!
            const stylesByRow = {};
            const richTextByRow = {};
            const formulasByRow = {};
            const hyperlinksByRow = {};
            
            for (const [key, value] of Object.entries(explorerState.cellStyles || {})) {
                const [keyRow, keyCol] = key.split('-');
                const rowIdx = parseInt(keyRow);
                if (rowIdx > 0) { // Nur Datenzeilen, nicht Header
                    if (!stylesByRow[rowIdx]) stylesByRow[rowIdx] = {};
                    // Direkte Referenz - wir verschieben nur, keine Modifikation
                    stylesByRow[rowIdx][keyCol] = value;
                }
            }
            
            for (const [key, value] of Object.entries(explorerState.richTextCells || {})) {
                const [keyRow, keyCol] = key.split('-');
                const rowIdx = parseInt(keyRow);
                if (rowIdx > 0) {
                    if (!richTextByRow[rowIdx]) richTextByRow[rowIdx] = {};
                    // Direkte Referenz - wir verschieben nur, keine Modifikation
                    richTextByRow[rowIdx][keyCol] = value;
                }
            }
            
            for (const [key, value] of Object.entries(explorerState.cellFormulas || {})) {
                const [keyRow, keyCol] = key.split('-');
                const rowIdx = parseInt(keyRow);
                if (rowIdx > 0) {
                    if (!formulasByRow[rowIdx]) formulasByRow[rowIdx] = {};
                    // Formeln sind Strings, brauchen keine tiefe Kopie
                    formulasByRow[rowIdx][keyCol] = value;
                }
            }
            
            for (const [key, value] of Object.entries(explorerState.cellHyperlinks || {})) {
                const [keyRow, keyCol] = key.split('-');
                const rowIdx = parseInt(keyRow);
                if (rowIdx > 0) {
                    if (!hyperlinksByRow[rowIdx]) hyperlinksByRow[rowIdx] = {};
                    // Hyperlinks sind Strings, brauchen keine tiefe Kopie
                    hyperlinksByRow[rowIdx][keyCol] = value;
                }
            }
            
            // EditedCells nach Zeile gruppieren (0-basierte Daten-Indizes, OHNE Header-Offset).
            // WICHTIG: Spezielle Marker-Keys ("_rowInserted" etc.) bleiben unangetastet
            // und werden NICHT verschoben — sie werden separat behandelt.
            const editedCellsByRow = {};
            const preservedMarkerEdits = new Map();
            for (const [key, value] of explorerState.editedCells.entries()) {
                if (typeof key === 'string' && key.startsWith('_')) {
                    preservedMarkerEdits.set(key, value);
                    continue;
                }
                const [keyRow, keyCol] = key.split('-');
                const rowIdx = parseInt(keyRow);
                if (Number.isInteger(rowIdx) && rowIdx >= 0) {
                    if (!editedCellsByRow[rowIdx]) editedCellsByRow[rowIdx] = {};
                    editedCellsByRow[rowIdx][keyCol] = value;
                }
            }
            
            // Sammle Zeilen-Bundles mit allen Metadaten (jetzt O(n))
            // EFFIZIENT: Direkte Referenzen, keine Kopien nötig da wir nur verschieben
            const totalRows = explorerState.data.length;
            const rowBundles = [];
            
            // Aktuelles rowMapping holen (oder initialisieren)
            const currentMapping = explorerState.rowMapping || explorerState.data.map((_, i) => i);
            
            for (let i = 0; i < totalRows; i++) {
                const styleRowIdx = i + 1; // Style-Key verwendet +1 (Header ist 0)
                const bundle = {
                    dataRef: explorerState.data[i],       // Direkte Referenz
                    originalRef: explorerState.originalData[i], // Direkte Referenz
                    originalExcelRow: currentMapping[i],  // Original Excel-Zeilen-Index (0-basiert in Daten)
                    styles: stylesByRow[styleRowIdx] || {},
                    richText: richTextByRow[styleRowIdx] || {},
                    formulas: formulasByRow[styleRowIdx] || {},
                    hyperlinks: hyperlinksByRow[styleRowIdx] || {},
                    editedCells: editedCellsByRow[i] || {},  // 0-basierte Daten-Zeile (kein Header-Offset)
                    isHidden: explorerState.hiddenRows?.has(i) || false,
                    highlight: explorerState.rowHighlights.get(i) || null
                };
                
                rowBundles.push(bundle);
            }
            
            // Extrahiere die zu verschiebenden Bundles
            const bundlesToMove = selectedIndices.map(idx => rowBundles[idx]);
            
            // Entferne sie aus der Liste (von hinten nach vorne)
            const sortedDescending = [...selectedIndices].sort((a, b) => b - a);
            sortedDescending.forEach(idx => {
                rowBundles.splice(idx, 1);
            });
            
            // Berechne den neuen Zielindex
            let newTargetIndex = targetIndex;
            selectedIndices.forEach(idx => {
                if (idx < targetIndex) {
                    newTargetIndex--;
                }
            });
            
            if (movePosition === 'after') {
                newTargetIndex++;
            }
            
            // Füge die Bundles an der neuen Position ein
            bundlesToMove.forEach((bundle, i) => {
                rowBundles.splice(newTargetIndex + i, 0, bundle);
            });
            
            // Rekonstruiere alle States aus den Bundles
            explorerState.data = [];
            explorerState.originalData = [];
            explorerState.cellStyles = {};
            explorerState.richTextCells = {};
            explorerState.cellFormulas = {};
            explorerState.cellHyperlinks = {};
            explorerState.hiddenRows = new Set();
            explorerState.rowHighlights = new Map();
            // EditedCells komplett leeren — Marker-Keys ("_rowsReordered" etc.) gleich wieder setzen,
            // damit Per-Row-Edits unten mit korrekten neuen Keys eingefügt werden.
            explorerState.editedCells = new Map();
            preservedMarkerEdits.forEach((value, key) => {
                explorerState.editedCells.set(key, value);
            });
            
            // Header-Styles behalten (Key "0-x")
            for (const [key, value] of Object.entries(previousCellStyles)) {
                const [keyRow] = key.split('-');
                if (parseInt(keyRow) === 0) {
                    explorerState.cellStyles[key] = value;
                }
            }
            
            rowBundles.forEach((bundle, newRowIdx) => {
                const newStyleRowIdx = newRowIdx + 1;
                
                explorerState.data.push(bundle.dataRef);  // Direkte Referenz
                explorerState.originalData.push(bundle.originalRef);  // Direkte Referenz
                
                if (bundle.isHidden) {
                    explorerState.hiddenRows.add(newRowIdx);
                }
                
                // Highlight mit neuem Index
                if (bundle.highlight) {
                    explorerState.rowHighlights.set(newRowIdx, bundle.highlight);
                }
                
                // Styles mit neuem Key
                for (const [col, style] of Object.entries(bundle.styles)) {
                    explorerState.cellStyles[`${newStyleRowIdx}-${col}`] = style;
                }
                
                // RichText mit neuem Key
                for (const [col, rt] of Object.entries(bundle.richText)) {
                    explorerState.richTextCells[`${newStyleRowIdx}-${col}`] = rt;
                }
                
                // Formeln mit neuem Key
                for (const [col, formula] of Object.entries(bundle.formulas)) {
                    explorerState.cellFormulas[`${newStyleRowIdx}-${col}`] = formula;
                }
                
                // Hyperlinks mit neuem Key
                for (const [col, link] of Object.entries(bundle.hyperlinks)) {
                    explorerState.cellHyperlinks[`${newStyleRowIdx}-${col}`] = link;
                }
                
                // EditedCells mit neuem 0-basierten Daten-Index als Key (kein Header-Offset)
                for (const [col, val] of Object.entries(bundle.editedCells)) {
                    explorerState.editedCells.set(`${newRowIdx}-${col}`, val);
                }
            });
            
            // WICHTIG: rowMapping aktualisieren - trackt welche Original-Excel-Zeile an welcher neuen Position ist
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Live Session: Daten und Excel sind synchron, kein Mapping nötig
                explorerState.rowMapping = null;
            } else {
                explorerState.rowMapping = rowBundles.map(bundle => bundle.originalExcelRow);
            }
            
            // Speichere Undo-Aktion (keine tiefen Kopien mehr - Undo wird vereinfacht)
            pushExplorerUndo({
                type: 'moveRows',
                previousDataOrder: previousDataOrder,  // Nur die Reihenfolge speichern
                previousCellStyles: previousCellStyles,
                previousRichTextCells: previousRichTextCells,
                movedIndices: selectedIndices,
                targetRow: targetRow,
                movePosition: movePosition
            });
            
            // Auswahl zurücksetzen
            explorerState.selectedRows.clear();
            
            // Tracke alle betroffenen Zeilen (alt und neu) für Style-Reset beim Speichern
            // WICHTIG: affectedRows verwendet STYLE-Indizes (rowIndex + 1 weil Header = 0)
            if (!explorerState.affectedRows) {
                explorerState.affectedRows = new Set();
            }
            // Alle Zeilen zwischen min und max der Bewegung sind betroffen (Style-Indizes!)
            const minDataIdx = Math.min(...selectedIndices, newTargetIndex);
            const maxDataIdx = Math.max(...selectedIndices, newTargetIndex + bundlesToMove.length - 1);
            for (let i = minDataIdx; i <= maxDataIdx; i++) {
                explorerState.affectedRows.add(i + 1); // +1 für Style-Index
            }
            
            // FilteredData neu erstellen (mit allen Filtern: hidden rows, Suche, Spaltenfilter, Sortierung)
            filterExplorerData();
            
            // Markierung, dass Zeilen neu angeordnet wurden (strukturelle Änderung)
            explorerState.editedCells.set('_rowsReordered', true);
            
            // Markiere als geändert
            explorerState.hasUnsavedChanges = true;
            
            showNotification(`${selectedCount} Zeile(n) ${movePosition === 'before' ? 'vor' : 'nach'} Excel-Zeile ${targetRow} verschoben`, 'success');
            
            // Eingabefeld leeren
            targetRowInput.value = '';
        }
        
        // Löscht mehrere ausgewählte Zeilen
        async function deleteSelectedRows() {
            const selectedCount = explorerState.selectedRows.size;
            if (selectedCount === 0) {
                showNotification('Keine Zeilen ausgewählt', 'warning');
                return;
            }
            
            // Bestätigung anfordern
            const deleteLabel = currentLanguage === 'en' ? 'Delete' : 'Löschen';
            const cancelLabel = currentLanguage === 'en' ? 'Cancel' : 'Abbrechen';
            const titleText = currentLanguage === 'en' 
                ? `Delete ${selectedCount} Row(s)?` 
                : `${selectedCount} Zeile(n) löschen?`;
            const confirmText = currentLanguage === 'en'
                ? `Do you really want to delete the ${selectedCount} selected row(s)?\n\nThis action cannot be undone.`
                : `Möchten Sie die ${selectedCount} ausgewählten Zeile(n) wirklich löschen?\n\nDiese Aktion kann nicht rückgängig gemacht werden.`;
            
            const confirmed = await showConfirmDialog(
                titleText,
                confirmText,
                deleteLabel,
                cancelLabel
            );
            
            if (!confirmed) return;
            
            // LIVE SESSION: Lösche Zeilen sofort in Excel (von hinten nach vorne)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Übersetze data-array-Indizes in physische Excel-Positionen
                const indicesWithExcelPos = Array.from(explorerState.selectedRows).map(idx => ({
                    dataIndex: idx,
                    excelPos: getExcelRowPosition(idx)
                }));
                // Sortiere nach Excel-Position absteigend (von hinten nach vorne löschen)
                indicesWithExcelPos.sort((a, b) => b.excelPos - a.excelPos);
                for (const item of indicesWithExcelPos) {
                    await liveSessionExecute('deleteRow', item.excelPos);
                }
                console.log(`[LiveSession] ${indicesWithExcelPos.length} Zeilen in Excel gelöscht`);
            }
            
            // Speichere den aktuellen Zustand für Undo (nur die gelöschten Zeilen, nicht den gesamten Datensatz)
            const selectedIndices = Array.from(explorerState.selectedRows).sort((a, b) => b - a);
            const deletedRows = selectedIndices.map(idx => ({
                index: idx,
                data: explorerState.data[idx] ? [...explorerState.data[idx]] : [],
                originalData: explorerState.originalData[idx] ? [...explorerState.originalData[idx]] : [],
                highlight: explorerState.rowHighlights.get(idx),
                editedCells: []
            }));
            // Editierte Zellen der gelöschten Zeilen sichern
            const deletedSet = new Set(selectedIndices);
            for (const [key, value] of explorerState.editedCells) {
                const dashIdx = key.indexOf('-');
                if (dashIdx > 0) {
                    const rowIdx = parseInt(key.substring(0, dashIdx));
                    if (deletedSet.has(rowIdx)) {
                        const row = deletedRows.find(r => r.index === rowIdx);
                        if (row) row.editedCells.push({ key, value });
                    }
                }
            }
            const previousHighlights = new Map(explorerState.rowHighlights);
            const previousEditedCells = new Map(explorerState.editedCells);
            
            // Lösche die Zeilen von hinten nach vorne (selectedIndices already sorted desc)
            selectedIndices.forEach(idx => {
                explorerState.data.splice(idx, 1);
                explorerState.originalData.splice(idx, 1);
            });
            
            // Highlights anpassen
            const newHighlights = new Map();
            explorerState.rowHighlights.forEach((color, idx) => {
                if (deletedSet.has(idx)) return; // Gelöschte Zeile überspringen
                // Berechne wie viele gelöschte Zeilen vor dieser lagen
                let offset = 0;
                selectedIndices.forEach(delIdx => {
                    if (delIdx < idx) offset++;
                });
                newHighlights.set(idx - offset, color);
            });
            explorerState.rowHighlights = newHighlights;
            
            // cellStyles anpassen (Key ist "styleRowIdx-colIdx" wobei styleRowIdx = dataRowIdx + 1)
            const newCellStyles = {};
            for (const [key, value] of Object.entries(explorerState.cellStyles || {})) {
                const [rowStr, colStr] = key.split('-');
                const styleRow = parseInt(rowStr);
                const col = parseInt(colStr);
                if (styleRow === 0) {
                    // Header behalten
                    newCellStyles[key] = value;
                    continue;
                }
                const dataRow = styleRow - 1; // styleRow = dataRow + 1
                if (deletedSet.has(dataRow)) continue; // Gelöschte Zeile überspringen
                let offset = 0;
                selectedIndices.forEach(delIdx => {
                    if (delIdx < dataRow) offset++;
                });
                newCellStyles[`${styleRow - offset}-${col}`] = value;
            }
            explorerState.cellStyles = newCellStyles;
            
            // cellHyperlinks anpassen
            const newCellHyperlinks = {};
            for (const [key, value] of Object.entries(explorerState.cellHyperlinks || {})) {
                const [rowStr, colStr] = key.split('-');
                const styleRow = parseInt(rowStr);
                const col = parseInt(colStr);
                if (styleRow === 0) {
                    newCellHyperlinks[key] = value;
                    continue;
                }
                const dataRow = styleRow - 1;
                if (deletedSet.has(dataRow)) continue;
                let offset = 0;
                selectedIndices.forEach(delIdx => {
                    if (delIdx < dataRow) offset++;
                });
                newCellHyperlinks[`${styleRow - offset}-${col}`] = value;
            }
            explorerState.cellHyperlinks = newCellHyperlinks;
            
            // richTextCells anpassen
            const newRichTextCells = {};
            for (const [key, value] of Object.entries(explorerState.richTextCells || {})) {
                const [rowStr, colStr] = key.split('-');
                const styleRow = parseInt(rowStr);
                const col = parseInt(colStr);
                if (styleRow === 0) {
                    newRichTextCells[key] = value;
                    continue;
                }
                const dataRow = styleRow - 1;
                if (deletedSet.has(dataRow)) continue;
                let offset = 0;
                selectedIndices.forEach(delIdx => {
                    if (delIdx < dataRow) offset++;
                });
                newRichTextCells[`${styleRow - offset}-${col}`] = value;
            }
            explorerState.richTextCells = newRichTextCells;
            
            // cellFormulas anpassen
            const newCellFormulas = {};
            for (const [key, value] of Object.entries(explorerState.cellFormulas || {})) {
                const [rowStr, colStr] = key.split('-');
                const styleRow = parseInt(rowStr);
                const col = parseInt(colStr);
                if (styleRow === 0) {
                    newCellFormulas[key] = value;
                    continue;
                }
                const dataRow = styleRow - 1;
                if (deletedSet.has(dataRow)) continue;
                let offset = 0;
                selectedIndices.forEach(delIdx => {
                    if (delIdx < dataRow) offset++;
                });
                newCellFormulas[`${styleRow - offset}-${col}`] = value;
            }
            explorerState.cellFormulas = newCellFormulas;
            
            // EditedCells anpassen
            const newEditedCells = new Map();
            explorerState.editedCells.forEach((value, key) => {
                if (key.startsWith('_')) {
                    newEditedCells.set(key, value);
                    return;
                }
                const [rowStr, colStr] = key.split('-');
                const row = parseInt(rowStr);
                const col = parseInt(colStr);
                if (deletedSet.has(row)) return; // Gelöschte Zeile überspringen
                // Berechne wie viele gelöschte Zeilen vor dieser lagen
                let offset = 0;
                selectedIndices.forEach(delIdx => {
                    if (delIdx < row) offset++;
                });
                newEditedCells.set(`${row - offset}-${col}`, value);
            });
            explorerState.editedCells = newEditedCells;
            
            // Markierung, dass etwas gelöscht wurde
            explorerState.editedCells.set('_rowDeleted', true);
            
            // WICHTIG: Erfasse Original-Indizes BEVOR wir rowMapping ändern (analog zu _columnDeleted)
            const existingDeleted = explorerState.editedCells.get('_deletedRowIndices');
            let deletedOriginalIndices = [];
            if (existingDeleted && Array.isArray(existingDeleted.originalIndices)) {
                deletedOriginalIndices = existingDeleted.originalIndices.slice();
            }
            
            // Erfasse Original-Indizes aller ausgewählten Zeilen
            const currentMapping = explorerState.rowMapping && explorerState.rowMapping.length > 0 
                ? explorerState.rowMapping 
                : explorerState.data.map((_, i) => i); // Ohne Mapping: Index = Original
            // WICHTIG: Wir iterieren in aufsteigender Reihenfolge für korrektes Mapping
            const sortedAsc = [...selectedIndices].sort((a, b) => a - b);
            for (const idx of sortedAsc) {
                // Benutze das Mapping VOR der Löschung
                // Beachte: Wenn bereits Zeilen früher in dieser Sitzung gelöscht wurden,
                // könnte das Mapping bereits angepasst sein
                const originalIdx = currentMapping[idx];
                // -1 = eingefügte Zeile, die gibt es nicht im Original
                if (originalIdx !== undefined && originalIdx >= 0) {
                    deletedOriginalIndices.push(originalIdx);
                }
            }
            
            explorerState.editedCells.set('_deletedRowIndices', { 
                originalIndices: deletedOriginalIndices,  // Array der ORIGINAL-Zeilen-Indices (0-basiert)
                count: deletedOriginalIndices.length
            });
            
            // HiddenRows anpassen
            const newHiddenRows = new Set();
            const sortedDeletedAsc = selectedIndices.slice().sort((a, b) => a - b);
            explorerState.hiddenRows.forEach(idx => {
                if (deletedSet.has(idx)) return;
                let offset = 0;
                sortedDeletedAsc.forEach(delIdx => {
                    if (delIdx < idx) offset++;
                });
                newHiddenRows.add(idx - offset);
            });
            explorerState.hiddenRows = newHiddenRows;
            
            // WICHTIG: rowMapping aktualisieren nach Löschen
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                // Live Session: Daten und Excel sind synchron, kein Mapping nötig
                explorerState.rowMapping = null;
            } else if (explorerState.rowMapping && explorerState.rowMapping.length > 0) {
                const newRowMapping = [...explorerState.rowMapping];
                // Von hinten nach vorne löschen (selectedIndices ist bereits absteigend sortiert)
                for (const delIdx of selectedIndices) {
                    if (delIdx < newRowMapping.length) {
                        newRowMapping.splice(delIdx, 1);
                    }
                }
                explorerState.rowMapping = newRowMapping;
            } else {
                // Erstelle neues Mapping: neue Position -> Original Excel-Zeile
                // Nach dem Löschen sind die verbleibenden Zeilen an neuen Positionen
                explorerState.rowMapping = [];
                for (let i = 0; i < explorerState.data.length; i++) {
                    // Berechne welche Original-Position (vor dem Löschen) diese Zeile hatte
                    let originalIdx = i;
                    // selectedIndices ist absteigend sortiert, wir brauchen aufsteigend
                    const sortedDeleted = selectedIndices.slice().sort((a, b) => a - b);
                    sortedDeleted.forEach(delIdx => {
                        if (delIdx <= originalIdx) originalIdx++;
                    });
                    explorerState.rowMapping.push(originalIdx);
                }
            }
            
            // WICHTIG: mergedCells in der GUI aktualisieren
            // Die mergedCells haben 0-basierte Zeilen-Indizes (Datenzeilen, nicht Excel-Zeilen)
            if (explorerState.mergedCells && explorerState.mergedCells.length > 0) {
                const sortedDeleted = selectedIndices.slice().sort((a, b) => a - b);
                const newMergedCells = [];
                
                for (const merge of explorerState.mergedCells) {
                    // mergedCells sind 0-basiert (startRow=0 ist Header, startRow=1 ist erste Datenzeile)
                    // Prüfe ob die Merged Cell komplett oder teilweise gelöscht wird
                    let affectedRows = 0;
                    let allRowsDeleted = true;
                    
                    for (let row = merge.startRow; row <= merge.endRow; row++) {
                        // Prüfe ob diese Zeile gelöscht wurde (row-1 weil mergedCells Header-Zeile einschließt)
                        const dataRowIdx = row - 1; // -1 weil startRow=1 ist Datenzeile 0
                        if (dataRowIdx >= 0 && sortedDeleted.includes(dataRowIdx)) {
                            affectedRows++;
                        } else {
                            allRowsDeleted = false;
                        }
                    }
                    
                    // Wenn alle Zeilen gelöscht wurden, Merged Cell entfernen
                    if (allRowsDeleted && merge.startRow > 0) {
                        continue;
                    }
                    
                    // Berechne neue Positionen
                    let newStartRow = merge.startRow;
                    let newEndRow = merge.endRow;
                    
                    // Für jede gelöschte Zeile vor der Merged Cell: verschiebe nach oben
                    for (const delIdx of sortedDeleted) {
                        const delRow = delIdx + 1; // +1 weil delIdx 0-basiert für Daten, mergedCells 1-basiert für Daten
                        if (delRow < merge.startRow) {
                            newStartRow--;
                            newEndRow--;
                        } else if (delRow >= merge.startRow && delRow <= merge.endRow) {
                            // Zeile innerhalb der Merged Cell gelöscht
                            newEndRow--;
                        }
                    }
                    
                    // Nur hinzufügen wenn die Merged Cell noch gültig ist
                    if (newStartRow <= newEndRow && newEndRow >= 0) {
                        newMergedCells.push({
                            ...merge,
                            startRow: newStartRow,
                            endRow: newEndRow,
                            rowSpan: newEndRow - newStartRow + 1
                        });
                    }
                }
                
                explorerState.mergedCells = newMergedCells;
            }
            
            // Speichere Undo-Aktion (nur die gelöschten Zeilen, nicht den gesamten Datensatz)
            pushExplorerUndo({
                type: 'deleteRows',
                deletedRows: deletedRows,
                previousHighlights: previousHighlights,
                previousEditedCells: previousEditedCells,
                deletedIndices: selectedIndices
            });
            
            // Auswahl zurücksetzen
            explorerState.selectedRows.clear();
            
            // Markiere als geändert
            explorerState.hasUnsavedChanges = true;
            
            // FilteredData neu erstellen (inkl. hiddenRows, Suchfilter, Spaltenfilter, Sortierung)
            filterExplorerData();
            updateHiddenRowsIndicator();
            
            const successMsg = currentLanguage === 'en'
                ? `${selectedCount} row(s) deleted`
                : `${selectedCount} Zeile(n) gelöscht`;
            showNotification(successMsg, 'success');
        }
        
        // Blendet mehrere ausgewählte Zeilen aus
        function hideSelectedRows() {
            const selectedCount = explorerState.selectedRows.size;
            if (selectedCount === 0) {
                showNotification(currentLanguage === 'en' ? 'No rows selected' : 'Keine Zeilen ausgewählt', 'warning');
                return;
            }
            
            // LIVE SESSION: Verstecke Zeilen sofort in Excel (Batch für Performance)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                const indicesToHide = Array.from(explorerState.selectedRows).map(idx => getExcelRowPosition(idx));
                // Verwende Batch-Funktion statt einzelner Aufrufe
                window.electronAPI.liveSessionHideRowsBatch(indicesToHide, true)
                    .then(result => {
                        if (result && result.success) {
                            console.log(`[LiveSession] ${indicesToHide.length} Zeilen in Excel versteckt (Batch)`);
                        } else {
                            console.error('[LiveSession] hideRowsBatch failed:', result);
                        }
                    })
                    .catch(error => console.error('[LiveSession] hideRowsBatch error:', error));
            }
            
            // Füge alle ausgewählten Zeilen zu den ausgeblendeten Zeilen hinzu
            // selectedRows enthält originalIndex-Werte (0-basiert)
            // hiddenRows enthält ebenfalls originalIndex-Werte (0-basiert)
            explorerState.selectedRows.forEach(originalIndex => {
                explorerState.hiddenRows.add(originalIndex);
            });
            
            // Markierung, dass Zeilen-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_rowVisibilityChanged', true);
            
            // Auswahl zurücksetzen
            explorerState.selectedRows.clear();
            
            // Markiere als geändert
            explorerState.hasUnsavedChanges = true;
            
            // Gefilterte Daten und UI aktualisieren
            filterExplorerData();
            updateHiddenRowsIndicator();
            updateExplorerEditStatus();
            updateRowMoveToolbar();
            
            const successMsg = currentLanguage === 'en'
                ? `${selectedCount} row(s) hidden`
                : `${selectedCount} Zeile(n) ausgeblendet`;
            showNotification(successMsg, 'success');
        }
        
        // Löscht die Zeilenauswahl
        function clearRowSelection() {
            explorerState.selectedRows.clear();
            updateRowSelectionUI();
            updateRowMoveToolbar();
            
            // Select-All Checkbox zurücksetzen
            const selectAllCheckbox = document.getElementById('selectAllRows');
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = false;
            }
        }

        function updateExplorerEditStatus() {
            const editedCount = explorerState.editedCells.size;
            const resultCount = elements.explorerResultCount.textContent;
            // Aktualisiere die Anzeige wenn sich die Anzahl geänderter Zellen ändert
            if (editedCount > 0) {
                elements.explorerStatus.textContent = `${editedCount} ${t('cellsEdited')}`;
                elements.explorerStatus.style.color = 'var(--warning-color)';
            } else {
                elements.explorerStatus.textContent = '';
            }
        }
        
        function updateExplorerPagination(totalPages) {
            const paginationEl = document.getElementById('explorerPagination');
            const pageInfoEl = document.getElementById('explorerPageInfo');
            const firstBtn = document.getElementById('btnExplorerFirstPage');
            const prevBtn = document.getElementById('btnExplorerPrevPage');
            const nextBtn = document.getElementById('btnExplorerNextPage');
            const lastBtn = document.getElementById('btnExplorerLastPage');
            
            // Pagination nur anzeigen wenn mehr als eine Seite
            if (explorerState.filteredData.length > explorerState.pageSize) {
                paginationEl.style.display = 'flex';
                pageInfoEl.textContent = `${t('pageLabel')} ${explorerState.currentPage} ${t('ofLabel')} ${totalPages}`;
                
                // Buttons aktivieren/deaktivieren
                firstBtn.disabled = explorerState.currentPage === 1;
                prevBtn.disabled = explorerState.currentPage === 1;
                nextBtn.disabled = explorerState.currentPage === totalPages;
                lastBtn.disabled = explorerState.currentPage === totalPages;
            } else {
                paginationEl.style.display = 'none';
            }
        }
        
        function explorerGoToPage(page) {
            // Legacy: Pagination ersetzt durch Virtual Scrolling
            // Scroll zur entsprechenden Position
            const rowIndex = (page - 1) * explorerState.pageSize;
            scrollToVirtualRow(Math.min(rowIndex, explorerState.filteredData.length - 1));
        }
        
        function explorerChangePageSize(newSize) {
            explorerState.pageSize = parseInt(newSize);
            // Virtual Scrolling: kein Page-Wechsel nötig
            renderExplorerTable();
        }
        
        function toggleColumnPanel() {
            const panel = document.getElementById('columnTogglePanel');
            const isHidden = panel.style.display === 'none';
            panel.style.display = isHidden ? 'flex' : 'none';
            
            // Button-Farbe umschalten
            const btn = document.getElementById('btnToggleColumns');
            if (btn) {
                btn.classList.toggle('btn-primary', !isHidden);
                btn.classList.toggle('btn-info', isHidden);
            }
        }
        
        function updateColumnToggles() {
            const container = document.getElementById('columnToggles');
            container.innerHTML = explorerState.headers.map((header, i) => `
                <div class="column-toggle ${explorerState.visibleColumns.includes(i) ? '' : 'hidden-col'}">
                    <input type="checkbox" id="colToggle_${i}" ${explorerState.visibleColumns.includes(i) ? 'checked' : ''} 
                           onchange="toggleExplorerColumn(${i}, this.checked)">
                    <label for="colToggle_${i}">${escapeHtml(header || `Spalte ${i + 1}`)}</label>
                </div>
            `).join('');
        }
        
        window.toggleExplorerColumn = async function(colIndex, visible) {
            // Hinweis bei aktivem Blattschutz: Hide wirkt nicht im Export
            if (!visible && explorerState.sheetProtected && !explorerState._sheetProtectedWarned) {
                explorerState._sheetProtectedWarned = true;
                showFloatingStatus(
                    currentLanguage === 'en'
                        ? '⚠️ Sheet is protected — hidden columns may not apply on export. Remove protection in Excel first.'
                        : '⚠️ Blattschutz aktiv — versteckte Spalten werden im Export evtl. nicht übernommen. Bitte zuerst in Excel den Blattschutz aufheben.',
                    'warning'
                );
            }
            if (visible) {
                if (!explorerState.visibleColumns.includes(colIndex)) {
                    explorerState.visibleColumns.push(colIndex);
                    explorerState.visibleColumns.sort((a, b) => a - b);
                }
            } else {
                explorerState.visibleColumns = explorerState.visibleColumns.filter(i => i !== colIndex);
            }
            
            // Live-Session: Spalte in Excel ein-/ausblenden
            // Checkboxen während Command disablen um Klick-Stacking bei schnellen Klicks
            // zu verhindern (sonst hängt Excel durch überlappende COM-Calls).
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                const toggleContainer = document.getElementById('columnToggles');
                if (toggleContainer) {
                    toggleContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.disabled = true; });
                }
                try {
                    const excelColIndex = getExcelColumnPosition(colIndex);
                    await window.electronAPI.liveSessionHideColumn(excelColIndex, !visible);
                } catch (error) {
                    console.error('[LiveSession] Spalten-Sichtbarkeit Fehler:', error);
                } finally {
                    if (toggleContainer) {
                        toggleContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.disabled = false; });
                    }
                }
            }
            
            // Markierung, dass Spalten-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_columnVisibilityChanged', true);
            renderExplorerTable();
            updateColumnToggles();
            updateHiddenColumnsIndicator();
            updateExplorerEditStatus();
            showFloatingStatus(visible ? (currentLanguage === 'en' ? 'Column shown' : 'Spalte eingeblendet') : (currentLanguage === 'en' ? 'Column hidden' : 'Spalte ausgeblendet'));
        };
        
        async function showAllExplorerColumns() {
            // Live-Session: Alle Spalten in Excel einblenden (Batch — 1 COM-Call)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    const hiddenColumns = explorerState.headers
                        .map((_, i) => i)
                        .filter(i => !explorerState.visibleColumns.includes(i));
                    if (hiddenColumns.length > 0) {
                        const excelCols = hiddenColumns.map(c => getExcelColumnPosition(c));
                        await window.electronAPI.liveSessionHideColumnsBatch(excelCols, false);
                    }
                    console.log('[LiveSession] All columns shown');
                } catch (error) {
                    console.error('[LiveSession] showAllColumns error:', error);
                }
            }
            
            explorerState.visibleColumns = explorerState.headers.map((_, i) => i);
            // Markierung, dass Spalten-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_columnVisibilityChanged', true);
            renderExplorerTable();
            updateColumnToggles();
            updateHiddenColumnsIndicator();
            updateExplorerEditStatus();
            showFloatingStatus(currentLanguage === 'en' ? 'All columns shown' : 'Alle Spalten eingeblendet');
        }
        
        async function hideAllExplorerColumns() {
            // Live-Session: Alle Spalten in Excel ausblenden (Batch)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    if (explorerState.visibleColumns.length > 0) {
                        const excelCols = explorerState.visibleColumns.map(c => getExcelColumnPosition(c));
                        await window.electronAPI.liveSessionHideColumnsBatch(excelCols, true);
                    }
                    console.log('[LiveSession] All columns hidden');
                } catch (error) {
                    console.error('[LiveSession] hideAllColumns error:', error);
                }
            }
            
            explorerState.visibleColumns = [];
            // Markierung, dass Spalten-Sichtbarkeit geändert wurde
            explorerState.editedCells.set('_columnVisibilityChanged', true);
            renderExplorerTable();
            updateColumnToggles();
            updateHiddenColumnsIndicator();
            updateExplorerEditStatus();
        }
        
        function addExplorerFilter() {
            const template = document.getElementById('explorerFilterTemplate');
            const clone = template.content.cloneNode(true);
            const row = clone.querySelector('.explorer-filter-row');
            
            // Spalten-Dropdown befüllen
            const colSelect = row.querySelector('.filter-column');
            colSelect.innerHTML = `<option value="">${t('selectColumn')}</option>` + 
                explorerState.headers.map((h, i) => `<option value="${i}">${escapeHtml(h || `Spalte ${i + 1}`)}</option>`).join('');
            
            const operatorSelect = row.querySelector('.filter-operator');
            const valueInput = row.querySelector('.filter-value');
            const daysInput = row.querySelector('.filter-days');
            const dateFromInput = row.querySelector('.filter-date-from');
            const dateToInput = row.querySelector('.filter-date-to');
            
            // Operator-Änderung: passende Felder anzeigen/ausblenden
            operatorSelect.onchange = () => {
                const op = operatorSelect.value;
                const needsDays = op === 'dateInDays' || op === 'dateOverdueDays';
                const needsDateRange = op === 'dateBetween';
                const needsNoValue = ['dateToday', 'datePast', 'dateFuture', 'dateThisWeek', 'dateThisMonth', 'isEmpty', 'isNotEmpty'].includes(op);
                
                daysInput.style.display = needsDays ? 'block' : 'none';
                dateFromInput.style.display = needsDateRange ? 'block' : 'none';
                dateToInput.style.display = needsDateRange ? 'block' : 'none';
                valueInput.style.display = (needsNoValue || needsDays || needsDateRange) ? 'none' : 'block';
                
                if (needsNoValue) {
                    valueInput.value = '_no_value_required_';
                } else if (needsDays || needsDateRange) {
                    valueInput.value = '';
                }
                
                // Placeholders aktualisieren wenn dateBetween gewählt
                if (needsDateRange) {
                    updateDatePlaceholders(row);
                }
                
                updateFiltersFromDOM(true);
            };
            
            // Event-Listener
            row.querySelector('.sync-filter').onclick = () => {
                syncFiltersToExcel();
            };
            row.querySelector('.remove-filter').onclick = () => {
                row.remove();
                updateFiltersFromDOM(true);
            };
            row.querySelector('.filter-column').onchange = () => {
                // Placeholders aktualisieren wenn Spalte gewechselt
                const op = operatorSelect.value;
                if (op === 'dateBetween') {
                    updateDatePlaceholders(row);
                }
                updateFiltersFromDOM(true);
            };
            daysInput.oninput = updateFiltersFromDOM;
            // Datumsfelder: Nur GUI filtern beim Tippen, Excel-Sync nur per Enter
            dateFromInput.oninput = () => {
                updateFiltersFromDOM();
                if (updateFiltersFromDOM._syncTimer) {
                    clearTimeout(updateFiltersFromDOM._syncTimer);
                    updateFiltersFromDOM._syncTimer = null;
                }
            };
            dateToInput.oninput = () => {
                updateFiltersFromDOM();
                if (updateFiltersFromDOM._syncTimer) {
                    clearTimeout(updateFiltersFromDOM._syncTimer);
                    updateFiltersFromDOM._syncTimer = null;
                }
            };
            
            // Enter in Datums-Textfeldern sendet an Excel
            dateFromInput.onkeydown = (e) => {
                if (e.key === 'Enter') { e.preventDefault(); syncFiltersToExcel(); }
            };
            dateToInput.onkeydown = (e) => {
                if (e.key === 'Enter') { e.preventDefault(); syncFiltersToExcel(); }
            };
            
            // Filter-Value: Enter sendet an Excel
            const filterValueInput = row.querySelector('.filter-value');
            filterValueInput.oninput = updateFiltersFromDOM;
            filterValueInput.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    syncFiltersToExcel();
                }
            };
            
            document.getElementById('explorerFilters').appendChild(row);
            document.getElementById('btnClearExplorerFilters').disabled = false;
        }
        
        function updateFiltersFromDOM(syncNow) {
            const rows = document.querySelectorAll('.explorer-filter-row');
            explorerState.filters = Array.from(rows).map(row => {
                const operator = row.querySelector('.filter-operator').value;
                const needsDays = operator === 'dateInDays' || operator === 'dateOverdueDays';
                const needsDateRange = operator === 'dateBetween';
                const needsNoValue = ['dateToday', 'datePast', 'dateFuture', 'dateThisWeek', 'dateThisMonth', 'isEmpty', 'isNotEmpty'].includes(operator);
                
                let value = row.querySelector('.filter-value').value;
                const days = row.querySelector('.filter-days').value;
                const dateFrom = row.querySelector('.filter-date-from')?.value || '';
                const dateTo = row.querySelector('.filter-date-to')?.value || '';
                
                // Für Datums-Filter mit Tagen: days als value verwenden
                if (needsDays) {
                    value = days || '0';
                } else if (needsNoValue) {
                    value = '_no_value_required_';
                } else if (needsDateRange) {
                    value = dateFrom || dateTo ? '_date_range_' : '';
                }
                
                return {
                    column: row.querySelector('.filter-column').value,
                    operator: operator,
                    value: value,
                    days: days,
                    dateFrom: dateFrom,
                    dateTo: dateTo
                };
            }).filter(f => f.column && (f.value || f.days || f.dateFrom || f.dateTo));
            updateFilterBadge();
            filterExplorerData();
            
            // Filter automatisch an Excel senden wenn Live-Session aktiv
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                if (updateFiltersFromDOM._syncTimer) clearTimeout(updateFiltersFromDOM._syncTimer);
                if (syncNow) {
                    // Diskrete Aktion (✕ Entfernen, Operator-/Spaltenwechsel): sofort senden
                    syncFiltersToExcel();
                } else {
                    // Tastatureingabe: Debounced, damit nicht bei jedem Tastendruck gesendet wird
                    updateFiltersFromDOM._syncTimer = setTimeout(() => {
                        syncFiltersToExcel();
                    }, 600);
                }
            }
        }
        
        async function clearExplorerFilters() {
            // Pendente Debounce-Syncs abbrechen
            if (updateFiltersFromDOM._syncTimer) clearTimeout(updateFiltersFromDOM._syncTimer);
            document.getElementById('explorerFilters').innerHTML = '';
            explorerState.filters = [];
            document.getElementById('btnClearExplorerFilters').disabled = true;
            updateFilterBadge();
            
            // AutoFilter-versteckte Zeilen und Range zurücksetzen
            // (Non-Streaming Reader liest row.hidden korrekt aus, Streaming Reader nicht)
            explorerState.hiddenRows.clear();
            explorerState.autoFilterRange = null;
            updateAutoFilterIndicator();
            updateHiddenRowsIndicator();
            
            // LIVE SESSION: Nur App-eigene AutoFilter in Excel zurücksetzen
            // (pre-existierende Excel-Filter bleiben erhalten)
            if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                try {
                    console.log('[LiveSession] clearAppFilters: starte... (active=true, ready=true)');
                    showFloatingStatus('⏳ Filter in Excel werden zurückgesetzt...');
                    // setAutoFilter([]) = nur App-Filter entfernen, Excel-eigene behalten
                    const result = await window.electronAPI.liveSessionSetAutoFilter([]);
                    console.log('[LiveSession] clearAppFilters result:', JSON.stringify(result));
                    if (result && result.success) {
                        showFloatingStatus('✅ Filter in Excel zurückgesetzt');
                    } else {
                        const errMsg = result?.error || 'Unbekannter Fehler';
                        console.error('[LiveSession] clearAppFilters failed:', errMsg);
                        showFloatingStatus(`⚠️ Filter-Reset: ${errMsg}`, true);
                    }
                } catch (error) {
                    console.error('[LiveSession] clearAppFilters exception:', error);
                    showFloatingStatus('❌ Filter-Reset in Excel fehlgeschlagen', true);
                }
            } else {
                console.log(`[LiveSession] clearAppFilters: SKIP - active=${explorerState.liveSessionActive}, ready=${explorerState.liveSessionReady}`);
                showFloatingStatus('🧹 Filter zurückgesetzt (nur lokal, keine Live-Session)');
            }
            
            filterExplorerData();
        }
        
        // Filter-Bereich ein-/ausblenden
        function toggleFilterSection() {
            const filtersDiv = document.getElementById('explorerFilters');
            const toggleIcon = document.getElementById('filterToggleIcon');
            const countBadge = document.getElementById('filterCountBadge');
            
            if (!filtersDiv) return;
            
            const isHidden = filtersDiv.style.display === 'none';
            
            if (isHidden) {
                // Einblenden
                filtersDiv.style.display = 'flex';
                if (toggleIcon) toggleIcon.textContent = '▼';
                if (countBadge) countBadge.style.display = 'none';
            } else {
                // Ausblenden
                filtersDiv.style.display = 'none';
                if (toggleIcon) toggleIcon.textContent = '▶';
                // Zeige Anzahl der aktiven Filter als Badge
                const filterCount = explorerState.filters.length;
                if (filterCount > 0 && countBadge) {
                    countBadge.textContent = `${filterCount} Filter aktiv`;
                    countBadge.style.display = 'inline-block';
                }
            }
        }
        
        function toggleFilterPanel() {
            const panel = document.getElementById('explorerFilterControls');
            const btn = document.getElementById('btnToggleFilterPanel');
            if (!panel) return;
            
            const isHidden = panel.style.display === 'none';
            panel.style.display = isHidden ? 'flex' : 'none';
            
            if (btn) {
                btn.classList.toggle('btn-primary', !isHidden);
                btn.classList.toggle('btn-info', isHidden);
            }
        }
        
        // Aktualisiert das Filter-Badge wenn Filter hinzugefügt/entfernt werden
        function updateFilterBadge() {
            const filtersDiv = document.getElementById('explorerFilters');
            const countBadge = document.getElementById('filterCountBadge');
            
            if (!countBadge) return;
            
            // Nur Badge anzeigen wenn Filter-Bereich ausgeblendet ist
            if (filtersDiv && filtersDiv.style.display === 'none') {
                const filterCount = explorerState.filters.length;
                if (filterCount > 0) {
                    countBadge.textContent = `${filterCount} Filter aktiv`;
                    countBadge.style.display = 'inline-block';
                } else {
                    countBadge.style.display = 'none';
                }
            } else {
                countBadge.style.display = 'none';
            }
        }
        
        /**
         * Undo — macht die letzte Aktion in der Live-Session rückgängig.
         * Nutzt einen internen Undo-Stack (nicht Excels Undo).
         */
        async function explorerUndo() {
            if (!explorerState.liveSessionActive || !explorerState.liveSessionReady) {
                showFloatingStatus('⚠️ Undo nur bei aktiver Live-Session verfügbar', 'warning');
                return;
            }
            
            const btnUndo = document.getElementById('btnExplorerUndo');
            if (btnUndo) {
                btnUndo.disabled = true;
                btnUndo.textContent = '⏳ Undo...';
            }
            
            try {
                // 0. Ausstehende Batch-Syncs sofort flushen (damit Python-Stack synchron bleibt)
                if (_cellSyncBatchTimer) {
                    clearTimeout(_cellSyncBatchTimer);
                    _cellSyncBatchTimer = null;
                }
                if (_pendingCellSyncs.size > 0) {
                    _flushCellSyncBatch();
                }
                
                // 1. Undo an Python-Session senden (pop Undo-Stack + reverse)
                const undoResult = await window.electronAPI.liveSessionUndo();
                console.log('[Undo] Step 1 - Undo Result:', JSON.stringify(undoResult));
                
                if (!undoResult || !undoResult.success) {
                    showFloatingStatus(`❌ Undo: ${undoResult?.error || 'Unbekannter Fehler'}`, 'error');
                    return;
                }
                
                // 2. Spalten-Reihenfolge zurücksetzen
                explorerState.columnOrder = [];
                explorerState.editedCells.delete('_columnMoved');
                
                // 3. Fast-Path für Aktionen die kein Datei-Reload brauchen
                if (undoResult.action === 'move_column') {
                    // move_column: explorerState.data/headers unverändert, nur columnOrder reset
                    console.log('[Undo] move_column → Fast-Path (kein readExcelSheet)');
                    renderExplorerTable();
                } else if (undoResult.action === 'restore_cell_value' || undoResult.action === 'restore_cells_batch' || undoResult.action === 'find_replace') {
                    // Zellwert-Undo: Python hat Excel via COM bereits aktualisiert.
                    // Frontend-Daten lokal aus dem Frontend-Undo-Stack aktualisieren.
                    // KEIN readExcelSheet! (verursacht Hang auf Windows wegen Excel-Datei-Lock)
                    console.log('[Undo] Zellwert-Undo → Fast-Path (lokales Update aus Frontend-Stack)');
                    
                    if (undoRedoState.explorerUndoStack.length > 0) {
                        const frontendAction = undoRedoState.explorerUndoStack.pop();
                        undoRedoState.explorerRedoStack.push(frontendAction);
                        
                        if (frontendAction.type === 'multi') {
                            // Multi-Zellen Undo (z.B. Suchen/Ersetzen)
                            frontendAction.actions.forEach(subAction => {
                                const { rowIndex, colIndex, oldValue, originalValue } = subAction;
                                explorerState.data[rowIndex][colIndex] = oldValue;
                                const filteredItem = explorerState.filteredData.find(item => item.originalIndex === rowIndex);
                                if (filteredItem && filteredItem.row) filteredItem.row[colIndex] = oldValue;
                                const cellKey = `${rowIndex}-${colIndex}`;
                                if (oldValue === originalValue) {
                                    explorerState.editedCells.delete(cellKey);
                                } else {
                                    explorerState.editedCells.set(cellKey, oldValue);
                                }
                            });
                        } else {
                            // Einzelzellen-Undo
                            const { rowIndex, colIndex, oldValue } = frontendAction;
                            explorerState.data[rowIndex][colIndex] = oldValue;
                            const filteredItem = explorerState.filteredData.find(item => item.originalIndex === rowIndex);
                            if (filteredItem && filteredItem.row) filteredItem.row[colIndex] = oldValue;
                            const cellKey = `${rowIndex}-${colIndex}`;
                            if (oldValue === frontendAction.originalValue) {
                                explorerState.editedCells.delete(cellKey);
                            } else {
                                explorerState.editedCells.set(cellKey, oldValue);
                            }
                        }
                    }
                    renderExplorerTable();
                } else if (undoResult.action === 'hide_column') {
                    // Spalten-Sichtbarkeit Undo: Python hat Excel-Spalte bereits ein-/ausgeblendet
                    // Frontend visibleColumns + UI aktualisieren
                    const params = undoResult.params || {};
                    const excelColIndex = params.col_index;
                    const wasShown = params.hidden === false;  // hidden=false → Spalte wurde eingeblendet
                    
                    // Excel-Index → Frontend-Index übersetzen
                    let frontendColIndex = excelColIndex;
                    if (explorerState.columnOrder.length > 0 && excelColIndex < explorerState.columnOrder.length) {
                        frontendColIndex = explorerState.columnOrder[excelColIndex];
                    }
                    
                    console.log(`[Undo] hide_column → Fast-Path: col=${frontendColIndex} (excel=${excelColIndex}), shown=${wasShown}`);
                    
                    if (wasShown) {
                        // Spalte wurde eingeblendet → zu visibleColumns hinzufügen
                        if (!explorerState.visibleColumns.includes(frontendColIndex)) {
                            explorerState.visibleColumns.push(frontendColIndex);
                            explorerState.visibleColumns.sort((a, b) => a - b);
                        }
                    } else {
                        // Spalte wurde ausgeblendet → aus visibleColumns entfernen
                        explorerState.visibleColumns = explorerState.visibleColumns.filter(i => i !== frontendColIndex);
                    }
                    
                    renderExplorerTable();
                    updateColumnToggles();
                    updateHiddenColumnsIndicator();
                } else if (undoResult.action === 'hide_row' || undoResult.action === 'hide_rows_batch') {
                    // Zeilen-Sichtbarkeit Undo: Python hat Excel-Zeile(n) bereits ein-/ausgeblendet
                    const params = undoResult.params || {};
                    const wasShown = params.hidden === false;
                    
                    if (undoResult.action === 'hide_row') {
                        const rowIndex = params.row_index;
                        console.log(`[Undo] hide_row → Fast-Path: row=${rowIndex}, shown=${wasShown}`);
                        if (wasShown) {
                            explorerState.hiddenRows.delete(rowIndex);
                        } else {
                            explorerState.hiddenRows.add(rowIndex);
                        }
                    } else {
                        // hide_rows_batch
                        const rowIndices = params.row_indices || [];
                        console.log(`[Undo] hide_rows_batch → Fast-Path: ${rowIndices.length} rows, shown=${wasShown}`);
                        for (const rowIndex of rowIndices) {
                            if (wasShown) {
                                explorerState.hiddenRows.delete(rowIndex);
                            } else {
                                explorerState.hiddenRows.add(rowIndex);
                            }
                        }
                    }
                    
                    // filteredData neu berechnen
                    explorerState.filteredData = explorerState.data.map((row, idx) => ({
                        row: row,
                        originalIndex: idx
                    }));
                    if (explorerState.hiddenRows.size > 0) {
                        explorerState.filteredData = explorerState.filteredData.filter(
                            item => !explorerState.hiddenRows.has(item.originalIndex)
                        );
                    }
                    
                    renderExplorerTable();
                    updateHiddenRowsIndicator();
                } else {
                    // Andere Undo-Typen (delete_column, insert_column, delete_row etc.):
                    // Daten über COM (Live-Session) lesen — NICHT von Disk (readExcelSheet)!
                    // readExcelSheet nutzt fs.readFile(), was auf Windows durch Excel-File-Lock
                    // extrem langsam ist oder komplett hängt.
                    console.log('[Undo] Step 2 - liveSessionGetData aufrufen (COM, kein Disk-I/O)...');
                    explorerState.sheetDataCache.clear();
                    explorerState.editedCells.clear();
                    
                    const result = await window.electronAPI.liveSessionGetData();
                    
                    if (result && result.success && result.headers && result.headers.length > 0) {
                        const newHeaders = result.headers.map(h => h != null ? String(h) : '');
                        const rawData = result.data || [];
                        const newData = rawData.map(row => {
                            if (!Array.isArray(row)) return [row];
                            return row;
                        });
                        
                        const columnsChanged = newHeaders.length !== explorerState.headers.length ||
                            newHeaders.some((h, i) => h !== explorerState.headers[i]);
                        
                        explorerState.headers = newHeaders;
                        explorerState.data = newData;
                        explorerState.originalData = newData.map(row => [...row]);
                        
                        if (columnsChanged) {
                            console.log('[Undo] Spaltenstruktur geändert → Layout-Reset');
                            explorerState.visibleColumns = newHeaders.map((_, i) => i);
                            explorerState.sortColumn = null;
                            explorerState.sortDirection = null;
                            // Metadaten resetten (stimmen nach Snapshot nicht mehr)
                            explorerState.cellStyles = {};
                            explorerState.cellFormulas = {};
                            explorerState.cellHyperlinks = {};
                            explorerState.richTextCells = {};
                            explorerState.mergedCells = [];
                            explorerState.headerStyles = {};
                            explorerState.cellVmMap = {};
                            updateColumnToggles();
                        }
                        
                        explorerState.filteredData = newData.map((row, index) => ({ originalIndex: index, row: row }));
                        
                        if (explorerState.hiddenRows.size > 0) {
                            explorerState.filteredData = explorerState.filteredData.filter(
                                item => !explorerState.hiddenRows.has(item.originalIndex)
                            );
                        }
                        
                        if (explorerState.sortColumn !== null) {
                            applyExplorerSort();
                        }
                        
                        explorerState.rowMapping = newData.map((_, i) => i);
                        
                        renderExplorerTable();
                        updateHiddenRowsIndicator();
                        updateHiddenColumnsIndicator();
                        updateAutoFilterIndicator();
                    } else {
                        console.warn('[Undo] liveSessionGetData fehlgeschlagen:', result?.error);
                        showFloatingStatus('⚠️ Undo ausgeführt, aber Daten konnten nicht gelesen werden', 'warning');
                        renderExplorerTable();
                    }
                }
                
                const label = undoResult.undone || 'Letzte Aktion';
                const remaining = undoResult.undoCount != null ? ` (${undoResult.undoCount} verbleibend)` : '';
                showFloatingStatus(`↩️ Rückgängig: ${label}${remaining}`);
                updateExplorerEditStatus();
                
                // Find&Replace Match-Liste aktualisieren (falls Panel offen)
                if (findReplaceState.lastSearchTerm && document.getElementById('findReplacePanel')?.style.display !== 'none') {
                    performFind();
                }
                
            } catch (err) {
                console.error('[Undo] Fehler:', err);
                showFloatingStatus(`❌ Undo Fehler: ${err.message}`, 'error');
            } finally {
                if (btnUndo) {
                    btnUndo.disabled = false;
                    btnUndo.textContent = '↩️ Undo';
                }
            }
        }
        
        async function exportExplorerData() {
            if (!explorerState.filePath) {
                elements.explorerStatus.textContent = t('noFileLoaded');
                return;
            }
            
            // Aktuelles Sheet im Cache speichern
            saveCurrentSheetToCache();
            
            // Sheet-Auswahl-Dialog anzeigen (inkl. Passwort-Option)
            const dialogResult = await showSheetSelectionDialog('export');
            if (!dialogResult || !dialogResult.sheets || dialogResult.sheets.length === 0) return;
            
            const selectedSheets = dialogResult.sheets;
            const exportPassword = dialogResult.password;
            
            // Dateiname ohne doppelte .xlsx Endung
            let baseName = explorerState.fileName || 'Daten';
            baseName = baseName.replace(/\.xlsx$/i, '');
            
            const savePath = await window.electronAPI.saveFileDialog({
                title: 'Export speichern',
                defaultPath: getWorkingDirectoryPath() ? (getWorkingDirectoryPath() + `/Export_${baseName}.xlsx`) : `Export_${baseName}.xlsx`,
                filters: [{ name: 'Excel', extensions: ['xlsx'] }]
            });
            
            if (savePath) {
                elements.explorerStatus.textContent = exportPassword ? 'Exportiere mit Passwortschutz...' : 'Exportiere...';
                
                // SPEICHEROPTIMIERUNG: Undo-Stack leeren vor dem Export
                if (undoRedoState.explorerUndoStack.length > 0) {
                    undoRedoState.explorerUndoStack = [];
                    undoRedoState.explorerRedoStack = [];
                    console.log('[Memory] Undo/Redo Stack geleert für Export');
                }
                
                // Im Live-Session-Modus: Daten sind bereits in Excel — nur Save-Befehl nötig
                // Alle Daten wurden bereits synchronisiert:
                // - Data Join: via await _syncDataJoinToLiveSession() (insertColumn + setColumnValues)
                // - Manuelle Edits: via blur-Handler (sofortige Zell-Sync)
                // Kein erneutes Senden nötig!
                let result;
                let sheetsToExport = null;
                const _exportStart = performance.now();
                
                if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                    console.log('[Export] Live-Session aktiv - nur Save-Befehl nötig...');
                    
                    try {
                        // Ausstehende Zell-Syncs flushen bevor gespeichert wird
                        if (_cellSyncBatchTimer) {
                            clearTimeout(_cellSyncBatchTimer);
                            _cellSyncBatchTimer = null;
                        }
                        if (_pendingCellSyncs.size > 0) {
                            console.log('[Export] Flushe', _pendingCellSyncs.size, 'ausstehende Zell-Syncs...');
                            const flushCells = Array.from(_pendingCellSyncs.values());
                            _pendingCellSyncs.clear();
                            await window.electronAPI.liveSessionSetCellsBatch(_mapCellsBatchCols(flushCells));
                        }
                        
                        // Filter-Debounce-Timer flushen (sonst können gerade getippte
                        // Filter noch ausstehen und werden nicht mitgespeichert)
                        let filterSyncNeeded = false;
                        if (updateFiltersFromDOM._syncTimer) {
                            clearTimeout(updateFiltersFromDOM._syncTimer);
                            updateFiltersFromDOM._syncTimer = null;
                            filterSyncNeeded = true; // Debounce war ausstehend → muss noch gesynct werden
                        }
                        
                        // Filter NUR synchronisieren wenn Debounce ausstehend war
                        // (sonst wurden Filter bereits per Debounce an Excel gesendet
                        // und ein redundanter AutoFilter-COM-Call kann bei großen
                        // Dateien mit Pivot-Tabellen Excel zum Hängen bringen)
                        if (filterSyncNeeded && explorerState.filters.some(f => f.column)) {
                            console.log('[Export] Flushe ausstehenden Filter-Sync vor dem Speichern...');
                            const syncResult = await syncFiltersToExcel();
                            console.log('[Export] Filter-Sync abgeschlossen');
                        }
                        
                        // Passwort-Logik
                        let passwordToUse;
                        if (exportPassword === undefined) {
                            passwordToUse = undefined;
                        } else if (exportPassword === '') {
                            passwordToUse = null;
                        } else {
                            passwordToUse = exportPassword;
                        }
                        
                        result = await window.electronAPI.liveSessionSaveFile(savePath, passwordToUse, selectedSheets);
                        console.log(`[Export TIMING] liveSessionSaveFile: ${(performance.now() - _exportStart).toFixed(0)}ms`);
                        
                        if (result.success) {
                            result.method = 'Live-Session';
                            if (exportPassword) {
                                explorerState.filePassword = exportPassword;
                            } else if (exportPassword === '') {
                                explorerState.filePassword = null;
                            }
                        }
                    } catch (error) {
                        console.error('[Export] Live-Session save error:', error);
                        result = { success: false, error: error.message };
                    }
                }
                
                // FALLBACK MODUS: Daten aufbereiten und über Python/xlwings exportieren
                if (!result) {
                
                sheetsToExport = [];
                
                for (const sheetName of selectedSheets) {
                    let sheetData;
                    
                    // Prüfe ob Sheet Änderungen hat
                    const isCurrentSheet = sheetName === explorerState.selectedSheet;
                    const cachedSheet = explorerState.sheetDataCache.get(sheetName);
                    
                    // Prüfe ob Filter aktiv sind (Zeilen werden ausgefiltert)
                    const hasActiveFilters = isCurrentSheet && explorerState.filters && explorerState.filters.length > 0;
                    const hasSearchFilter = isCurrentSheet && explorerState.searchTerm && explorerState.searchTerm.trim() !== '';
                    const rowsFiltered = isCurrentSheet && explorerState.filteredData.length < explorerState.data.length;
                    const isFiltered = (hasActiveFilters || hasSearchFilter) && rowsFiltered;
                    
                    // Änderungen = editedCells ODER aktive Filter (gefilterte Daten exportieren)
                    const hasChanges = isCurrentSheet 
                        ? (explorerState.editedCells.size > 0 || isFiltered)
                        : (cachedSheet?.editedCells?.size > 0);
                    
                    if (isFiltered) {
                        console.log(`[Export] Filter aktiv für Sheet "${sheetName}": ${explorerState.filteredData.length} von ${explorerState.data.length} Zeilen werden exportiert`);
                    }
                    
                    // AutoFilter-Range berechnen: Wenn GUI-Filter aktiv sind,
                    // MUSS ein autoFilter in der Export-Datei gesetzt werden.
                    // Range = A1:{lastCol}{lastRow} (gesamter Datenbereich inkl. Header)
                    let effectiveAutoFilterRange = isCurrentSheet
                        ? explorerState.autoFilterRange
                        : (cachedSheet?.autoFilterRange || null);
                    
                    if (isCurrentSheet && hasActiveFilters && !effectiveAutoFilterRange) {
                        const numCols = explorerState.headers.length;
                        const numRows = explorerState.data.length + 1; // +1 für Header
                        if (numCols > 0 && numRows > 1) {
                            const lastColLetter = getColumnLetter(numCols);
                            effectiveAutoFilterRange = `A1:${lastColLetter}${numRows}`;
                            console.log(`[Export] AutoFilter aus GUI-Filtern berechnet: ${effectiveAutoFilterRange}`);
                        }
                    }
                    
                    if (!hasChanges) {
                        // Keine Änderungen - Backend soll direkt aus Datei lesen (spart Speicher!)
                        // OPTIMIERUNG: Bei großen Dateien (>50k Styles) keine Styles senden
                        // ExcelJS liest die meisten Fills korrekt aus der Datei
                        const stylesForSheet = isCurrentSheet 
                            ? explorerState.cellStyles 
                            : (cachedSheet?.cellStyles || {});
                        
                        // Bei großen Dateien keine Styles senden - ExcelJS hat sie bereits
                        const styleCount = Object.keys(stylesForSheet || {}).length;
                        const finalStyles = styleCount > 50000 ? {} : stylesForSheet;
                        if (styleCount > 50000) {
                            console.log(`[Export] Sheet "${sheetName}": ${styleCount} Styles übersprungen (Performance)`);
                        }
                        
                        // Versteckte Spalten berechnen
                        const visibleColsForSheet = isCurrentSheet 
                            ? explorerState.visibleColumns 
                            : (cachedSheet?.visibleColumns || []);
                        const headersForSheet = isCurrentSheet 
                            ? explorerState.headers 
                            : (cachedSheet?.headers || []);
                        const allCols = headersForSheet.map((_, i) => i);
                        const visibleSet = new Set(visibleColsForSheet || allCols);
                        const hiddenColsForSheet = allCols.filter(i => !visibleSet.has(i));
                        
                        sheetData = {
                            sheetName: sheetName,
                            fromFile: true,
                            cellStyles: finalStyles,
                            hiddenColumns: hiddenColsForSheet,
                            autoFilterRange: effectiveAutoFilterRange
                        };
                        console.log(`[Export] fromFile Sheet "${sheetName}": ${hiddenColsForSheet.length} versteckte Spalten, AutoFilter: ${effectiveAutoFilterRange || 'none'}`);
                    } else if (isCurrentSheet) {
                        // Aktuelles Sheet mit Änderungen
                        const affectedRowsArray = explorerState.affectedRows ? Array.from(explorerState.affectedRows) : [];
                        const hasRowMoves = affectedRowsArray.length > 0;
                        
                        // Prüfe ob Filter aktiv sind (Zeilen werden ausgefiltert)
                        const hasActiveFilters = explorerState.filters && explorerState.filters.length > 0;
                        const hasSearchFilter = explorerState.searchTerm && explorerState.searchTerm.trim() !== '';
                        const rowsFiltered = explorerState.filteredData.length < explorerState.data.length;
                        const isFiltered = (hasActiveFilters || hasSearchFilter) && rowsFiltered;
                        
                        // Strukturelle Änderungen erkennen (Spalte/Zeile gelöscht/eingefügt/verschoben)
                        // HINWEIS: _rowVisibilityChanged ist KEINE strukturelle Änderung!
                        // Versteckte Zeilen bleiben im Excel, nur das hidden-Flag wird gesetzt.
                        // Die Styles bleiben dadurch an Ort und Stelle.
                        // HINWEIS: _rowHighlightChanged ist auch KEINE strukturelle Änderung!
                        // Zeilen-Highlights ändern nur die Zellfarben, nicht die Struktur.
                        // HINWEIS: isFiltered ist KEINE strukturelle Änderung!
                        // Filter werden über hiddenRows (hidden-Attribut) + autoFilterRange (Dropdown-Pfeile)
                        // abgebildet. Alle Zeilen bleiben in der Datei, nur nicht-passende werden versteckt.
                        const hasStructuralChanges = hasRowMoves || 
                            explorerState.editedCells.has('_columnDeleted') || 
                            explorerState.editedCells.has('_columnInserted') ||
                            explorerState.editedCells.has('_columnMoved') ||
                            explorerState.editedCells.has('_rowsReordered') ||
                            explorerState.editedCells.has('_rowInserted') ||  // Zeile eingefügt
                            explorerState.editedCells.has('_rowDeleted');     // Zeile gelöscht
                        
                        // HiddenRows aus dem aktuellen State ermitteln
                        // WICHTIG: hiddenRows enthält Indizes relativ zu explorerState.data (nicht filteredData!)
                        // filteredData enthält die versteckten Zeilen NICHT mehr, deshalb müssen wir data verwenden
                        let currentHiddenRows = [];
                        if (explorerState.hiddenRows && explorerState.hiddenRows.size > 0) {
                            // Bei strukturellen Änderungen: Index in der Export-Daten-Reihenfolge
                            // Wir müssen den Index finden, den die versteckte Zeile im Export haben wird
                            if (hasStructuralChanges) {
                                // Bei strukturellen Änderungen exportieren wir filteredData
                                // Aber versteckte Zeilen sind nicht in filteredData!
                                // Wir müssen die Zeilen aus data nehmen und an der richtigen Position einfügen
                                
                                // Erstelle Mapping: originalIndex -> Position in Export-Daten
                                // Die Export-Daten sind data[] ohne die gelöschten Zeilen (die sind schon weg)
                                // Versteckte Zeilen SOLLEN im Export sein (nur hidden-Attribut setzen)
                                explorerState.data.forEach((row, idx) => {
                                    if (explorerState.hiddenRows.has(idx)) {
                                        currentHiddenRows.push(idx);
                                    }
                                });
                            } else {
                                // Bei Zell-Edits: Direkt die originalIndex-Werte verwenden
                                currentHiddenRows = Array.from(explorerState.hiddenRows);
                            }
                        }
                        
                        // Filter als versteckte Zeilen abbilden (KEINE strukturelle Änderung!)
                        // Alle Zeilen bleiben in der Datei, nur nicht-passende bekommen hidden="1".
                        // Zusammen mit autoFilterRange ergibt das korrekte Excel-Filter.
                        if (isFiltered && !hasStructuralChanges) {
                            const filteredOriginalIndices = new Set(
                                explorerState.filteredData.map(item => item.originalIndex)
                            );
                            const existingHidden = new Set(currentHiddenRows);
                            for (let i = 0; i < explorerState.data.length; i++) {
                                if (!filteredOriginalIndices.has(i) && !existingHidden.has(i)) {
                                    currentHiddenRows.push(i);
                                }
                            }
                            console.log(`[Export] Filter: ${currentHiddenRows.length} Zeilen versteckt (${explorerState.filteredData.length} von ${explorerState.data.length} sichtbar)`);
                        }
                        
                        // Versteckte Spalten berechnen
                        const allCols = explorerState.headers.map((_, i) => i);
                        const visibleSet = new Set(explorerState.visibleColumns || allCols);
                        const hiddenCols = allCols.filter(i => !visibleSet.has(i));
                        
                        // STRATEGIE: Bei strukturellen Änderungen FULL REWRITE
                        if (hasStructuralChanges) {
                            // Strukturelle Änderung: FULL REWRITE
                            // Bei Filter: Nur gefilterte Zeilen exportieren
                            // Bei versteckten Zeilen (ohne Filter): Alle Zeilen exportieren
                            let allData;
                            if (isFiltered) {
                                // Filter aktiv: Nur die sichtbaren/gefilterten Zeilen exportieren
                                allData = explorerState.filteredData.map(item => item.row);
                                console.log(`[Export] Filter aktiv: Exportiere ${allData.length} von ${explorerState.data.length} Zeilen`);
                            } else {
                                // Kein Filter: Alle Zeilen (inkl. versteckte)
                                allData = explorerState.data;
                            }
                            
                            let exportHeaders = [...explorerState.headers];
                            // WICHTIG: cellStyles werden normalerweise NICHT mitgesendet!
                            // xlwings übernimmt Styles automatisch aus der Datei.
                            // AUSNAHME: Bei Data Join werden die Styles der importierten Spalten gesendet.
                            // Prüfe ob Spalten eingefügt wurden (Data Join)
                            const columnInsertedInfoCheck = explorerState.editedCells.get('_columnInserted');
                            let exportCellStyles = {};
                            if (columnInsertedInfoCheck && explorerState.cellStyles) {
                                // Data Join - sende die Styles für die neuen Spalten
                                // WICHTIG: cellStyles nutzen 1-basierte Row-Keys → 0-basiert konvertieren
                                for (const [key, value] of Object.entries(explorerState.cellStyles)) {
                                    const [rowStr, colStr] = key.split('-');
                                    exportCellStyles[`${parseInt(rowStr) - 1}-${colStr}`] = value;
                                }
                            }
                            let exportCellFormulas = { ...(explorerState.cellFormulas || {}) };
                            let exportCellHyperlinks = { ...(explorerState.cellHyperlinks || {}) };
                            
                            // WICHTIG: explorerState.richTextCells nutzt 1-basierte Row-Keys (styleKey-Format: "(row+1)-col"),
                            // aber der Python-Writer (_apply_rich_text_xml) erwartet 0-basierte Keys (row_idx + 2 = Excel-Zeile).
                            // Konvertierung: "1-0" (1-basiert) → "0-0" (0-basiert) = erste Datenzeile
                            let exportRichTextCells = {};
                            if (explorerState.richTextCells) {
                                for (const [key, value] of Object.entries(explorerState.richTextCells)) {
                                    const [rowStr, colStr] = key.split('-');
                                    exportRichTextCells[`${parseInt(rowStr) - 1}-${colStr}`] = value;
                                }
                            }
                            let exportVisibleColumns = [...(explorerState.visibleColumns || [])];
                            let exportHiddenCols = [...hiddenCols];
                            
                            // Bei gefilterten Daten: Styles/Formeln/Hyperlinks auf neue Zeilen-Indizes mappen
                            // originalIndex -> neuer Index in filteredData
                            if (isFiltered) {
                                console.log(`[Export] Filter aktiv: Mappe ${explorerState.filteredData.length} von ${explorerState.data.length} Zeilen`);
                                
                                // Erstelle Mapping: originalIndex -> newIndex (0-basiert in exportierten Daten)
                                const rowIndexMap = new Map();
                                explorerState.filteredData.forEach((item, newIdx) => {
                                    rowIndexMap.set(item.originalIndex, newIdx);
                                });
                                
                                // Styles remappen
                                const remappedStyles = {};
                                for (const [key, style] of Object.entries(exportCellStyles)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const originalRowIdx = parseInt(rowStr);
                                    const colIdx = parseInt(colStr);
                                    const newRowIdx = rowIndexMap.get(originalRowIdx);
                                    if (newRowIdx !== undefined) {
                                        remappedStyles[`${newRowIdx}-${colIdx}`] = style;
                                    }
                                }
                                exportCellStyles = remappedStyles;
                                
                                // Formeln remappen
                                const remappedFormulas = {};
                                for (const [key, formula] of Object.entries(exportCellFormulas)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const originalRowIdx = parseInt(rowStr);
                                    const colIdx = parseInt(colStr);
                                    const newRowIdx = rowIndexMap.get(originalRowIdx);
                                    if (newRowIdx !== undefined) {
                                        remappedFormulas[`${newRowIdx}-${colIdx}`] = formula;
                                    }
                                }
                                exportCellFormulas = remappedFormulas;
                                
                                // Hyperlinks remappen
                                const remappedHyperlinks = {};
                                for (const [key, link] of Object.entries(exportCellHyperlinks)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const originalRowIdx = parseInt(rowStr);
                                    const colIdx = parseInt(colStr);
                                    const newRowIdx = rowIndexMap.get(originalRowIdx);
                                    if (newRowIdx !== undefined) {
                                        remappedHyperlinks[`${newRowIdx}-${colIdx}`] = link;
                                    }
                                }
                                exportCellHyperlinks = remappedHyperlinks;
                                
                                // RichText remappen
                                const remappedRichText = {};
                                for (const [key, rt] of Object.entries(exportRichTextCells)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const originalRowIdx = parseInt(rowStr);
                                    const colIdx = parseInt(colStr);
                                    const newRowIdx = rowIndexMap.get(originalRowIdx);
                                    if (newRowIdx !== undefined) {
                                        remappedRichText[`${newRowIdx}-${colIdx}`] = rt;
                                    }
                                }
                                exportRichTextCells = remappedRichText;
                            }
                            
                            // Bei Filter: rowMapping erstellen für physische Zeilen-Umordnung
                            // rowMapping[newPos] = originalPos (0-basiert für Datenzeilen)
                            // Nutzt den schnellen ZIP/REGEX-Pfad der direkt auf XML arbeitet
                            // und alle Cell-Styles (Fill, Font, Alignment, Border) preserviert.
                            let filterRowMapping = null;
                            if (isFiltered && !hasRowMoves) {
                                filterRowMapping = explorerState.filteredData.map(item => item.originalIndex);
                                
                                // Bei Filter-Mapping: KEINE Frontend-Styles senden
                                // Die Excel-Styles werden durch physische XML-Umordnung übernommen
                                exportCellStyles = {};
                                exportCellFormulas = {};
                                exportCellHyperlinks = {};
                                exportRichTextCells = {};
                            }
                            
                            // Bei Row-Move: ExcelJS ordnet die Zeilen direkt um (mit allen Styles)
                            // Kein Frontend-Remapping nötig!
                            if (hasRowMoves && explorerState.rowMapping && explorerState.rowMapping.length > 0) {
                                // rowMapping wird an Writer gesendet
                            }
                            
                            // Hole die Indizes der gelöschten Spalten falls vorhanden
                            const columnDeletedInfo = explorerState.editedCells.get('_columnDeleted');
                            
                            // Extrahiere die Original-Indizes - unterstütze neues Format und Legacy-Formate
                            let deletedColumnIndices = [];
                            if (columnDeletedInfo && typeof columnDeletedInfo === 'object') {
                                if ('originalIndices' in columnDeletedInfo && Array.isArray(columnDeletedInfo.originalIndices)) {
                                    // Neues Format mit Original-Indices
                                    deletedColumnIndices = columnDeletedInfo.originalIndices.filter(v => v != null);
                                } else if ('indices' in columnDeletedInfo && Array.isArray(columnDeletedInfo.indices)) {
                                    // Altes Array-Format
                                    deletedColumnIndices = columnDeletedInfo.indices.filter(v => v != null);
                                } else if ('index' in columnDeletedInfo && columnDeletedInfo.index != null) {
                                    // Legacy single-index Format
                                    deletedColumnIndices = [columnDeletedInfo.index];
                                }
                            } else if (typeof columnDeletedInfo === 'number') {
                                deletedColumnIndices = [columnDeletedInfo];
                            }
                            
                            // Hole Info über eingefügte Spalten falls vorhanden
                            const columnInsertedInfo = explorerState.editedCells.get('_columnInserted');
                            let insertedColumnInfo = null;
                            if (columnInsertedInfo && typeof columnInsertedInfo === 'object') {
                                // Unterstütze neues Format mit 'operations' und alte Formate
                                if (columnInsertedInfo.operations) {
                                    // Neues Format mit operations Array — filtere ungültige Einträge
                                    insertedColumnInfo = {
                                        operations: columnInsertedInfo.operations.filter(op => op && op.position != null)
                                    };
                                } else if (columnInsertedInfo.position != null) {
                                    // Altes Format mit position
                                    insertedColumnInfo = {
                                        operations: [{
                                            position: columnInsertedInfo.position,
                                            count: columnInsertedInfo.count || 1,
                                            headers: columnInsertedInfo.headers || []
                                        }]
                                    };
                                } else if (columnInsertedInfo.index !== undefined) {
                                    // Legacy Format mit index/name
                                    insertedColumnInfo = {
                                        operations: [{
                                            position: columnInsertedInfo.index,
                                            count: 1,
                                            headers: [columnInsertedInfo.name || 'New Column']
                                        }]
                                    };
                                }
                            }
                            
                            // Prüfe ob Spalten verschoben wurden
                            const hasColumnMoves = explorerState.editedCells.has('_columnMoved');
                            let columnOrder = hasColumnMoves && explorerState.columnOrder.length > 0 
                                ? [...explorerState.columnOrder]
                                : null;
                            
                            // ========== ZEILEN-PARAMETER (analog zu Spalten) ==========
                            // Hole die Indizes der gelöschten Zeilen falls vorhanden
                            const rowDeletedInfo = explorerState.editedCells.get('_deletedRowIndices');
                            let deletedRowIndices = [];
                            if (rowDeletedInfo && typeof rowDeletedInfo === 'object') {
                                if ('originalIndices' in rowDeletedInfo && Array.isArray(rowDeletedInfo.originalIndices)) {
                                    deletedRowIndices = rowDeletedInfo.originalIndices.filter(v => v != null);
                                }
                            }
                            
                            // Hole Info über eingefügte Zeilen falls vorhanden
                            const rowInsertedInfo = explorerState.editedCells.get('_insertedRowInfo');
                            let insertedRowInfo = null;
                            if (rowInsertedInfo && typeof rowInsertedInfo === 'object') {
                                if (rowInsertedInfo.operations) {
                                    insertedRowInfo = {
                                        operations: rowInsertedInfo.operations.filter(op => op && op.position != null)
                                    };
                                }
                            }
                            
                            // Prüfe ob Zeilen verschoben wurden
                            const hasRowMoves2 = explorerState.editedCells.has('_rowsReordered');
                            // rowOrder: nur senden wenn Zeilen verschoben wurden (nicht bei nur löschen/einfügen)
                            let rowOrder = hasRowMoves2 && explorerState.rowMapping && explorerState.rowMapping.length > 0 
                                ? [...explorerState.rowMapping]
                                : null;
                            // ========== ENDE ZEILEN-PARAMETER ==========
                            
                            // Wenn Spalten verschoben wurden, ordne die Daten im Frontend um
                            // WICHTIG: columnOrder wird trotzdem an Python gesendet,
                            // damit Python die Formatierung in der Excel-Datei umordnen kann!
                            // HINWEIS: Bei gelöschten Spalten ist columnOrder bereits angepasst
                            // (enthält nur noch die verbleibenden Spalten-Indizes)
                            if (columnOrder) {
                                console.log('[Export] Ordne Daten gemäß columnOrder um...');
                                
                                // Headers umordnen
                                exportHeaders = columnOrder.map(oldIdx => explorerState.headers[oldIdx]);
                                
                                // Daten umordnen
                                allData = allData.map(row => columnOrder.map(oldIdx => row[oldIdx]));
                                
                                // Styles umordnen (Keys sind "rowIdx-colIdx")
                                const newCellStyles = {};
                                for (const [key, style] of Object.entries(exportCellStyles)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const rowIdx = parseInt(rowStr);
                                    const oldColIdx = parseInt(colStr);
                                    const newColIdx = columnOrder.indexOf(oldColIdx);
                                    if (newColIdx !== -1) {
                                        newCellStyles[`${rowIdx}-${newColIdx}`] = style;
                                    }
                                }
                                exportCellStyles = newCellStyles;
                                
                                // Formeln umordnen
                                const newFormulas = {};
                                for (const [key, formula] of Object.entries(exportCellFormulas)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const rowIdx = parseInt(rowStr);
                                    const oldColIdx = parseInt(colStr);
                                    const newColIdx = columnOrder.indexOf(oldColIdx);
                                    if (newColIdx !== -1) {
                                        newFormulas[`${rowIdx}-${newColIdx}`] = formula;
                                    }
                                }
                                exportCellFormulas = newFormulas;
                                
                                // Hyperlinks umordnen
                                const newHyperlinks = {};
                                for (const [key, link] of Object.entries(exportCellHyperlinks)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const rowIdx = parseInt(rowStr);
                                    const oldColIdx = parseInt(colStr);
                                    const newColIdx = columnOrder.indexOf(oldColIdx);
                                    if (newColIdx !== -1) {
                                        newHyperlinks[`${rowIdx}-${newColIdx}`] = link;
                                    }
                                }
                                exportCellHyperlinks = newHyperlinks;
                                
                                // RichText umordnen
                                const newRichText = {};
                                for (const [key, rt] of Object.entries(exportRichTextCells)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const rowIdx = parseInt(rowStr);
                                    const oldColIdx = parseInt(colStr);
                                    const newColIdx = columnOrder.indexOf(oldColIdx);
                                    if (newColIdx !== -1) {
                                        newRichText[`${rowIdx}-${newColIdx}`] = rt;
                                    }
                                }
                                exportRichTextCells = newRichText;
                                
                                // VisibleColumns umordnen (jetzt sind es die neuen Indizes)
                                exportVisibleColumns = explorerState.visibleColumns.map(oldIdx => columnOrder.indexOf(oldIdx)).filter(idx => idx !== -1);
                                
                                // HiddenColumns umordnen
                                exportHiddenCols = hiddenCols.map(oldIdx => columnOrder.indexOf(oldIdx)).filter(idx => idx !== -1);
                                
                                console.log('[Export] Daten umgeordnet: Headers=' + exportHeaders.length + ', Rows=' + allData.length);
                                
                                // columnOrder wird an Python gesendet für Formatierungs-Umordnung
                                // NICHT auf null setzen!
                            }
                            
                            // RowHighlights für Export vorbereiten
                            // ALLE aktiven Markierungen senden (Original-Daten-Indizes verwenden)
                            // WICHTIG: Original-Index = Daten-Array-Index = Excel-Zeile minus 2
                            // filteredData-Indizes wären FALSCH wenn versteckte Zeilen existieren,
                            // weil die Excel-Zeilen ihre Originalpositionen behalten
                            const exportRowHighlights = {};
                            explorerState.rowHighlights.forEach((color, originalIndex) => {
                                exportRowHighlights[originalIndex] = color;
                            });
                            
                            if (Object.keys(exportRowHighlights).length > 0) {
                                console.log(`[Export] ${Object.keys(exportRowHighlights).length} Zeilen-Markierungen werden gesendet`);
                            }
                            
                            // ClearedRowHighlights: Zeilen die ursprünglich markiert waren, aber jetzt nicht mehr
                            const clearedRowHighlights = [];
                            explorerState.originalRowHighlights.forEach((color, originalIndex) => {
                                if (!explorerState.rowHighlights.has(originalIndex)) {
                                    clearedRowHighlights.push(originalIndex);
                                }
                            });
                            
                            if (clearedRowHighlights.length > 0) {
                                console.log(`[Export] ${clearedRowHighlights.length} Zeilen-Markierungen werden entfernt`);
                            }
                            
                            // WICHTIG: changedCells auch bei strukturellen Änderungen senden!
                            // Bei Data Join sind die eingefügten Daten NUR in editedCells, nicht in der Datei.
                            // xlwings schreibt keine Bulk-Daten, also müssen wir die editedCells mitsenden.
                            let changedCellsForStructural = {};
                            
                            // WICHTIG: Cell-Edit-Keys IMMER in vollen Daten-Koordinaten senden
                            // (NICHT in filteredData-Indizes umrechnen!).
                            // Begründung: Die Python COMBINED-Pipeline (_fp_real_edits) operiert direkt
                            // auf der Source-XLSX (Voll-Datei), und auch insertedRowInfo.position /
                            // deletedRowIndices verwenden Voll-Daten-Indizes. Eine Filter-Umrechnung
                            // hier würde Edits in falsche Excel-Zeilen schreiben (oder verlieren),
                            // sobald Zeilen-Inserts/-Deletes mit aktivem Filter kombiniert werden.
                            // Bei aktivem Filter werden nicht-sichtbare Edits trotzdem mitgesendet —
                            // das ist korrekt, weil die Source-XLSX die zugehörige Zeile besitzt.
                            for (const cellKey of explorerState.editedCells.keys()) {
                                // Überspringe spezielle Marker-Keys
                                if (cellKey.startsWith('_')) continue;
                                
                                const [rowStr, colStr] = cellKey.split('-');
                                const originalRowIdx = parseInt(rowStr);
                                const colIdx = parseInt(colStr);
                                
                                if (originalRowIdx >= 0 && colIdx >= 0 && explorerState.data[originalRowIdx]) {
                                    changedCellsForStructural[cellKey] = explorerState.data[originalRowIdx][colIdx];
                                }
                            }
                            
                            // Spalten-Remap: editedCells-Keys nutzen ORIGINAL/logische Spaltenindizes,
                            // aber Python erwartet sie in physischen Excel-Spalten (nach columnOrder).
                            // Ohne dieses Remap landet ein Edit in der falschen Spalte (z.B. in der
                            // verschobenen Spalte statt in der Nachbarspalte).
                            if (columnOrder) {
                                const remappedEdits = {};
                                for (const [key, val] of Object.entries(changedCellsForStructural)) {
                                    const [rowStr, colStr] = key.split('-');
                                    const oldColIdx = parseInt(colStr);
                                    const newColIdx = columnOrder.indexOf(oldColIdx);
                                    if (newColIdx !== -1) {
                                        remappedEdits[`${rowStr}-${newColIdx}`] = val;
                                    }
                                    // Spalte gelöscht (newColIdx === -1) → Edit verwerfen
                                }
                                changedCellsForStructural = remappedEdits;
                            }
                            
                            if (Object.keys(changedCellsForStructural).length > 0) {
                                console.log(`[Export] ${Object.keys(changedCellsForStructural).length} geänderte Zellen bei struktureller Änderung`);
                            }
                            
                            sheetData = {
                                sheetName: sheetName,
                                headers: exportHeaders,
                                data: allData,
                                changedCells: changedCellsForStructural,  // WICHTIG: editedCells für xlwings
                                visibleColumns: exportVisibleColumns,
                                hiddenRows: currentHiddenRows,
                                hiddenColumns: exportHiddenCols,
                                cellStyles: exportCellStyles,
                                cellFormulas: exportCellFormulas,
                                cellHyperlinks: exportCellHyperlinks,
                                richTextCells: exportRichTextCells,
                                numberFormats: explorerState.numberFormats || {},
                                cellFonts: explorerState.cellFonts || {},
                                // ALLE aktiven Markierungen senden (Original-Datei hat keine)
                                rowHighlights: exportRowHighlights,
                                clearedRowHighlights: clearedRowHighlights,
                                autoFilterRange: effectiveAutoFilterRange,
                                fullRewrite: false,  // Python Fast-Path entscheidet selbst
                                structuralChange: true,  // Signalisiert dass Styles komplett zurückgesetzt werden müssen
                                // Spalten-Operationen
                                deletedColumnIndices: deletedColumnIndices,  // Array der gelöschten Spalten für spliceColumns
                                insertedColumnInfo: insertedColumnInfo,  // Info über eingefügte Spalten für spliceColumns
                                columnOrder: columnOrder,  // Neue Spaltenreihenfolge (null = keine Änderung oder bereits angewendet)
                                // Zeilen-Operationen (NEU - analog zu Spalten)
                                deletedRowIndices: deletedRowIndices,  // Array der gelöschten Zeilen-Original-Indizes
                                insertedRowInfo: insertedRowInfo,  // Info über eingefügte Zeilen
                                rowOrder: rowOrder,  // Zeilen-Reihenfolge bei Verschiebung
                                affectedRows: affectedRowsArray,  // Betroffene Zeilen bei Row-Move für Style-Reset
                                // rowMapping wird gesendet bei: Filter, Row-Move ODER Row-Deleted (Legacy)
                                rowMapping: filterRowMapping || (explorerState.rowMapping && explorerState.rowMapping.length > 0 ? explorerState.rowMapping : null),
                                // Operations Queues für serielle Abarbeitung
                                columnOperationsQueue: explorerState.columnOperationsQueue,
                                rowOperationsQueue: explorerState.rowOperationsQueue,
                                // Merged Cells (vollständiger Zustand aus GUI)
                                mergedCells: explorerState.mergedCells || [],
                                // VM-Map für Bild-Zellen (Copy&Paste von Zellbildern)
                                vmCellMap: explorerState.cellVmMap || {}
                            };
                            // Prüfe ob rowMapping gesendet wird
                            const hasRowMapping = !!(filterRowMapping || (explorerState.rowMapping && explorerState.rowMapping.length > 0));
                            console.log(`[Export] Strukturelle Änderung: Full Rewrite (${allData.length} Zeilen, ${exportHeaders.length} Spalten, ${exportHiddenCols.length} versteckte Spalten, ${currentHiddenRows.length} versteckte Zeilen, ${Object.keys(exportRowHighlights).length} Zeilen-Highlights, deletedColumnIndices: ${JSON.stringify(deletedColumnIndices)}, insertedColumnInfo: ${JSON.stringify(insertedColumnInfo)}, columnOrder: ${columnOrder ? 'angepasst' : 'original/angewendet'}, deletedRowIndices: ${JSON.stringify(deletedRowIndices)}, insertedRowInfo: ${JSON.stringify(insertedRowInfo)}, rowOrder: ${rowOrder ? 'angepasst' : 'nein'}, affectedRows: ${affectedRowsArray.length}, rowMapping: ${hasRowMapping ? 'ja' : 'nein'}, AutoFilter: ${effectiveAutoFilterRange || 'none'})`);
                        } else {
                            // Nur Zell-Edits: changedCells
                            const changedCells = {};
                            for (const cellKey of explorerState.editedCells.keys()) {
                                // Überspringe spezielle Marker-Keys
                                if (cellKey.startsWith('_')) continue;
                                
                                const [rowStr, colStr] = cellKey.split('-');
                                const rowIdx = parseInt(rowStr);
                                const colIdx = parseInt(colStr);
                                if (rowIdx >= 0 && colIdx >= 0 && explorerState.data[rowIdx]) {
                                    changedCells[cellKey] = explorerState.data[rowIdx][colIdx];
                                }
                            }
                            
                            // HINWEIS: Kein Bulk-Write (fullRewrite) mehr nötig!
                            // FALL 3a (_direct_xml_cell_edit) arbeitet direkt auf XML
                            // und ist performant für beliebig viele Zellen.
                            // fullRewrite=true erzwingt den openpyxl-Roundtrip (FALL 2),
                            // der AutoFilter/Tabellen in table1.xml zerstört.
                            {
                                // RichText für geänderte Zellen
                                // WICHTIG: editedCells nutzt 0-basierte Keys ("row-col"),
                                // aber richTextCells nutzt 1-basierte Keys ("(row+1)-col") wegen Header-Zeile!
                                const richTextForChanged = {};
                                for (const cellKey of explorerState.editedCells.keys()) {
                                    // Überspringe spezielle Marker-Keys
                                    if (cellKey.startsWith('_')) continue;
                                    
                                    // Konvertiere 0-basierten cellKey zu 1-basiertem styleKey
                                    const [rowStr, colStr] = cellKey.split('-');
                                    const styleKey = `${parseInt(rowStr) + 1}-${colStr}`;
                                    
                                    if (explorerState.richTextCells && explorerState.richTextCells[styleKey]) {
                                        richTextForChanged[cellKey] = explorerState.richTextCells[styleKey];
                                    }
                                }
                                
                                // CellStyles für geänderte Zellen — KOMPLETT senden (Font, Fill, Alignment, Borders)
                                // Genau wie RichText: vollständiges Style-Objekt, nicht nur Teilinfos
                                const cellStylesForChanged = {};
                                for (const cellKey of explorerState.editedCells.keys()) {
                                    if (cellKey.startsWith('_')) continue;
                                    
                                    const [rowStr, colStr] = cellKey.split('-');
                                    const styleKey = `${parseInt(rowStr) + 1}-${colStr}`;
                                    
                                    if (explorerState.cellStyles && explorerState.cellStyles[styleKey]) {
                                        const style = explorerState.cellStyles[styleKey];
                                        // Komplettes Style-Objekt senden (alle Eigenschaften)
                                        if (Object.keys(style).length > 0) {
                                            cellStylesForChanged[cellKey] = style;
                                        }
                                    }
                                }
                                
                                // CellFonts: aus cellFonts oder als Fallback aus cellStyles extrahieren
                                const cellFontsForChanged = {};
                                for (const cellKey of explorerState.editedCells.keys()) {
                                    if (cellKey.startsWith('_')) continue;
                                    const [rowStr, colStr] = cellKey.split('-');
                                    const styleKey = `${parseInt(rowStr) + 1}-${colStr}`;
                                    
                                    // Primär aus cellFonts (explizit gesetzt, z.B. Data Join)
                                    if (explorerState.cellFonts && explorerState.cellFonts[styleKey]) {
                                        cellFontsForChanged[cellKey] = explorerState.cellFonts[styleKey];
                                    }
                                }
                                
                                // RowHighlights für Zell-Edits Export vorbereiten
                                const exportRowHighlights = {};
                                if (explorerState.rowHighlights && explorerState.rowHighlights.size > 0) {
                                    explorerState.rowHighlights.forEach((color, originalIndex) => {
                                        exportRowHighlights[originalIndex] = color;
                                    });
                                }
                                
                                // ClearedRowHighlights: Zeilen die ursprünglich markiert waren, aber jetzt nicht mehr
                                const clearedRowHighlights = [];
                                explorerState.originalRowHighlights.forEach((color, originalIndex) => {
                                    if (!explorerState.rowHighlights.has(originalIndex)) {
                                        clearedRowHighlights.push(originalIndex);
                                    }
                                });
                                
                                if (clearedRowHighlights.length > 0) {
                                    console.log(`[Export] ${clearedRowHighlights.length} Zeilen-Markierungen werden entfernt`);
                                }
                                
                                sheetData = {
                                    sheetName: sheetName,
                                    changedCells: changedCells,
                                    richTextCells: Object.keys(richTextForChanged).length > 0 ? richTextForChanged : undefined,
                                    cellStyles: Object.keys(cellStylesForChanged).length > 0 ? cellStylesForChanged : undefined,
                                    cellFonts: Object.keys(cellFontsForChanged).length > 0 ? cellFontsForChanged : undefined,
                                    hiddenColumns: hiddenCols,
                                    hiddenRows: currentHiddenRows,
                                    rowHighlights: exportRowHighlights,
                                    clearedRowHighlights: clearedRowHighlights,
                                    autoFilterRange: effectiveAutoFilterRange,
                                    fullRewrite: false,
                                    mergedCells: explorerState.mergedCells || [],
                                    hasFormatChanges: explorerState.editedCells.has('_hasFormatChanges'),
                                    vmCellMap: explorerState.cellVmMap || {}
                                };
                                console.log(`[Export] Zell-Edits: ${Object.keys(changedCells).length} geänderte Zellen, ${Object.keys(richTextForChanged).length} RichText, ${Object.keys(cellStylesForChanged).length} CellStyles, ${Object.keys(cellFontsForChanged).length} CellFonts, ${hiddenCols.length} versteckte Spalten, ${currentHiddenRows.length} versteckte Zeilen, ${Object.keys(exportRowHighlights).length} Zeilen-Highlights, ${clearedRowHighlights.length} entfernte Highlights, AutoFilter: ${effectiveAutoFilterRange || 'none'}`);
                            console.log(`[Export] mergedCells für Zell-Edits: ${(explorerState.mergedCells || []).length} Einträge`, JSON.stringify(explorerState.mergedCells || []));
                            }
                        }
                    } else if (cachedSheet) {
                        // Gecachtes Sheet — prüfe ob strukturelle Änderungen vorliegen
                        const cachedEditedCells = cachedSheet.editedCells;
                        const cachedHasStructuralChanges = 
                            cachedEditedCells?.has('_columnDeleted') || 
                            cachedEditedCells?.has('_columnInserted') ||
                            cachedEditedCells?.has('_columnMoved') ||
                            cachedEditedCells?.has('_rowsReordered') ||
                            cachedEditedCells?.has('_rowInserted') ||
                            cachedEditedCells?.has('_rowDeleted');
                        
                        if (cachedHasStructuralChanges) {
                            // ===== Strukturelle Änderungen im gecachten Sheet — Full Rewrite =====
                            console.log(`[Export] Gecachtes Sheet "${sheetName}" hat strukturelle Änderungen`);
                            
                            // Spalten-Parameter extrahieren (analog zum currentSheet-Pfad)
                            const columnDeletedInfo = cachedEditedCells.get('_columnDeleted');
                            let deletedColumnIndices = [];
                            if (columnDeletedInfo && typeof columnDeletedInfo === 'object') {
                                if ('originalIndices' in columnDeletedInfo && Array.isArray(columnDeletedInfo.originalIndices)) {
                                    deletedColumnIndices = columnDeletedInfo.originalIndices.filter(v => v != null);
                                } else if ('indices' in columnDeletedInfo && Array.isArray(columnDeletedInfo.indices)) {
                                    deletedColumnIndices = columnDeletedInfo.indices.filter(v => v != null);
                                } else if ('index' in columnDeletedInfo) {
                                    deletedColumnIndices = [columnDeletedInfo.index];
                                }
                            } else if (typeof columnDeletedInfo === 'number') {
                                deletedColumnIndices = [columnDeletedInfo];
                            }
                            
                            const columnInsertedInfo = cachedEditedCells.get('_columnInserted');
                            let insertedColumnInfo = null;
                            if (columnInsertedInfo && typeof columnInsertedInfo === 'object') {
                                if (columnInsertedInfo.operations) {
                                    insertedColumnInfo = { operations: columnInsertedInfo.operations.filter(op => op && op.position != null) };
                                } else if (columnInsertedInfo.position !== undefined) {
                                    insertedColumnInfo = {
                                        operations: [{
                                            position: columnInsertedInfo.position,
                                            count: columnInsertedInfo.count || 1,
                                            headers: columnInsertedInfo.headers || []
                                        }]
                                    };
                                } else if (columnInsertedInfo.index !== undefined) {
                                    insertedColumnInfo = {
                                        operations: [{
                                            position: columnInsertedInfo.index,
                                            count: 1,
                                            headers: [columnInsertedInfo.name || 'New Column']
                                        }]
                                    };
                                }
                            }
                            
                            const hasColumnMoves = cachedEditedCells.has('_columnMoved');
                            let columnOrder = hasColumnMoves && cachedSheet.columnOrder && cachedSheet.columnOrder.length > 0 
                                ? [...cachedSheet.columnOrder]
                                : null;
                            
                            // Zeilen-Parameter
                            const rowDeletedInfo = cachedEditedCells.get('_deletedRowIndices');
                            let deletedRowIndices = [];
                            if (rowDeletedInfo && typeof rowDeletedInfo === 'object') {
                                if ('originalIndices' in rowDeletedInfo && Array.isArray(rowDeletedInfo.originalIndices)) {
                                    deletedRowIndices = rowDeletedInfo.originalIndices.filter(v => v != null);
                                }
                            }
                            
                            const rowInsertedInfo = cachedEditedCells.get('_insertedRowInfo');
                            let insertedRowInfo = null;
                            if (rowInsertedInfo && typeof rowInsertedInfo === 'object') {
                                if (rowInsertedInfo.operations) {
                                    insertedRowInfo = { operations: rowInsertedInfo.operations.filter(op => op && op.position != null) };
                                }
                            }
                            
                            let rowOrder = null;
                            
                            // Daten und Headers vorbereiten
                            let cachedExportHeaders = [...(cachedSheet.headers || [])];
                            let cachedAllData = (cachedSheet.data || []).map(row => [...row]);
                            
                            // Bei columnOrder: Daten umordnen (analog zum currentSheet-Pfad)
                            if (columnOrder) {
                                console.log(`[Export] Gecachtes Sheet "${sheetName}": Ordne Daten gemäß columnOrder um...`);
                                cachedExportHeaders = columnOrder.map(oldIdx => (cachedSheet.headers || [])[oldIdx]);
                                cachedAllData = cachedAllData.map(row => columnOrder.map(oldIdx => row[oldIdx]));
                            }
                            
                            // Versteckte Spalten berechnen
                            const allColsCached = cachedExportHeaders.map((_, i) => i);
                            const visibleSetCached = new Set(cachedSheet.visibleColumns || allColsCached);
                            const hiddenColsCached = allColsCached.filter(i => !visibleSetCached.has(i));
                            
                            // changedCells für strukturelle Änderungen (ohne _ Prefix)
                            const changedCellsForCachedStructural = {};
                            if (cachedEditedCells) {
                                for (const [cellKey, value] of cachedEditedCells) {
                                    if (cellKey.startsWith('_')) continue;
                                    const [rowStr, colStr] = cellKey.split('-');
                                    const colIdx = parseInt(colStr);
                                    const originalRowIdx = parseInt(rowStr);
                                    if (originalRowIdx >= 0 && colIdx >= 0 && cachedSheet.data[originalRowIdx]) {
                                        changedCellsForCachedStructural[cellKey] = cachedSheet.data[originalRowIdx][colIdx];
                                    }
                                }
                            }
                            
                            // Row Highlights
                            const cachedExportRowHighlights = {};
                            if (cachedSheet.rowHighlights && cachedSheet.rowHighlights.size > 0) {
                                cachedSheet.rowHighlights.forEach((color, originalIndex) => {
                                    cachedExportRowHighlights[originalIndex] = color;
                                });
                            }
                            
                            // RichText/CellStyles: 1-basierte Keys zu 0-basiert konvertieren (analog zum currentSheet-Pfad)
                            const cachedExportRichText = {};
                            if (cachedSheet.richTextCells) {
                                for (const [key, value] of Object.entries(cachedSheet.richTextCells)) {
                                    const [rowStr, colStr] = key.split('-');
                                    cachedExportRichText[`${parseInt(rowStr) - 1}-${colStr}`] = value;
                                }
                            }
                            const cachedExportCellStyles = {};
                            if (cachedSheet.cellStyles) {
                                for (const [key, value] of Object.entries(cachedSheet.cellStyles)) {
                                    const [rowStr, colStr] = key.split('-');
                                    cachedExportCellStyles[`${parseInt(rowStr) - 1}-${colStr}`] = value;
                                }
                            }
                            
                            sheetData = {
                                sheetName: sheetName,
                                headers: cachedExportHeaders,
                                data: cachedAllData,
                                changedCells: changedCellsForCachedStructural,
                                visibleColumns: cachedSheet.visibleColumns || [],
                                hiddenRows: cachedSheet.hiddenRows ? Array.from(cachedSheet.hiddenRows) : [],
                                hiddenColumns: hiddenColsCached,
                                cellStyles: cachedExportCellStyles,
                                richTextCells: cachedExportRichText,
                                cellFonts: {},
                                rowHighlights: cachedExportRowHighlights,
                                clearedRowHighlights: [],
                                autoFilterRange: cachedSheet.autoFilterRange || null,
                                fullRewrite: false,  // Python Fast-Path entscheidet selbst
                                structuralChange: true,
                                deletedColumnIndices: deletedColumnIndices,
                                insertedColumnInfo: insertedColumnInfo,
                                columnOrder: columnOrder,
                                deletedRowIndices: deletedRowIndices,
                                insertedRowInfo: insertedRowInfo,
                                rowOrder: rowOrder,
                                affectedRows: [],
                                rowMapping: null,
                                columnOperationsQueue: [],
                                rowOperationsQueue: [],
                                mergedCells: cachedSheet.mergedCells || [],
                                vmCellMap: {}
                            };
                            console.log(`[Export] Gecachtes Sheet "${sheetName}" strukturell: deletedColumns=${JSON.stringify(deletedColumnIndices)}, insertedColumns=${insertedColumnInfo ? 'ja' : 'nein'}, columnOrder=${columnOrder ? 'ja' : 'nein'}, deletedRows=${JSON.stringify(deletedRowIndices)}`);
                        } else {
                            // Keine strukturellen Änderungen - einfacher Pfad (Zell-Edits only)
                            const cachedChangedCells = {};
                            if (cachedSheet.editedCells) {
                                for (const [cellKey, value] of cachedSheet.editedCells) {
                                    if (cellKey.startsWith('_')) continue;
                                    const [rowStr, colStr] = cellKey.split('-');
                                    const colIdx = parseInt(colStr);
                                    const rowIdx = parseInt(rowStr);
                                    if (rowIdx >= 0 && colIdx >= 0 && cachedSheet.data[rowIdx]) {
                                        cachedChangedCells[cellKey] = cachedSheet.data[rowIdx][colIdx];
                                    }
                                }
                            }
                            
                            // Versteckte Spalten berechnen
                            const allColsCached2 = (cachedSheet.headers || []).map((_, i) => i);
                            const visibleSetCached2 = new Set(cachedSheet.visibleColumns || allColsCached2);
                            const hiddenColsCached2 = allColsCached2.filter(i => !visibleSetCached2.has(i));
                            
                            // Row Highlights
                            const cachedRowHighlights = {};
                            if (cachedSheet.rowHighlights && cachedSheet.rowHighlights.size > 0) {
                                cachedSheet.rowHighlights.forEach((color, originalIndex) => {
                                    cachedRowHighlights[originalIndex] = color;
                                });
                            }
                            
                            // RichText/CellStyles: 1-basierte Keys zu 0-basiert konvertieren
                            const cachedRichTextConverted = {};
                            if (cachedSheet.richTextCells) {
                                for (const [key, value] of Object.entries(cachedSheet.richTextCells)) {
                                    const [rowStr, colStr] = key.split('-');
                                    cachedRichTextConverted[`${parseInt(rowStr) - 1}-${colStr}`] = value;
                                }
                            }
                            const cachedCellStylesConverted = {};
                            if (cachedSheet.cellStyles) {
                                for (const [key, value] of Object.entries(cachedSheet.cellStyles)) {
                                    const [rowStr, colStr] = key.split('-');
                                    cachedCellStylesConverted[`${parseInt(rowStr) - 1}-${colStr}`] = value;
                                }
                            }
                            
                            sheetData = {
                                sheetName: sheetName,
                                headers: cachedSheet.headers || [],
                                data: cachedSheet.data || [],
                                changedCells: cachedChangedCells,
                                visibleColumns: cachedSheet.visibleColumns || [],
                                hiddenRows: cachedSheet.hiddenRows ? Array.from(cachedSheet.hiddenRows) : [],
                                hiddenColumns: hiddenColsCached2,
                                hasChanges: true,
                                cellStyles: cachedCellStylesConverted,
                                richTextCells: cachedRichTextConverted,
                                cellFonts: {},
                                rowHighlights: cachedRowHighlights,
                                clearedRowHighlights: [],
                                affectedRows: [],
                                autoFilterRange: cachedSheet.autoFilterRange || null,
                                mergedCells: cachedSheet.mergedCells || [],
                                vmCellMap: {}
                            };
                        }
                    } else {
                        // Sheet nicht im Cache - Backend soll aus Datei lesen
                        // Hier sind keine cellStyles vorhanden, Backend muss sie selbst extrahieren
                        sheetData = {
                            sheetName: sheetName,
                            fromFile: true,
                            cellStyles: {},  // Leer - Backend wird sie aus der Datei extrahieren
                            autoFilterRange: null  // Backend wird es aus der Datei lesen
                        };
                    }
                    
                    sheetsToExport.push(sheetData);
                }
                
                // FALLBACK MODUS: openpyxl/Python für Export verwenden (kein Excel verfügbar)
                console.log('[Export] Fallback-Modus - verwende openpyxl/Python für Export...');
                
                // Warnung bei Pivot-Tabellen im Fallback-Modus (ohne Live-Session können Pivot-Tabellen beschädigt werden)
                console.log('[Export Fallback] explorerState.hasPivotTables =', explorerState.hasPivotTables);
                if (explorerState.hasPivotTables) {
                    const isEn = currentLanguage === 'en';
                    const confirmed = await showConfirmDialog(
                        '⚠️ ' + (isEn ? 'Pivot Tables detected' : 'Pivot-Tabellen erkannt'),
                        isEn
                            ? 'This file contains pivot tables!\n\nWithout Live Mode, pivot tables may be lost or corrupted when saving.\n\nRecommendation: Use Live Mode or create a backup copy.'
                            : 'Diese Datei enthält Pivot-Tabellen!\n\nOhne Live-Modus können Pivot-Tabellen beim Speichern verloren gehen oder beschädigt werden.\n\nEmpfehlung: Verwenden Sie den Live Modus oder erstellen Sie eine Sicherheitskopie.',
                        isEn ? 'Save anyway' : 'Trotzdem speichern',
                        isEn ? 'Cancel' : 'Abbrechen'
                    );
                    if (!confirmed) {
                        elements.explorerStatus.textContent = isEn ? 'Save cancelled (pivot tables).' : 'Speichern abgebrochen (Pivot-Tabellen).';
                        return;
                    }
                }
                
                // DIAGNOSTIC: Log export data summary
                for (const s of sheetsToExport) {
                    // rowMapping Analyse
                    let rmInfo = 'null';
                    if (s.rowMapping) {
                        let isIdentity = true;
                        let firstDiff = -1;
                        for (let i = 0; i < s.rowMapping.length; i++) {
                            if (s.rowMapping[i] !== i) { isIdentity = false; firstDiff = i; break; }
                        }
                        rmInfo = `len=${s.rowMapping.length}, identity=${isIdentity}${!isIdentity ? `, firstDiff@${firstDiff}: val=${s.rowMapping[firstDiff]}` : ''}`;
                    }
                    // rowHighlights Analyse
                    const hlKeys = s.rowHighlights ? Object.keys(s.rowHighlights) : [];
                    const hlInfo = hlKeys.length > 0 ? `${hlKeys.length} highlights, first5Keys=[${hlKeys.slice(0,5).join(',')}]` : '0 highlights';
                    console.log(`[DIAG-FE] Sheet "${s.sheetName}": fromFile=${s.fromFile}, fullRewrite=${s.fullRewrite}, structuralChange=${s.structuralChange}, ` +
                        `deletedColumnIndices=${JSON.stringify(s.deletedColumnIndices || [])}, ` +
                        `insertedColumnInfo=${!!s.insertedColumnInfo}, columnOrder=${s.columnOrder ? s.columnOrder.length : 'null'}, ` +
                        `colOpsQueue=${(s.columnOperationsQueue || []).length}, ` +
                        `changedCells=${s.changedCells ? Object.keys(s.changedCells).length : 'undef'}, ` +
                        `rowMapping=(${rmInfo}), ` +
                        `data=${s.data ? s.data.length : 0} rows, ` +
                        `${hlInfo}`);
                }
                console.log(`[DIAG-FE] pendingSheetOperations: ${JSON.stringify(explorerState.pendingSheetOperations || [])}`);
                
                result = await window.electronAPI.pythonExportMultipleSheets({
                    sourcePath: explorerState.filePath,
                    originalSourcePath: explorerState.originalFilePath,
                    targetPath: savePath,
                    sheets: sheetsToExport,
                    password: exportPassword,
                    sourcePassword: explorerState.filePassword,
                    pendingSheetOperations: explorerState.pendingSheetOperations || [],
                    enginePreference: localStorage.getItem('excelSyncEngine') || 'auto'
                });
                
                } // Ende if (!result) — Fallback-Block
                
                if (result.success) {
                    const pwInfo = exportPassword ? ' (passwortgeschützt)' : '';
                    const engineInfo = result.method ? ` [${result.method}]` : '';
                    elements.explorerStatus.textContent = `✓ ${selectedSheets.length} Arbeitsblatt/blätter exportiert${pwInfo}${engineInfo}: ${savePath}`;
                    
                    // Im Live-Modus: State bereinigen, aber Datei nicht neu laden
                    if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                        console.log('[Export] Live-Session: Änderungsmarkierungen zurücksetzen...');
                        
                        // Änderungsmarkierungen zurücksetzen
                        explorerState.editedCells.clear();
                        explorerState.affectedRows?.clear();
                        explorerState.rowMapping = null;
                        
                        // rowHighlights: Aktuelle sind jetzt "original"
                        explorerState.originalRowHighlights = new Map(explorerState.rowHighlights);
                        
                        // originalData aktualisieren (aktuelle Daten sind jetzt die "Originale")
                        const _copyStart = performance.now();
                        explorerState.originalData = explorerState.data.map(row => [...row]);
                        console.log(`[Export TIMING] originalData copy: ${(performance.now() - _copyStart).toFixed(0)}ms (${explorerState.data.length} Zeilen)`);
                        
                        // Dateipfad aktualisieren falls "Speichern unter"
                        // NICHT aktualisieren bei gefiltertem Export (nicht alle Sheets),
                        // da das Original-Workbook weiterhin am Original-Pfad geöffnet bleibt
                        const isFilteredExport = selectedSheets.length < explorerState.sheets.length;
                        if (savePath !== explorerState.filePath && !isFilteredExport) {
                            explorerState.filePath = savePath;
                            explorerState.fileName = savePath.split('/').pop();
                            document.getElementById('explorerFileName').textContent = explorerState.fileName;
                        }
                        
                        // UI aktualisieren
                        const _filterStart = performance.now();
                        filterExplorerData();
                        console.log(`[Export TIMING] filterExplorerData: ${(performance.now() - _filterStart).toFixed(0)}ms`);
                        console.log(`[Export TIMING] GESAMT: ${(performance.now() - _exportStart).toFixed(0)}ms`);
                        
                        showFloatingStatus(`✓ Export erfolgreich${pwInfo}${engineInfo}`, 'success');
                    } else {
                        // Fallback-Modus: Strukturelle Änderungen behandeln
                        // Prüfe ob strukturelle Änderungen vorlagen (Spalten eingefügt/gelöscht)
                        const exportedSheet = sheetsToExport ? sheetsToExport.find(s => s.sheetName === explorerState.selectedSheet) : null;
                        const hadStructuralChange = exportedSheet && exportedSheet.structuralChange;
                        // XML-DIREKT (xml-col-ops) braucht keinen Reload: Styles bleiben intakt,
                        // Frontend-Daten sind bereits korrekt umgeordnet
                        const usedXmlDirekt = result.method && result.method.startsWith('xml-col-ops');
                    
                        if (hadStructuralChange && !usedXmlDirekt) {
                            // Bei strukturellen Änderungen: Datei neu einlesen für korrekte Styles
                            console.log('[Export] Strukturelle Änderung - lade Datei neu für korrekte Styles...');
                            
                            // Speichere aktuelle Ansicht für Wiederherstellung
                            const currentSheet = explorerState.selectedSheet;
                            const currentPage = explorerState.currentPage;
                            const currentSort = { column: explorerState.sortColumn, direction: explorerState.sortDirection, type: explorerState.sortType };
                            
                            // Row-Highlights VOR Reload sichern — reloadExplorerFile() löscht sie,
                            // und detectRowHighlights() im Reader erkennt sie nicht immer zuverlässig
                            const savedRowHighlights = new Map(explorerState.rowHighlights);
                            
                            // Datei neu laden (lädt auch das aktuelle Sheet automatisch)
                            await reloadExplorerFile(savePath, explorerState.filePassword);
                            
                            // Row-Highlights wiederherstellen (Indizes stimmen bereits,
                            // da insert/delete/move sie im Frontend laufend anpasst)
                            if (savedRowHighlights.size > 0) {
                                explorerState.rowHighlights = savedRowHighlights;
                                explorerState.originalRowHighlights = new Map(savedRowHighlights);
                                // Re-Render damit die Highlights sichtbar werden
                                filterExplorerData();
                            }
                            
                            // WICHTIG: Nach Reload nochmal sicherstellen dass alles sauber ist
                            explorerState.editedCells.clear();
                            explorerState.sheetDataCache.clear();
                            explorerState.rowMapping = null;
                            explorerState.affectedRows?.clear();
                            explorerState.pendingSheetOperations = [];
                            explorerState.liveSheetChanges = 0;
                            explorerState.sheetDiskNameMap.clear();
                            
                            // Sortierung wiederherstellen
                            if (currentSort.column !== null) {
                                explorerState.sortColumn = currentSort.column;
                                explorerState.sortDirection = currentSort.direction;
                                explorerState.sortType = currentSort.type;
                                applyExplorerSort();
                                renderExplorerTable();
                            }
                            
                            // Scroll-Position wiederherstellen (Virtual Scrolling)
                            if (currentPage > 1) {
                                const rowIndex = (currentPage - 1) * explorerState.pageSize;
                                scrollToVirtualRow(Math.min(rowIndex, explorerState.filteredData.length - 1));
                            }
                        } else {
                            // Keine strukturelle Änderung: Schneller Refresh ohne Datei neu lesen
                            console.log('[Export] Keine strukturelle Änderung - schneller State-Refresh...');
                            
                            if (exportedSheet && exportedSheet.fullRewrite && exportedSheet.data) {
                                // Headers und Data übernehmen
                                explorerState.headers = [...exportedSheet.headers];
                                explorerState.data = exportedSheet.data.map(row => [...row]);
                                explorerState.originalData = exportedSheet.data.map(row => [...row]);
                                
                                // filteredData neu aufbauen
                                explorerState.filteredData = explorerState.data.map((row, index) => ({
                                    originalIndex: index,
                                    row: row
                                }));
                                
                                // rowHighlights: Neue wurden gesetzt, sind jetzt "original"
                                explorerState.originalRowHighlights = new Map(explorerState.rowHighlights);
                            }
                            
                            // Änderungsmarkierungen zurücksetzen
                            explorerState.editedCells.clear();
                            explorerState.affectedRows?.clear();
                            explorerState.rowMapping = null;
                            explorerState.pendingSheetOperations = [];
                            explorerState.liveSheetChanges = 0;
                            explorerState.sheetDiskNameMap.clear();
                            explorerState.columnOperationsQueue = [];
                            explorerState.rowOperationsQueue = [];
                            
                            // Cache-Sheets: editedCells ebenfalls zurücksetzen
                            for (const [, cached] of explorerState.sheetDataCache) {
                                cached.editedCells.clear();
                                // originalData = aktuelle Daten (sind jetzt die neuen "Originale")
                                if (cached.data) {
                                    cached.originalData = cached.data.map(row => [...row]);
                                }
                            }
                            
                            // Cache aktualisieren
                            saveCurrentSheetToCache();
                            
                            // Dateipfad aktualisieren falls "Speichern unter"
                            // WICHTIG: Ohne dieses Update liest der nächste Export von der
                            // Original-Datei statt von der exportierten Datei!
                            const isFilteredExport2 = selectedSheets.length < explorerState.sheets.length;
                            if (savePath !== explorerState.filePath && !isFilteredExport2) {
                                explorerState.filePath = savePath;
                                explorerState.fileName = savePath.split('/').pop();
                                document.getElementById('explorerFileName').textContent = explorerState.fileName;
                            }
                            
                            // UI aktualisieren
                            filterExplorerData();
                        }
                        
                        showFloatingStatus(`✓ Export erfolgreich${pwInfo}${engineInfo}`, 'success');
                    } // Ende else (Fallback-Modus)
                    
                    // ========== GEMEINSAME CLEANUP nach erfolgreichem Export ==========
                    // Sicherstellen dass ALLE Change-Marker zurückgesetzt werden,
                    // unabhängig davon welcher Export-Pfad (Live/Fallback/Strukturell) genommen wurde
                    explorerState.editedCells.clear();
                    explorerState.pendingSheetOperations = [];
                    explorerState.liveSheetChanges = 0;
                    explorerState.sheetDiskNameMap.clear();
                    explorerState.columnOperationsQueue = [];
                    explorerState.rowOperationsQueue = [];
                    explorerState.affectedRows?.clear();
                    explorerState.rowMapping = null;
                    // originalData mit aktuellen Daten synchronisieren (keine Differenz mehr)
                    explorerState.originalData = explorerState.data.map(row => [...row]);
                    explorerState.originalRowHighlights = new Map(explorerState.rowHighlights);
                    for (const [, cached] of explorerState.sheetDataCache) {
                        cached.editedCells.clear();
                        if (cached.data) {
                            cached.originalData = cached.data.map(row => [...row]);
                        }
                    }
                    // Cache mit bereinigtem State aktualisieren
                    saveCurrentSheetToCache();
                    // Recovery-Daten aktualisieren (keine Änderungen mehr)
                    saveExplorerRecoveryData();
                    console.log(`[Export] Alle Change-Marker zurückgesetzt (editedCells=${explorerState.editedCells.size}, pending=${explorerState.pendingSheetOperations.length}, countAll=${countAllChanges()})`);
                    
                } else {
                    elements.explorerStatus.textContent = `Fehler: ${result.error}`;
                }
            }
        }
        
        // Interne Funktion: Lädt eine Datei ohne "ungespeicherte Änderungen"-Prüfung
        // Wird nach dem Speichern verwendet um den State zu synchronisieren
        async function reloadExplorerFile(filePath, password = null) {
            elements.explorerStatus.textContent = t('loadingFile');
            
            let result = await window.electronAPI.readExcelFile(filePath, password);
            
            // Passwortgeschützte Datei? (z.B. nach Passwortschutz-Vergabe beim Export)
            if (!result.success && result.needsPassword) {
                const enteredPassword = await showPromptDialog(
                    '🔐 Passwort erforderlich',
                    'Die gespeicherte Datei ist passwortgeschützt.\nBitte geben Sie das Passwort ein:',
                    '',
                    'password'
                );
                
                if (enteredPassword === null) {
                    elements.explorerStatus.textContent = 'Neuladen abgebrochen';
                    return;
                }
                
                result = await window.electronAPI.readExcelFile(filePath, enteredPassword);
                
                if (!result.success) {
                    if (result.needsPassword) {
                        showFloatingStatus('❌ Falsches Passwort', 'error');
                    } else {
                        elements.explorerStatus.textContent = `Fehler beim Neuladen: ${result.error}`;
                    }
                    return;
                }
                
                explorerState.filePassword = enteredPassword;
                showFloatingStatus('🔓 Datei entsperrt');
            } else if (!result.success) {
                console.error('[Reload] Fehler beim Neuladen:', result.error);
                elements.explorerStatus.textContent = `Fehler beim Neuladen: ${result.error}`;
                return;
            }
            
            // Cache komplett leeren
            explorerState.sheetDataCache.clear();
            explorerState.editedCells.clear();
            explorerState.rowHighlights.clear();
            explorerState.affectedRows?.clear();
            explorerState.rowMapping = null;
            
            // State aktualisieren
            explorerState.filePath = filePath;
            explorerState.fileName = result.fileName;
            explorerState.sheets = result.sheets;
            explorerState.filePassword = password;
            
            // UI aktualisieren
            document.getElementById('explorerFileName').textContent = explorerState.fileName;
            
            // Sheet-Dropdown füllen (mit Markierung für ausgeblendete Sheets)
            elements.explorerSheetSelect.innerHTML = explorerState.sheets
                .map(s => {
                    const isHidden = explorerState.hiddenSheets && explorerState.hiddenSheets.has(s);
                    const label = isHidden ? `👁️‍🗨️ ${s} (ausgeblendet)` : s;
                    return `<option value="${s}">${label}</option>`;
                })
                .join('');
            
            // Aktuelles Sheet auswählen (falls es noch existiert)
            const currentSheet = explorerState.selectedSheet;
            if (currentSheet && explorerState.sheets.includes(currentSheet)) {
                elements.explorerSheetSelect.value = currentSheet;
                await loadExplorerSheet(currentSheet);
            } else if (explorerState.sheets.length > 0) {
                await loadExplorerSheet(explorerState.sheets[0]);
            }
            
            console.log('[Reload] Datei neu geladen:', filePath);
        }
        
        // Dialog zur Auswahl der Arbeitsblätter
        function showSheetSelectionDialog(mode = 'export') {
            return new Promise((resolve) => {
                // Prüfe ob bereits ein Dialog existiert
                const existingDialog = document.querySelector('.sheet-selection-overlay');
                if (existingDialog) existingDialog.remove();
                
                const title = mode === 'export' ? 'Arbeitsblätter exportieren' : 'Arbeitsblätter auswählen';
                const confirmText = mode === 'export' ? 'Exportieren' : 'OK';
                
                const overlay = document.createElement('div');
                overlay.className = 'sheet-selection-overlay';
                overlay.style.cssText = `
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(0,0,0,0.6);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10002;
                `;
                
                const dialog = document.createElement('div');
                dialog.className = 'sheet-selection-dialog';
                dialog.style.cssText = `
                    background: var(--bg-medium);
                    border: 1px solid var(--border);
                    border-radius: 8px;
                    padding: 20px;
                    min-width: 350px;
                    max-width: 500px;
                    max-height: 80vh;
                    overflow: hidden;
                    display: flex;
                    flex-direction: column;
                    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
                `;
                
                // Sheet-Liste erstellen
                let sheetListHtml = '';
                explorerState.sheets.forEach((sheetName, index) => {
                    const isCurrentSheet = sheetName === explorerState.selectedSheet;
                    const isCached = explorerState.sheetDataCache.has(sheetName);
                    const hasChanges = isCurrentSheet 
                        ? explorerState.editedCells.size > 0 
                        : (isCached && explorerState.sheetDataCache.get(sheetName).editedCells.size > 0);
                    
                    const changesBadge = hasChanges ? '<span style="color: var(--warning); margin-left: 8px; font-size: 11px;">● Änderungen</span>' : '';
                    const currentBadge = isCurrentSheet ? '<span style="color: var(--primary); margin-left: 8px; font-size: 11px;">(aktuell)</span>' : '';
                    
                    sheetListHtml += `
                        <label style="display: flex; align-items: center; padding: 10px 12px; background: var(--bg-light); border-radius: 4px; cursor: pointer; user-select: none;">
                            <input type="checkbox" class="sheet-checkbox" value="${escapeHtml(sheetName)}" 
                                ${isCurrentSheet ? 'checked' : ''} 
                                style="width: 18px; height: 18px; margin-right: 12px; cursor: pointer;">
                            <span style="flex: 1;">${escapeHtml(sheetName)}</span>
                            ${currentBadge}${changesBadge}
                        </label>
                    `;
                });
                
                dialog.innerHTML = `
                    <h3 style="margin: 0 0 15px 0; color: var(--text);">📑 ${title}</h3>
                    <p style="margin: 0 0 15px 0; color: var(--text-secondary); font-size: 13px;">
                        Wählen Sie die Arbeitsblätter aus, die exportiert werden sollen:
                    </p>
                    <div style="display: flex; gap: 10px; margin-bottom: 12px;">
                        <button class="btn btn-sm btn-secondary" id="selectAllSheets">Alle auswählen</button>
                        <button class="btn btn-sm btn-secondary" id="selectNoneSheets">Keine auswählen</button>
                    </div>
                    <div style="overflow-y: auto; max-height: 300px; display: flex; flex-direction: column; gap: 6px; padding-right: 5px;">
                        ${sheetListHtml}
                    </div>
                    ${mode === 'export' ? `
                    <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid var(--border);">
                        <label style="display: flex; align-items: center; gap: 10px; cursor: pointer; user-select: none;">
                            <input type="checkbox" id="exportPasswordCheckbox" style="width: 18px; height: 18px; cursor: pointer;">
                            <span style="color: var(--text);">🔐 Mit Passwortschutz exportieren</span>
                        </label>
                        <div id="passwordInputContainer" style="display: none; margin-top: 10px; padding-left: 28px;">
                            <input type="password" id="exportPasswordInput" placeholder="Passwort eingeben" 
                                style="width: 100%; padding: 8px 12px; background: var(--bg-light); border: 1px solid var(--border); border-radius: 4px; color: var(--text); font-size: 14px;">
                        </div>
                    </div>
                    ` : ''}
                    <div style="display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; padding-top: 15px; border-top: 1px solid var(--border);">
                        <button class="btn btn-secondary" id="sheetDialogCancel">Abbrechen</button>
                        <button class="btn btn-primary" id="sheetDialogConfirm">${confirmText}</button>
                    </div>
                `;
                
                overlay.appendChild(dialog);
                document.body.appendChild(overlay);
                
                // Event handlers
                const confirmBtn = dialog.querySelector('#sheetDialogConfirm');
                const cancelBtn = dialog.querySelector('#sheetDialogCancel');
                const selectAllBtn = dialog.querySelector('#selectAllSheets');
                const selectNoneBtn = dialog.querySelector('#selectNoneSheets');
                const checkboxes = dialog.querySelectorAll('.sheet-checkbox');
                const passwordCheckbox = dialog.querySelector('#exportPasswordCheckbox');
                const passwordContainer = dialog.querySelector('#passwordInputContainer');
                const passwordInput = dialog.querySelector('#exportPasswordInput');
                
                // Passwort-Checkbox Toggle (nur im Export-Modus)
                if (passwordCheckbox) {
                    passwordCheckbox.onchange = () => {
                        passwordContainer.style.display = passwordCheckbox.checked ? 'block' : 'none';
                        if (passwordCheckbox.checked) {
                            passwordInput.focus();
                        }
                    };
                }
                
                selectAllBtn.onclick = () => {
                    checkboxes.forEach(cb => cb.checked = true);
                };
                
                selectNoneBtn.onclick = () => {
                    checkboxes.forEach(cb => cb.checked = false);
                };
                
                confirmBtn.onclick = () => {
                    const selected = Array.from(checkboxes)
                        .filter(cb => cb.checked)
                        .map(cb => cb.value);
                    
                    // Bei Export: Passwort mit zurückgeben
                    // undefined = Checkbox nicht aktiviert (Passwort beibehalten)
                    // '' = Checkbox aktiviert aber leer (Passwort entfernen)
                    // 'xxx' = Checkbox aktiviert mit Passwort (neues Passwort setzen)
                    let password = undefined;
                    if (mode === 'export' && passwordCheckbox) {
                        if (passwordCheckbox.checked) {
                            // Checkbox aktiviert: leerer String oder Passwort
                            password = passwordInput?.value || '';
                        }
                        // Checkbox nicht aktiviert: password bleibt undefined
                    }
                    
                    overlay.remove();
                    
                    // Rückgabe als Objekt mit sheets und password
                    if (mode === 'export') {
                        resolve({ sheets: selected, password: password });
                    } else {
                        resolve(selected);
                    }
                };
                
                cancelBtn.onclick = () => {
                    overlay.remove();
                    resolve(null);
                };
                
                // ESC zum Abbrechen
                const escHandler = (e) => {
                    if (e.key === 'Escape') {
                        overlay.remove();
                        document.removeEventListener('keydown', escHandler);
                        resolve(null);
                    }
                };
                document.addEventListener('keydown', escHandler);
            });
        }
        
        // Dialog für Passwortschutz beim Speichern/Exportieren
        function showPasswordProtectionDialog(currentPassword = null, mode = 'save') {
            return new Promise((resolve) => {
                const existingDialog = document.querySelector('.password-dialog-overlay');
                if (existingDialog) existingDialog.remove();
                
                const overlay = document.createElement('div');
                overlay.className = 'password-dialog-overlay';
                overlay.style.cssText = `
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(0,0,0,0.6);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10002;
                `;
                
                const dialog = document.createElement('div');
                dialog.className = 'password-dialog';
                dialog.style.cssText = `
                    background: var(--bg-medium);
                    border: 1px solid var(--border);
                    border-radius: 8px;
                    padding: 20px;
                    max-width: 450px;
                    width: 90%;
                    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
                `;
                
                const title = mode === 'export' ? '🔐 Export-Passwortschutz' : '🔐 Datei-Passwortschutz';
                const hasCurrentPassword = !!currentPassword;
                
                dialog.innerHTML = `
                    <h3 style="margin: 0 0 15px 0; color: var(--text);">${title}</h3>
                    <p style="margin: 0 0 15px 0; color: var(--text-secondary); font-size: 13px;">
                        ${mode === 'export' 
                            ? 'Die exportierte Datei kann mit einem Passwort geschützt werden.'
                            : 'Die Datei kann mit einem Passwort geschützt werden.'}
                        <br>Excel-kompatible Verschlüsselung - keine zusätzlichen Tools nötig.
                    </p>
                    
                    <div style="margin-bottom: 15px;">
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 10px;">
                            <input type="radio" name="passwordOption" value="none" ${!hasCurrentPassword ? 'checked' : ''}>
                            <span style="color: var(--text);">Kein Passwortschutz</span>
                        </label>
                        ${hasCurrentPassword ? `
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 10px;">
                            <input type="radio" name="passwordOption" value="keep" checked>
                            <span style="color: var(--text);">Bestehendes Passwort beibehalten</span>
                        </label>` : ''}
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                            <input type="radio" name="passwordOption" value="new">
                            <span style="color: var(--text);">${hasCurrentPassword ? 'Neues Passwort setzen' : 'Mit Passwort schützen'}</span>
                        </label>
                    </div>
                    
                    <div id="newPasswordSection" style="display: none; margin-bottom: 15px; padding: 12px; background: var(--bg-dark); border-radius: 6px;">
                        <label style="display: block; margin-bottom: 6px; color: var(--text-muted); font-size: 12px;">Neues Passwort:</label>
                        <input type="password" id="newPasswordInput" placeholder="Passwort eingeben..." 
                               style="width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 4px; 
                                      background: var(--bg-light); color: var(--text); box-sizing: border-box; margin-bottom: 10px;">
                        <label style="display: block; margin-bottom: 6px; color: var(--text-muted); font-size: 12px;">Passwort bestätigen:</label>
                        <input type="password" id="confirmPasswordInput" placeholder="Passwort wiederholen..." 
                               style="width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 4px; 
                                      background: var(--bg-light); color: var(--text); box-sizing: border-box;">
                        <div id="passwordError" style="color: #F44336; font-size: 12px; margin-top: 8px; display: none;"></div>
                    </div>
                    
                    <div style="display: flex; gap: 10px; justify-content: flex-end;">
                        <button class="btn btn-secondary" id="pwDialogCancel">Abbrechen</button>
                        <button class="btn btn-success" id="pwDialogConfirm">Fortfahren</button>
                    </div>
                `;
                
                overlay.appendChild(dialog);
                document.body.appendChild(overlay);
                
                // Radio-Button Handler
                const radios = dialog.querySelectorAll('input[name="passwordOption"]');
                const newPwSection = dialog.querySelector('#newPasswordSection');
                const newPwInput = dialog.querySelector('#newPasswordInput');
                const confirmPwInput = dialog.querySelector('#confirmPasswordInput');
                const pwError = dialog.querySelector('#passwordError');
                
                radios.forEach(radio => {
                    radio.onchange = () => {
                        newPwSection.style.display = radio.value === 'new' && radio.checked ? 'block' : 'none';
                        if (radio.value === 'new' && radio.checked) {
                            setTimeout(() => newPwInput.focus(), 100);
                        }
                    };
                });
                
                // Confirm Button
                dialog.querySelector('#pwDialogConfirm').onclick = () => {
                    const selectedOption = dialog.querySelector('input[name="passwordOption"]:checked').value;
                    
                    if (selectedOption === 'none') {
                        overlay.remove();
                        resolve({ action: 'none', password: null });
                    } else if (selectedOption === 'keep') {
                        overlay.remove();
                        resolve({ action: 'keep', password: currentPassword });
                    } else if (selectedOption === 'new') {
                        const newPw = newPwInput.value;
                        const confirmPw = confirmPwInput.value;
                        
                        if (!newPw) {
                            pwError.textContent = 'Bitte Passwort eingeben';
                            pwError.style.display = 'block';
                            newPwInput.focus();
                            return;
                        }
                        if (newPw.length < 4) {
                            pwError.textContent = 'Passwort muss mindestens 4 Zeichen haben';
                            pwError.style.display = 'block';
                            newPwInput.focus();
                            return;
                        }
                        if (newPw !== confirmPw) {
                            pwError.textContent = 'Passwörter stimmen nicht überein';
                            pwError.style.display = 'block';
                            confirmPwInput.focus();
                            return;
                        }
                        
                        overlay.remove();
                        resolve({ action: 'new', password: newPw });
                    }
                };
                
                // Cancel Button
                dialog.querySelector('#pwDialogCancel').onclick = () => {
                    overlay.remove();
                    resolve(null);
                };
                
                // ESC zum Abbrechen
                const escHandler = (e) => {
                    if (e.key === 'Escape') {
                        overlay.remove();
                        document.removeEventListener('keydown', escHandler);
                        resolve(null);
                    }
                };
                document.addEventListener('keydown', escHandler);
            });
        }

        // Bestätigungsdialog anzeigen
        function showConfirmDialog(title, message, confirmText = 'OK', cancelText = 'Abbrechen') {
            return new Promise((resolve) => {
                // Prüfe ob bereits ein Dialog existiert
                const existingDialog = document.querySelector('.confirm-dialog-overlay');
                if (existingDialog) existingDialog.remove();
                
                const overlay = document.createElement('div');
                overlay.className = 'confirm-dialog-overlay';
                overlay.style.cssText = `
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(0,0,0,0.6);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10002;
                `;
                
                const dialog = document.createElement('div');
                dialog.className = 'confirm-dialog';
                dialog.style.cssText = `
                    background: var(--bg-medium);
                    border: 1px solid var(--border);
                    border-radius: 8px;
                    padding: 20px;
                    max-width: 400px;
                    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
                `;
                
                const buttonsHtml = cancelText 
                    ? `<button class="btn btn-secondary" id="confirmDialogCancel">${cancelText}</button>
                       <button class="btn btn-success" id="confirmDialogConfirm">${confirmText}</button>`
                    : `<button class="btn btn-success" id="confirmDialogConfirm">${confirmText}</button>`;
                
                dialog.innerHTML = `
                    <h3 style="margin: 0 0 15px 0; color: var(--text);">${title}</h3>
                    <p style="margin: 0 0 20px 0; color: var(--text-secondary); white-space: pre-line;">${message}</p>
                    <div style="display: flex; gap: 10px; justify-content: flex-end;">
                        ${buttonsHtml}
                    </div>
                `;
                
                overlay.appendChild(dialog);
                document.body.appendChild(overlay);
                
                // Event handlers
                const confirmBtn = dialog.querySelector('#confirmDialogConfirm');
                const cancelBtn = dialog.querySelector('#confirmDialogCancel');
                
                confirmBtn.onclick = () => {
                    overlay.remove();
                    resolve(true);
                };
                
                if (cancelBtn) {
                    cancelBtn.onclick = () => {
                        overlay.remove();
                        resolve(false);
                    };
                }
                
                // ESC zum Abbrechen
                const escHandler = (e) => {
                    if (e.key === 'Escape') {
                        overlay.remove();
                        document.removeEventListener('keydown', escHandler);
                        resolve(false);
                    }
                };
                document.addEventListener('keydown', escHandler);
            });
        }
        
        // ==================== Datei-Info Modal ====================
        async function showFileInfoModal() {
            if (!explorerState.filePath) return;
            
            const result = await window.electronAPI.getFileMetadata(explorerState.filePath);
            if (!result.success) {
                showFloatingStatus('Fehler: ' + result.error, 'error');
                return;
            }
            
            const isEn = currentLanguage === 'en';
            const locale = isEn ? 'en-US' : 'de-DE';
            
            // Overlay
            const existingDialog = document.querySelector('.fileinfo-dialog-overlay');
            if (existingDialog) existingDialog.remove();
            
            const overlay = document.createElement('div');
            overlay.className = 'fileinfo-dialog-overlay';
            overlay.style.cssText = `position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.6);display:flex;align-items:center;justify-content:center;z-index:10002;`;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `background:var(--bg-medium);border:1px solid var(--border);border-radius:10px;padding:0;max-width:600px;width:90vw;max-height:80vh;box-shadow:0 8px 32px rgba(0,0,0,0.4);display:flex;flex-direction:column;`;
            
            const formatSize = (bytes) => {
                if (bytes < 1024) return bytes + ' B';
                if (bytes < 1024*1024) return (bytes/1024).toFixed(1) + ' KB';
                return (bytes/(1024*1024)).toFixed(2) + ' MB';
            };
            
            const formatDate = (d) => {
                if (!d) return '—';
                try { return new Date(d).toLocaleString(locale); } catch { return d; }
            };
            
            const badge = (text, color) => `<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;background:${color};color:white;margin:2px;">${text}</span>`;
            
            // Features sammeln
            const features = [];
            if (result.hasPivotTables) features.push(badge(`${result.pivotTableCount} ${isEn ? 'Pivot Table' : 'Pivot-Tabelle'}${result.pivotTableCount > 1 ? (isEn ? 's' : 'n') : ''}`, '#e91e63'));
            if (result.hasCharts) features.push(badge(`${result.chartCount} ${isEn ? 'Chart' : 'Diagramm'}${result.chartCount > 1 ? (isEn ? 's' : 'e') : ''}`, '#2196F3'));
            if (result.hasTables) features.push(badge(`${result.tableCount} ${isEn ? 'Table' : 'Tabelle'}${result.tableCount > 1 ? (isEn ? 's' : 'n') : ''}`, '#009688'));
            if (result.hasMacros) features.push(badge(isEn ? 'Macros (VBA)' : 'Makros (VBA)', '#ff5722'));
            if (result.hasImages) features.push(badge(`${result.imageCount} ${isEn ? 'Image' : 'Bild'}${result.imageCount > 1 ? (isEn ? 's' : 'er') : ''}`, '#795548'));
            if (result.hasComments) features.push(badge(isEn ? 'Comments' : 'Kommentare', '#607d8b'));
            if (result.hasConditionalFormatting) features.push(badge(isEn ? 'Conditional Formatting' : 'Bedingte Formatierung', '#9c27b0'));
            if (result.hasDataValidations) features.push(badge(isEn ? 'Data Validation' : 'Datenvalidierung', '#ff9800'));
            if (result.hasExternalLinks) features.push(badge(`${result.externalLinkCount} Ext. Link${result.externalLinkCount > 1 ? 's' : ''}`, '#f44336'));
            if (result.isPasswordProtected) features.push(badge(isEn ? 'Password Protected' : 'Passwortgeschützt', '#d32f2f'));
            
            const row = (label, value) => value && value !== '—' ? `<tr><td style="padding:4px 12px 4px 0;color:var(--text-muted);white-space:nowrap;vertical-align:top;">${label}</td><td style="padding:4px 0;color:var(--text);word-break:break-all;">${value}</td></tr>` : '';
            
            // Sheet-Info mit Zeilen/Spalten
            const dims = result.sheetDimensions || {};
            const sheetsInfo = result.sheets.map(s => {
                const hidden = result.hiddenSheets.includes(s);
                const dim = dims[s];
                let dimText = '';
                if (dim) {
                    dimText = ` <span style="font-size:10px;color:var(--text-muted);">(${dim.rows.toLocaleString(locale)} ${isEn ? 'rows' : 'Zeilen'} × ${dim.cols} ${isEn ? 'cols' : 'Spalten'})</span>`;
                }
                return `<div style="display:inline-block;padding:3px 8px;margin:2px;border-radius:4px;font-size:12px;background:${hidden ? '#555' : 'var(--bg-lighter)'};color:${hidden ? '#999' : 'var(--text)'};border:1px solid var(--border);">${s}${hidden ? ' 👁️‍🗨️' : ''}${dimText}</div>`;
            }).join('');
            
            dialog.innerHTML = `
                <div style="display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid var(--border);">
                    <h3 style="margin:0;color:var(--text);font-size:16px;">ℹ️ ${isEn ? 'File Information' : 'Datei-Informationen'}</h3>
                    <button id="fileInfoClose" style="background:none;border:none;color:var(--text-muted);font-size:20px;cursor:pointer;padding:0 4px;">&times;</button>
                </div>
                <div style="padding:20px;overflow-y:auto;flex:1;">
                    <table style="width:100%;border-collapse:collapse;font-size:13px;">
                        <tr><td colspan="2" style="padding:6px 0 4px;font-weight:700;font-size:14px;color:var(--primary);border-bottom:1px solid var(--border);">📄 ${isEn ? 'File' : 'Datei'}</td></tr>
                        ${row(isEn ? 'File Name' : 'Dateiname', result.fileName)}
                        ${row(isEn ? 'Path' : 'Pfad', `<span style="font-size:11px;color:var(--text-muted);">${result.filePath}</span>`)}
                        ${row(isEn ? 'Size' : 'Größe', formatSize(result.fileSize) + ` <span style="color:var(--text-muted);font-size:11px;">(${result.fileSize.toLocaleString(locale)} Bytes, ${result.zipEntryCount} ${isEn ? 'ZIP entries' : 'ZIP-Einträge'})</span>`)}
                        ${row(isEn ? 'Created (File)' : 'Erstellt (Datei)', formatDate(result.created))}
                        ${row(isEn ? 'Modified (File)' : 'Geändert (Datei)', formatDate(result.modified))}
                        
                        <tr><td colspan="2" style="padding:14px 0 4px;font-weight:700;font-size:14px;color:var(--primary);border-bottom:1px solid var(--border);">👤 ${isEn ? 'Document Properties' : 'Dokument-Eigenschaften'}</td></tr>
                        ${row(isEn ? 'Creator' : 'Ersteller', result.creator || '—')}
                        ${row(isEn ? 'Last Modified By' : 'Zuletzt geändert von', result.lastModifiedBy || '—')}
                        ${row(isEn ? 'Created (Document)' : 'Erstellt (Dokument)', formatDate(result.createdDate))}
                        ${row(isEn ? 'Modified (Document)' : 'Geändert (Dokument)', formatDate(result.modifiedDate))}
                        ${row(isEn ? 'Title' : 'Titel', result.title || '')}
                        ${row(isEn ? 'Subject' : 'Betreff', result.subject || '')}
                        ${row(isEn ? 'Category' : 'Kategorie', result.category || '')}
                        ${row(isEn ? 'Keywords' : 'Stichwörter', result.keywords || '')}
                        ${row(isEn ? 'Company' : 'Firma', result.company || '')}
                        ${row(isEn ? 'Application' : 'Anwendung', (result.application || '—') + (result.appVersion ? ` (v${result.appVersion})` : ''))}
                        
                        <tr><td colspan="2" style="padding:14px 0 4px;font-weight:700;font-size:14px;color:var(--primary);border-bottom:1px solid var(--border);">📊 ${isEn ? 'Worksheets' : 'Arbeitsblätter'} (${result.sheets.length})</td></tr>
                        <tr><td colspan="2" style="padding:6px 0;">${sheetsInfo}</td></tr>
                        
                        <tr><td colspan="2" style="padding:14px 0 4px;font-weight:700;font-size:14px;color:var(--primary);border-bottom:1px solid var(--border);">🔧 Features</td></tr>
                        <tr><td colspan="2" style="padding:6px 0;">${features.length > 0 ? features.join(' ') : `<span style="color:var(--text-muted);">${isEn ? 'No special features detected' : 'Keine besonderen Features erkannt'}</span>`}</td></tr>
                        ${result.hasSharedStrings ? row('Shared Strings', `${result.sharedStringUniqueCount?.toLocaleString(locale) || '?'} ${isEn ? 'unique' : 'eindeutige'} / ${result.sharedStringCount?.toLocaleString(locale) || '?'} ${isEn ? 'total' : 'gesamt'}`) : ''}
                        
                        ${result.lockFile ? `
                        <tr><td colspan="2" style="padding:14px 0 4px;font-weight:700;font-size:14px;color:#f44336;border-bottom:1px solid var(--border);">🔒 ${isEn ? 'Lock File Detected' : 'Lock-Datei erkannt'}</td></tr>
                        ${row(isEn ? 'Lock File' : 'Lock-Datei', result.lockFile)}
                        ${row(isEn ? 'Locked By' : 'Gesperrt von', result.lockedByUser || (isEn ? 'Unknown' : 'Unbekannt'))}
                        <tr><td colspan="2" style="padding:4px 0;"><span style="font-size:11px;color:#ff9800;">⚠️ ${isEn ? 'This file may currently be edited by another user.' : 'Diese Datei wird möglicherweise gerade von einem anderen Benutzer bearbeitet.'}</span></td></tr>
                        ` : ''}
                    </table>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            // Close handlers
            dialog.querySelector('#fileInfoClose').onclick = () => overlay.remove();
            overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
            const escHandler = (e) => { if (e.key === 'Escape') { overlay.remove(); document.removeEventListener('keydown', escHandler); } };
            document.addEventListener('keydown', escHandler);
        }
        
        // ==================== New Month Modal Functions ====================
        function getSuggestedMonthFilename(templateName, referenceDate = new Date()) {
            let baseName = templateName || 'Vertragsliste.xlsx';
            baseName = baseName.replace(/\.(xlsx|xls)$/i, '');

            // Das Template-Präfix und einen eventuell bereits vorhandenen
            // Monats-/Tages-Suffix entfernen, bevor das neue Monatsende
            // vorangestellt wird.
            baseName = baseName
                .replace(/^template[_\-\s]*/i, '')
                .replace(/[_\-\s]*\d{4}-\d{2}(?:-\d{2})?$/, '')
                .replace(/^[_\-\s]+|[_\-\s]+$/g, '');

            const year = referenceDate.getFullYear();
            const monthIndex = referenceDate.getMonth();
            const lastDay = new Date(year, monthIndex + 1, 0);
            const datePrefix = [
                year,
                String(monthIndex + 1).padStart(2, '0'),
                String(lastDay.getDate()).padStart(2, '0')
            ].join('-');

            return `${datePrefix}_${baseName || 'Vertragsliste'}.xlsx`;
        }

        function openNewMonthModal() {
            if (!state.template.filePath && !state.template.data) {
                showStatus(elements.transferStatus, 'Bitte erst ein Template laden', 'error');
                return;
            }
            
            // Element direkt aus DOM holen (kann durch Sprachumschaltung ersetzt worden sein)
            const templateNameEl = document.getElementById('newMonthTemplateName');
            if (templateNameEl) {
                templateNameEl.textContent = state.template.name || '-';
            }
            
            elements.newMonthFilename.value = getSuggestedMonthFilename(state.template.name);
            
            elements.newMonthModal.classList.remove('hidden');
        }
        
        function closeNewMonthModal() {
            elements.newMonthModal.classList.add('hidden');
        }
        
        async function confirmNewMonth() {
            const filename = elements.newMonthFilename.value.trim();
            if (!filename) {
                showStatus(elements.transferStatus, 'Bitte einen Dateinamen eingeben', 'error');
                return;
            }
            
            const finalFilename = filename.endsWith('.xlsx') ? filename : filename + '.xlsx';
            
            if (state.template.filePath) {
                const savePath = await window.electronAPI.saveFileDialog({
                    title: 'Neue Monatsdatei speichern',
                    defaultPath: getWorkingDirectoryPath() ? (getWorkingDirectoryPath() + '/' + finalFilename) : finalFilename,
                    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
                });
                
                if (savePath) {
                    // Verwende copyExcelFile (wie in preload.js definiert)
                    const result = await window.electronAPI.copyExcelFile({
                        sourcePath: state.template.filePath,
                        targetPath: savePath
                    });
                    
                    if (result.success) {
                        closeNewMonthModal();
                        
                        // Neue Datei als Datei 2 laden
                        const readResult = await window.electronAPI.readExcelFile(savePath);
                        if (readResult.success) {
                            state.file2.name = readResult.fileName;
                            state.file2.filePath = savePath;
                            state.file2.sheets = readResult.sheets;
                            state.file2.workbook = { SheetNames: readResult.sheets };
                            
                            // Change-Request-Cache invalidieren (neues Verzeichnis)
                            invalidateChangeRequestCache();
                            
                            elements.selectSheet2.innerHTML = readResult.sheets.map(s => `<option value="${s}">${s}</option>`).join('');
                            elements.selectSheet2.disabled = false;
                            elements.file2Info.textContent = `✓ ${readResult.fileName}`;
                            elements.file2Info.classList.add('loaded');
                            
                            await loadSheet2Electron(readResult.sheets[0]);
                            
                            showStatus(elements.transferStatus, `✓ Neue Monatsdatei erstellt: ${readResult.fileName}`, 'success');
                        }
                    } else {
                        showStatus(elements.transferStatus, `Fehler: ${result.error}`, 'error');
                    }
                }
            }
        }
        
        // ==================== Flag/Comment Column Visibility ====================
        function isFlagEnabled() {
            return document.getElementById('enableFlagColumn')?.checked ?? true;
        }
        
        function isCommentEnabled() {
            return document.getElementById('enableCommentColumn')?.checked ?? true;
        }
        
        // Automatische Spaltenberechnung:
        // - Flag ist immer Spalte 1 (wenn aktiviert)
        // - Kommentar ist nach Flag (Spalte 1 oder 2)
        // - Daten beginnen nach Flag und Kommentar
        function getFlagColumn() {
            // Flag ist immer Spalte 1 (A)
            return 1;
        }
        
        function getCommentColumn() {
            // Kommentar kommt nach Flag:
            // - Wenn Flag aktiviert: Spalte 2 (B)
            // - Wenn Flag deaktiviert: Spalte 1 (A)
            return isFlagEnabled() ? 2 : 1;
        }
        
        function getDataStartColumn() {
            // Daten beginnen nach Flag und Kommentar:
            // - Beide aktiviert: Spalte 3 (C)
            // - Nur eines aktiviert: Spalte 2 (B)
            // - Beide deaktiviert: Spalte 1 (A)
            let startCol = 1;
            if (isFlagEnabled()) startCol++;
            if (isCommentEnabled()) startCol++;
            return startCol;
        }
        
        // Aktualisiert die Anzeige der automatischen Spaltenberechnung
        function updateColumnDisplays() {
            const flagEnabled = isFlagEnabled();
            const commentEnabled = isCommentEnabled();
            
            const flagDisplay = document.getElementById('flagColumnDisplay');
            const commentDisplay = document.getElementById('commentColumnDisplay');
            const startDisplay = document.getElementById('targetStartColumnDisplay');
            
            if (flagDisplay) {
                flagDisplay.textContent = flagEnabled ? '→ Spalte A' : '(deaktiviert)';
                flagDisplay.style.color = flagEnabled ? 'var(--excel-green)' : 'var(--text-muted)';
            }
            
            if (commentDisplay) {
                if (commentEnabled) {
                    const col = getCommentColumn();
                    commentDisplay.textContent = `→ Spalte ${String.fromCharCode(64 + col)}`;
                    commentDisplay.style.color = 'var(--excel-green)';
                } else {
                    commentDisplay.textContent = '(deaktiviert)';
                    commentDisplay.style.color = 'var(--text-muted)';
                }
            }
            
            if (startDisplay) {
                const startCol = getDataStartColumn();
                startDisplay.textContent = `Spalte ${String.fromCharCode(64 + startCol)} (automatisch berechnet)`;
            }
        }
        
        function getFlagValues() {
            const input = document.getElementById('flagValues')?.value || 'A,D,C,leer';
            return input.split(',').map(v => v.trim()).filter(v => v);
        }
        
        function getCommentPlaceholder() {
            return document.getElementById('commentPlaceholder')?.value || 'Freier Text...';
        }
        
        function updateFlagDropdownOptions() {
            const values = getFlagValues();
            
            // Standard-Labels für bekannte Flags
            const flagLabels = {
                'A': 'A (Add)',
                'D': 'D (Delete)',
                'C': 'C (Change)',
                'leer': 'Leerzeile'
            };
            
            // Optionen HTML generieren
            const optionsHtml = values.map(v => {
                const label = flagLabels[v] || v;
                return `<option value="${escapeHtml(v)}">${escapeHtml(label)}</option>`;
            }).join('');
            
            // Beide Flag-Dropdowns aktualisieren
            if (elements.transferFlag) {
                elements.transferFlag.innerHTML = optionsHtml;
            }
            if (elements.newRowFlag) {
                elements.newRowFlag.innerHTML = optionsHtml;
            }
        }
        
        function updateCommentPlaceholders() {
            const placeholder = getCommentPlaceholder();
            if (elements.transferComment) {
                elements.transferComment.placeholder = placeholder;
            }
            if (elements.newRowComment) {
                elements.newRowComment.placeholder = placeholder;
            }
        }
        
        function updateFlagCommentVisibility() {
            const flagEnabled = isFlagEnabled();
            const commentEnabled = isCommentEnabled();
            
            // Transfer Panel
            const transferFlagField = elements.transferFlag?.closest('.transfer-field');
            const transferCommentField = elements.transferComment?.closest('.transfer-field');
            
            if (transferFlagField) {
                transferFlagField.style.display = flagEnabled ? '' : 'none';
            }
            if (transferCommentField) {
                transferCommentField.style.display = commentEnabled ? '' : 'none';
            }
            
            // New Row Panel
            const newRowFlagField = elements.newRowFlag?.closest('.transfer-field');
            const newRowCommentField = elements.newRowComment?.closest('.transfer-field');
            
            if (newRowFlagField) {
                newRowFlagField.style.display = flagEnabled ? '' : 'none';
            }
            if (newRowCommentField) {
                newRowCommentField.style.display = commentEnabled ? '' : 'none';
            }
            
            // Modal Config Sections
            const flagConfig = document.getElementById('flagColumnConfig');
            const commentConfig = document.getElementById('commentColumnConfig');
            
            if (flagConfig) {
                flagConfig.style.display = flagEnabled ? '' : 'none';
            }
            if (commentConfig) {
                commentConfig.style.display = commentEnabled ? '' : 'none';
            }
            
            // Update Labels mit Spaltennummer
            const flagColumn = getFlagColumn();
            const commentColumn = getCommentColumn();
            const colLetters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
            
            if (transferFlagField) {
                const label = transferFlagField.querySelector('label');
                if (label) label.textContent = `Spalte ${colLetters[flagColumn]} - Flag:`;
            }
            if (transferCommentField) {
                const label = transferCommentField.querySelector('label');
                if (label) label.textContent = `Spalte ${colLetters[commentColumn]} - Kommentar:`;
            }
            if (newRowFlagField) {
                const label = newRowFlagField.querySelector('label');
                if (label) label.textContent = `Flag (Spalte ${colLetters[flagColumn]}):`;
            }
            if (newRowCommentField) {
                const label = newRowCommentField.querySelector('label');
                if (label) label.textContent = `Kommentar (Spalte ${colLetters[commentColumn]}):`;
            }
        }
        
        // ==================== Initialize App ====================
        async function initApp() {
            // Initialize theme and language first
            setTheme(currentTheme);
            setLanguage(currentLanguage);
            document.getElementById('selectLanguage').value = currentLanguage;
            document.getElementById('selectTheme').value = currentTheme;
            document.getElementById('selectLanguage').onchange = (e) => setLanguage(e.target.value);
            document.getElementById('selectTheme').onchange = (e) => setTheme(e.target.value);
            
            // Excel Engine Auswahl
            const engineSelect = document.getElementById('selectExcelEngine');
            const savedEngine = localStorage.getItem('excelSyncEngine') || 'auto';
            if (engineSelect) {
                engineSelect.value = savedEngine;
                engineSelect.onchange = (e) => {
                    localStorage.setItem('excelSyncEngine', e.target.value);
                    updateEngineStatus();
                };
            }
            
            // Excel Engine Status anzeigen
            async function updateEngineStatus() {
                const statusIcon = document.getElementById('engineStatusIcon');
                const statusText = document.getElementById('engineStatusText');
                const engineSelect = document.getElementById('selectExcelEngine');
                
                if (!statusIcon || !statusText) return;
                
                const selectedEngine = engineSelect ? engineSelect.value : 'auto';
                
                try {
                    const status = await window.electronAPI.checkExcelAvailable();
                    
                    if (selectedEngine === 'openpyxl') {
                        // Manuell openpyxl erzwungen
                        statusIcon.textContent = '📦';
                        statusText.innerHTML = '<span style="color: var(--info);">openpyxl</span> - Manuell gewählt';
                    } else if (selectedEngine === 'xlwings') {
                        // Manuell xlwings erzwungen
                        if (status.excelAvailable) {
                            statusIcon.textContent = '🔴';
                            statusText.innerHTML = '<span style="color: var(--success);">xlwings</span> - Live-Modus aktiv';
                        } else {
                            statusIcon.textContent = '⚠️';
                            statusText.innerHTML = '<span style="color: var(--warning);">xlwings</span> - Excel nicht gefunden!';
                        }
                    } else {
                        // Auto-Modus
                        if (status.excelAvailable) {
                            statusIcon.textContent = '✓';
                            statusText.innerHTML = '<span style="color: var(--success);">Auto: xlwings</span> - Excel erkannt';
                        } else {
                            statusIcon.textContent = '📦';
                            statusText.innerHTML = '<span style="color: var(--info);">Auto: openpyxl</span> - Kein Excel';
                        }
                    }
                } catch (error) {
                    statusIcon.textContent = '⚠️';
                    statusText.textContent = 'Status unbekannt';
                }
            }
            
            // Die Prüfung startet unter Windows gegebenenfalls Excel. Erst nach
            // dem ersten sichtbaren Frame ausführen, damit sie den App-Start
            // nicht mit dem Renderer konkurrieren lässt.
            const initializeEngineStatus = async () => {
                try {
                    const engineSetting = localStorage.getItem('excelSyncEngine') || 'auto';
                    if (engineSetting !== 'openpyxl') {
                        const status = await window.electronAPI.checkExcelAvailable();
                        if (status && status.excelAvailable) {
                            explorerState.engineMode = 'live';
                            updateLiveModeIndicator();
                        } else {
                            explorerState.engineMode = 'openpyxl';
                        }
                    } else {
                        explorerState.engineMode = 'openpyxl';
                    }
                } catch (e) {
                    explorerState.engineMode = 'openpyxl';
                }
                await updateEngineStatus();
            };
            setTimeout(() => { void initializeEngineStatus(); }, 1000);
            
            // Flag/Comment Column Options
            const flagCheckbox = document.getElementById('enableFlagColumn');
            const commentCheckbox = document.getElementById('enableCommentColumn');
            const flagValuesInput = document.getElementById('flagValues');
            const commentPlaceholderInput = document.getElementById('commentPlaceholder');
            
            // Load saved preferences
            const savedFlagEnabled = localStorage.getItem('excelSyncEnableFlag');
            const savedCommentEnabled = localStorage.getItem('excelSyncEnableComment');
            const savedFlagValues = localStorage.getItem('excelSyncFlagValues');
            const savedCommentPlaceholder = localStorage.getItem('excelSyncCommentPlaceholder');
            
            if (savedFlagEnabled !== null) {
                flagCheckbox.checked = savedFlagEnabled === 'true';
            }
            if (savedCommentEnabled !== null) {
                commentCheckbox.checked = savedCommentEnabled === 'true';
            }
            if (savedFlagValues) {
                flagValuesInput.value = savedFlagValues;
            }
            if (savedCommentPlaceholder) {
                commentPlaceholderInput.value = savedCommentPlaceholder;
            }
            
            // Apply initial settings
            updateFlagDropdownOptions();
            updateCommentPlaceholders();
            updateFlagCommentVisibility();
            updateColumnDisplays();
            
            flagCheckbox.onchange = () => {
                localStorage.setItem('excelSyncEnableFlag', flagCheckbox.checked);
                updateFlagCommentVisibility();
                updateColumnDisplays();
                renderMappingList();
            };
            
            commentCheckbox.onchange = () => {
                localStorage.setItem('excelSyncEnableComment', commentCheckbox.checked);
                updateFlagCommentVisibility();
                updateColumnDisplays();
                renderMappingList();
            };
            
            flagValuesInput.onchange = () => {
                localStorage.setItem('excelSyncFlagValues', flagValuesInput.value);
                updateFlagDropdownOptions();
            };
            
            commentPlaceholderInput.onchange = () => {
                localStorage.setItem('excelSyncCommentPlaceholder', commentPlaceholderInput.value);
                updateCommentPlaceholders();
            };
            
            // Globaler Keyboard-Handler für Undo/Redo
            document.addEventListener('keydown', (e) => {
                // Strg+Z (Windows/Linux) oder Cmd+Z (Mac)
                if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
                    // Prüfe welches Modal aktiv ist
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    
                    if (explorerOpen) {
                        // Bei aktiver Live-Session: explorerUndo() nutzen (wird weiter unten aufgerufen)
                        // NICHT zusätzlich undoExplorer() aufrufen → sonst Doppel-Undo + Race-Condition
                        if (explorerState.liveSessionActive && explorerState.liveSessionReady) {
                            // Wird vom zweiten Ctrl+Z Handler (explorerUndo) behandelt
                        } else {
                            e.preventDefault();
                            if (undoExplorer()) {
                                showUndoRedoFeedback('Rückgängig');
                            }
                        }
                    } else {
                        e.preventDefault();
                        if (undoSearch()) {
                            showUndoRedoFeedback('Rückgängig');
                        }
                    }
                }
                
                // Strg+Y oder Strg+Shift+Z (Redo)
                if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
                    e.preventDefault();
                    
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    
                    if (explorerOpen) {
                        if (redoExplorer()) {
                            showUndoRedoFeedback('Wiederherstellen');
                        }
                    } else {
                        if (redoSearch()) {
                            showUndoRedoFeedback('Wiederherstellen');
                        }
                    }
                }
                
                // Strg+F → Fokus auf Suchfeld
                if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    
                    if (explorerOpen) {
                        e.preventDefault();
                        elements.explorerSearch.focus();
                        elements.explorerSearch.select();
                    } else if (!document.querySelector('.modal-backdrop:not(.hidden)')) {
                        // Nur wenn kein Modal offen ist
                        e.preventDefault();
                        elements.searchInput.focus();
                        elements.searchInput.select();
                    }
                }
                
                // Strg+C → Zellen kopieren (im Explorer mit Zell-Auswahl)
                if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    if (explorerOpen && explorerState.selectedCells.size > 0) {
                        // Nicht verhindern wenn Text in einem contentEditable selektiert ist
                        const activeEl = document.activeElement;
                        if (activeEl && activeEl.isContentEditable) {
                            const sel = window.getSelection();
                            if (sel && sel.toString().length > 0) return; // Browser-Copy erlauben
                        }
                        e.preventDefault();
                        copySelectedCellsWithFormat();
                    }
                }
                
                // Strg+V → Zellen einfügen (im Explorer mit Zell-Auswahl)
                if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    if (explorerOpen && explorerState.selectedCells.size > 0) {
                        // Nicht verhindern wenn in contentEditable
                        const activeEl = document.activeElement;
                        if (activeEl && activeEl.isContentEditable) return;
                        e.preventDefault();
                        // Wenn Format-Clipboard vorhanden UND kein CF → mit Formatierung einfügen
                        if (copiedCellsWithFormat && !explorerState.hasConditionalFormatting) {
                            pasteSelectedCellsWithFormat();
                        } else {
                            pasteToSelectedCells();
                        }
                    }
                }
                
                // Strg+Z → Undo (im Explorer mit aktiver Live-Session)
                if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    if (explorerOpen && explorerState.liveSessionActive && explorerState.liveSessionReady) {
                        // Nicht verhindern wenn in contentEditable (Text-Editing)
                        const activeEl = document.activeElement;
                        if (activeEl && activeEl.isContentEditable) return;
                        e.preventDefault();
                        explorerUndo();
                    }
                }
                
                // Strg+S → Warteschlange exportieren
                if ((e.ctrlKey || e.metaKey) && e.key === 's') {
                    e.preventDefault();
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    
                    if (explorerOpen) {
                        // Im Explorer: Exportieren
                        if (explorerState.data.length > 0) {
                            exportExplorerData();
                        }
                    } else if (state.transferQueue.length > 0) {
                        // Warteschlange exportieren
                        transferQueueToExcel();
                    }
                }
                
                // Strg+Enter → Direkt übertragen
                if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                    e.preventDefault();
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    
                    if (!explorerOpen && state.selectedRows.length > 0) {
                        transferSelectedDirect();
                    }
                }
                
                // F5 → Suche wiederholen / Daten neu laden
                if (e.key === 'F5') {
                    e.preventDefault();
                    const explorerOpen = !elements.dataExplorerModal.classList.contains('hidden');
                    
                    if (explorerOpen && explorerState.selectedSheet) {
                        loadExplorerSheet(explorerState.selectedSheet);
                        showUndoRedoFeedback('Neu geladen');
                    } else if (elements.searchInput.value) {
                        search();
                        showUndoRedoFeedback('Suche aktualisiert');
                    }
                }
                
                // Escape → Modal schließen
                if (e.key === 'Escape') {
                    // Prüfe welche Modals offen sind und schließe das oberste
                    if (!elements.dataExplorerModal.classList.contains('hidden')) {
                        closeDataExplorer();
                    } else if (!document.getElementById('helpModal').classList.contains('hidden')) {
                        document.getElementById('helpModal').classList.add('hidden');
                        document.body.classList.remove('modal-open');
                    } else if (!document.getElementById('mappingModal').classList.contains('hidden')) {
                        closeMappingModal();
                    } else if (!document.getElementById('newMonthModal').classList.contains('hidden')) {
                        closeNewMonthModal();
                    } else if (!document.getElementById('licenseModal').classList.contains('hidden')) {
                        document.getElementById('licenseModal').classList.add('hidden');
                        document.body.classList.remove('modal-open');
                    } else if (!elements.newRowPanel.classList.contains('hidden')) {
                        closeNewRowPanel();
                    }
                }
            });
            
            try { await initDB(); } catch (e) { console.warn('IndexedDB nicht verfügbar'); }
            loadConfig(); updateMappingPreview(); checkReadyState();
            
            // Auto-Save starten und Recovery prüfen
            startAutoSave();
            setTimeout(() => checkAutoSaveRecovery(), 500); // Kurze Verzögerung für UI-Initialisierung
            
            // "Öffnen mit..." Handler - Datei vom Betriebssystem empfangen
            window.electronAPI.onOpenFile(async (filePath) => {
                console.log('[FileOpen] Datei per "Öffnen mit..." empfangen:', filePath);
                // Datenexplorer öffnen
                await openDataExplorer();
                // Datei laden
                await loadExplorerFileByPath(filePath);
            });
            
            // Prüfung beim Schließen der App
            window.electronAPI.onBeforeClose(() => {
                // Bei normalem Schließen: Recovery-Daten löschen (nur bei Crash sollen sie bleiben)
                clearExplorerRecoveryData();
                
                if (state.transferQueue.length > 0) {
                    const msg = `Es befinden sich noch ${state.transferQueue.length} Einträge in der Warteschlange.\n\nWirklich schließen? Ungespeicherte Daten gehen verloren!`;
                    if (confirm(msg)) {
                        window.electronAPI.confirmClose(true);
                    } else {
                        window.electronAPI.confirmClose(false);
                    }
                } else {
                    window.electronAPI.confirmClose(true);
                }
            });
            
            // Search Event Listeners - WICHTIG!
            elements.searchInput.addEventListener('keydown', (e) => {
                const dropdown = document.getElementById('searchHistoryDropdown');
                const isDropdownOpen = dropdown && dropdown.classList.contains('show');
                
                if (e.key === 'Enter') {
                    e.preventDefault();
                    hideSearchHistoryDropdown();
                    search();
                } else if (e.key === 'ArrowDown') {
                    if (!isDropdownOpen) {
                        showSearchHistoryDropdown();
                    }
                    navigateSearchHistory('down');
                    e.preventDefault();
                } else if (e.key === 'ArrowUp') {
                    navigateSearchHistory('up');
                    e.preventDefault();
                } else if (e.key === 'Escape') {
                    hideSearchHistoryDropdown();
                }
            });
            
            elements.searchInput.addEventListener('focus', () => {
                showSearchHistoryDropdown();
            });
            
            elements.searchInput.addEventListener('input', debounce(() => {
                renderSearchHistoryDropdown(elements.searchInput.value);
                const dropdown = document.getElementById('searchHistoryDropdown');
                const filtered = getSearchHistory().filter(item => 
                    !elements.searchInput.value || item.term.toLowerCase().includes(elements.searchInput.value.toLowerCase())
                );
                if (filtered.length > 0) {
                    dropdown.classList.add('show');
                }
            }, 150));
            
            // Dropdown schließen bei Klick außerhalb
            document.addEventListener('click', (e) => {
                if (!e.target.closest('.search-wrapper')) {
                    hideSearchHistoryDropdown();
                }
            });
            
            elements.btnSearch.onclick = search;
            
            elements.btnNewRow.onclick = openNewRowPanel;
            elements.btnCloseNewRow.onclick = closeNewRowPanel;
            elements.btnAddNewRowToQueue.onclick = addNewRowToQueue;
            elements.btnTransferNewRowDirect.onclick = transferNewRowDirect;
            document.getElementById('btnAddEmptyRow').onclick = addEmptyRowToQueue;
            elements.btnAddToQueue.onclick = addToQueue;
            elements.btnTransferDirect.onclick = transferSelectedDirect;
            elements.btnSelectAll.onclick = () => selectAllRows(true);
            elements.btnDeselectAll.onclick = () => selectAllRows(false);
            elements.btnClearQueue.onclick = clearQueue;
            elements.btnConfigMapping.onclick = openMappingModal;
            document.getElementById('btnCloseMappingModal').onclick = closeMappingModal;
            document.getElementById('btnCancelMapping').onclick = closeMappingModal;
            document.getElementById('btnSaveMapping').onclick = saveMapping;
            document.getElementById('btnAddMapping').onclick = addMappingColumn;
            
            // Data Explorer Event Listeners
            elements.btnDataExplorer.onclick = openDataExplorer;
            elements.btnCloseExplorerX.onclick = closeDataExplorer;
            elements.btnCloseExplorerFooter.onclick = closeDataExplorer;
            elements.btnExplorerFullscreen.onclick = toggleExplorerFullscreen;
            elements.btnExplorerOpenFile.onclick = loadExplorerFile;
            
            // Drag & Drop Event Listeners for Explorer
            setupExplorerDropZone();
            elements.explorerSheetSelect.onchange = (e) => loadExplorerSheet(e.target.value);
            
            // Debounced Explorer-Suche (300ms Verzögerung)
            const debouncedExplorerSearch = debounce(() => {
                explorerState.searchTerm = elements.explorerSearch.value;
                filterExplorerData();
            }, 300);
            
            elements.explorerSearch.oninput = debouncedExplorerSearch;
            elements.explorerSearch.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    // Bei Enter sofort suchen (ohne Debounce)
                    explorerState.searchTerm = e.target.value;
                    filterExplorerData();
                }
            };
            elements.btnExplorerSearch.onclick = () => {
                explorerState.searchTerm = elements.explorerSearch.value;
                filterExplorerData();
            };
            elements.btnToggleColumns.onclick = toggleColumnPanel;
            elements.btnShowAllColumns.onclick = showAllExplorerColumns;
            elements.btnHideAllColumns.onclick = hideAllExplorerColumns;
            elements.btnAddExplorerFilter.onclick = addExplorerFilter;
            elements.btnClearExplorerFilters.onclick = clearExplorerFilters;
            elements.btnExplorerExport.onclick = exportExplorerData;
            
            // Schließen-Buttons (✕) für Spalten- und Filter-Panel
            const btnCloseColumnPanel = document.getElementById('btnCloseColumnPanel');
            if (btnCloseColumnPanel) {
                btnCloseColumnPanel.onclick = function() {
                    toggleColumnPanel();
                    saveCurrentSheetToCache();
                };
            }
            const btnCloseFilterPanel = document.getElementById('btnCloseFilterPanel');
            if (btnCloseFilterPanel) {
                btnCloseFilterPanel.onclick = toggleFilterPanel;
            }
            
            // Undo-Button
            const btnUndo = document.getElementById('btnExplorerUndo');
            if (btnUndo) {
                btnUndo.onclick = explorerUndo;
            }
            
            // Filter-Bereich Toggle (interner Toggle im Panel)
            const filterSectionToggle = document.getElementById('filterSectionToggle');
            if (filterSectionToggle) {
                filterSectionToggle.onclick = toggleFilterSection;
            }
            
            // Filter-Panel Button in Toolbar
            const btnToggleFilterPanel = document.getElementById('btnToggleFilterPanel');
            if (btnToggleFilterPanel) {
                btnToggleFilterPanel.onclick = toggleFilterPanel;
            }
            
            // Filter an Excel senden Button
            const btnSyncFilters = document.getElementById('btnSyncFiltersToExcel');
            if (btnSyncFilters) {
                btnSyncFilters.onclick = () => {
                    syncFiltersToExcel();
                };
            }
            
            // Live Session Toggle
            // Live-Session wird jetzt automatisch aktiviert - kein manueller Toggle mehr nötig
            
            // Data Join Event Listeners
            elements.btnDataJoin.onclick = openDataJoinModal;
            elements.btnCloseDataJoin.onclick = closeDataJoinModal;
            elements.btnCancelDataJoin.onclick = closeDataJoinModal;
            elements.btnJoinSelectFile.onclick = loadDataJoinSourceFile;
            elements.joinSourceSheet.onchange = (e) => loadDataJoinSourceSheet(e.target.value);
            elements.joinTargetKeyColumn.onchange = () => {
                dataJoinState.previewCalculated = false;
                elements.joinPreviewContainer.style.display = 'none';
                updateJoinButtons();
            };
            elements.joinSourceKeyColumn.onchange = () => {
                dataJoinState.previewCalculated = false;
                elements.joinPreviewContainer.style.display = 'none';
                updateJoinButtons();
            };
            elements.btnPreviewDataJoin.onclick = calculateDataJoinPreview;
            elements.btnExecuteDataJoin.onclick = executeDataJoin;
            
            // Serial-Check
            const btnSerialCheck = document.getElementById('btnSerialCheck');
            if (btnSerialCheck) btnSerialCheck.onclick = openSerialCheckModal;
            
            // Value-Count
            const btnValueCount = document.getElementById('btnValueCount');
            if (btnValueCount) btnValueCount.onclick = openValueCountModal;
            
            // Data Join Drag & Drop Zone
            const joinDropZone = document.getElementById('joinDropZone');
            if (joinDropZone) {
                joinDropZone.addEventListener('click', loadDataJoinSourceFile);
                
                joinDropZone.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    joinDropZone.style.borderColor = 'var(--primary)';
                    joinDropZone.style.background = 'rgba(0, 122, 204, 0.1)';
                });
                
                joinDropZone.addEventListener('dragleave', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    joinDropZone.style.borderColor = 'var(--border)';
                    joinDropZone.style.background = 'var(--bg-dark)';
                });
                
                joinDropZone.addEventListener('drop', async (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    joinDropZone.style.borderColor = 'var(--border)';
                    joinDropZone.style.background = 'var(--bg-dark)';
                    
                    const files = e.dataTransfer.files;
                    if (files.length > 0) {
                        const file = files[0];
                        const ext = file.name.split('.').pop().toLowerCase();
                        if (ext === 'xlsx' || ext === 'xls') {
                            // Dateipfad über Electron API abrufen (contextIsolation-sicher)
                            const filePath = window.electronAPI.getPathForFile(file);
                            if (!filePath) {
                                showNotification('Dateipfad konnte nicht ermittelt werden', 'error');
                                return;
                            }
                            await loadDataJoinSourceFromPath(filePath);
                        } else {
                            showNotification('Bitte nur Excel-Dateien (.xlsx, .xls) ablegen', 'warning');
                        }
                    }
                });
            }
            
            // Sheet Management Event Listeners
            document.getElementById('btnSheetManage').onclick = openSheetManageModal;
            document.getElementById('btnFileInfo').onclick = showFileInfoModal;
            document.getElementById('btnCloseSheetManage').onclick = closeSheetManageModal;
            document.getElementById('btnSheetManageClose').onclick = closeSheetManageModal;
            document.getElementById('btnSheetAdd').onclick = addNewSheet;
            document.getElementById('btnSheetRename').onclick = renameSelectedSheet;
            document.getElementById('btnSheetClone').onclick = cloneSelectedSheet;
            document.getElementById('btnSheetMoveUp').onclick = () => moveSelectedSheet('up');
            document.getElementById('btnSheetMoveDown').onclick = () => moveSelectedSheet('down');
            document.getElementById('btnSheetDelete').onclick = deleteSelectedSheet;
            document.getElementById('btnSheetToggleVisibility').onclick = toggleSheetVisibility;
            
            // Excel Ein-/Ausblenden Button
            elements.btnToggleExcel.onclick = toggleExcelVisibility;
            if (elements.btnToggleExcelInteractive) {
                elements.btnToggleExcelInteractive.onclick = toggleExcelInteractive;
            }
            document.getElementById('btnClosePreviewX').onclick = closeExplorerPreview;
            document.getElementById('btnClosePreview').onclick = closeExplorerPreview;
            
            // Row Move Toolbar Event Listeners
            document.getElementById('btnExecuteMove').onclick = executeRowMove;
            document.getElementById('btnHideSelectedRows').onclick = hideSelectedRows;
            document.getElementById('btnDeleteSelectedRows').onclick = deleteSelectedRows;
            document.getElementById('btnClearRowSelection').onclick = clearRowSelection;
            
            // Pagination Event Listeners (Data Explorer)
            elements.btnExplorerFirstPage.onclick = () => explorerGoToPage(1);
            elements.btnExplorerPrevPage.onclick = () => explorerGoToPage(explorerState.currentPage - 1);
            elements.btnExplorerNextPage.onclick = () => explorerGoToPage(explorerState.currentPage + 1);
            elements.btnExplorerLastPage.onclick = () => explorerGoToPage(Math.ceil(explorerState.filteredData.length / explorerState.pageSize));
            elements.explorerPageSize.onchange = (e) => explorerChangePageSize(e.target.value);
            
            // Pagination Event Listeners (Suchergebnisse)
            document.getElementById('btnSearchFirstPage').onclick = () => searchGoToPage(1);
            document.getElementById('btnSearchPrevPage').onclick = () => searchGoToPage(state.searchPagination.currentPage - 1);
            document.getElementById('btnSearchNextPage').onclick = () => searchGoToPage(state.searchPagination.currentPage + 1);
            document.getElementById('btnSearchLastPage').onclick = () => searchGoToPage(Math.ceil(state.searchResults.length / state.searchPagination.pageSize));
            document.getElementById('searchPageSize').onchange = (e) => searchChangePageSize(e.target.value);
            
            // New Month Modal Event Listeners
            elements.btnNewMonth.onclick = openNewMonthModal;
            document.getElementById('btnCloseNewMonthModal').onclick = closeNewMonthModal;
            document.getElementById('btnCancelNewMonth').onclick = closeNewMonthModal;
            document.getElementById('btnConfirmNewMonth').onclick = confirmNewMonth;
            
            // Create Template Modal Event Listeners
            document.getElementById('btnCloseCreateTemplateModal').onclick = closeCreateTemplateModal;
            document.getElementById('btnCancelCreateTemplate').onclick = closeCreateTemplateModal;
            document.getElementById('btnConfirmCreateTemplate').onclick = confirmCreateTemplate;
            document.getElementById('btnSelectAllSheets').onclick = () => {
                document.querySelectorAll('.template-sheet-checkbox').forEach(cb => cb.checked = true);
            };
            document.getElementById('btnDeselectAllSheets').onclick = () => {
                document.querySelectorAll('.template-sheet-checkbox').forEach(cb => cb.checked = false);
            };
            
            // Help Modal Event Listeners
            elements.btnHelp.onclick = () => elements.helpModal.classList.remove('hidden');
            document.getElementById('btnCloseHelpModal').onclick = () => elements.helpModal.classList.add('hidden');
            document.getElementById('btnCloseHelp').onclick = () => elements.helpModal.classList.add('hidden');
            
            // License Modal Event Listeners
            document.getElementById('btnLicense').onclick = () => document.getElementById('licenseModal').classList.remove('hidden');
            document.getElementById('btnCloseLicenseModal').onclick = () => document.getElementById('licenseModal').classList.add('hidden');
            document.getElementById('btnCloseLicense').onclick = () => document.getElementById('licenseModal').classList.add('hidden');
            
            // Security Logs Modal Event Listeners
            document.getElementById('btnSecurityLogs').onclick = openSecurityLogsModal;
            document.getElementById('btnCloseSecurityLogsModal').onclick = closeLogsModal;
            document.getElementById('btnCloseSecurityLogs').onclick = closeLogsModal;
            document.getElementById('btnRefreshLogs').onclick = loadSecurityLogs;
            document.getElementById('btnVerifyLogs').onclick = verifySecurityLogs;
            document.getElementById('btnClearLogs').onclick = clearSecurityLogs;
            document.getElementById('logsLevelFilter').onchange = filterSecurityLogs;
            document.getElementById('logsSearchFilter').oninput = filterSecurityLogs;
            
            // Tab-Switching für Logs
            document.getElementById('tabLocalLogs').onclick = () => switchLogsTab('local');
            document.getElementById('tabNetworkLogs').onclick = () => switchLogsTab('network');
            
            // Network Logs Event Listeners (innerhalb des Security Modals)
            document.getElementById('btnRefreshNetworkLogs').onclick = loadNetworkLogs;
            document.getElementById('networkLogsHostFilter').onchange = filterNetworkLogs;
            document.getElementById('networkLogsSearchFilter').oninput = filterNetworkLogs;
            
            // Sidebar Toggle
            document.getElementById('sidebarToggle').onclick = () => {
                document.getElementById('sidebar').classList.toggle('collapsed');
            };
            
            // Keyboard shortcuts
            document.onkeydown = (e) => {
                if (e.key === 'F1') {
                    e.preventDefault();
                    elements.helpModal.classList.toggle('hidden');
                }
                if (e.key === 'Escape') {
                    elements.helpModal.classList.add('hidden');
                    elements.mappingModal.classList.add('hidden');
                    elements.newMonthModal.classList.add('hidden');
                    elements.createTemplateModal.classList.add('hidden');
                    document.getElementById('licenseModal').classList.add('hidden');
                    closeLogsModal();
                    closeDataExplorer();
                }
            };
            
            // Electron-Modus: Verwende Electron-API für Dateioperationen
            elements.btnLoadFile1.onclick = loadFile1Electron;
            elements.btnLoadFile2.onclick = loadFile2Electron;
            elements.btnLoadTemplate.onclick = loadTemplateElectron;
            elements.btnCreateTemplate.onclick = createTemplateFromSourceElectron;
            elements.selectSheet1.onchange = (e) => loadSheet1Electron(e.target.value);
            elements.selectSheet2.onchange = (e) => loadSheet2Electron(e.target.value);
            elements.btnImportConfig.onclick = loadConfigFromAppDirOrDialog;
            elements.btnExportConfig.onclick = exportConfig;
            elements.btnExportPS.onclick = showDiffPreview;  // Zeigt zuerst Vorschau, dann Export via Modal
            elements.btnPreviewTransfer.onclick = showDiffPreview;
            
            // Arbeitsordner Event Handler
            elements.btnSelectWorkingDir.onclick = selectWorkingDirectory;
            elements.btnClearWorkingDir.onclick = clearWorkingDirectory;
            
            // Arbeitsordner aus localStorage laden
            loadWorkingDirectoryFromStorage();
            
            // Diff-Vorschau Modal Event Handler
            document.getElementById('btnCloseDiffModal').onclick = closeDiffPreview;
            document.getElementById('btnCancelDiff').onclick = closeDiffPreview;
            document.getElementById('btnConfirmTransfer').onclick = confirmTransferFromDiff;
            
            // Prüfe ob eine Datei per "Öffnen mit..." übergeben wurde
            // Falls ja: Config-Dateien (Quell-/Zieldatei/Template) NICHT laden → Startup beschleunigen
            let _hasStartupFile = false;
            try {
                _hasStartupFile = !!(await window.electronAPI.hasStartupFile());
            } catch (e) { /* OK, API evtl. nicht vorhanden */ }
            
            // Automatisch config.json aus Programmordner oder Arbeitsordner laden beim Start
            console.log('[Config] Suche automatisch nach config.json...');
            console.log('[Config] Arbeitsordner:', state.workingDirectory || '(nicht gesetzt)');
            if (_hasStartupFile) {
                console.log('[Config] Startup-Datei erkannt → Überspringe Laden von Quell-/Zieldatei/Template');
            }
            try {
                const autoResult = await window.electronAPI.loadConfigFromAppDir(state.workingDirectory);
                console.log('[Config] loadConfigFromAppDir Ergebnis:', autoResult);
                
                if (autoResult.success && autoResult.config) {
                    await applyLoadedConfig(autoResult.config, _hasStartupFile);
                    console.log('[Config] config.json automatisch geladen:', autoResult.path);
                    console.log('[Config] Benutzerprofil:', autoResult.userId || '(unbekannt)');
                    console.log('[Config] Hat Benutzer-Abschnitt:', autoResult.hasUserSection);
                    console.log('[Config] Legacy-Format:', autoResult.isLegacyFormat);
                    
                    // Zeige Computer-spezifische Info
                    let statusMsg = `✓ config.json geladen: ${autoResult.path}`;
                    if (autoResult.userId && !autoResult.isLegacyFormat) {
                        statusMsg = autoResult.hasUserSection
                            ? `✓ Config für Benutzer „${autoResult.userId}“ geladen`
                            : `✓ Config geladen (Standard, kein Abschnitt für Benutzer „${autoResult.userId}“)`;
                    }
                    showStatus(elements.transferStatus, statusMsg, 'success');
                } else {
                    console.log('[Config] Keine config.json gefunden.');
                    if (autoResult.searchedPaths) {
                        console.log('[Config] Gesuchte Pfade:', autoResult.searchedPaths);
                    }
                }
            } catch (e) {
                console.error('[Config] Fehler beim automatischen Laden:', e);
            }
        }
        
        // ==================== SECURITY LOGS FUNCTIONS ====================
        
        let securityLogsCache = [];
        let currentLogsTab = 'local'; // 'local' oder 'network'
        
        async function openSecurityLogsModal() {
            document.getElementById('securityLogsModal').classList.remove('hidden');
            // Reset auf lokale Logs
            switchLogsTab('local');
        }
        
        function closeLogsModal() {
            document.getElementById('securityLogsModal').classList.add('hidden');
        }
        
        async function switchLogsTab(tab) {
            currentLogsTab = tab;
            const tabLocal = document.getElementById('tabLocalLogs');
            const tabNetwork = document.getElementById('tabNetworkLogs');
            const sectionLocal = document.getElementById('localLogsSection');
            const sectionNetwork = document.getElementById('networkLogsSection');
            const clearBtn = document.getElementById('btnClearLogs');
            
            if (tab === 'local') {
                // Lokale Logs Tab aktiv
                tabLocal.style.background = 'var(--primary)';
                tabLocal.style.color = 'white';
                tabLocal.style.borderColor = 'var(--primary)';
                tabNetwork.style.background = 'var(--bg-light)';
                tabNetwork.style.color = 'var(--text)';
                tabNetwork.style.borderColor = 'var(--border)';
                
                sectionLocal.classList.remove('hidden');
                sectionNetwork.classList.add('hidden');
                clearBtn.style.display = 'block';
                
                await loadSecurityLogs();
                await verifySecurityLogs();
            } else {
                // Netzwerk Logs Tab aktiv
                tabNetwork.style.background = 'var(--primary)';
                tabNetwork.style.color = 'white';
                tabNetwork.style.borderColor = 'var(--primary)';
                tabLocal.style.background = 'var(--bg-light)';
                tabLocal.style.color = 'var(--text)';
                tabLocal.style.borderColor = 'var(--border)';
                
                sectionNetwork.classList.remove('hidden');
                sectionLocal.classList.add('hidden');
                clearBtn.style.display = 'none'; // Netzwerk-Logs nicht löschbar
                
                await loadNetworkLogs();
            }
        }
        
        async function loadSecurityLogs() {
            try {
                const result = await window.electronAPI.getSecurityLogs({ fromFile: true, limit: 500 });
                
                if (result.success) {
                    securityLogsCache = result.entries;
                    document.getElementById('logsCount').textContent = result.totalCount;
                    document.getElementById('logsPath').textContent = result.logFilePath || '-';
                    renderSecurityLogs(result.entries);
                } else {
                    document.getElementById('securityLogsTableBody').innerHTML = 
                        `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--error);">Fehler: ${result.error}</td></tr>`;
                }
            } catch (e) {
                console.error('Fehler beim Laden der Security-Logs:', e);
            }
        }
        
        function renderSecurityLogs(entries) {
            const tbody = document.getElementById('securityLogsTableBody');
            
            if (!entries || entries.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--text-muted);">Keine Log-Einträge vorhanden</td></tr>`;
                return;
            }
            
            const levelColors = {
                'INFO': { bg: 'rgba(33, 150, 243, 0.2)', text: '#2196F3', icon: 'ℹ️' },
                'WARN': { bg: 'rgba(255, 152, 0, 0.2)', text: '#FF9800', icon: '⚠️' },
                'ERROR': { bg: 'rgba(244, 67, 54, 0.2)', text: '#F44336', icon: '❌' },
                'SECURITY': { bg: 'rgba(156, 39, 176, 0.2)', text: '#9C27B0', icon: '🛡️' }
            };
            
            tbody.innerHTML = entries.map(entry => {
                const level = levelColors[entry.level] || levelColors['INFO'];
                const timestamp = new Date(entry.timestamp).toLocaleString('de-DE');
                const details = entry.details ? Object.entries(entry.details)
                    .filter(([k, v]) => k !== 'pid')
                    .map(([k, v]) => `<span style="color: var(--text-muted);">${k}:</span> ${typeof v === 'object' ? JSON.stringify(v) : v}`)
                    .join(', ') : '-';
                
                // Signature-Check Icon
                const sigValid = entry.signature ? '✓' : '?';
                const sigColor = entry.signature ? 'var(--success)' : 'var(--text-muted)';
                
                return `
                    <tr style="border-bottom: 1px solid var(--border);">
                        <td style="padding: 8px; white-space: nowrap; font-family: monospace; font-size: 11px;">${timestamp}</td>
                        <td style="padding: 8px; text-align: center;">
                            <span style="background: ${level.bg}; color: ${level.text}; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: bold;">
                                ${level.icon} ${entry.level}
                            </span>
                        </td>
                        <td style="padding: 8px; font-weight: 500;">${entry.action}</td>
                        <td style="padding: 8px; font-size: 11px; max-width: 350px; overflow: hidden; text-overflow: ellipsis;" title="${details}">${details}</td>
                        <td style="padding: 8px; text-align: center; color: ${sigColor}; font-weight: bold;">${sigValid}</td>
                    </tr>
                `;
            }).join('');
        }
        
        function filterSecurityLogs() {
            const levelFilter = document.getElementById('logsLevelFilter').value;
            const searchFilter = document.getElementById('logsSearchFilter').value.toLowerCase();
            
            let filtered = securityLogsCache;
            
            if (levelFilter) {
                filtered = filtered.filter(e => e.level === levelFilter);
            }
            
            if (searchFilter) {
                filtered = filtered.filter(e => 
                    e.action.toLowerCase().includes(searchFilter) ||
                    JSON.stringify(e.details).toLowerCase().includes(searchFilter)
                );
            }
            
            renderSecurityLogs(filtered);
        }
        
        async function verifySecurityLogs() {
            const statusEl = document.getElementById('securityLogsStatus');
            const integrityBox = document.getElementById('securityLogsIntegrity');
            const errorsBox = document.getElementById('securityLogsErrors');
            const errorsList = document.getElementById('securityLogsErrorList');
            
            statusEl.textContent = 'Wird geprüft...';
            
            try {
                const result = await window.electronAPI.verifySecurityLogs();
                
                if (result.success) {
                    if (result.valid) {
                        statusEl.innerHTML = `<span style="color: var(--success); font-weight: bold;">✓ Alle ${result.verifiedEntries} Einträge verifiziert</span>`;
                        integrityBox.style.borderLeftColor = 'var(--success)';
                        errorsBox.classList.add('hidden');
                    } else {
                        statusEl.innerHTML = `<span style="color: var(--error); font-weight: bold;">⚠️ ${result.errors.length} Integritätsfehler</span>`;
                        integrityBox.style.borderLeftColor = 'var(--error)';
                        errorsBox.classList.remove('hidden');
                        errorsList.innerHTML = result.errors.map(err => `<li>${err}</li>`).join('');
                    }
                } else {
                    statusEl.innerHTML = `<span style="color: var(--error);">Fehler: ${result.error}</span>`;
                }
            } catch (e) {
                statusEl.innerHTML = `<span style="color: var(--error);">Fehler bei Verifikation</span>`;
                console.error('Verifikationsfehler:', e);
            }
        }
        
        async function clearSecurityLogs() {
            if (!confirm('Alle Security-Logs wirklich löschen?\n\nDies kann nicht rückgängig gemacht werden.')) {
                return;
            }
            
            try {
                const result = await window.electronAPI.clearSecurityLogs();
                
                if (result.success) {
                    await loadSecurityLogs();
                    await verifySecurityLogs();
                } else {
                    alert('Fehler beim Löschen: ' + result.error);
                }
            } catch (e) {
                console.error('Fehler beim Löschen der Logs:', e);
            }
        }
        
        // ==================== NETWORK LOGS FUNCTIONS ====================
        
        let networkLogsCache = [];
        let currentNetworkPath = null;
        
        /**
         * Prüft ob eine Datei kürzlich von einem anderen Rechner bearbeitet wurde
         * und zeigt ggf. eine Warnung an.
         * @param {string} filePath - Pfad zur Datei
         * @returns {{proceed: boolean}} - true wenn fortgefahren werden soll
         */
        async function checkAndWarnNetworkConflict(filePath) {
            try {
                const result = await window.electronAPI.checkNetworkConflict(filePath, 5);
                
                if (!result.success || !result.isNetworkPath) {
                    // Kein Netzlaufwerk oder Fehler - einfach fortfahren
                    return { proceed: true };
                }
                
                if (!result.conflict) {
                    // Kein Konflikt
                    return { proceed: true };
                }
                
                // Warnung zusammenstellen
                let message = '⚠️ Achtung: Möglicher Bearbeitungskonflikt!\n\n';
                
                if (result.activeLock) {
                    message += `Diese Datei ist möglicherweise bereits geöffnet:\n`;
                    message += `• Rechner: ${result.activeLock.hostname}\n`;
                    message += `• Seit: ${result.activeLock.ageMinutes} Minute(n)\n\n`;
                }
                
                if (result.recentActivity) {
                    message += `Diese Datei wurde kürzlich bearbeitet:\n`;
                    message += `• Rechner: ${result.recentActivity.hostname}\n`;
                    message += `• Aktion: ${result.recentActivity.action}\n`;
                    message += `• Vor: ${result.recentActivity.ageMinutes} Minute(n)\n\n`;
                }
                
                message += 'Wenn Sie die Datei gleichzeitig bearbeiten, können Änderungen verloren gehen.\n\n';
                message += 'Trotzdem öffnen?';
                
                const proceed = confirm(message);
                
                return { proceed };
            } catch (e) {
                console.error('Fehler bei Konfliktprüfung:', e);
                // Bei Fehler trotzdem fortfahren
                return { proceed: true };
            }
        }
        
        /**
         * Session-Locks beim Schließen der App entfernen
         */
        async function cleanupSessionLocks() {
            try {
                if (state.file1?.filePath) {
                    await window.electronAPI.removeSessionLock(state.file1.filePath);
                }
                if (state.file2?.filePath) {
                    await window.electronAPI.removeSessionLock(state.file2.filePath);
                }
            } catch (e) {
                console.error('Fehler beim Entfernen der Session-Locks:', e);
            }
        }
        
        // Session-Locks beim Verlassen der Seite aufräumen
        window.addEventListener('beforeunload', () => {
            cleanupSessionLocks();
            // Markiere sauberen Shutdown - AutoSave wird gelöscht
            markCleanShutdown();
        });
        
        async function loadNetworkLogs() {
            const tbody = document.getElementById('networkLogsTableBody');
            const countEl = document.getElementById('networkLogsCount');
            const pathEl = document.getElementById('networkLogsPath');
            const hostnameEl = document.getElementById('networkHostname');
            const hostFilterEl = document.getElementById('networkLogsHostFilter');
            
            // Ermittle den aktuellen Dateipfad (von file1 oder file2)
            const filePath = state.file1?.filePath || state.file2?.filePath;
            
            if (!filePath) {
                tbody.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--text-muted);">Bitte zuerst eine Datei laden</td></tr>`;
                countEl.textContent = '0';
                pathEl.textContent = '-';
                return;
            }
            
            try {
                // Hostname abrufen
                const networkInfo = await window.electronAPI.isNetworkPath(filePath);
                hostnameEl.textContent = networkInfo.hostname || '-';
                
                if (!networkInfo.isNetwork) {
                    tbody.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--text-muted);">Die aktuelle Datei liegt nicht auf einem Netzlaufwerk.<br>Netzwerk-Logs werden nur für Dateien auf Netzlaufwerken geführt.</td></tr>`;
                    countEl.textContent = '0';
                    pathEl.textContent = 'Kein Netzlaufwerk';
                    return;
                }
                
                // Netzwerk-Logs laden
                const result = await window.electronAPI.getNetworkLogs(filePath);
                
                if (result.success) {
                    networkLogsCache = result.entries;
                    currentNetworkPath = result.logFilePath;
                    countEl.textContent = result.totalCount;
                    pathEl.textContent = result.logFilePath || '-';
                    
                    // Host-Filter befüllen
                    const hosts = [...new Set(result.entries.map(e => e.hostname))];
                    hostFilterEl.innerHTML = '<option value="">Alle Rechner</option>' + 
                        hosts.map(h => `<option value="${h}">${h}</option>`).join('');
                    
                    renderNetworkLogs(result.entries);
                } else {
                    tbody.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--error);">Fehler: ${result.error}</td></tr>`;
                }
            } catch (e) {
                console.error('Fehler beim Laden der Netzwerk-Logs:', e);
                tbody.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--error);">Fehler beim Laden</td></tr>`;
            }
        }
        
        function renderNetworkLogs(entries) {
            const tbody = document.getElementById('networkLogsTableBody');
            
            if (!entries || entries.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" style="padding: 20px; text-align: center; color: var(--text-muted);">Keine Netzwerk-Log-Einträge vorhanden</td></tr>`;
                return;
            }
            
            const actionColors = {
                'EXCEL_FILE_SAVED': { bg: 'rgba(76, 175, 80, 0.2)', text: '#4CAF50', icon: '💾' },
                'EXCEL_EXPORT_SOURCE': { bg: 'rgba(33, 150, 243, 0.2)', text: '#2196F3', icon: '📤' },
                'EXCEL_EXPORT_TARGET': { bg: 'rgba(33, 150, 243, 0.2)', text: '#2196F3', icon: '📥' },
                'DATA_TRANSFER': { bg: 'rgba(156, 39, 176, 0.2)', text: '#9C27B0', icon: '🔀' },
                'DATA_TRANSFER_SOURCE': { bg: 'rgba(255, 152, 0, 0.2)', text: '#FF9800', icon: '📤' }
            };
            
            tbody.innerHTML = entries.map(entry => {
                const actionStyle = actionColors[entry.action] || { bg: 'rgba(158, 158, 158, 0.2)', text: '#9E9E9E', icon: '📋' };
                const timestamp = new Date(entry.timestamp).toLocaleString('de-DE');
                const details = entry.details ? Object.entries(entry.details)
                    .filter(([k, v]) => v !== null && v !== undefined)
                    .map(([k, v]) => `<span style="color: var(--text-muted);">${k}:</span> ${v}`)
                    .join(', ') : '-';
                
                return `
                    <tr style="border-bottom: 1px solid var(--border);">
                        <td style="padding: 8px; white-space: nowrap; font-family: monospace; font-size: 11px;">${timestamp}</td>
                        <td style="padding: 8px; font-family: monospace; font-size: 11px;">${entry.hostname}</td>
                        <td style="padding: 8px;">
                            <span style="background: ${actionStyle.bg}; color: ${actionStyle.text}; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: bold;">
                                ${actionStyle.icon} ${entry.action}
                            </span>
                        </td>
                        <td style="padding: 8px; font-weight: 500;">${entry.file || '-'}</td>
                        <td style="padding: 8px; font-size: 11px; max-width: 300px; overflow: hidden; text-overflow: ellipsis;" title="${details}">${details}</td>
                    </tr>
                `;
            }).join('');
        }
        
        function filterNetworkLogs() {
            const hostFilter = document.getElementById('networkLogsHostFilter').value;
            const searchFilter = document.getElementById('networkLogsSearchFilter').value.toLowerCase();
            
            let filtered = networkLogsCache;
            
            if (hostFilter) {
                filtered = filtered.filter(e => e.hostname === hostFilter);
            }
            
            if (searchFilter) {
                filtered = filtered.filter(e => 
                    e.action.toLowerCase().includes(searchFilter) ||
                    (e.file && e.file.toLowerCase().includes(searchFilter)) ||
                    JSON.stringify(e.details).toLowerCase().includes(searchFilter)
                );
            }
            
            renderNetworkLogs(filtered);
        }

        initApp();
