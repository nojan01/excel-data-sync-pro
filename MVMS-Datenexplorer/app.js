'use strict';

(function () {
    const fileInput = document.getElementById('fileInput');
    const searchInput = document.getElementById('searchInput');
    const addFilterButton = document.getElementById('addFilter');
    const clearFiltersButton = document.getElementById('clearFilters');
    const filtersContainer = document.getElementById('filters');
    const filterTemplate = document.getElementById('filterTemplate');
    const statusBox = document.getElementById('status');
    const tableHead = document.querySelector('#resultTable thead');
    const tableBody = document.querySelector('#resultTable tbody');
    const resultCount = document.getElementById('resultCount');
    const exportButton = document.getElementById('exportButton');
    const sheetSelect = document.getElementById('sheetSelect');
    const tableWrapper = document.getElementById('tableWrapper');
    const tableScrollTop = document.getElementById('tableScrollTop');
    const tableScrollTopContent = document.getElementById('tableScrollTopContent');
    const hiddenColumnsPanel = document.getElementById('hiddenColumnsPanel');
    const hiddenColumnsList = document.getElementById('hiddenColumnsList');
    const restoreAllColumnsButton = document.getElementById('restoreAllColumns');

    const DATE_COLUMNS = new Set([
        'end of current contract',
        "end of vendor's warranty",
        'requested mvms start',
        'requested mvms end'
    ]);

    const OPERATIONS = {
        contains: (value, testValue) => value.includes(testValue),
        equals: (value, testValue) => value === testValue,
        startsWith: (value, testValue) => value.startsWith(testValue),
        endsWith: (value, testValue) => value.endsWith(testValue)
    };

    const state = {
        fileName: '',
        sheetNames: [],
        currentSheet: '',
        columns: [],
        originalRows: [],
        filteredRows: [],
        searchTerm: '',
        filters: [],
        hiddenColumns: new Set(),
        lastFileBuffer: null,
        isProcessing: false,
        normalizationJobId: 0
    };

    class VirtualTable {
        constructor(tbodyElement, scrollContainer, topScrollContent) {
            this.tbody = tbodyElement;
            this.scrollContainer = scrollContainer;
            this.topScrollContent = topScrollContent;
            this.columns = [];
            this.data = [];
            this.rowHeight = 36;
            this.buffer = 12;
            this.renderedRange = { start: 0, end: 0 };
            this.pendingFrame = null;
            this.pendingForceUpdate = false;
            this.rowMeasured = false;
        }

        setColumns(columns) {
            this.columns = Array.isArray(columns) ? columns : [];
            this.rowMeasured = false;
            this.updateMeasurements(true);
        }

        setData(rows) {
            this.data = Array.isArray(rows) ? rows : [];
            this.renderedRange = { start: -1, end: -1 };
            this.rowMeasured = false;
            this.updateMeasurements(true);
        }

        requestViewportUpdate(force) {
            if (force) {
                this.pendingForceUpdate = true;
            }
            if (this.pendingFrame !== null) {
                return;
            }
            const scheduler = typeof requestAnimationFrame === 'function'
                ? requestAnimationFrame
                : (callback) => setTimeout(callback, 16);
            this.pendingFrame = scheduler(() => {
                this.pendingFrame = null;
                const shouldForce = this.pendingForceUpdate;
                this.pendingForceUpdate = false;
                this.updateViewport(Boolean(shouldForce));
            });
        }

        updateViewport(force) {
            if (!this.scrollContainer) {
                this.renderAll();
                return;
            }

            const total = this.data.length;
            if (total === 0) {
                this.tbody.innerHTML = '';
                if (this.topScrollContent) {
                    this.topScrollContent.style.width = '0px';
                }
                return;
            }

            const scrollTop = this.scrollContainer.scrollTop;
            const viewportHeight = this.scrollContainer.clientHeight || 1;
            const estimatedRowHeight = this.rowHeight || 36;
            const startIndex = Math.max(0, Math.floor(scrollTop / estimatedRowHeight) - this.buffer);
            const visibleCount = Math.ceil(viewportHeight / estimatedRowHeight) + this.buffer * 2;
            const endIndex = Math.min(total, startIndex + visibleCount);

            if (!force && startIndex === this.renderedRange.start && endIndex === this.renderedRange.end) {
                return;
            }

            this.renderRange(startIndex, endIndex);
        }

        renderRange(start, end) {
            const fragment = document.createDocumentFragment();
            const topSpacerHeight = Math.max(0, start) * this.rowHeight;
            fragment.appendChild(this.createSpacer(topSpacerHeight));

            for (let index = start; index < end; index += 1) {
                fragment.appendChild(this.createRow(this.data[index], index));
            }

            const bottomSpacerHeight = Math.max(0, this.data.length - end) * this.rowHeight;
            fragment.appendChild(this.createSpacer(bottomSpacerHeight));

            this.tbody.replaceChildren(fragment);
            this.renderedRange = { start, end };
            if (!this.rowMeasured) {
                this.measureRowHeight();
            }
            this.updateTopScrollWidth();
        }

        createSpacer(height) {
            const tr = document.createElement('tr');
            tr.className = 'virtual-spacer';
            const td = document.createElement('td');
            td.colSpan = Math.max(1, this.columns.length);
            td.style.height = `${height}px`;
            td.style.border = 'none';
            td.style.padding = '0';
            td.style.lineHeight = '0';
            td.style.background = 'transparent';
            tr.appendChild(td);
            return tr;
        }

        createRow(row, index) {
            const tr = document.createElement('tr');
            tr.dataset.rowIndex = String(index);
            this.columns.forEach((column) => {
                const td = document.createElement('td');
                td.textContent = row[column] ?? '';
                tr.appendChild(td);
            });
            return tr;
        }

        measureRowHeight() {
            const firstRow = this.tbody.querySelector('tr[data-row-index]');
            if (!firstRow) {
                return;
            }
            const height = firstRow.getBoundingClientRect().height;
            if (height > 0 && Math.abs(height - this.rowHeight) > 1) {
                this.rowHeight = height;
                this.requestViewportUpdate(true);
            }
            this.rowMeasured = true;
        }

        updateMeasurements(force) {
            if (force) {
                this.requestViewportUpdate(true);
            }
            this.updateTopScrollWidth();
        }

        updateTopScrollWidth() {
            if (!this.topScrollContent) {
                return;
            }
            const table = this.tbody.closest('table');
            if (!table) {
                this.topScrollContent.style.width = '0px';
                return;
            }
            const width = table.scrollWidth;
            this.topScrollContent.style.width = `${width}px`;
        }

        renderAll() {
            const fragment = document.createDocumentFragment();
            this.data.forEach((row, index) => {
                fragment.appendChild(this.createRow(row, index));
            });
            this.tbody.replaceChildren(fragment);
        }
    }

    let worker = createWorker();
    const virtualTable = new VirtualTable(tableBody, tableWrapper, tableScrollTopContent);
    const searchDebouncer = new Debouncer(() => applyFilters(), 160);
    let localWorkbook = null;

    let topSyncInProgress = false;
    let mainSyncInProgress = false;

    initialiseUI();
    registerEventHandlers();

    function createWorker() {
        if (window.location.protocol === 'file:') {
            console.info('Datei wurde über file:// geöffnet – Web Worker sind hier nicht verfügbar.');
            setStatus('Bereit – Hintergrundprozess deaktiviert (lokaler Dateizugriff).');
            enableControls();
            return null;
        }

        if (typeof Worker === 'undefined') {
            console.warn('Web Workers are not supported in this environment. Falling back to main thread parsing.');
            setStatus('Bereit – Hintergrundprozess nicht verfügbar, Datei wird direkt verarbeitet.');
            enableControls();
            return null;
        }

        try {
            const workerUrl = new URL('worker.js', window.location.href);
            const workerInstance = new Worker(workerUrl, { type: 'classic' });

            workerInstance.addEventListener('message', (event) => {
            const { type, payload } = event.data || {};
            switch (type) {
                case 'ready':
                    setStatus('Bereit. Bitte eine Datei auswählen.');
                    break;
                case 'workbook-meta':
                    handleWorkbookMeta(payload);
                    break;
                case 'sheet':
                    handleSheetData(payload);
                    break;
                case 'progress':
                    if (payload?.message) {
                        setStatus(payload.message);
                    }
                    break;
                case 'error':
                    handleWorkerError(payload);
                    break;
                default:
                    break;
            }
            });

            workerInstance.addEventListener('error', (error) => {
                console.error('Worker error', error);
                fallbackToMainThread('Hintergrundprozess nicht verfügbar. Datei wird direkt im Browser verarbeitet.');
            });

            return workerInstance;
        } catch (error) {
            console.warn('Worker initialisation failed, falling back to main thread parsing.', error);
            fallbackToMainThread('Hintergrundprozess konnte nicht gestartet werden. Datei wird direkt im Browser verarbeitet.');
            return null;
        }
    }

    function initialiseUI() {
        disableControls(false);
        setStatus('Initialisiere Oberfläche...');
        resultCount.textContent = 'Keine Daten geladen.';
        tableHead.innerHTML = '';
        virtualTable.setColumns([]);
        virtualTable.setData([]);
        state.hiddenColumns.clear();
        updateHiddenColumnsPanel();
        setUpScrollSync();
        if (typeof requestAnimationFrame === 'function') {
            requestAnimationFrame(() => setStatus('Bereit. Bitte eine Datei auswählen.'));
        } else {
            setStatus('Bereit. Bitte eine Datei auswählen.');
        }
    }

    function registerEventHandlers() {
        if (fileInput) {
            fileInput.addEventListener('change', handleFileSelection);
        }
        if (searchInput) {
            searchInput.addEventListener('input', () => {
                state.searchTerm = searchInput.value.trim().toLowerCase();
                searchDebouncer.run();
            });
        }
        if (addFilterButton) {
            addFilterButton.addEventListener('click', () => addFilterRow());
        }
        if (clearFiltersButton) {
            clearFiltersButton.addEventListener('click', clearFilters);
        }
        if (exportButton) {
            exportButton.addEventListener('click', exportCurrentView);
        }
        if (restoreAllColumnsButton) {
            restoreAllColumnsButton.addEventListener('click', () => {
                restoreAllColumns();
            });
        }
        if (sheetSelect) {
            sheetSelect.addEventListener('change', (event) => {
                const sheetName = event.target.value;
                if (!sheetName || sheetName === state.currentSheet) {
                    return;
                }
                requestSheet(sheetName);
            });
        }
        if (tableWrapper) {
            tableWrapper.addEventListener('scroll', () => {
                if (mainSyncInProgress) {
                    mainSyncInProgress = false;
                    return;
                }
                topSyncInProgress = true;
                if (tableScrollTop) {
                    tableScrollTop.scrollLeft = tableWrapper.scrollLeft;
                }
                virtualTable.requestViewportUpdate();
            });
        }
        if (tableScrollTop) {
            tableScrollTop.addEventListener('scroll', () => {
                if (topSyncInProgress) {
                    topSyncInProgress = false;
                    return;
                }
                mainSyncInProgress = true;
                tableWrapper.scrollLeft = tableScrollTop.scrollLeft;
            });
        }
        window.addEventListener('resize', () => {
            virtualTable.updateMeasurements(true);
        });
    }

    function handleFileSelection(event) {
        const input = event.target;
        const files = input && input.files ? input.files : null;
        const file = files && files.length > 0 ? files[0] : null;
        if (!file) {
            setStatus('Es wurde keine Datei ausgewählt.', true);
            return;
        }
        resetStateForNewFile(file.name);
        disableControls();
        setStatus(`Datei "${file.name}" wird analysiert...`);
        readFileAsArrayBuffer(file)
            .then((buffer) => {
                if (worker) {
                    state.lastFileBuffer = buffer;
                    worker.postMessage({
                        type: 'parse-file',
                        payload: {
                            buffer,
                            name: file.name
                        }
                    });
                } else {
                    processFileInMainThread(buffer, file.name);
                }
            })
            .catch((error) => {
                console.error(error);
                setStatus('Die Datei konnte nicht gelesen werden.', true);
                enableControls();
            });
    }

    function readFileAsArrayBuffer(file) {
        if (!file) {
            return Promise.reject(new Error('Keine Datei angegeben.'));
        }
        if (typeof file.arrayBuffer === 'function') {
            return file.arrayBuffer();
        }
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                resolve(reader.result);
            };
            reader.onerror = () => {
                reject(reader.error || new Error('Die Datei konnte nicht gelesen werden.'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function handleWorkbookMeta(meta) {
        const sheetNames = Array.isArray(meta?.sheetNames) ? meta.sheetNames : [];
        state.sheetNames = sheetNames;
        populateSheetSelect(sheetNames);
    }

    function handleSheetData(data) {
        if (!data) {
            setStatus('Es wurden keine Daten im Arbeitsblatt gefunden.', true);
            return;
        }
        const { columns, rows, sheetName } = data;
        state.columns = Array.isArray(columns) ? columns : [];
        state.originalRows = [];
        state.filteredRows = [];
        state.currentSheet = sheetName || '';
        state.hiddenColumns.clear();
        updateHiddenColumnsPanel();
        state.lastFileBuffer = null;
        state.isProcessing = true;
        state.normalizationJobId += 1;
        const jobId = state.normalizationJobId;
        updateHeader();
        updateFilterColumnOptions();
        resultCount.textContent = 'Daten werden verarbeitet...';
        setStatus(`Arbeitsblatt "${state.currentSheet}" wird verarbeitet...`);
        const totalRows = Array.isArray(rows) ? rows.length : 0;
        processRowsAsync(Array.isArray(rows) ? rows : [], jobId, totalRows, () => {
            if (state.normalizationJobId !== jobId) {
                return;
            }
            state.isProcessing = false;
            enableControls();
            applyFilters(true);
            if (sheetSelect && state.currentSheet) {
                sheetSelect.value = state.currentSheet;
            }
            setStatus(`Arbeitsblatt "${state.currentSheet}" geladen. (${state.originalRows.length} Zeilen)`);
        });
    }

    function handleWorkerError(payload) {
        const message = payload?.message || 'Unbekannter Fehler im Hintergrundprozess.';
        if (state.lastFileBuffer && state.fileName) {
            fallbackToMainThread(message, true);
        } else {
            setStatus(message, true);
            state.lastFileBuffer = null;
            enableControls();
        }
    }

    function requestSheet(sheetName) {
        setStatus(`Arbeitsblatt "${sheetName}" wird geladen...`);
        disableControls(false);
        if (worker) {
            worker.postMessage({ type: 'load-sheet', payload: { sheetName } });
        } else {
            loadSheetLocally(sheetName);
        }
    }

    function fallbackToMainThread(message, isError = false) {
        if (worker) {
            try {
                worker.terminate();
            } catch (terminationError) {
                console.warn('Worker termination failed, continuing with main thread parsing.', terminationError);
            }
            worker = null;
        }

        if (message) {
            setStatus(message, isError);
        } else {
            setStatus('Hintergrundverarbeitung nicht verfügbar. Datei wird direkt im Browser verarbeitet.');
        }

        const buffer = state.lastFileBuffer;
        const fileName = state.fileName;
        state.lastFileBuffer = null;

        if (buffer && fileName) {
            processFileInMainThread(buffer, fileName);
        } else {
            enableControls();
        }
    }

    function processFileInMainThread(buffer, fileName) {
        try {
            setStatus(`Datei "${fileName}" wird direkt im Browser verarbeitet...`);
            state.lastFileBuffer = null;
            const options = {
                type: 'array',
                cellDates: true,
                cellNF: false,
                cellText: false,
                dateNF: 'yyyy-MM-dd'
            };
            const normalizedBuffer = normalizeArrayBuffer(buffer);
            localWorkbook = XLSX.read(normalizedBuffer, options);
            const sheetNames = Array.isArray(localWorkbook.SheetNames) ? [...localWorkbook.SheetNames] : [];
            state.sheetNames = sheetNames;
            populateSheetSelect(sheetNames);

            if (sheetNames.length === 0) {
                state.currentSheet = '';
                state.columns = [];
                state.originalRows = [];
                state.filteredRows = [];
                resultCount.textContent = 'Keine Daten geladen.';
                setStatus('Die ausgewählte Arbeitsmappe enthält keine Tabellenblätter.', true);
                enableControls();
                return;
            }

            const preferred = sheetNames.includes(state.currentSheet) && state.currentSheet
                ? state.currentSheet
                : sheetNames[0];
            loadSheetLocally(preferred);
        } catch (error) {
            console.error('Fehler bei der Verarbeitung im Hauptthread', error);
            setStatus('Fehler beim Einlesen der Datei: ' + (error?.message || String(error)), true);
            enableControls();
        }
    }

    function normalizeArrayBuffer(input) {
        if (input instanceof Uint8Array) {
            return input;
        }
        if (input instanceof ArrayBuffer) {
            return new Uint8Array(input);
        }
        if (input && typeof input.byteLength === 'number') {
            return new Uint8Array(input);
        }
        throw new Error('Ungültiger Dateipuffer.');
    }

    function loadSheetLocally(sheetName) {
        if (!localWorkbook) {
            setStatus('Keine Arbeitsmappe im Speicher. Bitte Datei erneut laden.', true);
            enableControls();
            return;
        }
        if (!sheetName) {
            setStatus('Es wurde kein Arbeitsblatt ausgewählt.', true);
            enableControls();
            return;
        }

        const worksheet = localWorkbook.Sheets?.[sheetName];
        if (!worksheet) {
            setStatus(`Arbeitsblatt "${sheetName}" wurde nicht gefunden.`, true);
            enableControls();
            return;
        }

        try {
            const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: true });
            const columns = inferColumns(rows);
            handleSheetData({ sheetName, columns, rows });
        } catch (error) {
            console.error('Fehler beim Laden des Arbeitsblatts im Hauptthread', error);
            setStatus('Fehler beim Laden des Arbeitsblatts: ' + (error?.message || String(error)), true);
            enableControls();
        }
    }

    function inferColumns(rows) {
        const columnSet = new Set();
        rows.forEach((row) => {
            Object.keys(row || {}).forEach((column) => columnSet.add(column));
        });
        return Array.from(columnSet);
    }

    function resetStateForNewFile(fileName) {
        state.fileName = fileName;
        state.sheetNames = [];
        state.currentSheet = '';
        state.columns = [];
        state.originalRows = [];
        state.filteredRows = [];
        state.searchTerm = '';
        state.filters = [];
        state.hiddenColumns.clear();
        state.lastFileBuffer = null;
        state.isProcessing = false;
        filtersContainer.innerHTML = '';
        clearFiltersButton.disabled = true;
        if (searchInput) {
            searchInput.value = '';
        }
        if (sheetSelect) {
            sheetSelect.innerHTML = '<option value="">Arbeitsblätter werden geladen...</option>';
            sheetSelect.disabled = true;
        }
        tableHead.innerHTML = '';
        virtualTable.setColumns([]);
        virtualTable.setData([]);
        updateHiddenColumnsPanel();
        resultCount.textContent = 'Daten werden geladen...';
        localWorkbook = null;
    }

    function disableControls(includeFile = true) {
        if (fileInput) {
            fileInput.disabled = Boolean(includeFile);
        }
        if (searchInput) {
            searchInput.disabled = true;
        }
        if (addFilterButton) {
            addFilterButton.disabled = true;
        }
        if (clearFiltersButton) {
            clearFiltersButton.disabled = true;
        }
        if (exportButton) {
            exportButton.disabled = true;
        }
        if (sheetSelect) {
            sheetSelect.disabled = true;
        }
    }

    function enableControls() {
        if (fileInput) {
            fileInput.disabled = false;
        }
        if (searchInput) {
            searchInput.disabled = false;
        }
        if (addFilterButton) {
            addFilterButton.disabled = false;
        }
        if (clearFiltersButton) {
            clearFiltersButton.disabled = state.filters.length === 0;
        }
        if (exportButton) {
            exportButton.disabled = state.filteredRows.length === 0;
        }
        if (sheetSelect && state.sheetNames.length > 0) {
            sheetSelect.disabled = false;
        }
    }

    function populateSheetSelect(sheetNames) {
        if (!sheetSelect) {
            return;
        }
        sheetSelect.innerHTML = '';
        if (!Array.isArray(sheetNames) || sheetNames.length === 0) {
            const option = document.createElement('option');
            option.value = '';
            option.textContent = 'Keine Arbeitsblätter gefunden';
            sheetSelect.appendChild(option);
            sheetSelect.disabled = true;
            return;
        }
        sheetNames.forEach((name) => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });
        sheetSelect.disabled = false;
    }

    function updateHeader() {
        const visibleColumns = getVisibleColumns();
        tableHead.innerHTML = '';
        virtualTable.setColumns(visibleColumns);
        if (!Array.isArray(visibleColumns) || visibleColumns.length === 0) {
            virtualTable.updateMeasurements(true);
            return;
        }
        const headerRow = document.createElement('tr');
        const singleColumn = visibleColumns.length <= 1;
        visibleColumns.forEach((column) => {
            const th = document.createElement('th');
            th.textContent = column;
            th.dataset.column = column;
            if (singleColumn) {
                th.classList.add('locked-column');
                th.title = 'Mindestens eine Spalte muss sichtbar bleiben.';
            } else {
                th.title = `Spalte "${column}" ausblenden (erneut einblenden über "Ausgeblendete Spalten")`;
                th.addEventListener('click', () => {
                    toggleColumnVisibility(column);
                });
            }
            headerRow.appendChild(th);
        });
        tableHead.appendChild(headerRow);
        virtualTable.updateMeasurements(true);
    }

    function addFilterRow(initialValues) {
        if (!filterTemplate) {
            return;
        }
        const clone = filterTemplate.content.firstElementChild.cloneNode(true);
        const columnSelect = clone.querySelector('.filter-column');
        const operatorSelect = clone.querySelector('.filter-operator');
        const valueInput = clone.querySelector('.filter-value');
        const removeButton = clone.querySelector('.remove-filter');

        const filter = {
            id: createId(),
            column: initialValues?.column || '',
            operator: initialValues?.operator || 'contains',
            value: initialValues?.value || '',
            element: clone
        };

        populateColumnOptions(columnSelect, filter.column);
        operatorSelect.value = filter.operator;
        valueInput.value = filter.value;

        columnSelect.addEventListener('change', () => {
            filter.column = columnSelect.value;
            applyFilters();
        });
        operatorSelect.addEventListener('change', () => {
            filter.operator = operatorSelect.value;
            applyFilters();
        });
        valueInput.addEventListener('input', () => {
            filter.value = valueInput.value;
            applyFilters();
        });
        removeButton.addEventListener('click', () => {
            filtersContainer.removeChild(clone);
            state.filters = state.filters.filter((item) => item.id !== filter.id);
            clearFiltersButton.disabled = state.filters.length === 0;
            applyFilters();
        });

        filtersContainer.appendChild(clone);
        state.filters.push(filter);
        clearFiltersButton.disabled = false;
        applyFilters();
    }

    function clearFilters() {
        filtersContainer.innerHTML = '';
        state.filters = [];
        clearFiltersButton.disabled = true;
        applyFilters();
    }

    function populateColumnOptions(selectElement, currentValue) {
        if (!selectElement) {
            return;
        }
        selectElement.innerHTML = '';
        const placeholder = document.createElement('option');
        placeholder.value = '';
        placeholder.textContent = 'Spalte wählen';
        selectElement.appendChild(placeholder);
        state.columns.forEach((column) => {
            const option = document.createElement('option');
            option.value = column;
            option.textContent = column;
            selectElement.appendChild(option);
        });
        if (currentValue && state.columns.includes(currentValue)) {
            selectElement.value = currentValue;
        } else {
            selectElement.value = '';
        }
    }

    function updateFilterColumnOptions() {
        const selects = filtersContainer.querySelectorAll('.filter-column');
        selects.forEach((select) => {
            const previousValue = select.value;
            populateColumnOptions(select, previousValue);
        });
    }

    function getVisibleColumns() {
        if (!Array.isArray(state.columns)) {
            return [];
        }
        return state.columns.filter((column) => !state.hiddenColumns.has(column));
    }

    function toggleColumnVisibility(column) {
        if (!column) {
            return;
        }
        if (state.hiddenColumns.has(column)) {
            restoreColumn(column);
        } else {
            hideColumn(column);
        }
    }

    function hideColumn(column) {
        if (!column || state.hiddenColumns.has(column)) {
            return;
        }
        const visibleColumns = getVisibleColumns();
        if (visibleColumns.length <= 1) {
            setStatus('Mindestens eine Spalte muss sichtbar bleiben.', true);
            return;
        }
        state.hiddenColumns.add(column);
        updateHiddenColumnsPanel();
        updateHeader();
    setStatus(`Spalte "${column}" ausgeblendet. Über "Ausgeblendete Spalten" wieder einblenden.`);
    }

    function restoreColumn(column) {
        if (!column || !state.hiddenColumns.has(column)) {
            return;
        }
        state.hiddenColumns.delete(column);
        updateHiddenColumnsPanel();
        updateHeader();
    setStatus(`Spalte "${column}" wieder eingeblendet.`);
    }

    function restoreAllColumns() {
        if (state.hiddenColumns.size === 0) {
            return;
        }
        state.hiddenColumns.clear();
        updateHiddenColumnsPanel();
        updateHeader();
        setStatus('Alle Spalten wurden eingeblendet.');
    }

    function updateHiddenColumnsPanel() {
        if (!hiddenColumnsPanel || !hiddenColumnsList) {
            return;
        }
        hiddenColumnsList.replaceChildren();
        if (state.hiddenColumns.size === 0) {
            hiddenColumnsPanel.hidden = true;
            if (restoreAllColumnsButton) {
                restoreAllColumnsButton.disabled = true;
            }
            return;
        }

        hiddenColumnsPanel.hidden = false;
        if (restoreAllColumnsButton) {
            restoreAllColumnsButton.disabled = false;
        }
        const sorted = Array.from(state.hiddenColumns).sort((a, b) => a.localeCompare(b, 'de-DE'));
        sorted.forEach((column) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.textContent = column;
            button.addEventListener('click', () => {
                restoreColumn(column);
            });
            hiddenColumnsList.appendChild(button);
        });
    }

    function applyFilters(resetScroll = false) {
        if (state.isProcessing) {
            if (resultCount) {
                resultCount.textContent = 'Daten werden verarbeitet...';
            }
            return;
        }
        if (!Array.isArray(state.originalRows) || state.originalRows.length === 0) {
            state.filteredRows = [];
            updateResultView(resetScroll);
            return;
        }
        const activeFilters = state.filters.filter((filter) => filter.column && filter.value.trim() !== '');
        const searchTerm = state.searchTerm;
        const results = state.originalRows.filter((row) => {
            if (searchTerm && !row.__search.includes(searchTerm)) {
                return false;
            }
            return activeFilters.every((filter) => evaluateFilter(row, filter));
        });
        state.filteredRows = results;
        updateResultView(resetScroll);
    }

    function evaluateFilter(row, filter) {
        const operator = OPERATIONS[filter.operator] || OPERATIONS.contains;
        const cellValue = (row[filter.column] ?? '').toString().toLowerCase();
        const testValue = filter.value.trim().toLowerCase();
        if (testValue === '') {
            return true;
        }
        return operator(cellValue, testValue);
    }

    function updateResultView(resetScroll) {
        const total = state.originalRows.length;
        const visible = state.filteredRows.length;
        resultCount.textContent = total === 0
            ? 'Keine Daten geladen.'
            : `${visible} von ${total} Zeilen angezeigt`;
        if (exportButton) {
            exportButton.disabled = visible === 0;
        }
        if (resetScroll) {
            tableWrapper.scrollTop = 0;
        }
        virtualTable.setData(state.filteredRows);
    }

    function processRowsAsync(rows, jobId, totalRows, onComplete) {
        if (!Array.isArray(rows) || rows.length === 0) {
            state.originalRows = [];
            state.filteredRows = [];
            if (resultCount) {
                resultCount.textContent = 'Keine Daten im Arbeitsblatt.';
            }
            if (typeof onComplete === 'function') {
                onComplete();
            }
            return;
        }

        state.originalRows = [];
        const batchSize = 400;
        const timeSource = (typeof performance !== 'undefined' && performance && typeof performance.now === 'function')
            ? performance
            : { now: () => Date.now() };
        let sourceRows = rows;
        let index = 0;
        const total = typeof totalRows === 'number' ? totalRows : sourceRows.length;

        function updateProgress() {
            if (!resultCount || total <= 0) {
                return;
            }
            const percentage = Math.min(100, Math.round((index / total) * 100));
            resultCount.textContent = `Verarbeite Daten... (${percentage}%)`;
        }

        function scheduleNextChunk() {
            if (state.normalizationJobId !== jobId) {
                sourceRows = null;
                return;
            }
            if (typeof window.requestIdleCallback === 'function') {
                window.requestIdleCallback(processChunk, { timeout: 100 });
            } else {
                setTimeout(processChunk, 0);
            }
        }

        function processChunk(deadline) {
            if (state.normalizationJobId !== jobId) {
                sourceRows = null;
                return;
            }
            const startTime = timeSource.now();
            let processed = 0;
            while (index < total) {
                state.originalRows.push(normalizeRow(sourceRows[index]));
                index += 1;
                processed += 1;

                if (processed >= batchSize) {
                    break;
                }
                if (deadline && typeof deadline.timeRemaining === 'function' && deadline.timeRemaining() <= 4) {
                    break;
                }
                if (!deadline && timeSource.now() - startTime > 12) {
                    break;
                }
            }

            if (state.normalizationJobId !== jobId) {
                sourceRows = null;
                return;
            }

            if (index < total) {
                if (processed > 0) {
                    updateProgress();
                }
                scheduleNextChunk();
                return;
            }

            sourceRows = null;
            if (typeof onComplete === 'function') {
                onComplete();
            }
        }

        updateProgress();
        scheduleNextChunk();
    }

    function normalizeRow(row) {
        const normalized = {};
        const searchParts = [];
        state.columns.forEach((column) => {
            const rawValue = column in row ? row[column] : '';
            const { display, searchValue } = normalizeCellValue(column, rawValue);
            normalized[column] = display;
            searchParts.push(searchValue);
        });
        normalized.__search = searchParts.join(' ');
        return normalized;
    }

    function normalizeCellValue(column, value) {
        if (value === null || value === undefined) {
            return { display: '', searchValue: '' };
        }
        let displayValue = value;
        if (DATE_COLUMNS.has(column.toLowerCase())) {
            const formatted = formatDateValue(value);
            if (formatted) {
                displayValue = formatted;
            }
        }
        if (typeof displayValue === 'number' && !Number.isNaN(displayValue)) {
            displayValue = Number.isInteger(displayValue) ? displayValue.toString(10) : displayValue.toLocaleString('de-DE');
        }
        if (typeof displayValue === 'object') {
            displayValue = JSON.stringify(displayValue);
        }
        const text = displayValue.toString();
        return { display: text, searchValue: text.toLowerCase() };
    }

    function formatDateValue(value) {
        try {
            if (value instanceof Date && !Number.isNaN(value.getTime())) {
                return new Intl.DateTimeFormat('de-DE', { timeZone: 'UTC' }).format(value);
            }
            if (typeof value === 'number') {
                const parsed = XLSX.SSF.parse_date_code(value);
                if (parsed) {
                    const date = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
                    return new Intl.DateTimeFormat('de-DE', { timeZone: 'UTC' }).format(date);
                }
            }
            if (typeof value === 'string') {
                const sanitized = value.trim();
                if (!sanitized) {
                    return '';
                }
                const asNumber = Number(sanitized);
                if (!Number.isNaN(asNumber)) {
                    const parsedNumber = XLSX.SSF.parse_date_code(asNumber);
                    if (parsedNumber) {
                        const dateFromNumber = new Date(Date.UTC(parsedNumber.y, parsedNumber.m - 1, parsedNumber.d));
                        return new Intl.DateTimeFormat('de-DE', { timeZone: 'UTC' }).format(dateFromNumber);
                    }
                }
                const parsed = new Date(sanitized);
                if (!Number.isNaN(parsed.getTime())) {
                    return new Intl.DateTimeFormat('de-DE').format(parsed);
                }
            }
        } catch (error) {
            console.warn('Datum konnte nicht formatiert werden:', error);
        }
        return null;
    }

    function exportCurrentView() {
        if (!Array.isArray(state.filteredRows) || state.filteredRows.length === 0) {
            setStatus('Es sind keine Daten zum Export vorhanden.', true);
            return;
        }
        const exportRows = state.filteredRows.map((row) => {
            const copy = {};
            state.columns.forEach((column) => {
                copy[column] = row[column];
            });
            return copy;
        });
        const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: state.columns });
        const workbook = XLSX.utils.book_new();
        const sheetName = state.currentSheet || 'Export';
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.substring(0, 31));
        const timeStamp = new Date().toISOString().replace(/[:T]/g, '-').split('.')[0];
        const baseName = state.fileName ? state.fileName.replace(/\.[^.]+$/, '') : 'mvms-export';
        const fileName = `${baseName}-${sheetName || 'sheet'}-${timeStamp}.xlsx`;
        XLSX.writeFile(workbook, fileName);
        setStatus(`Export erfolgreich erstellt: ${fileName}`);
    }

    function setStatus(message, isError = false) {
        if (!statusBox) {
            return;
        }
        statusBox.textContent = message;
        statusBox.classList.toggle('error', Boolean(isError));
    }

    function setUpScrollSync() {
        if (!tableScrollTopContent) {
            return;
        }
        tableScrollTopContent.style.width = '0px';
    }

    function createId() {
        if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
            return crypto.randomUUID();
        }
        return `mvms-${Math.random().toString(36).slice(2, 10)}`;
    }

    function Debouncer(callback, delay) {
        this.callback = callback;
        this.delay = delay;
        this.timer = null;
    }

    Debouncer.prototype.run = function () {
        clearTimeout(this.timer);
        this.timer = setTimeout(() => {
            this.callback();
        }, this.delay);
    };

})();
