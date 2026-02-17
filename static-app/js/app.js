/**
 * Main Application Controller
 * Handles all UI interactions and coordinates between modules
 */

const app = {
    data: {
        lmrbData: [],
        tcData: [],
        mappings: [],
        matchedData: [],
        unmatchedLMRB: [],
        unmatchedTC: [],
        selectedChannel: '',
        selectedLMRBIndex: null,
        selectedTCIndex: null
    },

    /**
     * Initialize application
     */
    init() {
        console.log('Initializing Media Reconciliation System...');

        // Load saved data from localStorage
        this.loadFromStorage();

        // Setup tab navigation
        this.setupTabs();

        // Update UI with loaded data
        this.updateAllCounts();

        // Set default dates (last 30 days)
        const endDate = new Date();
        const startDate = new Date();
        startDate.setDate(startDate.getDate() - 30);

        document.getElementById('startDate').valueAsDate = startDate;
        document.getElementById('endDate').valueAsDate = endDate;

        console.log('Application initialized successfully');
    },

    /**
     * Setup tab navigation
     */
    setupTabs() {
        const tabButtons = document.querySelectorAll('.tab-btn');
        tabButtons.forEach(btn => {
            btn.addEventListener('click', () => {
                const tabName = btn.getAttribute('data-tab');
                this.switchTab(tabName);
            });
        });
    },

    /**
     * Switch between tabs
     */
    switchTab(tabName) {
        // Update tab buttons
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
            if (btn.getAttribute('data-tab') === tabName) {
                btn.classList.add('active');
            }
        });

        // Update tab content
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`tab-${tabName}`).classList.add('active');

        // Refresh data for specific tabs
        if (tabName === 'results') {
            this.displayResults();
        } else if (tabName === 'manual') {
            this.displayManualMatching();
        }
    },

    /**
     * Load LMRB data from Google Sheets
     */
    async loadLMRBData() {
        const apiKey = document.getElementById('apiKey').value.trim();
        const sheetId = document.getElementById('sheetId').value.trim();
        const startDate = document.getElementById('startDate').value;
        const endDate = document.getElementById('endDate').value;

        if (!apiKey || !sheetId) {
            this.showStatus('lmrbStatus', 'Please enter API Key and Sheet ID', 'error');
            return;
        }

        this.showStatus('lmrbStatus', 'Loading data from Google Sheets...', 'info');

        try {
            const data = await googleSheets.loadData(apiKey, sheetId, startDate, endDate);
            this.data.lmrbData = data;

            // Save to localStorage
            storage.save('lmrbData', data);
            storage.save('apiKey', apiKey);
            storage.save('sheetId', sheetId);

            this.showStatus('lmrbStatus', `✅ Loaded ${data.length} LMRB records successfully!`, 'success');
            this.updateAllCounts();
        } catch (error) {
            console.error('Error loading LMRB data:', error);
            this.showStatus('lmrbStatus', `❌ Error: ${error.message}`, 'error');
        }
    },

    /**
     * Update channel info when channel is selected
     */
    updateChannelInfo() {
        const channel = document.getElementById('channelSelect').value;
        this.data.selectedChannel = channel;

        if (channel) {
            const info = this.getChannelFormatInfo(channel);
            document.getElementById('channelFormatDesc').textContent = info;
            document.getElementById('channelInfo').style.display = 'block';
        } else {
            document.getElementById('channelInfo').style.display = 'none';
        }
    },

    /**
     * Get channel format information
     */
    getChannelFormatInfo(channel) {
        const formats = {
            sirasa: 'Sirasa TV: Standard format with Date | Program | Time | Theme | Duration',
            hiru: 'Hiru TV: Format details to be configured',
            derana: 'Derana TV: Format details to be configured',
            itv: 'ITN: Format details to be configured',
            tv1: 'TV1: Format details to be configured'
        };
        return formats[channel] || 'Channel format not configured';
    },

    /**
     * Convert TC PDF to data
     */
    async convertTCPdf() {
        const fileInput = document.getElementById('tcPdfFile');
        const channel = document.getElementById('channelSelect').value;

        if (!fileInput.files[0]) {
            this.showStatus('uploadStatus', 'Please select a PDF file', 'error');
            return;
        }

        if (!channel) {
            this.showStatus('uploadStatus', 'Please select a channel first', 'error');
            return;
        }

        this.showStatus('uploadStatus', 'Converting PDF...', 'info');

        try {
            const file = fileInput.files[0];
            const converter = this.getConverter(channel);

            if (!converter) {
                throw new Error(`No converter available for ${channel}`);
            }

            const data = await converter.convert(file);
            this.data.tcData = data;

            // Save to localStorage
            storage.save('tcData', data);

            this.showStatus('uploadStatus', `✅ Converted ${data.length} TC records successfully!`, 'success');
            this.updateAllCounts();
            this.displayTCPreview(data);
        } catch (error) {
            console.error('Error converting PDF:', error);
            this.showStatus('uploadStatus', `❌ Error: ${error.message}`, 'error');
        }
    },

    /**
     * Get converter for channel
     */
    getConverter(channel) {
        const converters = {
            sirasa: window.sirasaConverter,
            hiru: window.hiruConverter,
            derana: window.deranaConverter,
            itv: window.itvConverter,
            tv1: window.tv1Converter
        };
        return converters[channel];
    },

    /**
     * Display TC preview
     */
    displayTCPreview(data) {
        const container = document.getElementById('tcPreviewTable');
        const preview = data.slice(0, 10); // Show first 10 records

        let html = '<table><thead><tr>';
        html += '<th>Date</th><th>Program</th><th>Time</th><th>Theme</th><th>Duration</th>';
        html += '</tr></thead><tbody>';

        preview.forEach(row => {
            html += '<tr>';
            html += `<td>${row.Date || 'N/A'}</td>`;
            html += `<td>${row.Program || 'N/A'}</td>`;
            html += `<td>${row.Time || 'N/A'}</td>`;
            html += `<td>${row.Theme || 'N/A'}</td>`;
            html += `<td>${row.Duration || 'N/A'}s</td>`;
            html += '</tr>';
        });

        html += '</tbody></table>';
        container.innerHTML = html;
        document.getElementById('tcPreview').style.display = 'block';
    },

    /**
     * Add theme mapping
     */
    addMapping() {
        const mapping = {
            type: document.getElementById('mappingType').value,
            tcTheme: document.getElementById('tcTheme').value.trim(),
            tcDuration: parseInt(document.getElementById('tcDuration').value),
            lmrbTheme: document.getElementById('lmrbTheme').value.trim(),
            lmrbDuration: parseInt(document.getElementById('lmrbDuration').value)
        };

        if (!mapping.tcTheme || !mapping.lmrbTheme) {
            alert('Please fill in both TC Theme and LMRB Theme');
            return;
        }

        this.data.mappings.push(mapping);
        storage.save('mappings', this.data.mappings);

        // Clear form
        document.getElementById('tcTheme').value = '';
        document.getElementById('lmrbTheme').value = '';

        this.displayMappings();
        this.updateAllCounts();
    },

    /**
     * Display mappings list
     */
    displayMappings() {
        const container = document.getElementById('mappingsList');
        const mappings = this.data.mappings;

        if (mappings.length === 0) {
            container.innerHTML = '<p class="text-center">No mappings added yet</p>';
            return;
        }

        let html = '';
        mappings.forEach((mapping, index) => {
            html += `<div class="mapping-item">`;
            html += `<div class="mapping-info">`;
            html += `<span class="mapping-type">${mapping.type}</span>`;
            html += `<strong>${mapping.tcTheme}</strong> (${mapping.tcDuration}s) `;
            html += `→ <strong>${mapping.lmrbTheme}</strong> (${mapping.lmrbDuration}s)`;
            html += `</div>`;
            html += `<button class="mapping-delete" onclick="app.deleteMapping(${index})">Delete</button>`;
            html += `</div>`;
        });

        container.innerHTML = html;
        document.getElementById('mappingsCount').textContent = mappings.length;
    },

    /**
     * Delete mapping
     */
    deleteMapping(index) {
        this.data.mappings.splice(index, 1);
        storage.save('mappings', this.data.mappings);
        this.displayMappings();
        this.updateAllCounts();
    },

    /**
     * Run reconciliation
     */
    runReconciliation() {
        if (this.data.lmrbData.length === 0) {
            alert('Please load LMRB data first');
            return;
        }

        if (this.data.tcData.length === 0) {
            alert('Please upload and convert TC PDF first');
            return;
        }

        if (this.data.mappings.length === 0) {
            alert('Please add at least one theme mapping');
            return;
        }

        this.showStatus('reconcileStatus', 'Running reconciliation...', 'info');

        const timeTolerance = parseInt(document.getElementById('timeTolerance').value);
        const ignoreDate = document.getElementById('ignoreDate').checked;

        try {
            const results = reconciliation.match(
                this.data.lmrbData,
                this.data.tcData,
                this.data.mappings,
                timeTolerance,
                ignoreDate
            );

            this.data.matchedData = results.matched;
            this.data.unmatchedLMRB = results.unmatchedLMRB;
            this.data.unmatchedTC = results.unmatchedTC;

            // Save results
            storage.save('matchedData', this.data.matchedData);
            storage.save('unmatchedLMRB', this.data.unmatchedLMRB);
            storage.save('unmatchedTC', this.data.unmatchedTC);

            this.showStatus('reconcileStatus', '✅ Reconciliation completed!', 'success');
            this.displayReconciliationResults(results);
        } catch (error) {
            console.error('Error running reconciliation:', error);
            this.showStatus('reconcileStatus', `❌ Error: ${error.message}`, 'error');
        }
    },

    /**
     * Display reconciliation results summary
     */
    displayReconciliationResults(results) {
        document.getElementById('totalLMRB').textContent = this.data.lmrbData.length;
        document.getElementById('totalTC').textContent = this.data.tcData.length;
        document.getElementById('totalMatched').textContent = results.matched.length;
        document.getElementById('totalUnmatched').textContent = results.unmatchedLMRB.length;

        const matchRate = this.data.lmrbData.length > 0
            ? ((results.matched.length / this.data.lmrbData.length) * 100).toFixed(1)
            : 0;
        document.getElementById('matchRate').textContent = matchRate + '%';

        document.getElementById('reconcileResults').style.display = 'block';
    },

    /**
     * Display results in results tab
     */
    displayResults() {
        document.getElementById('matchedCountTab').textContent = this.data.matchedData.length;
        document.getElementById('unmatchedLMRBCountTab').textContent = this.data.unmatchedLMRB.length;
        document.getElementById('unmatchedTCCountTab').textContent = this.data.unmatchedTC.length;

        // Show matched by default
        this.showResultsTab('matched');
    },

    /**
     * Show specific results tab
     */
    showResultsTab(type) {
        // Update tab buttons
        document.querySelectorAll('.results-tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        event.target.classList.add('active');

        const container = document.getElementById('resultsTableContainer');
        let data = [];

        if (type === 'matched') {
            data = this.data.matchedData;
        } else if (type === 'unmatchedLMRB') {
            data = this.data.unmatchedLMRB;
        } else if (type === 'unmatchedTC') {
            data = this.data.unmatchedTC;
        }

        this.displayTable(container, data);
    },

    /**
     * Display data table
     */
    displayTable(container, data) {
        if (data.length === 0) {
            container.innerHTML = '<p class="text-center">No data to display</p>';
            return;
        }

        const headers = Object.keys(data[0]);
        let html = '<table><thead><tr>';

        headers.forEach(header => {
            html += `<th>${header}</th>`;
        });
        html += '</tr></thead><tbody>';

        data.forEach(row => {
            html += '<tr>';
            headers.forEach(header => {
                html += `<td>${row[header] || 'N/A'}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table>';
        container.innerHTML = html;
    },

    /**
     * Display manual matching interface
     */
    displayManualMatching() {
        const lmrbList = document.getElementById('manualLMRBList');
        const tcList = document.getElementById('manualTCList');

        // Display unmatched LMRB
        let lmrbHtml = '';
        this.data.unmatchedLMRB.forEach((item, index) => {
            lmrbHtml += `<div class="match-item" onclick="app.selectLMRB(${index})">`;
            lmrbHtml += `${item.Date} - ${item.Theme} - ${item.Time} (${item.Duration}s)`;
            lmrbHtml += `</div>`;
        });
        lmrbList.innerHTML = lmrbHtml || '<p class="text-center">No unmatched records</p>';

        // Display unmatched TC
        let tcHtml = '';
        this.data.unmatchedTC.forEach((item, index) => {
            tcHtml += `<div class="match-item" onclick="app.selectTC(${index})">`;
            tcHtml += `${item.Date} - ${item.Theme} - ${item.Time} (${item.Duration}s)`;
            tcHtml += `</div>`;
        });
        tcList.innerHTML = tcHtml || '<p class="text-center">No unmatched records</p>';

        document.getElementById('manualLMRBCount').textContent = this.data.unmatchedLMRB.length;
        document.getElementById('manualTCCount').textContent = this.data.unmatchedTC.length;
    },

    /**
     * Select LMRB item for manual matching
     */
    selectLMRB(index) {
        this.data.selectedLMRBIndex = index;

        document.querySelectorAll('#manualLMRBList .match-item').forEach((item, i) => {
            item.classList.toggle('selected', i === index);
        });

        this.updateManualMatchButton();
    },

    /**
     * Select TC item for manual matching
     */
    selectTC(index) {
        this.data.selectedTCIndex = index;

        document.querySelectorAll('#manualTCList .match-item').forEach((item, i) => {
            item.classList.toggle('selected', i === index);
        });

        this.updateManualMatchButton();
    },

    /**
     * Update manual match button state
     */
    updateManualMatchButton() {
        const btn = document.getElementById('manualMatchBtn');
        btn.disabled = this.data.selectedLMRBIndex === null || this.data.selectedTCIndex === null;
    },

    /**
     * Perform manual match
     */
    manualMatch() {
        if (this.data.selectedLMRBIndex === null || this.data.selectedTCIndex === null) {
            return;
        }

        const lmrbItem = this.data.unmatchedLMRB[this.data.selectedLMRBIndex];
        const tcItem = this.data.unmatchedTC[this.data.selectedTCIndex];

        // Merge records
        const matched = { ...lmrbItem, ...tcItem, matchType: 'MANUAL' };
        this.data.matchedData.push(matched);

        // Remove from unmatched
        this.data.unmatchedLMRB.splice(this.data.selectedLMRBIndex, 1);
        this.data.unmatchedTC.splice(this.data.selectedTCIndex, 1);

        // Save
        storage.save('matchedData', this.data.matchedData);
        storage.save('unmatchedLMRB', this.data.unmatchedLMRB);
        storage.save('unmatchedTC', this.data.unmatchedTC);

        // Reset selection
        this.data.selectedLMRBIndex = null;
        this.data.selectedTCIndex = null;

        // Refresh display
        this.displayManualMatching();
        alert('Records matched successfully!');
    },

    /**
     * Export to Excel
     */
    exportToExcel() {
        if (this.data.matchedData.length === 0) {
            alert('No data to export');
            return;
        }

        const wb = XLSX.utils.book_new();

        // Matched data sheet
        const ws1 = XLSX.utils.json_to_sheet(this.data.matchedData);
        XLSX.utils.book_append_sheet(wb, ws1, 'Matched');

        // Unmatched LMRB sheet
        if (this.data.unmatchedLMRB.length > 0) {
            const ws2 = XLSX.utils.json_to_sheet(this.data.unmatchedLMRB);
            XLSX.utils.book_append_sheet(wb, ws2, 'Unmatched LMRB');
        }

        // Unmatched TC sheet
        if (this.data.unmatchedTC.length > 0) {
            const ws3 = XLSX.utils.json_to_sheet(this.data.unmatchedTC);
            XLSX.utils.book_append_sheet(wb, ws3, 'Unmatched TC');
        }

        XLSX.writeFile(wb, `reconciliation_${new Date().toISOString().split('T')[0]}.xlsx`);
    },

    /**
     * Export to PDF
     */
    exportToPDF() {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        doc.setFontSize(18);
        doc.text('Media Reconciliation Report', 14, 20);

        doc.setFontSize(12);
        doc.text(`Date: ${new Date().toLocaleDateString()}`, 14, 30);

        doc.text(`Total LMRB Records: ${this.data.lmrbData.length}`, 14, 40);
        doc.text(`Total TC Records: ${this.data.tcData.length}`, 14, 47);
        doc.text(`Matched Records: ${this.data.matchedData.length}`, 14, 54);
        doc.text(`Unmatched LMRB: ${this.data.unmatchedLMRB.length}`, 14, 61);
        doc.text(`Unmatched TC: ${this.data.unmatchedTC.length}`, 14, 68);

        const matchRate = this.data.lmrbData.length > 0
            ? ((this.data.matchedData.length / this.data.lmrbData.length) * 100).toFixed(1)
            : 0;
        doc.text(`Match Rate: ${matchRate}%`, 14, 75);

        doc.save(`reconciliation_summary_${new Date().toISOString().split('T')[0]}.pdf`);
    },

    /**
     * Show status message
     */
    showStatus(elementId, message, type) {
        const element = document.getElementById(elementId);
        element.textContent = message;
        element.className = `status-message ${type}`;
        element.style.display = 'block';
    },

    /**
     * Update all counts in UI
     */
    updateAllCounts() {
        document.getElementById('lmrbCount').textContent = this.data.lmrbData.length;
        document.getElementById('tcCount').textContent = this.data.tcData.length;
        document.getElementById('mappingCount').textContent = this.data.mappings.length;

        this.displayMappings();
    },

    /**
     * Load data from localStorage
     */
    loadFromStorage() {
        this.data.lmrbData = storage.load('lmrbData') || [];
        this.data.tcData = storage.load('tcData') || [];
        this.data.mappings = storage.load('mappings') || [];
        this.data.matchedData = storage.load('matchedData') || [];
        this.data.unmatchedLMRB = storage.load('unmatchedLMRB') || [];
        this.data.unmatchedTC = storage.load('unmatchedTC') || [];

        // Restore API key and Sheet ID
        const apiKey = storage.load('apiKey');
        const sheetId = storage.load('sheetId');
        if (apiKey) document.getElementById('apiKey').value = apiKey;
        if (sheetId) document.getElementById('sheetId').value = sheetId;
    }
};
