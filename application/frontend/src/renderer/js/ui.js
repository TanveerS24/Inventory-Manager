/**
 * UI Service - Handles all DOM manipulation and user interactions
 */
class UIService {
    constructor() {
        this.selectedRecord = null;
        this.records = [];
        this.clerkOptions = [];
        this.statusOptions = [];
        this.filters = {};
        
        // Photos workflow data
        this.selectedDocPath = null;
        this.selectedPhotosFolder = null;
    }

    // Initialize UI
    init() {
        this.cacheElements();
        this.bindEvents();
        this.loadOptions();
        this.loadRecords();
    }

    // Cache DOM elements
    cacheElements() {
        // Table
        this.recordsTable = document.getElementById('records-tbody');
        this.emptyState = document.getElementById('empty-state');
        
        // Buttons
        this.btnAdd = document.getElementById('btn-add');
        this.btnSearch = document.getElementById('btn-search');
        this.btnClear = document.getElementById('btn-clear');
        this.btnEdit = document.getElementById('btn-edit');
        this.btnDelete = document.getElementById('btn-delete');
        this.btnPastePhotos = document.getElementById('btn-paste-photos');
        this.btnSelectDoc = document.getElementById('btn-select-doc');
        this.btnSelectPhotos = document.getElementById('btn-select-photos');
        this.btnGenerate = document.getElementById('btn-generate');
        
        // Info
        this.selectedInfo = document.getElementById('selected-info');
        this.connectionStatus = document.getElementById('connection-status');
        
        // Modals
        this.modalRecord = document.getElementById('modal-record');
        this.modalSearch = document.getElementById('modal-search');
        this.modalPhotos = document.getElementById('modal-photos');
        
        // Forms
        this.recordForm = document.getElementById('record-form');
        this.searchForm = document.getElementById('search-form');
        
        // Loading
        this.loadingOverlay = document.getElementById('loading-overlay');
        
        // Photos workflow
        this.docPathDisplay = document.getElementById('doc-path');
        this.photosPathDisplay = document.getElementById('photos-path');
        this.generationStatus = document.getElementById('generation-status');
        this.generationResult = document.getElementById('generation-result');
        this.resultPath = document.getElementById('result-path');
    }

    // Bind event listeners
    bindEvents() {
        // Button clicks
        this.btnAdd.addEventListener('click', () => this.openRecordModal());
        this.btnSearch.addEventListener('click', () => this.openSearchModal());
        this.btnClear.addEventListener('click', () => this.clearFilters());
        this.btnEdit.addEventListener('click', () => this.openEditModal());
        this.btnDelete.addEventListener('click', () => this.deleteRecord());
        this.btnPastePhotos.addEventListener('click', () => this.openPhotosModal());
        
        // Photos workflow
        this.btnSelectDoc.addEventListener('click', () => this.selectDocFile());
        this.btnSelectPhotos.addEventListener('click', () => this.selectPhotosFolder());
        this.btnGenerate.addEventListener('click', () => this.generateReport());
        
        // Form submissions
        this.recordForm.addEventListener('submit', (e) => this.handleRecordSubmit(e));
        this.searchForm.addEventListener('submit', (e) => this.handleSearchSubmit(e));
        
        // Modal close buttons
        document.querySelectorAll('.modal-close, .modal-cancel').forEach(btn => {
            btn.addEventListener('click', (e) => this.closeModal(e.target.closest('.modal')));
        });
        
        // Close modal on outside click
        document.querySelectorAll('.modal').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) this.closeModal(modal);
            });
        });
    }

    // Load dropdown options
    async loadOptions() {
        try {
            const [clerks, statuses] = await Promise.all([
                api.getClerkOptions(),
                api.getStatusOptions()
            ]);
            
            this.clerkOptions = clerks.clerks;
            this.statusOptions = statuses.statuses;
            
            this.populateSelect('clerk', this.clerkOptions, 'Select Clerk');
            this.populateSelect('status', this.statusOptions, 'Select Status');
            this.populateSelect('search-clerk', this.clerkOptions, 'Any Clerk', true);
            this.populateSelect('search-status', this.statusOptions, 'Any Status', true);
        } catch (error) {
            console.error('Failed to load options:', error);
        }
    }

    // Populate select dropdown
    populateSelect(id, options, placeholder, includeEmpty = false) {
        const select = document.getElementById(id);
        if (!select) return;
        
        select.innerHTML = '';
        
        if (includeEmpty) {
            const emptyOption = document.createElement('option');
            emptyOption.value = '';
            emptyOption.textContent = placeholder;
            select.appendChild(emptyOption);
        } else {
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = placeholder;
            select.appendChild(defaultOption);
        }
        
        options.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            select.appendChild(opt);
        });
    }

    // Load records
    async loadRecords(params = {}) {
        this.showLoading(true);
        
        try {
            const response = await api.getRecords(params);
            this.records = response.items;
            this.renderTable();
            this.updateConnectionStatus(true);
        } catch (error) {
            console.error('Failed to load records:', error);
            this.updateConnectionStatus(false);
            this.showMessage('error', 'Error', 'Failed to load records. Is the backend running?');
        } finally {
            this.showLoading(false);
        }
    }

    // Render table
    renderTable() {
        this.recordsTable.innerHTML = '';
        
        if (this.records.length === 0) {
            this.emptyState.style.display = 'block';
            return;
        }
        
        this.emptyState.style.display = 'none';
        
        this.records.forEach((record, index) => {
            const row = document.createElement('tr');
            row.dataset.id = record.id;
            
            if (this.selectedRecord && this.selectedRecord.id === record.id) {
                row.classList.add('selected');
            }
            
            row.innerHTML = `
                <td>${index + 1}</td>
                <td>${record.date}</td>
                <td>${record.clerk}</td>
                <td class="col-address">${record.property_address}</td>
                <td>${record.client}</td>
                <td>${record.inv_type}</td>
                <td>
                    <span class="status-badge ${record.status.toLowerCase().replace(' ', '-')}">
                        ${record.status}
                    </span>
                </td>
            `;
            
            row.addEventListener('click', () => this.selectRecord(record, row));
            this.recordsTable.appendChild(row);
        });
    }

    // Select record
    selectRecord(record, rowElement) {
        // Remove previous selection
        this.recordsTable.querySelectorAll('tr').forEach(row => {
            row.classList.remove('selected');
        });
        
        // Add selection to clicked row
        rowElement.classList.add('selected');
        this.selectedRecord = record;
        
        // Update UI
        this.selectedInfo.textContent = `Selected: #${record.id} - ${record.client} (${record.property_address})`;
        this.btnEdit.disabled = false;
        this.btnDelete.disabled = false;
        
        // Disable Paste Photos if already completed
        if (record.status === 'Completed') {
            this.btnPastePhotos.disabled = true;
            this.btnPastePhotos.title = 'Record already completed';
        } else {
            this.btnPastePhotos.disabled = false;
            this.btnPastePhotos.title = '';
        }
    }

    // Reset action buttons
    resetSelection() {
        this.selectedRecord = null;
        this.recordsTable.querySelectorAll('tr').forEach(row => {
            row.classList.remove('selected');
        });
        this.selectedInfo.textContent = 'No record selected';
        this.btnEdit.disabled = true;
        this.btnDelete.disabled = true;
        this.btnPastePhotos.disabled = true;
    }

    // Modal functions
    openModal(modal) {
        modal.classList.add('active');
    }

    closeModal(modal) {
        modal.classList.remove('active');
        
        // Reset forms
        if (modal === this.modalRecord) {
            this.recordForm.reset();
            document.getElementById('record-id').value = '';
            document.getElementById('modal-title').textContent = 'Add Record';
        } else if (modal === this.modalSearch) {
            this.searchForm.reset();
        } else if (modal === this.modalPhotos) {
            this.resetPhotosWorkflow();
        }
    }

    openRecordModal() {
        this.recordForm.reset();
        document.getElementById('record-id').value = '';
        document.getElementById('modal-title').textContent = 'Add Record';
        this.openModal(this.modalRecord);
    }

    openEditModal() {
        if (!this.selectedRecord) return;
        
        const record = this.selectedRecord;
        document.getElementById('record-id').value = record.id;
        document.getElementById('clerk').value = record.clerk;
        document.getElementById('address').value = record.property_address;
        document.getElementById('client').value = record.client;
        document.getElementById('inv-type').value = record.inv_type;
        document.getElementById('status').value = record.status;
        
        // Disable status if completed
        const statusSelect = document.getElementById('status');
        if (record.status === 'Completed') {
            statusSelect.disabled = true;
        } else {
            statusSelect.disabled = false;
        }
        
        document.getElementById('modal-title').textContent = 'Edit Record';
        this.openModal(this.modalRecord);
    }

    openSearchModal() {
        this.openModal(this.modalSearch);
    }

    openPhotosModal() {
        if (!this.selectedRecord) return;
        this.resetPhotosWorkflow();
        this.openModal(this.modalPhotos);
    }

    // Form handlers
    async handleRecordSubmit(e) {
        e.preventDefault();
        
        const id = document.getElementById('record-id').value;
        const data = {
            clerk: document.getElementById('clerk').value,
            property_address: document.getElementById('address').value,
            client: document.getElementById('client').value,
            inv_type: document.getElementById('inv-type').value,
            status: document.getElementById('status').value
        };
        
        this.showLoading(true);
        
        try {
            if (id) {
                await api.updateRecord(id, data);
                this.showMessage('info', 'Success', 'Record updated successfully');
            } else {
                await api.createRecord(data);
                this.showMessage('info', 'Success', 'Record added successfully');
            }
            
            this.closeModal(this.modalRecord);
            this.loadRecords();
        } catch (error) {
            this.showMessage('error', 'Error', error.message);
        } finally {
            this.showLoading(false);
        }
    }

    async handleSearchSubmit(e) {
        e.preventDefault();
        
        const params = {
            client: document.getElementById('search-client').value,
            clerk: document.getElementById('search-clerk').value,
            address: document.getElementById('search-address').value,
            status: document.getElementById('search-status').value
        };
        
        // Remove empty params
        Object.keys(params).forEach(key => {
            if (!params[key]) delete params[key];
        });
        
        this.filters = params;
        this.btnClear.disabled = Object.keys(params).length === 0;
        
        this.closeModal(this.modalSearch);
        this.loadRecords(params);
    }

    async deleteRecord() {
        if (!this.selectedRecord) return;
        
        const confirmed = await this.showConfirm(
            'Confirm Delete',
            `Are you sure you want to delete record #${this.selectedRecord.id}?`
        );
        
        if (!confirmed) return;
        
        this.showLoading(true);
        
        try {
            await api.deleteRecord(this.selectedRecord.id);
            this.showMessage('info', 'Success', 'Record deleted successfully');
            this.resetSelection();
            this.loadRecords();
        } catch (error) {
            this.showMessage('error', 'Error', error.message);
        } finally {
            this.showLoading(false);
        }
    }

    clearFilters() {
        this.filters = {};
        this.btnClear.disabled = true;
        this.loadRecords();
    }

    // Photos workflow
    async selectDocFile() {
        if (!window.electronAPI) {
            // Fallback for browser testing
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.docx,.doc';
            input.onchange = (e) => {
                const file = e.target.files[0];
                if (file) {
                    this.selectedDocPath = file;
                    this.docPathDisplay.textContent = file.name;
                    this.updateGenerateButton();
                }
            };
            input.click();
            return;
        }
        
        const path = await window.electronAPI.selectFile();
        if (path) {
            this.selectedDocPath = path;
            this.docPathDisplay.textContent = path;
            this.updateGenerateButton();
        }
    }

    async selectPhotosFolder() {
        if (!window.electronAPI) {
            this.showMessage('info', 'Note', 'Folder selection requires Electron. Please run the desktop app.');
            return;
        }
        
        const path = await window.electronAPI.selectFolder();
        if (path) {
            this.selectedPhotosFolder = path;
            this.photosPathDisplay.textContent = path;
            this.updateGenerateButton();
        }
    }

    updateGenerateButton() {
        const canGenerate = this.selectedDocPath && this.selectedPhotosFolder;
        this.btnGenerate.disabled = !canGenerate;
    }

    resetPhotosWorkflow() {
        this.selectedDocPath = null;
        this.selectedPhotosFolder = null;
        this.docPathDisplay.textContent = 'No file selected';
        this.photosPathDisplay.textContent = 'No folder selected';
        this.btnGenerate.disabled = true;
        this.generationStatus.style.display = 'none';
        this.generationResult.style.display = 'none';
    }

    async generateReport() {
        if (!this.selectedRecord || !this.selectedDocPath || !this.selectedPhotosFolder) return;
        
        this.btnGenerate.disabled = true;
        this.generationStatus.style.display = 'block';
        this.generationResult.style.display = 'none';
        
        try {
            // For Electron, we need to read the file
            let fileToUpload;
            
            if (typeof this.selectedDocPath === 'string') {
                // Electron path - fetch and convert to File
                const response = await fetch(`file://${this.selectedDocPath}`);
                const blob = await response.blob();
                const filename = this.selectedDocPath.split('\\').pop().split('/').pop();
                fileToUpload = new File([blob], filename, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            } else {
                // Already a File object (browser fallback)
                fileToUpload = this.selectedDocPath;
            }
            
            const result = await api.generateReport(
                this.selectedRecord.id,
                fileToUpload,
                this.selectedPhotosFolder
            );
            
            this.generationStatus.style.display = 'none';
            this.generationResult.style.display = 'block';
            this.resultPath.textContent = result.document_path;
            
            // Update record status in UI
            this.selectedRecord.status = 'Completed';
            this.renderTable();
            
            this.showMessage('info', 'Success', 'Report generated successfully!');
            
        } catch (error) {
            this.generationStatus.style.display = 'none';
            this.btnGenerate.disabled = false;
            this.showMessage('error', 'Error', error.message);
        }
    }

    // Utility functions
    showLoading(show) {
        this.loadingOverlay.classList.toggle('active', show);
    }

    async showMessage(type, title, message) {
        if (window.electronAPI) {
            await window.electronAPI.showMessage(type, title, message);
        } else {
            alert(`${title}: ${message}`);
        }
    }

    async showConfirm(title, message) {
        return confirm(`${title}\n\n${message}`);
    }

    updateConnectionStatus(connected) {
        this.connectionStatus.textContent = connected ? 'Online' : 'Offline';
        this.connectionStatus.className = `status-indicator ${connected ? 'online' : 'offline'}`;
    }
}

// Export singleton
const ui = new UIService();
