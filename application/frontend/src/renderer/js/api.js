/**
 * API Service - Handles all backend communication
 */
const API_BASE_URL = 'http://127.0.0.1:8000/api/v1';

class ApiService {
    constructor() {
        this.baseURL = API_BASE_URL;
    }

    async request(endpoint, options = {}) {
        const url = `${this.baseURL}${endpoint}`;
        
        const config = {
            headers: {
                'Content-Type': 'application/json',
            },
            ...options
        };

        try {
            const response = await fetch(url, config);
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.detail || `HTTP ${response.status}`);
            }

            return await response.json();
        } catch (error) {
            console.error('API Error:', error);
            throw error;
        }
    }

    // Records API
    async getRecords(params = {}) {
        const queryString = new URLSearchParams(params).toString();
        const endpoint = queryString ? `/records/?${queryString}` : '/records/';
        return this.request(endpoint);
    }

    async getRecord(id) {
        return this.request(`/records/${id}`);
    }

    async createRecord(data) {
        return this.request('/records/', {
            method: 'POST',
            body: JSON.stringify(data)
        });
    }

    async updateRecord(id, data) {
        return this.request(`/records/${id}`, {
            method: 'PUT',
            body: JSON.stringify(data)
        });
    }

    async deleteRecord(id) {
        return this.request(`/records/${id}`, {
            method: 'DELETE'
        });
    }

    // Options API
    async getClerkOptions() {
        return this.request('/records/options/clerks');
    }

    async getStatusOptions() {
        return this.request('/records/options/statuses');
    }

    // Document Generation API
    async generateReport(recordId, middleDocFile, photosFolder) {
        const formData = new FormData();
        formData.append('middle_doc', middleDocFile);
        formData.append('photos_folder', photosFolder);

        const response = await fetch(
            `${this.baseURL}/documents/generate-report/${recordId}`,
            {
                method: 'POST',
                body: formData
            }
        );

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || 'Failed to generate report');
        }

        return await response.json();
    }

    // Health Check
    async healthCheck() {
        try {
            const response = await fetch('http://127.0.0.1:8000/health');
            return response.ok;
        } catch {
            return false;
        }
    }
}

// Export singleton
const api = new ApiService();
