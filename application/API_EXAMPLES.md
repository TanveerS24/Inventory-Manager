# API Usage Examples

## Base URL
```
http://127.0.0.1:8000/api/v1
```

## Records API

### Get All Records
```bash
curl http://127.0.0.1:8000/api/v1/records/
```

### Search Records
```bash
# Search by client
curl "http://127.0.0.1:8000/api/v1/records/?client=Smith"

# Search by clerk
curl "http://127.0.0.1:8000/api/v1/records/?clerk=Kevin"

# Search by address
curl "http://127.0.0.1:8000/api/v1/records/?address=London"

# Search by status
curl "http://127.0.0.1:8000/api/v1/records/?status=Completed"

# Combined search
curl "http://127.0.0.1:8000/api/v1/records/?client=Smith&clerk=Kevin&status=Inspected"
```

### Create Record
```bash
curl -X POST http://127.0.0.1:8000/api/v1/records/ \
  -H "Content-Type: application/json" \
  -d '{
    "clerk": "Kevin Crack",
    "property_address": "123 Main Street, London",
    "client": "John Smith",
    "inv_type": "Inventory & Schedule",
    "status": "Inspected"
  }'
```

### Update Record
```bash
curl -X PUT http://127.0.0.1:8000/api/v1/records/1 \
  -H "Content-Type: application/json" \
  -d '{
    "clerk": "Tom Tyrrel",
    "property_address": "456 Oak Avenue, London",
    "client": "Jane Doe",
    "inv_type": "Check In"
  }'
```

### Delete Record
```bash
curl -X DELETE http://127.0.0.1:8000/api/v1/records/1
```

## Document Generation API

### Generate Complete Report
```bash
curl -X POST "http://127.0.0.1:8000/api/v1/documents/generate-report/1" \
  -F "middle_doc=@/path/to/transcription.docx" \
  -F "photos_folder=/path/to/photos/folder"
```

**Response:**
```json
{
  "success": true,
  "message": "Report generated successfully",
  "document_path": "/path/to/photos/folder/final_1_Client_Name_1234567890.docx",
  "filename": "final_1_Client_Name_1234567890.docx"
}
```

## Options API

### Get Clerk Options
```bash
curl http://127.0.0.1:8000/api/v1/records/options/clerks
```

**Response:**
```json
{
  "clerks": ["Tom Tyrrel", "Kevin Crack", "Bill West"]
}
```

### Get Status Options
```bash
curl http://127.0.0.1:8000/api/v1/records/options/statuses
```

**Response:**
```json
{
  "statuses": ["Inspected", "Audio Recorded", "Completed"]
}
```

## Health Check

```bash
curl http://127.0.0.1:8000/health
```

**Response:**
```json
{
  "status": "healthy",
  "app": "InventoryHouse Pro",
  "version": "2.0.0"
}
```

## JavaScript/Fetch Examples

### Get Records with Fetch
```javascript
async function getRecords() {
  const response = await fetch('http://127.0.0.1:8000/api/v1/records/');
  const data = await response.json();
  console.log(data.items);
  return data;
}
```

### Create Record with Fetch
```javascript
async function createRecord(recordData) {
  const response = await fetch('http://127.0.0.1:8000/api/v1/records/', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(recordData)
  });
  
  if (!response.ok) {
    throw new Error('Failed to create record');
  }
  
  return await response.json();
}

// Usage
const newRecord = {
  clerk: "Kevin Crack",
  property_address: "123 Main Street",
  client: "John Smith",
  inv_type: "Inventory",
  status: "Inspected"
};

createRecord(newRecord)
  .then(record => console.log('Created:', record))
  .catch(error => console.error('Error:', error));
```

### Generate Report with Fetch (FormData)
```javascript
async function generateReport(recordId, docFile, photosFolder) {
  const formData = new FormData();
  formData.append('middle_doc', docFile);
  formData.append('photos_folder', photosFolder);
  
  const response = await fetch(
    `http://127.0.0.1:8000/api/v1/documents/generate-report/${recordId}`,
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

// Usage with file input
const fileInput = document.getElementById('doc-input');
const file = fileInput.files[0];

generateReport(1, file, '/path/to/photos')
  .then(result => console.log('Generated:', result.document_path))
  .catch(error => console.error('Error:', error));
```

## Python Requests Examples

```python
import requests

BASE_URL = "http://127.0.0.1:8000/api/v1"

# Get all records
response = requests.get(f"{BASE_URL}/records/")
records = response.json()
print(f"Total records: {records['total']}")

# Create record
new_record = {
    "clerk": "Kevin Crack",
    "property_address": "123 Main Street, London",
    "client": "John Smith",
    "inv_type": "Inventory & Schedule",
    "status": "Inspected"
}
response = requests.post(f"{BASE_URL}/records/", json=new_record)
created = response.json()
print(f"Created record ID: {created['id']}")

# Generate report
with open('/path/to/transcription.docx', 'rb') as f:
    files = {'middle_doc': f}
    data = {'photos_folder': '/path/to/photos'}
    response = requests.post(
        f"{BASE_URL}/documents/generate-report/{created['id']}",
        files=files,
        data=data
    )
    result = response.json()
    print(f"Report saved to: {result['document_path']}")
```
