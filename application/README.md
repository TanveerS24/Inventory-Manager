# InventoryHouse Pro

A modern, full-stack inventory management system built with **FastAPI** (backend) and **Electron** (frontend).

## Features

- **Record Management**: Add, edit, delete, and search property inventory records
- **Document Generation**: Automatic generation of professional reports combining:
  - Branded template front page with company logo
  - Transcribed audio document (middle)
  - Photo index with numbered captions
- **Photo Workflow**: 4×2 grid layout, 8 photos per page, auto-numbered
- **Modern UI**: Clean, responsive interface with soft color theme
- **Desktop Application**: Native desktop experience via Electron

## Architecture

```
application/
├── backend/                    # FastAPI REST API
│   ├── app/
│   │   ├── api/               # API route handlers
│   │   ├── core/              # Configuration
│   │   ├── db/                # Database models & connection
│   │   ├── models/            # SQLAlchemy models
│   │   ├── schemas/           # Pydantic schemas
│   │   └── services/          # Business logic (document generation)
│   ├── assets/                # Logo and template images
│   ├── main.py               # FastAPI application entry
│   └── requirements.txt      # Python dependencies
│
├── frontend/                  # Electron Desktop App
│   ├── src/
│   │   ├── main/             # Main process (window management)
│   │   ├── preload/          # Preload scripts (IPC)
│   │   └── renderer/         # Renderer process (UI)
│   │       ├── index.html    # Main UI
│   │       ├── styles/       # CSS stylesheets
│   │       └── js/           # JavaScript modules
│   │           ├── api.js    # Backend API client
│   │           ├── ui.js     # UI controller
│   │           └── app.js    # Application entry
│   ├── assets/               # Application assets
│   └── package.json          # Node.js configuration
│
├── start.bat                # Windows startup script (CMD)
├── start.ps1                # Windows startup script (PowerShell)
└── README.md                # This file
```

## Prerequisites

- **Python** 3.9 or higher
- **Node.js** 16 or higher
- **Windows** (for desktop file dialogs)

## Installation

### 1. Clone or Extract the Application

Ensure the `application` folder is in your desired location.

### 2. Install Python Dependencies

```bash
cd application/backend
python -m venv venv

# Windows
venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Install Node.js Dependencies

```bash
cd application/frontend
npm install
```

## Running the Application

### Option 1: Using the Startup Script (Recommended)

```bash
# Using Command Prompt
cd application
start.bat

# Using PowerShell
cd application
.\start.ps1
```

### Option 2: Manual Start

**Terminal 1 - Backend:**
```bash
cd application/backend
venv\Scripts\activate
uvicorn main:app --host 127.0.0.1 --port 8000 --reload
```

**Terminal 2 - Frontend:**
```bash
cd application/frontend
npm start
```

### Access Points

- **Electron Desktop App**: Opens automatically
- **Backend API**: http://127.0.0.1:8000
- **API Documentation**: http://127.0.0.1:8000/api/v1/docs

## API Endpoints

### Records
- `GET /api/v1/records/` - List all records (with optional filters)
- `POST /api/v1/records/` - Create new record
- `GET /api/v1/records/{id}` - Get single record
- `PUT /api/v1/records/{id}` - Update record
- `DELETE /api/v1/records/{id}` - Delete record

### Documents
- `POST /api/v1/documents/generate-report/{record_id}` - Generate complete report

### Options
- `GET /api/v1/records/options/clerks` - Get clerk options
- `GET /api/v1/records/options/statuses` - Get status options

## Document Generation Workflow

The "Paste Photos" feature automates the complete document pipeline:

1. **Select Transcription** - Choose the middle Word document (.docx)
2. **Select Photos** - Choose folder containing property images
3. **Generate Report** - Backend processes:
   - Forces landscape orientation on transcription
   - Generates branded template front page
   - Creates photo index (4×2 grid, 8 per page)
   - Merges all documents in order: Template → Transcription → Photos
   - Saves final DOCX in photos folder
   - Updates record status to "Completed"

## Technology Stack

### Backend
- **FastAPI** - Modern, fast web framework
- **SQLAlchemy** - ORM for SQLite database
- **Pydantic** - Data validation
- **python-docx** - Word document generation
- **docxcompose** - Document merging
- **Pillow** - Image processing

### Frontend
- **Electron** - Cross-platform desktop framework
- **HTML5/CSS3** - Modern web technologies
- **Vanilla JavaScript** - No framework dependencies

## Data Model

```
PropertyRecord
├── id (Integer, PK)
├── date (String) - Creation date (DD-MM-YYYY)
├── clerk (String) - Inspector name
├── property_address (String)
├── client (String)
├── inv_type (String) - Inventory type
├── status (String) - Inspected / Audio Recorded / Completed
├── final_doc_path (String) - Path to generated document
├── created_at (DateTime)
└── updated_at (DateTime)
```

## Configuration

Edit `backend/app/core/config.py` to customize:

- Company information (name, phone, address)
- Clerk options list
- Photo grid settings (images per page, dimensions)
- Color scheme variables (in frontend CSS)

## Troubleshooting

### Backend won't start
- Ensure port 8000 is not in use: `netstat -ano | findstr 8000`
- Check Python version: `python --version` (must be 3.9+)

### Frontend won't start
- Ensure Node.js is installed: `node --version`
- Delete `node_modules` and run `npm install` again

### Document generation fails
- Ensure Microsoft Word is installed (for DOCX compatibility)
- Check that photo folder contains valid images (.jpg, .png)
- Verify transcription file is valid .docx format

## Building for Distribution

### Build Electron App

```bash
cd application/frontend
npm run build
```

Output will be in `frontend/dist/`.

### Package Backend

Use PyInstaller or similar to create standalone executable:

```bash
cd application/backend
pip install pyinstaller
pyinstaller --onefile --add-data "assets;assets" main.py
```

## Credits

Built with the original InventoryHouse logic, refactored into a modern full-stack architecture.

- **Backend**: FastAPI with original python-docx document generation logic
- **Frontend**: Electron with clean, responsive UI design
- **Document Engine**: Preserves all original photo grid and template generation code

## License

MIT License - See LICENSE file for details.
