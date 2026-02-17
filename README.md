# Media Reconciliation System

> Modern web application for reconciling Telecast Certificate (TC) data with LMRB broadcast records

## Overview

The Media Reconciliation System is a comprehensive solution for matching television advertisement broadcast data (LMRB) with Telecast Certificates (TC). The system features Google Sheets integration, automatic PDF conversion, intelligent matching algorithms, and comprehensive reporting.

### Key Features

- ✅ **Google Sheets Integration** - Fetch LMRB data directly from Google Sheets with date range filtering
- ✅ **TC PDF Auto-Conversion** - Upload PDF files and automatically extract structured data
- ✅ **Smart Theme Mapping** - Configure flexible mappings between TC and LMRB themes
- ✅ **Intelligent Matching** - Advanced reconciliation algorithm with time tolerance and date options
- ✅ **Manual Matching** - Manually match unmatched records through intuitive UI
- ✅ **Export Functionality** - Export results to Excel and PDF formats
- ✅ **Modern UI** - Clean, responsive interface built with React and Material-UI
- ✅ **No Schedule Dependency** - Streamlined workflow without schedule upload requirements

## Architecture

### Backend (Flask + Python)
- **Framework**: Flask 3.0
- **PDF Processing**: pdfplumber
- **Data Processing**: pandas, openpyxl
- **Google API**: google-api-python-client
- **Report Generation**: reportlab

### Frontend (React + Material-UI)
- **Framework**: React 18
- **UI Library**: Material-UI (MUI) 5
- **Routing**: React Router v6
- **Build Tool**: Vite
- **Date Handling**: dayjs, @mui/x-date-pickers

## Installation

### Prerequisites

- Python 3.8 or higher
- Node.js 16 or higher
- Google Cloud Project with Sheets API enabled (for production)

### Backend Setup

```bash
# Navigate to backend directory
cd backend

# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Linux/Mac:
source venv/bin/activate
# On Windows:
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Google Sheets Setup (Optional but Recommended)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable **Google Sheets API**
4. Create **Service Account** credentials
5. Download JSON key file
6. Save as `backend/credentials/service-account.json`

**Or** set environment variable:
```bash
export GOOGLE_CREDENTIALS_PATH=/path/to/service-account.json
```

> **Note**: If you don't have credentials, the system will use mock data for testing.

### Frontend Setup

```bash
# Navigate to frontend directory
cd frontend

# Install dependencies
npm install
```

## Running the Application

### Start Backend Server

```bash
cd backend
python app.py
```

Backend will start on `http://localhost:5000`

### Start Frontend Development Server

```bash
cd frontend
npm run dev
```

Frontend will start on `http://localhost:3000`

### Access the Application

Open your browser and navigate to:
```
http://localhost:3000
```

## Usage Guide

### Step 1: Connect to Google Sheet

1. Open the Dashboard
2. Enter your Google Sheet ID (found in the URL of your Google Sheet)
   - Example URL: `https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit`
3. Click "Connect to Google Sheet"

### Step 2: Load LMRB Data & Upload TC PDF

1. Navigate to **Upload TC** page
2. Select date range for LMRB data
3. Optionally filter by channel
4. Click "Load LMRB Data"
5. Upload TC PDF file
6. Click "Upload & Convert PDF"
7. Optionally filter by specific products

### Step 3: Configure Theme Mappings

1. Navigate to **Theme Mapping** page
2. Add mappings between TC themes and LMRB themes:
   - Select mapping type (Commercial Benefits / Sponsorship)
   - Enter TC theme and duration
   - Enter corresponding LMRB theme and duration
3. Click "Add Mapping"
4. Repeat for all theme relationships

> **Tip**: You can map multiple TC themes to the same LMRB theme for flexible matching

### Step 4: Run Reconciliation

1. Navigate to **Reconciliation** page
2. Configure matching parameters:
   - **Time Tolerance**: Maximum time difference in seconds (default: 30)
   - **Ignore Date**: Match only on time and theme, ignore date
3. Click "Run Reconciliation"
4. View summary statistics

### Step 5: Review Results

1. Navigate to **Results** page
2. View tabs:
   - **Matched**: Successfully matched records
   - **Unmatched LMRB**: LMRB records without TC match
   - **Unmatched TC**: TC records without LMRB match
3. Export results:
   - **Export Excel**: Full data with all sheets
   - **Export PDF**: Summary report

### Step 6: Manual Matching (Optional)

1. Navigate to **Manual Matching** page
2. Select one unmatched LMRB record
3. Select one unmatched TC record
4. Click "Match Selected Records"

## API Documentation

See [backend/README.md](backend/README.md) for detailed API documentation.

## Project Structure

```
Reconciliation/
├── backend/                 # Flask backend
│   ├── app/
│   │   ├── __init__.py
│   │   ├── google_sheets.py    # Google Sheets integration
│   │   └── pdf_converter.py    # TC PDF conversion
│   ├── app.py                   # Main Flask application
│   ├── requirements.txt         # Python dependencies
│   └── README.md               # Backend documentation
│
├── frontend/                # React frontend
│   ├── src/
│   │   ├── components/
│   │   │   └── Layout.jsx      # Main layout with navigation
│   │   ├── pages/
│   │   │   ├── Dashboard.jsx
│   │   │   ├── Upload.jsx
│   │   │   ├── ThemeMapping.jsx
│   │   │   ├── Reconciliation.jsx
│   │   │   ├── Results.jsx
│   │   │   └── ManualMatching.jsx
│   │   ├── services/
│   │   │   └── api.js           # API service layer
│   │   ├── App.jsx
│   │   ├── main.jsx
│   │   └── index.css
│   ├── package.json
│   ├── vite.config.js
│   └── index.html
│
├── func/                    # Shared business logic
│   ├── matching.py          # Reconciliation algorithm
│   ├── format.py            # Data formatting
│   ├── create_pdf.py        # PDF report generation
│   └── ...
│
├── Main.py                  # Legacy Tkinter app (deprecated)
└── README.md               # This file
```

## Key Changes from Old System

### ✅ What's New

1. **Web-Based**: Modern React web app instead of Tkinter desktop app
2. **Google Sheets**: Direct integration - no need to upload LMRB files manually
3. **Date Range Selection**: Filter LMRB data by date range
4. **Schedule Removed**: Simplified workflow without schedule upload/matching
5. **Modern UI**: Clean, responsive Material-UI design
6. **Better UX**: Step-by-step wizard-style workflow

### 🗑️ What's Removed

1. **Schedule Upload Tab**: Completely removed
2. **Schedule Matching**: No longer needed
3. **File-based LMRB Upload**: Replaced with Google Sheets
4. **Tkinter Desktop UI**: Replaced with React web interface

## Development

### Building for Production

**Frontend:**
```bash
cd frontend
npm run build
```

Build output will be in `frontend/dist/`

**Backend:**

The Flask backend can be deployed using:
- Gunicorn (recommended)
- uWSGI
- Docker

Example with Gunicorn:
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Environment Variables

Create `.env` file in backend directory:

```env
GOOGLE_CREDENTIALS_PATH=/path/to/service-account.json
FLASK_ENV=production
SECRET_KEY=your-secret-key
```

## Troubleshooting

### Google Sheets Connection Failed

- Verify Sheet ID is correct
- Check service account has access to the sheet
- Share your Google Sheet with the service account email
- Verify credentials file path

### TC PDF Conversion Failed

- Ensure PDF is a valid Telecast Certificate
- Check PDF text is extractable (not scanned image)
- Verify PDF format matches expected structure

### No Matches Found

- Check theme mappings are correctly configured
- Adjust time tolerance (increase if needed)
- Verify date formats are consistent
- Try enabling "Ignore Date" option

## Contributing

This system is designed for internal use. For issues or feature requests, contact the development team.

## License

Internal use only - All rights reserved

## Support

For questions or issues:
- Check the troubleshooting section above
- Review API documentation in `backend/README.md`
- Contact system administrator

---

**Version**: 2.0.0 (Modern Web Application)
**Last Updated**: 2024
**Maintained By**: Development Team
