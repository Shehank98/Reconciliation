# Media Reconciliation Backend API

Modern Flask-based backend for the Media Reconciliation System.

## Features

- **Google Sheets Integration**: Fetch LMRB data directly from Google Sheets
- **TC PDF Conversion**: Upload and convert TC PDF files to structured data
- **Theme Mapping**: Configure mappings between TC and LMRB themes
- **Reconciliation Engine**: Match TC records against LMRB data
- **Export Functionality**: Export results to Excel and PDF
- **Manual Matching**: Manually match unmatched records

## Setup

### 1. Install Dependencies

```bash
cd backend
pip install -r requirements.txt
```

### 2. Google Sheets Credentials

To use Google Sheets integration:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable Google Sheets API
4. Create Service Account credentials
5. Download JSON key file
6. Save as `credentials/service-account.json`

**Or** set environment variable:
```bash
export GOOGLE_CREDENTIALS_PATH=/path/to/service-account.json
```

### 3. Run the Server

```bash
python app.py
```

Server will start on `http://localhost:5000`

## API Endpoints

### Health Check
```
GET /api/health
```

### Google Sheets

**Connect to Google Sheet:**
```
POST /api/google-sheets/connect
Body: { "sheet_id": "your-sheet-id", "session_id": "unique-id" }
```

**Get LMRB Data:**
```
POST /api/google-sheets/lmrb-data
Body: {
  "session_id": "unique-id",
  "start_date": "2024-01-01",
  "end_date": "2024-12-31",
  "channel": "Sirasa TV"
}
```

### TC Upload

**Upload TC PDF:**
```
POST /api/upload-tc
Form Data:
  - file: PDF file
  - session_id: unique-id
```

### Theme Mappings

**Get Mappings:**
```
GET /api/theme-mappings?session_id=unique-id
```

**Add Mapping:**
```
POST /api/theme-mappings?session_id=unique-id
Body: {
  "tc_theme": "Coca Cola",
  "tc_duration": 30,
  "lmrb_theme": "Beverages",
  "lmrb_duration": 30,
  "mapping_type": "COMMERCIAL BENEFITS"
}
```

**Delete Mapping:**
```
DELETE /api/theme-mappings?session_id=unique-id
Body: { "index": 0 }
```

### Reconciliation

**Run Matching:**
```
POST /api/reconcile
Body: {
  "session_id": "unique-id",
  "time_tolerance": 30,
  "ignore_date": false,
  "products": ["Product1", "Product2"]
}
```

**Get Matched Results:**
```
GET /api/results/matched?session_id=unique-id&page=1&per_page=50
```

**Get Unmatched Results:**
```
GET /api/results/unmatched?session_id=unique-id&type=lmrb
```

### Export

**Export to Excel:**
```
POST /api/export/excel
Body: { "session_id": "unique-id" }
```

**Export to PDF:**
```
POST /api/export/pdf
Body: { "session_id": "unique-id", "channel": "Sirasa TV" }
```

### Manual Matching

**Manually Match Records:**
```
POST /api/manual-match
Body: {
  "session_id": "unique-id",
  "lmrb_index": 0,
  "tc_index": 0
}
```

## Development

### Without Google Sheets Credentials

The API will use mock data for testing if credentials are not available. This allows you to develop and test the frontend without Google Sheets access.

### Session Management

Currently uses in-memory sessions. For production:
- Use Redis for session storage
- Add session expiration
- Implement user authentication

## Error Handling

All endpoints return JSON responses:

**Success:**
```json
{
  "success": true,
  "data": { ... }
}
```

**Error:**
```json
{
  "error": "Error message"
}
```

HTTP status codes:
- 200: Success
- 400: Bad Request (missing parameters, invalid data)
- 500: Server Error
