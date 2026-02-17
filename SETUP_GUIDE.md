# Quick Setup Guide

## Prerequisites Checklist

- [ ] Python 3.8+ installed
- [ ] Node.js 16+ installed
- [ ] Google Sheet with LMRB data (optional but recommended)
- [ ] Google Cloud service account (optional, for Google Sheets)

## Quick Start (5 Minutes)

### 1. Backend Setup

```bash
# Clone and navigate
cd backend

# Install Python dependencies
pip install -r requirements.txt

# Start backend server
python app.py
```

Backend running at: `http://localhost:5000` ✅

### 2. Frontend Setup

**Open new terminal:**

```bash
# Navigate to frontend
cd frontend

# Install Node dependencies
npm install

# Start development server
npm run dev
```

Frontend running at: `http://localhost:3000` ✅

### 3. Access Application

Open browser: `http://localhost:3000`

## Google Sheets Setup (Optional)

### If You Have a Google Sheet:

1. **Get Sheet ID** from URL:
   ```
   https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit
   ```

2. **Create Service Account**:
   - Go to https://console.cloud.google.com/
   - Create project → Enable Google Sheets API
   - Create Service Account → Download JSON key
   - Save as `backend/credentials/service-account.json`

3. **Share Sheet**:
   - Open your Google Sheet
   - Click "Share"
   - Add service account email (from JSON file)
   - Give "Viewer" permissions

### Without Google Sheets:

The system will use mock data for testing. You can still test all features!

## First Time Usage

1. **Dashboard**: Enter Google Sheet ID (or skip to use mock data)
2. **Upload**:
   - Select date range
   - Upload TC PDF file
3. **Theme Mapping**: Add theme relationships
4. **Reconciliation**: Run matching
5. **Results**: View and export

## Common Issues

### Backend won't start:
```bash
# Check Python version
python --version  # Should be 3.8+

# Reinstall dependencies
pip install -r requirements.txt --force-reinstall
```

### Frontend won't start:
```bash
# Check Node version
node --version  # Should be 16+

# Clear and reinstall
rm -rf node_modules package-lock.json
npm install
```

### Google Sheets error:
- Verify Sheet ID is correct
- Check service account has access to sheet
- Use mock data mode for testing without credentials

## Need Help?

- Check main README.md for detailed documentation
- Review backend/README.md for API details
- Ensure both servers are running simultaneously

## Next Steps

Once running:
1. Connect to your Google Sheet
2. Upload a sample TC PDF
3. Configure theme mappings
4. Run reconciliation
5. Export results

---

**Quick Reference:**
- Backend: `http://localhost:5000`
- Frontend: `http://localhost:3000`
- API Health: `http://localhost:5000/api/health`
