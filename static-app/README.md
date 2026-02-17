# Media Reconciliation System - Static Web App

> 100% Client-Side Application - No Backend Required!

## Overview

This is a **pure static web application** that runs entirely in your browser. Perfect for GitHub Pages deployment with **zero server costs**.

### ✨ Key Features

- ✅ **No Backend** - All processing happens in the browser
- ✅ **GitHub Pages Ready** - Deploy for free in minutes
- ✅ **Google Sheets Integration** - Direct API access to LMRB data
- ✅ **Channel-Specific PDF Converters** - Modular system for each channel
- ✅ **Offline Capable** - Data stored in browser localStorage
- ✅ **Zero Cost** - No hosting fees, no server maintenance

## Quick Start

### Option 1: Open Locally

1. Download all files
2. Open `index.html` in your browser
3. That's it! No installation needed.

### Option 2: Deploy to GitHub Pages

1. Push code to GitHub repository
2. Go to repository Settings → Pages
3. Set Source: "Deploy from a branch"
4. Select branch: `claude/remove-schedule-section-LrGLj`
5. Click Save
6. Your app will be live at: `https://yourusername.github.io/Reconciliation/static-app/`

## Setup Guide

### 1. Google Sheets API Key

To access Google Sheets, you need an API key:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable **Google Sheets API**
4. Go to **Credentials** → Create Credentials → API Key
5. Copy your API key
6. **Make your Google Sheet publicly viewable** (Share → Anyone with link can view)

### 2. Google Sheet ID

From your Google Sheet URL:
```
https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit
```
Copy the `YOUR_SHEET_ID_HERE` part.

## Usage

### Step 1: Setup (First Tab)

1. Enter your Google Sheets API Key
2. Enter your Google Sheet ID
3. Select date range for LMRB data
4. Click "Load LMRB Data"

### Step 2: Upload TC (Second Tab)

1. Select your TV channel (Sirasa, Hiru, Derana, ITN, TV1)
2. Click "Select TC PDF File"
3. Choose your TC PDF file
4. Click "Convert PDF to Data"

### Step 3: Theme Mapping (Third Tab)

1. Select mapping type (Commercial Benefits or Sponsorship)
2. Enter TC Theme and select duration
3. Enter corresponding LMRB Theme and duration
4. Click "Add Mapping"
5. Repeat for all theme relationships

### Step 4: Reconciliation (Fourth Tab)

1. Set time tolerance (default: 30 seconds)
2. Optionally check "Ignore Date"
3. Click "Run Reconciliation"
4. View summary statistics

### Step 5: Results (Fifth Tab)

1. View matched records
2. View unmatched LMRB/TC records
3. Export to Excel or PDF

### Step 6: Manual Matching (Sixth Tab)

1. Select unmatched LMRB record
2. Select unmatched TC record
3. Click "Match Selected"

## Customizing Channel Converters

Each channel has its own PDF converter in `js/converters/`:

- **sirasa.js** - Fully implemented (use as reference)
- **hiru.js** - Template (customize for Hiru TV)
- **derana.js** - Template (customize for Derana TV)
- **itv.js** - Template (customize for ITN)
- **tv1.js** - Template (customize for TV1)

### How to Customize a Converter

1. Open the converter file (e.g., `js/converters/hiru.js`)
2. Modify the `parsePage()` method:

```javascript
parsePage(text) {
    const records = [];
    const lines = text.split('\n');

    lines.forEach(line => {
        // Update this pattern to match your PDF format
        const pattern = /YOUR_REGEX_PATTERN_HERE/;
        const match = line.match(pattern);

        if (match) {
            // Extract fields based on your pattern
            records.push({
                Date: '...',      // Extract date
                Program: '...',   // Extract program
                Time: '...',      // Extract time
                Theme: '...',     // Extract theme
                Duration: ...     // Extract duration
            });
        }
    });

    return records;
}
```

3. Test with a sample PDF from that channel
4. Adjust patterns as needed

### Example: Sirasa TV Pattern

```javascript
// Pattern: 01/01/2024 News 10:30:15:00 Coca Cola 30
const pattern = /(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(\d{2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)\s*$/;
```

## File Structure

```
static-app/
├── index.html              # Main HTML file
├── css/
│   └── styles.css         # All styling
├── js/
│   ├── app.js             # Main application logic
│   ├── utils/
│   │   ├── storage.js     # LocalStorage management
│   │   ├── googleSheets.js # Google Sheets API
│   │   └── reconciliation.js # Matching algorithm
│   └── converters/
│       ├── sirasa.js      # Sirasa TV converter (complete)
│       ├── hiru.js        # Hiru TV converter (template)
│       ├── derana.js      # Derana converter (template)
│       ├── itv.js         # ITN converter (template)
│       └── tv1.js         # TV1 converter (template)
└── README.md              # This file
```

## How It Works

### Data Flow

1. **Load LMRB**: Fetch from Google Sheets API → Store in localStorage
2. **Upload TC**: Read PDF → Extract text → Parse → Store in localStorage
3. **Map Themes**: Configure mappings → Store in localStorage
4. **Reconcile**: Match records → Store results in localStorage
5. **Export**: Generate Excel/PDF from localStorage data

### Technologies Used

- **Pure JavaScript** - No frameworks, no build process
- **PDF.js** - Mozilla's PDF rendering library
- **SheetJS (XLSX)** - Excel file generation
- **jsPDF** - PDF report generation
- **Google Sheets API v4** - Direct sheet access
- **LocalStorage** - Browser-based data persistence

## Deployment

### GitHub Pages

**Automatic Deployment:**

1. Push to GitHub
2. Enable GitHub Pages in repository settings
3. Select branch and folder
4. Access at: `https://yourusername.github.io/repo-name/static-app/`

**Custom Domain (Optional):**

1. Add CNAME file with your domain
2. Configure DNS settings
3. Enable custom domain in GitHub Pages settings

### Other Static Hosts

Works with any static file hosting:

- **Netlify** - Drop folder or connect GitHub
- **Vercel** - Import repository
- **Cloudflare Pages** - Connect GitHub
- **AWS S3** - Upload to bucket with static hosting
- **Firebase Hosting** - Deploy with CLI

## Browser Compatibility

- ✅ Chrome 90+
- ✅ Firefox 88+
- ✅ Safari 14+
- ✅ Edge 90+

Requires:
- localStorage support
- FileReader API
- Fetch API
- ES6+ JavaScript

## Security Notes

1. **API Key Security**:
   - Use browser-restricted API keys
   - Restrict to your domain only
   - Enable only Google Sheets API

2. **Data Privacy**:
   - All data stored in browser localStorage
   - No data sent to external servers
   - Clearing browser data deletes all records

3. **Google Sheet Access**:
   - Sheet must be publicly viewable OR
   - Use OAuth for private sheets (requires code modification)

## Troubleshooting

### "Failed to fetch data from Google Sheets"
- Check API key is correct
- Verify Sheet ID is correct
- Ensure sheet is publicly viewable
- Check API key restrictions allow your domain

### "PDF conversion failed"
- Verify PDF is not encrypted/password protected
- Check PDF contains extractable text (not scanned images)
- Try a different PDF file
- Check browser console for specific error

### "No data found in sheet"
- Verify sheet has data in first row (headers)
- Check sheet name is "Sheet1" or update code
- Ensure data starts from row 1

### Data not persisting
- Check localStorage is enabled in browser
- Verify not in private/incognito mode
- Check localStorage quota not exceeded

## Limitations

1. **File Size**: Browser memory limits (typically 100MB for PDFs)
2. **Google Sheets**: API quota (300 requests per minute)
3. **LocalStorage**: 5-10MB limit per domain
4. **PDF Format**: Must have extractable text

## Advantages Over Backend Version

| Feature | Static App | Backend App |
|---------|-----------|-------------|
| **Hosting Cost** | FREE | $5-20/month |
| **Setup Time** | 5 minutes | 30+ minutes |
| **Dependencies** | None | Python, Node, etc |
| **Deployment** | Drag & drop | Server config |
| **Updates** | Git push | Redeploy server |
| **Scaling** | Automatic (CDN) | Manual |

## Support

For issues or questions:

1. Check troubleshooting section above
2. Review browser console for errors
3. Verify API keys and permissions
4. Test with sample data first

## License

Internal use only - All rights reserved

---

**Version**: 2.0 Static
**Last Updated**: 2024
**Deployment**: GitHub Pages Ready
