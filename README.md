# Media Reconciliation System

> 📺 **Static Web Application for Media Reconciliation**
> 100% Client-Side | Zero Backend | GitHub Pages Ready

![Status](https://img.shields.io/badge/status-active-success.svg)
![Platform](https://img.shields.io/badge/platform-web-blue.svg)
![License](https://img.shields.io/badge/license-private-red.svg)

## 🎯 Overview

A **pure static web application** for reconciling TV advertising data (TC PDFs) against LMRB data from Google Sheets. Runs entirely in your browser with **zero hosting costs**.

### ✨ Key Features

- ✅ **100% Client-Side** - No backend server required
- ✅ **GitHub Pages Ready** - Deploy for free in 5 minutes
- ✅ **Channel-Specific PDF Converters** - Modular system for Sirasa, Hiru, Derana, ITN, TV1
- ✅ **Google Sheets Integration** - Direct API access to LMRB data
- ✅ **Browser-Based Storage** - All data stored in localStorage
- ✅ **Excel & PDF Export** - Generate reports client-side
- ✅ **Mobile Responsive** - Works on desktop, tablet, and mobile

## 🚀 Quick Start

### 1. Deploy to GitHub Pages (5 Minutes)

1. **Enable GitHub Pages:**
   - Go to repository Settings → Pages
   - Source: `claude/remove-schedule-section-LrGLj` branch
   - Folder: `/ (root)`
   - Click Save

2. **Your app will be live at:**
   ```
   https://shehank98.github.io/Reconciliation/static-app/
   ```

### 2. Setup Google Sheets API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create new project
3. Enable **Google Sheets API**
4. Create **API Key** (Credentials → Create Credentials)
5. Restrict to your domain and Google Sheets API only
6. Make your Google Sheet publicly viewable

### 3. Start Using

1. Open your deployed app
2. Enter API Key and Sheet ID
3. Select date range and load LMRB data
4. Upload TC PDF (select channel first)
5. Add theme mappings
6. Run reconciliation
7. Export results!

## 📁 Project Structure

```
Reconciliation/
├── static-app/              # Main application
│   ├── index.html          # Entry point
│   ├── css/
│   │   └── styles.css      # Styling
│   ├── js/
│   │   ├── app.js          # Main controller
│   │   ├── utils/
│   │   │   ├── storage.js           # localStorage management
│   │   │   ├── googleSheets.js      # Google Sheets API
│   │   │   └── reconciliation.js    # Matching algorithm
│   │   └── converters/              # Channel-specific PDF parsers
│   │       ├── sirasa.js   # ✅ Sirasa TV (fully working)
│   │       ├── hiru.js     # ⚠️ Hiru TV (template)
│   │       ├── derana.js   # ⚠️ Derana TV (template)
│   │       ├── itv.js      # ⚠️ ITN (template)
│   │       └── tv1.js      # ⚠️ TV1 (template)
│   └── README.md           # App documentation
├── STATIC_APP_GUIDE.md     # Deployment guide
└── README.md               # This file
```

## 🔧 Channel-Specific PDF Converters

### Modular System

Each TV channel has its own PDF converter that can be customized independently.

#### ✅ Sirasa TV
- **File:** `static-app/js/converters/sirasa.js`
- **Status:** Fully implemented with multiple regex patterns
- **Format:** Date | Program | Time | Theme | Duration

#### ⚠️ Other Channels (Ready for Customization)
- **Hiru TV:** `static-app/js/converters/hiru.js`
- **Derana TV:** `static-app/js/converters/derana.js`
- **ITN:** `static-app/js/converters/itv.js`
- **TV1:** `static-app/js/converters/tv1.js`

### How to Customize a Converter

1. Open the converter file (e.g., `static-app/js/converters/hiru.js`)
2. Find the `parsePage()` method with TODO markers
3. Update the regex pattern to match your channel's PDF format
4. Test with a sample PDF
5. Commit and push - GitHub Pages auto-updates!

**Example Pattern:**
```javascript
// For format: 01/01/2024 | News | 10:30:15 | Coca Cola | 30
const pattern = /(\d{2}\/\d{2}\/\d{4})\s*\|\s*(.+?)\s*\|\s*(\d{2}:\d{2}:\d{2})\s*\|\s*(.+?)\s*\|\s*(\d+)/;
```

## 🎨 Features

### Data Management
- Load LMRB data from Google Sheets
- Filter by date range
- Upload and parse TC PDFs
- Browser-based storage (localStorage)

### Theme Mapping
- Map TC themes to LMRB themes
- Support for different durations (15s, 30s, 45s, 60s)
- Commercial Benefits and Sponsorship types
- Easy add/delete mappings

### Reconciliation
- Automatic matching based on theme, date, time, duration
- Configurable time tolerance
- Option to ignore date matching
- Manual matching interface for unmatched records

### Export
- Export to Excel (multiple sheets)
- Export summary to PDF
- Download matched/unmatched reports

## 📊 How It Works

```
┌─────────────────┐
│   Your Browser  │
│                 │
│  ┌───────────┐  │
│  │ HTML/CSS  │  │  All processing
│  │ JavaScript│  │  happens here!
│  └─────┬─────┘  │
│        │        │
└────────┼────────┘
         │
         ├──────► Google Sheets API (fetch LMRB data)
         ├──────► PDF.js (parse TC PDF)
         ├──────► localStorage (save all data)
         ├──────► XLSX.js (export Excel)
         └──────► jsPDF (export PDF)
```

**No backend server needed!** Everything runs in your browser.

## 🌐 Technologies

- **Pure JavaScript** - No frameworks, no build process
- **PDF.js** - Mozilla's PDF parsing library
- **Google Sheets API v4** - Direct sheet access
- **SheetJS (XLSX)** - Excel file generation
- **jsPDF** - PDF report generation
- **LocalStorage** - Browser data persistence
- **CSS Grid/Flexbox** - Modern responsive layout

## 💰 Cost Comparison

| Deployment Method | Monthly Cost | Setup Time |
|-------------------|--------------|------------|
| **GitHub Pages** (This app) | **$0** | **5 min** |
| Vercel + Backend | $0-7 | 30 min |
| Railway Full Stack | $5+ | 20 min |
| AWS EC2 + S3 | $10+ | 60 min |

## 📖 Documentation

- **[STATIC_APP_GUIDE.md](STATIC_APP_GUIDE.md)** - Complete deployment guide
- **[static-app/README.md](static-app/README.md)** - App documentation
- **Inline Comments** - Every JavaScript file has detailed comments

## 🔒 Security

### API Key Management
- Use restricted API keys (domain-specific)
- Enable only Google Sheets API
- Never commit API keys to repository

### Data Privacy
- All data stored in browser localStorage
- Nothing sent to external servers (except Google Sheets API)
- Clear browser data = clear all records

### Google Sheet Access
- Sheet must be publicly viewable OR
- Use OAuth for private sheets (requires code modification)

## 🐛 Troubleshooting

### "Failed to fetch data from Google Sheets"
- ✓ Check API key is correct
- ✓ Verify Sheet ID is correct
- ✓ Ensure sheet is publicly viewable
- ✓ Check API restrictions allow your domain

### "PDF conversion failed"
- ✓ Verify PDF is not password-protected
- ✓ Check PDF contains extractable text
- ✓ Select correct channel before uploading
- ✓ Check browser console for errors

### Data not persisting
- ✓ Ensure not in private/incognito mode
- ✓ Check localStorage is enabled
- ✓ Verify localStorage quota not exceeded

## 📱 Browser Compatibility

| Browser | Minimum Version | Status |
|---------|----------------|---------|
| Chrome | 90+ | ✅ Fully Supported |
| Firefox | 88+ | ✅ Fully Supported |
| Safari | 14+ | ✅ Fully Supported |
| Edge | 90+ | ✅ Fully Supported |

**Requirements:**
- LocalStorage support
- FileReader API
- Fetch API
- ES6+ JavaScript

## 🎯 Workflow

1. **Setup** → Enter API key, Sheet ID, date range
2. **Load LMRB** → Fetch data from Google Sheets
3. **Upload TC** → Select channel, upload PDF, convert
4. **Map Themes** → Create TC ↔ LMRB mappings
5. **Reconcile** → Run automatic matching
6. **Review** → Check matched/unmatched records
7. **Manual Match** → Match remaining records manually
8. **Export** → Download Excel/PDF reports

## 🚀 Deployment Options

### GitHub Pages (Recommended)
- ✅ Free hosting
- ✅ Automatic HTTPS
- ✅ CDN delivery
- ✅ Auto-deploy on push

### Alternative Static Hosts
- **Netlify** - Drop folder or connect GitHub
- **Vercel** - Import repository
- **Cloudflare Pages** - Connect GitHub
- **AWS S3** - Static website hosting
- **Firebase Hosting** - Google's CDN

## 📝 Next Steps

### Immediate
1. ✅ Deploy to GitHub Pages
2. ✅ Get Google API key
3. ✅ Test with Sirasa TV PDF

### When You Have Time
4. 📧 Provide sample PDFs from other channels
5. ⚡ Customize converters for each channel
6. 🎉 All channels working!

## 🤝 Contributing

This is a private project. For modifications:

1. Edit converter files for new channels
2. Test with sample PDFs
3. Commit and push
4. GitHub Pages auto-updates

## 📄 License

Internal use only - All rights reserved

---

## 🎉 Live Demo

**After deploying to GitHub Pages:**

```
https://shehank98.github.io/Reconciliation/static-app/
```

**Features:**
- ✨ Modern UI with tab navigation
- 📊 Real-time data preview
- 🔄 Automatic reconciliation
- 📥 Excel & PDF export
- ✋ Manual matching interface
- 📱 Mobile responsive

---

**Version:** 2.0 Static
**Type:** Pure Static Web Application
**Backend:** None Required
**Hosting:** GitHub Pages
**Cost:** $0/month

For detailed setup instructions, see **[STATIC_APP_GUIDE.md](STATIC_APP_GUIDE.md)**
