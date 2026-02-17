# Static App Deployment Guide

## 🎉 Your Static App is Ready!

The `static-app/` folder contains a **100% client-side application** that runs entirely in the browser. Perfect for GitHub Pages!

## 🚀 Deploy to GitHub Pages (5 Minutes)

### Step 1: Push to GitHub

Your code is already in the repository. Just make sure it's pushed:

```bash
git push origin claude/remove-schedule-section-LrGLj
```

### Step 2: Enable GitHub Pages

1. Go to your repository on GitHub:
   ```
   https://github.com/Shehank98/Reconciliation
   ```

2. Click **Settings** (top menu)

3. Click **Pages** (left sidebar)

4. Under "Source":
   - Branch: Select `claude/remove-schedule-section-LrGLj`
   - Folder: Select `/ (root)`
   - Click **Save**

5. Wait 1-2 minutes for deployment

6. Your app will be live at:
   ```
   https://shehank98.github.io/Reconciliation/static-app/
   ```

### Step 3: Get Google API Key

1. Go to [Google Cloud Console](https://console.cloud.google.com/)

2. Create new project or select existing

3. Enable **Google Sheets API**:
   - Click "Enable APIs and Services"
   - Search "Google Sheets API"
   - Click Enable

4. Create API Key:
   - Go to Credentials
   - Click "Create Credentials" → "API Key"
   - Copy your key

5. **Restrict the API Key** (Important!):
   - Click on the key to edit
   - Under "Application restrictions":
     - Select "HTTP referrers"
     - Add: `https://shehank98.github.io/*`
   - Under "API restrictions":
     - Select "Restrict key"
     - Select only "Google Sheets API"
   - Save

6. **Make your Google Sheet public**:
   - Open your Google Sheet
   - Click "Share"
   - Change to "Anyone with the link can view"
   - Copy Sheet ID from URL

## 📝 First Time Usage

1. Open your deployed app:
   ```
   https://shehank98.github.io/Reconciliation/static-app/
   ```

2. Enter API Key and Sheet ID

3. Load LMRB data

4. Upload TC PDF (select channel first)

5. Add theme mappings

6. Run reconciliation

7. Export results!

## 🎨 Customizing Channel Converters

You have 5 channel converters ready:

### Working Converter
- ✅ **Sirasa TV** (`js/converters/sirasa.js`) - Fully implemented

### Templates (Need Customization)
- ⚠️ **Hiru TV** (`js/converters/hiru.js`)
- ⚠️ **Derana TV** (`js/converters/derana.js`)
- ⚠️ **ITN** (`js/converters/itv.js`)
- ⚠️ **TV1** (`js/converters/tv1.js`)

### How to Add Your Channel's Parser

1. **Get a sample PDF** from the channel

2. **Open the PDF** and note the format:
   ```
   Example: Is it like this?
   01/01/2024 | News | 10:30:15 | Coca Cola | 30

   Or like this?
   Date: 01/01/2024
   Program: News
   Time: 10:30:15
   Theme: Coca Cola
   Duration: 30 seconds
   ```

3. **Edit the converter** (e.g., `static-app/js/converters/hiru.js`):

```javascript
parsePage(text) {
    const records = [];
    const lines = text.split('\n');

    lines.forEach(line => {
        // YOUR PATTERN HERE
        // Example for pipe-separated:
        const pattern = /(\d{2}\/\d{2}\/\d{4})\s*\|\s*(.+?)\s*\|\s*(\d{2}:\d{2}:\d{2})\s*\|\s*(.+?)\s*\|\s*(\d+)/;

        // Example for tab-separated:
        // const pattern = /(\d{2}\/\d{2}\/\d{4})\t(.+?)\t(\d{2}:\d{2}:\d{2})\t(.+?)\t(\d+)/;

        const match = line.match(pattern);

        if (match) {
            const [_, date, program, time, theme, duration] = match;
            records.push({
                Date: date,
                Program: program.trim(),
                Time: time,
                Theme: theme.trim(),
                Duration: parseInt(duration)
            });
        }
    });

    return records;
}
```

4. **Test it**:
   - Upload your sample PDF
   - Check if data appears in preview
   - Adjust pattern if needed

5. **Commit and push**:
```bash
git add static-app/js/converters/hiru.js
git commit -m "Add Hiru TV PDF converter"
git push
```

6. GitHub Pages will auto-update in 1-2 minutes!

## 🔧 Updating Your Converter Later

When you need to update a converter with new PDF parsing code:

1. Edit the file locally: `static-app/js/converters/channelname.js`

2. Update the `parsePage()` method with your new pattern

3. Commit and push:
```bash
git add static-app/js/converters/channelname.js
git commit -m "Update channel converter for new format"
git push
```

4. Changes will be live on GitHub Pages in 1-2 minutes!

## 📊 How to Send Me Channel PDFs

To help you create parsers for each channel, send me:

1. **Sample PDF** from each channel
2. **Channel name** (Hiru, Derana, ITN, TV1)
3. **Format description** if you know it

I can create the regex patterns for you!

Example message:
```
"Here's a Hiru TV TC PDF. The format is:
Date | Program | Start Time | Ad Theme | Duration
Each field is separated by pipes (|)"
```

## ✅ Benefits of Static App

### vs Backend App

| Feature | Static App | Backend App |
|---------|-----------|--------------|
| **Cost** | $0 (GitHub Pages free) | $5-20/month |
| **Setup** | 5 minutes | 30+ minutes |
| **Updates** | Git push = instant | Redeploy servers |
| **Speed** | ⚡ Instant (CDN) | Depends on server |
| **Maintenance** | Zero | Regular updates |

### How It Works

```
┌─────────────┐
│   Browser   │
│             │
│  ┌───────┐  │
│  │ HTML  │  │  All processing
│  │ CSS   │  │  happens here!
│  │ JS    │  │
│  └───┬───┘  │
│      │      │
└──────┼──────┘
       │
       ├─────► Google Sheets API (read LMRB)
       ├─────► PDF.js (parse TC PDF)
       ├─────► localStorage (save data)
       └─────► XLSX (export Excel)
```

No server needed! Everything runs in your browser.

## 🐛 Troubleshooting

### Can't access GitHub Pages URL

**Problem**: 404 error when visiting the URL

**Solution**:
1. Check GitHub Pages is enabled in Settings
2. Verify branch and folder are correct
3. Wait 2-3 minutes for first deployment
4. Try: `https://shehank98.github.io/Reconciliation/static-app/index.html`

### "Failed to fetch data from Google Sheets"

**Problem**: Can't load LMRB data

**Solution**:
1. Verify API key is correct
2. Check Google Sheet is publicly viewable
3. Ensure API restrictions allow your GitHub Pages domain
4. Test API key with a simple request first

### PDF Converter Not Working

**Problem**: "PDF conversion failed" error

**Solution**:
1. Check channel is selected before uploading
2. Verify PDF is not password-protected
3. Ensure PDF has extractable text (not scanned image)
4. Check browser console for specific error
5. Try with a different PDF file

### Data Disappears After Refresh

**Problem**: Loaded data is gone after closing browser

**Solution**:
- Data is saved in localStorage automatically
- If in Private/Incognito mode, localStorage is cleared on close
- Use regular browser window
- Check localStorage is enabled in browser settings

## 📱 Mobile Support

The app works on mobile browsers! But for best experience:

- Use tablet or desktop for complex operations
- PDF upload may be slower on mobile
- Large PDFs may cause memory issues on mobile

## 🔐 Security

**Data Privacy**:
- All data stays in YOUR browser
- Nothing sent to external servers (except Google Sheets API)
- Clear browser data = clear all reconciliation data

**API Key**:
- Use restricted API keys (domain-specific)
- Never share your API key publicly
- Regenerate if compromised

## 🎯 Next Steps

1. ✅ Deploy to GitHub Pages (5 min)
2. ✅ Get Google API key (10 min)
3. ✅ Test with Sirasa TV PDF
4. 📧 Send me other channel PDFs for parsing
5. 🚀 Start using the app!

---

**Your App URL** (after deployment):
```
https://shehank98.github.io/Reconciliation/static-app/
```

**Questions?** Check the README in `static-app/` folder!
