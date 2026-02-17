# Deployment Guide

## Quick Deployment Options

### 🚀 Option 1: Vercel (Frontend) + Render (Backend) - RECOMMENDED

**Why**: Both have generous free tiers, easy setup, automatic deployments from GitHub.

#### Step 1: Deploy Backend to Render

1. **Create Render Account**: Go to [render.com](https://render.com) and sign up with GitHub

2. **Create New Web Service**:
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Select the `Reconciliation` repo

3. **Configure Service**:
   ```
   Name: reconciliation-api
   Region: Choose closest to you
   Branch: claude/remove-schedule-section-LrGLj
   Root Directory: backend
   Runtime: Python 3
   Build Command: pip install -r requirements.txt && pip install gunicorn
   Start Command: gunicorn app:app
   ```

4. **Add Environment Variables**:
   - Click "Environment"
   - Add: `GOOGLE_CREDENTIALS_PATH` = (leave blank for now, add JSON content later)
   - Or upload service account JSON as a secret file

5. **Deploy**: Click "Create Web Service"
   - Wait 2-3 minutes for build
   - Copy your backend URL: `https://reconciliation-api.onrender.com`

#### Step 2: Deploy Frontend to Vercel

1. **Create Vercel Account**: Go to [vercel.com](https://vercel.com) and sign up with GitHub

2. **Import Repository**:
   - Click "Add New" → "Project"
   - Import your `Reconciliation` repo
   - Framework Preset: **Vite**

3. **Configure Build**:
   ```
   Root Directory: frontend
   Build Command: npm run build
   Output Directory: dist
   Install Command: npm install
   ```

4. **Add Environment Variable**:
   - Click "Environment Variables"
   - Add: `VITE_API_URL` = `https://reconciliation-api.onrender.com/api`

5. **Deploy**: Click "Deploy"
   - Wait 1-2 minutes
   - Your frontend URL: `https://reconciliation-xxx.vercel.app`

#### Step 3: Update Frontend API URL

The frontend needs to know where the backend is. Update the API URL:

```javascript
// frontend/src/services/api.js
const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5000/api'
```

This is already in your code! Just set the environment variable in Vercel.

#### Step 4: Configure CORS

Your backend needs to allow requests from Vercel. Update backend/app.py:

```python
# Update CORS configuration
CORS(app, resources={
    r"/api/*": {
        "origins": [
            "http://localhost:3000",
            "https://reconciliation-xxx.vercel.app",  # Your Vercel URL
            "https://*.vercel.app"  # All Vercel preview deployments
        ]
    }
})
```

---

### 🚀 Option 2: Railway (Full Stack) - EASIEST

**Why**: Deploy both frontend and backend from one place, $5/month after free trial.

1. **Create Account**: Go to [railway.app](https://railway.app)

2. **Deploy Backend**:
   - Click "New Project" → "Deploy from GitHub repo"
   - Select your repo
   - Railway auto-detects Python
   - Add environment variables
   - Get backend URL

3. **Deploy Frontend**:
   - Click "New" → "Service" in same project
   - Add frontend service
   - Set root directory to `frontend`
   - Railway auto-detects Vite
   - Get frontend URL

4. **Link Services**:
   - Frontend gets backend URL automatically
   - Both services in one project

---

### 🚀 Option 3: Netlify (Frontend) + PythonAnywhere (Backend)

**Backend on PythonAnywhere** (FREE tier available):

1. Go to [pythonanywhere.com](https://www.pythonanywhere.com)
2. Create free account
3. Upload your backend code
4. Set up Flask app
5. Get URL: `yourusername.pythonanywhere.com`

**Frontend on Netlify**:

1. Go to [netlify.com](https://www.netlify.com)
2. Drag & drop your `frontend/dist` folder after building locally
3. Or connect GitHub repo
4. Set environment variable: `VITE_API_URL`

---

## 📦 For Production: Update Backend Requirements

Add gunicorn to requirements.txt:

```bash
cd backend
echo "gunicorn==21.2.0" >> requirements.txt
```

---

## 🔐 Google Sheets Credentials in Production

### Option 1: Environment Variable (JSON as string)

1. Copy entire JSON content from your service account file
2. In hosting platform (Render/Vercel), create environment variable:
   - Name: `GOOGLE_CREDENTIALS_JSON`
   - Value: Paste entire JSON

3. Update backend code to read from env:

```python
import os
import json
from google.oauth2 import service_account

# In google_sheets.py
credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
if credentials_json:
    credentials_info = json.loads(credentials_json)
    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=SCOPES
    )
```

### Option 2: Secret Files (Render/Railway)

Upload service account JSON as a secret file in dashboard.

---

## 🧪 Testing Deployment

1. **Backend Health Check**:
   ```
   https://your-backend.onrender.com/api/health
   ```
   Should return: `{"status": "healthy"}`

2. **Frontend Access**:
   ```
   https://your-frontend.vercel.app
   ```
   Should load the dashboard

3. **Full Test**:
   - Connect to Google Sheet
   - Upload TC PDF
   - Run reconciliation

---

## 💰 Cost Comparison

| Service | Free Tier | Paid |
|---------|-----------|------|
| **Vercel** | Unlimited hobby projects | $20/mo Pro |
| **Render** | 750 hours/month free | $7/mo |
| **Railway** | $5 trial credit | $5/mo usage-based |
| **Netlify** | 100GB bandwidth/month | $19/mo |
| **PythonAnywhere** | 1 app free | $5/mo |

**Recommended FREE combo**: Vercel (Frontend) + Render (Backend) = $0/month

---

## 🚀 Quick Deploy Commands

### Deploy to Vercel (from CLI)

```bash
# Install Vercel CLI
npm i -g vercel

# Deploy frontend
cd frontend
vercel

# Follow prompts
```

### Deploy to Render

Just push to GitHub - Render auto-deploys!

---

## ⚠️ Important Notes

1. **Free Tier Limitations**:
   - Render free tier: App sleeps after 15 min of inactivity
   - First request after sleep takes ~30 seconds to wake up
   - Upgrade to paid ($7/mo) for always-on

2. **CORS Configuration**:
   - Must update CORS settings in backend to allow frontend domain

3. **Environment Variables**:
   - Set `VITE_API_URL` in frontend
   - Set `GOOGLE_CREDENTIALS_JSON` in backend

4. **Build Time**:
   - First deploy: 5-10 minutes
   - Subsequent deploys: 2-3 minutes
   - Auto-deploy on git push

---

## 🆘 Troubleshooting

### Backend returns 502

- Check logs in hosting platform
- Verify gunicorn is installed
- Check Python version matches

### Frontend can't connect to backend

- Check CORS configuration
- Verify API_URL environment variable
- Check backend is running (health check endpoint)

### Google Sheets not working

- Verify credentials are set correctly
- Check service account has access to sheet
- View backend logs for auth errors

---

## Need Help?

1. Check hosting platform docs:
   - [Vercel Docs](https://vercel.com/docs)
   - [Render Docs](https://render.com/docs)
   - [Railway Docs](https://docs.railway.app)

2. Common issues in this guide
3. Check backend/frontend logs in platform dashboard

---

**Recommended Path**: Start with Vercel + Render free tiers, upgrade if needed!
