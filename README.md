# MC Charkhi Dadri - Work Management System (DadriWorks)

A Progressive Web App (PWA) designed for Municipal Council Charkhi Dadri to track public works projects, manage Performance Bank Guarantees (PBG), generate official Word (.docx) documents, and export reports in PDF and Excel format.

---

## Architecture Overview

- **Frontend:** Single-Page Application (HTML5, Vanilla JS, Bootstrap 5, FontAwesome) hosted statically on **GitHub Pages**.
- **Backend:** Google Apps Script Web App interacting with Google Sheets (Master Sheet & BG Register) and Google Drive (for PBG file storage).
- **Offline / PWA:** Service Worker caching core assets for fast loads and offline access.

---

## Directory Structure

```text
Dashboard/
├── index.html            # Main dashboard interface
├── style.css             # Application stylesheet
├── config.js             # Global configuration (Apps Script URL, version)
├── service-worker.js     # PWA offline cache worker
├── manifest.json         # Web app manifest for mobile/desktop install
├── document-templates.js # Word (.docx) templates (agreements, certificates)
├── metadata.json         # PWA metadata & permissions
├── icons/                # App icons and vector assets
├── server.js             # Optional local development server & proxy
├── package.json          # Node config & test scripts
└── test/                 # Automated test suite
    └── calculations.test.js
```

---

## Deployment on GitHub Pages

1. Push this repository to your GitHub account (`main` branch).
2. Go to **Settings > Pages** in your repository.
3. Under **Build and deployment > Source**, select **Deploy from a branch**.
4. Select `main` branch and `/ (root)` folder, then click **Save**.
5. Your app will be live at `https://<your-username>.github.io/<repo-name>/`.

---

## Google Apps Script Setup & Updates

The backend script files are kept in the `app script/` folder:
- `app script/Code.js`: Core Google Apps Script backend.
- `app script/appsscript.json`: OAuth scopes and manifest.

### Updating Apps Script:
1. Open your master Google Sheet.
2. Click **Extensions > Apps Script**.
3. Copy the contents of `app script/Code.js` and paste them into the code editor.
4. Click **Deploy > Manage deployments > Edit > New version**, then click **Deploy**.
5. If the deployment URL changes, update `SCRIPT_URL` in `config.js`.

---

## Configuration (`config.js`)

Central configuration is stored in `config.js`:

```javascript
window.APP_CONFIG = {
  SCRIPT_URL: 'https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec',
  VERSION: '20261026_v30'
};
```

Whenever you update `config.js` or styles, increment the `VERSION` string to notify the service worker to refresh client caches.

---

## Local Development & Testing

You can run the app locally using Node.js:

```bash
# Start local server on http://localhost:3000
npm start

# Run automated tests
npm test

# Run syntax checks
npm run lint
```
