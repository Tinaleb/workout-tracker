# Workout Tracker

A mobile-first workout tracking app. Log sessions, sets, reps, weight, and PRs. Export to PDF or Excel/CSV. All data stays locally in the browser — no accounts, no servers.

## Quick Start

```bash
# Install dependencies
npm install

# Run locally
npm run dev
```

Open [http://localhost:5173](http://localhost:5173) in your browser.

## Deploy to GitHub Pages

### First-time setup

1. Create a new repository on GitHub
2. Push this project to it:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git push -u origin main
   ```

3. Open `package.json` and update the `homepage` field:
   ```json
   "homepage": "https://YOUR_USERNAME.github.io/YOUR_REPO_NAME"
   ```

4. Deploy:
   ```bash
   npm run deploy
   ```

5. In your GitHub repo → **Settings → Pages**, set the source branch to `gh-pages`.

Your app will be live at `https://YOUR_USERNAME.github.io/YOUR_REPO_NAME`

### Subsequent deployments (after making changes)

```bash
npm run deploy
```

## Project Structure

```
workout-tracker/
├── index.html          # HTML entry point (viewport, fonts, base styles)
├── vite.config.js      # Vite build config
├── package.json        # Dependencies and scripts
├── .gitignore
└── src/
    ├── main.jsx        # React entry point
    └── App.jsx         # Main application (all components and logic)
```

## Features

- 📋 Multiple workout sessions per export
- 🏋️ Log exercise type, sets, reps, weight (lbs/kg), duration, notes
- ⏱ Built-in rest timer with audio + visual alert
- 🏆 PR (Personal Record) flagging per exercise
- 📊 Live session summary banner
- 📂 Upload previous CSV/Excel for side-by-side progress comparison
- 💾 Auto-save to browser localStorage — data persists between visits
- 🔒 Privacy consent modal on first visit
- ⚠️ Leave-page warning if data hasn't been exported
- ⬇️ Export to PDF (print dialog) or CSV (opens in Excel / Google Sheets)
- 📱 Fully responsive — mobile, tablet, and desktop
