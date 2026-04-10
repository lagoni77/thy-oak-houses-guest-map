# 🌍 Thy Oak Houses — Guest World Map

A single-page Airbnb guest map that automatically pulls reservation data from Gmail via Google Apps Script and displays it on a Leaflet.js world map.

## Files

| File | Purpose |
|------|---------|
| `index.html` | The frontend map (deploy this to Vercel) |
| `Code.gs` | Google Apps Script backend (paste into Google Sheets) |

---

## Setup Guide

### Step 1 — Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Open **Extensions → Apps Script**
3. Paste the contents of `Code.gs` into the editor
4. Run `setupSheet()` once manually (creates the header row)

### Step 2 — Enable Gmail API

In the Apps Script sidebar: **Services (+) → Gmail API → Add**

### Step 3 — Deploy as Web App

1. **Deploy → New Deployment**
2. Type: **Web App**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Copy the deployment URL — you'll need it in Step 5

### Step 4 — Set the Trigger

1. Click the **Clock icon** (Triggers) in the left sidebar
2. Add trigger: `syncAirbnbEmails` | Time-driven | **Every hour**

### Step 5 — Connect to the Frontend

Open `index.html` and replace:
```js
const API_URL = 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
```
with your Web App URL from Step 3.

### Step 6 — Deploy to Vercel

Push this repo to GitHub, then import it in [Vercel](https://vercel.com).  
Set the **root directory** to `guest map` (or deploy from the repo root).

---

## Manual Testing

In Apps Script, run `testAddSampleGuest()` to add a test entry and verify the map works before emails arrive.

## How the Sync Works

```
Gmail (Airbnb emails)
        ↓  syncAirbnbEmails() — runs every hour
Google Sheet (GuestData)
        ↓  doGet() Web App API
Frontend (index.html on Vercel)
        ↓  Leaflet.js map
```

- **Duplicate prevention**: checked via unique Gmail Message ID
- **Geocoding**: cached per location — never re-geocodes the same city
- **Incremental sync**: `PropertiesService` stores last scan date
