# Pelago Value Story Generator — Setup Guide

## What This Does

A hosted web app where anyone on your team can:

1. Select a channel partner (auto-sets pricing model)
2. Enter eligible lives and prospect name
3. Check/uncheck which substances (TUD, AUD, CUD, OUD) to show
4. Preview all the calculated numbers
5. Click **Generate Slide** — it copies the template into their Google Drive with every value filled in

## Setup Steps (10 minutes)

### 1. Create the Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **New project**
3. Rename it to "Pelago Value Story Generator"

### 2. Add the Code

**Code.gs** (replace the default contents):
- Open the `Code.gs` file provided and paste the entire contents into the script editor

**Index.html** (add a new file):
- Click the **+** next to "Files" → select **HTML**
- Name it `Index` (no extension)
- Paste the entire contents of the `Index.html` file provided

### 3. Update the Template ID

In `Code.gs`, find the `CONFIG` object at the top and update:

- `TEMPLATE_ID` — the Google Slides template with X placeholders. Use the file ID from its URL (the part between `/d/` and `/edit`)
- `SHEET_ID` — your Channel Partner Configurations sheet ID (already set to your current sheet)

**Important:** The Google account deploying this must have edit access to the template and the config sheet.

### 4. Deploy as Web App

1. Click **Deploy** → **New deployment**
2. Click the gear icon → select **Web app**
3. Set:
   - **Description**: "Value Story Generator v1"
   - **Execute as**: "User accessing the web app" ← this is key so slides go into each person's own Drive
   - **Who has access**: "Anyone within [your org]" (or specific people)
4. Click **Deploy**
5. **Authorize** when prompted (grant access to Drive, Slides, Sheets)
6. Copy the web app URL — that's what you share with Todd's team

### 5. Share the Template

Make sure the template presentation is shared with "Anyone in your org can view" (or at minimum, with everyone who will use the tool). They need at least view access to make a copy.

### 6. Share the Config Sheet

Same — the config sheet needs to be viewable by anyone using the tool so the partner list loads.

## How It Works

- **Partner dropdown** pulls live from your Google Sheet (only "Active" partners)
- **Preview Numbers** runs the calculation server-side and shows a visual preview
- **Generate Slide** calls `makeCopy()` on the template, then uses `replaceAllText()` via the Slides API to fill every placeholder
- Unselected substances get their tag shapes deleted from the copy
- The new presentation lands in the user's own Google Drive

## Adding New Partners

Just add a row to the [Channel Partner Configurations](https://docs.google.com/spreadsheets/d/1YnoKU-WW7J8dfbID-ee8_lG6g2h-Wdbk3xmJsOWewJk) sheet:

| Partner Name | Pricing Model | Status | Notes |
|---|---|---|---|
| NewPartner | Hawaii | Active | Any notes here |

It shows up in the dropdown immediately — no code changes needed.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Backend: calculations, Slides API, partner lookup |
| `Index.html` | Frontend: calculator form + slide preview |
