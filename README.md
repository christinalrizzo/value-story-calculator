# value-story-calculator-slide

Pelago Value Story Generator â a Google Apps Script web app that calculates ROI projections and generates branded Google Slides from a template.

## What It Does

1. User enters channel partner name and prospect name
2. Enters eligible lives and selects pricing model
3. Checks/unchecks which substances (TUD, AUD, CUD, OUD) to include
4. Previews all calculated numbers before generating
5. Clicks **Generate Slide** â duplicates the template into their Google Drive with every value populated

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Apps Script backend â calculations, Slides API, slide generation |
| `Index.html` | Web app frontend â calculator form, number preview, generate button |
| `appsscript.json` | Apps Script manifest â OAuth scopes, runtime config |
| `.clasp.json` | clasp config â links this repo to your Apps Script project |

## Configuration

In `Code.gs`, update the `CONFIG` object:

- `TEMPLATE_ID` â Google Slides template file ID (the part between `/d/` and `/edit` in the URL)

### Pricing Models

**Hawaii:** Support $995, Manage $2,995, Treat $3,995

**Zanzibar:** Support $495, Manage $2,995, Treat $3,495

### Standard Constants

- SUD prevalence: 20%
- Engagement rate: 10%
- Tier split: Support 56%, Manage 37%, Treat 6%
- Avg savings per member: $11,289

## Setup & Deployment

### Prerequisites

- Node.js installed
- A Google account with access to the template slide and config sheet

### First-time Setup

```bash
# 1. Install clasp globally
npm install -g @google/clasp

# 2. Log in to your Google account
clasp login

# 3. Create a new Apps Script project (run from repo root)
clasp create --type webapp --title "Pelago Value Story Generator"
```

This creates a `.clasp.json` file with your project's `scriptId`. Commit it to the repo.

### Push Code to Apps Script

```bash
# Push all files to your Apps Script project
clasp push
```

This uploads `Code.gs`, `Index.html`, and `appsscript.json` to the linked Apps Script project.

### Deploy the Web App

```bash
# Create a new deployment
clasp deploy --description "v1.0"
```

Or deploy from the Apps Script editor:

1. `clasp open` (opens the project in your browser)
2. **Deploy** â **New deployment** â **Web app**
   - Execute as: **User accessing the web app**
   - Access: **Anyone within your org**
3. Authorize when prompted
4. Share the web app URL with the team

### Ongoing Workflow

```bash
# Edit files locally in your repo, then push
clasp push

# Update an existing deployment
clasp deploy --deploymentId <ID> --description "v1.1"

# Pull remote changes (if someone edits in the Apps Script editor)
clasp pull
```

### Required Permissions

The deploying account needs:
- **View** access to the template presentation
- Users need **View** access to the template (so they can copy it)

## How the Slide Population Works

The template slide has placeholder text (X's) that get replaced:

1. **Unique patterns** are replaced globally across the slide (e.g. `X.X : 1 NET ROI`, `$XXM in savings`)
2. **Ambiguous patterns** (e.g. `XXX`, `$XXXXXX`) are replaced per-shape by identifying each text box by its context (looks for keywords like "Support", "56%", "program spend", etc.)
3. **Substance tags** â shapes containing unselected substance text are deleted from the copy
