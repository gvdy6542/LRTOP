# LRTOP Google Apps Script

This repository contains a Google Apps Script project that provides a barcode based interface for store inventory and PPR (product receipt) processing. The script uses multiple HTML pages for the user interface and `Code.js` for the server side logic.

## Purpose
The application helps store staff scan barcodes, record quantities, generate PPR spreadsheets and view waste reports. Files are stored on Google Drive and manipulated through Apps Script.

## Deploying with clasp
1. Install [clasp](https://github.com/google/clasp) and authenticate.
2. Ensure `.clasp.json` contains the correct script ID:
```json
{
  "scriptId": "1JapSNYzL9MSIr6POHrL61qSbJ2Jo8uFeV6075STOGorI6rtD1q5LS2-C",
  "rootDir": ""
}
```
3. Push the code:
```bash
clasp push --force
```
A GitHub Actions workflow (`.github/workflows/deploy.yml`) performs these steps automatically on every push to `main`.

## Required OAuth scopes
The project requests the following scopes as defined in `appsscript.json`:
```
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/drive.readonly
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/script.scriptapp
https://www.googleapis.com/auth/script.external_request
https://www.googleapis.com/auth/userinfo.email
https://www.googleapis.com/auth/script.send_mail
```
These allow access to Drive and Spreadsheets, external requests and sending mail.

## Pages and key functions
- **index.html** – main barcode scanning UI with modals for PPR and waste reports.
- **interface.html** – embeds a Drive folder for quick access.
- **reference.html** – utilities for processing Drive files.

Notable functions in `Code.js` include:
- `doGet()` which serves `index.html`.
- `listRevisions()` and `getRevisionData()` for managing revision spreadsheets.
- `processBarcode()` to look up products.
- `savePPRData()`, `getPprData()` and `updatePprData()` for PPR creation and editing.
- `getWasteReport()` to generate monthly waste summaries.
