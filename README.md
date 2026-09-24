# Foxy Responsive Admin Dashboard

Google Apps Script dashboard with Google Sheets as the existing data store.

## Architecture

- Apps Script files in this repository remain the backend and deployment source.
- Google Sheets remains the data store; this repository contains no sheet data export.
- Keep this repository private because the source includes authentication and data-access logic.

## Local development

Install the Apps Script CLI and authenticate with the owning Google account:

```powershell
npm install --global @google/clasp
clasp login
```

Pull or push Apps Script source from the repository folder with `clasp pull` and `clasp push`.
