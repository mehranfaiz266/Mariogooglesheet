# Mariogooglesheet

This Google Apps Script manages EMRG leads in a Google Spreadsheet and syncs them with GoHighLevel (LeadConnector).

## Setup

1. Open the Apps Script editor for this project.
2. In **Project Properties → Script Properties**, add:
   - `GHL_PRIVATE_TOKEN` – your API token.
   - `GHL_LOCATION_ID` – your location ID.
3. Attach the project to a spreadsheet and run **⚖️ EMRG Tools → Initialize Leads Sheet** to create the required sheets.
4. Authorize the script when prompted.

## Usage

- **📊 Show Dashboard** – view summary metrics in a sidebar.
- **➕ Add New Lead** – open a form to create a lead and log it to the sheet.
- **🔄 Re-sync All Rows** – update GoHighLevel with changes from the sheet.
- **📈 Build Dashboard Sheet** – generate a `Dashboard` sheet with charts.
- **🛠️ Initialize Leads Sheet** – rebuild the sheet structure.

Editing a row marks it as `Pending` so it will be synchronized from the menu or by the optional 15‑minute trigger.
