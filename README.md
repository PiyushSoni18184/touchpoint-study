# Touchpoint Study

The static game is served from `index.html`. The **Save Telemetry CSV File** action downloads the CSV and sends the session, including skipped trials, and the final survey responses to the linked Google Sheet.

## Enable anonymous Google Sheets submissions

1. Open the supplied spreadsheet and choose **Extensions → Apps Script**.
2. Replace the editor contents with the current `Code.gs` from this repository and save. It creates a `GameTelemetry` tab on first submission and adds the skip/confidence columns to an existing tab.
3. Open **Deploy → Manage deployments**, edit the existing web-app deployment, select **New version**, and deploy. Keep **Execute as** set to your account and **Who has access** set to **Anyone**.
4. The current `/exec` URL is already configured in `index.html`. If the deployment URL changes, update `GOOGLE_APPS_SCRIPT_URL` and publish the change to `main`.

The endpoint is intentionally anonymous: anyone with its URL can submit rows to the spreadsheet. Do not treat the URL as private or use this endpoint for sensitive data. The game uses a browser no-CORS request so it cannot read the Apps Script response; after a request is sent, verify rows in `GameTelemetry`. CSV download remains available regardless of endpoint availability.

## Publish with GitHub Pages

The repository is configured to publish from the `main` branch at the repository root. Once GitHub Pages finishes building, the game is available at `https://piyushsoni18184.github.io/touchpoint-study/`.
