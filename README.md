# Touchpoint Study

The static game is served from `index.html`. The **Save Telemetry CSV File** action always downloads the CSV. After you configure the Apps Script URL in `index.html`, it also sends the complete session and the final survey responses to the linked Google Sheet.

## Enable anonymous Google Sheets submissions

1. Open the supplied spreadsheet and choose **Extensions → Apps Script**.
2. Replace the editor contents with `Code.gs` from this repository and save. The script creates a `GameTelemetry` tab with fixed headers on the first successful submission.
3. Choose **Deploy → New deployment → Web app**. Set **Execute as** to your account and **Who has access** to **Anyone**, then deploy and authorize it.
4. Copy the deployed web app URL ending in `/exec`. In `index.html`, set `GOOGLE_APPS_SCRIPT_URL` to that URL and push the update to `main`.

The endpoint is intentionally anonymous: anyone with its URL can submit rows to the spreadsheet. Do not treat the URL as private or use this endpoint for sensitive data. The game uses a browser no-CORS request so it cannot read the Apps Script response; after a request is sent, verify rows in `GameTelemetry`. CSV download remains available regardless of endpoint configuration or availability.

## Publish with GitHub Pages

The repository is configured to publish from the `main` branch at the repository root. Once GitHub Pages finishes building, the game is available at `https://piyushsoni18184.github.io/touchpoint-study/`.
