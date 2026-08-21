markdown# Google Sheets UI Integration

This is a custom web app interface built with Google Apps Script that connects directly to a Google Sheet database. I built this project to experiment with custom UI/UX design layouts while keeping a live backend connection to process and store data.

## What it does
* **Custom Frontend:** A clean, responsive UI layout built to make data entry easier than typing straight into rows.
* **Sheet Automation:** Custom script logic that maps web form inputs directly into spreadsheet columns.
* **Media Handling:** Includes dedicated logic to process images and file attachments directly via Google Drive.

## Code Structure
* `Index.html` - The user interface design, forms, and custom styling.
* `Integration Guide.js` - Backend JavaScript managing the spreadsheet connections and form processing.
* `Photos.js` - Scripts built to handle image uploads and media assets.
* `appsscript.json` - Google Apps Script project configuration.
