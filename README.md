# CSV-2-Spreadsheet
Simple script to import &amp; remap CSV files from your Google Drive into a spreadsheet. 
This script works with Tiller Money but is not created, published or maintained by Tiller Money.

This is a simple import script with limited error checking. Import rules must be manually created for each import-CSV file type. There are no provisions in the current build to recognize various csv formats or host multiple import rules (though these improvements would not be difficult to build out).

This script performs the following actions:
- Requests filename of CSV file stored in user's Google Drive
- Loads and parses requested CSV file
- Maps imported data to Transactions sheet headers per rules defined in `mapping`, `metadata` and `invertAmountPolarity`
- Creates new rows at the top of the Transactions sheet
- Inserts new data into the Transactions sheet

## Setup

This script runs within a Google Sheets spreadsheet as a bound script. 

## Permissions
This script requests the following Google permissions:
- See and download all your Google Drive files
- View and manage spreadsheets that this application has been installed in
- Connect to an external service
- Display and run third-party web content in prompts and sidebars inside Google applications
