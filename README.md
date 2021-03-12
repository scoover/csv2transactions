# CSV 2 Spreadsheet
Simple script to import &amp; remap CSV files from your Google Drive into a spreadsheet. 
This script works with Tiller Money spreadsheets but is not created, published or maintained by Tiller Money.

This is a simple import script with limited error checking. Import rules must be manually created for each import-CSV file type. There are no provisions in the current build to recognize various csv formats or host multiple import rules (though these improvements would not be difficult to build out).

This script performs the following actions:
- Requests filename of CSV file stored in user's Google Drive
- Loads and parses requested CSV file
- Maps imported data to Transactions sheet headers per rules defined in `mapping`, `metadata` and `invertAmountPolarity`
- Creates new rows at the top of the Transactions sheet
- Inserts new data into the Transactions sheet

## Setup & Configuration

This script runs within a Google Sheets spreadsheet as a bound script. 

## Installation
1. Open your Google Sheets spreadsheet.
2. Click on Tools/Script editor in the menu bar.
3. Open the `csv2spreadsheet.js` file in this project and copy it into the clipboard.
4. Paste the clipboard contents into the `Code.gs` file in the Script Editor.
5. Save the `Code.gs` file. (You will be prompted for a filename.)
6. Click back to the Google Sheets spreadsheet and reload the spreadsheet. 

Once the spreadsheet tab reloads (be patient), you will see a new option in the menu bar "Simple CSV". You have installed the script.

## Accept Permissions
When you first run the script by clicking Simple CSV/Import from Google Drive, you will be asked to accept the following Google permissions:
- See and download all your Google Drive files
- View and manage spreadsheets that this application has been installed in
- Connect to an external service
- Display and run third-party web content in prompts and sidebars inside Google applications
