# CSV2Transactions
Simple script to import &amp; remap CSV (comma separated values) files from your Google Drive into a spreadsheet. 
*This script works with Tiller Money spreadsheets but is not created, published or maintained by Tiller Money.*


## What It Does
This script performs the following actions:
- Requests filename of CSV file stored in user's Google Drive
- Loads and parses requested CSV file
- Maps imported data to Transactions sheet headers per rules defined in `mapping`, `metadata` and `invertAmountPolarity`
- Creates new rows at the top of the Transactions sheet
- Inserts new data into the Transactions sheet

## Limitations
This is a very simple first-pass at the problem of importing CSV line-items into a Tiller-Money-compatible Transactions sheet.
- Extremely limited error checking
- Import rules must be manually created for each import-CSV file type
- No provisions to recognize various csv formats 
- No provisions to for multiple import rules

## Setup

This script runs within a Google Sheets spreadsheet as a bound script. 

### Installation
1. Open your Google Sheets spreadsheet.
2. Click on Tools/Script editor in the menu bar.
3. Open the `csv2transactions.js` file in this project and copy it into the clipboard.
4. Paste the clipboard contents into the `Code.gs` file in the Script Editor.
5. Save the `Code.gs` file. (You will be prompted for a filename.)
6. Click back to the Google Sheets spreadsheet and reload the spreadsheet. 

Once the spreadsheet tab reloads (be patient), you will see a new option in the menu bar "Simple CSV". You have installed the script.

### Accept Permissions
When you first run the script by clicking `Simple CSV/Import from Google Drive`, you will be asked to accept the following Google permissions:
- See and download all your Google Drive files
- View and manage spreadsheets that this application has been installed in
- Connect to an external service
- Display and run third-party web content in prompts and sidebars inside Google applications

## Usage

### Customizing Column Mapping
Three global variables at the top of the `csv2transactions.js` file must be customized to map your CSV file to the proper columns in your Transactions sheet. Make sure to save your changes before running the script using the `Simple CSV/Import from Google Drive` menu flow.

#### Mapping
`mapping` is a javascript object with key values pair. The key is a column header exactly as it appears in your Transactions sheet. The value is a column header exactly as it appears in your import file. You can use `null` or delete the keys for columns to be left blank. 

```
var mapping = {
  'Date': 'Trans Date',
  'Amount': 'Amount',
  'Description': 'Merchant Name',
  ...
};
```

When a key maps to a column name that is not found in the import file, the value text is inserted into the column. For example, a key such as the one below would insert the text "Imported using csv2transactions" in the `Note` column of all imported rows.
```
  'Note': 'Imported using csv2transactions',
```
#### Metadata
`metadata` is an array that defines which imported headers will be mapped to JSON object in `Metadata` column (if present). The `Metadata` column is a great catchall for any data that doesn't fit neatly into one of the core columns but which you don't want to lose.

```
var metadata = [ 'Posting Date', 'Type', 'Transaction Type' ];
```

Metadata is written to your spreadsheet in as a stringified JSON object like this:
```
{"Posting Date":"1/1/2021","Type":"Debit","Transaction Type":"Regular Sales Draft"}
```

#### InvertAmountPolarity
`invertAmountPolarity` allows you to reverse the polarity of your Amount column value. Set this to true to reverse the polarity of the values in the import file.

### Importing CSV Files
1. Customize your column mapping (per the instructions above).
2. Download a CSV file from your data source. Save it to your local computer.
3. Rename the CSV file to a unique name.
4. Upload the CSV file to your Google Drive.
5. In your spreadsheet, run the script using the `Simple CSV/Import from Google Drive` menu flow.
6. Provide the unique filename for the uploaded CSV at the prompt.
7. Watch at the CSV file is mapped into your Transactions sheet.

Note that the script will use the first file with the provided filename that it finds. The results may be unpredictable if there are multiple files with the same name.
