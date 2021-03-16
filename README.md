# CSV2Transactions
Simple script to remap imported CSV (comma separated values) files into the Transactions sheet of a Tiller-Money-compatible spreadsheet. 
*This script works with Tiller Money spreadsheets but is not created, published or maintained by Tiller Money.*


## What It Does
This script performs the following actions:
- Requests sheet name of your imported CSV data
- Maps imported data to Transactions sheet headers per the column mapping in the first row (see below)
- Creates new rows at the top of the Transactions sheet &amp; inserts the data

## Limitations
This is a very simple first-pass at the problem of importing CSV line-items into a Tiller-Money-compatible Transactions sheet.
- Extremely limited error checking
- Data remapping rules must be manually created 
- No provisions to automatically recognize various data formats 

## Setup

This script runs within a Google Sheets spreadsheet as a bound script. 

There are a few steps to get up and running. *Most of these steps are one-time steps.* Once you have installed the scripts and mapped the columns for your CSV data sources, using this workflow will be quick and easy.

### Installation
1. Open the `csv2transactions.js` file in this Github project and copy the contents into the clipboard.
2. Open your Google Sheets spreadsheet.
3. Click on `Tools/Script` editor in the menu bar.
4. Select all contents of the file and overwrite by pasting the clipboard contents into the `Code.gs` file in the Script Editor.
5. Save the `Code.gs` file. 
6. Open the `select-sheet.html` file above in this Github project and copy the contents into the clipboard.
7. In the Script Editor, create a new HTML file using the `File/New/HTML file` menu flow.
8. Enter the filename "select-sheet.html" for the new file.
9. Select all contents of the file and overwrite by pasting the clipboard contents in the Script Editor.
10. Save the `select-sheet.html` file.
11. In the Script Editor, navigate to and click the `Project Settings` gear in the left side bar. Check the box that says "Show `appsscript.json` manifest file in editor". 
12. Click back to the `Editor` (`<>`) pane in the left side bar in the Script Editor.
13. Open the `appsscript.json` fil  e in this Github project and copy the contents into the clipboard.
14. In the Script Editor, open `appsscript.json` using the `Using/Show manifest file` menu flow.
15. Select all contents of the file and overwrite by pasting the clipboard contents in the Script Editor.
16. Save the `appsscript.json` file.
17. Click back to the Google Sheets spreadsheet tab.
18. Reload the browser tab with the spreadsheet.

Once the spreadsheet tab reloads (be patient), you will see a new option in the menu bar "Simple CSV". You have installed the script.

### Accept Permissions
When you first run the script by clicking `Simple CSV/Import transactions`, you will be asked to accept the following Google account permissions:
- View and manage spreadsheets that this application has been installed in
- Display and run third-party web content in prompts and sidebars inside Google applications

## Usage

There are three steps to importing a CSV file using the `Simple CSV` workflow:
1. Import your CSV file into the spreadsheet.
2. Add your column mapping to the sheet header.
3. Run the Simple CSV script.

### Import Your CSV File into the Spreadsheet
1. Open your Tiller-Money-compatible spreadsheet. 
2. In the file menu, click `File/Import`.
3. Navigate to the Upload option in the upper right of the modal dialog.
4. Drag the CSV file from your data source to into the "Drag a file here" container.
5. Set `Import Location` to "Insert new sheet(s)".
6. Leave `Separator type` set to "Detect automatically" and `Convert text to numbers, dates and formulas` set to "Yes".
7. Click `Import Data`.

### Add Your Column Mapping to the Sheet Header
1. Select the newly-import sheet in your Tiller-Money-compatible spreadsheet. 
2. Insert a new row at the top of the sheet (above the header from the imported data).
3. Above each import-file header, type the (exact) name of the Transactions-sheet column where you'd like the data to land. Use Tiller-Money reserved column names like "Date", "Amount", "Description", "Account #", etc.

#### Save and Reuse Your Import Settings
Pro tip for power users! If you frequently depend on a CSV data source, you can reuse its column-mapping by keeping the imported-dataa sheet in your spreadsheet. Whenever you want to import new data, put the cursor on cell `A2` in the sheet (the cell containing the first header cell of the imported data) and, instead of importing with the "Insert new sheet(s)" setting, select "Replace data at selected cell". If you have derived or formula-driven columns to the right of your data. These will be preserved— a workflow that works great with `ARRAYFORMULA()` formulas (e.g. `={"Concatenated Fields";ARRAYFORMULA(IF(ISBLANK(A2:A),IFERROR(1/0),A2:A&"/"&B2:B))}` or `={"Data Source";ARRAYFORMULA(IF(ISBLANK(A2:A),IFERROR(1/0),"Imported with CSV2Transactions"))}`) driven from the header row 1.

#### Inverting Numbers
Sometimes numbers require polarity inversion to match the polarity of the Transactions sheet. Adding a minus symbol ('-') before a Tiller-Money reserved column names like "Amount" (e.g. "-Amount") will invert the value when the script is run.

#### Special Meta-column Headers
Several special Tiller-Money reserved column names will be generated automatically if they are present in your Transactions sheet.
- Month
- Week
- Date Added
- Metadata
You do not need to create mappings or custom content for them to populate.

#### Metadata Column
The (optional) Metadata column in the Transactions sheet is a great catchall for any data that you don't want to lose that doesn't fit neatly into one of the core columns. Multiple mappings to the Transactions sheet's "Metadata" column are allowed in the column mapping. All columns marked to map to the Metadata coulm will be captured in a [JSON object](https://en.wikipedia.org/wiki/JSON) the written in "stringified" form in Metadata column (if present). 

#### Adding Static or Formula-driven Values
Before running the script, you can new columns with formula-driven values or static text to your imported data sheet if you would like to include concatenated data or static text to your remapped data in the Transactions sheet.

#### Additional Notes
- The column mapping name & text case much match exactly to be recognized by the script (e.g "date" will not map to the Transaction sheet's "Date" column.)
- Avoid extra spaces (e.g. "_Date" will not map to "Date".).
- Other than the special case of "Metadata", if a column mapping (e.g. "Description") is used more than once, the first mapped column will be used.

### Run the Simple CSV Script
1. Run the script using the `Simple CSV/Import transactions` menu flow. (If it is your first time running the script, you will be asked to accept the required permissions and will need to run the script again.)
2. At the prompt, select the sheet name containing your imported data (with your manual header mapping row).
3. Click the `Submit` button.
4. Watch as the imported data is mapped into your Transactions sheet.
