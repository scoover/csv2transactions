////////////////////////////////////////////////////
// 
// CSV IMPORT RULES
// ----------------
// Change this section to work with your CSV import file.
// (Example below set up for Target RedCard credit card.)
//
// `mapping` object defines mapping from Tiller spreadsheet header to csv-import header
// the first value is the Tiller Transactions sheet header, the second value is the csv-import header
var mapping = {
  'Date': 'Trans Date',
  'Category': null,
  'Amount': 'Amount',
  'Transaction ID': null,
  'Description': 'Merchant Name',
  'Full Description': 'Merchant Name',
  'Category Hint': null,
  'Account': null,
  'Account #': 'Originating Account Last 4',
  'Account ID': null,
  'Institution': null,
};
// `metadata` array defines imported headers that will be added to JSON object in Metadata column if present
var metadata = [ 'Posting Date', 'Type', 'Category', 'Merchant City', 'Merchant State', 'Description', 'Transaction Type' ];
// set `invertAmountPolarity` to true to invert polarity of Amount column value
var invertAmountPolarity = true;
// 
////////////////////////////////////////////////////

// onOpen: setup menu bar navigation
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Simple CSV')
    .addItem('Import from Google Drive', 'importCsv')
    .addToUi();
};

// importCsv: import csv file (from Google Drive) into the spreadsheet based on `mapping`, `metadata` and `invertAmountPolarity`
function importCsv() {
  var ui = SpreadsheetApp.getUi();

  // load the moment library for date/time handling (used for Month + Week columns and sheet name timestamp)
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js').getContentText());

  // fetch the Transactions sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transactions = ss.getSheetByName('Transactions');

  // report an error if the Transactions sheet is not present
  if(!transactions) {
    return ui.alert('Missing Sheet', 'A Transactions sheet must be present to run this workflow.', ui.ButtonSet.OK);
  }
  
  var t = transactions.getRange(1, 1, 2, transactions.getMaxColumns()).getValues();

  // prompt user for Google Drive filename (csv file for import)
  // csv file must be first uploaded to the users Google Drive
  var response = ui.prompt('CSV Filename', 'Google Drive filename:', ui.ButtonSet.OK_CANCEL);
  // cancel execution if user presses cancel or the filename provided is blank
  if((response.getSelectedButton() == ui.Button.CANCEL) || !response.getResponseText().length) { return; }
  // scan the user's Google Drive for the file and ACCEPT THE FIRST MATCH!
  if(!DriveApp.getFilesByName(response.getResponseText()).hasNext()) {
    return ui.alert('Not Found', 'No file with name \''+response.getResponseText()+'\' found in your Google Drive.', ui.ButtonSet.OK);
  }

  // load and parse the first matching file with the provided filename
  var file = DriveApp.getFilesByName(response.getResponseText()).next();
  var data = Utilities.parseCsv(file.getBlob().getDataAsString());

  // create an output array for imported data
  var output = [ ];

  // iterate through the imported rows in `data` (skip header row)
  for(var row = 1; row < data.length; row++) {
    var newRow = [ ];
    var value = null;
    var hasValues = false;

    // iterate through the columns in the order of the Transactions sheet header
    for(var col = 0; col < t[0].length; col++) {
      // get the header name from the column at index `col`
      var columnName = t[0][col];

      // skip over columns where the column header is not present in the `mapping` object
      if(!mapping.hasOwnProperty(columnName)) { 
        newRow.push(null);
        continue;
      }

      // set `mapValue` to the imported data column header name with the relevant data
      var mapValue = mapping[columnName];
      
      // skip columns that are mapped to data that do not exist in the imported data
      // if a match is not found in the imported data headers, use the literal string provided in the mapping object
      if((typeof mapValue !== 'string') || (data[0].indexOf(mapValue) == -1)) { 
        newRow.push(mapValue);
        continue;
      }

      // fetch the field value based on the column remapping
      value = data[row][data[0].indexOf(mapValue)];

      // if the cell is an empty string, push null into the newRow array and don't set `hasValues` (to true) for the new row
      if((typeof value === 'string') && !value.length) { 
        newRow.push(null);
        continue;
      }
      // when processing the 'Amount' column, convert the parsed string "value" to a number
      else if(columnName == 'Amount') {
        // if the first character is a parenthesis, assume the value is wrapped in parentheses to indicate a negative value
        var parenthesis = value[0] == '(';
        // use a regex to strip the string "value" to numbers, decimals and negative signs then convert to number
        // if the `invertAmountPolarity` is true, invert the polarity of the Amount column value
        value = Number(value.replace(/[^0-9.-]+/g,'')) * (parenthesis? -1:1) * (invertAmountPolarity? -1:1);
      }

      // if we made it this far, the column's value is non blank, so mark `hasValues` as true (the row will be written)
      hasValues = true;
      // push the new (column) value to the new row array
      newRow.push(value);
    }

    // if `hasValues` was never set to true, the row has no non-blank values, so skip writing it to the output array
    if(!hasValues) { continue; }
    
    // now that all the values have been ingested from the import row, add values for special columns: Month, Week, Date Added and Metadata
    // iterate over the columns in `newRow`
    for(var col = 0; col < newRow.length; col++) {
      // set `columnName` to the Transactions column header
      var columnName = t[0][col];

      // if the column is "Month" or "Week" (and the Transactions sheet includes a "Date" column)
      if(((columnName == 'Month') || (columnName == 'Week')) && (t[0].indexOf('Date') != -1)) {
        // start by fetching the date value from `newRow`'s Date column
        var dateValue = new Date(newRow[t[0].indexOf('Date')]);

        // skip if 'Date' column value is not a valid date (not possible to generate a "Month" or "Week")
        if(!(dateValue instanceof Date) || isNaN(dateValue.valueOf())) { continue; }

        // create the date value for the "Month" column
        if(columnName == 'Month') {
          newRow[col] = moment(dateValue).startOf('month').format('M/D/YYYY');
        }
        // create the date value for the "Week" column
        else if (columnName == 'Week') {
          newRow[col] = moment(dateValue).startOf('week').format('M/D/YYYY');
        }
      }
      // if the column is "Date Added", insert today's date
      else if(columnName == 'Date Added') {
        newRow[col] = moment().format('M/D/YYYY');
      }
      // if the column is "Metadata", create a stringified JSON object with the column headers in the `metadata` array
      else if(columnName == 'Metadata') {
        var m = { };

        metadata.forEach(function (header) {      
          if(data[0].indexOf(header) != -1) { m[header] = data[row][data[0].indexOf(header)]; }
        });

        newRow[col] = JSON.stringify(m);
      }
    }

    // push `newRow` into the output array
    output.push(newRow);
  }

  // write the `output` array to the destination sheet
  transactions.insertRowsBefore(2, output.length);
  transactions.getRange(2, 1, output.length, output[0].length).setValues(output);
}