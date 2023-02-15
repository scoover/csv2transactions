// onOpen: setup menu bar navigation
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Simple Import')
    .addItem('Import transactions', 'selectSourceSheet')
    .addToUi();
};

// selectSourceSheet: prompt user to select sheet containing imported data
function selectSourceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = [ ];

  // get the names of all sheets in the spreadsheet 
  // exclude sheets with known names (i.e. not likely to be the imported data)
  ss.getSheets().forEach(function(sheet) {      
    if(['Transactions', 'Balance History', 'Categories', 'Monthly Budget', 'Yearly Budget', 'Accounts', 'Balances', 'Insights']
      .indexOf(sheet.getName()) == -1)
    allSheets.push(sheet.getName());
  });

  // create a modal prompt for the user to select the sheet name from a dropdown
  var form = HtmlService.createTemplateFromFile('select-sheet');
  form.sheets = JSON.stringify(allSheets.sort());
  SpreadsheetApp.getUi().showModalDialog(
    form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(150).setWidth(400),
    'Select Sheet');
}

// importTransactions: remap the transactions from `importSheetName` to the Transactions sheet
function importTransactions(importSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  // get the Transactions sheet
  var transactions = ss.getSheetByName('Transactions');
  // get the user-selected import data sheet from selectSourceSheet()
  var importedData = ss.getSheetByName(importSheetName);
  var msg = "";
  // report an error if the Transactions sheet is not present
  if(!transactions) {
    msg = "Transactions sheet not found";
    console.error(msg);
    return ui.alert('Missing Sheet', 'A Transactions sheet must be present to run this workflow.', ui.ButtonSet.OK);
  }
  // report an error if the imported data sheet is no longer present
  else if(!importedData) {
    msg = `Selected sheet for importing [${importSheetName}] not found`;
    console.error(msg);
    return ui.alert('Missing Sheet', 'The sheet you selected for import could no longer be found.', ui.ButtonSet.OK);
  }
   // check timezone settings in script and sheet and report an error if they don't match
  var validTimezone = checkTimezone(ss)  
  // function returns an array with a boolean value and a message
  if (!validTimezone[0]) {
    //log issue and prompt user to stop import or continue anyway
    console.error(validTimezone[1])
    var prompt = " Stop import?"
    var response = ui.alert('Timezone Error - stop import?', validTimezone[1] +prompt, ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      msg = "User stopped import due to timezone descrepancy"
      console.log(msg)
      return ui.alert('Timezone Discrepancy', 'Import cancelled', ui.ButtonSet.OK);
    }
    if (response == ui.Button.NO) {
      msg = "User selected to continue import despite timezone descrepancy"
      console.warn(msg)
      ui.alert('Timezone Discrepancy Ignored', 'Continuing Import', ui.ButtonSet.OK);
    }
  }

  // get the (value) contents of the Transactions sheet - need headers to map data
  var t = transactions.getRange(1, 1, 1, transactions.getMaxColumns()).getValues();
  // get the (value) contents of the import data sheet
  var data = importedData.getRange(1, 1, importedData.getMaxRows(), importedData.getMaxColumns()).getValues();

  // create an output array for imported data
  var output = [ ];

  // iterate through the imported rows in `data` (skip 2 rows - mapping row + import header row)
  for(var row = 2; row < data.length; row++) {
    var newRow = [ ];
    var value = null;

    // iterate through the columns in the order of the Transactions sheet header
    for(var col = 0; col < t[0].length; col++) {
      // get the header name from the column at index `col`
      var columnName = t[0][col];

      // set `mapColumn` to the (first) index of the column in `importedData` mapped to `columnName`
      var mapColumn = data[0].indexOf(columnName);
      // if no `mapColumn` is found, try with a minus prefix (indicating a mapping with inverted polarity)
      if(mapColumn == -1) { mapColumn = data[0].indexOf('-'+columnName) }

      // skip over columns where the column header is not mapped in `importedData` or where column name is reserved
      if(mapColumn == -1 || (['Month','Week','Date Added','Metadata'].indexOf(columnName) != -1)) { 
        newRow.push(null);
        continue;
      }

      // fetch the field value based on the column remapping
      value = data[row][mapColumn];

      // if the cell is empty, push null into the newRow array
      if((typeof value === 'string') && !value.length) { 
        newRow.push(null);
        continue;
      }
      // invert polarity of columns mapped with header prefix '-'
      else if((typeof value === 'number') && (data[0][mapColumn][0] == '-')) { value *= -1; }

      // push the new (column) value to the new row array
      newRow.push(value);
    }

    // if all columns contain blank values, skip writing it to the output array
    if(newRow.every(function(v){ return v == null; })) { continue; }

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
          newRow[col] = new Date(dateValue.getFullYear(), dateValue.getMonth(), 1);
        }
        // create the date value for the "Week" column
        else if (columnName == 'Week') {
          newRow[col] = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate() - dateValue.getDay());
        }
      }
      // if the column is "Date Added", insert today's date
      else if(columnName == 'Date Added') {
        newRow[col] = new Date();
      }
      // if the column is "Metadata", create a stringified JSON object with the all columns mapped to ="Metadata"
      else if(columnName == 'Metadata') {
        var m = { };

        // iterate through all columns in the imported data
        for(var c = 0; c < data[0].length; c++) {
          // if the column has header "Metadata", add the cell data to the `m` object
          if(data[0][c] == 'Metadata') { m[data[1][c]] = data[row][c]; }
        }

        // write a stringified `m` to the row if it contains data
        newRow[col] = Object.keys(m).length? JSON.stringify(m):null;
      }
    }

    // push `newRow` into the output array
    output.push(newRow);
  }
  // log total number of rows in output array
  msg = `Total of [${output.length}] new transaction rows to be added from import sheet [${importSheetName}]`
  console.log(msg)

  // write the `output` array to the destination sheet
  transactions.insertRowsBefore(2, output.length);
  transactions.getRange(2, 1, output.length, output[0].length).setValues(output);
  // flush SpreadsheetApp to update the data
  SpreadsheetApp.flush()
  // make Transactions the active sheet
  SpreadsheetApp.setActiveSheet(transactions);
  // report success
  msg = `Imported [${output.length}] rows from [${importSheetName}] into Transactions sheet.`
  console.log(msg)
  return ui.alert('Success', msg, ui.ButtonSet.OK);
}
// check if the script and sheet timezones match
function checkTimezone(ss) {
  // get the timezone of the spreadsheet and the script
  var sheetTimezone = ss.getSpreadsheetTimeZone();
  var scriptTimezone = Session.getScriptTimeZone();
 
  if (sheetTimezone !== scriptTimezone) {
    var msg = `Timezone discrepancy detected! Sheet timezone: [${sheetTimezone}] Script timezone: [${scriptTimezone}]`
    console.warn(msg);
    //return array with false and message
    return [false, msg];
  } else {
    var msg = `Timezone settings match: Sheet timezone: [${sheetTimezone}] Script timezone: [${scriptTimezone}]`
    console.log(msg);
    //return array with true and message
    return [true, msg];
  }
}