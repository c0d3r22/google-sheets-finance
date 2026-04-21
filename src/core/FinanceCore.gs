// ============================================================
// Finance Tracker - Core Module
// Main application logic for Google Sheets finance management
// ============================================================

// TODO::getAndSet, make another just for get/log
function getPayPeriods(year, month) {
  var startDate = new Date('January 5, 2024');
  var properties = PropertiesService.getScriptProperties();

  // Calculate the paydays for the year
  var paydays = [];
  var currentDate = new Date(startDate);
  while (currentDate.getFullYear() <= year) {
    if (currentDate.getFullYear() === year) {
      paydays.push(new Date(currentDate));
    }
    currentDate.setDate(currentDate.getDate() + 14); // Move to the next payday (fortnightly)
  }

  // Filter the paydays for the specified month (Note: JavaScript months are 0-indexed)
  var monthPaydays = paydays.filter(function(date) {
    return date.getMonth() === month - 1;
  });

  // Store the pay periods in properties
  var key = 'payPeriods_' + year + '_' + month;
  properties.setProperty(key, JSON.stringify(monthPaydays));

  return monthPaydays;
}

function propertiesTest()
{
  var data = PropertiesService.getDocumentProperties();
  Logger.log(data.getKeys());
  /*
  [defaultStatuses, previousExpenseDates, expenses, plaidLinkToken]
  */

  var userData = PropertiesService.getUserProperties();
  Logger.log(userData.getKeys());
  /*
  [plaidClientId, InterestCheckingBalance, OnlineSavingsBalance_AM, plaidAccessToken, OnlineSavingsBalance_M, interestCheckingID, onlineSavingsID_M, plaidSecretId, onlineSavingsID_A&M, bankVariables, bankColumnVariables, expenses]
  */

  var scriptData = PropertiesService.getScriptProperties();
  Logger.log(scriptData.getKeys());
  /*
  [payPeriods_2024_11, payPeriods_2024_7, payPeriods_2024_12, payPeriods_2024_6, payPeriods_2024_9, payPeriods_2024_10, payPeriods_2024_8, payPeriods_2024_3, payPeriods_2024_2, payPeriods_2024_5, payPeriods_2024_4, payPeriods_2024_1, bankVariables]
  */
  //Logger.log(scriptData.getProperty('payPeriods_2024_1'));

  /*
  // Example usage
  var year = 2024;
  var month = 1;
  var payPeriods = getPayPeriods(year, month);
  Logger.log(payPeriods);
  */
}
function tempSetter()
{
  var year = 2024;
  for (var month = 1; month <= 12; month++)
  {
    var payPeriods = getPayPeriods(year, month);
  }
}

// TODO::Update the createSheet to account for multiple banks separated left to right
//       This could simply be with a properties service value tied to the SHEET that keeps
//       track of the number of banks added. Update HTML to create multiple pages from inputs ("+")
//       The HTML will create the sheets with the number of banks specified by the user.
//       - Use the input values when sheet is created to set properties service keys that are
//       <sheetName>NumberOfExpenses and <sheetName>NumberOfBanks to track their corresponding values.
function deleteEmptyRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const numRows = sheet.getMaxRows();
  const rowsToDelete = numRows - lastRow;
  if (rowsToDelete > 0) {
    sheet.deleteRows(lastRow + 1, rowsToDelete);
  }
}

function deleteEmptyColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastColumn = sheet.getLastColumn();
  const numColumns = sheet.getMaxColumns();
  const columnsToDelete = numColumns - lastColumn;
  if (columnsToDelete > 0) {
    sheet.deleteColumns(lastColumn + 1, columnsToDelete);
  }
}

function setExampleExpenses()
{
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('expenses', JSON.stringify({
    "Office Supplies": 100,
    "Coffee Meeting": 50,
    "Phone Bill": 75,
    "Business Lunch": 120,
    "Website Hosting": 80,
    "Office Snacks": 25,
    "Team Building Activity": 90,
    "Travel Expenses": 150,
    "Conference Registration": 200,
    "Advertising": 100,
    "Office Furniture": 75,
    "Software Subscriptions": 120,
    "Marketing Expenses": 80,
    "Printing Costs": 25
  }));
}

function storePreviousExpenseDates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 5, 1).getValues();
  var dates = data.map(function(row) {
    return row[0].getTime();
  });
  PropertiesService.getDocumentProperties().setProperty('previousExpenseDates', JSON.stringify(dates));
  Logger.log(dates);
}

function setPreviousExpenseDates(numExpenses) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var datesString = PropertiesService.getDocumentProperties().getProperty('previousExpenseDates');
  if (!datesString) {
    return;
  }
  var dates = JSON.parse(datesString).map(function(timestamp) {
    return new Date(timestamp);
  });
  var range = sheet.getRange(2, 1, dates.length, 1);
  range.setValues(dates.map(function(date) {
    return [date];
  }));
}

function setDefaultStatuses(newSheet = false) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = newSheet === true ? sheet.getRange(2, 4, sheet.getLastRow() - 1, 1) :
                                  sheet.getRange(2, 4, sheet.getLastRow() - 5, 1);
  var statuses = new Array(range.getNumRows()).fill(['Pending']);
  range.setValues(statuses);
  // PropertiesService.getDocumentProperties().setProperty('defaultStatuses', JSON.stringify(statuses));
}

// function setDefaultStatuses() {
//   var sheet = SpreadsheetApp.getActiveSheet();
//   var statusesString = PropertiesService.getDocumentProperties().getProperty('defaultStatuses');
//   if (!statusesString) {
//     return;
//   }
//   var statuses = JSON.parse(statusesString);
//   var range = sheet.getRange(2, 4, statuses.length, 1);
//   range.setValues(statuses);
// }

function deleteAllProperties()
{
  var props = PropertiesService.getDocumentProperties();
  props.deleteAllProperties();
  // props.setProperty('expenses', JSON.stringify({}));
}

// TODO::Mull over how the Projected Balance works...
function createSheet(sheetName, numBanks, numExpenses) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getUserProperties();

  var sheet = ss.getSheetByName(sheetName);
  if (sheet === null) {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = ["Date", "Amount", "Expense Description", "Status"];
  sheet.getRange("A1:D1").setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center");

  // var expenseProperty = sheet.getSheetName() + "NumberOfExpenses";
  // props.setProperty(numExpensesName, JSON.stringify({}));

  // var bankProperty = sheet.getSheetName() + "NumberOfBanks";
  // props.setProperty(numBanksName, JSON.stringify({}));

  // TODO::Set dates from preset dates and include pay days through some function or map
  // setPreviousExpenseDates(numExpenses);
  // TODO::Create number of rows based on numExpenses value with todays date
  const datesRange = sheet.getRange(2, 1, numExpenses, 1);
  const numRows = datesRange.getNumRows();
  const numCols = datesRange.getNumColumns();

  // Parse sheet name to get month and year, then create first day of that month
  var dateParts = sheetName.split(' ');
  var monthName = dateParts[0];
  var year = parseInt(dateParts[1]);
  var month = getMonthFromString(monthName);
  var firstDayOfMonth = new Date(year, month, 1);

  const dateValues = Array(numRows).fill().map(() => Array(numCols).fill(firstDayOfMonth));
  datesRange.setValues(dateValues);

  // Set default statuses to "Pending"
  setDefaultStatuses(true);

  sheet.getRange(1, 1, sheet.getLastRow(), 4).setHorizontalAlignment("center");

  // Store expenses as a document property
  // - initially blank to be filled in by user
  var props = PropertiesService.getUserProperties();
  if (props.getProperty('expenses') === null) {
    Logger.log("Expenses property did not exist, creating it...");
    props.setProperty('expenses', JSON.stringify({}));
  }

  // Set dropdown data validation for expense description column
  updateExpenseDescriptionRule(sheet, true);

  // Set dropdown data validation for status column and set width of columns
  updateColumnWidths(sheet, true);

  var totalRow = sheet.getLastRow() + 2;
  sheet.getRange("C" + totalRow).setValue("Total").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D" + totalRow).setNumberFormat("$#,##0.00").setHorizontalAlignment("center");

  var endingBalanceRow = totalRow + 1;
  sheet.getRange("C" + endingBalanceRow).setValue("Ending Balance").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D" + endingBalanceRow).setNumberFormat("$#,##0.00").setHorizontalAlignment("center");

  var currentBalanceRow = endingBalanceRow + 1;
  sheet.getRange("C" + currentBalanceRow).setValue("Projected Balance").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D" + currentBalanceRow).setNumberFormat("$#,##0.00").setHorizontalAlignment("center");

  deleteEmptyRows();
  deleteEmptyColumns();

  // make sure the amount column is set to be currency
  sheet.getRange(2,2, sheet.getLastRow() - 3, 1).setNumberFormat("$#,##0.00");

  updateBalances('interestCheckingID'); // TODO::Expand for other accounts

  sortSheetsByMonth();
}

function createThisMonthSheet(numBanks, numExpenses) {
  //TODO::Figure out the best way to pass in the numBanks and numExpenses
  numBanks = 1;
  numExpenses = 10;

  var today = new Date();
  var nextMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var sheetName = Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), "MMMM yyyy");

  if (sheetExists(sheetName)) {
    SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' already exists!  Please choose another name");
  } else {
    createSheet(sheetName, numBanks, numExpenses);
  }
}

function createNextMonthSheet(numBanks, numExpenses) {
  //TODO::Figure out the best way to pass in the numBanks and numExpenses
  numBanks = 1;
  numExpenses = 10;

  var today = new Date();
  var nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  var sheetName = Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), "MMMM yyyy");

  if (sheetExists(sheetName)) {
    SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' already exists! Please choose another name");
  } else {
    createSheet(sheetName, numBanks, numExpenses);
  }
}

function createCombinedMonthSheet(accountID) {
  accountID = accountID || 'interestCheckingID';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var monthSheets = sheets.filter(function(sheet) {
    return !sheet.isSheetHidden() && isMonthYearSheetName(sheet.getName());
  });

  if (monthSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No active month/year sheets found to combine.');
    return;
  }

  var combinedSheetName = 'Combined Months';
  var existing = ss.getSheetByName(combinedSheetName);
  if (existing) {
    ss.deleteSheet(existing);
  }

  // Get a fresh list of sheets after deletion
  sheets = ss.getSheets();
  var visibleSheets = sheets.filter(function(sheet) {
    return !sheet.isSheetHidden();
  });
  if (visibleSheets.length > 0) {
    ss.setActiveSheet(visibleSheets[visibleSheets.length - 1]);
  }

  monthSheets.sort(function(a, b) {
    return getMonthYearValue(a.getName()) - getMonthYearValue(b.getName());
  });

  var combinedSheet = ss.insertSheet(combinedSheetName);
  var headerValues = monthSheets[0].getRange('A1:D1').getValues();
  combinedSheet.getRange(1, 1, 1, 4).setValues(headerValues).setFontWeight('bold').setHorizontalAlignment('center');

  var combinedValues = [];
  var validationRule3 = null;
  var validationRule4 = null;

  monthSheets.forEach(function(sheet) {
    var sourceLastRow = sheet.getLastRow();
    var expenseRowCount = Math.max(0, sourceLastRow - 5);
    if (expenseRowCount > 0) {
      var values = sheet.getRange(2, 1, expenseRowCount, 4).getValues();
      combinedValues = combinedValues.concat(values);

      if (!validationRule3) {
        validationRule3 = sheet.getRange('C2:C2').getDataValidations()[0][0];
      }
      if (!validationRule4) {
        validationRule4 = sheet.getRange('D2:D2').getDataValidations()[0][0];
      }
    }
  });

  if (combinedValues.length > 0) {
    combinedSheet.getRange(2, 1, combinedValues.length, 4).setValues(combinedValues);
    combinedSheet.getRange(2, 2, combinedValues.length, 1).setNumberFormat('$#,##0.00');

    if (validationRule3) {
      combinedSheet.getRange(2, 3, combinedValues.length, 1).setDataValidation(validationRule3);
    }
    if (validationRule4) {
      combinedSheet.getRange(2, 4, combinedValues.length, 1).setDataValidation(validationRule4);
    }
  }

  var bufferRow = combinedValues.length + 2;
  combinedSheet.getRange(bufferRow, 1, 1, 4).clearContent();

  var totalRow = bufferRow + 1;
  var endingBalanceRow = bufferRow + 2;
  var projectedBalanceRow = bufferRow + 3;

  combinedSheet.getRange(totalRow, 3).setValue('Total').setFontWeight('bold').setHorizontalAlignment('center');
  combinedSheet.getRange(totalRow, 4).setNumberFormat('$#,##0.00').setHorizontalAlignment('center');

  combinedSheet.getRange(endingBalanceRow, 3).setValue('Ending Balance').setFontWeight('bold').setHorizontalAlignment('center');
  combinedSheet.getRange(endingBalanceRow, 4).setNumberFormat('$#,##0.00').setHorizontalAlignment('center');

  combinedSheet.getRange(projectedBalanceRow, 3).setValue('Projected Balance').setFontWeight('bold').setHorizontalAlignment('center');
  combinedSheet.getRange(projectedBalanceRow, 4).setNumberFormat('$#,##0.00').setHorizontalAlignment('center');

  var usedRows = projectedBalanceRow;
  combinedSheet.getRange(1, 1, usedRows, 4).setHorizontalAlignment('center');

  combinedSheet.autoResizeColumns(1, 2);
  combinedSheet.setColumnWidth(3, 150);
  combinedSheet.setColumnWidth(4, 150);

  ss.setActiveSheet(combinedSheet);
  updateBalances(accountID);
  deleteEmptyRows();
  deleteEmptyColumns();
  ss.moveActiveSheet(ss.getSheets().length);
}

function isMonthYearSheetName(sheetName) {
  return /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}$/.test(sheetName);
}

function getMonthYearValue(sheetName) {
  var parts = sheetName.split(' ');
  var month = getMonthFromString(parts[0]);
  var year = parseInt(parts[1], 10);
  return year * 100 + month;
}

function sheetExists(sheetName) {
  // OPTIMIZATION: Use getSheetByName() instead of iterating through all sheets
  // Returns null if not found, which is faster than manual loop
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) !== null;
}

function updateColumnWidths(sheet, newSheet = false)
{
  // Set dropdown data validation for status column
  var statusColumn = newSheet === true ? sheet.getRange(2,4,sheet.getLastRow() - 1, 1) :
                                         sheet.getRange(2,4,sheet.getLastRow() - 5, 1);
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(["Paid", "Pending", "Project-Transaction"]).build();
  statusColumn.setDataValidation(statusRule);

  // OPTIMIZATION: Parse expenses JSON once and cache status options
  var statusOptions = ["Paid", "Pending", "Project-Transaction"];
  var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
  var expenseOptions = Object.keys(expenses);

  // Find maximum length among all options
  var maxLength = 0;
  statusOptions.forEach(function(option) {
    maxLength = Math.max(maxLength, option.length);
  });
  expenseOptions.forEach(function(option) {
    maxLength = Math.max(maxLength, option.length);
  });
  maxLength = Math.max(maxLength, "Expense Description".length);

  sheet.setColumnWidth(3, maxLength * 8);
  sheet.setColumnWidth(4, maxLength * 8);

  sheet.autoResizeColumns(1, 2);
}

function onEdit(e) {
  onEditedItem(e);
  // Also check for any empty rows that might need setup after any edit
  setupEmptyRowsIfNeeded();
}

function onChange(e) {
  if (e.changeType === 'INSERT_ROW') {
    // Check if we need to set up any new rows after insertion
    setupEmptyRowsIfNeeded();
  }
}

// TODO::Somehow make this not trigger when rows are being deleted
function onEditedItem(e) {
  var sheet = e.range.getSheet();
  var editedCol = e.range.getColumn();
  var editedRow = e.range.getRow();

  if (rowNeedsSetup(editedRow)) {
    setupRowAtPosition(editedRow);
  }

  // If expense item, or status columns are edited
  if (editedCol == 1 || editedCol == 3 || editedCol == 4) {
    // If expense item column is edited
    if (editedCol == 3) {
      var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
      var expenseItem = sheet.getRange(editedRow, 3).getValue();
      sheet.getRange(editedRow, 2).setValue(expenses[expenseItem]);
      // Update status column to pending when there is no expense description
      if (expenseItem === "") {
        sheet.getRange(editedRow, 4).setValue('Pending');
      }
    }

    // TODO::Is this the right place? So when I edit the date, the expense description or the status...
    updateBalances('interestCheckingID');
  }

  // resize columns to fit the expense amounts
  sheet.autoResizeColumns(1, 2);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Edit Expenses')
    .addItem('Edit/Update Expense Items', 'editExpenses')
    .addItem('Create New Sheet', 'showModal')
    .addItem('Create This Months Sheet', 'createThisMonthSheet')
    .addItem('Create Next Months Sheet', 'createNextMonthSheet')
    .addItem('Combine Monthly Sheets', 'createCombinedMonthSheet')
    .addItem('Plaid - Link Bank Account', 'linkSetup')
    .addItem('Update transactions', 'dailyUpdateTrigger')
    .addItem('Unhide All Sheets', 'unhideAllSheets')
    .addItem('Organize Sheets', 'sortSheetsByMonth')
    .addToUi();

  // Ensure onChange trigger is installed
  installOnChangeTrigger();
}

function installOnChangeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var hasOnChange = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() === ScriptApp.EventType.ON_CHANGE) {
      hasOnChange = true;
      break;
    }
  }
  
  if (!hasOnChange) {
    ScriptApp.newTrigger('onChange')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onChange()
      .create();
  }
}

function showModal() {
  var html = HtmlService.createHtmlOutputFromFile('ui/CreateSheetModal')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Sheet');
}

// called from the modal.html, creates the new sheet
function processForm(form) {
  var sheetName = form.sheetName;
  var numBanks = form.numBanks;
  var numExpenses = form.defaultExpenses;

  if (sheetExists(sheetName)) {
    SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' already exists! Please choose another name");
  } else {
    createSheet(sheetName, numBanks, numExpenses);
  }
}

function editExpenses() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('expenses') === null) {
    Logger.log("Expenses property did not exist, creating...");
    props.setProperty('expenses', JSON.stringify({}));
  }

  var expenses = JSON.parse(props.getProperty('expenses'));
  var expenseItems = Object.keys(expenses).sort();

  var html = HtmlService.createTemplateFromFile('ui/EditExpenses');
  html.expenseItems = expenseItems;
  html.expenses = expenses;

  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), 'Edit Expenses');
}

function updateExpenses(expenseValues) {
  var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
  for (var expenseItem in expenseValues) {
    expenses[expenseItem] = expenseValues[expenseItem];
  }
  PropertiesService.getDocumentProperties().setProperty('expenses', JSON.stringify(expenses));
  refreshAmountColumn();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  updateExpenseDescriptionRule(sheet);

  updateColumnWidths(sheet);

  updateBalances('interestCheckingID'); // TODO::allow multiple accounts
}

function refreshAmountColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // OPTIMIZATION: Parse expenses JSON once
  var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
  var range = sheet.getRange(2, 2, sheet.getLastRow() - 5, 1);
  var descriptionRange = sheet.getRange(2, 3, sheet.getLastRow() - 5, 1);
  
  // OPTIMIZATION: Get all descriptions in one batch call
  var descriptions = descriptionRange.getValues();
  
  // Build array of amounts to set all at once
  var amounts = [];
  for (var i = 0; i < descriptions.length; i++) {
    var expenseItem = descriptions[i][0];
    amounts.push([expenses[expenseItem] || '']);
  }
  
  // Set all amounts in one batch operation
  range.setValues(amounts);
}

function deleteExpense(expenseItem) {
  var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
  delete expenses[expenseItem];
  PropertiesService.getDocumentProperties().setProperty('expenses', JSON.stringify(expenses));
  refreshAmountColumn();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  updateExpenseDescriptionRule(sheet);

  updateColumnWidths(sheet);
}

function onSelectionChange(e) {
  var sheet = e.source.getActiveSheet();

  // If the expense descriptions have not been updated yet
  updateExpenseDescriptionRule(sheet);

  // If the expense amounts have not been updated yet
  refreshAmountColumn();
}

function updateExpenseDescriptionRule(sheet, newSheet = false) {
  try {
    // OPTIMIZATION: Get expenses once and cache
    var expensesStr = PropertiesService.getDocumentProperties().getProperty('expenses');
    if (!expensesStr) {
      Logger.log("Expenses property not set");
      return;
    }
    
    var expenses = JSON.parse(expensesStr);
    var expenseDescriptionOptions = Object.keys(expenses).sort();
    
    if (expenseDescriptionOptions.length === 0) {
      Logger.log("No expense options available");
      return;
    }
    
    var expenseDescriptionColumn = newSheet === true ? sheet.getRange(2,3,sheet.getLastRow()-1,1) :
                                                       sheet.getRange(2,3,sheet.getLastRow()-5,1);
    var expenseDescriptionRule = SpreadsheetApp.newDataValidation().requireValueInList(expenseDescriptionOptions).build();
    expenseDescriptionColumn.setDataValidation(expenseDescriptionRule);
  } catch (e) {
    Logger.log("Error updating expense description rule: " + e.message);
  }
}

function onNewRowAdded() {
  // Specify the sheet name and range to monitor
  var sheetName = "Test2";
  var range = "A1:D20";

  // Get the spreadsheet object and sheet object
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Get the previous number of rows in the sheet
  var numRows = sheet.getLastRow();

  // Set up a trigger to run when a new row is added to the sheet
  ScriptApp.newTrigger("myFunction")
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  // Wait for a new row to be added to the sheet
  while (sheet.getLastRow() == numRows) {
    Utilities.sleep(1000);
  }
}

function myFunction() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var lastRow = sheet.getLastRow() - 3; // Subtract 3 from the last row

  for (var i = startRow; i < startRow + numRows; i++) {
    if (i <= lastRow) { // Add this condition to exclude the 4th row from the bottom
      var newRange = sheet.getRange(i, 1, 1, 4);
      var today = new Date();

      // Insert today's date in column 1 of the new row if it is empty
      var dateCell = newRange.getCell(1, 1);
      if (dateCell.isBlank() && (i != lastRow)) {
        dateCell.setValue(today);
      }

      // Copy the data validations from columns 3 and 4 to columns 3 and 4 of the new row if it is empty
      var sourceRange1 = sheet.getRange("C2:C2");
      var rule1 = sourceRange1.getDataValidations()[0][0];
      if (rule1 != null) {
        var expenseValidationCell = newRange.getCell(1, 3);
        if (expenseValidationCell.isBlank() && (i != lastRow)) {
          expenseValidationCell.setDataValidation(rule1);
        }
      }

      var sourceRange2 = sheet.getRange("D2:D2");
      var rule2 = sourceRange2.getDataValidations()[0][0];
      if (rule2 != null) {
        var statusValidationCell = newRange.getCell(1, 4);
        if (statusValidationCell.isBlank() && (i != lastRow)) {
          statusValidationCell.setDataValidation(rule2);
        }
      }

      // Insert "Pending" into column 4 of the new row if it is empty
      var statusCell = newRange.getCell(1, 4);
      if (statusCell.isBlank() && (i != lastRow)) {
        statusCell.setValue("Pending");
      }

      // Center the contents of all cells in the new row
      newRange.setHorizontalAlignment("center");
    }
  }
}

function rowNeedsSetup(rowNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var bufferRow = lastRow - 3; // the intentional blank row before summary rows

  if (rowNum < 2 || rowNum > bufferRow) {
    return false;
  }

  var rowRange = sheet.getRange(rowNum, 1, 1, 4);
  var rowValues = rowRange.getValues()[0];
  var dateCell = sheet.getRange(rowNum, 1);
  var expenseCell = sheet.getRange(rowNum, 3);
  var statusCell = sheet.getRange(rowNum, 4);

  var expenseRule = expenseCell.getDataValidations()[0][0];
  var statusRule = statusCell.getDataValidations()[0][0];

  var isRowEmpty = rowValues.every(function(cell) {
    return cell === '' || cell === null || cell === undefined;
  });

  // Keep the intentionally blank buffer row until the user starts typing into it.
  if (rowNum === bufferRow && isRowEmpty) {
    return false;
  }

  return isRowEmpty || dateCell.isBlank() || statusCell.isBlank() || !expenseRule || !statusRule;
}

function setupEmptyRowsIfNeeded() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var bufferRow = lastRow - 3; // the intentional blank row before summary rows

  for (var row = 2; row < bufferRow; row++) {
    if (rowNeedsSetup(row)) {
      setupRowAtPosition(row);
    }
  }

  if (rowNeedsSetup(bufferRow)) {
    setupRowAtPosition(bufferRow);
  }
}

function setupRowAtPosition(rowNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newRange = sheet.getRange(rowNum, 1, 1, 4);
  var today = new Date();

  // Insert today's date in column 1 of the new row if it's empty
  var dateCell = newRange.getCell(1, 1);
  if (dateCell.isBlank()) {
    dateCell.setValue(today);
  }

  // Copy data validations from columns 3 and 4 to the new row if they exist and cell is empty
  var sourceRange1 = sheet.getRange("C2:C2");
  var rule1 = sourceRange1.getDataValidations()[0][0];
  if (rule1 != null) {
    var expenseValidationCell = newRange.getCell(1, 3);
    if (expenseValidationCell.isBlank()) {
      expenseValidationCell.setDataValidation(rule1);
    }
  }

  var sourceRange2 = sheet.getRange("D2:D2");
  var rule2 = sourceRange2.getDataValidations()[0][0];
  if (rule2 != null) {
    var statusValidationCell = newRange.getCell(1, 4);
    if (statusValidationCell.isBlank()) {
      statusValidationCell.setDataValidation(rule2);
    }
  }

  // Insert "Pending" into column 4 of the new row if it's empty
  var statusCell = newRange.getCell(1, 4);
  if (statusCell.isBlank()) {
    statusCell.setValue("Pending");
  }

  // Center the contents of all cells in the new row
  newRange.setHorizontalAlignment("center");
}

function setupNewRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  // Empty row is always at lastRow - 3 (between expenses 2-N and summaries)
  var targetRow = lastRow - 3;
  setupRowAtPosition(targetRow);
}

function sortSheetsByMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // Make all sheets visible before sorting
  sheets.forEach(function(sheet) {
    if (sheet.getName() != 'Balances') {
      if (sheet.isSheetHidden()) {
        sheet.showSheet();
      }
    }
  });

  var sheetNames = sheets.filter(function(sheet) {
    return sheet.getName() !== 'Balances';}).map(function(sheet) {
      return sheet.getName();
  });
  //var sheetNames = sheets.map(function(sheet) {
  //  return sheet.getName();
  //});

  // Separate the Combined Months sheet from regular month sheets
  var combinedSheetName = 'Combined Months';
  var combinedSheetIndex = sheetNames.indexOf(combinedSheetName);
  var combinedSheet = null;
  if (combinedSheetIndex !== -1) {
    combinedSheet = sheets[combinedSheetIndex];
    // Remove from arrays for separate handling
    sheets.splice(combinedSheetIndex, 1);
    sheetNames.splice(combinedSheetIndex, 1);
  }

  // Function to parse sheet names and create a date object for comparison
  function parseSheetName(sheetName) {
    var parts = sheetName.split(' - ');
    var dateStr = parts[0];
    var archive = parts.length > 1;
    var dateParts = dateStr.split(' ');
    var month = getMonthFromString(dateParts[0]);
    var year = parseInt(dateParts[1]);
    var dateObj = new Date(year, month, 1);
    return {
      dateObj: dateObj,
      archive: archive,
      originalName: sheetName
    };
  }

  // Function to compare sheets by date and archive status
  function sheetComparator(a, b) {
    if (a.dateObj.getFullYear() !== b.dateObj.getFullYear()) {
      return b.dateObj.getFullYear() - a.dateObj.getFullYear();
    }
    if (a.dateObj.getMonth() !== b.dateObj.getMonth()) {
      return a.dateObj.getMonth() - b.dateObj.getMonth();
    }
    return a.archive - b.archive;
  }

  // Map sheet names to an array of objects containing date and archive info
  var sheetData = sheetNames.map(parseSheetName);

  // Sort the sheet data based on the comparator function
  sheetData.sort(sheetComparator);

  // Reorder and hide archived sheets in the spreadsheet
  sheetData.forEach(function(sheetInfo, index) {
    var sheet = ss.getSheetByName(sheetInfo.originalName);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
    // Hide the sheet if it is archived
    if (sheetInfo.archive) {
      sheet.hideSheet();
    }
  });

  // Move the Combined Months sheet to the end (after all sorted month sheets)
  if (combinedSheet) {
    ss.setActiveSheet(combinedSheet);
    ss.moveActiveSheet(sheets.length + 1); // Move to the end
  }
}

function getMonthFromString(mon) {
  return new Date(Date.parse(mon + " 1, 2023")).getMonth();
}

function unhideAllSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  sheets.forEach(function(sheet) {
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
  });
}
/*
function organizeSheetsAlphabetically() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.sort(function(a, b) {
    var nameA = a.getName().toUpperCase();
    var nameB = b.getName().toUpperCase();
    if (nameA < nameB) {
      return -1;
    }
    if (nameA > nameB) {
      return 1;
    }
    return 0;
  });

  for (var i = 0; i < sheets.length; i++) {
    ss.setActiveSheet(sheets[i]);
    ss.moveActiveSheet(i + 1);
  }
}

function organizeSheetsByMonth() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  // Map month names to numbers for sorting
  var monthNames = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12
  };

  // Create a list of sheets with additional data for sorting
  var sheetData = sheets.map(function(sheet) {
    var name = sheet.getName();
    var parts = name.split(" ");
    var month = monthNames[parts[0]] || 0;
    var year = parseInt(parts[1]) || 0;
    var hidden = sheet.isSheetHidden();

    return {
      sheet: sheet,
      month: month,
      year: year,
      hidden: hidden
    };
  });

  // Sort the list of sheet data
  sheetData.sort(function(a, b) {
    if(a.hidden !== b.hidden) {
      return a.hidden ? 1 : -1;
    }

    if(a.year !== b.year) {
      return a.year - b.year;
    }

    return a.month - b.month;
  });

  // Reorder the sheets and re-hide if necessary
  for(var i = 0; i < sheetData.length; i++) {
    spreadsheet.setActiveSheet(sheetData[i].sheet);
    spreadsheet.moveActiveSheet(i + 1);
    if(sheetData[i].hidden) {
      sheetData[i].sheet.hideSheet();
    }
  }
}
*/