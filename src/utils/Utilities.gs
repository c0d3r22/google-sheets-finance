// ============================================================
// Finance Tracker - Utility Functions Module
// Helper functions and test utilities
// ============================================================

/*
   The following functions are to be used with the html that creates only one dropdown and one input box.
   This may not be wanted as a user may want to edit multiple values at the same time and save them all
   at one time. This also makes the code more complicated which may not be warranted. These would use the
   editExpensesDropdown.html so the editExpenses() function would need to be updated to use that html.
*/
function getExpenseValue(expenseDescription) {
  var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
  return expenses[expenseDescription];
}

function updateExpense(expenseItem, expenseValue) {
  var expenses = JSON.parse(PropertiesService.getDocumentProperties().getProperty('expenses'));
  expenses[expenseItem] = expenseValue;
  PropertiesService.getDocumentProperties().setProperty('expenses', JSON.stringify(expenses));
  refreshAmountColumn();
}

function createNextMonthSheetWithData(numBanks, numExpenses) {
  // Get the current sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var numRows = sheet.getLastRow();
  var numCols = sheet.getLastColumn();
  var dataRange = sheet.getRange(1, 1, numRows-4, numCols);
  var data = dataRange.getValues();

  // Find the index of the Date and Expense Description columns
  var dateColumnIndex = -1;
  var expenseColumnIndex = -1;
  var headers = data[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === "Date") {
      dateColumnIndex = i;
    } else if (headers[i] === "Expense Description") {
      expenseColumnIndex = i;
    }
  }

  // Make sure we found the necessary columns
  if (dateColumnIndex < 0 || expenseColumnIndex < 0) {
    SpreadsheetApp.getUi().alert("Could not find Date and/or Expense Description columns!");
    return;
  }

  // Build an array of next month's dates and expense descriptions
  var nextMonthData = [];
  var today = new Date();
  for (var i = 1; i < data.length; i++) { // Start at index 1 to skip header row
    var row = data[i];
    var date = new Date(row[dateColumnIndex]);
    var description = row[expenseColumnIndex];
    //if (date.getMonth() === today.getMonth() && date.getFullYear() === today.getFullYear()) {
      // Only include expenses from the current month
      nextMonthData.push([
        new Date(date.getFullYear(), date.getMonth() + 1, date.getDate()),
        description
      ]);
    //}
  }

  // Create the new sheet with the next month's data
  var sheetName = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth() + 1, 1), Session.getScriptTimeZone(), "MMMM yyyy");
  if (sheetExists(sheetName)) {
    SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' already exists! Please choose another name");
  } else {
    createSheetWithData(sheetName, 1, nextMonthData.length, nextMonthData);
  }
}

function createSheetWithData(sheetName, numBanks, numExpenses, expenseData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getUserProperties();

  var sheet = ss.getSheetByName(sheetName);
  if (sheet === null) {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = ["Date", "Amount", "Expense Description", "Status"];
  sheet.getRange("A1:D1").setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center");

  const numRows = numExpenses;
  const numCols = headers.length;
  const dataRange = sheet.getRange(2, 1, numRows, numCols);
  const dateColumnIndex = 0;
  const descriptionColumnIndex = 2;
  for (var i = 0; i < expenseData.length; i++) {
    var row = expenseData[i];
    dataRange
    .getCell(i + 1, dateColumnIndex + 1)
    .setValue(row[0])
    .setNumberFormat("M/d/yyyy");
    dataRange
    .getCell(i + 1, descriptionColumnIndex + 1)
    .setValue(row[1]);
  }

  // Set default statuses to "Pending"
  setDefaultStatuses(true);

  sheet.getRange(1, 1, sheet.getLastRow(), numCols).setHorizontalAlignment("center");

  // Store expenses as a document property
  // - initially blank to be filled in by user
  // - carry over from already defined expenses
  var props = PropertiesService.getUserProperties();
  if (props.getProperty('expenses') === null) {
    Logger.log("Expenses property did not exist, creating it...");
    props.setProperty('expenses', JSON.stringify({}));
  }

  // Set dropdown data validation for expense description column
  updateExpenseDescriptionRule(sheet, true);

  // Set dropdown data validation for status column and set width of columns
  updateColumnWidths(sheet, true);
}