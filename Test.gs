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

  organizeSheetsAlphabetically();
}
///////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////
function testStore()
{
  var numBanks = 5;
  storeVariables(numBanks);
}

function testRetrieve()
{
  var key = 1;
  var bankValues = retrieveVariablesByKey(key);
  Logger.log(bankValues);
}

function storeVariables(numBanks) {
  var output = [];

  for (var i = 1; i <= numBanks; i++) {
    var dateColumn = (i - 1) * 5 + 1;
    var amountColumn = (i - 1) * 5 + 2;
    var expenseDescriptionColumn = (i - 1) * 5 + 3;
    var statusColumn = (i - 1) * 5 + 4;

    var bankVariables = [dateColumn, amountColumn, expenseDescriptionColumn, statusColumn];
    output.push([i, bankVariables]);
  }

  var properties = PropertiesService.getUserProperties();
  properties.setProperty('bankColumnVariables', JSON.stringify(output));
}

function retrieveVariablesByKey(key) {
  var properties = PropertiesService.getUserProperties();
  var bankVariables = JSON.parse(properties.getProperty('bankColumnVariables'));

  for (var i = 0; i < bankVariables.length; i++) {
    if (bankVariables[i][0] === key) {
      return bankVariables[i][1];
    }
  }

  return null; // Key not found
}
