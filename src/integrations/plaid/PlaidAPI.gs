// ============================================================
// Finance Tracker - Plaid Integration Module
// Handles Plaid API interactions for bank account linking and transactions
// ============================================================

// Account Properties:
// - interestCheckingID
// - onlineSavingsID_M
// - onlineSavingsID_A&M

function test()
{
  accountID = 'InterestCheckingBalance';
  accountBalance = 5701.20;
  PropertiesService.getUserProperties().setProperty(accountID, accountBalance);
  //getTransactionHistory('2023-04-01', '2023-04-26', PropertiesService.getUserProperties().getProperty('interestCheckingID'));

  // var matchingTransactions = searchTransactionByNames('2023-04-01', '2023-04-26', PropertiesService.getUserProperties().getProperty('interestCheckingID'), ["Internet transfer to", "Interest Paid"]);
  // matchingTransactions.forEach(function (transaction) {
  //     Logger.log('Date: ' + transaction.date);
  //     Logger.log('Amount: ' + transaction.amount);
  //     Logger.log('Name: ' + transaction.name);
  // });

  //markTransactionsAsPaid('interestCheckingID');

  //updateBalances('interestCheckingID');
}

function dailyUpdateTrigger()
{
  // Setup the account identifiers
  var account = 'interestCheckingID';
  var accountID = 'InterestCheckingBalance';

  // Get the balance from the 'Balances' sheet
  var balanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Balances');
  var accountBalance = balanceSheet.getRange(1, 2).getValue();

  // Set the properties service value with the account balance
  PropertiesService.getUserProperties().setProperty(accountID, accountBalance);

  // Update the balances
  updateBalances(account);
  /* Deprecated until Plaid doesn't cost freaking money...
  var account = 'interestCheckingID';
  // get the latest bank balances and update the properties service values
  storeBalances();

  // TODO::This can expand later to include other accounts
  // mark off any items as paid that have bank transactions
  markTransactionsAsPaid(account);

  updateBalances(account);
  */
}

function markTransactionsAsPaid(accountID) {
  // Get the account id
  const ACCOUNT = PropertiesService.getUserProperties().getProperty(accountID);

  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet data
  var sheet = spreadsheet.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Find the earliest and latest dates in the 'Date' column
  var startDate = null;
  var endDate = null;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var date = new Date(row[0]); // Assuming 'Date' column is at index 0

    if (startDate === null || date < startDate) {
      startDate = date;
    }

    if (endDate === null || date > endDate) {
      endDate = date;
    }
  }

  // Retrieve transaction data for the date range
  var transactions = getTransactionHistory(formatDate(startDate), formatDate(endDate), ACCOUNT);

  // OPTIMIZATION: Batch all status updates instead of cell-by-cell
  var statusUpdates = [];
  var updateRows = [];

  // Iterate through the spreadsheet data, starting from row 2
  for (var i = 1; i < data.length - 4; i++) {
    var row = data[i];
    var expenseDescription = row[2]; // Assuming 'Expense Description' column is at index 2
    var status = row[3]; // Assuming 'Status' column is at index 3

    if (expenseDescription !== "") {
      // Check for a match between transaction name and expense description
      for (var j = 0; j < transactions.length; j++) {
        var transaction = transactions[j];
        var transactionName = transaction.name.toLowerCase();
        var expenseDescriptionLower = expenseDescription.toLowerCase();

        var shouldMark = false;
        
        // Handle specific corner cases
        if (expenseDescriptionLower === "fidelity cc" && transactionName.includes("cardmember serv")) {
          shouldMark = true;
        } else if (expenseDescriptionLower === "xfinity" && transactionName.includes("apple card")) {
          shouldMark = true;
        } else if (expenseDescriptionLower === "pay day" && transactionName.includes("northrop grumman")) {
          shouldMark = true;
        } else if (expenseDescriptionLower === "apple card (incl. rent)" && transactionName.includes("applecard")) {
          shouldMark = true;
        } else if (transactionName.includes(expenseDescriptionLower)) {
          shouldMark = true;
        }
        
        if (shouldMark) {
          statusUpdates.push(['Paid']);
          updateRows.push(i + 1);
          break;
        }
      }
    }
  }

  // Apply all status updates in a single batch if there are any
  if (statusUpdates.length > 0) {
    for (var k = 0; k < updateRows.length; k++) {
      sheet.getRange(updateRows[k], 4).setValue(statusUpdates[k][0]);
    }
  }
}

function getTransactionHistory(startDate, endDate, accountID) {
  const ACCESS_TOKEN = PropertiesService.getUserProperties().getProperty('plaidAccessToken');
  const CLIENT_ID    = PropertiesService.getUserProperties().getProperty('plaidClientId');
  const SECRET       = PropertiesService.getUserProperties().getProperty('plaidSecretId');
  const COUNT        = 100; // number of default transactions, max is 500

  // headers are a parameter plaid requires for the post request
  // plaid takes a contentType parameter
  // google app script takes a content-type parameter
  var headers = {
    'contentType' : 'application/json',
    'Content-Type': 'application/json',
  };

  // data is a parameter plaid requires for the post request
  // created via the plaid quickstart app (node)
  var data = {
    'access_token': ACCESS_TOKEN,
    'client_id'   : CLIENT_ID,
    'secret'      : SECRET,
    'start_date'  : startDate,
    'end_date'    : endDate,
    'options'     : {count: COUNT, offset: 0,}
  };

  // pass in the necessary headers
  // pass the payload as a json object
  var parameters = {
    'headers'           : headers,
    'payload'           : JSON.stringify(data),
    'method'            : 'post',
    'muteHttpExceptions': true,
  };

  // api host + endpoint
  var url      = "https://development.plaid.com/transactions/get";
  var response = UrlFetchApp.fetch(url, parameters);

  // parse the response into a JSON object
  var json_data = JSON.parse(response);

  // get the transactions from the JSON
  var transactions = json_data.transactions;

  var transactionsOut = [];

  // loop through the transactions
  transactions.forEach(function (transactionObj) {
    if (transactionObj.account_id === accountID)
    {
      var transaction = processTransaction(transactionObj);
      transactionsOut.push(transaction);
      // Logger.log('Date: ' + transaction.date);
      // Logger.log('Amount: ' + transaction.amount);
      // Logger.log('Name: ' + transaction.name);
    }
  });

  return transactionsOut;
}

function processTransaction(transactionObj) {
  var transaction = {
    date: transactionObj.date,
    amount: -transactionObj.amount,
    name: transactionObj.name
  };

  return transaction;
}

function searchTransactionByNames(startDate, endDate, accountID, inputTextArray) {
  var matchingTransactions = [];

  const ACCESS_TOKEN = PropertiesService.getUserProperties().getProperty('plaidAccessToken');
  const CLIENT_ID = PropertiesService.getUserProperties().getProperty('plaidClientId');
  const SECRET = PropertiesService.getUserProperties().getProperty('plaidSecretId');
  const COUNT = 100; // number of default transactions, max is 500

  var headers = {
    'contentType': 'application/json',
    'Content-Type': 'application/json',
  };

  var data = {
    'access_token': ACCESS_TOKEN,
    'client_id': CLIENT_ID,
    'secret': SECRET,
    'start_date': startDate,
    'end_date': endDate,
    'options': {
      count: COUNT,
      offset: 0,
    }
  };

  var parameters = {
    'headers': headers,
    'payload': JSON.stringify(data),
    'method': 'post',
    'muteHttpExceptions': true,
  };

  var url = "https://development.plaid.com/transactions/get";
  var response = UrlFetchApp.fetch(url, parameters);

  var json_data = JSON.parse(response);
  var transactions = json_data.transactions;

  transactions.forEach(function (transactionObj) {
    if (transactionObj.account_id === accountID && containsAny(transactionObj.name.toLowerCase(), inputTextArray)) {
      var transaction = processTransaction(transactionObj);
      matchingTransactions.push(transaction);
    }
  });

  // Return the matching transactions
  return matchingTransactions;
}

function containsAny(str, substrings) {
  for (var i = 0; i < substrings.length; i++) {
    var substring = substrings[i].toLowerCase();
    if (str.includes(substring)) {
      return true;
    }
  }
  return false;
}

function printTransactionDetails(transactionObj) {
  // Accessing individual properties from the transaction object
  var paymentChannel = transactionObj.payment_channel;
  var transactionType = transactionObj.transaction_type;
  var pendingTransactionId = transactionObj.pending_transaction_id;
  var pending = transactionObj.pending;
  var paymentMeta = transactionObj.payment_meta;
  var date = transactionObj.date;
  var amount = transactionObj.amount;
  var authorizedDate = transactionObj.authorized_date;
  var merchantName = transactionObj.merchant_name;
  var datetime = transactionObj.datetime;
  var personalFinanceCategory = transactionObj.personal_finance_category;
  var authorizedDatetime = transactionObj.authorized_datetime;
  var category = transactionObj.category;
  var name = transactionObj.name;
  var accountOwner = transactionObj.account_owner;
  var transactionId = transactionObj.transaction_id;
  var categoryId = transactionObj.category_id;
  var isoCurrencyCode = transactionObj.iso_currency_code;
  var accountId = transactionObj.account_id;
  var unofficialCurrencyCode = transactionObj.unofficial_currency_code;
  var location = transactionObj.location;

  // Logging the values for testing
  Logger.log('Payment Channel: ' + paymentChannel);
  Logger.log('Transaction Type: ' + transactionType);
  Logger.log('Pending Transaction ID: ' + pendingTransactionId);
  Logger.log('Pending: ' + pending);
  Logger.log('Payment Meta: ' + paymentMeta);
  Logger.log('Date: ' + date);
  Logger.log('Amount: ' + amount);
  Logger.log('Authorized Date: ' + authorizedDate);
  Logger.log('Merchant Name: ' + merchantName);
  Logger.log('Datetime: ' + datetime);
  Logger.log('Personal Finance Category: ' + personalFinanceCategory);
  Logger.log('Authorized Datetime: ' + authorizedDatetime);
  Logger.log('Category: ' + category);
  Logger.log('Name: ' + name);
  Logger.log('Account Owner: ' + accountOwner);
  Logger.log('Transaction ID: ' + transactionId);
  Logger.log('Category ID: ' + categoryId);
  Logger.log('ISO Currency Code: ' + isoCurrencyCode);
  Logger.log('Account ID: ' + accountId);
  Logger.log('Unofficial Currency Code: ' + unofficialCurrencyCode);
  Logger.log('Location: ' + location);
}
