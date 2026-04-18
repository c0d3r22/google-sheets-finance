// ============================================================
// Finance Tracker - Balance Calculations Module
// Handles balance calculations and bank account integrations
// ============================================================

// TODO::Projected Balance is weird lol Figure it out
// Current Balance: The current balance of the bank account
// Ending Balance: The current balance minus all the transactions for the month
// Projected Balance: The current balance minus the projected items
function updateBalances(accountID) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // OPTIMIZATION: Only fetch the data range we need (avoid fetching empty cells)
  var dataRange = sheet.getRange(2, 2, lastRow - 5, 3); // Columns B, C, D (Amount, Description, Status)
  var data = dataRange.getValues();
  
  var paidAmount = 0;
  var totalAmount = 0;
  var projectedBalance = 0;

  // Calculate paid amount, total amount and projected balance
  for (var i = 0; i < data.length; i++) {
    var amount = data[i][0]; // Column B - Amount (now at index 0)
    var status = data[i][2];  // Column D - Status (now at index 2)

    if (typeof amount === 'number') {
      if (status === 'Paid') {
        paidAmount += amount;
      } else if (status === 'Pending') {
        totalAmount += amount;
      } else if (status === 'Project-Transaction') {
        totalAmount += amount;
        projectedBalance += amount;
      }
    }
  }

  // get bank balance from properties service for efficiency
  var accountBalance = getBalanceFromProperties(accountID);

  // We are adding because of negative cost vs positive income
  var endingBalance = parseFloat(accountBalance) + parseFloat(totalAmount);
  projectedBalance = parseFloat(accountBalance) + parseFloat(projectedBalance);

  // OPTIMIZATION: Update all three balance cells in a single batch operation
  var balanceRow = lastRow - 2;
  var balanceUpdates = [
    [accountBalance],
    [endingBalance],
    [projectedBalance]
  ];
  
  sheet.getRange(balanceRow, 4, 3, 1)
    .setNumberFormat("$#,##0.00")
    .setValues(balanceUpdates);
}

function isNegativeNumber(number) {
  if (typeof number !== 'number') {
    throw new Error('Input is not a valid number');
  }

  if (number < 0) {
    return true;
  } else {
    return false;
  }
}

function getBalanceFromProperties(accountID) {
  var properties = PropertiesService.getUserProperties();

  if (accountID === 'interestCheckingID') {
    return properties.getProperty('InterestCheckingBalance');
  } else if (accountID === 'onlineSavingsID_M') {
    return properties.getProperty('OnlineSavingsBalance_M');
  } else if (accountID === 'onlineSavingsID_A&M') {
    return properties.getProperty('OnlineSavingsBalance_AM');
  }

  return null; // Return null if the account ID is not recognized
}

function setAccountBalance(accountID, accountBalance)
{
  accountID = 'InterestCheckingBalance';
  accountBalance = 5811.53;
  PropertiesService.getUserProperties().setProperty(accountID, accountBalance);
}

function printBalances()
{
  var checkingBalance = getAccountBalance('interestCheckingID');
  Logger.log('Interest Checking Balance: $' + checkingBalance);

  var savingsBalance_M = getAccountBalance('onlineSavingsID_M');
  Logger.log('Online Savings (M) Balance: $' + savingsBalance_M);

  var savingsBalance_AM = getAccountBalance('onlineSavingsID_A&M');
  Logger.log('Online Savings (A&M) Balance: $' + savingsBalance_AM);
}

function storeBalances() {
  var checkingBalance = getAccountBalance('interestCheckingID');
  PropertiesService.getUserProperties().setProperty('InterestCheckingBalance', checkingBalance);

  var savingsBalance_M = getAccountBalance('onlineSavingsID_M');
  PropertiesService.getUserProperties().setProperty('OnlineSavingsBalance_M', savingsBalance_M);

  var savingsBalance_AM = getAccountBalance('onlineSavingsID_A&M');
  PropertiesService.getUserProperties().setProperty('OnlineSavingsBalance_AM', savingsBalance_AM);
}

function getAccountBalance(accountID) {
  return getBalanceFromProperties(accountID);
  /* Deprecated because Plaid sucks...
  const ACCESS_TOKEN = PropertiesService.getUserProperties().getProperty('plaidAccessToken');
  const CLIENT_ID    = PropertiesService.getUserProperties().getProperty('plaidClientId');
  const SECRET       = PropertiesService.getUserProperties().getProperty('plaidSecretId');
  const ACCOUNT      = PropertiesService.getUserProperties().getProperty(accountID);

  // Headers are a parameter plaid requires for the post request
  // Plaid takes a contentType parameter
  // Google app script takes a content-type parameter
  var headers = {
    'contentType' : 'application/json',
    'Content-Type': 'application/json',
  };

  // Data is a parameter plaid requires for the post request
  var data = {
    'access_token': ACCESS_TOKEN,
    'client_id'   : CLIENT_ID,
    'secret'      : SECRET,
    'options'     : {account_ids: [ACCOUNT],} // account ID
  };

  // Pass in the necessary headers
  // Pass the payload as a json object
  var parameters = {
    'headers'           : headers,
    'payload'           : JSON.stringify(data),
    'method'            : 'post',
    'muteHttpExceptions': true,
  };

  // API host + endpoint
  var url      = "https://development.plaid.com/accounts/balance/get";
  var response = UrlFetchApp.fetch(url, parameters);

  // Parse the response into a JSON object
  var json_data = JSON.parse(response);

  // Get the account balance from the JSON
  var accounts = json_data.accounts;
  var balance  = accounts[0].balances.available;

  return balance;
  */
}