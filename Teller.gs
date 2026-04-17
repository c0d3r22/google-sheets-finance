// ============================================================
// Teller.io Daily Balance Tracker — Google Apps Script
// ============================================================
// Setup:
//   1. In Apps Script, go to Project Settings > Script Properties
//   2. Add the following properties:
//      - TELLER_ACCESS_TOKEN  : Your Teller access token
//      - TELLER_CERT          : Your Teller certificate (PEM, single-line with \n)
//      - TELLER_PRIVATE_KEY   : Your Teller private key (PEM, single-line with \n)
//   3. Run setupDailyTrigger() once to enable daily automation
// ============================================================

const TELLER_BASE_URL = "https://api.teller.io";

// ------------------------------------------------------------
// Core: Fetch all accounts and their balances
// ------------------------------------------------------------

function fetchAccountBalances() {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty("TELLER_ACCESS_TOKEN");
  const cert = props.getProperty("TELLER_CERT");
  const privateKey = props.getProperty("TELLER_PRIVATE_KEY");

  if (!accessToken || !cert || !privateKey) {
    throw new Error(
      "Missing one or more required Script Properties: " +
      "TELLER_ACCESS_TOKEN, TELLER_CERT, TELLER_PRIVATE_KEY"
    );
  }

  // Teller requires mTLS (mutual TLS). Apps Script's UrlFetchApp
  // supports client certificates via the 'mutualtls' option.
  const options = {
    method: "GET",
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(accessToken + ":"),
      "Content-Type": "application/json",
    },
    // Provide the client certificate and private key for mTLS
    mutualtls: {
      cert: cert.replace(/\\n/g, "\n"),
      key: privateKey.replace(/\\n/g, "\n"),
    },
    muteHttpExceptions: true,
  };

  // 1. Get all linked accounts
  const accountsResp = UrlFetchApp.fetch(`${TELLER_BASE_URL}/accounts`, options);
  handleApiError(accountsResp, "accounts");
  const accounts = JSON.parse(accountsResp.getContentText());

  // 2. For each account, fetch its balance
  const balances = accounts.map((account) => {
    const balResp = UrlFetchApp.fetch(
      `${TELLER_BASE_URL}/accounts/${account.id}/balances`,
      options
    );
    handleApiError(balResp, `balance for account ${account.id}`);
    const balData = JSON.parse(balResp.getContentText());

    return {
      accountId: account.id,
      institutionName: account.institution?.name ?? "Unknown",
      accountName: account.name,
      accountType: account.type,
      subtype: account.subtype,
      currency: account.currency,
      ledgerBalance: balData.ledger,
      availableBalance: balData.available,
      lastUpdated: new Date().toISOString(),
    };
  });

  return balances;
}

// ------------------------------------------------------------
// Log balances to a Google Sheet
// ------------------------------------------------------------

function logBalancesToSheet(balances) {
  const SHEET_NAME = "Daily Balances";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "Date",
      "Account ID",
      "Institution",
      "Account Name",
      "Type",
      "Subtype",
      "Currency",
      "Ledger Balance",
      "Available Balance",
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4A86E8")
      .setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  balances.forEach((b) => {
    sheet.appendRow([
      today,
      b.accountId,
      b.institutionName,
      b.accountName,
      b.accountType,
      b.subtype,
      b.currency,
      b.ledgerBalance,
      b.availableBalance,
    ]);
  });

  Logger.log(`Logged ${balances.length} account balance(s) for ${today}.`);
}

// ------------------------------------------------------------
// Main daily job — called by the time-based trigger
// ------------------------------------------------------------

function dailyBalanceSync() {
  try {
    Logger.log("Starting daily Teller balance sync...");
    const balances = fetchAccountBalances();
    logBalancesToSheet(balances);
    Logger.log("Daily sync complete.");
  } catch (err) {
    Logger.log("ERROR during daily sync: " + err.message);
    sendErrorEmail(err.message);
  }
}

// ------------------------------------------------------------
// Optional: Email alert on failure
// ------------------------------------------------------------

function sendErrorEmail(errorMessage) {
  const recipient = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: recipient,
    subject: "⚠️ Teller Balance Sync Failed",
    body:
      "The daily Teller.io balance sync encountered an error:\n\n" +
      errorMessage +
      "\n\nCheck the Apps Script logs for details.",
  });
}

// ------------------------------------------------------------
// Utility: Surface API errors clearly
// ------------------------------------------------------------

function handleApiError(response, context) {
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    const body = response.getContentText();
    throw new Error(
      `Teller API error fetching ${context} — HTTP ${code}: ${body}`
    );
  }
}

// ------------------------------------------------------------
// One-time setup: Create a daily time-based trigger
// ------------------------------------------------------------

function setupDailyTrigger() {
  // Remove any existing triggers for dailyBalanceSync to avoid duplicates
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "dailyBalanceSync") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Schedule to run every day at 7 AM (script timezone)
  ScriptApp.newTrigger("dailyBalanceSync")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  Logger.log("Daily trigger created: dailyBalanceSync will run every day at 7 AM.");
}

// ------------------------------------------------------------
// Optional: Remove all triggers (for cleanup/reset)
// ------------------------------------------------------------

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "dailyBalanceSync") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Trigger removed.");
    }
  });
}
