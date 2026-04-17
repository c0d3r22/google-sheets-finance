# Google Sheets Finance Tracker

A comprehensive personal finance management application built with Google Apps Script and Google Sheets.

## Features

- **Monthly Budget Tracking**: Create monthly expense sheets with automatic balance calculations
- **Bank Integration**: Connect with Plaid or Teller.io for automatic transaction syncing
- **Expense Management**: Customizable expense categories with predefined amounts
- **Balance Projections**: Calculate ending and projected balances based on pending transactions
- **Sheet Organization**: Automatic sorting and management of monthly sheets

## Project Structure

### Reorganized Code (Recommended)
```
src/
├── core/                    # Core application logic
│   ├── FinanceCore.gs      # Main application functions (sheet creation, UI, etc.)
│   └── BalanceCalculations.gs # Balance calculation and bank account management
├── integrations/           # Third-party API integrations
│   ├── plaid/              # Plaid banking API integration
│   │   ├── PlaidAPI.gs     # Transaction fetching and processing
│   │   └── PlaidLinking.gs # Account linking and token management
│   └── teller/             # Teller.io banking API integration
│       └── TellerAPI.gs    # mTLS-based balance tracking
├── ui/                     # HTML user interface components
│   ├── CreateSheetModal.html    # Modal for creating new sheets
│   ├── EditExpenses.html        # Expense management interface
│   └── PlaidLink.html           # Plaid account linking interface
└── utils/                  # Utility functions and helpers
    └── Utilities.gs        # Helper functions and test utilities
```

### Original Code Archive
```
original-code/              # Archived original codebase (flat structure)
├── Code.gs                # Original monolithic main script
├── plaidAPI.gs            # Original Plaid integration
├── accountLinkingPlaid.gs # Original Plaid account linking
├── Teller.gs              # Original Teller.io integration
├── Test.gs                # Original test utilities
├── htmlDropdownFunctions.gs # Original HTML helper functions
├── modal.html             # Original sheet creation modal
├── editExpenses.html      # Original expense editing interface
├── editExpensesDropdown.html # Original dropdown-based editing
├── PlaidLink.html         # Original Plaid linking interface
└── README.md              # Documentation for original code
```

The `original-code/` folder contains the working legacy codebase in its original flat file structure. This is preserved for reference and gradual migration. The `src/` folder contains the new modular, maintainable structure.

## Setup

### Option 1: Use Reorganized Code (Recommended)
1. **Create a new Google Sheet**
2. **Open Apps Script Editor**: Extensions → Apps Script
3. **Copy the organized code**: Copy files from the `src/` directory into your Apps Script project
4. **Configure Properties**:
   - For Plaid integration: Set up `plaidClientId`, `plaidSecretId`, and `plaidAccessToken`
   - For Teller integration: Set up `TELLER_ACCESS_TOKEN`, `TELLER_CERT`, and `TELLER_PRIVATE_KEY`

### Option 2: Use Original Code (Legacy)
1. **Create a new Google Sheet**
2. **Open Apps Script Editor**: Extensions → Apps Script
3. **Copy the original code**: Copy all files from the `original-code/` directory into your Apps Script project
4. **Configure Properties**: Same as above

The reorganized code in `src/` is recommended for new implementations as it provides better maintainability and modularity. The original code in `original-code/` is preserved for existing users who prefer to stick with the working legacy version.

## Usage

### Basic Operations
- **Create Monthly Sheets**: Use the "Edit Expenses" menu to create new monthly budget sheets
- **Manage Expenses**: Edit expense categories and their default amounts
- **Track Transactions**: Mark transactions as Paid, Pending, or Project-Transaction

### Bank Integration
- **Plaid Setup**: Use "Plaid - Link Bank Account" to connect your bank accounts
- **Teller Setup**: Configure mTLS certificates for Teller.io integration
- **Auto-Sync**: Set up daily triggers for automatic balance updates

### Balance Calculations
- **Current Balance**: Your actual bank account balance
- **Ending Balance**: Current balance minus all pending transactions
- **Projected Balance**: Current balance minus projected transactions

## API Integrations

### Plaid
- Transaction history fetching
- Account balance retrieval
- Automatic transaction status updates

### Teller.io
- Daily balance tracking with mTLS authentication
- Multi-account support
- Automated email alerts on sync failures

## Development

The codebase is organized into logical modules for maintainability:

- **Core**: Essential application functionality
- **Integrations**: Third-party API wrappers
- **UI**: HTML interfaces for user interaction
- **Utils**: Helper functions and utilities

Each module is self-contained and can be developed independently.
