function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Team Budget Logger')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HELPER: Get spreadsheet using Script Property or fall back to Active Spreadsheet
function getSpreadsheet() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// UTILITY: Set Spreadsheet ID script property
function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  return "Spreadsheet ID set successfully to: " + id;
}

// GET: Fetch budgets, exchange rates, and user transactions
function getUserBudgetData() {
  var ss = getSpreadsheet();
  var activeUserEmail = Session.getActiveUser().getEmail();
  var targetEmail = activeUserEmail.toLowerCase().trim();

  // 1. Fetch Currency Exchange Rates from Settings sheet
  var settingsSheet = ss.getSheetByName("Settings");
  var rates = { "ZAR": 1.0, "BWP": 1.15 }; // Default fallback values
  if (settingsSheet) {
    var settingsData = settingsSheet.getDataRange().getValues();
    for (var r = 1; r < settingsData.length; r++) {
      if (settingsData[r][0]) {
        rates[settingsData[r][0].toString().toUpperCase().trim()] = parseFloat(settingsData[r][1]) || 1.0;
      }
    }
  }

  // 2. Fetch Budgets
  var budgetSheet = ss.getSheetByName("Budget_Master");
  var userBudgets = {};
  if (budgetSheet) {
    var budgetData = budgetSheet.getDataRange().getValues();
    for (var i = 1; i < budgetData.length; i++) {
      if (!budgetData[i][0]) continue;
      var rowEmail = budgetData[i][0].toString().toLowerCase().trim();
      if (rowEmail === targetEmail) {
        var category = budgetData[i][1];
        var allocated = parseFloat(budgetData[i][2]) || 0;
        var remaining = parseFloat(budgetData[i][3]) || 0;
        var spent = allocated - remaining;

        if (category) {
          userBudgets[category] = { allocated: allocated, remaining: remaining, spent: spent };
        }
      }
    }
  }

  // 3. Fetch User Transactions (Newest first, searching Expense_Log)
  // Headers: User Email | Row_ID | Date | Category | Price | Reason | Currency | Amount(ZAR)
  var expenseSheet = ss.getSheetByName("Expense_Log");
  var userTransactions = [];
  if (expenseSheet) {
    var expenseData = expenseSheet.getDataRange().getValues();
    for (var j = expenseData.length - 1; j >= 1; j--) {
      if (!expenseData[j][0]) continue;
      var expEmail = expenseData[j][0].toString().toLowerCase().trim();
      if (expEmail === targetEmail) {
        userTransactions.push({
          rowId: expenseData[j][1] ? expenseData[j][1].toString() : '',
          date: expenseData[j][2],
          category: expenseData[j][3],
          price: parseFloat(expenseData[j][4]) || 0,
          reason: expenseData[j][5] || '',
          currency: expenseData[j][6] || 'ZAR',
          priceZar: parseFloat(expenseData[j][7]) || parseFloat(expenseData[j][4]) || 0
        });
      }
    }
  }

  return {
    email: activeUserEmail,
    budgets: userBudgets,
    transactions: userTransactions,
    rates: rates
  };
}

// POST: Log a new transaction
function logExpense(category, price, reason, currency) {
  var ss = getSpreadsheet();
  var activeUserEmail = Session.getActiveUser().getEmail();
  var sheet = ss.getSheetByName("Expense_Log");
  if (!sheet) throw new Error("Sheet 'Expense_Log' not found.");
  
  var settingsSheet = ss.getSheetByName("Settings");
  var rate = 1.0;
  if (currency === 'BWP' && settingsSheet) {
    rate = parseFloat(settingsSheet.getRange("B3").getValue()) || 1.15;
  }

  var priceZar = price * rate;
  var today = new Date().toLocaleDateString("en-ZA");
  var rowId = Utilities.getUuid();
  
  // Headers: User Email | Row_ID | Date | Category | Price | Reason | Currency | Amount(ZAR)
  sheet.appendRow([activeUserEmail, rowId, today, category, price, reason, currency, priceZar]);
  return "Success";
}

// PUT: Update existing transaction by Row_ID
function updateExpense(rowId, category, price, reason, currency) {
  var ss = getSpreadsheet();
  var activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  var sheet = ss.getSheetByName("Expense_Log");
  if (!sheet) throw new Error("Sheet 'Expense_Log' not found.");

  var expenseData = sheet.getDataRange().getValues();
  var targetRowIndex = -1;
  for (var i = 1; i < expenseData.length; i++) {
    if (expenseData[i][1] && expenseData[i][1].toString() === rowId.toString()) {
      targetRowIndex = i + 1; // 1-based index
      break;
    }
  }

  if (targetRowIndex === -1) throw new Error("Transaction not found.");

  var rowEmail = sheet.getRange(targetRowIndex, 1).getValue().toString().toLowerCase().trim();
  if (rowEmail !== activeUserEmail) throw new Error("Unauthorized update request.");

  var settingsSheet = ss.getSheetByName("Settings");
  var rate = 1.0;
  if (currency === 'BWP' && settingsSheet) {
    rate = parseFloat(settingsSheet.getRange("B3").getValue()) || 1.15;
  }
  var priceZar = price * rate;

  // Update Category (Col 4), Price (Col 5), Reason (Col 6), Currency (Col 7), Amount(ZAR) (Col 8)
  sheet.getRange(targetRowIndex, 4, 1, 5).setValues([[category, price, reason, currency, priceZar]]);
  return "Success";
}

// DELETE: Delete transaction by Row_ID
function deleteExpense(rowId) {
  var ss = getSpreadsheet();
  var activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  var sheet = ss.getSheetByName("Expense_Log");
  if (!sheet) throw new Error("Sheet 'Expense_Log' not found.");

  var expenseData = sheet.getDataRange().getValues();
  var targetRowIndex = -1;
  for (var i = 1; i < expenseData.length; i++) {
    if (expenseData[i][1] && expenseData[i][1].toString() === rowId.toString()) {
      targetRowIndex = i + 1; // 1-based index
      break;
    }
  }

  if (targetRowIndex === -1) throw new Error("Transaction not found.");

  var rowEmail = sheet.getRange(targetRowIndex, 1).getValue().toString().toLowerCase().trim();
  if (rowEmail !== activeUserEmail) throw new Error("Unauthorized delete request.");

  sheet.deleteRow(targetRowIndex);
  return "Success";
}
