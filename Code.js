function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Team Budget Logger')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// GET: Fetch budgets, exchange rates, and user transactions
function getUserBudgetData() {
  var activeUserEmail = Session.getActiveUser().getEmail();
  var targetEmail = activeUserEmail.toLowerCase().trim();

  // 1. Fetch Currency Exchange Rates from Settings sheet
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
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
  var budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Budget_Master");
  var budgetData = budgetSheet.getDataRange().getValues();
  var userBudgets = {};
  
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

  // 3. Fetch User Transactions (Newest first, passing row index)
  var expenseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense_Log");
  var userTransactions = [];
  if (expenseSheet) {
    var expenseData = expenseSheet.getDataRange().getValues();
    for (var j = expenseData.length - 1; j >= 1; j--) {
      if (!expenseData[j][0]) continue;
      var expEmail = expenseData[j][0].toString().toLowerCase().trim();
      if (expEmail === targetEmail) {
        userTransactions.push({
          rowIndex: j + 1, // 1-based row index in Google Sheet
          date: expenseData[j][1],
          category: expenseData[j][2],
          price: parseFloat(expenseData[j][3]) || 0,
          reason: expenseData[j][4] || '',
          currency: expenseData[j][5] || 'ZAR',
          priceZar: parseFloat(expenseData[j][6]) || parseFloat(expenseData[j][3]) || 0
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
  var activeUserEmail = Session.getActiveUser().getEmail();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense_Log");
  if (!sheet) throw new Error("Sheet 'Expense_Log' not found.");
  
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var rate = 1.0;
  if (currency === 'BWP' && settingsSheet) {
    rate = parseFloat(settingsSheet.getRange("B3").getValue()) || 1.15;
  }

  var priceZar = price * rate;
  var today = new Date().toLocaleDateString("en-ZA");
  
  // Appends: User Email | Date | Category | Price | Reason | Currency | Price (ZAR)
  sheet.appendRow([activeUserEmail, today, category, price, reason, currency, priceZar]);
  return "Success";
}

// PUT: Update existing transaction by row index
function updateExpense(rowIndex, category, price, reason, currency) {
  var activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense_Log");
  
  var rowEmail = sheet.getRange(rowIndex, 1).getValue().toString().toLowerCase().trim();
  if (rowEmail !== activeUserEmail) throw new Error("Unauthorized update request.");

  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var rate = 1.0;
  if (currency === 'BWP' && settingsSheet) {
    rate = parseFloat(settingsSheet.getRange("B3").getValue()) || 1.15;
  }
  var priceZar = price * rate;

  sheet.getRange(rowIndex, 3, 1, 5).setValues([[category, price, reason, currency, priceZar]]);
  return "Success";
}

// DELETE: Delete transaction by row index
function deleteExpense(rowIndex) {
  var activeUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense_Log");
  
  var rowEmail = sheet.getRange(rowIndex, 1).getValue().toString().toLowerCase().trim();
  if (rowEmail !== activeUserEmail) throw new Error("Unauthorized delete request.");

  sheet.deleteRow(rowIndex);
  return "Success";
}