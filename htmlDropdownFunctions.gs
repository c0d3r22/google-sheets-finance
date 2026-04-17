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
