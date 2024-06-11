function doItAll() {
  // get ID of debts list
  const debtSheetID = getCurrentDebtsSheetID();
  
  // make a sheet object containing just the sheet we're interested in
  const currentDebtsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[debtSheetID];
  
  // find the column that has the dates
  const debtsDateCol = findColumnByHeader(currentDebtsSheet, "due date w/ year")
  Logger.log("debt date column number is " + debtsDateCol);
  
  // find the column that has the expense name
  const expenseCol = findColumnByHeader(currentDebtsSheet, "expense")
  Logger.log("debt name column number is " + expenseCol);
  
  // find the colum that has the cost
  const costCol = findColumnByHeader(currentDebtsSheet, "monthly bill")
  Logger.log("debt value column number is " + costCol);

  // reset calendar
  financeCal = getFinanceCalendar();
  clearAllEvents(financeCal);

  // for each row, get bill date
  // Get the data range (excluding header row)
  var dataRange = currentDebtsSheet.getDataRange();
  var data = dataRange.getValues().slice(1); // Skip the first row (header)

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateValue = row[debtsDateCol - 1];
    var expense = row[expenseCol - 1];
    var currencyValue = row[costCol - 1];

    // Validate date
    if (!isValidDate(dateValue)) {
      Logger.log("Row " + (i + 2) + ", Column " + debtsDateCol + ": Invalid Date");
      continue; // Skip to next row
    }

    // Validate value in column
    if (expense === "") {
      Logger.log("Row " + (i + 2) + ", Column " + expenseColCol + ": Missing Value");
      continue;
    }

    // Validate currency
    if (!isValidCurrency(currencyValue)) {
      Logger.log("Row " + (i + 2) + ", Column " + costColCol + ": Invalid Currency");
      continue;
    }

    // All validations pass, call createBillEvent
    createBillEvent(financeCal, expense, dateValue, currencyValue);
  }
}

function getFinanceCalendar() {
  const calendars = CalendarApp.getCalendarsByName('Finance');
  const financeCalendar = calendars[0];
  return financeCalendar;
}

function clearAllEvents(calendar) {
  // Set start and end dates (one year ago and year 2100)
  var oneYearAgo = new Date();
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
  var year2100 = new Date();
  year2100.setFullYear(2100);

  // Get events within the specified time range
  var events = calendar.getEvents(oneYearAgo, year2100);

  // Delete each event
  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
}

function createBillEvent(calendar, billName, billDate, billAmount) {
  // Set start and end date/time (9:00 AM and 1 hour later)
  var startDate = new Date();
  startDate.setFullYear(billDate.getFullYear(), billDate.getMonth(), billDate.getDate());
  startDate.setHours(9, 0, 0, 0); // Set time to 9 AM
  var endDate = new Date(startDate.getTime() + 60 * 60 * 1000);

  // Create a new event object with start and end times
  var event = calendar.createEvent(
      billName, // Event title
      startDate, // Start date
      endDate, // End date
      {
        description: "Bill amount: $" + billAmount, // Add bill amount to details
      }
  );
  
  Logger.log("Bill event created for: " + billName + ", on: " + billDate + ", amount: $" + billAmount);
}


function getCurrentDebtsSheetID() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let sheetNames = [];
  let debtID = '';
  sheets.forEach(function (sheet, i) {
    if (sheet.getName() == 'current debts') {
        debtID = i;
    }
  });
  return debtID;
}

function findColumnByHeader(sheet, header) {
  // Get all values in the header row.
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the index of the header value.
  const columnIndex = headerRow.indexOf(header);
  
  // Check if header is found, return -1 otherwise.
  return columnIndex === -1 ? -1 : columnIndex + 1;
}

// Function to check if a value is a valid date
function isValidDate(value) {
  return value !== "" && !isNaN(new Date(value));
}

// Function to check if a value is a valid currency (basic check)
function isValidCurrency(value) {
  // You can customize this function for more complex currency validation
  return !isNaN(parseFloat(value));
}
