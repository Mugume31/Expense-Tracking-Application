const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Add Expense/Revenue')
    .addItem('Open Form', 'showForm')
    .addToUi();
}

// Function to show the HTML form as a sidebar
function showForm() {
  var html = HtmlService.createTemplateFromFile('Form')
    .evaluate()  
    .setTitle("Add inventory")
    .setWidth(2000)
    .setHeight(700);  
  SpreadsheetApp.getUi().showModalDialog(html,"Add Expense/Revenue")
}

function doGet() {
  return  HtmlService.createTemplateFromFile('Form')
    .evaluate()  
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function include(filename) {
  var content = HtmlService.createHtmlOutputFromFile(filename).getContent();
  return content
}
// Function to get a list of unique departments from the 'Particulars' sheet
function getDepartments() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Particulars");
  if (!sheet) {
    Logger.log("Sheet 'Particulars' not found.");
    return [];
  }

  var data = sheet.getDataRange().getValues(); // Get all rows including headers
  var headerRow = data[0];
  var deptColIndex = headerRow.indexOf("Department");

  if (deptColIndex === -1) return []; // No 'Department' column found

  var departments = data.slice(1).map(row => row[deptColIndex]); // Get all values in the Department column
  var uniqueDepartments = [...new Set(departments)].filter(String); // Remove duplicates and blanks

  return uniqueDepartments;
}

// Function to retrieve particulars based on Type and Department
function getParticulars(type, department) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Particulars');
  if (!sheet) {
    Logger.log("Sheet 'Particulars' not found.");
    return [];
  }
  var data = sheet.getDataRange().getValues();

  var particulars = data.filter(function(row) {
    return row[0] === type && row[1] === department;  // Match Type and Department
  }).map(function(row) {
    return row[2];  // Return the "Particular" column
  });

  var uniqueParticulars = Array.from(new Set(particulars)); 
  uniqueParticulars.sort(); 
  return uniqueParticulars;
}

// Function to submit the form data to the "Transactions" sheet
function submitEntries(entries) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Transactions');
  if (!sheet) {
    Logger.log("Sheet 'Transactions' not found.");
    return;
  }

  entries.forEach(function(entry) {
    if (!entry.date || !entry.type || !entry.department || !entry.particular || entry.unitPrice <= 0 || entry.quantity <= 0) {
      Logger.log("Skipping entry due to missing or invalid fields: " + JSON.stringify(entry));
      return;
    }

    // Calculate amount (Unit Price * Quantity)
    entry.amount = entry.unitPrice * entry.quantity;

    // Append the entry data to the sheet
    sheet.appendRow([
      entry.date,          // Date
      entry.type,          // Type (Cost/Income)
      entry.department,    // Department (Bar, Kitchen, etc.)
      entry.particular,    // Particular
      entry.unitPrice,     // Unit Price
      entry.quantity,      // Quantity
      entry.amount,        // Amount (Unit Price * Quantity)
      entry.comment || ''  // Comment (default empty if null)
    ]);
  });
}




