/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
}
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("add-row").onclick = openDialog;
    document.getElementById("filter-selectCell").onclick = filterTableSelectCell;
  }
});


// js for splash screen
const splash = document.querySelector(".splash");
document.addEventListener("DOMContentLoaded", (e)=>{
  setTimeout(()=>{
    splash.classList.add("display-none");
  },2000);
})

export async function createTable() {
  try {
    await Excel.run(async (context) => {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable= currentWorksheet.tables.add("A1:D1", true);
      expensesTable.name= "ExpensesTable";
  
      expensesTable.getHeaderRowRange().values = [["Date","Merchant", "Category", "Amount"]];
      expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company GP", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
      ]);
   
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();
 
      await context.sync();
    });
  } catch (error) {
    console.error(error);
    if(error instanceof OfficeExtension.Error){
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}

// Start: Add new Row
// addRow function takes 4 arguments to add a new row to the existing table.
// we'll use a  dialog box to get these arguments as user inserted data. 

async function addRow(date, merchant, category, amount) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.getItem("ExpensesTable");
  
      expensesTable.rows.add(
          null, // index, Adds rows to the end of the table.
          [
              [date, merchant,category, amount]
          ], 
          true, // alwaysInsert, Specifies that the new rows be inserted into the table.
      );
  
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
 
      await context.sync();
    });
  } catch (error) {
    console.error(error);
    if(error instanceof OfficeExtension.Error){
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}

// openDialog() function will open a dialog box to take 4 user input data 
//and then the data gets stored into 4 diffrent variales 
// finally it'll call the addRow() function  
function openDialog() {
  // TODO1: Call the Office Common API that opens a dialog
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    {height: 45, width: 55},
  
    // TODO2: Add callback parameter.
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
  );
}

var date, merchant, category, price; 
function processMessage(arg) {
  const messageFromDialog = JSON.parse(arg.message);
  date= messageFromDialog.a;
  merchant = messageFromDialog.b;
  category = messageFromDialog.c;
  price = messageFromDialog.d;
  addRow(date, merchant, category,price);

  dialog.close();
 
}
var dialog = null;

// End : Add new Row


//Start: Select and Filter the Table
async function filterTableSelectCell() {
  try {
    await Excel.run(async (context) => {
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      let range = context.workbook.getSelectedRange();
      range.load(["text"]);
      await context.sync();

      var filterVal = range.text;
      var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
      var categoryFilter = expensesTable.columns.getItem('Category').filter;
      categoryFilter.applyValuesFilter([`${filterVal}`]);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
    if(error instanceof OfficeExtension.Error){
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}
//End: Select and Filter the Table