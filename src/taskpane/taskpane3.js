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
      document.getElementById("detectBlankCell-button").onclick = detectBlankCell;
      document.getElementById("emptyCell-button").onclick = emptyCell;
      document.getElementById("entireRow-button").onclick = entireRow;
      document.getElementById("create-table").onclick = createTable;
    }
  });
  
  function createTable(){
    Excel.run(function(context){
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable= currentWorksheet.tables.add("A1:D1", true);
      expensesTable.name= "ExpensesTable";
  
      expensesTable.getHeaderRowRange().values = [["Date","Merchant", "Category", "Amount"]];
      expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company GP", "Communications", ""],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["", "", "", ""],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
      ]);
   
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();
  
      return context.sync();
    })
    .catch( function(error){
      if(error){
        console.log("Error" + error);
        if(error instanceof OfficeExtension.Error){
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      }
    });
  }
  
  async function detectBlankCell() {
    try {
      await Excel.run(async (context) => {
        var rng = context.workbook.getSelectedRange();
        rng.load(["columnCount", "rowCount", "values"]);
        await context.sync();
  
        var RowCount = rng.rowCount  
        var ColCount = rng.columnCount;
        for (var i = 0; i < RowCount; ++i) {
          for(var j = 0; j < ColCount; ++j ){
            var rcellval = rng .values[i][j]
            if(rcellval == ""){
              console.log('emtpy')
             rng.getCell(i,j).format.fill.color = 'red';
            //rng.getCell(i,j).getEntireRow().rowHidden = true;
          }
            }
        }
        await context.sync();
  
      });
    } catch (error) {
      console.error(error);
    }
  }
  
  
  async function emptyCell() {
    try {
      await Excel.run(async (context) => {
        var rng = context.workbook.getSelectedRange();
        rng.load(["columnCount", "rowCount", "values"]);
        await context.sync();
  
        var RowCount = rng.rowCount  
        var ColCount = rng.columnCount;
        for (var i = 0; i < RowCount; ++i) {
          for(var j = 0; j < ColCount; ++j ){
            var rcellval = rng .values[i][j]
            if(rcellval == ""){
              console.log('emtpy')
             rng.getCell(i,j).format.fill.color = 'red';
            rng.getCell(i,j).getEntireRow().rowHidden = true;
          }
            }
        }
        await context.sync();
  
      });
    } catch (error) {
      console.error(error);
    }
  }
  
  
  async function entireRow() {
    try {
      await Excel.run(async (context) => {
        var rng = context.workbook.getSelectedRange();
        rng.load(["columnCount", "rowCount", "values"]);
        await context.sync();
  
        var RowCount = rng.rowCount  
        var ColCount = rng.columnCount;
        for (var i = 0; i < RowCount; ++i) {
         
            var rcellval = rng .values[i][0]
            if(rcellval == ""){
              console.log('emtpy')
             rng.getCell(i,0).format.fill.color = 'red';
             var entireRow = rng.getCell(i,0).getEntireRow();
             var checkRow = context.workbook.functions.countA(entireRow); 
             checkRow.load("value");
             await context.sync();
             console.log(checkRow.value);
             if(checkRow.value == 0){
             // entireRow.format.fill.color = 'green';
              entireRow.rowHidden = true;
             }
              //rng.getCell(i,j).getEntireRow().rowHidden = true;
          
            }
        }
        await context.sync();
  
      });
    } catch (error) {
      console.error(error);
    }
  }
  
  
  
  // js for splash screen
  const splash = document.querySelector(".splash");
  document.addEventListener("DOMContentLoaded", (e)=>{
    setTimeout(()=>{
      splash.classList.add("display-none");
    },2000);
  })