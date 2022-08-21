export function createTable(){
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