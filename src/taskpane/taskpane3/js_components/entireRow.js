export async function entireRow() {
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