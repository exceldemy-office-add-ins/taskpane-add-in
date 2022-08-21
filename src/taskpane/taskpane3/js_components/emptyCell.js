export async function emptyCell() {
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