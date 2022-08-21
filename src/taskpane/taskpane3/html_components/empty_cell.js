export const empty_cell = document.createElement('template');
empty_cell.innerHTML = `
<div class="row flex justify-content-center mb-3">
<div class="col-md-12 col-sm-12 col-12">
   
    <div class="card text-center">
        <div class="card-body">
           <h5 class="card-title"style="font-weight:600">Hide Rows with Blank Cell(s)</h5> 
          <ol>
            <li>Select a range of cells from your dataset.</li>
            <li>Now, click the following button to hide the rows with blank cell(s).
            </li>
            <p> <span style="color:red ;">**Note:</span> It'll also hide rows that are completely blank.</p>
          </ol>
          <button id="emptyCell-button" class="btn btn-sm btn-danger">Click to Hide</button>
        </div>
      </div>
</div>
</div>
`;
document.getElementById('empty_cell').appendChild(empty_cell.content); 