export const detect_blank_cell = document.createElement('template');
detect_blank_cell.innerHTML = `
<div class="row flex justify-content-center mb-3">
<div class="col-md-12 col-sm-12 col-12">
   
    <div class="card text-center">
        <div class="card-body">
           <h5 class="card-title"style="font-weight:600">Detect Blank Cell(s) </h5> 
          <ol>
            <li>Select a range of cells from your dataset.</li>
            <li>Then click the following button to highlight the black cells in the selected range.
            </li>
            <p> <span style="color:red ;">**Note:</span> It'll also highlight rows that are completely blank.</p>
          </ol>
          <button id="detectBlankCell-button" class="btn btn-sm btn-warning">Click to highlight</button>
        </div>
      </div>
</div>
</div>
`;
document.getElementById('detect_blank_cell').appendChild(detect_blank_cell.content);