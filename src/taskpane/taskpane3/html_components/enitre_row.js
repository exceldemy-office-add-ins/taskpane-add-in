export const entire_row = document.createElement('template');
entire_row.innerHTML = `
<div class="row flex justify-content-center mb-3">
<div class="col-md-12 col-sm-12 col-12">
    <div class="card text-center">
        <div class="card-body">
          <h5 class="card-title"style="font-weight:600">Hide Entirely Blank Rows </h5>
          <ol>
            <li>Select a range of cells from your dataset.</li>
            <li>Now, click the following button to hide the rows that are entirely blank.
            </li>
          </ol>
          <button id="entireRow-button" class="btn btn-sm btn-success">Hide Blank Rows</button>
        </div>
      </div>
</div>
</div>
`;
document.getElementById('entire_row').appendChild(entire_row.content); 