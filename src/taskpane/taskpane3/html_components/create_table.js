export const create_table = document.createElement('template');
create_table.innerHTML = `
<div class="row flex justify-content-center mb-3">
<div class="col-md-12 col-sm-12 col-12">
      <h2 class="text-center text-black" style="font-weight:600" >Hide Blank Rows </h2>
      <div class="card text-center">
          <div class="card-body">
            <h5 class="card-title"style="font-weight:600">Let's Create a Table with sample data.</h5>
            <button id="create-table" class="btn btn-sm btn-success">Create Now!</button>
          </div>
        </div>
  </div>
</div>
`;
document.getElementById('create_table').appendChild(create_table.content); 