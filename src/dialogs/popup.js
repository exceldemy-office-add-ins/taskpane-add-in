(function () {
    "use strict";
      Office.onReady()
        .then(function() {
    
          // TODO1: Assign handler to the OK button.
          document.getElementById("ok-button").onclick = sendStringToParentPage;
    
        });
    
      // TODO2: Create the OK button handler
      function sendStringToParentPage() {
        var date = document.getElementById("date").value;
        var merchant = document.getElementById("merchant").value;
        var category = document.getElementById("category").value;
        var price = document.getElementById("price").value;
        var data = {
           a : date,
          b : merchant,
          c : category,
          d : price
        }
        Office.context.ui.messageParent(JSON.stringify(data));

    }

    }());