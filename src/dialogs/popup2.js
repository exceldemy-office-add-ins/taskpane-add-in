(function () {
    "use strict";
      Office.onReady()
        .then(function() {
    
          // TODO1: Assign handler to the OK button.
          
          document.getElementById("okay-button").onclick = sendStringToParentPage;
    
        });
    
      // TODO2: Create the OK button handler
      function sendStringToParentPage() {
        const radioButtons= document.querySelectorAll('input[name="dateFormat"]');
        console.log(radioButtons);
        let selectedValue;
        for(const rb of radioButtons){
          if(rb.checked){
            selectedValue = rb.value;
            console.log(rb.value);
          }
          
        }
        Office.context.ui.messageParent(selectedValue);
    }

    }());