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
      document.getElementById("okay-button").onclick = changeDateFormat;

    }
  });
  
  export async function changeDateFormat() {
    try {
      await Excel.run(async (context) => {
        const radioButtons= document.querySelectorAll('input[name="dateFormat"]');
        console.log(radioButtons);
        let selectedValue;
        for(const rb of radioButtons){
          if(rb.checked){
            selectedValue = rb.value;
            console.log(rb.value);
          }
        }
        let rng = context.workbook.getSelectedRange();
          rng.numberFormat = selectedValue;
          rng.format.autofitColumns();
          rng.format.autofitRows();
        await context.sync();
      });
    } catch (error) {
      console.error(error);
      if(error instanceof OfficeExtension.Error){
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    }
  }

  
