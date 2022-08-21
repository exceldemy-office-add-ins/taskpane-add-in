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
      document.getElementById("detectBlankCell-button").onclick = detectBlankCell;
      document.getElementById("emptyCell-button").onclick = emptyCell;
      document.getElementById("entireRow-button").onclick = entireRow;
      document.getElementById("create-table").onclick = createTable;
    }
  });
  
  import {createTable} from './js_components/createTable'
  import {detectBlankCell} from './js_components/detectBlankCell'
  import { emptyCell } from './js_components/emptyCell';
  import { entireRow } from './js_components/entireRow';
  
  
  
  // js for splash screen
  const splash = document.querySelector(".splash");
  document.addEventListener("DOMContentLoaded", (e)=>{
    setTimeout(()=>{
      splash.classList.add("display-none");
    },2000);
  })