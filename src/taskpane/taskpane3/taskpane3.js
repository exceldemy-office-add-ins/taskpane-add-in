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
  //Js Components Imports
  import {createTable} from './js_components/createTable'
  import {detectBlankCell} from './js_components/detectBlankCell'
  import { emptyCell } from './js_components/emptyCell';
  import { entireRow } from './js_components/entireRow';

//HTML Comoponents Imports
import { create_table } from './html_components/create_table';
import { detect_blank_cell } from './html_components/detect_blank_cell';
import { empty_cell } from './html_components/empty_cell';
import { entire_row } from './html_components/enitre_row';
  
  
  
// js for splash screen
  const splash = document.querySelector(".splash");
  document.addEventListener("DOMContentLoaded", (e)=>{
    setTimeout(()=>{
      splash.classList.add("display-none");
    },2000);
  })


// Sample Nav Bar 

  $('.btn_menu').click(function(){
    $(this).toggleClass("click");
    $('.sidebar').toggleClass("show");
  });
    $('.feat-btn').click(function(){
      $('nav ul .feat-show').toggleClass("show");
      $('nav ul .first').toggleClass("rotate");
    });
    $('.serv-btn').click(function(){
      $('nav ul .serv-show').toggleClass("show1");
      $('nav ul .second').toggleClass("rotate");
    });
    $('nav ul li').click(function(){
      $(this).addClass("active").siblings().removeClass("active");
    });

