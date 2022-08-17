/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

export async function toggleProtection(args) {
  try {
    await Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('protection/protected');
      await context.sync();
  
        if (sheet.protection.protected) {
          sheet.protection.unprotect();
        } else {
          sheet.protection.protect();
        }
        await context.sync();

    });
  } catch (error) {
    console.error(error);
    if(error instanceof OfficeExtension.Error){
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
  args.completed();
}


export async function conditonalColoring(args) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      const criteria = {
        minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "blue" },
        midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percent, color: "yellow" },
        maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "red" }
      };
      conditionalFormat.colorScale.criteria = criteria;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
    if(error instanceof OfficeExtension.Error){
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
  args.completed();
}


function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
g.toggleProtection = toggleProtection;
g.conditonalColoring = conditonalColoring;