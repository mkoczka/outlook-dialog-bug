/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

$(document).ready(() => {
  $("#run").click(run);

  $("#auth").click(auth);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
  $("#sideload-msg").hide();
  $("#app-body").show();
};

async function auth() {
  Office.context.ui.displayDialogAsync(
    location.origin + "/dialog.html",
    dialog => {
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, a =>
        console.log("This will not  get called")
      );
      dialog.addEventHandler(Office.EventType.DialogEventReceived, a =>
        console.log("This will not  get called")
      );
    }
  );
}

async function run() {
  /**
   * Insert your Outlook code here
   */
}
