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
    result => {
      dialog = result.value;
      dialog.addEventHandler(
        Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
        a => console.log("This will not  get called", a)
      );
      dialog.addEventHandler(
        Microsoft.Office.WebExtension.EventType.DialogEventReceived,
        a => console.log("This will not  get called", a)
      );
    }
  );
}

async function run() {
  /**
   * Insert your Outlook code here
   */
}
