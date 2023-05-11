/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function insertParagraph() {
  await Word.run(async (context) => {
    // Queue commands to insert a paragraph into the document.
    const docBody = context.document.body;
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
      Word.InsertLocation.start
    );

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
