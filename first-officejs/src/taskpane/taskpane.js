/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("create-welcome-message").onclick = createWelcomeMessage;
    document.getElementById("create-table").onclick = createTable;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

export async function createWelcomeMessage() {
  return Word.run(async (context) => {
    // insert a welcome message at the end of the document.
    const welcomeMessage = context.document.body.insertHtml(
      "<h1>Welcome to first <b>Office Add-ins</b><h1>",
      Word.InsertLocation.end
    );

    // change the welcome message color to blue.
    welcomeMessage.font.color = "green";
    welcomeMessage.style = "italic";

    await context.sync();
  });
}

export async function createTable() {
  return Word.run(async (context) => {
    // insert a 3*3 table at the end of the document.
    context.document.body.insertTable(3, 3, Word.InsertLocation.end);

    await context.sync();
  });
}
