/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { base64Image } from "../../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
    document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl);

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

async function applyStyle() {
  await Word.run(async (context) => {
    // Queue commands to style text.
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;

    await context.sync();
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {
    // Queue commands to apply the custom style.
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";

    await context.sync();
  });
}

async function changeFont() {
  await Word.run(async (context) => {
    // Queue commands to apply a different font.
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });

    await context.sync();
  });
}

/**
 * The function is intended to insert the abbreviation ["(M365)"]
 * into the end of the Range whose text is "Microsoft 365". It makes a
 * simplifying assumption that the string is present and the user has selected it.
 */
async function insertTextIntoRange() {
  await Word.run(async (context) => {
    // 1. Queue commands to insert text into a selected range.
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (M365)", Word.InsertLocation.end);

    // 2. Load the text of the range and sync so that the current range text can be read.
    originalRange.load("text");
    await context.sync();

    // 3. Queue commands to repeat the text of the original range at the end of the document.
    doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);

    await context.sync();
  });
}

/**
 * The function is intended to add a range whose text is "Office 2019, "
 * before the range with text "Microsoft 364". It makes an assumption that
 * the string is present and the user has selected it.
 */
async function insertTextBeforeRange() {
  await Word.run(async (context) => {
    // 1. Queue commands to insert a new range before the selected range.
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", Word.InsertLocation.before);

    // 2. Load the text of the original range and sync so that the range text can be read and inserted.
    originalRange.load("text");
    await context.sync();

    // 3. Queue commands to insert the original range as a paragraph at the end of the document.
    // This new paragraph will demonstrate the fact that the new text is not part of the original selected range. The original range still has only the text it had when it was selected.
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, Word.InsertLocation.end);

    // 4. Make a final call of context.sync here and ensure that it runs after the insertParagraph has been queued.
    await context.sync();
  });
}

/**
 * The function is intended to replace the string "several"
 * with the string "many". It makes a simplifying assumption
 * that the string is present and the user has selected it.
 */
async function replaceText() {
  await Word.run(async (context) => {
    // 1. Queue commands to replace the text.
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", Word.InsertLocation.replace);

    await context.sync();
  });
}

async function insertImage() {
  await Word.run(async (context) => {
    // 1. Queue commands to insert an image.
    context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);

    await context.sync();
  });
}

async function insertHTML() {
  await Word.run(async (context) => {
    // 1. Queue commands to insert a string of HTML.
    // Adds a blank paragraph to the end of the document
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);

    // Inserts a string of HTML at the end of the paragraph;
    // specifically two paragraphs, one formatted with the Verdana font,
    // the other with the default styling of the Word document. (As you saw
    // in the insertImage method earlier, the context.document.body object also has the insert* methods.)
    blankParagraph.insertHtml(
      '<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>',
      Word.InsertLocation.end
    );

    await context.sync();
  });
}

async function insertTable() {
  await Word.run(async (context) => {
    // 1. Queue commands to get a reference to the paragraph that will precede the table.
    // Note: this line uses the ParagraphCollection.getFirst method to get a reference to the first paragraph and then uses the Paragraph.getNext method to get a reference to the second paragraph.
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

    // 2. Queue commands to create a table and populate it with data.

    /**
     * The first two parameters of the insertTable method specify the number of rows and columns.
     * The third parameter specifies where to insert the table, in this case after the paragraph.
     * The fourth parameter is a two-dimensional array that sets the values of the table cells.
     * The table will have plain default styling, but the insertTable method returns a Table object with many members, some of which are used to style the table.
     */
    const tableData = [
      ["Name", "ID", "Birth City"],
      ["Bob", "434", "Chicago"],
      ["Sue", "719", "Havana"],
    ];
    secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);

    await context.sync();
  });
}

/**
 * This code is intended to wrap the phrase "Microsoft 365" in a content control. It makes a simplifying assumption that the string is present and the user has selected it.
 */
async function createContentControl() {
  await Word.run(async (context) => {
    // 1. Queue commands to create a content control.

    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();

    // The ContentControl.title property specifies the visible title of the content control.
    serviceNameContentControl.title = "Service Name";
    // The ContentControl.tag property specifies an tag that can be used to get a reference to a content control using the ContentControlCollection.getByTag method, which you'll use in a later function.
    serviceNameContentControl.tag = "serviceName";
    // The ContentControl.appearance property specifies the visual look of the control. Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title. Other possible values are "BoundingBox" and "None".
    serviceNameContentControl.appearance = "Tags";
    // The ContentControl.color property specifies the color of the tags or the border of the bounding box.
    serviceNameContentControl.color = "blue";

    await context.sync();
  });
}

async function replaceContentInControl() {
  await Word.run(async (context) => {
    // 1. Queue commands to replace the text in the Service Name content control.
    // The ContentControlCollection.getByTag method returns a ContentControlCollection of all content controls of the specified tag. We use getFirst to get a reference to the desired control.
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    // eslint-disable-next-line no-undef
    console.error(error);
  }
}
