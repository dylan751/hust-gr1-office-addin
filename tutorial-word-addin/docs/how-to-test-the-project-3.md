## Test the Add-in

1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in. (See [README.md](../README.md))

2. If the add-in task pane isn't already open in Word, go to the `Home` tab and choose the `Show Taskpane` button in the ribbon to open it.

3. In the task pane, choose the Insert Paragraph button at least three times to ensure that there are a few paragraphs in the document.

4. Choose the `Insert Image` button and note that an image is inserted at the end of the document.

5. Choose the `Insert HTML` button and note that two paragraphs are inserted at the end of the document, and that the first one has the Verdana font.

6. Choose the `Insert Table` button and note that a table is inserted after the second paragraph.

![expected-output-image](../assets/how-to-test-the-project/word-tutorial-insert-image-html-table-2.png)

7. Clear all the document, proceed to the [How to test the project (Part 4)](./how-to-test-the-project-4.md)
