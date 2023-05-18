## Test the Add-in

1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in. (See [README.md](../README.md))

2. If the add-in task pane isn't already open in Word, go to the `Home` tab and choose the `Show Taskpane` button in the ribbon to open it.

3. In the task pane, choose the Insert Paragraph button to ensure that there's a paragraph at the start of the document.

4. Within the document, select the phrase "Microsoft 365 subscription". Be careful not to include the preceding space or following comma in the selection.

5. Choose the Insert Abbreviation button. Note that " (M365)" is added. Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.

6. Within the document, select the phrase "Microsoft 365". Be careful not to include the preceding or following space in the selection.

7. Choose the Add Version Info button. Note that "Office 2019, " is inserted between "Office 2016" and "Microsoft 365". Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.

8. Within the document, select the word "several". Be careful not to include the preceding or following space in the selection.

9. Choose the Change Quantity Term button. Note that "many" replaces the selected text.

![expected-output-image](../assets/how-to-test-the-project/word-tutorial-text-replace-2.png)

10. Clear all the document, proceed to the [How to test the project (Part 3)](./how-to-test-the-project-3.md)
