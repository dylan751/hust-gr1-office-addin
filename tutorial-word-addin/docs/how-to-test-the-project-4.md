## Test the Add-in

1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in. (See [README.md](../README.md))

2. If the add-in task pane isn't already open in Word, go to the `Home` tab and choose the `Show Taskpane` button in the ribbon to open it.

3. In the task pane, choose the Insert Paragraph button to ensure that there's a paragraph with "Microsoft 365" at the top of the document.

4. In the document, select the text "Microsoft 365" and then choose the `Create Content Control` button. Note that the phrase is wrapped in tags labelled "Service Name".

5. Choose the `Rename Service` button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".

![expected-output-image](../assets/how-to-test-the-project/word-tutorial-content-control-2.png)

6. Clear all the document.
