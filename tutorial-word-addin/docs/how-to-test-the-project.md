## How to test this project

1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in. (See [README.md](../README.md))

2. If the add-in task pane isn't already open in Word, go to the `Home` tab and choose the `Show Taskpane` button in the ribbon to open it.

3. Be sure there are at least three paragraphs in the document. You can choose the `Insert Paragraph` button three times. Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.

4. In Word, create a custom style named "MyCustomStyle". It can have any formatting that you want.

5. Choose the `Apply Style` button. The first paragraph will be styled with the built-in style `Intense Reference`.

6. Choose the `Apply Custom Style` button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)

7. Choose the `Change Font` button. The font of the second paragraph changes to 18 pt., bold, Courier New.
