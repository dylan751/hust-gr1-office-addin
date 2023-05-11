## Prerequisites

[NodeJS](https://nodejs.org/en) (The lastest version)

The lastest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview)

Install these tool globally:

```
npm install -g yo
```

## Create the add-in project

Install generator-office:

```
npm install generator-office
```

Init the project:

```
yo office
```

Choose the templates (As guided)

- Choose a project type: `Office Add-in Task Pane project`
- Choose a script type: `JavaScript`
- What do you want to name your add-in? `tutorial-word-addin`
- Which Office client application would you like to support? `Word`
  ![Yeoman Template](https://learn.microsoft.com/en-us/office/dev/add-ins/images/yo-office-word.png)

Install office-addin-debugging:

```
npm install office-addin-debugging
```

## How to run the project

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

```
npm run dev-server
```

- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Word with your add-in loaded.

- Open another terminal -> start the localhost/3000:

  ```
  npm start
  ```

- To test your add-in in Word on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace `{url}` with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.
  ```
  npm run start:web -- --document {url}
  ```

In Word, if the "My Office Add-in" task pane isn't already open, open a new document, choose the Home tab, and then choose the Show Taskpane button in the ribbon to open the add-in task pane.officejs
<br />

![image-2](https://learn.microsoft.com/en-us/office/dev/add-ins/images/word-quickstart-addin-2b.png)

At the bottom of the task pane, choose the Run link to add the text "Hello World" to the document in blue font.
<br />
![image-3](https://learn.microsoft.com/en-us/office/dev/add-ins/images/word-quickstart-addin-1c.png)
