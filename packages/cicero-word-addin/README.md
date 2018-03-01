# Accord Project Cicero Word Add-in
This is a Word task pane add-in helps you identify legal text that is suitable for conversion to a Accord Project Cicero Smart Clause Template.

## Try it out
### Get web application running
1. Open command prompt and navigate to project directory
2. Run 'npm install' in root directory
3. Run 'gulp serve-static' to run the node web server

### Deploy the Add-in manifest

```
cp accord-project-cicero-template-add-in-manifest.xml /Users/<USER>/Library/Containers/com.microsoft.Word/Data/Documents/wef
```
    
### Use the Add-in in Word

Press the Insert Tab
Press the down arrow to the right of the My Add-ins button
Select the Accord Project Add-in

If you close the Task Pane it can be reopened by clicking the Home tab and then the "Show Taskpane" button.
