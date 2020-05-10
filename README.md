Get web application running
Open command prompt and navigate to project directory
1.Run 'npm install' in root directory
Run 'npm start' to run the node web server
Deploy add-in manifest
The simplest way to deploy and test your add-in is to copy the files to a network share.

Create a folder on a network share (for example, \\MyShare\addins) and copy all the files in the Code Editor folder.

Edit the element of the manifest file so that it points to the share location from step 1.

Copy the manifest (manifest.xml) to a network share (for example, \\MyShare\addins).

Add the share location that contains the manifest as a trusted app catalog in Excel.

a. Launch Excel and open a blank spreadsheet.

b. Choose the File tab, and then choose Options.

c. Choose Trust Center, and then choose the Trust Center Settings button.

d. Choose Trusted Add-in Catalogs.

e. In the Catalog Url box, enter the path to the network share you created in step 3, and then choose Add Catalog.

f. Select the Show in Menu check box, and then choose OK. A message appears to inform you that your settings will be applied the next time you start Office.

Test add-in in Word
a.  In the **Insert tab** in Excel 2016, choose **My Add-ins**. 

b.  In the **Office Add-ins** dialog box, choose **Shared Folder**.

c.  Choose **Citations Sample**>**Insert**. The add-in opens in a task pane and shows the add-in. 
This 