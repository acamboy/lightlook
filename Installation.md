It must be done in 3steps:


**Module installation**
  * download latest version of modules1.bas from donwload page
  * start outlook
  * press ALT+F11 to open "VBA editor"
  * press CTRL+R to show project panels on the left (if not already visible)
  * right click on **modules** and select **import from file**
http://lightlook.googlecode.com/files/outlook-vba-module.JPG
  * load module1.bas  (download page: http://code.google.com/p/lightlook/downloads/list )


**Form installation**
  * right click on **form** and select **import from file**
  * load frmProgress.frm file (download page: http://code.google.com/p/lightlook/downloads/list )


**Folder creation**
  * create on your hard disk a folder which will contain all your attachments, for example "C:\OutlookAttachments\2012"
  * in the VBA editor, look for **RootFolder** in the text, and change its value to the path of the folder
  * Create folder "config" inside  your "C:\OutlookAttachments\2012" folder, and copy or create in it a text file named "Attachments Removed" with no file extension.'''


Now the macro is installed on your system.
To test if it is properly working, close "VBA editor" window and get back to outlook, press ALT+F8: you should see listed macros "StripAttachments\_Explorer" and "ShowFolderName".

A final step is required: you have to **specify how outbox folder is named** in your language: look for SENT\_FOLDER\_NAME in the macro text in VBA editor, and change it to that name; to know the exact name of the folder, just select a message in the folder itself, and start the macro named **ShowFolderName**

See dedicated help page to figure out how to use the macro --> [instructions](Usage.md)



---

For advanced users:
  * if you want to create the form by yourself, select INSERT instead:
http://lightlook.googlecode.com/files/outlook-add-form.JPG
  * please remember that form must have properties shown in picture below:
http://lightlook.googlecode.com/files/outlook-form.JPG

It must contain 4 labels: two are used to simulate progress bars, two to display percentage.