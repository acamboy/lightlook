Outlook .pst file, containing all your sent and received e-mails together with attachments, tends to grow and grow, up to several MBs or even GBs, if you like keeping ALL of your messages for future references; this can cause Outlook slowing down a lot while processing and even displaying messages.

This can be avoided if you remove the attachment from each message, as a simple message, even in HTML format, requires a few KB on your hard disk.

Outlook allows easy saving attachments on your hard disk, but unfortunately it also keep a copy of them in the .PST file; just deleting the attachment is not the better choice, as this would cause the paperclip icon in the message to disappear, thus preventing you from determine which messages had attached file and which didn't.

This macro solves all of these issues:
it saves on your hard disk all attachments it finds into selected messages;
it keeps all attachments of the same message grouped into one single folder;
it add a dummy attachment in place of the removed one(s), thus preventing the paperclip icon from disappearing;
it adds to each "stripped" message a link to the folder containing the attchments, and links to each attachment, which thus can be opened at a click of your mouse.

![http://lightlook.googlecode.com/files/outlook-email-stripped.jpg](http://lightlook.googlecode.com/files/outlook-email-stripped.jpg)

![http://lightlook.googlecode.com/files/outlook-email-folder-list.jpg](http://lightlook.googlecode.com/files/outlook-email-folder-list.jpg)

All of these features are not an idea of mine: you have to thank Carl from http://manage-this.com/strip-outlook-attachments-vba/ ,which had this and other ideas (another cool one is the macro which prevents you from sending by mistake an email with a missing attachment: http://manage-this.com/handy-outlook-attachment-reminder-macro/ )

Carl kindly authorized me to put his opensource project on GoogleCode, to collect other programmers' suggestion to improve the macro.

Added features till now:

- double progress bar, for both current message and whole job being executed

- fixed bug of "#" character in message subject (although it is a valid character for path, outlook can't handle it)
