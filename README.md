# KiepLogClipboard
### A C# WPF program to download e-mail contents via clipboard

This is a tool is used together with a [Tobii Dynavox Communicator](http://www.tobiidynavox.com/) page set: [logdownload.cdd](logdownload.cdd). The problem is that Tobii Communicator can start external tools, but not pass the contents of a text box via command line options for example. 

This experiment is trying to workaround that in the following way:
 - Put the data of the Tobii Communicator page set in the Windows Clipboard
 - Start the external tool 
 - Parse the Clipboard
 - Execute the program
 - Put the results back in the Clipboard 
 - Send a keypress (that is attached to a button) back to Tobii Communicator to readout the Clipboard and put it in a text box
 
This program was designed to compile a logfile of lots of small e-mails, grouped by date. Like the [KiepEmailExport project](https://github.com/Joozt/KiepEmailExport). Together with a page set that sends an e-mail every time you clear the text input field.

Enter your Gmail credentials in [KiepMail.cs](KiepMail.cs).
