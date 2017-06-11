# WindowsUpdateHider
A script to automatically hide a set of Windows Updates, in order to prevent being installed.

Use this script carefully, as it may hide updates that you don't wish to be hidden!
Always check the text file that contains the KB ids list, to make sure you are hiding the updates you want to!

This list and the script were tested on Windows 7 x64 ONLY!


HOW TO RUN THE SCRIPT:

-save both files (hwu.vbs and HideUpdatesList.txt) in a directory on your hard-drive.
-make sure both files are in the same place and your account has enough privileges both to run .vbs files AND to alter Windows Update configuration (eg. hide updates).
-check that HideUpdatesList.txt has the KB IDs of the updates that you wish to hide listed (one KB ID per line).
ATTENTION! - Line MUST START WITH KBxxxxxxx
-double click on hwu.vbs and follow steps. If you want to cancel, just close the console window.


Use this at your own risk! I am no liable of any damage that may occur by you using this script.
