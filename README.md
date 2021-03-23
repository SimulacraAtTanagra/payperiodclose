## The purpose of this project is as follows:
This project takes pay period close files from SOTA systems software and filters, formats, and/or sends the data.
## Here's some back story on why I needed to build this:
The pay period close process, including all of its constitutent steps and communication needs, is laborious, time consuming, and procedural. It was the perfect candidate for automation.
## This project uses only python built-in functions and data types.

## In order to use this, you'll first need do the following:
The user will need a system for generating NPAY502 files and pay period reports. This doesn't have to be AEMS or PR-Assist but this system was built with those in mind. The user will need to either change the address of the folder for the completed (constructed) NPAY502 files (or I will need to do better programming). The user will need to update the mailing list or associated JSON (depending on when you are reading this, dear user) to reflect the appropriate addresses and message body text. Lastly, the user will need to be working in outlook as my emailautosend functions are written to access Outlook via win32com.
## The expected frequency for running this code is as follows:
Monthly