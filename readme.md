# Office Licensing Fixer
Easy and simple GUI to clear Office 365 accounts from Office to enable a new account to be entered. This script uses the Office Software Protection Platform script found in Program Files\Microsoft Office\Office16 folder. 

This program is designed to be used when Office is un-licensed with someone else's account and will not allow you to enter in another account. It will find all the keys/accounts and remove them from Office in one click, allowing you to re-enter your office 365 details to license the program.


## How to use the office license fixer
1. Download the .exe file from [here](https://github.com/danielcummins/Office-licensing-fixer/blob/master/Office%20licensing%20fixer/Office%20licensing%20fixer/bin/Debug/Office%20licensing%20fixer.exe)
2. Run the .exe file as admin on the target machine
3. Click the "Fix office license" button to fix
3. Close all office apps then reopen to check the license!


## How it actually works

This program automates the following commands from the office ospp.vbs file:

Firstly this command to retrive all the keys
`cscript ospp.vbs /dstatus`

Then for each of the keys in that response this command is run
`cscript ospp.vbs /unpkey:<key from command>`
