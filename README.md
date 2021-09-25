# WMILister

## Description and Use
**WMILister** is intended to help find and remove THREATS using WMI to persist on a computer.


**ScanAllIPsOnNetwork** can help you quickly run WMILister against other computers on your network.  ScanAll has a simple user interface to allow you to swiftly configure it and start scanning your network.  Simply download both WMILister and ScanAll to the same directory, then open a command prompt as domain admin, change directories to where you saved the utilities, then run the command:

`powershell -ExecutionPolicy Bypass -File ScanAllIPsOnNetwork.ps1`

If any odd scripts are found, you will be prompted if you want to remove them.  It is best to review the log which will be saved inside of a Log folder in the same folder the utility was run from.

Updated WMILister to v3.4.  Improved cleaning, improved error handling, Typo in output was corrected, new handling for a new WMI threat.

 

### Run this command as admin:

`cscript //nologo WMILister.vbs`


If scripts are found, you will be prompted to remove them.  The prompt will remove all scripts it finds if you tell it to.  Here is an example output for scripts found with likely compromised IP Addresses:

![WMILister_ShowingDetection](/images/WMILister_ShowingDetection.png)

Example of compromised IPs inside of log file:

![LogShowingCompromised](/images/LogShowingCompromised.png)

## Advanced use:

This version has command line switches.  Use this command to see possible switches:

`cscript //nologo WMILister.vbs /?`

These are the possible commands to scan and clean remote machines (Port 135 inbound and port 445 outbound both need to be open on remote machine.  Same open ports are seemingly used for malware to spread, so infected computers likely already have these ports open).

### Examples of switch usage are:

**Machine Name:**

`cscript //nologo WMILister.vbs /s:MachineName`

**IP Address:**

`cscript //nologo WMILister.vbs /s:10.20.30.40`

**Force Cleaning with no prompt (use at own risk as this risks removal of non malicious WMI Scripts):**

`cscript //nologo WMILister.vbs /f`

`cscript //nologo WMILister.vbs /s:MachineName /f`

`cscript //nologo WMILister.vbs /s:10.20.30.40 /f`

![WMILister_Help](/images/WMILister_Help.png)
