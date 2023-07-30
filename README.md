# ACLDump
Lists all folders in the specified directory and its subdirectories, retrieves the ACL information for each folder, and exports the ACL report to an Excel worksheet 
- Rename \\servername\location to preferred path

# IPtoDNS
- Attach is both options where needing to import an list of IPs from an .xlsx and from .txt to .xlsx
 Notes:
- Be sure the top row of IP list has "ipAddress" on the .xlsx file.
- Replace the import path of your IP list.
- Replace the export path of where to save it.

# Duplicate-Cleanup
Clean up a folder containing files with duplicate files or unwanted characters in their names
- Change items: "-", "_", "()" to what ever you device
- Replace items in this line as well: $newName = $file.Name -replace "-|_|(\()|(\))|",""

# Install Module
Third-party module that provides cmdlets for importing and exporting data to and from Excel files.

# MacAddressDump
1) In the excel list set the first column as ComputerName
2) Import excel file of computer list
3) Import the Excel Module
4) Set the path to the input Excel file
5) Set the name of the output Excel file
6) Run

# Olderthan30days
Script sets a source folder and a destination folder, and then retrieves a list of files from the source folder that were last modified more than 30 days ago. Then it loops through the list of files and moves them to the destination folder.
Change source and destination folders to the appropriate locations:
- \\Server1\SharedFolder,   \\Server2\ArchiveFolder
-  change the desired number, edit the number: (-30)

# NetworkInfo
Creates a folder on Desktop named Network Dump which everything listed below is saved in .xlsx
* ARP Table
* DNS Cache
* Established TCP Connections
* Network Adapters
* Routing table
* TCP Connections

# PendingReboot
Import computer list from Excel and dump the pending reboot status exported back to Excel
1) Set the import path
2) Set the export path
3) Run

# RemotelyClearWinCreds
The script first prompts for the administrator password of each remote computer using Get-Credential. It creates a new PSSession to each remote computer and runs a ScriptBlock that uses the cmdkey command to list all stored credentials.
You will need to replace the $computers variable with the names of the remote computers you want to run this script on. 

# OSBuildInfo2023 remotely
Retrieves the OS version, build number, hotfixid installed date filtered in 2023.



