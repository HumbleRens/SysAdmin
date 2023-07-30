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

# OSBuildInfo2023 
Remotely retrieves the OS version, build number, hotfixid installed date filtered in 2023.

# Incident Response Information Dump
The purpose of this script is to start you in the right direction to assist in detecting possible malicious activity on a local Windows 10/11 computer with quickly grabbing and dumping information to be observed. Some small-medium organizations do not have appropriate SIEM/SOAR/IR/Tools/Training in place.

I Highly recommend running THOR Lite + LOKI first for IoC

Notes:
- Tested on Windows 10 and Windows 11
- Must have PSExcel module installed
- Must run as administrator to successfully get everything, if certain things do not exist, the .xlsx file will be empty

What does this do? 
Attempts to get and export all of this information into Excel individually, saves it as the name and date it was ran: 

-	Process: Retrieves information about running processes on the local computer.
-	Services: Retrieves information about services on the local computer.
-	Scheduled Tasks: Retrieves information about scheduled tasks on the local computer.
-	Startup Programs: Retrieves information about programs that are set to run at startup on the local computer.
-	Download Folder logs: Retrieves list of files in the current user's Downloads folder.
- Copies Windows Event Application/Security Logs.
-	Local User Account: Retrieves information about local user accounts on the local computer.
-	Installed Software x64: Retrieves information about installed software on the local computer (64-bit) from the “Uninstall” registry key. 
- Installed Software x32: Retrieves information about installed software on the local computer (32-bit) from the “Uninstall” registry key.
-	Temp File Directory Listing: Retrieves a list of files in the current user's temporary directory.
-	Recent USB Usage: Retrieves information about recent USB usage on the local computer.
-	NetAdapter: Retrieves information about network adapters on the local computer. 
-	DNS Cache: Retrieves the contents of the local DNS cache.
-	Windows Firewall Rules: Retrieves information about Windows Firewall rules on the local computer.
- Registry Keys: Retrieves all child items of the HKLM:\SOFTWARE key, which includes all the subkeys and values within the SOFTWARE hive.
-	TCP Connections: Retrieves information about current TCP connections on the local computer.
-	ARP Cache: Retrieves the contents of the ARP cache on the local computer.
-	Memory Dump: Retrieves a memory dump of the lsass.exe process on the local computer.
-	Hotfix: Retrieves a list of installed Microsoft hotfixes on the local computer.
-	ARP Table: ARP table entries for the IPv4 address family, selects the IPAddress, LinkLayerAddress, State, and InterfaceIndex properties. 
-	System Information: Get general system information such as name, primary owner name, domain, model, manufacturer.
-	SmbConnection: Get information about SMB connections.
-	NetRoute: Get information about the routing table.
- UDP Endpoint Connections: Network protocol that is commonly used for real time applications such as video conferencing, streaming and gaming.
- WindowsDriver: List all currently installed Windows drivers.
- Roaming Hashes: Receives MD5 and SHA256 file hashes for files in the Roaming folder.




