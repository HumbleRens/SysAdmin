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
- What does this do? Attempts to get and export all of this information into Excel individually, saves it as the name and date it was ran: 

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
- 
# EVTX to MITRE Att@ck Filter
-.PS1 Filter to scan common IoC within Windows Event logs. Orginal project: https://github.com/sbousseaden/EVTX-ATTACK-SAMPLES
- Event IDs filtered: $eventIDs = @(1151, 1116, 1149, 4699, 4698, 46, 984, 699, 7000, 7009, 4728, 29, 4756, 57, 4732, 33, 4742, 4738, 4723, 4724, 47, 284, 756, 4722, 4741, 4743, 4726, 4720, 4781, 15457, 18457, 19, 20, 21, 4622, 150, 770, 4673, 4717, 4704, 7045, 4697, 6416, 808, 354, 321, 104, 1102, 4616, 4964, 60, 5137, 5141, 5143, 5124, 5136, 5007, 4719, 4739, 4908, 2003, 4950, 2004, 4663, 325, 327, 4794, 472, 847, 564, 732, 4771, 33205, 70, 5382, 4768, 6004, 4625, 131, 4799, 4662, 600, 4769, 4661, 4825, 4648, 5140, 5145, 5142, 4674, 4656, 10, 4624, 13, 800, 4103, 4104, 4688, 11)



