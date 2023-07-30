# Load the Excel module
Import-Module -Name "Excel"

# Set the path to the input Excel file
$inputFile = "C:\path\to\output\File.xlsx"

# Set the name of the output Excel file
$outputFile = "C:\path\to\output\MacAddressDump.xlsx"

# Set the worksheet name in the input Excel file
$worksheetName = "Sheet1"

# Set the column name in the input Excel file that contains the computer names
$computerNameColumn = "ComputerName"

# Open the input Excel file
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($inputFile)
$worksheet = $workbook.Worksheets.Item($worksheetName)

# Set up an array to hold the output data
$outputData = @()

# Loop through each row in the worksheet and retrieve the adapter name, MAC address, and network connection status for each computer
for ($i = 2; $i -le $worksheet.UsedRange.Rows.Count; $i++) {
    $computerName = $worksheet.Range("$computerNameColumn$i").Value2
    
    # Retrieve the network adapter information for the computer
    $adapterInfo = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computerName | Where-Object { $_.IPEnabled -eq $true }
    
    # Loop through each adapter and retrieve the adapter name, MAC address, and network connection status
    foreach ($adapter in $adapterInfo) {
        $adapterName = $adapter.Description
        $macAddress = $adapter.MACAddress
        $netConnectionStatus = $adapter.NetConnectionStatus
        
        # Add the data to the output array
        $outputData += [PSCustomObject]@{
            "ComputerName" = $computerName
            "AdapterName" = $adapterName
            "MACAddress" = $macAddress
            "NetConnectionStatus" = $netConnectionStatus
        }
    }
}

# Close the input Excel file
$workbook.Close($false)
$excel.Quit()

# Export the output data to a new Excel file
$outputData | Export-Excel -Path $outputFile -AutoSize -WorksheetName "MacAddressDump"
