# Import computer list from Excel file
$computers = Import-Excel -Path "C:\path\to\file.xlsx" | Select-Object -ExpandProperty A

# Check pending reboot status for each computer
$results = foreach ($computer in $computers) {
    $pendingReboot = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computer).RebootPending
    [PSCustomObject]@{
        ComputerName = $computer
        PendingReboot = $pendingReboot
    }
}

# Export results to Excel file
$results | Export-Excel -Path "C:\path\to\results.xlsx" -AutoSize -AutoFilter -WorksheetName "PendingRebootList"
