# Import IP addresses from text file
$ipList = Get-Content -Path "C:\insert\location\here\ips.txt"

# Create array to store results
$results = @()

# Loop through each IP address and get corresponding computer name
foreach ($ipAddress in $ipList) {
    
    Write-Host "Getting computer name for IP address $ipAddress"
    
    $computerName = (Resolve-DnsName -Name $ipAddress -Type PTR -ErrorAction SilentlyContinue | Select-Object -ExpandProperty NameHost).TrimEnd('.')
    
    if ($computerName -ne $null) {
        # Create object with IP address and computer name
        $resultObject = New-Object -TypeName PSObject -Property @{
            IPAddress = $ipAddress
            ComputerName = $computerName
        }

              # Add result object to array
        $results += $resultObject
    } else {
        Write-Warning "Unable to resolve computer name for IP address $ipAddress"
    }
}

# Export results to Excel file
$results | Export-Excel -Path "C:\insert\location\here\IPDNS.xlsx" -AutoSize -AutoFilter
