# Set the output directory for the Excel files
$outputDirectory = "$([Environment]::GetFolderPath('Desktop'))\Network Dump"
# Create the output directory if it doesn't exist
if (!(Test-Path -Path $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
# Get NetTCPConnection information 
$netTCPConnections = Get-NetTCPConnection | Select-Object -Property State,LocalAddress,LocalPort,RemoteAddress,RemotePort
$netTCPConnections | Export-Excel -Path "$outputDirectory\TCP Connections.xlsx" -AutoSize -BoldTopRow

# Get ARP table information and export to Excel
$arpTable = Get-NetNeighbor | Where-Object { $_.State -eq 'Reachable' } | Select-Object -Property IPAddress,LinkLayerAddress
$arpTable | Export-Excel -Path "$outputDirectory\ARP Table.xlsx" -AutoSize -BoldTopRow

# Get DNS cache information 
$dnsCache = Get-DnsClientCache
$dnsCache | Export-Excel -Path "$outputDirectory\DNS Cache.xlsx" -AutoSize -BoldTopRow

# Get TCP connection information
$tcpConnections = Get-NetTCPConnection -State Established | Select-Object -Property State,LocalAddress,LocalPort,RemoteAddress,RemotePort
$tcpConnections | Export-Excel -Path "$outputDirectory\Established TCP Connections.xlsx" -AutoSize -BoldTopRow

# Get the routing table
$routeTable = Get-NetRoute | Select-Object -Property DestinationPrefix,NextHop,RouteMetric,InterfaceAlias,AddressFamily,Type
$routeTable | Export-Excel -Path "$outputDirectory\Routing Table.xlsx" -AutoSize -BoldTopRow

# Get the network adapters information
$networkAdapters = Get-CimInstance -Class Win32_NetworkAdapter | Select-Object -Property Name,AdapterType,MACAddress,Manufacturer,NetConnectionID,Speed
$networkAdapters | Export-Excel -Path "$outputDirectory\Network Adapters.xlsx" -AutoSize -BoldTopRow
