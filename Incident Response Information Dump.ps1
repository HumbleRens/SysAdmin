Start-Process powershell -Verb RunAs
# Install PSExcel for Current User
Install-Module -Name ImportExcel
Install-Module -Name PSExcel
Import-Module -Name PSExcel
Get-Module -Name PSExcel

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Computer Information Dump"
$Form.Width = 700
$Form.Height = 600
$Form.StartPosition = "CenterScreen"

# Set the minimum and maximum size to prevent resizing
$Form.MinimumSize = New-Object System.Drawing.Size($Form.Width, $Form.Height)
$Form.MaximumSize = New-Object System.Drawing.Size($Form.Width, $Form.Height)

$CheckBoxAll = New-Object System.Windows.Forms.CheckBox
$CheckBoxAll.Text = "Select All"
$CheckBoxAll.AutoSize = $true
$CheckBoxAll.Location = New-Object System.Drawing.Point(40, 20)
$CheckBoxAll.Add_Click({
    foreach ($checkbox in $form.Controls | Where-Object {$_.GetType().Name -eq "CheckBox" -and $_.Name -ne "checkBoxAll"}) {
        $checkbox.Checked = $CheckBoxAll.Checked
    }
})
$form.Controls.Add($CheckBoxAll)

$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Location = New-Object System.Drawing.Point(150, 20)
$Form.Controls.Add($TextBox1)

$Button1 = New-Object System.Windows.Forms.Button
$Button1.Text = "Browse"
$Button1.Location = New-Object System.Drawing.Point(150, 50)
$Button1.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $Result = $FolderBrowser.ShowDialog()
    if ($Result -eq "OK") {
        $TextBox1.Text = $FolderBrowser.SelectedPath
    }
})
$Form.Controls.Add($Button1)
$CheckBox1 = New-Object System.Windows.Forms.CheckBox
$CheckBox1.Text = "Processes"
$CheckBox1.AutoSize = $true
$CheckBox1.Location = New-Object System.Drawing.Point(150, 110)
$Form.Controls.Add($CheckBox1)

$CheckBox2 = New-Object System.Windows.Forms.CheckBox
$CheckBox2.Text = "Services"
$CheckBox2.AutoSize = $true
$CheckBox2.Location = New-Object System.Drawing.Point(150, 130)
$Form.Controls.Add($CheckBox2)

$CheckBox3 = New-Object System.Windows.Forms.CheckBox
$CheckBox3.Text = "Scheduled Tasks"
$CheckBox3.AutoSize = $true
$CheckBox3.Location = New-Object System.Drawing.Point(150, 150)
$Form.Controls.Add($CheckBox3)

$CheckBox4 = New-Object System.Windows.Forms.CheckBox
$CheckBox4.Text = "Startup Programs"
$CheckBox4.AutoSize = $true
$CheckBox4.Location = New-Object System.Drawing.Point(150, 170)
$Form.Controls.Add($CheckBox4)

$CheckBox5 = New-Object System.Windows.Forms.CheckBox
$CheckBox5.Text = "Download Logs"
$CheckBox5.AutoSize = $true
$CheckBox5.Location = New-Object System.Drawing.Point(150, 190)
$Form.Controls.Add($CheckBox5)

$CheckBox6 = New-Object System.Windows.Forms.CheckBox
$CheckBox6.Text = "EV: Application Logs"
$CheckBox6.AutoSize = $true
$CheckBox6.Location = New-Object System.Drawing.Point(150, 210)
$Form.Controls.Add($CheckBox6)

$CheckBox7 = New-Object System.Windows.Forms.CheckBox
$CheckBox7.Text = "EV: Security Logs"
$CheckBox7.AutoSize = $true
$CheckBox7.Location = New-Object System.Drawing.Point(150, 230)
$Form.Controls.Add($CheckBox7)

$CheckBox9 = New-Object System.Windows.Forms.CheckBox
$CheckBox9.Text = "Local User Accounts"
$CheckBox9.AutoSize = $true
$CheckBox9.Location = New-Object System.Drawing.Point(150, 250)
$Form.Controls.Add($CheckBox9)

$CheckBox10 = New-Object System.Windows.Forms.CheckBox
$CheckBox10.Text = "64x Installed Software"
$CheckBox10.AutoSize = $true
$CheckBox10.Location = New-Object System.Drawing.Point(350, 110)
$Form.Controls.Add($CheckBox10)

$CheckBox11= New-Object System.Windows.Forms.CheckBox
$CheckBox11.Text = "UDP Connections"
$CheckBox11.AutoSize = $true
$CheckBox11.Location = New-Object System.Drawing.Point(350, 130)
$Form.Controls.Add($CheckBox11)

$CheckBox12= New-Object System.Windows.Forms.CheckBox
$CheckBox12.Text = "Temp File Directory"
$CheckBox12.AutoSize = $true
$CheckBox12.Location = New-Object System.Drawing.Point(350, 150)
$Form.Controls.Add($CheckBox12)

$CheckBox13= New-Object System.Windows.Forms.CheckBox
$CheckBox13.Text = "Recent USB Usage"
$CheckBox13.AutoSize = $true
$CheckBox13.Location = New-Object System.Drawing.Point(350, 170)
$Form.Controls.Add($CheckBox13)

$CheckBox14= New-Object System.Windows.Forms.CheckBox
$CheckBox14.Text = "Network Adapters"
$CheckBox14.AutoSize = $true
$CheckBox14.Location = New-Object System.Drawing.Point(350, 190)
$Form.Controls.Add($CheckBox14)

$CheckBox15= New-Object System.Windows.Forms.CheckBox
$CheckBox15.Text = "DNS Cache"
$CheckBox15.AutoSize = $true
$CheckBox15.Location = New-Object System.Drawing.Point(350, 210)
$Form.Controls.Add($CheckBox15)

$CheckBox16= New-Object System.Windows.Forms.CheckBox
$CheckBox16.Text = "Windows Firewall Rules"
$CheckBox16.AutoSize = $true
$CheckBox16.Location = New-Object System.Drawing.Point(350, 230)
$Form.Controls.Add($CheckBox16)

$CheckBox17= New-Object System.Windows.Forms.CheckBox
$CheckBox17.Text = "Registry Keys"
$CheckBox17.AutoSize = $true
$CheckBox17.Location = New-Object System.Drawing.Point(350, 250)
$Form.Controls.Add($CheckBox17)

$CheckBox18= New-Object System.Windows.Forms.CheckBox
$CheckBox18.Text = "TCP Connections"
$CheckBox18.AutoSize = $true
$CheckBox18.Location = New-Object System.Drawing.Point(350, 270)
$Form.Controls.Add($CheckBox18)

$CheckBox19= New-Object System.Windows.Forms.CheckBox
$CheckBox19.Text = "ARP Cache"
$CheckBox19.AutoSize = $true
$CheckBox19.Location = New-Object System.Drawing.Point(350, 290)
$Form.Controls.Add($CheckBox19)

$CheckBox20= New-Object System.Windows.Forms.CheckBox
$CheckBox20.Text = "Memory Dump"
$CheckBox20.AutoSize = $true
$CheckBox20.Location = New-Object System.Drawing.Point(150, 270)
$Form.Controls.Add($CheckBox20)

$CheckBox21= New-Object System.Windows.Forms.CheckBox
$CheckBox21.Text = "Hot Fix List"
$CheckBox21.AutoSize = $true
$CheckBox21.Location = New-Object System.Drawing.Point(150, 290)
$Form.Controls.Add($CheckBox21)

$CheckBox20= New-Object System.Windows.Forms.CheckBox
$CheckBox20.Text = "SMB Connections"
$CheckBox20.AutoSize = $true
$CheckBox20.Location = New-Object System.Drawing.Point(150, 310)
$Form.Controls.Add($CheckBox20)

$CheckBox21= New-Object System.Windows.Forms.CheckBox
$CheckBox21.Text = "ARP Table"
$CheckBox21.AutoSize = $true
$CheckBox21.Location = New-Object System.Drawing.Point(350, 310)
$Form.Controls.Add($CheckBox21)

$CheckBox22= New-Object System.Windows.Forms.CheckBox
$CheckBox22.Text = "System Information"
$CheckBox22.AutoSize = $true
$CheckBox22.Location = New-Object System.Drawing.Point(150, 330)
$Form.Controls.Add($CheckBox22)

$CheckBox23= New-Object System.Windows.Forms.CheckBox
$CheckBox23.Text = "Routing Table"
$CheckBox23.AutoSize = $true
$CheckBox23.Location = New-Object System.Drawing.Point(350, 330)
$Form.Controls.Add($CheckBox23)

$CheckBox24= New-Object System.Windows.Forms.CheckBox
$CheckBox24.Text = "Roaming Hashes"
$CheckBox24.AutoSize = $true
$CheckBox24.Location = New-Object System.Drawing.Point(150, 350)
$Form.Controls.Add($CheckBox24)

$CheckBox25= New-Object System.Windows.Forms.CheckBox
$CheckBox25.Text = "Windows Driver"
$CheckBox25.AutoSize = $true
$CheckBox25.Location = New-Object System.Drawing.Point(350, 350)
$Form.Controls.Add($CheckBox25)


$Button2 = New-Object System.Windows.Forms.Button
$Button2.Text = "Save"
$Button2.Location = New-Object System.Drawing.Point(250, 50)
$Button2.Add_Click({
    $OutputFolder = $TextBox1.Text

    # Create Excel COM object
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false

    if ($CheckBox1.Checked) {
        # Get Processes
        $Processes = Get-Process | Select-Object ProcessName, Id, Count | Sort Count -Descending
        $Date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $FilePath = Join-Path $OutputFolder ("Processes  " + $Date + ".xlsx")
        $Processes | Export-Excel -Path $FilePath -AutoSize -AutoFilter  
        # Get Services
        $Services = Get-Service | Select-Object Name, DisplayName, Status, StartType
        $FilePath = Join-Path $OutputFolder ("Services  " + $Date + ".xlsx")
        $Services | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Scheduled Tasks
        $ScheduledTask = Get-ScheduledTask | Select-Object TaskName, TaskPath, State, Actions
        $FilePath = Join-Path $OutputFolder ("Scheduled Tasks  " + $Date + ".xlsx")
        $ScheduledTask | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Startup Programs
        $StartupPrograms = Get-CimInstance -Class Win32_StartupCommand | Select-Object Name, Command, User
        $FilePath = Join-Path $OutputFolder ("Startup Programs  " + $Date + ".xlsx")
        $StartupPrograms | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Download Logs
        $DownloadLogs = Get-ChildItem $env:USERPROFILE\Downloads | Sort-Object LastWriteTime -Descending | Select-Object Name, LastWriteTime
        $FilePath = Join-Path $OutputFolder ("Download Logs  " + $Date + ".xlsx")
        $DownloadLogs | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Application log
        $source = "C:\Windows\System32\winevt\Logs\"
        $FilePath = Join-Path $OutputFolder ("Application Logs " + $Date + ".evtx")
        Copy-Item -Path $source"Application.evtx" -Destination $FilePath
        # Get Security log
        $source = "C:\Windows\System32\winevt\Logs\"
        $FilePath = Join-Path $OutputFolder ("Security Logs " + $Date + ".evtx")
        Copy-Item -Path $source"Security.evtx" -Destination $FilePath
        # Get Local User Accounts
        $LocalUsers = Get-LocalUser | Select-Object Name, Description, AccountType
        $Date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $FilePath = Join-Path $OutputFolder ("Local User Accounts  " + $Date + ".xlsx")
        $LocalUsers | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get x64 Installed Software
        $Wow6432Node = Get-ItemProperty -Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        $Date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $FilePath = Join-Path $OutputFolder ("x64 Installed Software  " + $Date + ".xlsx")
        $Wow6432Node | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get x32 Installed Software
        $Wow6432Node = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        $Date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $FilePath = Join-Path $OutputFolder ("x32 Installed Software  " + $Date + ".xlsx")
        $Wow6432Node | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get UDP Connections
        $connections = Get-NetUDPEndpoint
        $connections | Format-Table -AutoSize
        $FilePath = Join-Path $OutputFolder ("UDP Connections " + $Date + ".xlsx")
        $connections | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Temp File Directory
        $TempDir = Get-ChildItem -Path $env:TEMP -Recurse | Select-Object Name, LastWriteTime
        $FilePath = Join-Path $OutputFolder ("Temp Files Dump " + $Date + ".xlsx")
        $TempDir  | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Recent USB Used
        $USBStor = Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR\*\* | Select FriendlyName
        $FilePath = Join-Path $OutputFolder ("Recent USB Usage " + $Date + ".xlsx")
        $USBStor  | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get NetAdapter
        $netAdapters = Get-NetAdapter | Where-Object status -eq "up" | Select-Object Name, InterfaceIndex, InterfaceDescription, Status
        $filePath = Join-Path $OutputFolder ("Network Adapters " + $Date + ".xlsx")
        $netAdapters | Export-Excel -Path $filePath -AutoSize -AutoFilter
        # Get DNSClientCache
        $DNSClientCache = Get-DnsClientCache -Status 'Success' | Select-Object Name, Data
        $FilePath = Join-Path $OutputFolder ("DNS Cache " + $Date + ".xlsx")
        $DNSClientCache | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Windows Firewall Rules
        $FwPolicy = New-Object -ComObject HNetCfg.FwPolicy2 
        $Rules = $FwPolicy.Rules
        $FilePath = Join-Path $OutputFolder ("Windows Firewall Rules " + $Date + ".xlsx")
        $Rules | Select-Object DisplayName, Direction, Action, Protocol, LocalPorts, RemotePorts, LocalAddresses, RemoteAddresses, Enabled, Grouping | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get list of registry keys under HKLM\SOFTWARE
        $keys = Get-ChildItem HKLM:\SOFTWARE
        # Sort keys by creation time, with newest first
        $keys = $keys | Sort-Object -Property PSChildName -Descending
        $FilePath = Join-Path $OutputFolder ("Registry Keys " + $Date + ".xlsx")
        $keys | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get TCPConnections
        $connections = Get-NetTCPConnection
        $connections | Format-Table -AutoSize
        $FilePath = Join-Path $OutputFolder ("TCP Connections " + $Date + ".xlsx")
        $connections | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get ARPCache
        $arpCache = Get-NetNeighbor -AddressFamily IPv4
        $arpCache | Format-Table -AutoSize
        $FilePath = Join-Path $OutputFolder ("ARP Cache " + $Date + ".xlsx")
        $arpCache | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # MemoryDump
        # Get memory usage information for all processes
        $processes = Get-Process | Where-Object {$_.Name -ne "Idle" -and $_.Name -ne "System"} # Exclude the "Idle" and "System" processes
        $processMemoryUsage = foreach ($process in $processes) {
            [PSCustomObject]@{
                Name = $process.Name
                Id = $process.Id
                MemoryUsageMB = [math]::Round($process.WorkingSet / 1MB, 2)
            }
        }
        $FilePath = Join-Path $OutputFolder ("Memory Usage Dump " + $Date + ".xlsx")
        $processMemoryUsage | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # HotFixID
        $HotFixID = Get-HotFix | Select-Object HotFixID,Description,InstalledOn
        $HotFixID | Format-Table -AutoSize
        $FilePath = Join-Path $OutputFolder ("HotFix List " + $Date + ".xlsx")
        $HotFixID | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # GetConnections
        $GetConnections = Get-SmbConnection | Select-Object *
        $GetConnections | Format-Table -AutoSize
        $FilePath = Join-Path $OutputFolder ("SMB Connections " + $Date + ".xlsx")
        $GetConnections | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # GetARPTable
        $arpTable = Get-NetNeighbor -AddressFamily IPv4 | Select-Object IPAddress, LinkLayerAddress, State, InterfaceIndex, InterfaceAlias
        $arpTable | Format-Table -AutoSize
        $FilePath = Join-Path $OutputFolder ("ARP Table " + $Date + ".xlsx")
        $arpTable | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get system information
        $computerInfo = Get-ComputerInfo  | Select-Object WindowsBuildLabEx, WindowsCurrentVersion, WindowsVersion, WindowsEditionId, WindowsInstallationType, WindowsInstallDateFromRegistry, WindowsProductId, WindowsProductName, WindowsRegisteredOrganization, WindowsRegisteredOwner, OsProductType, OsRegisteredUser, OsSerialNumber, BiosBIOSVersion, BiosName, CsProcessors, CsNetworkAdapters, OsName, OsTyoe, OsVersion, OsBuildNumber, OsArchitecture
        $computerInfo | Format-Table -AutoSize
        $filePath = Join-Path $OutputFolder ("System Information " + $Date + ".xlsx")
        $computerInfo | Export-Excel -Path $filePath -AutoSize -AutoFilter
        # Get Roaming Hashes
        $folder = "$env:APPDATA"
        Get-ChildItem -Path $folder -File | ForEach-Object {
            $file = $_.FullName
            $md5 = (Get-FileHash -Path $file -Algorithm MD5).Hash
            $sha256 = (Get-FileHash -Path $file -Algorithm SHA256).Hash
            [PSCustomObject]@{ Name = $_.Name 
                Path = $_.FullName
                MD5 = $md5
                SHA256 = $sha256
            }
        }
        $filePath = Join-Path $OutputFolder ("Roaming Hashes " + $Date + "xlsx")
            }
        # Get Routing Table
        $routeTable = Get-NetRoute | Select-Object -Property DestinationPrefix,NextHop,RouteMetric,InterfaceAlias,AddressFamily,Type
        $FilePath = Join-Path $OutputFolder ("Routing Table " + $Date + ".xlsx")
        $routeTable | Export-Excel -Path $FilePath -AutoSize -AutoFilter
        # Get Windows Driver
        $windowsDriver = Get-WindowsDriver -Online -All | Select-Object Driver,OrginialFileName,Inbox,ClassName,BootCrtical,ProviderName,Date,Version
        # Sort the driver data by newest date first
        $sortedWindowsDriver = $windowsDriver | Sort-Object Date -Descending
        $filePath = Join-Path $OutputFolder ("Windows Drivers " + $Date + ".xlsx")
        $windowsDriver | Export-Excel -Path $filePath



        # Display a message indicating that the script has completed
        Write-Host "Script completed."
    })
$Form.Controls.Add($Button2)
$Form.ShowDialog() | Out-Null
