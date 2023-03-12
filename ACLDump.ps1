# List all folders in the specified directory and its subdirectories
$FolderPath = Get-ChildItem -Path "\\servername\location" -Directory -Recurse -Force

# Create an empty array to store the ACL information
$Report = @()

# Loop through each folder and retrieve the ACL information
foreach ($Folder in $FolderPath) {
    $Acl = Get-Acl -Path $Folder.FullName
    foreach ($Access in $Acl.Access) {
        $Properties = [ordered]@{
            'FolderName' = $Folder.FullName
            'AD Group or User' = $Access.IdentityReference
            'Permissions' = $Access.FileSystemRights
            'Inherited' = $Access.IsInherited
        }
        $Report += New-Object -TypeName PSObject -Property $Properties
    }
}

# Export the ACL report to the Excel worksheet
$Report | Export-Excel -WorksheetName 'ACL Report' -AutoSize -BoldTopRow

# Close and save the Excel workbook
Close-PowerShellExcel -Save