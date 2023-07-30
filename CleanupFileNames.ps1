# Set the path to the folder containing the files you want to rename
$folderPath = "C:\path\to\file"

# Get all the files in the folder
$files = Get-ChildItem $folderPath

# Loop through each file
foreach ($file in $files) {
    # Remove any duplicate files by checking for files with the same name
    $duplicates = Get-ChildItem $folderPath | Where-Object { $_.Name -eq $file.Name -and $_.FullName -ne $file.FullName }
    if ($duplicates.Count -gt 0) {
        foreach ($duplicate in $duplicates) {
            Remove-Item $duplicate.FullName
        }
    }

    # Rename the file by removing "-", "_", "()"
    $newName = $file.Name -replace "-|_|(\()|(\))|",""
    Rename-Item $file.FullName -NewName $newName
}
