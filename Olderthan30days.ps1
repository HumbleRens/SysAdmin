# Set the source and destination folders
$sourceFolder = '\\Server1\SharedFolder'
$destinationFolder = '\\Server2\ArchiveFolder'

# Get a list of files in the source folder that were last modified more than 30 days ago
$filesToMove = Get-ChildItem -Path $sourceFolder -Recurse | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) }

# Loop through the list of files and move them to the destination folder
foreach ($file in $filesToMove) {
    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $file.Name
    Move-Item -Path $file.FullName -Destination $destinationPath
}