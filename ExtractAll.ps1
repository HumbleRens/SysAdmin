# Set the path to the directory containing zip files
$sourceDirectory = "C:\path\of\all\zips"

# Set the path to the directory where you want to extract the files
$destinationDirectory = "C:\Desktop\Study\extracted"

# Get a list of all zip files in the source directory
$zipFiles = Get-ChildItem -Path $sourceDirectory -Filter "*.zip" -File

# Loop through each zip file and extract its contents
foreach ($zipFile in $zipFiles) {
    $zipFileName = $zipFile.FullName
    Expand-Archive -Path $zipFileName -DestinationPath $destinationDirectory
}

Write-Host "Extraction completed."
