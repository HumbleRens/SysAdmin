Install-Module -Name ImportExcel
Get-InstalledModule
Install-Module -Name PSExcel
Import-Module -Name PSExcel
Get-Module -Name PSExcel
$env:PSModulePath.Split(';')
Install-Module -Name PSExcel -Scope CurrentUser
