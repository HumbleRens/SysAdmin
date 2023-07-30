$computers = "COMPUTER1", "COMPUTER2"  # Replace with the remote computer names

foreach ($computer in $computers) {
    $creds = Get-Credential -UserName "$computer\Administrator" -Message "Enter password for $computer"
    $session = New-PSSession -ComputerName $computer -Credential $creds

    Invoke-Command -Session $session -ScriptBlock {
        cmdkey /list | ForEach-Object {
            if ($_ -like "*Target:*") {
                $target = $_.Split(":", 2)[1].Trim()
                cmdkey /delete:$target
            }
        }
    }

    Remove-PSSession $session
}