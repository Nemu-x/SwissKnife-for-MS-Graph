# Read action: rows are objects; their properties become columns.
param($Mode, $Inputs, $Change)

Get-Mailbox -ResultSize 1000 -Filter 'LitigationHoldEnabled -eq $true' | ForEach-Object {
    [pscustomobject]@{
        name     = $_.DisplayName
        address  = [string]$_.PrimarySmtpAddress
        duration = [string]$_.LitigationHoldDuration
    }
}
