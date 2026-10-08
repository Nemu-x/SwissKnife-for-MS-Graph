# Read action: rows are objects; their properties become columns.
param($Mode, $Inputs, $Change)

# Unlimited: a hold report that silently stops at N mailboxes would mislead.
Get-Mailbox -ResultSize Unlimited -Filter 'LitigationHoldEnabled -eq $true' | ForEach-Object {
    [pscustomobject]@{
        name     = $_.DisplayName
        address  = [string]$_.PrimarySmtpAddress
        duration = [string]$_.LitigationHoldDuration
    }
}
