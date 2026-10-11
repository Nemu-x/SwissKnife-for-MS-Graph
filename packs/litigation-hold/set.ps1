# Write action. Mode "plan" must not change anything: it returns what would
# change ({target, field, op, before, after, ref}). Mode "apply" receives one
# of those objects as $Change and makes it so. Inputs arrive as data
# ($Inputs.mailbox), never as code.
param($Mode, $Inputs, $Change)

switch ($Mode) {
    'plan' {
        $m = Get-Mailbox -Identity $Inputs.mailbox
        $before = if ($m.LitigationHoldEnabled) { 'on' } else { 'off' }
        $op = if ($before -eq $Inputs.state) { 'none' } else { 'set' }
        @{
            target = [string]$m.PrimarySmtpAddress
            field  = 'litigationHold'
            op     = $op
            before = $before
            after  = $Inputs.state
            # The object id when Exchange knows it, else the address.
            ref    = @{ id = $(if ($m.ExternalDirectoryObjectId) { [string]$m.ExternalDirectoryObjectId } else { [string]$m.PrimarySmtpAddress }) }
        }
    }
    'apply' {
        Set-Mailbox -Identity $Change.ref.id -LitigationHoldEnabled ($Change.after -eq 'on')
    }
}
