# SwissKnife PowerShell host (ADR-008). Started with -EncodedCommand, so no
# script file exists to be cleaned up or blocked by an AllSigned policy. Reads one JSON request per line on
# stdin, answers one line on stdout prefixed with a marker so module banners
# and warnings that also reach stdout can never be mistaken for a reply.
#
# Requests:
#   {"id":1,"op":"init","allow":["Get-Mailbox", ...]}
#   {"id":2,"op":"connect","family":"exo","token":"...","organization":"t","upn":""}
#   {"id":3,"op":"invoke","cmdlet":"Get-Mailbox","params":{...},"select":["Name"]}
#   {"id":4,"op":"script","script":"param($Mode) ...","params":{...}}
#
# A script (from a trusted action pack) runs only if the SHA-256 of its text
# was in the init request: the trust decision is made before the host starts.
# Replies: {"id":N,"ok":true,"data":[...]} or {"id":N,"ok":false,"error":{...}}
#
# A host either runs allow-listed cmdlets (built-in actions) or trusted pack
# scripts (its own process, started with an empty allow-list) — never both.
# Parameters are splatted from the decoded object: input values are never
# parsed as PowerShell.

# The console code page (866 on Russian Windows, 437 in the US) would garble
# every non-ASCII name in both directions: the protocol is UTF-8.
$utf8 = [Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$marker = [char]1 + 'SKG '
$allow = @{}
$scripts = @{}
$initialized = $false

function Reply($obj) {
    # The leading newline keeps the marker at a line start even after module
    # output written with -NoNewline.
    [Console]::Out.WriteLine("`n" + $marker + ($obj | ConvertTo-Json -Depth 8 -Compress -EnumsAsStrings))
    [Console]::Out.Flush()
}

function Fail($id, $err) {
    Reply @{ id = $id; ok = $false; error = @{
        message  = $err.Exception.Message
        type     = $err.Exception.GetType().Name
        category = [string]$err.CategoryInfo.Category
        errorId  = [string]$err.FullyQualifiedErrorId
    } }
}


while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    if ($line.Trim() -eq '') { continue }
    $req = $null
    try {
        $req = $line | ConvertFrom-Json -AsHashtable
        switch ($req.op) {
            'init' {
                # One-shot: the allow-list cannot be widened later.
                if ($initialized) { throw 'already initialized' }
                $initialized = $true
                foreach ($c in $req.allow) { $allow[$c] = $true }
                foreach ($h in $req.scripts) { $scripts[$h] = $true }
                Reply @{ id = $req.id; ok = $true; data = @(@{ version = $PSVersionTable.PSVersion.ToString() }) }
            }
            'connect' {
                switch ($req.family) {
                    'exo' {
                        Import-Module ExchangeOnlineManagement
                        $p = @{ AccessToken = $req.token; ShowBanner = $false; SkipLoadingFormatData = $true }
                        if ($req.upn) { $p.UserPrincipalName = $req.upn }
                        elseif ($req.delegatedOrg) { $p.DelegatedOrganization = $req.delegatedOrg }
                        else { $p.Organization = $req.organization }
                        Connect-ExchangeOnline @p | Out-Null
                        Remove-Variable p
                    }
                    'ipps' {
                        Import-Module ExchangeOnlineManagement
                        $p = @{ AccessToken = $req.token; ShowBanner = $false }
                        if ($req.upn) { $p.UserPrincipalName = $req.upn }
                        elseif ($req.delegatedOrg) { $p.DelegatedOrganization = $req.delegatedOrg }
                        else { $p.Organization = $req.organization }
                        Connect-IPPSSession @p | Out-Null
                        Remove-Variable p
                    }
                    'teams' {
                        Import-Module MicrosoftTeams
                        Connect-MicrosoftTeams -AccessTokens @($req.graphToken, $req.token) | Out-Null
                    }
                    default { throw "unknown module family: $($req.family)" }
                }
                Reply @{ id = $req.id; ok = $true; data = @() }
            }
            'invoke' {
                if (-not $allow.ContainsKey($req.cmdlet)) { throw "cmdlet not allowed: $($req.cmdlet)" }
                # Splatting binds booleans to [switch] and [bool] parameters alike.
                $splat = @{}
                if ($req.params) { $splat = $req.params }
                $out = & $req.cmdlet @splat
                if ($req.select) { $out = $out | Select-Object -Property $req.select }
                Reply @{ id = $req.id; ok = $true; data = @($out) }
            }
            'script' {
                $bytes = [Text.Encoding]::UTF8.GetBytes($req.script)
                $hash = -join ([Security.Cryptography.SHA256]::HashData($bytes) | ForEach-Object { $_.ToString('x2') })
                if (-not $scripts.ContainsKey($hash)) { throw 'script not trusted' }
                $splat = @{}
                if ($req.params) { $splat = $req.params }
                $out = & ([scriptblock]::Create($req.script)) @splat
                Reply @{ id = $req.id; ok = $true; data = @($out) }
            }
            default { throw "unknown op: $($req.op)" }
        }
    } catch {
        $id = 0
        if ($req -and $req.id) { $id = $req.id }
        Fail $id $_
    }
}
