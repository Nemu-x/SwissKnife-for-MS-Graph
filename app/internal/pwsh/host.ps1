# SwissKnife PowerShell host (ADR-008). A short -EncodedCommand bootstrap
# reads this script from stdin, so no script file exists to be cleaned up or
# blocked by an AllSigned policy. Reads one JSON request per line on
# stdin, answers one line on stdout prefixed with a marker so module banners
# and warnings that also reach stdout can never be mistaken for a reply.
#
# Requests:
#   {"id":1,"op":"init","allow":["Get-Mailbox", ...],"scripts":[{"hash":"…","cmdlets":["Get-Mailbox"]}]}
#   {"id":2,"op":"connect","family":"exo","token":"...","organization":"t","upn":""}
#   {"id":3,"op":"invoke","cmdlet":"Get-Mailbox","params":{...},"select":["Name"]}
#   {"id":4,"op":"script","script":"param($Mode) ...","params":{...}}
#
# A script (from a trusted action pack) runs only if the SHA-256 of its text
# was in the init request: the trust decision is made before the host starts.
# It may call only the commands its pack declared for it (plus a few that
# reach neither the network nor the disk), named literally, and it runs in
# ConstrainedLanguage: no .NET calls, no Add-Type, no dynamic invocation. The
# module cmdlets it calls keep running as their modules do.
# Replies: {"id":N,"ok":true,"data":[...]} or {"id":N,"ok":false,"error":{...}}
#
# A host either runs allow-listed cmdlets (built-in actions) or trusted pack
# scripts (its own process, started with an empty allow-list) — never both.
# Parameters are splatted from the decoded object: input values are never
# parsed as PowerShell.

# The console is already UTF-8: the bootstrap (host.go) set it before it read
# this script from stdin — setting it again would drop buffered input.

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
# Everything below runs in a private scope, not the global one: a pack
# script (bound to an empty module, see 'script') resolves names through that
# module and the global scope only, so the host's state is out of its reach.
& {
$marker = [char]1 + 'SKG '
$allow = @{}
$scripts = @{}   # hash → { command name → $true } the script may call

# Commands every pack script may use: they touch neither the network nor the
# disk. Never: commands that run code, start processes or jobs, or load modules,
# whatever a pack declares (the app refuses such packs before this).
$packSafe = @{}
foreach ($c in 'Where-Object', 'ForEach-Object', 'Select-Object', 'Sort-Object', 'Group-Object', 'Measure-Object',
    'Compare-Object', 'Write-Output', 'Write-Verbose', 'Write-Warning', 'Write-Error', 'Out-Null', 'Out-String',
    'ConvertTo-Json', 'ConvertFrom-Json', 'Get-Date', 'New-TimeSpan', 'Join-String',
    '?', '%', 'where', 'foreach', 'select', 'sort', 'group', 'measure', 'echo') { $packSafe[$c] = $true }
$packNever = @{}
foreach ($c in 'Invoke-Expression', 'iex', 'Invoke-Command', 'icm', 'Add-Type', 'Start-Process', 'saps', 'start',
    'Start-Job', 'sajb', 'Start-ThreadJob', 'Receive-Job', 'Import-Module', 'ipmo', 'New-Module', 'Set-Alias', 'New-Alias',
    'Register-ObjectEvent', 'Register-EngineEvent', 'Set-ExecutionPolicy', 'New-PSSession', 'Enter-PSSession',
    'Invoke-Item', 'Set-Variable', 'Get-Variable', 'Remove-Variable', 'Clear-Variable', 'Get-Command') { $packNever[$c] = $true }

$langMode = [scriptblock].GetProperty('LanguageMode', [Reflection.BindingFlags]'NonPublic,Instance')
$sandbox = New-Module -Name SkgPackSandbox -ScriptBlock { }

# Test-PackScript refuses a script that calls anything it did not declare,
# calls a command by a computed name, or reaches the engine behind the
# language mode. The script is parsed, never run, to check it.
function Test-PackScript([string]$text, [hashtable]$declared) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw "the pack script does not parse: $($errors[0].Message)" }
    $defined = @{}
    foreach ($f in $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        # A scoped name (global:...) would outlive the run or shadow the host.
        if ($f.Name.Contains(':')) { throw "pack scripts cannot define $($f.Name)" }
        $defined[$f.Name] = $true
    }
    foreach ($n in $ast.FindAll({ $args[0] -is [Management.Automation.Language.UsingStatementAst] -or
                $args[0] -is [Management.Automation.Language.ConfigurationDefinitionAst] -or
                $args[0] -is [Management.Automation.Language.DynamicKeywordStatementAst] }, $true)) {
        throw "pack scripts cannot use 'using', configurations or dynamic keywords"
    }
    foreach ($c in $ast.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] }, $true)) {
        $name = $c.GetCommandName()
        if (-not $name) { throw "pack scripts must name the commands they call (line $($c.Extent.StartLineNumber))" }
        if ($c.InvocationOperator -eq 'Dot') { throw "pack scripts cannot dot-source (line $($c.Extent.StartLineNumber))" }
        if ($packNever.ContainsKey($name)) { throw "pack scripts cannot call $name" }
        if (-not ($declared.ContainsKey($name) -or $packSafe.ContainsKey($name) -or $defined.ContainsKey($name))) {
            throw "the pack did not declare the command $name"
        }
        # ForEach-Object by member name invokes methods past the language
        # mode: only script blocks.
        if ($name -in 'ForEach-Object', '%', 'foreach') {
            $args0 = @($c.CommandElements | Select-Object -Skip 1)
            for ($i = 0; $i -lt $args0.Count; $i++) {
                $e = $args0[$i]
                if ($e -is [Management.Automation.Language.CommandParameterAst]) {
                    $pn = $e.ParameterName.ToLower()
                    if ($pn.StartsWith('m') -or $pn.StartsWith('ar')) { throw "pack scripts cannot use ForEach-Object -$($e.ParameterName)" }
                    # The value after -InputObject or a common parameter is data.
                    if ($null -eq $e.Argument -and $pn -in 'inputobject', 'erroraction', 'warningaction', 'informationaction',
                        'errorvariable', 'warningvariable', 'informationvariable', 'outvariable', 'outbuffer', 'pipelinevariable') { $i++ }
                    continue
                }
                if ($e -isnot [Management.Automation.Language.ScriptBlockExpressionAst]) {
                    throw "pack scripts give ForEach-Object a script block (line $($e.Extent.StartLineNumber))"
                }
            }
        }
        foreach ($e in $c.CommandElements) {
            # -Parallel and -AsJob run the block outside this language mode.
            if ($e -is [Management.Automation.Language.CommandParameterAst] -and
                ('parallel'.StartsWith($e.ParameterName.ToLower()) -and $e.ParameterName.Length -ge 2 -or 'asjob' -eq $e.ParameterName.ToLower())) {
                throw "pack scripts cannot use -$($e.ParameterName)"
            }
        }
    }
    foreach ($m in $ast.FindAll({ $args[0] -is [Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
        if ($m.Member -isnot [Management.Automation.Language.StringConstantExpressionAst]) {
            throw "pack scripts must name the methods they call (line $($m.Extent.StartLineNumber))"
        }
        # .ForEach('Name') / .Where('Name') invoke members past the language mode.
        if ($m.Member.Value -in 'ForEach', 'Where' -and ($m.Arguments.Count -eq 0 -or $m.Arguments[0] -isnot [Management.Automation.Language.ScriptBlockExpressionAst])) {
            throw "pack scripts give .$($m.Member.Value)() a script block (line $($m.Extent.StartLineNumber))"
        }
    }
    foreach ($v in $ast.FindAll({ $args[0] -is [Management.Automation.Language.VariableExpressionAst] }, $true)) {
        $path = $v.VariablePath
        if ($path.IsDriveQualified -and $path.DriveName -ne 'env') { throw "pack scripts cannot read `$$($path.UserPath)" }
        # Global and script variables would outlive the run.
        if ($path.IsGlobal -or $path.IsScript) { throw "pack scripts cannot use `$$($path.UserPath)" }
        $bare = $path.UserPath -replace '^(?i)(global|local|script|private|using):', ''
        if ($bare -in 'ExecutionContext', 'Host') { throw "pack scripts cannot use `$$bare" }
    }
}
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
                foreach ($sc in $req.scripts) {
                    $set = @{}
                    foreach ($c in $sc.cmdlets) { $set[$c] = $true }
                    $scripts[$sc.hash] = $set
                }
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
                Test-PackScript $req.script $scripts[$hash]
                if (-not $langMode) { throw 'this PowerShell cannot run pack scripts in ConstrainedLanguage' }
                $sb = $sandbox.NewBoundScriptBlock([scriptblock]::Create($req.script))
                $langMode.SetValue($sb, [Management.Automation.PSLanguageMode]::ConstrainedLanguage)
                if ($langMode.GetValue($sb) -ne [Management.Automation.PSLanguageMode]::ConstrainedLanguage) {
                    throw 'this PowerShell cannot run pack scripts in ConstrainedLanguage'
                }
                $splat = @{}
                if ($req.params) { $splat = $req.params }
                $out = & $sb @splat
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
} # private scope
