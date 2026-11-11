<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x

About Script :
Author : Fardin Barashi
Title : psToDo tool - health check
Description : Checks the settings, the SMTP connection, the data files, when the alert and report scripts
last ran and the scheduled tasks. Changes nothing and sends nothing.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Test-Health.ps1

Version : 1.0-onprem
Release day : 2026-11-11
#>

[CmdletBinding()]
param()

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

$repoRoot       = (Resolve-Path (Join-Path $scriptRoot '..\..')).Path
$functionFolder = [IO.Path]::Combine($repoRoot, 'Settings', 'Functions', 'psToDo')
$settingsPath   = [IO.Path]::Combine($repoRoot, 'Settings', 'Config', 'ExchangeOnpremSettings.json')
$reportSettings = [IO.Path]::Combine($repoRoot, 'Settings', 'Config', 'report-settings.json')
$graphSettings  = [IO.Path]::Combine($repoRoot, 'Settings', 'Config', 'MsGraphSettings.json')
$dbPath         = [IO.Path]::Combine($repoRoot, 'Files', 'db', 'monitorobjects.json')
$statePath      = [IO.Path]::Combine($repoRoot, 'Files', 'state', 'sent-state.json')
$reportFolder   = [IO.Path]::Combine($repoRoot, 'Files', 'report')
$logFolder      = [IO.Path]::Combine($repoRoot, 'Logs')

$script:fail = 0; $script:warn = 0
function Write-Check([string] $Level, [string] $Name, [string] $Detail) {
    if ($Level -eq 'FAIL') { $script:fail++ } elseif ($Level -eq 'WARN') { $script:warn++ }
    Write-Host ("[{0}] {1} : {2}" -f $Level.PadRight(4), $Name, $Detail)
}
function Get-Age([datetime] $When) {
    $s = (Get-Date) - $When
    if ($s.TotalHours -lt 1) { return "$([int]$s.TotalMinutes) min ago" }
    if ($s.TotalHours -lt 48) { return "$([int]$s.TotalHours) h ago" }
    return "$([int]$s.TotalDays) days ago"
}

Write-Host "psToDo health check  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  on $([Environment]::MachineName) as $([Environment]::UserName)"
Write-Host "PowerShell $($PSVersionTable.PSVersion)"
Write-Host ''

Write-Host '--- Mail (on-prem Exchange) ---'
$settings = $null
try {
    foreach ($file in Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File) { . $file.FullName }
    $settings = Get-OnpremExchangeSettings -Path $settingsPath -RepoRoot $repoRoot
    Write-Check 'OK' 'Settings' "$($settings.SmtpServer):$($settings.SmtpPort), TLS=$($settings.UseTls), AuthMode=$($settings.AuthMode)"
    if (-not $settings.UseTls) { Write-Check 'WARN' 'TLS' 'UseTls is false. Mail and passwords are sent unencrypted.' }
    if ($settings.TlsAllowUntrustedCertificate) { Write-Check 'WARN' 'TLS' 'TlsAllowUntrustedCertificate is true. Use it in a lab only.' }
}
catch { Write-Check 'FAIL' 'Settings' $_.Exception.Message }

if ($settings -and $settings.AuthMode -eq 'Credential') {
    try {
        $cred = Import-Clixml -Path $settings.CredentialXmlPath
        if ($cred -is [pscredential] -and $cred.GetNetworkCredential().Password) { Write-Check 'OK' 'Credential file' "readable, user $($cred.UserName)" }
        else { Write-Check 'FAIL' 'Credential file' 'The file has no password.' }
    }
    catch { Write-Check 'FAIL' 'Credential file' "Cannot be read by $([Environment]::UserName): $($_.Exception.Message)" }
}

if ($settings) {
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $tcp.BeginConnect($settings.SmtpServer, $settings.SmtpPort, $null, $null)
        if ($async.AsyncWaitHandle.WaitOne(5000) -and $tcp.Connected) {
            $tcp.EndConnect($async)
            Write-Check 'OK' 'SMTP port' "$($settings.SmtpServer):$($settings.SmtpPort) answers"
        }
        else { Write-Check 'FAIL' 'SMTP port' "$($settings.SmtpServer):$($settings.SmtpPort) does not answer within 5 seconds" }
    }
    catch { Write-Check 'FAIL' 'SMTP port' "$($settings.SmtpServer):$($settings.SmtpPort) - $($_.Exception.Message)" }
    finally { $tcp.Close() }

    try {
        Connect-OnpremExchangeSmtp -Settings $settings | Out-Null
        Write-Check 'OK' 'SMTP connection' 'Connect-OnpremExchangeSmtp succeeded'
    }
    catch { Write-Check 'FAIL' 'SMTP connection' $_.Exception.Message }
    finally { try { Disconnect-OnpremExchangeSmtp } catch { } }
}

Write-Host ''
Write-Host '--- Data ---'
$objects = @()
try {
    $objects = @((Get-Content -Raw -Encoding UTF8 $dbPath | ConvertFrom-Json))
    Write-Check 'OK' 'monitorobjects.json' "$($objects.Count) object(s), changed $(Get-Age (Get-Item $dbPath).LastWriteTime)"
    $noMail = @($objects | Where-Object { $_.status -ne 'Completed' -and -not $_.notifyMethodbyMail })
    if ($noMail) { Write-Check 'WARN' 'Objects without mail' "$($noMail.Count): $((($noMail | Select-Object -First 8 | ForEach-Object { $_.id }) -join ', '))" }
    $noRcpt = @($objects | Where-Object { $_.notifyMethodbyMail -and -not (@($_.mail.mailRecipients) | Where-Object { $_ }) })
    if ($noRcpt) { Write-Check 'WARN' 'Objects without recipients' "$($noRcpt.Count): $((($noRcpt | Select-Object -First 8 | ForEach-Object { $_.id }) -join ', '))" }
    $dup = @($objects | Group-Object { "$($_.id)" } | Where-Object Count -gt 1)
    if ($dup) { Write-Check 'FAIL' 'Duplicate ids' (($dup | ForEach-Object Name) -join ', ') }
}
catch { Write-Check 'FAIL' 'monitorobjects.json' $_.Exception.Message }

if (Test-Path $statePath) {
    try {
        $raw = Get-Content -Raw -Encoding UTF8 $statePath
        $n = if ([string]::IsNullOrWhiteSpace($raw)) { 0 } else { @(($raw | ConvertFrom-Json).PSObject.Properties).Count }
        Write-Check 'OK' 'sent-state.json' "$n alert(s) recorded, changed $(Get-Age (Get-Item $statePath).LastWriteTime)"
    }
    catch { Write-Check 'FAIL' 'sent-state.json' $_.Exception.Message }
}
else { Write-Check 'WARN' 'sent-state.json' 'Not found. The alert script has not sent anything yet.' }

try {
    $drive = New-Object System.IO.DriveInfo([IO.Path]::GetPathRoot($repoRoot))
    $freeGb = [math]::Round($drive.AvailableFreeSpace / 1GB, 1)
    Write-Check $(if ($freeGb -lt 1) { 'WARN' } else { 'OK' }) 'Disk' "$freeGb GB free on $($drive.Name)"
}
catch { }

Write-Host ''
Write-Host '--- Last runs ---'
$staleHours = 24
try { $rs = Get-Content -Raw -Encoding UTF8 $reportSettings | ConvertFrom-Json; if ($rs.StaleAfterHours -gt 0) { $staleHours = [double]$rs.StaleAfterHours } } catch { }

$alertLog = Get-ChildItem -Path $logFolder -Filter 'psToDo-Onprem-v*.txt' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($alertLog) {
    $old = ((Get-Date) - $alertLog.LastWriteTime).TotalHours -gt 26
    Write-Check $(if ($old) { 'WARN' } else { 'OK' }) 'Alert script' "last ran $(Get-Age $alertLog.LastWriteTime) ($($alertLog.Name))"
}
else { Write-Check 'WARN' 'Alert script' 'No log found in Logs. Has psToDo-Onprem-v1.0.ps1 run?' }

foreach ($page in 'index.html', 'edit.html') {
    $p = [IO.Path]::Combine($reportFolder, $page)
    if (Test-Path $p) {
        $old = ((Get-Date) - (Get-Item $p).LastWriteTime).TotalHours -gt $staleHours
        Write-Check $(if ($old) { 'WARN' } else { 'OK' }) $page "written $(Get-Age (Get-Item $p).LastWriteTime)"
    }
    else { Write-Check 'WARN' $page 'Not found in Files\report.' }
}

Write-Host ''
Write-Host '--- Scheduled tasks ---'
if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
    try {
        $tasks = @(Get-ScheduledTask | Where-Object { ($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -match 'psToDo' })
        if (-not $tasks) { Write-Check 'WARN' 'Scheduled tasks' 'No task that runs a psToDo script was found.' }
        foreach ($t in $tasks) {
            $i = $t | Get-ScheduledTaskInfo
            $level = if ($t.State -eq 'Disabled') { 'WARN' } elseif ($i.LastTaskResult -in 0, 267009, 267011) { 'OK' } else { 'WARN' }
            $last  = if ($i.LastRunTime -and $i.LastRunTime.Year -gt 2000) { Get-Age $i.LastRunTime } else { 'never' }
            Write-Check $level $t.TaskName "state $($t.State), last run $last, result $($i.LastTaskResult), runs as $($t.Principal.UserId)"
        }
    }
    catch { Write-Check 'WARN' 'Scheduled tasks' "Could not be read: $($_.Exception.Message)" }
}
else { Write-Check 'WARN' 'Scheduled tasks' 'Get-ScheduledTask is not available on this computer.' }

Write-Host ''
Write-Host '--- Importers ---'
if ($settings -and $settings.ExchangePowerShellUri) { Write-Check 'OK' 'Exchange PowerShell' "$($settings.ExchangePowerShellUri) ($($settings.ExchangeAuthentication))" }
else { Write-Check 'WARN' 'Exchange PowerShell' 'ExchangePowerShellUri is empty. The Exchange certificate importer cannot run.' }
if (Get-Module -ListAvailable -Name ActiveDirectory) { Write-Check 'OK' 'ActiveDirectory module' 'installed' }
else { Write-Check 'WARN' 'ActiveDirectory module' 'Not installed. The AD account importer cannot run.' }
try {
    $g = Get-Content -Raw -Encoding UTF8 $graphSettings | ConvertFrom-Json
    if ("$($g.TenantId)" -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { Write-Check 'OK' 'Microsoft Graph settings' 'MsGraphSettings.json is filled in (used by the Entra and Message Center importers)' }
    else { Write-Check 'WARN' 'Microsoft Graph settings' 'MsGraphSettings.json is not filled in. The Entra and Message Center importers cannot run.' }
}
catch { Write-Check 'WARN' 'Microsoft Graph settings' 'MsGraphSettings.json not found or not readable.' }

Write-Host ''
Write-Host "Result: $($script:fail) failed, $($script:warn) warning(s)."
if ($script:fail) { exit 1 }
exit 0
