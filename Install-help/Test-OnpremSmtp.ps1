<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x

About Script :
Author : Fardin Barashi
Title : psToDo setup - SMTP test
Description : Sends one test mail with the settings in ExchangeOnpremSettings.json.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Test-OnpremSmtp.ps1 -To you@lab.local
    .\Test-OnpremSmtp.ps1 -To you@lab.local -From pstodo-noreply@lab.local

Version : 1.0-onprem
Release day : 2026-09-08
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $To,
    [string] $From
)

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

$ErrorActionPreference = 'Stop'

$repoRoot       = (Resolve-Path (Join-Path $scriptRoot '..')).Path
$functionFolder = Join-Path $repoRoot 'Settings\Functions\psToDo'
$settingsPath   = Join-Path $repoRoot 'Settings\Config\ExchangeOnpremSettings.json'
$dbPath         = Join-Path $repoRoot 'Files\db\monitorobjects.json'

foreach ($file in Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File) { . $file.FullName }

if (-not $From -and (Test-Path $dbPath)) {
    $From = @((Get-Content -Raw -Encoding UTF8 $dbPath | ConvertFrom-Json) | ForEach-Object { $_.mail.mailSender } |
              Where-Object { $_ -and $_ -notmatch 'REPLACE_ME' }) | Select-Object -First 1
}
if (-not $From) { throw 'No sender found. Pass -From.' }

$settings = Get-OnpremExchangeSettings -Path $settingsPath -RepoRoot $repoRoot
Write-Host "Settings : $($settings.SmtpServer):$($settings.SmtpPort) TLS=$($settings.UseTls) Auth=$($settings.AuthMode)" -ForegroundColor Cyan

Connect-OnpremExchangeSmtp -Settings $settings | Out-Null

$alert = [pscustomobject]@{
    id = 'TEST'; name = 'psToDo on-prem SMTP test'; servername = $env:COMPUTERNAME; template = 'Test'
    environment = 'test'; action = 'No action - this is a connectivity test from Install-help\Test-OnpremSmtp.ps1.'
    expireDate = (Get-Date).AddDays(10).Date; daysLeft = 10; trigger = 30; allTriggers = @(30, 15, 7)
    urgency = 'ok'; stateKey = 'test'
    mailOn = $true; mailSender = $From; mailSubject = 'psToDo on-prem SMTP test'
    mailBody = "If you can read this, the psToDo host can submit mail through $($settings.SmtpServer)."
    mailRecipients = @($To)
}

try {
    $ok = Send-AlertMail -Alert $alert
    if ($ok) { Write-Host "Test mail submitted to $To from $From." -ForegroundColor Green }
    else     { Write-Host 'Test mail was NOT sent. See the warning above.' -ForegroundColor Red }
}
finally { Disconnect-OnpremExchangeSmtp }
