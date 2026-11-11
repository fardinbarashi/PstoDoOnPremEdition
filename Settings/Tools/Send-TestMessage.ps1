<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x

About Script :
Author : Fardin Barashi
Title : psToDo tool - send test message
Description : Sends the alert mail of one object as a test. With -To the mail goes to that address
instead of the object's recipients. Nothing is recorded in sent-state.json.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Send-TestMessage.ps1 -ObjectId 4
    .\Send-TestMessage.ps1 -ObjectId 4 -To you@lab.local

Version : 1.0-onprem
Release day : 2026-11-11
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ObjectId,
    [string] $To
)

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

$ErrorActionPreference = 'Stop'
trap { Write-Host "Result: $($_.Exception.Message)"; exit 1 }

$repoRoot       = (Resolve-Path (Join-Path $scriptRoot '..\..')).Path
$functionFolder = [IO.Path]::Combine($repoRoot, 'Settings', 'Functions', 'psToDo')
$settingsPath   = [IO.Path]::Combine($repoRoot, 'Settings', 'Config', 'ExchangeOnpremSettings.json')
$dbPath         = [IO.Path]::Combine($repoRoot, 'Files', 'db', 'monitorobjects.json')
$triggerKeys    = @('1dateTrigger', '2dateTrigger', '3dateTrigger')

foreach ($file in Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File) { . $file.FullName }

$object = @((Get-Content -Raw -Encoding UTF8 $dbPath | ConvertFrom-Json)) | Where-Object { "$($_.id)" -eq $ObjectId } | Select-Object -First 1
if (-not $object) { throw "There is no object with id $ObjectId." }
Write-Host "Object     : $($object.id) - $($object.name)"

$expireDate = (Get-Date).Date
if (-not [string]::IsNullOrWhiteSpace($object.expireDate)) {
    $expireDate = [datetime]::ParseExact($object.expireDate, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture).Date
}
$daysLeft = ($expireDate - (Get-Date).Date).Days

[int[]] $triggers = @($triggerKeys | ForEach-Object { $object.$_ }) |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [int]$_ } | Sort-Object -Unique
$hit = $triggers | Where-Object { $daysLeft -le $_ -and $daysLeft -ge 0 } | Select-Object -First 1
$urgency = if     ($daysLeft -lt 0)                                               { 'expired'  }
           elseif ($triggers.Count -ge 1 -and $daysLeft -le $triggers[0])                   { 'critical' }
           elseif ($triggers.Count -ge 2 -and $daysLeft -le $triggers[1])         { 'warning'  }
           else                                                                   { 'ok'       }

$recipients = if ($To) { @($To) } else { @($object.mail.mailRecipients) }
if (-not ($recipients | Where-Object { $_ })) { throw 'The object has no mailRecipients. Give an address in "Send to".' }
if ([string]::IsNullOrWhiteSpace($object.mail.mailSender)) { throw 'The object has no mailSender.' }
Write-Host "Recipients : $($recipients -join ', ')"

$alert = [pscustomobject]@{
    id = "$($object.id)"; name = "$($object.name)"; servername = "$($object.servername)"; template = "$($object.template)"
    environment = "$($object.environment)"; action = "$($object.description)"
    expireDate = $expireDate; daysLeft = $daysLeft; trigger = $hit; allTriggers = @($triggers)
    urgency = $urgency; stateKey = 'test'
    mailOn = $true; mailSender = "$($object.mail.mailSender)"; mailSubject = "[TEST] $($object.mail.mailSubject)"
    mailBody = "$($object.mail.mailBody)"; mailRecipients = $recipients
}

$settings = Get-OnpremExchangeSettings -Path $settingsPath -RepoRoot $repoRoot
Write-Host "SMTP       : $($settings.SmtpServer):$($settings.SmtpPort) TLS=$($settings.UseTls) Auth=$($settings.AuthMode)"
Connect-OnpremExchangeSmtp -Settings $settings | Out-Null

$ok = $false
try { $ok = Send-AlertMail -Alert $alert }
finally { Disconnect-OnpremExchangeSmtp }

if ($ok) { Write-Host "Result: test message sent to $($recipients -join ', ')."; exit 0 }
Write-Host 'Result: the test message was NOT sent. See the warning above.'
exit 1
