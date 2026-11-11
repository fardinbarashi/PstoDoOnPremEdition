<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x

About Script :
Author : Fardin Barashi
Title : psToDo tool - send alert again
Description : Removes one object's entries from sent-state.json, so the object alerts again at the next
run of psToDo-Onprem-v1.0.ps1. A backup of sent-state.json is taken first.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Reset-AlertState.ps1 -ObjectId 4

Version : 1.0-onprem
Release day : 2026-11-11
#>

[CmdletBinding()]
param([Parameter(Mandatory)] [string] $ObjectId)

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

$ErrorActionPreference = 'Stop'
trap { Write-Host "Result: $($_.Exception.Message)"; exit 1 }

$repoRoot     = (Resolve-Path (Join-Path $scriptRoot '..\..')).Path
$statePath    = [IO.Path]::Combine($repoRoot, 'Files', 'state', 'sent-state.json')
$backupFolder = [IO.Path]::Combine($repoRoot, 'Files', 'backup', 'psToDo')

if (-not (Test-Path $statePath)) { Write-Host 'Result: there is no sent-state.json yet, nothing to reset.'; exit 0 }

$raw = Get-Content -Raw -Encoding UTF8 $statePath
$state = [ordered]@{}
if (-not [string]::IsNullOrWhiteSpace($raw)) { ($raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $state[$_.Name] = $_.Value } }

$keys = @($state.Keys | Where-Object { ($_ -split '_')[0] -eq $ObjectId })
if (-not $keys) { Write-Host "Result: object $ObjectId has no recorded alerts, nothing to reset."; exit 0 }

if (-not (Test-Path $backupFolder)) { New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null }
$backup = [IO.Path]::Combine($backupFolder, "sendstate-$(Get-Date -Format 'yyyy-MM-dd_HH.mm.ss').json")
Copy-Item -Path $statePath -Destination $backup -Force
Write-Host "Backup : $backup"

foreach ($k in $keys) { $state.Remove($k); Write-Host "Removed: $k" }

if ($state.Count) { $state | ConvertTo-Json -Depth 3 | Set-Content -Path $statePath -Encoding UTF8 }
else              { '{}' | Set-Content -Path $statePath -Encoding UTF8 }

Write-Host "Result: $($keys.Count) entry(ies) removed. Object $ObjectId alerts again on the next run of the alert script, if it is inside a trigger window."
