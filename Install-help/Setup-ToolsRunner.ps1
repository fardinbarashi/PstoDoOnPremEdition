<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x, run as Administrator on the psToDo server

About Script :
Author : Fardin Barashi
Title : psToDo setup - Tools runner
Description : Creates Files\queue and the scheduled task "psToDo Tools", which runs psToDo-Tools-v1.0.ps1
every minute. Use the same account that runs psToDo-Onprem-v1.0.ps1. The IIS part (queue folder for the
app pool) is done by Setup-IisReportEditor.ps1. Safe to run again.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Setup-ToolsRunner.ps1 -Credential (Get-Credential CONTOSO\svc-pstodo) -WhatIf
    .\Setup-ToolsRunner.ps1 -Credential (Get-Credential CONTOSO\svc-pstodo)
    .\Setup-ToolsRunner.ps1 -GmsaAccount 'CONTOSO\svc-pstodo$'

Version : 1.0-onprem
Release day : 2026-11-11
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [pscredential] $Credential,
    [string] $GmsaAccount,
    [string] $TaskName = 'psToDo Tools',
    [string] $PowerShellPath,
    [string] $QueueFolder
)

$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }
$repoRoot = (Resolve-Path (Join-Path $scriptRoot '..')).Path
$runner   = Join-Path $repoRoot 'psToDo-Tools-v1.0.ps1'
if ([string]::IsNullOrWhiteSpace($QueueFolder))    { $QueueFolder    = Join-Path $repoRoot 'Files\queue' }
if ([string]::IsNullOrWhiteSpace($PowerShellPath)) {
    $hostExe = (Get-Process -Id $PID).Path
    if ([IO.Path]::GetFileNameWithoutExtension($hostExe) -notin 'powershell', 'pwsh') {
        $hostExe = Join-Path $PSHOME $(if ($PSVersionTable.PSVersion.Major -ge 6) { 'pwsh.exe' } else { 'powershell.exe' })
    }
    $PowerShellPath = $hostExe
}

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run this script as Administrator.' }
if (-not (Test-Path $runner)) { throw "Not found: $runner" }
if (-not $Credential -and [string]::IsNullOrWhiteSpace($GmsaAccount)) { throw 'Give -Credential or -GmsaAccount.' }
if ($Credential -and $GmsaAccount) { throw 'Give -Credential or -GmsaAccount, not both.' }
$account = if ($Credential) { $Credential.UserName } else { $GmsaAccount }

#------------------------------- 1. Queue folder -------------------------------

Write-Host '1. Queue folder' -ForegroundColor Cyan
if ($PSCmdlet.ShouldProcess($QueueFolder, "Create and grant Modify to $account")) {
    if (-not (Test-Path $QueueFolder)) { New-Item -ItemType Directory -Path $QueueFolder -Force | Out-Null }
    & icacls.exe $QueueFolder /grant "$($account):(OI)(CI)M" | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Warning "icacls failed for $account on $QueueFolder (exit $LASTEXITCODE)." }
    else { Write-Host "   $account -> $QueueFolder" -ForegroundColor Green }
}

#------------------------------- 2. Scheduled task -------------------------------

Write-Host '2. Scheduled task' -ForegroundColor Cyan
$action   = New-ScheduledTaskAction -Execute $PowerShellPath -WorkingDirectory $repoRoot `
                -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$runner`""
$trigger  = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Minutes 1)
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable `
                -ExecutionTimeLimit (New-TimeSpan -Hours 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

if ($PSCmdlet.ShouldProcess($TaskName, "Register scheduled task as $account (every minute)")) {
    if ($Credential) {
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force `
            -User $Credential.UserName -Password $Credential.GetNetworkCredential().Password `
            -Description 'Runs the tools ordered in the psToDo dashboard.' | Out-Null
    }
    else {
        $taskPrincipal = New-ScheduledTaskPrincipal -UserId $GmsaAccount -LogonType Password
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force `
            -Principal $taskPrincipal -Description 'Runs the tools ordered in the psToDo dashboard.' | Out-Null
    }
    Write-Host "   Task '$TaskName' registered: $PowerShellPath -File $runner" -ForegroundColor Green
    Start-ScheduledTask -TaskName $TaskName
}

#------------------------------- Summary -------------------------------

Write-Host ''
Write-Host 'Done. Next:' -ForegroundColor Cyan
Write-Host '  1. Run .\Install-help\Setup-IisReportEditor.ps1 again if it was run before this version:'
Write-Host '     it gives the IIS app pool access to the queue folder and writes psToDo:QueueFolder.'
Write-Host '  2. Run .\psToDo-HTML-Report-Editable-Onprem-v1.0.ps1 once (copies the current psToDo-api.ashx).'
Write-Host '  3. Open the dashboard as an editor: Tools is in the menu under Edit.'
Write-Host "  4. $QueueFolder\runner.json shows when the task last ran."
