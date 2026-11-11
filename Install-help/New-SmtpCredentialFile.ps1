<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x
Run as the account that runs the scheduled task (the file can only be read by that account on that server).

About Script :
Author : Fardin Barashi
Title : psToDo setup - SMTP credential file
Description : Saves the SMTP username and password for AuthMode Credential to an encrypted (DPAPI) file.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\New-SmtpCredentialFile.ps1
    .\New-SmtpCredentialFile.ps1 -UserName LAB\svc-pstodo

Version : 1.0-onprem
Release day : 2026-09-08
#>

[CmdletBinding()]
param(
    [string] $OutFile = '',
    [string] $UserName
)

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($OutFile)) { $OutFile = (Join-Path (Resolve-Path (Join-Path $scriptRoot '..')).Path 'Settings\Config\smtp-credential.xml') }

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) {
    throw 'DPAPI-protected credential files are only supported on Windows.'
}

Write-Host "Running as $($env:USERDOMAIN)\$($env:USERNAME) on $($env:COMPUTERNAME)." -ForegroundColor Cyan
Write-Host 'The file will only be readable by this account on this machine.' -ForegroundColor DarkGray

$cred = if ($UserName) { Get-Credential -UserName $UserName -Message 'Password for the SMTP / Exchange account' }
        else            { Get-Credential -Message 'SMTP / Exchange account (DOMAIN\user or user@domain)' }
if (-not $cred) { throw 'No credential entered.' }

$dir = Split-Path $OutFile -Parent
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$cred | Export-Clixml -Path $OutFile
Write-Host "Credential for $($cred.UserName) written to $OutFile" -ForegroundColor Green

# Verify it round-trips for this account
$check = Import-Clixml -Path $OutFile
if ($check.UserName -ne $cred.UserName) { throw 'Round-trip check failed.' }
Write-Host 'Round-trip check OK.' -ForegroundColor Green
Write-Host 'Keep this file out of git and out of any web root.' -ForegroundColor Yellow
