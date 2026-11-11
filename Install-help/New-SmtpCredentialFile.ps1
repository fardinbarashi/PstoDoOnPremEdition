<#
psToDo setup helper : New-SmtpCredentialFile (ON-PREMISES EDITION)
    Author  : Fardin Barashi
    Purpose : Store a username + password for AuthMode = Credential (SMTP) or
              ExchangeAuthentication = Basic (Exchange remote PowerShell) as an
              encrypted CliXml file.

    The file is protected with Windows DPAPI: it can ONLY be decrypted by the
    SAME Windows account on the SAME machine that created it. So run this
    script AS THE ACCOUNT that will run the psToDo scheduled task
    (e.g.  runas /user:LAB\svc-pstodo "pwsh -File New-SmtpCredentialFile.ps1").

Parameters
    -OutFile   Where to write the file. Default: Settings\Config\smtp-credential.xml
               For Exchange remote PowerShell use Settings\Config\exchange-credential.xml
    -UserName  Pre-fill the user name (you are still prompted for the password).

Examples
    .\New-SmtpCredentialFile.ps1
    .\New-SmtpCredentialFile.ps1 -UserName LAB\svc-pstodo
    .\New-SmtpCredentialFile.ps1 -OutFile ..\Settings\Config\exchange-credential.xml -UserName LAB\svc-pstodo
#>

[CmdletBinding()]
param(
    [string] $OutFile = '',
    [string] $UserName
)

# Resolve the folder this script lives in. $PSScriptRoot is empty when the code is
# run with F8 / "Run Selection" in ISE or VS Code, or pasted into a console, which
# turns every path into "\Files\..." on the current drive. Fall back to the script
# path, then to the current folder.
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
