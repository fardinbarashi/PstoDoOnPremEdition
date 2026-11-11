<#
System requirements
PSVersion : Windows PowerShell 5.1 or PowerShell 7.x
Rights : Exchange RBAC role "View-Only Configuration"

About Script :
Author : Fardin Barashi
Title : psToDo plugin - Exchange certificates
Description : Imports the certificates on the on-prem Exchange servers into monitorobjects.json.
Defaults: Settings\plugin\exchangecert\config\exchangecert-defaults.json
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1 -UseMockData -WhatIf
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1 -Server mail-01.lab.local

Version : 1.0-onprem
Release day : 2026-09-08
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]   $UseMockData,
    [switch]   $NoDateRefresh,
    [string[]] $Server
)

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

Remove-TypeData -TypeName System.Array -ErrorAction SilentlyContinue

$ErrorActionPreference = 'Stop'

#------------------------------- Paths -------------------------------

$pluginRoot = $scriptRoot
$repoRoot   = (Resolve-Path (Join-Path $scriptRoot '..\..')).Path

$configFolder          = Join-Path (Join-Path $pluginRoot 'exchangecert') 'config'
$monitoringObjectsPath = Join-Path $repoRoot 'Files\db\monitorobjects.json'
$exchangeSettingsPath  = Join-Path $repoRoot 'Settings\Config\ExchangeOnpremSettings.json'
$defaultsPath          = Join-Path $configFolder 'exchangecert-defaults.json'
$versionPath           = Join-Path $configFolder 'version.json'
$backupFolder          = Join-Path $repoRoot 'Files\backup\psToDo'
$functionFolder        = Join-Path $repoRoot 'Settings\Functions\psToDo'

$expireDateFormat = 'yyyy-MM-dd'

#------------------------------- Functions -----------------------------

function Expand-Tokens {
    param(
        [string] $Text,
        [hashtable] $Tokens
    )
    if ($null -eq $Text) { return $Text }
    foreach ($key in $Tokens.Keys) {
        $Text = $Text -replace [regex]::Escape("{{$key}}"), [string]$Tokens[$key]
    }
    return $Text
}

function New-MonitorObjectFromCertificate {
    param(
        [Parameter(Mandatory)] $Defaults,
        [Parameter(Mandatory)] $Certificate,
        [Parameter(Mandatory)] [int] $Id,
        [string] $Environment
    )

    $expire = ([datetime]$Certificate.ExpireDate).ToString($expireDateFormat)
    $env    = if ([string]::IsNullOrWhiteSpace($Environment)) { $Defaults.environment } else { $Environment }

    $tokens = @{
        name        = $Certificate.Name
        expireDate  = $expire
        servername  = $Certificate.ServerName
        environment = $env
        thumbprint  = $Certificate.Thumbprint
        subject     = $Certificate.Subject
        services    = $Certificate.Services
    }

    $mail = [ordered]@{
        mailSender     = $Defaults.mail.mailSender
        mailSubject    = Expand-Tokens $Defaults.mail.mailSubject $tokens
        mailBody       = Expand-Tokens $Defaults.mail.mailBody    $tokens
        mailRecipients = @($Defaults.mail.mailRecipients)
    }

    $teams = [ordered]@{
        teamSubject    = Expand-Tokens $Defaults.teams.teamSubject $tokens
        teamBody       = Expand-Tokens $Defaults.teams.teamBody    $tokens
        teamWebhookUrl = $Defaults.teams.teamWebhookUrl
    }

    return [ordered]@{
        id                  = [string]$Id
        name                = $Certificate.Name
        expireDate          = $expire
        template            = $Defaults.template
        servername          = $Certificate.ServerName
        environment         = $env
        status              = $Defaults.status
        description         = Expand-Tokens $Defaults.description $tokens
        '1dateTrigger'      = $Defaults.'1dateTrigger'
        '2dateTrigger'      = $Defaults.'2dateTrigger'
        '3dateTrigger'      = $Defaults.'3dateTrigger'
        notifyMethodbyMail  = $Defaults.notifyMethodbyMail
        notifyMethodbyTeams = $false
        mail                = $mail
        teams               = $teams
        source              = 'ExchangeCert'
        exchangeCertKey     = [string]$Certificate.Key
        thumbprint          = [string]$Certificate.Thumbprint
        subject             = [string]$Certificate.Subject
        services            = [string]$Certificate.Services
    }
}

function Get-CertificateList {
    param(
        [Parameter(Mandatory)] $RawCertificates,
        [bool] $IncludeSelfSigned,
        [bool] $IncludeServicesNone
    )

    $list = [System.Collections.Generic.List[object]]::new()

    foreach ($c in $RawCertificates) {
        if ($null -eq $c.NotAfter) { continue }

        $services = "$($c.Services)".Trim()
        if (-not $IncludeSelfSigned   -and $c.IsSelfSigned)                          { Write-Verbose "Skip self-signed $($c.Thumbprint) on $($c.ServerFqdn)"; continue }
        if (-not $IncludeServicesNone -and ($services -eq 'None' -or $services -eq '')) { Write-Verbose "Skip unbound $($c.Thumbprint) on $($c.ServerFqdn)";     continue }

        $label = if (-not [string]::IsNullOrWhiteSpace($c.FriendlyName)) { $c.FriendlyName }
                 elseif ("$($c.Subject)" -match 'CN=([^,]+)')            { $Matches[1] }
                 else                                                    { $c.Thumbprint }

        $list.Add([pscustomobject]@{
            Name       = "$($c.ServerFqdn) - $label ($services)"
            ExpireDate = $c.NotAfter
            ServerName = $c.ServerFqdn
            Key        = "$($c.ServerFqdn)|$($c.Thumbprint)"
            Thumbprint = "$($c.Thumbprint)"
            Subject    = "$($c.Subject)"
            Services   = $services
        })
    }

    return $list
}

function Get-MockCertificates {
    [pscustomobject]@{ ServerFqdn = 'mail-01.lab.local'; Thumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678'; FriendlyName = 'mail.lab.local (SAN)'; Subject = 'CN=mail.lab.local, O=Lab'; NotAfter = (Get-Date).AddDays(45);  Services = 'IMAP, POP, IIS, SMTP'; IsSelfSigned = $false }
    [pscustomobject]@{ ServerFqdn = 'mail-01.lab.local'; Thumbprint = 'FFEEDDCCBBAA99887766554433221100FFEEDDCC'; FriendlyName = 'Microsoft Exchange Server Auth Certificate'; Subject = 'CN=Microsoft Exchange Server Auth Certificate'; NotAfter = (Get-Date).AddDays(400); Services = 'SMTP'; IsSelfSigned = $true }
    [pscustomobject]@{ ServerFqdn = 'mail-02.lab.local'; Thumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678'; FriendlyName = 'mail.lab.local (SAN)'; Subject = 'CN=mail.lab.local, O=Lab'; NotAfter = (Get-Date).AddDays(45);  Services = 'IMAP, POP, IIS, SMTP'; IsSelfSigned = $false }
    [pscustomobject]@{ ServerFqdn = 'mail-02.lab.local'; Thumbprint = '0123456789ABCDEF0123456789ABCDEF01234567'; FriendlyName = 'old-wildcard'; Subject = 'CN=*.lab.local'; NotAfter = (Get-Date).AddDays(10); Services = 'None'; IsSelfSigned = $false }
}

#------------------------------- Load defaults & existing data -------------------------------

Write-Host 'psToDo :: Import-ExchangeCertificates (on-prem)' -ForegroundColor Cyan

if (Test-Path $versionPath) {
    $version = Get-Content -Raw -Encoding UTF8 $versionPath | ConvertFrom-Json
    Write-Host "- Version : $($version.ExchangeCertVersion)" -ForegroundColor DarkGray
    if ($version.Changes) { Write-Host "- Changes : $($version.Changes)" -ForegroundColor DarkGray }
}

if (-not (Test-Path $defaultsPath)) { throw "Defaults template not found: $defaultsPath" }
$defaults = Get-Content -Raw -Encoding UTF8 $defaultsPath | ConvertFrom-Json

if (Test-Path $monitoringObjectsPath) {
    $existing = @((Get-Content -Raw -Encoding UTF8 $monitoringObjectsPath | ConvertFrom-Json))
} else {
    Write-Warning "monitorobjects.json not found at $monitoringObjectsPath - starting from an empty list."
    $existing = @()
}
Write-Host "- Existing objects in file : $($existing.Count)" -ForegroundColor DarkGray

#------------------------------- Fetch certificates -------------------------------

$exSession = $null
if ($UseMockData) {
    Write-Host '- Using MOCK data (no Exchange call).' -ForegroundColor Yellow
    $rawCerts = @(Get-MockCertificates)
    $orgName  = 'Lab (mock)'
} else {
    $functionFiles = Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File
    foreach ($file in $functionFiles) { . $file.FullName }

    $settings = Get-OnpremExchangeSettings -Path $exchangeSettingsPath -RepoRoot $repoRoot
    if ([string]::IsNullOrWhiteSpace($settings.ExchangePowerShellUri)) { throw 'ExchangeOnpremSettings.json is missing ExchangePowerShellUri' }

    Write-Host "- Connecting to Exchange remote PowerShell at $($settings.ExchangePowerShellUri) ($($settings.ExchangeAuthentication))..." -ForegroundColor DarkGray
    $sessionArgs = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $settings.ExchangePowerShellUri
        Authentication    = $settings.ExchangeAuthentication
        ErrorAction       = 'Stop'
    }
    if ($settings.ExchangeAuthentication -eq 'Basic') {
        if (-not $settings.ExchangeCredentialXmlPath -or -not (Test-Path $settings.ExchangeCredentialXmlPath)) {
            throw "ExchangeAuthentication is Basic but the credential file is missing: $($settings.ExchangeCredentialXmlPath)"
        }
        $sessionArgs.Credential = Import-Clixml -Path $settings.ExchangeCredentialXmlPath
        if ($settings.ExchangePowerShellUri -notmatch '^https://') { Write-Warning 'Basic auth over http sends the password in clear text. Use https.' }
    }

    try {
        $exSession = New-PSSession @sessionArgs
        Import-PSSession -Session $exSession -CommandName Get-ExchangeServer, Get-ExchangeCertificate -AllowClobber -DisableNameChecking | Out-Null

        $servers = @(Get-ExchangeServer | Where-Object { $_.ServerRole -notmatch 'Edge' })
        if ($Server) { $servers = @($servers | Where-Object { $_.Fqdn -in $Server -or $_.Name -in $Server }) }
        if (-not $servers) { throw 'Get-ExchangeServer returned no matching servers.' }
        Write-Host "- Exchange servers to read : $($servers.Count)" -ForegroundColor DarkGray

        $rawCerts = [System.Collections.Generic.List[object]]::new()
        foreach ($s in $servers) {
            try {
                $certs = @(Get-ExchangeCertificate -Server $s.Fqdn -ErrorAction Stop)
                foreach ($c in $certs) {
                    $rawCerts.Add([pscustomobject]@{
                        ServerFqdn   = $s.Fqdn
                        Thumbprint   = $c.Thumbprint
                        FriendlyName = $c.FriendlyName
                        Subject      = $c.Subject
                        NotAfter     = $c.NotAfter
                        Services     = "$($c.Services)"
                        IsSelfSigned = [bool]$c.IsSelfSigned
                    })
                }
                Write-Host "  $($s.Fqdn): $($certs.Count) certificate(s)" -ForegroundColor DarkGray
            }
            catch { Write-Warning "  $($s.Fqdn): Get-ExchangeCertificate failed - $($_.Exception.Message)" }
        }
        $orgName = try { ($servers | Select-Object -First 1).Domain } catch { 'Exchange' }
        if ([string]::IsNullOrWhiteSpace($orgName)) { $orgName = 'Exchange' }
    }
    finally {
        if ($exSession) { Remove-PSSession $exSession -ErrorAction SilentlyContinue }
    }
}

$exchangeEnv = "Exchange on-prem - $orgName"
Write-Host "- Environment tag : $exchangeEnv" -ForegroundColor DarkGray

$certificates = Get-CertificateList -RawCertificates $rawCerts `
                    -IncludeSelfSigned   ([bool]$defaults.includeSelfSigned) `
                    -IncludeServicesNone ([bool]$defaults.includeServicesNone)
Write-Host "- Certificates to monitor : $($certificates.Count) (of $(@($rawCerts).Count) found)" -ForegroundColor DarkGray

#------------------------------- Merge -------------------------------

$existingByKey = @{}
foreach ($obj in $existing) {
    if ($obj.PSObject.Properties.Name -contains 'exchangeCertKey' -and $obj.exchangeCertKey) {
        $existingByKey[[string]$obj.exchangeCertKey] = $obj
    }
}

$maxId = 0
foreach ($obj in $existing) {
    $n = 0
    if ([int]::TryParse([string]$obj.id, [ref]$n) -and $n -gt $maxId) { $maxId = $n }
}

$result = [System.Collections.Generic.List[object]]::new()
foreach ($obj in $existing) { $result.Add($obj) }

$added = 0; $updated = 0
foreach ($cert in $certificates) {
    if ($existingByKey.ContainsKey($cert.Key)) {
        $target = $existingByKey[$cert.Key]
        $newExpire = ([datetime]$cert.ExpireDate).ToString($expireDateFormat)
        if (-not $NoDateRefresh -and [string]$target.expireDate -ne $newExpire) {
            $target.expireDate = $newExpire
            $updated++
        }
    } else {
        $maxId++
        $result.Add((New-MonitorObjectFromCertificate -Defaults $defaults -Certificate $cert -Id $maxId -Environment $exchangeEnv))
        $added++
    }
}

Write-Host "- New objects to add      : $added"   -ForegroundColor Green
Write-Host "- Existing objects updated: $updated" -ForegroundColor Green
Write-Host "- Untouched objects       : $($existing.Count - $updated)" -ForegroundColor DarkGray

if ($added -eq 0 -and $updated -eq 0) {
    Write-Host 'Nothing to change. File left as is.' -ForegroundColor Yellow
    return
}

#------------------------------- Backup, then write -------------------------------

if ($PSCmdlet.ShouldProcess($monitoringObjectsPath, "Backup and write $($result.Count) monitored objects")) {

    if (Test-Path $monitoringObjectsPath) {
        if (-not (Test-Path $backupFolder)) { New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null }
        $backupfileDate = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
        $backupTarget   = Join-Path $backupFolder "monitorobjects-$backupfileDate.json"
        Copy-Item -Path $monitoringObjectsPath -Destination $backupTarget -Force
        Write-Host "- Backup written : $backupTarget" -ForegroundColor DarkGray
    }

    $dbDir = Split-Path $monitoringObjectsPath -Parent
    if ($dbDir -and -not (Test-Path $dbDir)) { New-Item -ItemType Directory -Path $dbDir -Force | Out-Null }

    $json = ConvertTo-Json -InputObject @($result) -Depth 20
    Set-Content -Path $monitoringObjectsPath -Value $json -Encoding UTF8
    Write-Host "- Wrote $($result.Count) objects to $monitoringObjectsPath" -ForegroundColor Green
}
