<#
System requirements
PSVersion : PowerShell 7.3.1 or later

About Script :
Author : Fardin Barashi
Title : psToDo plugin - Entra ID app registrations
Description : Imports client secret and certificate expiry dates of the app registrations into monitorobjects.json (app-only auth).
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Import-EntraAppRegistrations-v1.3.ps1 -UseMockData -WhatIf
    .\Import-EntraAppRegistrations-v1.3.ps1

Version : 1.3
Release day : 2026-09-08
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $UseMockData,
    [switch] $NoDateRefresh
)

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

Remove-TypeData -TypeName System.Array -ErrorAction SilentlyContinue

$ErrorActionPreference = 'Stop'

#------------------------------- Paths -------------------------------

$pluginRoot = $scriptRoot
$repoRoot   = (Resolve-Path (Join-Path $scriptRoot '..\..')).Path

$configFolder          = Join-Path (Join-Path $pluginRoot 'entra') 'config'
$monitoringObjectsPath = Join-Path $repoRoot 'Files\db\monitorobjects.json'
$graphSettingsPath     = Join-Path $repoRoot 'Settings\Config\MsGraphSettings.json'
$defaultsPath          = Join-Path $configFolder 'entra-app-defaults.json'
$versionPath           = Join-Path $configFolder 'version.json'
$backupFolder          = Join-Path $repoRoot 'Files\backup\psToDo'
$functionFolder        = Join-Path $repoRoot 'Settings\Functions\cloud-plugins'

$expireDateFormat = 'yyyy-MM-dd'

#------------------------------- Function -----------------------------

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

function New-MonitorObjectFromCredential {
    param(
        [Parameter(Mandatory)] $Defaults,
        [Parameter(Mandatory)] $Credential,
        [Parameter(Mandatory)] [int] $Id,
        [string] $Environment
    )

    $isEmpty = [bool]$Credential.NoCredentials
    $expire  = if ($isEmpty) { '' } else { ([datetime]$Credential.ExpireDate).ToString($expireDateFormat) }
    $env     = if ([string]::IsNullOrWhiteSpace($Environment)) { $Defaults.environment } else { $Environment }

    $tokens = @{
        name       = $Credential.Name
        expireDate = $expire
        servername = $Credential.ServerName
        environment = $env
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
        name                = $Credential.Name
        expireDate          = $expire
        template            = $Defaults.template
        servername          = $Credential.ServerName
        appId               = $Credential.ServerName
        environment         = $env
        status              = $Defaults.status
        description         = if ($isEmpty) { 'This app registration has no secrets or certificates to monitor.' } else { $Defaults.description }
        '1dateTrigger'      = $Defaults.'1dateTrigger'
        '2dateTrigger'      = $Defaults.'2dateTrigger'
        '3dateTrigger'      = $Defaults.'3dateTrigger'
        notifyMethodbyMail  = if ($isEmpty) { $false } else { $Defaults.notifyMethodbyMail }
        notifyMethodbyTeams = if ($isEmpty) { $false } else { $Defaults.notifyMethodbyTeams }
        mail                = $mail
        teams               = $teams
        source              = 'EntraAppReg'
        entraKeyId          = [string]$Credential.KeyId
        noCredentials       = $isEmpty
    }
}

function Get-CredentialListFromApps {
    param([Parameter(Mandatory)] $Apps)

    $list = [System.Collections.Generic.List[object]]::new()

    foreach ($app in $Apps) {
        $appName = $app.DisplayName
        $appId   = $app.AppId
        $added   = 0

        foreach ($secret in @($app.PasswordCredentials)) {
            if ($null -eq $secret.EndDateTime) { continue }
            $label = if ([string]::IsNullOrWhiteSpace($secret.DisplayName)) { 'secret' } else { $secret.DisplayName }
            $list.Add([pscustomobject]@{
                Name       = "$appName - $label (secret)"
                ExpireDate = $secret.EndDateTime
                ServerName = $appId
                KeyId      = [string]$secret.KeyId
                NoCredentials = $false
            })
            $added++
        }

        foreach ($cert in @($app.KeyCredentials)) {
            if ($null -eq $cert.EndDateTime) { continue }
            $label = if ([string]::IsNullOrWhiteSpace($cert.DisplayName)) { 'certificate' } else { $cert.DisplayName }
            $list.Add([pscustomobject]@{
                Name       = "$appName - $label (certificate)"
                ExpireDate = $cert.EndDateTime
                ServerName = $appId
                KeyId      = [string]$cert.KeyId
                NoCredentials = $false
            })
            $added++
        }

        if ($added -eq 0) {
            $list.Add([pscustomobject]@{
                Name       = "$appName - (no credentials)"
                ExpireDate = $null
                ServerName = $appId
                KeyId      = "app_$appId"
                NoCredentials = $true
            })
        }
    }

    return $list
}

function Get-MockApps {
    [pscustomobject]@{
        DisplayName = 'psToDo-Graph-Connector'
        AppId       = '11111111-1111-1111-1111-111111111111'
        PasswordCredentials = @(
            [pscustomobject]@{ KeyId = 'aaaaaaaa-0000-0000-0000-000000000001'; DisplayName = 'rotation-2026'; EndDateTime = (Get-Date).AddDays(45) }
        )
        KeyCredentials = @(
            [pscustomobject]@{ KeyId = 'aaaaaaaa-0000-0000-0000-000000000002'; DisplayName = 'signing-cert'; EndDateTime = (Get-Date).AddDays(120) }
        )
    }
    [pscustomobject]@{
        DisplayName = 'psToDo-Automation-App'
        AppId       = '22222222-2222-2222-2222-222222222222'
        PasswordCredentials = @(
            [pscustomobject]@{ KeyId = 'bbbbbbbb-0000-0000-0000-000000000001'; DisplayName = ''; EndDateTime = (Get-Date).AddDays(10) }
        )
        KeyCredentials = @()
    }
    [pscustomobject]@{
        DisplayName = 'psToDo-Empty-App'
        AppId       = '33333333-3333-3333-3333-333333333333'
        PasswordCredentials = @()
        KeyCredentials = @()
    }
}

#------------------------------- Load defaults & existing data -------------------------------

Write-Host 'psToDo :: Import-EntraAppRegistrations' -ForegroundColor Cyan

if (Test-Path $versionPath) {
    $version = Get-Content -Raw -Encoding UTF8 $versionPath | ConvertFrom-Json
    Write-Host "- Version : $($version.Entraversion)" -ForegroundColor DarkGray
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

#------------------------------- Fetch credentials -------------------------------

if ($UseMockData) {
    Write-Host '- Using MOCK data (no Graph call).' -ForegroundColor Yellow
    $apps = Get-MockApps
    $tenantId = '11111111-1111-1111-1111-111111111111'
    $tenantName = 'Contoso (mock)'
} else {
    $functionFiles = Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File
    foreach ($file in $functionFiles) { . $file.FullName }

    Initialize-RequiredModules -Modules @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.DirectoryManagement')

    if (-not (Test-Path $graphSettingsPath)) { throw "Cannot find Graph settings file: $graphSettingsPath" }
    $settings = Get-Content -Raw -Encoding UTF8 $graphSettingsPath | ConvertFrom-Json
    foreach ($k in 'TenantId', 'AppId', 'CertificateThumbprint') {
        if ([string]::IsNullOrWhiteSpace($settings.$k)) { throw "MsGraphSettings.json is missing $k" }
    }

    $certificate = Get-ChildItem "Cert:\LocalMachine\My\$($settings.CertificateThumbprint)" -ErrorAction SilentlyContinue
    if (-not $certificate) { throw "Certificate $($settings.CertificateThumbprint) not found in Cert:\LocalMachine\My" }

    Connect-CalenderReminderGraph -TenantId $settings.TenantId -AppId $settings.AppId -Certificate $certificate | Out-Null

    if ('Application.Read.All' -notin (Get-MgContext).Scopes) {
        Write-Warning "Token has no Application.Read.All scope. Reading applications may fail. Current scopes: $((Get-MgContext).Scopes -join ', ')"
    }

    Write-Host '- Reading app registrations from Entra ID...' -ForegroundColor DarkGray
    $apps = Get-MgApplication -All -Property 'id,appId,displayName,passwordCredentials,keyCredentials'

    $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
    $tenantId   = if ($org -and $org.Id)          { $org.Id }          else { (Get-MgContext).TenantId }
    $tenantName = if ($org -and $org.DisplayName) { $org.DisplayName } else { $tenantId }
    if (-not $org) { Write-Warning "Could not read the organization name (needs Organization.Read.All or Directory.Read.All). Using tenant id only." }
}

$entraEnv = "Entra - $tenantName - $tenantId"
Write-Host "- Environment tag : $entraEnv" -ForegroundColor DarkGray

$credentials = Get-CredentialListFromApps -Apps $apps
Write-Host "- Credentials with an expiry date found : $($credentials.Count)" -ForegroundColor DarkGray

#------------------------------- Merge -------------------------------

$existingByKeyId = @{}
foreach ($obj in $existing) {
    if ($obj.PSObject.Properties.Name -contains 'entraKeyId' -and $obj.entraKeyId) {
        $existingByKeyId[[string]$obj.entraKeyId] = $obj
    }
}

$maxId = 0
foreach ($obj in $existing) {
    $n = 0
    if ([int]::TryParse([string]$obj.id, [ref]$n) -and $n -gt $maxId) { $maxId = $n }
}

$result  = [System.Collections.Generic.List[object]]::new()
foreach ($obj in $existing) { $result.Add($obj) }

$added = 0; $updated = 0
foreach ($cred in $credentials) {
    if ($existingByKeyId.ContainsKey($cred.KeyId)) {
        $target = $existingByKeyId[$cred.KeyId]
        if ($cred.NoCredentials -or $null -eq $cred.ExpireDate) { continue }
        $newExpire = ([datetime]$cred.ExpireDate).ToString($expireDateFormat)
        if (-not $NoDateRefresh -and [string]$target.expireDate -ne $newExpire) {
            $target.expireDate = $newExpire
            $updated++
        }
    } else {
        $maxId++
        $result.Add((New-MonitorObjectFromCredential -Defaults $defaults -Credential $cred -Id $maxId -Environment $entraEnv))
        $added++
    }
}

Write-Host "- New objects to add      : $added"     -ForegroundColor Green
Write-Host "- Existing objects updated: $updated"   -ForegroundColor Green
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

    $json = ConvertTo-Json -InputObject @($result) -Depth 20
    Set-Content -Path $monitoringObjectsPath -Value $json -Encoding UTF8
    Write-Host "- Wrote $($result.Count) objects to $monitoringObjectsPath" -ForegroundColor Green
}
