<#
System requirements
    PSVersion  5.1 (Desktop) or 7.x (Core)
    Host       Domain-joined Windows host with network access to the Exchange
               remote PowerShell endpoint (http://<server>/PowerShell/).
    Rights     The account running the script needs the Exchange RBAC role
               "View-Only Configuration" (Get-ExchangeServer, Get-ExchangeCertificate).

psToDo plugin : Import-ExchangeCertificates (ON-PREMISES EDITION)
    Author  : Fardin Barashi
    Title   : psToDo - on-prem Exchange certificate importer
    Purpose : Read every certificate installed on every on-prem Exchange server
              (Get-ExchangeCertificate) via Exchange remote PowerShell and merge
              their expiry dates into Files\db\monitorobjects.json.
              On-prem counterpart of the cloud edition's Entra app registration
              importer: same merge/backup behaviour, different source.

Behaviour
    1. Connects to the Exchange remote PowerShell endpoint named in
       Settings\Config\ExchangeOnpremSettings.json (ExchangePowerShellUri) with
       Kerberos (default) or Basic auth, imports ONLY Get-ExchangeServer and
       Get-ExchangeCertificate, and reads the certificates on every server.
       Nothing else on the network is contacted.
    2. By default self-signed certificates and certificates bound to no
       service are skipped (they are noise). Change includeSelfSigned /
       includeServicesNone in exchangecert-defaults.json to keep them.
    3. Fields Exchange does NOT provide (template, environment, triggers,
       mail, ...) are taken from Settings\plugin\exchangecert\config\exchangecert-defaults.json.
       The tokens {{name}}, {{expireDate}}, {{servername}}, {{environment}},
       {{thumbprint}}, {{subject}} and {{services}} are replaced per object.
    4. BEFORE writing anything it takes a timestamped backup of
       monitorobjects.json into Files\backup\psToDo\.
    5. It MERGES rather than overwrites:
         - existing manual objects are never touched
         - objects previously imported by this plugin (matched on
           exchangeCertKey = "<server>|<thumbprint>") get their expireDate
           refreshed (unless -NoDateRefresh)
         - new certificates are appended with a fresh numeric id
       Imported objects carry the extra fields "source", "exchangeCertKey",
       "thumbprint", "subject" and "services", so re-runs are idempotent.
       The psToDo scripts ignore unknown fields, so this is safe.

Switches
    -UseMockData   Skip Exchange completely and use built-in sample certificates.
                   Lets you test the backup/merge logic without an Exchange server.
    -NoDateRefresh Do not update expireDate on already-imported objects.
    -Server        Only read the named server(s) instead of every Exchange server.
    -WhatIf        Show what would change without writing (SupportsShouldProcess).

Examples
    # Lab test, no Exchange needed:
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1 -UseMockData -WhatIf
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1 -UseMockData

    # Live against the organisation:
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1
    .\Import-ExchangeCertificates-Onprem-v1.0.ps1 -Server mail-01.lab.local, mail-02.lab.local
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]   $UseMockData,
    [switch]   $NoDateRefresh,
    [string[]] $Server
)

# Resolve the folder this script lives in. $PSScriptRoot is empty when the code is
# run with F8 / "Run Selection" in ISE or VS Code, or pasted into a console, which
# turns every path into "\Files\..." on the current drive. Fall back to the script
# path, then to the current folder.
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

# Windows PowerShell 5.1: its built-in type data for System.Array makes ConvertTo-Json
# write arrays as {"value":[...],"Count":n}. Removing it gives plain JSON arrays, like PS 7.
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

#------------------------------- Helpers -----------------------------

# Replace {{token}} placeholders in any string using a lookup table.
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

# Build one ordered monitored object from the defaults template + one Exchange certificate.
function New-MonitorObjectFromCertificate {
    param(
        [Parameter(Mandatory)] $Defaults,
        [Parameter(Mandatory)] $Certificate,   # PSCustomObject: Name, ExpireDate, ServerName, Key, Thumbprint, Subject, Services
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

# Flatten raw Get-ExchangeCertificate output (already tagged with ServerFqdn) into a list to monitor.
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

# Built-in sample data so the merge/backup logic can be tested without an Exchange server.
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
        # Import only the two cmdlets we need. -DisableNameChecking silences the unapproved-verb noise.
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

# Index existing imported objects by their exchangeCertKey.
$existingByKey = @{}
foreach ($obj in $existing) {
    if ($obj.PSObject.Properties.Name -contains 'exchangeCertKey' -and $obj.exchangeCertKey) {
        $existingByKey[[string]$obj.exchangeCertKey] = $obj
    }
}

# Next free numeric id (ids in the file are numeric strings).
$maxId = 0
foreach ($obj in $existing) {
    $n = 0
    if ([int]::TryParse([string]$obj.id, [ref]$n) -and $n -gt $maxId) { $maxId = $n }
}

$result = [System.Collections.Generic.List[object]]::new()
foreach ($obj in $existing) { $result.Add($obj) }   # keep every existing object, order preserved

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

    # -InputObject @(...) instead of -AsArray (PS 7 only): a JSON array on both 5.1 and 7.
    $json = ConvertTo-Json -InputObject @($result) -Depth 20
    Set-Content -Path $monitoringObjectsPath -Value $json -Encoding UTF8
    Write-Host "- Wrote $($result.Count) objects to $monitoringObjectsPath" -ForegroundColor Green
}
