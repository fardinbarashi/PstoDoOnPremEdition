<#
System requirements
    PSVersion  5.1 (Desktop) or 7.x (Core)
    Module     ActiveDirectory (RSAT: Active Directory Domain Services and
               Lightweight Directory Services Tools). No download is attempted -
               install RSAT on the host first:
                 Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
    Rights     Any domain user can read the attributes used here.

psToDo plugin : Import-ADAccountExpiry (ON-PREMISES EDITION)
    Author  : Fardin Barashi
    Title   : psToDo - Active Directory account & password expiry importer
    Purpose : On-prem counterpart of the cloud edition's Entra app registration
              importer. Where the cloud edition watches app secrets and
              certificates expiring in Entra ID, this plugin watches
                 (a) user accounts with an AccountExpirationDate, and
                 (b) service accounts whose password will expire
                     (msDS-UserPasswordExpiryTimeComputed),
              and merges those dates into Files\db\monitorobjects.json.
              It only talks to a domain controller (LDAP) - nothing else.

Behaviour
    1. Reads accounts from the OUs / groups configured in
       Settings\plugin\adaccount\config\adaccount-defaults.json.
    2. Each expiring thing becomes ONE monitored object:
         - "<name> - account expiry"   (accountExpires)
         - "<name> - password expiry"  (password policy / FGPP)
       Accounts with PasswordNeverExpires and accounts that never expire are
       ignored. Disabled accounts are skipped when excludeDisabled is true.
    3. Fields AD does NOT provide (template, environment, triggers, mail, ...)
       are taken from adaccount-defaults.json. The tokens {{name}},
       {{expireDate}}, {{servername}}, {{environment}}, {{samAccountName}},
       {{distinguishedName}} and {{expiryType}} are replaced per object.
    4. BEFORE writing anything it takes a timestamped backup of
       monitorobjects.json into Files\backup\psToDo\.
    5. It MERGES rather than overwrites:
         - existing manual objects are never touched
         - objects previously imported by this plugin (matched on
           adAccountKey = "<objectGUID>|account" or "<objectGUID>|password")
           get their expireDate refreshed (unless -NoDateRefresh)
         - new items are appended with a fresh numeric id
       Imported objects carry the extra fields "source", "adAccountKey",
       "samAccountName", "distinguishedName" and "expiryType", so re-runs are
       idempotent. The psToDo scripts ignore unknown fields, so this is safe.

Switches
    -UseMockData   Skip AD completely and use built-in sample accounts.
    -NoDateRefresh Do not update expireDate on already-imported objects.
    -DomainController  Query a specific DC instead of the closest one.
    -WhatIf        Show what would change without writing (SupportsShouldProcess).

Examples
    .\Import-ADAccountExpiry-Onprem-v1.0.ps1 -UseMockData -WhatIf
    .\Import-ADAccountExpiry-Onprem-v1.0.ps1 -UseMockData
    .\Import-ADAccountExpiry-Onprem-v1.0.ps1
    .\Import-ADAccountExpiry-Onprem-v1.0.ps1 -DomainController dc-01.lab.local
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $UseMockData,
    [switch] $NoDateRefresh,
    [string] $DomainController
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

$configFolder          = Join-Path (Join-Path $pluginRoot 'adaccount') 'config'
$monitoringObjectsPath = Join-Path $repoRoot 'Files\db\monitorobjects.json'
$defaultsPath          = Join-Path $configFolder 'adaccount-defaults.json'
$versionPath           = Join-Path $configFolder 'version.json'
$backupFolder          = Join-Path $repoRoot 'Files\backup\psToDo'
$functionFolder        = Join-Path $repoRoot 'Settings\Functions\psToDo'

$expireDateFormat = 'yyyy-MM-dd'

#------------------------------- Helpers -----------------------------

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

# Build one ordered monitored object from the defaults template + one expiring AD item.
function New-MonitorObjectFromAccount {
    param(
        [Parameter(Mandatory)] $Defaults,
        [Parameter(Mandatory)] $Item,      # PSCustomObject: Name, ExpireDate, ServerName, Key, SamAccountName, DistinguishedName, ExpiryType
        [Parameter(Mandatory)] [int] $Id,
        [string] $Environment
    )

    $expire = ([datetime]$Item.ExpireDate).ToString($expireDateFormat)
    $env    = if ([string]::IsNullOrWhiteSpace($Environment)) { $Defaults.environment } else { $Environment }

    $tokens = @{
        name              = $Item.Name
        expireDate        = $expire
        servername        = $Item.ServerName
        environment       = $env
        samAccountName    = $Item.SamAccountName
        distinguishedName = $Item.DistinguishedName
        expiryType        = $Item.ExpiryType
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
        name                = "$($Item.Name) - $($Item.ExpiryType)"
        expireDate          = $expire
        template            = $Defaults.template
        servername          = $Item.ServerName
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
        source              = 'ADAccount'
        adAccountKey        = [string]$Item.Key
        samAccountName      = [string]$Item.SamAccountName
        distinguishedName   = [string]$Item.DistinguishedName
        expiryType          = [string]$Item.ExpiryType
    }
}

# Turn raw user records into a flat list of expiring items.
# A raw record has: ObjectGuid, Name, SamAccountName, DistinguishedName, Enabled,
# AccountExpirationDate, PasswordNeverExpires, PasswordExpiry (datetime or $null), WatchPassword (bool), Domain
function Get-ExpiryItemList {
    param(
        [Parameter(Mandatory)] $Users,
        [bool] $IncludeAccountExpiry,
        [bool] $IncludePasswordExpiry,
        [bool] $ExcludeDisabled
    )

    $list = [System.Collections.Generic.List[object]]::new()

    foreach ($u in $Users) {
        if ($ExcludeDisabled -and -not $u.Enabled) { Write-Verbose "Skip disabled $($u.SamAccountName)"; continue }

        if ($IncludeAccountExpiry -and $null -ne $u.AccountExpirationDate) {
            $list.Add([pscustomobject]@{
                Name              = $u.Name
                ExpireDate        = $u.AccountExpirationDate
                ServerName        = $u.Domain
                Key               = "$($u.ObjectGuid)|account"
                SamAccountName    = $u.SamAccountName
                DistinguishedName = $u.DistinguishedName
                ExpiryType        = 'account expiry'
            })
        }

        if ($IncludePasswordExpiry -and $u.WatchPassword -and -not $u.PasswordNeverExpires -and $null -ne $u.PasswordExpiry) {
            $list.Add([pscustomobject]@{
                Name              = $u.Name
                ExpireDate        = $u.PasswordExpiry
                ServerName        = $u.Domain
                Key               = "$($u.ObjectGuid)|password"
                SamAccountName    = $u.SamAccountName
                DistinguishedName = $u.DistinguishedName
                ExpiryType        = 'password expiry'
            })
        }
    }

    return $list
}

# Built-in sample data so the merge/backup logic can be tested without a domain.
function Get-MockUsers {
    [pscustomobject]@{ ObjectGuid = '10000000-0000-0000-0000-000000000001'; Name = 'svc-backup';    SamAccountName = 'svc-backup';    DistinguishedName = 'CN=svc-backup,OU=Service Accounts,DC=lab,DC=local';    Enabled = $true;  AccountExpirationDate = $null;                    PasswordNeverExpires = $false; PasswordExpiry = (Get-Date).AddDays(12);  WatchPassword = $true;  Domain = 'lab.local' }
    [pscustomobject]@{ ObjectGuid = '10000000-0000-0000-0000-000000000002'; Name = 'svc-sql-agent'; SamAccountName = 'svc-sql-agent'; DistinguishedName = 'CN=svc-sql-agent,OU=Service Accounts,DC=lab,DC=local'; Enabled = $true;  AccountExpirationDate = $null;                    PasswordNeverExpires = $true;  PasswordExpiry = $null;                   WatchPassword = $true;  Domain = 'lab.local' }
    [pscustomobject]@{ ObjectGuid = '10000000-0000-0000-0000-000000000003'; Name = 'Consultant Anna'; SamAccountName = 'ext-anna';    DistinguishedName = 'CN=Consultant Anna,OU=External,DC=lab,DC=local';      Enabled = $true;  AccountExpirationDate = (Get-Date).AddDays(40);   PasswordNeverExpires = $false; PasswordExpiry = (Get-Date).AddDays(60);  WatchPassword = $false; Domain = 'lab.local' }
    [pscustomobject]@{ ObjectGuid = '10000000-0000-0000-0000-000000000004'; Name = 'Old Vendor';     SamAccountName = 'ext-vendor';   DistinguishedName = 'CN=Old Vendor,OU=External,DC=lab,DC=local';           Enabled = $false; AccountExpirationDate = (Get-Date).AddDays(5);    PasswordNeverExpires = $false; PasswordExpiry = $null;                   WatchPassword = $false; Domain = 'lab.local' }
}

#------------------------------- Load defaults & existing data -------------------------------

Write-Host 'psToDo :: Import-ADAccountExpiry (on-prem)' -ForegroundColor Cyan

if (Test-Path $versionPath) {
    $version = Get-Content -Raw -Encoding UTF8 $versionPath | ConvertFrom-Json
    Write-Host "- Version : $($version.ADAccountVersion)" -ForegroundColor DarkGray
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

$includeAccountExpiry  = if ($null -eq $defaults.includeAccountExpiry)  { $true } else { [bool]$defaults.includeAccountExpiry }
$includePasswordExpiry = if ($null -eq $defaults.includePasswordExpiry) { $true } else { [bool]$defaults.includePasswordExpiry }
$excludeDisabled       = if ($null -eq $defaults.excludeDisabled)       { $true } else { [bool]$defaults.excludeDisabled }

#------------------------------- Fetch accounts -------------------------------

if ($UseMockData) {
    Write-Host '- Using MOCK data (no AD call).' -ForegroundColor Yellow
    $users      = @(Get-MockUsers)
    $domainName = 'lab.local (mock)'
} else {
    $functionFiles = Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File
    foreach ($file in $functionFiles) { . $file.FullName }

    Test-RequiredModule -Modules @('ActiveDirectory') -InstallHint @{
        ActiveDirectory = 'Install RSAT: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
    }

    $adArgs = @{ ErrorAction = 'Stop' }
    if ($DomainController) { $adArgs.Server = $DomainController }

    $domain     = Get-ADDomain @adArgs
    $domainName = $domain.DNSRoot
    Write-Host "- Domain : $domainName" -ForegroundColor DarkGray

    $props = 'Name', 'SamAccountName', 'DistinguishedName', 'Enabled', 'AccountExpirationDate', 'PasswordNeverExpires', 'ObjectGUID', 'msDS-UserPasswordExpiryTimeComputed'

    # (a) account-expiry candidates: every enabled user in the search bases with accountExpires set
    $searchBases = @($defaults.searchBases | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not $searchBases) { $searchBases = @($domain.DistinguishedName) }

    $byGuid = @{}
    if ($includeAccountExpiry) {
        foreach ($sb in $searchBases) {
            Write-Host "- Reading accounts with an expiry date under $sb" -ForegroundColor DarkGray
            $found = @(Get-ADUser @adArgs -SearchBase $sb -LDAPFilter '(&(accountExpires>=1)(!(accountExpires=9223372036854775807)))' -Properties $props)
            foreach ($u in $found) { $byGuid["$($u.ObjectGUID)"] = $u }
        }
    }

    # (b) password-expiry candidates: users in the configured OUs / groups
    $pwWatch = @{}
    if ($includePasswordExpiry) {
        foreach ($sb in @($defaults.passwordExpirySearchBases | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            Write-Host "- Reading password expiry for users under $sb" -ForegroundColor DarkGray
            foreach ($u in @(Get-ADUser @adArgs -SearchBase $sb -Filter * -Properties $props)) {
                $byGuid["$($u.ObjectGUID)"] = $u; $pwWatch["$($u.ObjectGUID)"] = $true
            }
        }
        foreach ($g in @($defaults.passwordExpiryGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            Write-Host "- Reading password expiry for members of group $g" -ForegroundColor DarkGray
            foreach ($m in @(Get-ADGroupMember @adArgs -Identity $g | Where-Object objectClass -eq 'user')) {
                $u = Get-ADUser @adArgs -Identity $m.distinguishedName -Properties $props
                $byGuid["$($u.ObjectGUID)"] = $u; $pwWatch["$($u.ObjectGUID)"] = $true
            }
        }
    }

    $users = foreach ($u in $byGuid.Values) {
        $pwRaw = $u.'msDS-UserPasswordExpiryTimeComputed'
        $pwExpiry = $null
        if ($pwRaw -and $pwRaw -gt 0 -and $pwRaw -lt [long]::MaxValue) {
            try { $pwExpiry = [datetime]::FromFileTime([long]$pwRaw) } catch { $pwExpiry = $null }
        }
        [pscustomobject]@{
            ObjectGuid            = "$($u.ObjectGUID)"
            Name                  = $u.Name
            SamAccountName        = $u.SamAccountName
            DistinguishedName     = $u.DistinguishedName
            Enabled               = [bool]$u.Enabled
            AccountExpirationDate = $u.AccountExpirationDate
            PasswordNeverExpires  = [bool]$u.PasswordNeverExpires
            PasswordExpiry        = $pwExpiry
            WatchPassword         = $pwWatch.ContainsKey("$($u.ObjectGUID)")
            Domain                = $domainName
        }
    }
    $users = @($users)
}

$adEnv = "AD on-prem - $domainName"
Write-Host "- Environment tag : $adEnv" -ForegroundColor DarkGray

$items = Get-ExpiryItemList -Users $users -IncludeAccountExpiry $includeAccountExpiry -IncludePasswordExpiry $includePasswordExpiry -ExcludeDisabled $excludeDisabled
Write-Host "- Expiring items found : $($items.Count) (from $($users.Count) account(s))" -ForegroundColor DarkGray

#------------------------------- Merge -------------------------------

$existingByKey = @{}
foreach ($obj in $existing) {
    if ($obj.PSObject.Properties.Name -contains 'adAccountKey' -and $obj.adAccountKey) {
        $existingByKey[[string]$obj.adAccountKey] = $obj
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
foreach ($item in $items) {
    if ($existingByKey.ContainsKey($item.Key)) {
        $target = $existingByKey[$item.Key]
        $newExpire = ([datetime]$item.ExpireDate).ToString($expireDateFormat)
        if (-not $NoDateRefresh -and [string]$target.expireDate -ne $newExpire) {
            $target.expireDate = $newExpire
            $updated++
        }
    } else {
        $maxId++
        $result.Add((New-MonitorObjectFromAccount -Defaults $defaults -Item $item -Id $maxId -Environment $adEnv))
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
