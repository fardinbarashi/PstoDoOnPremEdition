<#
System requirements
PSVersion                      7.3.1
PSEdition                      Core

About Script :
Author : Fardin Barashi
Title : psToDo - ON-PREMISES EDITION
Description : This script sends calendar reminders to users based on specific criteria in the monitoring objects file.
Alerts are delivered as e-mail through an ON-PREM EXCHANGE SERVER over SMTP (STARTTLS).

What is different from the cloud edition (psToDo-v1.3.ps1)
  - No Microsoft Graph, no Entra ID app registration, no certificate auth,
    no PowerShell Gallery downloads. The host only needs .NET (ships with PowerShell).
  - The ONLY network endpoint this script talks to is the SMTP host named in
    Settings\Config\ExchangeOnpremSettings.json.
  - Teams webhooks (a cloud service) are not implemented. notifyMethodbyTeams is
    read for file compatibility but ignored.
  - Plugins in Settings\plugin import Exchange certificates and AD account expiry
    instead of Entra app registrations and the M365 Message Center.
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Version : 1.0-onprem (based on psToDo 1.3)
Release day : 2026-09-08

#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # Wipe the state file before running. Every open window alerts again.
    [switch] $ResetState
)

# Resolve the folder this script lives in. $PSScriptRoot is empty when the code is
# run with F8 / "Run Selection" in ISE or VS Code, or pasted into a console, which
# turns every path into "\Files\..." on the current drive. Fall back to the script
# path, then to the current folder.
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }


#------------------------------- Settings -------------------------------

# Filename and Folderspath
 $scriptName   = $MyInvocation.MyCommand.Name # Scriptname
 $logFolder    = "$scriptRoot\Logs" # Log Path
 $reportFolder = "$scriptRoot\Files\report" # Report Path
 $stateFolder  = "$scriptRoot\Files\state" # State Path

# Monitoring objects, Configration and State file path
 $monitoringObjectsPath = "$scriptRoot\Files\db\monitorobjects.json"
 $exchangeSettingsPath  = "$scriptRoot\Settings\Config\ExchangeOnpremSettings.json"
 $statePath             = "$stateFolder\sent-state.json"
 $scriptSettingsPath = "$scriptRoot\Settings\Config\version.json"

# backup paths for monitorobjects.json, sent-state.json
  $backupMonitorObjectPath = "$scriptRoot\Files\backup\psToDo\"
  $backupSendStateJsonPath = "$scriptRoot\Files\backup\psToDo\"
  $backupExchangeSettingsPath = "$scriptRoot\Files\backup\psToDo\"

# Ensure the folders the script writes to exist (first-run safe)
 foreach ($dir in @($logFolder, $stateFolder, $reportFolder, $backupMonitorObjectPath)) {
     if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
 }

# The three trigger keys, in the order they appear in the JSON
 $triggerKeys = @('1dateTrigger', '2dateTrigger', '3dateTrigger')

# Date format used inside monitorobjects.json
 $expireDateFormat = 'yyyy-MM-dd'
 $today            = (Get-Date).Date

# How often to notify about objects that have already expired
$expiredNotifyDays = 7

# Transcript
$logFileDate       = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
$tranScriptLogFile = "$logFolder\$scriptName - $logFileDate.txt"

Start-Transcript -Path $tranScriptLogFile -Force | Out-Null
Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
Write-Host '.. Starting TranScript'

# Error settings
$ErrorActionPreference = 'Continue'

#------------------------------- Functions List -------------------------------

Write-Host 'Checking required Functions...' -ForegroundColor Yellow

$functionFolder = "$scriptRoot\Settings\Functions\psToDo"
if (-not (Test-Path $functionFolder)) { throw "Function folder not found: $functionFolder"}
$functionFiles = Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File
if (-not $functionFiles) { throw "No .ps1 files found in $functionFolder"}

foreach ($file in $functionFiles) 
{ 
 try {
        . $file.FullName
        Write-Host "- Loaded $($file.Name)" -ForegroundColor DarkGray
    }
catch { throw "Failed to load function file '$($file.Name)': $($_.Exception.Message)" }
}
Write-Host "All $($functionFiles.Count) function file(s) loaded." -ForegroundColor Green

# On-prem edition: no external modules are needed. System.Net.Mail ships with .NET.

Write-host ""
#------------------------------- Section 1 -------------------------------
$Section = 'Section 1 : Read the monitorobjects.json file and the sent state'
try {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "Start $Section ... 0%" -ForegroundColor Yellow
    if (-not (Test-Path $monitoringObjectsPath)) { throw "Cannot find monitoring objects file: $monitoringObjectsPath"}
    try { 
         $monitoringObjects = Get-Content -Raw -Encoding UTF8 $monitoringObjectsPath | ConvertFrom-Json
         Write-host ""
   
        }
    catch { throw "Failed to parse JSON in '$monitoringObjectsPath': $($_.Exception.Message)"}
    if (-not $monitoringObjects) { throw 'The monitoring objects file is empty or contains no objects.'}
    Write-Host "Loaded $(@($monitoringObjects).Count) object(s)."

    if ($ResetState -and (Test-Path $statePath)) {
        if ($PSCmdlet.ShouldProcess($statePath, 'Delete state file')) {
            Remove-Item $statePath -Force
            Write-Warning 'State file deleted. Every open trigger window will alert again.'
        }
    }

    $state = Get-SentState -Path $statePath

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "ERROR on $Section" -ForegroundColor Red
    Write-Host 'ERROR:' $_.Exception.Message
    Write-Host 'Stopping Transcript and Script!' -ForegroundColor Red
    Stop-Transcript
    exit 1
}


#------------------------------- Section 2 -------------------------------
$Section = 'Section 2 : Find objects inside a trigger window that has not alerted yet'
try {
      Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
      Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

      $alerts    = [System.Collections.Generic.List[object]]::new()
      $liveKeys  = [System.Collections.Generic.List[string]]::new()
      foreach ($object in $monitoringObjects) 
       {
         $label = "Object $($object.id) - $($object.name)"
        # --- Validate and parse expireDate ---
        if ([string]::IsNullOrWhiteSpace($object.expireDate)) {
         Write-Warning "--> $label is missing expireDate - skipping."
         continue
        }

        try {
             $expireDate = [datetime]::ParseExact(
             $object.expireDate,
             $expireDateFormat,
             [cultureinfo]::InvariantCulture
            ).Date
        }
        catch {
            Write-Warning "--> $label has an invalid expireDate: '$($object.expireDate)' (expected $expireDateFormat) - skipping."
            continue
        }

        $daysLeft = ($expireDate - $today).Days

        # Prefix for every state key belonging to this object at this expiry date.
        # Renewing the changes the prefix, retiring the old keys.
        $keyPrefix = "$($object.id)_$($object.expireDate)"
        $liveKeys.Add($keyPrefix)

        # --- Collect triggers. Quoted access: the keys start with a digit. ---
        $rawTriggers = foreach ($key in $triggerKeys) { $object.$key }

        [int[]] $triggers = $rawTriggers |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                            ForEach-Object { [int]$_ } |
                            Sort-Object -Unique

        if (-not $triggers) {
            Write-Warning "--> $label has no valid date triggers - skipping."
            continue
        }

        # ascending order      
        $declared = @($rawTriggers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [int]$_ })
        if (($declared -join ',') -ne (($declared | Sort-Object) -join ',')) {
            Write-Warning "--> $label declares triggers out of order ($($declared -join ', ')). Sorted automatically." # order-independent, but a mismatch is flagging.
        }

        # --- The narrowest open window. Triggers are ascending, so the first
        $hit = $triggers | Where-Object { $daysLeft -le $_ -and $daysLeft -ge 0 } | Select-Object -First 1  #  match is the smallest trigger that still covers daysLeft. ---
        $expiredNotifyDay = ($daysLeft -lt 0) -and ($daysLeft % $expiredNotifyDays -eq 0)
        $expiryDay  = ($daysLeft -eq 0)

        # Expiry-day and expired reminders are time based, not window based
        $stateKey = if     ($expiryDay)       { "${keyPrefix}_expiryday" }
                    elseif ($daysLeft -lt 0)  { "${keyPrefix}_expired-$($today.ToString('yyyy-MM-dd'))" }
                    elseif ($null -ne $hit)   { "${keyPrefix}_$hit" }
                    else                      { $null }

        if ($null -eq $stateKey) { Write-Host "  $label - $daysLeft days left, no trigger window open." -ForegroundColor DarkGray
            continue }

        if (-not $expiryDay -and $daysLeft -lt 0 -and -not $expiredNotifyDay) { Write-Host "  $label - expired $([Math]::Abs($daysLeft)) days ago, not a reminder day." -ForegroundColor DarkGray
            continue}

        if ($state.ContainsKey($stateKey)) {
            $when = try { ([datetime]$state[$stateKey]).ToString('yyyy-MM-dd HH:mm') } catch { 'earlier' }
            Write-Host "  $label - $daysLeft days left, already alerted ($stateKey) on $when." -ForegroundColor DarkGray
            continue
        }

        # Smallest trigger sits closest to expiry, so it is the most urgent
        $urgency = if     ($daysLeft -lt 0)            { 'expired'  }
                   elseif ($daysLeft -le $triggers[0]) { 'critical' }
                   elseif ($triggers.Count -ge 2 -and $daysLeft -le $triggers[1]) { 'warning' }
                   else                                { 'ok'       }

        $alerts.Add([pscustomobject]@{
            id             = "$($object.id)"
            name           = "$($object.name)"
            servername     = "$($object.servername)"
            template       = "$($object.template)"
            environment    = "$($object.environment)"
            action         = "$($object.description)"
            expireDate     = $expireDate
            daysLeft       = $daysLeft
            trigger        = $hit
            allTriggers    = @($triggers)
            urgency        = $urgency
            stateKey       = $stateKey

            mailOn         = [bool]$object.notifyMethodbyMail
            mailSender     = "$($object.mail.mailSender)"
            mailSubject    = "$($object.mail.mailSubject)"
            mailBody       = "$($object.mail.mailBody)"
            mailRecipients = @($object.mail.mailRecipients)

            teamsOn        = [bool]$object.notifyMethodbyTeams
            teamSubject    = "$($object.teams.teamSubject)"
            teamBody       = "$($object.teams.teamBody)"
            teamWebhookUrl = "$($object.teams.teamWebhookUrl)"
        })

        $why = if     ($expiryDay)      { 'expiry day'        }
               elseif ($daysLeft -lt 0) { 'expired reminder'  }
               else                     { "$hit-day window"   }

        Write-Host "  $($urgency.ToUpper()): $label - $daysLeft days left ($why)." -ForegroundColor Red
    }

    Write-Host ''
    Write-Host "$($alerts.Count) object(s) need an alert today." -ForegroundColor Yellow
    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "ERROR on $Section" -ForegroundColor Red
    Write-Host 'ERROR:' $_.Exception.Message
    Write-Host 'Stopping Transcript and Script!' -ForegroundColor Red
    Stop-Transcript
    exit 1
}

#------------------------------- Section 3 -------------------------------

$Section = 'Section 3 : Connect to the on-prem Exchange server (SMTP)'

# Mail is the only channel in the on-prem edition, so a connection is only
# needed when at least one alert has mail enabled.
$needsSmtp = @($alerts | Where-Object mailOn).Count -gt 0

$teamsOnly = @($alerts | Where-Object { $_.teamsOn -and -not $_.mailOn }).Count
if ($teamsOnly) { Write-Warning "$teamsOnly object(s) have ONLY Teams enabled. Teams is not available in the on-prem edition - enable notifyMethodbyMail for them." }

if (-not $needsSmtp) {
    Write-Host "Skipping $Section - no mail to send today." -ForegroundColor DarkGray
    Write-Host ''
}
else {
    try {
        Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
        Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

        $settings = Get-OnpremExchangeSettings -Path $exchangeSettingsPath -RepoRoot $scriptRoot

        Connect-OnpremExchangeSmtp -Settings $settings | Out-Null

        Write-Host "Start $Section ... 100%" -ForegroundColor Green
        Write-Host ''
    }
    catch {
        Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host 'ERROR:' $_.Exception.Message
        Write-Host 'Cannot send mail without a connection to the on-prem Exchange server. Stopping.' -ForegroundColor Red
        Stop-Transcript
        exit 1
    }
}

#------------------------------- Section 4 -------------------------------

$Section = 'Section 4 : Send alerts on the channels each object asks for'

if ($alerts.Count) {
    try {
        Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
        Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

        $results = [System.Collections.Generic.List[object]]::new()
        $groups = $alerts | Group-Object trigger | Sort-Object { [int]($_.Name -as [int]) }
        foreach ($group in $groups) 
        {
            $heading = if ($group.Name) { "$($group.Name)-day window" } else { 'Expired / expiry day' }
            Write-Host "-- $heading ($($group.Count) object(s))" -ForegroundColor Cyan
            foreach ($alert in $group.Group) 
            {
                Write-Host "   $($alert.id) - $($alert.name) [$($alert.urgency)]" -ForegroundColor Cyan
                $r = Send-AlertNotification -Alert $alert
                $results.Add($r)
                if (-not $WhatIfPreference -and ($r.MailSent -or $r.TeamsSent)) { $state[$alert.stateKey] = (Get-Date).ToString('o') }# Record only what actually went out. A failed send stays unrecorded, so the next run tries again. -WhatIf never records.
            }
            Write-Host ''
        }

        $mailSent   = @($results | Where-Object MailSent).Count
        $skipped    = @($results | Where-Object Skipped).Count
        $mailTried  = @($alerts  | Where-Object mailOn).Count

        Write-Host "Mail:  $mailSent of $mailTried sent (via on-prem Exchange)" -ForegroundColor Yellow

        if ($skipped) { Write-Warning "$skipped object(s) had mail disabled and were not notified (Teams is not available on-prem)."}

        Write-Host "Start $Section ... 100%" -ForegroundColor Green
        Write-Host ''
    }
    catch 
    {
        Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host 'ERROR:' $_.Exception.Message
    }
    finally { Disconnect-OnpremExchangeSmtp }
}
else {
    Write-Host "Skipping $Section - nothing to send today." -ForegroundColor DarkGray
    Write-Host ''
}

#------------------------------- Section 5 -------------------------------

$Section = 'Section 5 : Save the sent state'
try {
     Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
     Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

    if ($WhatIfPreference) { Write-Host 'WhatIf: the state file was not written.' -ForegroundColor DarkGray }
    else { Save-SentState -State $state -Path $statePath -LiveKeyPrefixes $liveKeys }

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "ERROR on $Section" -ForegroundColor Red
    Write-Host 'ERROR:' $_.Exception.Message
}

#------------------------------- Section 6 -------------------------------

$Section = 'Section 6 : Create backup copies of monitorobjects.json, sent-state.json and ExchangeOnpremSettings.json'
try {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

    # Create a backup of the monitorobjects.json and sent-state.json files with a timestamp in the filename.
    # Each copy is guarded: a file that does not exist yet (e.g. sent-state.json on a first run) is skipped, not fatal.
    $backupfileDate = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    if (-not (Test-Path $backupMonitorObjectPath)) { New-Item -ItemType Directory -Path $backupMonitorObjectPath -Force | Out-Null }

    if (Test-Path $monitoringObjectsPath) { Copy-Item -Path $monitoringObjectsPath -Destination "$backupMonitorObjectPath\monitorobjects-$($backupfileDate).json" -Force -Verbose }
    else { Write-Warning "Skipped backup: $monitoringObjectsPath not found." }

    if (Test-Path $statePath) { Copy-Item -Path $statePath -Destination "$backupSendStateJsonPath\sendstate-$($backupfileDate).json" -Force -Verbose }
    else { Write-Host "No state file to back up yet ($statePath)." -ForegroundColor DarkGray }

    if (Test-Path $exchangeSettingsPath) { Copy-Item -Path $exchangeSettingsPath -Destination "$backupExchangeSettingsPath\ExchangeOnpremSettings-$($backupfileDate).json" -Force -Verbose }
    else { Write-Warning "Skipped backup: $exchangeSettingsPath not found." }

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "ERROR on $Section" -ForegroundColor Red
    Write-Host 'ERROR:' $_.Exception.Message
}


#------------------------------- End -------------------------------

Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
Write-Host 'Script finished.' -ForegroundColor Green
Stop-Transcript
