<#
.DESCRIPTION
    SYSTEM REQUIREMENTS : Windows PowerShell 5.1 or PowerShell 7.x. No modules.

    psToDo (ON-PREMISES EDITION) - status dashboard.

    Reads monitorobjects.json and sent-state.json and writes index.html: a self-contained
    status page for IIS (no external files, works on an isolated network).

    What the page shows
      - Header with your organisation name and logo (Settings\Config\report-settings.json).
      - Key figures, and a warning when the report is older than StaleAfterHours
        (= the scheduled report run has stopped).
      - A timeline of everything that expires in the next TimelineDays days.
      - Tables per tab: Monitored, Completed, No channel, No credentials, Not monitored.
        Columns: owner, expiry, days left, urgency, NEXT alert (date + channel) and
        LAST alert (from sent-state.json).
      - A detail view per object with its alert history: sent / due / waiting / skipped.
      - CSV export, print view, light and dark theme following the Windows setting,
        English or Swedish (Language in report-settings.json).

    Urgency follows the alert script: triggers are sorted ascending and the smallest is
    the most urgent because it sits closest to expiry.
        daysLeft < 0              -> expired
        daysLeft <= smallest      -> critical
        daysLeft <= second        -> warning
        otherwise                 -> ok

    Completed objects are COPIED (not moved) to Files\config\MonitorObjectComplete.json,
    and the Completed tab is built from that file, exactly as before.

.EXAMPLE
    .\psToDo-HTML-Report-Onprem-v1.0.ps1
    .\psToDo-HTML-Report-Onprem-v1.0.ps1 -Language sv -OrganizationName 'Contoso IT'
    .\psToDo-HTML-Report-Onprem-v1.0.ps1 -websitePath 'C:\inetpub\pstodo\index.html'

.NOTES
    Author      : Fardin Barashi
    Title       : psToDo Calendar Reminder - HTML report, ON-PREMISES EDITION
    Version     : 1.1-onprem (based on psToDo 1.3)
    Release day : 2026-11-06
    Github Link : https://github.com/fardinbarashi
    File name kept at -v1.0 so existing scheduled tasks keep working.
#>

#----------------------------------- Settings ------------------------------------------
[CmdletBinding()]
param(
    [string] $JsonPath = '',
    [string] $CompleteJsonPath = '',
    [string] $StatePath = '',
    [string] $ScriptSettingsPath = '',
    [string] $ReportSettingsPath = '',
    [string] $TemplatePath = '',
    [string] $websitePath = '',
    [string] $backupWebsitePath = '',
    [string] $Language = '',
    [string] $OrganizationName = '',
    [string] $LogoPath = '',
    [string] $DateFormat = 'yyyy-MM-dd',
    [string] $fileDate = (Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'),
    [string[]] $TriggerKeys = @('1dateTrigger', '2dateTrigger', '3dateTrigger')
)

# Resolve the folder this script lives in. $PSScriptRoot is empty when the code is
# run with F8 / "Run Selection" in ISE or VS Code, or pasted into a console.
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

# Windows PowerShell 5.1: its built-in type data for System.Array makes ConvertTo-Json
# write arrays as {"value":[...],"Count":n}. Removing it gives plain JSON arrays, like PS 7.
Remove-TypeData -TypeName System.Array -ErrorAction SilentlyContinue

if ([string]::IsNullOrWhiteSpace($JsonPath))           { $JsonPath           = "$scriptRoot\Files\db\monitorobjects.json" }
if ([string]::IsNullOrWhiteSpace($CompleteJsonPath))   { $CompleteJsonPath   = "$scriptRoot\Files\config\MonitorObjectComplete.json" }
if ([string]::IsNullOrWhiteSpace($StatePath))          { $StatePath          = "$scriptRoot\Files\state\sent-state.json" }
if ([string]::IsNullOrWhiteSpace($ScriptSettingsPath)) { $ScriptSettingsPath = "$scriptRoot\Settings\Config\version.json" }
if ([string]::IsNullOrWhiteSpace($ReportSettingsPath)) { $ReportSettingsPath = "$scriptRoot\Settings\Config\report-settings.json" }
if ([string]::IsNullOrWhiteSpace($TemplatePath))       { $TemplatePath       = "$scriptRoot\Settings\Web\report-template.html" }
if ([string]::IsNullOrWhiteSpace($websitePath))        { $websitePath        = "$scriptRoot\Files\report\index.html" }
if ([string]::IsNullOrWhiteSpace($backupWebsitePath))  { $backupWebsitePath  = "$scriptRoot\Files\backup\psToDo-HTML-Report\index-$($fileDate).html" }

# .NET file calls ([IO.File]) resolve relative paths against the PROCESS folder, not the
# PowerShell location. Make every path absolute first, so -websitePath .\x.html works.
$JsonPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($JsonPath)
$CompleteJsonPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CompleteJsonPath)
$StatePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($StatePath)
$ScriptSettingsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ScriptSettingsPath)
$ReportSettingsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportSettingsPath)
$TemplatePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TemplatePath)
$websitePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($websitePath)
$backupWebsitePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($backupWebsitePath)

# Transcript
$ScriptName        = $MyInvocation.MyCommand.Name
$LogFileDate       = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'   # no '/' - it is a path separator
$TranScriptLogFile = "$scriptRoot\Logs\$ScriptName - $LogFileDate.txt"
Start-Transcript -Path $TranScriptLogFile -Force | Out-Null
Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
Write-Host '.. Starting TranScript'

$ErrorActionPreference = 'Continue'

#------------------------------- Functions List -------------------------------

Write-Host 'Checking required Functions...' -ForegroundColor Yellow
$functionFolder = "$scriptRoot\Settings\Functions\psToDo-HtmlReport"
if (-not (Test-Path $functionFolder)) { throw "Function folder not found: $functionFolder" }
$functionFiles = Get-ChildItem -Path $functionFolder -Filter '*.ps1' -File
if (-not $functionFiles) { throw "No .ps1 files found in $functionFolder" }
foreach ($file in $functionFiles) {
    try   { . $file.FullName; Write-Host "- Loaded $($file.Name)" -ForegroundColor DarkGray }
    catch { throw "Failed to load function file '$($file.Name)': $($_.Exception.Message)" }
}
Write-Host "All $($functionFiles.Count) function file(s) loaded." -ForegroundColor Green
Write-Host ''

# Turn a value from ConvertFrom-Json into an ISO 8601 string (PS 7 turns ISO strings into DateTime).
function ConvertTo-IsoText {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToString('o') }
    return "$Value"
}

function Stop-OnError {
    param([string] $Section, $ErrorRecord)
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "ERROR on $Section" -ForegroundColor Red
    Write-Host 'ERROR:' $ErrorRecord.Exception.Message
    Write-Host 'Stopping Transcript and Script!' -ForegroundColor Red
    Stop-Transcript
    exit 1
}

#----------------------------------- Section 1 ------------------------------------------
$Section = 'Section 1 : Read settings, monitorobjects.json, sent state and archive completed objects'
try {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

    # --- report settings (file, then parameters override)
    $rs = $null
    if (Test-Path $ReportSettingsPath) {
        try   { $rs = Get-Content -Raw -Encoding UTF8 $ReportSettingsPath | ConvertFrom-Json }
        catch { Write-Warning "Could not read $ReportSettingsPath - using defaults. $($_.Exception.Message)" }
    }
    function Get-Setting($name, $default) {
        if ($rs -and $rs.PSObject.Properties.Name -contains $name -and "$($rs.$name)" -ne '') { return $rs.$name }
        return $default
    }
    if ([string]::IsNullOrWhiteSpace($Language))         { $Language         = [string](Get-Setting 'Language' 'en') }
    if ([string]::IsNullOrWhiteSpace($OrganizationName)) { $OrganizationName = [string](Get-Setting 'OrganizationName' '') }
    if ([string]::IsNullOrWhiteSpace($LogoPath))         { $LogoPath         = [string](Get-Setting 'LogoPath' '') }
    $staleHours        = [int](Get-Setting 'StaleAfterHours' 24)
    $timelineDays      = [int](Get-Setting 'TimelineDays' 90)
    $expiredNotifyDays = [int](Get-Setting 'ExpiredNotifyDays' 7)
    $Language = $Language.Trim().ToLowerInvariant()
    if ($Language -notin 'en', 'sv') { Write-Warning "Language '$Language' is not supported (en, sv) - using en."; $Language = 'en' }

    # --- logo, embedded so the page needs no other file
    $logo = ''
    if (-not [string]::IsNullOrWhiteSpace($LogoPath)) {
        $logoFull = if ([System.IO.Path]::IsPathRooted($LogoPath)) { $LogoPath } else { [System.IO.Path]::Combine($scriptRoot, $LogoPath) }
        if (-not (Test-Path $logoFull)) { Write-Warning "Logo not found: $logoFull" }
        else {
            $bytes = [System.IO.File]::ReadAllBytes($logoFull)
            $mime = switch ([System.IO.Path]::GetExtension($logoFull).ToLowerInvariant()) {
                '.png' { 'image/png' } '.jpg' { 'image/jpeg' } '.jpeg' { 'image/jpeg' } '.gif' { 'image/gif' } '.svg' { 'image/svg+xml' } default { '' } }
            if (-not $mime) { Write-Warning "Logo must be .png, .jpg, .gif or .svg: $logoFull" }
            elseif ($bytes.Length -gt 512KB) { Write-Warning "Logo is $([int]($bytes.Length / 1KB)) KB; keep it under 512 KB (it is embedded in the page). Skipped." }
            else { $logo = "data:$mime;base64," + [Convert]::ToBase64String($bytes) }
        }
    }

    # --- version
    $version = ''
    if (Test-Path $ScriptSettingsPath) {
        try { $version = "$((Get-Content -Raw -Encoding UTF8 $ScriptSettingsPath | ConvertFrom-Json).version)" } catch { }
    }

    # --- objects
    if (-not (Test-Path $JsonPath)) { throw "Cannot find the monitoring objects file: $JsonPath" }
    try   { $monitoringObjects = @((Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json)) }
    catch { throw "Failed to parse JSON in '$JsonPath': $($_.Exception.Message)" }
    if (-not $monitoringObjects) { throw 'The monitoring objects file is empty or contains no objects.' }
    Write-Host "- $($monitoringObjects.Count) object(s) read." -ForegroundColor DarkGray

    # --- sent state (what the alert script has already sent)
    $state = @{}
    if (Test-Path $StatePath) {
        try {
            $rawState = Get-Content -Raw -Encoding UTF8 $StatePath
            if (-not [string]::IsNullOrWhiteSpace($rawState)) {
                ($rawState | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $state[$_.Name] = ConvertTo-IsoText $_.Value }
            }
            Write-Host "- Sent state: $($state.Count) alert(s) recorded." -ForegroundColor DarkGray
        }
        catch { Write-Warning "Could not read $StatePath - 'last alert' will be empty. $($_.Exception.Message)" }
    }
    else { Write-Host "- No sent state yet ($StatePath)." -ForegroundColor DarkGray }

    # --- completed objects are COPIED to MonitorObjectComplete.json (unchanged behaviour)
    $completedObjectsRaw = @($monitoringObjects | Where-Object { "$($_.status)" -ieq 'Completed' })
    $monitoredObjectsRaw = @($monitoringObjects | Where-Object { "$($_.status)" -ine 'Completed' })

    function Get-CompleteKey {
        param($o)
        foreach ($k in @($o.exchangeCertKey, $o.adAccountKey, $o.entraKeyId, $o.messageId, $o.appId)) {
            if (-not [string]::IsNullOrWhiteSpace($k)) { return "$k" }
        }
        return ('{0}|{1}|{2}' -f $o.name, $o.servername, $o.expireDate)
    }

    $existingCopyDates = @{}
    if (Test-Path $CompleteJsonPath) {
        try {
            $prevRaw = Get-Content -Raw -Encoding UTF8 $CompleteJsonPath
            $prev    = if ([string]::IsNullOrWhiteSpace($prevRaw)) { @() } else { @(($prevRaw | ConvertFrom-Json)) }
            foreach ($p in $prev) {
                if (-not [string]::IsNullOrWhiteSpace($p.copiedDate)) { $existingCopyDates[(Get-CompleteKey $p)] = "$($p.copiedDate)" }
            }
        } catch { Write-Warning "Could not read existing $CompleteJsonPath - copy dates will be reset. $($_.Exception.Message)" }
    }
    $copyStamp = Get-Date -Format 'yyyy-MM-dd'
    foreach ($o in $completedObjectsRaw) {
        $key  = Get-CompleteKey $o
        $date = if ($existingCopyDates.ContainsKey($key)) { $existingCopyDates[$key] } else { $copyStamp }
        $o | Add-Member -NotePropertyName copiedDate -NotePropertyValue $date -Force
    }
    $completeDir = Split-Path $CompleteJsonPath -Parent
    if ($completeDir -and -not (Test-Path $completeDir)) { New-Item -ItemType Directory -Force -Path $completeDir | Out-Null }
    $completeJson = if ($completedObjectsRaw.Count) { ConvertTo-Json -InputObject @($completedObjectsRaw) -Depth 20 } else { '[]' }
    Set-Content -Encoding UTF8 -Path $CompleteJsonPath -Value $completeJson
    Write-Host "- Copied $($completedObjectsRaw.Count) completed object(s) to $CompleteJsonPath" -ForegroundColor DarkGray

    $completeRaw = if (Test-Path $CompleteJsonPath) { Get-Content -Raw -Encoding UTF8 $CompleteJsonPath } else { $null }
    $completedObjects = if ([string]::IsNullOrWhiteSpace($completeRaw)) { @() } else { @(($completeRaw | ConvertFrom-Json)) }

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch { Stop-OnError $Section $_ }

#----------------------------------- Section 2 ------------------------------------------
$Section = 'Section 2 : Build the report data'
try {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

    $today         = (Get-Date).Date
    $rows          = [System.Collections.Generic.List[object]]::new()
    $completedRows = [System.Collections.Generic.List[object]]::new()
    $invalid       = [System.Collections.Generic.List[object]]::new()
    $noCreds       = [System.Collections.Generic.List[object]]::new()
    $placeholder   = 0

    # One display row. Objects without a usable date or triggers get an "inventory" row
    # (shown, never alerts) - used for Completed objects.
    function New-ReportRow {
        param($Object, [switch] $Completed)

        $base = [ordered]@{
            id         = "$($Object.id)"
            name       = "$($Object.name)"
            owner      = "$($Object.owner)"
            status     = $(if ([string]::IsNullOrWhiteSpace($Object.status)) { $(if ($Completed) { 'Completed' } else { 'Backlog' }) } else { "$($Object.status)" })
            template   = "$($Object.template)"
            server     = "$($Object.servername)"
            env        = "$($Object.environment)"
            desc       = "$($Object.description)"
            source     = "$($Object.source)"
            severity   = "$($Object.severity)"
            link       = "$($Object.messageLink)"
            copied     = "$($Object.copiedDate)"
            expire     = ''
            days       = $null
            triggers   = @()
            urgency    = 'none'
            mailOn     = [bool](Test-NotifyFlag $Object.notifyMethodbyMail)
            sender     = "$($Object.mail.mailSender)"
            recipients = @(@($Object.mail.mailRecipients) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { "$_" })
            subject    = "$($Object.mail.mailSubject)"
            schedule   = @()
            last       = $null
            next       = $null
        }

        if ([string]::IsNullOrWhiteSpace($Object.expireDate)) { return $base }
        try   { $expireDate = [datetime]::ParseExact("$($Object.expireDate)", $DateFormat, [cultureinfo]::InvariantCulture).Date }
        catch { return $base }

        $rawTriggers = foreach ($key in $TriggerKeys) { $Object.$key }
        [int[]] $triggers = @($rawTriggers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [int]$_ } | Sort-Object -Unique)

        $base.expire = $expireDate.ToString('yyyy-MM-dd')
        $base.days   = ($expireDate - $today).Days
        if (-not $triggers) { return $base }

        $base.triggers = @($triggers)
        $base.urgency  = Get-Urgency -DaysLeft $base.days -Triggers $triggers

        $plan = Get-AlertSchedule -Id "$($Object.id)" -ExpireText "$($Object.expireDate)" -ExpireDate $expireDate `
                                  -Triggers $triggers -Today $today -State $state -ExpiredNotifyDays $expiredNotifyDays
        $base.schedule = @($plan.Schedule)
        $base.last     = $plan.Last
        $base.next     = $plan.Next
        return $base
    }

    foreach ($object in $monitoredObjectsRaw) {
        $label = "Object $($object.id) - $($object.name)"

        # App registrations with no secret and no certificate get their own tab.
        if ($object.noCredentials -eq $true -or ($object.source -eq 'EntraAppReg' -and [string]::IsNullOrWhiteSpace($object.expireDate))) {
            $noCreds.Add([ordered]@{
                id = "$($object.id)"; name = "$($object.name)"; owner = "$($object.owner)"; server = "$($object.servername)"
                env = "$($object.environment)"; status = $(if ([string]::IsNullOrWhiteSpace($object.status)) { 'Backlog' } else { "$($object.status)" })
                desc = "$($object.description)"; source = "$($object.source)"
            })
            continue
        }

        if ([string]::IsNullOrWhiteSpace($object.expireDate)) {
            Write-Warning "$label is missing expireDate - not monitored."
            $invalid.Add([ordered]@{ id = "$($object.id)"; name = "$($object.name)"; owner = "$($object.owner)"; reason = 'missing-date'; value = '' })
            continue
        }
        $parsed = [datetime]::MinValue
        if (-not [datetime]::TryParseExact("$($object.expireDate)", $DateFormat, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref] $parsed)) {
            Write-Warning "$label has an invalid expireDate: '$($object.expireDate)' (expected $DateFormat) - not monitored."
            $invalid.Add([ordered]@{ id = "$($object.id)"; name = "$($object.name)"; owner = "$($object.owner)"; reason = 'bad-date'; value = "$($object.expireDate)" })
            continue
        }

        $rawTriggers = foreach ($key in $TriggerKeys) { $object.$key }
        $declared = @($rawTriggers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [int]$_ })
        if (-not $declared) {
            Write-Warning "$label has no valid date triggers - not monitored."
            $invalid.Add([ordered]@{ id = "$($object.id)"; name = "$($object.name)"; owner = "$($object.owner)"; reason = 'no-triggers'; value = '' })
            continue
        }
        if (($declared -join ',') -ne (($declared | Sort-Object) -join ',')) { Write-Warning "$label declares triggers out of order ($($declared -join ', ')). Sorted automatically." }

        if (-not (Test-NotifyFlag $object.notifyMethodbyMail) -and "$($object.notifyMethodbyMail)" -match 'True or null') { $placeholder++ }

        $rows.Add((New-ReportRow $object))
    }

    foreach ($object in $completedObjects) { $completedRows.Add((New-ReportRow $object -Completed)) }

    if (-not ($rows.Count + $completedRows.Count + $noCreds.Count + $invalid.Count)) { throw 'No valid objects to report on. Check the warnings above.' }
    Write-Host "Processed $($rows.Count) monitored + $($completedRows.Count) completed object(s)." -ForegroundColor Green
    $rows | Group-Object { $_.urgency } | Sort-Object Name | ForEach-Object { Write-Host ("  {0,-9} {1}" -f $_.Name, $_.Count) }
    if ($invalid.Count) { Write-Warning "$($invalid.Count) object(s) are NOT being monitored." }
    if ($placeholder)   { Write-Warning "$placeholder notify field(s) still hold the placeholder string. Replace them with true / false." }

    $completeDisplay = $CompleteJsonPath
    if ($completeDisplay.StartsWith($scriptRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $completeDisplay = $completeDisplay.Substring($scriptRoot.Length).TrimStart('\', '/')
    }

    $editExists = Test-Path ([System.IO.Path]::Combine((Split-Path $websitePath -Parent), 'edit.html'))

    $data = [ordered]@{
        meta = [ordered]@{
            org               = $OrganizationName
            logo              = $logo
            lang              = $Language
            version           = $version
            generated         = (Get-Date).ToString('o')
            staleHours        = $staleHours
            timelineDays      = $timelineDays
            expiredNotifyDays = $expiredNotifyDays
            completeFile      = $completeDisplay
            editLink          = [bool]$editExists
            channel           = 'mail'
        }
        monitored = @($rows)
        completed = @($completedRows)
        noCreds   = @($noCreds)
        invalid   = @($invalid)
    }
    $dataJson = ConvertTo-Json -InputObject $data -Depth 10 -Compress
    $dataB64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($dataJson))

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch { Stop-OnError $Section $_ }

#----------------------------------- Section 3 ------------------------------------------
$Section = 'Section 3 : Build index.html from the template'
try {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

    if (-not (Test-Path $TemplatePath)) { throw "Template not found: $TemplatePath" }
    $template = Get-Content -Raw -Encoding UTF8 $TemplatePath
    foreach ($token in '__PSTODO_DATA_B64__', '__PSTODO_LANG__') {
        if ($template.IndexOf($token) -lt 0) { throw "The template is missing the placeholder $token." }
    }
    # .Replace (literal), not -replace (regex): the data must never be read as a pattern.
    $html = $template.Replace('__PSTODO_DATA_B64__', $dataB64)
    $html = $html.Replace('__PSTODO_LANG__', $Language)

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch { Stop-OnError $Section $_ }

#----------------------------------- Section 4 ------------------------------------------
$Section = 'Section 4 : Write index.html and a backup'
try {
    Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    Write-Host "Start $Section ... 0%" -ForegroundColor Yellow

    $outDir = Split-Path $websitePath -Parent
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
    [System.IO.File]::WriteAllText($websitePath, $html, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Report written to $websitePath" -ForegroundColor Green

    $backupDir = Split-Path $backupWebsitePath -Parent
    if ($backupDir -and -not (Test-Path $backupDir)) { New-Item -Path $backupDir -ItemType Directory -Force | Out-Null }
    Copy-Item -Path $websitePath -Destination $backupWebsitePath -Force
    Write-Host "Backup written to $backupWebsitePath" -ForegroundColor DarkGray

    Write-Host "Start $Section ... 100%" -ForegroundColor Green
    Write-Host ''
}
catch { Stop-OnError $Section $_ }

#----------------------------------- End Script ------------------------------------------
Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
Write-Host 'Script finished.' -ForegroundColor Green
Stop-Transcript
