function Get-AlertSchedule {
<#
    Works out, for ONE object, what the alert script (psToDo-Onprem) has sent and
    will send - using exactly the same rules and the same keys in sent-state.json.

    Rules of the alert script, run once a day:
      daysLeft = 0          -> "expiry day" alert          key <id>_<expireDate>_expiryday
      daysLeft < 0          -> reminder every N days        key <id>_<expireDate>_expired-<yyyy-MM-dd>
                               (when -daysLeft is a multiple of N, N = $ExpiredNotifyDays)
      0 < daysLeft <= t     -> the SMALLEST trigger t that still covers daysLeft
                               key <id>_<expireDate>_<t>
    So window t is in effect while daysLeft is in (previous smaller trigger, t].
    A trigger of 0 never fires on its own: day 0 is the expiry day.

    Every window gets a state:
      sent     - its key is in sent-state.json (sentAt = when)
      due      - the window is open now and has not been sent: the next run sends it
      pending  - opens later (date = the day it opens)
      skipped  - it closed without being sent (object added late, or the job was not running)

    Returns @{ Schedule; Last; Next } with dates as yyyy-MM-dd strings and times as ISO 8601.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]   $Id,
        [Parameter(Mandatory)] [string]   $ExpireText,      # the raw expireDate string - part of the state key
        [Parameter(Mandatory)] [datetime] $ExpireDate,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [int[]] $Triggers,   # ascending, unique
        [Parameter(Mandatory)] [datetime] $Today,
        [Parameter(Mandatory)] [hashtable] $State,
        [int] $ExpiredNotifyDays = 7
    )

    $fmt      = 'yyyy-MM-dd'
    $prefix   = "$($Id)_$($ExpireText)"
    $daysLeft = ($ExpireDate - $Today).Days
    $items    = [System.Collections.Generic.List[object]]::new()

    $stateOf = {
        param($key, [datetime] $from, [datetime] $to)
        if ($State.ContainsKey($key)) { return 'sent' }
        if ($Today -lt $from)         { return 'pending' }
        if ($Today -le $to)           { return 'due' }
        return 'skipped'
    }

    # --- trigger windows
    $effective = @($Triggers | Where-Object { $_ -gt 0 } | Sort-Object)
    for ($i = 0; $i -lt $effective.Count; $i++) {
        $t     = $effective[$i]
        $prevT = if ($i -eq 0) { 0 } else { $effective[$i - 1] }
        $open  = $ExpireDate.AddDays(-$t)
        $close = $ExpireDate.AddDays(-($prevT + 1))
        $key   = "$($prefix)_$t"
        $items.Add([ordered]@{
            kind    = 'window'
            trigger = $t
            date    = $open.ToString($fmt)
            state   = (& $stateOf $key $open $close)
            sentAt  = $(if ($State.ContainsKey($key)) { $State[$key] } else { $null })
        })
    }

    # --- expiry day
    $key = "$($prefix)_expiryday"
    $items.Add([ordered]@{
        kind    = 'expiry'
        trigger = $null
        date    = $ExpireDate.ToString($fmt)
        state   = (& $stateOf $key $ExpireDate $ExpireDate)
        sentAt  = $(if ($State.ContainsKey($key)) { $State[$key] } else { $null })
    })

    # --- reminders after expiry: the ones already sent, then the next one
    $pattern = '^' + [regex]::Escape("$($prefix)_expired-") + '(\d{4}-\d{2}-\d{2})$'
    foreach ($k in @($State.Keys | Where-Object { $_ -match $pattern } | Sort-Object)) {
        $null = $k -match $pattern
        $items.Add([ordered]@{ kind = 'reminder'; trigger = $null; date = $Matches[1]; state = 'sent'; sentAt = $State[$k] })
    }
    $n = [Math]::Max(1, $ExpiredNotifyDays)
    if ($daysLeft -lt 0) {
        $over = -$daysLeft
        if ($over % $n -eq 0) {
            $todayKey = "$($prefix)_expired-$($Today.ToString($fmt))"
            if ($State.ContainsKey($todayKey)) { $nextDate = $Today.AddDays($n); $nextState = 'pending' }
            else                               { $nextDate = $Today;             $nextState = 'due' }
        }
        else { $nextDate = $Today.AddDays($n - ($over % $n)); $nextState = 'pending' }
    }
    else { $nextDate = $ExpireDate.AddDays($n); $nextState = 'pending' }
    $items.Add([ordered]@{ kind = 'reminder'; trigger = $null; date = $nextDate.ToString($fmt); state = $nextState; sentAt = $null })

    $sorted = @($items | Sort-Object { $_.date }, { if ($_.kind -eq 'window') { 0 } elseif ($_.kind -eq 'expiry') { 1 } else { 2 } })

    # --- last alert: the newest timestamp of any key for this object at this expiry date
    $last = $null
    foreach ($k in @($State.Keys | Where-Object { $_.StartsWith("$($prefix)_") })) {
        $at = $State[$k]
        $atTime = [datetime]::MinValue
        if ([datetime]::TryParse($at, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref] $atTime)) {
            if ($null -eq $last -or $atTime -gt $last.time) { $last = @{ time = $atTime; at = $at; key = $k } }
        }
    }
    $lastOut = $null
    if ($last) {
        $item = $sorted | Where-Object { $_.sentAt -eq $last.at } | Select-Object -First 1
        $lastOut = [ordered]@{ at = $last.at; kind = $(if ($item) { $item.kind } else { 'window' }); trigger = $(if ($item) { $item.trigger } else { $null }) }
    }

    # --- next alert: what is due now, else the earliest pending item
    $next = $sorted | Where-Object { $_.state -eq 'due' } | Select-Object -First 1
    if (-not $next) { $next = $sorted | Where-Object { $_.state -eq 'pending' } | Select-Object -First 1 }
    $nextOut = $null
    if ($next) { $nextOut = [ordered]@{ date = $next.date; kind = $next.kind; trigger = $next.trigger; state = $next.state } }

    return @{ Schedule = $sorted; Last = $lastOut; Next = $nextOut }
}
