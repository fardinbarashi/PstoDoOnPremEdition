function Get-Urgency {
# Triggers must be sorted ascending.
    param([int] $DaysLeft, [int[]] $Triggers)

    if ($DaysLeft -lt 0) { return 'expired' }
    if ($Triggers.Count -ge 1 -and $DaysLeft -le $Triggers[0]) { return 'critical' }
    if ($Triggers.Count -ge 2 -and $DaysLeft -le $Triggers[1]) { return 'warning'  }
    return 'ok'
}