function Send-AlertNotification {
    <#
        On-prem edition: mail is the only channel. Teams (a cloud webhook) is
        intentionally not implemented so that the script never talks to
        anything but the on-prem Exchange server.

        notifyMethodbyTeams is still read from monitorobjects.json for file
        compatibility with the cloud edition, but an object that ONLY has
        Teams enabled is reported as skipped.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)] $Alert)

    $result = [pscustomobject]@{
        Id        = $Alert.id
        Name      = $Alert.name
        Trigger   = $Alert.trigger
        Urgency   = $Alert.urgency
        StateKey  = $Alert.stateKey
        MailSent  = $false
        TeamsSent = $false   # always false in the on-prem edition
        Skipped   = $false
    }

    if ($Alert.teamsOn) {
        Write-Verbose "  Object $($Alert.id): notifyMethodbyTeams is true but Teams is not available in the on-prem edition. Ignored."
    }

    if (-not $Alert.mailOn) {
        Write-Warning "  Object $($Alert.id) - $($Alert.name): notifyMethodbyMail is false (on-prem edition has no Teams channel). Nothing sent."
        $result.Skipped = $true
        return $result
    }

    $result.MailSent = Send-AlertMail -Alert $Alert -WhatIf:$WhatIfPreference

    return $result
}
