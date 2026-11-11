<#
System requirements
    Run in the Exchange Management Shell on an on-prem Exchange server, or in a
    session created with New-PSSession -ConfigurationName Microsoft.Exchange.
    Needs the "Receive Connectors" RBAC role (Organization Management).

psToDo setup helper : Setup-ExchangeRelayReceiveConnector (ON-PREMISES EDITION)
    Author  : Fardin Barashi
    Purpose : Create a dedicated, IP-scoped anonymous relay Receive connector so
              the psToDo host can submit alert mail with AuthMode = Anonymous.
              This is the on-prem equivalent of the cloud edition's
              Setup-SenderMailbox-and-AccessPolicy: it limits WHO can send
              (only the psToDo host's IP) instead of granting a tenant-wide right.

What it does
    1. Creates a Frontend Transport Receive connector on the given Exchange
       server, bound to port 25, that ONLY accepts connections from the psToDo
       host IP(s) in -RemoteIPRanges. Exchange picks the most specific
       connector for a source IP, so the default connector is untouched.
    2. Grants NT AUTHORITY\ANONYMOUS LOGON the ms-Exch-SMTP-Accept-Any-Recipient
       right on it (accepted domains only unless -AllowExternalRelay).
    3. Optionally requires TLS (recommended - psToDo uses STARTTLS by default).

    Every step is skipped if the object already exists, so it is safe to re-run.

Parameters
    -Server           Exchange server that will host the connector (required)
    -RemoteIPRanges   IP address(es) of the host running psToDo, e.g. 10.10.5.20  (required)
    -Name             Connector name (default 'psToDo Relay')
    -Fqdn             FQDN announced in the banner / used for the TLS cert (default: server FQDN)
    -AllowExternalRelay  Also allow relaying to addresses OUTSIDE the accepted domains.
    -RequireTls       Reject unencrypted sessions.

Examples
    .\Setup-ExchangeRelayReceiveConnector.ps1 -Server mail-01 -RemoteIPRanges 10.10.5.20 -WhatIf
    .\Setup-ExchangeRelayReceiveConnector.ps1 -Server mail-01 -RemoteIPRanges 10.10.5.20 -RequireTls
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string]   $Server,
    [Parameter(Mandatory)] [string[]] $RemoteIPRanges,
    [string] $Name = 'psToDo Relay',
    [string] $Fqdn,
    [switch] $AllowExternalRelay,
    [switch] $RequireTls
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command Get-ReceiveConnector -ErrorAction SilentlyContinue)) {
    throw 'Exchange cmdlets not found. Run this in the Exchange Management Shell or import an Exchange remote session first.'
}

$identity = "$Server\$Name"
$existing = Get-ReceiveConnector -Identity $identity -ErrorAction SilentlyContinue

if ($existing) {
    Write-Host "Receive connector '$identity' already exists - leaving it as is." -ForegroundColor Yellow
}
elseif ($PSCmdlet.ShouldProcess($identity, 'Create anonymous relay Receive connector')) {
    $params = @{
        Name            = $Name
        Server          = $Server
        TransportRole   = 'FrontendTransport'
        Usage           = 'Custom'
        Bindings        = '0.0.0.0:25'
        RemoteIPRanges  = $RemoteIPRanges
        AuthMechanism   = if ($RequireTls) { 'Tls' } else { 'Tls, Integrated, BasicAuth, BasicAuthRequireTLS' }
        PermissionGroups = 'AnonymousUsers'
        RequireTLS      = [bool]$RequireTls
        Comment         = 'psToDo alert relay - accepts mail ONLY from the psToDo host IP(s).'
    }
    if ($Fqdn) { $params.Fqdn = $Fqdn }

    $existing = New-ReceiveConnector @params
    Write-Host "Created Receive connector '$identity' for $($RemoteIPRanges -join ', ')." -ForegroundColor Green
}

# Relay permission. Without ms-Exch-SMTP-Accept-Any-Recipient, anonymous senders may only
# deliver to recipients in accepted domains (which is enough for most psToDo setups).
if ($AllowExternalRelay -and $existing) {
    $ace = Get-ADPermission -Identity $identity -User 'NT AUTHORITY\ANONYMOUS LOGON' -ErrorAction SilentlyContinue |
           Where-Object { $_.ExtendedRights -match 'ms-Exch-SMTP-Accept-Any-Recipient' }
    if ($ace) {
        Write-Host 'Anonymous relay right already granted.' -ForegroundColor Yellow
    }
    elseif ($PSCmdlet.ShouldProcess($identity, 'Grant ms-Exch-SMTP-Accept-Any-Recipient to ANONYMOUS LOGON')) {
        Get-ReceiveConnector -Identity $identity |
            Add-ADPermission -User 'NT AUTHORITY\ANONYMOUS LOGON' -ExtendedRights 'ms-Exch-SMTP-Accept-Any-Recipient' | Out-Null
        Write-Host 'Granted external relay right to ANONYMOUS LOGON on the connector.' -ForegroundColor Green
    }
}

Write-Host ''
Write-Host 'Next steps on the psToDo host:' -ForegroundColor Cyan
Write-Host "  1. Set SmtpServer = <FQDN of $Server>, SmtpPort = 25, AuthMode = Anonymous in Settings\Config\ExchangeOnpremSettings.json"
Write-Host '  2. Run Install-help\Test-OnpremSmtp.ps1 -To you@yourdomain to verify delivery.'
