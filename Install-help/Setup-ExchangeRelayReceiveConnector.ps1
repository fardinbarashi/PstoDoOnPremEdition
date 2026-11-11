<#
System requirements
Exchange Management Shell on the Exchange server

About Script :
Author : Fardin Barashi
Title : psToDo setup - Exchange relay Receive connector
Description : Creates an anonymous relay Receive connector that only accepts the psToDo server IP (AuthMode Anonymous).
For more information, please visit the GitHub repository: https://github.com/fardinbarashi/psToDo

Examples
    .\Setup-ExchangeRelayReceiveConnector.ps1 -Server mail-01 -RemoteIPRanges 10.10.5.20 -WhatIf
    .\Setup-ExchangeRelayReceiveConnector.ps1 -Server mail-01 -RemoteIPRanges 10.10.5.20 -RequireTls

Version : 1.0-onprem
Release day : 2026-09-08
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
