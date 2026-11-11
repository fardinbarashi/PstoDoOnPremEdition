function Connect-OnpremExchangeSmtp {
    # On-prem replacement for Connect-CalenderReminderGraph.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Settings,
        [int] $ConnectTimeoutSeconds = 10
    )

    # --- 1. Is the SMTP port reachable at all? Fail early with a clear message. ---
    $tcp = [System.Net.Sockets.TcpClient]::new()
    try {
        $async = $tcp.BeginConnect($Settings.SmtpServer, $Settings.SmtpPort, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($ConnectTimeoutSeconds))) {
            throw "Timed out after $ConnectTimeoutSeconds s connecting to $($Settings.SmtpServer):$($Settings.SmtpPort). Check the hostname, the firewall and that a Receive connector listens on that port."
        }
        $tcp.EndConnect($async)
    }
    catch {
        throw "Cannot reach the on-prem Exchange SMTP endpoint $($Settings.SmtpServer):$($Settings.SmtpPort) - $($_.Exception.Message)"
    }
    finally { $tcp.Close() }

    # --- 2. TLS version. Windows PowerShell 5.1 (.NET Framework) may still default to TLS 1.0/1.1,
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    } catch { Write-Verbose "Could not enable TLS 1.2: $($_.Exception.Message)" }

    # --- 2b. Certificate trust for STARTTLS ---
    if ($Settings.UseTls -and $Settings.TlsAllowUntrustedCertificate) {
        Write-Warning 'TlsAllowUntrustedCertificate is true: the Exchange TLS certificate will NOT be validated. Lab use only.'
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
    elseif (-not $Settings.UseTls) {
        Write-Warning 'UseTls is false: mail will be submitted in clear text on the LAN.'
    }

    # --- 3. Build the client ---
    $client = [System.Net.Mail.SmtpClient]::new($Settings.SmtpServer, $Settings.SmtpPort)
    $client.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
    $client.EnableSsl      = [bool]$Settings.UseTls
    $client.Timeout        = 60000

    switch ($Settings.AuthMode) {
        'Anonymous' {
            $client.UseDefaultCredentials = $false
            $client.Credentials = $null
        }
        'CurrentUser' {
            $client.UseDefaultCredentials = $true
        }
        'Credential' {
            try   { $cred = Import-Clixml -Path $Settings.CredentialXmlPath }
            catch { throw "Could not read the SMTP credential file $($Settings.CredentialXmlPath) as $($env:USERNAME). It is DPAPI-bound to the account and machine that created it: $($_.Exception.Message)" }
            if ($cred -isnot [System.Management.Automation.PSCredential]) { throw "The file $($Settings.CredentialXmlPath) does not contain a PSCredential." }
            $client.UseDefaultCredentials = $false
            $client.Credentials = $cred.GetNetworkCredential()
            Write-Host "  Authenticating to Exchange as $($cred.UserName)" -ForegroundColor DarkGray
        }
    }

    $script:OnpremSmtpClient = $client

    $tlsText = if ($Settings.UseTls) { 'STARTTLS' } else { 'no TLS' }
    Write-Host "Connected to on-prem Exchange SMTP $($Settings.SmtpServer):$($Settings.SmtpPort) ($tlsText, auth: $($Settings.AuthMode))." -ForegroundColor Green
    return $client
}

function Disconnect-OnpremExchangeSmtp {
    [CmdletBinding()]
    param()
    if ($script:OnpremSmtpClient) {
        try { $script:OnpremSmtpClient.Dispose() } catch { }
        $script:OnpremSmtpClient = $null
    }
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
}
