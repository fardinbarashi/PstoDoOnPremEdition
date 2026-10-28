function Send-AlertMail {
    <#
        On-prem edition: submits the alert to the on-prem Exchange server over
        SMTP using the client prepared by Connect-OnpremExchangeSmtp.

        The sender (mailSender in monitorobjects.json) must be an address the
        connector/account is allowed to send as:
          - Anonymous relay connector : any address in an accepted domain
          - CurrentUser / Credential  : the account's own primary SMTP address,
                                        or a mailbox it has Send-As on.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)] $Alert
    )

    if (-not $Alert.mailOn) {
        Write-Verbose "Object $($Alert.id): notifyMethodbyMail is false - skipping mail."
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($Alert.mailSender)) {
        Write-Warning "  Object $($Alert.id) has no mailSender - skipping mail."
        return $false
    }

    $recipients = @($Alert.mailRecipients) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (-not $recipients) {
        Write-Warning "  Object $($Alert.id) has no mailRecipients - skipping mail."
        return $false
    }

    if (-not $script:OnpremSmtpClient) {
        Write-Warning "  Object $($Alert.id): no SMTP connection. Run Connect-OnpremExchangeSmtp first."
        return $false
    }

    $p      = Get-AlertPresentation -Alert $Alert
    $target = "$($recipients -join ', ') for object $($Alert.id) - $($Alert.name)"

    if (-not $PSCmdlet.ShouldProcess($target, 'Send alert mail')) { return $false }

    $message = $null
    try {
        $message = [System.Net.Mail.MailMessage]::new()
        $message.From = [System.Net.Mail.MailAddress]::new($Alert.mailSender, 'psToDo')
        foreach ($r in $recipients) { $message.To.Add([System.Net.Mail.MailAddress]::new($r)) }

        $message.Subject         = "$($p.Prefix) $($Alert.mailSubject)"
        $message.SubjectEncoding = [System.Text.Encoding]::UTF8
        $message.Body            = Format-AlertMailBody -Alert $Alert
        $message.BodyEncoding    = [System.Text.Encoding]::UTF8
        $message.IsBodyHtml      = $true
        $message.Priority        = if ($Alert.urgency -in 'expired', 'critical') { [System.Net.Mail.MailPriority]::High } else { [System.Net.Mail.MailPriority]::Normal }
        $message.Headers.Add('X-psToDo-Object',   "$($Alert.id)")
        $message.Headers.Add('X-psToDo-StateKey', "$($Alert.stateKey)")

        $script:OnpremSmtpClient.Send($message)

        Write-Host "  Mail sent to $($recipients -join ', ')" -ForegroundColor Green
        return $true
    }
    catch {
        # SmtpException carries the Exchange reply (e.g. "5.7.1 Client does not have permissions to send as this sender")
        $detail = $_.Exception.Message
        if ($_.Exception.InnerException) { $detail += ' / ' + $_.Exception.InnerException.Message }
        Write-Warning "  Mail to $($recipients -join ', ') failed: $detail"
        return $false
    }
    finally {
        if ($message) { $message.Dispose() }
    }
}
