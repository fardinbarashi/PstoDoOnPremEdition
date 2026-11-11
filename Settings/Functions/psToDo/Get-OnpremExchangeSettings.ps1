function Get-OnpremExchangeSettings {
    <#
        Reads Settings\Config\ExchangeOnpremSettings.json, validates it and
        resolves relative paths (credential files) against the repo root.

        Returns a PSCustomObject with the effective settings. Nothing is
        contacted here - this is pure parsing/validation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $RepoRoot
    )

    if (-not (Test-Path $Path)) { throw "Cannot find the on-prem Exchange settings file: $Path" }

    try   { $raw = Get-Content -Raw -Encoding UTF8 $Path | ConvertFrom-Json }
    catch { throw "Failed to parse JSON in '$Path': $($_.Exception.Message)" }

    if ([string]::IsNullOrWhiteSpace($raw.SmtpServer)) { throw 'ExchangeOnpremSettings.json is missing SmtpServer' }

    $port = 25
    if ($null -ne $raw.SmtpPort -and "$($raw.SmtpPort)" -ne '') { $port = [int]$raw.SmtpPort }
    if ($port -lt 1 -or $port -gt 65535) { throw "ExchangeOnpremSettings.json has an invalid SmtpPort: $($raw.SmtpPort)" }

    $authMode = if ([string]::IsNullOrWhiteSpace($raw.AuthMode)) { 'Anonymous' } else { "$($raw.AuthMode)".Trim() }
    if ($authMode -notin 'Anonymous', 'CurrentUser', 'Credential') {
        throw "ExchangeOnpremSettings.json has an invalid AuthMode '$authMode'. Use Anonymous, CurrentUser or Credential."
    }

    $resolve = {
        param($p)
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }
        if ([System.IO.Path]::IsPathRooted($p)) { return $p }
        return (Join-Path $RepoRoot $p)
    }

    $credentialPath = & $resolve $raw.CredentialXmlPath
    if ($authMode -eq 'Credential') {
        if (-not $credentialPath) { throw 'AuthMode is Credential but CredentialXmlPath is empty.' }
        if (-not (Test-Path $credentialPath)) {
            throw "AuthMode is Credential but the credential file does not exist: $credentialPath. Create it with Install-help\New-SmtpCredentialFile.ps1 (run as the same account that runs psToDo)."
        }
    }

    $exAuth = if ([string]::IsNullOrWhiteSpace($raw.ExchangeAuthentication)) { 'Kerberos' } else { "$($raw.ExchangeAuthentication)".Trim() }
    if ($exAuth -notin 'Kerberos', 'Basic') {
        throw "ExchangeOnpremSettings.json has an invalid ExchangeAuthentication '$exAuth'. Use Kerberos or Basic."
    }

    [pscustomobject]@{
        SmtpServer                   = "$($raw.SmtpServer)".Trim()
        SmtpPort                     = $port
        UseTls                       = if ($null -eq $raw.UseTls) { $true } else { [bool]$raw.UseTls }
        TlsAllowUntrustedCertificate = [bool]$raw.TlsAllowUntrustedCertificate
        AuthMode                     = $authMode
        CredentialXmlPath            = $credentialPath
        ExchangePowerShellUri        = "$($raw.ExchangePowerShellUri)".Trim()
        ExchangeAuthentication       = $exAuth
        ExchangeCredentialXmlPath    = & $resolve $raw.ExchangeCredentialXmlPath
    }
}
