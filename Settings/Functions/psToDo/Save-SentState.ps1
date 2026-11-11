function Save-SentState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $State,
        [Parameter(Mandatory)] [string]    $Path,
        [Parameter(Mandatory)] [string[]]  $LiveKeyPrefixes
    )

    foreach ($key in @($State.Keys)) {
        $prefix = ($key -split '_')[0..1] -join '_'
        if ($prefix -notin $LiveKeyPrefixes) {
            $State.Remove($key)
            Write-Verbose "Retired stale state key: $key"
        }
    }

    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    try {
        if ($State.Count) {
            $State | ConvertTo-Json -Depth 3 | Set-Content -Path $Path -Encoding UTF8
        }
        else {
            '{}' | Set-Content -Path $Path -Encoding UTF8
        }
        Write-Host "State saved: $($State.Count) window(s) recorded." -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "Could not write the state file. The next run may resend: $($_.Exception.Message)"
    }
}