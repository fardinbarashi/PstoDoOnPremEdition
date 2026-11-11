function Test-RequiredModule {
    # On-prem replacement for Initialize-RequiredModules.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $Modules,
        [hashtable] $InstallHint = @{}
    )

    $missing = @()
    foreach ($module in $Modules) {
        Write-Host "Checking module: $module" -ForegroundColor Cyan
        if (Get-Module -ListAvailable -Name $module) {
            Import-Module $module -ErrorAction Stop
            Write-Host '- Module found and imported.' -ForegroundColor Green
        }
        else {
            $hint = if ($InstallHint.ContainsKey($module)) { $InstallHint[$module] } else { 'Install it on this host (no automatic download in the on-prem edition).' }
            Write-Host "- Module NOT found. $hint" -ForegroundColor Red
            $missing += $module
        }
    }

    if ($missing) { throw "Missing module(s): $($missing -join ', ')" }
    Write-Host 'All modules are ready.' -ForegroundColor Green
}
