# ServiceNow-Search.ps1
# Purpose: Open ServiceNow in an isolated browser window with a search (global or incident list).
# Notes:
# - If -Query is omitted and -FromClipboard is set, it will try clipboard first, then prompt.
# - Uses a unique temp profile per launch to avoid SSO/tab reuse and to help window placement.

[CmdletBinding()]
param(
    [string]$Instance = "https://your-instance.service-now.com",
    [string]$Query,
    [switch]$FromClipboard,
    [string]$IncidentNumber,
    [ValidateSet("Edge","Chrome","Firefox")]
    [string]$Browser = "Edge"
)

function New-TempProfileDir {
    $name = "SN_{0}_{1}_{2}" -f (Get-Date -Format "yyyyMMdd_HHmmss"), $env:USERNAME, (Get-Random -Maximum 99999)
    $dir  = Join-Path $env:TEMP $name
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

# Build URL
if ($IncidentNumber) {
    # Incident list filtered by number
    $encoded = [uri]::EscapeDataString("number=$IncidentNumber")
    $url = "$Instance/incident_list.do?sysparm_query=$encoded"
} else {
    if (-not $Query) {
        if ($FromClipboard) {
            try {
                $clip = Get-Clipboard -ErrorAction Stop
                if ($clip) { $Query = $clip }
            } catch { }
        }
        if (-not $Query) {
            $Query = Read-Host "ServiceNow global search (textsearch)"
        }
    }
    $encoded = [uri]::EscapeDataString($Query)
    $url = "$Instance/textsearch.do?sysparm_search=$encoded"
}

# Isolated profile to force a new window (helps placement)
$profileDir = New-TempProfileDir

switch ($Browser) {
    "Edge" {
        $exe  = "msedge.exe"
        $args = @("--new-window","--inprivate","--user-data-dir=$profileDir",$url)
    }
    "Chrome" {
        $exe  = "chrome.exe"
        $args = @("--new-window","--incognito","--user-data-dir=$profileDir",$url)
    }
    "Firefox" {
        $exe  = "firefox.exe"
        # Firefox has no simple user-data-dir switch; new private window is enough for isolation
        $args = @("-new-window","-private-window",$url)
    }
}

Start-Process -FilePath $exe -ArgumentList $args -WindowStyle Normal
Write-Host "Launched $Browser → $url" -ForegroundColor Green
