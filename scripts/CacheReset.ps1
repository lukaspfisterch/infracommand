# Cache Reset Tool
# Comprehensive cache cleanup for system and application caches
# Clears various system and application caches to resolve performance issues

param(
    [switch]$Verbose = $false,
    [switch]$Force = $false,
    [switch]$All = $false
)

# Funktion für saubere Ausgabe
function Write-Header {
    param([string]$Title)
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "     $Title" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
}

# Funktion für Fehlerbehandlung
function Write-Error-Info {
    param([string]$Message)
    Write-Host "  [!] $Message" -ForegroundColor Red
}

# Funktion für Erfolgsmeldung
function Write-Success-Info {
    param([string]$Message)
    Write-Host "  -> $Message" -ForegroundColor Green
}

# Funktion für Warnung
function Write-Warning-Info {
    param([string]$Message)
    Write-Host "  [!] $Message" -ForegroundColor Yellow
}

# Funktion für Cache-Größe berechnen
function Get-FolderSize {
    param([string]$Path)
    
    if (Test-Path $Path) {
        try {
            $size = (Get-ChildItem $Path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            return [math]::Round($size / 1MB, 2)
        } catch {
            return 0
        }
    }
    return 0
}

# Funktion für Cache löschen
function Clear-CacheFolder {
    param(
        [string]$Path,
        [string]$Description,
        [switch]$Recursive = $true
    )
    
    if (Test-Path $Path) {
        $sizeBefore = Get-FolderSize -Path $Path
        Write-Host "  $Description (${sizeBefore} MB)..." -ForegroundColor Yellow
        
        try {
            if ($Recursive) {
                Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            } else {
                Get-ChildItem $Path -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            }
            
            $sizeAfter = Get-FolderSize -Path $Path
            $freed = $sizeBefore - $sizeAfter
            Write-Success-Info "$Description bereinigt (${freed} MB freigegeben)"
            return $freed
        } catch {
            Write-Error-Info "Fehler beim Bereinigen von $Description"
            return 0
        }
    } else {
        Write-Warning-Info "$Description nicht gefunden"
        return 0
    }
}

# Funktion für Browser-Cache bereinigen
function Clear-BrowserCache {
    Write-Header "BROWSER-CACHE BEREINIGUNG"
    
    $totalFreed = 0
    
    # Chrome Cache
    $chromePaths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\ShaderCache"
    )
    
    foreach ($path in $chromePaths) {
        $totalFreed += Clear-CacheFolder -Path $path -Description "Chrome Cache"
    }
    
    # Edge Cache
    $edgePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache"
    )
    
    foreach ($path in $edgePaths) {
        $totalFreed += Clear-CacheFolder -Path $path -Description "Edge Cache"
    }
    
    # Firefox Cache
    $firefoxPaths = @(
        "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*\cache2",
        "$env:APPDATA\Mozilla\Firefox\Profiles\*\cache2"
    )
    
    foreach ($pattern in $firefoxPaths) {
        $paths = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue
        foreach ($path in $paths) {
            $totalFreed += Clear-CacheFolder -Path $path.FullName -Description "Firefox Cache"
        }
    }
    
    Write-Host "`nBrowser-Cache: $totalFreed MB freigegeben" -ForegroundColor Green
    return $totalFreed
}

# Funktion für Windows-Cache bereinigen
function Clear-WindowsCache {
    Write-Header "WINDOWS-CACHE BEREINIGUNG"
    
    $totalFreed = 0
    
    # Windows Temp
    $totalFreed += Clear-CacheFolder -Path "$env:TEMP" -Description "Windows Temp"
    $totalFreed += Clear-CacheFolder -Path "C:\Windows\Temp" -Description "Windows System Temp"
    
    # Windows Update Cache
    $totalFreed += Clear-CacheFolder -Path "C:\Windows\SoftwareDistribution\Download" -Description "Windows Update Cache"
    
    # Windows Store Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Windows\INetCache" -Description "Windows Store Cache"
    
    # Windows Defender Cache
    $totalFreed += Clear-CacheFolder -Path "C:\ProgramData\Microsoft\Windows Defender\Scans\History" -Description "Windows Defender Cache"
    
    # Event Logs (nur wenn Force)
    if ($Force) {
        $totalFreed += Clear-CacheFolder -Path "C:\Windows\Logs" -Description "Windows Logs" -Recursive:$false
    }
    
    Write-Host "`nWindows-Cache: $totalFreed MB freigegeben" -ForegroundColor Green
    return $totalFreed
}

# Funktion für Office-Cache bereinigen
function Clear-OfficeCache {
    Write-Header "OFFICE-CACHE BEREINIGUNG"
    
    $totalFreed = 0
    
    # Office Cache
    $officePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache",
        "$env:LOCALAPPDATA\Microsoft\Office\15.0\OfficeFileCache",
        "$env:APPDATA\Microsoft\Office\Recent",
        "$env:LOCALAPPDATA\Microsoft\Office\UnsavedFiles"
    )
    
    foreach ($path in $officePaths) {
        $totalFreed += Clear-CacheFolder -Path $path -Description "Office Cache"
    }
    
    # Outlook Cache
    $outlookPaths = @(
        "$env:LOCALAPPDATA\Microsoft\Outlook\RoamCache",
        "$env:LOCALAPPDATA\Microsoft\Outlook\Offline Address Books"
    )
    
    foreach ($path in $outlookPaths) {
        $totalFreed += Clear-CacheFolder -Path $path -Description "Outlook Cache"
    }
    
    Write-Host "`nOffice-Cache: $totalFreed MB freigegeben" -ForegroundColor Green
    return $totalFreed
}

# Funktion für Anwendungs-Cache bereinigen
function Clear-ApplicationCache {
    Write-Header "ANWENDUNGS-CACHE BEREINIGUNG"
    
    $totalFreed = 0
    
    # Teams Cache
    $totalFreed += Clear-CacheFolder -Path "$env:APPDATA\Microsoft\Teams\Cache" -Description "Teams Cache"
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Teams\Cache" -Description "Teams Local Cache"
    
    # OneDrive Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\OneDrive\logs" -Description "OneDrive Cache"
    
    # Adobe Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Adobe\Common\Media Cache Files" -Description "Adobe Cache"
    
    # Visual Studio Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\VisualStudio\16.0\ComponentModelCache" -Description "Visual Studio Cache"
    
    # .NET Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\VisualStudio\Packages" -Description ".NET Cache"
    
    Write-Host "`nAnwendungs-Cache: $totalFreed MB freigegeben" -ForegroundColor Green
    return $totalFreed
}

# Funktion für System-Cache bereinigen
function Clear-SystemCache {
    Write-Header "SYSTEM-CACHE BEREINIGUNG"
    
    $totalFreed = 0
    
    # DNS Cache leeren
    Write-Host "  DNS Cache leeren..." -ForegroundColor Yellow
    try {
        ipconfig /flushdns | Out-Null
        Write-Success-Info "DNS Cache geleert"
    } catch {
        Write-Error-Info "DNS Cache leeren fehlgeschlagen"
    }
    
    # Windows Store Cache
    Write-Host "  Windows Store Cache leeren..." -ForegroundColor Yellow
    try {
        Get-AppxPackage -AllUsers | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"} | Out-Null
        Write-Success-Info "Windows Store Cache geleert"
    } catch {
        Write-Error-Info "Windows Store Cache leeren fehlgeschlagen"
    }
    
    # Thumbnail Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Description "Thumbnail Cache"
    
    # Font Cache
    $totalFreed += Clear-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Windows\Fonts" -Description "Font Cache"
    
    Write-Host "`nSystem-Cache: $totalFreed MB freigegeben" -ForegroundColor Green
    return $totalFreed
}

# Hauptmenü
function Show-Menu {
    Write-Header "CACHE RESET TOOL - AUSWAHLMENÜ"
    Write-Host ""
    Write-Host "[1] Browser-Cache (Chrome, Edge, Firefox)" -ForegroundColor White
    Write-Host "[2] Windows-Cache (Temp, Updates, Logs)" -ForegroundColor White
    Write-Host "[3] Office-Cache (Word, Excel, Outlook)" -ForegroundColor White
    Write-Host "[4] Anwendungs-Cache (Teams, OneDrive, Adobe)" -ForegroundColor White
    Write-Host "[5] System-Cache (DNS, Store, Thumbnails)" -ForegroundColor White
    Write-Host "[6] Alle Caches bereinigen" -ForegroundColor White
    Write-Host "[7] Cache-Größen anzeigen" -ForegroundColor White
    Write-Host "[0] Abbrechen" -ForegroundColor White
    Write-Host ""
}

# Funktion für Cache-Größen anzeigen
function Show-CacheSizes {
    Write-Header "CACHE-GRÖSSEN ÜBERSICHT"
    
    $caches = @(
        @{Path="$env:TEMP"; Name="Windows Temp"},
        @{Path="C:\Windows\Temp"; Name="System Temp"},
        @{Path="$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"; Name="Chrome Cache"},
        @{Path="$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"; Name="Edge Cache"},
        @{Path="$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"; Name="Office Cache"},
        @{Path="$env:APPDATA\Microsoft\Teams\Cache"; Name="Teams Cache"},
        @{Path="$env:LOCALAPPDATA\Microsoft\OneDrive\logs"; Name="OneDrive Cache"}
    )
    
    $totalSize = 0
    foreach ($cache in $caches) {
        $size = Get-FolderSize -Path $cache.Path
        if ($size -gt 0) {
            Write-Host "  $($cache.Name): $size MB" -ForegroundColor White
            $totalSize += $size
        }
    }
    
    Write-Host "`nGesamtgröße: $totalSize MB" -ForegroundColor Green
}

# Hauptprogramm
Write-Header "CACHE RESET TOOL"

# Prüfe ob als Administrator ausgeführt
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Warning-Info "Warnung: Nicht als Administrator ausgeführt. Einige Funktionen könnten fehlschlagen."
    if (-not $Force) {
        $continue = Read-Host "Trotzdem fortfahren? (j/n)"
        if ($continue -ne "j" -and $continue -ne "J") {
            exit 0
        }
    }
}

# Menü-Schleife
do {
    Show-Menu
    $choice = Read-Host "Bitte Option wählen"
    
    $totalFreed = 0
    
    switch ($choice) {
        "1" { $totalFreed = Clear-BrowserCache; break }
        "2" { $totalFreed = Clear-WindowsCache; break }
        "3" { $totalFreed = Clear-OfficeCache; break }
        "4" { $totalFreed = Clear-ApplicationCache; break }
        "5" { $totalFreed = Clear-SystemCache; break }
        "6" { 
            Write-Header "ALLE CACHES BEREINIGEN"
            $totalFreed += Clear-BrowserCache
            $totalFreed += Clear-WindowsCache
            $totalFreed += Clear-OfficeCache
            $totalFreed += Clear-ApplicationCache
            $totalFreed += Clear-SystemCache
            break
        }
        "7" { 
            Show-CacheSizes
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
            continue
        }
        "0" { 
            Write-Host "Abgebrochen." -ForegroundColor Yellow
            exit 0
        }
        default { 
            Write-Host "Ungültige Option. Bitte erneut wählen." -ForegroundColor Red
            Start-Sleep -Seconds 2
            continue
        }
    }
    
    if ($choice -in @("1", "2", "3", "4", "5", "6")) {
        Write-Host "`n============================================" -ForegroundColor Cyan
        Write-Host "Cache-Bereinigung abgeschlossen!" -ForegroundColor Green
        Write-Host "Gesamt freigegeben: $totalFreed MB" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Cyan
        
        if ($Verbose) {
            Read-Host "`nDrücken Sie Enter zum Beenden..."
        } else {
            Start-Sleep -Seconds 2
        }
        break
    }
} while ($true)

# Sauberer Exit - wechsle zu C:\Windows\System32
Write-Host "`nBereinige temporäre Variablen..." -ForegroundColor Gray
Remove-Variable -Name "totalFreed" -ErrorAction SilentlyContinue
Remove-Variable -Name "isAdmin" -ErrorAction SilentlyContinue
Remove-Variable -Name "choice" -ErrorAction SilentlyContinue

# Garantiert zu C:\Windows\System32 wechseln
Write-Host "Wechsle zu Systemverzeichnis..." -ForegroundColor Gray

# Mehrere Versuche um sicherzustellen, dass wir in System32 landen
$attempts = 0
$maxAttempts = 5

do {
    $attempts++
    try {
        Set-Location "C:\Windows\System32" | Out-Null
        $currentPath = Get-Location
        if ($currentPath.Path -eq "C:\Windows\System32") {
            Write-Host "Erfolgreich zu C:\Windows\System32 gewechselt" -ForegroundColor Green
            break
        }
    } catch {
        Write-Host "Versuch $attempts fehlgeschlagen, versuche erneut..." -ForegroundColor Yellow
        Start-Sleep -Milliseconds 100
    }
} while ($attempts -lt $maxAttempts)

# Finale Bestätigung
$finalPath = Get-Location
Write-Host "Finaler Pfad: $($finalPath.Path)" -ForegroundColor Cyan

Write-Host "`nSkript erfolgreich beendet." -ForegroundColor Green
exit 0
