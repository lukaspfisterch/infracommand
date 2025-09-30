# Outlook Profile Reset Tool
# Advanced PowerShell version with menu system
# Comprehensive Outlook profile troubleshooting and reset functionality

param(
    [switch]$Verbose = $false,
    [switch]$Force = $false
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

# Funktion für Office-Version erkennen
function Get-OfficeVersion {
    $versions = @("23.0", "22.0", "21.0", "20.0", "16.0", "15.0", "14.0")
    
    foreach ($version in $versions) {
        $regPath = "HKCU:\Software\Microsoft\Office\$version\Outlook"
        if (Test-Path $regPath) {
            return $version
        }
    }
    return $null
}

# Funktion für Backup erstellen
function New-OutlookBackup {
    param([string]$BackupDir)
    
    Write-Host "Erstelle Backup..." -ForegroundColor Yellow
    
    try {
        # Registry Backup
        $officeVersion = Get-OfficeVersion
        if ($officeVersion) {
            $officeKeyPath = "HKCU:\Software\Microsoft\Office\$officeVersion"
            
            # Outlook Registry exportieren
            $outlookRegPath = "$officeKeyPath\Outlook"
            if (Test-Path $outlookRegPath) {
                reg export $outlookRegPath "$BackupDir\HKCU_Outlook.reg" /y | Out-Null
                Write-Success-Info "Outlook Registry gesichert"
            }
            
            # Outlook Profiles Registry exportieren
            $profilesRegPath = "$officeKeyPath\Outlook\Profiles"
            if (Test-Path $profilesRegPath) {
                reg export $profilesRegPath "$BackupDir\OutlookProfiles.reg" /y | Out-Null
                Write-Success-Info "Outlook Profiles Registry gesichert"
            }
        }
        
        # Signatures sichern
        $signaturesPath = "$env:APPDATA\Microsoft\Signatures"
        if (Test-Path $signaturesPath) {
            Copy-Item $signaturesPath "$BackupDir\Signatures" -Recurse -Force
            Write-Success-Info "Outlook Signaturen gesichert"
        } else {
            Write-Warning-Info "Keine Signaturen gefunden"
        }
        
        # Outlook-Dateien sichern
        $outlookDataPath = "$env:APPDATA\Microsoft\Outlook"
        if (Test-Path $outlookDataPath) {
            Copy-Item $outlookDataPath "$BackupDir\OutlookData" -Recurse -Force
            Write-Success-Info "Outlook-Daten gesichert"
        }
        
        return $true
    } catch {
        Write-Error-Info "Backup fehlgeschlagen: $($_.Exception.Message)"
        return $false
    }
}

# Funktion für Outlook beenden
function Stop-OutlookProcess {
    Write-Host "Beende Outlook-Prozesse..." -ForegroundColor Yellow
    
    try {
        $outlookProcesses = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
        if ($outlookProcesses) {
            $outlookProcesses | Stop-Process -Force
            Write-Success-Info "Outlook-Prozesse beendet"
        } else {
            Write-Success-Info "Keine Outlook-Prozesse gefunden"
        }
    } catch {
        Write-Error-Info "Fehler beim Beenden von Outlook: $($_.Exception.Message)"
    }
}

# Funktion für vollständige Bereinigung
function Start-FullReset {
    Write-Header "VOLLSTÄNDIGE OUTLOOK-BEREINIGUNG"
    
    # Backup-Verzeichnis erstellen
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $backupDir = "C:\Temp\Informatik\Outlook_$env:USERNAME`_$timestamp"
    
    try {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        Write-Success-Info "Backup-Verzeichnis erstellt: $backupDir"
    } catch {
        Write-Error-Info "Backup-Verzeichnis konnte nicht erstellt werden"
        return
    }
    
    # Backup erstellen
    if (-not (New-OutlookBackup -BackupDir $backupDir)) {
        Write-Error-Info "Backup fehlgeschlagen - Abbruch"
        return
    }
    
    # Outlook beenden
    Stop-OutlookProcess
    
    # Lokale Outlook-Daten löschen
    Write-Host "Lösche lokale Outlook-Daten..." -ForegroundColor Yellow
    
    $pathsToDelete = @(
        "$env:LOCALAPPDATA\Microsoft\Outlook\RoamCache",
        "$env:LOCALAPPDATA\Microsoft\Outlook\Offline Address Books",
        "$env:LOCALAPPDATA\Microsoft\Outlook",
        "$env:APPDATA\Microsoft\Outlook"
    )
    
    foreach ($path in $pathsToDelete) {
        if (Test-Path $path) {
            try {
                Remove-Item $path -Recurse -Force
                Write-Success-Info "Gelöscht: $path"
            } catch {
                Write-Error-Info "Fehler beim Löschen von $path"
            }
        }
    }
    
    # Registry bereinigen
    Write-Host "Bereinige Registry..." -ForegroundColor Yellow
    
    $officeVersion = Get-OfficeVersion
    if ($officeVersion) {
        $officeKeyPath = "HKCU:\Software\Microsoft\Office\$officeVersion"
        
        $regPathsToDelete = @(
            "$officeKeyPath\Outlook\Profiles",
            "$officeKeyPath\Common\MailSettings",
            "$officeKeyPath\Outlook\AutoDiscover"
        )
        
        foreach ($regPath in $regPathsToDelete) {
            if (Test-Path $regPath) {
                try {
                    Remove-Item $regPath -Recurse -Force
                    Write-Success-Info "Registry gelöscht: $regPath"
                } catch {
                    Write-Error-Info "Fehler beim Löschen der Registry: $regPath"
                }
            }
        }
    }
    
    # Weitere Registry-Bereinigung
    $additionalRegPaths = @(
        "HKCU:\Software\Microsoft\Exchange",
        "HKCU:\Software\Microsoft\Office\Common\UserInfo"
    )
    
    foreach ($regPath in $additionalRegPaths) {
        if (Test-Path $regPath) {
            try {
                Remove-Item $regPath -Recurse -Force
                Write-Success-Info "Registry gelöscht: $regPath"
            } catch {
                Write-Error-Info "Fehler beim Löschen der Registry: $regPath"
            }
        }
    }
    
    # Signaturen wiederherstellen
    Write-Host "Stelle Signaturen wieder her..." -ForegroundColor Yellow
    $signaturesBackup = "$backupDir\Signatures"
    if (Test-Path $signaturesBackup) {
        try {
            Copy-Item $signaturesBackup "$env:APPDATA\Microsoft\Signatures" -Recurse -Force
            Write-Success-Info "Signaturen wiederhergestellt"
        } catch {
            Write-Error-Info "Fehler beim Wiederherstellen der Signaturen"
        }
    }
    
    Write-Host "`nVollständige Bereinigung abgeschlossen!" -ForegroundColor Green
    Write-Host "Backup-Verzeichnis: $backupDir" -ForegroundColor Cyan
}

# Funktion für GAL-Probleme
function Start-GALReset {
    Write-Header "GAL-PROBLEME BEREINIGUNG"
    
    Stop-OutlookProcess
    
    Write-Host "Lösche GAL-bezogene Komponenten..." -ForegroundColor Yellow
    
    $galPaths = @(
        "$env:LOCALAPPDATA\Microsoft\Outlook\RoamCache",
        "$env:LOCALAPPDATA\Microsoft\Outlook\Offline Address Books"
    )
    
    foreach ($path in $galPaths) {
        if (Test-Path $path) {
            try {
                Remove-Item $path -Recurse -Force
                Write-Success-Info "Gelöscht: $path"
            } catch {
                Write-Error-Info "Fehler beim Löschen von $path"
            }
        }
    }
    
    Write-Host "`nGAL-Bereinigung abgeschlossen!" -ForegroundColor Green
}

# Funktion für minimalen Reset
function Start-MinimalReset {
    Write-Header "MINIMALER PROFIL-RESET"
    
    Stop-OutlookProcess
    
    Write-Host "Lösche Outlook-Daten und Profile..." -ForegroundColor Yellow
    
    $minimalPaths = @(
        "$env:APPDATA\Microsoft\Outlook",
        "$env:LOCALAPPDATA\Microsoft\Outlook"
    )
    
    foreach ($path in $minimalPaths) {
        if (Test-Path $path) {
            try {
                Remove-Item $path -Recurse -Force
                Write-Success-Info "Gelöscht: $path"
            } catch {
                Write-Error-Info "Fehler beim Löschen von $path"
            }
        }
    }
    
    # Registry Profile löschen
    $officeVersion = Get-OfficeVersion
    if ($officeVersion) {
        $profilesRegPath = "HKCU:\Software\Microsoft\Office\$officeVersion\Outlook\Profiles"
        if (Test-Path $profilesRegPath) {
            try {
                Remove-Item $profilesRegPath -Recurse -Force
                Write-Success-Info "Outlook Profile Registry gelöscht"
            } catch {
                Write-Error-Info "Fehler beim Löschen der Profile Registry"
            }
        }
    }
    
    Write-Host "`nMinimaler Reset abgeschlossen!" -ForegroundColor Green
}

# Hauptmenü
function Show-Menu {
    Write-Header "OUTLOOK-PROFILBEREINIGUNG - AUSWAHLMENÜ"
    Write-Host ""
    Write-Host "[1] Vollständige Bereinigung (alles löschen)" -ForegroundColor White
    Write-Host "[2] Nur GAL-Probleme (RoamCache, OAB)" -ForegroundColor White
    Write-Host "[3] Minimal-Reset (Profil und Daten, keine Registry)" -ForegroundColor White
    Write-Host "[4] Office-Version anzeigen" -ForegroundColor White
    Write-Host "[5] Outlook-Prozesse beenden" -ForegroundColor White
    Write-Host "[0] Abbrechen" -ForegroundColor White
    Write-Host ""
}

# Hauptprogramm
Write-Header "OUTLOOK PROFIL-RESET TOOL"

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

# Office-Version prüfen
$officeVersion = Get-OfficeVersion
if ($officeVersion) {
    Write-Success-Info "Office-Version erkannt: $officeVersion"
} else {
    Write-Warning-Info "Keine Office-Installation erkannt"
}

# Menü-Schleife
do {
    Show-Menu
    $choice = Read-Host "Bitte Option wählen"
    
    switch ($choice) {
        "1" { Start-FullReset; break }
        "2" { Start-GALReset; break }
        "3" { Start-MinimalReset; break }
        "4" { 
            $version = Get-OfficeVersion
            if ($version) {
                Write-Host "Office-Version: $version" -ForegroundColor Green
            } else {
                Write-Host "Keine Office-Installation gefunden" -ForegroundColor Red
            }
            Read-Host "Drücken Sie Enter zum Fortfahren..."
        }
        "5" { 
            Stop-OutlookProcess
            Read-Host "Drücken Sie Enter zum Fortfahren..."
        }
        "0" { 
            Write-Host "Abgebrochen." -ForegroundColor Yellow
            exit 0
        }
        default { 
            Write-Host "Ungültige Option. Bitte erneut wählen." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
    if ($choice -in @("1", "2", "3")) {
        Write-Host "`n============================================" -ForegroundColor Cyan
        Write-Host "Outlook-Bereinigung abgeschlossen!" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Cyan
        
        if ($Verbose) {
            Read-Host "`nDrücken Sie Enter zum Beenden..."
        } else {
            Start-Sleep -Seconds 2
        }
        break
    }
} while ($true)

# Sauberer Exit - wechsle zu Systemverzeichnis
try {
    Set-Location "C:\Windows\System32" | Out-Null
    Write-Host "`nWechsle zu Systemverzeichnis: C:\Windows\System32" -ForegroundColor Gray
} catch {
    try {
        Set-Location "C:\Windows" | Out-Null
        Write-Host "`nWechsle zu Systemverzeichnis: C:\Windows" -ForegroundColor Gray
    } catch {
        Set-Location "C:\" | Out-Null
        Write-Host "`nWechsle zu Systemverzeichnis: C:\" -ForegroundColor Gray
    }
}

# Finale Bereinigung
Write-Host "Bereinige temporäre Variablen..." -ForegroundColor Gray
Remove-Variable -Name "officeVersion" -ErrorAction SilentlyContinue
Remove-Variable -Name "backupDir" -ErrorAction SilentlyContinue
Remove-Variable -Name "isAdmin" -ErrorAction SilentlyContinue

Write-Host "`nSkript erfolgreich beendet." -ForegroundColor Green
exit 0
