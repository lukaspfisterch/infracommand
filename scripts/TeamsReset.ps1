# Microsoft Teams Repair Tool
# PowerShell version with interactive menu system
# Comprehensive Teams troubleshooting and reset functionality

param(
    [switch]$FullReset = $false
)

function Show-Menu {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "     Microsoft Teams Repair Tool" -ForegroundColor Cyan
    Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select an option:" -ForegroundColor White
    Write-Host ""
    Write-Host "[1] Stop Teams processes" -ForegroundColor Yellow
    Write-Host "[2] Clear cache and logs (safe)" -ForegroundColor Green
    Write-Host "[3] Backup settings" -ForegroundColor Blue
    Write-Host "[4] Restart Teams" -ForegroundColor Yellow
    Write-Host "[5] Full reset (CRITICAL!)" -ForegroundColor Red
    Write-Host "[6] Run everything (Cache + Restart)" -ForegroundColor Magenta
    Write-Host "[7] Kompletter Reset mit Bestätigung" -ForegroundColor Red
    Write-Host "[0] Beenden" -ForegroundColor Gray
    Write-Host ""
}

function Confirm-CriticalAction {
    param([string]$Action, [string]$Warning)
    
    Write-Host "`n⚠️  WARNUNG: $Warning" -ForegroundColor Red
    Write-Host "Aktion: $Action" -ForegroundColor Yellow
    Write-Host ""
    
    $confirm1 = Read-Host "Sind Sie SICHER, dass Sie fortfahren möchten? (JA/nein)"
    if ($confirm1 -ne "JA") {
        Write-Host "❌ Aktion abgebrochen." -ForegroundColor Red
        return $false
    }
    
    $confirm2 = Read-Host "Letzte Bestätigung - geben Sie 'BESTÄTIGEN' ein"
    if ($confirm2 -ne "BESTÄTIGEN") {
        Write-Host "❌ Aktion abgebrochen." -ForegroundColor Red
        return $false
    }
    
    Write-Host "✅ Bestätigung erhalten. Führe Aktion aus..." -ForegroundColor Green
    return $true
}

function Stop-TeamsProcesses {
    Write-Host "`n[*] Beende alle Teams-Prozesse..." -ForegroundColor Yellow
    try {
        Get-Process -Name "Teams" -ErrorAction SilentlyContinue | Stop-Process -Force
        Get-Process -Name "Update" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "     -> Teams-Prozesse beendet" -ForegroundColor Green
    } catch {
        Write-Host "     -> Keine Teams-Prozesse gefunden" -ForegroundColor Gray
    }
}

function Clear-TeamsCache {
    Write-Host "`n[*] Entferne Cache, Temp und Logs..." -ForegroundColor Yellow
    $TeamsConfig = "$env:APPDATA\Microsoft\Teams"
    $CacheDirs = @(
        "Cache",
        "blob_storage", 
        "databases",
        "GPUCache",
        "IndexedDB",
        "Local Storage",
        "tmp",
        "Logs"
    )

    foreach ($dir in $CacheDirs) {
        $fullPath = "$TeamsConfig\$dir"
        if (Test-Path $fullPath) {
            try {
                Remove-Item $fullPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "     -> $dir entfernt" -ForegroundColor Green
            } catch {
                Write-Host "     -> Fehler beim Entfernen von $dir" -ForegroundColor Red
            }
        }
    }
}

function Backup-TeamsSettings {
    Write-Host "`n[*] Sichere Teams-Einstellungen..." -ForegroundColor Yellow
    $BackupDir = "$env:TEMP\TeamsBackup_$env:USERNAME"
    if (-not (Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
    }
    
    $TeamsConfig = "$env:APPDATA\Microsoft\Teams"
    $SettingsFile = "$TeamsConfig\settings.json"

    if (Test-Path $SettingsFile) {
        try {
            Copy-Item $SettingsFile "$BackupDir\settings_backup.json" -Force
            Write-Host "     -> settings.json gesichert nach: $BackupDir" -ForegroundColor Green
        } catch {
            Write-Host "     -> Fehler beim Sichern der settings.json" -ForegroundColor Red
        }
    } else {
        Write-Host "     -> settings.json nicht gefunden" -ForegroundColor Gray
    }
}

function Start-Teams {
    Write-Host "`n[*] Starte Teams neu..." -ForegroundColor Yellow
    $TeamsExe = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
    if (Test-Path $TeamsExe) {
        try {
            Start-Process -FilePath $TeamsExe -ArgumentList "--processStart", "Teams.exe"
            Write-Host "     -> Teams gestartet" -ForegroundColor Green
        } catch {
            Write-Host "     -> Fehler beim Starten von Teams" -ForegroundColor Red
        }
    } else {
        Write-Host "     -> Teams-Update.exe nicht gefunden" -ForegroundColor Red
    }
}

function Reset-TeamsCompletely {
    if (-not (Confirm-CriticalAction -Action "Vollständiger Teams-Reset" -Warning "Dies wird ALLE Teams-Daten löschen inklusive Login! Sie müssen sich neu anmelden.")) {
        return
    }
    
    Write-Host "`n[*] Führe vollständigen Reset durch..." -ForegroundColor Yellow
    $TeamsConfig = "$env:APPDATA\Microsoft\Teams"
    if (Test-Path $TeamsConfig) {
        try {
            Remove-Item $TeamsConfig -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "     -> Komplettes Teams-Verzeichnis entfernt" -ForegroundColor Green
        } catch {
            Write-Host "     -> Fehler beim vollständigen Reset" -ForegroundColor Red
        }
    }
}

# Hauptmenü-Schleife
do {
    Show-Menu
    $choice = Read-Host "Ihre Wahl (0-7)"
    
    switch ($choice) {
        "1" {
            Stop-TeamsProcesses
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "2" {
            Clear-TeamsCache
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "3" {
            Backup-TeamsSettings
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "4" {
            Start-Teams
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "5" {
            Reset-TeamsCompletely
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "6" {
            Write-Host "`n[*] Führe sichere Reparatur aus..." -ForegroundColor Magenta
            Stop-TeamsProcesses
            Clear-TeamsCache
            Start-Teams
            Write-Host "`n✅ Reparatur abgeschlossen!" -ForegroundColor Green
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "7" {
            Write-Host "`n[*] Führe kompletten Reset mit Bestätigung aus..." -ForegroundColor Red
            Stop-TeamsProcesses
            Backup-TeamsSettings
            Reset-TeamsCompletely
            Start-Teams
            Write-Host "`n✅ Kompletter Reset abgeschlossen!" -ForegroundColor Green
            Read-Host "`nDrücken Sie Enter zum Fortfahren..."
        }
        "0" {
            Write-Host "`n👋 Auf Wiedersehen!" -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "`n❌ Ungültige Auswahl. Bitte wählen Sie 0-7." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne "0")
