# Microsoft Teams Health Check
# Comprehensive diagnosis for Microsoft Teams issues
# Supports both New Teams (2.0) and Classic Teams

param(
    [switch]$Verbose = $false,
    [switch]$Detailed = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     TEAMS HEALTH CHECK" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "     Mode: $(if($Detailed){'Detailed'}else{'Standard'})" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. Teams Installation Check
Write-Host "`n=== TEAMS INSTALLATION ===" -ForegroundColor Yellow

try {
    # Check for New Teams (Windows Store/AppX)
    $newTeams = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
    $teamsVersion = "Nicht gefunden"
    $teamsType = "Unknown"
    $teamsPath = $null
    
    if ($newTeams) {
        $teamsVersion = $newTeams.Version
        $teamsType = "New Teams (AppX)"
        $teamsPath = $newTeams.InstallLocation
        Write-Host "  Teams Version: $teamsVersion" -ForegroundColor Green
        Write-Host "  Teams Type: $teamsType" -ForegroundColor White
        Write-Host "  Teams Path: $teamsPath" -ForegroundColor White
        $info += "Teams: New Teams ($teamsVersion)"
    } else {
        # Check for Classic Teams (Desktop)
        $classicTeamsPaths = @(
            "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk",
            "${env:LOCALAPPDATA}\Microsoft\Teams\Update.exe",
            "${env:PROGRAMFILES}\Microsoft\Teams\current\Teams.exe",
            "${env:PROGRAMFILES(x86)}\Microsoft\Teams\current\Teams.exe"
        )
        
        foreach ($path in $classicTeamsPaths) {
            if (Test-Path $path) {
                if ($path -like "*.lnk") {
                    # Resolve shortcut
                    $shell = New-Object -ComObject WScript.Shell
                    $shortcut = $shell.CreateShortcut($path)
                    $teamsPath = $shortcut.TargetPath
                } else {
                    $teamsPath = $path
                }
                
                if (Test-Path $teamsPath) {
                    try {
                        $fileVersion = (Get-ItemProperty $teamsPath).VersionInfo
                        $teamsVersion = $fileVersion.FileVersion
                        $teamsType = "Classic Teams (Desktop)"
                        Write-Host "  Teams Version: $teamsVersion" -ForegroundColor Green
                        Write-Host "  Teams Type: $teamsType" -ForegroundColor White
                        Write-Host "  Teams Path: $teamsPath" -ForegroundColor White
                        $info += "Teams: Classic Teams ($teamsVersion)"
                        break
                    } catch {
                        Write-Host "  Teams Version: Nicht lesbar" -ForegroundColor Yellow
                        $teamsType = "Classic Teams (Desktop)"
                        $info += "Teams: Classic Teams (Version nicht lesbar)"
                        break
                    }
                }
            }
        }
        
        if (-not $teamsPath) {
            # Check if Teams is running (fallback)
            $runningTeams = Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue
            if ($runningTeams) {
                Write-Host "  Teams: Läuft (Pfad nicht ermittelbar)" -ForegroundColor Yellow
                $warnings += "Teams: Läuft aber Installation nicht gefunden"
                $info += "Teams: Läuft (Installation nicht ermittelbar)"
            } else {
                $issues += "KRITISCH: Teams Installation nicht gefunden!"
                Write-Host "  Teams: Nicht gefunden" -ForegroundColor Red
            }
        }
    }
    
} catch {
    $warnings += "Teams Installation Check fehlgeschlagen"
    Write-Host "  Installation Check: Fehlgeschlagen" -ForegroundColor Red
}

# 2. Teams Process Check
Write-Host "`n=== TEAMS PROCESS ===" -ForegroundColor Yellow

try {
    $teamsProcesses = @()
    
    # Check for New Teams processes
    $newTeamsProcesses = Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue
    if ($newTeamsProcesses) {
        $teamsProcesses += $newTeamsProcesses
    }
    
    # Check for Classic Teams processes
    $classicTeamsProcesses = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
    if ($classicTeamsProcesses) {
        $teamsProcesses += $classicTeamsProcesses
    }
    
    if ($teamsProcesses.Count -gt 0) {
        foreach ($process in $teamsProcesses) {
            $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
            Write-Host "  Teams läuft (PID: $($process.Id), Memory: $memoryMB MB)" -ForegroundColor Green
            
            # Check if Teams is responsive
            try {
                $process.Refresh()
                if ($process.Responding) {
                    Write-Host "    Status: Responsive" -ForegroundColor Green
                } else {
                    $warnings += "Teams läuft aber reagiert nicht"
                    Write-Host "    Status: Nicht responsive" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "    Status: Unbekannt" -ForegroundColor Gray
            }
        }
        $info += "Teams: Aktiv ($($teamsProcesses.Count) Prozess(e))"
    } else {
        Write-Host "  Teams: Nicht aktiv" -ForegroundColor Yellow
        $warnings += "Teams: Nicht aktiv"
    }
} catch {
    Write-Host "  Process Check: Fehlgeschlagen" -ForegroundColor Red
}

# 3. Teams Cache Check
Write-Host "`n=== TEAMS CACHE ===" -ForegroundColor Yellow

try {
    $cachePaths = @(
        "$env:APPDATA\Microsoft\Teams",
        "$env:LOCALAPPDATA\Microsoft\Teams",
        "$env:APPDATA\Microsoft\Teams\Cache",
        "$env:APPDATA\Microsoft\Teams\Application Cache",
        "$env:APPDATA\Microsoft\Teams\blob_storage",
        "$env:APPDATA\Microsoft\Teams\databases",
        "$env:APPDATA\Microsoft\Teams\IndexedDB",
        "$env:APPDATA\Microsoft\Teams\Local Storage",
        "$env:APPDATA\Microsoft\Teams\Session Storage"
    )
    
    $totalCacheSize = 0
    $cacheFound = $false
    
    foreach ($cachePath in $cachePaths) {
        if (Test-Path $cachePath) {
            $cacheSize = (Get-ChildItem $cachePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $totalCacheSize += $cacheSize
            $cacheFound = $true
            
            $cacheSizeMB = [math]::Round($cacheSize/1MB, 2)
            $folderName = Split-Path $cachePath -Leaf
            Write-Host "  ${folderName}: $cacheSizeMB MB" -ForegroundColor White
        }
    }
    
    if ($cacheFound) {
        $totalCacheMB = [math]::Round($totalCacheSize/1MB, 2)
        Write-Host "  Gesamt Cache: $totalCacheMB MB" -ForegroundColor White
        
        if ($totalCacheMB -gt 1000) {
            $warnings += "Teams Cache groß: $totalCacheMB MB (Cleanup empfohlen)"
            Write-Host "  Status: GROSS - Cleanup empfohlen" -ForegroundColor Yellow
        } elseif ($totalCacheMB -gt 500) {
            $warnings += "Teams Cache mittel: $totalCacheMB MB (Cleanup optional)"
            Write-Host "  Status: MITTEL - Cleanup optional" -ForegroundColor Yellow
        } else {
            Write-Host "  Status: OK" -ForegroundColor Green
        }
        $info += "Teams Cache: $totalCacheMB MB"
    } else {
        # Check if Teams is running but cache not found (VDI scenario)
        $runningTeams = Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue
        if ($runningTeams) {
            Write-Host "  Cache: Nicht gefunden (Teams läuft - möglicherweise VDI/Remote)" -ForegroundColor Yellow
            $warnings += "Teams Cache: Nicht gefunden (Teams läuft - VDI/Remote möglich)"
        } else {
            Write-Host "  Cache: Nicht gefunden" -ForegroundColor Gray
        }
    }
    
} catch {
    Write-Host "  Cache Check: Fehlgeschlagen" -ForegroundColor Red
}

# 4. Teams Settings & Configuration
Write-Host "`n=== TEAMS CONFIGURATION ===" -ForegroundColor Yellow

try {
    $teamsConfigPath = "$env:APPDATA\Microsoft\Teams"
    $settingsFile = "$teamsConfigPath\settings.json"
    
    if (Test-Path $settingsFile) {
        try {
            $settings = Get-Content $settingsFile -Raw | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($settings) {
                Write-Host "  Settings: Gefunden" -ForegroundColor Green
                
                # Check for specific settings
                if ($settings.appPreferenceSettings) {
                    Write-Host "    App Preferences: Konfiguriert" -ForegroundColor White
                }
                if ($settings.userPreferenceSettings) {
                    Write-Host "    User Preferences: Konfiguriert" -ForegroundColor White
                }
                $info += "Teams: Settings konfiguriert"
            } else {
                Write-Host "  Settings: Leer oder beschädigt" -ForegroundColor Yellow
                $warnings += "Teams: Settings leer oder beschädigt"
            }
        } catch {
            Write-Host "  Settings: Nicht lesbar" -ForegroundColor Yellow
            $warnings += "Teams: Settings nicht lesbar"
        }
    } else {
        # Check if Teams is running but settings not found (VDI scenario)
        $runningTeams = Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue
        if ($runningTeams) {
            Write-Host "  Settings: Nicht gefunden (Teams läuft - möglicherweise VDI/Remote)" -ForegroundColor Yellow
            $warnings += "Teams Settings: Nicht gefunden (Teams läuft - VDI/Remote möglich)"
        } else {
            Write-Host "  Settings: Nicht gefunden" -ForegroundColor Yellow
            $warnings += "Teams: Settings nicht gefunden"
        }
    }
    
    # Check for user data
    $userDataPath = "$teamsConfigPath\storage"
    if (Test-Path $userDataPath) {
        $userDataSize = (Get-ChildItem $userDataPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $userDataMB = [math]::Round($userDataSize/1MB, 2)
        Write-Host "  User Data: $userDataMB MB" -ForegroundColor White
        $info += "Teams User Data: $userDataMB MB"
    } else {
        Write-Host "  User Data: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Configuration Check: Fehlgeschlagen" -ForegroundColor Red
}

# 5. Teams Services Check
Write-Host "`n=== TEAMS SERVICES ===" -ForegroundColor Yellow

try {
    $teamsServices = @(
        "TeamsMachineInstaller",
        "TeamsUpdaterService"
    )
    
    $runningServices = 0
    foreach ($serviceName in $teamsServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.Status -eq "Running") {
                Write-Host "  ${serviceName}: $($service.Status)" -ForegroundColor Green
                $runningServices++
            } else {
                Write-Host "  ${serviceName}: $($service.Status)" -ForegroundColor Yellow
            }
        }
    }
    
    if ($runningServices -gt 0) {
        $info += "Teams Services: $runningServices von $($teamsServices.Count) aktiv"
    } else {
        Write-Host "  Teams Services: Keine gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Services Check: Fehlgeschlagen" -ForegroundColor Red
}

# 6. WebView2 Check (for New Teams)
Write-Host "`n=== WEBVIEW2 CHECK ===" -ForegroundColor Yellow

try {
    $webView2Paths = @(
        "$env:LOCALAPPDATA\Microsoft\Edge\WebView2",
        "$env:PROGRAMFILES\Microsoft\Edge\Application",
        "$env:PROGRAMFILES(x86)}\Microsoft\Edge\Application"
    )
    
    $webView2Found = $false
    foreach ($path in $webView2Paths) {
        if (Test-Path $path) {
            $webView2Found = $true
            $webView2Size = (Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $webView2MB = [math]::Round($webView2Size/1MB, 2)
            $folderName = Split-Path $path -Leaf
            Write-Host "  ${folderName}: $webView2MB MB" -ForegroundColor White
        }
    }
    
    if ($webView2Found) {
        $info += "WebView2: Gefunden"
    } else {
        Write-Host "  WebView2: Nicht gefunden" -ForegroundColor Yellow
        $warnings += "WebView2: Nicht gefunden (benötigt für New Teams)"
    }
    
} catch {
    Write-Host "  WebView2 Check: Fehlgeschlagen" -ForegroundColor Red
}

# 7. Teams Logs Check
Write-Host "`n=== TEAMS LOGS ===" -ForegroundColor Yellow

try {
    $logsPath = "$env:APPDATA\Microsoft\Teams\logs"
    if (Test-Path $logsPath) {
        $logFiles = Get-ChildItem $logsPath -Filter "*.log" -ErrorAction SilentlyContinue
        if ($logFiles) {
            $logCount = $logFiles.Count
            $totalLogSize = ($logFiles | Measure-Object -Property Length -Sum).Sum
            $logSizeMB = [math]::Round($totalLogSize/1MB, 2)
            
            Write-Host "  Log Files: $logCount gefunden" -ForegroundColor Green
            Write-Host "  Log Size: $logSizeMB MB" -ForegroundColor White
            
            # Check for recent errors
            $recentLogs = $logFiles | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-1) }
            if ($recentLogs) {
                Write-Host "  Recent Logs: $($recentLogs.Count) in letzten 24h" -ForegroundColor White
            }
            
            $info += "Teams Logs: $logCount Dateien ($logSizeMB MB)"
        } else {
            Write-Host "  Log Files: Keine gefunden" -ForegroundColor Gray
        }
    } else {
        Write-Host "  Logs Path: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Logs Check: Fehlgeschlagen" -ForegroundColor Red
}

# 8. Network Connectivity (Detailed Mode)
if ($Detailed) {
    Write-Host "`n=== NETWORK CONNECTIVITY ===" -ForegroundColor Yellow
    
    try {
        # Teams/Office365 Endpoints
        $teamsEndpoints = @(
            "teams.microsoft.com",
            "teams.live.com",
            "teams.events.data.microsoft.com",
            "teams.office.com"
        )
        
        foreach ($endpoint in $teamsEndpoints) {
            try {
                $result = Test-NetConnection -ComputerName $endpoint -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
                if ($result) {
                    Write-Host "  ${endpoint}: Erreichbar" -ForegroundColor Green
                } else {
                    Write-Host "  ${endpoint}: Nicht erreichbar" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "  ${endpoint}: Check fehlgeschlagen" -ForegroundColor Gray
            }
        }
        
    } catch {
        Write-Host "  Network Check: Fehlgeschlagen" -ForegroundColor Red
    }
}

# 9. Teams Performance Check
Write-Host "`n=== TEAMS PERFORMANCE ===" -ForegroundColor Yellow

try {
    $teamsProcesses = Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue
    if ($teamsProcesses) {
        $totalMemory = ($teamsProcesses | Measure-Object -Property WorkingSet -Sum).Sum
        $totalMemoryMB = [math]::Round($totalMemory/1MB, 2)
        
        Write-Host "  Total Memory: $totalMemoryMB MB" -ForegroundColor White
        
        if ($totalMemoryMB -gt 1000) {
            $warnings += "Teams: Hoher Memory-Verbrauch ($totalMemoryMB MB)"
            Write-Host "  Status: HOCH - Memory-Optimierung empfohlen" -ForegroundColor Yellow
        } elseif ($totalMemoryMB -gt 500) {
            Write-Host "  Status: MITTEL" -ForegroundColor Yellow
        } else {
            Write-Host "  Status: OK" -ForegroundColor Green
        }
        
        $info += "Teams Memory: $totalMemoryMB MB"
    } else {
        Write-Host "  Performance: Teams nicht aktiv" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Performance Check: Fehlgeschlagen" -ForegroundColor Red
}

# 10. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     TEAMS HEALTH SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

if ($issues.Count -gt 0) {
    Write-Host "`n🚨 KRITISCHE PROBLEME ($($issues.Count)):" -ForegroundColor Red
    foreach ($issue in $issues) {
        Write-Host "  • $issue" -ForegroundColor Red
    }
}

if ($warnings.Count -gt 0) {
    Write-Host "`n⚠️ WARNUNGEN ($($warnings.Count)):" -ForegroundColor Yellow
    foreach ($warning in $warnings) {
        Write-Host "  • $warning" -ForegroundColor Yellow
    }
}

if ($info.Count -gt 0) {
    Write-Host "`nℹ️ INFORMATIONEN ($($info.Count)):" -ForegroundColor Cyan
    foreach ($infoItem in $info) {
        Write-Host "  • $infoItem" -ForegroundColor Cyan
    }
}

# Overall Status
Write-Host "`n📊 GESAMTSTATUS:" -ForegroundColor White
if ($issues.Count -gt 0) {
    Write-Host "  Status: KRITISCH - $($issues.Count) kritische Probleme gefunden!" -ForegroundColor Red
} elseif ($warnings.Count -gt 0) {
    Write-Host "  Status: WARNUNG - $($warnings.Count) Warnungen gefunden" -ForegroundColor Yellow
} else {
    Write-Host "  Status: OK - Teams scheint gesund zu sein" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Teams Health Check completed!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# Clean exit - change to neutral system directory
Set-Location "C:\Windows\System32" | Out-Null

if ($Verbose) {
    Read-Host "`nDrücken Sie Enter zum Beenden..."
} else {
    # Brief pause to ensure output is visible
    Start-Sleep -Milliseconds 500
}

exit 0
