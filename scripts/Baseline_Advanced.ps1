# Advanced System Baseline Check
# Comprehensive system health assessment for detailed troubleshooting
# Extended checks for critical issues with detailed reporting

param(
    [switch]$Verbose = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     Advanced System Baseline Check" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. System Health Check (erweitert)
Write-Host "`n=== SYSTEM HEALTH (ADVANCED) ===" -ForegroundColor Yellow

# RAM Check (detailliert)
try {
    $totalRAM = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory/1GB, 2)
    $freeRAM = [math]::Round((Get-WmiObject -Class Win32_OperatingSystem).FreePhysicalMemory/1MB, 2)
    $usedRAM = $totalRAM - $freeRAM
    $ramPercent = [math]::Round(($usedRAM/$totalRAM)*100, 2)
    
    Write-Host "  RAM: $totalRAM GB total, $usedRAM GB used, $freeRAM GB free ($ramPercent%)" -ForegroundColor White
    
    if ($ramPercent -gt 90) {
        $issues += "KRITISCH: RAM Usage $ramPercent% - VDI Performance Problem!"
        Write-Host "  Status: KRITISCH" -ForegroundColor Red
    } elseif ($ramPercent -gt 80) {
        $warnings += "WARNUNG: RAM Usage $ramPercent% - Performance könnte beeinträchtigt sein"
        Write-Host "  Status: WARNUNG" -ForegroundColor Yellow
    } else {
        Write-Host "  Status: OK" -ForegroundColor Green
    }
} catch {
    $warnings += "RAM Check fehlgeschlagen"
    Write-Host "  RAM: Check fehlgeschlagen" -ForegroundColor Gray
}

# CPU Check (detailliert)
try {
    $cpuUsage = Get-WmiObject -Class Win32_Processor | Measure-Object -Property LoadPercentage -Average
    $cpuCount = (Get-WmiObject -Class Win32_Processor).Count
    $cpuName = (Get-WmiObject -Class Win32_Processor).Name
    
    Write-Host "  CPU: $cpuName ($cpuCount cores)" -ForegroundColor White
    Write-Host "  CPU Usage: $($cpuUsage.Average)%" -ForegroundColor White
    
    if ($cpuUsage.Average -gt 80) {
        $issues += "KRITISCH: CPU Usage $($cpuUsage.Average)% - VDI Performance Problem!"
        Write-Host "  Status: KRITISCH" -ForegroundColor Red
    } elseif ($cpuUsage.Average -gt 60) {
        $warnings += "WARNUNG: CPU Usage $($cpuUsage.Average)% - Performance könnte beeinträchtigt sein"
        Write-Host "  Status: WARNUNG" -ForegroundColor Yellow
    } else {
        Write-Host "  Status: OK" -ForegroundColor Green
    }
} catch {
    $warnings += "CPU Check fehlgeschlagen"
    Write-Host "  CPU: Check fehlgeschlagen" -ForegroundColor Gray
}

# Disk Check (alle Laufwerke)
try {
    $disks = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    Write-Host "  Disk Usage:" -ForegroundColor White
    foreach ($disk in $disks) {
        $sizeGB = [math]::Round($disk.Size/1GB, 2)
        $freeGB = [math]::Round($disk.FreeSpace/1GB, 2)
        $percentFree = [math]::Round(($disk.FreeSpace/$disk.Size)*100, 2)
        
        Write-Host "    $($disk.DeviceID) - $sizeGB GB total, $freeGB GB free ($percentFree%)" -ForegroundColor White
        
        if ($percentFree -lt 5) {
            $issues += "KRITISCH: $($disk.DeviceID) fast voll! ($percentFree% frei)"
        } elseif ($percentFree -lt 15) {
            $warnings += "WARNUNG: $($disk.DeviceID) wird voll ($percentFree% frei)"
        }
    }
} catch {
    $warnings += "Disk Check fehlgeschlagen"
    Write-Host "  Disk: Check fehlgeschlagen" -ForegroundColor Gray
}

# 2. P-Laufwerk Check (erweitert)
Write-Host "`n=== P-LAUFWERK CHECK (ADVANCED) ===" -ForegroundColor Yellow
try {
    $pDrive = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='P:'"
    if ($pDrive) {
        $pSizeGB = [math]::Round($pDrive.Size/1GB, 2)
        $pFreeGB = [math]::Round($pDrive.FreeSpace/1GB, 2)
        $pPercentFree = [math]::Round(($pDrive.FreeSpace/$pDrive.Size)*100, 2)
        
        Write-Host "  P: - $pSizeGB GB total, $pFreeGB GB frei ($pPercentFree%)" -ForegroundColor White
        
        # Standard 5GB Check
        if ($pSizeGB -eq 5) {
            $info += "P-Laufwerk: Standard 5GB (nicht erweitert)"
            Write-Host "  Status: Standard 5GB" -ForegroundColor Green
        } elseif ($pSizeGB -gt 5) {
            $info += "P-Laufwerk: Erweitert auf $pSizeGB GB (User hat Bestellung aufgegeben)"
            Write-Host "  Status: Erweitert auf $pSizeGB GB" -ForegroundColor Cyan
        } else {
            $warnings += "P-Laufwerk: $pSizeGB GB (kleiner als Standard 5GB)"
            Write-Host "  Status: $pSizeGB GB (kleiner als Standard)" -ForegroundColor Yellow
        }
        
        # Speicherplatz-Warnungen
        if ($pPercentFree -lt 5) {
            $issues += "KRITISCH: P-Laufwerk fast voll! ($pPercentFree% frei)"
            Write-Host "  Speicherplatz: KRITISCH ($pPercentFree% frei)" -ForegroundColor Red
        } elseif ($pPercentFree -lt 15) {
            $warnings += "WARNUNG: P-Laufwerk wird voll ($pPercentFree% frei)"
            Write-Host "  Speicherplatz: WARNUNG ($pPercentFree% frei)" -ForegroundColor Yellow
        } else {
            Write-Host "  Speicherplatz: OK ($pPercentFree% frei)" -ForegroundColor Green
        }
        
        # Screenpresso Check (bekanntes Problem)
        $screenpressoPath = "P:\Screenpresso"
        if (Test-Path $screenpressoPath) {
            $screenpressoSize = (Get-ChildItem $screenpressoPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $screenpressoSizeGB = [math]::Round($screenpressoSize/1GB, 2)
            Write-Host "  Screenpresso: $screenpressoSizeGB GB" -ForegroundColor White
            if ($screenpressoSizeGB -gt 2) {
                $warnings += "WARNUNG: Screenpresso verbraucht viel Speicherplatz ($screenpressoSizeGB GB)"
                Write-Host "    Screenpresso: GROSS ($screenpressoSizeGB GB)" -ForegroundColor Yellow
            }
        }
        
        # P-Laufwerk Inhalt Analyse
        Write-Host "  P-Laufwerk Inhalt:" -ForegroundColor White
        $pContents = Get-ChildItem "P:\" -ErrorAction SilentlyContinue | Sort-Object Length -Descending | Select-Object -First 10
        foreach ($item in $pContents) {
            $itemSizeGB = [math]::Round($item.Length/1GB, 2)
            Write-Host "    $($item.Name): $itemSizeGB GB" -ForegroundColor White
        }
        
    } else {
        $info += "P-Laufwerk: Nicht verfügbar (Administrativer Account)"
        Write-Host "  P: - Nicht verfügbar (Admin Account)" -ForegroundColor Gray
    }
} catch {
    $warnings += "P-Laufwerk Check fehlgeschlagen"
    Write-Host "  P-Laufwerk: Check fehlgeschlagen" -ForegroundColor Red
}

# 3. Healthcare Apps Check (erweitert)
Write-Host "`n=== HEALTHCARE APPS (ADVANCED) ===" -ForegroundColor Yellow

# Teams Check (detailliert)
try {
    $teamsProcess = Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue
    if ($teamsProcess) {
        $teamsMemory = [math]::Round($teamsProcess.WorkingSet/1MB, 2)
        Write-Host "  Teams: Läuft (PID: $($teamsProcess.Id), Memory: $teamsMemory MB)" -ForegroundColor Green
        $info += "Teams: Aktiv ($teamsMemory MB)"
        
        # Teams Cache Check
        $teamsCachePaths = @(
            "$env:APPDATA\Microsoft\Teams",
            "$env:LOCALAPPDATA\Microsoft\Teams"
        )
        
        $totalTeamsCache = 0
        foreach ($cachePath in $teamsCachePaths) {
            if (Test-Path $cachePath) {
                $cacheSize = (Get-ChildItem $cachePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                $totalTeamsCache += $cacheSize
            }
        }
        
        if ($totalTeamsCache -gt 0) {
            $teamsCacheGB = [math]::Round($totalTeamsCache/1GB, 2)
            Write-Host "    Teams Cache: $teamsCacheGB GB" -ForegroundColor White
            if ($teamsCacheGB -gt 2) {
                $warnings += "Teams Cache groß: $teamsCacheGB GB (Cleanup empfohlen)"
                Write-Host "      Teams Cache: GROSS ($teamsCacheGB GB)" -ForegroundColor Yellow
            }
        }
    } else {
        $warnings += "Teams: Nicht aktiv (Kommunikation möglicherweise eingeschränkt)"
        Write-Host "  Teams: Nicht aktiv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Teams: Check fehlgeschlagen" -ForegroundColor Gray
}

# Outlook Check (detailliert)
try {
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        $outlookMemory = [math]::Round($outlookProcess.WorkingSet/1MB, 2)
        Write-Host "  Outlook: Läuft (PID: $($outlookProcess.Id), Memory: $outlookMemory MB)" -ForegroundColor Green
        $info += "Outlook: Aktiv ($outlookMemory MB)"
        
        # Outlook OST Check
        $ostPath = "$env:LOCALAPPDATA\Microsoft\Outlook"
        if (Test-Path $ostPath) {
            $ostFiles = Get-ChildItem $ostPath -Filter "*.ost" -ErrorAction SilentlyContinue
            if ($ostFiles) {
                Write-Host "    Outlook OST Files:" -ForegroundColor White
                foreach ($ost in $ostFiles) {
                    $ostSizeGB = [math]::Round($ost.Length/1GB, 2)
                    Write-Host "      $($ost.Name): $ostSizeGB GB" -ForegroundColor White
                    if ($ostSizeGB -gt 2) {
                        $warnings += "Outlook OST groß: $($ost.Name) - $ostSizeGB GB (Cleanup empfohlen)"
                        Write-Host "        OST: GROSS ($ostSizeGB GB)" -ForegroundColor Yellow
                    }
                }
            }
        }
    } else {
        $warnings += "Outlook: Nicht aktiv (Email möglicherweise eingeschränkt)"
        Write-Host "  Outlook: Nicht aktiv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Outlook: Check fehlgeschlagen" -ForegroundColor Gray
}

# OneDrive Check (detailliert)
try {
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDriveProcess) {
        $oneDriveMemory = [math]::Round($oneDriveProcess.WorkingSet/1MB, 2)
        Write-Host "  OneDrive: Läuft (PID: $($oneDriveProcess.Id), Memory: $oneDriveMemory MB)" -ForegroundColor Green
        $info += "OneDrive: Aktiv ($oneDriveMemory MB)"
        
        # OneDrive Sync Status
        $oneDrivePath = "$env:USERPROFILE\OneDrive - USZ"
        if (Test-Path $oneDrivePath) {
            Write-Host "    OneDrive Path: Gefunden" -ForegroundColor Green
        } else {
            Write-Host "    OneDrive Path: Nicht gefunden" -ForegroundColor Yellow
        }
    } else {
        $warnings += "OneDrive: Nicht aktiv (File Sync möglicherweise eingeschränkt)"
        Write-Host "  OneDrive: Nicht aktiv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  OneDrive: Check fehlgeschlagen" -ForegroundColor Gray
}

# 4. Omnissa Check (Healthcare-spezifisch) - DEAKTIVIERT
# Der Omnissa Check wurde deaktiviert, da er Permission-Fehler verursacht
# und sehr spezifisch für bestimmte VDI-Umgebungen ist.
# 
# Write-Host "`n=== OMNISSA CHECK (HEALTHCARE) ===" -ForegroundColor Yellow
# try {
#     # Omnissa Service Check
#     $omnissaServices = Get-Service | Where-Object { 
#         $_.Name -like "*omnissa*" -or $_.DisplayName -like "*omnissa*" 
#     }
#     if ($omnissaServices) {
#         Write-Host "  Omnissa Services:" -ForegroundColor White
#         foreach ($service in $omnissaServices) {
#             Write-Host "    $($service.Name): $($service.Status)" -ForegroundColor White
#         }
#         $info += "Omnissa: Services gefunden"
#     } else {
#         Write-Host "  Omnissa Services: Nicht gefunden" -ForegroundColor Gray
#     }
#     
#     # Omnissa Process Check
#     $omnissaProcesses = Get-Process | Where-Object { 
#         $_.ProcessName -like "*omnissa*" 
#     }
#     if ($omnissaProcesses) {
#         Write-Host "  Omnissa Prozesse:" -ForegroundColor White
#         foreach ($process in $omnissaProcesses) {
#             $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
#             Write-Host "    $($process.ProcessName) (PID: $($process.Id), Memory: $memoryMB MB)" -ForegroundColor White
#         }
#         $info += "Omnissa: Prozesse aktiv"
#     } else {
#         Write-Host "  Omnissa Prozesse: Nicht aktiv" -ForegroundColor Gray
#     }
#     
#     # Omnissa Registry Check
#     $omnissaRegPaths = @(
#         "HKLM:\SOFTWARE\Omnissa",
#         "HKLM:\SOFTWARE\WOW6432Node\Omnissa",
#         "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
#     )
#     
#     $omnissaFound = $false
#     foreach ($regPath in $omnissaRegPaths) {
#         try {
#             if ($regPath -like "*Uninstall*") {
#                 $items = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { 
#                     $_.DisplayName -like "*omnissa*" 
#                 }
#                 if ($items) {
#                     Write-Host "  Omnissa Installation:" -ForegroundColor White
#                     foreach ($item in $items) {
#                         Write-Host "    $($item.DisplayName) - $($item.DisplayVersion)" -ForegroundColor White
#                     }
#                     $omnissaFound = $true
#                 }
#             } else {
#                 $reg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
#                 if ($reg) {
#                     Write-Host "  Omnissa Registry: $regPath gefunden" -ForegroundColor Green
#                     $omnissaFound = $true
#                 }
#             }
#         } catch { }
#     }
#     
#     if (-not $omnissaFound) {
#         Write-Host "  Omnissa Registry: Nicht gefunden" -ForegroundColor Gray
#     }
#     
#     # Gesamtstatus
#     if ($omnissaServices -or $omnissaProcesses -or $omnissaFound) {
#         $info += "Omnissa: Gefunden"
#     } else {
#         $warnings += "Omnissa: Nicht gefunden (Healthcare-System möglicherweise nicht verfügbar)"
#     }
#     
# } catch {
#     Write-Host "  Omnissa Check: Fehlgeschlagen" -ForegroundColor Gray
# }

# 5. VDI Environment Check
Write-Host "`n=== VDI ENVIRONMENT ===" -ForegroundColor Yellow
try {
    $vmwareTools = Get-Service -Name "VMTools" -ErrorAction SilentlyContinue
    if ($vmwareTools) {
        Write-Host "  VMware Tools: $($vmwareTools.Status)" -ForegroundColor Green
        $info += "VDI: VMware VDI erkannt"
    } else {
        Write-Host "  VMware Tools: Nicht gefunden" -ForegroundColor Gray
    }
    
    # Check for VDI processes
    $vdiProcesses = @("vmtoolsd", "citrix", "teradici")
    $foundVDI = $false
    foreach ($process in $vdiProcesses) {
        $proc = Get-Process | Where-Object { $_.ProcessName -like "*$process*" }
        if ($proc) {
            $foundVDI = $true
            Write-Host "  VDI Prozess: $($proc.ProcessName)" -ForegroundColor Green
        }
    }
    
    if (-not $foundVDI) {
        Write-Host "  VDI Prozesse: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  VDI Check: Fehlgeschlagen" -ForegroundColor Gray
}

# 5. Security Check (VDI-optimiert)
Write-Host "`n=== SECURITY CHECK ===" -ForegroundColor Yellow

# UAC Check
try {
    $uacReg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -ErrorAction SilentlyContinue
    if ($uacReg.EnableLUA -eq 1) {
        Write-Host "  UAC: Aktiviert" -ForegroundColor Green
        $info += "UAC: Aktiviert (Standard)"
    } else {
        $issues += "KRITISCH: UAC deaktiviert (NICHT STANDARD!)"
        Write-Host "  UAC: Deaktiviert (KRITISCH)" -ForegroundColor Red
    }
} catch {
    Write-Host "  UAC: Check fehlgeschlagen" -ForegroundColor Gray
}

# Domain Check
try {
    $domain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
    if ($domain -and $domain -ne $env:COMPUTERNAME) {
        Write-Host "  Domain: $domain" -ForegroundColor Green
        $info += "Domain: $domain"
    } else {
        $warnings += "Domain: Nicht verfügbar oder nicht korrekt konfiguriert"
        Write-Host "  Domain: Nicht verfügbar" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Domain: Check fehlgeschlagen" -ForegroundColor Gray
}

# 6. Advanced Checks (nur wenn -Advanced)
if ($Advanced) {
    Write-Host "`n=== ADVANCED CHECKS ===" -ForegroundColor Yellow
    
    # Teams Cache Check
    try {
        $teamsCachePaths = @(
            "$env:APPDATA\Microsoft\Teams",
            "$env:LOCALAPPDATA\Microsoft\Teams"
        )
        
        $totalTeamsCache = 0
        foreach ($cachePath in $teamsCachePaths) {
            if (Test-Path $cachePath) {
                $cacheSize = (Get-ChildItem $cachePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                $totalTeamsCache += $cacheSize
            }
        }
        
        if ($totalTeamsCache -gt 0) {
            $teamsCacheGB = [math]::Round($totalTeamsCache/1GB, 2)
            Write-Host "  Teams Cache: $teamsCacheGB GB" -ForegroundColor White
            if ($teamsCacheGB -gt 2) {
                $warnings += "Teams Cache groß: $teamsCacheGB GB (Cleanup empfohlen)"
                Write-Host "    Teams Cache: GROSS ($teamsCacheGB GB)" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  Teams Cache: Check fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Omnissa Check (Healthcare-spezifisch)
    Write-Host "`n  === OMNISSA CHECK (HEALTHCARE) ===" -ForegroundColor Yellow
    try {
        # Omnissa Service Check
        $omnissaServices = Get-Service | Where-Object { 
            $_.Name -like "*omnissa*" -or $_.DisplayName -like "*omnissa*" 
        }
        if ($omnissaServices) {
            Write-Host "    Omnissa Services:" -ForegroundColor White
            foreach ($service in $omnissaServices) {
                Write-Host "      $($service.Name): $($service.Status)" -ForegroundColor White
            }
            $info += "Omnissa: Services gefunden"
        } else {
            Write-Host "    Omnissa Services: Nicht gefunden" -ForegroundColor Gray
        }
        
        # Omnissa Process Check
        $omnissaProcesses = Get-Process | Where-Object { 
            $_.ProcessName -like "*omnissa*" 
        }
        if ($omnissaProcesses) {
            Write-Host "    Omnissa Prozesse:" -ForegroundColor White
            foreach ($process in $omnissaProcesses) {
                $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
                Write-Host "      $($process.ProcessName) (PID: $($process.Id), Memory: $memoryMB MB)" -ForegroundColor White
            }
            $info += "Omnissa: Prozesse aktiv"
        } else {
            Write-Host "    Omnissa Prozesse: Nicht aktiv" -ForegroundColor Gray
        }
        
        # Omnissa Registry Check
        $omnissaRegPaths = @(
            "HKLM:\SOFTWARE\Omnissa",
            "HKLM:\SOFTWARE\WOW6432Node\Omnissa",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        $omnissaFound = $false
        foreach ($regPath in $omnissaRegPaths) {
            try {
                if ($regPath -like "*Uninstall*") {
                    $items = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { 
                        $_.DisplayName -like "*omnissa*" 
                    }
                    if ($items) {
                        Write-Host "    Omnissa Installation:" -ForegroundColor White
                        foreach ($item in $items) {
                            Write-Host "      $($item.DisplayName) - $($item.DisplayVersion)" -ForegroundColor White
                        }
                        $omnissaFound = $true
                    }
                } else {
                    $reg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                    if ($reg) {
                        Write-Host "    Omnissa Registry: $regPath gefunden" -ForegroundColor Green
                        $omnissaFound = $true
                    }
                }
            } catch { }
        }
        
        if (-not $omnissaFound) {
            Write-Host "    Omnissa Registry: Nicht gefunden" -ForegroundColor Gray
        }
        
        # Gesamtstatus
        if ($omnissaServices -or $omnissaProcesses -or $omnissaFound) {
            $info += "Omnissa: Gefunden"
        } else {
            $warnings += "Omnissa: Nicht gefunden (Healthcare-System möglicherweise nicht verfügbar)"
        }
        
    } catch {
        Write-Host "    Omnissa Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Java Check
    try {
        $javaVersion = java -version 2>&1 | Select-String "version"
        if ($javaVersion) {
            Write-Host "  Java: $($javaVersion.Line)" -ForegroundColor White
            $info += "Java: $($javaVersion.Line)"
        } else {
            $warnings += "Java: Nicht verfügbar (Healthcare-Apps benötigen möglicherweise Java)"
            Write-Host "  Java: Nicht verfügbar" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  Java: Check fehlgeschlagen" -ForegroundColor Gray
    }
}

# 7. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     BASELINE SUMMARY" -ForegroundColor Cyan
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
    Write-Host "  Status: OK - Keine kritischen Probleme gefunden" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Baseline Check completed!" -ForegroundColor Green
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