# System Baseline Check
# Quick system health assessment for IT support and troubleshooting
# Focuses on critical issues with minimal performance impact

param(
    [switch]$Advanced = $false,
    [switch]$Verbose = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     System Baseline Check" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "     Mode: $(if($Advanced){'Advanced'}else{'Standard'})" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. System Health Check
Write-Host "`n=== SYSTEM HEALTH ===" -ForegroundColor Yellow

# RAM Check
try {
    $totalRAM = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory/1GB, 2)
    $freeRAM = [math]::Round((Get-WmiObject -Class Win32_OperatingSystem).FreePhysicalMemory/1MB, 2)
    $usedRAM = $totalRAM - $freeRAM
    $ramPercent = [math]::Round(($usedRAM/$totalRAM)*100, 2)
    
    if ($ramPercent -gt 90) {
        $issues += "CRITICAL: RAM Usage $ramPercent% - Performance issue detected!"
        Write-Host "  RAM: $ramPercent% (CRITICAL)" -ForegroundColor Red
    } elseif ($ramPercent -gt 80) {
        $warnings += "WARNING: RAM Usage $ramPercent% - Performance may be impacted"
        Write-Host "  RAM: $ramPercent% (WARNING)" -ForegroundColor Yellow
    } else {
        Write-Host "  RAM: $ramPercent% (OK)" -ForegroundColor Green
    }
} catch {
    $warnings += "RAM Check failed"
    Write-Host "  RAM: Check failed" -ForegroundColor Gray
}

# CPU Check
try {
    $cpuUsage = Get-WmiObject -Class Win32_Processor | Measure-Object -Property LoadPercentage -Average
    if ($cpuUsage.Average -gt 80) {
        $issues += "CRITICAL: CPU Usage $($cpuUsage.Average)% - Performance issue detected!"
        Write-Host "  CPU: $($cpuUsage.Average)% (CRITICAL)" -ForegroundColor Red
    } elseif ($cpuUsage.Average -gt 60) {
        $warnings += "WARNING: CPU Usage $($cpuUsage.Average)% - Performance may be impacted"
        Write-Host "  CPU: $($cpuUsage.Average)% (WARNING)" -ForegroundColor Yellow
    } else {
        Write-Host "  CPU: $($cpuUsage.Average)% (OK)" -ForegroundColor Green
    }
} catch {
    $warnings += "CPU Check failed"
    Write-Host "  CPU: Check failed" -ForegroundColor Gray
}

# 2. Home Drive Check (P: Drive)
Write-Host "`n=== P-LAUFWERK CHECK ===" -ForegroundColor Yellow
try {
    $pDrive = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='P:'"
    if ($pDrive) {
        $pSizeGB = [math]::Round($pDrive.Size/1GB, 2)
        $pFreeGB = [math]::Round($pDrive.FreeSpace/1GB, 2)
        $pPercentFree = [math]::Round(($pDrive.FreeSpace/$pDrive.Size)*100, 2)
        
        Write-Host "  P: - $pSizeGB GB total, $pFreeGB GB frei ($pPercentFree%)" -ForegroundColor White
        
        # Standard 5GB Check
        if ($pSizeGB -eq 5) {
            $info += "Home Drive: Standard 5GB (not extended)"
            Write-Host "  Status: Standard 5GB" -ForegroundColor Green
        } elseif ($pSizeGB -gt 5) {
            $info += "Home Drive: Extended to $pSizeGB GB (User requested extension)"
            Write-Host "  Status: Erweitert auf $pSizeGB GB" -ForegroundColor Cyan
        } else {
            $warnings += "Home Drive: $pSizeGB GB (smaller than standard 5GB)"
            Write-Host "  Status: $pSizeGB GB (kleiner als Standard)" -ForegroundColor Yellow
        }
        
        # Speicherplatz-Warnungen
        if ($pPercentFree -lt 5) {
            $issues += "CRITICAL: Home drive almost full! ($pPercentFree% free)"
            Write-Host "  Free Space: CRITICAL ($pPercentFree% free)" -ForegroundColor Red
        } elseif ($pPercentFree -lt 15) {
            $warnings += "WARNING: Home drive getting full ($pPercentFree% free)"
            Write-Host "  Free Space: WARNING ($pPercentFree% free)" -ForegroundColor Yellow
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
                $warnings += "WARNING: Screenpresso consuming significant space ($screenpressoSizeGB GB)"
                Write-Host "    Screenpresso: GROSS ($screenpressoSizeGB GB)" -ForegroundColor Yellow
            }
        }
        
    } else {
        $info += "Home Drive: Not available (Administrative Account)"
        Write-Host "  P: - Nicht verfügbar (Admin Account)" -ForegroundColor Gray
    }
} catch {
    $warnings += "Home Drive Check failed"
    Write-Host "  Home Drive: Check failed" -ForegroundColor Red
}

# 3. Enterprise Apps Check
Write-Host "`n=== HEALTHCARE APPS ===" -ForegroundColor Yellow

# Teams Check
try {
    $teamsProcess = Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue
    if ($teamsProcess) {
        Write-Host "  Teams: Läuft (PID: $($teamsProcess.Id))" -ForegroundColor Green
        $info += "Teams: Aktiv"
    } else {
        $warnings += "Teams: Nicht aktiv (Kommunikation möglicherweise eingeschränkt)"
        Write-Host "  Teams: Nicht aktiv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Teams: Check failed" -ForegroundColor Gray
}

# Outlook Check
try {
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Host "  Outlook: Läuft (PID: $($outlookProcess.Id))" -ForegroundColor Green
        $info += "Outlook: Aktiv"
    } else {
        $warnings += "Outlook: Nicht aktiv (Email möglicherweise eingeschränkt)"
        Write-Host "  Outlook: Nicht aktiv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Outlook: Check failed" -ForegroundColor Gray
}

# OneDrive Check
try {
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDriveProcess) {
        Write-Host "  OneDrive: Läuft (PID: $($oneDriveProcess.Id))" -ForegroundColor Green
        $info += "OneDrive: Aktiv"
    } else {
        $warnings += "OneDrive: Nicht aktiv (File Sync möglicherweise eingeschränkt)"
        Write-Host "  OneDrive: Nicht aktiv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  OneDrive: Check failed" -ForegroundColor Gray
}

# 4. Virtual Environment Check
Write-Host "`n=== VIRTUAL ENVIRONMENT ===" -ForegroundColor Yellow
try {
    $vmwareTools = Get-Service -Name "VMTools" -ErrorAction SilentlyContinue
    if ($vmwareTools) {
        Write-Host "  VMware Tools: $($vmwareTools.Status)" -ForegroundColor Green
        $info += "Virtual: VMware environment detected"
    } else {
        Write-Host "  VMware Tools: Nicht gefunden" -ForegroundColor Gray
    }
    
    # Check for virtual environment processes
    $vdiProcesses = @("vmtoolsd", "citrix", "teradici")
    $foundVDI = $false
    foreach ($process in $vdiProcesses) {
        $proc = Get-Process | Where-Object { $_.ProcessName -like "*$process*" }
        if ($proc) {
            $foundVDI = $true
            Write-Host "  Virtual Process: $($proc.ProcessName)" -ForegroundColor Green
        }
    }
    
    if (-not $foundVDI) {
        Write-Host "  Virtual Processes: Not found" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Virtual Check: Failed" -ForegroundColor Gray
}

# 5. Security Check (Virtual Environment)
Write-Host "`n=== SECURITY CHECK ===" -ForegroundColor Yellow

# UAC Check
try {
    $uacReg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -ErrorAction SilentlyContinue
    if ($uacReg.EnableLUA -eq 1) {
        Write-Host "  UAC: Aktiviert" -ForegroundColor Green
        $info += "UAC: Aktiviert (Standard)"
    } else {
        $issues += "CRITICAL: UAC disabled (NOT STANDARD!)"
        Write-Host "  UAC: Disabled (CRITICAL)" -ForegroundColor Red
    }
} catch {
    Write-Host "  UAC: Check failed" -ForegroundColor Gray
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
    Write-Host "  Domain: Check failed" -ForegroundColor Gray
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
        Write-Host "  Teams Cache: Check failed" -ForegroundColor Gray
    }
    
    # Java Check
    try {
        $javaVersion = java -version 2>&1 | Select-String "version"
        if ($javaVersion) {
            Write-Host "  Java: $($javaVersion.Line)" -ForegroundColor White
            $info += "Java: $($javaVersion.Line)"
  } else {
            $warnings += "Java: Not available (Enterprise apps may require Java)"
            Write-Host "  Java: Nicht verfügbar" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  Java: Check failed" -ForegroundColor Gray
    }
}

# 7. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     BASELINE SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

if ($issues.Count -gt 0) {
    Write-Host "`n🚨 CRITICAL ISSUES ($($issues.Count)):" -ForegroundColor Red
    foreach ($issue in $issues) {
        Write-Host "  • $issue" -ForegroundColor Red
    }
}

if ($warnings.Count -gt 0) {
    Write-Host "`n⚠️ WARNINGS ($($warnings.Count)):" -ForegroundColor Yellow
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
    Write-Host "  Status: CRITICAL - $($issues.Count) critical issues found!" -ForegroundColor Red
} elseif ($warnings.Count -gt 0) {
    Write-Host "  Status: WARNING - $($warnings.Count) warnings found" -ForegroundColor Yellow
} else {
    Write-Host "  Status: OK - Keine kritischen Probleme gefunden" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Baseline Check completed!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# Clean exit - prevent directory listing
if ($Verbose) {
    Read-Host "`nDrücken Sie Enter zum Beenden..."
} else {
    # Brief pause to ensure output is visible
    Start-Sleep -Milliseconds 500
}

# Ensure clean exit - change to a neutral system directory
Set-Location "C:\Windows\System32" | Out-Null
exit 0