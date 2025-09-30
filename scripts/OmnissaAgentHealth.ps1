# Omnissa Agent Health Check
# Comprehensive diagnosis for Omnissa/VMware Horizon Agent issues
# Virtual Desktop Infrastructure (VDI) agent troubleshooting

param(
    [switch]$Verbose = $false,
    [switch]$Detailed = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     OMNISSA AGENT HEALTH CHECK" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "     Mode: $(if($Detailed){'Detailed'}else{'Standard'})" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. Omnissa Agent Installation Check
Write-Host "`n=== OMNISSA AGENT INSTALLATION ===" -ForegroundColor Yellow

try {
    # Check Omnissa Registry
    $omnissaRegPaths = @(
        "HKLM:\SOFTWARE\Omnissa",
        "HKLM:\SOFTWARE\WOW6432Node\Omnissa",
        "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM"
    )
    
    $omnissaInstalled = $false
    $omnissaVersion = "Nicht gefunden"
    $omnissaType = "Unknown"
    
    foreach ($regPath in $omnissaRegPaths) {
        if (Test-Path $regPath) {
            try {
                $reg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                if ($reg) {
                    $omnissaInstalled = $true
                    Write-Host "  Registry: $regPath gefunden" -ForegroundColor Green
                    
                    # Try to get version from multiple sources
                    if ($reg.Version) {
                        $omnissaVersion = $reg.Version
                    } elseif ($reg.DisplayVersion) {
                        $omnissaVersion = $reg.DisplayVersion
                    } elseif ($reg.ProductVersion) {
                        $omnissaVersion = $reg.ProductVersion
                    } elseif ($reg.VersionInfo) {
                        $omnissaVersion = $reg.VersionInfo
                    }
                    
                    # If still no version, try to get it from file
                    if ($omnissaVersion -eq "Nicht gefunden" -and $omnissaPath) {
                        try {
                            $exeFiles = Get-ChildItem $omnissaPath -Filter "*.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                            if ($exeFiles) {
                                $fileVersion = (Get-ItemProperty $exeFiles.FullName).VersionInfo
                                if ($fileVersion.FileVersion) {
                                    $omnissaVersion = $fileVersion.FileVersion
                                }
                            }
                        } catch { }
                    }
                    
                    if ($regPath -like "*Omnissa*") {
                        $omnissaType = "Omnissa Horizon"
                    } elseif ($regPath -like "*VMware*") {
                        $omnissaType = "VMware Horizon"
                    }
                    
                    Write-Host "  Version: $omnissaVersion" -ForegroundColor White
                    Write-Host "  Type: $omnissaType" -ForegroundColor White
                    $info += "Omnissa: $omnissaType ($omnissaVersion)"
                    break
                }
            } catch {
                Write-Host "  Registry: $regPath (nicht lesbar)" -ForegroundColor Yellow
            }
        }
    }
    
    if (-not $omnissaInstalled) {
        Write-Host "  Registry: Nicht gefunden" -ForegroundColor Red
        $issues += "KRITISCH: Omnissa Agent nicht installiert!"
    }
    
    # Check Installation Paths
    $omnissaPaths = @(
        "${env:ProgramFiles}\VMware\VDM",
        "${env:ProgramFiles(x86)}\VMware\VDM",
        "${env:ProgramFiles}\Omnissa",
        "${env:ProgramFiles(x86)}\Omnissa"
    )
    
    $omnissaPath = $null
    foreach ($path in $omnissaPaths) {
        if (Test-Path $path) {
            $omnissaPath = $path
            Write-Host "  Installation Path: $path" -ForegroundColor White
            $info += "Omnissa: Installation gefunden ($path)"
            break
        }
    }
    
    if (-not $omnissaPath -and $omnissaInstalled) {
        Write-Host "  Installation Path: Nicht gefunden (Registry vorhanden)" -ForegroundColor Yellow
        $warnings += "Omnissa: Registry vorhanden aber Installation nicht gefunden"
    }
    
} catch {
    $warnings += "Omnissa Installation Check fehlgeschlagen"
    Write-Host "  Installation Check: Fehlgeschlagen" -ForegroundColor Red
}

# 2. Omnissa Services Check
Write-Host "`n=== OMNISSA SERVICES ===" -ForegroundColor Yellow

try {
    $omnissaServices = @(
        "VMTools",
        "vmtoolsd",
        "vmware-view",
        "vmware-usbarbitrator64",
        "vmware-viewagent",
        "horizon-agent",
        "omnissa-agent"
    )
    
    $runningServices = 0
    $totalServices = 0
    $foundServices = @()
    
    foreach ($serviceName in $omnissaServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            $totalServices++
            $foundServices += $serviceName
            if ($service.Status -eq "Running") {
                Write-Host "  ${serviceName}: $($service.Status)" -ForegroundColor Green
                $runningServices++
            } else {
                Write-Host "  ${serviceName}: $($service.Status)" -ForegroundColor Yellow
                $warnings += "Omnissa Service: ${serviceName} nicht aktiv ($($service.Status))"
            }
        }
    }
    
    if ($totalServices -gt 0) {
        Write-Host "  Services: $runningServices von $totalServices aktiv" -ForegroundColor White
        Write-Host "  Gefundene Services: $($foundServices -join ', ')" -ForegroundColor White
        $info += "Omnissa Services: $runningServices von $totalServices aktiv"
        
        if ($runningServices -eq 0) {
            $issues += "KRITISCH: Keine Omnissa Services aktiv!"
        } elseif ($runningServices -lt $totalServices) {
            $warnings += "Omnissa: Nicht alle Services aktiv"
        }
    } else {
        Write-Host "  Services: Keine gefunden" -ForegroundColor Red
        $issues += "KRITISCH: Keine Omnissa Services gefunden!"
    }
    
} catch {
    Write-Host "  Services Check: Fehlgeschlagen" -ForegroundColor Red
}

# 3. Omnissa Processes Check
Write-Host "`n=== OMNISSA PROCESSES ===" -ForegroundColor Yellow

try {
    $omnissaProcesses = @(
        "vmtoolsd",
        "vmware-view",
        "vmware-usbarbitrator64",
        "vmware-viewagent",
        "horizon-agent",
        "omnissa-agent",
        "wswcagent",
        "wswc_daemon"
    )
    
    $runningProcesses = 0
    $foundProcesses = @()
    
    foreach ($processName in $omnissaProcesses) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($processes) {
            $runningProcesses += $processes.Count
            $foundProcesses += $processName
            foreach ($process in $processes) {
                $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
                Write-Host "  ${processName}: PID $($process.Id), Memory: $memoryMB MB" -ForegroundColor Green
            }
        }
    }
    
    if ($runningProcesses -gt 0) {
        Write-Host "  Processes: $runningProcesses aktiv" -ForegroundColor White
        Write-Host "  Gefundene Processes: $($foundProcesses -join ', ')" -ForegroundColor White
        $info += "Omnissa Processes: $runningProcesses aktiv"
    } else {
        Write-Host "  Processes: Keine aktiv" -ForegroundColor Yellow
        $warnings += "Omnissa: Keine Prozesse aktiv"
    }
    
} catch {
    Write-Host "  Processes Check: Fehlgeschlagen" -ForegroundColor Red
}

# 4. VMware Tools Check
Write-Host "`n=== VMWARE TOOLS ===" -ForegroundColor Yellow

try {
    # Check VMware Tools Service
    $vmwareTools = Get-Service -Name "VMTools" -ErrorAction SilentlyContinue
    if ($vmwareTools) {
        Write-Host "  VMware Tools Service: $($vmwareTools.Status)" -ForegroundColor Green
        $info += "VMware Tools: $($vmwareTools.Status)"
        
        if ($vmwareTools.Status -ne "Running") {
            $warnings += "VMware Tools: Nicht aktiv ($($vmwareTools.Status))"
        }
    } else {
        Write-Host "  VMware Tools Service: Nicht gefunden" -ForegroundColor Yellow
        $warnings += "VMware Tools: Service nicht gefunden"
    }
    
    # Check VMware Tools Process
    $vmwareToolsProcesses = Get-Process -Name "vmtoolsd" -ErrorAction SilentlyContinue
    if ($vmwareToolsProcesses) {
        $totalMemory = ($vmwareToolsProcesses | Measure-Object -Property WorkingSet -Sum).Sum
        $totalMemoryMB = [math]::Round($totalMemory/1MB, 2)
        Write-Host "  VMware Tools Process: $($vmwareToolsProcesses.Count) Prozess(e), Memory: $totalMemoryMB MB" -ForegroundColor Green
        $info += "VMware Tools Process: $($vmwareToolsProcesses.Count) aktiv ($totalMemoryMB MB)"
    } else {
        Write-Host "  VMware Tools Process: Nicht aktiv" -ForegroundColor Yellow
        $warnings += "VMware Tools: Process nicht aktiv"
    }
    
} catch {
    Write-Host "  VMware Tools Check: Fehlgeschlagen" -ForegroundColor Red
}

# 5. Horizon Agent Configuration
Write-Host "`n=== HORIZON AGENT CONFIGURATION ===" -ForegroundColor Yellow

try {
    $configPaths = @(
        "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent\Configuration",
        "HKLM:\SOFTWARE\Omnissa\Agent\Configuration"
    )
    
    $configFound = $false
    foreach ($configPath in $configPaths) {
        if (Test-Path $configPath) {
            Write-Host "  Configuration: $configPath gefunden" -ForegroundColor Green
            
            # Check key configuration values
            $configKeys = @(
                "ServerURL",
                "ServerName",
                "DomainName",
                "EnableUSB",
                "EnableClientDriveRedirection",
                "EnablePrinterRedirection"
            )
            
            foreach ($key in $configKeys) {
                try {
                    $value = Get-ItemProperty $configPath -Name $key -ErrorAction SilentlyContinue
                    if ($value) {
                        Write-Host "    ${key}: $($value.$key)" -ForegroundColor White
                    }
                } catch { }
            }
            
            $configFound = $true
            $info += "Omnissa: Configuration gefunden"
            break
        }
    }
    
    if (-not $configFound) {
        # Check if this is a VDI environment where config might be different
        $vdiProcesses = Get-Process | Where-Object { $_.ProcessName -like "*vmtoolsd*" -or $_.ProcessName -like "*vmware*" }
        if ($vdiProcesses) {
            Write-Host "  Configuration: Nicht gefunden (VDI-Umgebung - normal)" -ForegroundColor Yellow
            $warnings += "Omnissa: Configuration nicht gefunden (VDI-Umgebung - normal)"
        } else {
            Write-Host "  Configuration: Nicht gefunden" -ForegroundColor Yellow
            $warnings += "Omnissa: Configuration nicht gefunden"
        }
    }
    
} catch {
    Write-Host "  Configuration Check: Fehlgeschlagen" -ForegroundColor Red
}

# 6. Horizon Logs Check
Write-Host "`n=== HORIZON LOGS ===" -ForegroundColor Yellow

try {
    $logsPaths = @(
        "C:\ProgramData\VMware\VDM\logs",
        "C:\ProgramData\Omnissa\logs",
        "C:\Users\$env:USERNAME\AppData\Local\VMware\VDM\logs"
    )
    
    $logsFound = $false
    foreach ($logsPath in $logsPaths) {
        if (Test-Path $logsPath) {
            $logFiles = Get-ChildItem $logsPath -Filter "*.log" -ErrorAction SilentlyContinue
            if ($logFiles) {
                $logCount = $logFiles.Count
                $totalLogSize = ($logFiles | Measure-Object -Property Length -Sum).Sum
                $logSizeMB = [math]::Round($totalLogSize/1MB, 2)
                
                Write-Host "  Log Path: $logsPath" -ForegroundColor Green
                Write-Host "  Log Files: $logCount gefunden" -ForegroundColor White
                Write-Host "  Log Size: $logSizeMB MB" -ForegroundColor White
                
                # Check for recent errors
                $recentLogs = $logFiles | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-1) }
                if ($recentLogs) {
                    Write-Host "  Recent Logs: $($recentLogs.Count) in letzten 24h" -ForegroundColor White
                    
                    # Check for error patterns in recent logs
                    $errorCount = 0
                    foreach ($log in $recentLogs) {
                        try {
                            $content = Get-Content $log.FullName -ErrorAction SilentlyContinue | Select-String -Pattern "error|fail|exception|disconnect" -CaseSensitive:$false
                            if ($content) {
                                $errorCount += $content.Count
                            }
                        } catch { }
                    }
                    
                    if ($errorCount -gt 0) {
                        Write-Host "  Errors in Logs: $errorCount gefunden" -ForegroundColor Yellow
                        $warnings += "Omnissa: $errorCount Fehler in Logs gefunden"
                    } else {
                        Write-Host "  Errors in Logs: Keine gefunden" -ForegroundColor Green
                    }
                }
                
                $logsFound = $true
                $info += "Omnissa Logs: $logCount Dateien ($logSizeMB MB)"
                break
            }
        }
    }
    
    if (-not $logsFound) {
        Write-Host "  Logs: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Logs Check: Fehlgeschlagen" -ForegroundColor Red
}

# 7. VDI Environment Check
Write-Host "`n=== VDI ENVIRONMENT ===" -ForegroundColor Yellow

try {
    # Check for VDI-specific processes
    $vdiProcesses = @(
        "vmtoolsd",
        "vmware-view",
        "citrix",
        "teradici",
        "rdp"
    )
    
    $vdiFound = $false
    foreach ($process in $vdiProcesses) {
        $proc = Get-Process | Where-Object { $_.ProcessName -like "*$process*" }
        if ($proc) {
            $vdiFound = $true
            Write-Host "  VDI Process: $($proc.ProcessName)" -ForegroundColor Green
        }
    }
    
    if ($vdiFound) {
        $info += "VDI: VDI-Umgebung erkannt"
    } else {
        Write-Host "  VDI Processes: Nicht gefunden" -ForegroundColor Yellow
        $warnings += "VDI: VDI-Prozesse nicht gefunden"
    }
    
    # Check for VDI-specific registry entries
    $vdiRegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server",
        "HKLM:\SOFTWARE\Citrix",
        "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM"
    )
    
    $vdiRegFound = $false
    foreach ($regPath in $vdiRegPaths) {
        if (Test-Path $regPath) {
            $vdiRegFound = $true
            Write-Host "  VDI Registry: $regPath gefunden" -ForegroundColor Green
        }
    }
    
    if ($vdiRegFound) {
        $info += "VDI: Registry-Einträge gefunden"
    }
    
} catch {
    Write-Host "  VDI Check: Fehlgeschlagen" -ForegroundColor Red
}

# 8. Performance Check (Detailed Mode)
if ($Detailed) {
    Write-Host "`n=== PERFORMANCE CHECK ===" -ForegroundColor Yellow
    
    try {
        # Check Omnissa process performance
        $omnissaProcesses = Get-Process | Where-Object { $_.ProcessName -like "*vmware*" -or $_.ProcessName -like "*horizon*" -or $_.ProcessName -like "*omnissa*" }
        if ($omnissaProcesses) {
            $totalMemory = ($omnissaProcesses | Measure-Object -Property WorkingSet -Sum).Sum
            $totalMemoryMB = [math]::Round($totalMemory/1MB, 2)
            
            Write-Host "  Omnissa Memory: $totalMemoryMB MB" -ForegroundColor White
            
            if ($totalMemoryMB -gt 200) {
                $warnings += "Omnissa: Hoher Memory-Verbrauch ($totalMemoryMB MB)"
                Write-Host "  Status: HOCH - Performance-Überwachung empfohlen" -ForegroundColor Yellow
            } else {
                Write-Host "  Status: OK" -ForegroundColor Green
            }
            
            $info += "Omnissa Memory: $totalMemoryMB MB"
        } else {
            Write-Host "  Omnissa Processes: Nicht aktiv" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "  Performance Check: Fehlgeschlagen" -ForegroundColor Red
    }
}

# 9. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     OMNISSA AGENT HEALTH SUMMARY" -ForegroundColor Cyan
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
    Write-Host "  Status: OK - Omnissa Agent scheint gesund zu sein" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Omnissa Agent Health Check completed!" -ForegroundColor Green
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
