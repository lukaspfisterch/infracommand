# FSLogix Health Check
# Comprehensive diagnosis for FSLogix profile management issues
# Virtual Desktop Infrastructure (VDI) profile container management

param(
    [switch]$Verbose = $false,
    [switch]$Detailed = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     FSLOGIX HEALTH CHECK" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "     Mode: $(if($Detailed){'Detailed'}else{'Standard'})" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. FSLogix Installation Check
Write-Host "`n=== FSLOGIX INSTALLATION ===" -ForegroundColor Yellow

try {
    # Check FSLogix Registry
    $fslogixRegPath = "HKLM:\SOFTWARE\FSLogix"
    $fslogixInstalled = $false
    $fslogixVersion = "Nicht gefunden"
    
    if (Test-Path $fslogixRegPath) {
        try {
            $fslogixReg = Get-ItemProperty $fslogixRegPath -ErrorAction SilentlyContinue
            if ($fslogixReg) {
                $fslogixInstalled = $true
                Write-Host "  FSLogix Registry: Gefunden" -ForegroundColor Green
                
                # Try to get version from registry
                if ($fslogixReg.Version) {
                    $fslogixVersion = $fslogixReg.Version
                    Write-Host "  Version: $fslogixVersion" -ForegroundColor White
                }
                
                $info += "FSLogix: Registry gefunden"
            }
        } catch {
            Write-Host "  Registry: Gefunden aber nicht lesbar" -ForegroundColor Yellow
            $warnings += "FSLogix: Registry nicht lesbar"
        }
    } else {
        Write-Host "  FSLogix Registry: Nicht gefunden" -ForegroundColor Red
        $issues += "KRITISCH: FSLogix nicht installiert!"
    }
    
    # Check FSLogix Installation Path
    $fslogixPaths = @(
        "${env:ProgramFiles}\FSLogix",
        "${env:ProgramFiles(x86)}\FSLogix",
        "C:\Program Files\FSLogix",
        "C:\Program Files (x86)\FSLogix"
    )
    
    $fslogixPath = $null
    foreach ($path in $fslogixPaths) {
        if (Test-Path $path) {
            $fslogixPath = $path
            Write-Host "  Installation Path: $path" -ForegroundColor White
            $info += "FSLogix: Installation gefunden ($path)"
            break
        }
    }
    
    if (-not $fslogixPath -and $fslogixInstalled) {
        Write-Host "  Installation Path: Nicht gefunden (Registry vorhanden)" -ForegroundColor Yellow
        $warnings += "FSLogix: Registry vorhanden aber Installation nicht gefunden"
    }
    
} catch {
    $warnings += "FSLogix Installation Check fehlgeschlagen"
    Write-Host "  Installation Check: Fehlgeschlagen" -ForegroundColor Red
}

# 2. FSLogix Services Check
Write-Host "`n=== FSLOGIX SERVICES ===" -ForegroundColor Yellow

try {
    $fslogixServices = @(
        "frxsvc",
        "frxdrv",
        "frxdrvvt"
    )
    
    $runningServices = 0
    $totalServices = 0
    
    foreach ($serviceName in $fslogixServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            $totalServices++
            if ($service.Status -eq "Running") {
                Write-Host "  ${serviceName}: $($service.Status)" -ForegroundColor Green
                $runningServices++
            } else {
                Write-Host "  ${serviceName}: $($service.Status)" -ForegroundColor Yellow
                $warnings += "FSLogix Service: ${serviceName} nicht aktiv ($($service.Status))"
            }
        }
    }
    
    if ($totalServices -gt 0) {
        Write-Host "  Services: $runningServices von $totalServices aktiv" -ForegroundColor White
        $info += "FSLogix Services: $runningServices von $totalServices aktiv"
        
        if ($runningServices -eq 0) {
            $issues += "KRITISCH: Keine FSLogix Services aktiv!"
        } elseif ($runningServices -lt $totalServices) {
            $warnings += "FSLogix: Nicht alle Services aktiv"
        }
    } else {
        Write-Host "  Services: Keine gefunden" -ForegroundColor Red
        $issues += "KRITISCH: Keine FSLogix Services gefunden!"
    }
    
} catch {
    Write-Host "  Services Check: Fehlgeschlagen" -ForegroundColor Red
}

# 3. FSLogix Drivers Check
Write-Host "`n=== FSLOGIX DRIVERS ===" -ForegroundColor Yellow

try {
    $fslogixDrivers = @(
        "frxdrv",
        "frxdrvvt"
    )
    
    $loadedDrivers = 0
    foreach ($driverName in $fslogixDrivers) {
        $driver = Get-WmiObject -Class Win32_SystemDriver | Where-Object { $_.Name -eq $driverName }
        if ($driver) {
            if ($driver.State -eq "Running") {
                Write-Host "  ${driverName}: $($driver.State)" -ForegroundColor Green
                $loadedDrivers++
            } else {
                Write-Host "  ${driverName}: $($driver.State)" -ForegroundColor Yellow
                $warnings += "FSLogix Driver: ${driverName} nicht aktiv ($($driver.State))"
            }
        } else {
            Write-Host "  ${driverName}: Nicht gefunden" -ForegroundColor Red
            $warnings += "FSLogix Driver: ${driverName} nicht gefunden"
        }
    }
    
    if ($loadedDrivers -gt 0) {
        $info += "FSLogix Drivers: $loadedDrivers von $($fslogixDrivers.Count) aktiv"
    } else {
        $warnings += "FSLogix: Keine Drivers aktiv"
    }
    
} catch {
    Write-Host "  Drivers Check: Fehlgeschlagen" -ForegroundColor Red
}

# 4. FSLogix Configuration Check
Write-Host "`n=== FSLOGIX CONFIGURATION ===" -ForegroundColor Yellow

try {
    $configRegPath = "HKLM:\SOFTWARE\FSLogix\Profiles"
    if (Test-Path $configRegPath) {
        Write-Host "  Configuration: Gefunden" -ForegroundColor Green
        
        # Check key configuration values
        $configKeys = @(
            "Enabled",
            "VHDLocations",
            "VolumeType",
            "SizeInMBs",
            "FlipFlopProfileDirectoryName",
            "DeleteLocalProfileWhenVHDShouldApply",
            "PreventLoginWithFailure",
            "PreventLoginWithTempProfile"
        )
        
        foreach ($key in $configKeys) {
            try {
                $value = Get-ItemProperty $configRegPath -Name $key -ErrorAction SilentlyContinue
                if ($value) {
                    $displayValue = $value.$key
                    if ($key -eq "VHDLocations" -and $displayValue) {
                        $displayValue = $displayValue -join "; "
                    }
                    Write-Host "    ${key}: $displayValue" -ForegroundColor White
                }
            } catch {
                Write-Host "    ${key}: Nicht konfiguriert" -ForegroundColor Gray
            }
        }
        
        $info += "FSLogix: Configuration gefunden"
    } else {
        Write-Host "  Configuration: Nicht gefunden" -ForegroundColor Yellow
        $warnings += "FSLogix: Configuration nicht gefunden"
    }
    
} catch {
    Write-Host "  Configuration Check: Fehlgeschlagen" -ForegroundColor Red
}

# 5. FSLogix Logs Check
Write-Host "`n=== FSLOGIX LOGS ===" -ForegroundColor Yellow

try {
    $logsPath = "C:\ProgramData\FSLogix\Logs"
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
                
                # Check for error patterns in recent logs
                $errorCount = 0
                foreach ($log in $recentLogs) {
                    try {
                        $content = Get-Content $log.FullName -ErrorAction SilentlyContinue | Select-String -Pattern "error|fail|exception" -CaseSensitive:$false
                        if ($content) {
                            $errorCount += $content.Count
                        }
                    } catch { }
                }
                
                if ($errorCount -gt 0) {
                    Write-Host "  Errors in Logs: $errorCount gefunden" -ForegroundColor Yellow
                    $warnings += "FSLogix: $errorCount Fehler in Logs gefunden"
                } else {
                    Write-Host "  Errors in Logs: Keine gefunden" -ForegroundColor Green
                }
            }
            
            $info += "FSLogix Logs: $logCount Dateien ($logSizeMB MB)"
        } else {
            Write-Host "  Log Files: Keine gefunden" -ForegroundColor Gray
        }
    } else {
        Write-Host "  Logs Path: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Logs Check: Fehlgeschlagen" -ForegroundColor Red
}

# 6. VHD/VHDX Files Check
Write-Host "`n=== VHD/VHDX FILES ===" -ForegroundColor Yellow

try {
    # Check for VHD locations in registry
    $vhdLocations = @()
    $configRegPath = "HKLM:\SOFTWARE\FSLogix\Profiles"
    
    if (Test-Path $configRegPath) {
        try {
            $vhdLocationsReg = Get-ItemProperty $configRegPath -Name "VHDLocations" -ErrorAction SilentlyContinue
            if ($vhdLocationsReg.VHDLocations) {
                $vhdLocations = $vhdLocationsReg.VHDLocations
            }
        } catch { }
    }
    
    # Default VHD locations if not configured
    if ($vhdLocations.Count -eq 0) {
        $vhdLocations = @(
            "\\fs-group\usz_daten\FSLogixProfiles",
            "\\fs-group\usz_daten\FSLogix",
            "C:\FSLogixProfiles"
        )
    }
    
    $vhdFound = $false
    foreach ($location in $vhdLocations) {
        if (Test-Path $location) {
            Write-Host "  VHD Location: $location" -ForegroundColor Green
            
            # Check for VHD/VHDX files
            $vhdFiles = Get-ChildItem $location -Filter "*.vhd*" -ErrorAction SilentlyContinue
            if ($vhdFiles) {
                $vhdCount = $vhdFiles.Count
                $totalVhdSize = ($vhdFiles | Measure-Object -Property Length -Sum).Sum
                $vhdSizeGB = [math]::Round($totalVhdSize/1GB, 2)
                
                Write-Host "    VHD Files: $vhdCount gefunden" -ForegroundColor White
                Write-Host "    VHD Size: $vhdSizeGB GB" -ForegroundColor White
                
                $vhdFound = $true
                $info += "FSLogix VHD: $vhdCount Dateien ($vhdSizeGB GB) in $location"
            } else {
                Write-Host "    VHD Files: Keine gefunden" -ForegroundColor Gray
            }
        } else {
            Write-Host "  VHD Location: $location (Nicht erreichbar)" -ForegroundColor Yellow
        }
    }
    
    if (-not $vhdFound) {
        $warnings += "FSLogix: Keine VHD-Dateien gefunden"
    }
    
} catch {
    Write-Host "  VHD Check: Fehlgeschlagen" -ForegroundColor Red
}

# 7. Profile Status Check
Write-Host "`n=== PROFILE STATUS ===" -ForegroundColor Yellow

try {
    # Check current user profile status
    $currentUser = $env:USERNAME
    $profileRegPath = "HKCU:\SOFTWARE\FSLogix"
    
    if (Test-Path $profileRegPath) {
        Write-Host "  User Profile: FSLogix aktiv" -ForegroundColor Green
        
        # Check profile-specific settings
        $profileSettings = @(
            "ProfileType",
            "VHDLocation",
            "ProfileSize",
            "LastLogin"
        )
        
        foreach ($setting in $profileSettings) {
            try {
                $value = Get-ItemProperty $profileRegPath -Name $setting -ErrorAction SilentlyContinue
                if ($value) {
                    Write-Host "    ${setting}: $($value.$setting)" -ForegroundColor White
                }
            } catch { }
        }
        
        $info += "FSLogix: User Profile aktiv"
    } else {
        Write-Host "  User Profile: FSLogix nicht aktiv" -ForegroundColor Yellow
        $warnings += "FSLogix: User Profile nicht aktiv"
    }
    
} catch {
    Write-Host "  Profile Check: Fehlgeschlagen" -ForegroundColor Red
}

# 8. Performance Check (Detailed Mode)
if ($Detailed) {
    Write-Host "`n=== PERFORMANCE CHECK ===" -ForegroundColor Yellow
    
    try {
        # Check FSLogix process performance
        $fslogixProcesses = Get-Process | Where-Object { $_.ProcessName -like "*fslogix*" -or $_.ProcessName -like "*frx*" }
        if ($fslogixProcesses) {
            $totalMemory = ($fslogixProcesses | Measure-Object -Property WorkingSet -Sum).Sum
            $totalMemoryMB = [math]::Round($totalMemory/1MB, 2)
            
            Write-Host "  FSLogix Memory: $totalMemoryMB MB" -ForegroundColor White
            
            if ($totalMemoryMB -gt 100) {
                $warnings += "FSLogix: Hoher Memory-Verbrauch ($totalMemoryMB MB)"
                Write-Host "  Status: HOCH - Performance-Überwachung empfohlen" -ForegroundColor Yellow
            } else {
                Write-Host "  Status: OK" -ForegroundColor Green
            }
            
            $info += "FSLogix Memory: $totalMemoryMB MB"
        } else {
            Write-Host "  FSLogix Processes: Nicht aktiv" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "  Performance Check: Fehlgeschlagen" -ForegroundColor Red
    }
}

# 9. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     FSLOGIX HEALTH SUMMARY" -ForegroundColor Cyan
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
    Write-Host "  Status: OK - FSLogix scheint gesund zu sein" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "FSLogix Health Check completed!" -ForegroundColor Green
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
