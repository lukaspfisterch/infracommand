# Outlook Health Check
# Comprehensive diagnosis for Microsoft Outlook issues
# Supports Office 2016, 2019, 2021, and Office 365

param(
    [switch]$Verbose = $false,
    [switch]$Detailed = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     OUTLOOK HEALTH CHECK" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "     Mode: $(if($Detailed){'Detailed'}else{'Standard'})" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. Outlook Installation Check
Write-Host "`n=== OUTLOOK INSTALLATION ===" -ForegroundColor Yellow

try {
    # Office Version Detection
    $officeVersions = @()
    $officePaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook",
        "HKLM:\SOFTWARE\Microsoft\Office\14.0\Outlook"
    )
    
    $outlookVersion = "Nicht gefunden"
    $officeType = "Unknown"
    
    foreach ($path in $officePaths) {
        try {
            $reg = Get-ItemProperty $path -ErrorAction SilentlyContinue
            if ($reg) {
                $version = $path.Split('\')[4]
                $officeVersions += $version
                
                if ($version -eq "16.0") {
                    $officeType = "Office 365 / Office 2019/2021"
                    $outlookVersion = "Office 365/2019/2021"
                } elseif ($version -eq "15.0") {
                    $officeType = "Office 2013"
                    $outlookVersion = "Office 2013"
                } elseif ($version -eq "14.0") {
                    $officeType = "Office 2010"
                    $outlookVersion = "Office 2010"
                }
                break
            }
        } catch { }
    }
    
    Write-Host "  Outlook Version: $outlookVersion" -ForegroundColor Green
    Write-Host "  Office Type: $officeType" -ForegroundColor White
    
    # Outlook Executable Check
    $outlookExe = Get-Command "outlook.exe" -ErrorAction SilentlyContinue
    $outlookPath = $null
    
    if ($outlookExe) {
        $outlookPath = $outlookExe.Source
    } else {
        # Fallback: Try common Office paths
        $commonPaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office15\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office15\OUTLOOK.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office14\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office14\OUTLOOK.EXE"
        )
        
        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                $outlookPath = $path
                break
            }
        }
    }
    
    if ($outlookPath) {
        Write-Host "  Outlook Path: $outlookPath" -ForegroundColor White
        
        # Get File Version
        try {
            $fileVersion = (Get-ItemProperty $outlookPath).VersionInfo
            Write-Host "  Build: $($fileVersion.FileVersion)" -ForegroundColor White
            $info += "Outlook: $outlookVersion ($($fileVersion.FileVersion))"
        } catch {
            Write-Host "  Build: Version nicht lesbar" -ForegroundColor Yellow
            $info += "Outlook: $outlookVersion (Version nicht lesbar)"
        }
    } else {
        # Check if Outlook is running (fallback)
        $runningOutlook = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if ($runningOutlook) {
            Write-Host "  Outlook Path: Läuft (Pfad nicht ermittelbar)" -ForegroundColor Yellow
            $warnings += "Outlook: Läuft aber Pfad nicht ermittelbar"
            $info += "Outlook: $outlookVersion (läuft)"
        } else {
            $issues += "KRITISCH: Outlook.exe nicht gefunden!"
            Write-Host "  Outlook.exe: Nicht gefunden" -ForegroundColor Red
        }
    }
    
} catch {
    $warnings += "Outlook Installation Check fehlgeschlagen"
    Write-Host "  Installation Check: Fehlgeschlagen" -ForegroundColor Red
}

# 2. Outlook Process Check
Write-Host "`n=== OUTLOOK PROCESS ===" -ForegroundColor Yellow

try {
    $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcesses) {
        foreach ($process in $outlookProcesses) {
            $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
            Write-Host "  Outlook läuft (PID: $($process.Id), Memory: $memoryMB MB)" -ForegroundColor Green
            
            # Check if Outlook is responsive
            try {
                $process.Refresh()
                if ($process.Responding) {
                    Write-Host "    Status: Responsive" -ForegroundColor Green
                } else {
                    $warnings += "Outlook läuft aber reagiert nicht"
                    Write-Host "    Status: Nicht responsive" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "    Status: Unbekannt" -ForegroundColor Gray
            }
        }
        $info += "Outlook: Aktiv ($($outlookProcesses.Count) Prozess(e))"
    } else {
        Write-Host "  Outlook: Nicht aktiv" -ForegroundColor Yellow
        $warnings += "Outlook: Nicht aktiv"
    }
} catch {
    Write-Host "  Process Check: Fehlgeschlagen" -ForegroundColor Red
}

# 3. Outlook Profile Check
Write-Host "`n=== OUTLOOK PROFILES ===" -ForegroundColor Yellow

try {
    $outlookProfiles = @()
    
    # Check for Outlook profiles in registry
    $profileRegPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    if (-not (Test-Path $profileRegPath)) {
        $profileRegPath = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles"
    }
    if (-not (Test-Path $profileRegPath)) {
        $profileRegPath = "HKCU:\Software\Microsoft\Office\14.0\Outlook\Profiles"
    }
    
    if (Test-Path $profileRegPath) {
        $profiles = Get-ChildItem $profileRegPath -ErrorAction SilentlyContinue
        foreach ($profile in $profiles) {
            $outlookProfiles += $profile.PSChildName
        }
        
        if ($outlookProfiles.Count -gt 0) {
            Write-Host "  Profile gefunden: $($outlookProfiles.Count)" -ForegroundColor Green
            foreach ($profile in $outlookProfiles) {
                Write-Host "    - $profile" -ForegroundColor White
            }
            $info += "Outlook: $($outlookProfiles.Count) Profile konfiguriert"
        } else {
            $warnings += "Outlook: Keine Profile konfiguriert"
            Write-Host "  Profile: Keine gefunden" -ForegroundColor Yellow
        }
    } else {
        $warnings += "Outlook: Profile Registry nicht gefunden"
        Write-Host "  Profile Registry: Nicht gefunden" -ForegroundColor Yellow
    }
} catch {
    $warnings += "Outlook Profile Check fehlgeschlagen"
    Write-Host "  Profile Check: Fehlgeschlagen" -ForegroundColor Red
}

# 4. OST/PST Files Check (Optional - da global deaktiviert)
Write-Host "`n=== OST/PST FILES ===" -ForegroundColor Yellow

try {
    $ostPath = "$env:LOCALAPPDATA\Microsoft\Outlook"
    $pstPath = "$env:APPDATA\Microsoft\Outlook"
    
    $ostFiles = @()
    $pstFiles = @()
    
    if (Test-Path $ostPath) {
        $ostFiles = Get-ChildItem $ostPath -Filter "*.ost" -ErrorAction SilentlyContinue
    }
    if (Test-Path $pstPath) {
        $pstFiles = Get-ChildItem $pstPath -Filter "*.pst" -ErrorAction SilentlyContinue
    }
    
    if ($ostFiles.Count -gt 0) {
        Write-Host "  OST Files: $($ostFiles.Count) gefunden" -ForegroundColor Green
        foreach ($ost in $ostFiles) {
            $sizeMB = [math]::Round($ost.Length/1MB, 2)
            Write-Host "    - $($ost.Name) ($sizeMB MB)" -ForegroundColor White
        }
        $info += "OST: $($ostFiles.Count) Dateien gefunden"
    } else {
        Write-Host "  OST Files: Keine gefunden (OK - global deaktiviert)" -ForegroundColor Green
        $info += "OST: Deaktiviert (erwartet)"
    }
    
    if ($pstFiles.Count -gt 0) {
        Write-Host "  PST Files: $($pstFiles.Count) gefunden" -ForegroundColor Green
        foreach ($pst in $pstFiles) {
            $sizeMB = [math]::Round($pst.Length/1MB, 2)
            Write-Host "    - $($pst.Name) ($sizeMB MB)" -ForegroundColor White
        }
        $info += "PST: $($pstFiles.Count) Dateien gefunden"
    } else {
        Write-Host "  PST Files: Keine gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  OST/PST Check: Fehlgeschlagen" -ForegroundColor Red
}

# 5. Outlook Cache Check
Write-Host "`n=== OUTLOOK CACHE ===" -ForegroundColor Yellow

try {
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Outlook\RoamCache",
        "$env:LOCALAPPDATA\Microsoft\Outlook\WebView2",
        "$env:APPDATA\Microsoft\Outlook\RoamCache"
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
        
        if ($totalCacheMB -gt 500) {
            $warnings += "Outlook Cache groß: $totalCacheMB MB (Cleanup empfohlen)"
            Write-Host "  Status: GROSS - Cleanup empfohlen" -ForegroundColor Yellow
        } else {
            Write-Host "  Status: OK" -ForegroundColor Green
        }
        $info += "Outlook Cache: $totalCacheMB MB"
    } else {
        Write-Host "  Cache: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Cache Check: Fehlgeschlagen" -ForegroundColor Red
}

# 6. Office Services Check
Write-Host "`n=== OFFICE SERVICES ===" -ForegroundColor Yellow

try {
    $officeServices = @(
        "ClickToRunSvc",
        "OfficeSvc",
        "SstpSvc"
    )
    
    $runningServices = 0
    foreach ($serviceName in $officeServices) {
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
        $info += "Office Services: $runningServices von $($officeServices.Count) aktiv"
    } else {
        Write-Host "  Office Services: Keine gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Services Check: Fehlgeschlagen" -ForegroundColor Red
}

# 7. Network Connectivity (Detailed Mode)
if ($Detailed) {
    Write-Host "`n=== NETWORK CONNECTIVITY ===" -ForegroundColor Yellow
    
    try {
        # Office365 Endpoints
        $o365Endpoints = @(
            "outlook.office365.com",
            "outlook.office.com",
            "login.microsoftonline.com"
        )
        
        foreach ($endpoint in $o365Endpoints) {
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

# 8. Outlook Add-ins Check
Write-Host "`n=== OUTLOOK ADD-INS ===" -ForegroundColor Yellow

try {
    $addinsRegPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems"
    if (-not (Test-Path $addinsRegPath)) {
        $addinsRegPath = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Resiliency\DisabledItems"
    }
    
    if (Test-Path $addinsRegPath) {
        $disabledAddins = Get-ItemProperty $addinsRegPath -ErrorAction SilentlyContinue
        if ($disabledAddins) {
            $disabledCount = ($disabledAddins.PSObject.Properties | Where-Object { $_.Name -ne "PSPath" -and $_.Name -ne "PSParentPath" -and $_.Name -ne "PSChildName" -and $_.Name -ne "PSDrive" -and $_.Name -ne "PSProvider" }).Count
            if ($disabledCount -gt 0) {
                Write-Host "  Deaktivierte Add-ins: $disabledCount" -ForegroundColor Yellow
                $warnings += "Outlook: $disabledCount Add-ins deaktiviert"
            } else {
                Write-Host "  Deaktivierte Add-ins: Keine" -ForegroundColor Green
            }
        } else {
            Write-Host "  Deaktivierte Add-ins: Keine" -ForegroundColor Green
        }
    } else {
        Write-Host "  Add-ins Registry: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Add-ins Check: Fehlgeschlagen" -ForegroundColor Red
}

# 9. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     OUTLOOK HEALTH SUMMARY" -ForegroundColor Cyan
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
    Write-Host "  Status: OK - Outlook scheint gesund zu sein" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Outlook Health Check completed!" -ForegroundColor Green
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
