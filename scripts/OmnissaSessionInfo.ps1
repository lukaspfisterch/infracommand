# Omnissa Session Info Health Check
# Comprehensive diagnosis for session/broker information
# Virtual Desktop Infrastructure (VDI) session troubleshooting

param(
    [switch]$Verbose = $false,
    [switch]$Detailed = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     OMNISSA SESSION INFO CHECK" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "     Mode: $(if($Detailed){'Detailed'}else{'Standard'})" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$issues = @()
$warnings = @()
$info = @()

# 1. Session Information Check
Write-Host "`n=== SESSION INFORMATION ===" -ForegroundColor Yellow

try {
    # Current Session Info
    $currentSession = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $sessionId = $currentSession.User.Value
    $sessionName = $currentSession.Name
    
    Write-Host "  Current Session: $sessionName" -ForegroundColor White
    Write-Host "  Session ID: $sessionId" -ForegroundColor White
    
    # Check if running in VDI session
    $isVDISession = $false
    $sessionType = "Unknown"
    
    # Check for RDP session
    $rdpSession = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction SilentlyContinue
    if ($rdpSession -and $rdpSession.fDenyTSConnections -eq 0) {
        $isVDISession = $true
        $sessionType = "RDP Session"
        Write-Host "  Session Type: RDP Session" -ForegroundColor Green
    }
    
    # Check for Citrix session
    $citrixSession = Get-Process -Name "*citrix*" -ErrorAction SilentlyContinue
    if ($citrixSession) {
        $isVDISession = $true
        $sessionType = "Citrix Session"
        Write-Host "  Session Type: Citrix Session" -ForegroundColor Green
    }
    
    # Check for VMware Horizon session
    $horizonSession = Get-Process -Name "*vmware*" -ErrorAction SilentlyContinue
    if ($horizonSession) {
        $isVDISession = $true
        $sessionType = "VMware Horizon Session"
        Write-Host "  Session Type: VMware Horizon Session" -ForegroundColor Green
    }
    
    # Check for vmtoolsd processes (VDI indicator)
    $vmtoolsProcesses = Get-Process -Name "vmtoolsd" -ErrorAction SilentlyContinue
    if ($vmtoolsProcesses -and -not $isVDISession) {
        $isVDISession = $true
        $sessionType = "VDI Session (VMware Tools)"
        Write-Host "  Session Type: VDI Session (VMware Tools)" -ForegroundColor Green
    }
    
    if (-not $isVDISession) {
        Write-Host "  Session Type: Local Session" -ForegroundColor Yellow
        $warnings += "Session: Lokale Session (kein VDI)"
    }
    
    $info += "Session: $sessionType"
    
} catch {
    $warnings += "Session Information Check fehlgeschlagen"
    Write-Host "  Session Check: Fehlgeschlagen" -ForegroundColor Red
}

# 2. Broker Connection Check
Write-Host "`n=== BROKER CONNECTION ===" -ForegroundColor Yellow

try {
    # Check for Horizon Broker processes
    $brokerProcesses = @(
        "vmware-view",
        "vmware-usbarbitrator64",
        "wswcagent",
        "wswc_daemon"
    )
    
    $activeBrokers = 0
    $foundBrokers = @()
    
    foreach ($processName in $brokerProcesses) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($processes) {
            $activeBrokers += $processes.Count
            $foundBrokers += $processName
            foreach ($process in $processes) {
                $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
                Write-Host "  ${processName}: PID $($process.Id), Memory: $memoryMB MB" -ForegroundColor Green
            }
        }
    }
    
    if ($activeBrokers -gt 0) {
        Write-Host "  Broker Processes: $activeBrokers aktiv" -ForegroundColor White
        Write-Host "  Gefundene Brokers: $($foundBrokers -join ', ')" -ForegroundColor White
        $info += "Broker: $activeBrokers Prozesse aktiv"
    } else {
        Write-Host "  Broker Processes: Keine aktiv" -ForegroundColor Yellow
        $warnings += "Broker: Keine Prozesse aktiv"
    }
    
    # Check for broker configuration
    $brokerConfigPaths = @(
        "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent\Configuration",
        "HKLM:\SOFTWARE\Omnissa\Agent\Configuration"
    )
    
    $brokerConfigFound = $false
    foreach ($configPath in $brokerConfigPaths) {
        if (Test-Path $configPath) {
            try {
                $serverURL = Get-ItemProperty $configPath -Name "ServerURL" -ErrorAction SilentlyContinue
                $serverName = Get-ItemProperty $configPath -Name "ServerName" -ErrorAction SilentlyContinue
                
                if ($serverURL) {
                    Write-Host "  Broker Server URL: $($serverURL.ServerURL)" -ForegroundColor White
                    $info += "Broker: Server URL konfiguriert"
                }
                if ($serverName) {
                    Write-Host "  Broker Server Name: $($serverName.ServerName)" -ForegroundColor White
                    $info += "Broker: Server Name konfiguriert"
                }
                
                $brokerConfigFound = $true
            } catch { }
        }
    }
    
    if (-not $brokerConfigFound) {
        # Check if this is a VDI environment where config might be different
        $vdiProcesses = Get-Process | Where-Object { $_.ProcessName -like "*vmware*" -or $_.ProcessName -like "*vmtoolsd*" }
        if ($vdiProcesses) {
            Write-Host "  Broker Configuration: Nicht gefunden (VDI-Umgebung - normal)" -ForegroundColor Yellow
            $warnings += "Broker: Configuration nicht gefunden (VDI-Umgebung - normal)"
        } else {
            Write-Host "  Broker Configuration: Nicht gefunden" -ForegroundColor Yellow
            $warnings += "Broker: Configuration nicht gefunden"
        }
    }
    
} catch {
    Write-Host "  Broker Check: Fehlgeschlagen" -ForegroundColor Red
}

# 3. Connection Server Check
Write-Host "`n=== CONNECTION SERVER ===" -ForegroundColor Yellow

try {
    # Check for connection server processes
    $connectionProcesses = Get-Process | Where-Object { 
        $_.ProcessName -like "*horizon*" -or 
        $_.ProcessName -like "*view*" -or 
        $_.ProcessName -like "*broker*" 
    }
    
    if ($connectionProcesses) {
        $totalMemory = ($connectionProcesses | Measure-Object -Property WorkingSet -Sum).Sum
        $totalMemoryMB = [math]::Round($totalMemory/1MB, 2)
        
        Write-Host "  Connection Processes: $($connectionProcesses.Count) gefunden" -ForegroundColor Green
        Write-Host "  Total Memory: $totalMemoryMB MB" -ForegroundColor White
        
        foreach ($process in $connectionProcesses) {
            $memoryMB = [math]::Round($process.WorkingSet/1MB, 2)
            Write-Host "    ${process.ProcessName}: PID $($process.Id), Memory: $memoryMB MB" -ForegroundColor White
        }
        
        $info += "Connection Server: $($connectionProcesses.Count) Prozesse ($totalMemoryMB MB)"
    } else {
        Write-Host "  Connection Processes: Keine gefunden" -ForegroundColor Yellow
        $warnings += "Connection Server: Keine Prozesse gefunden"
    }
    
    # Check for connection server registry entries
    $connectionRegPaths = @(
        "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent\Configuration",
        "HKLM:\SOFTWARE\Omnissa\Agent\Configuration"
    )
    
    $connectionRegFound = $false
    foreach ($regPath in $connectionRegPaths) {
        if (Test-Path $regPath) {
            try {
                $reg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                if ($reg) {
                    Write-Host "  Connection Registry: $regPath gefunden" -ForegroundColor Green
                    $connectionRegFound = $true
                    
                    # Display key connection settings
                    $connectionKeys = @(
                        "ServerURL",
                        "ServerName",
                        "DomainName",
                        "EnableUSB",
                        "EnableClientDriveRedirection"
                    )
                    
                    foreach ($key in $connectionKeys) {
                        try {
                            $value = Get-ItemProperty $regPath -Name $key -ErrorAction SilentlyContinue
                            if ($value) {
                                Write-Host "    ${key}: $($value.$key)" -ForegroundColor White
                            }
                        } catch { }
                    }
                }
            } catch { }
        }
    }
    
    if (-not $connectionRegFound) {
        # Check if this is a VDI environment where config might be different
        $vdiProcesses = Get-Process | Where-Object { $_.ProcessName -like "*vmware*" -or $_.ProcessName -like "*vmtoolsd*" }
        if ($vdiProcesses) {
            Write-Host "  Connection Registry: Nicht gefunden (VDI-Umgebung - normal)" -ForegroundColor Yellow
            $warnings += "Connection Server: Registry nicht gefunden (VDI-Umgebung - normal)"
        } else {
            Write-Host "  Connection Registry: Nicht gefunden" -ForegroundColor Yellow
            $warnings += "Connection Server: Registry nicht gefunden"
        }
    }
    
} catch {
    Write-Host "  Connection Server Check: Fehlgeschlagen" -ForegroundColor Red
}

# 4. Session State Check
Write-Host "`n=== SESSION STATE ===" -ForegroundColor Yellow

try {
    # Check current session state
    $sessionState = "Unknown"
    
    # Check for active user session
    $activeSessions = quser 2>$null
    if ($activeSessions) {
        $sessionState = "Active"
        Write-Host "  Session State: Active" -ForegroundColor Green
        
        # Parse session information
        $sessionLines = $activeSessions -split "`n" | Where-Object { $_.Trim() -ne "" }
        foreach ($line in $sessionLines) {
            if ($line -match "(\w+)\s+(\d+)\s+(\w+)\s+(.+)") {
                $user = $matches[1]
                $sessionId = $matches[2]
                $state = $matches[3]
                $info = $matches[4]
                
                Write-Host "    User: $user, Session: $sessionId, State: $state" -ForegroundColor White
            }
        }
        
        $info += "Session State: Active"
    } else {
        Write-Host "  Session State: Nicht verfügbar" -ForegroundColor Yellow
        $warnings += "Session State: Nicht verfügbar"
    }
    
    # Check for session timeout settings
    $timeoutRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"
    if (Test-Path $timeoutRegPath) {
        try {
            $idleTimeout = Get-ItemProperty $timeoutRegPath -Name "MaxIdleTime" -ErrorAction SilentlyContinue
            if ($idleTimeout) {
                $timeoutMinutes = [math]::Round($idleTimeout.MaxIdleTime / 60000, 0)
                Write-Host "  Idle Timeout: $timeoutMinutes Minuten" -ForegroundColor White
                $info += "Session: Idle Timeout $timeoutMinutes Min"
            }
        } catch { }
    }
    
} catch {
    Write-Host "  Session State Check: Fehlgeschlagen" -ForegroundColor Red
}

# 5. Network Connectivity Check
Write-Host "`n=== NETWORK CONNECTIVITY ===" -ForegroundColor Yellow

try {
    # Check for common VDI/VDI-related network connections
    $networkConnections = Get-NetTCPConnection -ErrorAction SilentlyContinue | Where-Object {
        $_.State -eq "Established" -and (
            $_.RemotePort -eq 443 -or 
            $_.RemotePort -eq 80 -or 
            $_.RemotePort -eq 3389 -or
            $_.RemotePort -eq 1494
        )
    }
    
    if ($networkConnections) {
        $connectionCount = $networkConnections.Count
        Write-Host "  Active Connections: $connectionCount" -ForegroundColor Green
        
        # Group by port
        $portGroups = $networkConnections | Group-Object RemotePort
        foreach ($group in $portGroups) {
            $port = $group.Name
            $count = $group.Count
            Write-Host "    Port ${port}: $count Verbindungen" -ForegroundColor White
        }
        
        $info += "Network: $connectionCount aktive Verbindungen"
    } else {
        Write-Host "  Active Connections: Keine gefunden" -ForegroundColor Yellow
        $warnings += "Network: Keine VDI-Verbindungen gefunden"
    }
    
    # Check for specific VDI endpoints (Detailed Mode)
    if ($Detailed) {
        $vdiEndpoints = @(
            "horizon.example.com",
            "broker.example.com",
            "connection.example.com"
        )
        
        foreach ($endpoint in $vdiEndpoints) {
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
    }
    
} catch {
    Write-Host "  Network Check: Fehlgeschlagen" -ForegroundColor Red
}

# 6. Session Logs Check
Write-Host "`n=== SESSION LOGS ===" -ForegroundColor Yellow

try {
    $logsPaths = @(
        "C:\ProgramData\VMware\VDM\logs",
        "C:\ProgramData\Omnissa\logs",
        "C:\Users\$env:USERNAME\AppData\Local\VMware\VDM\logs",
        "C:\Windows\Logs\TerminalServices"
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
                
                # Check for recent session logs
                $recentLogs = $logFiles | Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-1) }
                if ($recentLogs) {
                    Write-Host "  Recent Logs: $($recentLogs.Count) in letzten 24h" -ForegroundColor White
                    
                    # Check for session-related errors
                    $sessionErrorCount = 0
                    foreach ($log in $recentLogs) {
                        try {
                            $content = Get-Content $log.FullName -ErrorAction SilentlyContinue | Select-String -Pattern "session|disconnect|timeout|error" -CaseSensitive:$false
                            if ($content) {
                                $sessionErrorCount += $content.Count
                            }
                        } catch { }
                    }
                    
                    if ($sessionErrorCount -gt 0) {
                        Write-Host "  Session Errors: $sessionErrorCount gefunden" -ForegroundColor Yellow
                        $warnings += "Session Logs: $sessionErrorCount Fehler gefunden"
                    } else {
                        Write-Host "  Session Errors: Keine gefunden" -ForegroundColor Green
                    }
                }
                
                $logsFound = $true
                $info += "Session Logs: $logCount Dateien ($logSizeMB MB)"
                break
            }
        }
    }
    
    if (-not $logsFound) {
        Write-Host "  Session Logs: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Session Logs Check: Fehlgeschlagen" -ForegroundColor Red
}

# 7. Performance Check (Detailed Mode)
if ($Detailed) {
    Write-Host "`n=== PERFORMANCE CHECK ===" -ForegroundColor Yellow
    
    try {
        # Check session performance metrics
        $sessionProcesses = Get-Process | Where-Object { 
            $_.ProcessName -like "*vmware*" -or 
            $_.ProcessName -like "*horizon*" -or 
            $_.ProcessName -like "*view*" 
        }
        
        if ($sessionProcesses) {
            $totalMemory = ($sessionProcesses | Measure-Object -Property WorkingSet -Sum).Sum
            $totalMemoryMB = [math]::Round($totalMemory/1MB, 2)
            
            Write-Host "  Session Memory: $totalMemoryMB MB" -ForegroundColor White
            
            if ($totalMemoryMB -gt 300) {
                $warnings += "Session: Hoher Memory-Verbrauch ($totalMemoryMB MB)"
                Write-Host "  Status: HOCH - Performance-Überwachung empfohlen" -ForegroundColor Yellow
            } else {
                Write-Host "  Status: OK" -ForegroundColor Green
            }
            
            $info += "Session Memory: $totalMemoryMB MB"
        } else {
            Write-Host "  Session Processes: Nicht aktiv" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "  Performance Check: Fehlgeschlagen" -ForegroundColor Red
    }
}

# 8. Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     OMNISSA SESSION INFO SUMMARY" -ForegroundColor Cyan
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
    Write-Host "  Status: OK - Session Info scheint gesund zu sein" -ForegroundColor Green
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Omnissa Session Info Check completed!" -ForegroundColor Green
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
