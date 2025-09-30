# Enterprise Environment Scanner (OPTIMIZED VERSION)
# Fast registry-based queries instead of slow WMI calls
# Scans for enterprise software and system configuration

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     Enterprise Environment Scanner (FAST)" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# 1. System Information (fast)
Write-Host "`n=== SYSTEM INFORMATION ===" -ForegroundColor Yellow
try {
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $cs = Get-WmiObject -Class Win32_ComputerSystem
    Write-Host "OS: $($os.Caption) $($os.Version)" -ForegroundColor White
    Write-Host "RAM: $([math]::Round($cs.TotalPhysicalMemory/1GB, 2)) GB" -ForegroundColor White
    Write-Host "CPU: $($cs.NumberOfProcessors) processors" -ForegroundColor White
} catch {
    Write-Host "System info not available" -ForegroundColor Red
}

# 2. Installed Software (Registry - VERY FAST)
Write-Host "`n=== INSTALLED ENTERPRISE SOFTWARE ===" -ForegroundColor Yellow
$enterpriseKeywords = @(
    "Omnissa", "FSLogix", "ServiceNow", "Teams", "Office", "Outlook",
    "OneDrive", "SharePoint", "Citrix", "VMware", "Defender", "McAfee",
    "Symantec", "Trend", "Sophos", "CrowdStrike", "Adobe", "Java",
    "Microsoft 365", "Office 365", "Office 2016", "Office 2019", "Office 2021"
)

$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$foundSoftware = @()
foreach ($path in $regPaths) {
    try {
        $items = Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { 
            $_.DisplayName -and $_.DisplayName -notlike "*Update*" -and $_.DisplayName -notlike "*KB*"
        }
        foreach ($item in $items) {
            foreach ($keyword in $healthcareKeywords) {
                if ($item.DisplayName -like "*$keyword*") {
                    $foundSoftware += [PSCustomObject]@{
                        Name = $item.DisplayName
                        Version = $item.DisplayVersion
                        Publisher = $item.Publisher
                    }
                    break
                }
            }
        }
    } catch { }
}

if ($foundSoftware) {
    $foundSoftware | Sort-Object Name | Format-Table -AutoSize
} else {
    Write-Host "Keine Healthcare-Software gefunden" -ForegroundColor Gray
}

# 3. Laufende Prozesse (schnell)
Write-Host "`n=== RUNNING HEALTHCARE PROCESSES ===" -ForegroundColor Yellow
$healthcareProcesses = @(
    "omnissa", "fslogix", "kis", "servicenow", "epic", "cerner", "allscripts",
    "meditech", "nextgen", "athena", "eclinical", "greenway", "mckesson",
    "siemens", "ge", "philips", "teams", "outlook", "onedrive", "citrix",
    "vmware", "rdp", "vpn", "defender", "mcafee", "symantec", "trend",
    "sophos", "crowdstrike", "adobe", "java", "chrome", "firefox", "edge"
)

$runningProcesses = Get-Process | Where-Object { 
    $processName = $_.ProcessName.ToLower()
    $healthcareProcesses | Where-Object { $processName -like "*$_*" }
} | Select-Object ProcessName, Id, @{Name="Memory(MB)";Expression={[math]::Round($_.WorkingSet/1MB,2)}} | Sort-Object ProcessName

if ($runningProcesses) {
    $runningProcesses | Format-Table -AutoSize
} else {
    Write-Host "Keine Healthcare-Prozesse gefunden" -ForegroundColor Gray
}

# 4. Services (schnell)
Write-Host "`n=== HEALTHCARE SERVICES ===" -ForegroundColor Yellow
$healthcareServices = @(
    "omnissa", "fslogix", "kis", "servicenow", "epic", "cerner", "allscripts",
    "meditech", "nextgen", "athena", "eclinical", "greenway", "mckesson",
    "siemens", "ge", "philips", "teams", "outlook", "onedrive", "citrix",
    "vmware", "rdp", "vpn", "defender", "mcafee", "symantec", "trend",
    "sophos", "crowdstrike"
)

$foundServices = Get-Service | Where-Object { 
    $serviceName = $_.Name.ToLower()
    $displayName = $_.DisplayName.ToLower()
    $healthcareServices | Where-Object { 
        $serviceName -like "*$_*" -or $displayName -like "*$_*" 
    }
} | Select-Object Name, DisplayName, Status, StartType | Sort-Object Name

if ($foundServices) {
    $foundServices | Format-Table -AutoSize
} else {
    Write-Host "Keine Healthcare-Services gefunden" -ForegroundColor Gray
}

# 5. Disk-Usage (schnell) + P-Laufwerk Check
Write-Host "`n=== DISK USAGE ===" -ForegroundColor Yellow
try {
    $disks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"
    foreach ($disk in $disks) {
        $sizeGB = [math]::Round($disk.Size/1GB, 2)
        $freeGB = [math]::Round($disk.FreeSpace/1GB, 2)
        $percentFree = [math]::Round(($disk.FreeSpace/$disk.Size)*100, 2)
        $status = if ($percentFree -lt 10) { "KRITISCH" } elseif ($percentFree -lt 20) { "WARNUNG" } else { "OK" }
        
        Write-Host "$($disk.DeviceID) - $sizeGB GB total, $freeGB GB frei ($percentFree%) $status" -ForegroundColor White
    }
} catch {
    Write-Host "Disk-Info nicht verfügbar" -ForegroundColor Red
}

# P-Laufwerk spezifischer Check (Healthcare Home Drive)
Write-Host "`n=== P-LAUFWERK CHECK (Healthcare Home Drive) ===" -ForegroundColor Yellow
try {
    $pDrive = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='P:'"
    if ($pDrive) {
        $pSizeGB = [math]::Round($pDrive.Size/1GB, 2)
        $pFreeGB = [math]::Round($pDrive.FreeSpace/1GB, 2)
        $pPercentFree = [math]::Round(($pDrive.FreeSpace/$pDrive.Size)*100, 2)
        
        Write-Host "P: - $pSizeGB GB total, $pFreeGB GB frei ($pPercentFree%)" -ForegroundColor White
        
        # Standard 5GB Check
        if ($pSizeGB -eq 5) {
            Write-Host "  OK: Standard 5GB P-Laufwerk (nicht erweitert)" -ForegroundColor Green
        } elseif ($pSizeGB -gt 5) {
            Write-Host "  INFO: P-Laufwerk erweitert: $pSizeGB GB (User hat Bestellung aufgegeben)" -ForegroundColor Cyan
        } else {
            Write-Host "  WARNUNG: P-Laufwerk kleiner als Standard: $pSizeGB GB" -ForegroundColor Yellow
        }
        
        # Speicherplatz-Warnungen
        if ($pPercentFree -lt 5) {
            Write-Host "  KRITISCH: P-Laufwerk fast voll! ($pPercentFree% frei)" -ForegroundColor Red
        } elseif ($pPercentFree -lt 15) {
            Write-Host "  WARNUNG: P-Laufwerk wird voll ($pPercentFree% frei)" -ForegroundColor Yellow
        } else {
            Write-Host "  OK: P-Laufwerk Speicherplatz OK ($pPercentFree% frei)" -ForegroundColor Green
        }
        
        # Screenpresso Check (bekanntes Problem)
        $screenpressoPath = "P:\Screenpresso"
        if (Test-Path $screenpressoPath) {
            $screenpressoSize = (Get-ChildItem $screenpressoPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $screenpressoSizeGB = [math]::Round($screenpressoSize/1GB, 2)
            Write-Host "  Screenpresso Ordner: $screenpressoSizeGB GB" -ForegroundColor White
            if ($screenpressoSizeGB -gt 2) {
                Write-Host "    WARNUNG: Screenpresso verbraucht viel Speicherplatz!" -ForegroundColor Yellow
            }
        }
        
    } else {
        Write-Host "P: - Nicht verfügbar (Administrativer Account oder nicht gemappt)" -ForegroundColor Gray
    }
} catch {
    Write-Host "P-Laufwerk Check fehlgeschlagen" -ForegroundColor Red
}

# 6. Netzwerk-Verbindungen (schnell - nur wichtige)
Write-Host "`n=== NETWORK CONNECTIONS (Top 10) ===" -ForegroundColor Yellow
try {
    $connections = Get-NetTCPConnection | Where-Object { $_.State -eq "Established" } | 
    Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort | 
    Sort-Object LocalPort | Select-Object -First 10
    $connections | Format-Table -AutoSize
} catch {
    Write-Host "Netzwerk-Info nicht verfügbar" -ForegroundColor Red
}

# 7. Event Logs (nur letzte 5 kritische Fehler)
Write-Host "`n=== RECENT CRITICAL ERRORS (Last 5) ===" -ForegroundColor Yellow
try {
    $events = Get-WinEvent -FilterHashtable @{LogName='System','Application'; Level=1,2; StartTime=(Get-Date).AddHours(-6)} -MaxEvents 5 -ErrorAction SilentlyContinue
    if ($events) {
        foreach ($event in $events) {
            Write-Host "$($event.TimeCreated) - $($event.LevelDisplayName) - $($event.ProviderName)" -ForegroundColor White
            Write-Host "  $($event.Message.Substring(0, [Math]::Min(100, $event.Message.Length)))..." -ForegroundColor Gray
        }
    } else {
        Write-Host "Keine kritischen Events in den letzten 6h" -ForegroundColor Gray
    }
} catch {
    Write-Host "Event Logs nicht verfügbar" -ForegroundColor Red
}

# 8. PowerShell-Module (schnell)
Write-Host "`n=== POWERSHELL MODULES ===" -ForegroundColor Yellow
try {
    $modules = Get-Module -ListAvailable | Where-Object { 
        $_.Name -like "*healthcare*" -or $_.Name -like "*medical*" -or 
        $_.Name -like "*epic*" -or $_.Name -like "*cerner*" -or
        $_.Name -like "*teams*" -or $_.Name -like "*office*"
    } | Select-Object Name, Version
    if ($modules) {
        $modules | Format-Table -AutoSize
    } else {
        Write-Host "Keine relevanten Module gefunden" -ForegroundColor Gray
    }
} catch {
    Write-Host "Module-Info nicht verfügbar" -ForegroundColor Red
}

# 9. Healthcare-spezifische Checks
Write-Host "`n=== HEALTHCARE-SPEZIFISCHE CHECKS ===" -ForegroundColor Yellow

# FSLogix Check
Write-Host "`nFSLogix Status:" -ForegroundColor White
try {
    $fslogixService = Get-Service -Name "*fslogix*" -ErrorAction SilentlyContinue
    if ($fslogixService) {
        Write-Host "  FSLogix Service: $($fslogixService.Status)" -ForegroundColor White
    } else {
        Write-Host "  FSLogix Service: Nicht gefunden" -ForegroundColor Gray
    }
    
    # FSLogix Registry Check
    $fslogixReg = Get-ItemProperty "HKLM:\SOFTWARE\FSLogix" -ErrorAction SilentlyContinue
    if ($fslogixReg) {
        Write-Host "  FSLogix Registry: Gefunden" -ForegroundColor Green
    } else {
        Write-Host "  FSLogix Registry: Nicht gefunden" -ForegroundColor Gray
    }
} catch {
    Write-Host "  FSLogix Check fehlgeschlagen" -ForegroundColor Red
}

# Citrix Check
Write-Host "`nCitrix Status:" -ForegroundColor White
try {
    $citrixProcesses = Get-Process | Where-Object { $_.ProcessName -like "*citrix*" -or $_.ProcessName -like "*receiver*" }
    if ($citrixProcesses) {
        Write-Host "  Citrix Prozesse: $($citrixProcesses.Count) gefunden" -ForegroundColor White
        $citrixProcesses | ForEach-Object { Write-Host "    - $($_.ProcessName)" -ForegroundColor Gray }
    } else {
        Write-Host "  Citrix: Nicht aktiv" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Citrix Check fehlgeschlagen" -ForegroundColor Red
}

# Teams Check (Healthcare Communication) - New Teams 2.0
Write-Host "`nMicrosoft Teams Status (New Teams 2.0):" -ForegroundColor White
try {
    # Check for New Teams processes
    $teamsProcesses = @("ms-teams", "teams", "Teams")
    $foundTeams = $false
    
    foreach ($teamsName in $teamsProcesses) {
        $teamsProcess = Get-Process -Name $teamsName -ErrorAction SilentlyContinue
        if ($teamsProcess) {
            Write-Host "  Teams ($teamsName): Läuft (PID: $($teamsProcess.Id))" -ForegroundColor Green
            $foundTeams = $true
        }
    }
    
    if (-not $foundTeams) {
        Write-Host "  Teams: Nicht aktiv" -ForegroundColor Gray
    }
    
    # New Teams Cache Check (different paths)
    $teamsCachePaths = @(
        "$env:APPDATA\Microsoft\Teams",
        "$env:LOCALAPPDATA\Microsoft\Teams",
        "$env:APPDATA\Microsoft\Teams (work or school)"
    )
    
    $totalCacheSize = 0
    foreach ($cachePath in $teamsCachePaths) {
        if (Test-Path $cachePath) {
            $cacheSize = (Get-ChildItem $cachePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $totalCacheSize += $cacheSize
            $cacheSizeGB = [math]::Round($cacheSize/1GB, 2)
            Write-Host "  Teams Cache ($(Split-Path $cachePath -Leaf)): $cacheSizeGB GB" -ForegroundColor White
        }
    }
    
    if ($totalCacheSize -gt 0) {
        $totalCacheSizeGB = [math]::Round($totalCacheSize/1GB, 2)
        Write-Host "  Teams Gesamt-Cache: $totalCacheSizeGB GB" -ForegroundColor White
        if ($totalCacheSizeGB -gt 2) {
            Write-Host "    WARNUNG: Teams Cache groß - möglicherweise Cleanup nötig" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Teams Cache: Nicht gefunden" -ForegroundColor Gray
    }
    
    # Check for New Teams installation
    $teamsRegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $teamsInstalled = $false
    foreach ($regPath in $teamsRegPaths) {
        try {
            $teamsReg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { 
                $_.DisplayName -like "*Teams*" -and $_.DisplayName -notlike "*Update*"
            }
            if ($teamsReg) {
                Write-Host "  Teams Installation: $($teamsReg.DisplayName)" -ForegroundColor Green
                $teamsInstalled = $true
                break
            }
        } catch { }
    }
    
    if (-not $teamsInstalled) {
        Write-Host "  Teams Installation: Nicht gefunden" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Teams Check fehlgeschlagen" -ForegroundColor Red
}

# Outlook Check (Healthcare Email)
Write-Host "`nMicrosoft Outlook Status:" -ForegroundColor White
try {
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Host "  Outlook: Läuft (PID: $($outlookProcess.Id))" -ForegroundColor Green
    } else {
        Write-Host "  Outlook: Nicht aktiv" -ForegroundColor Gray
    }
    
    # Outlook OST Check
    $ostPath = "$env:LOCALAPPDATA\Microsoft\Outlook"
    if (Test-Path $ostPath) {
        $ostFiles = Get-ChildItem $ostPath -Filter "*.ost" -ErrorAction SilentlyContinue
        if ($ostFiles) {
            foreach ($ost in $ostFiles) {
                $ostSizeGB = [math]::Round($ost.Length/1GB, 2)
                Write-Host "  OST Datei: $($ost.Name) - $ostSizeGB GB" -ForegroundColor White
                if ($ostSizeGB -gt 2) {
                    Write-Host "    WARNUNG: Große OST-Datei - möglicherweise Cleanup nötig" -ForegroundColor Yellow
                }
            }
        }
    }
} catch {
    Write-Host "  Outlook Check fehlgeschlagen" -ForegroundColor Red
}

# 10. Standard Corporate Checks (Schweiz)
Write-Host "`n=== STANDARD CORPORATE CHECKS (Schweiz) ===" -ForegroundColor Yellow

# Windows Update Status
Write-Host "`nWindows Update Status:" -ForegroundColor White
try {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $updateSearcher.Search("IsInstalled=0")
    Write-Host "  Pending Updates: $($searchResult.Updates.Count)" -ForegroundColor White
    if ($searchResult.Updates.Count -gt 10) {
        Write-Host "    WARNUNG: Viele ausstehende Updates!" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Update-Info nicht verfügbar" -ForegroundColor Gray
}

# Antivirus Status
Write-Host "`nAntivirus Status:" -ForegroundColor White
try {
    $antivirus = Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiVirusProduct -ErrorAction SilentlyContinue
    if ($antivirus) {
        foreach ($av in $antivirus) {
            Write-Host "  $($av.displayName) - $($av.productState)" -ForegroundColor White
        }
    } else {
        Write-Host "  Antivirus-Info nicht verfügbar" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Antivirus-Info nicht verfügbar" -ForegroundColor Gray
}

# Windows Defender Status
Write-Host "`nWindows Defender Status:" -ForegroundColor White
try {
    $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defenderStatus) {
        Write-Host "  Real-time Protection: $($defenderStatus.RealTimeProtectionEnabled)" -ForegroundColor White
        Write-Host "  Antivirus Enabled: $($defenderStatus.AntivirusEnabled)" -ForegroundColor White
        Write-Host "  Last Quick Scan: $($defenderStatus.QuickScanStartTime)" -ForegroundColor White
        Write-Host "  Last Full Scan: $($defenderStatus.FullScanStartTime)" -ForegroundColor White
    } else {
        Write-Host "  Defender-Info nicht verfügbar" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Defender-Info nicht verfügbar" -ForegroundColor Gray
}

# UAC Status (Standard Corporate)
Write-Host "`nUAC Status:" -ForegroundColor White
try {
    $uacReg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -ErrorAction SilentlyContinue
    if ($uacReg.EnableLUA -eq 1) {
        Write-Host "  UAC: Aktiviert (Standard)" -ForegroundColor Green
    } else {
        Write-Host "  UAC: Deaktiviert (NICHT STANDARD!)" -ForegroundColor Red
    }
} catch {
    Write-Host "  UAC-Info nicht verfügbar" -ForegroundColor Gray
}

# Firewall Status
Write-Host "`nWindows Firewall Status:" -ForegroundColor White
try {
    $firewallProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
    foreach ($profile in $firewallProfiles) {
        Write-Host "  $($profile.Name): $($profile.Enabled)" -ForegroundColor White
    }
} catch {
    Write-Host "  Firewall-Info nicht verfügbar" -ForegroundColor Gray
}

# BitLocker Status (Corporate Standard)
Write-Host "`nBitLocker Status:" -ForegroundColor White
try {
    $bitlocker = Get-WmiObject -Namespace "root\cimv2\security\microsoftvolumeencryption" -Class "Win32_EncryptableVolume" -ErrorAction SilentlyContinue
    if ($bitlocker) {
        foreach ($volume in $bitlocker) {
            $driveLetter = $volume.DriveLetter
            $protectionStatus = $volume.GetProtectionStatus()
            Write-Host "  $driveLetter`: Protection Status: $($protectionStatus.ProtectionStatus)" -ForegroundColor White
        }
    } else {
        Write-Host "  BitLocker-Info nicht verfügbar" -ForegroundColor Gray
    }
} catch {
    Write-Host "  BitLocker-Info nicht verfügbar" -ForegroundColor Gray
}

# Group Policy Status
Write-Host "`nGroup Policy Status:" -ForegroundColor White
try {
    $gpResult = gpresult /r 2>$null
    if ($gpResult) {
        $appliedGPOs = $gpResult | Select-String "Applied Group Policy Objects"
        if ($appliedGPOs) {
            Write-Host "  GPOs: Angewendet" -ForegroundColor Green
        } else {
            Write-Host "  GPOs: Keine gefunden" -ForegroundColor Gray
        }
    } else {
        Write-Host "  GPO-Info nicht verfügbar" -ForegroundColor Gray
    }
} catch {
    Write-Host "  GPO-Info nicht verfügbar" -ForegroundColor Gray
}

# Domain Status
Write-Host "`nDomain Status:" -ForegroundColor White
try {
    $domain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
    $workgroup = (Get-WmiObject -Class Win32_ComputerSystem).Workgroup
    if ($domain -and $domain -ne $env:COMPUTERNAME) {
        Write-Host "  Domain: $domain" -ForegroundColor Green
    } elseif ($workgroup) {
        Write-Host "  Workgroup: $workgroup" -ForegroundColor Yellow
    } else {
        Write-Host "  Domain: Nicht verfügbar" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Domain-Info nicht verfügbar" -ForegroundColor Gray
}

# 11. Microsoft Office Check (O365 + Office 2016)
Write-Host "`n=== MICROSOFT OFFICE STATUS ===" -ForegroundColor Yellow

# Office 365 / Microsoft 365 Check
Write-Host "`nOffice 365 / Microsoft 365 Status:" -ForegroundColor White
try {
    $office365Processes = @("winword", "excel", "powerpnt", "outlook", "msaccess", "mspub", "onenote")
    $runningOffice365 = @()
    
    foreach ($process in $office365Processes) {
        $proc = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($proc) {
            $runningOffice365 += $process
        }
    }
    
    if ($runningOffice365) {
        Write-Host "  Office Apps laufen: $($runningOffice365 -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "  Office Apps: Nicht aktiv" -ForegroundColor Gray
    }
    
    # Office 365 Installation Check
    try {
        $office365Reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | 
        Where-Object { $_.DisplayName -like "*Microsoft 365*" -or $_.DisplayName -like "*Office 365*" }
        
        if ($office365Reg) {
            Write-Host "  Office 365 Installation: $($office365Reg.DisplayName)" -ForegroundColor Green
            Write-Host "  Version: $($office365Reg.DisplayVersion)" -ForegroundColor White
        } else {
            Write-Host "  Office 365: Nicht installiert" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  Office 365 Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Office 2016 Check
    try {
        $office2016Reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | 
        Where-Object { $_.DisplayName -like "*Office 2016*" }
        
        if ($office2016Reg) {
            Write-Host "  Office 2016 Installation: $($office2016Reg.DisplayName)" -ForegroundColor Green
            Write-Host "  Version: $($office2016Reg.DisplayVersion)" -ForegroundColor White
        } else {
            Write-Host "  Office 2016: Nicht installiert" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  Office 2016 Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Office Cache Check (häufiges Problem)
    $officeCachePaths = @(
        "$env:APPDATA\Microsoft\Office\16.0\OfficeFileCache",
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache",
        "$env:APPDATA\Microsoft\Office\UnsavedFiles",
        "$env:LOCALAPPDATA\Microsoft\Office\UnsavedFiles"
    )
    
    $totalOfficeCache = 0
    foreach ($cachePath in $officeCachePaths) {
        if (Test-Path $cachePath) {
            $cacheSize = (Get-ChildItem $cachePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $totalOfficeCache += $cacheSize
        }
    }
    
    if ($totalOfficeCache -gt 0) {
        $officeCacheGB = [math]::Round($totalOfficeCache/1GB, 2)
        Write-Host "  Office Cache: $officeCacheGB GB" -ForegroundColor White
        if ($officeCacheGB -gt 1) {
            Write-Host "    WARNUNG: Office Cache groß - möglicherweise Cleanup nötig" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Office Cache: Nicht gefunden" -ForegroundColor Gray
    }
    
    # Office Update Status
    Write-Host "`nOffice Update Status:" -ForegroundColor White
    try {
        $officeUpdateReg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
        if ($officeUpdateReg) {
            Write-Host "  Click-to-Run: Aktiviert" -ForegroundColor Green
            Write-Host "  Update Channel: $($officeUpdateReg.UpdateChannel)" -ForegroundColor White
        } else {
            Write-Host "  Click-to-Run: Nicht verfügbar (MSI-Installation?)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  Office Update Info: Nicht verfügbar" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "  Office Check fehlgeschlagen" -ForegroundColor Red
}

# 12. VDI Environment Check (IC, ICG, WS, WN)
Write-Host "`n=== VDI ENVIRONMENT CHECK ===" -ForegroundColor Yellow

# VDI Type Detection
Write-Host "`nVDI Type Detection:" -ForegroundColor White
try {
    # Check for VDI indicators
    $vdiIndicators = @()
    
    # Check for VMware VDI
    try {
        $vmwareTools = Get-Service -Name "VMTools" -ErrorAction SilentlyContinue
        if ($vmwareTools) {
            $vdiIndicators += "VMware Tools: $($vmwareTools.Status)"
        }
    } catch {
        Write-Host "  VMware Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Check for Citrix VDI
    try {
        $citrixServices = Get-Service | Where-Object { $_.Name -like "*citrix*" -or $_.DisplayName -like "*citrix*" }
        if ($citrixServices) {
            $vdiIndicators += "Citrix Services: $($citrixServices.Count) gefunden"
        }
    } catch {
        Write-Host "  Citrix Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Check for VDI-specific processes
    try {
        $vdiProcesses = @("vmware", "vmtoolsd", "citrix", "teradici", "pcoip")
        $foundVDIProcesses = @()
        foreach ($process in $vdiProcesses) {
            try {
                $proc = Get-Process | Where-Object { $_.ProcessName -like "*$process*" }
                if ($proc) {
                    $foundVDIProcesses += $process
                }
            } catch { }
        }
        
        if ($foundVDIProcesses) {
            $vdiIndicators += "VDI Prozesse: $($foundVDIProcesses -join ', ')"
        }
    } catch {
        Write-Host "  VDI Prozess Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    # Check for VDI-specific registry entries
    try {
        $vdiRegPaths = @(
            "HKLM:\SOFTWARE\VMware, Inc.\VMware Tools",
            "HKLM:\SOFTWARE\Citrix",
            "HKLM:\SOFTWARE\Teradici",
            "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        )
        
        foreach ($regPath in $vdiRegPaths) {
            try {
                if (Test-Path $regPath) {
                    $vdiIndicators += "Registry: $(Split-Path $regPath -Leaf)"
                }
            } catch { }
        }
    } catch {
        Write-Host "  Registry Check: Fehlgeschlagen" -ForegroundColor Gray
    }
    
    if ($vdiIndicators) {
        Write-Host "  VDI Indicators gefunden:" -ForegroundColor Green
        foreach ($indicator in $vdiIndicators) {
            Write-Host "    - $indicator" -ForegroundColor White
        }
    } else {
        Write-Host "  VDI Indicators: Nicht gefunden (möglicherweise physischer PC)" -ForegroundColor Gray
    }
    
    # Check for GPU (ICG vs IC)
    Write-Host "`nGPU Status (ICG vs IC):" -ForegroundColor White
    try {
        $gpuInfo = Get-WmiObject -Class Win32_VideoController | Where-Object { $_.Name -notlike "*Basic*" -and $_.Name -notlike "*Standard*" }
        if ($gpuInfo) {
            foreach ($gpu in $gpuInfo) {
                Write-Host "  GPU: $($gpu.Name)" -ForegroundColor White
                Write-Host "    VRAM: $([math]::Round($gpu.AdapterRAM/1GB, 2)) GB" -ForegroundColor White
                if ($gpu.Name -like "*NVIDIA*" -or $gpu.Name -like "*AMD*" -or $gpu.Name -like "*Radeon*") {
                    Write-Host "    Typ: ICG (mit GPU)" -ForegroundColor Green
                } else {
                    Write-Host "    Typ: IC (ohne dedizierte GPU)" -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "  GPU: Keine dedizierte GPU gefunden (IC-Typ)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  GPU Info: Nicht verfügbar" -ForegroundColor Gray
    }
    
    # Check for VDI-specific performance issues
    Write-Host "`nVDI Performance Check:" -ForegroundColor White
    try {
        # Check for high CPU usage
        $cpuUsage = Get-WmiObject -Class Win32_Processor | Measure-Object -Property LoadPercentage -Average
        if ($cpuUsage.Average -gt 80) {
            Write-Host "  CPU Usage: $($cpuUsage.Average)% (HOCH - VDI Performance Problem)" -ForegroundColor Red
        } elseif ($cpuUsage.Average -gt 60) {
            Write-Host "  CPU Usage: $($cpuUsage.Average)% (MITTEL)" -ForegroundColor Yellow
        } else {
            Write-Host "  CPU Usage: $($cpuUsage.Average)% (OK)" -ForegroundColor Green
        }
        
        # Check for high memory usage
        $totalRAM = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory/1GB, 2)
        $freeRAM = [math]::Round((Get-WmiObject -Class Win32_OperatingSystem).FreePhysicalMemory/1MB, 2)
        $usedRAM = $totalRAM - $freeRAM
        $ramPercent = [math]::Round(($usedRAM/$totalRAM)*100, 2)
        
        if ($ramPercent -gt 90) {
            Write-Host "  RAM Usage: $ramPercent% (KRITISCH - VDI Performance Problem)" -ForegroundColor Red
        } elseif ($ramPercent -gt 80) {
            Write-Host "  RAM Usage: $ramPercent% (HOCH)" -ForegroundColor Yellow
        } else {
            Write-Host "  RAM Usage: $ramPercent% (OK)" -ForegroundColor Green
        }
        
    } catch {
        Write-Host "  Performance Check: Fehlgeschlagen" -ForegroundColor Red
    }
    
    # Check for VDI-specific network issues
    Write-Host "`nVDI Network Check:" -ForegroundColor White
    try {
        $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        $vdiAdapters = $networkAdapters | Where-Object { $_.Name -like "*VMware*" -or $_.Name -like "*Citrix*" -or $_.Name -like "*Virtual*" }
        
        if ($vdiAdapters) {
            Write-Host "  VDI Network Adapter: $($vdiAdapters.Name)" -ForegroundColor Green
            foreach ($adapter in $vdiAdapters) {
                Write-Host "    Link Speed: $($adapter.LinkSpeed)" -ForegroundColor White
            }
        } else {
            Write-Host "  VDI Network: Standard Adapter" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  Network Check: Fehlgeschlagen" -ForegroundColor Red
    }
    
} catch {
    Write-Host "  VDI Check fehlgeschlagen" -ForegroundColor Red
}

# 13. Healthcare-spezifische Zusatz-Checks
Write-Host "`n=== HEALTHCARE ZUSATZ-CHECKS ===" -ForegroundColor Yellow

# Java Version Check (wichtig für Healthcare-Apps)
Write-Host "`nJava Status:" -ForegroundColor White
try {
    $javaVersion = java -version 2>&1 | Select-String "version"
    if ($javaVersion) {
        Write-Host "  Java: $($javaVersion.Line)" -ForegroundColor White
    } else {
        Write-Host "  Java: Nicht installiert oder nicht im PATH" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Java: Nicht verfügbar" -ForegroundColor Gray
}

# Adobe Reader Check (häufig in Healthcare)
Write-Host "`nAdobe Reader Status:" -ForegroundColor White
try {
    $adobeReader = Get-Process -Name "AcroRd32" -ErrorAction SilentlyContinue
    if ($adobeReader) {
        Write-Host "  Adobe Reader: Läuft (PID: $($adobeReader.Id))" -ForegroundColor Green
    } else {
        Write-Host "  Adobe Reader: Nicht aktiv" -ForegroundColor Gray
    }
    
    # Adobe Reader Installation Check
    $adobeReg = Get-ItemProperty "HKLM:\SOFTWARE\Adobe\Acrobat Reader" -ErrorAction SilentlyContinue
    if ($adobeReg) {
        Write-Host "  Adobe Reader: Installiert" -ForegroundColor Green
    } else {
        Write-Host "  Adobe Reader: Nicht installiert" -ForegroundColor Gray
    }
} catch {
    Write-Host "  Adobe Reader Check fehlgeschlagen" -ForegroundColor Red
}

# Browser Check (Healthcare Web-Apps)
Write-Host "`nBrowser Status:" -ForegroundColor White
$browsers = @("chrome", "firefox", "msedge", "iexplore")
foreach ($browser in $browsers) {
    try {
        $browserProcess = Get-Process -Name $browser -ErrorAction SilentlyContinue
        if ($browserProcess) {
            Write-Host "  $browser`: Läuft (PID: $($browserProcess.Id))" -ForegroundColor Green
        } else {
            Write-Host "  $browser`: Nicht aktiv" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  $browser`: Check fehlgeschlagen" -ForegroundColor Red
    }
}

# VPN Check (häufig in Healthcare)
Write-Host "`nVPN Status:" -ForegroundColor White
$vpnProcesses = @("openvpn", "forticlient", "cisco", "pulse", "globalprotect", "anyconnect")
$foundVPN = $false
foreach ($vpn in $vpnProcesses) {
    try {
        $vpnProcess = Get-Process | Where-Object { $_.ProcessName -like "*$vpn*" }
        if ($vpnProcess) {
            Write-Host "  VPN ($vpn): Läuft" -ForegroundColor Green
            $foundVPN = $true
        }
    } catch { }
}
if (-not $foundVPN) {
    Write-Host "  VPN: Keine aktiven VPN-Client gefunden" -ForegroundColor Gray
}

# Remote Desktop Check
Write-Host "`nRemote Desktop Status:" -ForegroundColor White
try {
    $rdpService = Get-Service -Name "TermService" -ErrorAction SilentlyContinue
    if ($rdpService) {
        Write-Host "  RDP Service: $($rdpService.Status)" -ForegroundColor White
    }
    
    $rdpProcess = Get-Process -Name "mstsc" -ErrorAction SilentlyContinue
    if ($rdpProcess) {
        Write-Host "  RDP Client: Läuft" -ForegroundColor Green
    } else {
        Write-Host "  RDP Client: Nicht aktiv" -ForegroundColor Gray
    }
} catch {
    Write-Host "  RDP Check fehlgeschlagen" -ForegroundColor Red
}

# OneDrive Check (Corporate File Sync)
Write-Host "`nOneDrive Status:" -ForegroundColor White
try {
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDriveProcess) {
        Write-Host "  OneDrive: Läuft (PID: $($oneDriveProcess.Id))" -ForegroundColor Green
    } else {
        Write-Host "  OneDrive: Nicht aktiv" -ForegroundColor Gray
    }
    
    # OneDrive Sync Status
    $oneDrivePath = "$env:USERPROFILE\OneDrive - USZ"
    if (Test-Path $oneDrivePath) {
        Write-Host "  OneDrive Ordner: Gefunden" -ForegroundColor Green
    } else {
        Write-Host "  OneDrive Ordner: Nicht gefunden" -ForegroundColor Gray
    }
} catch {
    Write-Host "  OneDrive Check fehlgeschlagen" -ForegroundColor Red
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "FAST Environment Scan completed!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# Clean exit - change to neutral system directory
Set-Location "C:\Windows\System32" | Out-Null

Read-Host "`nDrücken Sie Enter zum Beenden..."

exit 0