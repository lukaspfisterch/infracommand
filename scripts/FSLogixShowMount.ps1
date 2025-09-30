# FSLogix Show Mount Status
# Basic VHD mount status check for FSLogix profile containers
# Displays mounted VHD files and basic FSLogix configuration

param(
    [switch]$Verbose = $false
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     FSLOGIX SHOW MOUNT STATUS" -ForegroundColor Cyan
Write-Host "     User: $env:USERNAME" -ForegroundColor Cyan
Write-Host "     Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "`nNOTE: This is a basic implementation" -ForegroundColor Yellow
Write-Host "Full implementation will follow in future version" -ForegroundColor Yellow

# Basic VHD Mount Check
Write-Host "`n=== BASIC VHD CHECK ===" -ForegroundColor Yellow

try {
    $mountedVHDs = Get-Disk -ErrorAction SilentlyContinue | Where-Object { $_.BusType -eq "File Backed Virtual" }
    
    if ($mountedVHDs) {
        Write-Host "  Mounted VHDs: $($mountedVHDs.Count) found" -ForegroundColor Green
        
        foreach ($disk in $mountedVHDs) {
            $diskNumber = $disk.Number
            $diskSize = [math]::Round($disk.Size/1GB, 2)
            Write-Host "    Disk ${diskNumber}: ${diskSize} GB" -ForegroundColor White
        }
    } else {
        Write-Host "  Mounted VHDs: None found" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  VHD Check: Failed" -ForegroundColor Red
}

# Basic Registry Check
Write-Host "`n=== BASIC REGISTRY CHECK ===" -ForegroundColor Yellow

try {
    $configRegPath = "HKLM:\SOFTWARE\FSLogix\Profiles"
    if (Test-Path $configRegPath) {
        Write-Host "  FSLogix Registry: Found" -ForegroundColor Green
    } else {
        Write-Host "  FSLogix Registry: Not found" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Registry Check: Failed" -ForegroundColor Red
}

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "Basic check completed!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# Clean exit
Set-Location "C:\Windows\System32" | Out-Null

if ($Verbose) {
    Read-Host "`nPress Enter to exit..."
} else {
    Start-Sleep -Milliseconds 500
}

exit 0