# System Baseline Scripts

## Overview

These scripts are designed for virtual environments and provide fast, focused system checks for IT administrators.

## Scripts

### 1. Baseline.ps1 (Standard)
**Purpose:** Quick, basic system checks for daily use

**Features:**
- System Health (RAM, CPU)
- Home Drive Check (P: drive)
- Office Apps (Teams, Outlook, OneDrive)
- Virtual Environment Detection
- Security Check (UAC, Domain)
- Advanced Checks (optional with `-Advanced`)

**Usage:**
```powershell
# Standard
.\Baseline.ps1

# With Advanced Checks
.\Baseline.ps1 -Advanced

# With Verbose Output
.\Baseline.ps1 -Verbose
```

### 2. Baseline_Advanced.ps1 (Extended)
**Purpose:** Comprehensive system analysis for detailed diagnosis

**Features:**
- All Standard Features
- Detailed System Information
- Extended Virtual Environment Performance Checks
- Windows Defender Status
- Java Version Check (for Enterprise Apps)
- Detailed Cache Analysis
- Home Drive Content Analysis

**Usage:**
```powershell
# Standard
.\Baseline_Advanced.ps1

# With Verbose Output
.\Baseline_Advanced.ps1 -Verbose
```

## Virtual Environment Specific Checks

### Home Drive (P: Drive)
- **Standard:** 5GB (not extended)
- **Extended:** >5GB (User requested extension)
- **Warnings:** <15% free, <5% critical
- **Screenshot Check:** Known issue with large screenshots

### Virtual Environment
- **VMware Tools:** Virtual Environment Detection
- **Virtual Processes:** vmtoolsd, citrix, teradici
- **Performance:** CPU/RAM Usage for virtual environment optimization

### Enterprise Apps
- **Teams:** New Teams 2.0 (ms-teams process)
- **Outlook:** OST File Size Check
- **OneDrive:** Sync Status
- **Java:** Version Check (Enterprise Apps require Java 11+)

## Output Format

### Status Levels
- 🚨 **CRITICAL:** Immediate attention required
- ⚠️ **WARNING:** Monitoring recommended
- ℹ️ **INFO:** Normal information
- ✅ **OK:** Everything in order

### Summary
- **CRITICAL:** Red issues
- **WARNING:** Yellow warnings  
- **INFO:** Blue information
- **OVERALL STATUS:** Overall assessment

## InfraCommand Integration

Both scripts are already integrated in `grid.config.json`:
- **Baseline:** Standard check
- **Baseline Advanced:** Extended analysis

## Virtual Environment Optimization

The scripts are optimized for virtual environments:
- Fast checks (fewer WMI queries)
- Virtual environment specific performance metrics
- Enterprise specific warnings
- Robust error handling

## Error Handling

- All checks are in `try-catch` blocks
- Script doesn't break on individual errors
- Detailed error messages
- Graceful degradation
