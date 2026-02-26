<#
.SYNOPSIS
    Full cleanup of ManageEngine UEMS Agent before reinstall.
.NOTES
    Run as Administrator. Reboot after execution.
#>

#Requires -RunAsAdministrator

$guid = "{6AD2231F-FF48-4D59-AC26-405AFAE23DB7}" # MSI GUID for UEMS_Agent

$ErrorActionPreference = "SilentlyContinue"
$log = "C:\Temp\UEMS_Cleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null

function Write-Log {
    param([string]$msg)
    $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"
    Write-Host $line
    Add-Content -Path $log -Value $line
}

Write-Log "=== UEMS Agent Cleanup Started ==="

# 1. STOP & DISABLE SERVICES
Write-Log "Stopping services..."
@("ManageEngine UEMS Agent", "DCFAService", "DCFA") | ForEach-Object {
    $svc = Get-Service -Name $_ -ErrorAction SilentlyContinue
    if ($svc) {
        Stop-Service -Name $_ -Force
        Set-Service -Name $_ -StartupType Disabled
        Write-Log "  Stopped: $_"
    }
}

# 2. KILL PROCESSES
Write-Log "Killing processes..."
@("dcagentservice", "dcagent", "DCFA", "uems") | ForEach-Object {
    $procs = Get-Process -Name $_ -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Stop-Process -Force
        Write-Log "  Killed: $_"
    }
}

# 3. MSI UNINSTALL
Write-Log "Attempting MSI uninstall..."
Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn REBOOT=ReallySuppress" -Wait
Write-Log "  MSI uninstall attempted for $guid"

# 4. DELETE INSTALL FOLDERS
Write-Log "Removing install folders..."
@(
    "C:\Program Files (x86)\ManageEngine\UEMS_Agent",
    "C:\Program Files\ManageEngine\UEMS_Agent",
    "C:\ProgramData\ManageEngine\UEMS_Agent"
) | ForEach-Object {
    if (Test-Path $_) {
        Remove-Item -Path $_ -Recurse -Force
        Write-Log "  Deleted: $_"
    }
}

# 5. CLEAN UNINSTALL REGISTRY ENTRIES
Write-Log "Cleaning Uninstall registry keys..."
@(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
) | ForEach-Object {
    Get-ChildItem $_ | ForEach-Object {
        $name = $_.GetValue("DisplayName")
        if ($name -like "*ManageEngine*" -or $name -like "*UEMS*" -or $name -like "*Zoho*") {
            Remove-Item $_.PSPath -Recurse -Force
            Write-Log "  Removed uninstall key: $name"
        }
    }
}

# 6. CLEAN MSI PRODUCT CACHE
Write-Log "Cleaning MSI product cache..."
@(
    "HKLM:\SOFTWARE\Classes\Installer\Products",
    "HKLM:\SOFTWARE\Classes\Installer\Features"
) | ForEach-Object {
    $basePath = $_
    Get-ChildItem $basePath | ForEach-Object {
        $name = $_.GetValue("ProductName")
        if ($name -like "*ManageEngine*" -or $name -like "*UEMS*") {
            Remove-Item $_.PSPath -Recurse -Force
            Write-Log "  Removed MSI cache key: $name"
        }
    }
}

# 7. CLEAN SERVICE & GENERAL REGISTRY
Write-Log "Cleaning service and product registry keys..."
@(
    "HKLM:\SOFTWARE\ManageEngine\UEMS",
    "HKLM:\SOFTWARE\WOW6432Node\ManageEngine\UEMS",
    "HKLM:\SYSTEM\CurrentControlSet\Services\DCFAService",
    "HKLM:\SYSTEM\CurrentControlSet\Services\DCFA"
) | ForEach-Object {
    if (Test-Path $_) {
        Remove-Item -Path $_ -Recurse -Force
        Write-Log "  Removed: $_"
    }
}

Write-Log "=== Cleanup Complete. Log saved to: $log ==="
Write-Log ">>> REBOOT REQUIRED <<<"