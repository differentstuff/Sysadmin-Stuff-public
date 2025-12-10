# Parameter to accept action from command line
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Install", "Uninstall")]
    [string]$Action
)

# Function to manage Windows 11 context menu
function Update-Win11ContextMenu {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Install", "Uninstall")]
        [string]$Action
    )
    
    # Check if Windows 11 (build 22000+)
    if ((Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber -lt 22000) {
        return
    }
    
    # Define registry paths
    $regPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
    $inprocServer32Path = "$regPath\InprocServer32"
    
    if ($Action -eq "Install") {
        # Create registry keys if they don't exist
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        if (-not (Test-Path $inprocServer32Path)) { New-Item -Path $inprocServer32Path -Force | Out-Null }
        
        # Set default values to empty string
        Set-ItemProperty -Path $regPath -Name "(Default)" -Value "" -Force | Out-Null
        Set-ItemProperty -Path $inprocServer32Path -Name "(Default)" -Value "" -Force | Out-Null
    }
    else { # Uninstall
        # Remove the registry keys
        if (Test-Path $regPath) { Remove-Item -Path $regPath -Recurse -Force | Out-Null }
    }
    
    # Restart File Explorer to apply changes
    Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process "explorer.exe"
}

# Call the function with the parameter from command line
Update-Win11ContextMenu -Action $Action

# Example usage:
# To install (enable classic context menu):
# powershell.exe -ExecutionPolicy Bypass -WindowStyle hidden -File Update-Win11ContextMenu.ps1 -Action "Install"

# To uninstall (restore default context menu):
# powershell.exe -ExecutionPolicy Bypass -WindowStyle hidden -File Update-Win11ContextMenu.ps1 -Action "Uninstall"