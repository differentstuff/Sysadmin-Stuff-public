function Disable-PowerSavingForNetworkAdapters {

    # $SpecificAdapterName = "" # may be a param? needs Get-AllAdapterNames to get the correct name
    $PnPCapabilitiesValue = 0 # 0 = disable all

    # Requires elevated privileges (run as administrator)
    if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "This script requires administrator rights. Please run as administrator."
        exit
    }

    # Registration path for the adapter
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
    
    # Browse all subfolders to find the correct adapter
    $adapters = Get-ChildItem -Path $regPath 2>$null
    
    # Filter to only include keys that end with numeric values (like 0000, 0001, etc.)
    $adapters = $adapters | Where-Object {$_.PSChildName -match '^\d+$'}

    if($SpecificAdapterName){
        $adapters = $adapters | Where-Object {$_.PSChildName -eq $SpecificAdapterName}
        Write-Host "Found $($adapters.count) adapter registry entries"
    }

    # Retrieve all network adapters
    $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }

    foreach ($adapter in $networkAdapters) {
        Write-Host "Disable power saving options for: $($adapter.Name)"
    
        foreach ($adapterPath in $adapters) {
            if ($adapterPath.PSPath) {
                $instanceID = (Get-ItemProperty -Path $adapterPath.PSPath -Name "DriverDesc" -ErrorAction SilentlyContinue).DriverDesc 2>$null

                # Disable energy saving options if necessary
                if ($instanceID -eq $adapter.InterfaceDescription) {
                    $PnPCapabilitiesCurrent = Get-ItemProperty -Path $adapterPath.PSPath -Name "PnPCapabilities" -ErrorAction SilentlyContinue

                    if ($PnPCapabilitiesCurrent -ne $null -and $PnPCapabilitiesValue -ne $PnPCapabilitiesCurrent.PnPCapabilities) {
                        Write-Host "Disabling energy saving options for: $($adapter.Name) (Current value: $($PnPCapabilitiesCurrent.PnPCapabilities))" -ForegroundColor Green
                        Set-ItemProperty -Path $adapterPath.PSPath -Name "PnPCapabilities" -Value $PnPCapabilitiesValue -Type DWord -Force
                    } elseif ($PnPCapabilitiesCurrent -eq $null) {
                        Write-Host "Setting energy saving options for: $($adapter.Name) (No current value)" -ForegroundColor Green
                        New-ItemProperty -Path $adapterPath.PSPath -Name "PnPCapabilities" -Value $PnPCapabilitiesValue -PropertyType DWord -Force
                    } else {
                        Write-Host "Energy saving options already disabled for: $($adapter.Name)" -ForegroundColor Cyan
                    }
                }
            }
        }
    }

    Write-Host "`nChanges will take effect after a reboot." -ForegroundColor Yellow
    $restart = Read-Host -Prompt "Do you want to restart now? [Y/N]"

    if ($restart.ToUpper() -eq "J") {
        Write-Host "Restarting your computer in 10 seconds." -ForegroundColor Red
        Restart-Computer -Wait 10
    } else {
        Write-Host "Please restart your computer for the changes to take effect." -ForegroundColor Yellow
    }
}