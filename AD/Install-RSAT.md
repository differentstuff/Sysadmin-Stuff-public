# as local admin
$regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
Set-ItemProperty -Path $regPath -Name "UseWUServer" -Value 0
Restart-Service wuauserv

## Install RSAT
Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online

## re-activate WSUS
Set-ItemProperty -Path $regPath -Name "UseWUServer" -Value 1
Restart-Service wuauserv


# Mount Win11 iso in D:\
$source = "D:\sources\sxs"

# v1
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -Source $source -LimitAccess

# v2
New-Item -ItemType Directory -Path "C:\WIM_Mount"
dism /Mount-Wim /WimFile:"D:\sources\install.wim" /Index:6 /MountDir:"C:\WIM_Mount" /ReadOnly

# v3
DISM /Online /Add-Capability /CapabilityName:"Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" /LimitAccess /Source:"D:\sources\sxs"

# v4
Add-WindowsCapability -Online `
  -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" `
  -Source "C:\WIM_Mount" `
  -LimitAccess
