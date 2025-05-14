<#
.SYNOPSIS
    Disables the "HiberbootEnabled" registry setting on remote computers.

.DESCRIPTION
    This script connects to remote computers in an Active Directory environment and disables the 
    "HiberbootEnabled" registry setting, which controls Fast Startup. It filters computers based on 
    a specified name pattern and ensures proper error handling and logging throughout the process. 
    The script requires administrative privileges and uses PowerShell remoting to modify the registry 
    on each target computer.

.PARAMETER Filter
    Specifies the name filter for selecting computers in Active Directory. 
    Default: "*" (matches all computer)
    Alt: "Company*" (matches all computer names starting with 'Company')

.PARAMETER ErrorActionPreference
    Sets the error handling behavior for the script. Default: "Stop".

.EXAMPLE
    .\Disable-HibernateRemoteAll.ps1
    Disables the "HiberbootEnabled" registry setting on all computers matching the default filter "TC*".

.EXAMPLE
    .\Disable-HibernateRemoteAll.ps1 -Filter "PC*"
    Disables the "HiberbootEnabled" registry setting on all computers matching the filter "PC*".

.NOTES
    File Name      : Disable-HibernateRemoteAll.ps1
    Author         : Jean-Marie Heck (hecjex)
    Prerequisite   : Active Directory PowerShell Module, PowerShell Remoting
    Version        : 1.1
    Last Modified  : 06.02.2025

    Change Log:
    v1.0 - Initial release

.LINK
    https://learn.microsoft.com/en-us/powershell/scripting/overview
    https://learn.microsoft.com/en-us/windows/win32/power/system-power-management

.COMPONENT
    PowerShell Remoting, Active Directory PowerShell Module

.FUNCTIONALITY
    - Connects to remote computers using PowerShell remoting
    - Checks the "HiberbootEnabled" registry setting
    - Disables "HiberbootEnabled" if it is enabled
    - Logs all actions and errors
    - Skips offline computers
#>

#region Param
$Filter = "*" ##MODIFY HERE
$ErrorActionPreference = "Stop"  # Use Stop to ensure proper error handling
#endregion Param

#region Constants
$HKEY_LM_PATH = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
#endregion Constants

#region Logging
function Write-LogMessage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('Info','Warning','Error')]
        [string]$Type = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    switch($Type) {
        'Info'    { Write-Host $logMessage -ForegroundColor Gray }
        'Warning' { Write-Warning $logMessage }
        'Error'   { Write-Error $logMessage }
    }
}
#endregion Logging

#region Functions

function Test-AdminPrivileges {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#endregion Functions

#region Main

# Check for admin privileges
if (-not (Test-AdminPrivileges)) {
    Write-LogMessage "This script requires administrative privileges. Please run as administrator." -Type Error
    exit 1
}

$allPCs = Get-ADComputer -Filter {Enabled -eq $true -and Name -like $Filter} -Properties Name | Select-Object -ExpandProperty Name

foreach($pc in $allPCs){
    $sess = $null
    #$VerbosePreference = "Continue"
    try{

        # Test if computer is online first
        if (Test-Connection $pc -Count 1 -Quiet) {
            Write-Host "Connecting to $pc..."

            # Reduce timeout for Session
            $psoptions = New-PSSessionOption -OpenTimeout 30000 -OperationTimeout 30000 -IdleTimeout 60000
            $sess = New-PSSession -ComputerName $pc -SessionOption $psoptions -ErrorAction Stop

            Write-LogMessage "--------------------------------------------------------" -Type Info
            
            # Replace the file check (around line 64) with:
            $regValue = Invoke-Command -Session $sess -ScriptBlock {
                param($HKEY_LM_PATH)
                try {
                    $value = (Get-ItemProperty -Path $HKEY_LM_PATH -Name "HiberbootEnabled" -ErrorAction Stop).HiberbootEnabled
                    return $value
                } catch {
                    return $null
                }
            } -ArgumentList $HKEY_LM_PATH

            if($regValue -eq 0){
                Write-LogMessage "HiberbootEnabled already disabled on $pc - skipping" -Type Info
                $OperationOK = $True
                Write-LogMessage "----" -Type Info
            }
            else{
                Write-LogMessage "Starting registry modification on $pc" -Type Info
                Remove-Variable OperationOK -ErrorAction SilentlyContinue
            }

            # Replace the code block around line 76 with:
            if(!($OperationOK -eq $True)){
                Invoke-Command -Session $sess -ScriptBlock {
                    param($HKEY_LM_PATH)
                    Set-ItemProperty -Path $HKEY_LM_PATH -Name "HiberbootEnabled" -Value 0 -Type DWord
                } -ArgumentList $HKEY_LM_PATH
            }

            if(!($OperationOK -eq $True)){
                

                #################### End

                Write-LogMessage "Registry configuration completed on $pc" -Type Info
                Write-LogMessage "--------------------------------------------------------" -Type Info
            }
        }
        else {
            Remove-Variable OperationOK -ErrorAction SilentlyContinue
            Write-Host "Computer $pc is offline - skipping" -ForegroundColor Yellow
        }

        if($OperationOK -eq $True){
            Remove-Variable OperationOK -ErrorAction SilentlyContinue
            Write-Host "Computer $pc is done - skipped" -ForegroundColor Green
        }

    }
    catch{
        Write-Host "Error on $pc : $_" -ForegroundColor Red
    }
    finally{
        if ($sess) {
            Remove-PSSession $sess
        }
    }
}

Write-Host
Write-Host "--- All Computers processed ---" -ForegroundColor Green
Write-Host
#endregion Main