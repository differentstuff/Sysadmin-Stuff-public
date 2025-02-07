<#
.SYNOPSIS
    Executes a provided PowerShell script block on remote computers in Active Directory.

.DESCRIPTION
    This script connects to remote computers in Active Directory matching a specified name filter
    and executes a provided PowerShell script block on each computer. It includes comprehensive
    logging, error handling, and requires administrative privileges. The script uses PowerShell
    remoting with optimized session timeouts and validates computer availability before attempting
    connections.

.PARAMETER Filter
    Specifies the name pattern for filtering Active Directory computers.
    Default value: "" (No Filter: matches all computer names)

.PARAMETER ErrorActionPreference 
    Determines how the script responds to errors.
    Default value: "Stop" (ensures proper error handling by stopping on errors)

.PARAMETER RemoteScriptBlock
    The PowerShell script block to execute on each remote computer.
    This parameter is mandatory and must be provided as a string.

.EXAMPLE
    .\Execute-RemoteScript.ps1 -RemoteScriptBlock '$env:COMPUTERNAME'
    Retrieves the computer name from all machines matching the default filter

.EXAMPLE
    .\Execute-RemoteScript.ps1 -Filter "PC*" -RemoteScriptBlock 'Get-Service | Where-Object {$_.Status -eq "Running"}'
    Lists running services on all computers matching "PC*"

.EXAMPLE
    $script = 'Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name HiberbootEnabled'
    .\Execute-RemoteScript.ps1 -RemoteScriptBlock $script
    Checks the HiberbootEnabled registry value on remote computers

.NOTES
    File Name      : Execute-RemoteScript.ps1
    Author         : Jean-Marie Heck (differentstuff)
    Prerequisite   : Active Directory PowerShell Module, PowerShell Remoting, Admin Rights
    Version        : 1.0
    Last Modified  : 07.02.2025

    Change Log:
    v1.0 - Initial release

.LINK
    https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/running-remote-commands

.COMPONENT
    PowerShell Remoting
    Active Directory PowerShell Module

.FUNCTIONALITY
    - Filters and retrieves computers from Active Directory
    - Tests computer availability before connection
    - Establishes optimized remote PowerShell sessions
    - Executes custom script blocks remotely
    - Provides detailed logging of operations
    - Handles errors and session cleanup
    - Validates administrative privileges
#>


#region Param

param(
    [Parameter(Mandatory=$false)]
    [string]$Filter = "",  # Prefix/Suffix of your Device Names as RegEx: PC-* , DC01 , *PC* , ...
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.ActionPreference]$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop,  # Stop to ensure proper error handling
    [Parameter(Mandatory=$true)]
    [string]$RemoteScriptBlock = ""

)

#endregion Param


#region Logging
function Write-LogMessage {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Message = "",
        [ValidateSet('Info','Warning','Error','Empty')]
        [string]$Type = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    switch($Type) {
        'Info'    { Write-Host $logMessage -ForegroundColor Gray }
        'Warning' { Write-Warning $logMessage }
        'Error'   { Write-Error $logMessage }
        'Empty'   { Write-Host }
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
            Write-LogMessage "Connecting to $($pc)..." -Type Info

            # Reduce timeout for Session
            $psoptions = New-PSSessionOption -OpenTimeout 30000 -OperationTimeout 30000 -IdleTimeout 60000
            $sess = New-PSSession -ComputerName $pc -SessionOption $psoptions -ErrorAction Stop

            Write-LogMessage "--------------------------------------------------------" -Type Info
            
            # Replace the file check (around line 64) with:
            $scriptResult = Invoke-Command -Session $sess -ScriptBlock {
                param($ScriptToRun)
                try {
                    # Convert string to scriptblock and execute
                    $scriptBlock = [ScriptBlock]::Create($ScriptToRun)
                    return & $scriptBlock
                }
                catch {
                    return @{
                        Status = "Error"
                        Error = $_.Exception.Message
                    }
                }
            } -ArgumentList $RemoteScriptBlock
        }

        # Log the result
        if ($scriptResult.Status -eq "Success") {
            Write-LogMessage "Script executed successfully on $pc" -Type Info
            Write-LogMessage "Result: $($scriptResult.Result)" -Type Info
        }
        else {
            Write-LogMessage "Script execution failed on $pc" -Type Error
            Write-LogMessage "Error: $($scriptResult.Error)" -Type Error
        }
    }
    catch{
        Write-LogMessage "Error executing script on $($pc): $_" -Type Error
    }
    finally{
        if ($sess) {
            Remove-PSSession $sess
        }
    }
}

Write-LogMessage -Type Empty
Write-LogMessage "--- All Computers processed ---" -Type Info
Write-LogMessage -Type Empty

#endregion Main