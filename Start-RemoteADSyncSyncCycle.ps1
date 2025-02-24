function Start-RemoteADSyncSyncCycle {
<#
.SYNOPSIS
    Initiates an AD sync cycle on a remote Azure AD Connect server.

.DESCRIPTION
    This function starts an Azure AD synchronization cycle on a remote server running Azure AD Connect.
    It handles credential management, retries on busy conditions, and provides detailed error reporting.

.PARAMETER RemoteComputer
    The name of the remote server running Azure AD Connect.

.PARAMETER Credential
    PSCredential object containing administrative credentials for the remote server.
    If not provided, the script will prompt for credentials.

.PARAMETER Policy
    The type of sync cycle to run. Valid values are "Delta" (incremental) or "Full" (complete).
    Defaults to "Delta".

.EXAMPLE
    Start-RemoteADSyncSyncCycle -RemoteComputer "ADConnectServer"
    Starts a delta sync on the specified server, prompting for credentials.

.EXAMPLE
    Start-RemoteADSyncSyncCycle -RemoteComputer "ADConnectServer" -Policy Full -Credential $cred
    Starts a full sync using pre-existing credentials.

.NOTES
    Requires:
    - WinRM enabled on the remote server
    - Administrative privileges
    - ADSync PowerShell module installed on the remote server
    
    Author: differentstuff
    Version: 1.0
    Last Modified: 24.02.2025
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage="Enter the AD Connect server name")]
        [ValidateNotNullOrEmpty()]
        [string]$RemoteComputer,
        
        [Parameter(Mandatory=$false,
                   HelpMessage="Provide admin credentials for AD Sync operation")]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential,
        [Parameter(Mandatory=$false,
                   HelpMessage="Enter the AD Sync Policy")]
        [ValidateSet("Delta","Full")]
        [string]$Policy = "Delta"
    )

    Write-Output "- Starting Remote AD Sync on $RemoteComputer -"

    # Credential handling
    if (-not $Credential) {
        try {
            $Credential = Get-Credential -Message "Enter admin credentials for AD Sync operation"
            if (-not $Credential) {
                throw "Credentials are required for AD Sync operation"
            }
        }
        catch {
            Write-Output "Error getting credentials: $_"
            return
        }
    }

    # Define retry parameters
    $maxRetries = 3
    $retryCount = 0
    $retryDelay = 15 # seconds

    do {
        try {
            Write-Verbose "Attempting to connect to $RemoteComputer (Attempt $($retryCount + 1) of $maxRetries)"
            
            $invokeParams = @{
                ComputerName = $RemoteComputer
                ScriptBlock = { 
                    Import-Module 'ADSync'
                    Start-ADSyncSyncCycle -PolicyType $Policy
                }
                ArgumentList = $Policy
                Credential = $Credential
                ErrorAction = 'Stop'
            }

            # Execute sync remotely
            $result = Invoke-Command @invokeParams
            
            Write-Output "Sync initiated at $(Get-Date -Format 'HH:mm:ss')"
            Write-Output "Result: $($result.Result)"
            break # Success, exit the loop
        }
        catch {
            $retryCount++
            
            # Check specifically for System.InvalidOperationException
            if ($_.Exception.GetType().FullName -eq 'System.InvalidOperationException') {
                if ($retryCount -lt $maxRetries) {
                    Write-Output "AD Sync is currently busy. Waiting $retryDelay seconds before retry ($retryCount of $maxRetries)..."
                    Write-Verbose "Exception details: $($_.Exception.Message)"
                    Start-Sleep -Seconds $retryDelay
                    continue
                }
                else {
                    Write-Output "Error: AD Sync is busy. Maximum retry attempts reached. Please try again later."
                }
            }
            else {
                Write-Output "Error during remote AD Sync:"
                Write-Output "Message: $($_.Exception.Message)"
                Write-Output "`nPlease ensure:"
                Write-Output "- WinRM is enabled on the remote server"
                Write-Output "- Your credentials have appropriate permissions"
                Write-Output "- The ADSync PowerShell module is installed on the remote server"
                Write-Verbose "Full Error Details: $_"
            }
            return
        }
    } while ($retryCount -lt $maxRetries)
}
