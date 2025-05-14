<#
.SYNOPSIS
    Manages and removes keyboard layouts in the Windows registry.

.DESCRIPTION
    This script allows users to manage and remove keyboard layouts from the Windows registry. 
    It provides functionality to remove specific layouts, including the primary layout, while 
    ensuring the system remains in a valid state. The script supports backup creation, 
    restoration, and validation of registry changes to prevent errors.

.PARAMETER Force
    Skips confirmation prompts for removal actions.

.PARAMETER NoBackup
    Prevents the creation of a backup before making changes to the registry.

.PARAMETER IndexesToRemove
    Specifies the indexes of keyboard layouts to remove. Must be numeric values.

.PARAMETER WhatIf
    Displays what actions the script would take without making any changes.

.EXAMPLE
    .\Update-KeyboardLanguage.ps1
    Displays the current keyboard layouts and prompts the user to select layouts for removal.

.EXAMPLE
    .\Update-KeyboardLanguage.ps1 -IndexesToRemove 2,3
    Removes the keyboard layouts with indexes 2 and 3.

.EXAMPLE
    .\Update-KeyboardLanguage.ps1 -Force -NoBackup
    Removes keyboard layouts without confirmation and skips backup creation.

.EXAMPLE
    .\Update-KeyboardLanguage.ps1 -WhatIf
    Displays the actions the script would take without making any changes.

.NOTES
    File Name      : Update-KeyboardLanguage.ps1
    Author         : Jean-Marie Heck (hecjex)
    Prerequisite   : Administrative privileges
    Version        : 1.0
    Last Modified  : 07.02.2025

    Change Log:
    v1.0 - Initial release

.LINK
    https://learn.microsoft.com/en-us/powershell/scripting/overview
    https://learn.microsoft.com/en-us/windows/win32/intl/keyboard-layouts

.COMPONENT
    Windows Registry Management

.FUNCTIONALITY
    - Displays current keyboard layouts
    - Removes specified keyboard layouts
    - Handles primary layout replacement
    - Creates and restores backups
    - Validates registry changes
    - Provides detailed error handling and logging
#>


#region Parameters

[CmdletBinding()]
param(
    # Parameter help description
    [Parameter(Mandatory=$false)]
    [switch]$Force,              # Skip confirmations

    [Parameter(Mandatory=$false)]
    [switch]$NoBackup,           # Skip backup creation

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_ -match '^\d+$') { 
            return $true 
        }
        throw "Index '$_' must be a number"
    })]
    [string[]]$IndexesToRemove,  # Directly specify indexes to remove

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf              # Show what would happen
)

#endregion Parameters


#region Constants

$PRELOADPATH = "Registry::HKEY_USERS\.DEFAULT\Keyboard Layout\Preload"
$PRELOADBACKUPPATH = "Registry::HKEY_USERS\.DEFAULT\Keyboard Layout\Preload_Backup"
$KEYBOARDLAYOUTSPATH = "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layouts"

#endregion Constants


#region Functions

function Test-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-CurrentKeyboardLayouts {
    $currentEntries = @()
    $preloadEntries = Get-ItemProperty -Path $PRELOADPATH -ErrorAction Stop | 
        Select-Object -Property * -ExcludeProperty PS*

    if (-not $preloadEntries) {
        Write-Host "No keyboard layouts found." -ForegroundColor Yellow
        Exit-Script -Code 0
    }
    
    if ($preloadEntries) {
        $preloadEntries | ForEach-Object {
            $_.PSObject.Properties | ForEach-Object {
                $layoutID = $_.Value
                $layoutName = Get-KeyboardLayoutName $layoutID
                $currentEntries += [PSCustomObject]@{
                    Index = $_.Name
                    LayoutID = $layoutID
                    LayoutName = $layoutName
                }
            }
        }
    }
    return $currentEntries
}

function Get-KeyboardLayoutName {
    [CmdletBinding()]
    param (
        [string]$layoutID
    )
    Write-Verbose "Getting layout name for ID: $layoutID"
    try {
        # Validate the layout ID format (8 characters, hexadecimal)
        if (-not $layoutID -or $layoutID -notmatch '^00[0-9a-f]{6}$') {
            return "Invalid Layout ID"
        }

        $layoutPath = "$KEYBOARDLAYOUTSPATH\$layoutID"
        $layoutName = (Get-ItemProperty -Path $layoutPath -ErrorAction Stop)."Layout Text"
        return $layoutName
    }
    catch {
        return "Unknown Layout"
    }
}

function Test-RegistryAccess {
    param(
        [string]$Path
    )
    try {
        $null = Get-ItemProperty -Path $Path -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log "Cannot access registry path: $Path" -Level 'ERROR'
        return $false
    }
}

function Remove-BackupIfExists {
    if (Test-Path $PRELOADBACKUPPATH) {
        Write-Verbose "Removing existing backup"
        try {
            Remove-Item -Path $PRELOADBACKUPPATH -Recurse -Force -ErrorAction Stop
            Write-Verbose "Backup removed successfully"
            return $true
        }
        catch {
            Write-Log "Failed to remove backup: $_" -Level 'ERROR'
            return $false
        }
    }
    return $true
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    
    switch ($Level) {
        'INFO'  { Write-Host $Message -ForegroundColor White }
        'WARN'  { Write-Host $Message -ForegroundColor Yellow }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
    }
}

function Test-StateValidity {
    param(
        [array]$CurrentEntries,
        [array]$NewEntries
    )
    # Ensure we always have at least one layout
    if ($NewEntries.Count -eq 0) {
        Write-Log "Invalid state: No layouts would remain" -Level 'ERROR'
        return $false
    }
    # Ensure layout '1' exists
    if (-not ($NewEntries | Where-Object { $_.Index -eq "1" })) {
        Write-Log "Invalid state: Primary layout would be removed" -Level 'ERROR'
        return $false
    }
    return $true
}

function Exit-Script {
    param(
        [Parameter(Mandatory=$true)]
        [int]$Code
    )
    Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
        $null = Read-Host
        exit $Code
}
#endregion Functions


#region Main

# Check for admin rights
if (-not (Test-AdminRights)) {
    Write-Host "This script requires administrative privileges. Please run as administrator." -ForegroundColor Red
    Exit-Script -Code 1
}

# Test registry access for all required paths
$requiredPaths = @($PRELOADPATH, $KEYBOARDLAYOUTSPATH)
foreach ($path in $requiredPaths) {
    if (-not (Test-RegistryAccess -Path $path)) {
        Write-Host "Cannot access required registry path: $path" -ForegroundColor Red
        Exit-Script -Code 1
    }
}

try {
    # Get all entries and create a list with their details
    $entries = Get-CurrentKeyboardLayouts

    # Display all entries
    Write-Host "`nCurrent keyboard layouts:" -ForegroundColor Cyan
    $entries | Format-Table -AutoSize

    # If only one entry exists, prevent removal
    if ($entries.Count -eq 1) {
        Write-Host "Only one keyboard layout exists. Cannot remove the last layout." -ForegroundColor Yellow
        Exit-Script -Code 0
    }

    # Get valid indexes for removal (excluding 1)
    $validIndexes = $entries | Select-Object -ExpandProperty Index

    if ($validIndexes.Count -eq 0) {
        Write-Host "No layouts available for removal." -ForegroundColor Yellow
        Exit-Script -Code 0
    }

    # Use IndexesToRemove parameter if provided, otherwise ask user
    if ($IndexesToRemove) {
        $toRemove = $IndexesToRemove
        $validInput = $true
        foreach ($index in $toRemove) {
            if (-not ($index -in $validIndexes)) {
                Write-Host "Invalid index provided as parameter: $index" -ForegroundColor Red
                Exit-Script -Code 1
            }
        }
    }
    else {
    # Ask user which entries to remove
        do {
            Write-Host "`nAvailable indexes for removal: $($validIndexes -join ', ')" -ForegroundColor Yellow
            Write-Host "Enter the indexes to remove (comma-separated, e.g., '2,3' or 'all' for all except 1): " -ForegroundColor Yellow
            $response = Read-Host

            if ($response -eq "all") {
                $toRemove = $validIndexes | Where-Object { $_ -ne "1" }  # Exclude 1 only for 'all' case
                $validInput = $true
            }
            else {
                $toRemove = $response.Split(",").Trim()
                $validInput = $true
    
                # Validate each index
                foreach ($index in $toRemove) {
                    if (-not ($index -in $validIndexes)) {
                        Write-Host "Invalid index: $index" -ForegroundColor Red
                        $validInput = $false
                        break
                    }
                }
            }
        } while (-not $validInput)
    }

    # Check if index 1 is being removed
    if ($toRemove -contains "1") {
        Write-Host "`nYou have chosen to remove the primary layout (Index 1)." -ForegroundColor Yellow

        # Get the remaining valid indexes (excluding those marked for removal)
        $remainingIndexes = $validIndexes | Where-Object { $_ -notin $toRemove }

        if ($remainingIndexes.Count -eq 0) {
            Write-Host "No valid layouts are available to replace the primary layout. Aborting." -ForegroundColor Red
            Exit-Script -Code 1
        }

        # Display remaining layouts
        Write-Host "`nAvailable layouts to set as the new primary layout:" -ForegroundColor Cyan
        $entries | Where-Object { $_.Index -in $remainingIndexes } | Format-Table -AutoSize

        # Prompt the user to select a replacement for index 1
        do {
            Write-Host "Enter the index of the layout to set as the new primary layout: " -ForegroundColor Yellow
            $newPrimaryIndex = Read-Host

            if (-not ($newPrimaryIndex -in $remainingIndexes)) {
                Write-Host "Invalid index: $newPrimaryIndex" -ForegroundColor Red
                $validInput = $false
            }
            else {
                $validInput = $true
            }
        } while (-not $validInput)

        # Update the registry to set the new primary layout
        try {
            # Store the original layout ID of index 1
            $originalPrimaryLayoutID = ($entries | Where-Object { $_.Index -eq "1" }).LayoutID
            
            # Get the new primary layout ID
            $newPrimaryLayoutID = ($entries | Where-Object { $_.Index -eq $newPrimaryIndex }).LayoutID
            # Perform the swap in registry
            Set-ItemProperty -Path $PRELOADPATH -Name "1" -Value $newPrimaryLayoutID -ErrorAction Stop
            Set-ItemProperty -Path $PRELOADPATH -Name $newPrimaryIndex -Value $originalPrimaryLayoutID -ErrorAction Stop
            
            $oldPrimaryLayoutName = $($entries | Where-Object { $_.Index -eq "1" }).LayoutName
            $newPrimaryLayoutName = ($entries | Where-Object { $_.Index -eq $newPrimaryIndex }).LayoutName
            Write-Host "Primary layout has been changed from $oldPrimaryLayoutName to $newPrimaryLayoutName" -ForegroundColor Green

            # Update $toRemove to reflect the new position of the old index 1 layout
            $toRemove = $toRemove | ForEach-Object {
                if ($_ -eq "1") {
                    $newPrimaryIndex  # This is where the old index 1 layout moved to
                } else {
                    $_  # Keep other indexes unchanged
                }
            }

            # Refresh entries after the swap
            $entries = Get-CurrentKeyboardLayouts
            
        }
        catch {
            Write-Host "Failed to update the primary layout: $_" -ForegroundColor Red
            Exit-Script -Code 1
        }
    }

    # Calculate new state and validate it
    $newEntries = $entries | Where-Object { $_.Index -notin $toRemove }
    if (-not (Test-StateValidity -CurrentEntries $entries -NewEntries $newEntries)) {
        Write-Host "The requested changes would result in an invalid state." -ForegroundColor Red
        Exit-Script -Code 1
    }

    # Confirm removal unless Force switch is used
    if (-not $Force) {
        Write-Host "`nYou are about to remove these entries:" -ForegroundColor Yellow
        $entries | Where-Object { $_.Index -in $toRemove } | Format-Table -AutoSize

        do {
            $confirm = Read-Host "Are you sure? (y/n)"
        } while ($confirm -notmatch '^[yn]$')

        if ($confirm -ne "y") {
            Write-Host "Operation cancelled" -ForegroundColor Yellow
            Exit-Script -Code 0
        }
    }

    # Handle backup
    if (-not $NoBackup) {
        if (-not (Remove-BackupIfExists)) {
            Write-Host "Failed to remove existing backup. Aborting for safety." -ForegroundColor Red
            Exit-Script -Code 1
        }
        Copy-Item -Path $PRELOADPATH -Destination $PRELOADBACKUPPATH -Recurse
    }

    # Remove entries
    foreach ($index in $toRemove) {
        try {
            if ($WhatIf) {
                Write-Host "WhatIf: Would remove entry $index" -ForegroundColor Yellow
            }
            else {
                Remove-ItemProperty -Path $PRELOADPATH -Name $index -ErrorAction Stop
                Write-Host "Removed entry $index" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Failed to remove entry $($index): $_" -ForegroundColor Red
            if (-not $NoBackup) {
                Write-Host "Attempting to restore from backup..." -ForegroundColor Yellow
                Remove-Item -Path $PRELOADPATH -Recurse -Force
                Copy-Item -Path $PRELOADBACKUPPATH -Destination $PRELOADPATH -Recurse
                Write-Host "Restored from backup" -ForegroundColor Green
            }
            Exit-Script -Code 1
        }
    }

    # Cleanup backup if successful and not in WhatIf mode
    if (-not $NoBackup -and -not $WhatIf) {
        $null = Remove-BackupIfExists
    }

    # Show final state
    if (-not $WhatIf) {
        Write-Host "`nFinal keyboard layouts:" -ForegroundColor Cyan
        Get-CurrentKeyboardLayouts | Format-Table -AutoSize
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    if (-not $NoBackup) {
        if (Test-Path $PRELOADBACKUPPATH) {
            Write-Host "Attempting to restore from backup..." -ForegroundColor Yellow
            try {
                Remove-Item -Path $PRELOADPATH -Recurse -Force -ErrorAction Stop
                Copy-Item -Path $PRELOADBACKUPPATH -Destination $PRELOADPATH -Recurse -ErrorAction Stop
                Write-Host "Restored from backup" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to restore from backup: $_" -ForegroundColor Red
                Write-Host "`nBackup values (for manual restoration):" -ForegroundColor Yellow
                try {
                    $backupEntries = Get-ItemProperty -Path $PRELOADBACKUPPATH -ErrorAction Stop | 
                        Select-Object -Property * -ExcludeProperty PS*

                    Write-Host "`nRegistry Path: $PRELOADPATH" -ForegroundColor Cyan
                    Write-Host "Values to set:" -ForegroundColor Cyan
                    $backupEntries.PSObject.Properties | ForEach-Object {
                        $layoutName = Get-KeyboardLayoutName $_.Value
                        Write-Host "Name: $($_.Name), Value: $($_.Value) ($layoutName)" -ForegroundColor White
                    }
                }
                catch {
                    Write-Host "Failed to read backup values: $_" -ForegroundColor Red
                }
            }
        }
    }
    else {
        Write-Host "No backup was created (-NoBackup was specified)" -ForegroundColor Yellow
        Write-Host "`nOriginal values (for manual restoration):" -ForegroundColor Yellow
        try {
            Write-Host "`nRegistry Path: $PRELOADPATH" -ForegroundColor Cyan
            Write-Host "Values to set:" -ForegroundColor Cyan
            $entries | ForEach-Object {
                Write-Host "Name: $($_.Index), Value: $($_.LayoutID) ($($_.LayoutName))" -ForegroundColor White
            }
        }
        catch {
            Write-Host "Failed to display original values: $_" -ForegroundColor Red
        }
    }
    Exit-Script -Code 1
}

#endregion Main